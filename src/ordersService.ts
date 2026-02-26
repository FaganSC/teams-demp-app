import { TableClient, TableServiceClient, TableTransaction } from "@azure/data-tables";

export type OrderStatus = "Submitted" | "Pending" | "Processing" | "Shipped" | "Delivered" | "Cancelled";

export interface Order {
  id: string;
  customer: string;
  amount: number;
  status: OrderStatus;
  date: string;
}

const TABLE_NAME = "Orders";
const PARTITION_KEY = "Orders";

const CUSTOMERS = [
  "Contoso Ltd.",       "Fabrikam Inc.",        "Northwind Traders",
  "Adventure Works",    "Tailspin Toys",         "Woodgrove Bank",
  "Proseware Inc.",     "Lucerne Publishing",    "Humongous Insurance",
  "Wide World Importers", "Fourth Coffee",       "Alpine Ski House",
  "Coho Winery",        "Relecloud",             "Trey Research",
];

const STATUSES: OrderStatus[] = ["Submitted", "Pending", "Processing", "Shipped", "Delivered", "Cancelled"];

function randomDate(daysBack = 365): string {
  const now = new Date("2026-02-24");
  const offset = Math.floor(Math.random() * daysBack * 24 * 60 * 60 * 1000);
  return new Date(now.getTime() - offset).toISOString().slice(0, 10);
}

function generateOrders(count: number): Array<{
  partitionKey: string; rowKey: string;
  id: string; customer: string; amount: number; status: string; date: string;
}> {
  const orders = [];
  for (let i = 1; i <= count; i++) {
    const id = `ORD-${String(i).padStart(3, "0")}`;
    orders.push({
      partitionKey: PARTITION_KEY,
      rowKey: id,
      id,
      customer: CUSTOMERS[Math.floor(Math.random() * CUSTOMERS.length)],
      amount: Math.round((Math.random() * 9900 + 100) * 100) / 100,
      status: STATUSES[Math.floor(Math.random() * STATUSES.length)],
      date: randomDate(),
    });
  }
  return orders;
}

function getClients(connectionString: string) {
  return {
    service: TableServiceClient.fromConnectionString(connectionString),
    table: TableClient.fromConnectionString(connectionString, TABLE_NAME),
  };
}

/** Creates the Orders table and seeds 100 rows if it is empty. */
export async function seedIfEmpty(connectionString: string): Promise<void> {
  const { service, table } = getClients(connectionString);

  await service.createTable(TABLE_NAME).catch(() => { /* already exists */ });

  const iter = table.listEntities({ queryOptions: { select: ["rowKey"] } });
  const first = await iter.next();
  if (!first.done) {
    return; // already has data
  }

  const transaction = new TableTransaction();
  for (const order of generateOrders(100)) {
    transaction.createEntity(order);
  }
  await table.submitTransaction(transaction.actions);
}

/** Returns all orders sorted by ID descending. */
export async function listOrders(connectionString: string): Promise<Order[]> {
  const { table } = getClients(connectionString);
  const orders: Order[] = [];

  for await (const entity of table.listEntities<{
    id: string; customer: string; amount: number; status: string; date: string;
  }>()) {
    orders.push({
      id: entity.rowKey!,
      customer: entity.customer,
      amount: Number(entity.amount),
      status: entity.status as OrderStatus,
      date: entity.date,
    });
  }

  return orders.sort((a, b) => b.id.localeCompare(a.id));
}

/** Creates a new order, auto-incrementing the numeric portion of the ID. */
export async function createOrder(
  connectionString: string,
  data: Omit<Order, "id">,
  notify?: (order: Order) => void,
): Promise<Order> {
  const { table } = getClients(connectionString);

  let maxNum = 0;
  for await (const entity of table.listEntities({ queryOptions: { select: ["rowKey"] } })) {
    const match = entity.rowKey?.match(/^ORD-(\d+)$/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1], 10));
  }
  const id = `ORD-${String(maxNum + 1).padStart(3, "0")}`;

  await table.createEntity({
    partitionKey: PARTITION_KEY,
    rowKey: id,
    id,
    customer: data.customer,
    amount: data.amount,
    status: data.status,
    date: data.date,
  });

  const order: Order = { id, ...data, amount: Number(data.amount) };
  notify?.(order);
  return order;
}

/** Renames a customer on all their orders. Returns the number of orders updated. */
export async function renameCustomer(
  connectionString: string,
  oldName: string,
  newName: string,
): Promise<number> {
  const { table } = getClients(connectionString);
  const toUpdate: Order[] = [];

  for await (const entity of table.listEntities<{
    id: string; customer: string; amount: number; status: string; date: string;
  }>()) {
    if (entity.customer === oldName) {
      toUpdate.push({
        id: entity.rowKey!,
        customer: entity.customer,
        amount: Number(entity.amount),
        status: entity.status as OrderStatus,
        date: entity.date,
      });
    }
  }

  for (const order of toUpdate) {
    await table.updateEntity({
      partitionKey: PARTITION_KEY,
      rowKey: order.id,
      id: order.id,
      customer: newName,
      amount: order.amount,
      status: order.status,
      date: order.date,
    }, "Replace");
  }

  return toUpdate.length;
}

/** Updates an existing order in Table Storage. */
export async function updateOrder(
  connectionString: string,
  id: string,
  patch: Partial<Omit<Order, "id">>
): Promise<Order> {
  const { table } = getClients(connectionString);

  const existing = await table.getEntity<{
    id: string; customer: string; amount: number; status: string; date: string;
  }>(PARTITION_KEY, id);

  const updated = {
    partitionKey: PARTITION_KEY,
    rowKey: id,
    id,
    customer:  patch.customer  ?? existing.customer,
    amount:    patch.amount    ?? existing.amount,
    status:    patch.status    ?? existing.status,
    date:      patch.date      ?? existing.date,
  };

  await table.updateEntity(updated, "Replace");

  return {
    id,
    customer: updated.customer,
    amount:   Number(updated.amount),
    status:   updated.status as OrderStatus,
    date:     updated.date,
  };
}
