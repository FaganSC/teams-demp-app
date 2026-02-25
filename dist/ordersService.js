'use strict';

var dataTables = require('@azure/data-tables');

const TABLE_NAME = "Orders";
const PARTITION_KEY = "Orders";
const CUSTOMERS = [
  "Contoso Ltd.",
  "Fabrikam Inc.",
  "Northwind Traders",
  "Adventure Works",
  "Tailspin Toys",
  "Woodgrove Bank",
  "Proseware Inc.",
  "Lucerne Publishing",
  "Humongous Insurance",
  "Wide World Importers",
  "Fourth Coffee",
  "Alpine Ski House",
  "Coho Winery",
  "Relecloud",
  "Trey Research"
];
const STATUSES = ["Pending", "Processing", "Shipped", "Delivered", "Cancelled"];
function randomDate(daysBack = 365) {
  const now = /* @__PURE__ */ new Date("2026-02-24");
  const offset = Math.floor(Math.random() * daysBack * 24 * 60 * 60 * 1e3);
  return new Date(now.getTime() - offset).toISOString().slice(0, 10);
}
function generateOrders(count) {
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
      date: randomDate()
    });
  }
  return orders;
}
function getClients(connectionString) {
  return {
    service: dataTables.TableServiceClient.fromConnectionString(connectionString),
    table: dataTables.TableClient.fromConnectionString(connectionString, TABLE_NAME)
  };
}
async function seedIfEmpty(connectionString) {
  const { service, table } = getClients(connectionString);
  await service.createTable(TABLE_NAME).catch(() => {
  });
  const iter = table.listEntities({ queryOptions: { select: ["rowKey"] } });
  const first = await iter.next();
  if (!first.done) {
    return;
  }
  const transaction = new dataTables.TableTransaction();
  for (const order of generateOrders(100)) {
    transaction.createEntity(order);
  }
  await table.submitTransaction(transaction.actions);
}
async function listOrders(connectionString) {
  const { table } = getClients(connectionString);
  const orders = [];
  for await (const entity of table.listEntities()) {
    orders.push({
      id: entity.rowKey,
      customer: entity.customer,
      amount: entity.amount,
      status: entity.status,
      date: entity.date
    });
  }
  return orders.sort((a, b) => a.id.localeCompare(b.id));
}
async function renameCustomer(connectionString, oldName, newName) {
  const { table } = getClients(connectionString);
  const toUpdate = [];
  for await (const entity of table.listEntities()) {
    if (entity.customer === oldName) {
      toUpdate.push({
        id: entity.rowKey,
        customer: entity.customer,
        amount: entity.amount,
        status: entity.status,
        date: entity.date
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
      date: order.date
    }, "Replace");
  }
  return toUpdate.length;
}
async function updateOrder(connectionString, id, patch) {
  const { table } = getClients(connectionString);
  const existing = await table.getEntity(PARTITION_KEY, id);
  const updated = {
    partitionKey: PARTITION_KEY,
    rowKey: id,
    id,
    customer: patch.customer ?? existing.customer,
    amount: patch.amount ?? existing.amount,
    status: patch.status ?? existing.status,
    date: patch.date ?? existing.date
  };
  await table.updateEntity(updated, "Replace");
  return {
    id,
    customer: updated.customer,
    amount: updated.amount,
    status: updated.status,
    date: updated.date
  };
}

exports.listOrders = listOrders;
exports.renameCustomer = renameCustomer;
exports.seedIfEmpty = seedIfEmpty;
exports.updateOrder = updateOrder;
//# sourceMappingURL=ordersService.js.map
//# sourceMappingURL=ordersService.js.map