import React from "react";
import * as teamsJs from "@microsoft/teams-js";
import {
  FluentProvider,
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsLightTheme,
  Input,
  Spinner,
  MessageBar,
  MessageBarBody,
  Badge,
  Text,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  TableCellLayout,
  useTableFeatures,
  useTableSort,
  type TableColumnDefinition,
  createTableColumn,
  Label,
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogContent,
  DialogActions,
  Button,
  Field,
  type DialogOpenChangeData,
} from "@fluentui/react-components";
import { SearchRegular, EditRegular, OpenRegular } from "@fluentui/react-icons";

type TeamsTheme = "default" | "dark" | "contrast";

type OrderStatus = "Submitted" | "Pending" | "Processing" | "Shipped" | "Delivered" | "Cancelled";

const STATUS_APPEARANCE: Record<OrderStatus, "warning" | "informative" | "success" | "important"> = {
  Submitted:  "informative",
  Pending:    "warning",
  Processing: "informative",
  Shipped:    "success",
  Delivered:  "success",
  Cancelled:  "important",
};

interface Order {
  id: string;
  customer: string;
  amount: number;
  status: OrderStatus;
  date: string;
}

interface CustomerSummary {
  customer: string;
  totalAmount: number;
  orderCount: number;
  latestOrderDate: string;
}

const FLUENT_THEME: Record<TeamsTheme, typeof teamsLightTheme> = {
  default:  teamsLightTheme,
  dark:     teamsDarkTheme,
  contrast: teamsHighContrastTheme,
};

const columns: TableColumnDefinition<CustomerSummary>[] = [
  createTableColumn<CustomerSummary>({ columnId: "customer",        compare: (a, b) => a.customer.localeCompare(b.customer) }),
  createTableColumn<CustomerSummary>({ columnId: "orderCount",      compare: (a, b) => a.orderCount - b.orderCount }),
  createTableColumn<CustomerSummary>({ columnId: "totalAmount",     compare: (a, b) => a.totalAmount - b.totalAmount }),
  createTableColumn<CustomerSummary>({ columnId: "latestOrderDate", compare: (a, b) => a.latestOrderDate.localeCompare(b.latestOrderDate) }),
];

const COLUMN_LABELS: Record<string, string> = {
  customer:        "Customer",
  orderCount:      "Orders",
  totalAmount:     "Total Amount",
  latestOrderDate: "Latest Order",
};

function aggregateOrders(orders: Order[]): CustomerSummary[] {
  const map = new Map<string, CustomerSummary>();
  for (const order of orders) {
    const existing = map.get(order.customer);
    if (existing) {
      existing.totalAmount += order.amount;
      existing.orderCount += 1;
      if (order.date > existing.latestOrderDate) {
        existing.latestOrderDate = order.date;
      }
    } else {
      map.set(order.customer, {
        customer: order.customer,
        totalAmount: order.amount,
        orderCount: 1,
        latestOrderDate: order.date,
      });
    }
  }
  return Array.from(map.values());
}

export default function App() {
  const [theme, setTheme] = React.useState<TeamsTheme>("default");
  const [customers, setCustomers] = React.useState<CustomerSummary[]>([]);
  const [allOrders, setAllOrders] = React.useState<Order[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [fetchError, setFetchError] = React.useState<string | null>(null);
  const [filter, setFilter] = React.useState("");

  // Edit customer dialog
  const [editCustomer, setEditCustomer] = React.useState<CustomerSummary | null>(null);
  const [editName, setEditName] = React.useState("");
  const [saving, setSaving] = React.useState(false);
  const [saveError, setSaveError] = React.useState<string | null>(null);

  // View orders dialog
  const [ordersCustomer, setOrdersCustomer] = React.useState<CustomerSummary | null>(null);

  const openEdit = (c: CustomerSummary) => {
    setEditCustomer(c);
    setEditName(c.customer);
    setSaveError(null);
  };
  const closeEdit = () => { setEditCustomer(null); setSaveError(null); };

  const handleSaveCustomer = async () => {
    if (!editCustomer || !editName.trim()) return;
    setSaving(true);
    setSaveError(null);
    try {
      const res = await fetch(`/api/customers/${encodeURIComponent(editCustomer.customer)}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ newName: editName.trim() }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const newName = editName.trim();
      setAllOrders((prev) => prev.map((o) => o.customer === editCustomer.customer ? { ...o, customer: newName } : o));
      setCustomers((prev) =>
        prev.map((c) =>
          c.customer === editCustomer.customer ? { ...c, customer: newName } : c
        )
      );
      closeEdit();
    } catch (err) {
      setSaveError(String(err));
    } finally {
      setSaving(false);
    }
  };

  React.useEffect(() => {
    teamsJs.app.initialize()
      .then(() => {
        teamsJs.app.getContext().then((ctx) => {
          const t = (ctx.app.theme ?? "default") as TeamsTheme;
          setTheme(t in FLUENT_THEME ? t : "default");
        });
        teamsJs.app.registerOnThemeChangeHandler((t) => {
          setTheme(t in FLUENT_THEME ? (t as TeamsTheme) : "default");
        });
      })
      .catch(() => {/* running outside Teams */});
  }, []);

  React.useEffect(() => {
    setLoading(true);
    setFetchError(null);
    fetch("/api/orders")
      .then((r) => {
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        return r.json() as Promise<Order[]>;
      })
      .then((data) => {
        setAllOrders(data);
        setCustomers(aggregateOrders(data));
      })
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setLoading(false));
  }, []);

  const filtered = React.useMemo(
    () =>
      customers.filter(
        (c) =>
          filter === "" ||
          c.customer.toLowerCase().includes(filter.toLowerCase())
      ),
    [customers, filter]
  );

  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    { columns, items: filtered },
    [useTableSort({ defaultSortState: { sortColumn: "customer", sortDirection: "ascending" } })]
  );

  const rows = sort(getRows());

  const customerOrders = React.useMemo(
    () => ordersCustomer ? allOrders.filter((o) => o.customer === ordersCustomer.customer) : [],
    [allOrders, ordersCustomer]
  );

  return (
    <FluentProvider theme={FLUENT_THEME[theme]} style={{ minHeight: "100vh", padding: "1.5rem" }}>
      <Text as="h1" size={700} weight="semibold" block style={{ marginBottom: "1rem" }}>
        Customers
      </Text>

      {/* ── Edit customer dialog ── */}
      <Dialog
        open={!!editCustomer}
        onOpenChange={(_e: React.SyntheticEvent, data: DialogOpenChangeData) => {
          if (!data.open) closeEdit();
        }}
      >
        <DialogSurface>
          <DialogTitle>Edit Customer</DialogTitle>
          <DialogBody>
            <DialogContent>
              {saveError && (
                <MessageBar intent="error" style={{ marginBottom: "0.75rem" }}>
                  <MessageBarBody>{saveError}</MessageBarBody>
                </MessageBar>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: "0.75rem" }}>
                <Field label="Customer Name" required>
                  <Input
                    value={editName}
                    onChange={(_e, d) => setEditName(d.value)}
                    autoFocus
                  />
                </Field>
                <Text size={200} style={{ color: "var(--colorNeutralForeground3)" }}>
                  Renaming will update all orders for this customer.
                </Text>
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={handleSaveCustomer}
                disabled={saving || !editName.trim() || editName.trim() === editCustomer?.customer}
              >
                {saving ? "Saving…" : "Save"}
              </Button>
              <Button appearance="secondary" onClick={closeEdit} disabled={saving}>
                Cancel
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* ── View orders dialog ── */}
      <Dialog
        open={!!ordersCustomer}
        onOpenChange={(_e: React.SyntheticEvent, data: DialogOpenChangeData) => {
          if (!data.open) setOrdersCustomer(null);
        }}
      >
        <DialogSurface style={{ maxWidth: "760px", width: "90vw" }}>
          <DialogTitle>{ordersCustomer?.customer} — Orders</DialogTitle>
          <DialogBody>
            <DialogContent>
              {customerOrders.length === 0 ? (
                <Text italic>No orders found.</Text>
              ) : (
                <Table aria-label="Customer orders" style={{ width: "100%" }}>
                  <TableHeader>
                    <TableRow>
                      <TableHeaderCell>Order ID</TableHeaderCell>
                      <TableHeaderCell style={{ textAlign: "right" }}>Amount</TableHeaderCell>
                      <TableHeaderCell>Status</TableHeaderCell>
                      <TableHeaderCell>Date</TableHeaderCell>
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {customerOrders
                      .slice()
                      .sort((a, b) => b.date.localeCompare(a.date))
                      .map((o) => (
                        <TableRow key={o.id}>
                          <TableCell>
                            <Text font="monospace">{o.id}</Text>
                          </TableCell>
                          <TableCell style={{ textAlign: "right" }}>
                            {o.amount.toLocaleString("en-US", { style: "currency", currency: "USD" })}
                          </TableCell>
                          <TableCell>
                            <Badge
                              appearance="tint"
                              color={STATUS_APPEARANCE[o.status]}
                              shape="rounded"
                            >
                              {o.status}
                            </Badge>
                          </TableCell>
                          <TableCell>{o.date}</TableCell>
                        </TableRow>
                      ))}
                  </TableBody>
                </Table>
              )}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setOrdersCustomer(null)}>
                Close
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {fetchError && (
        <MessageBar intent="error" style={{ marginBottom: "1rem" }}>
          <MessageBarBody>{fetchError}</MessageBarBody>
        </MessageBar>
      )}

      <div style={{ display: "flex", flexWrap: "wrap", gap: "1rem", alignItems: "flex-end", marginBottom: "1rem" }}>
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <Label htmlFor="customer-search">Customer</Label>
          <Input
            id="customer-search"
            contentBefore={<SearchRegular />}
            placeholder="Search by customer…"
            value={filter}
            onChange={(_e, d) => setFilter(d.value)}
            style={{ width: "240px" }}
          />
        </div>
      </div>

      {loading ? (
        <Spinner label="Loading customers…" />
      ) : (
        <>
          <Table sortable aria-label="Customers table">
            <TableHeader>
              <TableRow>
                {columns.map((col) => (
                  <TableHeaderCell
                    key={col.columnId as string}
                    sortDirection={getSortDirection(col.columnId)}
                    onClick={(e) => toggleColumnSort(e, col.columnId)}
                    style={
                      col.columnId === "totalAmount" || col.columnId === "orderCount"
                        ? { textAlign: "right" }
                        : undefined
                    }
                  >
                    {COLUMN_LABELS[col.columnId as string]}
                  </TableHeaderCell>
                ))}
                <TableHeaderCell style={{ width: "80px" }} />
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.length === 0 ? (
                <TableRow>
                  <TableCell colSpan={4}>
                    <Text italic>No customers found.</Text>
                  </TableCell>
                </TableRow>
              ) : (
                rows.map(({ item: c }) => (
                  <TableRow key={c.customer}>
                    <TableCell>
                      <TableCellLayout>{c.customer}</TableCellLayout>
                    </TableCell>
                    <TableCell style={{ textAlign: "right" }}>
                      {c.orderCount}
                    </TableCell>
                    <TableCell style={{ textAlign: "right" }}>
                      {c.totalAmount.toLocaleString("en-US", { style: "currency", currency: "USD" })}
                    </TableCell>
                    <TableCell>{c.latestOrderDate}</TableCell>
                    <TableCell style={{ width: "80px" }}>
                      <div style={{ display: "flex", gap: "2px" }}>
                        <Button
                          appearance="subtle"
                          icon={<EditRegular />}
                          aria-label={`Edit ${c.customer}`}
                          title="Edit customer"
                          onClick={() => openEdit(c)}
                        />
                        <Button
                          appearance="subtle"
                          icon={<OpenRegular />}
                          aria-label={`View orders for ${c.customer}`}
                          title="View orders"
                          onClick={() => setOrdersCustomer(c)}
                        />
                      </div>
                    </TableCell>
                  </TableRow>
                ))
              )}
            </TableBody>
          </Table>
          <Text size={200} style={{ marginTop: "0.75rem", display: "block" }}>
            {rows.length} of {customers.length} customers
          </Text>
        </>
      )}
    </FluentProvider>
  );
}
