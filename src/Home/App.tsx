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
  Dropdown,
  Option,
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
import { SearchRegular, EditRegular, AddRegular } from "@fluentui/react-icons";

type TeamsTheme = "default" | "dark" | "contrast";

type OrderStatus = "Submitted" | "Pending" | "Processing" | "Shipped" | "Delivered" | "Cancelled";

interface Order {
  id: string;
  customer: string;
  amount: number;
  status: OrderStatus;
  date: string;
}

const ALL_STATUSES: OrderStatus[] = ["Submitted", "Pending", "Processing", "Shipped", "Delivered", "Cancelled"];

const STATUS_APPEARANCE: Record<OrderStatus, "warning" | "informative" | "success" | "important"> = {
  Submitted:  "informative",
  Pending:    "warning",
  Processing: "informative",
  Shipped:    "success",
  Delivered:  "success",
  Cancelled:  "important",
};

const FLUENT_THEME: Record<TeamsTheme, typeof teamsLightTheme> = {
  default:  teamsLightTheme,
  dark:     teamsDarkTheme,
  contrast: teamsHighContrastTheme,
};

const columns: TableColumnDefinition<Order>[] = [
  createTableColumn<Order>({ columnId: "id",       compare: (a, b) => a.id.localeCompare(b.id) }),
  createTableColumn<Order>({ columnId: "customer",  compare: (a, b) => a.customer.localeCompare(b.customer) }),
  createTableColumn<Order>({ columnId: "amount",    compare: (a, b) => a.amount - b.amount }),
  createTableColumn<Order>({ columnId: "status",    compare: (a, b) => a.status.localeCompare(b.status) }),
  createTableColumn<Order>({ columnId: "date",      compare: (a, b) => a.date.localeCompare(b.date) }),
];

const COLUMN_LABELS: Record<string, string> = {
  id: "ID", customer: "Customer", amount: "Amount", status: "Status", date: "Date",
};

export default function App() {
  const [theme, setTheme] = React.useState<TeamsTheme>("default");
  const [orders, setOrders] = React.useState<Order[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [fetchError, setFetchError] = React.useState<string | null>(null);
  const [filter, setFilter] = React.useState("");
  const [customerFilter, setCustomerFilter] = React.useState("");
  const [statusFilter, setStatusFilter] = React.useState<OrderStatus[]>([]);

  // Edit dialog state
  const [editOrder, setEditOrder] = React.useState<Order | null>(null);
  const [editDraft, setEditDraft] = React.useState<Order | null>(null);
  const [saving, setSaving] = React.useState(false);
  const [saveError, setSaveError] = React.useState<string | null>(null);

  // New order dialog state
  const emptyDraft = (): { customer: string; amount: number } => ({ customer: "", amount: 0 });
  const [newOrderOpen, setNewOrderOpen] = React.useState(false);
  const [newDraft, setNewDraft] = React.useState<{ customer: string; amount: number }>(emptyDraft());
  const [creating, setCreating] = React.useState(false);
  const [createError, setCreateError] = React.useState<string | null>(null);

  const openEdit = (order: Order) => {
    setEditOrder(order);
    setEditDraft({ ...order });
    setSaveError(null);
  };

  const closeEdit = () => {
    setEditOrder(null);
    setEditDraft(null);
    setSaveError(null);
  };

  const handleCreate = async () => {
    if (!newDraft.customer.trim()) return;
    setCreating(true);
    setCreateError(null);
    try {
      const res = await fetch("/api/orders", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(newDraft),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const created: Order = await res.json();
      // Also add locally immediately (SSE may deduplicate)
      setOrders((prev) => prev.some((o) => o.id === created.id) ? prev : [...prev, created]);
      setNewOrderOpen(false);
      setNewDraft(emptyDraft());
    } catch (err) {
      setCreateError(String(err));
    } finally {
      setCreating(false);
    }
  };

  const handleSave = async () => {
    if (!editDraft) return;
    setSaving(true);
    setSaveError(null);
    try {
      const res = await fetch(`/api/orders/${editDraft.id}`, {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          customer: editDraft.customer,
          amount: editDraft.amount,
          status: editDraft.status,
          date: editDraft.date,
        }),
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const updated: Order = await res.json();
      setOrders((prev) => prev.map((o) => (o.id === updated.id ? updated : o)));
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
      .then((data) => setOrders(data))
      .catch((err: unknown) => setFetchError(String(err)))
      .finally(() => setLoading(false));
  }, []);

  // SSE – auto-append any new order pushed by the server
  React.useEffect(() => {
    const es = new EventSource("/api/orders/events");
    es.onmessage = (e) => {
      try {
        const order: Order = JSON.parse(e.data as string);
        setOrders((prev) => prev.some((o) => o.id === order.id) ? prev : [...prev, order]);
      } catch { /* ignore */ }
    };
    return () => es.close();
  }, []);

  const filtered = React.useMemo(
    () => orders.filter((o) => {
      const matchesSearch =
        filter === "" ||
        o.id.toLowerCase().includes(filter.toLowerCase());
      const matchesCustomer =
        customerFilter === "" ||
        o.customer.toLowerCase().includes(customerFilter.toLowerCase());
      const matchesStatus =
        statusFilter.length === 0 || statusFilter.includes(o.status);
      return matchesSearch && matchesCustomer && matchesStatus;
    }),
    [orders, filter, customerFilter, statusFilter]
  );

  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    { columns, items: filtered },
    [useTableSort({ defaultSortState: { sortColumn: "date", sortDirection: "descending" } })]
  );

  const rows = sort(getRows());

  return (
    <FluentProvider theme={FLUENT_THEME[theme]} style={{ minHeight: "100vh", padding: "1.5rem" }}>
      <div style={{ display: "flex", alignItems: "center", gap: "1rem", marginBottom: "1rem" }}>
        <Text as="h1" size={700} weight="semibold" block style={{ margin: 0 }}>
          Orders
        </Text>
        <Button
          appearance="primary"
          icon={<AddRegular />}
          onClick={() => { setNewDraft(emptyDraft()); setCreateError(null); setNewOrderOpen(true); }}
        >
          New Order
        </Button>
      </div>

      {/* New Order dialog */}
      <Dialog
        open={newOrderOpen}
        onOpenChange={(_e: React.SyntheticEvent, data: DialogOpenChangeData) => {
          if (!data.open) { setNewOrderOpen(false); setCreateError(null); }
        }}
      >
        <DialogSurface>
          <DialogTitle>New Order</DialogTitle>
          <DialogBody>
            <DialogContent>
              {createError && (
                <MessageBar intent="error" style={{ marginBottom: "0.75rem" }}>
                  <MessageBarBody>{createError}</MessageBarBody>
                </MessageBar>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: "0.75rem" }}>
                <Field label="Customer" required>
                  <Input
                    value={newDraft.customer}
                    onChange={(_e, d) => setNewDraft((p) => ({ ...p, customer: d.value }))}
                    autoFocus
                  />
                </Field>
                <Field label="Amount" required>
                  <Input
                    type="number"
                    value={String(newDraft.amount)}
                    onChange={(_e, d) =>
                      setNewDraft((p) => ({ ...p, amount: parseFloat(d.value) || 0 }))
                    }
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="primary"
                onClick={handleCreate}
                disabled={creating || !newDraft.customer.trim()}
              >
                {creating ? "Creating…" : "Create"}
              </Button>
              <Button appearance="secondary" onClick={() => setNewOrderOpen(false)} disabled={creating}>
                Cancel
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Edit dialog */}
      <Dialog
        open={!!editOrder}
        onOpenChange={(_e: React.SyntheticEvent, data: DialogOpenChangeData) => {
          if (!data.open) closeEdit();
        }}
      >
        <DialogSurface>
          <DialogTitle>Edit Order {editOrder?.id}</DialogTitle>
          <DialogBody>
            <DialogContent>
              {saveError && (
                <MessageBar intent="error" style={{ marginBottom: "0.75rem" }}>
                  <MessageBarBody>{saveError}</MessageBarBody>
                </MessageBar>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: "0.75rem" }}>
                <Field label="Customer" required>
                  <Input
                    value={editDraft?.customer ?? ""}
                    onChange={(_e, d) =>
                      setEditDraft((prev) => prev ? { ...prev, customer: d.value } : prev)
                    }
                  />
                </Field>
                <Field label="Amount" required>
                  <Input
                    type="number"
                    value={String(editDraft?.amount ?? "")}
                    onChange={(_e, d) =>
                      setEditDraft((prev) =>
                        prev ? { ...prev, amount: parseFloat(d.value) || 0 } : prev
                      )
                    }
                  />
                </Field>
                <Field label="Status" required>
                  <Dropdown
                    value={editDraft?.status ?? ""}
                    selectedOptions={editDraft ? [editDraft.status] : []}
                    onOptionSelect={(_e, data) =>
                      setEditDraft((prev) =>
                        prev ? { ...prev, status: data.optionValue as OrderStatus } : prev
                      )
                    }
                  >
                    {ALL_STATUSES.map((s) => (
                      <Option key={s} value={s}>{s}</Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="Date" required>
                  <Input
                    type="date"
                    value={editDraft?.date ?? ""}
                    onChange={(_e, d) =>
                      setEditDraft((prev) => prev ? { ...prev, date: d.value } : prev)
                    }
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="primary" onClick={handleSave} disabled={saving}>
                {saving ? "Saving…" : "Save"}
              </Button>
              <Button appearance="secondary" onClick={closeEdit} disabled={saving}>
                Cancel
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
          <Label htmlFor="search-input">Order ID</Label>
          <Input
            id="search-input"
            contentBefore={<SearchRegular />}
            placeholder="Search by ID…"
            value={filter}
            onChange={(_e, d) => setFilter(d.value)}
            style={{ width: "220px" }}
          />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <Label htmlFor="customer-filter">Customer</Label>
          <Input
            id="customer-filter"
            placeholder="Filter by customer…"
            value={customerFilter}
            onChange={(_e, d) => setCustomerFilter(d.value)}
            style={{ width: "200px" }}
          />
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          <Label>Status</Label>
          <Dropdown
            multiselect
            placeholder="All statuses"
            selectedOptions={statusFilter}
            onOptionSelect={(_e, data) =>
              setStatusFilter(data.selectedOptions as OrderStatus[])
            }
            style={{ minWidth: "160px" }}
          >
            {ALL_STATUSES.map((s) => (
              <Option key={s} value={s} checkIcon={<Badge appearance="tint" color={STATUS_APPEARANCE[s]} shape="rounded" style={{ fontSize: "0.7rem" }}>{s}</Badge>}>
                {s}
              </Option>
            ))}
          </Dropdown>
        </div>
      </div>

      {loading ? (
        <Spinner label="Loading orders…" />
      ) : (
        <>
          <Table sortable aria-label="Orders table">
            <TableHeader>
              <TableRow>
                  {columns.map((col) => (
                  <TableHeaderCell
                    key={col.columnId as string}
                    sortDirection={getSortDirection(col.columnId)}
                    onClick={(e) => toggleColumnSort(e, col.columnId)}
                    style={col.columnId === "amount" ? { textAlign: "right" } : undefined}
                  >
                    {COLUMN_LABELS[col.columnId as string]}
                  </TableHeaderCell>
                ))}
                <TableHeaderCell style={{ width: "48px" }} />
              </TableRow>
            </TableHeader>
            <TableBody>
              {rows.length === 0 ? (
                <TableRow>
                  <TableCell colSpan={5}>
                    <Text italic>No orders found.</Text>
                  </TableCell>
                </TableRow>
              ) : (
                rows.map(({ item: order }) => (
                  <TableRow key={order.id}>
                    <TableCell>
                      <TableCellLayout>
                        <Text font="monospace">{order.id}</Text>
                      </TableCellLayout>
                    </TableCell>
                    <TableCell>{order.customer}</TableCell>
                    <TableCell style={{ textAlign: "right" }}>
                      {order.amount.toLocaleString("en-US", { style: "currency", currency: "USD" })}
                    </TableCell>
                    <TableCell>
                      <Badge
                        appearance="tint"
                        color={STATUS_APPEARANCE[order.status]}
                        shape="rounded"
                      >
                        {order.status}
                      </Badge>
                    </TableCell>
                    <TableCell>{order.date}</TableCell>
                    <TableCell style={{ width: "48px" }}>
                      <Button
                        appearance="subtle"
                        icon={<EditRegular />}
                        aria-label={`Edit ${order.id}`}
                        onClick={() => openEdit(order)}
                      />
                    </TableCell>
                  </TableRow>
                ))
              )}
            </TableBody>
          </Table>
          <Text size={200} style={{ marginTop: "0.75rem", display: "block" }}>
            {rows.length} of {orders.length} orders
          </Text>
        </>
      )}
    </FluentProvider>
  );
}
