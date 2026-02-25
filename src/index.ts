import fs from "fs";
import https from "https";
import path from "path";

import { App, HttpPlugin, IPlugin } from "@microsoft/teams.apps";
import { ConsoleLogger } from "@microsoft/teams.common/logging";
import { DevtoolsPlugin } from "@microsoft/teams.dev";

import { createOrder, listOrders, renameCustomer, seedIfEmpty, updateOrder, type Order, type OrderStatus } from "./ordersService.js";

const STORAGE_CONNECTION =
  process.env.AZURE_STORAGE_CONNECTION_STRING ?? "UseDevelopmentStorage=true";

const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : undefined,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : undefined,
};
const plugins: IPlugin[] = [new DevtoolsPlugin()];
if (sslOptions.cert && sslOptions.key) {
  plugins.push(new HttpPlugin(https.createServer(sslOptions)));
}
const app = new App({
  logger: new ConsoleLogger("tab", { level: "debug" }),
  plugins: plugins,
});

app.tab("home", path.join(__dirname, "./Home"));
app.tab("customers", path.join(__dirname, "./Customers"));

// REST API – orders
app.http.use(require("express").json());

// SSE – push new orders to all connected clients
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const sseClients = new Set<any>();

function broadcastNewOrder(order: Order): void {
  const payload = `data: ${JSON.stringify(order)}\n\n`;
  for (const client of sseClients) {
    if (!client.writableEnded) client.write(payload);
  }
}

app.http.get("/api/orders/events", (req: any, res: any) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();
  sseClients.add(res);
  req.on("close", () => sseClients.delete(res));
});

app.http.post("/api/orders", async (req: any, res: any) => {
  try {
    const { customer, amount } = req.body as { customer: string; amount: number };
    const today = new Date().toISOString().slice(0, 10);
    const created = await createOrder(STORAGE_CONNECTION, {
      customer,
      amount,
      status: "Submitted",
      date: today,
    });
    broadcastNewOrder(created);
    res.status(201).json(created);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

app.http.get("/api/orders", async (_req, res) => {
  try {
    const orders = await listOrders(STORAGE_CONNECTION);
    res.json(orders);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

app.http.get("/api/customers/:name/orders", async (req, res) => {
  try {
    const { name } = req.params as { name: string };
    const orders = await listOrders(STORAGE_CONNECTION);
    res.json(orders.filter((o) => o.customer === decodeURIComponent(name)));
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

app.http.put("/api/customers/:name", async (req, res) => {
  try {
    const { name } = req.params as { name: string };
    const { newName } = req.body as { newName: string };
    const updated = await renameCustomer(STORAGE_CONNECTION, decodeURIComponent(name), newName);
    res.json({ updated });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

app.http.put("/api/orders/:id", async (req, res) => {
  try {
    const { id } = req.params as { id: string };
    const patch = req.body as { customer?: string; amount?: number; status?: OrderStatus; date?: string };
    const updated = await updateOrder(STORAGE_CONNECTION, id, patch);
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

(async () => {
  await seedIfEmpty(STORAGE_CONNECTION);
  await app.start(+(process.env.PORT || 3978));
})();
