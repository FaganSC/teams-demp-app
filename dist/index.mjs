import fs from 'fs';
import https from 'https';
import path from 'path';
import { HttpPlugin, App } from '@microsoft/teams.apps';
import { ConsoleLogger } from '@microsoft/teams.common/logging';
import { DevtoolsPlugin } from '@microsoft/teams.dev';
import { listOrders, renameCustomer, updateOrder, seedIfEmpty } from './ordersService.js';

const STORAGE_CONNECTION = process.env.AZURE_STORAGE_CONNECTION_STRING ?? "UseDevelopmentStorage=true";
const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs.readFileSync(process.env.SSL_KEY_FILE) : void 0,
  cert: process.env.SSL_CRT_FILE ? fs.readFileSync(process.env.SSL_CRT_FILE) : void 0
};
const plugins = [new DevtoolsPlugin()];
if (sslOptions.cert && sslOptions.key) {
  plugins.push(new HttpPlugin(https.createServer(sslOptions)));
}
const app = new App({
  logger: new ConsoleLogger("tab", { level: "debug" }),
  plugins
});
app.tab("home", path.join(__dirname, "./Home"));
app.tab("customers", path.join(__dirname, "./Customers"));
app.http.use(require("express").json());
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
    const { name } = req.params;
    const orders = await listOrders(STORAGE_CONNECTION);
    res.json(orders.filter((o) => o.customer === decodeURIComponent(name)));
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
app.http.put("/api/customers/:name", async (req, res) => {
  try {
    const { name } = req.params;
    const { newName } = req.body;
    const updated = await renameCustomer(STORAGE_CONNECTION, decodeURIComponent(name), newName);
    res.json({ updated });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
app.http.put("/api/orders/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const patch = req.body;
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
//# sourceMappingURL=index.mjs.map
//# sourceMappingURL=index.mjs.map