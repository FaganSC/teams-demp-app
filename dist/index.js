'use strict';

var fs = require('fs');
var https = require('https');
var path = require('path');
var teams_apps = require('@microsoft/teams.apps');
var logging = require('@microsoft/teams.common/logging');
var teams_dev = require('@microsoft/teams.dev');
var ordersService_js = require('./ordersService.js');

function _interopDefault (e) { return e && e.__esModule ? e : { default: e }; }

var fs__default = /*#__PURE__*/_interopDefault(fs);
var https__default = /*#__PURE__*/_interopDefault(https);
var path__default = /*#__PURE__*/_interopDefault(path);

const STORAGE_CONNECTION = process.env.AZURE_STORAGE_CONNECTION_STRING ?? "UseDevelopmentStorage=true";
const sslOptions = {
  key: process.env.SSL_KEY_FILE ? fs__default.default.readFileSync(process.env.SSL_KEY_FILE) : void 0,
  cert: process.env.SSL_CRT_FILE ? fs__default.default.readFileSync(process.env.SSL_CRT_FILE) : void 0
};
const plugins = [new teams_dev.DevtoolsPlugin()];
if (sslOptions.cert && sslOptions.key) {
  plugins.push(new teams_apps.HttpPlugin(https__default.default.createServer(sslOptions)));
}
const app = new teams_apps.App({
  logger: new logging.ConsoleLogger("tab", { level: "debug" }),
  plugins
});
app.tab("home", path__default.default.join(__dirname, "./Home"));
app.tab("customers", path__default.default.join(__dirname, "./Customers"));
app.http.use(require("express").json());
app.http.get("/api/orders", async (_req, res) => {
  try {
    const orders = await ordersService_js.listOrders(STORAGE_CONNECTION);
    res.json(orders);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
app.http.get("/api/customers/:name/orders", async (req, res) => {
  try {
    const { name } = req.params;
    const orders = await ordersService_js.listOrders(STORAGE_CONNECTION);
    res.json(orders.filter((o) => o.customer === decodeURIComponent(name)));
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
app.http.put("/api/customers/:name", async (req, res) => {
  try {
    const { name } = req.params;
    const { newName } = req.body;
    const updated = await ordersService_js.renameCustomer(STORAGE_CONNECTION, decodeURIComponent(name), newName);
    res.json({ updated });
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
app.http.put("/api/orders/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const patch = req.body;
    const updated = await ordersService_js.updateOrder(STORAGE_CONNECTION, id, patch);
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});
(async () => {
  await ordersService_js.seedIfEmpty(STORAGE_CONNECTION);
  await app.start(+(process.env.PORT || 3978));
})();
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map