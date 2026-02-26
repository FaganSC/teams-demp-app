import { TableClient, TableServiceClient } from "@azure/data-tables";
import {
  AdaptiveCard,
  TextBlock,
  FactSet,
  Fact,
  ExecuteAction,
  ActionSet,
  type TextColor,
} from "@microsoft/teams.cards";

import type { Order, OrderStatus } from "./ordersService.js";

// ── Table storage for bot conversation IDs ─────────────────────────────────

const SUB_TABLE = "BotSubscriptions";
const SUB_PARTITION = "BotSub";

function subTable(connectionString: string): TableClient {
  return TableClient.fromConnectionString(connectionString, SUB_TABLE);
}

export async function ensureSubTable(connectionString: string): Promise<void> {
  const svc = TableServiceClient.fromConnectionString(connectionString);
  await svc.createTable(SUB_TABLE).catch(() => { /* already exists */ });
}

/** Persist a conversation ID so we can send proactive messages to it. */
export async function saveConversation(
  connectionString: string,
  conversationId: string,
  serviceUrl: string,
): Promise<void> {
  const client = subTable(connectionString);
  const rowKey = Buffer.from(conversationId).toString("base64").replace(/[/+=]/g, "_");
  await client.upsertEntity({
    partitionKey: SUB_PARTITION,
    rowKey,
    conversationId,
    serviceUrl,
  }, "Replace");
}

/** Remove a stored conversation (bot uninstalled). */
export async function removeConversation(
  connectionString: string,
  conversationId: string,
): Promise<void> {
  const client = subTable(connectionString);
  const rowKey = Buffer.from(conversationId).toString("base64").replace(/[/+=]/g, "_");
  await client.deleteEntity(SUB_PARTITION, rowKey).catch(() => { /* already gone */ });
}

/** Return all saved conversation IDs. */
export async function listConversations(
  connectionString: string,
): Promise<string[]> {
  const client = subTable(connectionString);
  const ids: string[] = [];
  for await (const entity of client.listEntities<{ conversationId: string }>()) {
    if (entity.conversationId) ids.push(entity.conversationId);
  }
  return ids;
}

// ── Adaptive card builders ──────────────────────────────────────────────────

const STATUS_COLORS: Record<OrderStatus, TextColor> = {
  Submitted:  "Accent",
  Pending:    "Warning",
  Processing: "Accent",
  Shipped:    "Good",
  Delivered:  "Good",
  Cancelled:  "Attention",
};

/** Build the initial "new order received" card with Accept / Cancel actions. */
export function buildNewOrderCard(order: Order): AdaptiveCard {
  const header = new TextBlock("📦 New Order Received", {
    size: "Large",
    weight: "Bolder",
    color: "Accent",
  });

  const facts = new FactSet();
  facts.facts = [
    new Fact("Order ID",  order.id),
    new Fact("Customer",  order.customer),
    new Fact("Amount",   `$${Number(order.amount).toFixed(2)}`),
    new Fact("Date",     order.date),
    new Fact("Status",   order.status),
  ];

  const accept = new ExecuteAction();
  accept.title = "✅ Accept";
  accept.verb  = "order.accept";
  accept.data  = { orderId: order.id };
  accept.style = "positive";

  const cancel = new ExecuteAction();
  cancel.title = "❌ Cancel";
  cancel.verb  = "order.cancel";
  cancel.data  = { orderId: order.id };
  cancel.style = "destructive";

  const actions = new ActionSet();
  actions.actions = [accept, cancel];

  const card = new AdaptiveCard();
  card.$schema = "http://adaptivecards.io/schemas/adaptive-card.json";
  card.version = "1.5";
  card.body = [header, facts, actions];
  return card;
}

/** Build a confirmation card shown after the user acts on an order. */
export function buildConfirmedCard(order: Order, actedBy: string): AdaptiveCard {
  const label =
    order.status === "Pending"   ? "✅ Order Accepted"  :
    order.status === "Cancelled" ? "❌ Order Cancelled" :
    `📋 Order ${order.status}`;

  const header = new TextBlock(label, {
    size:   "Large",
    weight: "Bolder",
    color:  STATUS_COLORS[order.status] ?? "Default",
  });

  const facts = new FactSet();
  facts.facts = [
    new Fact("Order ID",   order.id),
    new Fact("Customer",   order.customer),
    new Fact("Amount",     `$${Number(order.amount).toFixed(2)}`),
    new Fact("Date",       order.date),
    new Fact("New Status", order.status),
    new Fact("Updated by", actedBy),
  ];

  const note = new TextBlock("No further actions available.", {
    isSubtle: true,
    size: "Small",
  });

  const card = new AdaptiveCard();
  card.$schema = "http://adaptivecards.io/schemas/adaptive-card.json";
  card.version = "1.5";
  card.body = [header, facts, note];
  return card;
}