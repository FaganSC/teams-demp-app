type OrderStatus = "Pending" | "Processing" | "Shipped" | "Delivered" | "Cancelled";
interface Order {
    id: string;
    customer: string;
    amount: number;
    status: OrderStatus;
    date: string;
}
/** Creates the Orders table and seeds 100 rows if it is empty. */
declare function seedIfEmpty(connectionString: string): Promise<void>;
/** Returns all orders sorted by ID. */
declare function listOrders(connectionString: string): Promise<Order[]>;
/** Renames a customer on all their orders. Returns the number of orders updated. */
declare function renameCustomer(connectionString: string, oldName: string, newName: string): Promise<number>;
/** Updates an existing order in Table Storage. */
declare function updateOrder(connectionString: string, id: string, patch: Partial<Omit<Order, "id">>): Promise<Order>;

export { type Order, type OrderStatus, listOrders, renameCustomer, seedIfEmpty, updateOrder };
