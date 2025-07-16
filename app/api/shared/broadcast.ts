// Shared broadcast utility for real-time updates

// Store connected clients for finished products stream
const finishedProductsClients = new Set<ReadableStreamDefaultController>();

// Store connected clients for live products stream
const liveProductsClients = new Set<ReadableStreamDefaultController>();

// Store connected clients for RFD status stream
const rfdStatusClients = new Set<ReadableStreamDefaultController>();

// Function to broadcast when a product is added to live products
export function broadcastProductAddedToLive(productId: string, productName: string, machineName: string) {
  const data = JSON.stringify({
    type: 'product_added_to_live',
    data: {
      productId,
      productName,
      machineName,
      timestamp: new Date().toISOString()
    }
  });

  finishedProductsClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to finished products client:', error);
      finishedProductsClients.delete(client);
    }
  });
}

// Function to broadcast when a product is moved to past
export function broadcastProductMovedToPast(productId: string, productName: string) {
  const data = JSON.stringify({
    type: 'product_moved_to_past',
    data: {
      productId,
      productName,
      timestamp: new Date().toISOString()
    }
  });

  // Broadcast to both streams
  finishedProductsClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to finished products client:', error);
      finishedProductsClients.delete(client);
    }
  });

  liveProductsClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to live products client:', error);
      liveProductsClients.delete(client);
    }
  });
}

// Function to broadcast when finished products list is updated
export function broadcastFinishedProductsUpdate() {
  const data = JSON.stringify({
    type: 'finished_products_update',
    data: {
      timestamp: new Date().toISOString()
    }
  });

  finishedProductsClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to finished products client:', error);
      finishedProductsClients.delete(client);
    }
  });
}

// Function to broadcast dispatchable products update
export function broadcastDispatchableProducts(products: any[]) {
  const data = JSON.stringify({
    type: 'dispatchable_products_update',
    data: products,
    timestamp: new Date().toISOString()
  });

  liveProductsClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to live products client:', error);
      liveProductsClients.delete(client);
    }
  });
}

// Function to broadcast RFD status updates
export function broadcastRFDStatusUpdate(updateData: any) {
  const data = JSON.stringify({
    type: 'rfd_status_update',
    data: updateData,
    timestamp: new Date().toISOString()
  });

  rfdStatusClients.forEach(client => {
    try {
      client.enqueue(new TextEncoder().encode(`data: ${data}\n\n`));
    } catch (error) {
      console.error('Error broadcasting to RFD status client:', error);
      rfdStatusClients.delete(client);
    }
  });
}

// Export client sets for use in stream files
export { finishedProductsClients, liveProductsClients, rfdStatusClients }; 