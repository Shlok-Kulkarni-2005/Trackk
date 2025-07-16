import { NextRequest } from 'next/server';
import { SSEEvent } from '../types';

export interface SSEClient {
  id: string;
  controller: ReadableStreamDefaultController;
  lastPing: number;
  isAlive: boolean;
}

export class SSEManager {
  private clients: Map<string, SSEClient> = new Map();
  private keepAliveInterval: NodeJS.Timeout | null = null;
  private cleanupInterval: NodeJS.Timeout | null = null;
  private readonly KEEP_ALIVE_INTERVAL = 30000; // 30 seconds
  private readonly CLEANUP_INTERVAL = 60000; // 1 minute
  private readonly CLIENT_TIMEOUT = 120000; // 2 minutes

  constructor() {
    this.startKeepAlive();
    this.startCleanup();
  }

  private startKeepAlive(): void {
    this.keepAliveInterval = setInterval(() => {
      this.broadcastToAll({
        type: 'ping',
        timestamp: new Date().toISOString()
      });
    }, this.KEEP_ALIVE_INTERVAL);
  }

  private startCleanup(): void {
    this.cleanupInterval = setInterval(() => {
      this.cleanupDeadConnections();
    }, this.CLEANUP_INTERVAL);
  }

  private cleanupDeadConnections(): void {
    const now = Date.now();
    const deadClients: string[] = [];

    this.clients.forEach((client, id) => {
      if (now - client.lastPing > this.CLIENT_TIMEOUT) {
        deadClients.push(id);
      }
    });

    deadClients.forEach(id => {
      this.removeClient(id);
    });

    if (deadClients.length > 0) {
      console.log(`🧹 Cleaned up ${deadClients.length} dead SSE connections`);
    }
  }

  public addClient(id: string, controller: ReadableStreamDefaultController): void {
    this.clients.set(id, {
      id,
      controller,
      lastPing: Date.now(),
      isAlive: true
    });

    console.log(`🔗 SSE client connected: ${id} (total: ${this.clients.size})`);
  }

  public removeClient(id: string): void {
    const client = this.clients.get(id);
    if (client) {
      try {
        client.controller.close();
      } catch (error) {
        console.error(`Error closing SSE client ${id}:`, error);
      }
      this.clients.delete(id);
      console.log(`🔌 SSE client disconnected: ${id} (total: ${this.clients.size})`);
    }
  }

  public updateClientPing(id: string): void {
    const client = this.clients.get(id);
    if (client) {
      client.lastPing = Date.now();
    }
  }

  public broadcastToAll(event: SSEEvent): void {
    const deadClients: string[] = [];

    this.clients.forEach((client, id) => {
      try {
        const message = `data: ${JSON.stringify(event)}\n\n`;
        client.controller.enqueue(new TextEncoder().encode(message));
      } catch (error) {
        console.error(`Error sending SSE message to client ${id}:`, error);
        deadClients.push(id);
      }
    });

    // Remove dead clients
    deadClients.forEach(id => this.removeClient(id));
  }

  public broadcastToClient(clientId: string, event: SSEEvent): boolean {
    const client = this.clients.get(clientId);
    if (!client) {
      return false;
    }

    try {
      const message = `data: ${JSON.stringify(event)}\n\n`;
      client.controller.enqueue(new TextEncoder().encode(message));
      return true;
    } catch (error) {
      console.error(`Error sending SSE message to client ${clientId}:`, error);
      this.removeClient(clientId);
      return false;
    }
  }

  public getClientCount(): number {
    return this.clients.size;
  }

  public getClientIds(): string[] {
    return Array.from(this.clients.keys());
  }

  public destroy(): void {
    if (this.keepAliveInterval) {
      clearInterval(this.keepAliveInterval);
      this.keepAliveInterval = null;
    }

    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = null;
    }

    // Close all client connections
    this.clients.forEach((client, id) => {
      try {
        client.controller.close();
      } catch (error) {
        console.error(`Error closing SSE client ${id} during destroy:`, error);
      }
    });

    this.clients.clear();
    console.log('🧹 SSE Manager destroyed');
  }
}

// Global SSE manager instance
export const sseManager = new SSEManager();

// Cleanup on process exit
process.on('SIGINT', () => {
  console.log('🛑 Shutting down SSE Manager...');
  sseManager.destroy();
  process.exit(0);
});

process.on('SIGTERM', () => {
  console.log('🛑 Shutting down SSE Manager...');
  sseManager.destroy();
  process.exit(0);
});

export function createSSEStream(clientId: string): ReadableStream {
  const encoder = new TextEncoder();

  return new ReadableStream({
    start(controller) {
      // Add client to manager
      sseManager.addClient(clientId, controller);

      // Send initial connection message
      const connectionMessage = `data: ${JSON.stringify({
        type: 'connection_established',
        clientId,
        timestamp: new Date().toISOString()
      })}\n\n`;
      
      controller.enqueue(encoder.encode(connectionMessage));

      // Send current client count
      const clientCountMessage = `data: ${JSON.stringify({
        type: 'client_count',
        count: sseManager.getClientCount(),
        timestamp: new Date().toISOString()
      })}\n\n`;
      
      controller.enqueue(encoder.encode(clientCountMessage));
    },

    cancel() {
      // Remove client when connection is cancelled
      sseManager.removeClient(clientId);
    }
  });
}

export function handleSSERequest(req: NextRequest): Response {
  const clientId = generateClientId();
  
  return new Response(createSSEStream(clientId), {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Cache-Control'
    }
  });
}

function generateClientId(): string {
  return `client_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

export function broadcastEvent(event: SSEEvent): void {
  sseManager.broadcastToAll(event);
}

export function broadcastToClient(clientId: string, event: SSEEvent): boolean {
  return sseManager.broadcastToClient(clientId, event);
}

// Utility functions for common events
export function broadcastProductUpdate(product: any, action: string): void {
  broadcastEvent({
    type: 'product_updated',
    data: { product, action },
    timestamp: new Date().toISOString()
  });
}

export function broadcastAlertUpdate(alert: any, action: string): void {
  broadcastEvent({
    type: 'alert_updated',
    data: { alert, action },
    timestamp: new Date().toISOString()
  });
}

export function broadcastJobUpdate(job: any, action: string): void {
  broadcastEvent({
    type: 'job_updated',
    data: { job, action },
    timestamp: new Date().toISOString()
  });
}

export function broadcastSystemMessage(message: string, level: 'info' | 'warning' | 'error' = 'info'): void {
  broadcastEvent({
    type: 'system_message',
    data: { message, level },
    timestamp: new Date().toISOString()
  });
} 