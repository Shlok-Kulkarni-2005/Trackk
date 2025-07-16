import { NextRequest } from 'next/server';

const clients = new Set<ReadableStreamDefaultController>();

export function broadcastProductCount(productCount: any) {
  const message = `data: ${JSON.stringify(productCount)}\n\n`;
  
  clients.forEach((client) => {
    try {
      client.enqueue(new TextEncoder().encode(message));
    } catch (error) {
      console.error('Error broadcasting to client:', error);
    }
  });
}

export async function GET(req: NextRequest) {
  const stream = new ReadableStream({
    start(controller) {
      clients.add(controller);
      
      // Send initial connection message
      controller.enqueue(new TextEncoder().encode('data: {"type": "connected"}\n\n'));
      
      req.signal.addEventListener('abort', () => {
        clients.delete(controller);
        controller.close();
      });
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    },
  });
} 