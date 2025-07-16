import { NextRequest } from 'next/server';
import { liveProductsClients } from '../../shared/broadcast';

// SSE endpoint for real-time updates
export async function GET(req: NextRequest) {
  const stream = new ReadableStream({
    start(controller) {
      // Add client to the set
      liveProductsClients.add(controller);

      // Send initial connection message
      const initialData = JSON.stringify({
        type: 'connection_established',
        message: 'Connected to dispatchable products stream',
        timestamp: new Date().toISOString()
      });

      controller.enqueue(new TextEncoder().encode(`data: ${initialData}\n\n`));

      // Handle client disconnect
      req.signal.addEventListener('abort', () => {
        liveProductsClients.delete(controller);
        controller.close();
      });
    }
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Cache-Control',
    },
  });
} 