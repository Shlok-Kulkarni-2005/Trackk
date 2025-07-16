import { NextRequest } from 'next/server';
import { finishedProductsClients } from '../../../shared/broadcast';

// GET: Server-Sent Events endpoint for real-time updates
export async function GET(req: NextRequest) {
  const stream = new ReadableStream({
    start(controller) {
      // Add client to the set
      finishedProductsClients.add(controller);

      // Send initial connection message
      const connectionMessage = JSON.stringify({
        type: 'connection_established',
        data: {
          message: 'Connected to finished products stream',
          timestamp: new Date().toISOString()
        }
      });

      controller.enqueue(new TextEncoder().encode(`data: ${connectionMessage}\n\n`));

      // Handle client disconnect
      req.signal.addEventListener('abort', () => {
        finishedProductsClients.delete(controller);
        console.log('Client disconnected from finished products stream');
      });
    }
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET',
      'Access-Control-Allow-Headers': 'Cache-Control',
    },
  });
} 