import { NextRequest, NextResponse } from 'next/server';
import { rfdStatusClients } from '../../shared/broadcast';

// GET: Server-Sent Events stream for RFD status updates
export async function GET(req: NextRequest) {
  try {
    const stream = new ReadableStream({
      start(controller) {
        // Send initial connection message
        const initialData = JSON.stringify({
          type: 'connection_established',
          data: {
            message: 'RFD status stream connected',
            timestamp: new Date().toISOString()
          }
        });
        
        controller.enqueue(new TextEncoder().encode(`data: ${initialData}\n\n`));
        
        // Add client to the set
        rfdStatusClients.add(controller);
        
        // Handle client disconnect
        req.signal.addEventListener('abort', () => {
          rfdStatusClients.delete(controller);
          console.log('RFD status client disconnected');
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
        'Access-Control-Allow-Headers': 'Cache-Control'
      }
    });

  } catch (error) {
    console.error('Error setting up RFD status stream:', error);
    return NextResponse.json({ 
      error: 'Failed to establish RFD status stream' 
    }, { status: 500 });
  }
} 