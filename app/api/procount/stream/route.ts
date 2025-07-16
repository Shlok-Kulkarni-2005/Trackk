import { NextRequest } from 'next/server';


export async function GET(req: NextRequest) {
  const stream = new ReadableStream({
    start(controller) {
      // Store controller reference for broadcasting
      (req as any).controller = controller;
      
      // Send initial connection message
      controller.enqueue(new TextEncoder().encode('data: {"type": "connected"}\n\n'));
      
      req.signal.addEventListener('abort', () => {
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