import { NextRequest } from 'next/server';

export async function GET(req: NextRequest) {
  const { readable, writable } = new TransformStream();
  const writer = writable.getWriter();

  // Write initial comment to keep connection open
  writer.write(new TextEncoder().encode(':ok\n\n'));

  // No-op for addClient/removeClient since sse.ts is empty

  req.signal.addEventListener('abort', () => {
    writer.close();
  });

  return new Response(readable, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      Connection: 'keep-alive',
    },
  });
} 