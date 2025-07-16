import { NextRequest } from 'next/server';

// Store connected clients
const clients: WritableStreamDefaultWriter[] = [];

export function addAlertClient(client: WritableStreamDefaultWriter) {
  clients.push(client);
}

export function removeAlertClient(client: WritableStreamDefaultWriter) {
  const idx = clients.indexOf(client);
  if (idx !== -1) {
    clients.splice(idx, 1);
  }
}

export function broadcastAlert(alert: any) {
  const data = `data: ${JSON.stringify(alert)}\n\n`;
  for (const client of clients) {
    client.write(new TextEncoder().encode(data));
  }
}

export async function GET(req: NextRequest) {
  const { readable, writable } = new TransformStream();
  const writer = writable.getWriter();

  // Keep connection open
  writer.write(new TextEncoder().encode(':ok\n\n'));
  addAlertClient(writer);

  req.signal.addEventListener('abort', () => {
    removeAlertClient(writer);
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