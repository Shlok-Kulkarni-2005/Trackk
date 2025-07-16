import { NextRequest } from 'next/server';

const clients: WritableStreamDefaultWriter[] = [];

export function addJobClient(client: WritableStreamDefaultWriter) {
  clients.push(client);
}

export function removeJobClient(client: WritableStreamDefaultWriter) {
  const idx = clients.indexOf(client);
  if (idx !== -1) {
    clients.splice(idx, 1);
  }
}

export function broadcastJob(job: any) {
  const data = `data: ${JSON.stringify(job)}\n\n`;
  for (const client of clients) {
    client.write(new TextEncoder().encode(data));
  }
}

export async function GET(req: NextRequest) {
  const { readable, writable } = new TransformStream();
  const writer = writable.getWriter();
  writer.write(new TextEncoder().encode(':ok\n\n'));
  addJobClient(writer);
  req.signal.addEventListener('abort', () => {
    removeJobClient(writer);
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