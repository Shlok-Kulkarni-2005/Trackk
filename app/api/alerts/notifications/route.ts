import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { getCurrentOperatorId } from '../../../lib/operator-auth';

const prisma = new PrismaClient();

export async function GET(req: NextRequest) {
  try {
    const operatorId = await getCurrentOperatorId();
    
    if (!operatorId) {
      return NextResponse.json({ error: 'Operator not authenticated' }, { status: 401 });
    }

    const stream = new ReadableStream({
      start(controller) {
        const sendEvent = (data: any) => {
          try {
            controller.enqueue(`data: ${JSON.stringify(data)}\n\n`);
          } catch (error) {
            console.error('Error sending SSE event:', error);
          }
        };

        // Send initial unread count
        const sendInitialCount = async () => {
          try {
            const unreadCount = await prisma.operatorAlertStatus.count({
              where: {
                operatorId: operatorId,
                read: false,
              },
            });
            sendEvent({ type: 'unreadCount', count: unreadCount });
          } catch (error) {
            console.error('Error fetching initial unread count:', error);
            sendEvent({ type: 'error', message: 'Failed to fetch unread count' });
          }
        };

        sendInitialCount();

        // Set up interval to check for new alerts
        const interval = setInterval(async () => {
          try {
            const unreadCount = await prisma.operatorAlertStatus.count({
              where: {
                operatorId: operatorId,
                read: false,
              },
            });
            sendEvent({ type: 'unreadCount', count: unreadCount });
          } catch (error) {
            console.error('Error checking unread count:', error);
            // Don't send error events repeatedly, just log them
          }
        }, 5000); // Check every 5 seconds

        // Clean up on close
        req.signal.addEventListener('abort', () => {
          clearInterval(interval);
          try {
            controller.close();
          } catch (error) {
            console.error('Error closing SSE stream:', error);
          }
        });
      },
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
  } catch (error) {
    console.error('SSE endpoint error:', error);
    return NextResponse.json({ error: 'Internal server error' }, { status: 500 });
  }
} 