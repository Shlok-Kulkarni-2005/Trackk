import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { broadcastAlert } from './stream';

const prisma = new PrismaClient();

// GET: Fetch all alerts
export async function GET() {
  try {
    const alerts = await prisma.alert.findMany({
      orderBy: { createdAt: 'desc' },
    });
    return NextResponse.json({ alerts }, { status: 200 });
  } catch (error) {
    console.error('Error fetching alerts:', error);
    return NextResponse.json({ error: 'Failed to fetch alerts' }, { status: 500 });
  }
}

// POST: Send a new alert
export async function POST(req: NextRequest) {
  try {
    const { message, icon, senderId } = await req.json();
    if (!message) {
      return NextResponse.json({ error: 'Message is required' }, { status: 400 });
    }

    const operators = await prisma.operator.findMany({ select: { id: true } });

    const alert = await prisma.alert.create({
      data: {
        message,
        icon,
        senderId,
        recipients: {
          create: operators.map(op => ({
            operatorId: op.id,
            read: false,
          })),
        },
      },
    });

    broadcastAlert(alert);
    return NextResponse.json({ alert }, { status: 201 });
  } catch (error) {
    console.error('Error sending alert:', error);
    return NextResponse.json({ error: 'Failed to send alert' }, { status: 500 });
  }
} 