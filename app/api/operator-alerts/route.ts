import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { getCurrentOperatorId } from '../../lib/operator-auth';

const prisma = new PrismaClient();

// Get unread alert count for the current operator
export async function GET() {
  try {
    const operatorId = await getCurrentOperatorId();
    
    if (!operatorId) {
      return NextResponse.json({ error: 'Operator not authenticated' }, { status: 401 });
    }

    const unreadCount = await prisma.operatorAlertStatus.count({
      where: {
        operatorId: operatorId,
        read: false,
      },
    });
    return NextResponse.json({ success: true, unreadCount });
  } catch {
    return NextResponse.json({ success: false, error: 'Failed to fetch unread count' }, { status: 500 });
  }
}

// Mark alerts as read for the current operator
export async function POST() {
  try {
    const operatorId = await getCurrentOperatorId();
    
    if (!operatorId) {
      return NextResponse.json({ error: 'Operator not authenticated' }, { status: 401 });
    }

    await prisma.operatorAlertStatus.updateMany({
      where: {
        operatorId: operatorId,
        read: false,
      },
      data: {
        read: true,
      },
    });
    return NextResponse.json({ success: true, message: 'Alerts marked as read.' });
  } catch {
    return NextResponse.json({ success: false, error: 'Failed to mark alerts as read' }, { status: 500 });
  }
} 