import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch detailed information about dispatched items
export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const productId = searchParams.get('productId');

    const whereClause: Record<string, any> = {
      dispatchStatus: 'Dispatched',
    };

    if (productId) {
      whereClause.product = productId;
    }

    const dispatchedDetails = await prisma.operatorProductUpdate.findMany({
      where: whereClause,
      include: {
        operator: {
          select: {
            id: true,
            username: true,
            phone: true,
          },
        },
      },
      orderBy: {
        createdAt: 'desc',
      },
    });

    // Transform the data to include more detailed information
    const transformedDetails = dispatchedDetails.map(item => ({
      id: item.id.toString(),
      product: item.product,
      processSteps: item.processSteps,
      dispatchStatus: item.dispatchStatus,
      dispatchedCost: item.dispatchedCost,
      dispatchedAt: item.createdAt.toISOString(),
      operator: item.operator ? {
        id: item.operator.id,
        username: item.operator.username || 'Unknown',
        phone: item.operator.phone,
      } : null,
    }));

    return NextResponse.json({ 
      success: true, 
      dispatchedDetails: transformedDetails 
    });
  } catch (error) {
    console.error('Error fetching dispatched details:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
} 