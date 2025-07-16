import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// POST: Reduce quantity for a specific product when dispatched
export async function POST(req: NextRequest) {
  try {
    const { productId, quantity } = await req.json();
    
    if (!productId || !quantity || quantity <= 0) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID and valid quantity are required' 
      }, { status: 400 });
    }

    // Get all jobs for this product with RFD machine and OFF status
    const jobs = await prisma.job.findMany({
      where: {
        productId: parseInt(productId),
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      },
      orderBy: { createdAt: 'desc' }
    });

    if (jobs.length < quantity) {
      return NextResponse.json({ 
        success: false, 
        error: `Insufficient quantity. Available: ${jobs.length}, Requested: ${quantity}` 
      }, { status: 400 });
    }

    // Delete the oldest jobs (FIFO - First In, First Out) to reduce the quantity
    const jobsToDelete = jobs.slice(0, quantity);
    const jobIdsToDelete = jobsToDelete.map(job => job.id);

    await prisma.job.deleteMany({
      where: {
        id: {
          in: jobIdsToDelete
        }
      }
    });

    console.log(`Reduced quantity for product ${productId} by ${quantity}. Deleted ${jobIdsToDelete.length} jobs.`);

    return NextResponse.json({ 
      success: true, 
      message: `Successfully reduced quantity by ${quantity}`,
      remainingQuantity: jobs.length - quantity
    }, { status: 200 });

  } catch (error) {
    console.error('Error reducing product quantity:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to reduce product quantity',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 