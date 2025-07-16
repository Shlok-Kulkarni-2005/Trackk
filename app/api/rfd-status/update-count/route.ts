import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// POST: Update RFD status count when products are dispatched
export async function POST(req: NextRequest) {
  try {
    const { productId, quantity } = await req.json();
    
    if (!productId || !quantity || quantity <= 0) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID and valid quantity are required' 
      }, { status: 400 });
    }

    // Get the product name from the product ID
    const product = await prisma.product.findUnique({
      where: { id: parseInt(productId) }
    });

    if (!product) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product not found' 
      }, { status: 404 });
    }

    // Get all jobs for this product with RFD machine and OFF status
    const rfdJobs = await prisma.job.findMany({
      where: {
        productId: parseInt(productId),
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      },
      orderBy: { createdAt: 'desc' }
    });

    if (rfdJobs.length < quantity) {
      return NextResponse.json({ 
        success: false, 
        error: `Insufficient RFD quantity. Available: ${rfdJobs.length}, Requested: ${quantity}` 
      }, { status: 400 });
    }

    // Delete the oldest RFD jobs (FIFO - First In, First Out) to reduce the quantity
    const jobsToDelete = rfdJobs.slice(0, quantity);
    const jobIdsToDelete = jobsToDelete.map(job => job.id);

    await prisma.job.deleteMany({
      where: {
        id: {
          in: jobIdsToDelete
        }
      }
    });

    console.log(`Reduced RFD count for product ${product.name} by ${quantity}. Deleted ${jobIdsToDelete.length} RFD jobs.`);

    // Broadcast the update to connected clients
    try {
      await fetch(`${req.nextUrl.origin}/api/rfd-status/broadcast`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          type: 'rfd_count_updated',
          productId: parseInt(productId),
          productName: product.name,
          quantityReduced: quantity,
          remainingCount: rfdJobs.length - quantity
        }),
      });
    } catch (broadcastError) {
      console.warn('Failed to broadcast RFD count update:', broadcastError);
    }

    return NextResponse.json({ 
      success: true, 
      message: `Successfully reduced RFD count by ${quantity}`,
      remainingCount: rfdJobs.length - quantity,
      productName: product.name
    }, { status: 200 });

  } catch (error) {
    console.error('Error updating RFD status count:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to update RFD status count',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
}

// GET: Get current RFD count for a specific product
export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const productId = searchParams.get('productId');
    
    if (!productId) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID parameter is required' 
      }, { status: 400 });
    }

    // Count RFD jobs for this product
    const rfdCount = await prisma.job.count({
      where: {
        productId: parseInt(productId),
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      }
    });

    return NextResponse.json({ 
      success: true, 
      productId: parseInt(productId),
      rfdCount: rfdCount
    }, { status: 200 });

  } catch (error) {
    console.error('Error getting RFD count:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
} 