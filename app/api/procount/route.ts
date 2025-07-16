import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch all product counts
export async function GET() {
  try {
    // Get all jobs with RFD machine and OFF status, grouped by product
    const jobs = await prisma.job.findMany({
      where: {
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      },
      include: {
        product: true,
        machine: true
      },
      orderBy: { createdAt: 'desc' }
    });

    // Group by product and count occurrences
    const productCountsMap = new Map();
    
    jobs.forEach(job => {
      const productId = job.product.id;
      if (!productCountsMap.has(productId)) {
        productCountsMap.set(productId, {
          id: productId,
          product: job.product,
          count: 0,
          machine: job.machine.name,
          stage: job.stage,
          state: job.state,
          createdAt: job.createdAt,
          updatedAt: job.createdAt
        });
      }
      productCountsMap.get(productId).count++;
    });

    const productCounts = Array.from(productCountsMap.values());
    
    return NextResponse.json({ productCounts }, { status: 200 });
  } catch (error) {
    console.error('Failed to fetch product counts:', error);
    return NextResponse.json({ 
      error: 'Failed to fetch product counts',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
}

// POST: Only validate input and return success, do not create jobs
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { productName, machine, state } = body;

    // Only track when RFD machine is selected and status is OFF
    if (machine !== 'RFD' || state !== 'OFF') {
      return NextResponse.json({ 
        message: 'Product count only tracked for RFD machine with OFF status' 
      }, { status: 200 });
    }

    if (!productName || !machine || !state) {
      return NextResponse.json({ 
        error: 'All fields are required' 
      }, { status: 400 });
    }

    // No job creation here! Just acknowledge
    return NextResponse.json({ 
      message: 'Product count notification received (no job created)' 
    }, { status: 201 });

  } catch (error) {
    console.error('Failed to process product count notification:', error);
    return NextResponse.json({ 
      error: 'Failed to process product count notification',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
}

// DELETE: Clear all product counts (delete all RFD machine OFF jobs)
export async function DELETE() {
  try {
    await prisma.job.deleteMany({
      where: {
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      }
    });
    return NextResponse.json({ 
      message: 'All RFD machine OFF jobs deleted successfully'
    }, { status: 200 });
  } catch (error) {
    return NextResponse.json({ 
      error: 'Failed to clear product counts',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 