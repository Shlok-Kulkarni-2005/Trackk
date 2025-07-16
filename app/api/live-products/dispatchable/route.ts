import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch only live products (state ON) that can be dispatched
export async function GET() {
  try {
    const jobs = await prisma.job.findMany({
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
    });

    // Group jobs by productId and keep only the latest job for each product
    const latestJobByProduct: { [productId: string]: any } = {};
    jobs.forEach((job: any) => {
      const productId = job.product.id;
      if (
        !latestJobByProduct[productId] ||
        new Date(job.createdAt) > new Date(latestJobByProduct[productId].createdAt)
      ) {
        latestJobByProduct[productId] = job;
      }
    });

    // Filter only live products (state ON) that can be dispatched
    const dispatchableProducts = Object.values(latestJobByProduct)
      .filter((job: any) => job.state === 'ON')
      .map((job: any) => ({
        id: job.product.id.toString(),
        name: job.product.name,
        process: job.machine.name,
        status: job.state,
        createdAt: job.createdAt,
        jobId: job.id,
      }));

    return NextResponse.json({ 
      dispatchableProducts,
      count: dispatchableProducts.length 
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching dispatchable products:', error);
    return NextResponse.json({ error: 'Failed to fetch dispatchable products' }, { status: 500 });
  }
}

// POST: Move a product from live to past (change state from ON to OFF)
export async function POST(req: NextRequest) {
  try {
    const { productId, jobId } = await req.json();
    
    if (!productId || !jobId) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID and Job ID are required' 
      }, { status: 400 });
    }

    // Update the job state to OFF (moving from live to past)
    const updatedJob = await prisma.job.update({
      where: { id: parseInt(jobId) },
      data: { state: 'OFF' },
      include: {
        machine: true,
        product: true,
      },
    });

    // Also update the machine status if needed
    await prisma.machine.update({
      where: { id: updatedJob.machineId },
      data: { status: 'OFF' },
    });

    return NextResponse.json({ 
      success: true, 
      message: 'Product moved from live to past successfully',
      updatedJob 
    }, { status: 200 });
  } catch (error) {
    console.error('Error moving product to past:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to move product to past' 
    }, { status: 500 });
  }
} 