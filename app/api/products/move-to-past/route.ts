import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { broadcastProductMovedToPast } from '../../shared/broadcast';

const prisma = new PrismaClient();

// POST: Move a product from live to past automatically
export async function POST(req: NextRequest) {
  try {
    const { productId, jobId, reason = 'dispatched' } = await req.json();
    
    if (!productId || !jobId) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID and Job ID are required' 
      }, { status: 400 });
    }

    // First, verify the product is currently live (state ON)
    const currentJob = await prisma.job.findUnique({
      where: { id: parseInt(jobId) },
      include: {
        machine: true,
        product: true,
      },
    });

    if (!currentJob) {
      return NextResponse.json({ 
        success: false, 
        error: 'Job not found' 
      }, { status: 404 });
    }

    if (currentJob.state !== 'ON') {
      return NextResponse.json({ 
        success: false, 
        error: 'Product is not currently live' 
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

    // Archive all OperatorProductUpdate records for this product
    await prisma.operatorProductUpdate.updateMany({
      where: { product: updatedJob.product.name, archived: false },
      data: { archived: true },
    });

    // Update the machine status to OFF as well
    await prisma.machine.update({
      where: { id: updatedJob.machineId },
      data: { status: 'OFF' },
    });

    // Log the transition for tracking purposes
    console.log(`Product ${currentJob.product.name} moved from live to past. Reason: ${reason}`);

    // Broadcast the change to all connected clients
    broadcastProductMovedToPast(
      currentJob.product.id.toString(),
      currentJob.product.name
    );

    return NextResponse.json({ 
      success: true, 
      message: `Product moved from live to past successfully. Reason: ${reason}`,
      product: {
        id: updatedJob.product.id,
        name: updatedJob.product.name,
        previousState: 'ON',
        newState: 'OFF',
        machine: updatedJob.machine.name,
        transitionReason: reason,
        transitionTime: new Date().toISOString(),
      }
    }, { status: 200 });
  } catch (error) {
    console.error('Error moving product to past:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to move product to past' 
    }, { status: 500 });
  }
}

// GET: Get transition history for a product
export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const productId = searchParams.get('productId');
    
    if (!productId) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID is required' 
      }, { status: 400 });
    }

    // Get all jobs for this product to see transition history
    const jobs = await prisma.job.findMany({
      where: { productId: parseInt(productId) },
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
    });

    const transitionHistory = jobs.map(job => ({
      jobId: job.id,
      state: job.state,
      machine: job.machine.name,
      transitionTime: job.createdAt,
      stage: job.stage,
    }));

    return NextResponse.json({ 
      success: true,
      productId: parseInt(productId),
      transitionHistory,
      currentState: jobs[0]?.state || 'unknown'
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching transition history:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to fetch transition history' 
    }, { status: 500 });
  }
} 