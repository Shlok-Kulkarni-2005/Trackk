import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { broadcastProductMovedToPast } from '../../shared/broadcast';

const prisma = new PrismaClient();

// POST: Move a product from live to past automatically
export async function POST(req: NextRequest) {
  try {
    const { productId, jobId, reason = 'dispatched', quantity = 1 } = await req.json();
    
    if (!productId || (!jobId && !productId) || quantity <= 0) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID and valid quantity are required' 
      }, { status: 400 });
    }

    // Find the product
    const product = await prisma.product.findUnique({
      where: { id: parseInt(productId) },
    });
    if (!product) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product not found' 
      }, { status: 404 });
    }

    // Find all ON jobs for this product (optionally filter by jobId's machine if provided)
    let onJobs;
    if (jobId) {
      // If jobId is provided, use its machine for filtering
      const job = await prisma.job.findUnique({
        where: { id: parseInt(jobId) },
        include: { machine: true },
      });
      if (!job) {
        return NextResponse.json({ 
          success: false, 
          error: 'Job not found' 
        }, { status: 404 });
      }
      onJobs = await prisma.job.findMany({
        where: {
          productId: parseInt(productId),
          machineId: job.machineId,
          state: 'ON',
        },
        orderBy: { createdAt: 'asc' },
      });
    } else {
      // Fallback: all ON jobs for product
      onJobs = await prisma.job.findMany({
        where: {
          productId: parseInt(productId),
          state: 'ON',
        },
        orderBy: { createdAt: 'asc' },
      });
    }

    if (onJobs.length < quantity) {
      return NextResponse.json({ 
        success: false, 
        error: `Insufficient quantity. Only ${onJobs.length} available in ON state.` 
      }, { status: 400 });
    }

    // Update the oldest N ON jobs to OFF
    const jobsToUpdate = onJobs.slice(0, quantity);
    const updatedJobs = [];
    for (const job of jobsToUpdate) {
      const updatedJob = await prisma.job.update({
        where: { id: job.id },
        data: { state: 'OFF' },
        include: { machine: true, product: true },
      });
      updatedJobs.push(updatedJob);
      // Optionally update the machine status to OFF (if all jobs for this machine are now OFF)
      const stillOn = await prisma.job.count({
        where: { machineId: job.machineId, state: 'ON' },
      });
      if (stillOn === 0) {
        await prisma.machine.update({
          where: { id: job.machineId },
          data: { status: 'OFF' },
        });
      }
    }

    // Archive OperatorProductUpdate records for this product (if needed)
    await prisma.operatorProductUpdate.updateMany({
      where: { product: product.name, archived: false },
      data: { archived: true },
    });

    // Log the transition for tracking purposes
    console.log(`Product ${product.name} moved from live to past. Reason: ${reason}. Quantity: ${quantity}`);

    // Broadcast the change to all connected clients (use the first updated job for context)
    if (updatedJobs.length > 0) {
      broadcastProductMovedToPast(
        product.id.toString(),
        product.name
      );
    }

    return NextResponse.json({ 
      success: true, 
      message: `Moved ${quantity} quantity from live to past successfully. Reason: ${reason}`,
      product: {
        id: product.id,
        name: product.name,
        previousState: 'ON',
        newState: 'OFF',
        transitionReason: reason,
        transitionTime: new Date().toISOString(),
        quantityMoved: quantity,
      },
      jobsUpdated: updatedJobs.map(j => j.id),
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