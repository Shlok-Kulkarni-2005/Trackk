import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { broadcastProductMovedToPast } from '../../shared/broadcast';

const prisma = new PrismaClient();

// GET: Get product lifecycle information
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

    // Get all jobs for this product to see complete lifecycle
    const jobs = await prisma.job.findMany({
      where: { productId: parseInt(productId) },
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'asc' }, // Chronological order
    });

    if (jobs.length === 0) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product not found' 
      }, { status: 404 });
    }

    const lifecycle = {
      productId: parseInt(productId),
      productName: jobs[0].product.name,
      currentState: jobs[jobs.length - 1].state,
      currentMachine: jobs[jobs.length - 1].machine.name,
      totalStages: jobs.length,
      stages: jobs.map((job, index) => ({
        stageNumber: index + 1,
        machine: job.machine.name,
        state: job.state,
        stage: job.stage,
        timestamp: job.createdAt,
        duration: index > 0 ? 
          new Date(job.createdAt).getTime() - new Date(jobs[index - 1].createdAt).getTime() : 
          null
      })),
      isDispatchable: jobs[jobs.length - 1].state === 'ON',
      createdAt: jobs[0].createdAt,
      lastUpdated: jobs[jobs.length - 1].createdAt,
    };

    return NextResponse.json({ 
      success: true,
      lifecycle 
    }, { status: 200 });
  } catch {
    console.error('Error fetching product lifecycle');
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to fetch product lifecycle' 
    }, { status: 500 });
  }
}

// POST: Update product lifecycle (move between states)
export async function POST(req: NextRequest) {
  try {
    const { productId, jobId, action, newState, reason } = await req.json();
    
    if (!productId || !jobId || !action) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product ID, Job ID, and action are required' 
      }, { status: 400 });
    }

    // Get current job
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

    let updatedJob;
    let message = '';

    switch (action) {
      case 'move_to_past':
        if (currentJob.state !== 'ON') {
          return NextResponse.json({ 
            success: false, 
            error: 'Product is not currently live' 
          }, { status: 400 });
        }

        updatedJob = await prisma.job.update({
          where: { id: parseInt(jobId) },
          data: { state: 'OFF', updatedAt: new Date() },
          include: {
            machine: true,
            product: true,
          },
        });

        // Update machine status
        await prisma.machine.update({
          where: { id: updatedJob.machineId },
          data: { status: 'OFF' },
        });

        // Broadcast the change
        broadcastProductMovedToPast(
          currentJob.product.id.toString(),
          currentJob.product.name
        );

        // Also remove from finished products if it exists there
        try {
          await prisma.operatorProductUpdate.deleteMany({
            where: {
              product: currentJob.product.name
            }
          });
          console.log(`Removed product ${currentJob.product.name} from finished products`);
        } catch (error) {
          console.log(`Product ${currentJob.product.name} not found in finished products or already removed`);
        }

        message = `Product moved from live to past. Reason: ${reason || 'manual_action'}`;
        break;

      case 'reactivate':
        if (currentJob.state !== 'OFF') {
          return NextResponse.json({ 
            success: false, 
            error: 'Product is not currently in past state' 
          }, { status: 400 });
        }

        updatedJob = await prisma.job.update({
          where: { id: parseInt(jobId) },
          data: { state: 'ON' },
          include: {
            machine: true,
            product: true,
          },
        });

        // Update machine status
        await prisma.machine.update({
          where: { id: updatedJob.machineId },
          data: { status: 'ON' },
        });

        message = `Product reactivated from past to live. Reason: ${reason || 'manual_action'}`;
        break;

      case 'update_state':
        if (!newState) {
          return NextResponse.json({ 
            success: false, 
            error: 'New state is required for update_state action' 
          }, { status: 400 });
        }

        updatedJob = await prisma.job.update({
          where: { id: parseInt(jobId) },
          data: { state: newState, updatedAt: new Date() },
          include: {
            machine: true,
            product: true,
          },
        });

        // Update machine status to match
        await prisma.machine.update({
          where: { id: updatedJob.machineId },
          data: { status: newState },
        });

        // If moving to OFF, broadcast the change and remove from finished products
        if (newState === 'OFF' && currentJob.state === 'ON') {
          broadcastProductMovedToPast(
            currentJob.product.id.toString(),
            currentJob.product.name
          );

          // Also remove from finished products if it exists there
          try {
            await prisma.operatorProductUpdate.deleteMany({
              where: {
                product: currentJob.product.name
              }
            });
            console.log(`Removed product ${currentJob.product.name} from finished products`);
          } catch (error) {
            console.log(`Product ${currentJob.product.name} not found in finished products or already removed`);
          }
        }

        message = `Product state updated to ${newState}. Reason: ${reason || 'manual_action'}`;
        break;

      default:
        return NextResponse.json({ 
          success: false, 
          error: 'Invalid action. Supported actions: move_to_past, reactivate, update_state' 
        }, { status: 400 });
    }

    console.log(`Product lifecycle updated: ${message}`);

    return NextResponse.json({ 
      success: true, 
      message,
      updatedJob,
      previousState: currentJob.state,
      newState: updatedJob.state,
      action,
      reason: reason || 'manual_action',
      timestamp: new Date().toISOString(),
    }, { status: 200 });
  } catch {
    console.error('Error updating product lifecycle');
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to update product lifecycle' 
    }, { status: 500 });
  }
} 