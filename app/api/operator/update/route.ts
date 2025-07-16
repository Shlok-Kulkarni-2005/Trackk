import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// POST: Update product details from operator panel
export async function POST(req: NextRequest) {
  try {
    const { id, processSteps, dispatchStatus, quantity } = await req.json();
    
    if (!id || !processSteps || !dispatchStatus) {
      return NextResponse.json({ 
        success: false, 
        error: 'ID, process steps, and dispatch status are required' 
      }, { status: 400 });
    }

    // Set default quantity if not provided
    const finalQuantity = quantity || 1;

    let updatedRecord;

    // Check if this is a live product (from Job table) or finished product (from OperatorProductUpdate table)
    if (typeof id === 'string' && id.startsWith('job_')) {
      // This is a live product from Job table - create a new OperatorProductUpdate record
      const jobId = parseInt(id.replace('job_', ''));
      
      // Get the job details
      const job = await prisma.job.findUnique({
        where: { id: jobId },
        include: {
          product: true,
          machine: true,
        }
      });

      if (!job) {
        return NextResponse.json({ 
          success: false, 
          error: 'Job not found' 
        }, { status: 404 });
      }

      // Create a new OperatorProductUpdate record
      updatedRecord = await prisma.operatorProductUpdate.create({
        data: {
          operatorId: 12345, // Default operator ID
          product: job.product.name,
          processSteps,
          dispatchStatus,
          dispatchedCost: 0, // Default cost
          quantity: finalQuantity,
        }
      });

    } else {
      // This could be a product ID from our new recently finished products API
      // First, try to find the product and get its latest job
      const product = await prisma.product.findUnique({
        where: { id: parseInt(id) },
        include: {
          jobs: {
            orderBy: { createdAt: 'desc' },
            take: 1,
            include: {
              machine: true
            }
          }
        }
      });

      if (!product) {
        return NextResponse.json({ 
          success: false, 
          error: 'Product not found' 
        }, { status: 404 });
      }

      const latestJob = product.jobs[0];
      if (!latestJob) {
        return NextResponse.json({ 
          success: false, 
          error: 'No job found for this product' 
        }, { status: 404 });
      }

      // Check if there's already an OperatorProductUpdate for this product
      const existingUpdate = await prisma.operatorProductUpdate.findFirst({
        where: { 
          product: product.name,
          archived: false
        },
        orderBy: { createdAt: 'desc' }
      });

      if (existingUpdate) {
        // Update the existing record
        updatedRecord = await prisma.operatorProductUpdate.update({
          where: { id: existingUpdate.id },
          data: {
            processSteps,
            dispatchStatus,
            dispatchedCost: existingUpdate.dispatchedCost,
            quantity: finalQuantity,
          }
        });
      } else {
        // Find a valid operatorId
        const operator = await prisma.operator.findFirst();
        if (!operator) {
          return NextResponse.json({
            success: false,
            error: 'No operator found in the database. Please create an operator first.'
          }, { status: 500 });
        }
        // Create a new OperatorProductUpdate record
        updatedRecord = await prisma.operatorProductUpdate.create({
          data: {
            operatorId: operator.id, // Use a real operator ID
            product: product.name,
            processSteps,
            dispatchStatus,
            dispatchedCost: 0, // Default cost
            quantity: finalQuantity,
          }
        });
      }
    }

    // Since we're always dispatching (status is now always 'Pending'), reduce the quantity from RFD status
    try {
      // First, validate that we have enough quantity available and reduce RFD status count
      const rfdResponse = await fetch(`${req.nextUrl.origin}/api/rfd-status/update-count`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          productId: typeof id === 'string' && id.startsWith('job_') ? 
            (await prisma.job.findUnique({ where: { id: parseInt(id.replace('job_', '')) }, include: { product: true } }))?.productId :
            parseInt(id),
          quantity: finalQuantity
        }),
      });

      if (!rfdResponse.ok) {
        const rfdError = await rfdResponse.json();
        return NextResponse.json({ 
          success: false, 
          error: rfdError.error || 'Failed to reduce product quantity' 
        }, { status: 400 });
      }

      const rfdResult = await rfdResponse.json();
      console.log('RFD status count reduced:', rfdResult.message);

      // Now add to dispatched items
      const productName = updatedRecord.product;
      const dispatchResponse = await fetch(`${req.nextUrl.origin}/api/dispatched/update-quantity`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          product: productName,
          quantity: finalQuantity,
          operatorId: updatedRecord.operatorId
        }),
      });

      if (dispatchResponse.ok) {
        const dispatchResult = await dispatchResponse.json();
        console.log('Product automatically added to dispatched items:', dispatchResult.message);
      } else {
        console.warn('Failed to automatically add product to dispatched items');
      }
    } catch (error) {
      console.warn('Error processing dispatch:', error);
      return NextResponse.json({ 
        success: false, 
        error: 'Failed to process dispatch' 
      }, { status: 500 });
    }

    return NextResponse.json({ 
      success: true, 
      message: 'Product updated successfully',
      update: updatedRecord
    }, { status: 200 });

  } catch (error) {
    console.error('Error updating product:', error);
    console.error('Error details:', {
      message: error instanceof Error ? error.message : 'Unknown error',
      stack: error instanceof Error ? error.stack : undefined
    });
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to update product',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 