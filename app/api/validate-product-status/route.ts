import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// POST: Validate if a product can be set to OFF status
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { productName, machine, quantity = 1 } = body;

    if (!productName || !machine) {
      return NextResponse.json({ 
        error: 'Product name and machine are required' 
      }, { status: 400 });
    }

    // Validate quantity
    const requestedQuantity = parseInt(quantity);
    if (isNaN(requestedQuantity) || requestedQuantity <= 0 || requestedQuantity > 1000) {
      return NextResponse.json({ 
        canSetOff: false,
        reason: 'Invalid quantity. Must be between 1 and 1000'
      }, { status: 200 });
    }

    // Find the product
    const product = await prisma.product.findFirst({
      where: { name: { equals: productName.trim(), mode: 'insensitive' } }
    });

    if (!product) {
      return NextResponse.json({ 
        canSetOff: false,
        reason: 'Product not found in system'
      }, { status: 200 });
    }

    // Find the machine
    const machineRecord = await prisma.machine.findFirst({
      where: { name: machine }
    });

    if (!machineRecord) {
      return NextResponse.json({ 
        canSetOff: false,
        reason: 'Machine not found in system'
      }, { status: 200 });
    }

    // Get all jobs for this product on the selected machine and stage, ordered by createdAt DESC
    const jobs = await prisma.job.findMany({
      where: {
        productId: product.id,
        machineId: machineRecord.id,
        stage: body.stage, // Only match jobs with the same process/stage
      },
      include: {
        machine: true
      },
      orderBy: { createdAt: 'desc' }
    });

    // Count how many jobs are in ON state for this product on this machine and stage
    const onJobs = jobs.filter(job => job.state === 'ON');
    const availableQuantity = onJobs.length;

    if (availableQuantity === 0) {
      return NextResponse.json({ 
        canSetOff: false,
        reason: 'Please turn on the product on this machine and process first'
      }, { status: 200 });
    }

    // Check if requested quantity exceeds available quantity
    if (requestedQuantity > availableQuantity) {
      return NextResponse.json({ 
        canSetOff: false,
        reason: `Cannot turn OFF ${requestedQuantity} quantities. Only ${availableQuantity} quantities are available in ON state on this machine and process.`
      }, { status: 200 });
    }

    // No need to check other machines or processes; allow independent OFF

    // If we reach here, product is ON only on the selected machine and process and quantity is sufficient
    return NextResponse.json({ 
      canSetOff: true,
      reason: `Product can be turned OFF. ${availableQuantity} quantities available, ${requestedQuantity} requested.`,
      availableQuantity: availableQuantity,
      requestedQuantity: requestedQuantity,
      liveProduct: {
        id: onJobs[0]?.id,
        name: product.name,
        machine: machineRecord.name,
        stage: body.stage,
        state: 'ON',
        createdAt: onJobs[0]?.createdAt
      }
    }, { status: 200 });

  } catch (error) {
    console.error('Error validating product status:', error);
    return NextResponse.json({ 
      error: 'Failed to validate product status',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 