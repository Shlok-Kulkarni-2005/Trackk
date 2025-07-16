import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { broadcastJob } from './stream';
import { broadcastProduct } from '../products/stream';
import { broadcastProductCount } from '../product-count/stream';
import { broadcastProductAddedToLive } from '../shared/broadcast';


const prisma = new PrismaClient();

// GET: Fetch all jobs with machine and product info
export async function GET() {
  try {
    const jobs = await prisma.job.findMany({
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
    });
    return NextResponse.json({ jobs }, { status: 200 });
  } catch {
    return NextResponse.json({ error: 'Failed to fetch jobs' }, { status: 500 });
  }
}

// POST: Add a new job
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { machine, product, state, stage, quantity = 1 } = body;
    if (!machine || !product || !state || !stage) {
      console.error('Job upload failed: Missing fields', { machine, product, state, stage, body });
      return NextResponse.json({ error: 'All fields are required', received: body }, { status: 400 });
    }

    // Validate quantity
    const parsedQuantity = parseInt(quantity);
    if (isNaN(parsedQuantity) || parsedQuantity <= 0 || parsedQuantity > 1000) {
      return NextResponse.json({ error: 'Quantity must be between 1 and 1000' }, { status: 400 });
    }
    const jobQuantity = parsedQuantity;
    // Find or create machine
    let dbMachine = await prisma.machine.findFirst({ where: { name: machine } });
    if (!dbMachine) {
      dbMachine = await prisma.machine.create({ data: { name: machine, status: state } });
    } else {
      // Optionally update status
      await prisma.machine.update({ where: { id: dbMachine.id }, data: { status: state } });
    }
    // Find or create product (case-insensitive, trimmed)
    let dbProduct = await prisma.product.findFirst({
      where: { name: { equals: product.trim(), mode: 'insensitive' } }
    });
    if (!dbProduct) {
      dbProduct = await prisma.product.create({ data: { name: product.trim() } });
    }

    // Handle job creation/updates based on state
    const jobs = [];
    
    if (state === 'OFF') {
      // For OFF state, we need to update existing ON jobs to OFF
      // First, validate that we have enough ON jobs to turn OFF
      const existingOnJobs = await prisma.job.findMany({
        where: {
          productId: dbProduct.id,
          machineId: dbMachine.id,
          state: 'ON'
        },
        orderBy: { createdAt: 'desc' }
      });

      if (existingOnJobs.length < jobQuantity) {
        return NextResponse.json({ 
          error: `Cannot turn OFF ${jobQuantity} quantities. Only ${existingOnJobs.length} quantities are available in ON state.` 
        }, { status: 400 });
      }

      // Update the oldest ON jobs to OFF (FIFO - First In, First Out)
      const jobsToUpdate = existingOnJobs.slice(0, jobQuantity);
      for (const jobToUpdate of jobsToUpdate) {
        const updatedJob = await prisma.job.update({
          where: { id: jobToUpdate.id },
          data: { state: 'OFF', updatedAt: new Date() },
          include: {
            machine: true,
            product: true,
          },
        });
        jobs.push(updatedJob);
        broadcastJob(updatedJob);
      }
    } else {
      // For ON state, create new jobs
      for (let i = 0; i < jobQuantity; i++) {
        const job = await prisma.job.create({
          data: {
            machineId: dbMachine.id,
            productId: dbProduct.id,
            state,
            stage,
          },
          include: {
            machine: true,
            product: true,
          },
        });
        jobs.push(job);
        broadcastJob(job);
      }
    }
    
    // If RFD machine and OFF status, update product count for each job created
    if (machine === 'RFD' && state === 'OFF') {
      for (let i = 0; i < jobQuantity; i++) {
        try {
          await fetch(`${req.nextUrl.origin}/api/procount`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              productName: dbProduct.name,
              machine: machine,
              stage: stage,
              state: state
            })
          });
        } catch (error) {
          console.error('Failed to update product count:', error);
        }
      }
      console.log(`Product count updated for ${dbProduct.name} with RFD machine and OFF status (${jobQuantity} times)`);
    }
    
    // Fetch the product with its latest job and broadcast for real-time product list
    const productWithLatestJob = await prisma.product.findUnique({
      where: { id: dbProduct.id },
      include: {
        jobs: {
          orderBy: { createdAt: 'desc' },
          take: 1,
        },
      },
    });
    if (productWithLatestJob) {
      const updatedProduct = {
        id: productWithLatestJob.id,
        name: productWithLatestJob.name,
        process: productWithLatestJob.jobs[0]?.stage || '',
        status: productWithLatestJob.jobs[0]?.state || '',
      };
      broadcastProduct(updatedProduct);
    }

    // If the product is added with state 'ON', broadcast to Recently Finished Products dropdown
    if (state === 'ON') {
      // Only broadcast once per product, not per job
      broadcastProductAddedToLive(
        dbProduct.id.toString(),
        dbProduct.name,
        machine
      );
      console.log(`Product ${dbProduct.name} added to live products and will appear in Recently Finished Products dropdown`);
    }

    // After broadcasting product, also broadcast product count if state is 'ON' and machine is 'CNC Finished'
    if (state === 'ON' && machine === 'CNC Finished') {
      const cncMachine = await prisma.machine.findFirst({
        where: { name: 'CNC Finished' },
        select: { id: true },
      });

      if (cncMachine) {
        const count = await prisma.job.count({
          where: {
            productId: dbProduct.id,
            machineId: cncMachine.id,
            state: 'ON',
          },
        });

        const productCount = {
          id: dbProduct.id,
          name: dbProduct.name,
          count: count,
          status: 'ON',
        };
        broadcastProductCount(productCount);
      }
    }
    return NextResponse.json({ 
      jobs, 
      count: jobs.length,
      message: `Successfully created ${jobs.length} job(s) for ${dbProduct.name}`
    }, { status: 201 });
  } catch {
    console.error('Job upload failed');
    return NextResponse.json({ error: 'Failed to add job' }, { status: 500 });
  }
}

export async function DELETE() {
  try {
    await prisma.job.deleteMany({});
    return NextResponse.json({ success: true });
  } catch {
    return NextResponse.json({ error: 'Failed to clear jobs' }, { status: 500 });
  }
} 