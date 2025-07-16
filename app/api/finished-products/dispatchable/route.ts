import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch finished products that should appear in Recently Finished Products dropdown
export async function GET() {
  try {
    console.log('Fetching finished products for dropdown...');
    
    // Get finished products from OperatorProductUpdate table
    const finishedProducts = await prisma.operatorProductUpdate.findMany({
      where: { archived: false },
      orderBy: { createdAt: 'desc' },
      take: 50, // Limit to recent 50 products
    });

    console.log(`Found ${finishedProducts.length} finished products from OperatorProductUpdate`);

    // Get live products from Job table that should appear in Recently Finished Products
    const liveProducts = await prisma.job.findMany({
      where: {
        state: 'ON', // Only include products that are currently live
      },
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
      take: 50,
    });

    console.log(`Found ${liveProducts.length} live products from Job table`);

    // Get all jobs to determine the latest state for each product
    const allJobs = await prisma.job.findMany({
      orderBy: { createdAt: 'desc' },
      include: { product: true },
    });
    const latestJobStateByProduct: Record<string, string> = {};
    allJobs.forEach(job => {
      if (!latestJobStateByProduct[job.product.name]) {
        latestJobStateByProduct[job.product.name] = job.state;
      }
    });

    // Get the latest job for each productId
    const latestJobByProductId: Record<number, any> = {};
    allJobs.forEach(job => {
      const productId = job.productId;
      if (!latestJobByProductId[productId] || new Date(job.createdAt) > new Date(latestJobByProductId[productId].createdAt)) {
        latestJobByProductId[productId] = job;
      }
    });

    // Group by productId and get the latest update for each product
    const latestByProduct: { [productId: number]: any } = {};
    // Process finished products from OperatorProductUpdate
    finishedProducts.forEach((update: any) => {
      // Find the productId for this product name (case-insensitive, trimmed)
      const matchingJob = allJobs.find(j => j.product.name.trim().toLowerCase() === update.product.trim().toLowerCase());
      if (!matchingJob) return;
      const productId = matchingJob.productId;
      // Exclude if latest job is not ON
      if (!latestJobByProductId[productId] || latestJobByProductId[productId].state !== 'ON') return;
      if (
        !latestByProduct[productId] ||
        new Date(update.createdAt) > new Date(latestByProduct[productId].createdAt)
      ) {
        latestByProduct[productId] = {
          id: update.id,
          name: update.product,
          process: 'Finished',
          state: update.dispatchStatus || 'ON',
          date: new Date(update.createdAt).toLocaleDateString(),
          createdAt: update.createdAt,
          source: 'finished'
        };
      }
    });
    // Process live products from Job table - these should appear in Recently Finished Products
    liveProducts.forEach((job: any) => {
      const productId = job.productId;
      if (
        !latestByProduct[productId] ||
        new Date(job.createdAt) > new Date(latestByProduct[productId].createdAt)
      ) {
        latestByProduct[productId] = {
          id: `job_${job.id}`,
          name: job.product.name,
          process: job.machine.name,
          state: job.state,
          date: new Date(job.createdAt).toLocaleDateString(),
          createdAt: job.createdAt,
          source: 'live'
        };
      }
    });

    // Convert to the format expected by the dropdown
    const dispatchableFinishedProducts = Object.values(latestByProduct).map((product: any) => ({
      id: product.id,
      name: product.name,
      process: product.process,
      state: product.state,
      date: product.date,
      createdAt: product.createdAt,
      source: product.source, // Keep track of source for debugging
    }));

    // Sort by creation date (most recent first)
    dispatchableFinishedProducts.sort((a, b) => 
      new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
    );

    // Debug: Print all jobs and their latest state by productId
    console.log('--- ALL JOBS ---');
    allJobs.forEach(job => {
      console.log(`JobID: ${job.id}, ProductID: ${job.productId}, ProductName: ${job.product.name}, State: ${job.state}, CreatedAt: ${job.createdAt}`);
    });
    console.log('--- LATEST JOB BY PRODUCTID ---');
    Object.entries(latestJobByProductId).forEach(([pid, job]) => {
      console.log(`ProductID: ${pid}, ProductName: ${job.product.name}, LatestState: ${job.state}, CreatedAt: ${job.createdAt}`);
    });
    // Debug: Print all OperatorProductUpdate records
    console.log('--- ALL OperatorProductUpdate ---');
    finishedProducts.forEach(update => {
      console.log(`UpdateID: ${update.id}, Product: ${update.product}, Archived: ${update.archived}, CreatedAt: ${update.createdAt}`);
    });
    // Debug: Print final list of products being returned
    console.log('--- FINAL PRODUCTS FOR DROPDOWN ---');
    dispatchableFinishedProducts.forEach(product => {
      console.log(`Product: ${product.name}, Source: ${product.source}, State: ${product.state}, ID: ${product.id}`);
    });

    console.log(`Returning ${dispatchableFinishedProducts.length} total products for dropdown`);

    return NextResponse.json({ 
      finishedProducts: dispatchableFinishedProducts,
      count: dispatchableFinishedProducts.length 
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching finished products:', error);
    return NextResponse.json({ 
      error: 'Failed to fetch finished products',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 