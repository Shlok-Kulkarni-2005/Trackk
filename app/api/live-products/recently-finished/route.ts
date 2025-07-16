import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch recently finished products for the dropdown (only live products)
export async function GET() {
  try {
    console.log('Fetching recently finished products for dropdown (live products only)...');
    
    // Get all jobs to determine the latest state for each product
    const allJobs = await prisma.job.findMany({
      orderBy: { createdAt: 'desc' },
      include: { 
        product: true,
        machine: true 
      },
    });

    // Get the latest job for each productId
    const latestJobByProductId: Record<number, any> = {};
    allJobs.forEach(job => {
      const productId = job.productId;
      if (!latestJobByProductId[productId] || new Date(job.createdAt) > new Date(latestJobByProductId[productId].createdAt)) {
        latestJobByProductId[productId] = job;
      }
    });

    // Filter only products that are currently live (state 'ON')
    const liveProducts = Object.values(latestJobByProductId)
      .filter((job: any) => job.state === 'ON')
      .map((job: any) => ({
        id: job.product.id,
        name: job.product.name,
        process: job.machine.name,
        state: job.state,
        date: new Date(job.createdAt).toLocaleDateString(),
        createdAt: job.createdAt,
        machineName: job.machine.name,
        productName: job.product.name
      }));

    // Sort by creation date (most recent first)
    liveProducts.sort((a, b) => 
      new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
    );

    // Format the display name as requested: "ProductName - MachineName (Date)"
    const formattedProducts = liveProducts.map((product: any) => ({
      ...product,
      displayName: `${product.productName} - ${product.machineName} (${product.date})`,
      // Add availability count (for now, showing 3 as default, can be enhanced later)
      availableCount: 3
    }));

    console.log(`Returning ${formattedProducts.length} live products for recently finished dropdown`);

    return NextResponse.json({ 
      recentlyFinishedProducts: formattedProducts,
      count: formattedProducts.length 
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching recently finished products:', error);
    return NextResponse.json({ 
      error: 'Failed to fetch recently finished products',
      details: error instanceof Error ? error.message : 'Unknown error'
    }, { status: 500 });
  }
} 