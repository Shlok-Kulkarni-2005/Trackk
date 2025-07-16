import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch both live products and finished products for the dropdown
export async function GET() {
  try {
    // Get all jobs with machine and product info
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

    // Only include products whose latest job is NOT OFF
    const liveProducts = Object.values(latestJobByProduct)
      .filter((job: any) => job.state !== 'OFF')
      .map((job: any) => ({
        id: `job_${job.id}`,
        name: job.product.name,
        process: job.machine.name,
        state: job.state,
        date: new Date(job.createdAt).toLocaleDateString(),
        createdAt: job.createdAt,
        type: 'live'
      }));

    // Get finished products from OperatorProductUpdate table first
    const finishedProducts = await prisma.operatorProductUpdate.findMany({
      orderBy: { createdAt: 'desc' },
      take: 20, // Limit to recent 20 products
    });

    // Include all live products (don't filter out processed ones)
    const filteredLiveProducts = liveProducts;

    // Group by product name and get the latest update for each product
    const latestByProduct: { [productName: string]: any } = {};
    finishedProducts.forEach((update: any) => {
      const productName = update.product;
      if (
        !latestByProduct[productName] ||
        new Date(update.createdAt) > new Date(latestByProduct[productName].createdAt)
      ) {
        latestByProduct[productName] = update;
      }
    });

    const finishedProductsList = Object.values(latestByProduct).map((update: any) => ({
      id: update.id, // Use the update record ID
      name: update.product,
      process: 'Finished',
      state: update.dispatchStatus,
      date: new Date(update.createdAt).toLocaleDateString(),
      createdAt: update.createdAt,
      type: 'finished'
    }));

    // Combine both lists and sort by creation date
    const allProducts = [...filteredLiveProducts, ...finishedProductsList]
      .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

    return NextResponse.json({ liveProducts: allProducts }, { status: 200 });
  } catch (error) {
    console.error('Error fetching products:', error);
    return NextResponse.json({ error: 'Failed to fetch products' }, { status: 500 });
  }
} 