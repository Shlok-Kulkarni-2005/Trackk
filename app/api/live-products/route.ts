import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch live products (state ON) for the updatedetailsop page
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

    // Filter only live products (state ON)
    const liveProducts = Object.values(latestJobByProduct)
      .filter((job: any) => job.state === 'ON')
      .map((job: any) => ({
        id: job.product.id,
        name: job.product.name,
        process: job.machine.name,
        state: job.state,
        date: new Date(job.createdAt).toLocaleDateString(),
        createdAt: job.createdAt,
      }));

    return NextResponse.json({ liveProducts }, { status: 200 });
  } catch (error) {
    console.error('Error fetching live products:', error);
    return NextResponse.json({ error: 'Failed to fetch live products' }, { status: 500 });
  }
} 