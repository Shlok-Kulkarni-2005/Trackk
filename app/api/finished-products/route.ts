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

    // Aggregate jobs by productId for jobs with state 'OFF'
    const pastProductMap: { [productId: string]: { name: string, process: string, count: number, latestCreatedAt: Date } } = {};
    jobs.forEach((job: any) => {
      if (job.state === 'OFF') {
        const productId = job.product.id;
        if (!pastProductMap[productId]) {
          pastProductMap[productId] = {
            name: job.product.name,
            process: job.machine.name,
            count: 1,
            latestCreatedAt: new Date(job.createdAt),
          };
        } else {
          pastProductMap[productId].count += 1;
          // Update latestCreatedAt if this job is newer
          if (new Date(job.createdAt) > pastProductMap[productId].latestCreatedAt) {
            pastProductMap[productId].latestCreatedAt = new Date(job.createdAt);
            pastProductMap[productId].process = job.machine.name;
          }
        }
      }
    });

    // Aggregate jobs by productId for jobs with state 'ON'
    const liveProductMap: { [productId: string]: { name: string, process: string, count: number, latestCreatedAt: Date } } = {};
    jobs.forEach((job: any) => {
      if (job.state === 'ON') {
        const productId = job.product.id;
        if (!liveProductMap[productId]) {
          liveProductMap[productId] = {
            name: job.product.name,
            process: job.machine.name,
            count: 1,
            latestCreatedAt: new Date(job.createdAt),
          };
        } else {
          liveProductMap[productId].count += 1;
          // Update latestCreatedAt if this job is newer
          if (new Date(job.createdAt) > liveProductMap[productId].latestCreatedAt) {
            liveProductMap[productId].latestCreatedAt = new Date(job.createdAt);
            liveProductMap[productId].process = job.machine.name;
          }
        }
      }
    });

    // Prepare past products array
    const pastProducts = Object.entries(pastProductMap).map(([productId, data]) => ({
      id: productId,
      name: data.name,
      process: data.process,
      state: 'OFF',
      quantity: data.count,
      date: data.latestCreatedAt.toLocaleDateString(),
      createdAt: data.latestCreatedAt,
      type: 'past',
    }));

    // Prepare live products array
    const liveProducts = Object.entries(liveProductMap).map(([productId, data]) => ({
      id: productId,
      name: data.name,
      process: data.process,
      state: 'ON',
      quantity: data.count,
      date: data.latestCreatedAt.toLocaleDateString(),
      createdAt: data.latestCreatedAt,
      type: 'live',
    }));

    // Sort both lists by createdAt descending
    liveProducts.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
    pastProducts.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

    // Debug: Print all jobs
    console.log('--- ALL JOBS ---');
    jobs.forEach(job => {
      console.log(`JobID: ${job.id}, ProductID: ${job.product.id}, ProductName: ${job.product.name}, State: ${job.state}, CreatedAt: ${job.createdAt}`);
    });

    // Debug: Print computed liveProducts
    console.log('--- COMPUTED LIVE PRODUCTS ---');
    liveProducts.forEach(product => {
      console.log(`ProductID: ${product.id}, Name: ${product.name}, Quantity: ${product.quantity}, LatestTime: ${product.createdAt}`);
    });

    // Debug: Print computed pastProducts
    console.log('--- COMPUTED PAST PRODUCTS ---');
    pastProducts.forEach(product => {
      console.log(`ProductID: ${product.id}, Name: ${product.name}, Quantity: ${product.quantity}, LatestTime: ${product.createdAt}`);
    });

    return NextResponse.json({ liveProducts, pastProducts }, { status: 200 });
  } catch (error) {
    console.error('Error fetching products:', error);
    return NextResponse.json({ error: 'Failed to fetch products' }, { status: 500 });
  }
} 