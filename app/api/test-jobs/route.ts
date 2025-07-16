import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

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

    // Group jobs by product name to see all states for each product
    const jobsByProduct: { [productName: string]: any[] } = {};
    jobs.forEach((job: any) => {
      const productName = job.product.name;
      if (!jobsByProduct[productName]) {
        jobsByProduct[productName] = [];
      }
      jobsByProduct[productName].push(job);
    });

    // Get the latest job for each product
    const latestJobs: { [productName: string]: any } = {};
    Object.keys(jobsByProduct).forEach((productName: string) => {
      const productJobs = jobsByProduct[productName];
      const latestJob = productJobs[0]; // Already sorted by createdAt desc
      latestJobs[productName] = latestJob;
    });

    return NextResponse.json({ 
      success: true,
      allJobs: jobs,
      jobsByProduct,
      latestJobs
    });
  } catch (error) {
    console.error('Error testing jobs:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
} 