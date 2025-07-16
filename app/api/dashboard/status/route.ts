import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Get comprehensive dashboard status
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
    jobs.forEach((job) => {
      const productId = job.product.id;
      if (
        !latestJobByProduct[productId] ||
        new Date(job.createdAt) > new Date(latestJobByProduct[productId].createdAt)
      ) {
        latestJobByProduct[productId] = job;
      }
    });

    // Categorize products
    const liveProducts = Object.values(latestJobByProduct)
      .filter((job) => job.state === 'ON')
      .map((job) => ({
        id: job.product.id.toString(),
        name: job.product.name,
        process: job.machine.name,
        status: job.state,
        createdAt: job.createdAt,
        jobId: job.id,
        isDispatchable: true,
      }));

    const pastProducts = Object.values(latestJobByProduct)
      .filter((job) => job.state === 'OFF')
      .map((job) => ({
        id: job.product.id.toString(),
        name: job.product.name,
        process: job.machine.name,
        status: job.state,
        createdAt: job.createdAt,
        jobId: job.id,
        isDispatchable: false,
      }));

    // Get machine status
    const machines = await prisma.machine.findMany({
      include: {
        jobs: {
          orderBy: { createdAt: 'desc' },
          take: 1,
          include: {
            product: true,
          },
        },
      },
    });

    const machineStatus = machines.map(machine => ({
      id: machine.id,
      name: machine.name,
      status: machine.status,
      currentProduct: machine.jobs[0]?.product?.name || null,
      lastActivity: machine.jobs[0]?.createdAt || null,
    }));

    // Calculate statistics
    const stats = {
      totalProducts: Object.keys(latestJobByProduct).length,
      liveProducts: liveProducts.length,
      pastProducts: pastProducts.length,
      dispatchableProducts: liveProducts.length,
      totalMachines: machines.length,
      activeMachines: machines.filter(m => m.status === 'ON').length,
      inactiveMachines: machines.filter(m => m.status === 'OFF').length,
    };

    // Get recent activity (last 10 jobs)
    const recentActivity = jobs.slice(0, 10).map(job => ({
      id: job.id,
      productName: job.product.name,
      machineName: job.machine.name,
      state: job.state,
      stage: job.stage,
      timestamp: job.createdAt,
      action: job.state === 'ON' ? 'Started' : 'Completed',
    }));

    return NextResponse.json({
      success: true,
      dashboard: {
        stats,
        liveProducts,
        pastProducts,
        machineStatus,
        recentActivity,
        lastUpdated: new Date().toISOString(),
      }
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching dashboard status:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to fetch dashboard status' 
    }, { status: 500 });
  }
} 