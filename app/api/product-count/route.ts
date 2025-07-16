import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';


const prisma = new PrismaClient();

// GET: Fetch all products with their count (number of jobs with state 'ON' on 'CNC Finished' machine)
export async function GET() {
  try {
    const cncMachine = await prisma.machine.findFirst({
      where: { name: 'CNC Finished' },
      select: { id: true },
    });

    // If CNC machine doesn't exist, all counts are 0
    if (!cncMachine) {
      const products = await prisma.product.findMany();
      return NextResponse.json(
        {
          products: products.map((p) => ({
            id: p.id,
            name: p.name,
            count: 0,
            status: '',
          })),
        },
        { status: 200 }
      );
    }

    const products = await prisma.product.findMany({
      include: {
        _count: {
          select: {
            jobs: {
              where: {
                state: 'ON',
                machineId: cncMachine.id,
              },
            },
          },
        },
        jobs: {
          orderBy: { createdAt: 'desc' },
          take: 1,
        },
      },
    });

    const result = products.map((p) => ({
      id: p.id,
      name: p.name,
      count: p._count.jobs,
      status: p.jobs[0]?.state || '',
    }));

    return NextResponse.json({ products: result }, { status: 200 });
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : 'An unknown error occurred';
    console.error('Failed to fetch product counts:', errorMessage);
    return NextResponse.json(
      { error: 'Failed to fetch product counts' },
      { status: 500 }
    );
  }
} 