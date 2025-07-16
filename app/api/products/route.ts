import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';


const prisma = new PrismaClient();

// GET: Fetch all products with their latest job (process/stage and status)
export async function GET() {
  try {
    // Get all products and their latest job
    const products = await prisma.product.findMany({
      include: {
        jobs: {
          orderBy: { createdAt: 'desc' },
          take: 1,
        },
      },
    });
    // Map to include process and status from latest job
    const result = products.map((p) => ({
      id: p.id,
      name: p.name,
      process: p.jobs[0]?.stage || '',
      status: p.jobs[0]?.state || '',
    }));
    return NextResponse.json({ products: result }, { status: 200 });
  } catch {
    return NextResponse.json({ error: 'Failed to fetch products' }, { status: 500 });
  }
} 