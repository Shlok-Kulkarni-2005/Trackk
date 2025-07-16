import { NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export async function GET() {
  try {
    const machines = await prisma.machine.findMany({ orderBy: { name: 'asc' } });
    return NextResponse.json({ machines }, { status: 200 });
  } catch {
    return NextResponse.json({ error: 'Failed to fetch machines' }, { status: 500 });
  }
} 