import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export async function POST(req: NextRequest) {
  const { phone, username, profileImage } = await req.json();

  if (!phone || !username || !profileImage) {
    return NextResponse.json({ success: false, error: 'All fields are required.' }, { status: 400 });
  }

  // Find operator by phone
  const operator = await prisma.operator.findUnique({ where: { phone } });
  if (!operator) {
    return NextResponse.json({ success: false, error: 'Operator not found.' }, { status: 404 });
  }

  // Update operator
  await prisma.operator.update({
    where: { phone },
    data: { username, profileImage },
  });

  return NextResponse.json({ success: true, message: 'Operator account setup complete.' });
} 