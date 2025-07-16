import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import jwt from 'jsonwebtoken';

const prisma = new PrismaClient();

export async function POST(req: NextRequest) {
  const { phone, otp } = await req.json();

  // 1. Find operator by phone
  const operator = await prisma.operator.findUnique({ where: { phone } });
  if (!operator) {
    return NextResponse.json({ success: false, error: 'Invalid phone or OTP' }, { status: 401 });
  }

  // 2. Find valid OTP for operator
  const otpRecord = await prisma.operatorOTP.findFirst({
    where: {
      operatorId: operator.id,
      code: otp,
      used: false,
      expiresAt: { gt: new Date() },
    },
  });

  if (!otpRecord) {
    return NextResponse.json({ success: false, error: 'Invalid or expired OTP' }, { status: 401 });
  }

  // 3. Mark OTP as used
  await prisma.operatorOTP.update({
    where: { id: otpRecord.id },
    data: { used: true },
  });

  // 4. Generate JWT
  const token = jwt.sign(
    { phone: operator.phone },
    process.env.JWT_SECRET!,
    { expiresIn: '7d' }
  );

  // 5. Set JWT as HTTP-only cookie
  const response = NextResponse.json({
    success: true,
    firstTime: !(operator.username && operator.profileImage),
    username: operator.username,
    profileImage: operator.profileImage,
  });
  response.cookies.set('token', token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    path: '/',
    maxAge: 60 * 60 * 24 * 7, // 7 days
  });

  return response;
} 