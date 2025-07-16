import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import jwt from 'jsonwebtoken';

const prisma = new PrismaClient();

export async function POST(req: NextRequest) {
  const { email, otp } = await req.json();

  // 1. Find user by email
  const user = await prisma.user.findUnique({ where: { email } });
  if (!user) {
    return NextResponse.json({ success: false, error: 'Invalid email or OTP' }, { status: 401 });
  }

  // 2. Find valid OTP for user
  const otpRecord = await prisma.oTP.findFirst({
    where: {
      userId: user.id,
      code: otp,
      used: false,
      expiresAt: { gt: new Date() },
    },
  });

  if (!otpRecord) {
    return NextResponse.json({ success: false, error: 'Invalid or expired OTP' }, { status: 401 });
  }

  // 3. Mark OTP as used
  await prisma.oTP.update({
    where: { id: otpRecord.id },
    data: { used: true },
  });

  // 4. Generate JWT
  const token = jwt.sign(
    { userId: user.id, email: user.email },
    process.env.JWT_SECRET!,
    { expiresIn: '7d' }
  );

  // 5. Set JWT as HTTP-only cookie using NextResponse
  const response = NextResponse.json({ success: true, message: 'OTP verified successfully.' });
  response.cookies.set('token', token, {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax',
    path: '/',
    maxAge: 60 * 60 * 24 * 7, // 7 days
  });

  return response;
} 