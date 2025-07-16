import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import { sendEmail } from '../../lib/email'; // adjust path if needed
import bcrypt from 'bcryptjs';

const prisma = new PrismaClient();

function generateOTP(length = 6) {
  return Math.floor(100000 + Math.random() * 900000).toString().slice(0, length);
}

export async function POST(req: NextRequest) {
  const { email, password } = await req.json();

  // 1. Find user by email
  const user = await prisma.user.findUnique({ where: { email } });
  if (!user) {
    return NextResponse.json({ success: false, error: 'Invalid credentials' }, { status: 401 });
  }

  // 2. Check password
  const isMatch = await bcrypt.compare(password, user.password);
  if (!isMatch) {
    return NextResponse.json({ success: false, error: 'Invalid credentials' }, { status: 401 });
  }

  // 3. Generate OTP
  const otp = generateOTP();
  const expiresAt = new Date(Date.now() + 10 * 60 * 1000); // 10 minutes from now

  // 4. Save OTP to DB
  await prisma.oTP.create({
    data: {
      userId: user.id,
      code: otp,
      expiresAt,
    },
  });

  // 5. Send OTP email
  await sendEmail({
    to: user.email,
    subject: 'Your OTP Code',
    html: `<p>Your OTP is: <b>${otp}</b></p>`,
  });

  return NextResponse.json({ success: true, message: 'OTP sent to your email.' });
}