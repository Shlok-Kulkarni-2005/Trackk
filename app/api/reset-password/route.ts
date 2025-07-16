import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import bcrypt from 'bcryptjs';

const prisma = new PrismaClient();

export async function POST(req: NextRequest) {
  const { email, token, newPassword } = await req.json();

  // 1. Find user by email
  const user = await prisma.user.findUnique({ where: { email } });
  if (!user) {
    return NextResponse.json({ success: false, error: 'Invalid email or token' }, { status: 400 });
  }

  // 2. Find valid password reset token
  const resetToken = await prisma.passwordResetToken.findFirst({
    where: {
      userId: user.id,
      token,
      used: false,
      expiresAt: { gt: new Date() },
    },
  });

  if (!resetToken) {
    return NextResponse.json({ success: false, error: 'Invalid or expired token' }, { status: 400 });
  }

  // 3. Hash new password
  const hashedPassword = await bcrypt.hash(newPassword, 10);

  // 4. Update user's password
  await prisma.user.update({
    where: { id: user.id },
    data: { password: hashedPassword },
  });

  // 5. Mark token as used
  await prisma.passwordResetToken.update({
    where: { id: resetToken.id },
    data: { used: true },
  });

  // 6. Respond with success
  return NextResponse.json({ success: true, message: 'Password reset successful.' });
} 