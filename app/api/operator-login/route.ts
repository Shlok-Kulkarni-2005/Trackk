import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import twilio from 'twilio';

const prisma = new PrismaClient();

const twilioClient = twilio(
  process.env.TWILIO_ACCOUNT_SID!,
  process.env.TWILIO_AUTH_TOKEN!
);
const twilioPhone = process.env.TWILIO_PHONE_NUMBER!;

function generateOTP(length = 6) {
  return Math.floor(100000 + Math.random() * 900000).toString().slice(0, length);
}

export async function POST(req: NextRequest) {
  const { phone } = await req.json();
  if (!phone) {
    return NextResponse.json({ success: false, error: 'Phone number is required.' }, { status: 400 });
  }

  // Find or create operator
  let operator = await prisma.operator.findUnique({ where: { phone } });
  if (!operator) {
    operator = await prisma.operator.create({ data: { phone } });
  }

  // Generate OTP
  const otp = generateOTP();
  const expiresAt = new Date(Date.now() + 10 * 60 * 1000); // 10 minutes

  // Store OTP
  await prisma.operatorOTP.create({
    data: {
      operatorId: operator.id,
      code: otp,
      expiresAt,
    },
  });

  // Send OTP via Twilio
  try {
    await twilioClient.messages.create({
      body: `Your TrackOPZ operator login OTP is: ${otp}`,
      from: twilioPhone,
      to: phone,
    });
  } catch (err) {
    console.error(err);
    return NextResponse.json({ success: false, error: 'Failed to send OTP.' }, { status: 500 });
  }

  // Always respond with a generic message
  return NextResponse.json({ success: true, message: 'If the phone number is valid, an OTP has been sent.' });
} 