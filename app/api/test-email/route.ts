import { NextRequest, NextResponse } from 'next/server';
import { sendEmail } from '../../lib/email'; // adjust path if needed

export async function GET() {
  try {
    await sendEmail({
      to: process.env.EMAIL_USER!, // send to yourself for testing
      subject: 'Test Email from TrackOpz',
      html: '<p>This is a test email sent from your Next.js app!</p>',
    });
    return NextResponse.json({ success: true, message: 'Email sent!' });
  } catch (error) {
    return NextResponse.json({ success: false, error: (error as Error).message }, { status: 500 });
  }
}