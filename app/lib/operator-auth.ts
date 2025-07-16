import { cookies } from 'next/headers';
import jwt from 'jsonwebtoken';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export async function getCurrentOperator() {
  const cookieStore = await cookies();
  const token = cookieStore.get('token')?.value;

  if (!token) {
    return null;
  }

  try {
    const decoded: any = jwt.verify(token, process.env.JWT_SECRET!);
    const phone = decoded.phone;
    
    if (!phone) {
      return null;
    }

    const operator = await prisma.operator.findUnique({
      where: { phone },
      select: { id: true, phone: true, username: true, profileImage: true }
    });

    return operator;
  } catch (error) {
    console.error('Error decoding operator token:', error);
    return null;
  }
}

export async function getCurrentOperatorId(): Promise<number | null> {
  const operator = await getCurrentOperator();
  return operator?.id || null;
} 