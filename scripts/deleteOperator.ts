// scripts/deleteOperator.ts
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

async function main() {
  const phone = '+918956747898'; // Replace with the operator's number

  // Find the operator
  const operator = await prisma.operator.findUnique({ where: { phone } });
  if (!operator) {
    console.log('Operator not found');
    return;
  }

  // Delete all related OTPs
  await prisma.operatorOTP.deleteMany({
    where: { operatorId: operator.id },
  });

  // Now delete the operator
  const deleted = await prisma.operator.delete({
    where: { phone },
  });

  console.log('Deleted operator:', deleted);
}

main()
  .catch((e) => {
    console.error(e);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
