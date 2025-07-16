const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function fixUpdatedAt() {
  try {
    console.log('Starting to fix updatedAt field for existing jobs...');
    
    // Use raw SQL to update all jobs that don't have updatedAt set
    const result = await prisma.$executeRaw`
      UPDATE "Job" 
      SET "updatedAt" = "createdAt" 
      WHERE "updatedAt" IS NULL
    `;
    
    console.log(`Updated ${result} jobs with updatedAt field!`);
    
    // Verify the update worked by checking a few jobs
    const sampleJobs = await prisma.$queryRaw`
      SELECT id, "createdAt", "updatedAt" 
      FROM "Job" 
      LIMIT 5
    `;
    
    console.log('Sample jobs after update:');
    sampleJobs.forEach(job => {
      console.log(`Job ${job.id}: createdAt=${job.createdAt}, updatedAt=${job.updatedAt}`);
    });
    
  } catch (error) {
    console.error('Error fixing updatedAt field:', error);
  } finally {
    await prisma.$disconnect();
  }
}

fixUpdatedAt(); 