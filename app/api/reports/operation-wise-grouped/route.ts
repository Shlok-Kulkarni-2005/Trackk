import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

function formatTimeHHMM(dateStr: string) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);
  const operationType = searchParams.get('process');
  const startDateParam = searchParams.get('startDate');
  const endDateParam = searchParams.get('endDate');

  if (!operationType || !startDateParam || !endDateParam) {
    return NextResponse.json({ success: false, error: 'Missing required parameters.' }, { status: 400 });
  }

  const startDate = new Date(startDateParam);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(endDateParam);
  endDate.setHours(23, 59, 59, 999);

  try {
    // Get all jobs for the operation type and date range
    const jobs = await prisma.job.findMany({
      where: {
        machine: { name: { startsWith: operationType } },
        createdAt: { gte: startDate, lte: endDate },
      },
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'asc' },
    });
    // Group jobs by product, machine, and ON/OFF pairs
    const jobGroups: Record<string, any[]> = {};
    jobs.forEach(job => {
      const key = `${job.productId}__${job.machineId}`;
      if (!jobGroups[key]) jobGroups[key] = [];
      jobGroups[key].push(job);
    });
    // Collect all report rows (ON/OFF pairs and unmatched ON jobs)
    const allRows: any[] = [];
    Object.values(jobGroups).forEach((group: any) => {
      group.sort((a: any, b: any) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
      let onJobs: any[] = [];
      let offJobs: any[] = [];
      for (const job of group) {
        if (job.state === 'ON') {
          onJobs.push(job);
        } else if (job.state === 'OFF') {
          offJobs.push(job);
        }
      }
      let i = 0;
      while (i < Math.min(onJobs.length, offJobs.length)) {
        const onJob = onJobs[i];
        const offJob = offJobs[i];
        const onTime = new Date(onJob.createdAt);
        const offTime = new Date(offJob.updatedAt || offJob.createdAt);
        const dateStr = onTime.toISOString().split('T')[0];
        const totalTime = Math.round((offTime.getTime() - onTime.getTime()) / 60000);
        let machineNumber = '';
        const match = onJob.machine.name.match(/#(.*)$/);
        if (match) machineNumber = match[1].trim();
        allRows.push({
          productId: onJob.product.name,
          machineNumber,
          date: dateStr,
          onTime: formatTimeHHMM(onJob.createdAt),
          offTime: formatTimeHHMM(offJob.updatedAt || offJob.createdAt),
          totalTime: totalTime ? totalTime : '',
          quantity: 1
        });
        i++;
      }
      for (let j = i; j < onJobs.length; j++) {
        const onJob = onJobs[j];
        const onTime = new Date(onJob.createdAt);
        const dateStr = onTime.toISOString().split('T')[0];
        let machineNumber = '';
        const match = onJob.machine.name.match(/#(.*)$/);
        if (match) machineNumber = match[1].trim();
        allRows.push({
          productId: onJob.product.name,
          machineNumber,
          date: dateStr,
          onTime: formatTimeHHMM(onJob.createdAt),
          offTime: '',
          totalTime: '',
          quantity: 1
        });
      }
    });
    // Group rows by all parameters except quantity
    const groupedRows = new Map();
    for (const row of allRows) {
      const rowKey = [
        row.productId,
        row.machineNumber,
        row.date,
        row.onTime,
        row.offTime,
        row.totalTime
      ].join('|');
      if (!groupedRows.has(rowKey)) {
        groupedRows.set(rowKey, { ...row });
      } else {
        groupedRows.get(rowKey).quantity += row.quantity;
      }
    }
    const result = Array.from(groupedRows.values());
    return NextResponse.json({ success: true, data: result });
  } catch (error) {
    console.error('Error generating grouped report:', error);
    return NextResponse.json({ success: false, error: 'Internal server error.' }, { status: 500 });
  }
} 