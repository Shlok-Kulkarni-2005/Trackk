import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import ExcelJS from 'exceljs';

const prisma = new PrismaClient();

function formatTimeHHMM(dateStr: string) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function getMachineNumber(machineName: string) {
  if (!machineName) return '';
  const match = machineName.match(/#(.*)$/);
  if (match) return match[1].trim();
  return '';
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
    // Get all jobs for the date range (do not filter by machine in the DB query)
    const jobs = await prisma.job.findMany({
      where: {
        createdAt: { gte: startDate, lte: endDate },
      },
      include: {
        machine: true,
        product: {
          include: {
            jobs: {
              orderBy: { createdAt: 'asc' },
              take: 1,
            },
          },
        },
      },
      orderBy: { createdAt: 'asc' },
    });
    // Normalize function for machine names
    function normalizeMachineName(name) {
      return name.split('#')[0].replace(/[-\s]+/g, ' ').trim().toLowerCase();
    }
    const normalizedOperationType = normalizeMachineName(operationType);
    // Filter jobs to only those whose machine base name matches the selected operation
    const jobsForOperation = jobs.filter(job => {
      if (!job.machine?.name) return false;
      return normalizeMachineName(job.machine.name) === normalizedOperationType;
    });
    // Remove the filter for product's first job created today
    // Use all jobs for the selected machine and date range
    const filteredJobs = jobsForOperation;
    // Group jobs by product name and normalized machine name
    const jobGroups: Record<string, any[]> = {};
    filteredJobs.forEach(job => {
      const normalizedProductName = job.product?.name?.trim().toLowerCase() || '';
      const normalizedMachineName = normalizeMachineName(job.machine?.name || '');
      const key = `${normalizedProductName}__${normalizedMachineName}`;
      if (!jobGroups[key]) jobGroups[key] = [];
      jobGroups[key].push(job);
    });
    // For every job with state OFF, use createdAt as ON Time and updatedAt as OFF Time, and show total time in minutes only
    const allRows: any[] = [];
    // Helper to format time as HH:MM (ignore seconds)
    function formatTimeHHMMOnly(date) {
      if (!date) return '';
      const d = new Date(date);
      const pad = (n) => n.toString().padStart(2, '0');
      return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }

    // When building allRows, use formatTimeHHMMOnly for onTime and offTime
    Object.values(jobGroups).forEach((group: any) => {
      group.sort((a: any, b: any) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
      group.forEach((job: any) => {
        if (job.state === 'OFF') {
          const normalizedProductName = job.product?.name?.trim().toLowerCase() || '';
          const normalizedMachineName = normalizeMachineName(job.machine?.name || '');
          const dateStr = new Date(job.createdAt).toISOString().split('T')[0];
          const machineNumber = getMachineNumber(job.machine.name);
          const totalMinutes = Math.round((new Date(job.updatedAt || job.createdAt).getTime() - new Date(job.createdAt).getTime()) / 60000) || 0;
          allRows.push({
            productId: job.product.name,
            quantity: job.quantity || 1,
            machineNumber,
            date: dateStr,
            onTime: formatTimeHHMMOnly(job.createdAt),
            offTime: formatTimeHHMMOnly(job.updatedAt || job.createdAt),
            totalTime: totalMinutes,
          });
        }
      });
    });
    // Group rows by all parameters except quantity
    // When grouping, use the new onTime and offTime (HH:MM only)
    const groupedRows = new Map();
    for (const row of allRows) {
      const rowKey = [
        row.productId,
        row.machineNumber,
        row.date
      ].join('|');
      if (!groupedRows.has(rowKey)) {
        groupedRows.set(rowKey, {
          ...row,
          quantity: row.quantity,
          totalTime: row.totalTime,
          onTimes: [row.onTime],
          offTimes: [row.offTime]
        });
      } else {
        const group = groupedRows.get(rowKey);
        group.quantity += row.quantity;
        group.totalTime += row.totalTime;
        if (row.onTime) group.onTimes.push(row.onTime);
        if (row.offTime) group.offTimes.push(row.offTime);
      }
    }
    // Prepare final result: show earliest ON Time and latest OFF Time for each group, and compute total time as their difference
    const result = Array.from(groupedRows.values()).map(group => {
      // Find earliest ON and latest OFF as Date objects
      const onTimes = group.onTimes.filter(Boolean).map(t => {
        const [h, m] = t.split(':');
        return new Date(`${group.date}T${h.padStart(2, '0')}:${m.padStart(2, '0')}:00`);
      });
      const offTimes = group.offTimes.filter(Boolean).map(t => {
        const [h, m] = t.split(':');
        return new Date(`${group.date}T${h.padStart(2, '0')}:${m.padStart(2, '0')}:00`);
      });
      const earliestOn = onTimes.length ? new Date(Math.min(...onTimes.map(d => d.getTime()))) : null;
      const latestOff = offTimes.length ? new Date(Math.max(...offTimes.map(d => d.getTime()))) : null;
      let totalTime = '';
      if (earliestOn && latestOff && latestOff > earliestOn) {
        totalTime = Math.round((latestOff.getTime() - earliestOn.getTime()) / 60000);
      }
      return {
        ...group,
        onTime: group.onTimes.filter(Boolean).sort()[0] || '',
        offTime: group.offTimes.filter(Boolean).sort().reverse()[0] || '',
        totalTime,
      };
    });

    // Helper to format time as HH:MM:SS
    function formatTimeHMS(dateOrString: Date | string | undefined) {
      if (!dateOrString) return '';
      const date = new Date(dateOrString);
      const pad = (n: number) => n.toString().padStart(2, '0');
      return `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    // Helper to format total time as HH:MM:SS
    function formatTotalTimeHMS(totalMinutes: number) {
      if (!totalMinutes || isNaN(totalMinutes) || totalMinutes <= 0) return '';
      const totalSeconds = Math.round(totalMinutes * 60);
      const hours = Math.floor(totalSeconds / 3600);
      const minutes = Math.floor((totalSeconds % 3600) / 60);
      const seconds = totalSeconds % 60;
      return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
    }

    // Create Excel
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Process Wise Report');
    worksheet.addRow([
      'Product ID', 'Machine Number', 'Date', 'ON Time', 'OFF Time', 'Total Time (min)', 'Quantity'
    ]);
    result.forEach(row => {
      worksheet.addRow([
        row.productId,
        row.machineNumber,
        row.date,
        row.onTime,
        row.offTime,
        row.totalTime,
        row.quantity
      ]);
    });
    worksheet.columns.forEach(column => {
      column.width = 18;
    });
    const buffer = await workbook.xlsx.writeBuffer();
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="process-wise-report_${operationType}_${startDateParam}_${endDateParam}.xlsx"`,
      },
    });
  } catch (error) {
    console.error('Error generating process wise report:', error);
    return NextResponse.json({ success: false, error: 'Internal server error.' }, { status: 500 });
  }
} 