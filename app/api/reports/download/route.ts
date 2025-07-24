import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';
import ExcelJS from 'exceljs';

const prisma = new PrismaClient();

// Helper to get date ranges
const getDateRange = (reportType: string, startDateParam?: string, endDateParam?: string) => {
  const now = new Date();
  let startDate = new Date(now);
  let endDate = new Date(now);

  if (startDateParam && endDateParam) {
    // Use provided date range
    startDate = new Date(startDateParam);
    endDate = new Date(endDateParam);
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);
  } else if (reportType === 'daily') {
    // For Date Wise - use today's date range
    startDate.setHours(0, 0, 0, 0);
  } else if (reportType === 'weekly') {
    // For Weekly - use current week
    startDate.setDate(now.getDate() - now.getDay());
    startDate.setHours(0, 0, 0, 0);
  } else if (reportType === 'monthly') {
    // For Monthly - use current month
    startDate.setDate(1);
    startDate.setHours(0, 0, 0, 0);
  }
  endDate.setHours(23, 59, 59, 999);
  return { startDate, endDate };
};

function formatDateTimeNoMs(date) {
  // Format as 'YYYY-MM-DD HH:mm:ss'
  const pad = (n) => n.toString().padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function getIsoDateTimeToSecond(date) {
  // Returns 'YYYY-MM-DDTHH:mm:ss' (ISO, no ms)
  const pad = (n) => n.toString().padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function formatTimeHHMMNoSeconds(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  const pad = n => n.toString().padStart(2, '0');
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}
function getMachineNumber(machineName) {
  if (!machineName) return '';
  const match = machineName.match(/#(\d+)/);
  if (match) return match[1];
  // fallback: if no #number, try to extract last number
  const numMatch = machineName.match(/(\d+)$/);
  if (numMatch) return numMatch[1];
  return '';
}

// Main GET handler
export async function GET(req: NextRequest) {
  // --- DEBUG LOGGING: API route is being hit ---
  console.log('DEBUG: /api/reports/download called');
  const { searchParams } = new URL(req.url);
  const reportType = searchParams.get('reportType') || 'daily';
  const startDateParam = searchParams.get('startDate') || undefined;
  const endDateParam = searchParams.get('endDate') || undefined;

  try {
    const { startDate, endDate } = getDateRange(reportType, startDateParam, endDateParam);
    
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (reportType === 'processWise') {
      // Operation/machine-wise report
      const operationType = searchParams.get('process');
      if (!operationType) {
        return NextResponse.json({ success: false, error: 'Operation/Machine is required for process-wise report.' }, { status: 400 });
      }
      worksheet = workbook.addWorksheet('Operation Wise Report');
      // Ensure columns match screenshot order and names
      worksheet.columns = [
        { header: 'Product ID', key: 'productId', width: 20 },
        { header: 'Quantity', key: 'quantity', width: 12 },
        { header: 'Machine Number', key: 'machineNumber', width: 18 },
        { header: 'Date', key: 'date', width: 15 },
        { header: 'ON Time', key: 'onTime', width: 10 },
        { header: 'OFF Time', key: 'offTime', width: 10 },
        { header: 'Total Time (min)', key: 'totalTime', width: 20 },
      ];
      // --- DEBUG LOGGING ---
      // Log all ON and OFF jobs for the selected process/date
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
      const allOn = jobs.filter(j => j.state === 'ON');
      const allOff = jobs.filter(j => j.state === 'OFF');
      console.log('DEBUG: ALL ON JOBS:', allOn.map(j => ({ id: j.id, productId: j.productId, machineId: j.machineId, createdAt: j.createdAt, quantity: j.quantity })));
      console.log('DEBUG: ALL OFF JOBS:', allOff.map(j => ({ id: j.id, productId: j.productId, machineId: j.machineId, createdAt: j.createdAt, updatedAt: j.updatedAt, quantity: j.quantity })));
      // --- END DEBUG LOGGING ---
      // ---
      // All time formatting for ON/OFF is handled below. Always output as HH:mm (no seconds, no ms)
      // ---
      // Get all jobs for the operation type and date range
      const jobGroups = {};
      jobs.forEach(job => {
        // Only skip if product or machine is truly missing
        if (!job.product || !job.machine) {
          console.warn('Skipping job with missing product or machine:', job.id);
          return;
        }
        // Group by productId and machineId
        const key = `${job.productId}__${job.machineId}`;
        if (!jobGroups[key]) jobGroups[key] = [];
        jobGroups[key].push(job);
      });
      const allRows = [];
      Object.values(jobGroups).forEach((group) => {
        group.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
        // Separate ON and OFF jobs
        let onJobs = group.filter(j => j.state === 'ON').map(j => ({...j, remaining: j.quantity}));
        let offJobs = group.filter(j => j.state === 'OFF').map(j => ({...j, remaining: j.quantity}));
        let onIdx = 0, offIdx = 0;
        while (onIdx < onJobs.length && offIdx < offJobs.length) {
          let onJob = onJobs[onIdx];
          let offJob = offJobs[offIdx];
          let pairQty = Math.min(onJob.remaining, offJob.remaining);
          if (pairQty > 0) {
            // Log each ON/OFF pair
            console.log('DEBUG: PAIRING ON/OFF', {
              onId: onJob.id, offId: offJob.id, pairQty,
              onCreatedAt: onJob.createdAt, offCreatedAt: offJob.createdAt, offUpdatedAt: offJob.updatedAt
            });
            // Use real ON and OFF times
            const onTime = new Date(onJob.createdAt);
            onTime.setSeconds(0, 0);
            const offTime = new Date(offJob.updatedAt || offJob.createdAt);
            offTime.setSeconds(0, 0);
            const dateStr = onTime.toISOString().split('T')[0];
            // Always use getMachineNumber
            const machineNumber = getMachineNumber(onJob.machine.name);
            const onTimeStr = `${onTime.getHours().toString().padStart(2, '0')}:${onTime.getMinutes().toString().padStart(2, '0')}`;
            const offTimeStr = `${offTime.getHours().toString().padStart(2, '0')}:${offTime.getMinutes().toString().padStart(2, '0')}`;
            const totalTime = Math.round((offTime.getTime() - onTime.getTime()) / 60000);
            allRows.push({
              productId: onJob.product.name || onJob.productId,
              quantity: pairQty,
              machineNumber,
              date: dateStr,
              onTime: onTimeStr,
              offTime: offTimeStr,
              totalTime: totalTime ? totalTime : '',
            });
            onJob.remaining -= pairQty;
            offJob.remaining -= pairQty;
          }
          // Always increment at least one index to avoid infinite loop
          if (onJob.remaining === 0 && offJob.remaining === 0) {
            onIdx++;
            offIdx++;
          } else if (onJob.remaining === 0) {
            onIdx++;
          } else if (offJob.remaining === 0) {
            offIdx++;
          } else {
            break;
          }
        }
        // Any remaining ON jobs (not yet OFF)
        for (; onIdx < onJobs.length; onIdx++) {
          let onJob = onJobs[onIdx];
          if (onJob.remaining > 0) {
            // Log unmatched ON job
            console.log('DEBUG: UNMATCHED ON JOB', {
              onId: onJob.id, remaining: onJob.remaining, onCreatedAt: onJob.createdAt
            });
            const onTime = new Date(onJob.createdAt);
            onTime.setSeconds(0, 0);
            const dateStr = onTime.toISOString().split('T')[0];
            const machineNumber = getMachineNumber(onJob.machine.name);
            const onTimeStr = `${onTime.getHours().toString().padStart(2, '0')}:${onTime.getMinutes().toString().padStart(2, '0')}`;
            // Fallback: unmatched ON job, blank OFF/total time
            allRows.push({
              productId: onJob.product.name || onJob.productId,
              quantity: onJob.remaining,
              machineNumber,
              date: dateStr,
              onTime: onTimeStr,
              offTime: '',
              totalTime: '',
            });
          }
        }
      });
      // Group rows by all parameters except quantity
      const groupedRowsObj = {};
      for (const row of allRows) {
        const groupKey = [
          row.productId,
          row.machineNumber,
          row.date,
          row.onTime,
          row.offTime,
          row.totalTime
        ].join('|');
        if (!groupedRowsObj[groupKey]) {
          groupedRowsObj[groupKey] = { ...row };
        } else {
          groupedRowsObj[groupKey].quantity += row.quantity;
        }
      }
      // Add grouped rows to worksheet
      let addedRows = 0;
      for (const key in groupedRowsObj) {
        worksheet.addRow(groupedRowsObj[key]);
        addedRows++;
      }
      // Fallback: if no rows were added but jobs exist, add a row for each job with available info
      if (addedRows === 0 && jobs.length > 0) {
        jobs.forEach(job => {
          // Always use getMachineNumber for fallback
          const onTimeStr = job.createdAt ? `${new Date(job.createdAt).getHours().toString().padStart(2, '0')}:${new Date(job.createdAt).getMinutes().toString().padStart(2, '0')}` : '';
          const offTimeStr = job.updatedAt ? `${new Date(job.updatedAt).getHours().toString().padStart(2, '0')}:${new Date(job.updatedAt).getMinutes().toString().padStart(2, '0')}` : '';
          worksheet.addRow({
            productId: (job.product && job.product.name) || job.productId || 'Unknown',
            quantity: job.quantity || 1,
            machineNumber: getMachineNumber(job.machine && job.machine.name),
            date: job.createdAt ? new Date(job.createdAt).toISOString().split('T')[0] : '',
            onTime: onTimeStr,
            offTime: offTimeStr,
            totalTime: '',
          });
        });
        addedRows = jobs.length;
      }
      if (addedRows === 0) {
        worksheet.addRow({ productId: 'No data found for the selected criteria.' });
      }
      worksheet.addRow([]);
      // Use Object.values for total product calculation
      const totalQuantity = Object.values(groupedRowsObj).reduce((sum, row) => sum + row.quantity, 0);
      const summaryRow = worksheet.addRow(['Total Products:', totalQuantity, '', '', '', '', '']);
      summaryRow.getCell('A').font = { bold: true };
      summaryRow.getCell('B').font = { bold: true };
      const startDateStr = startDateParam || startDate.toISOString().split('T')[0];
      const endDateStr = endDateParam || endDate.toISOString().split('T')[0];
      const reportName = `${operationType}_${startDateStr}_${endDateStr}_report`;
      await prisma.reportDownload.create({ data: { reportName } });
      const buffer = await workbook.xlsx.writeBuffer();
      const headers = new Headers({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${reportName.replace(/ /g, '_')}.xlsx"`,
      });
      return new NextResponse(buffer, { status: 200, headers });
    }

    // Only get dispatched products (where dispatch button was clicked)
    const whereClause = { 
      createdAt: { gte: startDate, lte: endDate },
      dispatchStatus: 'Pending' // Only products that were dispatched
    };

    worksheet = workbook.addWorksheet('Dispatched Products Report');
    
    // Get all dispatched products with their quantities and dates
    const dispatchedProducts = await prisma.operatorProductUpdate.findMany({ 
      where: whereClause, 
      orderBy: { createdAt: 'desc' },
      select: {
        id: true,
        product: true,
        quantity: true,
        createdAt: true,
        processSteps: true
      }
    });

    if (!dispatchedProducts || dispatchedProducts.length === 0) {
      return NextResponse.json({ success: false, error: 'No dispatched products found for the selected date range.' }, { status: 404 });
    }

    // Set up columns for the report
    worksheet.columns = [
      { header: 'Product', key: 'product', width: 30 },
      { header: 'Quantity', key: 'quantity', width: 15 },
      { header: 'Date', key: 'date', width: 25, style: { numFmt: 'yyyy-mm-dd hh:mm:ss' } },
    ];

    // Add data rows
    dispatchedProducts.forEach(product => {
      worksheet.addRow({
        product: product.product,
        quantity: product.quantity,
        date: product.createdAt
      });
    });

    // Add summary row
    worksheet.addRow([]);
    const totalQuantity = dispatchedProducts.reduce((sum, product) => sum + product.quantity, 0);
    const summaryRow = worksheet.addRow(['Total Products Dispatched:', totalQuantity, '']);
    summaryRow.getCell('A').font = { bold: true };
    summaryRow.getCell('B').font = { bold: true };

    // Log and Send File
    const reportName = `${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Dispatched Products Report`;
    await prisma.reportDownload.create({ data: { reportName } });

    const buffer = await workbook.xlsx.writeBuffer();
    const headers = new Headers({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${reportName.replace(/ /g, '_')}.xlsx"`,
    });

    return new NextResponse(buffer, { status: 200, headers });
  } catch (error) {
    console.error('Error generating report:', error);
    const errorMessage = error instanceof Error ? error.message : 'An unknown error occurred';
    return NextResponse.json({ success: false, error: errorMessage }, { status: 500 });
  }
} 