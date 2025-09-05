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

// function formatDateTimeNoMs(date) {
//   // Format as 'YYYY-MM-DD HH:mm:ss'
//   const pad = (n) => n.toString().padStart(2, '0');
//   return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
// }

// function getIsoDateTimeToSecond(date) {
//   // Returns 'YYYY-MM-DDTHH:mm:ss' (ISO, no ms)
//   const pad = (n) => n.toString().padStart(2, '0');
//   return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
// }

// function formatTimeHHMMNoSeconds(dateStr) {
//   if (!dateStr) return '';
//   const date = new Date(dateStr);
//   const pad = n => n.toString().padStart(2, '0');
//   return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
// }
function getMachineNumber(machineName: string) {
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
    let worksheet: any;

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
        { header: 'Machine Number', key: 'machineNumber', width: 15 },
        { header: 'Date', key: 'date', width: 15 },
        { header: 'ON Time', key: 'onTime', width: 15 },
        { header: 'OFF Time', key: 'offTime', width: 15 },
        { header: 'Total Time (min)', key: 'totalTime', width: 18 },
        { header: 'Quantity', key: 'quantity', width: 12 }
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
      console.log('DEBUG: ALL ON JOBS:', allOn.map(j => ({ id: j.id, productId: j.productId, machineId: j.machineId, createdAt: j.createdAt })));
      console.log('DEBUG: ALL OFF JOBS:', allOff.map(j => ({ id: j.id, productId: j.productId, machineId: j.machineId, createdAt: j.createdAt, updatedAt: j.updatedAt })));
      // --- END DEBUG LOGGING ---
      // ---
      // All time formatting for ON/OFF is handled below. Always output as HH:mm (no seconds, no ms)
      // ---
      // Get all jobs for the operation type and date range
              const jobGroups: { [key: string]: any[] } = {};
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
      const allRows: any[] = [];
      Object.values(jobGroups).forEach((group) => {
        group.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
        // Separate ON and OFF jobs
        const onJobs = group.filter(j => j.state === 'ON');
        const offJobs = group.filter(j => j.state === 'OFF');
        let onIdx = 0, offIdx = 0;
        while (onIdx < onJobs.length && offIdx < offJobs.length) {
          const onJob = onJobs[onIdx];
          const offJob = offJobs[offIdx];
          
          // Log each ON/OFF pair
          console.log('DEBUG: PAIRING ON/OFF', {
            onId: onJob.id, offId: offJob.id,
            onCreatedAt: onJob.createdAt, offCreatedAt: offJob.createdAt, offUpdatedAt: offJob.updatedAt
          });
          // Use real ON and OFF times
          const onTime = new Date(onJob.createdAt);
          onTime.setSeconds(0, 0);
          const offTime = new Date(offJob.updatedAt || offJob.createdAt);
          offTime.setSeconds(0, 0);
          const dateStr = onTime.toISOString().split('T')[0];
          // Always use getMachineNumber
          // const machineNumber = getMachineNumber(onJob.machine.name);
          const onTimeStr = `${onTime.getHours().toString().padStart(2, '0')}:${onTime.getMinutes().toString().padStart(2, '0')}`;
          const offTimeStr = `${offTime.getHours().toString().padStart(2, '0')}:${offTime.getMinutes().toString().padStart(2, '0')}`;
          const totalTime = Math.round((offTime.getTime() - onTime.getTime()) / 60000);
          allRows.push({
            productId: onJob.product.name,
            machineNumber: getMachineNumber(onJob.machine.name),
            date: dateStr,
            onTime: onTimeStr,
            offTime: offTimeStr,
            totalTime: totalTime ? totalTime : '',
            quantity: onJob.quantity || 1
          });
          onIdx++;
          offIdx++;
        }
        // Any remaining ON jobs (not yet OFF)
        for (; onIdx < onJobs.length; onIdx++) {
          const onJob = onJobs[onIdx];
          // Log unmatched ON job
          console.log('DEBUG: UNMATCHED ON JOB', {
            onId: onJob.id, onCreatedAt: onJob.createdAt
          });
          const onTime = new Date(onJob.createdAt);
          onTime.setSeconds(0, 0);
          const dateStr = onTime.toISOString().split('T')[0];
          const machineNumber = getMachineNumber(onJob.machine.name);
          const onTimeStr = `${onTime.getHours().toString().padStart(2, '0')}:${onTime.getMinutes().toString().padStart(2, '0')}`;
          // Fallback: unmatched ON job, blank OFF/total time
          allRows.push({
            productId: onJob.product.name || onJob.productId,
            machineNumber,
            date: dateStr,
            onTime: onTimeStr,
            offTime: '',
            totalTime: '',
            quantity: onJob.quantity || 1
          });
        }
      });
      // Group rows by all parameters
              const groupedRowsObj: { [key: string]: any } = {};
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
            machineNumber: getMachineNumber(job.machine && job.machine.name),
            date: job.createdAt ? new Date(job.createdAt).toISOString().split('T')[0] : '',
            onTime: onTimeStr,
            offTime: offTimeStr,
            totalTime: '',
            quantity: 1
          });
        });
        addedRows = jobs.length;
      }
      if (addedRows === 0) {
        worksheet.addRow({ productId: 'No data found for the selected criteria.' });
      }
      worksheet.addRow([]);
      // Use Object.values for total product calculation
      const totalProducts = Object.keys(groupedRowsObj).length;
      const summaryRow = worksheet.addRow(['Total Products:', totalProducts, '', '', '', '', '']);
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

    if (reportType === 'dateWiseAllMachines') {
      // Date-wise report for all machines
      const selectedDate = searchParams.get('date');
      if (!selectedDate) {
        return NextResponse.json({ success: false, error: 'Date is required for Date-wise All Machines report.' }, { status: 400 });
      }

      // Parse the selected date and create date range for that day
      const reportDate = new Date(selectedDate);
      reportDate.setHours(0, 0, 0, 0);
      const nextDay = new Date(reportDate);
      nextDay.setDate(reportDate.getDate() + 1);
      nextDay.setHours(0, 0, 0, 0);

      worksheet = workbook.addWorksheet('Date-wise All Machines Report');
      
      // Set up columns for the report
      worksheet.columns = [
        { header: 'Product ID', key: 'productId', width: 20 },
        { header: 'Machine Number', key: 'machineNumber', width: 15 },
        { header: 'Process', key: 'process', width: 20 },
        { header: 'Quantity', key: 'quantity', width: 12 },
        { header: 'ON Time', key: 'onTime', width: 15 },
        { header: 'OFF Time', key: 'offTime', width: 15 },
        { header: 'Total Time (min)', key: 'totalTime', width: 18 }
      ];

      // Style the header row
      const headerRow = worksheet.getRow(1);
      headerRow.font = { bold: true };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } }; // Light gray background
      headerRow.eachCell((cell: any) => {
        cell.border = { 
          top: { style: 'thin' }, 
          bottom: { style: 'thin' }, 
          left: { style: 'thin' }, 
          right: { style: 'thin' } 
        };
      });

      // Get all jobs for the selected date across all machines
      const allJobs = await prisma.job.findMany({
        where: {
          createdAt: { gte: reportDate, lt: nextDay },
        },
        include: {
          machine: true,
          product: true,
        },
        orderBy: [
          { machine: { name: 'asc' } },
          { createdAt: 'asc' }
        ],
      });

      if (!allJobs || allJobs.length === 0) {
        worksheet.addRow(['No jobs found for the selected date.']);
        const reportName = `Date-wise_All_Machines_${selectedDate}_report`;
        await prisma.reportDownload.create({ data: { reportName } });
        const buffer = await workbook.xlsx.writeBuffer();
        const headers = new Headers({
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Content-Disposition': `attachment; filename="${reportName}.xlsx"`,
        });
        return new NextResponse(buffer, { status: 200, headers });
      }

      // Group jobs by machine and product for better organization
              const machineGroups: { [key: string]: any } = {};
      allJobs.forEach(job => {
        if (!job.machine || !job.product) return;
        
        const machineName = job.machine.name;
        const processName = machineName.split('#')[0].trim(); // Extract process name (e.g., "Cutting", "Gun Drilling")
        const machineNumber = getMachineNumber(machineName);
        
        if (!machineGroups[machineName]) {
          machineGroups[machineName] = {
            processName,
            machineNumber,
            products: {}
          };
        }
        
        // Group by product name
        const productName = job.product.name;
        if (!machineGroups[machineName].products[productName]) {
          machineGroups[machineName].products[productName] = {
            jobs: [],
            totalQuantity: 0,
            onTimes: [],
            offTimes: [],
            onJobs: [],
            offJobs: []
          };
        }
        
        machineGroups[machineName].products[productName].jobs.push(job);
        machineGroups[machineName].products[productName].totalQuantity += 1;
        
        if (job.state === 'ON') {
          machineGroups[machineName].products[productName].onTimes.push(new Date(job.createdAt));
          machineGroups[machineName].products[productName].onJobs.push(job);
        } else if (job.state === 'OFF') {
          machineGroups[machineName].products[productName].offTimes.push(new Date(job.updatedAt || job.createdAt));
          machineGroups[machineName].products[productName].offJobs.push(job);
        }
      });

      // Process each machine group
      let totalRows = 0;
      Object.entries(machineGroups).forEach(([machineName, machineData]: [string, any]) => {
        // Add machine header row
        worksheet.addRow([]);
        worksheet.addRow([`Machine: ${machineName}`, '', '', '', '', '', '']);
        const machineHeaderRow = worksheet.getRow(worksheet.rowCount);
        machineHeaderRow.getCell(1).font = { bold: true, color: { argb: 'FF0000FF' } }; // Blue color
        machineHeaderRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } }; // Light gray background
        // Add border to the machine header
        machineHeaderRow.getCell(1).border = { 
          top: { style: 'thin' }, 
          bottom: { style: 'thin' }, 
          left: { style: 'thin' }, 
          right: { style: 'thin' } 
        };
        
        // Process products for this machine
        const { products, processName, machineNumber } = machineData;
        
        Object.entries(products).forEach(([productName, productData]: [string, any]) => {
          const { totalQuantity, onTimes, offTimes, onJobs, offJobs } = productData;
          
          let onTimeStr = '';
          let offTimeStr = '';
          let totalTime = '';
          
          // Try to pair ON/OFF jobs first
          if (onJobs.length > 0 && offJobs.length > 0) {
            // Sort jobs by time to find the best pairs
            const sortedOnJobs = onJobs.sort((a: any, b: any) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
            const sortedOffJobs = offJobs.sort((a: any, b: any) => new Date(a.updatedAt || a.createdAt).getTime() - new Date(b.updatedAt || b.createdAt).getTime());
            
            // Find the earliest ON and latest OFF time
            const earliestOnTime = new Date(sortedOnJobs[0].createdAt);
            const latestOffTime = new Date(sortedOffJobs[sortedOffJobs.length - 1].updatedAt || sortedOffJobs[sortedOffJobs.length - 1].createdAt);
            
            onTimeStr = `${earliestOnTime.getHours().toString().padStart(2, '0')}:${earliestOnTime.getMinutes().toString().padStart(2, '0')}`;
            offTimeStr = `${latestOffTime.getHours().toString().padStart(2, '0')}:${latestOffTime.getMinutes().toString().padStart(2, '0')}`;
            
            // Calculate total time
            if (latestOffTime > earliestOnTime) {
              totalTime = Math.round((latestOffTime.getTime() - earliestOnTime.getTime()) / 60000).toString();
            }
          } else if (onJobs.length > 0) {
            // Only ON jobs exist
            const earliestOnTime = new Date(Math.min(...onTimes.map((t: Date) => t.getTime())));
            onTimeStr = `${earliestOnTime.getHours().toString().padStart(2, '0')}:${earliestOnTime.getMinutes().toString().padStart(2, '0')}`;
            
            // Try to estimate OFF time from the latest ON job + some buffer
            const latestOnJob = onJobs.sort((a: any, b: any) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0];
            const estimatedOffTime = new Date(latestOnJob.createdAt.getTime() + 30 * 60000); // Add 30 minutes buffer
            offTimeStr = `${estimatedOffTime.getHours().toString().padStart(2, '0')}:${estimatedOffTime.getMinutes().toString().padStart(2, '0')}`;
            
          } else if (offJobs.length > 0) {
            // Only OFF jobs exist - try to estimate ON time
            const latestOffTime = new Date(Math.max(...offTimes.map((t: Date) => t.getTime())));
            offTimeStr = `${latestOffTime.getHours().toString().padStart(2, '0')}:${latestOffTime.getMinutes().toString().padStart(2, '0')}`;
            
            // Estimate ON time by subtracting some time from OFF time
            const estimatedOnTime = new Date(latestOffTime.getTime() - 30 * 60000); // Subtract 30 minutes
            onTimeStr = `${estimatedOnTime.getHours().toString().padStart(2, '0')}:${estimatedOnTime.getMinutes().toString().padStart(2, '0')}`;
            
            // Calculate estimated total time
            totalTime = '30'; // Estimated 30 minutes
          }
          
          worksheet.addRow([
            productName,
            machineNumber,
            processName,
            totalQuantity,
            onTimeStr,
            offTimeStr,
            totalTime || ''
          ]);
          totalRows++;
        });
      });

      // Add summary
      worksheet.addRow([]);
      const summaryRow = worksheet.addRow(['Total Jobs:', totalRows, '', '', '', '', '']);
      summaryRow.getCell('A').font = { bold: true };
      summaryRow.getCell('B').font = { bold: true };

      const reportName = `Date-wise_All_Machines_${selectedDate}_report`;
      await prisma.reportDownload.create({ data: { reportName } });
      const buffer = await workbook.xlsx.writeBuffer();
      const headers = new Headers({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${reportName}.xlsx"`,
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
    const totalProducts = dispatchedProducts.length;
    const summaryRow = worksheet.addRow(['Total Products Dispatched:', totalProducts, '']);
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