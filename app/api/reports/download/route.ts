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

// Main GET handler
export async function GET(req: NextRequest) {
  const { searchParams } = new URL(req.url);
  const reportType = searchParams.get('reportType') || 'daily';
  const startDateParam = searchParams.get('startDate') || undefined;
  const endDateParam = searchParams.get('endDate') || undefined;

  try {
    const { startDate, endDate } = getDateRange(reportType, startDateParam, endDateParam);
    
    // Only get dispatched products (where dispatch button was clicked)
    const whereClause = { 
      createdAt: { gte: startDate, lte: endDate },
      dispatchStatus: 'Pending' // Only products that were dispatched
    };

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Dispatched Products Report');
    
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