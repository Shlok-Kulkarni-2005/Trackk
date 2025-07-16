import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch all dispatched items and RFD products
export async function GET() {
  try {
    // Get dispatched items from OperatorProductUpdate
    const dispatchedItems = await prisma.operatorProductUpdate.findMany({
      where: {
        dispatchStatus: 'Dispatched',
        archived: false,
      },
      orderBy: {
        createdAt: 'desc',
      },
    });

    // Get RFD products that are turned off
    const rfdProducts = await prisma.job.findMany({
      where: {
        machine: {
          name: 'RFD'
        },
        state: 'OFF'
      },
      include: {
        product: true,
        machine: true
      },
      orderBy: { createdAt: 'desc' }
    });

    // Get today's date for filtering (local date)
    const today = new Date();
    const todayString = today.toLocaleDateString('en-CA'); // 'YYYY-MM-DD'
    
    // Group dispatched items by product name and date, counting quantities
    const groupedItems: { [productName: string]: any } = {};
    const todayGroupedItems: { [productName: string]: any } = {};
    
    dispatchedItems.forEach(item => {
      const itemDate = item.createdAt.toLocaleDateString('en-CA');
      const isToday = itemDate === todayString;
      
      // Group for today's dispatches
      if (isToday) {
        if (!todayGroupedItems[item.product]) {
          todayGroupedItems[item.product] = {
            id: item.id.toString(),
            product: item.product,
            quantity: 1,
            destination: 'Dispatched',
            notes: '',
            dispatchedAt: item.createdAt.toISOString(),
            status: item.dispatchStatus,
            totalQuantity: 1,
            lastUpdated: item.createdAt.toISOString(),
            date: itemDate,
            type: 'dispatched'
          };
        } else {
          todayGroupedItems[item.product].totalQuantity += 1;
          if (new Date(item.createdAt) > new Date(todayGroupedItems[item.product].lastUpdated)) {
            todayGroupedItems[item.product].dispatchedAt = item.createdAt.toISOString();
            todayGroupedItems[item.product].lastUpdated = item.createdAt.toISOString();
          }
        }
      }
      
      // Group for all dispatches (history)
      if (!groupedItems[item.product]) {
        groupedItems[item.product] = {
          id: item.id.toString(),
          product: item.product,
          quantity: 1,
          destination: 'Dispatched',
          notes: '',
          dispatchedAt: item.createdAt.toISOString(),
          status: item.dispatchStatus,
          totalQuantity: 1,
          lastUpdated: item.createdAt.toISOString(),
          date: itemDate,
          type: 'dispatched'
        };
      } else {
        groupedItems[item.product].totalQuantity += 1;
        if (new Date(item.createdAt) > new Date(groupedItems[item.product].lastUpdated)) {
          groupedItems[item.product].dispatchedAt = item.createdAt.toISOString();
          groupedItems[item.product].lastUpdated = item.createdAt.toISOString();
        }
      }
    });

    // Add RFD products and increment quantity for each job, grouped by product and date
    rfdProducts.forEach(job => {
      const productName = job.product.name;
      const jobDate = job.createdAt.toLocaleDateString('en-CA');
      const isToday = jobDate === todayString;
      const key = `${productName}__${jobDate}`;
      if (isToday) {
        if (!todayGroupedItems[key]) {
          todayGroupedItems[key] = {
            id: `rfd_${job.id}`,
            product: productName,
            quantity: 1,
            destination: 'RFD',
            notes: 'RFD machine turned off',
            dispatchedAt: job.createdAt.toISOString(),
            status: 'OFF',
            totalQuantity: 1,
            lastUpdated: job.createdAt.toISOString(),
            date: jobDate,
            type: 'rfd'
          };
        } else {
          todayGroupedItems[key].quantity += 1;
          todayGroupedItems[key].totalQuantity += 1;
          if (new Date(job.createdAt) > new Date(todayGroupedItems[key].lastUpdated)) {
            todayGroupedItems[key].dispatchedAt = job.createdAt.toISOString();
            todayGroupedItems[key].lastUpdated = job.createdAt.toISOString();
          }
        }
      } else {
        if (!groupedItems[key]) {
          groupedItems[key] = {
            id: `rfd_${job.id}`,
            product: productName,
            quantity: 1,
            destination: 'RFD',
            notes: 'RFD machine turned off',
            dispatchedAt: job.createdAt.toISOString(),
            status: 'OFF',
            totalQuantity: 1,
            lastUpdated: job.createdAt.toISOString(),
            date: jobDate,
            type: 'rfd'
          };
        } else {
          groupedItems[key].quantity += 1;
          groupedItems[key].totalQuantity += 1;
          if (new Date(job.createdAt) > new Date(groupedItems[key].lastUpdated)) {
            groupedItems[key].dispatchedAt = job.createdAt.toISOString();
            groupedItems[key].lastUpdated = job.createdAt.toISOString();
          }
        }
      }
    });

    // Convert grouped items to arrays and sort by last updated
    // Exclude any product from history if it is present in today's
    const todayProductNames = new Set(Object.keys(todayGroupedItems));
    const transformedItems = Object.values(groupedItems)
      .filter(item => !todayProductNames.has(item.product))
      .map(item => ({
        ...item,
        quantity: item.totalQuantity,
      }))
      .sort((a, b) => new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime());

    const todayTransformedItems = Object.values(todayGroupedItems)
      .map(item => ({
        ...item,
        quantity: item.totalQuantity,
      }))
      .sort((a, b) => new Date(b.lastUpdated).getTime() - new Date(a.lastUpdated).getTime());

    return NextResponse.json({ 
      success: true, 
      dispatchedItems: transformedItems,
      todayDispatchedItems: todayTransformedItems
    });
  } catch (error) {
    console.error('Error fetching dispatched items:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
}

// POST: Create a new dispatched item
export async function POST(req: NextRequest) {
  try {
    const { product, destination } = await req.json();

    if (!product || !destination) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product and destination are required' 
      }, { status: 400 });
    }

    // Create a new dispatched item
    const dispatchedItem = await prisma.operatorProductUpdate.create({
      data: {
        operatorId: 1, // Default operator ID since we don't have authentication
        product,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched',
        dispatchedCost: 0, // Default cost
      },
    });

    return NextResponse.json({ 
      success: true, 
      dispatchedItem 
    }, { status: 201 });
  } catch (error) {
    console.error('Error creating dispatched item:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
} 