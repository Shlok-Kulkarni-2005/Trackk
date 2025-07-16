import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// POST: Add or update dispatched product quantity
export async function POST(req: NextRequest) {
  try {
    const { product, operatorId = 12345 } = await req.json();
    
    if (!product) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product name is required' 
      }, { status: 400 });
    }

    // Check if there's already a dispatched item for this product
    const existingDispatchedItem = await prisma.operatorProductUpdate.findFirst({
      where: {
        product: product,
        dispatchStatus: 'Dispatched',
        archived: false
      },
      orderBy: {
        createdAt: 'desc'
      }
    });

    let result;

    if (existingDispatchedItem) {
      // Product already exists in dispatched items
      // Create a new record to represent the additional quantity
      result = await prisma.operatorProductUpdate.create({
        data: {
          operatorId: operatorId,
          product: product,
          processSteps: {
            deburring: true,
            finalInspect: true,
            oiling: true
          },
          dispatchStatus: 'Dispatched',
          dispatchedCost: 0, // Default cost
        }
      });

      console.log(`Increased quantity for product "${product}" in dispatched items`);
    } else {
      // Product doesn't exist in dispatched items, create new entry
      result = await prisma.operatorProductUpdate.create({
        data: {
          operatorId: operatorId,
          product: product,
          processSteps: {
            deburring: true,
            finalInspect: true,
            oiling: true
          },
          dispatchStatus: 'Dispatched',
          dispatchedCost: 0, // Default cost
        }
      });

      console.log(`Added new product "${product}" to dispatched items`);
    }

    return NextResponse.json({ 
      success: true, 
      message: existingDispatchedItem ? 'Product quantity increased in dispatched items' : 'Product added to dispatched items',
      dispatchedItem: result
    }, { status: 200 });

  } catch (error) {
    console.error('Error updating dispatched product quantity:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
}

// GET: Get dispatched product quantity for a specific product
export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const product = searchParams.get('product');
    
    if (!product) {
      return NextResponse.json({ 
        success: false, 
        error: 'Product parameter is required' 
      }, { status: 400 });
    }

    // Count all dispatched items for this product
    const dispatchedCount = await prisma.operatorProductUpdate.count({
      where: {
        product: product,
        dispatchStatus: 'Dispatched',
        archived: false
      }
    });

    return NextResponse.json({ 
      success: true, 
      product: product,
      quantity: dispatchedCount
    }, { status: 200 });

  } catch (error) {
    console.error('Error getting dispatched product quantity:', error);
    return NextResponse.json({ 
      success: false, 
      error: error instanceof Error ? error.message : String(error) 
    }, { status: 500 });
  }
} 