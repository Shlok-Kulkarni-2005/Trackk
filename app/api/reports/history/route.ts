import { NextRequest, NextResponse } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

// GET: Fetch recent report downloads
export async function GET(req: NextRequest) {
  try {
    const { searchParams } = new URL(req.url);
    const limit = parseInt(searchParams.get('limit') || '10');
    
    const recentDownloads = await prisma.reportDownload.findMany({
      orderBy: { downloadedAt: 'desc' },
      take: limit,
    });

    return NextResponse.json({ 
      downloads: recentDownloads,
      total: recentDownloads.length 
    }, { status: 200 });
  } catch (error) {
    console.error('Error fetching recent downloads:', error);
    return NextResponse.json({ 
      error: 'Failed to fetch recent downloads' 
    }, { status: 500 });
  }
} 