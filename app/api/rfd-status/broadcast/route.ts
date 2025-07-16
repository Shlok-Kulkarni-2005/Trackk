import { NextRequest, NextResponse } from 'next/server';
import { broadcastRFDStatusUpdate } from '../../shared/broadcast';

// POST: Broadcast RFD status updates to connected clients
export async function POST(req: NextRequest) {
  try {
    const updateData = await req.json();
    
    // Broadcast the RFD status update
    broadcastRFDStatusUpdate(updateData);
    
    return NextResponse.json({ 
      success: true, 
      message: 'RFD status update broadcasted successfully' 
    }, { status: 200 });

  } catch (error) {
    console.error('Error broadcasting RFD status update:', error);
    return NextResponse.json({ 
      success: false, 
      error: 'Failed to broadcast RFD status update' 
    }, { status: 500 });
  }
} 