import { NextRequest } from 'next/server';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export async function GET(request: NextRequest) {
  const encoder = new TextEncoder();
  
  const stream = new ReadableStream({
    async start(controller) {
      let keepAliveInterval: NodeJS.Timeout | null = null;
      
      try {
        // Send initial connection message
        const connectionMessage = `data: ${JSON.stringify({
          type: 'connection_established',
          timestamp: new Date().toISOString()
        })}\n\n`;
        controller.enqueue(encoder.encode(connectionMessage));

        // Send initial data immediately
        const allJobs = await prisma.job.findMany({
          orderBy: { createdAt: 'desc' },
          include: { 
            product: true,
            machine: true 
          },
        });

        const latestJobByProductId: Record<number, any> = {};
        allJobs.forEach(job => {
          const productId = job.productId;
          if (!latestJobByProductId[productId] || new Date(job.createdAt) > new Date(latestJobByProductId[productId].createdAt)) {
            latestJobByProductId[productId] = job;
          }
        });

        const liveProducts = Object.values(latestJobByProductId)
          .filter((job: any) => job.state === 'ON')
          .map((job: any) => ({
            id: job.product.id,
            name: job.product.name,
            process: job.machine.name,
            state: job.state,
            date: new Date(job.createdAt).toLocaleDateString(),
            createdAt: job.createdAt,
            machineName: job.machine.name,
            productName: job.product.name,
            displayName: `${job.product.name} - ${job.machine.name} (${new Date(job.createdAt).toLocaleDateString()})`,
            availableCount: 3
          }));

        const initialMessage = `data: ${JSON.stringify({
          type: 'live_products_update',
          data: {
            liveProducts: liveProducts,
            count: liveProducts.length
          },
          timestamp: new Date().toISOString()
        })}\n\n`;
        controller.enqueue(encoder.encode(initialMessage));

        // Keep the connection alive with periodic updates
        keepAliveInterval = setInterval(async () => {
          try {
            // Check if controller is still open
            if (controller.desiredSize === null) {
              clearInterval(keepAliveInterval!);
              return;
            }

            // Get current live products
            const allJobs = await prisma.job.findMany({
              orderBy: { createdAt: 'desc' },
              include: { 
                product: true,
                machine: true 
              },
            });

            // Get the latest job for each productId
            const latestJobByProductId: Record<number, any> = {};
            allJobs.forEach(job => {
              const productId = job.productId;
              if (!latestJobByProductId[productId] || new Date(job.createdAt) > new Date(latestJobByProductId[productId].createdAt)) {
                latestJobByProductId[productId] = job;
              }
            });

            // Filter only products that are currently live (state 'ON')
            const liveProducts = Object.values(latestJobByProductId)
              .filter((job: any) => job.state === 'ON')
              .map((job: any) => ({
                id: job.product.id,
                name: job.product.name,
                process: job.machine.name,
                state: job.state,
                date: new Date(job.createdAt).toLocaleDateString(),
                createdAt: job.createdAt,
                machineName: job.machine.name,
                productName: job.product.name,
                displayName: `${job.product.name} - ${job.machine.name} (${new Date(job.createdAt).toLocaleDateString()})`,
                availableCount: 3
              }));

            // Send update message
            const updateMessage = `data: ${JSON.stringify({
              type: 'live_products_update',
              data: {
                liveProducts: liveProducts,
                count: liveProducts.length
              },
              timestamp: new Date().toISOString()
            })}\n\n`;
            controller.enqueue(encoder.encode(updateMessage));

          } catch {
            console.error('Error in live products stream update');
            // Don't send error messages to avoid controller issues
          }
        }, 30000); // Update every 30 seconds instead of 10

        // Handle connection close
        request.signal.addEventListener('abort', () => {
          if (keepAliveInterval) {
            clearInterval(keepAliveInterval);
          }
          try {
            controller.close();
          } catch (error) {
            // Ignore errors when closing
          }
        });

              } catch {
          console.error('Error in live products stream');
          try {
            const errorMessage = `data: ${JSON.stringify({
              type: 'error',
              error: 'Failed to initialize live products stream',
              timestamp: new Date().toISOString()
            })}\n\n`;
            controller.enqueue(encoder.encode(errorMessage));
            controller.close();
          } catch {
            // Ignore errors when closing
          }
        }
    }
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Headers': 'Cache-Control',
    },
  });
} 