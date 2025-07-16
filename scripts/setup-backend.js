const { PrismaClient } = require('@prisma/client');
const fetch = require('node-fetch');

const prisma = new PrismaClient();
const BASE_URL = 'http://localhost:3000/api';

async function setupBackend() {
  console.log('🚀 Setting up TrackOPZ Backend System...\n');

  try {
    // Step 1: Test database connection
    console.log('1. Testing database connection...');
    await prisma.$connect();
    console.log('✅ Database connection successful\n');

    // Step 2: Create sample data if none exists
    console.log('2. Checking for existing data...');
    const existingJobs = await prisma.job.count();
    const existingProducts = await prisma.product.count();
    const existingMachines = await prisma.machine.count();

    if (existingJobs === 0 && existingProducts === 0 && existingMachines === 0) {
      console.log('📝 No existing data found. Creating sample data...');
      
      // Create sample machines
      const machines = await Promise.all([
        prisma.machine.create({ data: { name: 'Cutting MC/1', status: 'ON' } }),
        prisma.machine.create({ data: { name: 'Milling 1', status: 'OFF' } }),
        prisma.machine.create({ data: { name: 'Milling 2', status: 'ON' } }),
        prisma.machine.create({ data: { name: 'Drilling', status: 'OFF' } }),
        prisma.machine.create({ data: { name: 'CNC Finish', status: 'ON' } })
      ]);

      // Create sample products
      const products = await Promise.all([
        prisma.product.create({ data: { name: 'Product A' } }),
        prisma.product.create({ data: { name: 'Product B' } }),
        prisma.product.create({ data: { name: 'Product C' } })
      ]);

      // Create sample jobs (some live, some past)
      const jobs = await Promise.all([
        // Live products
        prisma.job.create({
          data: {
            machineId: machines[0].id,
            productId: products[0].id,
            state: 'ON',
            stage: 'Cutting'
          }
        }),
        prisma.job.create({
          data: {
            machineId: machines[2].id,
            productId: products[1].id,
            state: 'ON',
            stage: 'Milling'
          }
        }),
        // Past products
        prisma.job.create({
          data: {
            machineId: machines[4].id,
            productId: products[2].id,
            state: 'OFF',
            stage: 'Finished'
          }
        })
      ]);

      console.log(`✅ Created ${machines.length} machines, ${products.length} products, and ${jobs.length} jobs\n`);
    } else {
      console.log(`✅ Found existing data: ${existingJobs} jobs, ${existingProducts} products, ${existingMachines} machines\n`);
    }

    // Step 3: Test API endpoints
    console.log('3. Testing API endpoints...');
    
    const endpoints = [
      { name: 'Jobs', url: '/jobs' },
      { name: 'Dispatchable Products', url: '/live-products/dispatchable' },
      { name: 'Dashboard Status', url: '/dashboard/status' }
    ];

    for (const endpoint of endpoints) {
      try {
        const response = await fetch(`${BASE_URL}${endpoint.url}`);
        if (response.ok) {
          console.log(`✅ ${endpoint.name} endpoint: Working`);
        } else {
          console.log(`❌ ${endpoint.name} endpoint: Failed (${response.status})`);
        }
      } catch (error) {
        console.log(`❌ ${endpoint.name} endpoint: Error - ${error.message}`);
      }
    }

    console.log('\n4. Testing real-time functionality...');
    try {
      const streamResponse = await fetch(`${BASE_URL}/live-products/dispatchable/stream`);
      if (streamResponse.ok) {
        console.log('✅ Real-time stream endpoint: Working');
      } else {
        console.log(`❌ Real-time stream endpoint: Failed (${streamResponse.status})`);
      }
    } catch (error) {
      console.log(`❌ Real-time stream endpoint: Error - ${error.message}`);
    }

    // Step 4: Display system status
    console.log('\n5. Current system status:');
    const dashboardResponse = await fetch(`${BASE_URL}/dashboard/status`);
    if (dashboardResponse.ok) {
      const dashboard = await dashboardResponse.json();
      if (dashboard.success) {
        const stats = dashboard.dashboard.stats;
        console.log(`   📊 Total Products: ${stats.totalProducts}`);
        console.log(`   🟢 Live Products: ${stats.liveProducts}`);
        console.log(`   🔴 Past Products: ${stats.pastProducts}`);
        console.log(`   📦 Dispatchable Products: ${stats.dispatchableProducts}`);
        console.log(`   ⚙️  Active Machines: ${stats.activeMachines}/${stats.totalMachines}`);
      }
    }

    console.log('\n🎉 Backend setup completed successfully!');
    console.log('\n📋 Next steps:');
    console.log('1. Start your Next.js development server: npm run dev');
    console.log('2. Visit http://localhost:3000 to access the application');
    console.log('3. Navigate to /workpanel to see live and past products');
    console.log('4. Navigate to /dispatch to test the dispatch functionality');
    console.log('5. Run the test script: node scripts/test-backend.js');

  } catch (error) {
    console.error('❌ Setup failed:', error.message);
    console.log('\n🔧 Troubleshooting:');
    console.log('1. Make sure your database is running and accessible');
    console.log('2. Check your DATABASE_URL environment variable');
    console.log('3. Run "npx prisma generate" to generate Prisma client');
    console.log('4. Run "npx prisma db push" to sync database schema');
  } finally {
    await prisma.$disconnect();
  }
}

// Run the setup
setupBackend(); 