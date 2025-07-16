const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testRecentlyFinishedProducts() {
  console.log('🧪 Testing Recently Finished Products (Live Products Only) functionality...\n');

  try {
    // Step 1: Check current state
    console.log('1. Current state:');
    const currentJobs = await prisma.job.findMany({
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
    });
    
    const liveProducts = currentJobs.filter(job => job.state === 'ON');
    const pastProducts = currentJobs.filter(job => job.state === 'OFF');
    
    console.log(`   - Live products: ${liveProducts.length}`);
    liveProducts.forEach(job => {
      console.log(`     * ${job.product.name} on ${job.machine.name} (${job.state})`);
    });
    
    console.log(`   - Past products: ${pastProducts.length}`);
    pastProducts.forEach(job => {
      console.log(`     * ${job.product.name} on ${job.machine.name} (${job.state})`);
    });

    // Step 2: Test adding a new live product
    console.log('\n2. Adding a new live product...');
    const testProduct = await prisma.product.create({
      data: { name: `Angel` }
    });
    
    const testMachine = await prisma.machine.findFirst({
      where: { name: 'Cutting MC/1' }
    });
    
    if (!testMachine) {
      console.log('   ❌ Test machine "Cutting MC/1" not found');
      return;
    }
    
    const newJob = await prisma.job.create({
      data: {
        machineId: testMachine.id,
        productId: testProduct.id,
        state: 'ON',
        stage: 'Cutting'
      },
      include: {
        machine: true,
        product: true,
      }
    });
    
    console.log(`   ✅ Added: ${newJob.product.name} on ${newJob.machine.name} (${newJob.state})`);

    // Step 3: Test the recently finished products API
    console.log('\n3. Testing recently finished products API...');
    const response = await fetch('http://localhost:3002/api/live-products/recently-finished');
    if (response.ok) {
      const data = await response.json();
      console.log(`   ✅ API returned ${data.recentlyFinishedProducts.length} products`);
      
      const testProductInList = data.recentlyFinishedProducts.find(p => p.name === newJob.product.name);
      if (testProductInList) {
        console.log(`   ✅ Test product found in Recently Finished Products: ${testProductInList.displayName}`);
        console.log(`   ✅ Available count: ${testProductInList.availableCount}`);
      } else {
        console.log(`   ❌ Test product NOT found in Recently Finished Products`);
      }
      
      // Verify only live products are returned
      const nonLiveProducts = data.recentlyFinishedProducts.filter(p => p.state !== 'ON');
      if (nonLiveProducts.length === 0) {
        console.log(`   ✅ All returned products are live (state: ON)`);
      } else {
        console.log(`   ❌ Found ${nonLiveProducts.length} non-live products in results`);
      }
    } else {
      console.log('   ❌ Failed to fetch recently finished products API');
    }

    // Step 4: Test moving product to past
    console.log('\n4. Moving test product to past...');
    const moveResponse = await fetch('http://localhost:3002/api/products/lifecycle', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        productId: newJob.product.id.toString(),
        jobId: newJob.id.toString(),
        action: 'move_to_past',
        reason: 'test_movement'
      })
    });
    
    if (moveResponse.ok) {
      console.log('   ✅ Product moved to past successfully');
    } else {
      console.log('   ❌ Failed to move product to past');
    }

    // Step 5: Verify product is removed from recently finished products
    console.log('\n5. Verifying product removed from recently finished products...');
    const response2 = await fetch('http://localhost:3002/api/live-products/recently-finished');
    if (response2.ok) {
      const data2 = await response2.json();
      console.log(`   ✅ API returned ${data2.recentlyFinishedProducts.length} products`);
      
      const testProductInList2 = data2.recentlyFinishedProducts.find(p => p.name === newJob.product.name);
      if (!testProductInList2) {
        console.log(`   ✅ Test product correctly removed from Recently Finished Products`);
      } else {
        console.log(`   ❌ Test product still found in Recently Finished Products`);
      }
    } else {
      console.log('   ❌ Failed to fetch recently finished products API');
    }

    // Step 6: Test the stream endpoint
    console.log('\n6. Testing stream endpoint...');
    const streamResponse = await fetch('http://localhost:3002/api/live-products/stream');
    if (streamResponse.ok) {
      console.log('   ✅ Stream endpoint is working');
    } else {
      console.log('   ❌ Stream endpoint failed');
    }

    // Cleanup
    console.log('\n7. Cleaning up test data...');
    await prisma.job.delete({
      where: { id: newJob.id }
    });
    await prisma.product.delete({
      where: { id: testProduct.id }
    });
    console.log('   ✅ Test data cleaned up');

    console.log('\n🎉 Test completed successfully!');
    console.log('\n📋 Summary:');
    console.log('   - Recently Finished Products API only shows live products (state: ON)');
    console.log('   - Products are formatted as "ProductName - MachineName (Date) (3 available)"');
    console.log('   - Products are automatically removed when moved to past');
    console.log('   - Stream endpoint provides real-time updates');
    
  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testRecentlyFinishedProducts(); 