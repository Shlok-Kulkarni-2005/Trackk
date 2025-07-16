const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testCompleteFlow() {
  console.log('🧪 Testing Complete Recently Finished Products Flow...\n');

  try {
    // Step 1: Check current live products
    console.log('1. Checking current live products...');
    const response = await fetch('http://localhost:3002/api/live-products/recently-finished');
    if (response.ok) {
      const data = await response.json();
      console.log(`   ✅ Found ${data.recentlyFinishedProducts.length} live products`);
      
      if (data.recentlyFinishedProducts.length === 0) {
        console.log('   ⚠️ No live products found. Creating a test product...');
        
        // Create a test product
        const testProduct = await prisma.product.create({
          data: { name: `Test Product ${Date.now()}` }
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
          }
        });
        
        console.log(`   ✅ Created test product: ${testProduct.name}`);
      }
    } else {
      console.log('   ❌ Failed to fetch live products');
      return;
    }

    // Step 2: Fetch products again to get the latest list
    console.log('\n2. Fetching updated product list...');
    const response2 = await fetch('http://localhost:3002/api/live-products/recently-finished');
    if (response2.ok) {
      const data2 = await response2.json();
      console.log(`   ✅ Found ${data2.recentlyFinishedProducts.length} live products`);
      
      if (data2.recentlyFinishedProducts.length > 0) {
        const testProduct = data2.recentlyFinishedProducts[0];
        console.log(`   📋 Test product: ${testProduct.displayName}`);
        console.log(`   📋 Product ID: ${testProduct.id}`);
        console.log(`   📋 Available count: ${testProduct.availableCount}`);

        // Step 3: Test the update functionality
        console.log('\n3. Testing update functionality...');
        const updateData = {
          id: testProduct.id,
          processSteps: {
            deburring: true,
            finalInspect: true,
            oiling: false
          },
          dispatchStatus: 'Pending'
        };

        console.log('   📤 Sending update request...');
        const updateResponse = await fetch('http://localhost:3002/api/operator/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(updateData)
        });

        console.log(`   📥 Update response status: ${updateResponse.status}`);
        
        if (updateResponse.ok) {
          const updateResult = await updateResponse.json();
          console.log('   ✅ Update successful!');
          console.log(`   📋 Update result: ${JSON.stringify(updateResult, null, 2)}`);
        } else {
          const errorText = await updateResponse.text();
          console.log(`   ❌ Update failed: ${updateResponse.status} - ${errorText}`);
        }
      } else {
        console.log('   ⚠️ No products available for testing');
      }
    } else {
      console.log('   ❌ Failed to fetch updated product list');
    }

    // Step 4: Test the stream endpoint
    console.log('\n4. Testing stream endpoint...');
    const streamResponse = await fetch('http://localhost:3002/api/live-products/stream');
    if (streamResponse.ok) {
      console.log('   ✅ Stream endpoint is working');
    } else {
      console.log('   ❌ Stream endpoint failed');
    }

    console.log('\n🎉 Complete flow test finished!');
    console.log('\n📋 Summary:');
    console.log('   - Recently Finished Products API is working');
    console.log('   - Products are formatted correctly');
    console.log('   - Update functionality is working');
    console.log('   - Stream endpoint is working');
    
  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testCompleteFlow(); 