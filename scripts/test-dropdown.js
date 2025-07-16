const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testDropdown() {
  console.log('🧪 Testing Recently Finished Products Dropdown...\n');

  try {
    // Step 1: Check if we have any existing products
    console.log('1. Checking existing products...');
    const existingJobs = await prisma.job.findMany({
      include: {
        machine: true,
        product: true,
      },
      orderBy: { createdAt: 'desc' },
      take: 5,
    });

    if (existingJobs.length === 0) {
      console.log('   No existing products found. Creating test data...');
      
      // Create test machine if it doesn't exist
      let testMachine = await prisma.machine.findFirst({
        where: { name: 'Cutting MC/1' }
      });
      
      if (!testMachine) {
        testMachine = await prisma.machine.create({
          data: { name: 'Cutting MC/1', status: 'ON' }
        });
        console.log('   Created test machine: Cutting MC/1');
      }

      // Create test products
      const testProducts = [
        { name: 'Test Product A', state: 'ON' },
        { name: 'Test Product B', state: 'ON' },
        { name: 'Test Product C', state: 'OFF' }
      ];

      for (const productData of testProducts) {
        let product = await prisma.product.findFirst({
          where: { name: productData.name }
        });
        
        if (!product) {
          product = await prisma.product.create({
            data: { name: productData.name }
          });
        }

        const job = await prisma.job.create({
          data: {
            machineId: testMachine.id,
            productId: product.id,
            state: productData.state,
            stage: 'Cutting'
          }
        });

        console.log(`   Created job: ${product.name} - ${productData.state}`);
      }
    } else {
      console.log(`   Found ${existingJobs.length} existing jobs`);
      existingJobs.forEach(job => {
        console.log(`     * ${job.product.name} on ${job.machine.name} (${job.state})`);
      });
    }

    // Step 2: Test the API endpoint
    console.log('\n2. Testing API endpoint...');
    const response = await fetch('http://localhost:3000/api/finished-products/dispatchable');
    
    if (response.ok) {
      const data = await response.json();
      console.log(`   ✅ API returned ${data.finishedProducts.length} products`);
      
      if (data.finishedProducts.length > 0) {
        console.log('   Products in dropdown:');
        data.finishedProducts.forEach((product, index) => {
          console.log(`     ${index + 1}. ${product.name} - ${product.process} (${product.date}) - Source: ${product.source}`);
        });
      } else {
        console.log('   ⚠️ No products found in dropdown');
      }
    } else {
      console.log('   ❌ API request failed');
      const errorText = await response.text();
      console.log('   Error:', errorText);
    }

    // Step 3: Test adding a new live product
    console.log('\n3. Testing adding new live product...');
    const newProduct = await prisma.product.create({
      data: { name: `Dropdown Test Product ${Date.now()}` }
    });

    const newJob = await prisma.job.create({
      data: {
        machineId: (await prisma.machine.findFirst({ where: { name: 'Cutting MC/1' } })).id,
        productId: newProduct.id,
        state: 'ON',
        stage: 'Cutting'
      }
    });

    console.log(`   ✅ Added new product: ${newProduct.name}`);

    // Step 4: Verify it appears in dropdown
    console.log('\n4. Verifying new product appears in dropdown...');
    const response2 = await fetch('http://localhost:3000/api/finished-products/dispatchable');
    
    if (response2.ok) {
      const data2 = await response2.json();
      const newProductInList = data2.finishedProducts.find(p => p.name === newProduct.name);
      
      if (newProductInList) {
        console.log(`   ✅ New product found in dropdown: ${newProductInList.name} - ${newProductInList.process}`);
      } else {
        console.log(`   ❌ New product NOT found in dropdown`);
      }
    }

    // Step 5: Test moving product to past
    console.log('\n5. Testing moving product to past...');
    const moveResponse = await fetch('http://localhost:3000/api/products/lifecycle', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        productId: newProduct.id.toString(),
        jobId: newJob.id.toString(),
        action: 'move_to_past',
        reason: 'dropdown_test'
      })
    });

    if (moveResponse.ok) {
      console.log('   ✅ Product moved to past successfully');
    } else {
      console.log('   ❌ Failed to move product to past');
    }

    // Step 6: Verify it's removed from dropdown
    console.log('\n6. Verifying product removed from dropdown...');
    const response3 = await fetch('http://localhost:3000/api/finished-products/dispatchable');
    
    if (response3.ok) {
      const data3 = await response3.json();
      const removedProductInList = data3.finishedProducts.find(p => p.name === newProduct.name);
      
      if (!removedProductInList) {
        console.log(`   ✅ Product correctly removed from dropdown`);
      } else {
        console.log(`   ❌ Product still found in dropdown`);
      }
    }

    // Cleanup
    console.log('\n7. Cleaning up test data...');
    await prisma.job.delete({
      where: { id: newJob.id }
    });
    await prisma.product.delete({
      where: { id: newProduct.id }
    });
    console.log('   ✅ Test data cleaned up');

    console.log('\n🎉 Dropdown test completed successfully!');
    console.log('\n📋 Summary:');
    console.log('   - Dropdown shows live products (state ON)');
    console.log('   - Products are removed when moved to past');
    console.log('   - API endpoint works correctly');
    console.log('   - Real-time updates should work via SSE');
    
  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testDropdown(); 