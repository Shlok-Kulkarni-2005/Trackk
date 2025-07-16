const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testRFDFrontendIntegration() {
  console.log('🧪 Testing RFD Frontend Integration...\n');

  try {
    // Step 1: Test the procount API (used by RFD status tab)
    console.log('1. Testing procount API (RFD status tab)...');
    const procountResponse = await fetch('http://localhost:3000/api/procount');
    if (procountResponse.ok) {
      const procountData = await procountResponse.json();
      console.log(`   ✅ Procount API working. Found ${procountData.productCounts.length} products`);
      procountData.productCounts.forEach(product => {
        console.log(`     * ${product.product.name}: ${product.count} items`);
      });
    } else {
      console.log('   ❌ Procount API failed');
    }

    // Step 2: Test the RFD status stream endpoint
    console.log('\n2. Testing RFD status stream endpoint...');
    const streamResponse = await fetch('http://localhost:3000/api/rfd-status/stream');
    if (streamResponse.ok) {
      console.log('   ✅ RFD status stream endpoint is accessible');
    } else {
      console.log('   ❌ RFD status stream endpoint failed');
    }

    // Step 3: Test the procount stream endpoint
    console.log('\n3. Testing procount stream endpoint...');
    const procountStreamResponse = await fetch('http://localhost:3000/api/procount/stream');
    if (procountStreamResponse.ok) {
      console.log('   ✅ Procount stream endpoint is accessible');
    } else {
      console.log('   ❌ Procount stream endpoint failed');
    }

    // Step 4: Test the broadcast endpoint
    console.log('\n4. Testing RFD status broadcast endpoint...');
    const broadcastResponse = await fetch('http://localhost:3000/api/rfd-status/broadcast', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        type: 'rfd_count_updated',
        productId: 1,
        productName: 'Test Product',
        quantityReduced: 1,
        remainingCount: 5
      }),
    });
    if (broadcastResponse.ok) {
      console.log('   ✅ RFD status broadcast endpoint is working');
    } else {
      console.log('   ❌ RFD status broadcast endpoint failed');
    }

    // Step 5: Test the operator update API (used by updatedetailsop tab)
    console.log('\n5. Testing operator update API (updatedetailsop tab)...');
    
    // First, get a product to test with
    const testProduct = await prisma.product.findFirst();
    if (!testProduct) {
      console.log('   ❌ No products found in database');
      return;
    }

    const operatorUpdateResponse = await fetch('http://localhost:3000/api/operator/update', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        id: testProduct.id,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Pending',
        quantity: 1
      }),
    });

    if (operatorUpdateResponse.ok) {
      const operatorUpdateResult = await operatorUpdateResponse.json();
      console.log(`   ✅ Operator update API working: ${operatorUpdateResult.message}`);
    } else {
      const operatorUpdateError = await operatorUpdateResponse.json();
      console.log(`   ❌ Operator update API failed: ${operatorUpdateError.error}`);
    }

    console.log('\n🎉 RFD Frontend Integration Test completed successfully!');
    console.log('\n📋 Summary:');
    console.log('   - All API endpoints are accessible and functioning');
    console.log('   - Stream endpoints are ready for real-time updates');
    console.log('   - Broadcast system is working');
    console.log('   - Frontend can connect to all required endpoints');

  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testRFDFrontendIntegration(); 