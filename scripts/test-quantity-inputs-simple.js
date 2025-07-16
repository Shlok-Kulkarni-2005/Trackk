// Using built-in fetch (Node.js 18+)

async function testQuantityInputsSimple() {
  console.log('🧪 Testing Quantity Input Fixes (Backend Verification)...\n');
  
  try {
    // Test 1: Add Jobs API with large quantity
    console.log('📝 Testing Add Jobs API with quantity 350...');
    const addJobsResponse = await fetch('http://localhost:3000/api/jobs', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        productId: 1,
        machineId: 1,
        action: 'ON',
        quantity: 350
      })
    });
    
    if (addJobsResponse.ok) {
      const addJobsData = await addJobsResponse.json();
      console.log('   ✅ Add Jobs API accepts quantity 350');
      console.log(`   Created ${addJobsData.jobs?.length || 0} jobs`);
    } else {
      console.log('   ❌ Add Jobs API failed with quantity 350');
    }
    
    // Test 2: Test different quantities
    console.log('\n📝 Testing Different Quantities...');
    const testQuantities = [2, 10, 50, 100, 500, 1000];
    
    for (const quantity of testQuantities) {
      const response = await fetch('http://localhost:3000/api/jobs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          productId: 1,
          machineId: 1,
          action: 'ON',
          quantity: quantity
        })
      });
      
      if (response.ok) {
        const data = await response.json();
        console.log(`   ✅ Quantity ${quantity}: ${data.jobs?.length || 0} jobs created`);
      } else {
        console.log(`   ❌ Quantity ${quantity}: Failed`);
      }
    }
    
    // Test 3: Test operator update with large quantity
    console.log('\n📝 Testing Operator Update API with quantity 250...');
    const updateResponse = await fetch('http://localhost:3000/api/operator/update', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        id: 1,
        processSteps: {
          deburring: true,
          oiling: true,
          finalInspect: true
        },
        dispatchStatus: 'Pending',
        quantity: 250
      })
    });
    
    if (updateResponse.ok) {
      console.log('   ✅ Operator Update API accepts quantity 250');
    } else {
      console.log('   ❌ Operator Update API failed with quantity 250');
    }
    
    // Test 4: Test dispatch API with large quantity
    console.log('\n📝 Testing Dispatch API with quantity 175...');
    const dispatchResponse = await fetch('http://localhost:3000/api/admin/dispatched-items', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        product: 'Test Product',
        quantity: 175,
        destination: 'Test Destination',
        notes: 'Test Notes'
      })
    });
    
    if (dispatchResponse.ok) {
      console.log('   ✅ Dispatch API accepts quantity 175');
    } else {
      console.log('   ❌ Dispatch API failed with quantity 175');
    }
    
    console.log('\n🎉 Quantity input tests completed!');
    console.log('\n📋 Summary:');
    console.log('   - All backend APIs should accept large quantities (2, 10, 50, 100, 500, 1000)');
    console.log('   - Frontend quantity inputs should allow free typing');
    console.log('   - Constraints should only apply on blur/validation');
    console.log('   - No more "stuck at 1" issues');
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Run the test
if (require.main === module) {
  testQuantityInputsSimple();
}

module.exports = { testQuantityInputsSimple }; 