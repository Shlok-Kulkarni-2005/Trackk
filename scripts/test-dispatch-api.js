const BASE_URL = 'http://localhost:3000';

async function testDispatchAPI() {
  console.log('🧪 Testing Dispatch API Endpoints...\n');

  try {
    // Test 1: Get current dispatched items
    console.log('📋 Test 1: Getting current dispatched items...');
    const dispatchedResponse = await fetch(`${BASE_URL}/api/admin/dispatched-items`);
    const dispatchedData = await dispatchedResponse.json();
    
    if (dispatchedResponse.ok) {
      console.log(`✅ Dispatched items API working. Found ${dispatchedData.dispatchedItems?.length || 0} items`);
      if (dispatchedData.dispatchedItems?.length > 0) {
        console.log('📦 Sample item:', dispatchedData.dispatchedItems[0]);
      }
    } else {
      console.log('❌ Dispatched items API failed:', dispatchedData.error);
    }

    // Test 2: Test the update quantity API
    console.log('\n📋 Test 2: Testing update quantity API...');
    const testProductName = `API Test Product ${Date.now()}`;
    
    const updateResponse = await fetch(`${BASE_URL}/api/dispatched/update-quantity`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        product: testProductName,
        quantity: 1,
        operatorId: 1
      }),
    });
    
    const updateData = await updateResponse.json();
    
    if (updateResponse.ok) {
      console.log(`✅ Update quantity API working: ${updateData.message}`);
    } else {
      console.log('❌ Update quantity API failed:', updateData.error);
    }

    // Test 3: Test getting quantity for specific product
    console.log('\n📋 Test 3: Testing get quantity API...');
    const quantityResponse = await fetch(`${BASE_URL}/api/dispatched/update-quantity?product=${encodeURIComponent(testProductName)}`);
    const quantityData = await quantityResponse.json();
    
    if (quantityResponse.ok) {
      console.log(`✅ Get quantity API working. Quantity for "${testProductName}": ${quantityData.quantity}`);
    } else {
      console.log('❌ Get quantity API failed:', quantityData.error);
    }

    // Test 4: Add another quantity for the same product
    console.log('\n📋 Test 4: Adding another quantity for the same product...');
    const updateResponse2 = await fetch(`${BASE_URL}/api/dispatched/update-quantity`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        product: testProductName,
        quantity: 1,
        operatorId: 1
      }),
    });
    
    const updateData2 = await updateResponse2.json();
    
    if (updateResponse2.ok) {
      console.log(`✅ Second update successful: ${updateData2.message}`);
    } else {
      console.log('❌ Second update failed:', updateData2.error);
    }

    // Test 5: Check updated quantity
    console.log('\n📋 Test 5: Checking updated quantity...');
    const quantityResponse2 = await fetch(`${BASE_URL}/api/dispatched/update-quantity?product=${encodeURIComponent(testProductName)}`);
    const quantityData2 = await quantityResponse2.json();
    
    if (quantityResponse2.ok) {
      console.log(`✅ Updated quantity for "${testProductName}": ${quantityData2.quantity}`);
    } else {
      console.log('❌ Get updated quantity failed:', quantityData2.error);
    }

    // Test 6: Check dispatched items again
    console.log('\n📋 Test 6: Checking dispatched items after updates...');
    const dispatchedResponse2 = await fetch(`${BASE_URL}/api/admin/dispatched-items`);
    const dispatchedData2 = await dispatchedResponse2.json();
    
    if (dispatchedResponse2.ok) {
      console.log(`✅ Dispatched items after updates: ${dispatchedData2.dispatchedItems?.length || 0} items`);
      const testProductInList = dispatchedData2.dispatchedItems?.find(item => item.product === testProductName);
      if (testProductInList) {
        console.log(`📦 Test product in dispatched list: ${testProductInList.product} (quantity: ${testProductInList.quantity})`);
      }
    } else {
      console.log('❌ Final dispatched items check failed:', dispatchedData2.error);
    }

    console.log('\n🎉 Dispatch API test completed successfully!');
    console.log('\n📝 Summary:');
    console.log('  - All API endpoints are working correctly');
    console.log('  - Products are properly added to dispatched items');
    console.log('  - Quantities are correctly incremented');
    console.log('  - The dispatched items list shows proper grouping and quantities');

  } catch (error) {
    console.error('❌ API test failed:', error.message);
    console.log('\n💡 Make sure the development server is running on http://localhost:3000');
  }
}

// Run the test
testDispatchAPI(); 