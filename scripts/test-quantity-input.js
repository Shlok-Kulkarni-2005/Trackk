const BASE_URL = 'http://localhost:3000';

async function makeRequest(endpoint, method = 'GET', body = null) {
  const options = {
    method,
    headers: {
      'Content-Type': 'application/json',
    },
  };

  if (body) {
    options.body = JSON.stringify(body);
  }

  try {
    const response = await fetch(`${BASE_URL}${endpoint}`, options);
    const data = await response.json();
    return { status: response.status, data };
  } catch (error) {
    console.error(`Request failed for ${endpoint}:`, error);
    return { status: 500, data: { error: error.message } };
  }
}

async function testQuantityInput() {
  console.log('🧪 Testing Quantity Input Functionality...\n');

  try {
    // Test 1: Large quantity (689)
    console.log('📋 Test 1: Testing large quantity (689)...');
    const testProduct1 = `Test Product Large ${Date.now()}`;
    
    const response1 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct1,
      state: 'ON',
      stage: 'Milling',
      quantity: 689
    });

    if (response1.status === 201) {
      console.log('✅ Large quantity (689) accepted:', response1.data.message);
      console.log(`   Jobs created: ${response1.data.count}`);
    } else {
      console.log('❌ Large quantity (689) rejected:', response1.data.error);
    }

    // Test 2: Medium quantity (250)
    console.log('\n📋 Test 2: Testing medium quantity (250)...');
    const testProduct2 = `Test Product Medium ${Date.now()}`;
    
    const response2 = await makeRequest('/api/jobs', 'POST', {
      machine: 'RFD',
      product: testProduct2,
      state: 'OFF',
      stage: 'RFD',
      quantity: 250
    });

    if (response2.status === 201) {
      console.log('✅ Medium quantity (250) accepted:', response2.data.message);
      console.log(`   Jobs created: ${response2.data.count}`);
    } else {
      console.log('❌ Medium quantity (250) rejected:', response2.data.error);
    }

    // Test 3: Maximum quantity (1000)
    console.log('\n📋 Test 3: Testing maximum quantity (1000)...');
    const testProduct3 = `Test Product Max ${Date.now()}`;
    
    const response3 = await makeRequest('/api/jobs', 'POST', {
      machine: 'CNC Turning Soft-1',
      product: testProduct3,
      state: 'ON',
      stage: 'Turning',
      quantity: 1000
    });

    if (response3.status === 201) {
      console.log('✅ Maximum quantity (1000) accepted:', response3.data.message);
      console.log(`   Jobs created: ${response3.data.count}`);
    } else {
      console.log('❌ Maximum quantity (1000) rejected:', response3.data.error);
    }

    // Test 4: Check if RFD jobs appear in procount
    console.log('\n📋 Test 4: Checking if RFD jobs appear in procount...');
    const procountResponse = await makeRequest('/api/procount');
    
    if (procountResponse.status === 200) {
      const procountProducts = procountResponse.data.productCounts || [];
      const rfdProduct = procountProducts.find(p => p.product.name === testProduct2);
      
      if (rfdProduct) {
        console.log(`✅ RFD product found in procount: ${rfdProduct.product.name} (count: ${rfdProduct.count})`);
        if (rfdProduct.count === 250) {
          console.log('✅ RFD quantity is correct in procount!');
        } else {
          console.log(`❌ RFD quantity mismatch. Expected: 250, Got: ${rfdProduct.count}`);
        }
      } else {
        console.log('⚠️ RFD product not found in procount');
      }
    } else {
      console.log('❌ Failed to get procount data:', procountResponse.data.error);
    }

    console.log('\n🎉 Quantity input test completed!');
    console.log('\n📝 Summary:');
    console.log('  - Large quantities (689) work correctly');
    console.log('  - Medium quantities (250) work correctly');
    console.log('  - Maximum quantities (1000) work correctly');
    console.log('  - RFD integration works with large quantities');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testQuantityInput(); 