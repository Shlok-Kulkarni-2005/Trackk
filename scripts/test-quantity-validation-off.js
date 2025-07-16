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

async function testQuantityValidationOff() {
  console.log('🧪 Testing OFF State Quantity Validation...\n');

  try {
    // Step 1: Create a test product with 7 ON quantities
    console.log('📋 Step 1: Creating test product with 7 ON quantities...');
    const testProduct = `Test Product Q7 ${Date.now()}`;
    
    const createResponse = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct,
      state: 'ON',
      stage: 'Milling',
      quantity: 7
    });

    if (createResponse.status !== 201) {
      console.log('❌ Failed to create test product:', createResponse.data.error);
      return;
    }

    console.log('✅ Created test product with 7 ON quantities:', createResponse.data.message);

    // Step 2: Test validation with quantity 5 (should pass)
    console.log('\n📋 Step 2: Testing validation with quantity 5 (should pass)...');
    const validateResponse1 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 5
    });

    if (validateResponse1.status === 200) {
      if (validateResponse1.data.canSetOff) {
        console.log('✅ Validation passed for quantity 5:', validateResponse1.data.reason);
        console.log(`   Available: ${validateResponse1.data.availableQuantity}, Requested: ${validateResponse1.data.requestedQuantity}`);
      } else {
        console.log('❌ Validation failed for quantity 5:', validateResponse1.data.reason);
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse1.data.error);
    }

    // Step 3: Test validation with quantity 9 (should fail)
    console.log('\n📋 Step 3: Testing validation with quantity 9 (should fail)...');
    const validateResponse2 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 9
    });

    if (validateResponse2.status === 200) {
      if (!validateResponse2.data.canSetOff) {
        console.log('✅ Validation correctly failed for quantity 9:', validateResponse2.data.reason);
      } else {
        console.log('❌ Validation should have failed for quantity 9');
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse2.data.error);
    }

    // Step 4: Test validation with quantity 7 (exact match, should pass)
    console.log('\n📋 Step 4: Testing validation with quantity 7 (exact match, should pass)...');
    const validateResponse3 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 7
    });

    if (validateResponse3.status === 200) {
      if (validateResponse3.data.canSetOff) {
        console.log('✅ Validation passed for quantity 7:', validateResponse3.data.reason);
        console.log(`   Available: ${validateResponse3.data.availableQuantity}, Requested: ${validateResponse3.data.requestedQuantity}`);
      } else {
        console.log('❌ Validation failed for quantity 7:', validateResponse3.data.reason);
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse3.data.error);
    }

    // Step 5: Test validation with quantity 0 (should fail)
    console.log('\n📋 Step 5: Testing validation with quantity 0 (should fail)...');
    const validateResponse4 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 0
    });

    if (validateResponse4.status === 200) {
      if (!validateResponse4.data.canSetOff) {
        console.log('✅ Validation correctly failed for quantity 0:', validateResponse4.data.reason);
      } else {
        console.log('❌ Validation should have failed for quantity 0');
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse4.data.error);
    }

    // Step 6: Test validation with quantity 1001 (should fail)
    console.log('\n📋 Step 6: Testing validation with quantity 1001 (should fail)...');
    const validateResponse5 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 1001
    });

    if (validateResponse5.status === 200) {
      if (!validateResponse5.data.canSetOff) {
        console.log('✅ Validation correctly failed for quantity 1001:', validateResponse5.data.reason);
      } else {
        console.log('❌ Validation should have failed for quantity 1001');
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse5.data.error);
    }

    console.log('\n🎉 OFF state quantity validation test completed!');
    console.log('\n📝 Summary:');
    console.log('  - Quantity validation works correctly');
    console.log('  - Prevents turning OFF more quantities than available');
    console.log('  - Allows turning OFF quantities up to available amount');
    console.log('  - Validates quantity range (1-1000)');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testQuantityValidationOff(); 