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

async function testQuantityValidation() {
  console.log('🧪 Testing Quantity Validation...\n');

  try {
    // Test 1: Valid quantity (1)
    console.log('📋 Test 1: Valid quantity (1)...');
    const testProduct1 = `Test Product Valid1 ${Date.now()}`;
    
    const response1 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct1,
      state: 'ON',
      stage: 'Milling',
      quantity: 1
    });

    if (response1.status === 201) {
      console.log('✅ Valid quantity (1) accepted:', response1.data.message);
    } else {
      console.log('❌ Valid quantity (1) rejected:', response1.data.error);
    }

    // Test 2: Invalid quantity (0)
    console.log('\n📋 Test 2: Invalid quantity (0)...');
    const testProduct2 = `Test Product Invalid0 ${Date.now()}`;
    
    const response2 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct2,
      state: 'ON',
      stage: 'Milling',
      quantity: 0
    });

    if (response2.status === 400) {
      console.log('✅ Invalid quantity (0) correctly rejected:', response2.data.error);
    } else {
      console.log('❌ Invalid quantity (0) should have been rejected');
    }

    // Test 3: Invalid quantity (1001)
    console.log('\n📋 Test 3: Invalid quantity (1001)...');
    const testProduct3 = `Test Product Invalid1001 ${Date.now()}`;
    
    const response3 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct3,
      state: 'ON',
      stage: 'Milling',
      quantity: 1001
    });

    if (response3.status === 400) {
      console.log('✅ Invalid quantity (1001) correctly rejected:', response3.data.error);
    } else {
      console.log('❌ Invalid quantity (1001) should have been rejected');
    }

    // Test 4: Valid quantity (1000)
    console.log('\n📋 Test 4: Valid quantity (1000)...');
    const testProduct4 = `Test Product Valid1000 ${Date.now()}`;
    
    const response4 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct4,
      state: 'ON',
      stage: 'Milling',
      quantity: 1000
    });

    if (response4.status === 201) {
      console.log('✅ Valid quantity (1000) accepted:', response4.data.message);
    } else {
      console.log('❌ Valid quantity (1000) rejected:', response4.data.error);
    }

    // Test 5: No quantity specified (should default to 1)
    console.log('\n📋 Test 5: No quantity specified (should default to 1)...');
    const testProduct5 = `Test Product Default ${Date.now()}`;
    
    const response5 = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct5,
      state: 'ON',
      stage: 'Milling'
      // quantity not specified
    });

    if (response5.status === 201) {
      console.log('✅ No quantity specified accepted (defaults to 1):', response5.data.message);
    } else {
      console.log('❌ No quantity specified rejected:', response5.data.error);
    }

    console.log('\n🎉 Quantity validation test completed!');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testQuantityValidation(); 