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

async function testValidationAfterOff() {
  console.log('🧪 Testing Validation After Turning OFF Quantities...\n');

  try {
    // Step 1: Create a test product with 7 ON quantities
    console.log('📋 Step 1: Creating test product with 7 ON quantities...');
    const testProduct = `Test Product Validation ${Date.now()}`;
    
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

    // Step 2: Validate that we can turn OFF 7 quantities
    console.log('\n📋 Step 2: Validating that we can turn OFF 7 quantities...');
    const validateResponse1 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 7
    });

    if (validateResponse1.status === 200) {
      if (validateResponse1.data.canSetOff) {
        console.log('✅ Validation passed for quantity 7:', validateResponse1.data.reason);
        console.log(`   Available: ${validateResponse1.data.availableQuantity}, Requested: ${validateResponse1.data.requestedQuantity}`);
      } else {
        console.log('❌ Validation failed for quantity 7:', validateResponse1.data.reason);
        return;
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse1.data.error);
      return;
    }

    // Step 3: Turn OFF all 7 quantities
    console.log('\n📋 Step 3: Turning OFF all 7 quantities...');
    const turnOffResponse = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct,
      state: 'OFF',
      stage: 'Milling',
      quantity: 7
    });

    if (turnOffResponse.status !== 201) {
      console.log('❌ Failed to turn OFF quantities:', turnOffResponse.data.error);
      return;
    }

    console.log('✅ Successfully turned OFF 7 quantities:', turnOffResponse.data.message);

    // Step 4: Wait a moment for processing
    console.log('\n📋 Step 4: Waiting for processing...');
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Step 5: Try to validate turning OFF 7 more quantities (should fail)
    console.log('\n📋 Step 5: Trying to validate turning OFF 7 more quantities (should fail)...');
    const validateResponse2 = await makeRequest('/api/validate-product-status', 'POST', {
      productName: testProduct,
      machine: 'Cutting',
      quantity: 7
    });

    if (validateResponse2.status === 200) {
      if (!validateResponse2.data.canSetOff) {
        console.log('✅ Validation correctly failed after turning OFF all quantities:', validateResponse2.data.reason);
      } else {
        console.log('❌ Validation should have failed after turning OFF all quantities');
        console.log('   This indicates the validation is stale!');
      }
    } else {
      console.log('❌ Validation request failed:', validateResponse2.data.error);
    }

    // Step 6: Check how many ON quantities are actually available
    console.log('\n📋 Step 6: Checking actual ON quantities available...');
    const jobsResponse = await makeRequest('/api/jobs');
    
    if (jobsResponse.status === 200) {
      const jobs = jobsResponse.data.jobs || [];
      const onJobs = jobs.filter(job => 
        job.product.name === testProduct && 
        job.machine.name === 'Cutting' && 
        job.state === 'ON'
      );
      
      console.log(`📊 Actual ON quantities for ${testProduct} on Cutting machine: ${onJobs.length}`);
      
      if (onJobs.length === 0) {
        console.log('✅ Correct: No ON quantities available');
      } else {
        console.log('❌ Issue: Still have ON quantities available');
      }
    } else {
      console.log('❌ Failed to get jobs:', jobsResponse.data.error);
    }

    console.log('\n🎉 Validation after OFF test completed!');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testValidationAfterOff(); 