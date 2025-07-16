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

async function testAddJobsQuantity() {
  console.log('🧪 Testing Add Jobs Quantity Functionality...\n');

  try {
    // Step 1: Get initial job count
    console.log('📋 Step 1: Getting initial job count...');
    const initialJobsResponse = await makeRequest('/api/jobs');
    
    if (initialJobsResponse.status !== 200) {
      console.log('❌ Failed to get initial jobs:', initialJobsResponse.data.error);
      return;
    }

    const initialJobCount = initialJobsResponse.data.jobs?.length || 0;
    console.log(`✅ Initial job count: ${initialJobCount}`);

    // Step 2: Test adding 1 job (default quantity)
    console.log('\n📋 Step 2: Testing adding 1 job (default quantity)...');
    const testProduct1 = `Test Product Q1 ${Date.now()}`;
    
    const addJob1Response = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProduct1,
      state: 'ON',
      stage: 'Milling'
      // quantity not specified, should default to 1
    });

    if (addJob1Response.status !== 201) {
      console.log('❌ Failed to add 1 job:', addJob1Response.data.error);
      return;
    }

    console.log('✅ Successfully added 1 job:', addJob1Response.data.message);
    console.log(`   Jobs created: ${addJob1Response.data.count}`);

    // Step 3: Test adding 3 jobs (explicit quantity)
    console.log('\n📋 Step 3: Testing adding 3 jobs (explicit quantity)...');
    const testProduct3 = `Test Product Q3 ${Date.now()}`;
    
    const addJob3Response = await makeRequest('/api/jobs', 'POST', {
      machine: 'RFD',
      product: testProduct3,
      state: 'OFF',
      stage: 'RFD',
      quantity: 3
    });

    if (addJob3Response.status !== 201) {
      console.log('❌ Failed to add 3 jobs:', addJob3Response.data.error);
      return;
    }

    console.log('✅ Successfully added 3 jobs:', addJob3Response.data.message);
    console.log(`   Jobs created: ${addJob3Response.data.count}`);

    // Step 4: Test adding 5 jobs (larger quantity)
    console.log('\n📋 Step 4: Testing adding 5 jobs (larger quantity)...');
    const testProduct5 = `Test Product Q5 ${Date.now()}`;
    
    const addJob5Response = await makeRequest('/api/jobs', 'POST', {
      machine: 'CNC Turning Soft-1',
      product: testProduct5,
      state: 'ON',
      stage: 'Turning',
      quantity: 5
    });

    if (addJob5Response.status !== 201) {
      console.log('❌ Failed to add 5 jobs:', addJob5Response.data.error);
      return;
    }

    console.log('✅ Successfully added 5 jobs:', addJob5Response.data.message);
    console.log(`   Jobs created: ${addJob5Response.data.count}`);

    // Step 5: Test invalid quantity (0)
    console.log('\n📋 Step 5: Testing invalid quantity (0)...');
    const testProductInvalid = `Test Product Invalid ${Date.now()}`;
    
    const addJobInvalidResponse = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProductInvalid,
      state: 'ON',
      stage: 'Milling',
      quantity: 0
    });

    if (addJobInvalidResponse.status === 400) {
      console.log('✅ Correctly rejected invalid quantity (0):', addJobInvalidResponse.data.error);
    } else {
      console.log('❌ Should have rejected invalid quantity (0)');
    }

    // Step 6: Test invalid quantity (1001)
    console.log('\n📋 Step 6: Testing invalid quantity (1001)...');
    const addJobInvalid2Response = await makeRequest('/api/jobs', 'POST', {
      machine: 'Cutting',
      product: testProductInvalid,
      state: 'ON',
      stage: 'Milling',
      quantity: 1001
    });

    if (addJobInvalid2Response.status === 400) {
      console.log('✅ Correctly rejected invalid quantity (1001):', addJobInvalid2Response.data.error);
    } else {
      console.log('❌ Should have rejected invalid quantity (1001)');
    }

    // Step 7: Get final job count and verify
    console.log('\n📋 Step 7: Verifying final job count...');
    const finalJobsResponse = await makeRequest('/api/jobs');
    
    if (finalJobsResponse.status !== 200) {
      console.log('❌ Failed to get final jobs:', finalJobsResponse.data.error);
      return;
    }

    const finalJobCount = finalJobsResponse.data.jobs?.length || 0;
    const expectedNewJobs = 1 + 3 + 5; // 9 new jobs should be created
    const actualNewJobs = finalJobCount - initialJobCount;

    console.log(`📊 Results:`);
    console.log(`  Initial job count: ${initialJobCount}`);
    console.log(`  Final job count: ${finalJobCount}`);
    console.log(`  Expected new jobs: ${expectedNewJobs}`);
    console.log(`  Actual new jobs: ${actualNewJobs}`);

    if (actualNewJobs === expectedNewJobs) {
      console.log('✅ SUCCESS: All jobs were created correctly!');
    } else {
      console.log('❌ FAILED: Job count mismatch');
    }

    // Step 8: Check if RFD jobs appear in procount
    console.log('\n📋 Step 8: Checking if RFD jobs appear in procount...');
    const procountResponse = await makeRequest('/api/procount');
    
    if (procountResponse.status === 200) {
      const procountProducts = procountResponse.data.productCounts || [];
      const rfdProduct = procountProducts.find(p => p.product.name === testProduct3);
      
      if (rfdProduct) {
        console.log(`✅ RFD product found in procount: ${rfdProduct.product.name} (count: ${rfdProduct.count})`);
        if (rfdProduct.count === 3) {
          console.log('✅ RFD quantity is correct in procount!');
        } else {
          console.log(`❌ RFD quantity mismatch. Expected: 3, Got: ${rfdProduct.count}`);
        }
      } else {
        console.log('⚠️ RFD product not found in procount');
      }
    } else {
      console.log('❌ Failed to get procount data:', procountResponse.data.error);
    }

    console.log('\n🎉 Test completed!');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testAddJobsQuantity(); 