const fetch = require('node-fetch');

const BASE_URL = 'http://localhost:3000';

async function testProcountBackend() {
  console.log('🧪 Testing Procount Backend...\n');

  try {
    // Test 1: Add a job with RFD stage and OFF status
    console.log('1. Testing job creation with RFD stage and OFF status...');
    const jobResponse = await fetch(`${BASE_URL}/api/jobs`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        machine: 'RFD',
        product: 'Test Product A',
        state: 'OFF',
        stage: 'RFD'
      })
    });

    if (jobResponse.ok) {
      const jobData = await jobResponse.json();
      console.log('✅ Job created successfully:', jobData.job);
    } else {
      const error = await jobResponse.json();
      console.log('❌ Job creation failed:', error);
    }

    // Test 2: Add another job with RFD stage and OFF status (should increment count)
    console.log('\n2. Testing second job creation (should increment count)...');
    const jobResponse2 = await fetch(`${BASE_URL}/api/jobs`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        machine: 'RFD',
        product: 'Test Product A',
        state: 'OFF',
        stage: 'RFD'
      })
    });

    if (jobResponse2.ok) {
      const jobData2 = await jobResponse2.json();
      console.log('✅ Second job created successfully:', jobData2.job);
    } else {
      const error = await jobResponse2.json();
      console.log('❌ Second job creation failed:', error);
    }

    // Test 3: Add a job with different stage (should not affect count)
    console.log('\n3. Testing job creation with different stage (should not affect count)...');
    const jobResponse3 = await fetch(`${BASE_URL}/api/jobs`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        machine: 'RFD',
        product: 'Test Product B',
        state: 'OFF',
        stage: 'Milling'
      })
    });

    if (jobResponse3.ok) {
      const jobData3 = await jobResponse3.json();
      console.log('✅ Job with different stage created successfully:', jobData3.job);
    } else {
      const error = await jobResponse3.json();
      console.log('❌ Job with different stage creation failed:', error);
    }

    // Test 4: Fetch all product counts
    console.log('\n4. Fetching all product counts...');
    const procountResponse = await fetch(`${BASE_URL}/api/procount`);
    
    if (procountResponse.ok) {
      const procountData = await procountResponse.json();
      console.log('✅ Product counts fetched successfully:');
      console.log('Total products:', procountData.productCounts.length);
      procountData.productCounts.forEach((product, index) => {
        console.log(`  ${index + 1}. ${product.product.name}: ${product.count} times`);
      });
    } else {
      const error = await procountResponse.json();
      console.log('❌ Failed to fetch product counts:', error);
    }

    // Test 5: Test direct procount API call
    console.log('\n5. Testing direct procount API call...');
    const directProcountResponse = await fetch(`${BASE_URL}/api/procount`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        productName: 'Test Product C',
        machine: 'RFD',
        stage: 'RFD',
        state: 'OFF'
      })
    });

    if (directProcountResponse.ok) {
      const directData = await directProcountResponse.json();
      console.log('✅ Direct procount API call successful:', directData);
    } else {
      const error = await directProcountResponse.json();
      console.log('❌ Direct procount API call failed:', error);
    }

    console.log('\n🎉 Procount backend testing completed!');

  } catch (error) {
    console.error('❌ Test failed with error:', error.message);
  }
}

// Run the test
testProcountBackend(); 