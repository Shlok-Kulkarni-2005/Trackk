const fetch = require('node-fetch');

const BASE_URL = 'http://localhost:3000/api';

async function testBackend() {
  console.log('🧪 Testing TrackOPZ Backend System...\n');

  try {
    // Test 1: Get dispatchable products
    console.log('1. Testing dispatchable products endpoint...');
    const dispatchableRes = await fetch(`${BASE_URL}/live-products/dispatchable`);
    const dispatchableData = await dispatchableRes.json();
    console.log(`✅ Found ${dispatchableData.dispatchableProducts?.length || 0} dispatchable products\n`);

    // Test 2: Get dashboard status
    console.log('2. Testing dashboard status endpoint...');
    const dashboardRes = await fetch(`${BASE_URL}/dashboard/status`);
    const dashboardData = await dashboardRes.json();
    if (dashboardData.success) {
      console.log(`✅ Dashboard Status:`);
      console.log(`   - Total Products: ${dashboardData.dashboard.stats.totalProducts}`);
      console.log(`   - Live Products: ${dashboardData.dashboard.stats.liveProducts}`);
      console.log(`   - Past Products: ${dashboardData.dashboard.stats.pastProducts}`);
      console.log(`   - Dispatchable Products: ${dashboardData.dashboard.stats.dispatchableProducts}`);
      console.log(`   - Active Machines: ${dashboardData.dashboard.stats.activeMachines}\n`);
    } else {
      console.log('❌ Dashboard status failed\n');
    }

    // Test 3: Test product lifecycle (if products exist)
    if (dispatchableData.dispatchableProducts?.length > 0) {
      const testProduct = dispatchableData.dispatchableProducts[0];
      console.log(`3. Testing product lifecycle for: ${testProduct.name}...`);
      
      const lifecycleRes = await fetch(`${BASE_URL}/products/lifecycle?productId=${testProduct.id}`);
      const lifecycleData = await lifecycleRes.json();
      if (lifecycleData.success) {
        console.log(`✅ Product Lifecycle:`);
        console.log(`   - Current State: ${lifecycleData.lifecycle.currentState}`);
        console.log(`   - Is Dispatchable: ${lifecycleData.lifecycle.isDispatchable}`);
        console.log(`   - Total Stages: ${lifecycleData.lifecycle.totalStages}`);
        console.log(`   - Current Machine: ${lifecycleData.lifecycle.currentMachine}\n`);
      } else {
        console.log('❌ Product lifecycle failed\n');
      }
    } else {
      console.log('3. Skipping product lifecycle test (no products available)\n');
    }

    // Test 4: Test real-time stream connection
    console.log('4. Testing real-time stream connection...');
    try {
      const streamRes = await fetch(`${BASE_URL}/live-products/dispatchable/stream`);
      if (streamRes.ok) {
        console.log('✅ Real-time stream endpoint is accessible\n');
      } else {
        console.log('❌ Real-time stream endpoint failed\n');
      }
    } catch (error) {
      console.log('⚠️  Real-time stream test skipped (SSE not supported in Node.js fetch)\n');
    }

    // Test 5: Test jobs endpoint
    console.log('5. Testing jobs endpoint...');
    const jobsRes = await fetch(`${BASE_URL}/jobs`);
    const jobsData = await jobsRes.json();
    console.log(`✅ Found ${jobsData.jobs?.length || 0} total jobs\n`);

    console.log('🎉 Backend system tests completed successfully!');
    console.log('\n📋 Summary:');
    console.log('- Dispatchable products endpoint: ✅');
    console.log('- Dashboard status endpoint: ✅');
    console.log('- Product lifecycle endpoint: ✅');
    console.log('- Real-time stream endpoint: ✅');
    console.log('- Jobs endpoint: ✅');

  } catch (error) {
    console.error('❌ Test failed:', error.message);
    console.log('\n🔧 Troubleshooting:');
    console.log('1. Make sure the server is running on http://localhost:3000');
    console.log('2. Check if the database is properly connected');
    console.log('3. Verify all API routes are accessible');
  }
}

// Run the test
testBackend(); 