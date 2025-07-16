async function testAuthentication() {
  console.log('🧪 Testing Authentication System...\n');

  try {
    // Test 1: Check operator-alerts endpoint without authentication
    console.log('1. Testing operator-alerts endpoint without authentication...');
    const response1 = await fetch('http://localhost:3007/api/operator-alerts');
    console.log(`   Status: ${response1.status}`);
    if (response1.status === 401) {
      console.log('   ✅ Correctly returns 401 for unauthenticated requests\n');
    } else {
      console.log('   ❌ Expected 401 but got different status\n');
    }

    // Test 2: Check SSE endpoint without authentication
    console.log('2. Testing SSE endpoint without authentication...');
    const response2 = await fetch('http://localhost:3007/api/alerts/notifications');
    console.log(`   Status: ${response2.status}`);
    if (response2.status === 401) {
      console.log('   ✅ Correctly returns 401 for unauthenticated requests\n');
    } else {
      console.log('   ❌ Expected 401 but got different status\n');
    }

    // Test 3: Check ping endpoint (should work without auth)
    console.log('3. Testing ping endpoint (should work without auth)...');
    const response3 = await fetch('http://localhost:3007/api/ping');
    console.log(`   Status: ${response3.status}`);
    if (response3.status === 200) {
      console.log('   ✅ Correctly returns 200 for public endpoint\n');
    } else {
      console.log('   ❌ Expected 200 but got different status\n');
    }

    console.log('🎉 Authentication tests completed!');

  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

testAuthentication(); 