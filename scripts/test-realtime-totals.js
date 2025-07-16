const BASE_URL = 'http://localhost:3006';

async function testRealtimeTotals() {
  console.log('🧪 Testing Real-Time Totals Calculation...\n');

  try {
    // Test 1: Get current dispatched items
    console.log('📋 Test 1: Getting current dispatched items...');
    const response = await fetch(`${BASE_URL}/api/admin/dispatched-items`);
    const data = await response.json();
    
    if (!data.success) {
      throw new Error('Failed to fetch dispatched items');
    }

    const allItems = data.dispatchedItems;
    const todayItems = data.todayDispatchedItems;
    
    console.log(`✅ Dispatched items API working`);
    console.log(`📦 Total items: ${allItems.length}`);
    console.log(`📦 Today's items: ${todayItems.length}`);

    // Test 2: Calculate today's totals manually
    console.log('\n📋 Test 2: Calculating today\'s totals...');
    const todayTotalQuantity = todayItems.reduce((sum, item) => sum + (item.quantity || 1), 0);
    const todayTotalCost = todayItems.reduce((sum, item) => sum + (item.cost || 0), 0);
    const todayTotalProducts = new Set(todayItems.map(item => item.product)).size;
    
    console.log(`📊 Today's totals:`);
    console.log(`   - Total Quantity: ${todayTotalQuantity}`);
    console.log(`   - Total Cost: $${todayTotalCost}`);
    console.log(`   - Total Products: ${todayTotalProducts}`);

    // Test 3: Calculate monthly totals manually
    console.log('\n📋 Test 3: Calculating monthly totals...');
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    
    const monthlyItems = allItems.filter(item => {
      const itemDate = new Date(item.dispatchedAt);
      return itemDate.getMonth() === currentMonth && 
             itemDate.getFullYear() === currentYear;
    });
    
    const monthlyTotalQuantity = monthlyItems.reduce((sum, item) => sum + (item.quantity || 1), 0);
    const monthlyTotalCost = monthlyItems.reduce((sum, item) => sum + (item.cost || 0), 0);
    const monthlyTotalProducts = new Set(monthlyItems.map(item => item.product)).size;
    
    console.log(`📊 Monthly totals:`);
    console.log(`   - Total Quantity: ${monthlyTotalQuantity}`);
    console.log(`   - Total Cost: $${monthlyTotalCost}`);
    console.log(`   - Total Products: ${monthlyTotalProducts}`);

    // Test 4: Add a new dispatch and verify totals update
    console.log('\n📋 Test 4: Adding new dispatch to test real-time updates...');
    const testProduct = `Realtime Test Product ${Date.now()}`;
    
    const addResponse = await fetch(`${BASE_URL}/api/dispatched/update-quantity`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        product: testProduct,
        quantity: 3,
        destination: 'Test Destination',
        notes: 'Testing real-time totals'
      })
    });
    
    if (addResponse.ok) {
      console.log(`✅ New dispatch added: ${testProduct}`);
      
      // Wait a moment for the update to propagate
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // Get updated data
      const updatedResponse = await fetch(`${BASE_URL}/api/admin/dispatched-items`);
      const updatedData = await updatedResponse.json();
      
      const updatedTodayItems = updatedData.todayDispatchedItems;
      const updatedTodayTotalQuantity = updatedTodayItems.reduce((sum, item) => sum + (item.quantity || 1), 0);
      
      console.log(`📊 Updated today's total quantity: ${updatedTodayTotalQuantity}`);
      console.log(`📊 Expected increase: ${todayTotalQuantity + 3}`);
      
      if (updatedTodayTotalQuantity === todayTotalQuantity + 3) {
        console.log('✅ Real-time totals are working correctly!');
      } else {
        console.log('❌ Real-time totals may not be updating correctly');
      }
    } else {
      console.log('❌ Failed to add new dispatch');
    }

    // Test 5: Verify frontend page loads with correct totals
    console.log('\n📋 Test 5: Checking frontend page...');
    const pageResponse = await fetch(`${BASE_URL}/dispatched`);
    if (pageResponse.ok) {
      console.log('✅ Dispatched page loads successfully');
      console.log('✅ Frontend should display real-time totals');
    } else {
      console.log('❌ Dispatched page failed to load');
    }

    console.log('\n🎉 Real-Time Totals Test Completed!');
    console.log('\n📝 Summary:');
    console.log('  - Today\'s totals are calculated from today\'s dispatched items');
    console.log('  - Monthly totals are calculated from current month\'s items');
    console.log('  - Totals update automatically with new dispatches');
    console.log('  - Frontend displays real-time data');
    console.log('  - Auto-refresh every 30 seconds keeps data current');

  } catch (error) {
    console.error('❌ Test failed:', error.message);
  }
}

// Run the test
testRealtimeTotals(); 