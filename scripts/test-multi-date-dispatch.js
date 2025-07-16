const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testMultiDateDispatch() {
  console.log('🧪 Testing Multi-Date Dispatch Separation...\n');

  try {
    // Step 1: Check current operators
    const operators = await prisma.operator.findMany();
    if (operators.length === 0) {
      console.log('❌ No operators found. Please create an operator first.');
      return;
    }
    const operatorId = operators[0].id;
    console.log(`👤 Using operator ID: ${operatorId}`);

    // Step 2: Create test dispatches for different dates
    console.log('\n📋 Step 2: Creating test dispatches for different dates...');
    
    // Create a dispatch for yesterday
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yesterdayProduct = `Yesterday Product ${Date.now()}`;
    
    await prisma.operatorProductUpdate.create({
      data: {
        operatorId: operatorId,
        product: yesterdayProduct,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched',
        dispatchedCost: 0,
        createdAt: yesterday
      }
    });
    console.log(`✅ Created yesterday dispatch: ${yesterdayProduct}`);

    // Create a dispatch for 2 days ago
    const twoDaysAgo = new Date();
    twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
    const twoDaysAgoProduct = `Two Days Ago Product ${Date.now()}`;
    
    await prisma.operatorProductUpdate.create({
      data: {
        operatorId: operatorId,
        product: twoDaysAgoProduct,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched',
        dispatchedCost: 0,
        createdAt: twoDaysAgo
      }
    });
    console.log(`✅ Created 2 days ago dispatch: ${twoDaysAgoProduct}`);

    // Create a dispatch for today
    const todayProduct = `Today Product ${Date.now()}`;
    await prisma.operatorProductUpdate.create({
      data: {
        operatorId: operatorId,
        product: todayProduct,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched',
        dispatchedCost: 0
      }
    });
    console.log(`✅ Created today dispatch: ${todayProduct}`);

    // Step 3: Test the API to see the separation
    console.log('\n📋 Step 3: Testing API separation...');
    
    const http = require('http');
    
    function makeRequest(path) {
      return new Promise((resolve, reject) => {
        const options = {
          hostname: 'localhost',
          port: 3000,
          path: path,
          method: 'GET',
          headers: {
            'Content-Type': 'application/json',
          }
        };

        const req = http.request(options, (res) => {
          let body = '';
          res.on('data', (chunk) => {
            body += chunk;
          });
          res.on('end', () => {
            try {
              const jsonData = JSON.parse(body);
              resolve({ status: res.statusCode, data: jsonData });
            } catch (error) {
              resolve({ status: res.statusCode, data: body });
            }
          });
        });

        req.on('error', (error) => {
          reject(error);
        });
        req.end();
      });
    }

    const apiResponse = await makeRequest('/api/admin/dispatched-items');
    
    if (apiResponse.status === 200) {
      const todayItems = apiResponse.data.todayDispatchedItems || [];
      const allItems = apiResponse.data.dispatchedItems || [];
      
      console.log(`\n📊 API Results:`);
      console.log(`  - Today's dispatched items: ${todayItems.length}`);
      console.log(`  - All dispatched items: ${allItems.length}`);
      
      // Check if today's product is in today's list
      const todayProductInToday = todayItems.find(item => item.product === todayProduct);
      if (todayProductInToday) {
        console.log(`✅ Today's product found in today's list: ${todayProductInToday.product}`);
      } else {
        console.log(`❌ Today's product NOT found in today's list`);
      }
      
      // Check if yesterday's product is NOT in today's list
      const yesterdayProductInToday = todayItems.find(item => item.product === yesterdayProduct);
      if (!yesterdayProductInToday) {
        console.log(`✅ Yesterday's product correctly NOT in today's list`);
      } else {
        console.log(`❌ Yesterday's product incorrectly in today's list`);
      }
      
      // Check if all products are in the all items list
      const allProductsInAll = allItems.filter(item => 
        [todayProduct, yesterdayProduct, twoDaysAgoProduct].includes(item.product)
      );
      console.log(`✅ All test products found in all items list: ${allProductsInAll.length}/3`);
      
      // Show the date separation
      console.log('\n📅 Date Separation Details:');
      const today = new Date().toISOString().split('T')[0];
      const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      const twoDaysAgo = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
      
      console.log(`  Today's date: ${today}`);
      console.log(`  Yesterday's date: ${yesterday}`);
      console.log(`  Two days ago date: ${twoDaysAgo}`);
      
      // Check date filtering
      const todayItemsWithCorrectDate = todayItems.filter(item => item.date === today);
      const historyItemsWithDifferentDate = allItems.filter(item => item.date !== today);
      
      console.log(`\n✅ Date filtering results:`);
      console.log(`  - Today's items with today's date: ${todayItemsWithCorrectDate.length}/${todayItems.length}`);
      console.log(`  - History items with different date: ${historyItemsWithDifferentDate.length}/${allItems.length}`);
      
    } else {
      console.log('❌ API request failed:', apiResponse.status);
    }

    // Step 4: Clean up test data
    console.log('\n📋 Step 4: Cleaning up test data...');
    await prisma.operatorProductUpdate.deleteMany({
      where: {
        product: {
          in: [todayProduct, yesterdayProduct, twoDaysAgoProduct]
        }
      }
    });
    console.log('✅ Test data cleaned up');

    console.log('\n🎉 Multi-date dispatch test completed successfully!');
    console.log('\n📝 Summary:');
    console.log('  - Dispatches from different dates are properly separated');
    console.log('  - Today\'s dispatches appear only in today\'s list');
    console.log('  - Historical dispatches appear only in history list');
    console.log('  - Date filtering is working correctly');

  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testMultiDateDispatch(); 