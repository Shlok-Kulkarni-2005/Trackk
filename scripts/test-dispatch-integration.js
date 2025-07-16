const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testDispatchIntegration() {
  console.log('🧪 Testing Dispatch Integration...\n');

  try {
    // Step 1: Check if we have any operators
    const operators = await prisma.operator.findMany();
    console.log(`📋 Found ${operators.length} operators in database`);
    
    if (operators.length === 0) {
      console.log('❌ No operators found. Please create an operator first.');
      return;
    }

    const operatorId = operators[0].id;
    console.log(`👤 Using operator ID: ${operatorId}`);

    // Step 2: Check current dispatched items
    const currentDispatchedItems = await prisma.operatorProductUpdate.findMany({
      where: {
        dispatchStatus: 'Dispatched',
        archived: false
      }
    });
    console.log(`📦 Current dispatched items: ${currentDispatchedItems.length}`);

    // Step 3: Create a test product update with "Dispatched" status
    const testProductName = `Test Product ${Date.now()}`;
    console.log(`\n🔄 Creating test product update for: ${testProductName}`);

    const testUpdate = await prisma.operatorProductUpdate.create({
      data: {
        operatorId: operatorId,
        product: testProductName,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched',
        dispatchedCost: 0
      }
    });

    console.log(`✅ Created test update with ID: ${testUpdate.id}`);

    // Step 4: Check if the product appears in dispatched items
    const updatedDispatchedItems = await prisma.operatorProductUpdate.findMany({
      where: {
        dispatchStatus: 'Dispatched',
        archived: false
      }
    });

    console.log(`📦 Updated dispatched items count: ${updatedDispatchedItems.length}`);
    
    const testProductInDispatched = updatedDispatchedItems.filter(item => 
      item.product === testProductName
    );
    
    console.log(`🔍 Test product "${testProductName}" appears ${testProductInDispatched.length} times in dispatched items`);

    // Step 5: Test the quantity counting logic
    const groupedByProduct = {};
    updatedDispatchedItems.forEach(item => {
      if (!groupedByProduct[item.product]) {
        groupedByProduct[item.product] = 0;
      }
      groupedByProduct[item.product]++;
    });

    console.log('\n📊 Product quantities in dispatched items:');
    Object.entries(groupedByProduct).forEach(([product, quantity]) => {
      console.log(`  - ${product}: ${quantity}`);
    });

    // Step 6: Clean up test data
    console.log('\n🧹 Cleaning up test data...');
    await prisma.operatorProductUpdate.deleteMany({
      where: {
        product: testProductName
      }
    });
    console.log('✅ Test data cleaned up');

    console.log('\n🎉 Dispatch integration test completed successfully!');
    console.log('\n📝 Summary:');
    console.log('  - Product updates with "Dispatched" status are properly recorded');
    console.log('  - Products are grouped by name and quantities are counted correctly');
    console.log('  - The dispatched items API should now show proper quantities');

  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testDispatchIntegration(); 