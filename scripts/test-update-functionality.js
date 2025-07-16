// Using built-in fetch API (Node.js 18+)

async function testUpdateFunctionality() {
  const baseUrl = 'http://localhost:3002';
  
  console.log('🧪 Testing Update Functionality...\n');
  
  try {
    // 1. Test fetching finished products
    console.log('1. Fetching finished products...');
    const productsResponse = await fetch(`${baseUrl}/api/finished-products/dispatchable`);
    const productsData = await productsResponse.json();
    
    if (productsResponse.ok) {
      console.log(`✅ Found ${productsData.finishedProducts.length} products`);
      console.log('Sample products:');
      productsData.finishedProducts.slice(0, 3).forEach((product, index) => {
        console.log(`   ${index + 1}. ${product.name} (ID: ${product.id}, Source: ${product.source})`);
      });
    } else {
      console.log('❌ Failed to fetch products');
      return;
    }
    
    // 2. Test updating a live product (job_ prefixed ID)
    console.log('\n2. Testing update with live product (job_ prefixed ID)...');
    const liveProduct = productsData.finishedProducts.find(p => p.id.startsWith('job_'));
    
    if (liveProduct) {
      const updateData = {
        id: liveProduct.id,
        processSteps: {
          deburring: true,
          finalInspect: true,
          oiling: true
        },
        dispatchStatus: 'Dispatched'
      };
      
      const updateResponse = await fetch(`${baseUrl}/api/operator/update`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updateData)
      });
      
      const updateResult = await updateResponse.json();
      
      if (updateResponse.ok && updateResult.success) {
        console.log(`✅ Successfully updated live product: ${liveProduct.name}`);
        console.log(`   Created OperatorProductUpdate record with ID: ${updateResult.update.id}`);
      } else {
        console.log(`❌ Failed to update live product: ${updateResult.error}`);
      }
    } else {
      console.log('⚠️ No live products found to test');
    }
    
    // 3. Test updating a finished product (numeric ID)
    console.log('\n3. Testing update with finished product (numeric ID)...');
    const finishedProduct = productsData.finishedProducts.find(p => !p.id.startsWith('job_'));
    
    if (finishedProduct) {
      const updateData = {
        id: finishedProduct.id,
        processSteps: {
          deburring: false,
          finalInspect: true,
          oiling: false
        },
        dispatchStatus: 'Pending'
      };
      
      const updateResponse = await fetch(`${baseUrl}/api/operator/update`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updateData)
      });
      
      const updateResult = await updateResponse.json();
      
      if (updateResponse.ok && updateResult.success) {
        console.log(`✅ Successfully updated finished product: ${finishedProduct.name}`);
        console.log(`   Updated OperatorProductUpdate record with ID: ${updateResult.update.id}`);
      } else {
        console.log(`❌ Failed to update finished product: ${updateResult.error}`);
      }
    } else {
      console.log('⚠️ No finished products found to test');
    }
    
    // 4. Test error handling with invalid ID
    console.log('\n4. Testing error handling with invalid ID...');
    const invalidUpdateData = {
      id: 'invalid_id',
      processSteps: {
        deburring: true,
        finalInspect: true,
        oiling: true
      },
      dispatchStatus: 'Dispatched'
    };
    
    const invalidResponse = await fetch(`${baseUrl}/api/operator/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(invalidUpdateData)
    });
    
    const invalidResult = await invalidResponse.json();
    
    if (!invalidResponse.ok) {
      console.log(`✅ Correctly handled invalid ID: ${invalidResult.error}`);
    } else {
      console.log('❌ Should have failed with invalid ID');
    }
    
    console.log('\n🎉 Update functionality test completed successfully!');
    
  } catch (error) {
    console.error('❌ Test failed with error:', error.message);
  }
}

// Run the test
testUpdateFunctionality(); 