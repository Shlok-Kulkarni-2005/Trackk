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

async function testDispatchQuantityFix() {
  console.log('🧪 Testing Dispatch Quantity Fix...\n');

  try {
    // Step 1: Get initial procount data
    console.log('📋 Step 1: Getting initial procount data...');
    const initialProcountResponse = await makeRequest('/api/procount');
    
    if (initialProcountResponse.status !== 200) {
      console.log('❌ Failed to get initial procount data:', initialProcountResponse.data.error);
      return;
    }

    const initialProducts = initialProcountResponse.data.productCounts;
    console.log(`✅ Initial procount data loaded. Found ${initialProducts.length} products`);
    
    if (initialProducts.length === 0) {
      console.log('⚠️ No products found in procount. Cannot test dispatch.');
      return;
    }

    // Find a product with quantity > 1 for testing
    const testProduct = initialProducts.find(p => p.count > 1);
    if (!testProduct) {
      console.log('⚠️ No products with quantity > 1 found. Cannot test dispatch.');
      return;
    }

    console.log(`📦 Selected test product: ${testProduct.product.name} (quantity: ${testProduct.count})`);

    // Step 2: Dispatch 1 quantity
    console.log('\n📋 Step 2: Dispatching 1 quantity...');
    const dispatchResponse = await makeRequest('/api/operator/update', 'POST', {
      id: testProduct.id,
      processSteps: {
        deburring: true,
        finalInspect: true,
        oiling: true
      },
      dispatchStatus: 'Pending',
      quantity: 1
    });

    if (dispatchResponse.status !== 200) {
      console.log('❌ Dispatch failed:', dispatchResponse.data.error);
      return;
    }

    console.log('✅ Dispatch successful:', dispatchResponse.data.message);

    // Step 3: Wait a moment for processing
    console.log('\n📋 Step 3: Waiting for processing...');
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Step 4: Get updated procount data
    console.log('\n📋 Step 4: Getting updated procount data...');
    const updatedProcountResponse = await makeRequest('/api/procount');
    
    if (updatedProcountResponse.status !== 200) {
      console.log('❌ Failed to get updated procount data:', updatedProcountResponse.data.error);
      return;
    }

    const updatedProducts = updatedProcountResponse.data.productCounts;
    const updatedTestProduct = updatedProducts.find(p => p.id === testProduct.id);

    if (!updatedTestProduct) {
      console.log('⚠️ Test product not found in updated data. It may have been completely removed.');
      console.log('✅ This is expected if the quantity was reduced to 0.');
      return;
    }

    // Step 5: Calculate the difference
    const quantityReduced = testProduct.count - updatedTestProduct.count;
    console.log(`\n📊 Results:`);
    console.log(`  Initial quantity: ${testProduct.count}`);
    console.log(`  Updated quantity: ${updatedTestProduct.count}`);
    console.log(`  Quantity reduced: ${quantityReduced}`);

    if (quantityReduced === 1) {
      console.log('✅ SUCCESS: Only 1 quantity was reduced as expected!');
    } else if (quantityReduced === 2) {
      console.log('❌ FAILED: 2 quantities were reduced (the bug is still present)');
    } else {
      console.log(`⚠️ UNEXPECTED: ${quantityReduced} quantities were reduced`);
    }

    // Step 6: Check dispatched items
    console.log('\n📋 Step 6: Checking dispatched items...');
    const dispatchedResponse = await makeRequest('/api/admin/dispatched-items');
    
    if (dispatchedResponse.status === 200) {
      const dispatchedItems = dispatchedResponse.data.dispatchedItems || [];
      const testProductDispatched = dispatchedItems.find(item => item.product === testProduct.product.name);
      
      if (testProductDispatched) {
        console.log(`✅ Product found in dispatched items: ${testProductDispatched.product} (quantity: ${testProductDispatched.quantity})`);
      } else {
        console.log('⚠️ Product not found in dispatched items');
      }
    } else {
      console.log('❌ Failed to get dispatched items:', dispatchedResponse.data.error);
    }

    console.log('\n🎉 Test completed!');

  } catch (error) {
    console.error('❌ Test failed with error:', error);
  }
}

// Run the test
testDispatchQuantityFix(); 