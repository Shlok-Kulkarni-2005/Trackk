const puppeteer = require('puppeteer');

async function testAllQuantityInputs() {
  console.log('🧪 Testing All Quantity Input Fields Across Project...\n');
  
  const browser = await puppeteer.launch({ 
    headless: false, 
    defaultViewport: null,
    args: ['--start-maximized']
  });
  
  try {
    const page = await browser.newPage();
    
    // Test 1: Add Jobs Page Quantity Input
    console.log('📝 Testing Add Jobs Page Quantity Input...');
    await page.goto('http://localhost:3000/addjobs');
    await page.waitForTimeout(2000);
    
    // Find and test quantity input
    const addJobsQuantityInput = await page.$('input[type="number"][min="1"]');
    if (addJobsQuantityInput) {
      await addJobsQuantityInput.click();
      await addJobsQuantityInput.type('350');
      await page.waitForTimeout(500);
      
      const addJobsValue = await addJobsQuantityInput.evaluate(el => el.value);
      console.log(`   Add Jobs quantity input value: ${addJobsValue}`);
      
      if (addJobsValue === '350') {
        console.log('   ✅ Add Jobs quantity input works correctly');
      } else {
        console.log('   ❌ Add Jobs quantity input failed');
      }
    } else {
      console.log('   ⚠️  Add Jobs quantity input not found');
    }
    
    // Test 2: Updated Details OP Page Quantity Input
    console.log('\n📝 Testing Updated Details OP Page Quantity Input...');
    await page.goto('http://localhost:3000/updatedetailsop');
    await page.waitForTimeout(2000);
    
    // Find and test quantity input
    const updateDetailsQuantityInput = await page.$('input[type="number"][min="1"]');
    if (updateDetailsQuantityInput) {
      await updateDetailsQuantityInput.click();
      await updateDetailsQuantityInput.type('250');
      await page.waitForTimeout(500);
      
      const updateDetailsValue = await updateDetailsQuantityInput.evaluate(el => el.value);
      console.log(`   Updated Details quantity input value: ${updateDetailsValue}`);
      
      if (updateDetailsValue === '250') {
        console.log('   ✅ Updated Details quantity input works correctly');
      } else {
        console.log('   ❌ Updated Details quantity input failed');
      }
    } else {
      console.log('   ⚠️  Updated Details quantity input not found');
    }
    
    // Test 3: Dispatch Page Quantity Input
    console.log('\n📝 Testing Dispatch Page Quantity Input...');
    await page.goto('http://localhost:3000/dispatch');
    await page.waitForTimeout(2000);
    
    // Find and test quantity input
    const dispatchQuantityInput = await page.$('input[type="number"][min="1"]');
    if (dispatchQuantityInput) {
      await dispatchQuantityInput.click();
      await dispatchQuantityInput.type('175');
      await page.waitForTimeout(500);
      
      const dispatchValue = await dispatchQuantityInput.evaluate(el => el.value);
      console.log(`   Dispatch quantity input value: ${dispatchValue}`);
      
      if (dispatchValue === '175') {
        console.log('   ✅ Dispatch quantity input works correctly');
      } else {
        console.log('   ❌ Dispatch quantity input failed');
      }
    } else {
      console.log('   ⚠️  Dispatch quantity input not found');
    }
    
    // Test 4: Test different number ranges
    console.log('\n📝 Testing Different Number Ranges...');
    
    const testNumbers = [2, 10, 50, 100, 500, 1000];
    
    for (const testNum of testNumbers) {
      await page.goto('http://localhost:3000/addjobs');
      await page.waitForTimeout(1000);
      
      const quantityInput = await page.$('input[type="number"][min="1"]');
      if (quantityInput) {
        await quantityInput.click();
        await quantityInput.type(testNum.toString());
        await page.waitForTimeout(500);
        
        const value = await quantityInput.evaluate(el => el.value);
        console.log(`   Testing ${testNum}: ${value === testNum.toString() ? '✅' : '❌'} (${value})`);
      }
    }
    
    console.log('\n🎉 All quantity input tests completed!');
    console.log('\n📋 Summary:');
    console.log('   - Add Jobs Page: Quantity input should allow large numbers');
    console.log('   - Updated Details OP Page: Quantity input should allow large numbers');
    console.log('   - Dispatch Page: Quantity input should allow large numbers');
    console.log('   - All inputs should accept numbers like 2, 10, 50, 100, 500, 1000');
    
  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await browser.close();
  }
}

// Run the test
if (require.main === module) {
  testAllQuantityInputs();
}

module.exports = { testAllQuantityInputs }; 