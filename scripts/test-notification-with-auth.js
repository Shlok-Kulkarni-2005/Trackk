const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testNotificationWithAuth() {
  console.log('🧪 Testing Notification System with Authentication...\n');

  try {
    // Step 1: Create a test operator
    console.log('1. Creating test operator...');
    const operator = await prisma.operator.create({
      data: {
        phone: '+1234567890',
        username: 'TestOperator',
        profileImage: 'https://example.com/avatar.jpg'
      }
    });
    console.log(`✅ Created operator with ID: ${operator.id}`);

    // Step 2: Create a test alert
    console.log('\n2. Creating test alert...');
    const alert = await prisma.alert.create({
      data: {
        message: 'Test alert message',
        icon: '🔔',
        senderId: 1,
        recipients: {
          create: {
            operatorId: operator.id,
            read: false,
          },
        },
      },
    });
    console.log(`✅ Created alert with ID: ${alert.id}`);

    // Step 3: Check unread count for this specific operator
    console.log('\n3. Checking unread count for operator...');
    const unreadCount = await prisma.operatorAlertStatus.count({
      where: {
        operatorId: operator.id,
        read: false,
      },
    });
    console.log(`✅ Unread count for operator ${operator.id}: ${unreadCount}`);

    // Step 4: Test the API endpoint (this would normally require JWT token)
    console.log('\n4. Testing API endpoint (would require authentication)...');
    console.log('   Note: This would return 401 without proper JWT token');
    console.log('   To see notifications, you need to:');
    console.log('   1. Login as an operator at /operator-login');
    console.log('   2. Complete OTP verification');
    console.log('   3. Then notifications will appear');

    // Step 5: Show how to create a JWT token for testing
    console.log('\n5. JWT Token Structure:');
    console.log('   The system expects a JWT token with operator phone number');
    console.log('   Token payload: { phone: "+1234567890" }');
    console.log('   This token should be stored as an HTTP-only cookie named "token"');

    // Cleanup
    console.log('\n6. Cleaning up test data...');
    await prisma.operatorAlertStatus.deleteMany({
      where: { operatorId: operator.id }
    });
    await prisma.alert.delete({
      where: { id: alert.id }
    });
    await prisma.operator.delete({
      where: { id: operator.id }
    });
    console.log('✅ Test data cleaned up');

    console.log('\n🎉 Test completed!');
    console.log('\n📋 Summary:');
    console.log('   - Alerts are being created correctly');
    console.log('   - Notifications require operator authentication');
    console.log('   - To see notifications, login as an operator');
    console.log('   - Use /operator-login to authenticate');

  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the test
testNotificationWithAuth(); 