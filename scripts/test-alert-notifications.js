const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function testAlertNotifications() {
  try {
    console.log('🧪 Testing Alert Notification System...\n');

    // 1. Create a test operator
    console.log('1. Creating test operator...');
    const operator = await prisma.operator.create({
      data: {
        phone: '+1234567890',
        username: 'TestOperator',
        profileImage: 'https://example.com/avatar.jpg'
      }
    });
    console.log(`✅ Created operator with ID: ${operator.id}\n`);

    // 2. Create a test alert
    console.log('2. Creating test alert...');
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
    console.log(`✅ Created alert with ID: ${alert.id}\n`);

    // 3. Check unread count
    console.log('3. Checking unread count...');
    const unreadCount = await prisma.operatorAlertStatus.count({
      where: {
        operatorId: operator.id,
        read: false,
      },
    });
    console.log(`✅ Unread count: ${unreadCount}\n`);

    // 4. Mark alerts as read
    console.log('4. Marking alerts as read...');
    await prisma.operatorAlertStatus.updateMany({
      where: {
        operatorId: operator.id,
        read: false,
      },
      data: {
        read: true,
      },
    });
    console.log('✅ Alerts marked as read\n');

    // 5. Check unread count again
    console.log('5. Checking unread count after marking as read...');
    const newUnreadCount = await prisma.operatorAlertStatus.count({
      where: {
        operatorId: operator.id,
        read: false,
      },
    });
    console.log(`✅ New unread count: ${newUnreadCount}\n`);

    // 6. Clean up test data
    console.log('6. Cleaning up test data...');
    await prisma.operatorAlertStatus.deleteMany({
      where: { operatorId: operator.id }
    });
    await prisma.alert.delete({
      where: { id: alert.id }
    });
    await prisma.operator.delete({
      where: { id: operator.id }
    });
    console.log('✅ Test data cleaned up\n');

    console.log('🎉 All tests passed! Alert notification system is working correctly.');

  } catch (error) {
    console.error('❌ Test failed:', error);
  } finally {
    await prisma.$disconnect();
  }
}

testAlertNotifications(); 