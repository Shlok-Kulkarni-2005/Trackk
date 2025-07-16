const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function createTestOperator() {
  console.log('🔧 Creating Test Operator Account...\n');

  try {
    // Check if test operator already exists
    let operator = await prisma.operator.findUnique({
      where: { phone: '+1234567890' }
    });

    if (operator) {
      console.log('✅ Test operator already exists:');
      console.log(`   Phone: ${operator.phone}`);
      console.log(`   Username: ${operator.username || 'Not set'}`);
      console.log(`   ID: ${operator.id}`);
    } else {
      // Create new test operator
      operator = await prisma.operator.create({
        data: {
          phone: '+1234567890',
          username: 'TestOperator',
          profileImage: 'https://example.com/avatar.jpg'
        }
      });
      console.log('✅ Created new test operator:');
      console.log(`   Phone: ${operator.phone}`);
      console.log(`   Username: ${operator.username}`);
      console.log(`   ID: ${operator.id}`);
    }

    // Create a test alert for this operator
    const alert = await prisma.alert.create({
      data: {
        message: 'Welcome! This is a test alert.',
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

    console.log('\n✅ Created test alert:');
    console.log(`   Message: ${alert.message}`);
    console.log(`   Alert ID: ${alert.id}`);

    // Check unread count
    const unreadCount = await prisma.operatorAlertStatus.count({
      where: {
        operatorId: operator.id,
        read: false,
      },
    });

    console.log(`\n📊 Current unread count: ${unreadCount}`);

    console.log('\n🎯 To see notifications:');
    console.log('1. Go to: http://localhost:3000/operator-login');
    console.log('2. Enter phone: +1234567890');
    console.log('3. Complete OTP verification');
    console.log('4. You should see a red notification badge!');

    console.log('\n📝 Note: You may need to set up Twilio for OTP delivery,');
    console.log('   or check the database for the OTP code manually.');

  } catch (error) {
    console.error('❌ Error creating test operator:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the script
createTestOperator(); 