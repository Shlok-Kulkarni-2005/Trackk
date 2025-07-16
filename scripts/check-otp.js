const { PrismaClient } = require('@prisma/client');

const prisma = new PrismaClient();

async function checkOTP() {
  console.log('🔍 Checking OTP for Test Operator...\n');

  try {
    // Find the test operator
    const operator = await prisma.operator.findUnique({
      where: { phone: '+1234567890' }
    });

    if (!operator) {
      console.log('❌ Test operator not found. Run create-test-operator.js first.');
      return;
    }

    console.log(`✅ Found operator: ${operator.username} (${operator.phone})`);

    // Find the latest OTP for this operator
    const otpRecord = await prisma.operatorOTP.findFirst({
      where: {
        operatorId: operator.id,
        used: false,
        expiresAt: { gt: new Date() }
      },
      orderBy: { id: 'desc' }
    });

    if (otpRecord) {
      console.log('\n📱 OTP Details:');
      console.log(`   Code: ${otpRecord.code}`);
      console.log(`   Expires: ${otpRecord.expiresAt}`);
      console.log(`   Created: ${otpRecord.createdAt}`);
      
      console.log('\n🎯 To login:');
      console.log('1. Go to: http://localhost:3000/operator-login');
      console.log('2. Enter phone: +1234567890');
      console.log('3. Enter OTP: ' + otpRecord.code);
      console.log('4. You should see notifications!');
    } else {
      console.log('\n❌ No valid OTP found.');
      console.log('   You need to request an OTP first:');
      console.log('   1. Go to: http://localhost:3000/operator-login');
      console.log('   2. Enter phone: +1234567890');
      console.log('   3. Click "Send OTP"');
      console.log('   4. Then run this script again to get the code');
    }

  } catch (error) {
    console.error('❌ Error checking OTP:', error);
  } finally {
    await prisma.$disconnect();
  }
}

// Run the script
checkOTP(); 