const { Client } = require('pg');

// Set your database connection string here
const connectionString = 'postgresql://neondb_owner:npg_juMfEXq4g0Ym@ep-rapid-sky-a1kts7sw-pooler.ap-southeast-1.aws.neon.tech/neondb?sslmode=require';

const client = new Client({
  connectionString,
});

async function addUpdatedAtColumn() {
  try {
    await client.connect();
    console.log('Connected to database.');
    // Check if the column already exists
    const checkRes = await client.query(`
      SELECT column_name FROM information_schema.columns 
      WHERE table_name='Job' AND column_name='updatedAt';
    `);
    if (checkRes.rows.length > 0) {
      console.log('updatedAt column already exists.');
      return;
    }
    // Add the column
    await client.query('ALTER TABLE "Job" ADD COLUMN "updatedAt" TIMESTAMP DEFAULT NOW();');
    console.log('updatedAt column added to Job table.');
  } catch (err) {
    console.error('Error adding updatedAt column:', err);
  } finally {
    await client.end();
  }
}

addUpdatedAtColumn(); 