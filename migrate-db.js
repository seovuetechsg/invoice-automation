const { Pool } = require('pg');

const connectionString = "postgresql://neondb_owner:npg_jvM0kACEHGV8@ep-super-sun-aodsokmz.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require";

async function main() {
  const pool = new Pool({ connectionString });
  try {
    console.log("Running migration to add connection_type to erp_profiles table...");
    await pool.query(`
      ALTER TABLE erp_profiles 
      ADD COLUMN IF NOT EXISTS connection_type VARCHAR(20) DEFAULT 'proxy',
      ADD COLUMN IF NOT EXISTS warehouse VARCHAR(255);
    `);
    console.log("✓ Migration successful! connection_type column added.");
  } catch (err) {
    console.error("✗ Migration failed:", err);
  } finally {
    await pool.end();
  }
}

main();
