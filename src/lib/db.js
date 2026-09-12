import { Pool } from 'pg';

const connectionString = process.env.DATABASE_URL || "postgresql://neondb_owner:npg_jvM0kACEHGV8@ep-super-sun-aodsokmz.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require";

if (!global.pgPool) {
  global.pgPool = new Pool({
    connectionString,
    max: 10,
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 5000,
  });
}
const pool = global.pgPool;

export async function initRebateTables() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS rebate_exception_list (
        id SERIAL PRIMARY KEY,
        item_model_code VARCHAR(255) UNIQUE NOT NULL,
        created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
      );
    `);
    await pool.query(`
      CREATE TABLE IF NOT EXISTS rebate_records (
        id SERIAL PRIMARY KEY,
        date DATE NOT NULL,
        invoice_no VARCHAR(255) NOT NULL,
        item_model_code VARCHAR(255) NOT NULL,
        quantity NUMERIC(12, 4) NOT NULL DEFAULT 1,
        rate NUMERIC(12, 4) NOT NULL DEFAULT 0,
        real_rate NUMERIC(12, 4) NOT NULL DEFAULT 0,
        rebate_percentage NUMERIC(5, 4) NOT NULL DEFAULT 0,
        rebate_amount NUMERIC(12, 4) NOT NULL DEFAULT 0,
        company_profile VARCHAR(255) NOT NULL,
        created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
        supplier_invoice_no VARCHAR(255)
      );
    `);
    // Safely add column if it doesn't exist for existing tables
    await pool.query(`
      DO $$
      BEGIN
        IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='rebate_records' AND column_name='supplier_invoice_no') THEN
          ALTER TABLE rebate_records ADD COLUMN supplier_invoice_no VARCHAR(255);
        END IF;
      END $$;
    `);
  } catch (err) {
    console.error("Rebate tables initialization warning:", err.message);
  }
}

export default pool;
