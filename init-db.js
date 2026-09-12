const { Client } = require('pg');

const connectionString = "postgresql://neondb_owner:npg_jvM0kACEHGV8@ep-super-sun-aodsokmz.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require";

async function main() {
  const client = new Client({
    connectionString: connectionString,
  });

  try {
    console.log("Connecting to Neon Database...");
    await client.connect();
    console.log("Connected successfully!");

    console.log("Creating erp_profiles table...");
    await client.query(`
      CREATE TABLE IF NOT EXISTS erp_profiles (
          id SERIAL PRIMARY KEY,
          profile_name VARCHAR(100) UNIQUE NOT NULL,
          url VARCHAR(255) NOT NULL,
          auth_type VARCHAR(20) DEFAULT 'token',
          api_key VARCHAR(100),
          api_secret VARCHAR(100),
          username VARCHAR(100),
          password VARCHAR(255),
          company VARCHAR(100),
          expense_account VARCHAR(100),
          cost_center VARCHAR(100),
          payment_account VARCHAR(100),
          creditors_account VARCHAR(100),
          sync_doctype VARCHAR(50) DEFAULT 'Purchase Invoice',
          connection_type VARCHAR(20) DEFAULT 'proxy',
          warehouse VARCHAR(255)
      );
    `);

    console.log("Creating audit_logs table...");
    await client.query(`
      CREATE TABLE IF NOT EXISTS audit_logs (
          id SERIAL PRIMARY KEY,
          created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
          action_type VARCHAR(50) NOT NULL,
          profile_name VARCHAR(100),
          supplier VARCHAR(100),
          amount NUMERIC(15, 2) DEFAULT 0.0,
          doc_name VARCHAR(100),
          details JSONB
      );
    `);

    console.log("✓ Tables created successfully!");
  } catch (err) {
    console.error("✗ Initialization failed:", err);
  } finally {
    await client.end();
  }
}

main();
