import pool from '@/lib/db';
import { NextResponse } from 'next/server';

export async function GET() {
  try {
    await pool.query('ALTER TABLE erp_profiles ADD COLUMN IF NOT EXISTS tax_template VARCHAR(255)');
    await pool.query('ALTER TABLE erp_profiles ADD COLUMN IF NOT EXISTS enable_sn_tracking BOOLEAN DEFAULT FALSE');
    const res = await pool.query('SELECT * FROM erp_profiles ORDER BY profile_name ASC');
    const sanitized = res.rows.map(row => ({
      ...row,
      api_key: '',
      api_secret: '',
      username: '',
      password: ''
    }));
    return NextResponse.json({ success: true, data: sanitized });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}

export async function POST(req) {
  try {
    await pool.query('ALTER TABLE erp_profiles ADD COLUMN IF NOT EXISTS tax_template VARCHAR(255)');
    await pool.query('ALTER TABLE erp_profiles ADD COLUMN IF NOT EXISTS enable_sn_tracking BOOLEAN DEFAULT FALSE');
    const body = await req.json();
    const {
      profile_name,
      url,
      auth_type,
      api_key,
      api_secret,
      username,
      password,
      company,
      expense_account,
      cost_center,
      payment_account,
      creditors_account,
      sync_doctype,
      connection_type,
      warehouse,
      tax_template,
      enable_sn_tracking
    } = body;

    if (!profile_name || !url) {
      return NextResponse.json({ success: false, error: "Missing required fields" }, { status: 400 });
    }

    const query = `
      INSERT INTO erp_profiles (
        profile_name, url, auth_type, api_key, api_secret, username, password,
        company, expense_account, cost_center, payment_account, creditors_account, sync_doctype, connection_type, warehouse, tax_template, enable_sn_tracking
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17)
      ON CONFLICT (profile_name) DO UPDATE SET
        url = EXCLUDED.url,
        auth_type = EXCLUDED.auth_type,
        api_key = EXCLUDED.api_key,
        api_secret = EXCLUDED.api_secret,
        username = EXCLUDED.username,
        password = EXCLUDED.password,
        company = EXCLUDED.company,
        expense_account = EXCLUDED.expense_account,
        cost_center = EXCLUDED.cost_center,
        payment_account = EXCLUDED.payment_account,
        creditors_account = EXCLUDED.creditors_account,
        sync_doctype = EXCLUDED.sync_doctype,
        connection_type = EXCLUDED.connection_type,
        warehouse = EXCLUDED.warehouse,
        tax_template = EXCLUDED.tax_template,
        enable_sn_tracking = EXCLUDED.enable_sn_tracking
      RETURNING *
    `;

    const values = [
      profile_name,
      url,
      auth_type || 'token',
      api_key || null,
      api_secret || null,
      username || null,
      password || null,
      company || null,
      expense_account || null,
      cost_center || null,
      payment_account || null,
      creditors_account || null,
      sync_doctype || 'Purchase Order',
      connection_type || 'proxy',
      warehouse || null,
      tax_template || null,
      enable_sn_tracking || false
    ];

    const res = await pool.query(query, values);
    return NextResponse.json({ success: true, data: res.rows[0] });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}

export async function DELETE(req) {
  try {
    const { searchParams } = new URL(req.url);
    const profile_name = searchParams.get('profile_name');

    if (!profile_name) {
      return NextResponse.json({ success: false, error: "Missing profile_name parameter" }, { status: 400 });
    }

    const res = await pool.query('DELETE FROM erp_profiles WHERE profile_name = $1 RETURNING *', [profile_name]);
    
    if (res.rowCount === 0) {
      return NextResponse.json({ success: false, error: "Profile not found" }, { status: 404 });
    }

    return NextResponse.json({ success: true, message: "Profile deleted successfully" });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
