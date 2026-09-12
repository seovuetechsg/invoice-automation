import pool from '@/lib/db';
import { NextResponse } from 'next/server';

export async function GET() {
  try {
    // Get last 100 logs
    const logsRes = await pool.query('SELECT * FROM audit_logs ORDER BY created_at DESC LIMIT 100');
    
    // Get aggregated statistics
    const statsRes = await pool.query(`
      SELECT 
        COUNT(*) as total_count,
        COALESCE(SUM(CASE WHEN action_type = 'SYNC_SUCCESS' THEN amount ELSE 0 END), 0) as total_synced_amount,
        COALESCE(SUM(CASE WHEN action_type = 'PAYMENT_CREATED' THEN amount ELSE 0 END), 0) as total_paid_amount
      FROM audit_logs
    `);

    const stats = statsRes.rows[0];
    
    return NextResponse.json({
      success: true,
      data: logsRes.rows,
      metrics: {
        total_count: parseInt(stats.total_count || 0),
        total_synced_amount: parseFloat(stats.total_synced_amount || 0),
        total_paid_amount: parseFloat(stats.total_paid_amount || 0)
      }
    });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}

export async function POST(req) {
  try {
    const body = await req.json();
    const { action_type, profile_name, supplier, amount, doc_name, details } = body;

    if (!action_type) {
      return NextResponse.json({ success: false, error: "Missing action_type" }, { status: 400 });
    }

    const query = `
      INSERT INTO audit_logs (action_type, profile_name, supplier, amount, doc_name, details)
      VALUES ($1, $2, $3, $4, $5, $6)
      RETURNING *
    `;

    const values = [
      action_type,
      profile_name || 'Default Profile',
      supplier || null,
      amount ? parseFloat(amount) : 0.0,
      doc_name || null,
      details ? JSON.stringify(details) : null
    ];

    const res = await pool.query(query, values);
    return NextResponse.json({ success: true, data: res.rows[0] });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
