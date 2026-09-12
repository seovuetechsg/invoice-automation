import pool, { initRebateTables } from '@/lib/db';
import { NextResponse } from 'next/server';

export async function GET() {
  try {
    await initRebateTables();
    const res = await pool.query('SELECT item_model_code FROM rebate_exception_list ORDER BY item_model_code ASC');
    const items = res.rows.map(row => row.item_model_code);
    return NextResponse.json({ success: true, data: items, count: items.length });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}

export async function POST(req) {
  try {
    await initRebateTables();
    const body = await req.json();
    const items = body.items || [];

    // Filter and sanitize item codes
    const uniqueItems = [...new Set(
      items
        .map(i => String(i || '').trim())
        .filter(i => i.length > 0)
    )];

    // Truncate existing exception list
    await pool.query('TRUNCATE TABLE rebate_exception_list RESTART IDENTITY');

    if (uniqueItems.length > 0) {
      // Bulk insert new exception item codes
      const valueStrings = uniqueItems.map((_, idx) => `($${idx + 1})`).join(', ');
      const query = `INSERT INTO rebate_exception_list (item_model_code) VALUES ${valueStrings} ON CONFLICT (item_model_code) DO NOTHING`;
      await pool.query(query, uniqueItems);
    }

    return NextResponse.json({
      success: true,
      count: uniqueItems.length,
      message: `Successfully updated exception list with ${uniqueItems.length} items.`
    });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
