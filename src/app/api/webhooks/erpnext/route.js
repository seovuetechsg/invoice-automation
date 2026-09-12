import pool, { initRebateTables } from '@/lib/db';
import { NextResponse } from 'next/server';

export async function POST(req) {
  try {
    await initRebateTables();
    const payload = await req.json();
    const docName = payload.name;

    if (!docName) {
      return NextResponse.json({ success: false, error: 'No document name found in webhook payload' }, { status: 400 });
    }

    const deleteQuery = 'DELETE FROM rebate_records WHERE invoice_no = $1';
    const res = await pool.query(deleteQuery, [docName]);

    console.log('Webhook triggered: Deleted ' + res.rowCount + ' rebate records for cancelled/trashed document ' + docName);

    return NextResponse.json({ 
      success: true, 
      message: 'Deleted ' + res.rowCount + ' rebate records for document ' + docName
    });
  } catch (err) {
    console.error('ERPNext Webhook Error: ' + err.message);
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
