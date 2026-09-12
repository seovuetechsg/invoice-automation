import pool, { initRebateTables } from '@/lib/db';
import { NextResponse } from 'next/server';

export async function GET(req) {
  try {
    await initRebateTables();
    const { searchParams } = new URL(req.url);
    const company_profile = searchParams.get('company_profile') || '';
    const invoice_nos_str = searchParams.get('invoices') || '';

    let query = 'SELECT * FROM rebate_records';
    const params = [];
    const conditions = [];

    if (company_profile) {
      params.push(company_profile);
      conditions.push(`company_profile = $${params.length}`);
    }

    if (invoice_nos_str) {
      const invArray = invoice_nos_str
        .split(',')
        .map(i => i.trim())
        .filter(i => i.length > 0);
      
      if (invArray.length > 0) {
        params.push(invArray);
        conditions.push(`(invoice_no = ANY($${params.length}) OR supplier_invoice_no = ANY($${params.length}))`);
      }
    }

    if (conditions.length > 0) {
      query += ' WHERE ' + conditions.join(' AND ');
    }

    query += ' ORDER BY date DESC, id DESC';

    const res = await pool.query(query, params);
    return NextResponse.json({ success: true, data: res.rows, count: res.rows.length });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}

export async function POST(req) {
  try {
    await initRebateTables();
    const body = await req.json();
    const records = body.records || [];

    if (!Array.isArray(records) || records.length === 0) {
      return NextResponse.json({ success: true, count: 0, message: "No records to insert." });
    }

    // Step 1: Deduplicate - Clear existing records for the same invoice_no and company_profile
    const invoiceProfilePairs = new Set();
    records.forEach(rec => {
      if (rec.invoice_no && rec.company_profile) {
        invoiceProfilePairs.add(JSON.stringify({
          inv: String(rec.invoice_no).trim(),
          prof: String(rec.company_profile).trim()
        }));
      }
    });

    for (const pairStr of invoiceProfilePairs) {
      const { inv, prof } = JSON.parse(pairStr);
      await pool.query('DELETE FROM rebate_records WHERE company_profile = $1 AND invoice_no = $2', [prof, inv]);
    }

    // Step 2: Insert fresh rebate records
    const insertedRows = [];

    for (const rec of records) {
      const {
        date,
        invoice_no,
        item_model_code,
        quantity,
        rate,
        real_rate,
        rebate_percentage,
        rebate_amount,
        company_profile,
        supplier_invoice_no
      } = rec;

      if (!invoice_no || !item_model_code || !company_profile) continue;

      const query = `
        INSERT INTO rebate_records (
          date, invoice_no, item_model_code, quantity, rate, real_rate, rebate_percentage, rebate_amount, company_profile, supplier_invoice_no
        ) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10)
        RETURNING *
      `;

      const values = [
        date || new Date().toISOString().split('T')[0],
        String(invoice_no).trim(),
        String(item_model_code).trim(),
        parseFloat(quantity) || 1,
        parseFloat(rate) || 0,
        parseFloat(real_rate) || 0,
        parseFloat(rebate_percentage) || 0,
        parseFloat(rebate_amount) || 0,
        String(company_profile).trim(),
        supplier_invoice_no ? String(supplier_invoice_no).trim() : null
      ];

      const res = await pool.query(query, values);
      if (res.rows.length > 0) {
        insertedRows.push(res.rows[0]);
      }
    }

    return NextResponse.json({
      success: true,
      count: insertedRows.length,
      message: `Successfully saved ${insertedRows.length} rebate record(s) to Neon database (previous invoice entries updated).`
    });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}


export async function DELETE(req) {
  try {
    const { searchParams } = new URL(req.url);
    const id = searchParams.get('id');
    if (!id) return NextResponse.json({ success: false, error: 'Missing id parameter' }, { status: 400 });

    const res = await pool.query('DELETE FROM rebate_records WHERE id = $1', [id]);
    return NextResponse.json({ success: true, message: 'Deleted ' + res.rowCount + ' record(s)' });
  } catch (err) {
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
