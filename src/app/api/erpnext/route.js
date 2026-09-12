import pool from '@/lib/db';
import { NextResponse } from 'next/server';

export async function POST(req) {
  try {
    const body = await req.json();
    const { url, method, profile_name, data } = body;

    if (!url) {
      return NextResponse.json({ success: false, error: "Missing url parameter" }, { status: 400 });
    }

    let profile = null;
    if (body.profile) {
      profile = body.profile;
    } else {
      if (!profile_name) {
        return NextResponse.json({ success: false, error: "Missing profile_name" }, { status: 400 });
      }
      // Fetch profile credentials
      const profileRes = await pool.query('SELECT * FROM erp_profiles WHERE profile_name = $1', [profile_name]);
      if (profileRes.rowCount === 0) {
        return NextResponse.json({ success: false, error: `Profile '${profile_name}' not found.` }, { status: 404 });
      }
      profile = profileRes.rows[0];
    }
    const targetUrl = `${profile.url.replace(/\/$/, '')}${url}`;
    const headers = {};

    // Auth Header
    if (profile.auth_type === 'token') {
      if (!profile.api_key || !profile.api_secret) {
        return NextResponse.json({ success: false, error: "Profile missing Token API Key/Secret credentials." }, { status: 401 });
      }
      headers['Authorization'] = `token ${profile.api_key}:${profile.api_secret}`;
    } else {
      if (!profile.username || !profile.password) {
        return NextResponse.json({ success: false, error: "Profile missing Username/Password credentials." }, { status: 401 });
      }

      // Perform a session login request to ERPNext to get a session cookie (sid)
      const loginUrl = `${profile.url.replace(/\/$/, '')}/api/method/login`;
      console.log(`Performing password authentication to: ${loginUrl}`);

      // Use application/x-www-form-urlencoded to match standard browser form submit format
      const loginResponse = await fetch(loginUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          usr: profile.username,
          pwd: profile.password
        }).toString()
      });

      if (!loginResponse.ok) {
        const loginText = await loginResponse.text();
        console.error("Login request failed with status:", loginResponse.status, loginText);
        return NextResponse.json({ 
          success: false, 
          error: `ERPNext Login failed (HTTP ${loginResponse.status}): ${loginText || loginResponse.statusText}` 
        }, { status: loginResponse.status });
      }

      // Extract cookies from response (sid)
      const setCookies = loginResponse.headers.getSetCookie 
        ? loginResponse.headers.getSetCookie() 
        : [loginResponse.headers.get('set-cookie')].filter(Boolean);

      if (setCookies.length > 0) {
        // Strip out cookie metadata attributes (Path, Secure, HttpOnly, etc.)
        const parsedCookies = setCookies.map(cookieStr => cookieStr.split(';')[0]);
        headers['Cookie'] = parsedCookies.join('; ');
        console.log(`✓ Login success. Attached session cookie headers: ${headers['Cookie']}`);
      } else {
        return NextResponse.json({ success: false, error: "Authentication succeeded but no session cookies were returned by ERPNext." }, { status: 401 });
      }
    }


    if (method && method !== 'GET') {
      headers['Content-Type'] = 'application/json';
    }

    const fetchOptions = {
      method: method || 'GET',
      headers: headers
    };

    if (data && method && method !== 'GET') {
      fetchOptions.body = JSON.stringify(data);
    }

    console.log(`Proxying request: ${fetchOptions.method} ${targetUrl}`);
    const response = await fetch(targetUrl, fetchOptions);
    const responseText = await response.text();

    let responseData;
    try {
      responseData = JSON.parse(responseText);
    } catch (e) {
      responseData = { message: responseText };
    }

    return NextResponse.json(responseData, { status: response.status });

  } catch (err) {
    console.error("Proxy error:", err);
    return NextResponse.json({ success: false, error: err.message }, { status: 500 });
  }
}
