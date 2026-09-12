const { Pool } = require('pg');

const connectionString = "postgresql://neondb_owner:npg_jvM0kACEHGV8@ep-super-sun-aodsokmz.c-2.ap-southeast-1.aws.neon.tech/neondb?sslmode=require&channel_binding=require";

async function main() {
  const pool = new Pool({ connectionString });
  
  try {
    console.log("Fetching profiles from Neon Database...");
    const res = await pool.query('SELECT * FROM erp_profiles');
    
    if (res.rows.length === 0) {
      console.log("No profiles found in database.");
      return;
    }
    
    for (const profile of res.rows) {
      console.log("\n=========================================");
      console.log(`Testing Profile: "${profile.profile_name}"`);
      console.log(`URL: ${profile.url}`);
      console.log(`Auth Type: ${profile.auth_type}`);
      
      if (profile.auth_type === 'password') {
        console.log(`Username: ${profile.username}`);
        console.log(`Password length: ${profile.password ? profile.password.length : 0}`);
        
        const loginUrl = `${profile.url.replace(/\/$/, '')}/api/method/login`;
        console.log(`Sending POST to ${loginUrl}...`);
        
        try {
          const loginResponse = await fetch(loginUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              usr: profile.username,
              pwd: profile.password
            }).toString()
          });
          
          console.log(`Login Response Status: ${loginResponse.status} ${loginResponse.statusText}`);
          console.log("Login Response Headers:");
          loginResponse.headers.forEach((val, key) => {
            console.log(`  ${key}: ${val}`);
          });
          
          const loginText = await loginResponse.text();
          console.log(`Login Response Body: ${loginText}`);
          
          const setCookies = loginResponse.headers.getSetCookie 
            ? loginResponse.headers.getSetCookie() 
            : [loginResponse.headers.get('set-cookie')].filter(Boolean);
            
          console.log(`Extracted Set-Cookie:`, setCookies);
          
          if (setCookies.length > 0) {
            const parsedCookies = setCookies.map(cookieStr => cookieStr.split(';')[0]);
            const cookieHeader = parsedCookies.join('; ');
            console.log(`Formed Cookie Header: "${cookieHeader}"`);
            
            // Try fetching Company using the cookie
            const testUrl = `${profile.url.replace(/\/$/, '')}/api/resource/Company`;
            console.log(`\nTesting API access using cookie to: ${testUrl}...`);
            const testRes = await fetch(testUrl, {
              method: 'GET',
              headers: { 'Cookie': cookieHeader }
            });
            
            console.log(`Test API Status: ${testRes.status} ${testRes.statusText}`);
            const testText = await testRes.text();
            console.log(`Test API Body: ${testText.substring(0, 500)}`);
          } else {
            console.log("✗ Failed: No cookies returned from login.");
          }
        } catch (e) {
          console.error("✗ Login Request Error:", e);
        }
      } else {
        console.log("Token auth profile, skipping password test.");
      }
    }
  } catch (err) {
    console.error("Database connection failed:", err);
  } finally {
    await pool.end();
  }
}

main();
