/**
 * Cloudflare Worker with Static Assets & Hybrid D1 / Edge Cache Engine
 * Handles /api/gas proxy, accelerates reads via D1/Cache, dual-writes to Google Sheets + D1
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With',
  'Access-Control-Max-Age': '86400',
};

// Cache TTLs in seconds for read actions
const CACHE_RULES = {
  'getAuthPageInfo': 300,        // 5 min
  'getSystemSettings': 180,      // 3 min
  'getDashboardStats': 60,       // 1 min
  'getDailyFormBootstrap': 120,  // 2 min
  'getCommuneSummaryReportData': 60,
  'getProvinceSummaryReportData': 60,
  'getAvailableDates': 120,
};

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // 1. API Route: /api/gas proxy to Google Apps Script with D1 / Cache Acceleration
    if (url.pathname === '/api/gas') {
      if (request.method === 'OPTIONS') {
        return new Response(null, { status: 204, headers: CORS_HEADERS });
      }

      const GAS_URL = (env && env.GAS_EXEC_URL) || "https://script.google.com/macros/s/AKfycbz6tkV50ynsjBavfLrDBq34mBforpqM3bdB2FjxFuFqye31xyz_x1aGfwJYSgTaHlsk/exec";

      try {
        let body = '';
        if (request.method === 'POST') {
          body = await request.text();
        }

        let parsedRequest = null;
        try {
          if (body) parsedRequest = JSON.parse(body);
        } catch (_) {}

        const action = parsedRequest ? parsedRequest.action : null;
        const args = (parsedRequest && parsedRequest.args) ? parsedRequest.args : [];

        // Check if D1 database binding 'DB' is available
        const hasD1 = !!(env && env.DB && typeof env.DB.prepare === 'function');

        // --- READ ACCELERATION (D1 / Edge Cache) ---
        if (hasD1 && action && CACHE_RULES[action]) {
          try {
            const cacheKey = `${action}:${JSON.stringify(args)}`;
            const now = Math.floor(Date.now() / 1000);
            const cached = await env.DB.prepare(
              "SELECT cache_value FROM system_cache WHERE cache_key = ? AND expires_at > ?"
            ).bind(cacheKey, now).first();

            if (cached && cached.cache_value) {
              return new Response(cached.cache_value, {
                status: 200,
                headers: {
                  ...CORS_HEADERS,
                  'Content-Type': 'application/json;charset=utf-8',
                  'X-Data-Source': 'Cloudflare-D1-Cache',
                }
              });
            }
          } catch (d1Err) {
            console.warn("[D1 Cache Read Warning]:", d1Err.message);
          }
        }

        // --- FORWARD TO GOOGLE APPS SCRIPT (MASTER) ---
        const gasResponse = await fetch(GAS_URL, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8'
          },
          body: body,
          redirect: 'follow'
        });

        const text = await gasResponse.text();

        // Catch missing doPost on Google Apps Script
        if (text.includes('អនុគមន៏ស្គ្រីបរកមិនឃើញ') || text.includes('Script function not found') || text.includes('doPost')) {
          return new Response(JSON.stringify({
            __isGasError: true,
            message: "អនុគមន៏ស្គ្រីបរកមិនឃើញ៖ doPost! សូមចូលទៅ Google Apps Script (script.google.com) រួចចុច Deploy > Manage deployments > Edit > ជ្រើសរើស Version 'New version' > Deploy ដើម្បីឱ្យ Web App ទទួលស្គាល់កូដ doPost(e)។"
          }), {
            status: 200,
            headers: { ...CORS_HEADERS, 'Content-Type': 'application/json;charset=utf-8' }
          });
        }

        // Catch permissions / login redirect
        if (text.includes('accounts.google.com') || text.includes('Sign in - Google Accounts')) {
          return new Response(JSON.stringify({
            __isGasError: true,
            message: "Google Apps Script ទាមទារការ Login! សូមចូលទៅ script.google.com រួចចុច Deploy > Manage deployments > Edit > កំណត់ 'Who has access' ទៅជា 'Anyone' (នរណាក៏ដោយ)។"
          }), {
            status: 200,
            headers: { ...CORS_HEADERS, 'Content-Type': 'application/json;charset=utf-8' }
          });
        }

        // Try parsing response to verify valid JSON
        let responseData = null;
        let isJson = false;
        try {
          responseData = JSON.parse(text);
          isJson = true;
        } catch (_) {}

        // --- ASYNC D1 SYNC & CACHE UPDATE ---
        if (hasD1 && isJson && responseData && !responseData.__isGasError) {
          const syncTask = async () => {
            try {
              const now = Math.floor(Date.now() / 1000);

              // 1. Update Read Cache if applicable
              if (action && CACHE_RULES[action]) {
                const ttl = CACHE_RULES[action];
                const cacheKey = `${action}:${JSON.stringify(args)}`;
                await env.DB.prepare(
                  `INSERT INTO system_cache (cache_key, cache_value, updated_at, expires_at)
                   VALUES (?, ?, ?, ?)
                   ON CONFLICT(cache_key) DO UPDATE SET
                   cache_value = excluded.cache_value,
                   updated_at = excluded.updated_at,
                   expires_at = excluded.expires_at`
                ).bind(cacheKey, text, now, now + ttl).run();
              }

              // 2. Invalidate cache on write operations
              if (action && (action.startsWith('save') || action.startsWith('update') || action === 'approveUser' || action === 'deleteUser')) {
                await env.DB.prepare(
                  "DELETE FROM system_cache WHERE cache_key LIKE 'getDashboardStats%' OR cache_key LIKE 'getCommuneSummary%' OR cache_key LIKE 'getProvinceSummary%' OR cache_key LIKE 'getDailyReportData%'"
                ).run();

                // If saveDistrictDayEntries, record in daily_entries table as mirror
                if (action === 'saveDistrictDayEntries' && args.length >= 4) {
                  const currentUsername = args[0];
                  const date = args[1];
                  const district = args[2];
                  const rows = args[3] || [];

                  for (const row of rows) {
                    await env.DB.prepare(
                      `INSERT INTO daily_entries (entry_date, district, commune, values_json, note, created_by, created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
                    ).bind(
                      date, district, row.commune, JSON.stringify(row.values || {}),
                      row.note || '', currentUsername, now, now
                    ).run();
                  }

                  // Log to audit log
                  await env.DB.prepare(
                    `INSERT INTO audit_logs (action, username, details, status, timestamp)
                     VALUES (?, ?, ?, 'SUCCESS', ?)`
                  ).bind(
                    action, currentUsername, `Saved ${rows.length} communes for ${district} on ${date}`, now
                  ).run();
                }
              }
            } catch (syncErr) {
              console.warn("[D1 Sync Warning]:", syncErr.message);
            }
          };

          if (ctx && typeof ctx.waitUntil === 'function') {
            ctx.waitUntil(syncTask());
          } else {
            await syncTask();
          }
        }

        if (isJson) {
          return new Response(text, {
            status: 200,
            headers: {
              ...CORS_HEADERS,
              'Content-Type': 'application/json;charset=utf-8',
              'X-Data-Source': 'Google-Apps-Script',
            }
          });
        }

        if (text.trim().startsWith('<')) {
          return new Response(JSON.stringify({
            __isGasError: true,
            message: "Google Apps Script បានឆ្លើយតបជាទំព័រ HTML (កំហុស)៖ " + text.replace(/<[^>]*>?/gm, '').trim().substring(0, 150)
          }), {
            status: 200,
            headers: { ...CORS_HEADERS, 'Content-Type': 'application/json;charset=utf-8' }
          });
        }

        return new Response(text, {
          status: 200,
          headers: { ...CORS_HEADERS, 'Content-Type': 'text/plain;charset=utf-8' }
        });

      } catch (error) {
        return new Response(JSON.stringify({
          __isGasError: true,
          message: "Cloudflare Edge Proxy Error: " + error.message
        }), {
          status: 500,
          headers: { ...CORS_HEADERS, 'Content-Type': 'application/json;charset=utf-8' }
        });
      }
    }

    // 2. Static Assets fallback (SPA support)
    if (env && env.ASSETS) {
      const response = await env.ASSETS.fetch(request);
      if (response.status === 404 && !url.pathname.includes('.')) {
        return env.ASSETS.fetch(new Request(new URL('/index.html', request.url), request));
      }
      return response;
    }

    return new Response("Assets not configured", { status: 500 });
  }
};
