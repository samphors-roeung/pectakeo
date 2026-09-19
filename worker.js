/**
 * Cloudflare Worker with Static Assets
 * Handles /api/gas proxy and serves static assets for all other routes
 */
export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // 1. API Route: /api/gas proxy to Google Apps Script
    if (url.pathname === '/api/gas') {
      const corsHeaders = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With',
        'Access-Control-Max-Age': '86400',
      };

      if (request.method === 'OPTIONS') {
        return new Response(null, { status: 204, headers: corsHeaders });
      }

      const GAS_URL = (env && env.GAS_EXEC_URL) || "https://script.google.com/macros/s/AKfycbz6tkV50ynsjBavfLrDBq34mBforpqM3bdB2FjxFuFqye31xyz_x1aGfwJYSgTaHlsk/exec";

      try {
        let body = '';
        if (request.method === 'POST') {
          body = await request.text();
        }

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
            headers: { ...corsHeaders, 'Content-Type': 'application/json;charset=utf-8' }
          });
        }

        // Catch permissions / login redirect
        if (text.includes('accounts.google.com') || text.includes('Sign in - Google Accounts')) {
          return new Response(JSON.stringify({
            __isGasError: true,
            message: "Google Apps Script ទាមទារការ Login! សូមចូលទៅ script.google.com រួចចុច Deploy > Manage deployments > Edit > កំណត់ 'Who has access' ទៅជា 'Anyone' (នរណាក៏ដោយ)។"
          }), {
            status: 200,
            headers: { ...corsHeaders, 'Content-Type': 'application/json;charset=utf-8' }
          });
        }

        try {
          JSON.parse(text);
          return new Response(text, {
            status: 200,
            headers: { ...corsHeaders, 'Content-Type': 'application/json;charset=utf-8' }
          });
        } catch (e) {
          if (text.trim().startsWith('<')) {
            return new Response(JSON.stringify({
              __isGasError: true,
              message: "Google Apps Script បានឆ្លើយតបជាទំព័រ HTML (កំហុស)៖ " + text.replace(/<[^>]*>?/gm, '').trim().substring(0, 150)
            }), {
              status: 200,
              headers: { ...corsHeaders, 'Content-Type': 'application/json;charset=utf-8' }
            });
          }
          return new Response(text, {
            status: 200,
            headers: { ...corsHeaders, 'Content-Type': 'text/plain;charset=utf-8' }
          });
        }
      } catch (error) {
        return new Response(JSON.stringify({
          __isGasError: true,
          message: "Cloudflare Edge Proxy Error: " + error.message
        }), {
          status: 500,
          headers: { ...corsHeaders, 'Content-Type': 'application/json;charset=utf-8' }
        });
      }
    }

    // 2. Static Assets fallback (SPA support)
    if (env.ASSETS) {
      const response = await env.ASSETS.fetch(request);
      if (response.status === 404 && !url.pathname.includes('.')) {
        return env.ASSETS.fetch(new Request(new URL('/index.html', request.url), request));
      }
      return response;
    }

    return new Response("Assets not configured", { status: 500 });
  }
};
