module.exports = async (req, res) => {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const GAS_URL = process.env.GAS_EXEC_URL || "https://script.google.com/macros/s/AKfycbz6tkV50ynsjBavfLrDBq34mBforpqM3bdB2FjxFuFqye31xyz_x1aGfwJYSgTaHlsk/exec";

  try {
    let body = req.body;
    if (typeof body !== 'string') {
      body = JSON.stringify(body || {});
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
      return res.status(200).json({
        __isGasError: true,
        message: "អនុគមន៏ស្គ្រីបរកមិនឃើញ៖ doPost! សូមចូលទៅ Google Apps Script (script.google.com) រួចចុច Deploy > Manage deployments > Edit > ជ្រើសរើស Version 'New version' > Deploy ដើម្បីឱ្យ Web App ទទួលស្គាល់កូដ doPost(e)។"
      });
    }

    // Catch permissions / login redirect
    if (text.includes('accounts.google.com') || text.includes('Sign in - Google Accounts')) {
      return res.status(200).json({
        __isGasError: true,
        message: "Google Apps Script ទាមទារការ Login! សូមចូលទៅ script.google.com រួចចុច Deploy > Manage deployments > Edit > កំណត់ 'Who has access' ទៅជា 'Anyone' (នរណាក៏ដោយ)។"
      });
    }

    try {
      const data = JSON.parse(text);
      return res.status(200).json(data);
    } catch (e) {
      if (text.trim().startsWith('<')) {
        return res.status(200).json({
          __isGasError: true,
          message: "Google Apps Script បានឆ្លើយតបជាទំព័រ HTML (កំហុស)៖ " + text.replace(/<[^>]*>?/gm, '').trim().substring(0, 150)
        });
      }
      return res.status(200).send(text);
    }
  } catch (error) {
    return res.status(500).json({
      __isGasError: true,
      message: "Proxy Error: " + error.message
    });
  }
};
