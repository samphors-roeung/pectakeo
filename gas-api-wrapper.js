/**
 * Google Apps Script API Wrapper (Drop-in Polyfill for google.script.run)
 * Allows google.script.run to work seamlessly on Vercel, Localhost, etc. via fetch()
 */
(function(window) {
  const GAS_EXEC_URL = "https://script.google.com/macros/s/AKfycbxNBwlpLENSGCA-zyCyY_WHvYLDhVtku-9Hd46PTzLwCEsTqB_gXOXlYawlI7egoU-J/exec";

  // If running natively inside Google Apps Script iframe, use native runner
  if (window.google && window.google.script && window.google.script.run) {
    return;
  }

  window.google = window.google || {};
  window.google.script = window.google.script || {};

  function createGasRunner(successHandler, failureHandler, userObj) {
    return new Proxy({}, {
      get(target, prop) {
        if (prop === 'withSuccessHandler') {
          return function(fn) {
            return createGasRunner(fn, failureHandler, userObj);
          };
        }
        if (prop === 'withFailureHandler') {
          return function(fn) {
            return createGasRunner(successHandler, fn, userObj);
          };
        }
        if (prop === 'withUserObject') {
          return function(obj) {
            return createGasRunner(successHandler, failureHandler, obj);
          };
        }

        // Intercept any server function call (e.g. loginUser, getDashboardStats, etc.)
        return function(...args) {
          const actionName = prop;

          // text/plain avoids CORS preflight OPTIONS request
          fetch(GAS_EXEC_URL, {
            method: 'POST',
            headers: {
              'Content-Type': 'text/plain;charset=utf-8'
            },
            body: JSON.stringify({
              action: actionName,
              args: args
            }),
            mode: 'cors',
            redirect: 'follow'
          })
          .then(async (response) => {
            if (!response.ok) {
              throw new Error(`HTTP Error ${response.status}: ${response.statusText}`);
            }
            const text = await response.text();
            let data;
            try {
              data = JSON.parse(text);
            } catch (e) {
              if (text.includes('<html') || text.includes('<!DOCTYPE') || text.includes('Google Apps Script') || text.includes('doPost')) {
                throw new Error("Google Apps Script មិនទាន់បាន Deploy 'New Version' ដែលមាន doPost(e) នៅឡើយទេ។ សូមចូលទៅ script.google.com រួចចុច Deploy > Manage deployments > Edit > ជ្រើសរើស Version 'New version' > Deploy។");
              }
              throw new Error("ការឆ្លើយតបពី Server មិនមែនជាទម្រង់ JSON ត្រឹមត្រូវឡើយ៖ " + text.substring(0, 120));
            }
            return data;
          })
          .then((res) => {
            if (res && res.__isGasError) {
              const err = new Error(res.message || "កំហុសបច្ចេកទេសនៅលើ Apps Script Server");
              if (typeof failureHandler === 'function') {
                failureHandler(err, userObj);
              } else {
                console.error(`[GAS Server Error in ${actionName}]:`, err);
              }
              return;
            }

            if (typeof successHandler === 'function') {
              successHandler(res, userObj);
            }
          })
          .catch((err) => {
            console.error(`[Network/GAS Error in ${actionName}]:`, err);
            if (typeof failureHandler === 'function') {
              failureHandler(err, userObj);
            }
          });
        };
      }
    });
  }

  window.google.script.run = createGasRunner(null, function(err) {
    console.error("[Unhandled GAS Error]:", err);
  }, null);

  console.log("🚀 Google Apps Script API Wrapper initialized on Vercel!");
})(window);
