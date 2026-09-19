/**
 * Universal API Handler សម្រាប់ទទួល Request ពី Vercel
 * ដំណើរការរាល់ call ទាំងអស់ចេញពី fetch() លើ Vercel
 */
function doPost(e) {
  var output;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("មិនមានទិន្នន័យបញ្ជូនមកទេ (No post data received)");
    }

    var request = JSON.parse(e.postData.contents);
    var action = request.action;
    var args = request.args || [];

    var fn = null;
    if (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function') {
      fn = globalThis[action];
    } else if (typeof this !== 'undefined' && typeof this[action] === 'function') {
      fn = this[action];
    }

    if (!fn) {
      throw new Error("រកមិនឃើញអនុគមន៍ '" + action + "' នៅលើ Server ឡើយ។");
    }

    var result = fn.apply(null, args);
    output = ContentService.createTextOutput(JSON.stringify(result === undefined ? null : result));

  } catch (error) {
    Logger.log("[doPost Error] " + error.toString());
    output = ContentService.createTextOutput(JSON.stringify({
      __isGasError: true,
      message: error.message || error.toString()
    }));
  }

  return output.setMimeType(ContentService.MimeType.JSON);
}
