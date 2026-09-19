// ==================== Auth.gs — ការ Login/Logout, Session, គណនីអ្នកប្រើប្រាស់ ====================

// ==== FIX (Fix117, "ពេល Login ម្តងៗដំណើរការយឺតពេក")៖ មុននេះ Format ជួរឈរ M (ថ្ងៃផុតកំណត់) ត្រូវបានអនុវត្តរហូតដល់
// sheet.getMaxRows() ទាំងស្រុង (ជាធម្មតា ~១០០០ជួរដេក លំនាំដើមរបស់ Google Sheet ថ្មី) ទោះបីជាមានទិន្នន័យអ្នកប្រើប្រាស់
// ពិតប្រាកដត្រឹមតែពីរបីរយជួរដេកក៏ដោយ។ នេះជាកំហុសបច្ចេកទេសដ៏ល្បីមួយក្នុង Google Apps Script៖ ការ Format ក្រឡាទទេ
// (setNumberFormat/setBackground ។ល។ ទោះគ្មានទិន្នន័យផ្ទាល់) ក៏ "ពង្រីក" ដែនកំណត់ទិន្នន័យ (Used Range) របស់ Sheet
// ដែរ ធ្វើឲ្យ sheet.getDataRange()/getLastRow() ត្រឡប់ចំនួនជួរដេកច្រើនហួសហេតុជានិច្ច (ដល់ ~១០០០) ទោះក្រោយបំផុតមាន
// ទិន្នន័យពិតតែពីរបីរយជួរដេក — ជាហេតុធ្វើឲ្យ loginUser()/getUserRow_()/requireValidSession_() (ហៅរាល់ពេល Login
// ក៏ដូចជារាល់ការហៅ Server ស្ទើរតែទាំងអស់ក្រោយ Login) ត្រូវអាន+ស្កេនជួរដេក ~១០០០ ជានិច្ចរាល់ដងហៅ (ខណៈត្រូវការតែពីរបី
// រយ) — នេះជាមូលហេតុចម្បងមួយដែលធ្វើឲ្យ Login និងសកម្មភាពគ្រប់យ៉ាងក្នុងប្រព័ន្ធយឺត។ ដំណោះស្រាយ (២ផ្នែក)៖
//  ១. formatExpiryColumnBounded_() ខាងក្រោម — កំណត់ Format ត្រឹមតែជួរដេកទិន្នន័យបច្ចុប្បន្ន + ស្តុកទុក ៥០ជួរដេក
//     សម្រាប់ការរីកលូតលាស់ជិតៗនេះប៉ុណ្ណោះ (មិនមែនរហូតដល់ getMaxRows() ទាំងអស់ទេ) — ជៀសវាងកំហុសនេះកើតឡើងម្តងទៀត
//  ២. trimUsersSheetRowBloat_() ខាងក្រោម — កាត់បន្ថយជួរដេកលើសដែលធ្លាប់ត្រូវ Format ខុសរួចហើយ (ពីមុនកូដនេះ) ចេញពី
//     Sheet ដែលមានស្រាប់ (ដំណើរការតែម្តងគត់ក្នុងមួយ Deployment — ការពារដោយ Script Property) ====
function formatExpiryColumnBounded_(sheet) {
  var bufferRows = 50;
  var neededRows = Math.max(sheet.getLastRow(), 1) + bufferRows;
  var fmtRows = Math.min(sheet.getMaxRows() - 1, neededRows);
  if (fmtRows > 0) sheet.getRange(2, 12, fmtRows, 1).setNumberFormat("@");
}
function trimUsersSheetRowBloat_(sheet) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('USERS_SHEET_TRIMMED_V1') === '1') return; // ធ្លាប់កាត់រួចហើយ — ដំណើរការតែម្តងគត់
    var maxRows = sheet.getMaxRows();
    // ==== រកជួរដេកចុងក្រោយ "ពិតប្រាកដ" ដែលមានទិន្នន័យ Username (ជួរឈរ B) — មិនអាចទុកចិត្តលើ sheet.getLastRow()
    // បានទេ ព្រោះខ្លួនវាផ្ទាល់ក៏ត្រូវបានពង្រីកមិនត្រឹមត្រូវ ដោយសារកំហុស Format ដដែលនេះដែរ ====
    var usernames = sheet.getRange(1, 2, maxRows, 1).getDisplayValues();
    var trueLastRow = 1;
    for (var i = usernames.length - 1; i >= 1; i--) {
      if (String(usernames[i][0]).trim() !== '') { trueLastRow = i + 1; break; }
    }
    var bufferRows = 50;
    var keepUntilRow = Math.min(maxRows, trueLastRow + bufferRows);
    if (maxRows > keepUntilRow) {
      sheet.deleteRows(keepUntilRow + 1, maxRows - keepUntilRow);
    }
    props.setProperty('USERS_SHEET_TRIMMED_V1', '1');
  } catch (err) {}
}

function ensureUsersSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USERS);
    var headers = ["ID", "Username", "Password", "គោត្តនាម-នាម", "តួនាទី (Role)", "ស្រុក (បើមាន)", "ស្ថានភាព", "កាលបរិច្ឆេទបង្កើត", "Gmail (សម្រាប់ចូល Google Sheet ដោយផ្ទាល់)", "Key Number", "Session Token (ឧបករណ៍ដែលកំពុង Login)", "ថ្ងៃផុតកំណត់ (សម្រាប់សង្កេតការណ៍/គ្រូប្រចាំស្រុក)", "ឃុំ/សង្កាត់ (សម្រាប់គណនីឃុំសង្កាត់)"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ ធានាថាជួរឈរនេះរក្សាទុកជា Text ធម្មតាជានិច្ច
    // ("yyyy-MM-dd" ត្រូវនឹងតម្លៃពី <input type="date">) មិនអនុញ្ញាតឲ្យ Google Sheet បម្លែងទៅជា Date Object
    // ដោយស្វ័យប្រវត្តិទេ (ជៀសវាងបញ្ហាទម្រង់បង្ហាញខុសគ្នាទៅតាម Locale ពេលអានមកប្រៀបធៀបនឹងកាលបរិច្ឆេទបច្ចុប្បន្ន) ====
    formatExpiryColumnBounded_(sheet);
    // គណនី SuperAdmin លំនាំដើម (សូមប្តូរលេខសម្ងាត់ភ្លាមៗបន្ទាប់ពី Login លើកដំបូង!)
    sheet.appendRow([newId_(), "admin", makePasswordHash_("admin123"), "អ្នកគ្រប់គ្រងប្រព័ន្ធ", ROLE_SUPERADMIN, "", STATUS_ACTIVE, formatNow_(), "", generateKeyNumber_()]);
  } else {
    var lastCol = sheet.getLastColumn();
    // ==== ជួសជុល Sheet ចាស់ៗ ដែលមិនទាន់មានជួរឈរ "Session Token" (ត្រូវការសម្រាប់ការកំណត់ Login មួយឧបករណ៍) ====
    if (lastCol < 11) {
      sheet.getRange(1, 11).setValue("Session Token (ឧបករណ៍ដែលកំពុង Login)").setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    }
    // ==== ជួសជុល Sheet ចាស់ៗ ដែលមិនទាន់មានជួរឈរ "ថ្ងៃផុតកំណត់" (FIX — សុពលភាព/រយៈពេលប្រើប្រាស់) ====
    if (lastCol < 12) {
      sheet.getRange(1, 12).setValue("ថ្ងៃផុតកំណត់ (សម្រាប់សង្កេតការណ៍/គ្រូប្រចាំស្រុក)").setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
      formatExpiryColumnBounded_(sheet);
    }
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ជួសជុល Sheet ចាស់ៗ ដែលមិនទាន់មានជួរឈរ "ឃុំ/សង្កាត់" (ជួរឈរ M) —
    // ត្រូវការសម្រាប់កំណត់អត្តសញ្ញាណគណនីកម្រិតឃុំសង្កាត់ (ROLE_COMMUNE) ព្រោះឈ្មោះឃុំសង្កាត់មិនមែនតែមួយគត់ទូទាំងខេត្ត
    // ដូច្នេះមិនអាចប្រើ role=ឈ្មោះឃុំសង្កាត់ ដូចលំនាំគណនីស្រុកបានទេ ====
    if (lastCol < 13) {
      sheet.getRange(1, 13).setValue("ឃុំ/សង្កាត់ (សម្រាប់គណនីឃុំសង្កាត់)").setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    }
    // ==== FIX (សំណើថ្មី "បង្កើតគណនីឃុំសង្កាត់ដោយស្វ័យប្រវត្តិ")៖ ជួសជុល Sheet ចាស់ៗ ដែលមិនទាន់មានជួរឈរ "តម្រូវឲ្យ
    // ប្តូរលេខសម្ងាត់ដំបូង" (ជួរឈរ N) — Flag នេះកំណត់ដោយ bulkCreateAllCommuneAccounts_() ចំពោះគណនីដែលបានបង្កើតដោយ
    // ស្វ័យប្រវត្តិ (ចែករំលែងលេខសម្ងាត់លំនាំដើមតែមួយដូចគ្នា) — ត្រូវហាមឃាត់ loginUser() មិនឲ្យបន្តទៅ Session ពេញលេញ
    // រហូតដល់ប្តូរលេខសម្ងាត់+ឈ្មោះខ្លួនឯង (មើល changePasswordFirstLogin() ក្នុង Auth.gs) ====
    if (lastCol < 14) {
      sheet.getRange(1, 14).setValue("តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូង (សម្រាប់គណនីបង្កើតស្វ័យប្រវត្តិ)").setFontWeight("bold").setBackground("#7c3aed").setFontColor("#ffffff");
    }
    // ==== FIX (Fix117, "ពេល Login ម្តងៗដំណើរការយឺតពេក")៖ កាត់បន្ថយជួរដេកលើសដែលធ្លាប់ត្រូវ Format ខុសរួច (មើលការ
    // ពន្យល់លម្អិតខាងលើ) ចេញពី Sheet ដែលមានស្រាប់នេះ — ដំណើរការតែម្តងគត់ (ការពារដោយ Script Property) ====
    trimUsersSheetRowBloat_(sheet);
  }
  return sheet;
}

// បង្កើតលេខសម្គាល់ (Key Number) ៦ខ្ទង់ ដោយចៃដន្យ សម្រាប់គណនីថ្មីនីមួយៗ (ប្រើសម្រាប់ផ្ញើជូនដំណឹង/ផ្ទៀងផ្ទាត់ទៅ SuperAdmin)
function generateKeyNumber_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

// ==================== ០.១ ការ Hash លេខសម្ងាត់ (Password Hashing) — ការពារកុំឲ្យលេខសម្ងាត់ត្រូវបានផ្ទុកជា Plain Text ====
// ==== FIX (សំខាន់ណាស់ — Plaintext Password Storage)៖ មុននេះ ជួរឈរ "Password" ក្នុង Sheet(អ្នកប្រើប្រាស់) ផ្ទុក
// លេខសម្ងាត់ជាអក្សរធម្មតា (Plain Text) ដោយផ្ទាល់ — នរណាម្នាក់ដែលបើក Google Sheet នេះបាន (ឧ. Admin ណាមួយដែលមាន
// Gmail ជា Editor, ឬ SuperAdmin ណាមួយ) នឹងឃើញលេខសម្ងាត់ពិតប្រាកដរបស់អ្នកប្រើប្រាស់ទាំងអស់ភ្លាមៗ ដែលបង្កហានិភ័យ
// ធំធេង ព្រោះមនុស្សភាគច្រើនប្រើលេខសម្ងាត់ដដែលនៅកន្លែងផ្សេងទៀត (ឧ. Email ផ្ទាល់ខ្លួន) ។ ការជួសជុលនេះ ប្តូរទៅផ្ទុក
// "Hash" (លទ្ធផលគណនាតាមទិសដៅតែមួយ មិនអាចត្រឡប់ក្រោយទៅជាលេខសម្ងាត់ដើមវិញបានទេ) ជំនួសលេខសម្ងាត់ផ្ទាល់ ដោយប្រើ
// SHA-256 ធ្វើម្តងទៀតៗច្រើនដង (Iterated Hashing ស្រដៀង PBKDF2 សាមញ្ញ ព្រោះ Apps Script គ្មាន bcrypt/scrypt/argon2
// ជាប់មកជាមួយទេ) បូក Salt ចៃដន្យផ្ទាល់ខ្លួនក្នុងមួយគណនី (ការពារ Rainbow-table Attack) ។ ទម្រង់ដែលផ្ទុក៖
// "PH1$<salt hex>$<hash hex>" នៅក្នុងជួរឈរ "Password" ដដែល (មិនចាំបាច់បន្ថែមជួរឈរថ្មី ដើម្បីជៀសវាងកែប្រែលេខ
// Column Index នៅកន្លែងផ្សេងទៀតទាំងអស់ក្នុងប្រព័ន្ធ) ។ គណនីចាស់ៗ ដែលនៅមិនទាន់បាន Hash (Legacy Plaintext) នៅតែ
// ដំណើរការធម្មតា (verifyPassword_ ត្រួតពិនិត្យទម្រង់ភ្លាមៗ) រហូតដល់ពេលប្តូរទៅ Hash ដោយស្វ័យប្រវត្តិ (មើល
// migratePlaintextPasswordsIfNeeded_ និង loginUser) ====
var PASSWORD_HASH_PREFIX_ = "PH1$";
var PASSWORD_HASH_ITERATIONS_ = 3000; // តុល្យភាពរវាងសុវត្ថិភាព (លំបាកទាយ) និងល្បឿន (Login មិនយឺតដល់អ្នកប្រើ)
var PASSWORD_MIGRATION_FLAG_PROP_ = "passwordsHashMigratedV1";

function generateSalt_() {
  var bytes = [];
  for (var i = 0; i < 16; i++) bytes.push(Math.floor(Math.random() * 256));
  return bytesToHex_(bytes);
}
function bytesToHex_(bytes) {
  return bytes.map(function(b) { return ("0" + (b & 0xFF).toString(16)).slice(-2); }).join("");
}
function sha256Hex_(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return bytesToHex_(raw);
}
// ==== Iterated SHA-256 (ស្រដៀង PBKDF2 សាមញ្ញ)៖ ធ្វើ Hash ដដែលៗច្រើនដង ដើម្បីបង្កើនពេលវេលា/ថាមពលកុំព្យូទ័រ
// ចាំបាច់ក្នុងការទាយលេខសម្ងាត់ដោយវាយលេលើ Hash ដែលលេចធ្លាយ (Brute-force/Rainbow-table resistance) ====
function hashPasswordWithSalt_(plainPassword, salt) {
  var h = salt + ":" + plainPassword;
  for (var i = 0; i < PASSWORD_HASH_ITERATIONS_; i++) {
    h = sha256Hex_(h + salt);
  }
  return h;
}
function makePasswordHash_(plainPassword) {
  var salt = generateSalt_();
  var hash = hashPasswordWithSalt_(String(plainPassword || ""), salt);
  return PASSWORD_HASH_PREFIX_ + salt + "$" + hash;
}
// ==== FIX (ការពារ Timing Attack)៖ ប្រៀបធៀប String ពីរដោយប្រើពេលវេលាថេរ (មិនប្រែប្រួលទៅតាមចំនួនតួអក្សរដូចគ្នា
// ពីដើម) ដើម្បីកុំឲ្យអ្នកវាយប្រហារអាចទាយ Hash/Token ម្តងមួយតួអក្សរ ដោយវាស់ភាពខុសគ្នានៃពេលវេលាឆ្លើយតប។ ត្រូវធ្វើរង្វិលជុំ
// ប្រៀបធៀបគ្រប់តួអក្សរជានិច្ច (មិនបញ្ឈប់ភ្លាមពេលឃើញភាពខុសគ្នា) ====
function timingSafeEqual_(a, b) {
  a = String(a === null || a === undefined ? "" : a);
  b = String(b === null || b === undefined ? "" : b);
  var lenEqual = (a.length === b.length);
  // ==== ប្រៀបធៀបគ្រប់តួអក្សរជានិច្ច (រហូតដល់ប្រវែងវែងជាងគេ) ដោយមិនបញ្ឈប់ភ្លាមពេលឃើញភាពខុសគ្នា — ចៀសវាងកុំឲ្យពេលវេលា
  // ដំណើរការប្រែប្រួលទៅតាមចំណុចដែលចាប់ផ្តើមខុសគ្នា។ diff ចាប់ផ្តើមពី ១ បើប្រវែងខុសគ្នា ដូច្នេះលទ្ធផលមិនអាចស្មើ ០ បានឡើយ ====
  var diff = lenEqual ? 0 : 1;
  var cmpLen = Math.max(a.length, b.length);
  for (var i = 0; i < cmpLen; i++) {
    var ca = i < a.length ? a.charCodeAt(i) : 0;
    var cb = i < b.length ? b.charCodeAt(i) : 0;
    diff |= (ca ^ cb);
  }
  return diff === 0;
}
// ត្រូវផ្ទៀងផ្ទាត់លេខសម្ងាត់ដែលបញ្ចូល ធៀបនឹងតម្លៃដែលផ្ទុកក្នុង Sheet (អាចជា Hash ថ្មី ឬ Plaintext ចាស់ដែលមិនទាន់ប្តូរ)
function verifyPassword_(plainPassword, stored) {
  stored = String(stored || "");
  plainPassword = String(plainPassword || "").trim();
  if (stored.indexOf(PASSWORD_HASH_PREFIX_) !== 0) {
    // ==== គណនីចាស់ដែលមិនទាន់ Hash (Legacy Plaintext) — ប្រៀបធៀបផ្ទាល់ (ផ្លូវនេះនឹងលែងប្រើនៅពេលក្រោយ បន្ទាប់ពី
    // migratePlaintextPasswordsIfNeeded_ ឬ Login ជោគជ័យលើកបន្ទាប់ ដែលនឹង Hash ជាថ្មីភ្លាមៗ) ====
    return timingSafeEqual_(stored.trim(), plainPassword);
  }
  var rest = stored.slice(PASSWORD_HASH_PREFIX_.length);
  var sepIdx = rest.indexOf("$");
  if (sepIdx === -1) return false;
  var salt = rest.slice(0, sepIdx);
  var expectedHash = rest.slice(sepIdx + 1);
  return timingSafeEqual_(hashPasswordWithSalt_(plainPassword, salt), expectedHash);
}
function isPasswordHashed_(stored) {
  return String(stored || "").indexOf(PASSWORD_HASH_PREFIX_) === 0;
}

// ==== ការប្តូរលេខសម្ងាត់ចាស់ៗ (Plaintext) ទៅជា Hash ដោយស្វ័យប្រវត្តិ តែម្តងគត់ (Idempotent — ហៅច្រើនដងក៏មិនប៉ះពាល់)
// ហៅពី doGet() (មើល Utils.gs) ដោយប្រើ Script Property ជា Flag ដើម្បីកុំឲ្យស្កេន Sheet ឡើងវិញរាល់ការស្នើសុំ
// (មានតែម្តងគត់ ចាប់តាំងពី Deploy កំណែនេះដំបូង) — ក្រៅពីនេះ loginUser() ក៏ Hash ជាថ្មីភ្លាមៗនៅពេល Login ជោគជ័យ
// លើគណនីមួយណាដែលនៅសល់ជា Plaintext ដដែរ (ការពារពីរជាន់ បើ Migration នេះខកខានហេតុផលណាមួយ) ====
function migratePlaintextPasswordsIfNeeded_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PASSWORD_MIGRATION_FLAG_PROP_) === "1") return;
  try {
    var ss = getSS_();
    var sheet = ensureUsersSheet_(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var range = sheet.getRange(2, 3, lastRow - 1, 1); // ជួរឈរ C = Password
      var values = range.getValues();
      var changed = false;
      for (var i = 0; i < values.length; i++) {
        var pw = String(values[i][0] || "");
        if (pw && !isPasswordHashed_(pw)) {
          values[i][0] = makePasswordHash_(pw);
          changed = true;
        }
      }
      if (changed) range.setValues(values);
    }
    props.setProperty(PASSWORD_MIGRATION_FLAG_PROP_, "1");
  } catch (errMig) {
    // ==== បើបរាជ័យ (ឧ. Permission/Quota បណ្ដោះអាសន្ន) កុំកំណត់ Flag ថាបានធ្វើរួច — ព្យាយាមម្តងទៀត Request ក្រោយ ====
  }
}

// ==================== ០.១.១ សុពលភាព/រយៈពេលប្រើប្រាស់គណនី (Account Expiry) — សម្រាប់សង្កេតការណ៍/គ្រូប្រចាំក្រុងស្រុក ====
// ==== FIX (សំណើ)៖ SuperAdmin/Admin អាចកំណត់ "ថ្ងៃផុតកំណត់" ដាច់ដោយឡែកសម្រាប់គណនីម្នាក់ៗ (Observer/District តែប៉ុណ្ណោះ
// — គណនី SuperAdmin/Admin/PEC21 មិនអាចកំណត់ថ្ងៃផុតកំណត់បានទេ ដើម្បីជៀសវាងខ្លួនឯងចាក់សោខ្លួនឯងចេញ) ។ ទុកទទេ =
// មិនកំណត់ថ្ងៃផុតកំណត់ទេ (ប្រើប្រាស់បានគ្មានកំណត់ ដូចមុន) ។ ថ្ងៃដែលបញ្ចូល គឺជា "ថ្ងៃចុងក្រោយ" ដែលនៅប្រើប្រាស់បាន
// ធម្មតា (ដូចកាលបរិច្ឆេទផុតកំណត់លើប័ណ្ណសម្គាល់ខ្លួន) — ចាប់ផ្តើមថ្ងៃបន្ទាប់ គណនីនឹងត្រូវបានទប់ស្កាត់ភ្លាមៗ ទាំង
// ការ Login ថ្មី និង Session ដែលកំពុងបើកប្រើប្រាស់ស្រាប់ (មើល requireValidSession_/loginUser/validateSession) ====
function isExpiryApplicableRole_(role) {
  return role === ROLE_OBSERVER || isDistrictRole_(role) || role === ROLE_COMMUNE || isObserverTierRole_(role);
}
// ធ្វើឲ្យតម្លៃថ្ងៃផុតកំណត់ដែលបញ្ចូល ក្លាយទៅជាទម្រង់ "yyyy-MM-dd" ត្រឹមត្រូវ ឬ "" (ទទេ) បើមិនត្រឹមត្រូវ/មិនបានផ្តល់
function normalizeExpiryInput_(v) {
  var s = String(v || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : "";
}
// តើគណនីនេះបានផុតកំណត់រយៈពេលប្រើប្រាស់ហើយឬនៅ (ត្រូវផ្តល់ Role និងតម្លៃ "ថ្ងៃផុតកំណត់" ដែលទុកក្នុង Sheet)
function isAccountExpired_(role, expiryDate) {
  if (!isExpiryApplicableRole_(role)) return false; // អនុវត្តតែលើ សង្កេតការណ៍/គ្រូប្រចាំក្រុងស្រុក
  var exp = normalizeDateStr_(expiryDate);
  if (!exp) return false; // ទទេ = មិនកំណត់ (ប្រើប្រាស់បានគ្មានកំណត់)
  return exp < todayDateStr_();
}

// ---- Sheet(ឃុំ សង្កាត់) — បញ្ជីឈ្មោះឃុំ/សង្កាត់ តាមស្រុកនីមួយៗ (សម្រាប់ចម្រាញ់ Dropdown) ----
// អ្នកគ្រប់គ្រងត្រូវបើក Sheet នេះផ្ទាល់ ហើយបំពេញ/កែសម្រួលឈ្មោះឃុំ/សង្កាត់ជាក់ស្តែងឲ្យត្រូវនឹងស្រុកនីមួយៗ
// (មួយជួរ = មួយឃុំ/សង្កាត់)។ ទម្រង់បញ្ចូលទិន្នន័យប្រចាំថ្ងៃនឹងទាញយកបញ្ជីនេះ ដើម្បីបង្ហាញជា Dropdown ស្វ័យប្រវត្តិ។
// ==================== ០.២.១ ការអនុញ្ញាតឲ្យគ្រូប្រចាំស្រុក កែថ្ងៃដែលកន្លងហួសទៅហើយ ====================
// លំនាំដើម៖ គ្រូប្រចាំស្រុក (District role) កែ/លុបបានតែថ្ងៃបច្ចុប្បន្ន ឬអនាគតប៉ុណ្ណោះ — ថ្ងៃដែលកន្លងហួសហើយ
// ត្រូវបានចាក់សោស្វ័យប្រវត្តិ។ SuperAdmin/Admin អាចបើកសិទ្ធិឲ្យកែថ្ងៃជាក់លាក់ណាមួយបានជាបណ្តោះអាសន្ន
// (សម្រាប់ស្រុកជាក់លាក់ណាមួយ ឬគ្រប់ស្រុកទាំងអស់) ។ SuperAdmin/Admin ខ្លួនឯង មិនរងឥទ្ធិពលពីការចាក់សោនេះទេ។
function resetVisitCounter(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  try {
    PropertiesService.getScriptProperties().setProperty(VISIT_COUNTER_PROP_, "0");
    return { success: true, message: "សម្អាតចំនួនចូលប្រើជោគជ័យ!" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== FIX (Fix118, "ដំណើរការ Login Form ដើរយឺត")៖ skipStats (Parameter ថ្មី) — Function នេះដើមឡើយសម្រាប់តែទំព័រ
// Login សាធារណៈ (មុន Login — ត្រូវការ memberCount/visitCount ពិតប្រាកដ) ប៉ុន្តែក្រោយមក applyLoginSession() ក៏ហៅ
// Function នេះដដែល (សម្រាប់តែទាញ enabledLangs/Branding ឡើងវិញ — មើលការពន្យល់ត្រង់ applyLoginSession() ក្នុង
// Index.html) ដែលធ្វើឲ្យរាល់ Login ត្រូវអាន+ស្កេន Sheet(អ្នកប្រើប្រាស់) ទាំងមូលដដែលៗដោយឥតប្រយោជន៍ (memberCount មិន
// ដែលបង្ហាញនៅក្នុង App ក្រោយ Login ទេ — មានតែនៅទំព័រ Login សាធារណៈប៉ុណ្ណោះ)។ skipStats=true ឲ្យរំលងការគណនា
// activeCount ទាំងស្រុង (សូម្បីតែការបើក Spreadsheet — getSS_() — ក៏មិនចាំបាច់ទៀត) ត្រូវប្រើសម្រាប់រាល់ការហៅក្រោយ
// Login/ខណៈកំពុង Login រួចហើយ (Tab ការកំណត់ប្រព័ន្ធ ។ល។) ទុកតែការហៅពីទំព័រ Login សាធារណៈប៉ុណ្ណោះ ដែលនៅតែគណនាដដែល ====
function getAuthPageInfo(skipStats) {
  var activeCount = 0, visits = 0;
  if (!skipStats) {
    try {
      var sheet = ensureUsersSheet_(getSS_());
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) { if (data[i][6] === STATUS_ACTIVE) activeCount++; }
    } catch (err) {}
    try { visits = Number(PropertiesService.getScriptProperties().getProperty(VISIT_COUNTER_PROP_)) || 0; } catch (err) {}
  }
  var s = getSystemSettings_();
  return {
    success: true, districts: DISTRICT_LIST, memberCount: activeCount, visitCount: visits,
    orgName: s.orgName, orgSlogan: s.orgSlogan, logoUrl: s.logoUrl,
    telegramLink: s.telegramLink, youtubeLink: s.youtubeLink, enabledLangs: s.enabledLangs,
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ត្រូវការសម្រាប់ Dropdown ជ្រើសរើសឃុំសង្កាត់ (Cascading — តាមស្រុកដែលបានជ្រើស)
    // នៅក្នុងទម្រង់ស្នើសុំគណនីខ្លួនឯង (Panel Signup) — Function នេះជាសាធារណៈ (គ្មានតម្រូវឲ្យ Login) ហៅរួចហើយពេលបើក
    // ទំព័រ Login ដូច្នេះមិនចាំបាច់ហៅបន្ថែមទៀតទេ ====
    communeOrder: COMMUNE_ORDER,
    // ==== FIX (Fix127, "ជម្រើសបើក/បិទប្រើប្រាស់គណនីឃុំសង្កាត់")៖ ត្រូវការសម្រាប់លាក់ជម្រើស "អ្នកបញ្ចូលទិន្នន័យ
    // ឃុំសង្កាត់" ចេញពី Dropdown "ប្រភេទគណនី" នៅ Panel ស្នើសុំគណនីថ្មី (មុន Login) បើ Admin បិទមុខងារនេះ ====
    enableCommuneAccounts: isCommuneAccountsEnabled_(),
    // ==== FIX (សំណើថ្មី "បិទកន្លែងស្នើសុំបង្កើតគណនីថ្មីទាំងអស់")៖ ត្រូវការជានិច្ច (ទោះ skipStats=true ក៏ដោយ — ដូចគ្នា
    // នឹង orgName/logoUrl ខាងលើ) ព្រោះទំព័រ Login សាធារណៈតម្រូវការដឹងភ្លាមៗថាតើត្រូវលាក់តំណភ្ជាប់ "ស្នើសុំគណនីថ្មីទីនេះ"
    // ចេញឬអត់ — មើល loadAuthPageInfo()/showAuthPanel() ក្នុង Index.html ====
    enableSelfSignup: isSelfSignupEnabled_(),
    // ==== FIX (Fix129, "គណនីថ្មីទាំង៥ ត្រូវមាន Option បើក/បិទ ការបង្កើតគណនីនីមួយៗ")៖ ត្រូវការសម្រាប់លាក់/បង្ហាញ
    // ជម្រើសនីមួយៗក្នុងចំណោមតួនាទីថ្មីទាំង៥ ចេញពី Dropdown "ប្រភេទគណនី" នៅ Panel ស្នើសុំគណនីថ្មី (មុន Login) —
    // ត្រឡប់ជា Array {role, accountType, optionId, label, enabled} ដើម្បីឲ្យ Index.html អាច Loop កំណត់ Show/Hide
    // ដោយស្វ័យប្រវត្តិ (មិនចាំបាច់សរសេរ Field ដាច់ដោយឡែកម្នាក់ៗទេ — មើល loadAuthPageInfo ក្នុង Index.html) ====
    enabledObserverTierRoles: OBSERVER_TIER_ROLES_.map(function(cfg) {
      return { role: cfg.role, accountType: cfg.accountType, optionId: cfg.optionId, label: cfg.label, enabled: isObserverTierRoleEnabled_(cfg.role) };
    })
  };
}

// ==== FIX (សំណើថ្មី "កំណត់ចំនួនកំរិតគណនី")៖ កំណត់ចំនួនកំរិតអតិបរមានៃគណនីគ្រូប្រចាំក្រុងស្រុក ក្នុងមួយក្រុងស្រុក
// (២ ក្នុងមួយស្រុក × ១០ស្រុក = ២០ គណនីសរុបអតិបរមាទូទាំងខេត្ត) និងគណនីឃុំសង្កាត់ ក្នុងមួយឃុំសង្កាត់ (១ ក្នុងមួយឃុំសង្កាត់
// × ១០០ឃុំសង្កាត់ = ១០០ គណនីសរុបអតិបរមាទូទាំងខេត្ត) — អនុវត្តទាំងផ្លូវស្នើសុំខ្លួនឯង (requestSignup ខាងក្រោម) និងផ្លូវ
// Admin/គ្រូប្រចាំក្រុងស្រុកបង្កើតដោយផ្ទាល់ (addUserAccount) ។ ការរាប់ រាប់បញ្ចូលគណនីគ្រប់ស្ថានភាព (រង់ចាំអនុម័ត/សកម្ម/
// ផ្អាក) ព្រោះទាំងនេះនៅតែកាន់កាប់ "កន្លែង" មួយក្នុងចំនួនកំណត់ដដែល (ការស្នើសុំមួយ ទោះជាមិនទាន់អនុម័ត ក៏រារាំងគណនីទី៣/
// ទី២ បន្ថែមទៀតដែរ ដើម្បីជៀសវាងសំណើច្រើនក្នុងពេលដំណាលគ្នាឆ្លងកាត់ការត្រួតពិនិត្យទាំងអស់មុននឹងណាមួយត្រូវបានអនុម័ត) —
// មានតែការលុបគណនី (deleteUserAccount, ដែលលុបជួរដេកចោលពិតប្រាកដ) ទេ ទើបដោះលែងកន្លែងវិញ។ ការត្រួតពិនិត្យទាំង២ខាងក្រោម
// ត្រូវហៅពីខាងក្នុង LockService Critical Section ដដែល (ក្រោយអានទិន្នន័យ Sheet ចូល `data`) ដើម្បីជៀសវាង Race
// Condition (២សំណើស្ទើរតែក្នុងពេលតែមួយ ទាំង២ឆ្លងកាត់ការត្រួតពិនិត្យ ១ ដូចគ្នា មុននឹងណាមួយបាន appendRow) — មិនមែនជា
// ការធានាដាច់ខាត១០០%ចំពោះទិន្នន័យចាស់ដែលមានស្រាប់ពីមុនកំណត់នេះទេ (បើស្រុក/ឃុំសង្កាត់ណាមួយមានលើសចំនួនកំណត់រួចហើយ
// ពីមុន Fix នេះ មិនត្រូវបានលុប/ផ្អាកដោយស្វ័យប្រវត្តិទេ — កំណត់នេះរារាំងតែការបន្ថែមថ្មីទៀតតទៅមុខប៉ុណ្ណោះ) ====
var DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ = 2;
var COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ = 1;
// រាប់ចំនួនគណនីគ្រូប្រចាំក្រុងស្រុក (role === ឈ្មោះស្រុកខ្លួនឯង ដូចនៅក្នុង requestSignup/addUserAccount) សម្រាប់ស្រុក
// ដែលបានផ្តល់មក — data ត្រូវជា getDisplayValues()/getValues() ពេញលេញ (រួមទាំងជួរដេកចំណងជើង Header ជួរទី១)
function countDistrictRoleAccounts_(data, district) {
  var n = 0;
  for (var i = 1; i < data.length; i++) { if (data[i][4] === district) n++; }
  return n;
}
// រាប់ចំនួនគណនីឃុំសង្កាត់ (ROLE_COMMUNE) ដែលកំណត់ចំពោះ ស្រុក+ឃុំសង្កាត់ ជាក់លាក់ដែលបានផ្តល់មក (មិនមែនរាប់ឈ្មោះ
// ឃុំសង្កាត់ដាច់ដោយឡែកទេ ព្រោះឈ្មោះឃុំសង្កាត់ដូចគ្នាអាចមាននៅច្រើនស្រុកខុសគ្នា — ស្រុក+ឃុំសង្កាត់ ទើបជាគូតែមួយគត់)
function countCommuneRoleAccountsForCommune_(data, district, commune) {
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][4] === ROLE_COMMUNE && String(data[i][5]).trim() === district && String(data[i][12] || "").trim() === commune) n++;
  }
  return n;
}

// ស្នើសុំគណនីថ្មីដោយខ្លួនឯង (សាធារណៈ — មិនត្រូវការចូលប្រព័ន្ធជាមុនទេ) — គណនីនឹងមានស្ថានភាព "រង់ចាំអនុម័ត"
// រហូតដល់ Admin/SuperAdmin ចូលអនុម័ត (មើល approveUserAccount) ទើបអាចប្រើប្រាស់បាន — ការពារកុំឲ្យនរណាម្នាក់
// ចុះឈ្មោះខ្លួនឯងជា Admin/SuperAdmin/PEC21 បាន (ដាក់កម្រិតឲ្យតែជាគណនីគ្រូប្រចាំស្រុក ឬ "សង្កេតការណ៍" ប៉ុណ្ណោះ) ។
// newUser.accountType === "observer" → ស្នើសុំគណនីជាតួនាទី "សង្កេតការណ៍" (គ្មានស្រុកជាក់លាក់)
// បើមិនមែនទេ (លំនាំដើម) → ស្នើសុំគណនីជាគ្រូប្រចាំស្រុក (ដូចដើម) ។
function requestSignup(newUser) {
  // ==== FIX (សំណើថ្មី "បិទកន្លែងស្នើសុំបង្កើតគណនីថ្មីទាំងអស់")៖ ត្រួតពិនិត្យត្រង់នេះមុនគេ (មុនសូម្បីតែផ្ទៀងផ្ទាត់
  // ព័ត៌មានបំពេញខាងក្រោម) ជាការការពារបន្ថែម (Defense-in-depth) — ទោះបីជា Client (Index.html) លាក់តំណភ្ជាប់/Panel
  // នេះរួចហើយ ក៏ករណីនរណាម្នាក់ហៅ Function នេះដោយផ្ទាល់ (ឧ. តាមរយៈ Console/កូដផ្ទាល់ខ្លួន) នៅតែត្រូវបានបដិសេធដដែល —
  // មើល isSelfSignupEnabled_() ក្នុង Settings.gs ====
  if (!isSelfSignupEnabled_()) {
    return { success: false, message: "ប្រព័ន្ធបានបិទការស្នើសុំគណនីថ្មីដោយខ្លួនឯងជាបណ្ដោះអាសន្ន! សូមទាក់ទង SuperAdmin/Admin ប្រព័ន្ធ ដើម្បីស្នើសុំគណនីថ្មីដោយផ្ទាល់។" };
  }
  if (!newUser || !String(newUser.fullName || "").trim() || !String(newUser.username || "").trim() ||
      !String(newUser.password || "").trim()) {
    return { success: false, message: "សូមបំពេញព័ត៌មានឲ្យបានគ្រប់គ្រាន់!" };
  }
  var isObserver = (String(newUser.accountType || "") === "observer");
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ បន្ថែមប្រភេទគណនីទី៣ — ស្នើសុំគណនីកម្រិតឃុំសង្កាត់ដោយខ្លួនឯង (ត្រូវការស្រុក+
  // ឃុំសង្កាត់ជាក់លាក់ទាំង២ ព្រោះឈ្មោះឃុំសង្កាត់មិនតែមួយគត់ទូទាំងខេត្ត) — ស្ថិតក្នុងស្ថានភាព "រង់ចាំអនុម័ត" ដូចគ្នា
  // រហូតដល់គ្រូប្រចាំក្រុងស្រុកម្ចាស់ស្រុកនោះ ឬ SuperAdmin/Admin/PEC21 អនុម័ត (មើល approveUserAccount) ====
  var isCommuneSignup = (String(newUser.accountType || "") === "commune");
  // ==== FIX (Fix129, "គណនីថ្មីទាំង៥ ស្នើសុំខ្លួនឯងបានដែរ")៖ ប្រភេទគណនីទី៤ — ស្នើសុំគណនីតួនាទីមួយក្នុងចំណោមតួនាទីថ្មី
  // ទាំង៥ ដោយខ្លួនឯង (accountType = "necofficer"/"provinceadmin"/"districtadmin"/"armedforces"/"politicalparty" —
  // មើល OBSERVER_TIER_ROLES_ ក្នុង Utils.gs) — មិនត្រូវការស្រុក/ឃុំសង្កាត់ដូច សង្កេតការណ៍ ដែរ ព្រោះមើលបានគ្រប់ស្រុក ====
  var observerTierCfg = getObserverTierRoleConfigByAccountType_(String(newUser.accountType || ""));
  var role, district, commune = "";
  if (isObserver) {
    role = ROLE_OBSERVER;
    district = "";
  } else if (observerTierCfg) {
    if (!isObserverTierRoleEnabled_(observerTierCfg.role)) {
      return { success: false, message: "ប្រព័ន្ធមិនទាន់បើកអនុញ្ញាតឲ្យបង្កើតគណនី" + observerTierCfg.label + "ទេ! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
    }
    role = observerTierCfg.role;
    district = "";
  } else if (isCommuneSignup) {
    // ==== FIX (Fix127, "ជម្រើសបើក/បិទប្រើប្រាស់គណនីឃុំសង្កាត់")៖ បើ SuperAdmin/Admin បិទមុខងារនេះ (ការកំណត់ប្រព័ន្ធ)
    // មិនអនុញ្ញាតឲ្យស្នើសុំគណនីឃុំសង្កាត់ដោយខ្លួនឯងទៀតទេ — ត្រូវត្រួតពិនិត្យមុនការផ្ទៀងផ្ទាត់ស្រុក/ឃុំសង្កាត់ខាងក្រោម ====
    if (!isCommuneAccountsEnabled_()) {
      return { success: false, message: "ប្រព័ន្ធមិនទាន់បើកអនុញ្ញាតឲ្យប្រើប្រាស់គណនីឃុំសង្កាត់ទេ! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
    }
    if (!String(newUser.district || "").trim() || !String(newUser.commune || "").trim()) {
      return { success: false, message: "សូមបំពេញព័ត៌មានឲ្យបានគ្រប់គ្រាន់!" };
    }
    if (!isValidCommune_(newUser.district, newUser.commune)) {
      return { success: false, message: "សូមជ្រើសរើសស្រុក/ក្រុង និងឃុំ/សង្កាត់ឲ្យត្រឹមត្រូវ!" };
    }
    role = ROLE_COMMUNE;
    district = newUser.district;
    commune = newUser.commune;
  } else {
    if (!String(newUser.district || "").trim()) {
      return { success: false, message: "សូមបំពេញព័ត៌មានឲ្យបានគ្រប់គ្រាន់!" };
    }
    if (DISTRICT_LIST.indexOf(newUser.district) === -1) {
      return { success: false, message: "សូមជ្រើសរើសស្រុក/ក្រុងឲ្យត្រឹមត្រូវ!" };
    }
    role = newUser.district;
    district = newUser.district;
  }
  if (String(newUser.password).trim().length < 6) {
    return { success: false, message: "លេខសម្ងាត់ត្រូវមានយ៉ាងតិច ៦ តួអក្សរ!" };
  }
  // ==== Gmail ក្លាយជាជម្រើស — បើទុកទទេ គណនីនេះប្រើបានតែតាមរយៈ Form ក្នុង App ប៉ុណ្ណោះ (គ្មានសិទ្ធិចូល Google Sheet ដោយផ្ទាល់ទេ) ====
  // ត្រូវផ្ទៀងផ្ទាត់ទម្រង់ត្រឹមតែពេលមានការបំពេញ Gmail ចូលមកប៉ុណ្ណោះ
  var gmail = String(newUser.gmail || "").trim();
  if (gmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(gmail)) {
    return { success: false, message: "សូមបញ្ចូល Gmail ឲ្យត្រឹមត្រូវ (ឬទុកទទេ ប្រសិនបើមិនត្រូវការចូល Google Sheet ដោយផ្ទាល់)!" };
  }

  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  // ==== FIX (ការពារការប៉ះទង្គិចគ្នា — Race Condition)៖ ការអាន (ត្រួតពិនិត្យស្ទួន) → សរសេរ (appendRow) ខាងក្រោម
  // មិនមែនប្រតិបត្តិការតែមួយឥតដាច់ទេ — បើសំណើ ២ ស្នើសុំគណនីស្ទើរតែក្នុងពេលតែមួយសម្រាប់ឈ្មោះគណនីដូចគ្នា (ឧ. Script
  // ស្វ័យប្រវត្តិ ឬចៃដន្យ ២ Browser Tab) ទាំងពីរអាចឆ្លងកាត់ការត្រួតពិនិត្យស្ទួនដូចគ្នា មុននឹងម្នាក់ណាបាន appendRow
  // ធ្វើឲ្យមានជួរដេកគណនីស្ទួនគ្នា ២ ក្នុងឈ្មោះតែមួយ (ច្របូកច្របល់ក្រោយមក — មានតែជួរដេកទីមួយប៉ុណ្ណោះដែល loginUser
  // ចូលដល់) — ចាក់សោខ្លីមួយភ្លែត ធានាថាមានតែសំណើមួយប៉ុណ្ណោះ ត្រួតពិនិត្យ+សរសេរក្នុងពេលតែមួយ ====
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(15000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានការស្នើសុំគណនីផ្សេងទៀតកំពុងដំណើរការ)! សូមព្យាយាមម្តងទៀត។" };
  }
  try {
    var data = sheet.getDataRange().getDisplayValues();
    var inputUser = String(newUser.username).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === inputUser) {
        return { success: false, message: "ឈ្មោះគណនីនេះមានអ្នកប្រើរួចហើយ!" };
      }
    }
    // ==== FIX (សំណើថ្មី "កំណត់ចំនួនកំរិតគណនី")៖ មើលការពន្យល់ពេញលេញនៅ countDistrictRoleAccounts_()/
    // countCommuneRoleAccountsForCommune_() ខាងលើ — ត្រូវត្រួតពិនិត្យត្រង់នេះ (ក្នុង Critical Section ដដែលនឹងការ
    // ត្រួតពិនិត្យឈ្មោះគណនីស្ទួនខាងលើ) ដើម្បីជៀសវាង Race Condition ====
    if (isCommuneSignup) {
      if (countCommuneRoleAccountsForCommune_(data, district, commune) >= COMMUNE_ACCOUNT_MAX_PER_COMMUNE_) {
        return { success: false, message: "ឃុំ/សង្កាត់ \"" + commune + "\" មានគណនីឃុំសង្កាត់គ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ + " គណនីក្នុងមួយឃុំសង្កាត់)! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
      }
    } else if (!isObserver && !observerTierCfg) {
      // ==== នេះជាគណនីគ្រូប្រចាំក្រុងស្រុកធម្មតា (role = ឈ្មោះស្រុកខ្លួនឯង) — សង្កេតការណ៍/តួនាទីថ្មីទាំង៥ មិនកំណត់
      // ចំនួន (គ្មានស្រុកជាក់លាក់ ដូច្នេះមិនអាចរាប់តាមស្រុកបានទេ) ====
      if (countDistrictRoleAccounts_(data, district) >= DISTRICT_ACCOUNT_MAX_PER_DISTRICT_) {
        return { success: false, message: "ស្រុក/ក្រុង \"" + district + "\" មានគណនីគ្រូប្រចាំក្រុងស្រុកគ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ + " គណនីក្នុងមួយស្រុក)! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
      }
    }
    var keyNumber = generateKeyNumber_();
    sheet.appendRow([
      newId_(), newUser.username.trim(), makePasswordHash_(newUser.password), sanitizeForSheetCell_(newUser.fullName.trim()),
      role, district, STATUS_PENDING, formatNow_(), gmail, keyNumber, "", "", commune
    ]);
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
  // ចំណាំ៖ មិនហៅ syncSheetPermissions_() នៅត្រង់នេះទេ — គណនីនៅមិនទាន់អនុម័ត មិនត្រូវផ្តល់សិទ្ធិចូល Google Sheet នៅឡើយទេ
  var s = getSystemSettings_();
  var appName = s.orgName || "ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ ខេត្តតាកែវ";
  // ==== FIX (Fix118, "សូមឲ្យការផ្ញើសារ Key Number ... ភ្លាមគឺ ផ្ញើភ្លាមតែម្តង")៖ ប្រើ notifyTelegramImmediate_()
  // (ផ្ញើផ្ទាល់ភ្លាមៗ) ជំនួស notifyTelegram_() (ដាក់ជួរដេក) សម្រាប់តែសារ Key Number នេះប៉ុណ្ណោះ — ព្រោះជាព័ត៌មាន
  // បន្ទាន់ (SuperAdmin ត្រូវការឆាប់ជូន Key Number ទៅអ្នកស្នើសុំ) ហើយស្នើសុំគណនីជាសកម្មភាពកម្រកើតឡើង (មិនប៉ះពាល់
  // ល្បឿនទូទៅ) ====
  notifyTelegramImmediate_(
    "🔑 មានការស្នើសុំគណនីថ្មី៖\n" +
    "កម្មវិធី៖ " + appName + "\n" +
    "ឈ្មោះគណនី៖ " + newUser.username.trim() + "\n" +
    (isObserver
      ? "តួនាទី៖ ភ្នាក់ងារសង្កេតការណ៍ (Observer — មើលបានប៉ុណ្ណោះ)\n"
      : observerTierCfg
        ? ("តួនាទី៖ " + observerTierCfg.label + " (មើលបានប៉ុណ្ណោះ)\n")
        : isCommuneSignup
          ? ("ស្ថាប័ន៖ " + district + " → ឃុំ/សង្កាត់ " + commune + "\nតួនាទី៖ គណនីឃុំសង្កាត់\n")
          : ("ស្ថាប័ន៖ " + district + "\nតួនាទី៖ គ្រូប្រចាំស្រុក\n")) +
    "Key Number៖ " + keyNumber + "\n\n" +
    "សូមផ្ញើរ Key Number នេះទៅឲ្យអ្នកប្រើប្រាស់ខាងលើ ដើម្បីបញ្ចប់ការបង្កើតគណនី (ឬគ្រូប្រចាំក្រុងស្រុក/SuperAdmin/Admin/PEC21 អាចអនុម័តដោយផ្ទាល់ក៏បាន)។"
  );
  return { success: true, username: newUser.username.trim(), message: "សំណើចុះឈ្មោះបានផ្ញើទៅ SuperAdmin/គ្រូប្រចាំក្រុងស្រុករួចហើយ!" };
}

// បញ្ជាក់ Key Number (សាធារណៈ — មិនត្រូវការចូលប្រព័ន្ធជាមុនទេ) ដើម្បីបញ្ចប់ការបង្កើតគណនី ក្រោយពី SuperAdmin
// ផ្តល់ Key Number ជូនអ្នកស្នើសុំ (តាមរយៈទំនាក់ទំនងផ្ទាល់ខ្លួន ក្រៅពី App) ។ ត្រូវផ្គូផ្គងទាំង Username និង Key Number
// ត្រឹមត្រូវ ទើបប្តូរស្ថានភាពពី "រង់ចាំអនុម័ត" ទៅ "សកម្ម" ហើយធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ភ្លាមៗ។
// ==== FIX (សុវត្ថិភាព — ការទាយ Key Number ដោយស្វ័យប្រវត្តិ/Brute-force)៖ Function នេះមិនត្រូវការចូល Login ជាមុនទេ
// (សាធារណៈ) ហើយ Key Number ជាលេខ៦ខ្ទង់តែប៉ុណ្ណោះ (~១លានលទ្ធភាព) ដោយគ្មានការកំណត់ចំនួនដងសាកល្បងពីមុនមកទាល់តែសោះ —
// នរណាម្នាក់អាចស្នើសុំបង្កើតគណនីខ្លួនឯង (requestSignup) ជ្រើសរើសលេខសម្ងាត់ផ្ទាល់ខ្លួន រួចសរសេរ Script ហៅ Function
// នេះឡើងវិញរាប់ពាន់/ម៉ឺនដងស្វ័យប្រវត្តិ ដើម្បីទាយ Key Number ត្រឹមត្រូវ ដោយមិនចាំបាច់រង់ចាំ SuperAdmin ផ្តល់ជូនផ្ទាល់
// ដែលជាជំហានផ្ទៀងផ្ទាត់សំខាន់មុនអនុម័តគណនី — ត្រូវកំណត់ចំនួនដងសាកល្បងខុសក្នុងមួយ Username (CacheService អាចរក្សា
// ទុកបានរហូតដល់ ៦ម៉ោង) ដើម្បីទប់ស្កាត់ការទាយស្វ័យប្រវត្តិ ====
var KEY_NUMBER_MAX_ATTEMPTS_ = 10;
var KEY_NUMBER_LOCKOUT_SECONDS_ = 3600; // ១ម៉ោង

function confirmKeyNumber(username, keyNumber) {
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var data = sheet.getDataRange().getValues();
  var inputUser = String(username || "").trim().toLowerCase();

  var cache = CacheService.getScriptCache();
  var attemptKey = "keyNumAttempts_" + inputUser;
  var attempts = 0;
  try { attempts = Number(cache.get(attemptKey)) || 0; } catch (eCacheGet) {}
  if (attempts >= KEY_NUMBER_MAX_ATTEMPTS_) {
    return { success: false, message: "សាកល្បងខុសច្រើនដងពេក! សូមរង់ចាំមួយសន្ទុះ (រហូតដល់ ១ម៉ោង) ឬទាក់ទង SuperAdmin ដើម្បីទទួល Key Number ជាថ្មី។" };
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) {
      if (data[i][6] !== STATUS_PENDING) {
        return { success: false, message: "គណនីនេះលែងស្ថិតក្នុងស្ថានភាពរង់ចាំ Key Number ទៀតហើយ។ សូមសាកល្បងចូលប្រព័ន្ធវិញ។" };
      }
      var storedKey = String(data[i][9] || "").trim();
      if (!storedKey || !timingSafeEqual_(storedKey, String(keyNumber || "").trim())) {
        try { cache.put(attemptKey, String(attempts + 1), KEY_NUMBER_LOCKOUT_SECONDS_); } catch (eCachePut) {}
        return { success: false, message: "Key Number មិនត្រឹមត្រូវទេ! សូមពិនិត្យម្តងទៀត ឬទាក់ទង SuperAdmin។" };
      }
      try { cache.remove(attemptKey); } catch (eCacheRm) {}
      sheet.getRange(i + 1, 7).setValue(STATUS_ACTIVE);
      // ==== FIX (ល្បឿន)៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជាផ្ទៃខាងក្រោយ (មិនចាំបាច់ឲ្យអ្នកប្រើរង់ចាំ — មើល Utils.gs) ====
      scheduleSyncSheetPermissions_();
      return { success: true, message: "បញ្ជាក់ Key Number ជោគជ័យ! គណនីរបស់អ្នកបានដំណើរការហើយ សូមចូលប្រើប្រាស់។" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// អនុម័តគណនីស្នើសុំ (ប្តូរស្ថានភាពពី "រង់ចាំអនុម័ត" ទៅ "សកម្ម" + ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ភ្លាមៗ)
function approveUserAccount(currentUsername, userId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ឥឡូវអនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុក អនុម័តគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនបានដែរ
  // (មិនតម្រូវ isAdmin_ ទៀតទេ) — ត្រួតពិនិត្យបែបគ្រោងគ្រាំងនៅត្រង់នេះសិន, លម្អិត (ត្រូវជាគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួន
  // ប៉ុណ្ណោះ) ក្រោយពីរកឃើញជួរដេកគោលដៅខាងក្រោម ====
  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិអនុម័តគណនីទេ!" };
  }
  // ==== FIX (ល្បឿន)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_()/getUserRow_() ខាងលើ
  // (ដែលបានស្កេន Sheet ទាំងមូលរួចហើយម្តង) ជំនួសការអាន Sheet ទាំងមូលម្តងទៀត (ដូចគ្នានឹង Fix120 — មើល
  // getUsersList()/getCommuneUsersList()) ====
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (!canManageTargetAccount_(currentUsername, data[i][4], data[i][5])) {
        return { success: false, message: "គ្មានសិទ្ធិអនុម័តគណនីនេះទេ!" };
      }
      sheet.getRange(i + 1, 7).setValue(STATUS_ACTIVE);
      // ==== FIX (ល្បឿន)៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជាផ្ទៃខាងក្រោយ (មិនចាំបាច់ឲ្យ Admin រង់ចាំ — មើល Utils.gs) ====
      scheduleSyncSheetPermissions_();
      return { success: true, message: "អនុម័តគណនីជោគជ័យ! អ្នកប្រើអាចចូលប្រើប្រាស់បានហើយ។" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// ប្តូរស្ថានភាព "សកម្ម" ⇄ "ផ្អាក" (សម្រាប់គណនីដែលអនុម័តរួចហើយ — ប្រើកាត់ការចូលប្រើប្រាស់ជាបណ្តោះអាសន្ន)
function toggleUserSuspend(currentUsername, userId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ឥឡូវអនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុក ផ្អាក/ធ្វើសកម្មគណនីឃុំសង្កាត់ក្នុងស្រុក
  // ខ្លួនបានដែរ ====
  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិកែប្រែស្ថានភាពគណនីទេ!" };
  }
  // ==== FIX (ល្បឿន)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_()/getUserRow_() ខាងលើ
  // (ដែលបានស្កេន Sheet ទាំងមូលរួចហើយម្តង) ជំនួសការអាន Sheet ទាំងមូលម្តងទៀត ====
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (data[i][1] === currentUsername) return { success: false, message: "អ្នកមិនអាចផ្អាកគណនីខ្លួនឯងបានទេ!" };
      if (!canManageTargetAccount_(currentUsername, data[i][4], data[i][5])) {
        return { success: false, message: "គ្មានសិទ្ធិកែប្រែស្ថានភាពគណនីនេះទេ!" };
      }
      // ==== FIX (សុវត្ថិភាព — សំខាន់ណាស់)៖ ដូចគ្នានឹង updateUserAccount()/deleteUserAccount() ខាងលើ — Admin/PEC21
      // មិនត្រូវអាច Suspend គណនី SuperAdmin (អ្នកដទៃ) បានទេ (ធ្វើឲ្យ SuperAdmin ត្រូវជាប់គាំង Login មិនកើត) ====
      if (data[i][4] === ROLE_SUPERADMIN && !isSuperAdmin_(currentUsername)) {
        return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាចផ្អាកគណនី SuperAdmin បាន!" };
      }
      var newStatus = data[i][6] === STATUS_ACTIVE ? STATUS_SUSPENDED : STATUS_ACTIVE;
      sheet.getRange(i + 1, 7).setValue(newStatus);
      // ==== ចេតនាហៅផ្ទាល់ (Synchronous) នៅតែដដែល មិនប្តូរទៅ scheduleSyncSheetPermissions_() ដូចកន្លែងផ្សេងទេ — ព្រោះ
      // ការផ្អាកគណនី ត្រូវការដកសិទ្ធិ Google Sheet ភ្លាមៗ ជាហេតុផលសុវត្ថិភាព (កុំឲ្យគណនីដែលទើបផ្អាក នៅតែបើក Google
      // Sheet ដោយផ្ទាល់ (មិនកាត់តាម App) កែប្រែទិន្នន័យបានទៀតរយៈពេលខ្លីមួយ ខណៈកំពុងរង់ចាំ Trigger ដំណើរការ) ====
      try { syncSheetPermissions_(); } catch (err) {}
      return { success: true, message: newStatus === STATUS_ACTIVE ? "ធ្វើសកម្មគណនីជោគជ័យ!" : "ផ្អាកគណនីជោគជ័យ!" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// ==================== ៤. ការចូល (Login) ====================
// ==== ជូនដំណឹងទៅ SuperAdmin តាម Telegram ពេលគណនីគ្រូប្រចាំស្រុក (ឬគណនីឃុំសង្កាត់ — Fix123) Login ឬ Logout ====
// ==== FIX (Fix122, "សូមឲ្យការ Login ចូលមានដំណើរការលឿន...")៖ eventLabel អាចជា String តែមួយ (ដូចមុន — Logout ។ល។)
// ឬជា Array នៃ String ច្រើន (ករណី Login ដែលមានព្រឹត្តិការណ៍ជាច្រើនក្នុងពេលតែមួយ ឧ. "ឧបករណ៍ថ្មី" + "Login") — ដើម្បីបញ្ចូល
// គ្នាទៅជាសារ Telegram តែមួយ/ការហៅ notifyTelegramImmediate_() តែម្តងគត់ ជំនួសការហៅច្រើនដងជាប់គ្នា (ដែលនីមួយៗជា Blocking
// Call ទៅ Telegram API — ២ដងស្ទួនគ្នាធ្វើឲ្យ Login យឺតទ្វេដង ក្នុងករណីប្តូរឧបករណ៍ថ្មី — មើល loginUser() ផងដែរ) ====
// ==== FIX (Fix123, "នៅពេលដែលគណនីឃុំសង្កាត់ login/logout មិនបានផ្ញើសារមក SuperAdmin")៖ Parameter commune ថ្មី
// (Optional — ត្រូវការតែសម្រាប់គណនីឃុំសង្កាត់ — ROLE_COMMUNE) ដើម្បីឲ្យសារបង្ហាញឈ្មោះឃុំ/សង្កាត់ជាក់លាក់ជំនួសពាក្យ
// "គ្រូប្រចាំ..." ដែលមិនត្រឹមត្រូវសម្រាប់គណនីកម្រិតឃុំសង្កាត់ (ជាអ្នកបញ្ចូលទិន្នន័យ មិនមែនគ្រូប្រចាំក្រុងស្រុកទេ)។
// ==== FIX (Fix123, កំហុសមួយទៀតដែលរកឃើញ)៖ districtType មុននេះគណនាពី DISTRICT_TYPE[role] (សន្មតថា role = ឈ្មោះស្រុក
// ដូចលំនាំគណនីស្រុក) — ត្រឹមត្រូវសម្រាប់គណនីស្រុក ប៉ុន្តែខុសសម្រាប់គណនីឃុំសង្កាត់ (role = ROLE_COMMUNE ជានិច្ច មិនមែន
// ឈ្មោះស្រុកទេ) ធ្វើឲ្យ "ស្ថាប័ន៖" បង្ហាញ "ស្រុក" ខុសសម្រាប់គណនីឃុំសង្កាត់ក្នុងក្រុងដូនកែវ។ ត្រូវប្រើ getDistrictType_
// (district) ជំនួសវិញ (គណនាពីឈ្មោះស្រុកផ្ទាល់ ត្រឹមត្រូវសម្រាប់គណនីទាំង២ប្រភេទ) ====
function notifyLoginEvent_(eventLabel, fullName, username, district, role, deviceInfo, commune) {
  try {
    var s = getSystemSettings_();
    if (!s.tgBotToken || !s.tgChatId) return; // មិនបានកំណត់ Telegram — រំលង
    var districtType = getDistrictType_(district);
    var appName = s.orgName || "ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ ខេត្តតាកែវ";
    var tz = Session.getScriptTimeZone() || "Asia/Phnom_Penh";
    var now = Utilities.formatDate(new Date(), tz, "dd-MM-yyyy HH:mm:ss");
    var eventLabelText = Array.isArray(eventLabel) ? eventLabel.join("\n") : eventLabel;
    var roleLabel = isCommuneRole_(role)
      ? ("អ្នកបញ្ចូលទិន្នន័យ ឃុំ/សង្កាត់" + (commune ? (" " + commune) : "") + " (" + districtType + district + ")")
      : ("គ្រូប្រចាំ" + districtType + district);
    var text = eventLabelText + "\n"
      + "កម្មវិធី៖ " + appName + "\n"
      + "ឈ្មោះគណនី៖ " + fullName + " (" + username + ")\n"
      + "ស្ថាប័ន៖ " + districtType + district + "\n"
      + "តួនាទី៖ " + roleLabel + "\n"
      + "ព័ត៌មានឧបករណ៍៖ " + (deviceInfo || "មិនស្គាល់") + "\n"
      + "ពេលវេលា៖ " + now;
    // ==== FIX (សំណើថ្មី "សូមឲ្យ Login ដំណើរការលឿនជាងមុន")៖ ត្រឡប់ទៅប្រើជួរដេក (Queue+Trigger — notifyTelegram_())
    // វិញ ជំនួស notifyTelegramImmediate_() ដែលធ្លាប់ប្រើតាំងពី Fix121 (ការសម្រេចចិត្តពីមុន "ផ្ញើភ្លាមៗតែម្តង" ដើម្បីឲ្យ
    // SuperAdmin ដឹងភ្លាមៗ)។ Login/Logout គឺជាសកម្មភាពញឹកញាប់ជាងគេបំផុតក្នុងប្រព័ន្ធ (មិនដូច "ស្នើសុំគណនី" ដែលកម្រ —
    // មើល notifyTelegramImmediate_() ខាងលើ) ដូច្នេះការហៅ Telegram API ជា Synchronous (Blocking — ~១,៥វិនាទី
    // ក្នុងករណីអាក្រក់បំផុត) រាល់ពេល Login/Logout ជះឥទ្ធិពលដល់អ្នកប្រើប្រាស់ទាំងអស់ជានិច្ច។ អ្នកប្រើបានជ្រើសរើសដោយដឹងខ្លួន
    // ឲ្យ Login លឿនជាអាទិភាព (SuperAdmin ទទួលបានការជូនដំណឹងយឺតជាងមុនបន្តិច ជំនួសភ្លាមៗ — ជាធម្មតាក្នុងរយៈពេលពីរបីវិនាទី
    // ប៉ុន្តែអាចយូរជាងនេះក្នុងករណីកម្រ ព្រោះ Trigger មិនធានាពេលវេលាច្បាស់លាស់ទេ) ====
    notifyTelegram_(text);
  } catch (err) {}
}

function loginUser(username, password, deviceInfo) {
  if (!username || !password) return { success: false, message: "សូមបំពេញឈ្មោះគណនី និងលេខសម្ងាត់!" };
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var data = sheet.getDataRange().getDisplayValues();
  var inputUser = String(username).trim().toLowerCase();
  var inputPass = String(password).trim();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) {
      if (!verifyPassword_(inputPass, data[i][2])) {
        return { success: false, message: "ឈ្មោះគណនី ឬលេខសម្ងាត់មិនត្រឹមត្រូវទេ!" };
      }
      // ==== FIX (Password Hashing — ការប្តូរដោយស្វ័យប្រវត្តិ)៖ បើគណនីនេះនៅតែផ្ទុកលេខសម្ងាត់ជា Plaintext ចាស់
      // (មិនទាន់ត្រូវបាន Hash ដោយ migratePlaintextPasswordsIfNeeded_ ដោយហេតុផលណាមួយ) Login ជោគជ័យលើកនេះ ជា
      // ឱកាសល្អបំផុតដើម្បី Hash វាឥឡូវនេះភ្លាមៗ (ព្រោះយើងទើបតែផ្ទៀងផ្ទាត់ថាដឹងលេខសម្ងាត់ត្រឹមត្រូវមែន) ====
      if (!isPasswordHashed_(data[i][2])) {
        try { sheet.getRange(i + 1, 3).setValue(makePasswordHash_(inputPass)); } catch (errRehash) {}
      }
      if (data[i][6] === STATUS_PENDING) {
        return { success: false, pending: true, username: data[i][1], message: "គណនីរបស់អ្នកកំពុងរង់ចាំ Key Number ដើម្បីបញ្ចប់ការបង្កើតគណនី! សូមទាក់ទង SuperAdmin ដើម្បីទទួល Key Number ។" };
      }
      if (data[i][6] !== STATUS_ACTIVE) {
        return { success: false, message: "គណនីរបស់អ្នកត្រូវបានផ្អាក! សូមទាក់ទងអ្នកគ្រប់គ្រងប្រព័ន្ធ។" };
      }
      // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ គណនីសង្កេតការណ៍/គ្រូប្រចាំក្រុងស្រុក ដែលបានកន្លងផុត
      // ថ្ងៃផុតកំណត់ (កំណត់ដោយ SuperAdmin/Admin ក្នុងទំព័រគណនីអ្នកប្រើប្រាស់) មិនត្រូវអនុញ្ញាតឲ្យ Login ថ្មីបានទេ ====
      if (isAccountExpired_(data[i][4], data[i][11])) {
        return { success: false, message: "គណនីរបស់អ្នកបានផុតកំណត់រយៈពេលប្រើប្រាស់ហើយ! សូមទាក់ទងអ្នកគ្រប់គ្រងប្រព័ន្ធដើម្បីបន្ត។" };
      }
      // ==== FIX (Fix128, "គណនីឃុំសង្កាត់ដែលបង្កើតហើយ អាចជ្រើសរើសប្រើបាន ឬមិនអាចប្រើបានបន្តទៀត")៖ Admin អាចផ្អាកគណនី
      // ឃុំសង្កាត់ទាំងអស់ភ្លាមៗតែម្តង (កុងតាក់ធំ — មិនប៉ះពាល់ស្ថានភាព "សកម្ម"/"ផ្អាក" ដាច់ដោយឡែករបស់គណនីនីមួយៗទេ) —
      // ត្រូវត្រួតពិនិត្យត្រង់នេះ (មុន Login ជោគជ័យ) ដើម្បីបដិសេធភ្លាមៗជាមួយសារច្បាស់លាស់ ជំនួសឲ្យអនុញ្ញាត Login ដំបូង
      // ហើយក្រោយមកបរាជ័យរាល់សកម្មភាព (សារច្រឡំ) ====
      if (isCommuneRole_(data[i][4]) && !isCommuneAccountsUsageEnabled_()) {
        return { success: false, message: "គណនីឃុំសង្កាត់ត្រូវបានផ្អាកការប្រើប្រាស់ជាបណ្តោះអាសន្នដោយ Admin ប្រព័ន្ធ! សូមទាក់ទង Admin ប្រព័ន្ធ ឬគ្រូប្រចាំក្រុងស្រុករបស់អ្នក។" };
      }
      // ==== FIX (សំណើថ្មី "ត្រូវប្តូរលេខសម្ងាត់ដំបូង")៖ គណនីដែលបានបង្កើតដោយស្វ័យប្រវត្តិ (bulkCreateAllCommuneAccounts_
      // ខាងក្រោម) មានទង់ Flag នេះកំណត់ (ជួរឈរ N) — ត្រូវហាមឃាត់មិនឲ្យបន្តទៅជា Session ពេញលេញ (មិនបង្កើត sessionToken
      // ថ្មីទេ) រហូតដល់ប្តូរលេខសម្ងាត់+ឈ្មោះខ្លួនឯង (changePasswordFirstLogin ខាងក្រោម) ជាមុនសិន — ត្រូវត្រួតពិនិត្យ
      // ត្រង់នេះ (ក្រោយពីផ្ទៀងផ្ទាត់លេខសម្ងាត់/ស្ថានភាព/ថ្ងៃផុតកំណត់ខាងលើរួចរាល់ហើយ — ធានាថាមានតែអ្នកដឹងលេខសម្ងាត់
      // ត្រឹមត្រូវទេ ទើបដឹងថាគណនីនេះកំពុងតម្រូវឲ្យប្តូរ) ====
      if (String(data[i][13] || "") === "1") {
        return { success: false, mustChangePassword: true, username: data[i][1], fullName: data[i][3], message: "នេះជាការចូលប្រើប្រាស់លើកដំបូង! សូមប្តូរលេខសម្ងាត់ និងឈ្មោះគោត្តនាម-នាមជាមុនសិន។" };
      }

      // ==== ចាប់ផ្តើម Session ថ្មី — Login ត្រូវបានកំណត់ត្រឹមតែមួយឧបករណ៍ក្នុងពេលតែមួយ (Token ថ្មីជំនួសចាស់) ====
      var role = data[i][4];
      var district = data[i][5];
      var hadActiveSession = !!String(data[i][10] || "").trim(); // តើមានឧបករណ៍ផ្សេងកំពុង Login រួចហើយឬអត់ មុននឹង Login លើកនេះ
      var sessionToken = newId_();
      sheet.getRange(i + 1, 11).setValue(sessionToken);

      // ==== ជូនដំណឹង SuperAdmin (គណនីគ្រូប្រចាំស្រុក ឬគណនីឃុំសង្កាត់ — មិនរាប់បញ្ចូល SuperAdmin/Admin/PEC21 ខ្លួនឯងទេ) ====
      // ==== FIX (Fix123, "នៅពេលដែលគណនីឃុំសង្កាត់ login/logout មិនបានផ្ញើសារមក SuperAdmin")៖ លក្ខខណ្ឌមុននេះ
      // isDistrictRole_(role) ត្រឹមត្រូវសម្រាប់គណនីស្រុកតែប៉ុណ្ណោះ — គណនីឃុំសង្កាត់ (ROLE_COMMUNE) មិនស្ថិតក្នុង
      // DISTRICT_LIST ទេ ដូច្នេះ isDistrictRole_() ត្រឡប់ false ជានិច្ចសម្រាប់គណនីទាំងនេះ ធ្វើឲ្យគ្មានការជូនដំណឹង
      // Telegram ណាមួយផ្ញើសោះ ពេលគណនីឃុំសង្កាត់ Login/Logout — ត្រូវបន្ថែម isCommuneRole_(role) ផងដែរ ====
      if (isDistrictRole_(role) || isCommuneRole_(role)) {
        // ==== FIX (សំណើ)៖ បើគណនីនេះកំពុង Login នៅឧបករណ៍ផ្សេងស្រាប់ (hadActiveSession) មានន័យថាឧបករណ៍ចាស់នោះ
        // នឹងត្រូវ Logout ដោយស្វ័យប្រវត្តិភ្លាមៗ (Token ចាស់ត្រូវជំនួសដោយ Token ថ្មីខាងលើ) — ត្រូវជូនដំណឹង SuperAdmin
        // តាម Telegram ពីព្រឹត្តិការណ៍នេះផ្ទាល់ (ជំនួសការបង្ហាញ Popup ធំៗលើឧបករណ៍ចាស់ ដែលធ្វើឲ្យអ្នកប្រើភ័យស្លន់ស្លោ
        // ដោយឥតប្រយោជន៍ — មើល forceLogoutDueToNewDevice_() ក្នុង Index.html) ====
        // ==== FIX (Fix122)៖ មុននេះ ករណីប្តូរឧបករណ៍ថ្មី (hadActiveSession) ហៅ notifyLoginEvent_() ២ដងដាច់ដោយឡែក —
        // ម្តងៗជា Blocking Call ទៅ Telegram API (Fix121) ធ្វើឲ្យ Login យឺតទ្វេដងក្នុងករណីនេះ។ ឥឡូវបញ្ចូលគ្នាទៅជា
        // ការហៅតែម្តងគត់ (សារតែមួយមានព័ត៌មានទាំង២បន្ទាត់) ដើម្បីកាត់បន្ថយពេលវេលារង់ចាំ Telegram API ត្រឹមតែម្តង
        // ដដែល ខណៈ SuperAdmin នៅតែទទួលបានព័ត៌មានពេញលេញដូចមុន ====
        var loginEventLabels_ = [];
        if (hadActiveSession) loginEventLabels_.push("🔄 Login លើឧបករណ៍ថ្មី (ឧបករណ៍ចាស់ត្រូវបាន Logout ស្វ័យប្រវត្តិ)");
        loginEventLabels_.push("🟢 ចូលប្រើប្រាស់ (Login)");
        notifyLoginEvent_(loginEventLabels_, data[i][3], data[i][1], district, role, deviceInfo, data[i][12]);
      }

      return {
        success: true,
        wasAlreadyLoggedIn: hadActiveSession, // ជូនដំណឹងអតិថិជនប្រសិនបើគណនីនេះកំពុងបានប្រើនៅឧបករណ៍ផ្សេងរួចហើយ (ឥឡូវត្រូវបានជំនួស)
        id: data[i][0],
        username: data[i][1],
        fullName: data[i][3],
        role: role,
        district: district,
        // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ Client ត្រូវការដឹងឃុំសង្កាត់ខ្លួន ដើម្បីកំណត់ Table បញ្ចូលទិន្នន័យ
        // ត្រឹមតែជួរដេកឃុំសង្កាត់ខ្លួន (ជួរឈរ M — ទទេសម្រាប់តួនាទីផ្សេងទាំងអស់) ====
        commune: data[i][12] || "",
        districtType: DISTRICT_TYPE[role] || "ស្រុក",
        sessionToken: sessionToken
      };
    }
  }
  return { success: false, message: "ឈ្មោះគណនី ឬលេខសម្ងាត់មិនត្រឹមត្រូវទេ!" };
}

// ==== ត្រួតពិនិត្យ Session ដែលរក្សាទុកនៅឧបករណ៍នេះ នៅតែសុពលភាពដែរឬអត់ (សម្រាប់ Login ជាប់រហូត + រកឃើញភ្លាមៗ ពេលឧបករណ៍ផ្សេងបាន Login សារជាថ្មី) ====
function validateSession(username, sessionToken) {
  if (!username || !sessionToken) return { success: false, message: "Session មិនត្រឹមត្រូវទេ!" };
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var data = sheet.getDataRange().getDisplayValues();
  var inputUser = String(username).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) {
      if (data[i][6] !== STATUS_ACTIVE) return { success: false, message: "គណនីរបស់អ្នកត្រូវបានផ្អាក!" };
      // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ Session ដែលកំពុងបើកប្រើប្រាស់ស្រាប់ ត្រូវកាត់ចេញ
      // ភ្លាមៗ (មិនចាំបាច់រង់ចាំដល់ Login ថ្មី) ពេលកន្លងផុតថ្ងៃផុតកំណត់ — Client ត្រួតពិនិត្យមុខងារនេះរាល់ ៩០វិនាទី។
      // ចំណាំ៖ ចេតនាមិនដាក់ sessionExpired:true ទេ (ខុសពីករណី Login ឧបករណ៍ផ្សេង) ដើម្បីឲ្យ Client បង្ហាញសារ
      // "ផុតកំណត់" ជាក់លាក់នេះ ជំនួសសារទូទៅ "Login នៅឧបករណ៍ផ្សេង" (មើល startSessionPolling_ ក្នុង Index.html) ====
      if (isAccountExpired_(data[i][4], data[i][11])) {
        return { success: false, message: "គណនីរបស់អ្នកបានផុតកំណត់រយៈពេលប្រើប្រាស់ហើយ! សូមទាក់ទងអ្នកគ្រប់គ្រងប្រព័ន្ធដើម្បីបន្ត។" };
      }
      // ==== FIX (Fix128, "គណនីឃុំសង្កាត់ដែលបង្កើតហើយ អាចជ្រើសរើសប្រើបាន ឬមិនអាចប្រើបានបន្តទៀត")៖ បើ Admin ទើបតែ
      // បិទកុងតាក់នេះ ខណៈគណនីឃុំសង្កាត់កំពុង Login ប្រើប្រាស់ស្រាប់ — ត្រូវរកឃើញ ក្នុងការត្រួតពិនិត្យតាមកាលកំណត់
      // (startSessionPolling_ រាល់ ៩០វិនាទី ក្នុង Index.html) ហើយបង្ខំ Logout ចេញ ដូចគ្នានឹងករណី Suspend គណនីនោះ
      // ដោយផ្ទាល់ដែរ (មិនចាំបាច់រង់ចាំដល់ Refresh ដោយដៃ) ====
      if (isCommuneRole_(data[i][4]) && !isCommuneAccountsUsageEnabled_()) {
        return { success: false, message: "គណនីឃុំសង្កាត់ត្រូវបានផ្អាកការប្រើប្រាស់ជាបណ្តោះអាសន្នដោយ Admin ប្រព័ន្ធ! សូម Login ម្តងទៀត។" };
      }
      if (!timingSafeEqual_(String(data[i][10] || ""), String(sessionToken))) {
        return { success: false, sessionExpired: true, message: "គណនីនេះត្រូវបាន Login នៅឧបករណ៍ផ្សេងទៀត!" };
      }
      var role = data[i][4];
      return {
        success: true,
        id: data[i][0],
        username: data[i][1],
        fullName: data[i][3],
        role: role,
        district: data[i][5],
        commune: data[i][12] || "",
        districtType: DISTRICT_TYPE[role] || "ស្រុក",
        sessionToken: sessionToken
      };
    }
  }
  return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
}

// ==== Logout ដោយផ្ទាល់ (ចុចប៊ូតុង "ចាកចេញ") — សម្អាត Session Token ជូនដំណឹង SuperAdmin ====
function logoutUser(username, sessionToken, deviceInfo) {
  if (!username) return { success: false };
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var data = sheet.getDataRange().getDisplayValues();
  var inputUser = String(username).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) {
      // សម្អាត Token ត្រឹមតែក្នុងករណីវានៅតែជា Token របស់ឧបករណ៍នេះ (ជៀសវាងលុប Token ថ្មីជាង ប្រសិនបើ Login កន្លងទៅរួច)
      if (String(data[i][10] || "") === String(sessionToken)) {
        sheet.getRange(i + 1, 11).setValue("");
        // ==== FIX (កំហុសរាយការណ៍ "គណនីនៅតែបង្ហាញ Online ក្រោយ Logout" — ចំណុចទី៣)៖ សម្អាត Presence Cache ភ្លាមៗ
        // នៅទីនេះ (មិនរង់ចាំ TTL ធម្មតា ១៥វិនាទីទេ) ដូចគ្នានឹងការសម្អាត Token ខាងលើ — ត្រូវការពារក្នុង if ដូចគ្នា
        // (ត្រូវប្រាកដថា Token នេះពិតជា Token សកម្មបច្ចុប្បន្នមែន) ដើម្បីកុំឲ្យការហៅ logoutUser() ហួសសម័យ (ឧ. ពី Tab
        // ចាស់ដែល Token របស់វាត្រូវបានជំនួសរួចហើយដោយ Login ថ្មីនៅឧបករណ៍ផ្សេង) បំបាត់ Online ខុសសម្រាប់ Session
        // ថ្មីដែលកំពុងសកម្មពិតប្រាកដឡើយ ====
        clearPresence_(username);
      }
      var role = data[i][4];
      // ==== FIX (Fix123)៖ ដូចគ្នានឹង loginUser() ដែរ — ត្រូវបន្ថែម isCommuneRole_(role) ដើម្បីឲ្យគណនីឃុំសង្កាត់
      // ក៏ទទួលបានការជូនដំណឹង Logout ជូន SuperAdmin ដែរ (មិនត្រឹមតែគណនីស្រុកទេ) ====
      if (isDistrictRole_(role) || isCommuneRole_(role)) {
        notifyLoginEvent_("🔴 ចាកចេញ (Logout)", data[i][3], data[i][1], data[i][5], role, deviceInfo, data[i][12]);
      }
      return { success: true };
    }
  }
  return { success: false };
}

// ==== ល្បឿន៖ ការចាំ (Memoize) ត្រឹមតែក្នុងមួយ Execution (Request) តែមួយប៉ុណ្ណោះ — មិនចែករំលែងរវាង Request ផ្សេងគ្នាទេ ====
// getUserRole_/getUserDistrict_/isAdmin_/canViewAllDistricts_/canEditDistrict_ សុទ្ធតែហៅ getUserRow_ ដោយឡែកៗពីគ្នា
// ហើយមុនកន្លងមក ម្តងហៅម្តងអាន+ស្កេន Sheet(អ្នកប្រើប្រាស់) ទាំងមូលឡើងវិញរាល់ដង — សំណើមួយអាចហៅដល់ ២-៤ដងសម្រាប់អ្នកប្រើតែម្នាក់
// ==== FIX (Fix120, "ចាប់ទិន្នន័យគណនីទាំងអស់ឲ្យលឿន")៖ getUsersList()/getCommuneUsersList() ជាដើម ក៏ត្រូវការទិន្នន័យ
// Sheet(អ្នកប្រើប្រាស់) ទាំងមូលដែរ (មិនមែនអ្នកប្រើតែម្នាក់ដូច getUserRow_) — មុននេះនីមួយៗអាន Sheet ទាំងមូលដោយឡែក
// ដាច់ដោយឡែក ខណៈពេលដែល requireValidSession_() (ហៅរាល់ Server Function ជានិច្ច ជាមុនគេ) បានហៅ getUserRow_() ស្កេន
// Sheet ទាំងមូលរួចហើយផងដែរ — ស្មើនឹងអានទិន្នន័យតែមួយដដែលៗ ២ដងក្នុងសំណើតែមួយ។ ឥឡូវ ចែករំលែក "ទិន្នន័យទាំងមូល" នេះ
// តាមរយៈ getUsersSheetDisplayDataOnce_() ជំនួសវិញ — អានតែម្តងគត់ក្នុងមួយ Execution ====
var USERS_SHEET_DISPLAY_DATA_CACHE_ = null;
function getUsersSheetDisplayDataOnce_() {
  if (USERS_SHEET_DISPLAY_DATA_CACHE_) return USERS_SHEET_DISPLAY_DATA_CACHE_;
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  USERS_SHEET_DISPLAY_DATA_CACHE_ = { sheet: sheet, data: sheet.getDataRange().getDisplayValues() };
  return USERS_SHEET_DISPLAY_DATA_CACHE_;
}

var USER_ROW_CACHE_ = {};
function getUserRow_(username) {
  var inputUser = String(username || "").trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(USER_ROW_CACHE_, inputUser)) return USER_ROW_CACHE_[inputUser];
  var data = getUsersSheetDisplayDataOnce_().data;
  var found = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) { found = data[i]; break; }
  }
  USER_ROW_CACHE_[inputUser] = found;
  return found;
}
function getUserRole_(username) {
  var row = getUserRow_(username);
  return row ? row[4] : null;
}
function getUserDistrict_(username) {
  var row = getUserRow_(username);
  return row ? row[5] : null;
}
function getUserFullName_(username) {
  var row = getUserRow_(username);
  return row ? row[3] : null;
}
// ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាពគ្រប់ការបញ្ចូល/កែប្រែទិន្នន័យ" — ចំណុចទី៥)៖ ជួរឈរ "អ្នកបញ្ចូល" ក្នុង Period Sheet
// (DAILY_META_SUFFIX — មើល Utils.gs) ត្រូវបង្ហាញទាំង ឈ្មោះគណនី (Username) និង ឈ្មោះគោត្តនាម-នាម ព្រមគ្នា (តាមសំណើ
// ផ្ទាល់៖ "ជួរឈរអ្នកបញ្ចូលមាន(ឈ្មោះគណនី+ឈ្មោះគោត្តនាម នាម)") ជំនួសការបង្ហាញឈ្មោះម្នាក់ឯង (ដូចមុន)។ មុខងារនេះជា
// ចំណុចតែមួយគត់ (Single source of truth) ត្រូវប្រើទាំង saveDistrictDayEntries() និង upsertDailyEntry_() ទាំង២
// (PeriodSheets.gs) ដើម្បីធានាថាទម្រង់ដូចគ្នាបេះបិទរវាងផ្លូវសរសេរទាំង២ (ពីមុន ២ផ្លូវនេះមិនស៊ីគ្នា៖ Bulk-save សរសេរ
// ឈ្មោះម្នាក់ឯង ខណៈ Single-entry សរសេរ Username ឆៅម្នាក់ឯង) ==== Fallback៖ បើរកមិនឃើញឈ្មោះពេញ (ឧ. គណនីត្រូវបានលុប
// រួច) បង្ហាញត្រឹមតែ Username ដដែល (ជៀសវាងបង្ហាញ "username - username" ស្ទួនគ្នា) ====
function formatEnteredBy_(username) {
  var fullName = getUserFullName_(username);
  return fullName ? (username + ' - ' + fullName) : username;
}
function isAdmin_(username) {
  var role = getUserRole_(username);
  // ==== FIX៖ PEC21 ឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង (តាមសំណើ) — មុននេះ PEC21 មានតែសិទ្ធិមើលទាំងអស់
  // (canViewAllDistricts_) ប៉ុន្តែគ្មានសិទ្ធិកែប្រែ/គ្រប់គ្រងគណនី/ការកំណត់ប្រព័ន្ធ ដូច Admin ទេ ====
  return role === ROLE_ADMIN || role === ROLE_SUPERADMIN || role === ROLE_PEC21;
}
function isSuperAdmin_(username) {
  return getUserRole_(username) === ROLE_SUPERADMIN;
}
// SuperAdmin/Admin/PEC21/សង្កេតការណ៍ អាចមើលទិន្នន័យទាំងអស់ (គ្រប់ស្រុក) ។ គ្រូប្រចាំស្រុក មើលបានតែស្រុករបស់ខ្លួន
function canViewAllDistricts_(username) {
  var role = getUserRole_(username);
  return role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21 || role === ROLE_OBSERVER || isObserverTierRole_(role);
}
// បញ្ចូល/កែប្រែទិន្នន័យ៖ SuperAdmin/Admin/PEC21 បញ្ចូលបានគ្រប់ស្រុក, គ្រូប្រចាំស្រុកបញ្ចូលបានតែស្រុករបស់ខ្លួន, សង្កេតការណ៍ មើលបានប៉ុណ្ណោះ (មិនបញ្ចូល)
// ==== FIX៖ PEC21 ឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង (តាមសំណើ) — មុននេះ PEC21 = view-only ដូច សង្កេតការណ៍ ====
function canEditDistrict_(username, district) {
  var role = getUserRole_(username);
  if (role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21) return true;
  if (isDistrictRole_(role)) return role === district;
  return false; // សង្កេតការណ៍/គណនីឃុំសង្កាត់ = មិនអាចកែប្រែទិន្នន័យទាំងស្រុក (ឃុំសង្កាត់ប្រើ canEditCommune_ ខាងក្រោម)
}

// ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ជួរឈរ M ក្នុង Sheet(អ្នកប្រើប្រាស់) — ឃុំ/សង្កាត់ (មានតែសម្រាប់ ROLE_COMMUNE) ====
function getUserCommune_(username) {
  var row = getUserRow_(username);
  return row ? (row[12] || "") : null;
}

// ==== សិទ្ធិបញ្ចូល/កែប្រែទិន្នន័យប្រចាំថ្ងៃ ត្រឹមតែឃុំសង្កាត់ជាក់លាក់មួយ (ប្រើ Scope ការសរសេរសម្រាប់គណនីឃុំសង្កាត់
// ក្នុង PeriodSheets.gs) — SuperAdmin/Admin/PEC21 = គ្រប់ឃុំសង្កាត់ទាំងអស់, គ្រូប្រចាំក្រុងស្រុក = គ្រប់ឃុំសង្កាត់
// ក្នុងស្រុកខ្លួន (ដូចមុន, មិនប្តូរ), គណនីឃុំសង្កាត់ = ត្រឹមតែឃុំសង្កាត់ខ្លួនផ្ទាល់ប៉ុណ្ណោះ, សង្កេតការណ៍ = មិនអាចកែប្រែឡើយ ====
function canEditCommune_(username, district, commune) {
  var role = getUserRole_(username);
  if (role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21) return true;
  if (isDistrictRole_(role)) return role === district;
  if (isCommuneRole_(role)) {
    return getUserDistrict_(username) === district && getUserCommune_(username) === commune;
  }
  return false;
}

// ==== សិទ្ធិមើលរបាយការណ៍ ត្រឹមតែឃុំសង្កាត់ជាក់លាក់មួយ (ប្រើ Scope ការអានរបាយការណ៍សម្រាប់គណនីឃុំសង្កាត់ ក្នុង
// Reports.gs) — ដូច canEditCommune_ ប៉ុន្តែបន្ថែម SuperAdmin/Admin/PEC21/សង្កេតការណ៍ (មើលបានគ្រប់ទីកន្លែង) ====
function canViewCommune_(username, district, commune) {
  if (canViewAllDistricts_(username)) return true;
  var role = getUserRole_(username);
  if (isDistrictRole_(role)) return role === district;
  if (isCommuneRole_(role)) {
    return getUserDistrict_(username) === district && getUserCommune_(username) === commune;
  }
  return false;
}

// ==== សិទ្ធិគ្រប់គ្រងគណនីឃុំសង្កាត់ (បង្កើត/អនុម័ត/កែប្រែ/ផ្អាក/លុប) — SuperAdmin/Admin/PEC21 គ្រប់គ្រងបានគ្រប់ស្រុក,
// គ្រូប្រចាំក្រុងស្រុក គ្រប់គ្រងបានតែគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនប៉ុណ្ណោះ (ដូចសំណើ) ====
function canManageCommuneAccountsFor_(username, district) {
  if (isAdmin_(username)) return true;
  var role = getUserRole_(username);
  return isDistrictRole_(role) && role === district;
}

// ==== ត្រួតពិនិត្យលម្អិត៖ តើ currentUsername គ្រប់គ្រងគណនីគោលដៅ (targetRole/targetDistrict) នេះបានដែរឬទេ — ប្រើក្នុង
// approveUserAccount/updateUserAccount/deleteUserAccount/toggleUserSuspend ក្រោយពីរកឃើញជួរដេកគោលដៅ — Admin-tier
// គ្រប់គ្រងបានគ្រប់គណនី (ដូចមុន, មិនប្តូរ), គ្រូប្រចាំក្រុងស្រុក គ្រប់គ្រងបានតែគណនីឃុំសង្កាត់ (ROLE_COMMUNE) ក្នុងស្រុក
// ខ្លួនប៉ុណ្ណោះ (មិនអាចប៉ះគណនីប្រភេទផ្សេង ឬឃុំសង្កាត់ស្រុកដទៃបានឡើយ) ====
function canManageTargetAccount_(currentUsername, targetRole, targetDistrict) {
  if (isAdmin_(currentUsername)) return true;
  if (targetRole === ROLE_COMMUNE) return canManageCommuneAccountsFor_(currentUsername, targetDistrict);
  return false;
}

// ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ ធ្វើសញ្ញា Online/Offline ដូច Facebook/Messenger/
// Telegram សម្រាប់គណនីឃុំសង្កាត់/គ្រូប្រចាំក្រុងស្រុក (Tab "គណនីឃុំសង្កាត់"/"គណនីគ្រូប្រចាំក្រុងស្រុក") ។ ដោយសារ Google
// Apps Script HtmlService មិនគាំទ្រ WebSocket ទេ (ដូចគ្នានឹងបញ្ហា "ធ្វើបច្ចុប្បន្នភាពទិន្នន័យភ្លាមៗឆ្លងគណនី" ខាងលើ)
// ត្រូវប្រើ CacheService ជា "ស្លាកសំគាល់ពេលវេលាចុងក្រោយដែលឃើញសកម្មភាព" (Last-Seen Timestamp) ដូចគ្នា៖
//   • touchPresence_(username) — កត់ត្រា Date.now() ថ្មីសម្រាប់គណនីនោះ ជាមួយ TTL ខ្លីតែ ១៥វិនាទី (មិនប្រើ ២១៦០០
//     វិនាទីដូច Version-Tracking ខាងលើទេ ព្រោះ "Online ឥឡូវនេះ" ត្រូវការភាពស្រស់បំផុត — គណនីត្រូវលេចជា Offline
//     ភ្លាមៗក្រោយពីឈប់សកម្ម មិនមែនរង់ចាំរហូតដល់ ៦ម៉ោងទេ) — ==== FIX (សំណើថ្មី "សូមបន្ថយចំណុចនេះឲ្យត្រឹម ១៥វិនាទី
//     ពេលបិទ Browser ដោយផ្ទាល់")៖ ពីមុន ៤០វិនាទី ឥឡូវបន្ថយមក ១៥វិនាទីព្រម (heartbeatPresence()/startPresenceHeartbeat_
//     ក្នុង Index.html ក៏បន្ថយចន្លោះពេលពី ១៥វិនាទី → ៥វិនាទីស្របគ្នា ដើម្បីនៅតែមាន Margin ~២ដងគ្រប់គ្រាន់ទប់ Network
//     យឺត/Browser Throttle Background Tab មិនឲ្យ Online "ភ្លឹបភ្លែតៗ" ខុសពេលកំពុងសកម្មពិតប្រាកដ) — ចំណាំ៖ កាត់បន្ថយ
//     រយៈពេលនេះ = បង្កើន RPC Call ញឹកញាប់ជាងមុន (heartbeatPresence រាល់ ៥វិនាទី/គណនីសកម្មនីមួយៗ) ដូច្នេះជះឥទ្ធិពល
//     លើ Quota របស់ Google Apps Script ច្រើនជាងការកំណត់មុន — ត្រូវតាមដានប្រសិនបើមានគណនីសកម្មច្រើននាក់ក្នុងពេលតែមួយ
//   • isUserOnline_(username) — គ្រាន់តែពិនិត្យមើលថា Key នៅតែមាន (មិនផុតកំណត់) ក្នុង Cache ដែរឬអត់
//   • clearPresence_(username) — លុប Key ចោលភ្លាមៗ (មិនរង់ចាំ TTL ធម្មតាទេ) — ហៅពី logoutUser() ខាងក្រោម
//     ដើម្បីធានាថា Logout ច្បាស់លាស់ (ចុចប៊ូតុង "ចាកចេញ") ធ្វើឲ្យលេចជា Offline ភ្លាមៗ មិនចាំបាច់រង់ចាំ TTL ធម្មតា
//     ដល់ ១៥វិនាទីនោះទេ — ជួសជុលរបាយការណ៍ "គណនីនៅតែបង្ហាញ Online ក្រោយ Logout" ====
// touchPresence_() ត្រូវបានហៅពី ២ចំណុច៖ (១) requireValidSession_() ខាងក្រោម — ដូច្នេះរាល់ការហៅ Server ណាមួយ
// ដែលបានផ្ទៀងផ្ទាត់ Session ជោគជ័យ (ឧ. pollDistrictDayUpdates_ រាល់ ១០វិនាទី ពេលកំពុងបញ្ចូលទិន្នន័យប្រចាំថ្ងៃ,
// រក្សាទុកទិន្នន័យ, ប្តូរទំព័រ ។ល។) នឹងកត់សម្គាល់ថាគណនីនោះ "Online" ដោយស្វ័័យប្រវត្តិ ដោយមិនចាំបាច់ Call បន្ថែមទេ
// (២) heartbeatPresence() RPC ថ្មីខាងក្រោម — ហៅដោយ Client រៀងរាល់ ~៥វិនាទី ពេល App កំពុងបើក (មិនថាកំពុងនៅទំព័រ
// មួយណា) ធានាថាគណនីនៅតែបង្ហាញ Online ទោះជាទំព័រនោះមិនមាន Poll ផ្សេងទៀតក៏ដោយ (ឧ. កំពុងអាន Dashboard ធម្មតា)។
// ==== ចំណាំសំខាន់ (ដែនកំណត់ជាធម្មជាតិ — ដូចគ្នាទាំង Facebook/Messenger/Telegram ដែរ)៖ បើអ្នកប្រើប្រាស់បិទ Browser
// ដោយផ្ទាល់ (មិនចុច "ចាកចេញ") គ្មានវិធីណាដែល Server ដឹងភ្លាមៗបានទេ (គ្មាន Event "កំពុងបិទ" អាចទុកចិត្តបានឆ្លងកាត់
// google.script.run ដែលជា Asynchronous) — ក្នុងករណីនេះ គណនីនោះនឹងបន្តបង្ហាញ Online រហូតដល់ TTL ១៥វិនាទីផុតកំណត់
// ដោយខ្លួនឯង (មិនមែនកំហុសទេ — ជាឥរិយាបថដែលរំពឹងទុករបស់ប្រព័ន្ធ Presence ណាមួយក៏ដោយ) ====
function presenceCacheKey_(username) {
  return 'presence_' + String(username || '').trim().toLowerCase();
}
function touchPresence_(username) {
  try {
    CacheService.getScriptCache().put(presenceCacheKey_(username), String(Date.now()), 15);
  } catch (eTouchPresence_) {}
}
function isUserOnline_(username) {
  try {
    return !!CacheService.getScriptCache().get(presenceCacheKey_(username));
  } catch (eIsOnline_) { return false; }
}
function clearPresence_(username) {
  try {
    CacheService.getScriptCache().remove(presenceCacheKey_(username));
  } catch (eClearPresence_) {}
}

// ==================== ៤.១ ការផ្ទៀងផ្ទាត់ Session (Session Token Verification) — ការពារ Auth Bypass ====================
// ==== FIX (សំខាន់ណាស់បំផុត — Authorization Bypass)៖ មុននេះ មុខងារសំខាន់ៗស្ទើរតែទាំងអស់ (បង្កើត/លុប/កែគណនី,
// កំណត់ប្រព័ន្ធ, បញ្ចូល/កែ/លុបទិន្នន័យ, មតិយោបល់, Backup ។ល។) ជឿទុកចិត្តលើ Parameter "currentUsername" (String
// ធម្មតា) ដែលផ្ញើមកពី Browser ដោយផ្ទាល់ តាមរយៈ google.script.run ដោយគ្មានការផ្ទៀងផ្ទាត់ថាអ្នកហៅពិតជាបាន Login
// ជាអ្នកប្រើនោះមែនឬអត់ទេ — Google Apps Script អនុញ្ញាតឲ្យហៅមុខងារ Server ណាមួយក៏បាន (ទោះមិនមែនកូដ Client ខ្លួនឯង
// ហៅក៏ដោយ) ដូច្នេះនរណាម្នាក់ដែលបើក Web App នេះ អាចបើក Browser DevTools (F12) ហើយហៅ
// google.script.run.addUserAccount('admin', {...}) ដោយផ្ទាល់ (ដោយមិនចាំបាច់ដឹងលេខសម្ងាត់ធ្វើ Login សោះ សូម្បីតែ
// ម្តង) ដើម្បីបង្កើតគណនី SuperAdmin ថ្មីរបស់ខ្លួនផ្ទាល់ ឬលុប/កែគណនីអ្នកដទៃ គ្រាន់តែស្គាល់/ទាយ Username ត្រឹមត្រូវ
// ប៉ុណ្ណោះ (ឧ. "admin") — ការត្រួតពិនិត្យ isAdmin_(currentUsername) ដែលមានស្រាប់ ការពារបានតែក្នុងករណីអ្នកហៅមិន
// មែនជា Admin ប៉ុណ្ណោះ មិនបានផ្ទៀងផ្ទាត់ថា "អ្នកហៅ ពិតជាបានផ្ទៀងផ្ទាត់អត្តសញ្ញាណ (Authenticate) ជាអ្នកប្រើនោះ
// មែនឬអត់" ទេ។ ការជួសជុលនេះ តម្រូវឲ្យមុខងារសំខាន់ៗទាំងនោះ ទទួល sessionToken ជាប៉ារ៉ាម៉ែត្រចុងក្រោយបន្ថែម ហើយ
// ផ្ទៀងផ្ទាត់ថា Token នោះត្រូវគ្នានឹង Token ដែលរក្សាទុកសម្រាប់ Username នោះក្នុង Sheet(អ្នកប្រើប្រាស់) ជួរឈរ K
// (ដដែលនឹង Token ដែលបង្កើតដោយ loginUser() ហើយផ្ទៀងផ្ទាត់ដដែលដោយ validateSession() រាល់ ៩០វិនាទី) ជាមុនសិន
// មុននឹងបន្តអនុវត្តអ្វីផ្សេងទៀត — ស្មើនឹងទាមទារឲ្យអ្នកហៅ "កំពុង Login ជាអ្នកប្រើនោះពិតប្រាកដ" មិនមែនត្រឹមតែ "ដឹង
// Username" ទេ។ Token នេះក៏មិនអាចយកទៅប្រើឆ្លងគណនីបានដែរ ព្រោះមានតែមួយ Token សកម្មក្នុងមួយពេល ក្នុងមួយគណនី
// (Login ថ្មីលើឧបករណ៍ថ្មី = Token ចាស់ត្រូវលែងសុពលភាពភ្លាមៗ — មើល loginUser) ====
function requireValidSession_(username, sessionToken) {
  if (!username || !sessionToken) {
    return { valid: false, message: "សម័យប្រើប្រាស់របស់អ្នកមិនត្រឹមត្រូវ ឬបានផុតកំណត់! សូម Login ម្តងទៀត។" };
  }
  var row = getUserRow_(username);
  if (!row) {
    return { valid: false, message: "គណនីមិនត្រឹមត្រូវទេ! សូម Login ម្តងទៀត។" };
  }
  if (row[6] !== STATUS_ACTIVE) {
    return { valid: false, message: "គណនីរបស់អ្នកលែងសកម្មទៀតហើយ! សូម Login ម្តងទៀត។" };
  }
  // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ ការពារកុំឲ្យអនុវត្តមុខងារសំខាន់ៗណាមួយបន្ត ក្រោយពីគណនី
  // សង្កេតការណ៍/គ្រូប្រចាំក្រុងស្រុកនោះ បានផុតកំណត់រយៈពេលប្រើប្រាស់ (ដូចគ្នានឹងការត្រួតពិនិត្យ STATUS_ACTIVE ខាងលើ) ====
  if (isAccountExpired_(row[4], row[11])) {
    return { valid: false, message: "គណនីរបស់អ្នកបានផុតកំណត់រយៈពេលប្រើប្រាស់ហើយ! សូមទាក់ទងអ្នកគ្រប់គ្រងប្រព័ន្ធដើម្បីបន្ត។" };
  }
  // ==== FIX (Fix128, "គណនីឃុំសង្កាត់ដែលបង្កើតហើយ អាចជ្រើសរើសប្រើបាន ឬមិនអាចប្រើបានបន្តទៀត")៖ requireValidSession_()
  // ជាចំណុចត្រួតពិនិត្យកណ្តាលដែលមុខងារ Server សំខាន់ៗស្ទើរតែទាំងអស់ហៅមុនអនុវត្តអ្វីផ្សេង (បង្កើត/លុប/កែទិន្នន័យ,
  // មតិយោបល់, Backup ។ល។) — ត្រូវត្រួតពិនិត្យត្រង់នេះផងដែរ (មិនមែនត្រឹមតែ loginUser()/validateSession() ទេ) ដើម្បី
  // ធានាថាគណនីឃុំសង្កាត់ដែលកុងតាក់នេះទើបតែផ្អាក មិនអាចអនុវត្តសកម្មភាពណាមួយបន្តទៀតបានភ្លាមៗ (Defense in Depth —
  // មិនរង់ចាំដល់ការត្រួតពិនិត្យតាមកាលកំណត់ ៩០វិនាទីទេ) ====
  if (isCommuneRole_(row[4]) && !isCommuneAccountsUsageEnabled_()) {
    return { valid: false, message: "គណនីឃុំសង្កាត់ត្រូវបានផ្អាកការប្រើប្រាស់ជាបណ្តោះអាសន្នដោយ Admin ប្រព័ន្ធ! សូម Login ម្តងទៀត។" };
  }
  var storedToken = String(row[10] || "").trim();
  if (!storedToken || !timingSafeEqual_(storedToken, String(sessionToken).trim())) {
    return { valid: false, message: "សម័យប្រើប្រាស់របស់អ្នកបានផុតកំណត់ ឬបានប្តូរទៅ Login លើឧបករណ៍ផ្សេង! សូម Login ម្តងទៀត។" };
  }
  // ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ ចំណុចនេះជាកន្លែងតែមួយគត់ដែលមុខងារ Server ស្ទើរតែ
  // ទាំងអស់ហៅមុនអនុវត្តអ្វីផ្សេង ដូច្នេះកត់សម្គាល់ថាគណនីនេះ "Online" នៅទីនេះតែម្តង គ្របដណ្ដប់រាល់ការហៅ Server ណាមួយ
  // ដែលបានផ្ទៀងផ្ទាត់ Session ជោគជ័យ ដោយមិនចាំបាច់បន្ថែមកូដត្រង់ RPC នីមួយៗឡើយ — មើលការពន្យល់ពេញលេញនៅ
  // touchPresence_() ខាងលើ ====
  touchPresence_(username);
  return { valid: true };
}

// ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ RPC ស្រាលៗសម្រាប់ Client ហៅរៀងរាល់ ~២០វិនាទី
// (មើល startPresenceHeartbeat_ ក្នុង Index.html) ដើម្បីរក្សាគណនីនេះឲ្យបង្ហាញ "Online" ជានិច្ចពេល App កំពុងបើក ទោះជា
// កំពុងនៅទំព័រណាក៏ដោយ (requireValidSession_ ខាងលើ ជាអ្នកកត់សម្គាល់ Online ពិតប្រាកដ — Function នេះគ្រាន់តែជា
// មធ្យោបាយឲ្យ Client មាន "អ្វីមួយ" ត្រូវហៅជាទៀងទាត់ សូម្បីតែពេលមិននៅលើទំព័រមាន Poll ផ្សេងទៀត) ====
function heartbeatPresence(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };
  return { success: true };
}

// ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ RPC ស្រាលៗសម្រាប់ធ្វើបច្ចុប្បន្នភាព "ចំណុចខៀវ/ប្រផេះ"
// នៅតារាងគណនីឃុំសង្កាត់/គ្រូប្រចាំក្រុងស្រុក ដោយមិនចាំបាច់ផ្ទុកទិន្នន័យគណនីទាំងមូលឡើងវិញ (getCommuneUsersList/
// getDistrictTeacherUsersList) ជារៀងរាល់ដង — មើល pollAccountPresence_() ក្នុង Index.html — ត្រូវការសិទ្ធិដូចគ្នា
// បេះបិទនឹងអ្នកអាចមើលបញ្ជីគណនីទាំងនោះបាន (isAdmin_ ឬគ្រូប្រចាំក្រុងស្រុក — Defense in depth, កុំឲ្យអ្នកមិនពាក់ព័ន្ធ
// ដឹងថាគណនីណាកំពុង Online ដោយឥតការអនុញ្ញាត) ====
function getOnlinePresence(currentUsername, sessionToken, usernames) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };
  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  }
  var online = {};
  (usernames || []).forEach(function(u) { online[u] = isUserOnline_(String(u)); });
  return { success: true, online: online };
}

// ==================== ៥. គណនីអ្នកប្រើប្រាស់ (User Management — SuperAdmin/Admin only) ====================
function getUsersList(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  // ==== FIX (Fix120)៖ ប្រើទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ដែលបានអានរួចហើយពី requireValidSession_() ខាងលើ (ចែករំលែក
  // តាម getUsersSheetDisplayDataOnce_()) ជំនួសការអានទាំងមូលម្តងទៀត (ជៀសវាងស្ទួន — មើលពន្យល់លម្អិតនៅ getUserRow_) ====
  var data = getUsersSheetDisplayDataOnce_().data;
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][4];
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គណនីកម្រិតឃុំសង្កាត់ មិនបង្ហាញក្នុងបញ្ជីគណនីទូទៅនេះទេ — មានទំព័រគ្រប់គ្រង
    // ដាច់ដោយឡែកសម្រាប់ខ្លួន (មើល getCommuneUsersList) ដើម្បីកុំឲ្យលាយច្របល់ជាមួយគណនីគ្រូប្រចាំក្រុងស្រុក/Admin ====
    if (isCommuneRole_(role)) continue;
    // ==== FIX (Fix130, "Tab(គណនីគ្រូប្រចាំក្រុងស្រុក) ថ្មី")៖ គណនីគ្រូប្រចាំក្រុងស្រុកទាំង១០ ផ្លាស់ទីទៅ Tab ដាច់ដោយ
    // ឡែកថ្មី (មើល getDistrictTeacherUsersList) តម្រៀបតាមលេខរៀងក្រុងស្រុក (01-10) ជំនួសការលាយច្របល់ក្នុង Tab នេះ ====
    if (isDistrictRole_(role)) continue;
    // ==== FIX (Fix130, "Tab(គណនីភាគីពាក់ព័ន្ធ) ថ្មី")៖ ROLE_OBSERVER (ភ្នាក់ងារសង្កេតការណ៍) និងតួនាទីថ្មីទាំង៤ (មិន
    // រាប់បញ្ចូល "មន្ត្រី ល.ខ.ប តាកែវ" ដែលនៅតែផ្ទុកទីនេះដដែល) ផ្លាស់ទីទៅ Tab "គណនីភាគីពាក់ព័ន្ធ" ថ្មី (មើល
    // getStakeholderUsersList) ====
    if (role === ROLE_OBSERVER) continue;
    var __tierCfgForUsersTab_ = getObserverTierRoleConfigByRole_(role);
    if (__tierCfgForUsersTab_ && __tierCfgForUsersTab_.tab === 'stakeholder') continue;
    var expiryDate = normalizeDateStr_(data[i][11] || "");
    list.push({
      id: data[i][0], username: data[i][1], fullName: data[i][3],
      role: role, district: data[i][5], status: data[i][6], createdDate: data[i][7],
      gmail: data[i][8] || "", keyNumber: data[i][9] || "",
      // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ ព័ត៌មានសម្រាប់ UI បង្ហាញ/កែសម្រួល ====
      expiryDate: expiryDate, expiryApplicable: isExpiryApplicableRole_(role), expired: isAccountExpired_(role, expiryDate)
    });
  }
  // ==== FIX (Fix130, "Tab(គណនីគ្រូប្រចាំក្រុងស្រុក)/Tab(គណនីភាគីពាក់ព័ន្ធ) ថ្មី")៖ Tab(អ្នកប្រើប្រាស់) នេះឥឡូវផ្ទុក
  // តែ SuperAdmin/Admin/PEC21 + "មន្ត្រី ល.ខ.ប តាកែវ" ប៉ុណ្ណោះ (តួនាទីទាំង៥ផ្សេងទៀត + ROLE_OBSERVER + ស្រុកទាំង១០
  // ផ្លាស់ទីទៅ Tab ថ្មីរៀងៗខ្លួនហើយ — មើលខាងលើ) ====
  var usersTabTierRoleNames_ = OBSERVER_TIER_ROLES_.filter(function(cfg) { return cfg.tab === 'users'; }).map(function(cfg) { return cfg.role; });
  return {
    success: true, users: list, roles: [ROLE_SUPERADMIN, ROLE_ADMIN, ROLE_PEC21].concat(usersTabTierRoleNames_), districtType: DISTRICT_TYPE,
    expiryApplicableRoles: usersTabTierRoleNames_
  };
}

// ==================== ៥.០.១ គណនីគ្រូប្រចាំក្រុងស្រុក (District Teacher Accounts — SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ)
// — ទំព័រគ្រប់គ្រងដាច់ដោយឡែក (Fix130, "សុំបង្កើត Tab(គណនីគ្រូប្រចាំក្រុងស្រុក) ថ្មី") — តម្រៀបតាមលេខរៀងក្រុងស្រុកផ្លូវការ
// (01-10, ដូច DISTRICT_LIST ដែលរៀបតាមលំដាប់ត្រឹមត្រូវរួចហើយ — មើល districtOnlyComparator_ ក្នុង Utils.gs) ====================
function getDistrictTeacherUsersList(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  var data = getUsersSheetDisplayDataOnce_().data;
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][4];
    if (!isDistrictRole_(role)) continue;
    var expiryDate = normalizeDateStr_(data[i][11] || "");
    list.push({
      id: data[i][0], username: data[i][1], fullName: data[i][3],
      district: role, status: data[i][6], createdDate: data[i][7],
      gmail: data[i][8] || "", keyNumber: data[i][9] || "",
      expiryDate: expiryDate, expired: isAccountExpired_(role, expiryDate),
      // ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ មើល touchPresence_()/isUserOnline_() ខាងលើ ====
      online: isUserOnline_(data[i][1])
    });
  }
  list.sort(districtOnlyComparator_);
  return { success: true, users: list, districts: DISTRICT_LIST, districtType: DISTRICT_TYPE };
}

// ==================== ៥.០.២ គណនីភាគីពាក់ព័ន្ធ (Stakeholder Accounts — SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ, សិទ្ធិដូច
// ភ្នាក់ងារសង្កេតការណ៍ទាំងស្រុង) — ទំព័រគ្រប់គ្រងដាច់ដោយឡែក (Fix130, "សុំបង្កើត Tab(គណនីភាគីពាក់ព័ន្ធ) បន្ថែម") — រួម
// បញ្ចូល ROLE_OBSERVER (ភ្នាក់ងារសង្កេតការណ៍) + តួនាទីថ្មីទាំង៤ (មិនរាប់ "មន្ត្រី ល.ខ.ប តាកែវ" ដែលនៅ Tab(អ្នកប្រើប្រាស់)
// ដដែល — មើល OBSERVER_TIER_ROLES_.tab ក្នុង Utils.gs) — តម្រៀបតាមលំដាប់ដែលបានស្នើ៖ រដ្ឋបាលខេត្ត → រដ្ឋបាលក្រុងស្រុក
// → កងកម្លាំងទាំងបី → គណបក្សនយោបាយ → ភ្នាក់ងារសង្កេតការណ៍ (ចុងក្រោយគេ) ====================
function getStakeholderUsersList(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  var stakeholderConfigs_ = OBSERVER_TIER_ROLES_.filter(function(cfg) { return cfg.tab === 'stakeholder'; });
  var orderRoles_ = stakeholderConfigs_.map(function(cfg) { return cfg.role; }).concat([ROLE_OBSERVER]);

  var data = getUsersSheetDisplayDataOnce_().data;
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][4];
    if (orderRoles_.indexOf(role) === -1) continue;
    var expiryDate = normalizeDateStr_(data[i][11] || "");
    list.push({
      id: data[i][0], username: data[i][1], fullName: data[i][3],
      role: role, status: data[i][6], createdDate: data[i][7],
      gmail: data[i][8] || "", keyNumber: data[i][9] || "",
      expiryDate: expiryDate, expired: isAccountExpired_(role, expiryDate)
    });
  }
  list.sort(function(a, b) { return orderRoles_.indexOf(a.role) - orderRoles_.indexOf(b.role); });
  return {
    success: true, users: list,
    // ==== ROLE_OBSERVER គ្មាន Toggle បើក/បិទទេ (ដូចមុន Fix127-129 — Admin បង្កើតបានជានិច្ច) — enabled: true ជានិច្ច ====
    stakeholderRoles: stakeholderConfigs_.map(function(cfg) { return { role: cfg.role, enabled: isObserverTierRoleEnabled_(cfg.role) }; }).concat([{ role: ROLE_OBSERVER, enabled: true }])
  };
}

// ==================== ៥.០ គណនីឃុំសង្កាត់ (Commune Accounts — SuperAdmin/Admin/PEC21 គ្រប់ស្រុក, គ្រូប្រចាំក្រុងស្រុក
// ត្រឹមតែស្រុកខ្លួន) — ទំព័រគ្រប់គ្រងដាច់ដោយឡែក (មិនលាយច្របល់ជាមួយ Tab(គណនីអ្នកប្រើប្រាស់) ខាងលើទេ) ====================
function getCommuneUsersList(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var callerRole = getUserRole_(currentUsername);
  var isCallerAdmin = isAdmin_(currentUsername);
  var isCallerDistrict = isDistrictRole_(callerRole);
  if (!isCallerAdmin && !isCallerDistrict) {
    return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  }
  // ==== FIX (Fix120)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_() ដូចគ្នានឹង getUsersList()
  // ខាងលើ — មើលពន្យល់លម្អិតនៅ getUserRow_() ====
  var data = getUsersSheetDisplayDataOnce_().data;
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var role = data[i][4];
    if (!isCommuneRole_(role)) continue;
    var district = data[i][5];
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គ្រូប្រចាំក្រុងស្រុក មើលបានតែគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនប៉ុណ្ណោះ —
    // Server-side Scoping (Defense in depth, មិនទុកចិត្តតែ Client Filter ទេ) ====
    if (!isCallerAdmin && district !== callerRole) continue;
    var expiryDate = normalizeDateStr_(data[i][11] || "");
    var communeName_ = data[i][12] || "";
    list.push({
      id: data[i][0], username: data[i][1], fullName: data[i][3],
      district: district, commune: communeName_, status: data[i][6], createdDate: data[i][7],
      gmail: data[i][8] || "", keyNumber: data[i][9] || "",
      // ==== FIX144 (សំណើថ្មី "តារាងគណនីឃុំសង្កាត់ សូមមានកូដឃុំសង្កាត់ជួរឈរទី១")៖ ប្រើ communeCode_() ដែលមានស្រាប់
      // (Utils.gs — ដូចគ្នានឹងកន្លែងផ្សេងទៀតប្រើសម្រាប់បង្ហាញកូដ ឧ. 21-001) ====
      communeCode: communeCode_(district, communeName_),
      expiryDate: expiryDate, expired: isAccountExpired_(role, expiryDate),
      // ==== FIX (សំណើថ្មី "សញ្ញាបង្ហាញគណនីកំពុង Online" — ចំណុចទី៣)៖ មើល touchPresence_()/isUserOnline_() ខាងលើ ====
      online: isUserOnline_(data[i][1])
    });
  }
  // ==== FIX (Fix130, "រៀបចំបញ្ជីតាមលំដាប់កូដឃុំសង្កាត់ពី 001-100")៖ ប្រើ Comparator ដែលមានស្រាប់ (ប្រើដូចគ្នានឹង
  // getLinkedData() ស្រាប់) ====
  list.sort(districtCommuneComparator_);
  return {
    success: true, users: list,
    // ==== គ្រូប្រចាំក្រុងស្រុក ជ្រើសរើសបានតែស្រុកខ្លួន (ចាក់សោ Dropdown ត្រង់ Client) — Admin-tier ជ្រើសបានគ្រប់ស្រុក ====
    districts: isCallerAdmin ? DISTRICT_LIST : [callerRole],
    communeOrder: COMMUNE_ORDER, districtType: DISTRICT_TYPE, canPickDistrict: isCallerAdmin,
    // ==== FIX (Fix127, "ជម្រើសបើក/បិទប្រើប្រាស់គណនីឃុំសង្កាត់")៖ ត្រូវការសម្រាប់លាក់/បិទទម្រង់ "បង្កើតគណនីឃុំសង្កាត់
    // ថ្មី" (ត្រឹមតែផ្នែកបង្កើតថ្មីប៉ុណ្ណោះ — តារាងគណនីដែលមានស្រាប់ខាងក្រោម នៅតែគ្រប់គ្រង/មើលបានធម្មតា) នៅ Tab
    // "គណនីឃុំសង្កាត់" បើ Admin បិទមុខងារនេះ ====
    enableCommuneAccounts: isCommuneAccountsEnabled_()
  };
}

function addUserAccount(currentUsername, newUser, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!newUser || !newUser.username || !newUser.password || !newUser.fullName || !newUser.role) {
    return { success: false, message: "សូមបំពេញព័ត៌មានឲ្យបានគ្រប់គ្រាន់!" };
  }
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គ្រូប្រចាំក្រុងស្រុក ឥឡូវក៏អាចបង្កើតគណនីបានដែរ ប៉ុន្តែកំណត់ត្រឹមតែគណនីកម្រិត
  // ឃុំសង្កាត់ (ROLE_COMMUNE) សម្រាប់ស្រុកខ្លួនប៉ុណ្ណោះ (មិនអាចបង្កើតគណនីប្រភេទផ្សេង ឬសម្រាប់ស្រុកដទៃបានឡើយ) —
  // SuperAdmin/Admin/PEC21 (isAdmin_) នៅតែបង្កើតបានគ្រប់ប្រភេទ គ្រប់ស្រុក ដូចមុន ====
  var callerRole = getUserRole_(currentUsername);
  var callerIsDistrict = isDistrictRole_(callerRole);
  var callerCanCreate = isAdmin_(currentUsername) ||
    (callerIsDistrict && newUser.role === ROLE_COMMUNE && String(newUser.district || "").trim() === callerRole);
  if (!callerCanCreate) return { success: false, message: "គ្មានសិទ្ធិបង្កើតគណនីទេ!" };
  if (newUser.role === ROLE_SUPERADMIN && !isSuperAdmin_(currentUsername)) {
    return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាចបង្កើតគណនី SuperAdmin បាន!" };
  }
  // ==== Gmail ក្លាយជាជម្រើសសម្រាប់គណនីគ្រូប្រចាំស្រុក/ឃុំសង្កាត់ (មិនតម្រូវឲ្យបំពេញទៀតទេ) ====
  // បើទុកទទេ៖ គណនីនោះប្រើបានតែតាមរយៈ Form ក្នុង App ប៉ុណ្ណោះ គ្មានសិទ្ធិចូល Google Sheet ដោយផ្ទាល់ទេ
  var district = "", commune = "";
  if (isDistrictRole_(newUser.role)) {
    district = newUser.role;
  } else if (newUser.role === ROLE_COMMUNE) {
    // ==== FIX (Fix127, "ជម្រើសបើក/បិទប្រើប្រាស់គណនីឃុំសង្កាត់")៖ បើ SuperAdmin/Admin បិទមុខងារនេះ (ការកំណត់ប្រព័ន្ធ)
    // មិនអនុញ្ញាតឲ្យបង្កើតគណនីឃុំសង្កាត់ថ្មីទៀតទេ — អនុវត្តដូចគ្នាទាំងអ្នកបង្កើតជា SuperAdmin/Admin/PEC21 (isAdmin_)
    // និងគ្រូប្រចាំក្រុងស្រុក (មិនមានករណីលើកលែងទេ ដើម្បីឲ្យជម្រើសនេះជាការបិទ/បើកទាំងស្រុងសម្រាប់ខេត្ត/ស្រុកនោះ — បើ
    // SuperAdmin ចង់បង្កើតមួយពិសេស អាចបើកមុខងារនេះជាបណ្តោះអាសន្នវិញបាន) — ចំណាំ៖ គណនីឃុំសង្កាត់ដែលមានស្រាប់ មិនរង
    // ផលប៉ះពាល់ទេ (isCommuneAccountsEnabled_ ត្រួតពិនិត្យតែផ្លូវ "បង្កើត" ថ្មីប៉ុណ្ណោះ) ====
    if (!isCommuneAccountsEnabled_()) {
      return { success: false, message: "ប្រព័ន្ធមិនទាន់បើកអនុញ្ញាតឲ្យប្រើប្រាស់គណនីឃុំសង្កាត់ទេ! សូមចូល \"ការកំណត់ប្រព័ន្ធ\" ដើម្បីបើកមុខងារនេះជាមុនសិន (SuperAdmin/Admin ប៉ុណ្ណោះ)។" };
    }
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ត្រូវការស្រុក+ឃុំសង្កាត់ជាក់លាក់ទាំង២ ព្រោះឈ្មោះឃុំសង្កាត់មិនតែមួយគត់
    // ទូទាំងខេត្ត (មិនអាចប្រើ district=role ដូចគណនីស្រុកបានទេ) — Status = សកម្មភ្លាមៗ (បង្កើតដោយផ្ទាល់ ជឿទុកចិត្តលើ
    // អ្នកបង្កើត មិនចាំបាច់អនុម័តទៀត — ខុសពីលំហូរស្នើសុំខ្លួនឯង requestSignup ដែលនៅតែត្រូវការអនុម័ត) ====
    district = String(newUser.district || "").trim();
    commune = String(newUser.commune || "").trim();
    if (!isValidCommune_(district, commune)) {
      return { success: false, message: "សូមជ្រើសរើសស្រុក/ក្រុង និងឃុំ/សង្កាត់ឲ្យត្រឹមត្រូវ!" };
    }
    if (callerIsDistrict && district !== callerRole) {
      return { success: false, message: "អ្នកអាចបង្កើតគណនីឃុំសង្កាត់បានតែសម្រាប់ស្រុកខ្លួនប៉ុណ្ណោះ!" };
    }
  } else if (isObserverTierRole_(newUser.role)) {
    // ==== FIX (Fix129, "គណនីថ្មីទាំង៥ ត្រូវមាន Option បើក/បិទ ការបង្កើតគណនីនីមួយៗ")៖ ដូចគ្នានឹងគណនីឃុំសង្កាត់ខាងលើ —
    // ត្រូវត្រួតពិនិត្យមុនបង្កើត (callerCanCreate ខាងលើ ធានារួចហើយថាមានតែ SuperAdmin/Admin/PEC21 ទេ ដែលអាចមកដល់
    // ត្រង់នេះបាន ព្រោះគ្រូប្រចាំក្រុងស្រុកអនុញ្ញាតឲ្យបង្កើតបានតែ ROLE_COMMUNE ប៉ុណ្ណោះ) — គ្មានស្រុក/ឃុំសង្កាត់ត្រូវការទេ ====
    if (!isObserverTierRoleEnabled_(newUser.role)) {
      var __tierCfgAdd_ = getObserverTierRoleConfigByRole_(newUser.role);
      var __tierLabelAdd_ = __tierCfgAdd_ ? __tierCfgAdd_.label : newUser.role;
      return { success: false, message: "ប្រព័ន្ធមិនទាន់បើកអនុញ្ញាតឲ្យបង្កើតគណនី" + __tierLabelAdd_ + "ទេ! សូមចូល \"ការកំណត់ប្រព័ន្ធ\" ដើម្បីបើកមុខងារនេះជាមុនសិន (SuperAdmin/Admin ប៉ុណ្ណោះ)។" };
    }
  }
  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  // ==== FIX (ការពារការប៉ះទង្គិចគ្នា — Race Condition)៖ ចាក់សោ ដើម្បីកុំឲ្យ Admin ២នាក់បង្កើត Username ដូចគ្នាក្នុងពេលតែមួយ ====
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(15000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានការបង្កើតគណនីផ្សេងទៀតកំពុងដំណើរការ)! សូមព្យាយាមម្តងទៀត។" };
  }
  var keyNumber;
  try {
    var data = sheet.getDataRange().getDisplayValues();
    var inputUser = String(newUser.username).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === inputUser) {
        return { success: false, message: "ឈ្មោះគណនីនេះមានអ្នកប្រើរួចហើយ!" };
      }
    }
    // ==== FIX (សំណើថ្មី "កំណត់ចំនួនកំរិតគណនី")៖ ដូចគ្នានឹង requestSignup ខាងលើ — មើលការពន្យល់ពេញលេញនៅ
    // countDistrictRoleAccounts_()/countCommuneRoleAccountsForCommune_() (ខាងលើគេ) — អនុវត្តតែចំពោះគណនីគ្រូប្រចាំ
    // ក្រុងស្រុក/ឃុំសង្កាត់ប៉ុណ្ណោះ (គណនីតួនាទីថ្មីទាំង៥/សង្កេតការណ៍/SuperAdmin/Admin/PEC21 មិនកំណត់ចំនួន) ====
    // ==== FIX (សំណើថ្មី "លើកលែងសម្រាប់ SuperAdmin")៖ Admin/PEC21 ធម្មតា នៅតែត្រូវបានទប់ស្កាត់ត្រឹម
    // DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ ជាដាច់ខាត (ត្រូវទាក់ទង SuperAdmin) ប៉ុន្តែ SuperAdmin ខ្លួនឯង (isSuperAdmin_)
    // អាចបង្កើតគណនីគ្រូប្រចាំក្រុងស្រុកបន្ថែមលើសពីចំនួនកំណត់នេះដោយផ្ទាល់បាន (ករណីលើកលែងពិសេស) — មិនអនុវត្តចំពោះគណនី
    // ឃុំសង្កាត់ទេ (ខាងក្រោម នៅតែជាដាច់ខាតសម្រាប់អ្នកប្រើគ្រប់កម្រិត រួមទាំង SuperAdmin) ====
    if (isDistrictRole_(newUser.role)) {
      if (!isSuperAdmin_(currentUsername) && countDistrictRoleAccounts_(data, district) >= DISTRICT_ACCOUNT_MAX_PER_DISTRICT_) {
        return { success: false, message: "ស្រុក/ក្រុង \"" + district + "\" មានគណនីគ្រូប្រចាំក្រុងស្រុកគ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ + " គណនីក្នុងមួយស្រុក)! សូមទាក់ទង SuperAdmin ដើម្បីស្នើសុំបង្កើតបន្ថែម។" };
      }
    } else if (newUser.role === ROLE_COMMUNE) {
      if (countCommuneRoleAccountsForCommune_(data, district, commune) >= COMMUNE_ACCOUNT_MAX_PER_COMMUNE_) {
        return { success: false, message: "ឃុំ/សង្កាត់ \"" + commune + "\" មានគណនីឃុំសង្កាត់គ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ + " គណនីក្នុងមួយឃុំសង្កាត់)!" };
      }
    }
    keyNumber = generateKeyNumber_();
    // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ អនុវត្តតម្លៃថ្ងៃផុតកំណត់ តែចំពោះ Role ដែលអនុញ្ញាតប៉ុណ្ណោះ
    // (សង្កេតការណ៍/គ្រូប្រចាំក្រុងស្រុក) — ចោលចេញ បើផ្ញើមកសម្រាប់ SuperAdmin/Admin/PEC21 (ជៀសវាងទិន្នន័យច្របូកច្របល់) ====
    var expiryDate = isExpiryApplicableRole_(newUser.role) ? normalizeExpiryInput_(newUser.expiryDate) : "";
    sheet.appendRow([newId_(), newUser.username.trim(), makePasswordHash_(newUser.password), sanitizeForSheetCell_(newUser.fullName.trim()), newUser.role, district, STATUS_ACTIVE, formatNow_(), String(newUser.gmail || "").trim(), keyNumber, "", expiryDate, commune]);
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
  // ==== FIX (ល្បឿន)៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជាផ្ទៃខាងក្រោយ (មិនចាំបាច់ឲ្យ Admin រង់ចាំ — មើល Utils.gs) ====
  scheduleSyncSheetPermissions_();
  notifyTelegram_("✅ គណនីថ្មីត្រូវបានបង្កើត៖\n👤 " + newUser.fullName.trim() + " (" + newUser.role + ")\n🔑 Username: " + newUser.username.trim() + "\n🔢 Key Number: " + keyNumber);
  return { success: true, message: "បង្កើតគណនីជោគជ័យ! Key Number: " + keyNumber };
}

// ==================== ៥.០.១ បង្កើតគណនីឃុំសង្កាត់ដោយស្វ័យប្រវត្តិទាំងអស់ម្តង (សំណើថ្មី) ====================
// SuperAdmin/Admin/PEC21 ចុចប៊ូតុងតែម្តង ដើម្បីបង្កើតគណនីឃុំសង្កាត់ ១ គណនីសម្រាប់ ១០០ឃុំសង្កាត់ទាំងអស់ភ្លាមៗ
// (មិនចាំបាច់វាយបញ្ចូលដោយដៃម្នាក់ៗ ១០០ដងទៀតទេ)៖
//  ១. ឈ្មោះគណនី៖ com21XXXt001 (XXX = លេខកូដឃុំសង្កាត់ ៣ខ្ទង់ ស្រង់ចេញពី COMMUNE_ORDER[district][i].code
//     ដែលមានទម្រង់ "21-XXX" — ឧ. កូដ "21-001" → ឈ្មោះគណនី "com21001t001", កូដ "21-002" → "com21002t001")
//  ២. លេខសម្ងាត់លំនាំដើម៖ AUTO_COMMUNE_DEFAULT_PASSWORD_ ("pec21") ដូចគ្នាទាំង១០០គណនី (បណ្តោះអាសន្ន)
//  ៣. គោត្តនាម-នាម លំនាំដើម៖ ឈ្មោះឃុំសង្កាត់ខ្លួនឯង (ឧ. "អង្គរបូរី") — អាចប្តូរជាឈ្មោះផ្ទាល់ខ្លួនពេល Login លើកដំបូង
//  ៤. Status = សកម្មភ្លាមៗ (ជឿទុកចិត្តលើ Admin ដូចគ្នានឹង addUserAccount() ធម្មតា — មិនចាំបាច់អនុម័ត/Key Number
//     ទៀតទេ) ប៉ុន្តែសម្គាល់ Flag "តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូង" (ជួរឈរ N) ដើម្បីបង្ខំឲ្យប្តូរលេខសម្ងាត់+ឈ្មោះ ជាមុន
//     សិន មុននឹងអាចប្រើប្រាស់ប្រព័ន្ធធម្មតាបាន (មើល loginUser() ខាងលើ និង changePasswordFirstLogin() ខាងក្រោម)
// អាចហៅដដែលៗដោយសុវត្ថិភាព (Idempotent — ប្រើ COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ ដដែលនឹងកំណត់ចំនួនកំរិតគណនីខាងលើ)៖
// រំលងឃុំសង្កាត់ណាដែលមានគណនីរួចហើយ — មានប្រយោជន៍បើឃុំសង្កាត់ថ្មីត្រូវបានបន្ថែមទៅ COMMUNE_ORDER ពេលក្រោយ ឬបើគណនី
// ខ្លះត្រូវបានលុបចោល ចង់បង្កើតជំនួសម្តងទៀត ====
var AUTO_COMMUNE_DEFAULT_PASSWORD_ = "pec21";
var AUTO_COMMUNE_USERNAME_PREFIX_ = "com21";
var AUTO_COMMUNE_USERNAME_SUFFIX_ = "t001";
// ==== ស្រង់លេខកូដ ៣ខ្ទង់ចេញពីទម្រង់ "21-XXX" (ដូចផ្ទុកក្នុង COMMUNE_ORDER[district][i].code) ====
function autoCommuneUsername_(communeCodeFull) {
  var codeDigits = String(communeCodeFull || '').split('-')[1] || '';
  return AUTO_COMMUNE_USERNAME_PREFIX_ + codeDigits + AUTO_COMMUNE_USERNAME_SUFFIX_;
}

// ==== FIX (រាយការណ៍ថ្មី "ចុចបង្កើតគណនីឃុំសង្កាត់ ច្រានចោល 'កំពុងបង្កើត...' រហូត អត់ដឹងថាកំពុងដំណើរការឬគាំង")៖ ដូចគ្នា
// នឹង CACHE_KEY_RESET_PROGRESS_ (Settings.gs — resetAllData) — Key/TTL នេះប្រើសម្រាប់រក្សាទុក "វឌ្ឍនភាព" (ឃុំសង្កាត់
// ណាហើយ ប៉ុន្មានហើយ) ទៅ CacheService ភ្លាមៗក្រោយឃុំសង្កាត់នីមួយៗត្រូវបានត្រួតពិនិត្យ/បង្កើតចប់ — Client ស្ទង់មើល (Poll)
// តាមរយៈ getBulkCreateCommuneAccountsProgress() ជាទៀងទាត់ ខណៈកំពុងរង់ចាំ Call ចម្បង (bulkCreateAllCommuneAccounts_)
// បញ្ចប់ ដើម្បីបង្ហាញវឌ្ឍនភាពជាក់ស្តែង ជំនួសការស្ងាត់ស្ងៀមទាំងស្រុង (ដែលធ្វើឲ្យអ្នកប្រើស្មានថាគាំង) ====
var CACHE_KEY_COMMUNE_CREATE_PROGRESS_ = 'bulkCreateCommuneProgress_v1';
var CACHE_TTL_COMMUNE_CREATE_PROGRESS_ = 600; // វិនាទី (១០នាទី — លើសពេលវេលាដំណើរការជាក់ស្តែងច្រើន សម្រាប់ជាកម្រិតសុវត្ថិភាព)

// ==== FIX (រាយការណ៍ថ្មី "ឃើញ 'ឃុំសង្កាត់ទាំង 0 ចប់ហើយ!' ភ្លាមៗ ទាំងៗដែលមិនទាន់ចាប់ផ្តើម")៖ ចំនួនសរុបលំនាំដើម (Fallback —
// មុន Server ចាប់ផ្តើមសរសេរ "វឌ្ឍនភាព" ទៅ Cache ផ្ទាល់ខ្លួន) ត្រូវប្រើចំនួនឃុំសង្កាត់ពិតប្រាកដ (គណនាពី COMMUNE_ORDER)
// មិនមែន 0 ទេ — ដូចគ្នានឹង getResetAllDataProgress() ប្រើ DISTRICT_LIST.length ជា Fallback (មិនមែន 0) ដែរ។ បើ Fallback
// ជា 0 នោះ Client (done>=total ⇒ "finished") នឹងឃើញ 0>=0 ជាការពិតភ្លាមៗ ពេល Poll លើកទី១ មុននឹង Server ទាន់សរសេរ
// Cache ដំបូង (Race Condition) ធ្វើឲ្យបង្ហាញសារ "ចប់ហើយ!" ខុសឆ្គងទាំងស្រុង ====
function totalCommuneCount_() {
  var total = 0;
  DISTRICT_LIST.forEach(function(d) { total += (COMMUNE_ORDER[d] || []).length; });
  return total;
}

function getBulkCreateCommuneAccountsProgress(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  try {
    var raw = CacheService.getScriptCache().get(CACHE_KEY_COMMUNE_CREATE_PROGRESS_);
    if (!raw) return { success: true, done: 0, total: totalCommuneCount_(), commune: '', finished: false };
    var p = JSON.parse(raw);
    return { success: true, done: p.done || 0, total: p.total || totalCommuneCount_(), commune: p.commune || '', finished: !!p.finished };
  } catch (err) {
    return { success: true, done: 0, total: totalCommuneCount_(), commune: '', finished: false };
  }
}

// ==== FIX (fix183, "ចុចបង្កើតគណនីឃុំសង្កាត់មិនដំណើរការសោះ")៖ ត្រូវប្តូរឈ្មោះមុខងារនេះ ដកសញ្ញា "_" ចេញពីចុង
// ព្រោះ google.script.run របស់ Google Apps Script មិនអាចហៅមុខងារដែលឈ្មោះបញ្ចប់ដោយសញ្ញា "_" បានទេ (ចាត់ទុកជា
// "Private" ដោយស្វ័យប្រវត្តិ ហើយមិនបង្ហាញនៅក្នុង Client API Stub ទាល់តែសោះ) — នេះជាមូលហេតុពិតប្រាកដដែលបណ្តាលឲ្យ
// ប៊ូតុង "បង្កើតគណនីឃុំសង្កាត់" ជាប់គាំងរហូត ដោយសារការហៅពី Client (Index.html) មិនដែលទៅដល់ Server សោះ (ប៉ុន្តែ
// getBulkCreateCommuneAccountsProgress ដែលមិនមានសញ្ញា "_" ខាងចុង នៅតែហៅបានធម្មតា ទើបធ្វើឲ្យមើលទៅហាក់ដូចជា
// "កំពុង Poll ដំណើរការ ប៉ុន្តែសកម្មភាពពិតមិនដំណើរការ") ====
function bulkCreateAllCommuneAccounts(currentUsername, sessionToken, expiryDate) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) {
    return { success: false, message: "គ្មានសិទ្ធិបង្កើតគណនីឃុំសង្កាត់ស្វ័យប្រវត្តិទេ! (SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ)" };
  }
  if (!isCommuneAccountsEnabled_()) {
    return { success: false, message: "ប្រព័ន្ធមិនទាន់បើកអនុញ្ញាតឲ្យប្រើប្រាស់គណនីឃុំសង្កាត់ទេ! សូមចូល \"ការកំណត់ប្រព័ន្ធ\" ដើម្បីបើកមុខងារនេះជាមុនសិន។" };
  }
  // ==== FIX (សំណើថ្មី "ថ្ងៃផុតកំណត់មុនបង្កើត")៖ ធ្វើឲ្យស្តង់ដារតម្លៃថ្ងៃផុតកំណត់ (ស្រេចចិត្ត) តាមរបៀបដូចគ្នានឹង
  // addUserAccount ប្រើ — តម្លៃមិនត្រឹមត្រូវ ឬទទេ នឹងក្លាយជា "" (មិនកំណត់ថ្ងៃផុតកំណត់) ដោយស្វ័យប្រវត្តិ ====
  var normalizedExpiry_ = normalizeExpiryInput_(expiryDate);

  // ==== FIX (រាយការណ៍ថ្មី "ច្រានចោល 'កំពុងបង្កើត...' រហូត")៖ ចុះបញ្ជីឃុំសង្កាត់ទាំងអស់ (ស្រុក+ឃុំសង្កាត់) ជាបញ្ជីរាបស្មើ
  // តែមួយជាមុនសិន ដើម្បីឲ្យមានលេខរៀង (index) ច្បាស់លាស់ សម្រាប់សរសេរ "វឌ្ឍនភាព" ទៅ CacheService ក្រោយឃុំសង្កាត់
  // នីមួយៗចប់ — ដូចគ្នានឹង districtIdx_ ក្នុង resetAllData() (Settings.gs) ====
  var allCommuneEntries_ = [];
  DISTRICT_LIST.forEach(function(district) {
    (COMMUNE_ORDER[district] || []).forEach(function(entry) {
      allCommuneEntries_.push({ district: district, entry: entry });
    });
  });
  var totalCommunes_ = allCommuneEntries_.length;
  try { CacheService.getScriptCache().put(CACHE_KEY_COMMUNE_CREATE_PROGRESS_, JSON.stringify({ done: 0, total: totalCommunes_, commune: '', finished: false }), CACHE_TTL_COMMUNE_CREATE_PROGRESS_); } catch (eProg0) {}

  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  // ==== ចាក់សោ (ដូចគ្នានឹង requestSignup/addUserAccount ខាងលើ) — ការពារកុំឲ្យប្តូរម្តងទៀត (ឧ. ចុចប៊ូតុងជាន់គ្នា)
  // បង្កើតគណនីស្ទួនគ្នាសម្រាប់ឃុំសង្កាត់ដដែល — Timeout វែងជាងធម្មតា (២០វិនាទី) ព្រោះការហៅនេះអាចសរសេរដល់១០០ជួរដេក ====
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(20000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់! សូមព្យាយាមម្តងទៀត។" };
  }
  var createdCount = 0, skippedCount = 0, newRows = [];
  try {
    var data = sheet.getDataRange().getDisplayValues();
    var existingUsernames_ = {};
    for (var i = 1; i < data.length; i++) { existingUsernames_[String(data[i][1]).trim().toLowerCase()] = true; }
    allCommuneEntries_.forEach(function(item, idx_) {
      var district = item.district, entry = item.entry;
      if (countCommuneRoleAccountsForCommune_(data, district, entry.name) >= COMMUNE_ACCOUNT_MAX_PER_COMMUNE_) {
        skippedCount++;
      } else {
        var username = autoCommuneUsername_(entry.code);
        if (existingUsernames_[username.toLowerCase()]) {
          // ==== ករណីកម្រណាស់ (ឈ្មោះគណនីស្ទួន — ឧ. នរណាម្នាក់ធ្លាប់បង្កើតដោយដៃដោយប្រើទម្រង់ឈ្មោះដូចគ្នា) — រំលង
          // ជំនួសការបោះ Error ធ្វើឲ្យទាំងមូលគាំង (ឃុំសង្កាត់ដែលនៅសល់នៅតែបន្តបង្កើតធម្មតា) ====
          skippedCount++;
        } else {
          newRows.push([
            newId_(), username, makePasswordHash_(AUTO_COMMUNE_DEFAULT_PASSWORD_), sanitizeForSheetCell_(entry.name),
            ROLE_COMMUNE, district, STATUS_ACTIVE, formatNow_(), "", generateKeyNumber_(), "", normalizedExpiry_, entry.name, "1"
          ]);
          existingUsernames_[username.toLowerCase()] = true; // ការពារស្ទួនគ្នាខ្លួនឯងក្នុងវដ្តតែមួយនេះ
          createdCount++;
        }
      }
      // ==== FIX (រាយការណ៍ថ្មី)៖ ធ្វើបច្ចុប្បន្នភាព "វឌ្ឍនភាព" ភ្លាមៗក្រោយឃុំសង្កាត់នីមួយៗចប់ (មិនថាបង្កើត ឬរំលង) ដើម្បីឲ្យ
      // Client ស្ទង់មើល (Poll — getBulkCreateCommuneAccountsProgress) ឃើញដំណើរការជាក់ស្តែង ====
      try {
        CacheService.getScriptCache().put(CACHE_KEY_COMMUNE_CREATE_PROGRESS_, JSON.stringify({
          done: idx_ + 1, total: totalCommunes_, commune: entry.name, district: district, finished: false
        }), CACHE_TTL_COMMUNE_CREATE_PROGRESS_);
      } catch (eProg) {}
    });
    if (newRows.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      formatExpiryColumnBounded_(sheet);
    }
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
  if (createdCount) {
    scheduleSyncSheetPermissions_();
    notifyTelegram_("✅ បានបង្កើតគណនីឃុំសង្កាត់ស្វ័យប្រវត្តិ ចំនួន " + createdCount + " គណនី (ចាប់ពី com21001t001)\n🔑 លេខសម្ងាត់លំនាំដើម៖ " + AUTO_COMMUNE_DEFAULT_PASSWORD_ + " (នឹងតម្រូវឲ្យប្តូរនៅ Login លើកដំបូង)" + (normalizedExpiry_ ? ("\n📅 ថ្ងៃផុតកំណត់៖ " + normalizedExpiry_ + " (គណនីទាំងអស់)") : "") + (skippedCount ? ("\n⏭️ រំលង " + skippedCount + " ឃុំសង្កាត់ ព្រោះមានគណនីរួចហើយ") : ""));
  }
  try { CacheService.getScriptCache().put(CACHE_KEY_COMMUNE_CREATE_PROGRESS_, JSON.stringify({ done: totalCommunes_, total: totalCommunes_, commune: '', finished: true }), CACHE_TTL_COMMUNE_CREATE_PROGRESS_); } catch (eProgFin) {}
  return {
    success: true, created: createdCount, skipped: skippedCount,
    message: createdCount
      ? ("បង្កើតគណនីឃុំសង្កាត់ស្វ័យប្រវត្តិជោគជ័យ ចំនួន " + createdCount + " គណនី!" + (skippedCount ? (" (រំលង " + skippedCount + " ឃុំសង្កាត់ ព្រោះមានគណនីរួចហើយ)") : ""))
      : "គ្មានគណនីត្រូវបង្កើតថ្មីទេ — ឃុំសង្កាត់ទាំងអស់មានគណនីរួចហើយ!"
  };
}

// ==================== ៥.០.១.១ កំណត់គណនីឃុំសង្កាត់ជាលំនាំដើមវិញ (សម្រាប់ក្រោយសាកល្បង — Reset to Default) ====================
// ==== FIX (សំណើថ្មី "កំណត់ជាលំនាំដើមវិញ")៖ បើ Admin/គ្រូប្រចាំក្រុងស្រុកយកគណនីឃុំសង្កាត់ណាមួយទៅសាកល្បង (ដែលឆ្លងកាត់
// ការតម្រូវប្តូរលេខសម្ងាត់+ឈ្មោះលើកដំបូងរួច) មុខងារនេះស្តារគណនីនោះត្រឡប់ទៅដូចដើមបេះបិទ ដូចគ្នានឹងស្ថានភាពដែលបានពី
// bulkCreateAllCommuneAccounts_() លើកដំបូង៖ លេខសម្ងាត់ត្រឡប់ទៅ AUTO_COMMUNE_DEFAULT_PASSWORD_ ("pec21"), ឈ្មោះ
// គោត្តនាម-នាមត្រឡប់ទៅឈ្មោះឃុំសង្កាត់ខ្លួនឯង, តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូងម្តងទៀត (Column 14 = "1"), ហើយកាត់ Session
// ដែលកំពុងប្រើប្រាស់ស្រាប់ចោល (បើមាន — សម្អាត Column 11 ទៅទទេ ដើម្បីជៀសវាងអ្នកសាកល្បងនៅតែជាប់ Login ក្នុងស្ថានភាព
// ចាស់ក្រោយ Reset រួច) ។ ធ្វើការតែលើគណនី ROLE_COMMUNE ប៉ុណ្ណោះ (មិនអនុញ្ញាតលើគណនីប្រភេទផ្សេង ដើម្បីជៀសវាងការប្រើខុស) ====
// ==== FIX (fix184, "ប្តូរលេខសម្ងាត់ហើយ តែនៅតែ Login ចូលបានដោយលេខសម្ងាត់ចាស់ ក្រោយចុច (លំនាំដើម)")៖ ដូចគ្នានឹង fix183
// បេះបិទ — ឈ្មោះមុខងារនេះក៏បញ្ចប់ដោយសញ្ញា "_" ដែរ ហើយត្រូវបានហៅតាមរយៈ runUserAccountAction_() ដែលហៅ Server ជា
// Dynamic (google.script.run[serverFnName](...)) ដូច្នេះការហៅនេះក៏មិនដែលទៅដល់ Server ដែរ — ចុច "លំនាំដើម" មិនដែល
// កំណត់អ្វីទាំងអស់ សោះ (គ្មាន Error បង្ហាញផង ព្រោះ withSuccessHandler/withFailureHandler ក៏មិនដែនត្រូវហៅដែរ) ====
function resetCommuneUserToDefault(currentUsername, userId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិកំណត់គណនីនេះជាលំនាំដើមវិញទេ!" };
  }
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (data[i][4] !== ROLE_COMMUNE) {
        return { success: false, message: "សកម្មភាពនេះអនុវត្តបានតែលើគណនីឃុំសង្កាត់ប៉ុណ្ណោះ!" };
      }
      if (!canManageTargetAccount_(currentUsername, data[i][4], data[i][5])) {
        return { success: false, message: "គ្មានសិទ្ធិកំណត់គណនីនេះជាលំនាំដើមវិញទេ!" };
      }
      var communeName = String(data[i][12] || "").trim();
      if (!communeName) {
        return { success: false, message: "រកមិនឃើញឈ្មោះឃុំសង្កាត់សម្រាប់គណនីនេះទេ! មិនអាចកំណត់ជាលំនាំដើមវិញបានទេ។" };
      }
      sheet.getRange(i + 1, 3).setValue(makePasswordHash_(AUTO_COMMUNE_DEFAULT_PASSWORD_));
      sheet.getRange(i + 1, 4).setValue(sanitizeForSheetCell_(communeName));
      sheet.getRange(i + 1, 11).setValue(""); // សម្អាត Session Token ចាស់ (បង្ខំ Logout បើកំពុងប្រើប្រាស់ស្រាប់)
      sheet.getRange(i + 1, 14).setValue("1"); // តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូងម្តងទៀត
      scheduleSyncSheetPermissions_();
      return { success: true, message: "បានកំណត់គណនីនេះជាលំនាំដើមវិញជោគជ័យ! លេខសម្ងាត់៖ \"" + AUTO_COMMUNE_DEFAULT_PASSWORD_ + "\" (នឹងតម្រូវឲ្យប្តូរនៅ Login លើកដំបូង)" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// ==================== ៥.០.១.២ បង្កើតគណនីគ្រូប្រចាំក្រុងស្រុកដោយស្វ័យប្រវត្តិទាំងអស់ម្តង (សំណើថ្មី — ត្រាប់តាម
// លំនាំគណនីឃុំសង្កាត់ខាងលើ ៥.០.១ ទាំងស្រុង, ប៉ុន្តែប្រើ Prefix ឈ្មោះគណនីខុសគ្នា ដើម្បីកុំឲ្យប៉ះទង្គិច/ច្រឡំគ្នា)
// ==== FIX (សំណើកែប្រែថ្មី "២០គណនី ២ក្នុងមួយក្រុងស្រុក, អក្សរធំ")៖ ត្រូវការគណនី ២ ក្នុងមួយក្រុងស្រុក (សរុប ២០) —
// បន្ថែម Suffix ទី២ "R001" (ក្រៅពី "T001" ដើម) ព្រម​ទាំងប្តូរឈ្មោះគណនីទាំងអស់ជាអក្សរធំ (Uppercase) តាមសំណើ ====
// SuperAdmin/Admin/PEC21 ចុចប៊ូតុងតែម្តង ដើម្បីបង្កើតគណនីគ្រូប្រចាំក្រុងស្រុក ២ គណនីម្នាក់ៗ សម្រាប់ក្រុងស្រុកទាំង១០
// (សរុប ២០ គណនី) ភ្លាមៗ៖
//  ១. ឈ្មោះគណនី៖ PEC21XXT001 (គណនីទី១) និង PEC21XXR001 (គណនីទី២) — XX = លេខរៀងក្រុងស្រុក ២ខ្ទង់ តាមលំដាប់
//     DISTRICT_LIST ឧ. អង្គរបូរី(លេខរៀង១)→ PEC2101T001/PEC2101R001, បាទី(លេខរៀង២)→ PEC2102T001/PEC2102R001
//     ... ទ្រាំង(លេខរៀង១០)→ PEC2110T001/PEC2110R001 — ប្រើ Prefix "PEC21" ខុសពី Prefix ឃុំសង្កាត់ "com21"
//     ដោយចេតនា (សំណើច្បាស់លាស់ពីអ្នកប្រើ "កុំខុសកូដដូចឃុំសង្កាត់") — ឈ្មោះគណនីទាំងអស់សរសេរជាអក្សរធំ (Uppercase)
//     តាមសំណើ (ចំណាំ៖ ការ Login ប្រៀបធៀបឈ្មោះគណនីមិនប្រកាន់អក្សរតូច/ធំទេ — loginUser() ប្រើ toLowerCase() —
//     ដូច្នេះវាយអក្សរតូចក៏ Login ចូលបានដដែល ភាពខុសគ្នាគឺជាទម្រង់បង្ហាញ/រក្សាទុកតែប៉ុណ្ណោះ)
//  ២. លេខសម្ងាត់លំនាំដើម៖ AUTO_DISTRICT_DEFAULT_PASSWORD_ ("pec21") ដូចគ្នាទាំង២០គណនី (បណ្តោះអាសន្ន)
//  ៣. គោត្តនាម-នាម លំនាំដើម៖ ឈ្មោះក្រុងស្រុកខ្លួនឯង (ឧ. "អង្គរបូរី") ដូចគ្នាទាំង២គណនីក្នុងក្រុងស្រុកតែមួយ — អាចប្តូរ
//     ជាឈ្មោះផ្ទាល់ខ្លួនពេល Login លើកដំបូង
//  ៤. Status = សកម្មភ្លាមៗ ប៉ុន្តែសម្គាល់ Flag "តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូង" (ជួរឈរ ១៤) ដើម្បីបង្ខំឲ្យប្តូរលេខសម្ងាត់
//     +ឈ្មោះជាមុនសិន (ប្រើ changePasswordFirstLogin() ខាងក្រោម ដដែល — Function រួមស្រាប់ហើយ មិនចាំបាច់សរសេរថ្មីទេ)
// អាចហៅដដែលៗដោយសុវត្ថិភាព (Idempotent — ប្រើ DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ = 2 ដដែលនឹងកំណត់ចំនួនកំរិតគណនី
// ខាងលើ ដែលដំណើរការត្រូវគ្នាល្អឥតខ្ចោះជាមួយ ២គណនី/ក្រុងស្រុកនេះ)៖ រំលងក្រុងស្រុកណាដែលមានគណនីគ្រប់ចំនួនកំណត់រួចហើយ
// ==== FIX (ការត្រួតពិនិត្យត្រឹមត្រូវសំខាន់)៖ ដោយសារឥឡូវអាចបង្កើត ២ គណនីសម្រាប់ក្រុងស្រុកតែមួយក្នុងការហៅតែម្តង (មិនមែន
// ១ ដូចមុនទៀតទេ) ការរាប់ចំនួនគណនីបច្ចុប្បន្នមិនអាចពឹងផ្អែកតែលើ Snapshot ស្ថិត (data) ដដែលពេញមួយ Loop បានទៀតទេ —
// ត្រូវកត់ត្រា "ចំនួនកំពុងកើនឡើង" ក្នុងអង្គចងចាំ (districtCounts_) ដើម្បីកុំឲ្យបង្កើតលើសចំនួនកំណត់ (ឧ. ក្រុងស្រុកមាន
// គណនីស្រាប់ ១ រួចហើយ ការហៅតែម្តងនេះមិនត្រូវបង្កើតបន្ថែមលើសពី ១ ទៀតទេ ដើម្បីកុំឲ្យលើសកំណត់ ២) ====
// ==== FIX (ដកបទពិសោធន៍ពី fix183/fix184)៖ ឈ្មោះមុខងារ Server ត្រូវតែគ្មានសញ្ញា "_" ខាងចុងជាដាច់ខាត — google.script.run
// មិនអាចហៅមុខងារបញ្ចប់ដោយ "_" បានទេ ====
var AUTO_DISTRICT_DEFAULT_PASSWORD_ = "pec21";
var AUTO_DISTRICT_USERNAME_PREFIX_ = "PEC21";
var AUTO_DISTRICT_ACCOUNT_SUFFIXES_ = ["T001", "R001"]; // ២ គណនីក្នុងមួយក្រុងស្រុក
// ==== ស្រង់លេខរៀងក្រុងស្រុក (0-based index ក្នុង DISTRICT_LIST) ជាទម្រង់ ២ខ្ទង់ ឧ. index 0 → "01", index 9 → "10" ====
function autoDistrictUsername_(districtIndex, suffix) {
  var num = String(districtIndex + 1);
  if (num.length < 2) num = '0' + num;
  return AUTO_DISTRICT_USERNAME_PREFIX_ + num + suffix;
}

var CACHE_KEY_DISTRICT_CREATE_PROGRESS_ = 'bulkCreateDistrictProgress_v1';
var CACHE_TTL_DISTRICT_CREATE_PROGRESS_ = 600; // វិនាទី (១០នាទី)

// ==== ចំនួន "កន្លែង" សរុប (ក្រុងស្រុក ១០ x គណនី ២ = ២០) — ប្រើទាំង Fallback (មុន Cache ដំបូង) និង Total ពិតប្រាកដ ====
function totalDistrictAutoCreateSlots_() {
  return DISTRICT_LIST.length * AUTO_DISTRICT_ACCOUNT_SUFFIXES_.length;
}

function getBulkCreateDistrictAccountsProgress(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  try {
    var raw = CacheService.getScriptCache().get(CACHE_KEY_DISTRICT_CREATE_PROGRESS_);
    if (!raw) return { success: true, done: 0, total: totalDistrictAutoCreateSlots_(), district: '', finished: false };
    var p = JSON.parse(raw);
    return { success: true, done: p.done || 0, total: p.total || totalDistrictAutoCreateSlots_(), district: p.district || '', finished: !!p.finished };
  } catch (err) {
    return { success: true, done: 0, total: totalDistrictAutoCreateSlots_(), district: '', finished: false };
  }
}

function bulkCreateAllDistrictAccounts(currentUsername, sessionToken, expiryDate) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) {
    return { success: false, message: "គ្មានសិទ្ធិបង្កើតគណនីគ្រូប្រចាំក្រុងស្រុកស្វ័យប្រវត្តិទេ! (SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ)" };
  }
  var normalizedExpiry_ = normalizeExpiryInput_(expiryDate);

  var allSlots_ = [];
  DISTRICT_LIST.forEach(function(district, districtIdx_) {
    AUTO_DISTRICT_ACCOUNT_SUFFIXES_.forEach(function(suffix) {
      allSlots_.push({ district: district, districtIdx: districtIdx_, suffix: suffix });
    });
  });
  var totalSlots_ = allSlots_.length;
  try { CacheService.getScriptCache().put(CACHE_KEY_DISTRICT_CREATE_PROGRESS_, JSON.stringify({ done: 0, total: totalSlots_, district: '', finished: false }), CACHE_TTL_DISTRICT_CREATE_PROGRESS_); } catch (eProg0) {}

  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(20000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់! សូមព្យាយាមម្តងទៀត។" };
  }
  var createdCount = 0, skippedCount = 0, newRows = [];
  try {
    var data = sheet.getDataRange().getDisplayValues();
    var existingUsernames_ = {};
    for (var i = 1; i < data.length; i++) { existingUsernames_[String(data[i][1]).trim().toLowerCase()] = true; }
    // ==== FIX (ការត្រួតពិនិត្យត្រឹមត្រូវ — មើលពន្យល់ខាងលើ)៖ ចំនួនកំពុងកើនឡើងក្នុងអង្គចងចាំ ចាប់ផ្តើមពី Snapshot ស្ថិត ====
    var districtCounts_ = {};
    DISTRICT_LIST.forEach(function(district) { districtCounts_[district] = countDistrictRoleAccounts_(data, district); });
    allSlots_.forEach(function(slot, slotIdx_) {
      var district = slot.district;
      if (districtCounts_[district] >= DISTRICT_ACCOUNT_MAX_PER_DISTRICT_) {
        skippedCount++;
      } else {
        var username = autoDistrictUsername_(slot.districtIdx, slot.suffix);
        if (existingUsernames_[username.toLowerCase()]) {
          skippedCount++;
        } else {
          newRows.push([
            newId_(), username, makePasswordHash_(AUTO_DISTRICT_DEFAULT_PASSWORD_), sanitizeForSheetCell_(district),
            district, district, STATUS_ACTIVE, formatNow_(), "", generateKeyNumber_(), "", normalizedExpiry_, "", "1"
          ]);
          existingUsernames_[username.toLowerCase()] = true;
          districtCounts_[district]++;
          createdCount++;
        }
      }
      try {
        CacheService.getScriptCache().put(CACHE_KEY_DISTRICT_CREATE_PROGRESS_, JSON.stringify({
          done: slotIdx_ + 1, total: totalSlots_, district: district + ' (' + slot.suffix + ')', finished: false
        }), CACHE_TTL_DISTRICT_CREATE_PROGRESS_);
      } catch (eProg) {}
    });
    if (newRows.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
      formatExpiryColumnBounded_(sheet);
    }
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
  if (createdCount) {
    scheduleSyncSheetPermissions_();
    notifyTelegram_("✅ បានបង្កើតគណនីគ្រូប្រចាំក្រុងស្រុកស្វ័យប្រវត្តិ ចំនួន " + createdCount + " គណនី (ចាប់ពី PEC2101T001)\n🔑 លេខសម្ងាត់លំនាំដើម៖ " + AUTO_DISTRICT_DEFAULT_PASSWORD_ + " (នឹងតម្រូវឲ្យប្តូរនៅ Login លើកដំបូង)" + (normalizedExpiry_ ? ("\n📅 ថ្ងៃផុតកំណត់៖ " + normalizedExpiry_ + " (គណនីទាំងអស់)") : "") + (skippedCount ? ("\n⏭️ រំលង " + skippedCount + " គណនី ព្រោះក្រុងស្រុកនោះមានគណនីគ្រប់ចំនួនកំណត់រួចហើយ") : ""));
  }
  try { CacheService.getScriptCache().put(CACHE_KEY_DISTRICT_CREATE_PROGRESS_, JSON.stringify({ done: totalSlots_, total: totalSlots_, district: '', finished: true }), CACHE_TTL_DISTRICT_CREATE_PROGRESS_); } catch (eProgFin) {}
  return {
    success: true, created: createdCount, skipped: skippedCount,
    message: createdCount
      ? ("បង្កើតគណនីគ្រូប្រចាំក្រុងស្រុកស្វ័យប្រវត្តិជោគជ័យ ចំនួន " + createdCount + " គណនី!" + (skippedCount ? (" (រំលង " + skippedCount + " គណនី ព្រោះក្រុងស្រុកនោះមានគណនីគ្រប់ចំនួនកំណត់រួចហើយ)") : ""))
      : "គ្មានគណនីត្រូវបង្កើតថ្មីទេ — ក្រុងស្រុកទាំងអស់មានគណនីគ្រប់ចំនួនកំណត់រួចហើយ!"
  };
}

// ==== FIX (ត្រាប់តាម ៥.០.១.១)៖ កំណត់គណនីគ្រូប្រចាំក្រុងស្រុកជាលំនាំដើមវិញ (សម្រាប់ក្រោយសាកល្បង — Reset to Default) —
// SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ (មិនអនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុកម្នាក់ Reset គណនីគ្រូប្រចាំក្រុងស្រុកដទៃទៀតឡើយ —
// ខុសពីគណនីឃុំសង្កាត់ដែលអនុញ្ញាតឲ្យ Admin ស្រុកគ្រប់គ្រងបាន ព្រោះកម្រិតគណនីនេះខ្ពស់ជាង) ====
function resetDistrictUserToDefault(currentUsername, userId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) {
    return { success: false, message: "គ្មានសិទ្ធិកំណត់គណនីនេះជាលំនាំដើមវិញទេ!" };
  }
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      var role = data[i][4];
      if (!isDistrictRole_(role)) {
        return { success: false, message: "សកម្មភាពនេះអនុវត្តបានតែលើគណនីគ្រូប្រចាំក្រុងស្រុកប៉ុណ្ណោះ!" };
      }
      sheet.getRange(i + 1, 3).setValue(makePasswordHash_(AUTO_DISTRICT_DEFAULT_PASSWORD_));
      sheet.getRange(i + 1, 4).setValue(sanitizeForSheetCell_(role));
      sheet.getRange(i + 1, 11).setValue(""); // សម្អាត Session Token ចាស់ (បង្ខំ Logout បើកំពុងប្រើប្រាស់ស្រាប់)
      sheet.getRange(i + 1, 14).setValue("1"); // តម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូងម្តងទៀត
      scheduleSyncSheetPermissions_();
      return { success: true, message: "បានកំណត់គណនីនេះជាលំនាំដើមវិញជោគជ័យ! លេខសម្ងាត់៖ \"" + AUTO_DISTRICT_DEFAULT_PASSWORD_ + "\" (នឹងតម្រូវឲ្យប្តូរនៅ Login លើកដំបូង)" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// ==================== ៥.០.២ ប្តូរលេខសម្ងាត់ដំបូង (សម្រាប់គណនីបង្កើតដោយស្វ័យប្រវត្តិ — bulkCreateAllCommuneAccounts_) ====
// Function សាធារណៈ (មិនតម្រូវ Session ទេ — ព្រោះហៅមុននឹងមាន Session ណាមួយទាល់តែសោះ ស្រដៀងគ្នានឹង confirmKeyNumber())
// ត្រូវការតែឈ្មោះគណនី+លេខសម្ងាត់បច្ចុប្បន្ន (ត្រឹមត្រូវ) ជា "សិទ្ធិ" ដើម្បីអនុញ្ញាតឲ្យប្តូរ។
// ==== FIX (សុវត្ថិភាព)៖ ដូចគ្នានឹង confirmKeyNumber() ខាងលើ — ត្រូវកំណត់ចំនួនដងសាកល្បង (Rate-limit) ព្រោះនេះជា
// Function សាធារណៈ ហើយលេខសម្ងាត់លំនាំដើមចែករំលែងគ្នាទាំង១០០គណនី (ស្គាល់ជាទូទៅចំពោះនរណាម្នាក់ដឹងទម្រង់នេះ) — ដោយ
// គ្មានការកំណត់នេះ នរណាម្នាក់អាចព្យាយាមប្តូរលេខសម្ងាត់គណនីឃុំសង្កាត់ណាមួយ (ដែលម្ចាស់ពិតប្រាកដមិនទាន់ Login ប្តូរជាមុន)
// ដើម្បីលួច "កាន់កាប់" គណនីនោះ (Account Takeover) — ចំណាំសំខាន់៖ នេះជាការកាត់បន្ថយហានិភ័យ មិនមែនលុបបំបាត់ទាំងស្រុង
// ទេ ព្រោះលេខសម្ងាត់លំនាំដើមត្រូវបានគេស្គាល់ស្រាប់ (ជាទម្រង់រួម) — Admin គួរជូនដំណឹង Username ទៅសមាជិកនីមួយៗ ហើយ
// លើកទឹកចិត្តឲ្យប្តូរភ្លាមៗបន្ទាប់ពីទទួលបាន ដើម្បីកាត់បន្ថយរយៈពេលដែលងាយរងគ្រោះ ====
var FIRST_LOGIN_CHANGE_MAX_ATTEMPTS_ = 10;
var FIRST_LOGIN_CHANGE_LOCKOUT_SECONDS_ = 3600; // ១ម៉ោង

function changePasswordFirstLogin(username, oldPassword, newPassword, newFullName) {
  var inputUser = String(username || "").trim().toLowerCase();
  if (!inputUser || !oldPassword || !newPassword || !String(newFullName || "").trim()) {
    return { success: false, message: "សូមបំពេញព័ត៌មានឲ្យបានគ្រប់គ្រាន់!" };
  }
  if (String(newPassword).trim().length < 6) {
    return { success: false, message: "លេខសម្ងាត់ថ្មីត្រូវមានយ៉ាងតិច ៦ តួអក្សរ!" };
  }
  // ==== FIX (សំណើថ្មី "ការ Activate គណនីត្រូវពិតប្រាកដ")៖ មុននេះ Function នេះត្រួតពិនិត្យត្រឹមតែ newFullName មិនទទេ
  // (មិនបានប្រៀបធៀបថាខុសពីទម្រង់ដើមដែរឬទេ) ហើយមិនដែលត្រួតពិនិត្យទាល់តែសោះថា newPassword ខុសពី oldPassword — ដូច្នេះ
  // អ្នកប្រើអាចចុច "ប្តូរលេខសម្ងាត់ដំបូង" ដោយគ្រាន់តែវាយបញ្ចូលលេខសម្ងាត់លំនាំដើម ("pec21") ដដែលម្តងទៀត ហើយទុកឈ្មោះ
  // គោត្តនាម-នាមជាឈ្មោះឃុំសង្កាត់/ស្រុកលំនាំដើម (ដែល bulkCreateAllCommuneAccounts_()/resetCommuneUserToDefault()
  // កំណត់ជាតម្លៃដើមស្រាប់) ក៏អាច "ដំណើរការ" (Activate) បានដដែល ដោយពុំបានផ្លាស់ប្តូរអ្វីពិតប្រាកដសោះ — ត្រូវបង្ខំឲ្យ
  // ទាំងពីរផ្នែកខុសពីទម្រង់ដើមជាក់ស្តែង ទើបចាត់ទុកថាបានធ្វើការ Activate ពិតប្រាកដ (ត្រួតពិនិត្យ newPassword!==
  // oldPassword នៅទីនេះ មុននឹង Lock ឃើញថា Server មិនទាន់ស្កេន Sheet ទេ ព្រោះនេះជាការប្រៀបធៀបតម្លៃដែលអ្នកប្រើផ្ទាល់
  // បញ្ចូល មិនទាមទារទិន្នន័យ Sheet ទេ — ចំណែកការត្រួតពិនិត្យ newFullName ត្រូវធ្វើក្រោយរកឃើញជួរដេក ព្រោះត្រូវការតម្លៃ
  // ឈ្មោះលំនាំដើមបច្ចុប្បន្នពី Sheet ជាមុនសិន — មើលខាងក្រោម) ====
  if (String(newPassword).trim() === String(oldPassword).trim()) {
    return { success: false, message: "លេខសម្ងាត់ថ្មីត្រូវខុសពីលេខសម្ងាត់បច្ចុប្បន្ន (លំនាំដើម)! សូមកំណត់លេខសម្ងាត់ថ្មីផ្ទាល់ខ្លួន កុំប្រើលេខសម្ងាត់ដដែល។" };
  }
  var cache = CacheService.getScriptCache();
  var attemptKey = "fcpAttempts_" + inputUser;
  var attempts = 0;
  try { attempts = Number(cache.get(attemptKey)) || 0; } catch (eGet) {}
  if (attempts >= FIRST_LOGIN_CHANGE_MAX_ATTEMPTS_) {
    return { success: false, message: "សាកល្បងខុសច្រើនដងពេក! សូមរង់ចាំមួយសន្ទុះ (រហូតដល់ ១ម៉ោង) ឬទាក់ទង Admin ប្រព័ន្ធ។" };
  }

  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(10000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់! សូមព្យាយាមម្តងទៀត។" };
  }
  try {
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === inputUser) {
        if (String(data[i][13] || "") !== "1") {
          return { success: false, message: "គណនីនេះលែងតម្រូវឲ្យប្តូរលេខសម្ងាត់ដំបូងទៀតហើយ។ សូមចូលប្រើប្រាស់ធម្មតា។" };
        }
        if (!verifyPassword_(oldPassword, data[i][2])) {
          try { cache.put(attemptKey, String(attempts + 1), FIRST_LOGIN_CHANGE_LOCKOUT_SECONDS_); } catch (eCachePut) {}
          return { success: false, message: "លេខសម្ងាត់បច្ចុប្បន្នមិនត្រឹមត្រូវទេ!" };
        }
        // ==== FIX (សំណើថ្មី "ការ Activate គណនីត្រូវពិតប្រាកដ" — ត្រួតពិនិត្យផ្នែកឈ្មោះ)៖ data[i][3] នៅចំណុចនេះ
        // នៅតែជាតម្លៃលំនាំដើម (ឈ្មោះឃុំសង្កាត់/ស្រុកខ្លួនឯង) ព្រោះ Flag ជួរឈរ ១៤ ("1") បញ្ជាក់ថាគណនីនេះមិនទាន់ធ្វើ
        // ការប្តូរលេខសម្ងាត់ដំបូងដោយជោគជ័យសោះ (ចាប់តាំងពីបង្កើត ឬចាប់តាំងពី resetCommuneUserToDefault()/
        // resetDistrictUserToDefault() លុបចោលចុងក្រោយ) — ដូច្នេះការប្រៀបធៀប newFullName ជាមួយតម្លៃបច្ចុប្បន្ននេះ ស្មើនឹង
        // ការប្រៀបធៀបជាមួយទម្រង់ដើមពិតប្រាកដ ====
        if (String(newFullName).trim() === String(data[i][3] || "").trim()) {
          return { success: false, message: "សូមប្តូរឈ្មោះគោត្តនាម-នាមរបស់អ្នកឲ្យពិតប្រាកដ (មិនអាចទុកជាឈ្មោះឃុំសង្កាត់/ស្រុកលំនាំដើមបានទេ)!" };
        }
        try { cache.remove(attemptKey); } catch (eCacheRm) {}
        sheet.getRange(i + 1, 3).setValue(makePasswordHash_(newPassword));
        sheet.getRange(i + 1, 4).setValue(sanitizeForSheetCell_(String(newFullName).trim()));
        sheet.getRange(i + 1, 14).setValue("");
        return { success: true, message: "ប្តូរលេខសម្ងាត់ជោគជ័យ! សូម Login ម្តងទៀតដោយប្រើលេខសម្ងាត់ថ្មី។" };
      }
    }
    return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
}

// អនុញ្ញាតឲ្យអ្នកប្រើប្រាស់ណាម្នាក់ កែប្រែគោត្តនាម-នាម/លេខសម្ងាត់ខ្លួនឯង (មិនតម្រូវសិទ្ធិ Admin ទេ — គ្រប់តួនាទីអាចប្រើបាន)
// ការប្តូរលេខសម្ងាត់ ត្រូវផ្តល់លេខសម្ងាត់បច្ចុប្បន្នដើម្បីបញ្ជាក់ (ការពារកុំឲ្យអ្នកដទៃប្តូរបានបើទុកកម្មវិធីមិនបានចាកចេញ)
function updateOwnAccount(currentUsername, payload, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (ល្បឿន)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_()/getUserRow_() ខាងលើ
  // (ដែលបានស្កេន Sheet ទាំងមូលរួចហើយម្តង) ជំនួសការអាន Sheet ទាំងមូលម្តងទៀត ====
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  var inputUser = String(currentUsername || "").trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === inputUser) {
      if (payload && payload.fullName) sheet.getRange(i + 1, 4).setValue(String(payload.fullName).trim());
      if (payload && payload.password) {
        if (!verifyPassword_(payload.currentPassword, data[i][2])) {
          return { success: false, message: "លេខសម្ងាត់បច្ចុប្បន្នមិនត្រឹមត្រូវទេ!" };
        }
        if (String(payload.password).trim().length < 6) {
          return { success: false, message: "លេខសម្ងាត់ថ្មីត្រូវមានយ៉ាងតិច ៦ តួអក្សរ!" };
        }
        sheet.getRange(i + 1, 3).setValue(makePasswordHash_(payload.password));
      }
      return { success: true, message: "កែប្រែព័ត៌មានគណនីជោគជ័យ!" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីទេ!" };
}

function updateUserAccount(currentUsername, userObj, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ឥឡូវអនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុកកែប្រែគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនបានដែរ —
  // ត្រួតពិនិត្យបែបគ្រោងគ្រាំង (Coarse) នៅត្រង់នេះសិន (មិនមែន isAdmin_ ក៏ដេញចេញភ្លាមទេ បើជាគ្រូប្រចាំក្រុងស្រុក)
  // ការត្រួតពិនិត្យលម្អិត (Fine-grained — ត្រូវជាគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនប៉ុណ្ណោះ) នៅខាងក្រោម ក្រោយពីរកឃើញជួរដេក ====
  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិកែប្រែគណនីទេ!" };
  }
  // ==== FIX (សំខាន់ណាស់ — លុបចន្លោះប្រហោងសុវត្ថិភាព)៖ addUserAccount() ខាងលើ ការពាររួចហើយកុំឲ្យអ្នកមិនមែន SuperAdmin
  // បង្កើតគណនី SuperAdmin ថ្មីបាន ប៉ុន្តែ Function នេះ (កែប្រែគណនីដែលមានស្រាប់) មិនធ្លាប់មានការការពារដូចគ្នាទេ — Admin/
  // PEC21 ធម្មតា (ដែលឆ្លងកាត់ isAdmin_() ដូចគ្នា) អាចបើក Modal កែគណនី ហើយប្តូរតួនាទីខ្លួនឯង ឬអ្នកដទៃ ទៅជា SuperAdmin
  // បានដោយសេរី — ត្រូវការពារដូចគ្នាបេះបិទ ====
  if (userObj.role === ROLE_SUPERADMIN && !isSuperAdmin_(currentUsername)) {
    return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាចផ្តល់តួនាទី SuperAdmin បាន!" };
  }
  // ==== FIX (ល្បឿន)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_()/getUserRow_() ខាងលើ
  // (ដែលបានស្កេន Sheet ទាំងមូលរួចហើយម្តង) ជំនួសការអាន Sheet ទាំងមូលម្តងទៀត ====
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userObj.id)) {
      // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ អ្នកមិនមែន Admin-tier (ឧ. គ្រូប្រចាំក្រុងស្រុក) កែប្រែបានតែគណនីឃុំសង្កាត់
      // ក្នុងស្រុកខ្លួនប៉ុណ្ណោះ ====
      if (!canManageTargetAccount_(currentUsername, data[i][4], data[i][5])) {
        return { success: false, message: "គ្មានសិទ្ធិកែប្រែគណនីនេះទេ!" };
      }
      // ==== FIX (សុវត្ថិភាព — សំខាន់ណាស់, ការលើកកម្ពស់សិទ្ធិខ្លួនឯង)៖ addUserAccount() ការពាររួចហើយកុំឲ្យអ្នកមិនមែន
      // SuperAdmin *បង្កើត* គណនី SuperAdmin ថ្មី ហើយខាងលើ (userObj.role === ROLE_SUPERADMIN) ក៏ការពារកុំឲ្យ *ផ្តល់*
      // តួនាទី SuperAdmin ដល់អ្នកដទៃដែរ — ប៉ុន្តែមិនធ្លាប់មានការការពារ ករណីជួរដេកគោលដៅ (Target) *ជា SuperAdmin រួចស្រាប់*
      // ទេ! Admin/PEC21 ធម្មតា (ដែលឆ្លងកាត់ isAdmin_() ដូចគ្នា) អាចហៅ Function នេះ ដោយផ្ញើ userObj ដែលគ្មាន Field
      // "role" (រំលងការត្រួតពិនិត្យខាងលើ) ប៉ុន្តែមាន Field "password" ថ្មី ដើម្បីប្តូរលេខសម្ងាត់របស់ SuperAdmin ដោយផ្ទាល់
      // (ឬសូម្បីតែ Suspend/លុបគណនីនោះចោលក៏បាន — មើល toggleUserSuspend()/deleteUserAccount() ខាងក្រោម) — ស្មើនឹង
      // លួច Take Over គណនី SuperAdmin ទាំងស្រុង ====
      if (data[i][4] === ROLE_SUPERADMIN && !isSuperAdmin_(currentUsername)) {
        return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាចកែប្រែគណនី SuperAdmin បាន!" };
      }
      // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គ្រូប្រចាំក្រុងស្រុក (មិនមែន Admin-tier) មិនអាចប្តូរតួនាទីគណនីនេះទៅជា
      // អ្វីផ្សេងក្រៅពី ROLE_COMMUNE បានទេ ហើយក៏មិនអាចផ្ទេរគណនីនេះទៅស្រុកដទៃបានដែរ (ការពារការលើកកម្ពស់សិទ្ធិ/រំលង Scope) ====
      if (!isAdmin_(currentUsername)) {
        if (userObj.role !== undefined && userObj.role !== ROLE_COMMUNE) {
          return { success: false, message: "អ្នកមិនអាចប្តូរតួនាទីគណនីនេះបានទេ!" };
        }
        if (userObj.district !== undefined && String(userObj.district).trim() !== data[i][5]) {
          return { success: false, message: "អ្នកមិនអាចផ្ទេរគណនីនេះទៅស្រុកដទៃបានទេ!" };
        }
      }
      // ==== FIX (សំណើ)៖ អនុញ្ញាតឲ្យ Admin/SuperAdmin ប្តូរឈ្មោះគណនី (Username) បាន — មុននេះ ប្រអប់នេះត្រូវបានចាក់សោ
      // (Disabled) នៅ Client ជានិច្ចពេលកែគណនីដែលមានស្រាប់ ហើយ Function នេះក៏មិនធ្លាប់អាន userObj.username ទាល់តែសោះ
      // ដូច្នេះទោះបីជាបើកប្រអប់ឲ្យវាយបញ្ចូលបាន ក៏មិនដែលរក្សាទុកដែរ។ ត្រូវត្រួតពិនិត្យថាឈ្មោះថ្មីមិនស្ទួនគណនីណាមួយផ្សេង
      // ទៀត (Case-insensitive ដូច Login/addUserAccount) មុននឹងកែប្រែ — បើកំពុងប្តូរឈ្មោះខ្លួនឯង (ឬអ្នកកំពុង Login
      // ស្រាប់) Session នៅឧបករណ៍នោះនឹងលែងត្រូវគ្នា (ព្រោះ Username ចាស់លែងមានក្នុង Sheet) ហើយនឹងត្រូវ Logout ដោយ
      // ស្វ័យប្រវត្តិ (មើល validateSession/startSessionPolling_) ជាឥរិយាបថធម្មតារួចស្រាប់ — ត្រូវ Login ម្តងទៀតដោយ
      // ឈ្មោះថ្មី ====
      if (userObj.username) {
        var newUsername = String(userObj.username).trim();
        if (!newUsername) return { success: false, message: "ឈ្មោះគណនីមិនអាចទទេបានទេ!" };
        var newUsernameLower = newUsername.toLowerCase();
        if (newUsernameLower !== String(data[i][1]).trim().toLowerCase()) {
          // ==== FIX (សំណើថ្មី "ចាក់សោឈ្មោះគណនីឃុំសង្កាត់អចិន្ត្រៃយ៍" — ចំណុចទី៤)៖ ដើម្បីការពារកុំឲ្យទិន្នន័យ
          // លាយច្របល់រវាងឃុំសង្កាត់ដោយប្រការណាមួយ (ឧ. គ្រូប្រចាំក្រុងស្រុកកែច្រំប្តូរឈ្មោះគណនីខុសកន្លែងដោយចៃដន្យ ឬ
          // ដោយចេតនា) ឈ្មោះគណនីរបស់គណនីឃុំសង្កាត់ (ROLE_COMMUNE) ត្រូវបានចាក់សោអចិន្ត្រៃយ៍ បន្ទាប់ពីបង្កើតរួច —
          // មិនអាចប្តូរបានទៀតទេ ដោយអ្នកប្រើណាក៏ដោយ (រួមទាំង Admin/SuperAdmin ផងដែរ តាមសំណើច្បាស់លាស់របស់អ្នកប្រើ) —
          // ខុសពីមុខងារ resetCommuneUserToDefault() ដែលនៅតែអនុញ្ញាតឲ្យកែប្រែឈ្មោះពេញ/លេខសម្ងាត់ដដែល (មិនប៉ះពាល់
          // ដោយការចាក់សោនេះទេ ព្រោះវាមិនដែលសរសេរជួរឈរទី ២ ឈ្មោះគណនីទេ) — Client (Index.html) ក៏ចាក់សោប្រអប់នេះ
          // (Disabled) ជានិច្ចនៅពេលកែប្រែគណនីដែលមានស្រាប់ដែរ (មើល editCommuneUser()) — ការត្រួតពិនិត្យត្រង់នេះគឺជា
          // ការការពារបន្ថែម (Defense in depth) ក្នុងករណី Client ត្រូវបានរំលងដោយ DevTools ====
          if (data[i][4] === ROLE_COMMUNE) {
            return { success: false, message: "ឈ្មោះគណនីឃុំសង្កាត់ត្រូវបានចាក់សោអចិន្ត្រៃយ៍ បន្ទាប់ពីបង្កើតរួច មិនអាចប្តូរបានទៀតទេ (ដើម្បីការពារកុំឲ្យទិន្នន័យលាយច្របល់រវាងឃុំសង្កាត់)!" };
          }
          for (var k = 1; k < data.length; k++) {
            if (k !== i && String(data[k][1]).trim().toLowerCase() === newUsernameLower) {
              return { success: false, message: "ឈ្មោះគណនីនេះមានអ្នកប្រើរួចហើយ!" };
            }
          }
          sheet.getRange(i + 1, 2).setValue(newUsername);
        }
      }
      if (userObj.fullName) sheet.getRange(i + 1, 4).setValue(sanitizeForSheetCell_(userObj.fullName));
      if (userObj.role) {
        // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ROLE_COMMUNE ត្រូវការស្រុក+ឃុំសង្កាត់ដាច់ដោយឡែក (ជួរឈរ M) —
        // មិនអាចប្រើ district=role ដូចគណនីស្រុកបានទេ (ឈ្មោះឃុំសង្កាត់មិនតែមួយគត់ទូទាំងខេត្ត) ====
        var district, newCommuneVal = "";
        if (isDistrictRole_(userObj.role)) {
          district = userObj.role;
          // ==== FIX (សំណើថ្មី "លើកលែងសម្រាប់ SuperAdmin")៖ បិទចន្លោះប្រហោង — addUserAccount() ខាងលើ ត្រួតពិនិត្យ
          // កំណត់ចំនួន DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ ពេលបង្កើតថ្មី ប៉ុន្តែ Function នេះ (កែប្រែគណនីមានស្រាប់)
          // មិនធ្លាប់ត្រួតពិនិត្យទេ — Admin/PEC21 អាចបើក Modal កែគណនីធម្មតាមួយ ហើយប្តូរ Role វា ទៅជាគណនីស្រុកមួយ
          // ដែលពេញកំណត់រួចហើយ ដោយរំលងការទប់ស្កាត់ខាងលើបានទាំងស្រុង — ត្រូវត្រួតពិនិត្យត្រង់នេះដែរ ដូចគ្នា (រាប់ខ្លួនឯង
          // ចេញ ដោយត្រួតពិនិត្យតែពេលកំពុងផ្លាស់ប្តូរស្រុកជាក់ស្តែង — ការកែប្រែធម្មតា ឧ. ប្តូរឈ្មោះ/លេខសម្ងាត់ ដោយមិន
          // ប្តូរស្រុកសោះ មិនរងផលប៉ះពាល់ទេ — SuperAdmin ខ្លួនឯង នៅតែលើកលែងបានដូច addUserAccount()) ====
          if (district !== data[i][5] && !isSuperAdmin_(currentUsername) &&
              countDistrictRoleAccounts_(data, district) >= DISTRICT_ACCOUNT_MAX_PER_DISTRICT_) {
            return { success: false, message: "ស្រុក/ក្រុង \"" + district + "\" មានគណនីគ្រូប្រចាំក្រុងស្រុកគ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + DISTRICT_ACCOUNT_MAX_PER_DISTRICT_ + " គណនីក្នុងមួយស្រុក)! សូមទាក់ទង SuperAdmin ដើម្បីស្នើសុំបង្កើតបន្ថែម។" };
          }
        } else if (userObj.role === ROLE_COMMUNE) {
          district = String(userObj.district || data[i][5] || "").trim();
          newCommuneVal = String(userObj.commune || data[i][12] || "").trim();
          if (!isValidCommune_(district, newCommuneVal)) {
            return { success: false, message: "សូមជ្រើសរើសស្រុក/ក្រុង និងឃុំ/សង្កាត់ឲ្យត្រឹមត្រូវ!" };
          }
          // ==== FIX (ចំណុចទី៤ — ប្រការទី២)៖ addUserAccount() ខាងលើ ត្រួតពិនិត្យរួចហើយកុំឲ្យបង្កើតគណនីឃុំសង្កាត់ថ្មី
          // ហួសពី COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ ក្នុងមួយឃុំសង្កាត់ណាមួយ ប៉ុន្តែ Function នេះ (កែប្រែគណនីមានស្រាប់)
          // មិនធ្លាប់ត្រួតពិនិត្យទេ — អ្នកគ្រប់គ្រង (Admin/គ្រូប្រចាំក្រុងស្រុក) អាចបើក Modal កែគណនីឃុំសង្កាត់ណាមួយ ហើយ
          // ប្តូរឃុំសង្កាត់របស់វា ទៅជាឃុំសង្កាត់ដែលមានគណនីគ្រប់ចំនួនកំណត់រួចហើយ ដោយរំលងការទប់ស្កាត់ខាងលើបានទាំងស្រុង —
          // ជាហេតុនាំឲ្យទិន្នន័យពីរឃុំសង្កាត់ខុសគ្នាអាចច្រឡំគ្នា (ឬចូលរួមគ្នា ខុសពីគោលបំណង) — ត្រួតពិនិត្យតែពេលកំពុង
          // ផ្លាស់ប្តូរស្រុក/ឃុំសង្កាត់ជាក់ស្តែងប៉ុណ្ណោះ (ការកែប្រែធម្មតា ឧ. ប្តូរឈ្មោះពេញ/លេខសម្ងាត់ ដោយមិនប្តូរឃុំសង្កាត់
          // សោះ មិនរងផលប៉ះពាល់ទេ) ====
          if ((district !== String(data[i][5]).trim() || newCommuneVal !== String(data[i][12] || "").trim()) &&
              countCommuneRoleAccountsForCommune_(data, district, newCommuneVal) >= COMMUNE_ACCOUNT_MAX_PER_COMMUNE_) {
            return { success: false, message: "ឃុំ/សង្កាត់ \"" + newCommuneVal + "\" មានគណនីឃុំសង្កាត់គ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ + " គណនីក្នុងមួយឃុំសង្កាត់)! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
          }
        } else {
          district = "";
        }
        sheet.getRange(i + 1, 5).setValue(userObj.role);
        sheet.getRange(i + 1, 6).setValue(district);
        sheet.getRange(i + 1, 13).setValue(newCommuneVal);
      } else if (userObj.commune !== undefined && data[i][4] === ROLE_COMMUNE) {
        // ==== ប្តូរឃុំសង្កាត់ ដោយមិនប្តូរ Role (នៅតែស្រុកដដែល) — ឧ. គ្រូប្រចាំក្រុងស្រុកផ្ទេរគណនីនេះទៅឃុំសង្កាត់ផ្សេង ====
        var newCommuneOnly = String(userObj.commune || "").trim();
        if (!isValidCommune_(data[i][5], newCommuneOnly)) {
          return { success: false, message: "សូមជ្រើសរើសឃុំ/សង្កាត់ឲ្យត្រឹមត្រូវ!" };
        }
        // ==== FIX (ចំណុចទី៤ — ប្រការទី២)៖ ដូចគ្នានឹងខាងលើ (ករណីប្តូរ Role ជាមួយគ្នា) — ផ្លូវនេះ (ប្តូរឃុំសង្កាត់ ដោយឡែក
        // ពី Role) ក៏ត្រូវការពារកុំឲ្យផ្ទេរទៅឃុំសង្កាត់ដែលមានគណនីគ្រប់ចំនួនកំណត់រួចហើយដែរ ====
        if (newCommuneOnly !== String(data[i][12] || "").trim() &&
            countCommuneRoleAccountsForCommune_(data, data[i][5], newCommuneOnly) >= COMMUNE_ACCOUNT_MAX_PER_COMMUNE_) {
          return { success: false, message: "ឃុំ/សង្កាត់ \"" + newCommuneOnly + "\" មានគណនីឃុំសង្កាត់គ្រប់ចំនួនកំណត់រួចហើយ (អតិបរមា " + COMMUNE_ACCOUNT_MAX_PER_COMMUNE_ + " គណនីក្នុងមួយឃុំសង្កាត់)! សូមទាក់ទង Admin ប្រព័ន្ធ ឬ SuperAdmin។" };
        }
        sheet.getRange(i + 1, 13).setValue(newCommuneOnly);
      }
      if (userObj.password) sheet.getRange(i + 1, 3).setValue(makePasswordHash_(userObj.password));
      if (userObj.status) sheet.getRange(i + 1, 7).setValue(userObj.status);
      if (userObj.gmail !== undefined) sheet.getRange(i + 1, 9).setValue(String(userObj.gmail || "").trim());
      // ==== FIX (សុពលភាព/រយៈពេលប្រើប្រាស់ — Account Expiry)៖ អនុញ្ញាតឲ្យ Admin/SuperAdmin កែ/លុបថ្ងៃផុតកំណត់
      // (ផ្អែកលើ Role ថ្មី បើកំពុងប្តូរ Role ក្នុងពេលតែមួយ ឬ Role ដើម បើមិនប្តូរ) — ចោលចេញ បើ Role មិនអនុញ្ញាត ====
      if (userObj.expiryDate !== undefined) {
        var effectiveRole = userObj.role || data[i][4];
        var newExpiry = isExpiryApplicableRole_(effectiveRole) ? normalizeExpiryInput_(userObj.expiryDate) : "";
        sheet.getRange(i + 1, 12).setValue(newExpiry);
      }
      // ==== FIX (ល្បឿន)៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជាផ្ទៃខាងក្រោយ (មិនចាំបាច់ឲ្យ Admin រង់ចាំ — មើល Utils.gs) ====
      scheduleSyncSheetPermissions_();
      return { success: true, message: "កែប្រែគណនីជោគជ័យ!" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

function deleteUserAccount(currentUsername, userId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ឥឡូវអនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុក លុបគណនីឃុំសង្កាត់ក្នុងស្រុកខ្លួនបានដែរ ====
  if (!isAdmin_(currentUsername) && !isDistrictRole_(getUserRole_(currentUsername))) {
    return { success: false, message: "គ្មានសិទ្ធិលុបគណនីទេ!" };
  }
  // ==== FIX (ល្បឿន)៖ ចែករំលែងទិន្នន័យ Sheet(អ្នកប្រើប្រាស់) ជាមួយ requireValidSession_()/getUserRow_() ខាងលើ
  // (ដែលបានស្កេន Sheet ទាំងមូលរួចហើយម្តង) ជំនួសការអាន Sheet ទាំងមូលម្តងទៀត ====
  var __usd_ = getUsersSheetDisplayDataOnce_();
  var sheet = __usd_.sheet;
  var data = __usd_.data;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) {
      if (data[i][1] === currentUsername) return { success: false, message: "អ្នកមិនអាចលុបគណនីខ្លួនឯងបានទេ!" };
      if (!canManageTargetAccount_(currentUsername, data[i][4], data[i][5])) {
        return { success: false, message: "គ្មានសិទ្ធិលុបគណនីនេះទេ!" };
      }
      // ==== FIX (សុវត្ថិភាព — សំខាន់ណាស់)៖ ដូចគ្នានឹង updateUserAccount() ខាងលើ — Admin/PEC21 មិនត្រូវអាចលុបគណនី
      // SuperAdmin (អ្នកដទៃ) ចោលបានទេ ====
      if (data[i][4] === ROLE_SUPERADMIN && !isSuperAdmin_(currentUsername)) {
        return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាចលុបគណនី SuperAdmin បាន!" };
      }
      sheet.deleteRow(i + 1);
      // ==== ចេតនាហៅផ្ទាល់ (Synchronous) នៅតែដដែល មិនប្តូរទៅ scheduleSyncSheetPermissions_() ដូចកន្លែងផ្សេងទេ — ព្រោះ
      // ការលុបគណនី ត្រូវការដកសិទ្ធិ Google Sheet ភ្លាមៗ ដូចគ្នានឹងហេតុផលក្នុង toggleUserSuspend() ខាងលើដែរ ====
      try { syncSheetPermissions_(); } catch (err) {}
      return { success: true, message: "លុបគណនីជោគជ័យ!" };
    }
  }
  return { success: false, message: "រកមិនឃើញគណនីនេះទេ!" };
}

// ==================== ៥.១ ការធ្វើសមកាលកម្មសិទ្ធិ Google Sheet (Sharing + Protected Ranges) ====================
// ត្រូវហៅរាល់ពេលមានការ បន្ថែម/កែប្រែ/លុប គណនីអ្នកប្រើប្រាស់ ដើម្បីធ្វើឲ្យសិទ្ធិលើ Google Sheet ត្រូវគ្នានឹងគណនីក្នុង App
// Logic:
//  - SuperAdmin/Admin/PEC21 (មាន Gmail) → ក្លាយជា Editor លើឯកសារទាំងមូល + អាចកែប្រែគ្រប់ Tab ស្រុកទាំងអស់
//    (FIX៖ PEC21 ត្រូវបានលើកឡើងពី Viewer ទៅ Editor ដូច Admin ទាំងស្រុង តាមសំណើ)
//  - សង្កេតការណ៍ (មាន Gmail) → ក្លាយជា Viewer លើឯកសារ (មើលបានគ្រប់ Tab ប៉ុន្តែកែប្រែមិនបានទេ)
//  - គ្រូប្រចាំស្រុកម្នាក់ៗ ត្រូវបានបន្ថែមជា "Editor តាម Protected Range" នៅលើ Tab ស្រុករបស់ខ្លួនតែមួយគត់
//    (Protected Range Editor អាចផ្តល់សិទ្ធិកែប្រែបាន ទោះបីជាសិទ្ធិឯកសារទាំងមូលជា Viewer ក៏ដោយ)
// ==== FIX (សុវត្ថិភាព — សំខាន់ណាស់)៖ បិទ "ការចែករំលែកទូទៅ" (General Access / Link Sharing) របស់ Google Sheet ម្នាក់ៗ
// ឲ្យទៅជា "កំណត់" (Private/Restricted) ជានិច្ច — ដើម្បីកុំឲ្យនរណាម្នាក់អាចយកតំណភ្ជាប់ (Link) ទៅចែកបន្តឲ្យអ្នកដទៃ
// ណាមួយក៏បាន ចូលកែប្រែបាន ទោះមិនមែនជា Gmail ដែលបានផ្តល់សិទ្ធិដោយផ្ទាល់ (addEditor/addViewer ខាងក្រោម) ក៏ដោយ។
// មុននេះ Function នេះហៅតែ addEditor()/addViewer()/removeEditor() ប៉ុណ្ណោះ (កំណត់តែបញ្ជីអ្នកប្រើប្រាស់ជាក់លាក់)
// ប៉ុន្តែមិនដែលបានបិទ "ការចែករំលែកទូទៅ" ទាល់តែសោះ — ២យន្តការនេះឯករាជ្យពីគ្នា (Google Drive) ដូច្នេះបើ Sheet ណាមួយ
// កំណត់ទុកជា "អ្នកគ្រប់គ្នាដែលមានតំណភ្ជាប់អាចកែប្រែបាន" ស្រាប់ (ឧ. តាំងពីពេលបង្កើតដំបូង ឬត្រូវបានប្តូរដោយចៃដន្យ)
// នរណាម្នាក់ដែលមានតំណភ្ជាប់នោះ អាចកែប្រែបាន ទោះមិនស្ថិតក្នុងបញ្ជីអនុញ្ញាតក៏ដោយ — ជៀសវាង Error ក្នុងករណី Google
// Workspace ស្ថាប័នខ្លះដាក់កំហិត Domain Policy មិនអនុញ្ញាតឲ្យ Script ប្តូរកម្រិតនេះ (ដូចគ្នានឹង setSharing() ផ្នែក Logo) ====
function restrictFileToPrivate_(file) {
  try { file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (err) {}
}

function syncSheetPermissions_() {
  var ss = getSS_();
  var file = DriveApp.getFileById(getTargetSheetId_());
  restrictFileToPrivate_(file);
  var usersSheet = ensureUsersSheet_(ss);
  var data = usersSheet.getDataRange().getDisplayValues();

  var adminEmails = []; // SuperAdmin/Admin/PEC21 — Editor លើឯកសារទាំងអស់
  var viewerEmails = []; // សង្កេតការណ៍ ប៉ុណ្ណោះ — Viewer លើ Spreadsheet មេ
  var districtEmails = {};
  DISTRICT_LIST.forEach(function(d) { districtEmails[d] = []; });

  for (var i = 1; i < data.length; i++) {
    var role = data[i][4];
    var status = data[i][6];
    var gmail = String(data[i][8] || "").trim();
    if (!gmail || status !== STATUS_ACTIVE) continue;
    // ==== FIX៖ PEC21 ត្រូវបានផ្លាស់ចេញពី viewerEmails មករាប់ចូល adminEmails ព្រោះឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង ====
    if (role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21) {
      adminEmails.push(gmail);
    } else if (role === ROLE_OBSERVER || isObserverTierRole_(role)) {
      // ==== FIX (Fix129, "តួនាទីថ្មីទាំង៥")៖ មានសិទ្ធិដូចសង្កេតការណ៍ទាំងស្រុង — ទទួល Viewer លើ Spreadsheet មេដូចគ្នា ====
      viewerEmails.push(gmail);
    } else if (isDistrictRole_(role)) {
      // *** សំខាន់៖ គ្រូប្រចាំស្រុក លែងទទួល Viewer លើ Spreadsheet មេទៀត — ដើម្បីកុំឲ្យឃើញទិន្នន័យស្រុកដទៃ ***
      districtEmails[role].push(gmail);
    }
  }

  // ---- Spreadsheet មេ (សរុបរួម)៖ SuperAdmin/Admin/PEC21=Editor, សង្កេតការណ៍=Viewer, គ្រូប្រចាំស្រុក=គ្មានសិទ្ធិចូលទាល់តែសោះ ----
  var allowedOnMaster = adminEmails.concat(viewerEmails);
  adminEmails.forEach(function(e) { try { file.addEditor(e); } catch (err) {} });
  viewerEmails.forEach(function(e) { try { file.addViewer(e); } catch (err) {} });
  removeStaleFileAccess_(file, allowedOnMaster);

  // ---- Tab សរុបរួម/ស្ថិតិ ក្នុង Spreadsheet មេ — Read-only សម្រាប់មនុស្ស (សរសេរដោយ Script ប៉ុណ្ណោះ) ----
  protectSheetEditors_(ensureDailySheet_(ss), []);
  protectSheetEditors_(ensureCommuneSumSheet_(ss), []);
  protectSheetEditors_(ensureProvinceSumSheet_(ss), []);

  // ---- Spreadsheet ដាច់ដោយឡែក តាមស្រុកនីមួយៗ៖ SuperAdmin/Admin/PEC21 (Editor គ្រប់ស្រុក) + គ្រូប្រចាំស្រុកនោះ (Editor តែស្រុកខ្លួន) ----
  DISTRICT_LIST.forEach(function(d) {
    getDistrictSpreadsheet_(d); // ធានាថា Spreadsheet របស់ស្រុកនេះមានស្រាប់ (Tab ខាងក្នុងបង្កើតតាមតម្រូវការនៅពេលក្រោយ)
    var dSs = getDistrictSpreadsheet_(d);
    var dFile = DriveApp.getFileById(dSs.getId());
    restrictFileToPrivate_(dFile);
    var allowedOnDistrict = adminEmails.concat(districtEmails[d]);
    allowedOnDistrict.forEach(function(e) { try { dFile.addEditor(e); } catch (err) {} });
    removeStaleFileAccess_(dFile, allowedOnDistrict);
  });
}

// ដកអ្នកមានសិទ្ធិ (Editor/Viewer) ចាស់ដែលលែងស្ថិតក្នុងបញ្ជីអនុញ្ញាតថ្មីចេញ — លើកលែងម្ចាស់ File ខ្លួនឯង
// (ចាំបាច់ណាស់ សម្រាប់ធានាថា គណនីគ្រូប្រចាំស្រុកមុន ដែលឥឡូវលែងត្រូវការសិទ្ធិទៀត នឹងលែងឃើញ Spreadsheet នោះទាំងស្រុង)
function removeStaleFileAccess_(file, allowedEmails) {
  var owner;
  try { owner = file.getOwner().getEmail(); } catch (err) { owner = null; }
  var allowedSet = {};
  allowedEmails.forEach(function(e) { allowedSet[String(e).trim().toLowerCase()] = true; });
  [file.getEditors(), file.getViewers()].forEach(function(list) {
    list.forEach(function(u) {
      var email = u.getEmail();
      if (!email) return;
      var lower = email.toLowerCase();
      if (owner && lower === owner.toLowerCase()) return;
      if (!allowedSet[lower]) {
        try { file.removeEditor(email); } catch (err) {}
        try { file.removeViewer(email); } catch (err) {}
      }
    });
  });
}

// ចលនាតែមួយ៖ ដកអ្នកកែប្រែចាស់ចេញ ហើយកំណត់បញ្ជីអ្នកមានសិទ្ធិកែប្រែ Sheet ថ្មី
function protectSheetEditors_(sheet, emails) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  var protection = protections.length ? protections[0] : sheet.protect();
  protection.setDescription('គ្រប់គ្រងស្វ័យប្រវត្តិដោយប្រព័ន្ធ — កុំកែប្រែសិទ្ធិផ្ទាល់ក្នុង Google Sheet');
  try { protection.setDomainEdit(false); } catch (err) {}
  var currentEditors = protection.getEditors();
  if (currentEditors.length) {
    protection.removeEditors(currentEditors.map(function(u) { return u.getEmail(); }));
  }
  if (emails && emails.length) protection.addEditors(emails);
}

// ប៊ូតុងសម្រាប់ Admin ចុចដោយដៃ (ករណីលើកទីមួយ ត្រូវ Authorize សិទ្ធិ Drive ជាមុន ឬក្នុងករណីចង់បង្ខំធ្វើសមកាលកម្មភ្លាមៗ)
function syncSheetPermissions(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  try {
    syncSheetPermissions_();
    return { success: true, message: "ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជោគជ័យ!" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==================== ៦. ទិន្នន័យចុះឈ្មោះបោះឆ្នាំប្រចាំថ្ងៃ (Daily Entry — CRUD) ====================
// ចាប់ពីឥឡូវ Period Sheet (ខាងលើ) គឺជាប្រភពទិន្នន័យ "តែមួយគត់"៖ App សរសេរ/អាន ត្រង់ជួរដេកឃុំ/សង្កាត់
// ដែលមានស្រាប់ (បង្កើតជាមុនម្តងក្នុង Tab ថ្ងៃខែនោះ) ជំនួសការបន្ថែមជួរដេកថ្មីៗគ្មានទីបញ្ចប់ដូចមុន។

