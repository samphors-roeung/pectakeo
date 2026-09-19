// ==================== Comments.gs — មតិយោបល់/រាយការណ៍បញ្ហា (Observer → SuperAdmin/គ្រូប្រចាំស្រុក) ====================
// មុខងារនេះអនុញ្ញាតឲ្យតួនាទី "សង្កេតការណ៍" (Observer — មើលទិន្នន័យបានគ្រប់ស្រុក ប៉ុន្តែមិនអាចបញ្ចូល/កែ/លុបទិន្នន័យបានឡើយ)
// អាចសរសេរមតិយោបល់/រាយការណ៍បញ្ហា ទៅកាន់ SuperAdmin (ឃើញជានិច្ច) និងគ្រូប្រចាំស្រុកជាក់លាក់ណាមួយ (បើជ្រើសរើសស្រុកគោលដៅ) ។
// SuperAdmin/Admin និងគ្រូប្រចាំស្រុកដែលពាក់ព័ន្ធ អាចមើលឃើញ ឆ្លើយតប ឬសម្គាល់ថាដោះស្រាយរួច។ គ្រប់តួនាទីអាចប្រើមុខងារនេះ
// ដើម្បីសរសេរមតិយោបល់បាន (មិនកំណត់ត្រឹមតែ សង្កេតការណ៍ ទេ) ប៉ុន្តែមើលឃើញតែមតិយោបល់ដែលពាក់ព័ន្ធនឹងខ្លួនប៉ុណ្ណោះ (មើល getComments)។

var SHEET_COMMENTS_ = "មតិយោបល់";
var COMMENT_STATUS_NEW_ = "ថ្មី";
var COMMENT_STATUS_REPLIED_ = "បានឆ្លើយតប";
var COMMENT_STATUS_RESOLVED_ = "បានដោះស្រាយ";
var COMMENT_TARGET_GENERAL_LABEL_ = "ទូទៅ (SuperAdmin)";
var COMMENT_TARGET_ALL_SENTINEL_ = "__ALL__"; // តម្លៃពិសេសពី Client សម្រាប់ជម្រើស "គ្រប់ក្រុងស្រុកទាំងអស់" (Broadcast) — មិនមែនឈ្មោះស្រុកពិតទេ
// ==== FIX (Fix117, "សម្រាប់គណនីគ្រូ គោលដៅសុំឲ្យមាន SuperAdmin, Admin, PEC21 ផង")៖ តម្លៃពិសេសដែលអាចផ្ញើមកក្នុង
// Parameter "targetCommune" ជំនួសឈ្មោះឃុំ/សង្កាត់ពិតប្រាកដ — សម្គាល់ថាគ្រូប្រចាំក្រុងស្រុក (District-tier) ចង់សរសេរ
// មតិយោបល់ ដោយចង់ជូនដំណឹងច្បាស់លាស់ថាសម្រាប់តួនាទីណា (SuperAdmin/Admin/PEC21 សុទ្ធតែឃើញមតិយោបល់ទាំងអស់ស្រាប់ហើយ
// ដោយមិនគិតពីគោលដៅក៏ដោយ — នេះគ្រាន់តែជាព័ត៌មានបញ្ជាក់បន្ថែម/ស្លាកសម្គាល់ ដើម្បីភាពច្បាស់លាស់ប៉ុណ្ណោះ មិនប៉ះពាល់សិទ្ធិមើល
// ដែលមានស្រាប់ទេ)។ ក្រុងស្រុកគោលដៅ (targetDistrict) នៅតែជាក្រុងស្រុកខ្លួនជានិច្ច (Defense in depth ដដែល — មិនប្តូរ) ====
var COMMENT_TARGET_ROLE_SENTINELS_ = ["SuperAdmin", "Admin", "PEC21"];

// ==== ធានាថា Tab "មតិយោបល់" មានស្រាប់ក្នុង Spreadsheet មេ (បង្កើតដោយស្វ័យប្រវត្តិលើកដំបូង) ====
function ensureCommentsSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_COMMENTS_);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COMMENTS_);
    sheet.appendRow(["ID", "ពេលវេលា", "គណនីអ្នកសរសេរ", "ឈ្មោះអ្នកសរសេរ", "តួនាទី", "ស្រុកគោលដៅ", "សារ", "ស្ថានភាព", "ការឆ្លើយតប", "អ្នកឆ្លើយតប", "ពេលវេលាឆ្លើយតប", "ឃុំសង្កាត់គោលដៅ"]);
    sheet.getRange(1, 1, 1, 12).setFontWeight("bold").setBackground("#eef2f9");
    sheet.setFrozenRows(1);
    try { protectSheetEditors_(sheet, []); } catch (err) {} // Read-only សម្រាប់មនុស្ស — សរសេរដោយ Script ប៉ុណ្ណោះ
  } else {
    // ==== FIX (សំណើថ្មី "គោលដៅ ក្រុងស្រុក/ឃុំសង្កាត់ ២កម្រិត")៖ ការធ្វើឲ្យប្រសើរឡើង (Migration) — Sheet ចាស់ដែលមានតែ
    // ១១ជួរឈរ (បង្កើតមុននឹងមានមុខងារនេះ — គ្មានជួរឈរ "ឃុំសង្កាត់គោលដៅ" ទេ) ត្រូវបន្ថែមជួរឈរទី១២ដោយស្វ័យប្រវត្តិ
    // ដើម្បីកុំឲ្យទិន្នន័យចាស់ៗ (ជួរដេកមុនៗ) ខូច/ប្តូរទីតាំងខុសនៅពេលអានតទៅ ====
    try {
      if (sheet.getLastColumn() < 12) {
        sheet.getRange(1, 12).setValue("ឃុំសង្កាត់គោលដៅ").setFontWeight("bold").setBackground("#eef2f9");
      }
    } catch (err) {}
  }
  return sheet;
}

// ==== ព័ត៌មានចាំបាច់សម្រាប់សាងសង់ទម្រង់សរសេរមតិយោបល់ (ក្រុងស្រុក+ឃុំសង្កាត់គោលដៅ អាចជ្រើសរើសបាន តាមកម្រិតសិទ្ធិ) ====
// ==== FIX (សំណើថ្មី "គោលដៅ ក្រុងស្រុក/ឃុំសង្កាត់ ២កម្រិត")៖ ត្រឡប់ commune/communeOrder/districtType បន្ថែម
// ដើម្បីឲ្យ Client សាងសង់ Dropdown ជ្រើសរើសឃុំ/សង្កាត់ (Cascading — ស្របតាមកម្រិតសិទ្ធិ ៣ថ្នាក់) ====
function getCommentsBootstrap(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  return {
    success: true, role: role,
    district: getUserDistrict_(currentUsername) || "",
    commune: getUserCommune_(currentUsername) || "",
    districts: DISTRICT_LIST,
    communeOrder: COMMUNE_ORDER,
    districtType: DISTRICT_TYPE
  };
}

// ==== គ្រប់តួនាទីអាចសរសេរមតិយោបល់បាន — targetDistrict = "" មានន័យថា "ទូទៅ" (SuperAdmin ឃើញ) ====
// ==== FIX (សំណើថ្មី "គោលដៅ ក្រុងស្រុក/ឃុំសង្កាត់ ២កម្រិត")៖ បន្ថែម targetCommune — Server ត្រូវតែផ្ទៀងផ្ទាត់/ចាក់សោ
// គោលដៅ ដោយខ្លួនឯង ស្របតាមកម្រិតសិទ្ធិ ៣ថ្នាក់ខាងក្រោម (Defense in depth — មិនទុកចិត្តលើតម្លៃពី Client ទាំងស្រុងទេ ដូច
// Pattern saveDistrictDayEntries)៖
//  ១. គណនីមើលបានគ្រប់ក្រុងស្រុក (SuperAdmin/Admin/PEC21/សង្កេតការណ៍)៖ ជ្រើសរើសបានគ្រប់ក្រុងស្រុក+ឃុំសង្កាត់ណាមួយក៏បាន
//  ២. គ្រូប្រចាំក្រុងស្រុក៖ ក្រុងស្រុកគោលដៅជានិច្ចជាក្រុងស្រុកខ្លួន — ជ្រើសរើសបានតែឃុំសង្កាត់ក្នុងក្រុងស្រុកខ្លួនប៉ុណ្ណោះ
//  ៣. គណនីឃុំសង្កាត់៖ គោលដៅជានិច្ចជាឃុំសង្កាត់ខ្លួនផ្ទាល់ (+ក្រុងស្រុកខ្លួន) — មិនអាចប្តូរបានឡើយ ====
function submitComment(currentUsername, targetDistrict, targetCommune, message, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  message = String(message || "").trim();
  if (!message) return { success: false, message: "សូមសរសេរខ្លឹមសារមតិយោបល់!" };
  if (message.length > 2000) return { success: false, message: "សារវែងពេក (មិនលើសពី ២០០០ តួអក្សរ)!" };
  targetDistrict = String(targetDistrict || "").trim();
  targetCommune = String(targetCommune || "").trim();
  // ==== FIX (សុវត្ថិភាព — Formula/CSV Injection)៖ សារនេះឆ្លងកាត់អ្នកប្រើប្រាស់គ្រប់តួនាទី ហើយត្រូវបានសរសេរផ្ទាល់ចូល
  // ក្រឡា Google Sheet ខាងក្រោម — ការពារកុំឲ្យតម្លៃដែលចាប់ផ្តើមដោយ =/+/-/@ ដំណើរការជារូបមន្តស្វ័យប្រវត្តិ ពេល
  // SuperAdmin/Admin/PEC21 បើកមើល Google Sheet ដោយផ្ទាល់ (មើល sanitizeForSheetCell_ ក្នុង Utils.gs) — ចេតនាមិន
  // ប៉ះពាល់ "message" ដើម ដែលនៅតែប្រើសម្រាប់ Telegram/សារឆ្លើយតបទៅ Client ខាងក្រោម (មិនចាំបាច់ការពារទីនោះទេ) ====
  var messageForSheet = sanitizeForSheetCell_(message);

  // ==== FIX (Defense in depth)៖ ចាក់សោគោលដៅ ស្របតាមកម្រិតសិទ្ធិ ២/៣ ដោយខ្លួន Server ផ្ទាល់ (មិនទុកចិត្តលើ Client) ====
  if (isCommuneRole_(role)) {
    targetDistrict = getUserDistrict_(currentUsername) || "";
    targetCommune = getUserCommune_(currentUsername) || "";
  } else if (isDistrictRole_(role)) {
    targetDistrict = role;
    // ==== FIX (Fix117)៖ ការផ្ទៀងផ្ទាត់ targetCommune សម្រាប់តួនាទីនេះ ត្រូវបានផ្លាស់ទីទៅធ្វើនៅផ្នែកផ្ទៀងផ្ទាត់រួម
    // ខាងក្រោម (ជៀសវាងសរសេរតក្កវិជ្ជាដដែលពីរដង) ដែលឥឡូវលើកលែងតម្លៃពិសេស COMMENT_TARGET_ROLE_SENTINELS_ ផងដែរ ====
  }
  // (គណនីមើលបានគ្រប់ក្រុងស្រុក៖ ផ្ទៀងផ្ទាត់ធម្មតាខាងក្រោម ដូចមុន — ជ្រើសរើសបានគ្រប់ក្រុងស្រុក+ឃុំសង្កាត់ណាមួយក៏បាន)

  // ==== ជម្រើស "គ្រប់ក្រុងស្រុកទាំងអស់" (SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះ) — ដើម្បីកុំឲ្យត្រូវជ្រើសរើស ហើយផ្ញើម្តងមួយៗ
  // ម្តងមួយស្រុក ១០ដង SuperAdmin អាចផ្ញើសារតែម្តង ហើយប្រព័ន្ធបង្កើតជួរដេកមួយដាច់ដោយឡែកសម្រាប់ស្រុកនីមួយៗ (ដូចផ្ញើ
  // ដោយដៃម្តងមួយៗដដែល) ដើម្បីរក្សាតក្កវិជ្ជាឆ្លើយតប/ដោះស្រាយ/មើលឃើញ ដដែលទាំងស្រុង (ស្រុកនីមួយៗគ្រប់គ្រងឯករាជ្យពីគ្នា
  // ដូចមុន — បើស្រុកមួយឆ្លើយតបឬដោះស្រាយរួច មិនប៉ះពាល់ស្រុកឯទៀតទេ) — ករណីនេះមិនអាចជ្រើសរើសឃុំសង្កាត់ជាក់លាក់បានទេ
  // (ព្រោះជាការផ្ញើទៅកម្រិតក្រុងស្រុកទាំង១០ ព្រមគ្នា) ====
  if (targetDistrict === COMMENT_TARGET_ALL_SENTINEL_) {
    // ==== FIX៖ PEC21 ឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង ====
    if (role !== ROLE_SUPERADMIN && role !== ROLE_ADMIN && role !== ROLE_PEC21) {
      return { success: false, message: "មានតែ SuperAdmin/Admin/PEC21 ទេ ដែលអាចផ្ញើទៅគ្រប់ក្រុងស្រុកទាំងអស់បាន!" };
    }
    var ssAll = getSS_();
    var sheetAll = ensureCommentsSheet_(ssAll);
    var fullNameAll = getUserFullName_(currentUsername) || currentUsername;
    var nowAll = formatNow_();
    var rowsAll = DISTRICT_LIST.map(function(d) {
      return [newId_(), nowAll, currentUsername, fullNameAll, role, d, messageForSheet, COMMENT_STATUS_NEW_, "", "", "", ""];
    });
    sheetAll.getRange(sheetAll.getLastRow() + 1, 1, rowsAll.length, 12).setValues(rowsAll);
    // ==== FIX (Fix126, "ការផ្ញើសារមក SuperAdmin លុះត្រាតែគណនី...Logout ទើបបានផ្ញើ...សូមឲ្យផ្ញើភ្លាមៗ")៖ notifyTelegram_()
    // (ជួរដេក — queueTelegramNotify_/flushTelegramQueue_) ពឹងផ្អែកលើ Time-based Trigger "after(2000)" ដែល Google
    // Apps Script មិនធានាថាដំណើរការក្នុងរយៈពេល ២វិនាទីជាក់ស្តែងទេ (អាចយឺតដល់ជារាប់នាទី ឬយូរជាងនេះថែមទៀត — ដូចគ្នា
    // នឹងបញ្ហាដែលធ្លាប់រកឃើញ ហើយបានកែសម្រាប់ Login/Logout/Key Number ទៅហើយ Fix118/Fix121)។ សម្រាប់មតិយោបល់ដែល
    // ជាការឆ្លើយឆ្លងគ្នាភ្លាមៗរវាងគណនី ការយឺតបែបនេះមិនអាចទទួលយកបានទេ — ប្តូរទៅ notifyTelegramImmediate_() (ភ្លាមៗ
    // Synchronous ជាមួយ Retry Logic ស្រាប់ — Fix119) ជំនួសវិញ ដោយដឹងថាការសរសេរមតិយោបល់នឹងយឺតបន្តិច (បន្ថែមរយៈពេល
    // ហៅ Telegram API — ប្រហែល ១ វិនាទីកន្លះ ក្នុងករណីអាក្រក់បំផុត) ====
    try {
      notifyTelegramImmediate_(
        "🗣️ មតិយោបល់ថ្មី (ផ្ញើទៅគ្រប់ក្រុងស្រុកទាំង " + DISTRICT_LIST.length + ")!\n👤 អ្នកសរសេរ៖ " + fullNameAll + " (" + role + ")" +
        "\n💬 សារ៖ " + message
      );
    } catch (err) {}
    return { success: true, message: "ផ្ញើមតិយោបល់ទៅគ្រប់ក្រុងស្រុកទាំង " + DISTRICT_LIST.length + " ជោគជ័យ!" };
  }

  if (targetDistrict && DISTRICT_LIST.indexOf(targetDistrict) === -1) {
    return { success: false, message: "ស្រុកគោលដៅមិនត្រឹមត្រូវទេ!" };
  }
  // ==== FIX៖ ឃុំសង្កាត់គោលដៅ ត្រូវផ្ទៀងផ្ទាត់ថាមានពិតប្រាកដក្នុងក្រុងស្រុកគោលដៅនោះ (បើមានបញ្ជាក់) — លើកលែងតែជាតម្លៃ
  // ពិសេស COMMENT_TARGET_ROLE_SENTINELS_ (Fix117 — SuperAdmin/Admin/PEC21 ជាស្លាកសម្គាល់ មិនមែនឈ្មោះឃុំសង្កាត់ទេ) ====
  if (targetCommune && COMMENT_TARGET_ROLE_SENTINELS_.indexOf(targetCommune) === -1 &&
      (!targetDistrict || !isValidCommune_(targetDistrict, targetCommune))) {
    return { success: false, message: "ឃុំសង្កាត់គោលដៅមិនត្រឹមត្រូវទេ!" };
  }
  var ss = getSS_();
  var sheet = ensureCommentsSheet_(ss);
  var fullName = getUserFullName_(currentUsername) || currentUsername;
  var id = newId_();
  sheet.appendRow([id, formatNow_(), currentUsername, fullName, role, targetDistrict, messageForSheet, COMMENT_STATUS_NEW_, "", "", "", targetCommune]);
  // ==== FIX (Fix126)៖ notifyTelegram_() → notifyTelegramImmediate_() — ដូចគ្នានឹងខាងលើ (ករណី "__ALL__") ====
  try {
    var targetLabelForNotify_ = (targetDistrict || COMMENT_TARGET_GENERAL_LABEL_) + (targetCommune ? " - " + targetCommune : "");
    notifyTelegramImmediate_(
      "🗣️ មតិយោបល់ថ្មី!\n👤 អ្នកសរសេរ៖ " + fullName + " (" + role + ")\n📍 គោលដៅ៖ " + targetLabelForNotify_ +
      "\n💬 សារ៖ " + message
    );
  } catch (err) {}
  return { success: true, message: "ផ្ញើមតិយោបល់ជោគជ័យ!" };
}

// ==== ត្រឡប់មតិយោបល់ដែលពាក់ព័ន្ធនឹងអ្នកប្រើប្រាស់បច្ចុប្បន្ន (តាមតួនាទី) ====
// SuperAdmin/Admin/PEC21 ឃើញទាំងអស់ ។ គ្រូប្រចាំស្រុក ឃើញមតិយោបល់ដែលមានគោលដៅជាស្រុកខ្លួន (មិនថាមានឃុំសង្កាត់ជាក់លាក់
// ឬអត់) ឬដែលខ្លួនផ្ទាល់សរសេរ ។ គណនីឃុំសង្កាត់ ឃើញតែមតិយោបល់ដែលមានគោលដៅជាឃុំសង្កាត់ខ្លួនផ្ទាល់ (ជាក់លាក់) ឬដែលខ្លួន
// ផ្ទាល់សរសេរប៉ុណ្ណោះ (FIX សំណើថ្មី "គោលដៅ ក្រុងស្រុក/ឃុំសង្កាត់ ២កម្រិត") ។ សង្កេតការណ៍/តួនាទីផ្សេងទៀត ឃើញតែមតិយោបល់
// ដែលខ្លួនផ្ទាល់សរសេរប៉ុណ្ណោះ (ដើម្បីតាមដានការឆ្លើយតប)។
function getComments(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var district = getUserDistrict_(currentUsername);
  var commune = isCommuneRole_(role) ? (getUserCommune_(currentUsername) || "") : "";
  // ==== FIX៖ PEC21 ឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង (មើលឃើញ/ឆ្លើយតប/លុប បានគ្រប់មតិយោបល់) ====
  var isAdminRole = (role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21);
  var ss = getSS_();
  var sheet = ensureCommentsSheet_(ss);
  var lastRow = sheet.getLastRow();
  var list = [];
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
    data.forEach(function(r) {
      var targetDistrict = r[5];
      var targetCommune = r[11] || "";
      var fromUsername = r[2];
      var visible = isAdminRole ||
          fromUsername === currentUsername ||
          (isCommuneRole_(role)
            ? (targetDistrict === district && !!targetCommune && targetCommune === commune)
            : (isDistrictRole_(role) && targetDistrict === district));
      if (!visible) return;
      var canRespond = canRespondToComment_(currentUsername, targetDistrict, targetCommune);
      list.push({
        id: r[0], timestamp: r[1], fromUsername: fromUsername, fromFullName: r[3], fromRole: r[4],
        targetDistrict: targetDistrict, targetCommune: targetCommune, message: r[6], status: r[7],
        replyMessage: r[8], repliedBy: r[9], repliedAt: r[10],
        canRespond: canRespond, canDelete: isAdminRole // លុបបាន SuperAdmin/Admin ប៉ុណ្ណោះ (មិនថាស្ថានភាពអ្វីទេ)
      });
    });
  }
  list.reverse(); // ថ្មីបំផុតនៅលើគេ
  return { success: true, comments: list, myRole: role, myDistrict: district || "", myCommune: commune || "" };
}

function findCommentRow_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // ជួរដេកពិត (1-indexed + Header)
  }
  return -1;
}

// SuperAdmin/Admin/PEC21 អាចឆ្លើយតប/ដោះស្រាយបានគ្រប់មតិយោបល់ ។ គ្រូប្រចាំស្រុក អាចធ្វើបានចំពោះមតិយោបល់ដែលមានគោលដៅជា
// ស្រុកខ្លួនប៉ុណ្ណោះ (មិនថាមានឃុំសង្កាត់ជាក់លាក់ ឬអត់) ។ គណនីឃុំសង្កាត់ អាចធ្វើបានតែចំពោះមតិយោបល់ដែលមានគោលដៅជា
// ឃុំសង្កាត់ខ្លួនផ្ទាល់ (ជាក់លាក់) ប៉ុណ្ណោះ (FIX សំណើថ្មី "គោលដៅ ក្រុងស្រុក/ឃុំសង្កាត់ ២កម្រិត") ====
// ==== FIX៖ PEC21 ឥឡូវមានសិទ្ធិដូច Admin ទាំងស្រុង ====
function canRespondToComment_(currentUsername, targetDistrict, targetCommune) {
  var role = getUserRole_(currentUsername);
  if (role === ROLE_SUPERADMIN || role === ROLE_ADMIN || role === ROLE_PEC21) return true;
  if (isDistrictRole_(role)) return role === targetDistrict;
  if (isCommuneRole_(role)) {
    var myDistrict = getUserDistrict_(currentUsername) || "";
    var myCommune = getUserCommune_(currentUsername) || "";
    return targetDistrict === myDistrict && !!targetCommune && targetCommune === myCommune;
  }
  return false;
}

// ==== ការពារការប៉ះទង្គិចគ្នា (Race Condition)៖ replyToComment/resolveComment/deleteComment ខាងក្រោមទាំងអស់ ជាទម្រង់
// "រកជួរដេកជាមុន រួចធ្វើសកម្មភាពលើលេខជួរដេកនោះ" (Find-then-act) — បើសំណើពីរកើតឡើងស្ទើរតែក្នុងពេលតែមួយ (ឧ. Admin ២នាក់
// ឆ្លើយតប/លុបមតិយោបល់ក្នុងពេលជិតគ្នា) លេខជួរដេកដែលរកឃើញអាចលែងត្រឹមត្រូវ ទាល់តែសោះ ពេលដល់ពេលសរសេរ/លុបជាក់ស្តែង
// (ជួរដេកផ្សេងទៀតអាចត្រូវបានលុប ដែលរុញលេខជួរដេកទាំងអស់ខាងក្រោមប្តូរទីតាំង) ធ្វើឲ្យសកម្មភាពនោះអនុវត្តខុសទៅលើមតិយោបល់
// ដែលមិនមានចេតនា។ ចាក់សោខ្លីមួយភ្លែត (ដូចគ្នានឹង Pattern saveDistrictDayEntries ក្នុង PeriodSheets.gs) ដើម្បីធានាថា
// មានតែសំណើមួយប៉ុណ្ណោះ ធ្វើ Find-then-act នេះក្នុងពេលតែមួយ ====
function withCommentLock_(fn) {
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(10000); } catch (e) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានសកម្មភាពផ្សេងទៀតលើមតិយោបល់កំពុងដំណើរការ)! សូមព្យាយាមម្តងទៀត។" };
  }
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ==== FIX (Fix124, "ចំពោះសារឆ្លើយតប និងដំណោះស្រាយ មិនបានផ្ញើសារទៅ SuperAdmin")៖ មុននេះ replyToComment()/
// resolveComment() ខាងក្រោម មិនដែលហៅ notifyTelegram_() ទាល់តែសោះ (ខុសពី submitComment() ខាងលើ ដែលហៅរាល់ពេលមាន
// មតិយោបល់ថ្មី) — SuperAdmin/អ្នកតាមដានក្នុង Telegram ឃើញតែមតិយោបល់ដើម ប៉ុន្តែមិនដឹងថាមានការឆ្លើយតប/ដោះស្រាយ
// កើតឡើងឡើយ លុះត្រាតែបើក Google Sheet ដោយផ្ទាល់ ឬចូល Tab(មតិយោបល់) ក្នុងកម្មវិធីមើលឡើងវិញ។ ឥឡូវបន្ថែមការជូនដំណឹង
// ដូចគ្នា (notifyTelegram_ — ជួរដេក ដូច submitComment ដើម មិនប្រញាប់ភ្លាមៗ) ====
function replyToComment(currentUsername, commentId, replyText, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  replyText = String(replyText || "").trim();
  if (!replyText) return { success: false, message: "សូមសរសេរចម្លើយឆ្លើយតប!" };
  if (replyText.length > 2000) return { success: false, message: "ចម្លើយវែងពេក (មិនលើសពី ២០០០ តួអក្សរ)!" };
  return withCommentLock_(function() {
    var ss = getSS_();
    var sheet = ensureCommentsSheet_(ss);
    var row = findCommentRow_(sheet, commentId);
    if (row === -1) return { success: false, message: "រកមិនឃើញមតិយោបល់នេះទេ (អាចត្រូវបានលុបរួច)!" };
    // ==== FIX (Fix124)៖ អាន Row ទាំងមូលម្តងគត់ (ជំនួស Call ដាច់ដោយឡែក ២ដង — targetDistrict/targetCommune) ដើម្បី
    // យកទាំងព័ត៌មានគោលដៅ និងព័ត៌មានមតិយោបល់ដើម (អ្នកសរសេរ/សារ) សម្រាប់សារជូនដំណឹង Telegram ខាងក្រោមតែម្តង ====
    var rowData = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];
    var targetDistrict = rowData[5];
    var targetCommune = rowData[11] || "";
    if (!canRespondToComment_(currentUsername, targetDistrict, targetCommune)) {
      return { success: false, message: "អ្នកគ្មានសិទ្ធិឆ្លើយតបមតិយោបល់នេះទេ!" };
    }
    var fullName = getUserFullName_(currentUsername) || currentUsername;
    sheet.getRange(row, 8, 1, 4).setValues([[COMMENT_STATUS_REPLIED_, sanitizeForSheetCell_(replyText), fullName, formatNow_()]]);
    // ==== FIX (Fix126, "ការផ្ញើសារមក SuperAdmin លុះត្រាតែគណនី...Logout...សូមឲ្យផ្ញើភ្លាមៗ")៖ notifyTelegram_()
    // (ជួរដេក) → notifyTelegramImmediate_() (ភ្លាមៗ) — ដូចគ្នានឹងហេតុផលនៅ submitComment() ខាងលើ ====
    try {
      var role = getUserRole_(currentUsername);
      var targetLabelForNotify_ = (targetDistrict || COMMENT_TARGET_GENERAL_LABEL_) + (targetCommune ? " - " + targetCommune : "");
      notifyTelegramImmediate_(
        "↩️ ឆ្លើយតបមតិយោបល់!\n👤 អ្នកឆ្លើយតប៖ " + fullName + " (" + role + ")\n📍 គោលដៅ៖ " + targetLabelForNotify_ +
        "\n📝 មតិយោបល់ដើម ពី " + rowData[3] + "៖ " + rowData[6] +
        "\n💬 ចម្លើយ៖ " + replyText
      );
    } catch (err) {}
    return { success: true, message: "ឆ្លើយតបជោគជ័យ!" };
  });
}

function resolveComment(currentUsername, commentId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  return withCommentLock_(function() {
    var ss = getSS_();
    var sheet = ensureCommentsSheet_(ss);
    var row = findCommentRow_(sheet, commentId);
    if (row === -1) return { success: false, message: "រកមិនឃើញមតិយោបល់នេះទេ (អាចត្រូវបានលុបរួច)!" };
    var rowData = sheet.getRange(row, 1, 1, 12).getDisplayValues()[0];
    var targetDistrict = rowData[5];
    var targetCommune = rowData[11] || "";
    if (!canRespondToComment_(currentUsername, targetDistrict, targetCommune)) {
      return { success: false, message: "អ្នកគ្មានសិទ្ធិសម្គាល់មតិយោបល់នេះថាដោះស្រាយបានទេ!" };
    }
    sheet.getRange(row, 8).setValue(COMMENT_STATUS_RESOLVED_);
    if (!rowData[9]) {
      var fullName = getUserFullName_(currentUsername) || currentUsername;
      sheet.getRange(row, 10, 1, 2).setValues([[fullName, formatNow_()]]);
    }
    // ==== FIX (Fix126)៖ notifyTelegram_() → notifyTelegramImmediate_() — ដូចគ្នានឹងហេតុផលនៅ submitComment()/
    // replyToComment() ខាងលើ ====
    try {
      var role = getUserRole_(currentUsername);
      var resolverName = getUserFullName_(currentUsername) || currentUsername;
      var targetLabelForNotify_ = (targetDistrict || COMMENT_TARGET_GENERAL_LABEL_) + (targetCommune ? " - " + targetCommune : "");
      notifyTelegramImmediate_(
        "✅ បានដោះស្រាយមតិយោបល់!\n👤 អ្នកសម្គាល់៖ " + resolverName + " (" + role + ")\n📍 គោលដៅ៖ " + targetLabelForNotify_ +
        "\n📝 មតិយោបល់ដើម ពី " + rowData[3] + "៖ " + rowData[6]
      );
    } catch (err) {}
    return { success: true, message: "សម្គាល់ថាដោះស្រាយរួចរាល់!" };
  });
}

// ==== លុបមតិយោបល់ចាស់/មិនពាក់ព័ន្ធ — SuperAdmin/Admin ប៉ុណ្ណោះ (សម្រាប់សម្អាតបញ្ជីឲ្យស្អាត) ====
function deleteComment(currentUsername, commentId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "មានតែ SuperAdmin/Admin ទេ ដែលអាចលុបមតិយោបល់បាន!" };
  return withCommentLock_(function() {
    var ss = getSS_();
    var sheet = ensureCommentsSheet_(ss);
    var row = findCommentRow_(sheet, commentId);
    if (row === -1) return { success: false, message: "រកមិនឃើញមតិយោបល់នេះទេ (អាចត្រូវបានលុបរួច)!" };
    sheet.deleteRow(row);
    return { success: true, message: "លុបមតិយោបល់ជោគជ័យ!" };
  });
}
