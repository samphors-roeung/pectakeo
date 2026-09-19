// ==================== Settings.gs — ការកំណត់ប្រព័ន្ធ, ស្ថាប័ន, ទិន្នន័យភ្ជាប់, ទិន្នន័យបម្រុងទុក ====================

// ==================== ០.៣ ទិន្នន័យភ្ជាប់ (Linked Data) — Sheet តែមួយ សាមញ្ញ ផ្ទុកទិន្នន័យយោងទាំងអស់ ====================
// ==== ចំណាំ៖ ធ្វើឲ្យសាមញ្ញតាមដែលអាចធ្វើបាន (គ្មាន Header ច្រើនជួរ, គ្មាន Merge Cell, គ្មាន Trigger ពន្យារពេល) ====
// ដើម្បីជៀសវាងហានិភ័យណាមួយដែលធ្លាប់ជួបប្រទះ — រចនាសម្ព័ន្ធសាមញ្ញបំផុតតែងតែជាជម្រើសដែលអាចទុកចិត្តបានបំផុត។
function getLinkedData(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  if (!canViewAllDistricts_(currentUsername)) return { success: false, message: "គណនីគ្រូប្រចាំស្រុក មិនអាចចូលមើលទិន្នន័យនេះទេ!" };
  try {
    var ss = getSS_();
    var sheet = ensureLinkedDataSheet_(ss);
    var lastRow = sheet.getLastRow();
    var list = [];
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, LINKED_DATA_TOTAL_COLS).getValues();
      data.forEach(function(r) {
        var d = String(r[0] || "").trim(), c = String(r[1] || "").trim();
        if (!d || !c) return;
        list.push({
          district: d, commune: c,
          estimate: Number(r[2]) || 0, newStation: Number(r[3]) || 0, totalStation: Number(r[4]) || 0,
          total2025: Number(r[5]) || 0, female2025: Number(r[6]) || 0, code: r[7] || "",
          biometricTotal: Number(r[8]) || 0, biometricFemale: Number(r[9]) || 0,
          provinceEstimate: Number(r[10]) || 0
        });
      });
    }
    list.sort(districtCommuneComparator_);
    if (!canViewAllDistricts_(currentUsername)) {
      var myDistrict = getUserDistrict_(currentUsername);
      list = list.filter(function(r) { return r.district === myDistrict; });
    }
    return { success: true, list: list, canEdit: isAdmin_(currentUsername) };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

function saveLinkedData(currentUsername, rows, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "មានតែអ្នកគ្រប់គ្រង (Admin) ប៉ុណ្ណោះ ដែលអាចកែទិន្នន័យនេះបាន!" };
  try {
    var ss = getSS_();
    var sheet = ensureLinkedDataSheet_(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, message: "រក្សាទុកជោគជ័យ!" };
    var data = sheet.getRange(2, 1, lastRow - 1, LINKED_DATA_TOTAL_COLS).getValues();
    var rowIndexMap = {};
    data.forEach(function(r, idx) { rowIndexMap[String(r[0]).trim() + "|" + String(r[1]).trim()] = idx; });
    (rows || []).forEach(function(item) {
      var key = String(item.district || "").trim() + "|" + String(item.commune || "").trim();
      var idx = rowIndexMap[key];
      if (idx === undefined) return;
      data[idx][2] = Number(item.estimate) || 0;
      data[idx][3] = Number(item.newStation) || 0;
      data[idx][4] = Number(item.totalStation) || 0;
      data[idx][5] = Number(item.total2025) || 0;
      data[idx][6] = Number(item.female2025) || 0;
      data[idx][8] = Number(item.biometricTotal) || 0;
      data[idx][9] = Number(item.biometricFemale) || 0;
      data[idx][10] = Number(item.provinceEstimate) || 0;
    });
    sheet.getRange(2, 1, data.length, LINKED_DATA_TOTAL_COLS).setValues(data);
    clearLookupCache_();
    return { success: true, message: "រក្សាទុកទិន្នន័យភ្ជាប់ជោគជ័យ!" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

function getBaseline2025Map_() {
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY_BASELINE_MAP_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var ss = getSS_();
  var sheet = ensureLinkedDataSheet_(ss);
  var map = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, LINKED_DATA_TOTAL_COLS).getValues();
    data.forEach(function(r) {
      var d = String(r[0] || "").trim(), c = String(r[1] || "").trim();
      if (!d || !c) return;
      map[d + "|" + c] = { total: Number(r[5]) || 0, female: Number(r[6]) || 0 };
    });
  }

  try { CacheService.getScriptCache().put(CACHE_KEY_BASELINE_MAP_, JSON.stringify(map), CACHE_TTL_); } catch (err) {}
  return map;
}

// ==================== ០.៣ "ទិន្នន័យយោង" ត្រូវប្តូរខ្លួនស្វ័យប្រវត្តិ តាមកាលបរិច្ឆេទ ====================
// ថ្ងៃចាប់ផ្តើមដំណើរការ (ថ្ងៃដំបូងគេ) → ប្រើ "ទិន្នន័យ២០២៥" ជាក់ស្តែង (ពី Sheet "ទិន្នន័យភ្ជាប់")។
// ថ្ងៃទី២ តទៅ (ក្រោយថ្ងៃចាប់ផ្តើម) → ប្តូរឈ្មោះទៅជា "យោងចំនួនក្នុងបញ្ជីសរុបពីមុន" ដោយស្វ័យប្រវត្តិ ព្រមទាំង
// តម្លៃខ្លួនឯងក៏ត្រូវប្តូរទៅជា "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦" របស់ Period Sheet ចុងក្រោយបំផុត មុនកាលបរិច្ឆេទដែលកំពុងមើល
// (មិនមែនប្រើតម្លៃ ២០២៥ ថេរជានិច្ចទៀតទេ) — បន្តបែបនេះរហូតដល់ចប់ប្រតិបត្តិការទាំងមូល។
var REF_LABEL_DEFAULT_ = "យោងចំនួនក្នុងបញ្ជីសរុបពីមុន";

function isFirstOperationDay_(dateStr) {
  var opStart = getOperationStartDate_();
  var d = normalizeDateStr_(dateStr);
  return !opStart || d <= opStart;
}

// ត្រឡប់ { map: {"ស្រុក|ឃុំ": {total, female}}, isFirstDay: true/false, label: "..." } សម្រាប់កាលបរិច្ឆេទដែលបានផ្តល់ឲ្យ
function getRollingReferenceMap_(dateStr, districtFilter) {
  var d = normalizeDateStr_(dateStr);
  if (isFirstOperationDay_(d)) {
    return { map: getBaseline2025Map_(), isFirstDay: true, label: null }; // null = ប្រើ Label លំនាំដើម (ទិន្នន័យ២០២៥)
  }

  var upd26TCol0 = periodFieldColIndex_("update2026_total") - 1;
  var upd26FCol0 = periodFieldColIndex_("update2026_female") - 1;
  var communeIdx0 = DAILY_META_PREFIX.length - 1;
  var totalCols = periodSheetTotalCols_();
  var map = {};
  // ==== ល្បឿន៖ បើហៅសម្រាប់ស្រុកជាក់លាក់តែមួយ (ភាគច្រើនកន្លែងហៅ) ស្កេនតែស្រុកនោះប៉ុណ្ណោះ ជំនួសស្កេនគ្រប់ស្រុកទាំង១០ ====
  var districtsToScan = districtFilter ? [districtFilter] : DISTRICT_LIST;

  districtsToScan.forEach(function(district) {
    // ==== ល្បឿន៖ ប្រើ Map កាលបរិច្ឆេទ Tab ដែលបាន Cache ទុក (getPeriodSheetDateMap_) ជំនួសការអាន sheet.getRange(4,2).getValue()
    // ម្តងមួយៗ សម្រាប់ Tab ជារាប់សិប (Round-trip ទៅ Google Sheets ១ដងក្នុងមួយ Tab — យឺតណាស់ បើហៅញឹកញាប់) ====
    var dateMap = getPeriodSheetDateMap_(district);
    var bestSheet = null, bestDate = null;
    listDistrictPeriodSheets_(district).forEach(function(sheet) {
      var sheetDate = dateMap[sheet.getName()];
      if (sheetDate && sheetDate < d && (!bestDate || sheetDate > bestDate)) {
        bestDate = sheetDate; bestSheet = sheet;
      }
    });
    if (!bestSheet) return; // គ្មាន Period Sheet មុនកាលបរិច្ឆេទនេះទេ — ទុកទទេ (0)
    var lastRow2 = bestSheet.getLastRow();
    if (lastRow2 < 4) return;
    var data = bestSheet.getRange(4, 1, lastRow2 - 3, totalCols).getValues();
    data.forEach(function(row) {
      // ==== FIX (រាយការណ៍ថ្មី "ចំនួនឃុំ/សង្កាត់មិនត្រឹមត្រូវ")៖ ជួរដេក "សរុប" ស្វ័យប្រវត្តិ (writePeriodSheetTotalRow_
      // ក្នុង PeriodSheets.gs) មាន ID ទទេដោយចេតនា ប៉ុន្តែជួរឈរ "ឃុំ/សង្កាត់" របស់វា ("សរុបទាំងអស់") មិនទទេទេ — ការត្រួត
      // ពិនិត្យ "if (!commune) return;" ខាងក្រោម ក្រឡែងតែជួរដេកឈ្មោះឃុំទទេ មិនចាប់ជួរដេកនេះទេ ធ្វើឲ្យ Key ក្លែងក្លាយ
      // "ស្រុក|សរុបទាំងអស់" ចូលទៅក្នុង Map នេះ ដោយចៃដន្យ (ទោះមិនដែលមានកូដណាមួយអាន Key នេះមកប្រើផ្ទាល់ក៏ដោយ — ត្រូវ
      // ជៀសវាងជាមុន) ====
      if (!row[0]) return; // ជួរដេក "សរុប" (ID ទទេ) — មិនមែនឃុំ/សង្កាត់ពិត
      var commune = row[communeIdx0];
      if (!commune) return;
      map[district + "|" + commune] = {
        total: Number(row[upd26TCol0]) || 0,
        female: Number(row[upd26FCol0]) || 0
      };
    });
  });
  return { map: map, isFirstDay: false, label: REF_LABEL_DEFAULT_ };
}

// ---- Sheet(ការកំណត់ប្រព័ន្ធ) — Key/Value សាមញ្ញ (បច្ចុប្បន្នប្រើសម្រាប់ "ថ្ងៃចាប់ផ្តើមដំណើរការ") ----
function ensureSettingsSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SETTINGS);
    sheet.appendRow(["ការកំណត់", "តម្លៃ"]);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#111827").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.appendRow([SETTINGS_KEY_START_DATE, ""]);
  }
  return sheet;
}
// ==== ល្បឿន៖ ថត Cache ក្នុងការប្រតិបត្តិតែមួយ (Execution) ព្រោះ getSetting_ (ជាពិសេស getOperationStartDate_
// តាមរយៈ isFirstOperationDay_) ត្រូវបានហៅម្តងៗលើ Period Sheet នីមួយៗ គុណនឹងស្រុកទាំង១០ ក្នុង protectPastPeriodSheets_
// (ហៅរាល់ដងបង្កើត Sheet ថ្ងៃថ្មី) — មុននេះ រាល់ការហៅ សុទ្ធតែអាន Sheet(ការកំណត់ប្រព័ន្ធ) ថ្មីម្តងៗ ខណៈតម្លៃដដែលមិនប្តូរ
// ក្នុងការប្រតិបត្តិតែមួយឡើយ ====
var SETTING_CACHE_ = {};
function getSetting_(key) {
  if (Object.prototype.hasOwnProperty.call(SETTING_CACHE_, key)) return SETTING_CACHE_[key];
  var ss = getSS_();
  var sheet = ensureSettingsSheet_(ss);
  var lastRow = sheet.getLastRow();
  var value = "";
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    for (var i = 0; i < data.length; i++) { if (data[i][0] === key) { value = data[i][1]; break; } }
  }
  SETTING_CACHE_[key] = value;
  return value;
}
function setSetting_(key, value) {
  var ss = getSS_();
  var sheet = ensureSettingsSheet_(ss);
  var lastRow = sheet.getLastRow();
  var data = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues() : [];
  var written = false;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 2, 2).setValue(value); written = true; break; }
  }
  if (!written) sheet.appendRow([key, value]);
  SETTING_CACHE_[key] = String(value); // ធានាថាការហៅ getSetting_ បន្ទាប់ ក្នុងការប្រតិបត្តិតែមួយនេះ ឃើញតម្លៃថ្មីភ្លាមៗ
}
function getOperationStartDate_() {
  return normalizeDateStr_(getSetting_(SETTINGS_KEY_START_DATE));
}
function getOperationEndDate_() {
  return normalizeDateStr_(getSetting_(SETTINGS_KEY_END_DATE));
}
// ត្រឡប់ការកំណត់ប្រព័ន្ធ (ថ្ងៃចាប់ផ្តើម/ចុងបញ្ចប់ដំណើរការ + សិទ្ធិកែ) សម្រាប់ភ្ជាប់ទៅ Client
function getSystemSettings(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  return {
    success: true, operationStartDate: getOperationStartDate_(), operationEndDate: getOperationEndDate_(),
    isAdmin: isAdmin_(currentUsername), todayStr: formatNow_().slice(0, 10)
  };
}
// ==== ត្រឡប់ព័ត៌មាន "អំពីកម្មវិធី" — មើលបានសម្រាប់អ្នកប្រើប្រាស់គ្រប់រូបដែល Login រួច (មិនចាំបាច់ជា Admin ទេ)
// ត្រឡប់តែវាល aboutXxx ប៉ុណ្ណោះ (មិនមែនការកំណត់រសើបទាំងអស់ដូច getSystemSettingsAdmin ទេ) ====
function getAboutInfo(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var s = getSystemSettings_();
  return {
    // ==== "អំពីកម្មវិធី" កែប្រែបានតែ SuperAdmin ប៉ុណ្ណោះ (មិនរួមបញ្ចូល Admin ធម្មតា ឬ PEC21 ទេ) — ខុសពី
    // ការកំណត់ប្រព័ន្ធផ្សេងទៀត ដែល Admin នៅតែកែបាន ====
    success: true, isAdmin: isAdmin_(currentUsername), isSuperAdmin: isSuperAdmin_(currentUsername),
    appTitle: s.aboutAppTitle || "", appVersion: s.aboutAppVersion || "",
    creatorName: s.aboutCreatorName || "", creatorAddress: s.aboutCreatorAddress || "",
    creatorTelegram: s.aboutCreatorTelegram || "", creatorPhone: s.aboutCreatorPhone || ""
  };
}
// ==== រក្សាទុកព័ត៌មាន "អំពីកម្មវិធី" — មានតែ SuperAdmin ប៉ុណ្ណោះទើបកែប្រែបាន (Admin/PEC21/គណនីផ្សេងទៀត មិនអាចកែបានទេ
// ទោះបីជា Admin នៅតែកែការកំណត់ប្រព័ន្ធផ្សេងទៀត (Tab កំណត់ប្រព័ន្ធ) បាន តាមរយៈ saveSystemSettings ធម្មតាក៏ដោយ) ====
function saveAboutInfo(currentUsername, patch, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isSuperAdmin_(currentUsername)) return { success: false, message: "មានតែ SuperAdmin ប៉ុណ្ណោះ ដែលអាចកែប្រែព័ត៌មាន \"អំពីកម្មវិធី\" បានទេ!" };
  var props = PropertiesService.getScriptProperties();
  var aboutKeys_ = ["aboutAppTitle", "aboutAppVersion", "aboutCreatorName", "aboutCreatorAddress", "aboutCreatorTelegram", "aboutCreatorPhone"];
  Object.keys(patch || {}).forEach(function(k) {
    if (aboutKeys_.indexOf(k) === -1) return;
    props.setProperty(SETTINGS_PROP_PREFIX_ + k, String(patch[k] || ""));
  });
  return { success: true, message: "រក្សាទុកជោគជ័យ!" };
}

// កំណត់ "ថ្ងៃចាប់ផ្តើមដំណើរការ" — មានតែ Admin/SuperAdmin ប៉ុណ្ណោះទើបកំណត់បាន
function setOperationStartDate(currentUsername, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "មានតែអ្នកគ្រប់គ្រង (Admin) ប៉ុណ្ណោះ ដែលអាចកំណត់ថ្ងៃចាប់ផ្តើមដំណើរការបានទេ!" };
  setSetting_(SETTINGS_KEY_START_DATE, normalizeDateStr_(dateStr));
  return { success: true };
}
// កំណត់ "ថ្ងៃចុងបញ្ចប់ដំណើរការ" — មានតែ Admin/SuperAdmin ប៉ុណ្ណោះទើបកំណត់បាន
function setOperationEndDate(currentUsername, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "មានតែអ្នកគ្រប់គ្រង (Admin) ប៉ុណ្ណោះ ដែលអាចកំណត់ថ្ងៃចុងបញ្ចប់ដំណើរការបានទេ!" };
  setSetting_(SETTINGS_KEY_END_DATE, normalizeDateStr_(dateStr));
  return { success: true };
}

// ---- Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) — មួយជួរ (Row) = ១ ការបញ្ចូលប្រចាំថ្ងៃ តាមឃុំ/សង្កាត់ ----
var SETTINGS_PROP_PREFIX_ = "sysSetting_";
var SETTINGS_KEYS_ = ["orgName", "orgSlogan", "logoUrl", "tgBotToken", "tgChatId", "telegramLink", "youtubeLink", "enabledLangs", "showDuplicateFields", "showEstimate", "showNewStation", "showTotalStation", "showCommuneCode", "showProvinceEstimate", "showAggPercent1", "showAggPercent2", "showAggPercent3", "showAggPercent4",
  // ==== ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) — ប្រសិនបើមិនទាន់កំណត់ផ្ទាល់
  // ("") ទេ នឹងប្រើតម្លៃដដែលនឹង Group ខាងលើ (ទម្រង់បញ្ចូលទិន្នន័យ) ជាលំនាំដើម — មើល resolveGroupDisplaySettings_() ====
  "showEstimateDaily", "showProvinceEstimateDaily", "showNewStationDaily", "showTotalStationDaily", "showCommuneCodeDaily", "showDuplicateFieldsDaily",
  // ==== ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ Tab(របាយការណ៍សរុបតាមឃុំសង្កាត់/របាយការណ៍សរុបខេត្ត) — ចែករំលែក
  // ការកំណត់តែមួយរួមគ្នារវាងទាំង២ Tab នេះ ហើយក៏ត្រូវបានប្រើដោយ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទបូកយោង (លេខ៣.៤ —
  // cumCommune/cumProvince) + Excel Export ដូចគ្នាដែរ (ព្រោះទាំងអស់នេះមកពី getSummaries()/computeSummaries_()
  // ដូចគ្នា) — ជួរឈរ "ភាគរយ" ១-៤ ចែករំលែក Key ចាស់ (showAggPercent1-4) ព្រោះជួរឈរនេះមានន័យសម្រាប់តែរបាយការណ៍
  // ក្រុមនេះប៉ុណ្ណោះ (មិនចាំបាច់ Key ថ្មីទេ) ====
  "showEstimateSum", "showProvinceEstimateSum", "showNewStationSum", "showTotalStationSum", "showCommuneCodeSum", "showDuplicateFieldsSum",
  // ==== FIX (សំណើថ្មី "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)/ភាគរយ")៖ ជួរឈរបន្ថែមថ្មី ២ (តម្លៃយោង "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)"
  // ពី Tab "ទិន្នន័យភ្ជាប់" + ភាគរយធៀបនឹងចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រពិតប្រាកដ) ដាក់ជាប់បន្ទាប់ពីជួរឈរ ឃុំ/សង្កាត់
  // (ឬ ក្រុង/ស្រុក) ភ្លាមៗ — សម្រាប់ Tab(របាយការណ៍សរុបតាមឃុំសង្កាត់/សរុបខេត្ត) + "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទ
  // បូកយោង (លេខ៣.៤) ប៉ុណ្ណោះ (ចែករំលែក Key តែមួយ ដូចគ្នានឹងធីច "ភាគរយ" ១-៤ — មិនចាំបាច់ Suffix ដាច់ដោយឡែកទេ ព្រោះ
  // មានន័យសម្រាប់តែក្រុមនេះជាក់ស្តែង) — លំនាំដើមបង្ហាញ (ដូច showAggPercent1-4) ព្រោះជាមុខងារដែល Admin ស្នើសុំផ្ទាល់ ====
  "showBioTotalSum",
  // ==== FIX (សំណើថ្មី "១")៖ ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទប្រចាំថ្ងៃ
  // (លេខ១.២ — dailyCommune/dailyProvince) ប៉ុណ្ណោះ — មុននេះរបាយការណ៍ទាំង២នេះ ចែករំលែកការកំណត់ជាមួយក្រុម "Sum"
  // ខាងលើ (ព្រមទាំងរបាយការណ៍ ៣.៤) ដោយអចេតនា ។ បើមិនទាន់កំណត់ផ្ទាល់ ("") ទេ Fallback ទៅតាមអ្វីដែលក្រុម "Sum"
  // កំពុងបង្ហាញជាក់ស្តែងជាមុនសិន (មិនប៉ះពាល់ការបង្ហាញបច្ចុប្បន្នភ្លាមៗពេលដំឡើងកំណែនេះ) — មើល
  // resolvePrintDailyDisplaySettings_() ====
  "showEstimatePrintDaily", "showProvinceEstimatePrintDaily", "showNewStationPrintDaily", "showTotalStationPrintDaily", "showCommuneCodePrintDaily", "showDuplicateFieldsPrintDaily",
  // ==== FIX (សំណើថ្មី "៣.៤")៖ ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទ
  // ធ្វើជីវមាត្របូកយោង (លេខ៥.៦ — bioCommune/bioProvince) ប៉ុណ្ណោះ — ជួរឈរ "ស្ថិតិប៉ាន់ស្មាន" ក្នុងរបាយការណ៍នេះ
  // ពីមុនមកបង្ហាញជានិច្ចដោយ Hardcode (គ្មានកន្លែងកំណត់ទាល់តែសោះ) ប្រភពតែ "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" ។ ឥឡូវអាចជ្រើសរើស
  // "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" ឬ "ស្ថិតិប៉ាន់ស្មានខេត្ត" មួយណាក៏បាន ឬទាំងពីរក៏បាន — ជួរឈរ "ភាគរយ" ដើរតាមការជ្រើសរើសនេះ
  // ដោយស្វ័យប្រវត្តិ (មិនមានធីចដាច់ដោយឡែកទេ) — មើល resolveBioDisplaySettings_() ====
  "showEstimateBio", "showProvinceEstimateBio", "showCommuneCodeBio",
  "labelBaseline2025", "labelUpdate2026", "footerChairmanTitle", "footerChairmanName", "footerPreparerTitle", "footerPreparerName", "backupFolderId",
  "aboutAppTitle", "aboutAppVersion", "aboutCreatorName", "aboutCreatorAddress", "aboutCreatorTelegram", "aboutCreatorPhone",
  "reportLogoUrl", "reportCommitteeName", "reportSecretariatName",
  // ==== តំណភ្ជាប់ Google Drive/Google Sheet មេ ដែល Admin បិទភ្ជាប់ (Copy-Paste) ដោយផ្ទាល់ — មើល getMasterLinks()
  // ក្នុង PeriodSheets.gs (ប្រើតំណភ្ជាប់ទាំងនេះជានិច្ច ប្រសិនបើមាន ជំនួសការគណនា/ស្វែងរកដោយស្វ័យប្រវត្តិ) ====
  "masterDriveLink", "masterSheetLink",
  // ==== FIX (Fix127, "ចង់ឲ្យមានជម្រើសបើក/បិទប្រើប្រាស់គណនីឃុំសង្កាត់")៖ Key ថ្មី — ត្រូវជា Key ចុងក្រោយ (បន្ថែម
  // ដោយសុវត្ថិភាព មិនប៉ះពាល់លំដាប់ Key ចាស់ៗនានា) — មើល isCommuneAccountsEnabled_() ខាងក្រោម ====
  "enableCommuneAccounts",
  // ==== FIX (Fix128, "ចង់ឲ្យមានជម្រើសបន្តទៀត...គណនីឃុំសង្កាត់ដែលបង្កើតហើយ អាចជ្រើសរើសប្រើបាន ឬមិនអាចប្រើបានបន្តទៀត")៖
  // Key ថ្មីទី២ — ខុសពី "enableCommuneAccounts" (Fix127, ត្រួតពិនិត្យតែការ "បង្កើត" គណនីថ្មី) — Key នេះត្រួតពិនិត្យថាតើ
  // គណនីឃុំសង្កាត់ដែល "មានស្រាប់" រួចហើយ អាចបន្ត Login/ប្រើប្រាស់បានដែរឬអត់ — មើល isCommuneAccountsUsageEnabled_()
  // ខាងក្រោម ====
  "enableCommuneAccountsUsage",
  // ==== FIX (Fix129, "គណនីថ្មីទាំង៥ ត្រូវមាន Option បើក/បិទ ការបង្កើតគណនីនីមួយៗ")៖ Key ថ្មីទាំង៥ — មួយ Key
  // ក្នុងមួយតួនាទី (មើល OBSERVER_TIER_ROLES_ ក្នុង Utils.gs, isObserverTierRoleEnabled_ ខាងក្រោម) — ខុសពី
  // enableCommuneAccounts/enableCommuneAccountsUsage ខាងលើ (លំនាំដើម "អនុញ្ញាត" ព្រោះជាមុខងារមានស្រាប់រួចហើយ)៖
  // Key ទាំង៥នេះលំនាំដើម "បិទ" (ត្រូវប្រើ === 'true' មិនមែន !== 'false' ទេ) ព្រោះជាតួនាទីថ្មីទាំងស្រុង គ្មានការប្រើ
  // ប្រាស់ពីមុនត្រូវការពារ — SuperAdmin/Admin ត្រូវចូល "ការកំណត់ប្រព័ន្ធ" បើកជាក់លាក់ម្តងមួយៗ តាមតួនាទីដែលត្រូវការ ====
  "enableRoleNecOfficer", "enableRoleProvinceAdmin", "enableRoleDistrictAdmin", "enableRoleArmedForces", "enableRolePoliticalParty",
  // ==== FIX (សំណើថ្មី "បិទកន្លែងស្នើសុំបង្កើតគណនីថ្មីទាំងអស់")៖ Key ថ្មី — ត្រូវជា Key ចុងក្រោយ (បន្ថែមដោយសុវត្ថិភាព
  // មិនប៉ះពាល់លំដាប់ Key ចាស់ៗនានា) — កុងតាក់ធំតែមួយ គ្រប់គ្រង "Panel ស្នើសុំគណនីថ្មី" ទាំងមូលនៅទំព័រ Login សាធារណៈ
  // (រួមទាំងតំណភ្ជាប់ "ស្នើសុំគណនីថ្មីទីនេះ" ខ្លួនឯង) — ខុសពី enableCommuneAccounts/enableRoleXxx ខាងលើ (ដែលគ្រប់គ្រង
  // ត្រឹមតែប្រភេទគណនីជាក់លាក់ម្នាក់ៗ)៖ កុងតាក់នេះបិទផ្លូវ "ស្នើសុំខ្លួនឯង" (requestSignup) ទាំងអស់ម្តងតែម្តង មិនថា
  // ប្រភេទគណនីណា (ស្រុក/ឃុំសង្កាត់/សង្កេតការណ៍/តួនាទីថ្មីទាំង៥) ដើម្បីកុំឲ្យអ្នកដទៃប្រើ Link ទំព័រ Login ចូលមកបង្កើត
  // គណនីខ្លួនឯងបានទៀត — SuperAdmin/Admin/PEC21 នៅតែបង្កើតគណនីឲ្យអ្នកដទៃដោយផ្ទាល់បាន (addUserAccount) ដដែល
  // (មិនរងផលប៉ះពាល់ទេ ព្រោះជាមុខងារខុសគ្នាទាំងស្រុង) — លំនាំដើមអនុញ្ញាត (true) ព្រោះជាមុខងារមានស្រាប់រួចហើយ (ដូចគ្នា
  // នឹង enableCommuneAccounts) — មើល isSelfSignupEnabled_() ខាងក្រោម ====
  "enableSelfSignup"];

// ==== បម្លែង Logo URL ចាស់ (ទម្រង់ uc?export=view ឬ តំណភ្ជាប់ចែករំលែក .../file/d/ID/view) ទៅជាទម្រង់
// lh3.googleusercontent.com/d/{id} ដែលអាចទុកចិត្តបាននៅពេលបង្ហាញផ្ទាល់ក្នុង <img> ។ ធ្វើដូចនេះទាំង
// getSystemSettings_ (មិនមែនត្រឹមតែពេលបញ្ចូល Logo ថ្មីទេ) ដើម្បីឲ្យ Logo ចាស់ៗដែលធ្លាប់រក្សាទុកជាទម្រង់
// ខូចរួចហើយ ជួសជុលដោយស្វ័យប្រវត្តិផងដែរ ដោយមិនបាច់ឲ្យអ្នកគ្រប់គ្រងបញ្ចូល Logo ម្តងទៀតទេ។
function normalizeDriveImageUrl_(url) {
  url = String(url || "").trim();
  if (!url) return url;
  if (url.indexOf("lh3.googleusercontent.com") !== -1) return url;
  var m = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return "https://lh3.googleusercontent.com/d/" + m[1];
  return url;
}

function getSystemSettings_() {
  var props = PropertiesService.getScriptProperties();
  // ==== FIX (Fix120, "ល្បឿនចាប់ទិន្នន័យផ្សេងៗ")៖ SETTINGS_KEYS_ មានច្រើនជាង ៥០ Key — មុននេះហៅ props.getProperty()
  // ដាច់ដោយឡែកម្តងមួយៗ (ជាង ៥០ដងក្នុងការហៅតែម្តង) ។ getSystemSettings_() ត្រូវបានហៅស្ទើរតែរាល់ Server Function
  // សំខាន់ៗ (ទម្រង់បញ្ចូលទិន្នន័យ, ការកំណត់ប្រព័ន្ធ, ទំព័រ Login សាធារណៈ ។ល។) ដូច្នេះការហៅច្រើនដងនេះកកកុញឲ្យយឺត។
  // ឥឡូវហៅ props.getProperties() តែម្តងគត់ (ទាញយកគ្រប់ Key-Value ទាំងអស់ក្នុងការហៅតែម្តង) លឿនជាងច្រើន ====
  var allProps = props.getProperties();
  var out = {};
  SETTINGS_KEYS_.forEach(function(k) { out[k] = allProps[SETTINGS_PROP_PREFIX_ + k] || ""; });
  try { out.enabledLangs = out.enabledLangs ? JSON.parse(out.enabledLangs) : ["km", "en"]; }
  catch (err) { out.enabledLangs = ["km", "en"]; }
  if (out.logoUrl) out.logoUrl = normalizeDriveImageUrl_(out.logoUrl);
  if (out.reportLogoUrl) out.reportLogoUrl = normalizeDriveImageUrl_(out.reportLogoUrl);
  return out;
}

// ==== FIX (សំណើថ្មី "បិទកន្លែងស្នើសុំបង្កើតគណនីថ្មីទាំងអស់")៖ ត្រឡប់ថាតើប្រព័ន្ធអនុញ្ញាតឲ្យ Panel "ស្នើសុំគណនីថ្មី"
// (ទំព័រ Login សាធារណៈ — មិនតម្រូវឲ្យ Login ជាមុនទេ) ប្រើប្រាស់បានដែរឬអត់ — កុងតាក់ធំតែមួយ គ្រប់គ្រងទាំង Panel/តំណភ្ជាប់
// នៅ Client (មើល loadAuthPageInfo()/showAuthPanel() ក្នុង Index.html) និងខ្ទប់ requestSignup() ខាង Server ជាការពារ
// បន្ថែម (Defense-in-depth — ទោះបីជា Client ត្រូវបានបំបាត់/រំលង ក៏ Server នៅតែបដិសេធដដែល) ។ ខុសពី
// isCommuneAccountsEnabled_()/isObserverTierRoleEnabled_() ខាងក្រោម (ដែលគ្រប់គ្រងតែប្រភេទគណនីជាក់លាក់ម្នាក់ៗ)៖
// កុងតាក់នេះជា "មេ" បិទផ្លូវ "ស្នើសុំខ្លួនឯង" ទាំងអស់ម្តងតែម្តង មិនប៉ះពាល់ដល់ addUserAccount() (SuperAdmin/Admin/
// PEC21/គ្រូប្រចាំក្រុងស្រុកបង្កើតគណនីឲ្យអ្នកដទៃដោយផ្ទាល់) ដែលនៅតែដំណើរការធម្មតា — លំនាំដើមអនុញ្ញាត (true) ដូចគ្នា
// នឹង isCommuneAccountsEnabled_() ====
function isSelfSignupEnabled_() {
  return getSystemSettings_().enableSelfSignup !== 'false';
}

// ==== FIX (Fix127, "ចង់ឲ្យមានជម្រើសមួយដែលអាចប្រើប្រាស់គណនីឃុំសង្កាត់ក៏បាន ឬមិនប្រើក៏បាន")៖ ត្រឡប់ថាតើប្រព័ន្ធ
// អនុញ្ញាតឲ្យ "បង្កើត" គណនីឃុំសង្កាត់ (ROLE_COMMUNE) ថ្មីដែរឬអត់ — ទាំងផ្លូវស្នើសុំខ្លួនឯង (requestSignup) និងផ្លូវ
// គ្រូប្រចាំក្រុងស្រុក/SuperAdmin/Admin/PEC21 បង្កើតឲ្យ (addUserAccount) ។ លំនាំដើមអនុញ្ញាត (true) ប្រសិនបើ Admin
// មិនទាន់កំណត់ផ្ទាល់ ("" ទទេ) ដើម្បីកុំប៉ះពាល់ខេត្ត/ស្រុកដែលកំពុងប្រើប្រាស់មុខងារនេះស្រាប់ (Fix122-126) ភ្លាមៗពេល
// ដំឡើងកំណែនេះ — Admin ត្រូវចូល "ការកំណត់ប្រព័ន្ធ" រួចបិទដោយផ្ទាល់ (កំណត់ជា "false") ទើបបិទបាន (ដូចគ្នានឹងលំនាំ
// showAggPercent1-4) ។ ចំណាំសំខាន់៖ នេះជាការហាមឃាត់ការ "បង្កើត" គណនីថ្មីប៉ុណ្ណោះ — គណនីឃុំសង្កាត់ដែលមានស្រាប់ នៅតែ
// អាច Login/ប្រើប្រាស់/គ្រប់គ្រង (កែប្រែ/ផ្អាក/លុប) បានដដែល មិនរងផលប៉ះពាល់ពីការកំណត់នេះទេ ====
function isCommuneAccountsEnabled_() {
  return getSystemSettings_().enableCommuneAccounts !== 'false';
}

// ==== FIX (Fix128, "ចង់ឲ្យមានជម្រើសបន្តទៀត...គណនីឃុំសង្កាត់ដែលបង្កើតហើយ អាចជ្រើសរើសប្រើបាន ឬមិនអាចប្រើបានបន្តទៀត")៖
// ត្រឡប់ថាតើគណនីឃុំសង្កាត់ (ROLE_COMMUNE) ដែល "មានស្រាប់" រួចហើយ អាចបន្ត Login/ប្រើប្រាស់ប្រព័ន្ធបានដែរឬអត់ — ខុសពី
// isCommuneAccountsEnabled_() ខាងលើ (Fix127, ត្រួតពិនិត្យតែផ្លូវ "បង្កើត" គណនីថ្មីប៉ុណ្ណោះ) — Key នេះជា "កុងតាក់ធំ"
// មួយ ដែលអាចផ្អាកគណនីឃុំសង្កាត់ទាំងអស់ភ្លាមៗតែម្តង (Login ថ្មីត្រូវបដិសេធ + Session ដែលកំពុងប្រើប្រាស់ស្រាប់ត្រូវបាន
// កាត់ផ្តាច់ស្វ័យប្រវត្តិ) ដោយមិនចាំបាច់ចូល "ផ្អាក" ម្តងមួយៗគណនីតាមតារាង (ដែលនឹងលំបាកខ្លាំង បើមានគណនីច្រើន) ហើយ
// ក៏មិនប៉ះពាល់ស្ថានភាព "សកម្ម/ផ្អាក" ដែលកំណត់ដាច់ដោយឡែកសម្រាប់គណនីនីមួយៗក្នុង Sheet ទេ (អាចត្រឡប់មកបើកវិញបាន
// ភ្លាមៗ ដោយមិនចាំបាច់ធ្វើសកម្មភាពអ្វីម្តងទៀត) — លំនាំដើមអនុញ្ញាត (true) ដូចគ្នានឹងហេតុផល isCommuneAccountsEnabled_() ====
function isCommuneAccountsUsageEnabled_() {
  return getSystemSettings_().enableCommuneAccountsUsage !== 'false';
}

// ==== FIX (Fix129, "គណនីថ្មីទាំង៥ ត្រូវមាន Option បើក/បិទ ការបង្កើតគណនីនីមួយៗ")៖ ត្រឡប់ថាតើប្រព័ន្ធអនុញ្ញាតឲ្យ
// "បង្កើត" គណនីតួនាទីណាមួយក្នុងចំណោមតួនាទីថ្មីទាំង៥ (មើល OBSERVER_TIER_ROLES_ ក្នុង Utils.gs) ដែរឬអត់ — ត្រួតពិនិត្យ
// តាម role ជាក់លាក់ (រកមើល settingKey ដែលទាក់ទងតាម OBSERVER_TIER_ROLES_) ។ ខុសពី isCommuneAccountsEnabled_()/
// isCommuneAccountsUsageEnabled_() ខាងលើ (លំនាំដើម "អនុញ្ញាត")៖ តួនាទីទាំង៥នេះ លំនាំដើម "បិទ" ទាំងអស់ (ត្រូវ
// SuperAdmin/Admin ចូល "ការកំណត់ប្រព័ន្ធ" បើកជាក់លាក់ដោយផ្ទាល់ — === 'true' ទើបចាត់ទុកជាបើក) — អនុវត្តទាំងផ្លូវ
// ស្នើសុំខ្លួនឯង (requestSignup) និងផ្លូវ Admin បង្កើតឲ្យ (addUserAccount) ។ បើ role មិនមែនក្នុងចំណោមទាំង៥នេះ ត្រឡប់
// false ជានិច្ច (មិនត្រូវការធីចនេះទាល់តែសោះ) ====
function isObserverTierRoleEnabled_(role) {
  var cfg = getObserverTierRoleConfigByRole_(role);
  if (!cfg) return false;
  return getSystemSettings_()[cfg.settingKey] === 'true';
}

// ==== ត្រឡប់ការកំណត់ជួរឈរបន្ថែម (ស្ថិតិប៉ាន់ស្មាន/ការិ.ថ្មី/លេខកូដ/ស្ទួន/ភាគរយ) សម្រាប់ "ក្រុម" (Group) មួយ
// ក្នុងចំណោម ៣ ក្រុមឯករាជ្យពីគ្នា ដែលបានស្នើសុំ៖
//   suffix "" (ទទេ)  → ទម្រង់បញ្ចូលទិន្នន័យប្រចាំថ្ងៃ (Group A — Key ដើម មិនប្តូរ)
//   suffix "Daily"   → Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) (Group B — getDailyEntries/renderDailyTable)
//   suffix "Sum"     → Tab(របាយការណ៍សរុបតាមឃុំសង្កាត់/សរុបខេត្ត) + របាយការណ៍សម្រាប់បោះពុម្ព + Excel Export
//                      (Group C — ចែករំលែកគ្នា ព្រោះទាំងអស់នេះមកពី getSummaries()/computeSummaries_() ដូចគ្នា)
// ភាពឆបគ្នា/មិនប៉ះពាល់អ្នកប្រើចាស់៖ បើ Admin មិនទាន់ធ្លាប់កំណត់ Key ជាក់លាក់របស់ក្រុម Daily/Sum ផ្ទាល់ខ្លួនទេ
// (តម្លៃនៅតែជា "" ទទេ ដូចលំនាំដើម) — ត្រឡប់ទៅប្រើតម្លៃដដែលនឹងទម្រង់បញ្ចូលទិន្នន័យ (Group A) ជាមុនសិន ដូច្នេះការ
// បង្ហាញ/លាក់ជួរឈរនានា មិនប្តូរភ្លាមៗដោយឥតដឹងខ្លួន ភ្លាមៗពេលដំឡើងកំណែនេះឡើយ (រហូតដល់ Admin ជ្រើសរើសកំណត់ដាច់ដោយឡែក
// ជាថ្មីសម្រាប់ក្រុមណាមួយផ្ទាល់) ====
function resolveGroupDisplaySettings_(sysSettings, suffix) {
  suffix = suffix || '';
  function pick(baseKey) {
    var scoped = sysSettings[baseKey + suffix];
    if (suffix && scoped !== '' && scoped != null) return scoped === 'true';
    return sysSettings[baseKey] === 'true';
  }
  return {
    showEstimate: pick('showEstimate'),
    showProvinceEstimate: pick('showProvinceEstimate'),
    showNewStation: pick('showNewStation'),
    showTotalStation: pick('showTotalStation'),
    showCommuneCode: pick('showCommuneCode'),
    showDuplicateFields: pick('showDuplicateFields'),
    // ==== ជួរឈរ "ភាគរយ" ១-៤ ចែករំលែក Key តែមួយសម្រាប់គ្រប់ក្រុម (មានន័យសម្រាប់តែក្រុម "Sum" ប៉ុណ្ណោះជាក់ស្តែង) —
    // លំនាំដើមបង្ហាញ (ខុសពីធីចផ្សេងទៀត) — លាក់លុះត្រាតែកំណត់ជា "false" ជាក់ស្តែងប៉ុណ្ណោះ ====
    showAggPercent1: sysSettings.showAggPercent1 !== 'false',
    showAggPercent2: sysSettings.showAggPercent2 !== 'false',
    showAggPercent3: sysSettings.showAggPercent3 !== 'false',
    showAggPercent4: sysSettings.showAggPercent4 !== 'false',
    // ==== ជួរឈរ "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)" + "ភាគរយ" ចែករំលែក Key តែមួយ (មានន័យសម្រាប់តែក្រុម "Sum" ប៉ុណ្ណោះ
    // ជាក់ស្តែង) — លំនាំដើមបង្ហាញ ដូចគ្នានឹងធីច "ភាគរយ" ១-៤ ====
    showBioTotalSum: sysSettings.showBioTotalSum !== 'false'
  };
}

// ==== FIX (សំណើថ្មី "១")៖ ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទប្រចាំថ្ងៃ
// (លេខ១.២ — dailyCommune/dailyProvince) ខុសពី resolveGroupDisplaySettings_() ធម្មតា ត្រង់ថា មិន Fallback
// ទៅ Group (ក) ទម្រង់បញ្ចូលទិន្នន័យដោយផ្ទាល់ទេ (ព្រោះមុននេះរបាយការណ៍ ១.២ មិនដែលអានតម្លៃពី Group (ក) ទាល់តែសោះ)
// ប៉ុន្តែ Fallback ទៅតម្លៃដែលក្រុម "Sum" កំពុងបង្ហាញជាក់ស្តែងជាមុនសិន (ដែលជាអ្វីដែលរបាយការណ៍ ១.២ ធ្លាប់ប្រើប្រាស់
// រួមគ្នាជាមួយរបាយការណ៍ ៣.៤ ពីមុនមក) ដើម្បីធានាថាការបង្ហាញមិនប្រែប្រួលភ្លាមៗពេលដំឡើងកំណែនេះ រហូតដល់ Admin
// ជ្រើសរើសកំណត់ដាច់ដោយឡែកជាថ្មីសម្រាប់ក្រុមនេះផ្ទាល់ ====
function resolvePrintDailyDisplaySettings_(sysSettings) {
  var sumGroup_ = resolveGroupDisplaySettings_(sysSettings, 'Sum');
  function pickPD_(key, fallbackVal) {
    var v = sysSettings[key];
    if (v === '' || v == null) return fallbackVal;
    return v === 'true';
  }
  return {
    showEstimate: pickPD_('showEstimatePrintDaily', sumGroup_.showEstimate),
    showProvinceEstimate: pickPD_('showProvinceEstimatePrintDaily', sumGroup_.showProvinceEstimate),
    showNewStation: pickPD_('showNewStationPrintDaily', sumGroup_.showNewStation),
    showTotalStation: pickPD_('showTotalStationPrintDaily', sumGroup_.showTotalStation),
    showCommuneCode: pickPD_('showCommuneCodePrintDaily', sumGroup_.showCommuneCode),
    showDuplicateFields: pickPD_('showDuplicateFieldsPrintDaily', sumGroup_.showDuplicateFields)
  };
}

// ==== FIX (សំណើថ្មី "៣.៤")៖ ការកំណត់ជួរឈរបន្ថែម ដាច់ដោយឡែក សម្រាប់ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទធ្វើជីវមាត្រ
// បូកយោង (លេខ៥.៦ — bioCommune/bioProvince) ។ ខុសពី resolveGroupDisplaySettings_() ធម្មតា ត្រង់ថា មិន Fallback
// ទៅ Group (ក)/Base key ទេ ព្រោះជួរឈរ "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" ក្នុងរបាយការណ៍នេះ ពីមុនមកបង្ហាញជានិច្ចដោយ Hardcode
// (មិនអាចកំណត់បាន) ដូច្នេះលំនាំដើម (មិនទាន់កំណត់អ្វីទេ) ត្រូវតែបន្តបង្ហាញ "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" ជានិច្ចដដែល
// (មិនប្រែប្រួលការបង្ហាញបច្ចុប្បន្ន) រហូតដល់ Admin កំណត់ដោយខ្លួនឯង — "ស្ថិតិប៉ាន់ស្មានខេត្ត" ជាជួរឈរថ្មី ដូច្នេះលំនាំដើម
// លាក់ (មិនធ្លាប់មានពីមុនទេ) ។ "លេខកូដ" ឃុំ/សង្កាត់ Fallback ទៅ Base key ធម្មតា (ដូចនឹងអ្វីដែលរបាយការណ៍នេះធ្លាប់
// អានពីមុន — SHOW_COMMUNE_CODE_ គ្មាន Suffix) ====
function resolveBioDisplaySettings_(sysSettings) {
  function pickBio_(key, defaultVal) {
    var v = sysSettings[key];
    if (v === '' || v == null) return defaultVal;
    return v === 'true';
  }
  return {
    showEstimate: pickBio_('showEstimateBio', true),
    showProvinceEstimate: pickBio_('showProvinceEstimateBio', false),
    showCommuneCode: pickBio_('showCommuneCodeBio', sysSettings.showCommuneCode === 'true')
  };
}

// ត្រឡប់ការកំណត់ទាំងអស់ រួមទាំង Bot Token/Chat ID ដែលរសើប — សម្រាប់ទម្រង់កែប្រែក្នុងទំព័រកំណត់ប្រព័ន្ធ (Admin ប៉ុណ្ណោះ)
function getSystemSettingsAdmin(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលការកំណត់ប្រព័ន្ធទេ!" };
  return { success: true, settings: getSystemSettings_() };
}

function saveSystemSettings(currentUsername, patch, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិកែប្រែការកំណត់ប្រព័ន្ធទេ!" };
  // ==== FIX (សំខាន់ណាស់ — ការពារកុំឲ្យប្រព័ន្ធទាំងមូលដួលរលំ)៖ បើ Admin បិទភ្ជាប់ (masterSheetLink) ខុស ឬចង្អុលទៅ
  // Spreadsheet ដែលគ្មានសិទ្ធិចូល ហើយប្រព័ន្ធទទួលយកដោយមិនបានផ្ទៀងផ្ទាត់ជាមុន doGet() (ចំណុចចូលតែមួយគត់សម្រាប់អ្នកប្រើ
  // ប្រាស់ទាំងអស់) នឹងបោះ Error ភ្លាមៗ ធ្វើឲ្យអ្នកប្រើប្រាស់ទាំងអស់ (រួមទាំង Admin ខ្លួនឯង) មិនអាចបើកទំព័របានទៀត ដោយ
  // គ្មានផ្លូវសម្រួលដោយខ្លួនឯង (ត្រូវការចូល Apps Script Editor ដោយផ្ទាល់ ដើម្បីលុប Script Property ចោល) — ដូច្នេះ
  // ត្រូវសាកល្បងបើក Spreadsheet នោះជាមុនសិន មុននឹងទទួលយក ====
  if (Object.prototype.hasOwnProperty.call(patch || {}, 'masterSheetLink')) {
    var newLink = String(patch.masterSheetLink || '').trim();
    if (newLink) {
      var newId = extractSpreadsheetId_(newLink);
      if (!newId) {
        return { success: false, message: "តំណភ្ជាប់ Google Sheet មិនត្រឹមត្រូវទេ! សូមចម្លង URL ឬ ID របស់ Spreadsheet ឲ្យបានត្រឹមត្រូវ។" };
      }
      try {
        SpreadsheetApp.openById(newId);
      } catch (errOpen) {
        return { success: false, message: "មិនអាចបើក Google Sheet នេះបានទេ! សូមពិនិត្យ URL/ID ម្តងទៀត ហើយធានាថាបានចែករំលែក (Share) ជាមួយគណនីដែលដំណើរការ Script នេះ (ជាធម្មតាគណនីដែលបើក Apps Script Editor)។" };
      }
    }
  }
  var props = PropertiesService.getScriptProperties();
  Object.keys(patch || {}).forEach(function(k) {
    if (SETTINGS_KEYS_.indexOf(k) === -1) return;
    var v = patch[k];
    if (k === "enabledLangs") v = JSON.stringify((v && v.length) ? v : ["km"]);
    props.setProperty(SETTINGS_PROP_PREFIX_ + k, String(v || ""));
  });
  return { success: true, message: "រក្សាទុកការកំណត់ជោគជ័យ!" };
}

// បញ្ចូលរូបភាព Logo ស្ថាប័ន — រក្សាទុកក្នុង Folder តែមួយជាមួយ Spreadsheet មេ ហើយចែករំលែកជាសាធារណៈ (មើលបានតែមើលទេ)
function uploadOrgLogo(currentUsername, base64Data, mimeType, fileName, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិបញ្ចូល Logo ទេ!" };
  if (!base64Data) return { success: false, message: "មិនមានទិន្នន័យរូបភាពទេ! សូមព្យាយាមជ្រើសរើសរូបភាពម្តងទៀត។" };

  var file;
  try {
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || "image/png", fileName || "logo");
    var folder = getMasterParentFolder_() || DriveApp.getRootFolder();
    file = folder.createFile(blob);
  } catch (err) {
    return { success: false, message: "កំហុសពេលបញ្ចូលឯកសារ៖ " + err.message };
  }

  // ព្យាយាមចែករំលែកជាសាធារណៈដាច់ដោយឡែក — ព្រោះ Google Workspace ស្ថាប័នខ្លះ (ឧ. រដ្ឋាភិបាល) បិទសិទ្ធិនេះ
  // ថ្នាក់ Domain ទាំងមូល ដែលធ្វើឲ្យ setSharing() បោះ Error ។ បើករណីនេះកើតឡើង យើងនៅតែរក្សា Logo ទុក
  // (កុំបោះបង់ចោល ឬបង្ហាញកំហុសទាំងស្រុង) គ្រាន់តែជូនដំណឹងថា Logo អាចនឹងមិនបង្ហាញលើទំព័រ Login សម្រាប់
  // អ្នកមិនទាន់ចូលប្រព័ន្ធ (សាធារណៈ) ប៉ុណ្ណោះ។
  var sharedOk = true;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    sharedOk = false;
  }

  // ==== ប្រើទម្រង់ lh3.googleusercontent.com/d/{id} ជំនួស uc?export=view ព្រោះទម្រង់ចាស់ (uc?export=view)
  // Google តែងតែទប់ស្កាត់ការបង្ហាញផ្ទាល់ក្នុង <img> (បង្ហាញទំព័រព្រមាន/ស្កេនវីរុសជំនួសរូបភាព) ជាហេតុធ្វើឲ្យ
  // រូបភាព Logo មិនចេញ (ធ្លាក់ទៅរូប 🗳️ សម្រាប់ជំនួសវិញ) ជាពិសេសនៅលើរបារខាងឆ្វេង/ទំព័រ Login ពេលផ្ទុកទំព័រថ្មី។
  // ទម្រង់ lh3.googleusercontent.com គឺជាទម្រង់ដដែលប្រើសម្រាប់ URL ដែលបញ្ចូលដោយផ្ទាល់ (មើល normalizeDriveImageUrl_
  // ក្នុង Index.html) ដូច្នេះត្រូវឲ្យស្របគ្នា ដើម្បីទុកចិត្តបានថាបង្ហាញរូបភាពជានិច្ច។
  var url = "https://lh3.googleusercontent.com/d/" + file.getId();
  PropertiesService.getScriptProperties().setProperty(SETTINGS_PROP_PREFIX_ + "logoUrl", url);

  if (sharedOk) {
    return { success: true, message: "បញ្ចូល Logo ជោគជ័យ!", url: url };
  }
  return {
    success: true, url: url,
    message: "បញ្ចូល Logo ជោគជ័យ! ប៉ុន្តែគោលការណ៍ Google Workspace របស់ស្ថាប័នអ្នក រារាំងការចែករំលែក Drive " +
      "ជាសាធារណៈ — Logo អាចនឹងមិនបង្ហាញលើទំព័រ Login សម្រាប់អ្នកមកទស្សនាដែលមិនទាន់ចូលប្រព័ន្ធនោះទេ។ " +
      "សូមទាក់ទងអ្នកគ្រប់គ្រង Google Workspace ដើម្បីអនុញ្ញាតការចែករំលែក ឬប្រើ URL រូបភាពពីកន្លែងផ្សេង (ឧ. imgur.com) ជំនួសវិញ។"
  };
}

// បញ្ចូលរូបភាព Logo សម្រាប់ក្បាលរបាយការណ៍សម្រាប់បោះពុម្ព (ខុសពី Logo ស្ថាប័នខាងលើ ដែលប្រើតែលើទំព័រ Login/របារខាងឆ្វេង)
// — រក្សាទុកក្នុង Folder តែមួយជាមួយ Spreadsheet មេ ហើយចែករំលែកជាសាធារណៈ (មើលបានតែមើលទេ) ដូចគ្នា
function uploadReportLogo(currentUsername, base64Data, mimeType, fileName, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិបញ្ចូល Logo ទេ!" };
  if (!base64Data) return { success: false, message: "មិនមានទិន្នន័យរូបភាពទេ! សូមព្យាយាមជ្រើសរើសរូបភាពម្តងទៀត។" };

  var file;
  try {
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, mimeType || "image/png", fileName || "report-logo");
    var folder = getMasterParentFolder_() || DriveApp.getRootFolder();
    file = folder.createFile(blob);
  } catch (err) {
    return { success: false, message: "កំហុសពេលបញ្ចូលឯកសារ៖ " + err.message };
  }

  var sharedOk = true;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    sharedOk = false;
  }

  var url = "https://lh3.googleusercontent.com/d/" + file.getId();
  PropertiesService.getScriptProperties().setProperty(SETTINGS_PROP_PREFIX_ + "reportLogoUrl", url);

  if (sharedOk) {
    return { success: true, message: "បញ្ចូល Logo ជោគជ័យ!", url: url };
  }
  return {
    success: true, url: url,
    message: "បញ្ចូល Logo ជោគជ័យ! ប៉ុន្តែគោលការណ៍ Google Workspace របស់ស្ថាប័នអ្នក រារាំងការចែករំលែក Drive " +
      "ជាសាធារណៈ — Logo អាចនឹងមិនបង្ហាញលើរបាយការណ៍សម្រាប់អ្នកមកទស្សនាដែលមិនទាន់ចូលប្រព័ន្ធនោះទេ។ " +
      "សូមទាក់ទងអ្នកគ្រប់គ្រង Google Workspace ដើម្បីអនុញ្ញាតការចែករំលែក ឬប្រើ URL រូបភាពពីកន្លែងផ្សេង (ឧ. imgur.com) ជំនួសវិញ។"
  };
}

// ==== ការជូនដំណឹងតាម Telegram ====
function notifyTelegram_(text) {
  try {
    var s = getSystemSettings_();
    // ==== ល្បឿន៖ ដាក់ជាជួរដេក មិនផ្ញើភ្លាមៗទេ — ជៀសវាងឲ្យអ្នកប្រើរង់ចាំ Telegram API ខណៈកំពុងស្នើសុំគណនី/បង្កើតគណនី ====
    if (s.tgBotToken && s.tgChatId) queueTelegramNotify_(text);
  } catch (err) {}
}

// ==== FIX (Fix118, "សូមឲ្យការផ្ញើសារ Key Number នៅពេលចុច(ស្នើសុំគណនី)ភ្លាមគឺ ផ្ញើភ្លាមតែម្តង")៖ notifyTelegram_()
// ខាងលើ ដាក់ជាជួរដេក (Queue+Trigger — ចេតនាដើម្បីកុំឲ្យ Login/Logout/បង្កើតគណនីយឺត) ប៉ុន្តែសារ Key Number ពេលមាន
// ការស្នើសុំគណនីថ្មី (requestSignup) មានលក្ខណៈបន្ទាន់ខុសពីគេ — SuperAdmin ត្រូវការឃើញ+ផ្ញើ Key Number ជូនអ្នកស្នើសុំ
// ជាបន្ទាន់ ដើម្បីឲ្យគាត់អាចបញ្ចប់ការបង្កើតគណនីបានឆាប់ (Trigger ជួរដេក after(2000) មិនធានាថាដំណើរការក្នុងរយៈពេល
// ខ្លីនោះទេ — ជាក់ស្តែងអាចចំណាយពេលរាប់នាទី) ។ ដូច្នេះ Function នេះហៅ sendTelegramMessage_() ដោយផ្ទាល់ (Synchronous
// — មិនឆ្លងកាត់ជួរដេកទេ) ព្រោះស្នើសុំគណនីជាសកម្មភាពកម្រ (មិនញឹកញាប់ដូច Login/Logout ទេ) ដូច្នេះការរង់ចាំបន្ថែម
// ១-២វិនាទីសម្រាប់ការហៅ Telegram API តែម្តងនេះ សមហេតុផលជាងការឲ្យ Key Number ធ្វើដំណើរយឺត ====
function notifyTelegramImmediate_(text) {
  try {
    var s = getSystemSettings_();
    if (s.tgBotToken && s.tgChatId) sendTelegramMessage_(s.tgBotToken, s.tgChatId, text);
  } catch (err) {}
}

function sendTelegramTest(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var s = getSystemSettings_();
  if (!s.tgBotToken || !s.tgChatId) return { success: false, message: "សូមបំពេញ Bot Token និង Chat ID ជាមុនសិន!" };
  var ok = sendTelegramMessage_(s.tgBotToken, s.tgChatId, "🔔 សារសាកល្បងពី " + (s.orgName || "ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ") + " — ការតភ្ជាប់ Telegram ដំណើរការត្រឹមត្រូវ!");
  return ok ? { success: true, message: "ផ្ញើសារសាកល្បងជោគជ័យ! សូមពិនិត្យ Telegram។" }
            : { success: false, message: "ផ្ញើមិនបានទេ សូមពិនិត្យ Bot Token/Chat ID ឡើងវិញ។" };
}

// កំណត់ចំនួនចូលទស្សនាទំព័រ Login (ចូលប្រើ) ត្រឡប់ទៅ ០ វិញ — សម្រាប់ Admin ចង់សម្អាតស្ថិតិចាស់ចោល
var BACKUP_NAME_PREFIX_ = "ទិន្នន័យបម្រុងទុក - ";

// ==== FIX (ស្នើសុំដោយ SuperAdmin)៖ ពេលចុច Reset ត្រូវបង្កើត Google Sheet ស្រុកគ្រូប្រចាំក្រុងស្រុកទាំង១០ "ថ្មីទាំងអស់តែម្តង"
// (មិនមែនគ្រាន់តែសម្អាត Tab ខាងក្នុង Spreadsheet ចាស់ដដែលទេ) ព្រមទាំងសម្អាត Script Property/Trigger/Cache ចាស់ៗ
// ដែលពាក់ព័ន្ធនឹង Spreadsheet ស្រុកនោះឲ្យអស់ — ដើម្បីកុំឲ្យ Property ចាស់ ចង្អុលទៅ Spreadsheet ដែលលែងបើកកើត
// (ឧ. ត្រូវបានលុប/ផ្លាស់ទីដោយដៃដោយចៃដន្យ) ជាប់សល់នៅសេសសល់ ដែលធ្លាប់បណ្តាលឲ្យ getDistrictSpreadsheet_() បោះកំហុស
// ("មិនអាចបើក Spreadsheet របស់ស្រុក ... បានទេ") ហើយធ្វើឲ្យ Reset ទាំងមូលបរាជ័យ។
//
// អាកប្បកិរិយា៖ (១) បើ Spreadsheet ស្រុកចាស់នៅមាន/បើកបាន ត្រូវផ្ទេរទៅធុងសំរាម Google Drive (Trash — មិនលុបជា
// អចិន្ត្រៃយ៍ភ្លាមៗទេ អាចស្តារវិញបានក្នុងរយៈពេល ៣០ថ្ងៃ តាមប្រព័ន្ធ Google ខ្លួនឯង) ជាការការពារបន្ថែមតូចមួយ ទោះជា
// មិនទាមទារទុកទិន្នន័យចាស់ក៏ដោយ។ បើបើកមិនកើតទាល់តែសោះ (ឧ. ត្រូវបានលុបរួចជាស្រេច) មិនចាត់ទុកជាកំហុសទេ រំលងទៅជំហាន
// បន្ទាប់ភ្លាមៗ។ (២) លុប Trigger onEdit ចាស់ដែលភ្ជាប់ជាមួយ Spreadsheet ID ចាស់។ (៣) លុប Script Property ចាស់
// (districtSS_...) + Cache ក្នុងសតិ ដើម្បីឲ្យ getDistrictSpreadsheet_() ជឿថា "ស្រុកនេះមិនទាន់មាន Spreadsheet ស្រាប់ទេ"
// ហើយបង្កើតថ្មីស្អាតដោយស្វ័យប្រវត្តិ (រួមទាំងដំឡើង Trigger onEdit ថ្មីភ្ជាប់ស្រាប់ ដូចករណីបង្កើតលើកដំបូងធម្មតា)។
function recreateDistrictSpreadsheetFresh_(district, effectiveId, existingTriggers) {
  var props = PropertiesService.getScriptProperties();
  var key = DISTRICT_SS_PROP_PREFIX + effectiveId + '_' + district;
  var oldSsId = props.getProperty(key);

  if (oldSsId) {
    try {
      var file = DriveApp.getFileById(oldSsId);
      if (!file.isTrashed()) file.setTrashed(true);
    } catch (eTrash) {
      // Spreadsheet ចាស់បើកមិនកើតទាល់តែសោះ (ប្រហែលជាបានលុប/ផ្លាស់ទីរួចហើយ) — មិនអីទេ បន្តទៅសម្អាត Property ដដែល
    }

    (existingTriggers || []).forEach(function(t) {
      try {
        if (t.getHandlerFunction() === 'onDistrictSheetEdit_' && t.getTriggerSourceId() === oldSsId) {
          ScriptApp.deleteTrigger(t);
        }
      } catch (eDel) {}
    });

    try { props.deleteProperty(key); } catch (eProp) {}
    try { delete DISTRICT_SS_CACHE_[effectiveId + '|' + district]; } catch (eCache) {}
  }

  // ==== ចំណុចស្នូល៖ ដោយសារ Property ត្រូវបានលុបខាងលើរួច getDistrictSpreadsheet_() នឹង "គិត" ថាស្រុកនេះមិនទាន់
  // មាន Spreadsheet ស្រាប់ទេ ហើយបង្កើតថ្មីស្អាតដោយស្វ័យប្រវត្តិ (ជាមួយ Trigger onEdit ថ្មីភ្ជាប់ស្រាប់) ====
  var freshSs = getDistrictSpreadsheet_(district);

  // ==== សម្អាត Tab លំនាំដើម "Sheet1" ដែល Google បង្កើតឲ្យស្វ័យប្រវត្តិ ពេលបង្កើត Spreadsheet ថ្មី — មិនអាចលុបភ្លាមៗ
  // ទេ ព្រោះ Google មិនអនុញ្ញាតលុប Sheet ចុងក្រោយគេ (ត្រូវរង់ចាំ Period Sheet ថ្មីត្រូវបានបញ្ចូលសិន — មើល resetAllData) ====
  return freshSs;
}

// ==================== Reset ទិន្នន័យទាំងអស់ត្រឡប់ទៅដើម (SuperAdmin ប៉ុណ្ណោះ — សកម្មភាពមិនអាចត្រឡប់វិញបានទេ!) ====================
// លុប Period Sheet (Tab ថ្ងៃនីមួយៗ) ទាំងអស់ ក្នុង Google Sheet ស្រុកទាំង១០ + សម្អាតទិន្នន័យក្នុង Spreadsheet មេ
// (របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ/សរុបឃុំសង្កាត់/សរុបខេត្ត) ។ មិនប៉ះពាល់ដល់៖ ទិន្នន័យភ្ជាប់, អ្នកប្រើប្រាស់, បញ្ជីឃុំ/សង្កាត់ ឡើយ។
// ==== FIX៖ Google Sheet ស្រុកគ្រូប្រចាំក្រុងស្រុកទាំង១០ ត្រូវបានបង្កើត "ថ្មីទាំងស្រុង" ជារៀងរាល់ពេល Reset (មិនមែនគ្រាន់តែ
// សម្អាត Tab ខាងក្នុង Spreadsheet ចាស់ដដែលទៀតទេ) តាមរយៈ recreateDistrictSpreadsheetFresh_() ខាងលើ — ដោះស្រាយទាំង
// សំណើ "បង្កើត Sheet ស្រុកថ្មីទាំងអស់តែម្តង" ព្រមទាំងករណី Property ចាស់ខូច/ចង្អុលទៅ Spreadsheet ដែលលែងបើកកើត ====
// ==== FIX (សំណើថ្មី "Reset ស្ងាត់ៗ មិនដឹងដំណើរការកម្រិតណា")៖ resetAllData() ខាងក្រោម ជាប្រតិបត្តិការយឺតបំផុតមួយ
// ក្នុងប្រព័ន្ធ (បង្កើត Google Sheet ស្រុកថ្មីទាំង១០ ជាប់គ្នា — Google Apps Script មិនអាចដំណើរការស្របគ្នាបានទេ ក្នុង
// ការហៅតែម្តងគត់) ដូច្នេះ Key/TTL នេះប្រើសម្រាប់រក្សាទុក "វឌ្ឍនភាព" (ស្រុកណាហើយ ប៉ុន្មានហើយ) ទៅ CacheService ភ្លាមៗ
// ក្រោយស្រុកនីមួយៗចប់ — Client ស្ទង់មើល (Poll) តាមរយៈ getResetAllDataProgress() ជាទៀងទាត់ ខណៈកំពុងរង់ចាំ Call
// ចម្បង (resetAllData) បញ្ចប់ ដើម្បីបង្ហាញវឌ្ឍនភាពជាក់ស្តែង ជំនួសការស្ងាត់ស្ងៀមទាំងស្រុង ====
var CACHE_KEY_RESET_PROGRESS_ = 'resetAllDataProgress_v1';
var CACHE_TTL_RESET_PROGRESS_ = 600; // វិនាទី (១០នាទី — លើសពេលវេលាដំណើរការជាក់ស្តែងច្រើន សម្រាប់ជាកម្រិតសុវត្ថិភាព)

function getResetAllDataProgress(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isSuperAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  try {
    var raw = CacheService.getScriptCache().get(CACHE_KEY_RESET_PROGRESS_);
    if (!raw) return { success: true, done: 0, total: DISTRICT_LIST.length, district: '', finished: false };
    var p = JSON.parse(raw);
    return { success: true, done: p.done || 0, total: p.total || DISTRICT_LIST.length, district: p.district || '', finished: !!p.finished };
  } catch (err) {
    return { success: true, done: 0, total: DISTRICT_LIST.length, district: '', finished: false };
  }
}

function resetAllData(currentUsername, startDate, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isSuperAdmin_(currentUsername)) {
    return { success: false, message: "មានតែ SuperAdmin ទេ ដែលអាច Reset ទិន្នន័យទាំងអស់បាន!" };
  }
  var d = normalizeDateStr_(startDate);
  if (!d || !/^\d{4}-\d{2}-\d{2}$/.test(d)) {
    return { success: false, message: "សូមជ្រើសរើសកាលបរិច្ឆេទចាប់ផ្តើមឲ្យត្រឹមត្រូវ!" };
  }

  try {
    var recreatedCount = 0;
    var districtErrors = [];
    try { CacheService.getScriptCache().put(CACHE_KEY_RESET_PROGRESS_, JSON.stringify({ done: 0, total: DISTRICT_LIST.length, district: '', finished: false }), CACHE_TTL_RESET_PROGRESS_); } catch (eProg0) {}

    // ០. កំណត់ "ថ្ងៃចាប់ផ្តើមដំណើរការ" ជាមុនសិន (មុននឹងបង្កើត Sheet) ដើម្បីឲ្យតក្កវិជ្ជា "ថ្ងៃដំបូង" ត្រឹមត្រូវ ខណៈបង្កើត
    try { setSetting_(SETTINGS_KEY_START_DATE, d); } catch (e) {}
    try { CacheService.getScriptCache().removeAll([CACHE_KEY_BASELINE_MAP_, CACHE_KEY_UNLOCKS_]); } catch (e) {}

    // ១. បង្កើត Google Sheet ស្រុកគ្រូប្រចាំក្រុងស្រុកទាំង១០ "ថ្មីទាំងស្រុង" (ផ្ទេរចាស់ទៅធុងសំរាម + សម្អាត Property/
    //    Trigger/Cache ចាស់ + បង្កើតថ្មីស្អាត) រួចទើបបញ្ចូល Period Sheet ថ្ងៃចាប់ផ្តើម (ថ្ងៃទី១) ចូល Spreadsheet ថ្មីនោះ។
    //    ==== FIX៖ ដំណើរការម្តងម្នាក់ៗ ក្នុង try/catch ដាច់ដោយឡែក — បើស្រុកមួយមានបញ្ហា (ឧ. Spreadsheet ចាស់/ថ្មីមានបញ្ហា
    //    សិទ្ធិចូលប្រើបណ្តោះអាសន្ន) មិនត្រូវឲ្យប៉ះពាល់ដល់ស្រុកដទៃទៀតដែលដំណើរការធម្មតាឡើយ (ស្រុកមុនធ្លាប់ធ្វើឲ្យ Reset
    //    ទាំងមូលបរាជ័យ ព្រោះ forEach ឈប់ភ្លាមៗពេលជួប Exception ណាមួយ) ====
    var existingTriggers_;
    try { existingTriggers_ = ScriptApp.getProjectTriggers(); } catch (eTrig) { existingTriggers_ = []; }
    var effectiveId_ = getTargetSheetId_();

    DISTRICT_LIST.forEach(function(district, districtIdx_) {
      try {
        var dSs = recreateDistrictSpreadsheetFresh_(district, effectiveId_, existingTriggers_);
        createDistrictPeriodSheet_(district, d); // ថ្ងៃចាប់ផ្តើម (ថ្ងៃទី១) ប៉ុណ្ណោះ
        // សម្អាត Tab លំនាំដើម "Sheet1" ដែល Google បង្កើតឲ្យស្វ័យប្រវត្តិ (ឥឡូវអាចលុបបាន ព្រោះមាន Period Sheet
        // ថ្មីរួចហើយ — Google មិនអនុញ្ញាតលុប Sheet ចុងក្រោយគេទេ)
        try {
          var defaultSheet = dSs.getSheetByName('Sheet1');
          if (defaultSheet && dSs.getNumSheets() > 1) dSs.deleteSheet(defaultSheet);
        } catch (eDefault) {}
        recreatedCount++;
      } catch (eDistrict) {
        districtErrors.push(district + '៖ ' + eDistrict.message);
      }
      // ==== FIX (សំណើថ្មី "Reset ស្ងាត់ៗ")៖ ធ្វើបច្ចុប្បន្នភាព "វឌ្ឍនភាព" ភ្លាមៗក្រោយស្រុកនីមួយៗចប់ (មិនថាជោគជ័យ ឬបរាជ័យ)
      // ដើម្បីឲ្យ Client ស្ទង់មើល (Poll — getResetAllDataProgress) ឃើញដំណើរការជាក់ស្តែង ====
      try {
        CacheService.getScriptCache().put(CACHE_KEY_RESET_PROGRESS_, JSON.stringify({
          done: districtIdx_ + 1, total: DISTRICT_LIST.length, district: district, finished: false
        }), CACHE_TTL_RESET_PROGRESS_);
      } catch (eProg) {}
    });

    // ២. សម្អាតទិន្នន័យ (រក្សាទុក Header) ក្នុង Spreadsheet មេ៖ របាយការណ៍ប្រចាំថ្ងៃ, សរុបឃុំសង្កាត់, សរុបខេត្ត
    var ss = getSS_();
    [ensureDailySheet_(ss), ensureCommuneSumSheet_(ss), ensureProvinceSumSheet_(ss)].forEach(function(sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow >= 4) sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).clearContent();
    });

    // ៣. សម្អាតកំណត់ត្រា "ការអនុញ្ញាតកែថ្ងៃចាស់" ចាស់ៗ (ព្រោះថ្ងៃទាំងនោះលែងមានទិន្នន័យទៀតហើយ)
    var unlockSheet = ensureEditUnlocksSheet_(ss);
    var uLastRow = unlockSheet.getLastRow();
    if (uLastRow >= 2) unlockSheet.getRange(2, 1, uLastRow - 1, unlockSheet.getLastColumn()).clearContent();

    // ៤. កំណត់ "ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន" ទៅជាថ្ងៃទី១ ដោយផ្ទាល់ (ព្រោះជាការ Reset ចេតនា — មិនប្រៀបធៀបជាមួយថ្ងៃពិតតាម Server ទេ)
    try { PropertiesService.getScriptProperties().setProperty(CURRENT_PERIOD_DATE_PROP_, d); } catch (e) {}

    // ៥. អនុវត្តការចាក់សោ/លាក់ជួរឈរ ភ្លាមៗលើ Sheet ថ្មី (ថ្ងៃទី១ មិនត្រូវលាក់ ព្រោះជា Sheet ដំបូងបំផុត)
    try { protectPastPeriodSheets_(); } catch (e) {}

    scheduleRebuild_();
    try {
      CacheService.getScriptCache().put(CACHE_KEY_RESET_PROGRESS_, JSON.stringify({
        done: DISTRICT_LIST.length, total: DISTRICT_LIST.length, district: '', finished: true
      }), CACHE_TTL_RESET_PROGRESS_);
    } catch (eProgFin) {}
    var msg = "Reset ទិន្នន័យទាំងអស់ជោគជ័យ! បានបង្កើត Google Sheet ស្រុកគ្រូប្រចាំក្រុងស្រុកថ្មីទាំងស្រុងចំនួន " + recreatedCount + "/" + DISTRICT_LIST.length +
      " (Spreadsheet ចាស់ត្រូវបានផ្ទេរទៅធុងសំរាម Google Drive — មិនលុបជាអចិន្ត្រៃយ៍ភ្លាមៗទេ) ព្រមទាំងបានបង្កើត Sheet ថ្ងៃចាប់ផ្តើម (" + d + ") ជូនរួចរាល់" +
      " (សម្រាប់ថ្ងៃបន្ទាប់ សូមចុច \"បង្កើត Sheet ថ្មីបន្ទាប់\" នៅ Tab របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ — ទិន្នន័យភ្ជាប់ អ្នកប្រើប្រាស់ និងបញ្ជីឃុំ/សង្កាត់ មិនប៉ះពាល់ទេ)";
    if (districtErrors.length) {
      msg += " ⚠️ ស្រុកចំនួន " + districtErrors.length + " មានបញ្ហា៖ " + districtErrors.join(' | ');
    }
    return { success: districtErrors.length < DISTRICT_LIST.length, message: msg };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== Folder ដែលទុកទិន្នន័យបម្រុងទុក — Admin អាចកំណត់ខ្លួនឯង (Tab ទិន្នន័យបម្រុងទុក) តាមរយៈ ID ឬ URL របស់ថតឯកសារ (Folder)
// ក្នុង Google Drive ដដែល (មិនមែនប្តូរទៅគណនី Google ផ្សេងទេ — ការផ្ទេរទិន្នន័យទៅគណនីមួយផ្សេងទៀត ទាមទារការអនុញ្ញាត OAuth
// ដាច់ដោយឡែក ដែលមិនទាន់អនុវត្តនៅឡើយទេ)។ បើមិនកំណត់ទេ ប្រើ Folder ដដែលនឹង Google Sheet មេ (លំនាំដើមចាស់)។
function extractDriveFolderId_(idOrUrl) {
  var s = String(idOrUrl || '').trim();
  if (!s) return '';
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s; // ជា ID ត្រង់ៗរួចហើយ
  return '';
}

// ==== ស្រង់ Spreadsheet ID ចេញពីតំណភ្ជាប់ (URL) ឬ ID ត្រង់ៗ របស់ Google Sheet មេ ដែល Admin បិទភ្ជាប់ (Copy-Paste)
// ក្នុង Tab "ការកំណត់ប្រព័ន្ធ" (setting key: masterSheetLink) — មើល getTargetSheetId_() ក្នុង Utils.gs ====
function extractSpreadsheetId_(idOrUrl) {
  var s = String(idOrUrl || '').trim();
  if (!s) return '';
  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s; // ជា ID ត្រង់ៗរួចហើយ
  return '';
}

// ==== Diagnostic សម្រាប់ទំព័រ "ការកំណត់ប្រព័ន្ធ" ប៉ុណ្ណោះ — ពន្យល់ច្បាស់ថាហេតុអ្វីតំណភ្ជាប់ Google Drive ដែល Admin
// បិទភ្ជាប់ ប្រើមិនកើត (ជំនួសឲ្យការស្ងាត់ស្ងៀមប្តូរទៅរកតាមស្វ័យប្រវត្តិវិញ ដូច getMasterParentFolder_ ធ្វើសម្រាប់ប្រតិបត្តិការ
// ជាក់ស្តែង)។ មូលហេតុទូទៅបំផុតដែលធ្វើឲ្យ DriveApp.getFolderById() បរាជ័យ ទោះជា ID ត្រឹមត្រូវក៏ដោយ៖ ថតនោះមិនទាន់ត្រូវបាន
// ចែករំលែក (Share) ជាមួយគណនី Google ដែល App Script នេះកំពុងដំណើរការក្រោមឈ្មោះ (ជាធម្មតាដូចគណនីដែលបើក Apps Script
// Editor បាន) ====
function diagnoseMasterDriveLink_(masterDriveLink) {
  var customFolderId = extractDriveFolderId_(masterDriveLink);
  if (!customFolderId) return { hasCustomLink: false, error: '' };
  try {
    DriveApp.getFolderById(customFolderId);
    return { hasCustomLink: true, error: '' };
  } catch (err) {
    return {
      hasCustomLink: true,
      error: "⚠️ មិនអាចចូលប្រើថត Google Drive ដែលបានបិទភ្ជាប់នេះបានទេ (ID: " + customFolderId + ")។ " +
        "មូលហេតុទូទៅបំផុត៖ ថតនេះមិនទាន់ត្រូវបានចែករំលែក (Share) ជាមួយគណនី Google ដែលកំពុងដំណើរការ App នេះទេ។ " +
        "សូមបើកថតនោះក្នុង Google Drive ចុច Share រួចបន្ថែមគណនីតែមួយដដែលដែលអ្នកកំពុងប្រើ Login ចូល script.google.com " +
        "(សម្រាប់មើល/កែកូដ App) ជា Editor។ (កំហុសពិត៖ " + (err && err.message ? err.message : String(err)) + ")"
    };
  }
}

function getBackupFolder_() {
  var settings = getSystemSettings_();
  var id = extractDriveFolderId_(settings.backupFolderId);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (err) {}
  }
  return getMasterParentFolder_() || DriveApp.getRootFolder();
}

function createBackup(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិបង្កើតទិន្នន័យបម្រុងទុកទេ!" };
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy_HH:mm');
    var file = DriveApp.getFileById(getTargetSheetId_());
    var folder = getBackupFolder_();
    var copy = file.makeCopy(BACKUP_NAME_PREFIX_ + stamp, folder);
    // ==== កត់ត្រាថាអ្នកណា/តួនាទីអ្វី ជាអ្នកបង្កើតទិន្នន័យបម្រុងទុកនេះ (ដាក់ក្នុងការពិពណ៌នាឯកសារ — listBackups អានវិញនៅពេលក្រោយ) ====
    try { copy.setDescription(JSON.stringify({ by: currentUsername, role: getUserRole_(currentUsername) || '' })); } catch (e) {}
    return { success: true, message: "បង្កើតទិន្នន័យបម្រុងទុកជោគជ័យ!", url: copy.getUrl(), name: copy.getName() };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== ទាញយកទិន្នន័យបម្រុងទុក (Google Sheet មេ) ជាឯកសារ Excel (.xlsx) ចុះមកកុំព្យូទ័រផ្ទាល់ (មិនរក្សាទុកក្នុង Drive ទេ) ====
function exportBackupToComputer(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទាញយកទិន្នន័យបម្រុងទុកទេ!" };
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy_HH-mm'); // ប្រើ - ជំនួស : ព្រោះឈ្មោះឯកសារក្នុងកុំព្យូទ័រមិនអនុញ្ញាត :
    var exportUrl = "https://docs.google.com/spreadsheets/d/" + getTargetSheetId_() + "/export?format=xlsx";
    var res = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      return { success: false, message: "ទាញយកបរាជ័យ (កូដ " + res.getResponseCode() + ")! សូមព្យាយាមម្តងទៀត។" };
    }
    var base64 = Utilities.base64Encode(res.getBlob().getBytes());
    return { success: true, base64: base64, filename: BACKUP_NAME_PREFIX_ + stamp + ".xlsx" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== ត្រឡប់ backupFolderId ដើម (ដែល Admin បានបញ្ចូល) ភ្ជាប់ជាមួយ ដើម្បីឲ្យ Client មិនចាំបាច់ហៅ
// getSystemSettingsAdmin() ដាច់ដោយឡែកទៀត គ្រាន់តែសម្រាប់បំពេញប្រអប់បញ្ចូល Folder ID ក្នុងទំព័រ Backup
// (កាត់បន្ថយ round-trip មួយ ពេលបើកទំព័រ "ទិន្នន័យបម្រុងទុក") ====
function listBackups(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  try {
    var settings = getSystemSettings_();
    var folder = getBackupFolder_();
    var it = folder.getFiles();
    var files = [];
    while (it.hasNext()) {
      var f = it.next();
      if (f.getName().indexOf(BACKUP_NAME_PREFIX_) === 0) {
        var meta = { by: '', role: '' };
        try { var desc = f.getDescription(); if (desc) { var parsed = JSON.parse(desc); if (parsed) meta = parsed; } } catch (e) {}
        files.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), date: f.getDateCreated().toISOString(), by: meta.by || '', role: meta.role || '' });
      }
    }
    files.sort(function(a, b) { return b.date.localeCompare(a.date); });
    return { success: true, backups: files.slice(0, 30), folderUrl: folder.getUrl(), backupFolderId: settings.backupFolderId || '' };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== ស្តារទិន្នន័យបម្រុងទុក (ពី Google Drive) — ត្រូវការសុវត្ថិភាព ដូច្នេះមិនសរសេរជាន់លើ Google Sheet មេផ្ទាល់ទេ
// គ្រាន់តែចម្លងទិន្នន័យបម្រុងទុកនោះ ទៅជា Google Sheet ថ្មីមួយដាច់ដោយឡែក — Admin ត្រូវបើកមើល ហើយផ្លាស់ប្តូរទៅប្រើដោយខ្លួនឯង ====
function restoreBackupToNewSheet(currentUsername, fileId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិស្តារទិន្នន័យទេ!" };
  if (!fileId) return { success: false, message: "សូមជ្រើសរើសទិន្នន័យបម្រុងទុកជាមុនសិន!" };
  try {
    var src = DriveApp.getFileById(fileId);
    var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy_HH:mm');
    var folder = getBackupFolder_();
    var restored = src.makeCopy("ស្តារពី (" + src.getName() + ") - " + stamp, folder);
    return {
      success: true,
      message: "ស្តារទិន្នន័យជោគជ័យ! ប្រព័ន្ធបានបង្កើត Google Sheet ថ្មីមួយដាច់ដោយឡែក (មិនប៉ះពាល់ទិន្នន័យផ្ទាល់បច្ចុប្បន្នទេ) — សូមចុច \"បើក\" ដើម្បីពិនិត្យ រួចផ្លាស់ប្តូរទៅប្រើ Sheet ថ្មីនេះដោយផ្ទាល់ ប្រសិនបើត្រូវការ។",
      url: restored.getUrl(), name: restored.getName()
    };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== ស្តារទិន្នន័យបម្រុងទុក ពីឯកសារ .xlsx ដែល Upload ពីកុំព្យូទ័រ (ក៏បង្កើត Google Sheet ថ្មីដាច់ដោយឡែក ដូចគ្នា — មិនប៉ះពាល់ទិន្នន័យផ្ទាល់ទេ) ====
// ត្រូវការបើក "Drive API" (Advanced Service) ជាមុនសិន ក្នុង Apps Script Editor → Services (ធ្វើតែម្តងគត់) ទើបបំប្លែង
// ឯកសារ .xlsx ទៅជា Google Sheet វិញបាន — បើមិនទាន់បើកទេ នឹងបង្ហាញសារណែនាំឲ្យបើកជាមុន។
function restoreBackupFromUpload(currentUsername, base64Data, fileName, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិស្តារទិន្នន័យទេ!" };
  if (!base64Data) return { success: false, message: "មិនមានទិន្នន័យឯកសារទេ! សូមព្យាយាមជ្រើសរើសឯកសារម្តងទៀត។" };
  if (typeof Drive === 'undefined') {
    return {
      success: false,
      message: "មុខងារនេះទាមទារឲ្យបើក \"Drive API\" (Advanced Service) ជាមុនសិន៖ ក្នុង Apps Script Editor → ចុច \"Services\" (សញ្ញា +) នៅផ្នែកម្ខាង → ជ្រើសរើស \"Drive API\" → ចុច Add ។ ការនេះធ្វើតែម្តងគត់ សូមស្នើសុំអ្នកបច្ចេកទេសជួយកំណត់ជំហាននេះជាមុនសិន។"
    };
  }
  try {
    var bytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, fileName || "backup.xlsx");
    var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
    var stamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy_HH:mm');
    var folder = getBackupFolder_();
    var fileMeta = { name: "ស្តារពីកុំព្យូទ័រ - " + stamp, mimeType: MimeType.GOOGLE_SHEETS, parents: [folder.getId()] };
    var created = Drive.Files.create(fileMeta, blob);
    return {
      success: true,
      message: "ស្តារទិន្នន័យពីកុំព្យូទ័រជោគជ័យ! ប្រព័ន្ធបានបង្កើត Google Sheet ថ្មីមួយដាច់ដោយឡែក (មិនប៉ះពាល់ទិន្នន័យផ្ទាល់បច្ចុប្បន្នទេ) — សូមចុច \"បើក\" ដើម្បីពិនិត្យ រួចផ្លាស់ប្តូរទៅប្រើ Sheet ថ្មីនេះដោយផ្ទាល់ ប្រសិនបើត្រូវការ។",
      url: "https://docs.google.com/spreadsheets/d/" + created.id + "/edit", name: fileMeta.name
    };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message + " (សូមប្រាកដថាបានបើក \"Drive API\" Advanced Service រួចហើយ)" };
  }
}

function deleteBackup(currentUsername, fileId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិលុបទិន្នន័យបម្រុងទុកទេ!" };
  if (!fileId) return { success: false, message: "ទិន្នន័យមិនត្រឹមត្រូវទេ!" };
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true, message: "លុបទិន្នន័យបម្រុងទុកជោគជ័យ!" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==================== ១២. តំបន់គ្រោះថ្នាក់ (Danger Zone) — កំណត់ Spreadsheet ស្រុកទាំង១០ ឡើងវិញ ====================
// ប្រើសម្រាប់ចាប់ផ្តើមវដ្តចុះឈ្មោះថ្មី (ឧ. ឆមាស/ឆ្នាំថ្មី)៖ Admin ចង់ឲ្យ Spreadsheet ស្រុកនីមួយៗ (១០ស្រុក) ចាប់ផ្តើម
// ស្អាតឡើងវិញពីដើម។ ជំហាន៖ (១) ចម្លងទុក Spreadsheet ចាស់នីមួយៗទៅក្នុងថត Backup ស្រាប់ (getBackupFolder_ — កន្លែងតែមួយ
// ជាមួយ Backup Sheet មេ ងាយស្វែងរក) ជាមុនសិន — កុំពឹងផ្អែកតែលើ Google Drive Trash ដែលអាចត្រូវលុបចោលជាអចិន្ត្រៃយ៍ដោយ
// ស្វ័យប្រវត្តិក្រោយ៣០ថ្ងៃ ឬបើនរណាម្នាក់សម្អាតធុងសំរាមចោលមុនកំណត់ — (២) ទើបផ្ទេរ Spreadsheet ដើមទៅធុងសំរាម (៣) លុប
// Script Property ចាស់ + Trigger onEdit ចាស់ ដើម្បីឲ្យ getDistrictSpreadsheet_() បង្កើត Spreadsheet ថ្មីស្អាតដោយ
// ស្វ័យប្រវត្តិ (ជាមួយក្បាលតារាង/រូបមន្តត្រឹមត្រូវទាំងអស់) នៅពេលមានការបញ្ចូលទិន្នន័យលើកក្រោយ។ ទាមទារ Admin ជាក់ស្តែង
// + វាយបញ្ជាក់ឃ្លាត្រឹមត្រូវ (ការពារការចុចខុសដោយចៃដន្យ លើប្រតិបត្តិការមិនអាចត្រឡប់វិញបានដូចដើម ១០០%នេះ) ====
var RESET_DISTRICTS_CONFIRM_PHRASE_ = "លុបចោលទាំងអស់";
var DISTRICT_ARCHIVE_NAME_PREFIX_ = "ចម្លងទុកមុនកំណត់ឡើងវិញ_ស្រុក_";

function resetAllDistrictSpreadsheets(currentUsername, confirmPhrase, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សុវត្ថិភាព — ភាពមិនស៊ីសង្វាក់គ្នានៃសិទ្ធិ)៖ ប្រតិបត្តិការនេះលុប/ចម្លងទុក Google Sheet ស្រុកទាំង១០ថ្មីទាំងស្រុង
  // (មិនអាចត្រឡប់វិញបានទេ ចំពោះទិន្នន័យផ្ទាល់) ដូចគ្នានឹងកម្រិតគ្រោះថ្នាក់របស់ resetAllData() ខាងលើ ដែលកំណត់ឲ្យតែ
  // SuperAdmin ប៉ុណ្ណោះ — ប៉ុន្តែ Function នេះពីមុនប្រើ isAdmin_() ដែលរួមបញ្ចូល PEC21 ផងដែរ (តាមសំណើមុនដែលលើកកម្ពស់
  // PEC21 ឲ្យមានសិទ្ធិដូច Admin ទាំងស្រុងសម្រាប់ការងារធម្មតា) ធ្វើឲ្យ PEC21 តែម្នាក់ឯង អាចលុបទិន្នន័យស្រុកទាំង១០បាន
  // ដោយមិនចាំបាច់ឆ្លងកាត់ SuperAdmin ទាល់តែសោះ — តាមការសម្រេចរបស់អ្នកគ្រប់គ្រងប្រព័ន្ធ កំណត់ឲ្យតែ SuperAdmin ធ្វើបាន ====
  if (!isSuperAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិធ្វើសកម្មភាពនេះទេ! (សម្រាប់តែ SuperAdmin ប៉ុណ្ណោះ)" };
  if (String(confirmPhrase || '').trim() !== RESET_DISTRICTS_CONFIRM_PHRASE_) {
    return { success: false, message: "អក្សរបញ្ជាក់មិនត្រឹមត្រូវទេ! សូមវាយ \"" + RESET_DISTRICTS_CONFIRM_PHRASE_ + "\" ឲ្យបានត្រឹមត្រូវ (គ្មានចន្លោះ/អក្សរខុសសោះ) ដើម្បីបញ្ជាក់ថាអ្នកយល់ព្រម។" };
  }

  var effectiveId = getTargetSheetId_();
  var props = PropertiesService.getScriptProperties();
  var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
  var stamp = Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy_HH:mm');
  var folder;
  try { folder = getBackupFolder_(); } catch (eF) { folder = null; }

  var existingTriggers;
  try { existingTriggers = ScriptApp.getProjectTriggers(); } catch (eT) { existingTriggers = []; }

  var results = [];
  DISTRICT_LIST.forEach(function(district) {
    var key = DISTRICT_SS_PROP_PREFIX + effectiveId + '_' + district;
    var ssId = props.getProperty(key);
    if (!ssId) { results.push({ district: district, status: 'skipped', reason: 'គ្មាន Spreadsheet ស្រាប់ទេ (មិនទាន់បង្កើត)' }); return; }
    try {
      var file = DriveApp.getFileById(ssId);
      var archivedUrl = '';
      if (folder) {
        try {
          var copy = file.makeCopy(DISTRICT_ARCHIVE_NAME_PREFIX_ + district + '_' + stamp, folder);
          try { copy.setDescription(JSON.stringify({ by: currentUsername, role: getUserRole_(currentUsername) || '', originalId: ssId })); } catch (eDesc) {}
          archivedUrl = copy.getUrl();
        } catch (eCopy) {
          // ==== បើចម្លងទុកមិនកើត (ឧ. គ្មានទំហំ Drive សល់) — ត្រូវបញ្ឈប់ភ្លាមៗ មិនបន្តទៅផ្ទេរធុងសំរាមឡើយ សម្រាប់
          // ស្រុកនេះ (សុវត្ថិភាព៖ កុំចោលទិន្នន័យដោយគ្មាន Backup ត្រឹមត្រូវ) ====
          results.push({ district: district, status: 'error', reason: 'ចម្លងទុក Backup មិនកើត៖ ' + eCopy.message + ' (មិនបានផ្ទេរទៅធុងសំរាមទេ — ទិន្នន័យដើមនៅដដែល)' });
          return;
        }
      }
      // ==== លុប Trigger onEdit ចាស់ដែលភ្ជាប់នឹង Spreadsheet នេះ (មុននឹងផ្ទេរទៅធុងសំរាម) — ជៀសវាង Trigger ក្លាយជា
      // "សំណល់" ឥតប្រយោជន៍ (ប៉ះពាល់កូតា ២០ Trigger/User/Project ក្នុងរយៈពេលវែង បើកំណត់ឡើងវិញច្រើនលើក) ====
      (existingTriggers || []).forEach(function(t) {
        try {
          if (t.getHandlerFunction() === 'onDistrictSheetEdit_' && t.getTriggerSourceId() === ssId) {
            ScriptApp.deleteTrigger(t);
          }
        } catch (eDel) {}
      });
      file.setTrashed(true);
      props.deleteProperty(key);
      delete DISTRICT_SS_CACHE_[effectiveId + '|' + district];
      results.push({ district: district, status: 'reset', oldId: ssId, archivedUrl: archivedUrl });
    } catch (err) {
      results.push({ district: district, status: 'error', reason: err.message });
    }
  });

  var resetCount = results.filter(function(r) { return r.status === 'reset'; }).length;
  var errCount = results.filter(function(r) { return r.status === 'error'; }).length;
  var skipCount = results.filter(function(r) { return r.status === 'skipped'; }).length;
  var msg = "បានកំណត់ស្រុកចំនួន " + resetCount + "/" + DISTRICT_LIST.length + " ជោគជ័យ។";
  if (skipCount) msg += " (" + skipCount + "ស្រុក រំលង ព្រោះគ្មាន Spreadsheet ស្រាប់)";
  if (errCount) msg += " ⚠️ " + errCount + "ស្រុក មានបញ្ហា — សូមពិនិត្យលម្អិតខាងក្រោម (ទិន្នន័យស្រុកទាំងនោះមិនត្រូវបានប៉ះពាល់ទេ)។";
  if (folder && resetCount) msg += " ច្បាប់ចម្លងទុកមុនកំណត់ឡើងវិញ ត្រូវបានរក្សាទុកក្នុងថត Backup ដដែល។";
  return { success: true, message: msg, results: results, backupFolderUrl: folder ? folder.getUrl() : '' };
}

// ==================== ៦.១ តំបន់គ្រោះថ្នាក់ — លុបគណនីទាំងអស់ ទុកតែ SuperAdmin (Full Reset) ====================
// ==== FIX (សំណើថ្មី "Full Reset — លុប​គណនីទាំងអស់ ទុកតែ SuperAdmin")៖ សម្រាប់ត្រៀមប្រព័ន្ធថ្មីមួយសម្រាប់ការប្រើប្រាស់
// ផ្លូវការ (ឧ. ក្រោយចម្លង/ដាក់ Google Sheet ទៅទីតាំងថ្មី — Sheet ចាស់ទុកសម្រាប់តេស្ត/បណ្តុះបណ្តាលដដែល) — លុបគណនី
// Admin/PEC21/គ្រូប្រចាំក្រុងស្រុក/ឃុំសង្កាត់/សង្កេតការណ៍ *ទាំងអស់* ចេញភ្លាមៗ ទុកសល់តែជួរដេកតួនាទី SuperAdmin ប៉ុណ្ណោះ
// (រួមទាំង SuperAdmin ដទៃទៀត មិនត្រឹមតែអ្នកហៅផ្ទាល់ទេ)។ កម្រិតគ្រោះថ្នាក់ស្មើនឹង resetAllData()/
// resetAllDistrictSpreadsheets() ខាងលើ — កំណត់ឲ្យតែ SuperAdmin ប៉ុណ្ណោះ (មិនមែន isAdmin_() ទូទៅ ដែលរាប់បញ្ចូល
// Admin/PEC21 ផងដែរ) ព្រមទាំងតម្រូវឲ្យវាយបញ្ជាក់ឃ្លាដូចគ្នាដែរ ដើម្បីជៀសវាងការចុចខុសដោយចៃដន្យ។ ចំណាំ៖ មុខងារនេះប៉ះពាល់
// តែ Sheet(អ្នកប្រើប្រាស់) ប៉ុណ្ណោះ — មិនប៉ះពាល់ទិន្នន័យចុះឈ្មោះប្រចាំថ្ងៃ/ទិន្នន័យយោង/បញ្ជីឃុំសង្កាត់ឡើយ (ប្រើ
// resetAllData()/resetAllDistrictSpreadsheets() ដាច់ដោយឡែកសម្រាប់ទិន្នន័យទាំងនោះ) ====
var RESET_ALL_USERS_CONFIRM_PHRASE_ = "លុបគណនីទាំងអស់";

function resetAllUserAccountsExceptSuperAdmin(currentUsername, confirmPhrase, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isSuperAdmin_(currentUsername)) {
    return { success: false, message: "គ្មានសិទ្ធិធ្វើសកម្មភាពនេះទេ! (សម្រាប់តែ SuperAdmin ប៉ុណ្ណោះ)" };
  }
  if (String(confirmPhrase || '').trim() !== RESET_ALL_USERS_CONFIRM_PHRASE_) {
    return { success: false, message: "អក្សរបញ្ជាក់មិនត្រឹមត្រូវទេ! សូមវាយ \"" + RESET_ALL_USERS_CONFIRM_PHRASE_ + "\" ឲ្យបានត្រឹមត្រូវ (គ្មានចន្លោះ/អក្សរខុសសោះ) ដើម្បីបញ្ជាក់ថាអ្នកយល់ព្រម។" };
  }

  var ss = getSS_();
  var sheet = ensureUsersSheet_(ss);
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { gotLock = lock.tryLock(20000); } catch (eLock) {}
  if (!gotLock) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់! សូមព្យាយាមម្តងទៀត។" };
  }
  var deletedCount = 0, keptCount = 0;
  try {
    var data = sheet.getDataRange().getDisplayValues();
    var totalCols = data[0].length;
    var keptRows = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][4] === ROLE_SUPERADMIN) { keptRows.push(data[i]); keptCount++; }
      else deletedCount++;
    }
    var lastRow = sheet.getLastRow();
    // ==== សម្អាតជួរដេកទិន្នន័យទាំងអស់ (រក្សាទុកក្បាលតារាងជួរទី១) សិន រួចសរសេរជួរដេក SuperAdmin ដែលរក្សាទុកវិញ
    // ជាការហៅ setValues() តែម្តងគត់ (Batch) — ជៀសវាងហៅ deleteRow() ច្រើនដងឡែកៗពីគ្នា (យឺត + ប្រឈមបញ្ហាលេខ
    // ជួរដេកប្តូរខណៈកំពុងលុប) ====
    if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, totalCols).clearContent();
    if (keptRows.length) sheet.getRange(2, 1, keptRows.length, totalCols).setValues(keptRows);
  } finally {
    try { lock.releaseLock(); } catch (eRel) {}
  }
  // ==== ដកសិទ្ធិ Google Sheet ភ្លាមៗ (Synchronous — មិនប្តូរទៅ scheduleSyncSheetPermissions_() ទេ) ដូចគ្នានឹង
  // deleteUserAccount()/toggleUserSuspend() — ហេតុផលសុវត្ថិភាពដូចគ្នា (គណនីដែលទើបលុប មិនត្រូវនៅតែបើក Google
  // Sheet ដោយផ្ទាល់កែប្រែបានទៀតឡើយ សូម្បីតែរយៈពេលខ្លីមួយ) ====
  try { syncSheetPermissions_(); } catch (err) {}
  var msg = "បានលុបគណនីទាំងអស់ក្រៅពី SuperAdmin ជោគជ័យ! លុបចំនួន " + deletedCount + " គណនី (សល់ SuperAdmin " + keptCount + " គណនី)";
  try { notifyTelegram_("⚠️🗑️ SuperAdmin (" + currentUsername + ") បានលុបគណនីទាំងអស់ក្រៅពី SuperAdmin ចេញ!\nលុបចំនួន " + deletedCount + " គណនី (សល់ SuperAdmin " + keptCount + " គណនី)"); } catch (errNotify) {}
  return { success: true, deleted: deletedCount, kept: keptCount, message: msg };
}

// ==== ធ្វើបច្ចុប្បន្នភាព Tab សរុប៣ (Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ)/Sheet(របាយការណ៍សរុបតាមឃុំសង្កាត់)/
// Sheet(របាយការណ៍សរុបខេត្ត)) ភ្លាមៗ តាមការស្នើសុំរបស់ Admin ជំនួសរង់ចាំ Trigger ស្វ័យប្រវត្តិ (scheduleRebuild_ —
// ធម្មតារត់ក្រោយរក្សាទុកទិន្នន័យប្រចាំថ្ងៃ ~4វិនាទីក្រោយ)។ មិនមែនសកម្មភាពគ្រោះថ្នាក់ទេ (មិនប៉ះពាល់ Period Sheet
// ប្រភពទិន្នន័យពិតប្រាកដសោះឡើយ — គ្រាន់តែសរសេរ Tab សរុបដែលអាចគណនាឡើងវិញបានគ្រប់ពេលឡើងវិញតែប៉ុណ្ណោះ) ប៉ុន្តែ
// ជាការងារធ្ងន់ (ស្កេន Spreadsheet ស្រុកទាំង១០) ដូច្នេះអាចចំណាយពេលពីរបីវិនាទីទៅច្រើននាទី អាស្រ័យលើបរិមាណទិន្នន័យ ====
function forceRebuildDerivedDataNow(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិធ្វើសកម្មភាពនេះទេ! (សម្រាប់តែ Admin ប៉ុណ្ណោះ)" };
  try {
    rebuildAllDerivedData_();
    return { success: true, message: "បានធ្វើបច្ចុប្បន្នភាព Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ)/Tab(របាយការណ៍សរុបតាមឃុំសង្កាត់)/Tab(របាយការណ៍សរុបខេត្ត) ជោគជ័យ!" };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// អានលេខកូដឃុំ/សង្កាត់ (ឧ. "21-001") ដែលមានស្រាប់ក្នុង COMMUNE_ORDER — មិនចាំបាច់វាយបញ្ចូលដោយដៃថ្មីទេ
// អាន ស្ថិតិប៉ាន់ស្មាន/ការិ.បង្កើតថ្មី/ចំនួនការិ.សរុប ពី Sheet(ទិន្នន័យភ្ជាប់) ដោយផ្ទាល់ (គ្រប់ស្រុក/ឃុំសង្កាត់ ក្នុងការអានតែម្តង)
