// ==================== PeriodSheets.gs — ការបង្កើត/ការពារ Period Sheet, រូបមន្ត, ការកែប្រែទិន្នន័យប្រចាំថ្ងៃ ====================

// ==== ត្រឡប់តំណភ្ជាប់ Google Sheet មេ + ថត Google Drive (SuperAdmin/Admin ប៉ុណ្ណោះ) — ដើម្បីងាយស្រួល Copy ប្រើប្រាស់
// ចំណាំសំខាន់៖ លែងគ្រាន់តែជាតំណភ្ជាប់សម្រាប់បង្ហាញ (Display) ទៀតហើយ — បើ Admin បានបិទភ្ជាប់ (Copy-Paste) តំណភ្ជាប់ផ្ទាល់ខ្លួន
// ទុកជាមុនរួច (Tab "ការកំណត់ប្រព័ន្ធ") ប្រព័ន្ធទាំងមូលនឹងប្តូរទៅអាន/សរសេរទិន្នន័យ​ពី Google Sheet នោះជាក់ស្តែងភ្លាមៗ
// (មើល getTargetSheetId_() និង getMasterParentFolder_() ក្នុង Utils.gs — ចំណុចនេះជា Function ត្រឡប់តំណភ្ជាប់
// សម្រាប់បង្ហាញនៅ Tab ការកំណត់ប្រព័ន្ធតែប៉ុណ្ណោះ ការគណនាជាក់ស្តែងទាំងអស់ស្ថិតនៅក្នុង Utils.gs) ====
function getMasterLinks(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិចូលមើលទិន្នន័យនេះទេ!" };
  var settings = getSystemSettings_();
  var masterSheetUrl = "https://docs.google.com/spreadsheets/d/" + getTargetSheetId_() + "/edit";
  var folderUrl = "";
  try {
    var folder = getMasterParentFolder_();
    if (folder) folderUrl = folder.getUrl();
  } catch (err) {}
  // ==== FIX៖ ពេលមានតំណភ្ជាប់ផ្ទាល់ខ្លួនបិទភ្ជាប់ (masterDriveLink) ប៉ុន្តែប្រព័ន្ធនៅតែរកមិនឃើញថត (folderUrl ទទេ)
  // ត្រូវប្រាប់ Admin ច្បាស់ថាហេតុអ្វី (ជាធម្មតា Share Permission) ជំនួសឲ្យបង្ហាញសារទូទៅមិនច្បាស់លាស់ ====
  var folderError = "";
  if (!folderUrl) {
    var diag = diagnoseMasterDriveLink_(settings.masterDriveLink);
    if (diag.error) folderError = diag.error;
  }
  return {
    success: true, masterSheetUrl: masterSheetUrl, folderUrl: folderUrl, folderError: folderError,
    isCustomSheetLink: !!extractSpreadsheetId_(settings.masterSheetLink),
    isCustomDriveLink: !!extractDriveFolderId_(settings.masterDriveLink)
  };
}

// ==== FIX (Fix118, "ដំណើរការ Login Form ដើរយឺត")៖ សម្រាប់គ្រូប្រចាំក្រុងស្រុក (District-tier) — គណនីភាគច្រើនប្រើប្រាស់
// ក្នុងប្រព័ន្ធ — getDistrictSpreadsheet_() ខាងក្រោម (ក្នុង Branch isDistrictRole_) បើក Spreadsheet ដាច់ដោយឡែក
// (SpreadsheetApp.openById) របស់ស្រុកខ្លួន ដែលកកកុញ Tab ថ្ងៃរាប់រយកន្លងមក (ចាប់ផ្តើមប្រើប្រាស់តាំងពីដើមប្រព័ន្ធ) —
// ជាប្រតិបត្តិការយឺតបំផុតមួយក្នុង Google Apps Script។ ចំណុចនេះត្រូវបានហៅរាល់ពេល Login (loadSheetAccessButton() ក្នុង
// applyLoginSession()) គ្រាន់តែដើម្បីគណនា URL ប៊ូតុង "បើក Google Sheet" (មិនសំខាន់បន្ទាន់ - Feature ងាយស្រួលបន្ថែម
// មិនមែនទិន្នន័យស្នូលទេ) ។ ដំណោះស្រាយ៖ Cache លទ្ធផលចុងក្រោយ (URL+Label) ទុកតាម (ស្រុក + ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន)
// អតិបរមា ៦ម៉ោង (ដូចគ្នានឹងគោលការណ៍ getPeriodSheetDateMap_ ខាងលើ) — លុបចោលភ្លាមៗពេលមាន Tab ថ្ងៃថ្មីត្រូវបានបង្កើត
// សម្រាប់ស្រុកនោះ (មើល clearSheetAccessCache_() ហៅពី createDistrictPeriodSheet_) ដូច្នេះ Login ភាគច្រើន (ក្នុងចន្លោះ
// ៦ម៉ោងតែមួយ) លែងបើក Spreadsheet ស្រុកសោះ (លឿនភ្លាមៗ) ====
var CACHE_KEY_SHEET_ACCESS_PREFIX_ = 'sheetAccess_v1_';
var CACHE_TTL_SHEET_ACCESS_ = 21600; // វិនាទី (៦ ម៉ោង — អតិបរមា CacheService អនុញ្ញាត)
function sheetAccessCacheKey_(district) {
  return CACHE_KEY_SHEET_ACCESS_PREFIX_ + getTargetSheetId_() + '_' + district + '_' + getCurrentPeriodDate_();
}
function clearSheetAccessCache_(district) {
  try { CacheService.getScriptCache().remove(sheetAccessCacheKey_(district)); } catch (err) {}
}

function getSheetAccessInfo(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គណនីកម្រិតឃុំសង្កាត់ គ្មានសិទ្ធិចូល Google Sheet ណាមួយដោយផ្ទាល់ទាល់តែសោះ
  // (បញ្ចូល/មើលទិន្នន័យបានតែតាមរយៈ App ប៉ុណ្ណោះ តាមការកំណត់) — ត្រឡប់ url ទទេ ដើម្បីឲ្យ Client លាក់ប៊ូតុង "បើក Google
  // Sheet" ដោយស្វ័យប្រវត្តិ (មើល loadSheetAccessButton() ក្នុង Index.html — ធ្លាប់លាក់រួចស្រាប់ពេល url ទទេ) ====
  if (isCommuneRole_(role)) {
    return { success: true, url: "", label: "", noSheetAllowed: true };
  }
  var ss = getSS_();
  var sheet, label, spreadsheetId;
  if (isDistrictRole_(role)) {
    // ==== THE FIX (Fix118)៖ ត្រួតពិនិត្យ Cache មុននឹងបើក Spreadsheet ស្រុក (ថ្លៃបំផុត) ====
    var cKey = sheetAccessCacheKey_(role);
    try {
      var cachedRes = CacheService.getScriptCache().get(cKey);
      if (cachedRes) return JSON.parse(cachedRes);
    } catch (errCacheGet) {}

    var dSs = getDistrictSpreadsheet_(role); // ធានាថា Spreadsheet ស្រុកមានស្រាប់ (មិនប៉ះពាល់ Tab ណាមួយឡើយ)
    var opDateStr = getCurrentPeriodDate_(); // ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន (មិនមែនថ្ងៃពិតតាម Server ទេ)
    sheet = dSs.getSheetByName(periodSheetName_(opDateStr));
    var resultForCache;
    if (!sheet) {
      // ==== សំខាន់៖ លែងបង្កើត Tab ថ្ងៃនេះដោយស្វ័យប្រវត្តិទៀតហើយ ====
      // មុននេះ ការចុចប៊ូតុងនេះ (ឬសូម្បីតែគ្រាន់តែ Login) ធ្វើឲ្យ Tab ថ្ងៃនេះកកើតឡើងភ្លាមៗ ដោយគ្មានការអនុញ្ញាតពី Admin ។
      // ឥឡូវ Tab ត្រូវបង្កើតដោយ Admin/SuperAdmin/PEC21 ប៉ុណ្ណោះ (ចុច "បង្កើត Sheet ថ្មីបន្ទាប់") ។
      resultForCache = { success: true, url: "", label: "ថ្ងៃនេះមិនទាន់មាន Sheet ទេ — សូមទាក់ទង Admin ដើម្បីបង្កើតជាមុនសិន", noSheetYet: true };
    } else {
      spreadsheetId = dSs.getId();
      label = 'បញ្ចូលទិន្នន័យផ្ទាល់ក្នុង Google Sheet — ស្រុក ' + role + ' (Tab ថ្ងៃទី "' + periodSheetName_(opDateStr) + '")';
      resultForCache = { success: true, url: buildSheetTabUrl_(spreadsheetId, sheet), label: label };
    }
    try { CacheService.getScriptCache().put(cKey, JSON.stringify(resultForCache), CACHE_TTL_SHEET_ACCESS_); } catch (errCachePut) {}
    return resultForCache;
  } else {
    sheet = ensureDailySheet_(ss);
    spreadsheetId = getTargetSheetId_();
    label = 'របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ (សរុបគ្រប់ស្រុក — មើលបានប៉ុណ្ណោះ)';
  }
  return { success: true, url: buildSheetTabUrl_(spreadsheetId, sheet), label: label };
}

var SHEET_EDIT_UNLOCKS = "ការអនុញ្ញាតកែថ្ងៃចាស់";
var UNLOCK_ALL_DISTRICTS_ = "ទាំងអស់ (គ្រប់ស្រុក)";
var CACHE_KEY_UNLOCKS_ = "editUnlocks_v1";

function ensureEditUnlocksSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_EDIT_UNLOCKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EDIT_UNLOCKS);
    var headers = ["ID", "ស្រុក/ក្រុង", "កាលបរិច្ឆេទ", "បើកដោយ", "ពេលបើក"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f59e0b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUnlockList_() {
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY_UNLOCKS_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var ss = getSS_();
  var sheet = ensureEditUnlocksSheet_(ss);
  var list = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    data.forEach(function(r) {
      if (!r[0]) return;
      list.push({ id: r[0], district: r[1], date: normalizeDateStr_(r[2]), grantedBy: r[3], grantedAt: r[4] });
    });
  }
  try { CacheService.getScriptCache().put(CACHE_KEY_UNLOCKS_, JSON.stringify(list), CACHE_TTL_); } catch (err) {}
  return list;
}

function isDateUnlockedForDistrict_(district, dateStr) {
  var list = getUnlockList_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].date !== dateStr) continue;
    if (list[i].district === UNLOCK_ALL_DISTRICTS_ || list[i].district === district) return true;
  }
  return false;
}

// ==== "ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន" — ខុសពី "ថ្ងៃពិត" តាមកុំព្យូទ័រ/Server ====
// ប្រព័ន្ធនេះ ដំណើរការទៅតាមការសម្រេចចិត្តរបស់ Admin (ចុច "បង្កើត Sheet ថ្មីបន្ទាប់") មិនមែនតាមម៉ោង/ថ្ងៃពិតទេ
// (ជាពិសេសនៅពេលសាកល្បង ឬប្រើកាលបរិច្ឆេទមិនត្រូវនឹងប្រតិទិនជាក់ស្តែង)។ ដូច្នេះ "ថ្ងៃចាស់ត្រូវចាក់សោ" ត្រូវផ្អែក
// លើកាលបរិច្ឆេទ Sheet ចុងក្រោយដែល Admin បានបង្កើត ជា "បច្ចុប្បន្ន" ជំនួសឲ្យប្រៀបធៀបនឹងម៉ោង Server ។
var CURRENT_PERIOD_DATE_PROP_ = "currentPeriodDate_v1";

function getCurrentPeriodDate_() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(CURRENT_PERIOD_DATE_PROP_);
    if (v && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  } catch (err) {}
  return todayDateStr_(); // លើកដំបូង (មិនទាន់កំណត់) — ប្រើថ្ងៃពិតតាម Server ជាលំនាំដើម
}

function setCurrentPeriodDate_(dateStr) {
  try {
    var d = normalizeDateStr_(dateStr);
    if (!d) return;
    var cur = getCurrentPeriodDate_();
    if (d > cur) PropertiesService.getScriptProperties().setProperty(CURRENT_PERIOD_DATE_PROP_, d);
  } catch (err) {}
}

// ==== FIX (សំណើថ្មី "ការបង្កើត Sheet ត្រូវរៀបតាមលំដាប់ថ្ងៃខែ — មិនត្រូវអាស្រ័យលើកាលបរិច្ឆេទណាមួយក្នុង App ដាក់តាម
// ចិត្តទេ")៖ ត្រឡប់កាលបរិច្ឆេទ Period Sheet "ចុងក្រោយបំផុត" ជាក់ស្តែង ដោយស្កែនផ្ទាល់ពី Sheet ទាំងអស់ដែលមានស្រាប់ (គ្រប់
// ស្រុកទាំង១០ — យក Max ចំពោះកាលបរិច្ឆេទ) ជំនួសការពឹងផ្អែកលើ CURRENT_PERIOD_DATE_PROP_ (Property ខាងលើ) ដែលមានចំណុច
// ខ្សោយមួយ៖ setCurrentPeriodDate_() ខាងលើ កត់ត្រាតែពេល d > cur ប៉ុណ្ណោះ (ចេតនា — ដើម្បីកុំឲ្យ "ថ្ងៃប្រតិបត្តិការ
// បច្ចុប្បន្ន" ថយក្រោយវិញ) ដូច្នេះបើ Sheet ដំបូងបំផុតត្រូវបានបង្កើតសម្រាប់ថ្ងៃ ≤ ថ្ងៃពិតតាម Server (ករណីធម្មតាបំផុត —
// ជាធម្មតា Admin បង្កើត Sheet សម្រាប់ "ថ្ងៃនេះ") Property នេះនឹងមិនត្រូវបានកត់ត្រាទាល់តែសោះ ធ្វើឲ្យមិនអាចទុកចិត្តវា
// ១០០% សម្រាប់កំណត់ "តើ Sheet ចុងក្រោយពិតប្រាកដ ជាថ្ងៃណា" បានទេ។ Function នេះស្កែនផ្ទាល់ពី getPeriodSheetDateMap_()
// (ដែល Cache រួចហើយ — មិនចំណាយពេលបន្ថែមទេ) ធានាថាត្រឹមត្រូវ១០០% ជានិច្ច ទោះក្នុងស្ថានភាពណាក៏ដោយ ====
function getLatestExistingPeriodDate_() {
  var latest = null;
  DISTRICT_LIST.forEach(function(district) {
    var dateMap = getPeriodSheetDateMap_(district);
    Object.keys(dateMap).forEach(function(name) {
      var d = dateMap[name];
      if (d && (!latest || d > latest)) latest = d;
    });
  });
  return latest; // null = មិនទាន់មាន Period Sheet ណាមួយសោះ (មិនទាន់ចាប់ផ្តើមប្រតិបត្តិការ)
}

// ត្រឡប់ { isFirst, currentDate, nextDate } — "ថ្ងៃត្រឹមត្រូវបន្ទាប់" ដែល Sheet ថ្មីត្រូវបង្កើត (Sheet ចុងក្រោយ + ១ថ្ងៃ)
// ដោយអាស្រ័យលើ Sheet ចាស់ចុងក្រោយពិតប្រាកដប៉ុណ្ណោះ (មិនអាស្រ័យលើកាលបរិច្ឆេទណាមួយផ្សេងទៀតក្នុង App ទេ — ឧ. dEntryDate)
// ==== FIX (សំណើថ្មី "ថ្ងៃប្រជុំ/ថ្ងៃគ្មានទិន្នន័យ")៖ "ចំណុចយោងលំដាប់" (pointer) ឥឡូវយក Max រវាង Sheet ចាស់
// ចុងក្រោយ និង ថ្ងៃដែលបានសម្គាល់ថារំលងចុងក្រោយ (getLatestSkippedDate_) — ដូច្នេះបើថ្ងៃ 17-09 ត្រូវបានសម្គាល់ថា
// ជាថ្ងៃប្រជុំ (រំលង) ខណៈ Sheet ពិតប្រាកដចុងក្រោយនៅតែជា 16-09 នោះ "ថ្ងៃត្រឹមត្រូវបន្ទាប់" (nextDate) នឹងក្លាយជា
// 18-09 ដោយស្វ័យប្រវត្តិ (មិនតម្រូវឲ្យបង្កើត Sheet ថ្ងៃ 17-09 ជាមុនសិនទេ) — currentDate (សម្រាប់ "ជួសជុលរូបមន្ត")
// នៅតែជា Sheet ពិតប្រាកដចុងក្រោយជានិច្ច (ថ្ងៃសម្គាល់រំលង គ្មាន Sheet ឲ្យជួសជុលឡើយ) ====
function getNextExpectedPeriodDate_() {
  var latestSheet = getLatestExistingPeriodDate_();
  var latestSkip = getLatestSkippedDate_();
  var pointer = latestSheet;
  if (latestSkip && (!pointer || latestSkip > pointer)) pointer = latestSkip;
  if (!pointer) {
    var startDate = getOperationStartDate_();
    return { isFirst: true, currentDate: null, nextDate: startDate || todayDateStr_() };
  }
  return {
    isFirst: false, currentDate: latestSheet, nextDate: addDaysToDateStr_(pointer, 1),
    lastSkippedDate: (latestSkip && latestSkip === pointer) ? latestSkip : null
  };
}

// ត្រូវហៅពី Client មុននឹងបើកប្រអប់ "បង្កើត Sheet ថ្មីបន្ទាប់" — ដើម្បីបង្ហាញ/ណែនាំកាលបរិច្ឆេទត្រឹមត្រូវបន្ទាប់ ជំនួស
// ការប្រើ dEntryDate (ដែលអាចជាកាលបរិច្ឆេទផ្សេងទាំងស្រុង ដែលអ្នកប្រើកំពុងមើលទិន្នន័យ មិនទាក់ទងនឹង Sheet ចុងក្រោយ
// ដែលត្រូវបង្កើតបន្តទេ — នេះជាមូលហេតុពិតនៃបញ្ហា "ការទាញទិន្នន័យយោងមិនត្រឹមត្រូវ + យូរបំផុត" ដែលរាយការណ៍មក)
function getNextPeriodSheetInfo(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (role !== ROLE_SUPERADMIN && role !== ROLE_ADMIN && role !== ROLE_PEC21) {
    return { success: false, message: "មានតែ SuperAdmin, Admin, ឬ PEC21 ប៉ុណ្ណោះ ដែលអាចបង្កើត Sheet ថ្មីនេះបាន!" };
  }
  var info = getNextExpectedPeriodDate_();
  return {
    success: true, isFirst: info.isFirst, currentDate: info.currentDate, nextDate: info.nextDate,
    lastSkippedDate: info.lastSkippedDate || null
  };
}

// ==== FIX (សំណើថ្មី "ថ្ងៃប្រជុំ/ថ្ងៃដែលមិនត្រូវបញ្ចូលទិន្នន័យ")៖ អនុញ្ញាតឲ្យ SuperAdmin/Admin/PEC21 សម្គាល់ថ្ងៃជាក់លាក់
// ណាមួយថា "គ្មានទិន្នន័យ" (ឧ. ថ្ងៃប្រជុំ) ដោយមិនចាំបាច់បង្កើត Period Sheet សម្រាប់ថ្ងៃនោះទេ ប៉ុន្តែលំដាប់នៃការបង្កើត
// Sheet បន្តបន្ទាប់ (getNextExpectedPeriodDate_) នៅតែឆ្លងកាត់/រំលងថ្ងៃនេះដោយស្វ័យប្រវត្តិ (ឧ. 17-09=ថ្ងៃប្រជុំ →
// ថ្ងៃត្រឹមត្រូវបន្ទាប់ក្លាយជា 18-09 ដោយស្វ័យប្រវត្តិ) ។ រក្សាទុកជា Sheet ដាច់ដោយឡែក (ដូចគ្នានឹងគំរូ
// "ការអនុញ្ញាតកែថ្ងៃចាស់" ខាងលើ) ដើម្បីតាមដាន/បង្ហាញ/ដកការសម្គាល់វិញបាន ====
var SHEET_SKIPPED_DATES_ = "ថ្ងៃគ្មានទិន្នន័យ (ប្រជុំ)";
var CACHE_KEY_SKIPPED_DATES_ = "skippedPeriodDates_v1";

function ensureSkippedDatesSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_SKIPPED_DATES_);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SKIPPED_DATES_);
    var headers = ["ID", "កាលបរិច្ឆេទ", "មូលហេតុ", "សម្គាល់ដោយ", "ពេលសម្គាល់"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f59e0b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSkippedDatesList_() {
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY_SKIPPED_DATES_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var ss = getSS_();
  var sheet = ensureSkippedDatesSheet_(ss);
  var list = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    data.forEach(function(r) {
      if (!r[0]) return;
      list.push({ id: r[0], date: normalizeDateStr_(r[1]), note: r[2] || 'ថ្ងៃប្រជុំ', markedBy: r[3], markedAt: r[4] });
    });
  }
  try { CacheService.getScriptCache().put(CACHE_KEY_SKIPPED_DATES_, JSON.stringify(list), CACHE_TTL_); } catch (err) {}
  return list;
}

function clearSkippedDatesCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY_SKIPPED_DATES_); } catch (err) {}
}

function isDateSkipped_(dateStr) {
  var list = getSkippedDatesList_();
  for (var i = 0; i < list.length; i++) if (list[i].date === dateStr) return list[i];
  return null;
}

function getLatestSkippedDate_() {
  var latest = null;
  getSkippedDatesList_().forEach(function(s) { if (s.date && (!latest || s.date > latest)) latest = s.date; });
  return latest;
}

function markPeriodDateSkippedInternal_(dateStr, note, markedBy) {
  var ss = getSS_();
  var sheet = ensureSkippedDatesSheet_(ss);
  sheet.appendRow([newId_(), dateStr, note || 'ថ្ងៃប្រជុំ', markedBy, formatNow_()]);
  clearSkippedDatesCache_();
}

// ត្រូវហៅពី Client (ប៊ូតុង "កំណត់ថ្ងៃប្រជុំ") — សម្គាល់តែ "ថ្ងៃត្រឹមត្រូវបន្ទាប់" ប៉ុណ្ណោះ (រក្សាលំដាប់ដូចគ្នានឹងការ
// បង្កើត Sheet ជានិច្ច — មិនអនុញ្ញាតសម្គាល់រំលងកណ្តាល ឬថ្ងៃដែលមាន Sheet រួចហើយឡើយ)
function skipPeriodDate(currentUsername, dateStr, note, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (role !== ROLE_SUPERADMIN && role !== ROLE_ADMIN && role !== ROLE_PEC21) {
    return { success: false, message: "មានតែ SuperAdmin, Admin, ឬ PEC21 ប៉ុណ្ណោះ ដែលអាចសម្គាល់ថ្ងៃនេះបាន!" };
  }
  var d = normalizeDateStr_(dateStr);
  if (!d) return { success: false, message: "សូមជ្រើសរើសកាលបរិច្ឆេទ!" };
  if (isDateSkipped_(d)) return { success: false, message: "ថ្ងៃនេះត្រូវបានសម្គាល់ថា \"គ្មានទិន្នន័យ\" រួចហើយ!" };

  var nextInfo_ = getNextExpectedPeriodDate_();
  if (!nextInfo_.isFirst && d !== nextInfo_.nextDate) {
    return {
      success: false,
      message: "អាចសម្គាល់បានតែថ្ងៃ \"" + nextInfo_.nextDate + "\" (ថ្ងៃត្រឹមត្រូវបន្ទាប់តាមលំដាប់) ប៉ុណ្ណោះ! (Sheet ចុងក្រោយបំផុតបច្ចុប្បន្ន៖ \""
        + nextInfo_.currentDate + "\")"
    };
  }
  // ចំណាំ៖ d === nextInfo_.nextDate ធានាថាមិនទាន់មាន Sheet សម្រាប់ថ្ងៃនេះជាដាច់ខាត (nextDate កំណត់ជា Sheet
  // ចុងក្រោយបំផុត+១ថ្ងៃ ជានិច្ច) ដូច្នេះមិនចាំបាច់ចំណាយពេលបើក Spreadsheet ស្រុកទាំង១០ ដើម្បីត្រួតពិនិត្យម្តងទៀតទេ
  markPeriodDateSkippedInternal_(d, String(note || '').trim() || 'ថ្ងៃប្រជុំ', currentUsername);
  return { success: true, message: "សម្គាល់ថ្ងៃ \"" + d + "\" ថា \"គ្មានទិន្នន័យ\" ជោគជ័យ! ការបង្កើត Sheet ថ្មីបន្ទាប់ នឹងរំលងថ្ងៃនេះដោយស្វ័យប្រវត្តិ។" };
}

// ដកការសម្គាល់ — អនុញ្ញាតតែពេលវាជា "ថ្ងៃចុងក្រោយបំផុតក្នុងលំដាប់" ប៉ុណ្ណោះ (គ្មានអ្វីទៀតបន្តបន្ទាប់ពីវា ទោះជា Sheet
// ឬថ្ងៃសម្គាល់ផ្សេងទៀតក៏ដោយ) ដើម្បីការពារកុំឲ្យលំដាប់ខូច
function unskipPeriodDate(currentUsername, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var d = normalizeDateStr_(dateStr);
  if (!d) return { success: false, message: "កាលបរិច្ឆេទមិនត្រឹមត្រូវទេ!" };

  var latestSkip_ = getLatestSkippedDate_();
  var latestSheet_ = getLatestExistingPeriodDate_();
  if (d !== latestSkip_ || (latestSheet_ && latestSheet_ > d)) {
    return { success: false, message: "អាចដកការសម្គាល់បានតែថ្ងៃចុងក្រោយបំផុតក្នុងលំដាប់ប៉ុណ្ណោះ (មានថ្ងៃថ្មីជាងនេះរួចហើយក្នុងប្រព័ន្ធ)!" };
  }
  var list = getSkippedDatesList_();
  var target = null;
  for (var i = 0; i < list.length; i++) { if (list[i].date === d) { target = list[i]; break; } }
  if (!target) return { success: false, message: "រកមិនឃើញកំណត់ត្រានេះទេ!" };

  var ss = getSS_();
  var sheet = ensureSkippedDatesSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][0]) === String(target.id)) {
        sheet.deleteRow(r + 2);
        clearSkippedDatesCache_();
        return { success: true, message: "ដកការសម្គាល់ថ្ងៃ \"" + d + "\" ចេញជោគជ័យ!" };
      }
    }
  }
  return { success: false, message: "រកមិនឃើញកំណត់ត្រានេះទេ!" };
}

function listSkippedPeriodDates(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var list = getSkippedDatesList_().slice().sort(function(a, b) { return b.date.localeCompare(a.date); });
  return { success: true, skippedDates: list };
}

// ត្រឡប់ true ប្រសិនបើគណនីនេះ ត្រូវបានចាក់សោ មិនឲ្យកែ/លុបទិន្នន័យសម្រាប់ ស្រុក+កាលបរិច្ឆេទ នេះទេ
function isPastDateLocked_(currentUsername, district, dateStr) {
  if (isAdmin_(currentUsername)) return false; // SuperAdmin/Admin មិនរងឥទ្ធិពលឡើយ
  var d = normalizeDateStr_(dateStr);
  if (d >= getCurrentPeriodDate_()) return false; // ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន ឬអនាគត — មិនចាក់សោទេ
  // ចំណាំ៖ ថ្ងៃចាប់ផ្តើម (ថ្ងៃទី១) ក៏ត្រូវចាក់សោដូចថ្ងៃផ្សេងទៀតដែរ ម្តងណាមានថ្ងៃថ្មីជាងកើតឡើង (តាមការស្នើសុំចុងក្រោយ)
  return !isDateUnlockedForDistrict_(district, d);
}

// ==== ការការពារពិតប្រាកដលើ Google Sheet ខ្លួនឯង (Protect Sheet) — កុំឲ្យគ្រូអាចរំលងការចាក់សោ ដោយកែផ្ទាល់ក្នុង Sheet ====
// ការចាក់សោក្នុង App ខាងលើ ការពារតែផ្លូវចូលតាមរយៈ App ប៉ុណ្ណោះ។ ដោយសារគ្រូប្រចាំស្រុក មានសិទ្ធិចូល Google Sheet
// ដោយផ្ទាល់ (សម្រាប់បញ្ចូលទិន្នន័យផ្ទាល់ជាជម្រើស) ការចាក់សោត្រូវអនុវត្តផងដែរ នៅកម្រិត Google Sheet ខ្លួនឯង
// ដោយប្រើ Protection API — លុបសិទ្ធិកែប្រែរបស់អ្នកកែប្រែទាំងអស់ចេញ ហើយទុកតែម្ចាស់ Script (Admin) ប៉ុណ្ណោះដែលអាចកែបាន។
function applyOrRemoveSheetProtection_(sheet, shouldProtect) {
  try {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) {
      try { if (p.getDescription() === "ចាក់សោស្វ័យប្រវត្តិ — ថ្ងៃកន្លងផុត") p.remove(); } catch (e) {}
    });
    if (shouldProtect) {
      var protection = sheet.protect().setDescription("ចាក់សោស្វ័យប្រវត្តិ — ថ្ងៃកន្លងផុត");
      try { protection.removeEditors(protection.getEditors()); } catch (e) {}
      try { protection.setDomainEdit(false); } catch (e) {}
      try { protection.addEditor(Session.getEffectiveUser().getEmail()); } catch (e) {}
    }
  } catch (err) {}
}

// ត្រួតពិនិត្យ/អនុវត្តការការពារ សម្រាប់ ស្រុក+កាលបរិច្ឆេទ តែមួយ (ប្រើភ្លាមៗពេល Admin បើក/បិទសិទ្ធិ — លឿន មិនស្កេនអ្វីទាំងអស់)
function protectOnePeriodSheetIfNeeded_(district, dateStr) {
  var d = normalizeDateStr_(dateStr);
  var dSs = getDistrictSpreadsheet_(district);
  var sheet = dSs.getSheetByName(periodSheetName_(d));
  if (!sheet) return;
  var shouldProtect = d < getCurrentPeriodDate_() && !isDateUnlockedForDistrict_(district, d);
  applyOrRemoveSheetProtection_(sheet, shouldProtect);
}

// ស្កេនគ្រប់ Period Sheet ក្នុងស្រុកទាំង១០ ហើយអនុវត្ត/ដកការការពារ តាមក្បួនចាក់សោ — ហៅរាល់ថ្ងៃស្វ័យប្រវត្តិ (Trigger)
// ដើម្បីធានាថា Tab ដែលទើបប្រែក្លាយពី "ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន" ទៅជា "ថ្ងៃចាស់" ត្រូវបានចាក់សោដោយស្វ័យប្រវត្តិ
function protectPastPeriodSheets_() {
  try {
    var current = getCurrentPeriodDate_();
    DISTRICT_LIST.forEach(function(district) {
      // ==== ល្បឿន៖ ប្រើ Map កាលបរិច្ឆេទ Tab ដែលបាន Cache ទុក ជំនួសការអាន .getValue() ម្តងមួយៗ សម្រាប់ Tab ជារាប់សិប ====
      var dateMap = getPeriodSheetDateMap_(district);
      var sheetInfos = listDistrictPeriodSheets_(district).map(function(sheet) {
        return { sheet: sheet, date: dateMap[sheet.getName()] || null };
      }).filter(function(info) { return info.date; });

      sheetInfos.forEach(function(info) {
        var isFirstDay = isFirstOperationDay_(info.date);
        // ថ្ងៃចាប់ផ្តើម ត្រូវចាក់សោដូចថ្ងៃផ្សេងទៀតដែរ ម្តងណាមានថ្ងៃថ្មីជាងកើតឡើង (តាមការស្នើសុំចុងក្រោយ)
        var shouldProtect = info.date < current && !isDateUnlockedForDistrict_(district, info.date);
        applyOrRemoveSheetProtection_(info.sheet, shouldProtect);
        applyPeriodSheetColumnVisibility_(info.sheet, !isFirstDay); // ថ្ងៃចាប់ផ្តើម — នៅតែមិនលាក់ ស្ថិតិប៉ាន់ស្មាន ។ល។ (ខុសគ្នាពីការចាក់សោ)
        // ការពារជួរឈររូបមន្ត — អនុវត្តរាល់ថ្ងៃ សម្រាប់ Tab ទាំងអស់ (មិនថាថ្ងៃចាស់ ឬបច្ចុប្បន្ន/អនាគត) ដើម្បីជួសជុលខ្លួនឯង
        // (Self-healing) ចំពោះ Tab ណាមួយដែលការការពារបាត់ដោយចៃដន្យ (ឧ. ត្រូវបានលុបផ្ទាល់ក្នុង Google Sheet ដោយ Admin)
        protectPeriodSheetFormulaColumns_(info.sheet);
      });
    });
  } catch (err) {
    try { Logger.log("protectPastPeriodSheets_ error: " + err.message); } catch (e2) {}
  }
}

// ==== FIX (ល្បឿន — "សូមការបង្កើត Sheet ថ្មីសម្រាប់គ្រូប្រចាំក្រុងស្រុកលឿន")៖ protectPastPeriodSheets_() ខាងលើ ស្កេន
// Period Sheet "ចាស់" ទាំងអស់ (គ្រប់ថ្ងៃមុនៗ) × គ្រប់ស្រុកទាំង១០ ជានិច្ច (មិនមែនតែ Sheet ថ្មីទើបបង្កើតទេ) — កាន់តែយូរទៅៗ
// កាន់តែយឺត ព្រោះចំនួន Sheet ចាស់កើនឡើងជារៀងរាល់ថ្ងៃ (ឧ. ២០ថ្ងៃកន្លងទៅ × ១០ស្រុក = Sheet ២០០ ត្រូវពិនិត្យ/ការពារឡើងវិញ
// រាល់ដងបង្កើត Sheet ថ្មី)។ ការហៅដោយផ្ទាល់ (Synchronous) ក្នុង createNextDailyPeriodSheets() ធ្វើឲ្យគ្រូ/Admin ត្រូវរង់ចាំ
// ជំហាននេះបញ្ចប់ សិន ទើបឃើញសារ "ជោគជ័យ" ទោះបីជាការបង្កើត Sheet ថ្មីខ្លួនឯង (ជំហានសំខាន់ដែលអ្នកប្រើកំពុងរង់ចាំមើលលទ្ធផល
// ភ្លាមៗ) ចប់រួចហើយក៏ដោយ។ ជំហាននេះជា "ការពារខ្លួនឯង" (Self-heal) សម្រាប់ Sheet ចាស់ដែលធ្លាប់បានចាក់សោរួច (មិនប៉ះពាល់
// ដល់ទិន្នន័យ Sheet ថ្មីដែលទើបបង្កើតឡើយ — Sheet ថ្មីទទួលបានរូបមន្ត/ការការពារផ្ទាល់ខ្លួនរួចហើយពី applyPeriodSheetFormulasAndFormatting_
// ដូចដែលមានស្រាប់) ដូច្នេះមិនចាំបាច់ត្រូវបញ្ចប់មុននឹងឆ្លើយតបទៅអ្នកប្រើទេ — ពន្យារពេលឲ្យដំណើរការនៅផ្ទៃខាងក្រោយវិញ (Pattern
// ដដែលនឹង scheduleSyncSheetPermissions_/scheduleRebuild_ ដែលមានស្រាប់រួចហើយ — CacheService Flag ការពារកុំឲ្យបង្កើត
// Trigger ស្ទួនច្រើនដងក្នុងរយៈពេលខ្លី) ====
// ==== សំខាន់៖ ចាំបាច់ត្រូវប្រើឈ្មោះ Handler ដាច់ដោយឡែក ("runScheduledProtectPastPeriodSheets_") មិនមែនហៅ
// "protectPastPeriodSheets_" ដោយផ្ទាល់ត្រង់ Trigger ទេ — ព្រោះ Function នោះមាន Trigger អចិន្ត្រៃយ៍ (Persistent —
// ដំណើរការរាល់ថ្ងៃ ម៉ោង ១ព្រឹក ដំឡើងដោយ ensureDailyProtectionTrigger_ ខាងក្រោម) ដែលត្រូវការឲ្យនៅគង់វង្សអចិន្ត្រៃយ៍
// (មិនត្រូវលុបចោលដោយអចេតនាឡើយ)។ ប្រើ Wrapper ដាច់ដោយឡែក (ដូចគ្នានឹង scheduleSyncSheetPermissions_/
// runScheduledSyncSheetPermissions_ ដែលមានស្រាប់រួចហើយ) ដើម្បីអាចលុប Trigger បណ្តោះអាសន្នរបស់ខ្លួនឯងដោយសុវត្ថិភាព
// ដោយមិនប៉ះពាល់ Trigger អចិន្ត្រៃយ៍នោះឡើយ ====
function scheduleProtectPastPeriodSheets_() {
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('protectPastTriggerScheduled_')) return; // Trigger ស្រាប់ត្រូវបានគ្រោងទុករួចហើយ (ក្នុងរយៈពេលថ្មីៗនេះ)
    ScriptApp.newTrigger('runScheduledProtectPastPeriodSheets_').timeBased().after(2000).create();
    cache.put('protectPastTriggerScheduled_', '1', 30);
  } catch (err) {
    // ករណី Trigger បង្កើតមិនចេញ (ដែនកំណត់/សិទ្ធិ) — ធ្វើភ្លាមៗវិញ ដើម្បីធានាការការពារ Tab ចាស់មិនធ្លាក់ខាងក្រោយ
    try { protectPastPeriodSheets_(); } catch (errFallback) {}
  }
}

// ហៅដោយ Trigger ស្វ័យប្រវត្តិប៉ុណ្ណោះ (មិនត្រូវហៅផ្ទាល់ពីកន្លែងផ្សេងទេ — ប្រើ scheduleProtectPastPeriodSheets_() ជំនួសវិញ)
function runScheduledProtectPastPeriodSheets_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'runScheduledProtectPastPeriodSheets_') { try { ScriptApp.deleteTrigger(t); } catch (e) {} }
    });
  } catch (err) {}
  try { protectPastPeriodSheets_(); } catch (errSync) {}
}

// ដំឡើង Trigger ប្រចាំថ្ងៃ (ធ្វើតែម្តងគត់) ដើម្បីហៅ protectPastPeriodSheets_ រៀងរាល់ថ្ងៃ — ការពារ Tab ថ្មីៗ ដែលទើបក្លាយជាថ្ងៃចាស់
// ==== ល្បឿន៖ ScriptApp.getProjectTriggers() យឺតគួរឲ្យកត់សម្គាល់ (ត្រូវរាប់ Trigger ទាំងអស់ក្នុងគម្រោង រួមទាំង Trigger
// កែប្រែផ្ទាល់ក្នុង Google Sheet របស់ស្រុកទាំង១០ផងដែរ) ហើយមុននេះមុខងារនេះត្រូវបានហៅ *រាល់ doGet()* មានន័យថា *រាល់
// ដងបើកទំព័រ App* (សូម្បីតែមុននឹង Login ទៀត!) ត្រូវរង់ចាំការហៅនេះជានិច្ច ដែលជាមូលហេតុចម្បងបំផុតមួយ ដែលធ្វើឲ្យ
// ការបើកទំព័រដំបូងមានអារម្មណ៍យឺត។ ដំណោះស្រាយ៖ ចងចាំលទ្ធផលថា "បានផ្ទៀងផ្ទាត់រួចហើយ" ក្នុង Cache រយៈពេល ១០នាទី
// (ដូចគ្នានឹង CACHE_TTL_ ផ្សេងទៀតក្នុងគម្រោងនេះ) ដើម្បីរំលងការហៅ ScriptApp.getProjectTriggers() ថ្លៃថ្នូរនេះ ក្នុងករណី
// ភាគច្រើន (Trigger នេះកម្រនឹងត្រូវលុបចោលដោយចៃដន្យណាស់) — នៅតែផ្ទៀងផ្ទាត់ម្តងទៀតដោយស្វ័យប្រវត្តិ រៀងរាល់ ១០នាទី
// ដើម្បីជួសជុលខ្លួនឯង (Self-healing) បើ Trigger នោះបាត់ទៅវិញ ====
var DAILY_PROTECTION_TRIGGER_CACHE_KEY_ = 'dailyProtTriggerOk_v1';
function ensureDailyProtectionTrigger_() {
  try {
    if (CacheService.getScriptCache().get(DAILY_PROTECTION_TRIGGER_CACHE_KEY_)) return;
  } catch (err) {}
  try {
    var exists = ScriptApp.getProjectTriggers().some(function(t) {
      return t.getHandlerFunction() === "protectPastPeriodSheets_";
    });
    if (!exists) {
      ScriptApp.newTrigger("protectPastPeriodSheets_").timeBased().everyDays(1).atHour(1).create();
    }
    CacheService.getScriptCache().put(DAILY_PROTECTION_TRIGGER_CACHE_KEY_, '1', CACHE_TTL_);
  } catch (err) {}
}

// ==================== ការតាមដានអ្នកបញ្ចូល/កែប្រែ ដោយផ្ទាល់ក្នុង Google Sheet (onEdit) ====================
// លំនាំដើម ជួរឈរ "អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ត្រូវបានបំពេញ ត្រឹមតែពេលរក្សាទុកតាមរយៈ App ប៉ុណ្ណោះ។ ដើម្បីតាមដាន
// ការកែប្រែផ្ទាល់ក្នុង Google Sheet ដែរ (ជម្រើសមួយសម្រាប់គ្រូបញ្ចូលទិន្នន័យ) ត្រូវការ Trigger ដំឡើងដោយឡែក
// (Installable) ចងភ្ជាប់ទៅ Spreadsheet ស្រុកនីមួយៗដោយផ្ទាល់ (មិនមែន Trigger សាមញ្ញធម្មតាទេ ព្រោះត្រូវការសិទ្ធិ
// អាន Sheet(អ្នកប្រើប្រាស់) ក្នុង Spreadsheet មេ ដែលគ្រូខ្លួនឯងគ្មានសិទ្ធិចូលដោយផ្ទាល់ឡើយ)។

// រកឈ្មោះគណនី (Username) ក្នុងប្រព័ន្ធ ដោយផ្គូផ្គងតាម Gmail — ប្រើសម្រាប់បង្ហាញឈ្មោះគណនីជំនួស Email ឆៅ
// ==== ល្បឿន៖ ចងចាំ (Memoize) ផែនទី email→username ក្នុងមួយ Execution (ដូចគ្នានឹង USER_ROW_CACHE_ ក្នុង Auth.gs) —
// មុននេះ ស្កេន Sheet(អ្នកប្រើប្រាស់) ទាំងមូលរាល់ដងដែលមានការកែក្នុង Google Sheet ដោយផ្ទាល់ (onDistrictSheetEdit_) ====
var USERNAME_BY_GMAIL_CACHE_ = null;
function getUsernameByGmail_(email) {
  if (!email) return "";
  try {
    if (!USERNAME_BY_GMAIL_CACHE_) {
      USERNAME_BY_GMAIL_CACHE_ = {};
      var sheet = ensureUsersSheet_(getSS_());
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        var gmail = String(data[i][8] || "").trim().toLowerCase();
        if (gmail) USERNAME_BY_GMAIL_CACHE_[gmail] = data[i][1];
      }
    }
    var found = USERNAME_BY_GMAIL_CACHE_[String(email).trim().toLowerCase()];
    if (found) return found;
  } catch (err) {}
  return email;
}

// ដំឡើង Trigger onEdit (ធ្វើតែម្តងគត់) សម្រាប់ Spreadsheet ស្រុកមួយ
// ==== ល្បឿន៖ ទទួល `existingTriggers` ជាជម្រើស (Array ដែលហៅ ScriptApp.getProjectTriggers() រួចស្រាប់) ដើម្បីជៀសវាង
// ការហៅ API ថ្លៃថ្នូរនេះម្តងទៀត ក្នុងករណីហៅមុខងារនេះច្រើនដងជាប់គ្នា (ឧ. installAllDistrictEditTriggers លើស្រុកទាំង១០)។
// បើមិនបានផ្តល់មក (ឧ. ការហៅតែម្តងគត់ ពេលបង្កើត Spreadsheet ស្រុកថ្មីក្នុង getDistrictSpreadsheet_) នៅតែហៅដូចដើម ====
function ensureDistrictEditTrigger_(dSs, existingTriggers) {
  try {
    var ssId = dSs.getId();
    var triggers = existingTriggers || ScriptApp.getProjectTriggers();
    var exists = triggers.some(function(t) {
      return t.getHandlerFunction() === "onDistrictSheetEdit_" && t.getTriggerSourceId() === ssId;
    });
    if (!exists) {
      ScriptApp.newTrigger("onDistrictSheetEdit_").forSpreadsheet(dSs).onEdit().create();
    }
  } catch (err) {}
}

// ដំឡើង Trigger onEdit ជូនគ្រប់ស្រុកទាំង១០ ក្នុងលើកតែមួយ — ប្រើសម្រាប់ Spreadsheet ស្រុកចាស់ៗ ដែលបានបង្កើតរួច
// មុននឹង Trigger នេះមាន (Spreadsheet ថ្មីៗ ដំឡើងស្វ័យប្រវត្តិរួចហើយតាមរយៈ getDistrictSpreadsheet_)
function installAllDistrictEditTriggers(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var installed = 0;
  // ==== ល្បឿន៖ ហៅ ScriptApp.getProjectTriggers() តែម្តងគត់ ជំនួសការហៅ ១ដងក្នុងមួយស្រុក (១០ដង) — ថ្លៃណាស់ ព្រោះត្រូវរាប់
  // Trigger ទាំងអស់ក្នុងគម្រោង រួមទាំង Trigger កែប្រែផ្ទាល់ក្នុង Google Sheet ស្រុកទាំង១០ផងដែរ ====
  var triggers;
  try { triggers = ScriptApp.getProjectTriggers(); } catch (err) { triggers = null; }
  DISTRICT_LIST.forEach(function(district) {
    try {
      var dSs = getDistrictSpreadsheet_(district);
      ensureDistrictEditTrigger_(dSs, triggers);
      installed++;
    } catch (err) {}
  });
  return { success: true, message: "ដំឡើងការតាមដានសម្រាប់ស្រុកចំនួន " + installed + " ជោគជ័យ!" };
}

// ==== ការពារជួរឈររូបមន្ត (Range Protect) ភ្លាមៗ សម្រាប់ Period Sheet ទាំងអស់ គ្រប់ស្រុកទាំង១០ — ប្រើសម្រាប់ Tab ចាស់ៗ
// ដែលបានបង្កើតរួច មុននឹងលក្ខណៈនេះមាន (Tab ថ្មីៗ ទទួលបានការការពារនេះស្វ័យប្រវត្តិរួចហើយ ជារៀងរាល់ពេលមានការកែប្រែតាមរយៈ App
// ហើយក៏ត្រូវបានជួសជុលដោយស្វ័យប្រវត្តិផងដែរ តាម Trigger ប្រចាំថ្ងៃ protectPastPeriodSheets_) — ចុចប៊ូតុងនេះ ដើម្បីអនុវត្តភ្លាមៗ
// ដោយមិនចាំបាច់រង់ចាំដល់ថ្ងៃបន្ទាប់ ឬការកែប្រែលើកក្រោយតាមរយៈ App ទេ ====
function protectAllPeriodSheetFormulaColumnsNow(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var protectedCount = 0, sheetCount = 0;
  DISTRICT_LIST.forEach(function(district) {
    try {
      listDistrictPeriodSheets_(district).forEach(function(sheet) {
        try {
          protectPeriodSheetFormulaColumns_(sheet);
          sheetCount++;
        } catch (errSheet) {}
      });
      protectedCount++;
    } catch (errDistrict) {}
  });
  return { success: true, message: "បានការពារជួរឈររូបមន្តសម្រាប់ស្រុកចំនួន " + protectedCount + " (Sheet ចំនួន " + sheetCount + ") ជោគជ័យ!" };
}

// ==== FIX (សំណើថ្មី "រៀបតាមលំដាប់ថ្ងៃខែ")៖ រៀបចំលំដាប់ Tab (Period Sheet) ក្នុង Spreadsheet ស្រុកមួយ ឲ្យតាមលំដាប់
// ថ្ងៃខែពីតូចទៅធំ។ ប្រើវិធី "រើសយកតាមលំដាប់ថ្ងៃខែ ផ្លាស់ទីទៅចុងម្តងមួយៗ" — សុវត្ថិភាព និងសាមញ្ញ (មិនចាំបាច់គណនាលេខ
// ទីតាំង Index ស្មុគស្មាញ) ដោយមិនប៉ះពាល់ទិន្នន័យក្នុង Sheet ណាមួយឡើយ (ប្តូរតែលំដាប់ទីតាំង Tab)។ ត្រឡប់ចំនួន Sheet
// ដែលបានផ្លាស់ទី (0 = រៀបរួចហើយ មិនចាំបាច់ធ្វើអ្វី) ====
function reorderPeriodSheetTabsInSpreadsheet_(dSs, dateMap) {
  var sheets = dSs.getSheets();
  var periodEntries = [];
  sheets.forEach(function(s) {
    var d = dateMap[s.getName()];
    if (d) periodEntries.push({ sheet: s, date: d });
  });
  if (periodEntries.length < 2) return { moved: 0, total: periodEntries.length };
  var alreadySorted = true;
  for (var i = 1; i < periodEntries.length; i++) {
    if (periodEntries[i].date < periodEntries[i - 1].date) { alreadySorted = false; break; }
  }
  if (alreadySorted) return { moved: 0, total: periodEntries.length };
  var sortedByDate = periodEntries.slice().sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  var moved = 0;
  sortedByDate.forEach(function(e) {
    dSs.setActiveSheet(e.sheet);
    dSs.moveActiveSheet(dSs.getNumSheets());
    moved++;
  });
  return { moved: moved, total: periodEntries.length };
}

// ==== ជួសជុលភ្លាមៗ (Admin) ដើម្បីរៀបចំលំដាប់ Tab (Period Sheet) ឲ្យតាមលំដាប់ថ្ងៃខែ សម្រាប់ស្រុកទាំងអស់ភ្លាមៗ — ប្រើ
// សម្រាប់ជួសជុល Spreadsheet ចាស់ៗ ដែល Tab បានលេចធាតុមិនតាមលំដាប់រួចហើយ (Sheet ថ្មីៗពីនេះទៅ នឹងរៀបចំដោយស្វ័យប្រវត្តិ
// ជានិច្ចរួចហើយ តាមរយៈ createDistrictPeriodSheet_ ខាងលើ) ====
function reorderAllPeriodSheetTabsNow(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var totalMoved = 0, totalSheets = 0, districtsFixed = 0;
  DISTRICT_LIST.forEach(function(district) {
    try {
      clearPeriodSheetDateCache_(district);
      var dSs = getDistrictSpreadsheet_(district);
      var dateMap = getPeriodSheetDateMap_(district);
      var result = reorderPeriodSheetTabsInSpreadsheet_(dSs, dateMap);
      totalMoved += result.moved;
      totalSheets += result.total;
      if (result.moved > 0) districtsFixed++;
    } catch (errDistrict) {}
  });
  return { success: true, message: "បានពិនិត្យ Sheet ចំនួន " + totalSheets + " (ស្រុកចំនួន " + DISTRICT_LIST.length + ") — បានរៀបចំលំដាប់ថ្មីសម្រាប់ស្រុកចំនួន " + districtsFixed + " (ផ្លាស់ទី Tab ចំនួន " + totalMoved + ") ជោគជ័យ!" };
}

// ==== FIX (សំណើថ្មី "ជូនដំណឹង SuperAdmin ពេលគ្រូបញ្ចូលទិន្នន័យផ្ទាល់ក្នុង Google Sheet")៖ ពីមុន SuperAdmin គ្មានវិធីណា
// ដឹងថាគ្រូប្រចាំក្រុងស្រុកណាកំពុងបញ្ចូលទិន្នន័យដោយផ្ទាល់ក្នុង Google Sheet (មិនកន្លងតាម App) ទាល់តែសោះ (ឧ. ករណីស្រុក
// ត្រាំកក់ ដែលរាយការណ៍មក)។ FIX៖ ប្រើប្រព័ន្ធជូនដំណឹង Telegram ដែលមានស្រាប់ (notifyTelegramImmediate_ — ប្រើដដែលនឹង
// ការជូនដំណឹងគណនីថ្មី/Comment ។ល។ កំណត់រួចហើយក្នុង Tab "ការកំណត់ប្រព័ន្ធ" → "ការជូនដំណឹងតាម Telegram") ជូនដំណឹងភ្លាមៗ
// ពេលមានការកែដំបូង។ ជៀសវាងផ្ញើសារច្រើនដងស្ទួនគ្នា (គ្រូម្នាក់ បញ្ចូលទិន្នន័យច្រើនជួរដេក/ជួរឈរក្នុងសម័យតែមួយ អាចបង្កឲ្យ
// Trigger នេះដំណើរការច្រើនដង) ដោយប្រើ CacheService ជា "គំនូស" (Debounce) — ជូនដំណឹងម្តងគត់ក្នុងមួយ (ស្រុក+ថ្ងៃ) ក្នុង
// រយៈពេល ១០នាទី ====
var CACHE_KEY_SHEET_EDIT_NOTIFY_PREFIX_ = 'sheetEditNotified_';
var SHEET_EDIT_NOTIFY_COOLDOWN_ = 600; // ១០នាទី (វិនាទី)
function notifyDirectSheetEditIfNeeded_(district, dateStr, commune, who) {
  try {
    var cache = CacheService.getScriptCache();
    var key = CACHE_KEY_SHEET_EDIT_NOTIFY_PREFIX_ + district + '|' + dateStr;
    if (cache.get(key)) return; // ធ្លាប់ជូនដំណឹងរួចហើយក្នុងរយៈពេលថ្មីៗនេះ — រំលង (ជៀសវាងស្ទួន)
    cache.put(key, '1', SHEET_EDIT_NOTIFY_COOLDOWN_);
    var text = "📝 គ្រូប្រចាំក្រុងស្រុក \"" + district + "\" កំពុងបញ្ចូលទិន្នន័យដោយផ្ទាល់ក្នុង Google Sheet (មិនកន្លងតាម App)\n"
        + "ថ្ងៃ៖ " + periodSheetName_(dateStr) + (commune ? ("\nឃុំ/សង្កាត់៖ " + commune) : "")
        + (who ? ("\nអ្នកបញ្ចូល៖ " + who) : "");
    notifyTelegramImmediate_(text);
  } catch (err) {}
}

// ដំណើរការរាល់ពេលមានការកែក្នុង Google Sheet ស្រុកណាមួយដោយផ្ទាល់ (មិនកន្លងតាម App) — កត់ត្រា "អ្នកបញ្ចូល"/"ពេលបញ្ចូល"
function onDistrictSheetEdit_(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (!isPeriodSheetName_(sheet.getName())) return;
    var row = range.getRow();
    if (row < 4) return; // ជួរ Header — មិនពាក់ព័ន្ធ

    var n = DAILY_FIELDS.length;
    var editStartCol = periodFieldColIndex_(DAILY_FIELDS[0].key); // ជួរឈរដំបូងនៃ DAILY_FIELDS (1-indexed)
    var noteCol = editStartCol + n;                               // ជួរឈរ "ចំណាំ" (1-indexed)
    var editEndCol = noteCol;                                     // រួមទាំង "ចំណាំ" (មិនរាប់បញ្ចូល អ្នកបញ្ចូល/ពេលបញ្ចូល ខ្លួនឯង)
    var col1 = range.getColumn();
    var col2 = col1 + range.getNumColumns() - 1;
    if (col2 < editStartCol || col1 > editEndCol) return; // ការកែនេះ នៅក្រៅជួរឈរដែលពាក់ព័ន្ធ (ឧ. កែ ID ផ្ទាល់ដៃ) — រំលង

    var email = "";
    try { email = e.user ? e.user.getEmail() : Session.getActiveUser().getEmail(); } catch (err2) {}
    var who = email ? getUsernameByGmail_(email) : "មិនស្គាល់ (Sheet ដោយផ្ទាល់)";

    var byCol = noteCol + 1;
    sheet.getRange(row, byCol, 1, 2).setValues([[who, formatNow_()]]);

    // ==== FIX (សំណើថ្មី "ជូនដំណឹង SuperAdmin")៖ អានទិន្នន័យ Meta (ស្រុក/កាលបរិច្ឆេទ/ឃុំសង្កាត់) ផ្ទាល់ពីជួរដេកដែលទើប
    // កែនេះ (មិនអាស្រ័យលើ Lock/Cache របស់ជំហាន fastReindex ខាងក្រោម ដើម្បីធានាថាការជូនដំណឹងកើតឡើងជានិច្ច ទោះជា
    // ចាក់សោមិនបាន — ការហៅ Sheets API បន្ថែមទីនេះតូចតាច ១ជួរដេក ៥ជួរឈរប៉ុណ្ណោះ) ====
    try {
      var rowMeta_ = sheet.getRange(row, 1, 1, 5).getValues()[0]; // ID, កាលបរិច្ឆេទ, លេខកូដ, ស្រុក, ឃុំ/សង្កាត់
      var districtName_ = String(rowMeta_[3] || '').trim();
      var communeName_ = String(rowMeta_[4] || '').trim();
      var dateForNotify_ = normalizeDateStr_(rowMeta_[1]);
      if (districtName_ && dateForNotify_) notifyDirectSheetEditIfNeeded_(districtName_, dateForNotify_, communeName_, who);
    } catch (errNotify) {}

    // ==== FIX (ល្បឿន — "កែផ្ទាល់ក្នុង Google Sheet")៖ ធ្វើឲ្យលឿនស្មើនឹងការបញ្ចូលតាម App (Form) — App ខ្លួនឯង
    // (saveDistrictDayEntries) ហៅ fastReindexDistrictDay_() ភ្លាមៗបន្ទាប់ពីសរសេរ ដើម្បីឲ្យតារាង "របាយការណ៍ចុះឈ្មោះ
    // ប្រចាំថ្ងៃ" ក្នុង App ឃើញទិន្នន័យថ្មីភ្លាមៗ (មិនចាំបាច់រង់ចាំ scheduleRebuild_() ដែលជា Trigger ពេលវេលា ~4វិនាទី
    // + ត្រូវស្កេនឡើងវិញគ្រប់ស្រុកទាំង១០ — ការងារធំបំផុតក្នុងប្រព័ន្ធ)។ ការកែដោយផ្ទាល់ក្នុង Google Sheet ពីមុន ហៅតែ
    // scheduleRebuild_() ប៉ុណ្ណោះ ធ្វើឲ្យទិន្នន័យលេចមិនទាន់ភ្លាមៗក្នុង App ។ ត្រូវអាន Sheet ទាំងមូល (មិនមែនតែជួរដេក
    // ទើបកែ) ព្រោះ fastReindexDistrictDay_() សរសេរ Index ជា JSON មួយដុំគ្របដណ្ដប់ទិន្នន័យគ្រប់ឃុំ/សង្កាត់ក្នុង
    // ស្រុក+ថ្ងៃនេះ (ដូចគ្នានឹងរបៀបប្រើក្នុង saveDistrictDayEntries)។ ចាក់សោខ្លីមួយភ្លែត ដើម្បីការពារការប៉ះទង្គិចគ្នា
    // ជាមួយស្រុកផ្សេងទៀត ដែលអាចកែ Sheet ក្នុងពេលជិតគ្នា ហើយសរសេរជាន់ Index Sheet រួមគ្នា (fastReindexDistrictDay_
    // ខ្លួនឯង គ្មានការចាក់សោផ្ទាល់ខ្លួនទេ — ពឹងផ្អែកលើអ្នកហៅ ដូចគ្នានឹង saveDistrictDayEntries) ====
    try {
      var lastRow_ = sheet.getLastRow();
      if (lastRow_ >= 4) {
        var lock_ = LockService.getScriptLock();
        var gotLock_ = false;
        try { gotLock_ = lock_.tryLock(10000); } catch (eLock_) {}
        if (gotLock_) {
          try {
            // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ")៖ ត្រូវធានាថា Sheet នេះ (ដែលទើបត្រូវបាន
            // កែដោយផ្ទាល់ក្នុង Google Sheet — មិនចាំបាច់ឆ្លងកាត់ createDistrictPeriodSheet_ ទេ) មានជួរឈរនេះជានិច្ច
            // មុននឹងអាន totalCols_ ទទឹងពេញ (បើមិនដូច្នេះទេ Sheet ចាស់ៗ អាចធ្វើឲ្យ getRange ខាងក្រោមបរាជ័យ — ត្រូវបាន
            // ចាប់ (Catch) ដោយសុវត្ថិភាពរួចហើយនៅខាងក្រៅ ប៉ុន្តែជួសជុលអោយត្រឹមត្រូវនៅទីនេះ ប្រសើរជាង) ====
            try { ensurePeriodSheetHasEditCountColumn_(sheet); } catch (errEnsure4) {}
            var totalCols_ = periodSheetTotalCols_();
            var data_ = sheet.getRange(4, 1, lastRow_ - 3, totalCols_).getValues();
            var districtIdx0_ = DAILY_META_PREFIX.length - 2; // ជួរឈរ "ស្រុក" (កូឡុំទី៤)
            var district_ = String(data_[0][districtIdx0_] || "").trim();
            var dateStr_ = normalizeDateStr_(data_[0][1]); // ជួរឈរ "កាលបរិច្ឆេទ" (កូឡុំទី២)
            if (district_ && dateStr_) fastReindexDistrictDay_(district_, dateStr_, data_);
          } finally {
            try { lock_.releaseLock(); } catch (eRel_) {}
          }
        }
        // បើមិនអាចចាក់សោបាន (កម្រណាស់) — scheduleRebuild_() ខាងក្រោម នៅតែជា Fallback ធានាទិន្នន័យត្រឹមត្រូវទីបំផុត
      }
    } catch (errReindex) {}

    scheduleRebuild_();
  } catch (err) {
    try { Logger.log("onDistrictSheetEdit_ error: " + err.message); } catch (e2) {}
  }
}

// ==================== ការលាក់ជួរឈរខ្លះក្នុង Period Sheet សម្រាប់គ្រូប្រចាំស្រុក (មើលក្នុង Google Sheet ដោយផ្ទាល់) ====================
// ១. ID/កាលបរិច្ឆេទ/ស្រុក (ជួរឈរ ១-២,៤) — លាក់ជានិច្ច (ចាប់ពី Sheet ដំបូងបំផុត) ព្រោះមិនចាំបាច់ដល់គ្រូ
//    (ជួរឈរ ៣ លេខកូដ និងជួរឈរ ៥ ឃុំ/សង្កាត់ មិនប៉ះពាល់ទេ — មិនត្រូវលាក់ជាដាច់ខាត)
// ២. ស្ថិតិប៉ាន់ស្មាន/ការិ.បង្កើតថ្មី/ចំនួនការិ.សរុប (ជួរឈរ ៦-៨) — មិនលាក់ចំពោះ Sheet "ដំបូងបំផុត" របស់ស្រុកនោះ (ដើម្បីឲ្យគ្រូ
//    ឃើញទិន្នន័យយោងម្តងដំបូង) ប៉ុន្តែលាក់ ចាប់ពី Sheet ទី២ តទៅ
function applyPeriodSheetColumnVisibility_(sheet, hideStationStats) {
  try {
    // ==== ជួរឈរ ១=ID, ២=កាលបរិច្ឆេទ, ៣=លេខកូដ, ៤=ស្រុក, ៥=ឃុំ/សង្កាត់ ====
    // លាក់ ID+កាលបរិច្ឆេទ (១-២) និងស្រុក (៤) ដាច់ដោយឡែក ដើម្បីរំលងជួរឈរ "លេខកូដ" (៣) ដែលមិនត្រូវលាក់ជាដាច់ខាត
    sheet.hideColumns(1, 2); // ID, កាលបរិច្ឆេទ
    sheet.showColumns(3, 1); // លេខកូដ — មិនត្រូវលាក់ជាដាច់ខាត (ធានាថាមិនធ្លាប់លាក់ដោយអចេតនា)
    sheet.hideColumns(4, 1); // ស្រុក (ជួរឈរ ៥ ឃុំ/សង្កាត់ មិនប៉ះពាល់ទេ)
    if (hideStationStats) {
      sheet.hideColumns(6, 3); // ស្ថិតិប៉ាន់ស្មាន, ការិ.បង្កើតថ្មី, ចំនួនការិ.សរុប
    } else {
      sheet.showColumns(6, 3); // Sheet "ដំបូងបំផុត" របស់ស្រុក — បង្ហាញវិញ ក្នុងករណីធ្លាប់លាក់ពីមុន
    }
    // ចំណាំ៖ មិនប៉ះពាល់ជួរឈរផ្សេងទៀតឡើយ (ជាពិសេស ជួរឈរ "ស្ទួន" ដែលមានតក្កវិជ្ជាលាក់ដោយឡែកខ្លួនឯងរួចហើយ)
  } catch (err) {}
}


function grantDateUnlock(currentUsername, district, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var d = normalizeDateStr_(dateStr);
  if (!d) return { success: false, message: "សូមជ្រើសរើសកាលបរិច្ឆេទ!" };
  var districtLabel = district || UNLOCK_ALL_DISTRICTS_;
  if (isDateUnlockedForDistrict_(districtLabel, d) && getUnlockList_().some(function(u) { return u.district === districtLabel && u.date === d; })) {
    return { success: false, message: "បានបើកសិទ្ធិសម្រាប់ថ្ងៃនេះរួចហើយ!" };
  }
  var ss = getSS_();
  var sheet = ensureEditUnlocksSheet_(ss);
  sheet.appendRow([newId_(), districtLabel, d, currentUsername, formatNow_()]);
  try { CacheService.getScriptCache().remove(CACHE_KEY_UNLOCKS_); } catch (err) {}
  // អនុវត្តការផ្លាស់ប្តូរភ្លាមៗលើ Google Sheet ខ្លួនឯង (ដកការការពារចេញ) — ត្រូវការស្រុកជាក់លាក់ ឬគ្រប់ស្រុក
  try {
    if (districtLabel === UNLOCK_ALL_DISTRICTS_) {
      DISTRICT_LIST.forEach(function(dist) { protectOnePeriodSheetIfNeeded_(dist, d); });
    } else {
      protectOnePeriodSheetIfNeeded_(districtLabel, d);
    }
  } catch (err) {}
  return { success: true, message: "បើកសិទ្ធិកែថ្ងៃ " + d + " ជោគជ័យ!" };
}

function revokeDateUnlock(currentUsername, unlockId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var ss = getSS_();
  var sheet = ensureEditUnlocksSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(unlockId)) {
        var districtLabel = data[i][1];
        var d = normalizeDateStr_(data[i][2]);
        sheet.deleteRow(i + 2);
        try { CacheService.getScriptCache().remove(CACHE_KEY_UNLOCKS_); } catch (err) {}
        // អនុវត្តការចាក់សោឡើងវិញភ្លាមៗលើ Google Sheet (បើលែងបានអនុញ្ញាតទៀតហើយ)
        try {
          if (districtLabel === UNLOCK_ALL_DISTRICTS_) {
            DISTRICT_LIST.forEach(function(dist) { protectOnePeriodSheetIfNeeded_(dist, d); });
          } else {
            protectOnePeriodSheetIfNeeded_(districtLabel, d);
          }
        } catch (err) {}
        return { success: true, message: "បិទសិទ្ធិជោគជ័យ!" };
      }
    }
  }
  return { success: false, message: "រកមិនឃើញកំណត់ត្រានេះទេ!" };
}

function listDateUnlocks(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var list = getUnlockList_().slice().sort(function(a, b) { return b.date.localeCompare(a.date); });
  return { success: true, unlocks: list, districts: DISTRICT_LIST, allLabel: UNLOCK_ALL_DISTRICTS_ };
}

// ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព")៖ ចាក់សោដោយដៃ (Manual) សម្រាប់ ស្រុក/ក្រុង + កាលបរិច្ឆេទ
// ជាក់លាក់ណាមួយ — ខុសពី isPastDateLocked_() ខាងលើ (ដែលចាក់សោស្វ័យប្រវត្តិតែ "ថ្ងៃចាស់" ប៉ុណ្ណោះ) ព្រោះមុខងារនេះត្រូវ
// អាចចាក់សោសូម្បីតែ "ថ្ងៃបច្ចុប្បន្ន" ក៏ដោយ (ទោះមិនទាន់ចាស់ ឬសូម្បីតែមិនទាន់មាន Period Sheet សម្រាប់ថ្ងៃនោះក៏ដោយ) — ដើម្បី
// ការពារកុំឲ្យគណនីគ្រូប្រចាំក្រុងស្រុក/ឃុំសង្កាត់ បន្តកែប្រែទិន្នន័យដោយអចេតនា ខណៈ Admin កំពុងបោះពុម្ពរបាយការណ៍ប្រចាំថ្ងៃ។
// ស្ថាបនាតាមគំរូដូចគ្នាបេះបិទនឹង SHEET_EDIT_UNLOCKS/grantDateUnlock/revokeDateUnlock/listDateUnlocks ខាងលើ — រួមទាំង
// Sentinel "ទាំងអស់ (គ្រប់ស្រុក)" ដូចគ្នា (ប្រើ UNLOCK_ALL_DISTRICTS_ ដដែល — មិនបង្កើត Sentinel ថ្មីទេ ព្រោះជាទស្សនៈ
// "គ្រប់ស្រុកទាំងអស់" ដូចគ្នា) ដើម្បីអនុញ្ញាតទាំងចាក់សោស្រុកជាក់លាក់តែមួយ ឬចាក់សោគ្រប់ស្រុកទាំង១០ ក្នុងពេលតែម្តង (តាម
// ការស្នើសុំរបស់ Admin — "ប្រើបានទាំងពីរជម្រើស")។ ==== សំខាន់៖ SuperAdmin/Admin/PEC21 (isAdmin_) នៅតែកែប្រែទិន្នន័យ
// បានធម្មតា សូម្បីតែពេលចាក់សោសម្រាប់ Print ក៏ដោយ (ដូចគ្នានឹង isPastDateLocked_ បេះបិទ — ដើម្បីឲ្យ Admin ខ្លួនឯង
// នៅតែកែកំហុសបន្ទាន់មុនបោះពុម្ពបានក្នុងករណីចាំបាច់) — មិនដូច isSkippedDate ដែលទប់ស្កាត់សូម្បីតែ Admin ទេ ====
var SHEET_PRINT_LOCKS = "ការចាក់សោការពារពេលបោះពុម្ព";
var CACHE_KEY_PRINT_LOCKS_ = "printLocks_v1";

function ensurePrintLocksSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_PRINT_LOCKS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PRINT_LOCKS);
    var headers = ["ID", "ស្រុក/ក្រុង", "កាលបរិច្ឆេទ", "ចាក់សោដោយ", "ពេលចាក់សោ"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#dc2626").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getPrintLockList_() {
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY_PRINT_LOCKS_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var ss = getSS_();
  var sheet = ensurePrintLocksSheet_(ss);
  var list = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    data.forEach(function(r) {
      if (!r[0]) return;
      list.push({ id: r[0], district: r[1], date: normalizeDateStr_(r[2]), lockedBy: r[3], lockedAt: r[4] });
    });
  }
  try { CacheService.getScriptCache().put(CACHE_KEY_PRINT_LOCKS_, JSON.stringify(list), CACHE_TTL_); } catch (err) {}
  return list;
}

function isDateManuallyLockedForDistrict_(district, dateStr) {
  var list = getPrintLockList_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].date !== dateStr) continue;
    if (list[i].district === UNLOCK_ALL_DISTRICTS_ || list[i].district === district) return true;
  }
  return false;
}

// ត្រឡប់ true ប្រសិនបើគណនីនេះ ត្រូវបានចាក់សោដោយដៃ (មិនឲ្យកែ/លុបទិន្នន័យ) សម្រាប់ ស្រុក+កាលបរិច្ឆេទ នេះ — ការពារពេល Print
function isPrintLockedForDistrict_(currentUsername, district, dateStr) {
  if (isAdmin_(currentUsername)) return false; // SuperAdmin/Admin/PEC21 មិនរងឥទ្ធិពលឡើយ
  var d = normalizeDateStr_(dateStr);
  return isDateManuallyLockedForDistrict_(district, d);
}

function setPrintLock(currentUsername, district, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var d = normalizeDateStr_(dateStr);
  if (!d) return { success: false, message: "សូមជ្រើសរើសកាលបរិច្ឆេទ!" };
  var districtLabel = district || UNLOCK_ALL_DISTRICTS_;
  if (isDateManuallyLockedForDistrict_(districtLabel, d) && getPrintLockList_().some(function(x) { return x.district === districtLabel && x.date === d; })) {
    return { success: true, message: "បានចាក់សោសម្រាប់ថ្ងៃនេះរួចហើយ!", alreadyLocked: true };
  }
  var ss = getSS_();
  var sheet = ensurePrintLocksSheet_(ss);
  sheet.appendRow([newId_(), districtLabel, d, currentUsername, formatNow_()]);
  try { CacheService.getScriptCache().remove(CACHE_KEY_PRINT_LOCKS_); } catch (err) {}
  return { success: true, message: "ចាក់សោការពារទិន្នន័យសម្រាប់ថ្ងៃ " + d + " (" + districtLabel + ") ជោគជ័យ! គណនីគ្រូប្រចាំក្រុងស្រុក/ឃុំសង្កាត់ នឹងមិនអាចកែប្រែទិន្នន័យបានទៀតទេ រហូតដល់ដោះសោវិញ។" };
}

function clearPrintLock(currentUsername, district, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var d = normalizeDateStr_(dateStr);
  if (!d) return { success: false, message: "សូមជ្រើសរើសកាលបរិច្ឆេទ!" };
  var districtLabel = district || UNLOCK_ALL_DISTRICTS_;
  var ss = getSS_();
  var sheet = ensurePrintLocksSheet_(ss);
  var lastRow = sheet.getLastRow();
  var removedAny = false;
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var rDistrict = data[i][1];
      var rDate = normalizeDateStr_(data[i][2]);
      if (rDate === d && rDistrict === districtLabel) {
        sheet.deleteRow(i + 2);
        removedAny = true;
      }
    }
  }
  if (removedAny) { try { CacheService.getScriptCache().remove(CACHE_KEY_PRINT_LOCKS_); } catch (err) {} }
  return { success: true, message: removedAny ? ("ដោះសោថ្ងៃ " + d + " (" + districtLabel + ") ជោគជ័យ!") : "គ្មានការចាក់សោសម្រាប់ជម្រើសនេះទេ (ប្រហែលជាបានចាក់សោក្នុងវិសាលភាពផ្សេង)!" };
}

// ត្រឡប់ស្ថានភាពបច្ចុប្បន្ន សម្រាប់ ស្រុក+កាលបរិច្ឆេទ ជាក់លាក់មួយ — ប្រើដោយ Client ដើម្បីគូរប៊ូតុងចាក់សោ/ដោះសោ ឲ្យត្រឹមត្រូវ
// (ត្រូវញែក district-specific ដាច់ដោយឡែកពី "ទាំងអស់" ដើម្បីឲ្យ Client ដឹងច្បាស់ថាការចាក់សោបច្ចុប្បន្ន មកពីវិសាលភាពណា)
function getPrintLockStatus(currentUsername, district, dateStr, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!isAdmin_(currentUsername)) return { success: false, message: "គ្មានសិទ្ធិទេ!" };
  var d = normalizeDateStr_(dateStr);
  var list = getPrintLockList_();
  var districtLocked = list.some(function(x) { return x.date === d && x.district === district; });
  var allLocked = list.some(function(x) { return x.date === d && x.district === UNLOCK_ALL_DISTRICTS_; });
  return { success: true, districtLocked: districtLocked, allLocked: allLocked, locked: districtLocked || allLocked };
}


function buildPeriodSheetHeader_(sheet, totalCols, dateStr) {
  var isFirstDay = dateStr ? isFirstOperationDay_(dateStr) : true;
  var refLabelPrefix = isFirstDay ? "ទិន្នន័យ២០២៥" : REF_LABEL_DEFAULT_;
  var row1 = [], row2 = [], row3 = [];
  DAILY_META_PREFIX.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });
  PERIOD_EXTRA_FIELDS.forEach(function(f) {
    row1.push("ទិន្នន័យយោង"); row2.push("—");
    if (f.key === "baseline2025_total") row3.push(refLabelPrefix + " សរុប");
    else if (f.key === "baseline2025_female") row3.push(refLabelPrefix + " ស្រី");
    else row3.push(f.label);
  });
  DAILY_FIELDS.forEach(function(f) { row1.push(f.group || "—"); row2.push(f.sub || "—"); row3.push(f.unit || f.label); });
  DAILY_META_SUFFIX.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });

  sheet.getRange(1, 1, 1, totalCols).setValues([row1]);
  sheet.getRange(2, 1, 1, totalCols).setValues([row2]);
  sheet.getRange(3, 1, 1, totalCols).setValues([row3]);
  sheet.getRange(1, 1, 3, totalCols)
      .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(5);
  mergeConsecutiveHeaderCells_(sheet, 1, totalCols, row1);
  mergeConsecutiveHeaderCells_(sheet, 2, totalCols, row2);
  sheet.setColumnWidths(1, totalCols, 95);
}

// ត្រឡប់ផែនទី { ឈ្មោះឃុំសង្កាត់ → តម្លៃ } សម្រាប់ស្រុកមួយ ពី Sheet(ស្ថិតិប៉ាន់ស្មាន/ការិ.បង្កើតថ្មី/ចំនួនការិ.សរុប)
function fixPeriodSheetCodeColumn_(sheet, totalCols, lastRowExisting) {
  try {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 3) return; // Sheet ទទេ ឬតូចពេក — រំលង
    var header3 = sheet.getRange(3, 1, 1, lastCol).getValues()[0];
    if (header3[2] === "លេខកូដ") return; // ត្រឹមត្រូវរួចហើយ

    var codeAtLastCol = header3[lastCol - 1] === "លេខកូដ"; // រចនាសម្ព័ន្ធពីមុន (បន្ថែមនៅចុង)
    var codeValues = [];
    if (codeAtLastCol && lastRowExisting >= 4) {
      codeValues = sheet.getRange(4, lastCol, lastRowExisting - 3, 1).getValues();
    }

    sheet.insertColumnBefore(3);
    sheet.getRange(1, 3, 2, 1).setValue("—").setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    sheet.getRange(3, 3).setValue("លេខកូដ").setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");

    if (codeAtLastCol) {
      if (codeValues.length) sheet.getRange(4, 3, codeValues.length, 1).setValues(codeValues);
      sheet.deleteColumn(lastCol + 1); // ជួរឈរចាស់ (ឥឡូវរុញទៅស្តាំ ១ បន្ថែម ដោយសារ insertColumnBefore ខាងលើ)
    } else if (lastRowExisting >= 4) {
      var codeMapFix = getLinkedDataCodeMap_();
      var oldKeys = sheet.getRange(4, 4, lastRowExisting - 3, 2).getValues(); // ស្រុក, ឃុំ/សង្កាត់ (ឥឡូវនៅ ៤,៥ បន្ទាប់ពី Insert)
      var codeColFix = oldKeys.map(function(r) {
        return [codeMapFix[String(r[0]).trim() + "|" + String(r[1]).trim()] || ""];
      });
      sheet.getRange(4, 3, codeColFix.length, 1).setValues(codeColFix);
    }
  } catch (err) {}
}

// ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក")៖ ជួរដេក "សរុប" ស្វ័យប្រវត្តិ ដែលបន្ថែមនៅចុងក្រោយបំផុតរបស់ Period Sheet
// នីមួយៗ (មើលការពន្យល់ពេញលេញនៅ writePeriodSheetTotalRow_ ខាងក្រោម)។ សម្គាល់ដោយឈ្មោះពិសេសមួយនៅជួរឈរ "ឃុំ/សង្កាត់"
// (មិនដែលប៉ះទង្គិចនឹងឈ្មោះឃុំ/សង្កាត់ពិតប្រាកដណាមួយឡើយ) ====
var PERIOD_TOTAL_ROW_LABEL_ = "សរុបទាំងអស់";

// ត្រឡប់ថា ជួរដេកមួយ (rowIdx) ក្នុង Period Sheet ជាជួរដេក "សរុប" ស្វ័យប្រវត្តិដែរឬអត់ — ពិនិត្យតាមឈ្មោះពិសេសក្នុង
// ជួរឈរ "ឃុំ/សង្កាត់" (ដូចគ្នានឹងជួរឈរដែល findCommuneRowIndex_ ប្រើ)
function isPeriodSheetTotalRow_(sheet, rowIdx) {
  try {
    var communeCol = DAILY_META_PREFIX.length;
    return String(sheet.getRange(rowIdx, communeCol).getValue() || '').trim() === PERIOD_TOTAL_ROW_LABEL_;
  } catch (err) { return false; }
}

// ត្រឡប់ជួរដេកចុងក្រោយ "ពិតប្រាកដ" នៃទិន្នន័យឃុំ/សង្កាត់ (មិនរាប់បញ្ចូលជួរដេក "សរុប" ស្វ័យប្រវត្តិ បើមាន) — ប្រើជំនួស
// sheet.getLastRow() ត្រង់គ្រប់ចំណុចណាដែលត្រូវការគណនា/សរសេរជាន់ទិន្នន័យ "ទិន្នន័យយោង" ត្រឹមតែលើជួរដេកឃុំ/សង្កាត់ពិត
// ប៉ុណ្ណោះ (បើមិនដូច្នេះទេ ជួរដេក "សរុប" នឹងត្រូវបានគេយល់ច្រឡំថាជាឃុំ/សង្កាត់មួយ ហើយសរសេរតម្លៃខុសទៅក្នុងវា)
function getPeriodSheetCommuneLastRow_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= 4 && isPeriodSheetTotalRow_(sheet, lastRow)) return lastRow - 1;
  return lastRow;
}

// ធានាថា Header + តម្លៃ ក្នុងជួរឈរ "ទិន្នន័យយោង" (ទិន្នន័យ២០២៥ ឬ យោងចំនួនក្នុងបញ្ជីសរុបពីមុន) ត្រឹមត្រូវជានិច្ច សម្រាប់ Sheet ដែលមានរួចហើយ
// (ករណី Sheet ត្រូវបានបង្កើតមុននឹងលក្ខណៈនេះមាន ឬថ្ងៃចាប់ផ្តើមផ្លាស់ប្តូរ) — ជួរឈរនេះជាទិន្នន័យយោង/មិនមែនទិន្នន័យបញ្ចូលដោយដៃ
// ដូច្នេះការសរសេរជាន់ឡើងវិញរាល់ពេលមិនប៉ះពាល់ដល់ការងារពិតរបស់គ្រូឡើយ
// ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក")៖ ត្រូវប្រើ getPeriodSheetCommuneLastRow_() ជំនួសការទុកចិត្តលើ lastRow ដែល
// បញ្ជូនចូលដោយផ្ទាល់ — បើ Sheet នេះមានជួរដេក "សរុប" រួចហើយ lastRow ដើម (raw getLastRow()) នឹងរួមបញ្ចូលជួរដេកនោះ
// ខុសដោយចៃដន្យ ធ្វើឲ្យ Function នេះសរសេរតម្លៃ "ទិន្នន័យយោង" ២០២៥/យោងចាស់ ០ ខុសទៅក្នុងជួរដេក "សរុប" (ព្រោះរកមិនឃើញ
// ឈ្មោះ "សរុបទាំងអស់" ក្នុង refMap ដែលមានតែឈ្មោះឃុំ/សង្កាត់ពិតប្រាកដ) ====
function fixPeriodSheetReferenceColumns_(sheet, district, dateStr, lastRow) {
  try {
    var isFirstDay = isFirstOperationDay_(dateStr);
    var refLabelPrefix = isFirstDay ? "ទិន្នន័យ២០២៥" : REF_LABEL_DEFAULT_;
    var baseTCol = periodExtraColIndex_('baseline2025_total');
    var baseFCol = periodExtraColIndex_('baseline2025_female');

    sheet.getRange(3, baseTCol).setValue(refLabelPrefix + " សរុប");
    sheet.getRange(3, baseFCol).setValue(refLabelPrefix + " ស្រី");

    var communeLastRow = getPeriodSheetCommuneLastRow_(sheet);
    if (communeLastRow >= 4) {
      var communeIdx0 = DAILY_META_PREFIX.length - 1;
      var refMap = getRollingReferenceMap_(dateStr, district).map;
      var totalCols = periodSheetTotalCols_();
      var data = sheet.getRange(4, 1, communeLastRow - 3, totalCols).getValues();
      var newBaseValues = data.map(function(row) {
        var commune = row[communeIdx0];
        var ref = refMap[district + "|" + commune] || { total: 0, female: 0 };
        return [Number(ref.total) || 0, Number(ref.female) || 0];
      });
      sheet.getRange(4, baseTCol, newBaseValues.length, 2).setValues(newBaseValues);
    }
  } catch (err) {}
}

// ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក ដើម្បីងាយស្រួលពិនិត្យផ្ទៀងផ្ទាត់")៖ សរសេរ/ធ្វើបច្ចុប្បន្នភាពជួរដេក "សរុប" ស្វ័យប្រវត្តិ
// (រូបមន្ត SUM គ្របដណ្តប់ជួរដេកឃុំ/សង្កាត់ទាំងអស់) ភ្ជាប់ជាប់ជួរដេកចុងក្រោយបំផុតរបស់ Period Sheet មួយ — ដើម្បីឲ្យ Admin/
// PEC21/គ្រូប្រចាំក្រុងស្រុក អាចមើលឃើញផលបូកសរុបទាំងស្រុងភ្លាមៗ ដោយមិនចាំបាច់បូកដោយដៃ ឬចេញ App ទៅមើលរបាយការណ៍ដាច់ដោយឡែក។
// ==== សុវត្ថិភាពទិន្នន័យ (សំខាន់ណាស់)៖ ជួរដេកនេះមិនដែលមាន ID ទេ (ទុកទទេដោយចេតនា) — កូដស្ទើរគ្រប់ទីកន្លែងទាំងអស់ក្នុង
// ប្រព័ន្ធ ដែលអាន Period Sheet ជួរដេកឆៅសម្រាប់សរុប/របាយការណ៍ (getDailyEntries, rebuildAllDerivedData_ ។ល។) សុទ្ធតែ
// ត្រួតពិនិត្យ "if (!row[0]) return;" (រំលងជួរគ្មាន ID) ជាមុនស្រាប់ស្រាប់ជានិច្ច ដូច្នេះជួរដេកនេះនឹងមិនប៉ះពាល់/មិនត្រូវ
// បូកបញ្ចូលទ្វេដងចូលក្នុងរបាយការណ៍/ផ្ទាំងគ្រប់គ្រងកន្លែងណាមួយឡើយ (ផ្ទៀងផ្ទាត់ដោយប្រុងប្រយ័ត្នមុននឹងដាក់ឲ្យប្រើ) ។ ជួរឈរ
// "ទិន្នន័យយោង" (PERIOD_EXTRA_FIELDS) ក៏ត្រូវបានដាក់រូបមន្ត SUM ដូចគ្នាដែរ ប៉ុន្តែ fixPeriodSheetReferenceColumns_
// (ខាងលើ) នឹងសរសេរជាន់វិញជា ០ រាល់ពេលដំណើរការ self-heal (ព្រោះឈ្មោះ "សរុបទាំងអស់" រកមិនឃើញក្នុងផែនទីទិន្នន័យយោង) —
// ដូច្នេះ Function បន្ទាប់ (createDistrictPeriodSheet_) ត្រូវហៅ Function នេះជានិច្ចក្រោយ fixPeriodSheetReferenceColumns_
// ដើម្បីស្តារតម្លៃ SUM ត្រឹមត្រូវឡើងវិញភ្លាមៗ ==== ត្រូវហៅ Function នេះឡើងវិញ រាល់ពេលចំនួនជួរដេកឃុំ/សង្កាត់ផ្លាស់ប្តូរ
// (បង្កើត Sheet ថ្មី/self-heal Sheet ចាស់/បន្ថែមឃុំ/សង្កាត់ពិសេស) — មិនចាំបាច់ហៅឡើងវិញរាល់ពេលរក្សាទុកទិន្នន័យធម្មតាទេ
// ព្រោះរូបមន្ត SUM ស្វ័យគណនាឡើងវិញដោយខ្លួនឯងជានិច្ច (ដូចគ្នានឹងរូបមន្ត update2026 ក្នុងជួរដេកធម្មតានីមួយៗ) ====
// ==== FIX (ល្បឿន — "បង្កើត Sheet ថ្មី" នៅតែយឺត)៖ ពីមុន Function នេះហៅ setValues()/setFormulas() ដាច់ដោយឡែក ៣ដង
// (Prefix + រូបមន្តជាលេខ + Suffix — ៣ការហៅ Sheets API) រួម នឹងការហៅរចនាភាព/កម្ពស់ជួរដេកទៀត ២ដង — សរុប ~៥ការហៅ
// Sheets API ក្នុងមួយស្រុក × ១០ស្រុក = ~៥០ការហៅបន្ថែម រាល់ពេលបង្កើត Sheet ថ្ងៃថ្មី (លើសពីការហៅជាច្រើនទៀតដែលមានស្រាប់
// ពីមុន ដូចជាការការពារជួរឈររូបមន្ត) ។ FIX៖ ផ្គុំទិន្នន័យទាំងមូល (Prefix+រូបមន្ត+Suffix) ទៅជា Array តែមួយ គ្របដណ្តប់
// ជួរទាំងមូល (columns 1..totalCols) ហើយហៅ setValues() តែម្តងគត់ (Google Sheets កំណត់តម្លៃដែលចាប់ផ្តើមដោយ "="
// ជារូបមន្តដោយស្វ័យប្រវត្តិ ដូចគ្នានឹងវាយបញ្ចូលផ្ទាល់ដែរ — មិនចាំបាច់ហៅ setFormulas() ដាច់ដោយឡែកទេ) — កាត់បន្ថយពី
// ៣ការហៅ Sheets API មកនៅសល់ត្រឹមតែ ១ (សម្រាប់ផ្នែកទិន្នន័យ) ====
function writePeriodSheetTotalRow_(sheet, district, totalCols, communeLastRow) {
  try {
    if (communeLastRow < 4) return; // គ្មានជួរដេកឃុំ/សង្កាត់ណាមួយទេ — មិនចាំបាច់ជួរដេក "សរុប"
    var totalRowIdx = communeLastRow + 1;

    var numericKeys = PERIOD_EXTRA_FIELDS.map(function(f) { return f.key; })
        .concat(DAILY_FIELDS.map(function(f) { return f.key; }));
    var firstNumCol = periodExtraColIndex_(PERIOD_EXTRA_FIELDS[0].key);
    var formulasRow = numericKeys.map(function(key) {
      var isExtra = PERIOD_EXTRA_FIELDS.some(function(f) { return f.key === key; });
      var col = isExtra ? periodExtraColIndex_(key) : periodFieldColIndex_(key);
      var colL = colLetter_(col);
      return '=SUM(' + colL + '4:' + colL + communeLastRow + ')';
    });

    // ==== ជួរដេកទាំងមូល៖ Meta Prefix (ID ទទេដោយចេតនា — សុវត្ថិភាព, មើលខាងលើ) + រូបមន្ត SUM គ្រប់ជួរឈរជាលេខ +
    // Meta Suffix ទទេ — ត្រូវត្រូវនឹង totalCols បេះបិទ ====
    var fullRow = ["", "", "", district, PERIOD_TOTAL_ROW_LABEL_].concat(formulasRow, ["", "", ""]);
    sheet.getRange(totalRowIdx, 1, 1, totalCols).setValues([fullRow]);

    // ==== រចនាភាព៖ ពណ៌ផ្ទៃខាងក្រោយ + អក្សរដិត ដើម្បីសម្គាល់ច្បាស់ថាជាជួរដេក "សរុប" (ហៅជាប់គ្នា — ១ការហៅប្រសិទ្ធភាព) ====
    sheet.getRange(totalRowIdx, 1, 1, totalCols).setFontWeight("bold").setBackground("#fde68a");
  } catch (err) {}
}

// ==== FIX (សំណើ "ពិនិត្យរូបមន្តបច្ចុប្បន្នភាពបញ្ជីឆ្នាំ២០២៦")៖ ជួរឈរ "ទិន្នន័យយោង" (baseline) របស់ថ្ងៃមួយៗ គឺ
// "រូបថតឆាប់រហ័ស" (Snapshot — សរសេរតម្លៃចូល setValues() ធម្មតា) ដែលថតត្រឹមតែ "ពេលបង្កើត/ជួសជុល Sheet" ប៉ុណ្ណោះ
// (មិនមែនរូបមន្តភ្ជាប់រស់ (Live Cross-sheet Formula) ដែលនឹងធ្វើបច្ចុប្បន្នភាពដោយស្វ័យប្រវត្តិឡើយ)។ ដូច្នេះ បើ Admin/
// PEC21 បង្កើត Sheet ថ្ងៃបន្ទាប់ (ឬច្រើនថ្ងៃជាមុន) មុននឹងគ្រូបញ្ចូល/រក្សាទុកទិន្នន័យបន្ថែម/លុបពិតប្រាកដសម្រាប់ថ្ងៃមុននោះ
// ស្រេច — ជួរឈរយោងរបស់ថ្ងៃបន្ទាប់ (និងថ្ងៃៗបន្ត) នឹង "កកខ្ចាប់" នៅត្រឹមតម្លៃចាស់ (ដែលស្មើនឹងទិន្នន័យ២០២៥ បើជាថ្ងៃដើមខ្សែសង្វាក់)
// ជាអចិន្ត្រៃយ៍ ទោះបីជាក្រោយមកមានការបញ្ចូលទិន្នន័យពិតប្រាកដ (ចូល/ចេញ) សម្រាប់ថ្ងៃមុននោះហើយក៏ដោយ — នេះជាមូលហេតុពិត
// នៃរបាយការណ៍ "យោងចំនួនក្នុងបញ្ជីសរុបពីមុន ស្មើនឹង ទិន្នន័យ២០២៥" ជានិច្ច (មិនមែនកំហុសក្នុងរូបមន្ត ១-៣ ខ្លួនឯងទេ — ការគណនា
// ១-៣ ត្រឹមត្រូវ ១០០% បើទិន្នន័យយោងត្រូវបានយកនៅពេលត្រឹមត្រូវ)។
// FIX៖ រាល់ពេលទិន្នន័យប្រចាំថ្ងៃពិតប្រាកដ (ចូល/ចេញ) របស់ថ្ងៃមួយត្រូវបានប្តូរ (រក្សាទុក/កែប្រែ/លុប) ហៅ Function នេះ
// ដើម្បី "ជួសជុលឡើងវិញ" (self-heal) ជួរឈរយោងរបស់រាល់ Sheet ថ្ងៃបន្តបន្ទាប់ទាំងអស់ ដែលមានស្រាប់រួចហើយក្នុងស្រុកនោះ
// (តាមលំដាប់ថ្ងៃខែ ពីជិតទៅឆ្ងាយ ដូច្នេះថ្ងៃនីមួយៗទទួលបានតម្លៃត្រឹមត្រូវពីថ្ងៃមុនផ្ទាល់របស់វា ដែលទើបជួសជុលរួច) — ធានា
// ថាខ្សែសង្វាក់ទាំងមូលត្រឹមត្រូវជានិច្ច ទោះបីជា Sheet អនាគតត្រូវបានបង្កើតជាមុន (មុនពេលទិន្នន័យត្រូវបានបញ្ចូល) ក៏ដោយ។
// ==== ល្បឿន៖ ករណីធម្មតាស្ទើរ១០០% (គ្មាន Sheet អនាគតទាល់តែសោះ — ព្រោះជាទូទៅ Admin បង្កើត Sheet ថ្ងៃមួយម្តង តាមលំដាប់)
// Function នេះចាកចេញភ្លាមៗបន្ទាប់ពីអាន Sheet List + Date Map (Cache ភាគច្រើន — មិនចាំបាច់អាន Sheet ចាស់ម្តងមួយៗទេ)
// ដោយមិនធ្វើអ្វីបន្ថែម — ថ្លៃថ្កាសបន្ថែមកើតឡើងតែក្នុងករណីកម្រ (មាន Sheet អនាគតបង្កើតជាមុនស្រាប់) ប៉ុណ្ណោះ ====
function cascadeRefreshRollingReferenceForward_(district, fromDateStr) {
  try {
    var d = normalizeDateStr_(fromDateStr);
    if (!d) return;
    var dateMap = getPeriodSheetDateMap_(district);
    var sheets = listDistrictPeriodSheets_(district);
    var future = [];
    sheets.forEach(function(sh) {
      var dt = dateMap[sh.getName()];
      if (dt && dt > d) future.push({ date: dt, sheet: sh });
    });
    if (!future.length) return; // ករណីធម្មតា — គ្មាន Sheet អនាគតបង្កើតជាមុនទេ — ចាកចេញភ្លាមៗ
    future.sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
    future.forEach(function(item) {
      var lastRowN = item.sheet.getLastRow();
      fixPeriodSheetReferenceColumns_(item.sheet, district, item.date, lastRowN);
    });
  } catch (err) {}
}

// ==== FIX (រាយការណ៍ថ្មី "ការរៀបលំដាប់ Tab នៅតែខុសដដែល")៖ ត្រួតពិនិត្យលឿន (Metadata ប៉ុណ្ណោះ — sheet.getName(),
// មិនចាំបាច់ Sheets API បន្ថែមទេ) ថាតើលំដាប់ Tab (ទម្រង់ឈ្មោះ "dd-MM") បច្ចុប្បន្ន "មើលទៅ" ដូចជាត្រូវបានប្តូរលំដាប់ដោយ
// ដៃ (អូស Tab ក្នុង Google Sheet ដោយផ្ទាល់ — ក្រៅ App) ដែរឬអត់ — មិនអាចទុកចិត្ត ១០០% បានទេ (ឈ្មោះគ្មានឆ្នាំ) ប៉ុន្តែ
// គ្រប់គ្រាន់សម្រាប់ជា "សញ្ញាព្រមាន" ថោកមួយ ព្រោះក្នុងលំដាប់ត្រឹមត្រូវ (បង្កើតតាមលំដាប់ថ្ងៃខែជានិច្ច) គួរតែមានយ៉ាងច្រើន
// តែ ១ ចំណុចដែលខែ/ថ្ងៃ "ថយក្រោយ" ធៀបនឹង Tab មុន (ឆ្លងកាត់ចុងឆ្នាំ ធ្នូ→មករា ស្របច្បាប់) — បើលើសពី ១ ចំណុច ជាសញ្ញា
// ច្បាស់លាស់ថាមានការច្រឡំដោយប្រភពក្រៅ App ====
function periodSheetTabOrderLooksScrambled_(sheets) {
  try {
    var names = sheets.map(function(s) { return s.getName(); }).filter(isPeriodSheetName_);
    var decreases = 0;
    for (var i = 1; i < names.length; i++) {
      var prev = names[i - 1].split('-'); // [dd, MM]
      var cur = names[i].split('-');
      var prevKey = Number(prev[1]) * 100 + Number(prev[0]); // MM*100+DD — ប្រៀបធៀបខែមុនថ្ងៃ
      var curKey = Number(cur[1]) * 100 + Number(cur[0]);
      if (curKey < prevKey) decreases++;
    }
    return decreases > 1;
  } catch (err) { return false; } // កំហុសណាមួយ — កុំបង្ខំរៀបចំពេញលេញដោយច្រឡំ (សុវត្ថិភាព — ធ្លាក់ចុះទៅផ្លូវលឿនធម្មតា)
}

// ==== FIX (ល្បឿន — "សូមបង្កើនល្បឿនក្នុងការបង្កើត Sheet ថ្មី")៖ deferFormulaColumnProtection (ប៉ារ៉ាម៉ែត្រទី៤ថ្មី) —
// ពេល true ស្នើឲ្យ applyPeriodSheetFormulasAndFormatting_() ខាងក្រោម រំលងជំហារ protectPeriodSheetFormulaColumns_
// (ការហៅ Sheets API ~២៥-៣០ដងចំពោះការការពារជួរឈររូបមន្តកម្រិត Range) ដោយផ្ទាល់ — ប្រើតែពី createNextDailyPeriodSheets()
// (ដែលធានាថានឹងហៅ scheduleProtectPastPeriodSheets_() ភ្លាមៗបន្ទាប់ ដែលនឹងអនុវត្តការការពារនេះឡើងវិញនៅផ្ទៃខាងក្រោយ
// ក្នុងរយៈពេលខ្លីដដែល — មិនប៉ះពាល់សុវត្ថិភាពទិន្នន័យឡើយ) — Caller ផ្សេងទៀត (ឧ. upsertDailyEntry_ បង្កើត Sheet ថ្មីតែ
// មួយ ដោយគ្មានការហៅ scheduleProtectPastPeriodSheets_ តាមក្រោយ) មិនហុចប៉ារ៉ាម៉ែត្រនេះទេ (undefined = false) ធានាថា
// ការការពារនៅតែអនុវត្តភ្លាមៗដដែលសម្រាប់ផ្លូវនោះ ====
function createDistrictPeriodSheet_(district, dateStr, skipFormulaReapplyIfExisting, deferFormulaColumnProtection) {
  var dSs = getDistrictSpreadsheet_(district);
  var name = periodSheetName_(dateStr);
  var totalCols = periodSheetTotalCols_();
  var existing = dSs.getSheetByName(name);
  if (existing) {
    // Sheet មានរួចហើយ — គ្រាន់តែជួសជុល/ធ្វើបច្ចុប្បន្នភាពរូបមន្តគណនា (មិនប៉ះពាល់ទិន្នន័យបញ្ចូលដោយដៃឡើយ)
    // ==== FIX (សុវត្ថិភាពទិន្នន័យ)៖ ត្រូវធានាថា Sheet ចាស់ៗ (បង្កើតមុនលក្ខណៈ "សរុប") មានជួរឈរគ្រប់គ្រាន់ជានិច្ច
    // មុននឹងធ្វើអ្វីផ្សេងទៀត ព្រោះការរក្សាទុកធាតុថ្មីម្តងៗ (upsertDailyEntry_) សរសេរចូលទៅតាមទីតាំងជួរឈរតាមលំដាប់
    // DAILY_FIELDS បច្ចុប្បន្ន — បើ Sheet មិនទាន់មានជួរឈរទាំងនេះ វានឹងសរសេរជាន់លើទិន្នន័យ "ចំណាំ"/"អ្នកបញ្ចូល"
    // ចាស់ដោយចៃដន្យ (នេះជាមូលហេតុនៃបញ្ហាទិន្នន័យបាត់បង់ដែលធ្លាប់កើតឡើងកាលពីមុន) ====
    try { ensurePeriodSheetHasTotalColumns_(existing); } catch (errEnsure) { try { Logger.log('ensurePeriodSheetHasTotalColumns_ error (' + district + '/' + name + '): ' + (errEnsure && errEnsure.message)); } catch (eLog) {} }
    // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ចំណុចទី៥, ជួរឈរ "ចំនួនដងកែប្រែ")៖ ដូចគ្នានឹងចំណុចខាងលើ — ត្រូវធានាថា
    // Sheet ចាស់ៗ (បង្កើតមុនលក្ខណៈនេះ) មានជួរឈរនេះជានិច្ច មុននឹងធ្វើអ្វីផ្សេងទៀត (មើលការពន្យល់ពេញលេញនៅ
    // ensurePeriodSheetHasEditCountColumn_() ខាងក្រោម) ====
    try { ensurePeriodSheetHasEditCountColumn_(existing); } catch (errEnsure2) { try { Logger.log('ensurePeriodSheetHasEditCountColumn_ error (' + district + '/' + name + '): ' + (errEnsure2 && errEnsure2.message)); } catch (eLog2) {} }
    var lastRowExisting = existing.getLastRow();
    fixPeriodSheetCodeColumn_(existing, totalCols, lastRowExisting);
    // ==== FIX (ល្បឿន x2 — "រក្សាទុកទាំងអស់")៖ fixPeriodSheetReferenceColumns_() ជានិច្ចអាន (getRange().getValues())
    // ជួរដេកទិន្នន័យទាំងមូលឡើងវិញ (ស្ទួនបេះបិទនឹងការអានដដែលដែល saveDistrictDayEntries ធ្វើភ្លាមៗបន្ទាប់ពី Function
    // នេះត្រឡប់រួច) បូកនឹងសរសេរ (setValues) ជួរឈរ "ទិន្នន័យយោង" (Baseline) ឡើងវិញទាំងមូល — សរុប ~៣ការហៅ Sheets API
    // បន្ថែម (Header ២ + Baseline ១) ក្រៅពីការអានទិន្នន័យទាំងមូល ១ដង។ Function នេះជា "ធានា"/self-heal សម្រាប់ករណី
    // កម្រ (Sheet ចាស់មុននឹងលក្ខណៈនេះមាន ឬថ្ងៃចាប់ផ្តើមប្រតិបត្តិការផ្លាស់ប្តូរ) — មិនចាំបាច់ត្រូវពិនិត្យ/សរសេរជាន់
    // ឡើងវិញរាល់ពេលចុច "រក្សាទុកទាំងអស់" ទេ ព្រោះ Save All មិនកែប្រែជួរឈរ "ទិន្នន័យយោង" ទាល់តែសោះ (សុវត្ថិភាពទិន្នន័យ
    // នៅតែដដែល — Save All អាន-កែ-សរសេរត្រឹមតែវាល DAILY_FIELDS/ចំណាំ/អ្នកបញ្ចូល/ពេលបញ្ចូលប៉ុណ្ណោះ ទុកជួរឈរ "ទិន្នន័យ
    // យោង" ដដែលនៅក្នុង Array ដែលអានចូលមក)។ ប្រើ skipFormulaReapplyIfExisting ដដែល (Caller = saveDistrictDayEntries
    // ផ្លូវលឿន) ដើម្បីរំលងជំហាននេះផងដែរ — Caller ផ្សេងទៀត (ឧ. upsertDailyEntry_ ការកែឯកតា, admin បង្កើត Sheet)
    // នៅតែហៅផ្លូវពេញ (skip=false/undefined) ធានា Self-heal នៅតែដំណើរការសម្រាប់ករណីកម្រទាំងនោះដដែល ====
    if (!skipFormulaReapplyIfExisting) fixPeriodSheetReferenceColumns_(existing, district, dateStr, lastRowExisting); // ធានាថាជួរឈរយោង + Header ត្រឹមត្រូវជានិច្ច
    // ==== FIX (ខ្សែសង្វាក់ទិន្នន័យយោង)៖ ថ្ងៃនេះទើបជួសជុលរូបមន្ត/ជួរឈរយោងឡើងវិញ (ខាងលើ) — ត្រូវបន្តជួសជុល Sheet
    // ថ្ងៃបន្តបន្ទាប់ (បើមានស្រាប់រួចហើយ) ដែលអាចនៅតែផ្អែកលើតម្លៃចាស់ (មើលការពន្យល់ពេញលេញនៅ Function ខាងក្រោម) ====
    if (!skipFormulaReapplyIfExisting) cascadeRefreshRollingReferenceForward_(district, dateStr);
    // ==== FIX (Fix122, "សូម...ការបញ្ចូលទិន្នន័យ...ពេលចុច(រក្សាទុកទាំងអស់) សូមឲ្យដំណើរការលឿនផងដែរ")៖ Caller ខ្លះ
    // (ឧ. saveDistrictDayEntries) នឹងសរសេរទិន្នន័យថ្មីទាំងមូល ហើយហៅ applyPeriodSheetFormulasAndFormatting_() ដោយ
    // ខ្លួនឯងម្តងទៀតភ្លាមៗបន្ទាប់ពី Function នេះត្រឡប់ — ធ្វើឲ្យវដ្តរូបមន្ត+ការពារជួរឈរទាំងមូល (រួមទាំង
    // protectPeriodSheetFormulaColumns_ ដែលចំណាយ Sheets API ច្រើនសម្រាប់លុប+បង្កើតការពារ៦គូឡើងវិញ) ដំណើរការ
    // ២ដងស្ទួនគ្នាក្នុងសំណើតែមួយ (ម្តងទីនេះលើ Range ចាស់ ម្តងទៀតលើ Data ថ្មីភ្លាមៗបន្ទាប់) ដោយឥតប្រយោជន៍ — លទ្ធផល
    // ទីនេះនឹងត្រូវសរសេរជាន់ត្រឡប់ភ្លាមៗនៅឡើយ។ Caller ទាំងនោះអាចហុច skipFormulaReapplyIfExisting=true ដើម្បីរំលង
    // ជំហាននេះទីនេះ (សុវត្ថិភាពទិន្នន័យនៅតែដដែល ព្រោះ Caller ធានាថានឹងហៅផ្ទាល់ខ្លួនឯងវិញភ្លាមៗ) ====
    if (lastRowExisting >= 4 && !skipFormulaReapplyIfExisting) applyPeriodSheetFormulasAndFormatting_(existing, totalCols, 4, lastRowExisting, false, deferFormulaColumnProtection);
    // ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក")៖ ធានាថា Sheet ចាស់ៗ (បង្កើតមុនលក្ខណៈនេះមាន) ក៏ទទួលបានជួរដេក "សរុប" ដែរ
    // (self-heal) ព្រមទាំងស្តារតម្លៃ SUM ត្រឹមត្រូវឡើងវិញ ក្នុងករណីជំហានខាងលើ (applyPeriodSheetFormulasAndFormatting_/
    // fixPeriodSheetReferenceColumns_ ដែលមិនស្គាល់ជួរដេកនេះ) បានសរសេរជាន់វា ដោយចៃដន្យ ជា Row ឃុំ/សង្កាត់ធម្មតា —
    // ត្រូវហៅចុងក្រោយបំផុត (ក្រោយគ្រប់ជំហានផ្សេងទៀតរួចរាល់) ដើម្បីធានាលទ្ធផលចុងក្រោយត្រឹមត្រូវជានិច្ច។ ប្រើ
    // skipFormulaReapplyIfExisting ដដែល ដូចជំហានផ្សេងទៀតខាងលើ (មិនចាំបាច់រត់ឡើងវិញរាល់ពេលរក្សាទុកទិន្នន័យធម្មតា —
    // រូបមន្ត SUM ស្វ័យគណនាដោយខ្លួនឯងជានិច្ច) ====
    if (!skipFormulaReapplyIfExisting) writePeriodSheetTotalRow_(existing, district, totalCols, getPeriodSheetCommuneLastRow_(existing));
    // ==== FIX (ល្បឿន x3 — "រក្សាទុកទាំងអស់")៖ ត្រឡប់ Object Sheet ជាក់ស្តែង (sheet:) ជាមួយផងដែរ — មិនត្រឹមតែ Metadata
    // (created/name/url) ទេ — ដើម្បីឲ្យ getOrCreatePeriodSheet_() ខាងក្រោម មិនចាំបាច់ហៅ dSs.getSheetByName() ម្តងទៀត
    // (ការហៅ Sheets API ស្ទួនចោល ព្រោះ Function នេះខ្លួនឯងហៅរួចរាល់ខាងលើ (var existing = dSs.getSheetByName(name);)
    // ស្រាប់ហើយ ដើម្បីត្រួតពិនិត្យថា Sheet មានរួចហើយឬអត់) ====
    return { created: false, name: name, url: buildSheetTabUrl_(dSs.getId(), existing), sheet: existing };
  }

  var sheet = dSs.insertSheet(name);
  buildPeriodSheetHeader_(sheet, totalCols, dateStr);

  var estimateMap = getSingleValueMapForDistrict_('estimate', district);
  var newStationMap = getSingleValueMapForDistrict_('newStation', district);
  var totalStationMap = getSingleValueMapForDistrict_('totalStation', district);
  var baselineMap = getRollingReferenceMap_(dateStr, district).map; // ទិន្នន័យ២០២៥ (ថ្ងៃទី១) ឬ update2026 ថ្ងៃមុន (ថ្ងៃទី២ តទៅ)
  var codeMap = getLinkedDataCodeMap_();

  var communes = COMMUNE_ORDER[district] || [];
  var rows = communes.map(function(c) {
    var key = district + "|" + c.name;
    var base = baselineMap[key] || { total: 0, female: 0 };
    var row = [newId_(), dateStr, codeMap[key] || "", district, c.name]; // ID, កាលបរិច្ឆេទ, លេខកូដ, ស្រុក, ឃុំ/សង្កាត់
    row.push(Number(estimateMap[c.name]) || 0);
    row.push(Number(newStationMap[c.name]) || 0);
    row.push(Number(totalStationMap[c.name]) || 0);
    row.push(base.total, base.female);
    DAILY_FIELDS.forEach(function() { row.push(0); });
    DAILY_META_SUFFIX.forEach(function() { row.push(""); });
    return row;
  });
  if (rows.length) {
    sheet.getRange(4, 1, rows.length, totalCols).setValues(rows);
    applyPeriodSheetFormulasAndFormatting_(sheet, totalCols, 4, 3 + rows.length, false, deferFormulaColumnProtection);
    // ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក")៖ បន្ថែមជួរដេក "សរុប" ស្វ័យប្រវត្តិ ភ្ជាប់ជាប់ចុងក្រោយបំផុត ភ្លាមៗពេលបង្កើត
    // Sheet ថ្មី (មើលការពន្យល់ពេញលេញនៅ writePeriodSheetTotalRow_) ====
    writePeriodSheetTotalRow_(sheet, district, totalCols, 3 + rows.length);
  }

  try { sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); }); } catch (err) {}
  clearPeriodSheetDateCache_(district); // Tab ថ្មីទើបបង្កើត — Cache កាលបរិច្ឆេទចាស់ (getPeriodSheetDateMap_) លែងត្រឹមត្រូវទៀតហើយ
  // ==== FIX (Fix118)៖ Cache ប៊ូតុង "បើក Google Sheet" (getSheetAccessInfo) ក៏លែងត្រឹមត្រូវដែរ ពេល Tab ថ្មីទើបបង្កើត
  // (ឧ. ថ្ងៃនេះទើបមាន Sheet — ត្រូវបង្ហាញប៊ូតុងភ្លាមៗ មិនរង់ចាំ Cache ចាស់ផុតកំណត់ដល់ ៦ម៉ោងទេ) ====
  clearSheetAccessCache_(district);
  // ==== FIX (សំណើថ្មី "រៀបតាមលំដាប់ថ្ងៃខែ")៖ Google insertSheet() ដាក់ Tab ថ្មីនៅ "ចុងបញ្ជី" ជានិច្ច ដោយមិនគិតពី
  // លំដាប់ថ្ងៃខែទេ — បើអ្នកប្រើបង្កើត Sheet ថ្មីមិនតាមលំដាប់ (ឧ. រំលងថ្ងៃណាមួយ រួចមកបំពេញវិញក្រោយពេលថ្ងៃក្រោយៗមានរួចហើយ ឬ
  // ប្រើប៊ូតុង "កំណត់ថ្ងៃប្រជុំ" រំលងថ្ងៃណាមួយ) នោះ Tab នឹងលេចធាតុមិនតាមលំដាប់ (ឧ. 14-09, 18-09, 15-09, 16-09)។ ត្រូវ
  // រៀបចំតាមលំដាប់ថ្ងៃខែឡើងវិញភ្លាមៗ ជារៀងរាល់ពេលមាន Sheet ថ្មី ដើម្បីងាយស្រួលរកមើល (មិនប៉ះពាល់ទិន្នន័យក្នុង Sheet ណាមួយ
  // ឡើយ — ប្តូរតែលំដាប់ទីតាំង Tab ប៉ុណ្ណោះ)។
  // ==== FIX (ល្បឿន — "រៀបតាមលំដាប់ថ្ងៃខែ" ធ្វើឲ្យបង្កើត Sheet ថ្មីយឺតខ្លាំង)៖ ការហៅ getPeriodSheetDateMap_(district)
  // ដោយផ្ទាល់ត្រង់នេះ គឺជានិច្ចជា Cache Miss (ព្រោះ clearPeriodSheetDateCache_() ទើបហៅរួច ២បន្ទាត់ខាងលើ) — ធ្វើឲ្យត្រូវ
  // អាន Sheet ចាស់ៗ រាល់ Sheet (getLastRow + getRange(4,2).getValue() ម្តងមួយៗ) ដែលយឺតខ្លាំង (រាប់សិប/រាប់រយ Sheets
  // API Call) បើស្រុកនោះមាន Period Sheet ជាច្រើនរួចមកហើយ — សម្រាប់ករណីធម្មតាស្ទើរតែ១០០% (បន្ថែមថ្ងៃថ្មីតាមលំដាប់
  // ធម្មតា — កាលបរិច្ឆេទថ្មីនេះធំជាង/ស្មើនឹងកាលបរិច្ឆេទ Sheet ចុងក្រោយដែលមានស្រាប់ជានិច្ច) Sheet ថ្មីនេះ (ដែល insertSheet()
  // ដាក់នៅចុងបញ្ជីរួចហើយ) ស្ថិតនៅត្រឹមត្រូវតាមលំដាប់រួចហើយ ដោយមិនចាំបាច់ធ្វើអ្វីបន្ថែមទៀតទាល់តែសោះ — ត្រូវពិនិត្យត្រឹមតែ Sheet
  // មួយ (Sheet ដែលនៅមុន Sheet ថ្មីនេះ) ថាកាលបរិច្ឆេទចាស់ជាង/ស្មើ dateStr ថ្មីដែរឬអត់ (អាន Sheets API តែ១ដងប៉ុណ្ណោះ)។
  // ការហៅ getPeriodSheetDateMap_()/reorderPeriodSheetTabsInSpreadsheet_() ពេញលេញ (ថ្លៃថ្កាស — O(n) Sheets API
  // Call) ត្រូវការតែក្នុងករណីកម្រ (បំពេញទិន្នន័យថ្ងៃចាស់ត្រលប់ក្រោយ ក្រោយពីមានថ្ងៃក្រោយៗរួចហើយ) ប៉ុណ្ណោះ ====
  try {
    var allSheetsNow_ = dSs.getSheets();
    var newSheetIdx_ = allSheetsNow_.length - 1; // Sheet ថ្មីតែងតែនៅចុងបញ្ជី (insertSheet())
    var needFullReorder_ = false;
    if (newSheetIdx_ > 0) {
      var prevSheet_ = allSheetsNow_[newSheetIdx_ - 1];
      if (isPeriodSheetName_(prevSheet_.getName())) {
        var prevLastRow_ = prevSheet_.getLastRow();
        if (prevLastRow_ >= 4) {
          var prevDate_ = normalizeDateStr_(prevSheet_.getRange(4, 2).getValue());
          if (prevDate_ && prevDate_ > dateStr) needFullReorder_ = true; // Sheet ថ្មីមានកាលបរិច្ឆេទចាស់ជាង Sheet មុន — ចាំបាច់រៀបលំដាប់ពេញលេញ
        }
      }
    }
    // ==== FIX (រាយការណ៍ថ្មី "ការរៀបលំដាប់ Tab នៅតែខុសដដែល ត្រូវចុច 'រៀបតាមលំដាប់ថ្ងៃខែ' ដោយដៃរាល់ដង")៖ ការត្រួតពិនិត្យ
    // ខាងលើ ពិនិត្យតែ Sheet ១ (ដែលនៅជាប់មុន Sheet ថ្មី) ប៉ុណ្ណោះ — បើ Tab ផ្សេងទៀត (មិនជាប់នឹង Sheet ថ្មី) ធ្លាប់ត្រូវបាន
    // ប្តូរលំដាប់ដោយដៃផ្ទាល់ក្នុង Google Sheet (អូស Tab ដោយខ្លួនឯង — កើតឡើងក្រៅ App ទាំងស្រុង) មុននឹងមាន Sheet ថ្មីនេះ
    // ការត្រួតពិនិត្យតែ Sheet ១ ខាងលើ នឹងមិនដែលចាប់បានករណីនេះទេ (ព្រោះ Sheet ថ្មីតែងតែមានកាលបរិច្ឆេទចុងក្រោយបំផុតជានិច្ច
    // ស្របតាមលក្ខខណ្ឌចូល Function នេះ (getNextExpectedPeriodDate_ ខាងលើ) ដូច្នេះលក្ខខណ្ឌ "prevDate_ > dateStr" ស្ទើរតែ
    // មិនដែលពិតទេ ទោះ Tab ពីមុនៗច្រឡំកន្លែងណាមួយក៏ដោយ) — Tab ចាស់ៗនៅតែច្រឡំជាអចិន្ត្រៃយ៍ រហូតដល់មាននរណាម្នាក់ចុច
    // "រៀបតាមលំដាប់ថ្ងៃខែ" ដោយដៃ (Tab ការកំណត់)។ FIX៖ បន្ថែមការត្រួតពិនិត្យ "សម្រាប់ភាពសមហេតុផលទូទៅ" មួយបន្ថែម —
    // ឥតគិតថ្លៃ Sheets API ទាល់តែសោះ (ប្រើតែ sheet.getName() ដែលទទួលបានរួចហើយពី allSheetsNow_ ខាងលើ — មិនចាំបាច់
    // អាន getRange(4,2) បន្ថែមទេ) — ស្កេនឈ្មោះ Tab ទាំងអស់ (ទម្រង់ "dd-MM") តាមលំដាប់ជាក់ស្តែងបច្ចុប្បន្ន រាប់ចំនួន
    // ចំណុចដែលខែ/ថ្ងៃ "ថយក្រោយ" ធៀបនឹង Tab មុន — ក្នុងលំដាប់ត្រឹមត្រូវ (បង្កើតតាមលំដាប់ថ្ងៃជានិច្ច) គួរតែមានយ៉ាងច្រើន
    // តែ ១ ចំណុចប៉ុណ្ណោះ (ឆ្លងកាត់ចុងឆ្នាំ ធ្នូ→មករា ស្របច្បាប់) — បើលើសពី ១ ចំណុច មានន័យថា Tab ត្រូវបានប្តូរលំដាប់ដោយដៃ
    // ប្រាកដជាមិនមែនស្របនឹងលំដាប់ការបង្កើតធម្មតាទេ ត្រូវការរៀបចំពេញលេញឡើងវិញ ====
    if (!needFullReorder_ && periodSheetTabOrderLooksScrambled_(allSheetsNow_)) needFullReorder_ = true;
    if (needFullReorder_) {
      reorderPeriodSheetTabsInSpreadsheet_(dSs, getPeriodSheetDateMap_(district));
      // ==== FIX (ខ្សែសង្វាក់ទិន្នន័យយោង — ករណីបំពេញទិន្នន័យថ្ងៃចាស់ត្រលប់ក្រោយ/Backfill តែប៉ុណ្ណោះ)៖ Sheet ថ្មីនេះមាន
      // កាលបរិច្ឆេទចាស់ជាង Sheet ចុងក្រោយដែលមានស្រាប់ (ការណ៍កម្រ — កំពុងស្ថិតក្នុង if (needFullReorder_) រួចហើយ) —
      // Sheet ចាស់ៗដែលមកក្រោយកាលបរិច្ឆេទថ្មីនេះ អាចនឹងបានយក "ទិន្នន័យយោង" ខុស (ព្រោះ Sheet ថ្មីនេះមិនទាន់មាននៅពេល
      // ពួកវាត្រូវបានបង្កើត) — ត្រូវជួសជុលខ្សែសង្វាក់ថ្ងៃបន្តបន្ទាប់ពី Sheet ថ្មីនេះទៅមុខវិញ (មើលការពន្យល់ពេញលេញនៅ
      // cascadeRefreshRollingReferenceForward_)។ ដាក់នៅទីនេះ (ខាងក្នុង if ខាងលើ ក្រោយ getPeriodSheetDateMap_ រួច)
      // ដោយចេតនា — ជៀសវាងហៅ Sheets API ទាំងស្រុងក្នុងករណីធម្មតា (បន្ថែមថ្ងៃថ្មីតាមលំដាប់ធម្មតា ~១០០% នៃពេល) ដែល
      // needFullReorder_=false ជានិច្ច (Sheet ថ្មីជានិច្ចជា Sheet ចុងក្រោយបំផុត — គ្មាន Sheet អនាគតត្រូវជួសជុលទេ) —
      // FIX ពីមុន (ហៅដោយផ្ទាល់មិនគិតលក្ខខណ្ឌ ភ្លាមៗក្រោយ clearPeriodSheetDateCache_) បណ្តាលឲ្យបង្កើត Sheet ថ្មីយឺត
      // ខ្លាំងវិញម្តងទៀត (ត្រូវការ getPeriodSheetDateMap_ ដែលជានិច្ចជា Cache Miss ព្រោះទើប Clear — ដូចបញ្ហាដើមដែល
      // ជួសជុលរួចហើយក្នុង fix153) — ឥឡូវហៅតែក្នុងករណីកម្រនេះប៉ុណ្ណោះ ព្រមទាំងប្រើ Cache ដែល getPeriodSheetDateMap_
      // ខាងលើទើបញែកបំពេញរួច (មិនហៅម្តងទៀតទទេ) ====
      cascadeRefreshRollingReferenceForward_(district, dateStr);
    }
  } catch (errReorder) {}

  // ==== លាក់ជួរឈរខ្លះសម្រាប់គ្រូប្រចាំស្រុក (មើលក្នុង Google Sheet ដោយផ្ទាល់) ====
  // "ថ្ងៃចាប់ផ្តើមដំណើរការ" (កំណត់ក្នុង Tab កំណត់ប្រព័ន្ធ) → នៅតែបង្ហាញ ស្ថិតិប៉ាន់ស្មាន ។ល។ ដើម្បីឲ្យគ្រូឃើញម្តងដំបូង
  applyPeriodSheetColumnVisibility_(sheet, !isFirstOperationDay_(dateStr));

  return { created: true, name: name, url: buildSheetTabUrl_(dSs.getId(), sheet), sheet: sheet };
}

// ត្រឡប់ថា Tab មួយ ជា Period Sheet (ឈ្មោះទម្រង់ "dd-MM") ដែរឬទេ — ប្រើសម្រាប់ញែកចេញពី Tab ផ្សេងទៀត (Sheet1 ។ល។)
function getOrCreatePeriodSheet_(district, dateStr, skipFormulaReapplyIfExisting) {
  var d = normalizeDateStr_(dateStr);
  // ==== FIX (ល្បឿន x3 — "រក្សាទុកទាំងអស់")៖ ប្រើ Sheet Object ដែល createDistrictPeriodSheet_() ត្រឡប់មកផ្ទាល់
  // (r.sheet) ជំនួសការហៅ getDistrictSpreadsheet_()+dSs.getSheetByName() ម្តងទៀតនៅទីនេះ — មុននេះ Sheets API
  // getSheetByName() ត្រូវបានហៅ ២ដងស្ទួនគ្នាក្នុងសំណើតែមួយ (ម្តងក្នុង createDistrictPeriodSheet_ ខាងលើ ដើម្បីត្រួត
  // ពិនិត្យថាមានស្រាប់ឬអត់ ម្តងទៀតទីនេះ ដើម្បីទទួល Object ត្រឡប់មកវិញ) ====
  var r = createDistrictPeriodSheet_(district, d, skipFormulaReapplyIfExisting);
  return r.sheet;
}

// ស្វែងរកលេខជួរដេក (row index) ក្នុង Period Sheet មួយ ដែលឈ្មោះឃុំ/សង្កាត់ (ជួរឈរទី៤) ត្រូវគ្នា
function findCommuneRowIndex_(sheet, commune) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return -1;
  var target = String(commune || '').trim();
  var communeCol = DAILY_META_PREFIX.length; // ឃុំ/សង្កាត់ ជានិច្ចជាជួរឈរចុងក្រោយនៃ Prefix (1-indexed)
  var communes = sheet.getRange(4, communeCol, lastRow - 3, 1).getDisplayValues();
  for (var i = 0; i < communes.length; i++) {
    if (String(communes[i][0]).trim() === target) return 4 + i;
  }
  return -1;
}

// ករណីឃុំ/សង្កាត់ មិនទាន់មាននៅក្នុង COMMUNE_ORDER (ថ្មី/មិនធម្មតា) — បន្ថែមជួរដេកថ្មីនៅចុង Period Sheet
// ==== FIX (សំណើ "បូកសរុបក្រុងស្រុក")៖ បើ Sheet នេះមានជួរដេក "សរុប" ស្វ័យប្រវត្តិរួចហើយ (ភាគច្រើនតែងតែមាន — មើល
// writePeriodSheetTotalRow_) ត្រូវបញ្ចូលជួរដេកឃុំ/សង្កាត់ថ្មីនេះ "មុន" ជួរដេក "សរុប" (មិនមែនបន្ថែមខាងក្រោមវាទេ —
// បើមិនដូច្នេះទេ ជួរដេក "សរុប" លែងស្ថិតនៅចុងក្រោយបំផុតទៀតហើយ បំបែកលក្ខខណ្ឌសំខាន់ដែលកូដផ្សេងទៀតទុកចិត្ត) រួចសរសេរ
// រូបមន្ត SUM របស់ជួរដេក "សរុប" ឡើងវិញ ឲ្យគ្របដណ្តប់ដល់ជួរដេកឃុំ/សង្កាត់ថ្មីនេះផងដែរ ====
function appendCustomCommuneRow_(sheet, district, dateStr, commune) {
  var code = (getLinkedDataCodeMap_()[district + "|" + commune]) || "";
  var row = [newId_(), dateStr, code, district, commune];
  PERIOD_EXTRA_FIELDS.forEach(function() { row.push(0); });
  DAILY_FIELDS.forEach(function() { row.push(0); });
  DAILY_META_SUFFIX.forEach(function() { row.push(''); });

  var rawLastRow = sheet.getLastRow();
  if (rawLastRow >= 4 && isPeriodSheetTotalRow_(sheet, rawLastRow)) {
    var newRowIdx = rawLastRow; // ចូលជំនួសទីតាំងជួរដេក "សរុប" ចាស់ (ដែលរុញចុះក្រោម ១ ដោយស្វ័យប្រវត្តិ)
    sheet.insertRowBefore(newRowIdx);
    sheet.getRange(newRowIdx, 1, 1, row.length).setValues([row]);
    // ==== ចំណាំ៖ មិនចាំបាច់ហៅ applyPeriodSheetFormulasAndFormatting_() លើជួរដេកថ្មីនេះទីនេះទេ — Caller
    // (upsertDailyEntry_) សរសេរតម្លៃពិតប្រាកដជាន់ត្រឡប់ភ្លាមៗ ហើយហៅ Function នោះដោយខ្លួនឯងស្រាប់ (ដូចគ្នានឹងផ្លូវចាស់
    // ដែលគ្រាន់តែ appendRow() ទុកតម្លៃ ០ ចាំ Caller សរសេរជាន់ដដែរ) — ត្រូវការត្រឹមតែ writePeriodSheetTotalRow_
    // ខាងក្រោម ដើម្បីឲ្យរូបមន្ត SUM គ្របដណ្តប់ដល់ជួរដេកនេះផងប៉ុណ្ណោះ ====
    var totalCols = periodSheetTotalCols_();
    writePeriodSheetTotalRow_(sheet, district, totalCols, newRowIdx);
    return newRowIdx;
  }

  sheet.appendRow(row);
  return sheet.getLastRow();
}

// សម្អាតជួរដេកមួយ ក្នុង Period Sheet ត្រឡប់ទៅតម្លៃ ០/ទទេវិញ (ប្រើសម្រាប់ "លុប" ធាតុមួយ ឬពេលផ្លាស់ទី ID ទៅជួរដេកផ្សេង)
// ID/កាលបរិច្ឆេទ/ស្រុក/ឃុំសង្កាត់ និងជួរឈរ "ទិន្នន័យយោង" (PERIOD_EXTRA_FIELDS) មិនត្រូវប៉ះពាល់ទេ — លុបតែផ្នែកទិន្នន័យបញ្ចូលប្រចាំថ្ងៃប៉ុណ្ណោះ
function resetEntryRow_(sheet, rowIndex) {
  var zeros = DAILY_FIELDS.map(function() { return 0; });
  sheet.getRange(rowIndex, periodFieldColIndex_(DAILY_FIELDS[0].key), 1, DAILY_FIELDS.length).setValues([zeros]);
  var noteCol = periodFieldColIndex_(DAILY_FIELDS[DAILY_FIELDS.length - 1].key) + 1;
  // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ")៖ ការលុបធាតុមួយ ត្រូវចាត់ទុកថាជាការ "សម្អាត"
  // ជួរដេកនោះទាំងស្រុងឲ្យត្រឡប់ទៅស្ថានភាព "មិនទាន់ធ្លាប់មានទិន្នន័យ" វិញ (ដូចគ្នានឹង ចំណាំ/អ្នកបញ្ចូល/ពេលបញ្ចូល ដែល
  // ត្រូវបានសម្អាតរួចហើយខាងក្រោម) ដូច្នេះ "ចំនួនដងកែប្រែ" ក៏ត្រូវកំណត់ត្រឡប់ទៅ ០ វិញផងដែរ (មិនមែនកកកុញបន្តទៅមុខ
  // ហួសពីការលុប) ====
  sheet.getRange(rowIndex, noteCol, 1, 4).setValues([['', '', '', 0]]);
  applyPeriodSheetFormulasAndFormatting_(sheet, periodSheetTotalCols_(), rowIndex, rowIndex);
}

// ==================== ៦.២ រូបមន្តគណនាស្វ័យប្រវត្តិ ក្នុង Sheet ថ្មីតាមថ្ងៃខែ (Period Sheet) ====================
// កំណត់លេខជួរឈរ (1-indexed) នៃវាល DAILY_FIELDS ជាក់លាក់មួយ ក្នុង Period Sheet
// (រចនាសម្ព័ន្ធ Period Sheet = DAILY_META_PREFIX[4] + PERIOD_EXTRA_FIELDS[5] + DAILY_FIELDS[...] + DAILY_META_SUFFIX[4])
// ==== FIX (ល្បឿន — "សូមបង្កើនល្បឿនក្នុងការបង្កើត Sheet ថ្មី")៖ deferColumnProtectionOnly (ប៉ារ៉ាម៉ែត្រទី៦ថ្មី) — ខុសពី
// skipProtection (ប៉ារ៉ាម៉ែត្រទី៥ ដែលរំលងទាំងស្រុងសូម្បីតែការសរសេររូបមន្ត/ការតុបតែង — ប្រើសម្រាប់ជួរដេកចាស់ដែលធ្លាប់
// មានទម្រង់រួចហើយ) — deferColumnProtectionOnly រំលងតែជំហារ protectPeriodSheetFormulaColumns_() ចុងក្រោយបំផុត
// ប៉ុណ្ណោះ (ការការពារកម្រិត Range) ខណៈនៅតែសរសេររូបមន្ត + ពណ៌ផ្ទៃខាងក្រោយ/ទ្រេត + លាក់ជួរឈរស្ទួន ដូចធម្មតា — ត្រូវការ
// សម្រាប់ជួរដេកថ្មីដែលមិនធ្លាប់មានទម្រង់ពីមុនសោះ (ការតុបតែងត្រូវធ្វើភ្លាមៗ តែការការពារអាចពន្យារបន្តិចបាន) ====
function applyPeriodSheetFormulasAndFormatting_(sheet, totalCols, firstRow, lastRow, skipProtection, deferColumnProtectionOnly) {
  if (!sheet || lastRow < firstRow) return;
  var n = lastRow - firstRow + 1;

  var cAddNewT = colLetter_(periodFieldColIndex_('addNew_total')),           cAddNewF = colLetter_(periodFieldColIndex_('addNew_female'));
  var cAddTrInT = colLetter_(periodFieldColIndex_('addTransferIn_total')),   cAddTrInF = colLetter_(periodFieldColIndex_('addTransferIn_female'));
  var cAddChgT = colLetter_(periodFieldColIndex_('addChangeStn_total')),    cAddChgF = colLetter_(periodFieldColIndex_('addChangeStn_female'));
  var cAddTotT = periodFieldColIndex_('addTotal_total'),                    cAddTotF = periodFieldColIndex_('addTotal_female');
  var cDelDeathT = colLetter_(periodFieldColIndex_('delDeath_total')),      cDelDeathF = colLetter_(periodFieldColIndex_('delDeath_female'));
  var cDelTrOutT = colLetter_(periodFieldColIndex_('delTransferOut_total')),cDelTrOutF = colLetter_(periodFieldColIndex_('delTransferOut_female'));
  var cDelDupT = colLetter_(periodFieldColIndex_('delDuplicate_total')),    cDelDupF = colLetter_(periodFieldColIndex_('delDuplicate_female'));
  var cDelChgOutT = periodFieldColIndex_('delChangeStnOut_total'),          cDelChgOutF = periodFieldColIndex_('delChangeStnOut_female');
  var cDelTotT = periodFieldColIndex_('delTotal_total'),                    cDelTotF = periodFieldColIndex_('delTotal_female');
  var cUpd26T = periodFieldColIndex_('update2026_total'),                   cUpd26F = periodFieldColIndex_('update2026_female');
  var cBase25T = colLetter_(periodExtraColIndex_('baseline2025_total')),    cBase25F = colLetter_(periodExtraColIndex_('baseline2025_female'));
  var cIdNewT = colLetter_(periodFieldColIndex_('idNew_total')),           cIdNewF = colLetter_(periodFieldColIndex_('idNew_female'));
  var cIdTrInT = colLetter_(periodFieldColIndex_('idTransferIn_total')),   cIdTrInF = colLetter_(periodFieldColIndex_('idTransferIn_female'));
  var cIdTotT = periodFieldColIndex_('idTotal_total'),                     cIdTotF = periodFieldColIndex_('idTotal_female');
  var cCorrT = colLetter_(periodFieldColIndex_('correction_total')),       cCorrF = colLetter_(periodFieldColIndex_('correction_female'));
  var cBioT = colLetter_(periodFieldColIndex_('biometric_total')),        cBioF = colLetter_(periodFieldColIndex_('biometric_female'));
  var cCbtT = periodFieldColIndex_('correctionBiometricTotal_total'),     cCbtF = periodFieldColIndex_('correctionBiometricTotal_female');

  var addTotalF = [], delChgOutF = [], delTotalF = [], upd26F = [], idTotalF = [], cbTotalF = [];
  for (var i = 0; i < n; i++) {
    var r = firstRow + i;
    addTotalF.push(['=' + cAddNewT + r + '+' + cAddTrInT + r + '+' + cAddChgT + r, '=' + cAddNewF + r + '+' + cAddTrInF + r + '+' + cAddChgF + r]);
    delChgOutF.push(['=' + cAddChgT + r, '=' + cAddChgF + r]);
    delTotalF.push([
      '=' + cDelDeathT + r + '+' + cDelTrOutT + r + '+' + cDelDupT + r + '+' + colLetter_(cDelChgOutT) + r,
      '=' + cDelDeathF + r + '+' + cDelTrOutF + r + '+' + cDelDupF + r + '+' + colLetter_(cDelChgOutF) + r
    ]);
    upd26F.push([
      '=' + cBase25T + r + '+' + colLetter_(cAddTotT) + r + '-' + colLetter_(cDelTotT) + r,
      '=' + cBase25F + r + '+' + colLetter_(cAddTotF) + r + '-' + colLetter_(cDelTotF) + r
    ]);
    idTotalF.push(['=' + cIdNewT + r + '+' + cIdTrInT + r, '=' + cIdNewF + r + '+' + cIdTrInF + r]);
    // ==== ជួរឈរ "សរុប" ថ្មី = ចំនួនករណីកែតម្រូវសរុប + ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ (រូបមន្តផ្ទាល់ក្នុង Sheet —
    // សូមមើលមូលហេតុលម្អិតនៅ DAILY_FIELDS ក្នុង Utils.gs) ====
    cbTotalF.push(['=' + cCorrT + r + '+' + cBioT + r, '=' + cCorrF + r + '+' + cBioF + r]);
  }

  // ==== FIX (ល្បឿន x5 — "រក្សាទុកទាំងអស់" នៅតែយឺត)៖ ជួរឈររូបមន្ត ៣គូ (delChangeStnOut/delTotal/update2026 —
  // delChgOutF/delTotalF/upd26F ខាងលើ) ស្ថិតនៅជាប់គ្នាដោយផ្ទាល់ក្នុង Period Sheet ជានិច្ច (គ្មានជួរឈរផ្សេងបំបែក —
  // សូមមើលលំដាប់ក្នុង DAILY_FIELDS ក្នុង Utils.gs៖ delChangeStnOut → delTotal → update2026 ជាប់ៗគ្នា ៦ជួរឈរ)
  // ដូច្នេះអាចផ្គុំទិន្នន័យទាំង៣គូ ជា Range តែមួយ (៦ជួរឈរ) ហើយហៅ setFormulas() តែម្តងគត់ ជំនួសការហៅ Sheets API
  // ដាច់ដោយឡែក៣ដងដូចមុន (កាត់បន្ថយការហៅមួយទៀត ទាំងផ្លូវលឿន (skipProtection) និងផ្លូវពេញ)។ ត្រួតពិនិត្យ Adjacency
  // ជាមុនសិន (Defense in depth — បើលំដាប់ជួរឈរផ្លាស់ប្តូរនាពេលអនាគត ធ្លាក់ចុះទៅហៅដាច់ដោយឡែក៣ដងវិញដោយស្វ័យប្រវត្តិ
  // មិនសរសេរទិន្នន័យខុសជួរឈរដោយអចេតនាឡើយ) ====
  var canMergeMidF_ = (cDelTotT === cDelChgOutT + 2) && (cUpd26T === cDelChgOutT + 4);
  var mergedMidF_ = null;
  if (canMergeMidF_) {
    mergedMidF_ = [];
    for (var mi = 0; mi < n; mi++) {
      mergedMidF_.push(delChgOutF[mi].concat(delTotalF[mi], upd26F[mi]));
    }
  }

  // ==== FIX (ល្បឿន — "រក្សាទុកទាំងអស់")៖ ការគ្រប់គ្រងការការពារជួរឈររូបមន្ត (លុបជាមុន + protectPeriodSheetFormulaColumns_
  // បង្កើតជាថ្មីវិញ) ត្រូវការហៅ Sheets API ជាបន្តបន្ទាប់ជាច្រើនដង (protect/getEditors/removeEditors/setDomainEdit/
  // addEditor × ៦គូជួរឈររូបមន្ត — រហូតដល់ ~៣០ការហៅ) ដែលយឺតខ្លាំង ជាពិសេសពេលហៅពី saveDistrictDayEntries() (ចុច
  // "រក្សាទុកទាំងអស់") ។ Function នោះសរសេរជាន់ត្រឡប់លើជួរដេកដែលមានស្រាប់ស្រាប់ជានិច្ច (មិនដែលបន្ថែមជួរដេកថ្មីតាម
  // ផ្លូវនេះទេ — ការបន្ថែមជួរដេកថ្មី កើតឡើងតែតាមរយៈ appendCustomCommuneRow_ ក្នុង upsertDailyEntry_ ប៉ុណ្ណោះ ដែលនៅតែ
  // ហៅ Function នេះជាផ្លូវពេញ — skipProtection មិនកំណត់ទេ) ដូច្នេះជួរដេកដែលត្រូវការពារ (4→lastRow) មិនផ្លាស់ប្តូរ
  // ធៀបនឹងលើកមុនទេ — ការការពារដែលមានស្រាប់រួចហើយ (បង្កើតដោយការហៅលើកមុន) នៅតែត្រឹមត្រូវ អាចរំលងជំហាននេះបាន ដើម្បី
  // បង្កើនល្បឿន — Caller ហៅដោយ skipProtection=true ។ ការពារកុំឲ្យខូច៖ បើសរសេររូបមន្តបរាជ័យ (ករណីកម្រ — ឧ. Sheet ចាស់
  // ដែលមិនទាន់មានការការពារត្រឹមត្រូវ ធ្វើឲ្យសរសេរជាន់លើជួរឈរការពារបរាជ័យ) ត្រលប់ទៅផ្លូវពេញ (លុប+ការពារឡើងវិញ) វិញ
  // ដោយស្វ័យប្រវត្តិខាងក្រោម ====
  // ==== FIX (ល្បឿន x2 បន្ថែម — "រក្សាទុកទាំងអស់")៖ ផ្លូវលឿននេះ ពីមុនក៏ហៅ setBackground('#f1f5f9')/setFontStyle
  // ('italic') ១២ដង (៦គូ x ២) + hideColumns() ១ដង រាល់ពេលរក្សាទុក (សរុប ១៣ការហៅ Sheets API បន្ថែម) ។ ការតុបតែងទាំង
  // នេះ (ពណ៌ផ្ទៃខាងក្រោយ/ទ្រេត/លាក់ជួរឈរស្ទួន) ជាទម្រង់ (Format) មិនមែនតម្លៃ (Value/Formula) ទេ — setValues()/
  // setFormulas() មិនជះឥទ្ធិពល/លុបចោលទម្រង់ដែលមានស្រាប់ឡើយ ហើយជួរដេកទាំងនេះមានស្រាប់ស្រាប់ជានិច្ច (ដូចខាងលើ —
  // មិនដែលជួរដេកថ្មីតាមផ្លូវនេះទេ) ដូច្នេះទម្រង់ត្រូវបានកំណត់រួចជាស្រេចតាំងពីការហៅផ្លូវពេញលើកដំបូង (ពេលបង្កើត Sheet/
  // ជួរដេកនោះ) ហើយនៅតែត្រឹមត្រូវជានិច្ច — មិនចាំបាច់សរសេរជាន់ម្តងទៀតរាល់ពេលរក្សាទុកទេ (មានតែរូបមន្ត/setFormulas
  // ប៉ុណ្ណោះ ដែលត្រូវសរសេរជាថ្មីជានិច្ច ព្រោះ setValues() ខាងលើ (Caller) សរសេរជាន់ត្រឡប់លើក្រឡារូបមន្តទាំងនោះជាមួយ
  // តម្លៃគណនាស្ថិតិចុងក្រោយ (មិនមែនរូបមន្តទេ) ជានិច្ច — ត្រូវស្តាររូបមន្តវិញ) ====
  if (skipProtection) {
    try {
      sheet.getRange(firstRow, cAddTotT, n, 2).setFormulas(addTotalF);
      if (canMergeMidF_) {
        sheet.getRange(firstRow, cDelChgOutT, n, 6).setFormulas(mergedMidF_);
      } else {
        sheet.getRange(firstRow, cDelChgOutT, n, 2).setFormulas(delChgOutF);
        sheet.getRange(firstRow, cDelTotT, n, 2).setFormulas(delTotalF);
        sheet.getRange(firstRow, cUpd26T, n, 2).setFormulas(upd26F);
      }
      sheet.getRange(firstRow, cIdTotT, n, 2).setFormulas(idTotalF);
      sheet.getRange(firstRow, cCbtT, n, 2).setFormulas(cbTotalF);
      return;
    } catch (errFast) {
      // ធ្លាក់ចុះទៅផ្លូវពេញខាងក្រោម (កម្រណាស់ — គ្រាន់តែធានាសុវត្ថិភាព)
    }
  }

  // ==== លុបការការពារជួរឈររូបមន្តចាស់ជាមុនសិន (បើមាន) មុននឹងសរសេររូបមន្តជាថ្មី — ជៀសវាងកំហុសសិទ្ធិកែប្រែពេលសរសេរជាន់
  // ត្រូវបានបង្កើតឡើងវិញភ្លាមៗ ខាងក្រោម (protectPeriodSheetFormulaColumns_) ក្រោយពេលរូបមន្តត្រូវបានសរសេររួច ====
  try { sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); }); } catch (err) {}
  sheet.getRange(firstRow, cAddTotT, n, 2).setFormulas(addTotalF);
  if (canMergeMidF_) {
    sheet.getRange(firstRow, cDelChgOutT, n, 6).setFormulas(mergedMidF_);
  } else {
    sheet.getRange(firstRow, cDelChgOutT, n, 2).setFormulas(delChgOutF);
    sheet.getRange(firstRow, cDelTotT, n, 2).setFormulas(delTotalF);
    sheet.getRange(firstRow, cUpd26T, n, 2).setFormulas(upd26F);
  }
  sheet.getRange(firstRow, cIdTotT, n, 2).setFormulas(idTotalF);
  sheet.getRange(firstRow, cCbtT, n, 2).setFormulas(cbTotalF);

  // សម្គាល់ជួរឈររូបមន្ត (មិនត្រូវកែដោយដៃ) ជាពណ៌ប្រផេះ + ទ្រេត
  [[cAddTotT,2],[cDelChgOutT,2],[cDelTotT,2],[cUpd26T,2],[cIdTotT,2],[cCbtT,2]].forEach(function(pair) {
    sheet.getRange(firstRow, pair[0], n, pair[1]).setBackground('#f1f5f9').setFontStyle('italic');
  });

  // លាក់ជួរឈរ "ស្ទួន" សិន (តាមសំណើ)
  try {
    var dupCol = periodFieldColIndex_('delDuplicate_total');
    sheet.hideColumns(dupCol, 2);
  } catch (err) {}

  // ==== ការពារជួរឈររូបមន្តទាំង៦គូ (Range Protect) ពីការកែប្រែផ្ទាល់ដោយគ្រូប្រចាំក្រុងស្រុក ក្នុង Google Sheet —
  // ត្រូវប្រើ sheet.getLastRow() ជាក់ស្តែងបច្ចុប្បន្ន (មិនមែន firstRow/lastRow ដែលបញ្ជូនចូល Function នេះទេ) ព្រោះ
  // ការហៅមួយចំនួន (ឧ. resetEntryRow_) សរសេររូបមន្តតែមួយជួរដេក ប៉ុន្តែការការពារត្រូវគ្របដណ្តប់ជួរដេកទិន្នន័យទាំងអស់
  // ជានិច្ច (មិនមែនតែជួរដេកដែលទើបសរសេរនេះទេ) — ==== FIX (ល្បឿន — "សូមបង្កើនល្បឿនក្នុងការបង្កើត Sheet ថ្មី")៖ Caller
  // អាចហុច deferColumnProtectionOnly=true ដើម្បីរំលងជំហាននេះ (មើលការពន្យល់ពេញលេញនៅកន្លែងប្រកាស Function ខាងលើ) ====
  if (!deferColumnProtectionOnly) protectPeriodSheetFormulaColumns_(sheet);
}

// ==== ជួរឈររូបមន្តទាំង៦គូ (សរុប/total+ស្រី/female = ១២ជួរឈរ) ដែលត្រូវការពារ — សូមមើលរូបមន្តជាក់ស្តែងខាងលើ
// (applyPeriodSheetFormulasAndFormatting_) ====
var FORMULA_COL_PROTECTION_DESC_ = "ចាក់សោស្វ័យប្រវត្តិ — ជួរឈររូបមន្ត (កុំកែ)";
var FORMULA_COL_FIELD_KEYS_ = [
  'addTotal_total', 'delChangeStnOut_total', 'delTotal_total',
  'update2026_total', 'idTotal_total', 'correctionBiometricTotal_total'
];

// ==== ការពារ Google Sheet កម្រិត Range៖ ជួរឈររូបមន្ត (សរុប) ៦គូ ក្នុង Period Sheet មួយ — កុំឲ្យគ្រូប្រចាំក្រុងស្រុក
// កែប្រែផ្ទាល់ (បង្កចន្លោះទិន្នន័យខុសពីការគណនាជាក់ស្តែង) ខណៈជួរឈរបញ្ចូលទិន្នន័យឆៅផ្សេងទៀត នៅតែអាចកែបានធម្មតា។
// ត្រូវបានហៅរាល់ដងដែល applyPeriodSheetFormulasAndFormatting_ ដំណើរការ (ចំណុចចូលតែមួយគត់ ដែលសរសេររូបមន្តទាំង
// អស់) ដូច្នេះការការពារនេះកកើត/ធ្វើបច្ចុប្បន្នភាពដោយស្វ័យប្រវត្តិ ទាំង Sheet ចាស់ (ពេលចូលប្រើ/រក្សាទុកលើកក្រោយ តាមរយៈ
// createDistrictPeriodSheet_) និង Sheet ថ្មី — ដោយមិនចាំបាច់រត់ Function ដាច់ដោយឡែកទេ។ ==== ចំណាំសំខាន់៖ ដកសិទ្ធិកែប្រែ
// អ្នកកែប្រែទាំងអស់ចេញ ហើយទុកតែម្ចាស់ Script (Admin ដែល Deploy) ប៉ុណ្ណោះដែលអាចកែបាន ដូចគ្នានឹងគោលការណ៍ដែលប្រើ
// រួចហើយសម្រាប់ការចាក់សោ Sheet ទាំងមូល (applyOrRemoveSheetProtection_) ====
// ==== FIX (ល្បឿន — "បង្កើត Sheet ថ្មីនៅតែយឺត ១០នាទីជាង")៖ ពីមុន Function នេះហៅ .protect() ដាច់ដោយឡែក ៦ដង (មួយក្នុង
// មួយគូជួរឈររូបមន្ត) ក្នុងនោះមួយចំនួន (delChangeStnOut/delTotal/update2026) ជាប់គ្នាដោយផ្ទាល់ក្នុង Sheet ស្រាប់ (ដូច
// ការសង្កេតឃើញរួចហើយសម្រាប់ setFormulas() — មើល canMergeMidF_ ក្នុង applyPeriodSheetFormulasAndFormatting_) ប៉ុន្តែ
// ការការពារនៅតែធ្វើដាច់ដោយឡែក៦ដងដដែល (មិនបានផ្គុំ) ធ្វើឲ្យខាតបង់ការហៅ Sheets API ដោយឥតប្រយោជន៍ (protect/getEditors/
// removeEditors/setDomainEdit/addEditor ~៥ការហៅក្នុងមួយគូ × ៦គូ = ~៣០ការហៅ ក្នុងមួយស្រុក × ១០ស្រុក = ~៣០០ការហៅ
// រាល់ពេលបង្កើត Sheet ថ្ងៃថ្មី — នេះជាហេតុផលចម្បងមួយនៃរយៈពេល ១០នាទីជាងដែលរាយការណ៍មក ព្រោះ Range.protect() និង
// Function ពាក់ព័ន្ធរបស់វា ដឹងស្គាល់ថាយឺតជាងការហៅ Sheets API ធម្មតាច្រើន)។ FIX៖ ផ្គុំគូជួរឈររូបមន្តដែលជាប់គ្នាដោយ
// ផ្ទាល់ (គណនាដោយស្វ័យប្រវត្តិ តាមលេខជួរឈរជាក់ស្តែង — មិនកំណត់ដាច់ស្រេចតាមកូដ ២-៣ គូណាមួយទេ ដូច្នេះបើលំដាប់ជួរឈរ
// ផ្លាស់ប្តូរនាពេលអនាគត នៅតែផ្គុំត្រឹមត្រូវដោយស្វ័យប្រវត្តិ ឬធ្លាក់ចុះទៅជា Range ដាច់ដោយឡែកវិញដោយសុវត្ថិភាព — មិនដែល
// ការពារខុសជួរឈរ ឬលែងការពារជួរឈរណាមួយឡើយ) ទៅជា Range តែមួយ រួចហៅ .protect()/removeEditors/setDomainEdit/
// addEditor តែម្តងគត់ក្នុងមួយ Range ដែលបានផ្គុំ — ជួយកាត់បន្ថយពី ៦ Range មកនៅសល់ត្រឹមតែ ៤ Range (~៣៣% តិចជាង)
// ដោយគ្មានប៉ះពាល់សុវត្ថិភាព/លទ្ធផលការពារឡើយ (គ្របដណ្តប់ជួរឈរដូចគ្នាបេះបិទ គ្រាន់តែជា Range តូចជាងចំនួនច្រើនជាង ក្លាយ
// ជា Range ធំជាងចំនួនតិចជាង) ====
function periodSheetFormulaColProtectionRanges_() {
  var cols = FORMULA_COL_FIELD_KEYS_.map(function(key) { return periodFieldColIndex_(key); })
      .filter(function(c) { return c !== -1; })
      .sort(function(a, b) { return a - b; });
  var ranges = [];
  var i = 0;
  while (i < cols.length) {
    var startCol = cols[i];
    var endCol = startCol + 1; // គូនេះខ្លួនឯង (សរុប+ស្រី = ២ជួរឈរ)
    var j = i + 1;
    while (j < cols.length && cols[j] === endCol + 1) { // គូបន្ទាប់ចាប់ផ្តើមភ្លាមៗបន្ទាប់ពីគូនេះចប់ — ជាប់គ្នាដោយផ្ទាល់
      endCol = cols[j] + 1;
      j++;
    }
    ranges.push({ col: startCol, width: endCol - startCol + 1 });
    i = j;
  }
  return ranges;
}

function protectPeriodSheetFormulaColumns_(sheet) {
  try {
    // ==== FIX (សំខាន់ណាស់ — Bug កកកុញការការពារ)៖ Function នេះត្រូវបានហៅជារៀងរាល់ថ្ងៃដោយ Trigger ស្វ័យប្រវត្តិ
    // (protectPastPeriodSheets_) និងរាល់ពេលចុចប៊ូតុង Admin (protectAllPeriodSheetFormulaColumnsNow) ដោយផ្ទាល់
    // (មិនកន្លងតាម applyPeriodSheetFormulasAndFormatting_ ដែលជម្រះការការពារចាស់ចេញរួចទេ)។ Range.protect() របស់ Google
    // តែងតែបង្កើត Protection object ថ្មីរាល់ដង ទោះបីជា Range នោះមានការការពាររួចហើយក៏ដោយ — បើមិនជម្រះការការពារចាស់
    // (ដែលបានបង្កើតដោយ Function នេះផ្ទាល់ កំណត់ដោយ Description ដូចគ្នា) ជាមុនសិនទេ ការការពារនឹងកកកុញកើនឡើងរាល់ថ្ងៃ
    // សម្រាប់ Sheet ចាស់ៗ (រាប់សិបគូក្នុងមួយ Sheet ក្រោយពីរយៈពេលខ្លី) ធ្វើឲ្យប្រតិបត្តិការ Sheets API យឺត ហើយចុងក្រោយ
    // អាចនាំឲ្យ .protect() បរាជ័យស្ងាត់ស្ងៀម (ជាន់ Try/Catch) ដោះលែងការការពារពិតប្រាកដដោយអចេតនា។ ដូច្នេះ ត្រូវជម្រះ
    // ការការពារចាស់ (ត្រឹមតែដែលផលិតដោយ Function នេះ សម្គាល់ដោយ Description — មិនប៉ះពាល់ការការពារផ្សេងទៀត) ជានិច្ច
    // ត្រង់ចំណុចចូលរួមមួយនេះ ដើម្បីឲ្យគ្រប់ Caller ទាំងអស់ទទួលបានអត្ថប្រយោជន៍នេះដោយស្វ័យប្រវត្តិ ====
    try {
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
        try { if (p.getDescription() === FORMULA_COL_PROTECTION_DESC_) p.remove(); } catch (eRm) {}
      });
    } catch (eList) {}

    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return; // គ្មានជួរដេកទិន្នន័យទេ — មិនចាំបាច់ការពារ
    var editorEmail = "";
    try { editorEmail = Session.getEffectiveUser().getEmail(); } catch (e0) {}
    periodSheetFormulaColProtectionRanges_().forEach(function(g) {
      var protection = sheet.getRange(4, g.col, lastRow - 3, g.width).protect().setDescription(FORMULA_COL_PROTECTION_DESC_);
      try { protection.removeEditors(protection.getEditors()); } catch (e1) {}
      try { protection.setDomainEdit(false); } catch (e2) {}
      try { if (editorEmail) protection.addEditor(editorEmail); } catch (e3) {}
    });
  } catch (err) {}
}

// ==================== ៦.៣ ធានាថា Period Sheet មួយ មានជួរឈរ "សរុប" (correctionBiometricTotal) គ្រប់គ្រាន់ និងនៅត្រឹមទីតាំងត្រឹមត្រូវ ====================
// ==== សំខាន់ណាស់ (ប្រវត្តិកំហុស)៖ កំណែមុន (fix90) សរសេរ Header/រូបមន្តជាន់ដោយផ្ទាល់ទៅលើទីតាំងជួរឈរ "ចំណាំ"/
// "អ្នកបញ្ចូល" ចាស់ ដោយមិនបានផ្លាស់ទីទិន្នន័យទាំងនោះជាមុន (insertColumnsBefore) ធ្វើឲ្យទិន្នន័យ "ចំណាំ"/"អ្នកបញ្ចូល"
// ដែលធ្លាប់មានស្រាប់ត្រូវបានលុបចោល។ Function នេះជាកំណែកែតម្រូវ៖ ប្រើ sheet.insertColumnsBefore()/moveColumns()
// ដើម្បីរុញ/ផ្លាស់ទីទិន្នន័យចាស់ដោយសុវត្ថិភាព ជំនួសឲ្យការសរសេរជាន់លើគ្នា។ ត្រូវបានហៅពី ២កន្លែង៖ (១)
// createDistrictPeriodSheet_() គ្រប់ពេលចូលប្រើ/រក្សាទុក Sheet ថ្ងៃណាមួយ (ស្វ័យប្រវត្តិទាំងស្រុង) និង (២)
// migratePeriodSheetsAddTotalColumnHeaders_() សម្រាប់ជួសជុលជាបាច់ធំម្តងតែម្តង។
// ==== FIX (កំណែថ្មី — ជួរឈរ "សរុប" ផ្លាស់ទីតាំងទៅជាប់នឹង "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ")៖ ការវិភាគស្ថានភាព Sheet
// ឥឡូវមិនអាស្រ័យលើថាតើ "ចំណាំ" ស្ថិតនៅ totalColF+1 ជានិច្ចទៀតទេ (សន្មតថាជួរឈរ "សរុប" នៅចុងគេ) — ជំនួសមក ស្វែងរក
// ទីតាំង "ពេលបញ្ចូល" ជាក់ស្តែងជា Anchor សិន (មិនប្រែប្រួលដោយថាតើ Sheet នេះនៅលំដាប់ចាស់ ឬថ្មី) រួចទើបវិភាគថា ជួរឈរ
// "សរុប" ត្រូវការ (A) បញ្ចូលថ្មីទាំងស្រុង (Sheet ចាស់មិនទាន់មានវាសោះ) (B) ជួសជុលការសរសេរជាន់ចាស់ (D) ផ្លាស់ទីតាំង
// ដោយ moveColumns() (Sheet មានវារួច ប៉ុន្តែនៅតែស្ថិតនៅចុងគេ តាមលំដាប់ចាស់មុននឹងកំណែនេះ) ឬ (C) មិនធ្វើអ្វីទាំងអស់
// (ត្រឹមត្រូវរួចហើយ)។ moveColumns() ផ្លាស់ទីទិន្នន័យ/រូបមន្តទាំងអស់ក្នុងជួរឈរនោះទៅជាមួយ មិនមែនលុបចោលទេ ====
// ត្រឡប់ { action: 'none' | 'inserted_safe' | 'moved_to_new_position' | 'header_rewritten' | 'repaired_partial_loss' | 'unknown_skip' }
function ensurePeriodSheetHasTotalColumns_(sheet) {
  var totalColT = periodFieldColIndex_('correctionBiometricTotal_total');
  var totalCols = periodSheetTotalCols_();
  var noteLabel = DAILY_META_SUFFIX[0], enteredByLabel = DAILY_META_SUFFIX[1], enteredAtLabel = DAILY_META_SUFFIX[2];
  // ជួរឈរចុងក្រោយបំផុតនៃ DAILY_FIELDS បច្ចុប្បន្ន (មិនអាស្រ័យលើថាតើជួរឈរ "សរុប" ស្ថិតនៅចុងគេ ឬកណ្តាលក្នុងលំដាប់
  // បច្ចុប្បន្នឡើយ) — "ចំណាំ" តែងតែស្ថិតជាប់បន្ទាប់ពីជួរឈរនេះជានិច្ច
  var lastFieldCol = periodFieldColIndex_(DAILY_FIELDS[DAILY_FIELDS.length - 1].key);
  var expectedNoteCol = lastFieldCol + 1;
  var fTotal = DAILY_FIELDS.filter(function(f) { return f.key === 'correctionBiometricTotal_total'; })[0];
  var fFemale = DAILY_FIELDS.filter(function(f) { return f.key === 'correctionBiometricTotal_female'; })[0];

  var lastCol = sheet.getLastColumn();

  // ករណី C៖ ត្រឹមត្រូវរួចហើយ (ជួរឈរគ្រប់គ្រាន់ "ចំណាំ" ស្ថិតត្រឹមត្រូវ **និង** Header ជួរឈរ "សរុប" ជាក់ស្តែងស្ថិត
  // នៅទីតាំង totalColT ត្រឹមត្រូវផងដែរ) — Idempotent។ ==== សំខាន់៖ ត្រូវពិនិត្យ Header ត្រង់ totalColT ជាមួយផងដែរ
  // មិនអាចទុកចិត្តតែលើ ចំនួនជួរឈរសរុប + ទីតាំង "ចំណាំ" ម្នាក់ឯងបានទេ ព្រោះ Sheet ចាស់ (កំណែមុន correctionBiometricTotal
  // ផ្លាស់មកកណ្តាល — ត្រូវការ moveColumns/Case D) មានចំនួនជួរឈរសរុប **ដូចគ្នាបេះបិទ** ហើយ "ចំណាំ" ក៏ស្ថិតត្រង់
  // ទីតាំងដូចគ្នាបេះបិទដែរ (ព្រោះចំនួនវាលសរុប 34 មិនប្រែប្រួល មានតែលំដាប់ខាងក្នុងប្រែប្រួល) — បើគ្មានការពិនិត្យ Header
  // បន្ថែមនេះទេ Case D នឹងមិនដែលត្រូវបានរកឃើញ (ត្រូវបានរកឃើញ ហើយកែតម្រូវ ដោយការធ្វើតេស្តលម្អិត មុននឹង Deploy) ====
  if (lastCol >= totalCols && sheet.getRange(3, expectedNoteCol).getDisplayValue() === noteLabel
      && sheet.getRange(1, totalColT).getDisplayValue() === (fTotal.group || '—')) {
    return { action: 'none' };
  }

  // ==== ស្វែងរកទីតាំង "ពេលបញ្ចូល" ជាក់ស្តែងបច្ចុប្បន្ន — Anchor តែមួយគត់ដែលអាចទុកចិត្តបាន ដើម្បីវិភាគស្ថានភាពពិត
  // របស់ Sheet នេះ ដោយមិនអាស្រ័យលើថាតើជួរឈរ "សរុប" ស្ថិតនៅចុងគេ (លំដាប់ចាស់) ឬកណ្តាល (លំដាប់ថ្មី) ====
  var scanWidth = Math.max(lastCol, totalCols, expectedNoteCol) + 5;
  var enteredAtCol = -1;
  for (var c = 1; c <= scanWidth; c++) {
    if (sheet.getRange(3, c).getDisplayValue() === enteredAtLabel) { enteredAtCol = c; break; }
  }
  if (enteredAtCol === -1) {
    // រកមិនឃើញ "ពេលបញ្ចូល" ទាល់តែសោះ — រចនាសម្ព័ន្ធមិនធម្មតាខ្លាំង (ប្រហែល Sheet ត្រូវបានកែប្រែដោយដៃ) កុំប៉ះពាល់អ្វីទាំងអស់
    return { action: 'unknown_skip' };
  }

  try { sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); }); } catch (e0) {}

  var enteredByCol = enteredAtCol - 1;
  var noteCol = enteredAtCol - 2;
  var suffixLabelsIntact = noteCol >= 1
      && sheet.getRange(3, enteredByCol).getDisplayValue() === enteredByLabel
      && sheet.getRange(3, noteCol).getDisplayValue() === noteLabel;

  if (!suffixLabelsIntact) {
    // ករណី B៖ "ចំណាំ"/"អ្នកបញ្ចូល" ត្រង់មុន "ពេលបញ្ចូល" លែងត្រឹមត្រូវទៀតហើយ — ប្រហែលមុខងារ Migration ចាស់ (មានបញ្ហា)
    // បាន Run ជាន់លើវារួច (ទិន្នន័យ "ចំណាំ"/"អ្នកបញ្ចូល" ចាស់ ប្រហែលបានបាត់បង់ រកមកវិញបានតែពី Version History ប៉ុណ្ណោះ)។
    // ត្រង់នេះគ្រាន់តែជួសជុលរចនាសម្ព័ន្ធបន្តទៅមុខ ដោយផ្អែកលើ "ពេលបញ្ចូល" ដែលនៅសល់ (ជាធម្មតាមិនត្រូវបានប៉ះពាល់
    // ព្រោះមានតែ ២ជួរឈរដំបូងប៉ុណ្ណោះដែលត្រូវបានសរសេរជាន់) ដោយមិនប៉ះពាល់ដល់វាទៀតឡើយ
    if (sheet.getRange(3, enteredByCol).getDisplayValue() === enteredByLabel) {
      // "អ្នកបញ្ចូល" នៅសល់ (មិនទាន់ត្រូវប៉ះពាល់) — ខ្វះតែជួរឈរ "ចំណាំ" ១ប៉ុណ្ណោះ
      sheet.insertColumnsBefore(enteredByCol, 1);
      sheet.getRange(1, enteredByCol, 3, 1).setValues([['—'], ['—'], [noteLabel]]);
    } else {
      // ទាំង "ចំណាំ" និង "អ្នកបញ្ចូល" ត្រូវបានសរសេរជាន់រួច (ទិន្នន័យទាំង២នេះបាត់បង់) — បញ្ចូលជួរឈរទទេ ២ថ្មី មុន "ពេលបញ្ចូល"
      sheet.insertColumnsBefore(enteredAtCol, 2);
      sheet.getRange(1, enteredAtCol, 3, 2).setValues([['—', '—'], ['—', '—'], [noteLabel, enteredByLabel]]);
    }
    // ==== "ចំណាំ"/"អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ឥឡូវត្រឹមត្រូវហើយ — ត្រូវហៅខ្លួនឯងម្តងទៀត ដើម្បីវិភាគ/ជួសជុលជួរឈរ "សរុប"
    // ខ្លួនឯងដាច់ដោយឡែក (A/D/គ្មានអ្វី) ដោយផ្អែកលើទីតាំង "ចំណាំ" ដ៏ត្រឹមត្រូវថ្មីនេះ ====
    var recursed = ensurePeriodSheetHasTotalColumns_(sheet);
    return { action: (recursed.action === 'none' || recursed.action === 'header_rewritten') ? 'repaired_partial_loss' : recursed.action };
  }

  // (fTotal/fFemale ត្រូវបានគណនារួចហើយពីខាងលើ សម្រាប់ការពិនិត្យលឿន Case C)
  // ==== "ចំណាំ"/"អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ឥឡូវត្រឹមត្រូវ (Anchor ជឿទុកចិត្តបាន) — ពិនិត្យ Header ផ្ទាល់ (មិនមែនស្មាន
  // តាមការគណនាទីតាំងទេ ដើម្បីភាពត្រឹមត្រូវ ១០០%) នៅទីតាំង "ជាប់មុនចំណាំដោយផ្ទាល់ ២ជួរឈរ" — ដែលជាទីតាំងចាស់ដែល
  // ជួរឈរ "សរុប" ធ្លាប់ស្ថិតនៅជានិច្ចមុននឹងកំណែនេះ (ចាប់តាំងពីវាធ្លាប់ត្រូវបានបន្ថែម "នៅចុងគេ" ជានិច្ច) ====
  var oldTotalColT = noteCol - 2;
  var groupAtOldPos = (oldTotalColT >= 1) ? sheet.getRange(1, oldTotalColT).getDisplayValue() : '';

  var action;
  if (oldTotalColT === totalColT) {
    // ទីតាំងចាស់ (ជាប់មុនចំណាំ) ដូចគ្នាទៅនឹងទីតាំងគោលដៅថ្មីស្រាប់ (ចំនួនជួរឈរគ្រប់គ្រាន់ត្រឹមត្រូវរួច) ប៉ុន្តែមកដល់
    // ត្រង់នេះមានន័យថា Header ត្រង់នោះប្រហែលមិនត្រឹមត្រូវ (ករណីកម្រ) — សរសេរ Header ជាថ្មីខាងក្រោមដោយសុវត្ថិភាព
    // (មិនប៉ះពាល់ជួរដេកទិន្នន័យ ព្រោះ setFormulas សរសេរជាន់ត្រឹមតែជួរឈររូបមន្តដដែលនេះប៉ុណ្ណោះ)
    action = 'header_rewritten';
  } else if (groupAtOldPos === (fTotal.group || '')) {
    // ករណី D៖ ជួរឈរ "សរុប" មានរួចហើយ (ពេញលេញ) ប៉ុន្តែនៅតែស្ថិតនៅទីតាំងចាស់ (ជាប់មុន "ចំណាំ" ដោយផ្ទាល់ — Sheet ដែល
    // បានបង្កើត/បញ្ចូលទិន្នន័យរួច មុននឹងកំណែផ្លាស់ទីតាំងលើកនេះ) — ផ្លាស់ទី (moveColumns) ២ជួរឈរនោះ មករកទីតាំងថ្មី
    // ដោយ **រក្សាទុកទិន្នន័យ/រូបមន្តទាំងអស់ក្នុងជួរឈរនោះទាំងស្រុង** (moveColumns ផ្លាស់ទីទិន្នន័យទៅជាមួយ មិនមែនលុប
    // ចោលទេ)
    sheet.moveColumns(sheet.getRange(1, oldTotalColT, sheet.getMaxRows(), 2), totalColT);
    action = 'moved_to_new_position';
  } else if (sheet.getRange(1, totalColT).getDisplayValue() === (fTotal.group || '—')) {
    // ==== ការការពារ Idempotency (ស្វែងឃើញតាមរយៈការធ្វើតេស្តលម្អិត)៖ ជួរឈរ "សរុប" ត្រឹមត្រូវមានវត្តមានស្រាប់ត្រង់
    // ទីតាំងគោលដៅ (totalColT) រួចទៅហើយ ទោះបីជាទីតាំង "ជាប់មុនចំណាំ" (oldTotalColT) មិនត្រូវគ្នាក៏ដោយ (ឧទាហរណ៍៖
    // ក្រោយករណី B ជួសជុល ចំណាំ/អ្នកបញ្ចូល ដោយ insertColumnsBefore, អាចមានចន្លោះទំនេរ ១ ឬ ២ ជួរឈរ "កំសល់"
    // (Orphan) នៅចន្លោះ DAILY_FIELDS ចុងក្រោយ និង ចំណាំ) — ក្នុងករណីនេះ មិនត្រូវបញ្ចូលជួរឈរ "សរុប" ថ្មីទៀត ឬផ្លាស់ទី
    // អ្វីទាំងអស់ទេ (នឹងធ្វើឲ្យ Sheet កើនទទឹងឥតឈប់ឈរ រាល់ពេលហៅ Function នេះម្តងៗ — Bug ធ្ងន់ធ្ងរ) — គ្រាន់តែសរសេរ
    // Header/រូបមន្តជាថ្មីដោយសុវត្ថិភាព (Idempotent) ដូច Case ខាងលើ ====
    action = 'header_rewritten';
  } else if (oldTotalColT >= 1) {
    // ករណី A៖ Sheet ចាស់ទាំងស្រុង មិនទាន់មានជួរឈរ "សរុប" សោះឡើយ (ទីតាំង "ជាប់មុនចំណាំ" ជាជួរឈរដទៃទៀត ដូចជា
    // idNew_total/residenceCert_female ។ល។ មិនមែន "សរុប" ទេ) — បញ្ចូលជួរឈរទទេ ២ថ្មីត្រង់ទីតាំងគោលដៅ ដោយសុវត្ថិភាព
    // (insertColumnsBefore រុញ "ចំណាំ"/"អ្នកបញ្ចូល"/"ពេលបញ្ចូល" និងជួរឈរដទៃទៀត ដែលនៅចន្លោះនោះ ទៅខាងស្តាំ)
    sheet.insertColumnsBefore(totalColT, 2);
    action = 'inserted_safe';
  } else {
    // ស្ថានភាពមិនធម្មតា (Sheet ចង្អៀតជាងសូម្បីតែ "គ្មានជួរឈរសរុបទាំងស្រុង" — ប្រហែលខ្វះជួរឈរដទៃទៀតទៀត) — កុំប៉ះពាល់
    // អ្វីទាំងអស់ ត្រូវពិនិត្យដោយដៃ
    return { action: 'unknown_skip' };
  }

  // សរសេរក្បាលជួរឈរ "សរុប" ត្រឹមត្រូវ (ទីតាំងនេះឥឡូវជាទទេ ព្រោះទើបបញ្ចូល/ផ្លាស់ទីថ្មី ឬកំពុងតែត្រឹមត្រូវរួចហើយ)
  sheet.getRange(1, totalColT, 1, 2).setValues([[fTotal.group || '—', fFemale.group || '—']]);
  sheet.getRange(2, totalColT, 1, 2).setValues([[fTotal.sub || '—', fFemale.sub || '—']]);
  sheet.getRange(3, totalColT, 1, 2).setValues([[fTotal.unit || fTotal.label, fFemale.unit || fFemale.label]]);
  sheet.getRange(1, totalColT, 3, 2)
      .setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  try { sheet.getRange(1, totalColT, 1, 2).merge(); } catch (e1) {}
  sheet.setColumnWidths(totalColT, 2, 95);

  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) applyPeriodSheetFormulasAndFormatting_(sheet, periodSheetTotalCols_(), 4, lastRow);

  return { action: action };
}

// ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាពគ្រប់ការបញ្ចូល/កែប្រែទិន្នន័យ" — ចំណុចទី៥, កំណែកែតម្រូវតាមសំណើផ្ទាល់របស់អ្នកប្រើប្រាស់)៖
// អ្នកប្រើប្រាស់បដិសេធការសាងសង់ Sheet "កំណត់ត្រាសវនកម្ម" ដាច់ដោយឡែក (ការព្យាយាមលើកដំបូង) ហើយស្នើសុំជំនួសវិញនូវ
// ជួរឈរបន្ថែម ១ ថ្មីតែប៉ុណ្ណោះ ("ចំនួនដងកែប្រែ") ដាក់បញ្ចូលដោយផ្ទាល់នៅចុងគេនៃតារាងទិន្នន័យប្រចាំថ្ងៃដែលមានស្រាប់រួចហើយ
// (ដូចគ្នានឹង "ចំណាំ"/"អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ដែលមានរួចស្រាប់ — មើល DAILY_META_SUFFIX ក្នុង Utils.gs)។ ==== ខុសពី
// ensurePeriodSheetHasTotalColumns_() ខាងលើ (ជួរឈរ "សរុប" ត្រូវផ្លាស់ទីពីចុងគេទៅកណ្តាល DAILY_FIELDS ដ៏ស្មុគស្មាញ)
// ជួរឈរនេះ គ្រាន់តែបន្ថែមទៅ "ចុងគេបំផុត" ជានិច្ច (ព្រោះ DAILY_META_SUFFIX ជាក្រុមចុងក្រោយគេនៃលំដាប់ជួរឈរទាំងមូល)
// ដូច្នេះមិនចាំបាច់ moveColumns()/insertColumnsBefore() ដែលស្មុគស្មាញ ហើយប្រថុយប្រថានជាងទេ — គ្រាន់តែសរសេរទៅជួរឈរ
// បន្ទាប់ពី lastCol បច្ចុប្បន្នដោយផ្ទាល់ (setValues លើសពីទទឹងបច្ចុប្បន្ន ធ្វើឲ្យ Google Sheets ពង្រីកទទឹង Sheet ស្វ័យប្រវត្តិ)
// — សុវត្ថិភាពខ្ពស់ជាង ធានាថាមិនដែលរុញ/ផ្លាស់ទីជួរឈរទិន្នន័យផ្សេងទៀតណាមួយឡើយ ==== សុវត្ថិភាព (Idempotent +
// Anchor-checked, ដូចគ្នានឹងគោលការណ៍ ensurePeriodSheetHasTotalColumns_ ខាងលើ)៖
//   (១) ត្រឹមត្រូវរួចហើយ (Header ជួរឈរចុងគេ = ស្លាក "ចំនួនដងកែប្រែ" រួចហើយ) → Idempotent ១០០% អាចហៅច្រើនដងដោយសុវត្ថិភាព។
//   (២) ធម្មតា (Sheet ចាស់ មិនទាន់មានជួរឈរនេះ) → ត្រូវផ្ទៀងផ្ទាត់ "Anchor" ជាមុនសិន៖ ជួរឈរចុងគេបច្ចុប្បន្ន (lastCol)
//       ត្រូវតែមាន Header ត្រូវនឹង "ពេលបញ្ចូល" ពិតប្រាកដ (DAILY_META_SUFFIX[2]) ទើបទុកចិត្តបានថា Sheet នេះ មាន
//       រចនាសម្ព័ន្ធធម្មតា សុវត្ថិភាពសម្រាប់បន្ថែមជួរឈរថ្មីភ្លាមៗនៅចុងគេ (Header ៣ជួរដេក + តម្លៃ ០ លំនាំដើមសម្រាប់
//       ជួរដេកឃុំ/សង្កាត់ទាំងអស់ដែលមានរួចហើយ — ជៀសវាងក្រឡាទទេ ដែលអាចធ្វើឲ្យតក្កវិជ្ជា "អាន+បូក+សរសេរ" (read-
//       increment-write) ក្រោយមកច្រឡំ ព្រោះ "" ខុសពី 0 ក្នុងការប្រៀបធៀបមួយចំនួន)។
//   (៣) មិនធម្មតា (Anchor មិនត្រូវគ្នា — ប្រហែល Sheet ត្រូវបានកែប្រែដោយដៃ/រចនាសម្ព័ន្ធខុសប្រក្រតី) → រំលងដោយសុវត្ថិភាព
//       (unknown_skip) មិនប៉ះពាល់អ្វីទាំងអស់ ====
function ensurePeriodSheetHasEditCountColumn_(sheet) {
  var editCountLabel = DAILY_META_SUFFIX[DAILY_META_SUFFIX.length - 1]; // "ចំនួនដងកែប្រែ"
  var enteredAtLabel = DAILY_META_SUFFIX[2]; // "ពេលបញ្ចូល" — Anchor ដែលត្រូវពឹងផ្អែក (ជួរឈរចុងគេចាស់ មុនបន្ថែមថ្មី)
  var lastCol = sheet.getLastColumn();

  if (lastCol >= 1 && sheet.getRange(3, lastCol).getDisplayValue() === editCountLabel) {
    return { action: 'none' }; // ត្រឹមត្រូវរួចហើយ — Idempotent
  }

  if (lastCol < 1 || sheet.getRange(3, lastCol).getDisplayValue() !== enteredAtLabel) {
    // Anchor មិនត្រូវគ្នា (Sheet ចង្អៀតខុសពីរំពឹងទុក ឬរចនាសម្ព័ន្ធមិនធម្មតា) — កុំប៉ះពាល់អ្វីទាំងអស់
    return { action: 'unknown_skip' };
  }

  try { sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); }); } catch (e0) {}

  var newCol = lastCol + 1;
  sheet.getRange(1, newCol, 3, 1).setValues([['—'], ['—'], [editCountLabel]]);
  sheet.getRange(1, newCol, 3, 1)
      .setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setColumnWidth(newCol, 95);

  var lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    var numRows = lastRow - 3;
    var zeros = [];
    for (var i = 0; i < numRows; i++) zeros.push([0]);
    sheet.getRange(4, newCol, numRows, 1).setValues(zeros);
  }

  return { action: 'appended' };
}

// ត្រូវហៅពីប៊ូតុង "បង្កើត Sheet ថ្មីបន្ទាប់" លើ App — មានតែ SuperAdmin/Admin/PEC21 ប៉ុណ្ណោះទើបហៅបាន
// (ក៏ប្រើសម្រាប់ "ជួសជុល" បន្ថែមរូបមន្ត ក្នុង Sheet ដែលបានបង្កើតរួចហើយផងដែរ ដោយគ្រាន់តែហៅម្តងទៀតលើកាលបរិច្ឆេទដដែល)
// ==== FIX (សំណើថ្មី "សូមធ្វើការបង្ហាញលក្ខណៈនៃការបង្កើតដូចជានៅក្នុងការ Reset ទិន្នន័យដែរ")៖ ដូចគ្នានឹង resetAllData()
// (Settings.gs — CACHE_KEY_RESET_PROGRESS_) ដែលបង្កើត Google Sheet ស្រុកទាំង១០ជាប់គ្នាដែរ (Apps Script មិនអាច
// ដំណើរការស្របគ្នាបានទេ ក្នុងការហៅតែម្តងគត់) ដូច្នេះប្រើគំរូដូចគ្នា៖ រក្សាទុក "វឌ្ឍនភាព" (ស្រុកណាហើយ ប៉ុន្មានហើយ) ទៅ
// CacheService ភ្លាមៗក្រោយស្រុកនីមួយៗចប់ — Client ស្ទង់មើល (Poll) តាមរយៈ getCreateNextPeriodSheetsProgress() ជា
// ទៀងទាត់ ខណៈកំពុងរង់ចាំ Call ចម្បង (createNextDailyPeriodSheets) បញ្ចប់ ដើម្បីបង្ហាញវឌ្ឍនភាពជាក់ស្តែង ជំនួសការ
// ស្ងាត់ស្ងៀមទាំងស្រុង (ដូចគ្នានឹង "កំពុងបង្កើត..." ដែលមានស្រាប់លើប៊ូតុង ប៉ុន្តែមិនប្រាប់ថាដល់ស្រុកណាហើយ) ====
var CACHE_KEY_CREATE_SHEETS_PROGRESS_ = 'createSheetsProgress_v1';
var CACHE_TTL_CREATE_SHEETS_PROGRESS_ = 600; // វិនាទី (១០នាទី — លើសពេលវេលាដំណើរការជាក់ស្តែងច្រើន សម្រាប់ជាកម្រិតសុវត្ថិភាព)

function getCreateNextPeriodSheetsProgress(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };
  try {
    var raw = CacheService.getScriptCache().get(CACHE_KEY_CREATE_SHEETS_PROGRESS_);
    if (!raw) return { success: true, done: 0, total: DISTRICT_LIST.length, district: '', finished: false };
    var p = JSON.parse(raw);
    return { success: true, done: p.done || 0, total: p.total || DISTRICT_LIST.length, district: p.district || '', finished: !!p.finished };
  } catch (err) {
    return { success: true, done: 0, total: DISTRICT_LIST.length, district: '', finished: false };
  }
}

// បង្កើត Tab កាលបរិច្ឆេទថ្មី ក្នុង Spreadsheet ស្រុកទាំង១០ ក្នុងពេលតែមួយ
function createNextDailyPeriodSheets(currentUsername, dateStr, sessionToken, confirmSkipGaps) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (role !== ROLE_SUPERADMIN && role !== ROLE_ADMIN && role !== ROLE_PEC21) {
    return { success: false, message: "មានតែ SuperAdmin, Admin, ឬ PEC21 ប៉ុណ្ណោះ ដែលអាចបង្កើត Sheet ថ្មីនេះបាន!" };
  }
  // ==== FIX (វឌ្ឍនភាព)៖ សម្អាត "វឌ្ឍនភាព" ចាស់ (finished:true ពីលើកមុន) ជានិច្ច ភ្លាមៗពេលហៅ Function នេះ (មិនថា
  // ជោគជ័យ/ជាប់ការត្រួតពិនិត្យខាងក្រោមឬអត់) ដើម្បីកុំឲ្យ Client ស្ទង់មើលឃើញសារ "ចប់ហើយ" ចាស់ភ្លាមៗ ខណៈកំពុងចាប់ផ្តើម
  // ការហៅថ្មី (ការត្រួតពិនិត្យខាងក្រោម ធម្មតារហ័សខ្លាំង — មិនចាំបាច់ហៅ Sheets API ណាមួយទេ) ====
  try { CacheService.getScriptCache().put(CACHE_KEY_CREATE_SHEETS_PROGRESS_, JSON.stringify({ done: 0, total: DISTRICT_LIST.length, district: '', finished: false }), CACHE_TTL_CREATE_SHEETS_PROGRESS_); } catch (eProg0) {}
  var d = normalizeDateStr_(dateStr) || todayDateStr_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return { success: false, message: "កាលបរិច្ឆេទមិនត្រឹមត្រូវទេ!" };

  // ==== FIX (សំណើថ្មី "ថ្ងៃប្រជុំ")៖ ថ្ងៃដែលបានសម្គាល់ថារំលងរួចហើយ មិនអនុញ្ញាតឲ្យបង្កើត Sheet ដោយផ្ទាល់ទេ
  // (ត្រូវដកការសម្គាល់ជាមុនសិន បើពិតជាត្រូវការ Sheet សម្រាប់ថ្ងៃនេះជាក់ស្តែង) ====
  var skipInfo_ = isDateSkipped_(d);
  if (skipInfo_) {
    return {
      success: false,
      message: "ថ្ងៃ \"" + d + "\" ត្រូវបានសម្គាល់ថា \"" + (skipInfo_.note || 'ថ្ងៃប្រជុំ') + " - គ្មានទិន្នន័យ\" រួចហើយ! បើត្រូវការ Sheet សម្រាប់ថ្ងៃនេះជាក់ស្តែង សូមដកការសម្គាល់ចេញជាមុនសិន (Tab \"ការកំណត់ប្រព័ន្ធ\")។"
    };
  }

  // ==== FIX (សំណើថ្មី "ការបង្កើត Sheet ត្រូវរៀបតាមលំដាប់ថ្ងៃខែ ជាបន្តបន្ទាប់ ដូចជា 14-09→15-09→16-09→17-09 —
  // មិនត្រូវឲ្យបង្កើត Sheet ថ្មី ដាក់តាមចិត្តទេ")៖ ត្រួតពិនិត្យថា d ស្មើនឹង Sheet ចុងក្រោយបំផុតដែលមានស្រាប់ (ករណី
  // ចង់ "ជួសជុល" Sheet បច្ចុប្បន្ន — មើល Comment ខាងលើ) ឬ Sheet ចុងក្រោយ+១ថ្ងៃ (ករណីធម្មតា — បង្កើត Sheet ថ្ងៃបន្ទាប់)
  // ប៉ុណ្ណោះ — បដិសេធភ្លាមៗ (មុននឹងធ្វើអ្វីទាំងអស់ — មិនចាំបាច់បើក Spreadsheet ស្រុកទាំង១០ ដើម្បីព្យាយាម "ជួសជុល"
  // ដោយអចេតនាទេ) បើកាលបរិច្ឆេទមិនត្រូវលំដាប់ — ដោះស្រាយទាំង ២ បញ្ហាក្នុងពេលតែមួយ៖ (១) ទប់ស្កាត់ការរំលងថ្ងៃ/ការច្រឡំ
  // កាលបរិច្ឆេទ (ឧ. ជ្រើសតាមកាលបរិច្ឆេទផ្សេងក្នុង App ដែលមិនទាក់ទង) ដែលពីមុនអាចនាំឲ្យទាញទិន្នន័យ "យោង" ខុស ព្រោះ
  // Sheet ថ្មីមិនមែនពិតជាបន្តពី Sheet ចុងក្រោយ (២) ជៀសវាងការចូលទៅដល់ Branch "Sheet មានរួចហើយ" ដ៏យឺត (ជួសជុលពេញលេញ
  // ១០ស្រុក) ដោយអចេតនា ពេលចុចខុសកាលបរិច្ឆេទចាស់ណាមួយ — នេះជាមូលហេតុនៃ "យូរបំផុត" ដែលរាយការណ៍មក ====
  var nextInfo_ = getNextExpectedPeriodDate_();
  if (!nextInfo_.isFirst && d !== nextInfo_.currentDate && d !== nextInfo_.nextDate) {
    // ==== FIX (សំណើថ្មី "ថ្ងៃប្រជុំ")៖ ថ្ងៃស្នើ នៅឆ្ងាយជាង "ថ្ងៃត្រឹមត្រូវបន្ទាប់" — អាចមានថ្ងៃប្រជុំ/ថ្ងៃគ្មានទិន្នន័យ
    // ចន្លោះកណ្តាលដែលមិនទាន់បានសម្គាល់ ជំនួសការបដិសេធផ្ទាល់ ស្នើសុំការបញ្ជាក់ពី Client ជាមុនសិន (needsGapConfirm) —
    // ប្រសិនបើបញ្ជាក់ (confirmSkipGaps=true) សម្គាល់ថ្ងៃចន្លោះទាំងអស់ថារំលងស្វ័យប្រវត្តិ រួចបន្តបង្កើត Sheet ថ្ងៃ d ====
    if (d > nextInfo_.nextDate) {
      var gapDates_ = [];
      var g_ = nextInfo_.nextDate;
      while (g_ < d) { gapDates_.push(g_); g_ = addDaysToDateStr_(g_, 1); }
      if (!confirmSkipGaps) {
        return {
          success: false, needsGapConfirm: true, gapDates: gapDates_,
          message: "ចន្លោះថ្ងៃ \"" + nextInfo_.nextDate + "\" ដល់ \"" + gapDates_[gapDates_.length - 1] + "\" (សរុប "
            + gapDates_.length + "ថ្ងៃ) មិនទាន់មាន Sheet ឬមិនទាន់បានសម្គាល់ថារំលងទេ។ ប្រសិនបើថ្ងៃទាំងនេះជាថ្ងៃប្រជុំ/"
            + "ថ្ងៃដែលមិនត្រូវបញ្ចូលទិន្នន័យពិតប្រាកដ សូមបញ្ជាក់ដើម្បីសម្គាល់ថ្ងៃទាំងនោះជា \"គ្មានទិន្នន័យ\" ស្វ័យប្រវត្តិ រួចបន្តបង្កើត Sheet ថ្ងៃ \"" + d + "\"។"
        };
      }
      gapDates_.forEach(function(gd) {
        markPeriodDateSkippedInternal_(gd, 'រំលងស្វ័យប្រវត្តិ (ពេលបង្កើត Sheet ' + d + ')', currentUsername);
      });
    } else {
      return {
        success: false,
        message: "កាលបរិច្ឆេទមិនត្រូវតាមលំដាប់ទេ! Sheet ត្រូវបង្កើតតាមលំដាប់ថ្ងៃខែជាបន្តបន្ទាប់តែប៉ុណ្ណោះ (Sheet ចុងក្រោយបំផុតបច្ចុប្បន្ន៖ \""
          + nextInfo_.currentDate + "\")។ សូមជ្រើសរើសថ្ងៃ \"" + nextInfo_.nextDate + "\" ដើម្បីបង្កើត Sheet ថ្មីបន្ទាប់ ឬថ្ងៃ \""
          + nextInfo_.currentDate + "\" ប្រសិនបើគ្រាន់តែចង់ជួសជុលរូបមន្ត Sheet បច្ចុប្បន្នវិញ។"
      };
    }
  }

  var results = [];
  var anyCreated = false;
  DISTRICT_LIST.forEach(function(district, districtIdx_) {
    try {
      // ==== FIX (ល្បឿន — "សូមបង្កើនល្បឿនក្នុងការបង្កើត Sheet ថ្មី")៖ ការការពារជួរឈររូបមន្ត ៦គូ (protectPeriodSheetFormulaColumns_
      // ក្នុង applyPeriodSheetFormulasAndFormatting_ — .protect()/.removeEditors()/.setDomainEdit()/.addEditor() ×
      // ៦គូ = ~២៥-៣០ ការហៅ Sheets API ក្នុងមួយស្រុក) ត្រូវធ្វើតែម្តងទៀតដោយ scheduleProtectPastPeriodSheets_() ខាងក្រោម
      // (protectPastPeriodSheets_ ស្កេន Tab ទាំងអស់រួមទាំង Tab ថ្មីនេះ ហើយហៅ protectPeriodSheetFormulaColumns_ លើ Tab
      // ទាំងអស់ដោយមិនគិតលក្ខខណ្ឌ — មើលមូលហេតុនៅ protectPastPeriodSheets_ ខាងលើ) ដូច្នេះការហៅផ្ទាល់ទីនេះស្ទួន ២ដងដោយឥត
      // ប្រយោជន៍ — ហុច deferFormulaColumnProtection=true ដើម្បីរំលងជំហាននេះទីនេះ (រូបមន្ត/ការតុបតែងផ្សេងទៀត នៅតែសរសេរ
      // ភ្លាមៗដដែល — មានតែការការពារកម្រិត Range ប៉ុណ្ណោះដែលពន្យារពេលបន្តិច ២-៤វិនាទី ដល់ scheduleProtectPastPeriodSheets_
      // ខាងក្រោមរត់ — មិនប៉ះពាល់សុវត្ថិភាពទិន្នន័យឡើយ ព្រោះការហៅ App (Save/Delete) ការពារដោយ LockService/App logic ខ្លួន
      // ឯង មិនអាស្រ័យលើ Range Protection នេះទេ — ការពារនេះការពារតែការកែផ្ទាល់ក្នុង Google Sheet UI ប៉ុណ្ណោះ) ====
      var r = createDistrictPeriodSheet_(district, d, false, true);
      if (r.created) anyCreated = true;
      results.push({ district: district, name: r.name, created: r.created, url: r.url });
    } catch (err) {
      results.push({ district: district, error: String(err && err.message || err) });
    }
    // ==== FIX (វឌ្ឍនភាព — "សូមធ្វើការបង្ហាញលក្ខណៈនៃការបង្កើតដូចជានៅក្នុងការ Reset ទិន្នន័យដែរ")៖ ធ្វើបច្ចុប្បន្នភាព
    // "វឌ្ឍនភាព" ភ្លាមៗក្រោយស្រុកនីមួយៗចប់ (មិនថាជោគជ័យ ឬបរាជ័យ) ដូចគ្នានឹង resetAllData() ដើម្បីឲ្យ Client ស្ទង់មើល
    // (Poll — getCreateNextPeriodSheetsProgress) ឃើញដំណើរការជាក់ស្តែង (ស្រុកណាហើយ ប៉ុន្មានហើយ) ====
    try {
      CacheService.getScriptCache().put(CACHE_KEY_CREATE_SHEETS_PROGRESS_, JSON.stringify({
        done: districtIdx_ + 1, total: DISTRICT_LIST.length, district: district, finished: false
      }), CACHE_TTL_CREATE_SHEETS_PROGRESS_);
    } catch (eProg) {}
  });
  // ==== បន្តវឺទាហារ "ថ្ងៃប្រតិបត្តិការបច្ចុប្បន្ន" ទៅជាថ្ងៃថ្មីនេះ (មិនអាស្រ័យលើម៉ោង Server ទេ) ====
  setCurrentPeriodDate_(d);
  // ==== ចាក់សោ Tab ថ្ងៃចាស់ភ្លាមៗ ព្រោះការបង្កើត Sheet ថ្ងៃថ្មី គឺជាសញ្ញាច្បាស់ថាថ្ងៃមុនបានចប់ហើយ ====
  // (មិនចាំបាច់រង់ចាំដល់ Trigger ប្រចាំថ្ងៃ ម៉ោង ១ព្រឹក ទៀតទេ)
  // ==== FIX (ល្បឿន — "សូមការបង្កើត Sheet ថ្មីលឿន")៖ ធ្វើនៅផ្ទៃខាងក្រោយវិញ (មិនរង់ចាំ) — មើលការពន្យល់ពេញលេញនៅ
  // scheduleProtectPastPeriodSheets_() — ការហៅនេះក៏ជាអ្នកទទួលបន្ទុកបំពេញការការពារជួរឈររូបមន្ត ដែលពន្យារពេលពីខាងលើ
  // ផងដែរ (deferFormulaColumnProtection=true) ====
  scheduleProtectPastPeriodSheets_();
  // ==== វឌ្ឍនភាព៖ ស្រុកទាំង១០ចប់ហើយ (មិនទាន់រាប់បញ្ចូល scheduleProtectPastPeriodSheets_/scheduleSyncSheetPermissions_
  // ដែលរត់នៅផ្ទៃខាងក្រោយផ្សេងទៀត ព្រោះ Function នេះនឹង return ភ្លាមៗខាងក្រោម — Client លែងចាំបាច់ស្ទង់មើលទៀតហើយ) ====
  try {
    CacheService.getScriptCache().put(CACHE_KEY_CREATE_SHEETS_PROGRESS_, JSON.stringify({
      done: DISTRICT_LIST.length, total: DISTRICT_LIST.length, district: '', finished: true
    }), CACHE_TTL_CREATE_SHEETS_PROGRESS_);
  } catch (eProgFin) {}
  scheduleRebuild_(); // ==== ធានាថា "សន្ទស្សន៍" ត្រូវបានធ្វើបច្ចុប្បន្នភាព ដើម្បីឲ្យតារាងខាងក្រោមមើលឃើញឃុំ/សង្កាត់ថ្មីៗភ្លាមៗ ====
  // ==== FIX៖ ធានាថា Spreadsheet ស្រុកគ្រប់ស្រុក (រួមទាំង Spreadsheet ថ្មីៗ ដែលទើបបង្កើត ក្នុងករណីប្តូរទីតាំង Sheet មេ
  // ទៅតំណភ្ជាប់ថ្មី — មើល getTargetSheetId_()/getDistrictSpreadsheet_() ក្នុង Utils.gs) ត្រូវបានចែករំលែកសិទ្ធិ Google
  // Sheet ភ្លាមៗ ជូន SuperAdmin/Admin/PEC21 + គ្រូប្រចាំស្រុកដែលពាក់ព័ន្ធ — មុននេះ ការចែករំលែកសិទ្ធិកើតឡើងតែពេលមានការ
  // ផ្លាស់ប្តូរគណនីអ្នកប្រើប្រាស់ប៉ុណ្ណោះ ដូច្នេះ Spreadsheet ស្រុកថ្មីៗ ដែលបង្កើតដោយប៊ូតុងនេះ គ្មាននរណាម្នាក់ (សូម្បីតែ
  // SuperAdmin/Admin ខ្លួនឯង) ឃើញទាល់តែសោះ រហូតដល់មានគេចុច "ធ្វើសមកាលកម្មសិទ្ធិ" ដោយដៃ ====
  // ==== FIX (ល្បឿន)៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet ជាផ្ទៃខាងក្រោយ (មិនចាំបាច់ឲ្យ Admin រង់ចាំ — មើល Utils.gs) ====
  scheduleSyncSheetPermissions_();
  return {
    success: true, sheetName: periodSheetName_(d), results: results,
    message: anyCreated
      ? ('បានបង្កើត Sheet "' + periodSheetName_(d) + '" ក្នុង Google Sheet ស្រុកទាំង១០ រួចរាល់! (ការចាក់សោ Tab ថ្ងៃចាស់ និងសិទ្ធិចូល Google Sheet នឹងត្រូវបានធ្វើសមកាលកម្មដោយស្វ័យប្រវត្តិនៅផ្ទៃខាងក្រោយក្នុងពេលបន្តិចទៀត)')
      : ('Sheet "' + periodSheetName_(d) + '" មានរួចហើយក្នុងស្រុកទាំងអស់ — បានធ្វើបច្ចុប្បន្នភាពរូបមន្តគណនាឡើងវិញ (ទិន្នន័យបញ្ចូលដោយដៃមិនប៉ះពាល់ទេ)')
  };
}

// ==== ការធ្វើចំណាកស្រុកម្តងគត់ (Run ១ដងគត់ ដោយផ្ទាល់ពី Apps Script Editor — ជ្រើសរើសមុខងារនេះក្នុង Dropdown
// ខាងលើ "Run" រួចចុច Run) — បន្ថែមក្បាលជួរឈរ "សរុប" ថ្មី ទៅ Period Sheet ចាស់ៗ ដែលមានរួចហើយ (បង្កើតមុនលក្ខណៈនេះ)
// ក្នុង Spreadsheet ស្រុកទាំង១០។ ចាំបាច់ត្រូវ Run ព្រោះ createDistrictPeriodSheet_() មិនកសាងក្បាលជួរឈរ (Header)
// ឡើងវិញសម្រាប់ Sheet ដែលមានរួចហើយទេ (សាងសង់ Header តែពេលបង្កើត Sheet ថ្មីប៉ុណ្ណោះ) — រូបមន្តគណនាតម្លៃវិញ ត្រូវបាន
// បំពេញដោយស្វ័យប្រវត្តិដោយមុខងារនេះផ្ទាល់ (applyPeriodSheetFormulasAndFormatting_) ដូច្នេះមិនចាំបាច់រង់ចាំដល់ការ
// រក្សាទុកលើកក្រោយឡើយ។ សុវត្ថិភាព ១០០%៖ ប៉ះពាល់តែក្បាលជួរឈរ + ជួរឈរថ្មី ២ (សរុប/ស្រី) ប៉ុណ្ណោះ មិនប៉ះពាល់ទិន្នន័យ
// ចាស់ណាមួយ ក្នុងជួរឈរផ្សេងទៀត ឬជួរដេកទិន្នន័យណាមួយឡើយ ហើយអាច Run ម្តងទៀតដោយសុវត្ថិភាព (រំលង Sheet ដែលកែរួច) ====
function migratePeriodSheetsAddTotalColumnHeaders_() {
  var fixedCount = 0, skippedCount = 0, errorCount = 0, partialLossCount = 0, unknownCount = 0;
  var partialLossDetails = [];

  DISTRICT_LIST.forEach(function(district) {
    listDistrictPeriodSheets_(district).forEach(function(sheet) {
      try {
        // ==== FIX៖ ឥឡូវប្រើ ensurePeriodSheetHasTotalColumns_() ដែលចាំបាច់ត្រូវប្រើ insertColumnsBefore() ដើម្បីរុញ
        // ជួរឈរ "ចំណាំ"/"អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ចាស់ទៅខាងស្តាំមុន (មិនសរសេរជាន់លើទិន្នន័យទាំងនោះទៀតឡើយ) ====
        var res = ensurePeriodSheetHasTotalColumns_(sheet);
        if (res.action === 'none') { skippedCount++; }
        else if (res.action === 'inserted_safe') { fixedCount++; }
        else if (res.action === 'repaired_partial_loss') {
          fixedCount++; partialLossCount++;
          partialLossDetails.push(district + '/' + sheet.getName());
        } else { unknownCount++; }
      } catch (err) {
        errorCount++;
        try { Logger.log('migratePeriodSheetsAddTotalColumnHeaders_ error (' + district + '/' + sheet.getName() + '): ' + (err && err.message)); } catch (e2) {}
      }
    });
  });

  try { protectPastPeriodSheets_(); } catch (err) {} // ==== ស្តារការចាក់សោវិញសម្រាប់ Tab ថ្ងៃចាស់ៗ តាមក្បួនធម្មតា ====
  scheduleRebuild_(); // ==== ធ្វើបច្ចុប្បន្នភាព Sheet សរុបរួម/ឃុំសង្កាត់/ខេត្ត ក្នុង Spreadsheet មេ ភ្លាមៗ (ជួរឈរថ្មីនឹងបង្ហាញត្រឹមត្រូវ) ====
  var msg = 'បានជួសជុលរួចរាល់! Period Sheet ចំនួន ' + fixedCount + ' បានបន្ថែមជួរឈរ "សរុប" ថ្មី, ' + skippedCount + ' រំលង (ធ្លាប់ត្រឹមត្រូវរួចហើយ)'
      + (unknownCount ? (', ' + unknownCount + ' រំលង (រចនាសម្ព័ន្ធមិនធម្មតា — ត្រូវពិនិត្យដោយដៃ)') : '')
      + (errorCount ? (', ' + errorCount + ' មានបញ្ហា (មើល Execution log)') : '') + '។'
      + (partialLossCount ? (' ⚠️ ក្នុងនោះ ' + partialLossCount + ' Sheet ត្រូវបានជួសជុលរចនាសម្ព័ន្ធតែប៉ុណ្ណោះ ព្រោះទិន្នន័យ "ចំណាំ"/"អ្នកបញ្ចូល" ចាស់ត្រូវបានសរសេរជាន់ពីមុនរួចហើយ (ដោយ Migration កំណែចាស់) — សូមពិនិត្យ Version History នៃ Google Sheet ដើម្បីព្យាយាមសង្គ្រោះ៖ ' + partialLossDetails.join(', ')) : '');
  try { Logger.log(msg); } catch (e3) {}
  return { success: true, fixedCount: fixedCount, skippedCount: skippedCount, errorCount: errorCount, unknownCount: unknownCount, partialLossCount: partialLossCount, partialLossDetails: partialLossDetails, message: msg };
}

function buildDailySheetHeader_(sheet, totalCols) {
  var row1 = [], row2 = [], row3 = [];
  DAILY_META_PREFIX.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });
  // ==== ប្រើលំដាប់ "សម្រាប់បង្ហាញ" (getReportDisplayFields_) មិនមែន DAILY_FIELDS ដើមទេ — ដាក់ជួរឈរ "សរុប"
  // (correctionBiometricTotal) ជាប់នឹង "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ" ជំនួសនៅចុងគេ (Tab នេះជាទិន្នន័យសរុប
  // ស្វ័យប្រវត្តិសុទ្ធសាធ សរសេរឡើងវិញទាំងស្រុងរាល់ពេល ដូច្នេះមិនប៉ះពាល់ Period Sheet ដើមឡើយ) ====
  getReportDisplayFields_().forEach(function(f) { row1.push(f.group || "—"); row2.push(f.sub || "—"); row3.push(f.unit || f.label); });
  DAILY_META_SUFFIX.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });

  sheet.getRange(1, 1, 1, totalCols).setValues([row1]);
  sheet.getRange(2, 1, 1, totalCols).setValues([row2]);
  sheet.getRange(3, 1, 1, totalCols).setValues([row3]);
  sheet.getRange(1, 1, 3, totalCols)
      .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(4);
  mergeConsecutiveHeaderCells_(sheet, 1, totalCols, row1);
  mergeConsecutiveHeaderCells_(sheet, 2, totalCols, row2);
  sheet.setColumnWidths(1, totalCols, 95);
}

// ---- Sheet(ស្ថិតិរួមឃុំសង្កាត់(បូកយោង)) — បូកសរុបស្វ័យប្រវត្តិពី Sheet(ទិន្នន័យចុះឈ្មោះបោះឆ្នាំប្រចាំថ្ងៃ) តាមស្រុក+ឃុំ/សង្កាត់ ----
function buildSumSheetHeader_(sheet, totalCols, prefixHeaders, suffixHeaders, color) {
  var row1 = [], row2 = [], row3 = [];
  prefixHeaders.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });
  // ==== ប្រើលំដាប់ "សម្រាប់បង្ហាញ" (getReportDisplayFields_) — ដូចគ្នានឹង buildDailySheetHeader_() ខាងលើ ====
  getReportDisplayFields_().forEach(function(f) { row1.push(f.group || "—"); row2.push(f.sub || "—"); row3.push(f.unit || f.label); });
  suffixHeaders.forEach(function(h) { row1.push("—"); row2.push("—"); row3.push(h); });

  sheet.getRange(1, 1, 1, totalCols).setValues([row1]);
  sheet.getRange(2, 1, 1, totalCols).setValues([row2]);
  sheet.getRange(3, 1, 1, totalCols).setValues([row3]);
  sheet.getRange(1, 1, 3, totalCols)
      .setFontWeight("bold").setBackground(color).setFontColor("#ffffff")
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(prefixHeaders.length);
  mergeConsecutiveHeaderCells_(sheet, 1, totalCols, row1);
  mergeConsecutiveHeaderCells_(sheet, 2, totalCols, row2);
  sheet.setColumnWidths(1, totalCols, 95);
}

// បញ្ចូលគ្នា (merge) ក្រឡាដែលមានតម្លៃដូចគ្នា ជាប់ៗគ្នា នៅលើជួរដេកមួយ (សម្រាប់ក្បាលតារាង ក្រុមធំ/ក្រុមរង)
function mergeConsecutiveHeaderCells_(sheet, rowIndex, totalCols, values) {
  var startCol = 1;
  for (var c = 2; c <= totalCols + 1; c++) {
    var curVal = c <= totalCols ? values[c - 1] : null;
    var prevVal = values[startCol - 1];
    if (curVal !== prevVal || c > totalCols) {
      if (c - startCol > 1 && prevVal && prevVal !== "—") {
        sheet.getRange(rowIndex, startCol, 1, c - startCol).merge();
      }
      startCol = c;
    }
  }
}

// ==================== ៤.១ ព័ត៌មានទំព័រចូលប្រព័ន្ធ (មិនត្រូវការចូលប្រព័ន្ធជាមុនទេ) ====================
// ត្រឡប់បញ្ជីស្រុក (សម្រាប់ Dropdown ស្នើសុំគណនីថ្មី) + ស្ថិតិសាមញ្ញ (ចំនួនសមាជិកសកម្ម + ចំនួនចូលប្រើ)
// ==================== ៤.០ ការកំណត់ប្រព័ន្ធ (ឈ្មោះស្ថាប័ន/Logo/Telegram/តំណភ្ជាប់/ភាសា) ====================
// រក្សាទុកជា Script Properties (Key-Value សាមញ្ញ) — មិនចាំបាច់ Sheet ដាច់ដោយឡែកសម្រាប់ការកំណត់ចំនួនតិចតួចទាំងនេះទេ
function findDailyRowIndexById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return -1;
  var ids = sheet.getRange(4, 1, lastRow - 3, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return 4 + i;
  }
  return -1;
}

// ត្រឡប់ថា ជួរដេកមួយ ក្នុង Period Sheet មានទិន្នន័យពិតប្រាកដដែរឬទេ (មានតម្លៃ≠០ ណាមួយ ឬចំណាំ ឬអ្នកបញ្ចូល)
// ត្រូវការ ព្រោះគ្រប់ជួរឃុំ/សង្កាត់ទាំងអស់ ត្រូវបានបង្កើតជាមុនស្រាប់ក្នុង Period Sheet (សូម្បីមិនទាន់មានអ្នកបញ្ចូលក៏ដោយ)
// ==== ល្បឿន៖ វិធីចាស់ (findDailyEntryLocationSlow_) ស្កេន Period Sheet គ្រប់ស្រុកទាំង១០ × គ្រប់ថ្ងៃខែ (ការហៅ Sheets
// API ដាច់ដោយឡែករាប់សិបដង) រាល់ពេលកែ/លុបធាតុមួយ — យឺតគួរឲ្យកត់សម្គាល់ ជាពិសេសកាលណាទិន្នន័យកកកុញកាន់តែច្រើនថ្ងៃ។
// ដូច្នេះព្យាយាមរក ស្រុក+កាលបរិច្ឆេទ+ឃុំសង្កាត់ ពី "សន្ទស្សន៍" (loadDailyIndexOnce_ — Sheet តែមួយ ក្នុងមួយ Execution
// ស្រាប់ៗនៅក្នុងសតិ គ្មានការហៅ Sheets API បន្ថែមទេ) ជាមុនសិន រួចបើកតែ Sheet តែមួយ (មិនមែនទាំង១០ស្រុកទេ) ដើម្បីទាញយក
// rowIndex ពិតប្រាកដ (findCommuneRowIndex_ — ១ការហៅ Sheets API ប៉ុណ្ណោះ) ។ ត្រូវផ្ទៀងផ្ទាត់ ID ត្រង់ជួរដេកនោះឡើងវិញ
// ជានិច្ច (Sanity Check) ព្រោះសន្ទស្សន៍អាចហួសសម័យបន្តិចនៅក្នុងករណីកម្រ (ឧ. Trigger ពីក្រោយឆាកមិនទាន់រត់ចប់) — បើមិន
// ត្រូវគ្នាទេ ត្រលប់ទៅវិធីចាស់ (ស្កេនពេញលេញ) វិញ ដើម្បីធានាភាពត្រឹមត្រូវ ១០០% ====
function findDailyEntryLocationFast_(id) {
  var map = loadDailyIndexOnce_();
  for (var key in map) {
    if (!Object.prototype.hasOwnProperty.call(map, key)) continue;
    var entries = map[key] || [];
    for (var i = 0; i < entries.length; i++) {
      if (String(entries[i].id) === String(id)) {
        var sep = key.indexOf('|');
        if (sep < 0) continue;
        return { district: key.slice(0, sep), date: key.slice(sep + 1), commune: entries[i].commune };
      }
    }
  }
  return null;
}
function findDailyEntryLocationSlow_(id) {
  for (var i = 0; i < DISTRICT_LIST.length; i++) {
    var d = DISTRICT_LIST[i];
    var sheets = listDistrictPeriodSheets_(d);
    for (var j = 0; j < sheets.length; j++) {
      var rIdx = findDailyRowIndexById_(sheets[j], id);
      if (rIdx !== -1) return { sheet: sheets[j], rowIndex: rIdx, district: d };
    }
  }
  return null;
}
function findDailyEntryLocation_(id) {
  if (!id) return null;
  try {
    var fast = findDailyEntryLocationFast_(id);
    if (fast && fast.district && fast.date && fast.commune) {
      var dSs = getDistrictSpreadsheet_(fast.district);
      var sheet = dSs.getSheetByName(periodSheetName_(fast.date));
      if (sheet) {
        var rIdx = findCommuneRowIndex_(sheet, fast.commune);
        if (rIdx !== -1) {
          var actualId = sheet.getRange(rIdx, 1).getDisplayValue();
          if (String(actualId) === String(id)) {
            return { sheet: sheet, rowIndex: rIdx, district: fast.district };
          }
        }
      }
    }
  } catch (err) {}
  // ---- វិធីចាស់ (ត្រលប់ប្រើ ប្រសិនបើសន្ទស្សន៍រកមិនឃើញ ឬហួសសម័យ) ----
  return findDailyEntryLocationSlow_(id);
}

// ==== ល្បឿន៖ ទិន្នន័យសម្រាប់ "សាងសង់ទម្រង់" ដាច់ដោយឡែកពីទិន្នន័យ "បញ្ជីធាតុ" ====
// មុននេះ ទម្រង់បញ្ចូល (Dropdown ស្រុក + ឃុំ/សង្កាត់) ត្រូវរង់ចាំ getDailyEntries() ដែលស្កេន Period Sheet គ្រប់ថ្ងៃខែ
// ក្នុងស្រុកទាំង១០ ចប់សិន ទើបបង្ហាញបាន — ទោះបីជាឈ្មោះស្រុក/ឃុំសង្កាត់ អានតែពី Sheet តូចៗ ២ ប៉ុណ្ណោះក៏ដោយ។
// មុខងារនេះ ត្រឡប់តែអ្វីដែលចាំបាច់សម្រាប់សាងសង់ទម្រង់ (លឿនណាស់ — មិនប៉ះ Period Sheet ណាមួយឡើយ)
// ដូច្នេះ Dropdown លេចឡើងភ្លាមៗ ចំណែកបញ្ជីធាតុ (តារាងខាងក្រោម) ផ្ទុកបន្តនៅពីក្រោយដោយឡែក។
// ==== ល្បឿន៖ អាន "សន្ទស្សន៍" ខាងក្នុង (SHEET_DAILY_INDEX) ម្តងគត់ក្នុងមួយ Execution — Sheet តែមួយ ក្នុង Spreadsheet មេ
// ដែលធ្វើបច្ចុប្បន្នភាពដោយស្វ័យប្រវត្តិរាល់ពេលមានការកែប្រែ (មើល rebuildAllDerivedData_) ។ លឿនជាងច្រើន បើប្រៀបធៀបនឹង
// ការបើក Spreadsheet ដាច់ដោយឡែករបស់ស្រុកទាំង១០ ព្រោះអាន Sheet ១ ជំនួសបើក File ១០ (ដែលយឺតឡើងៗតាមចំនួន Tab
// កើនឡើងជារៀងរាល់ថ្ងៃ ក្នុង Spreadsheet ស្រុកនីមួយៗ)។ Key ទម្រង់ "ស្រុក|កាលបរិច្ឆេទ" → Array នៃធាតុគ្រប់ឃុំ/សង្កាត់ ====
var DAILY_INDEX_CACHE_ = null;
function loadDailyIndexOnce_() {
  if (DAILY_INDEX_CACHE_) return DAILY_INDEX_CACHE_;
  var map = {};
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET_DAILY_INDEX);
    var lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow >= 2) {
      var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // ស្រុក, កាលបរិច្ឆេទ, ទិន្នន័យ(JSON)
      data.forEach(function(r) {
        var dist = String(r[0] || '').trim();
        var dt = String(r[1] || '').trim();
        if (!dist || !dt) return;
        try { map[dist + '|' + dt] = JSON.parse(r[2]); } catch (e) {}
      });
    }
  } catch (err) {}
  DAILY_INDEX_CACHE_ = map;
  return map;
}
// ត្រឡប់ Array នៃធាតុ (បើមាននៅក្នុងសន្ទស្សន៍) ឬ null (បើសន្ទស្សន៍មិនទាន់ត្រូវបានធ្វើបច្ចុប្បន្នភាពសម្រាប់ស្រុក+ថ្ងៃនេះ
// — ករណីនេះ អ្នកហៅត្រូវត្រលប់ទៅវិធីចាស់ ដើម្បីធានាភាពត្រឹមត្រូវ)
// ==== FIX (សំណើថ្មី "Fix112 ដំណើរការយឺត")៖ getDailyEntries() ត្រូវការតែ ១ ជួរដេក (ស្រុក+ថ្ងៃជាក់លាក់មួយ) ប៉ុណ្ណោះ
// រាល់ពេលហៅ (Client ជានិច្ចផ្ញើកាលបរិច្ឆេទដែលបានជ្រើសរើសមកជាមួយ — មិនដែលទទេទេក្នុងការប្រើប្រាស់ជាក់ស្តែង) ប៉ុន្តែ
// មុនកំណែនេះ ហៅ loadDailyIndexOnce_() ខាងលើ ដែលអាន+Parse JSON គ្រប់ជួរដេកទាំងអស់ក្នុង Sheet(សន្ទស្សន៍) ជានិច្ច —
// មានន័យថា គ្រប់ស្រុកទាំង១០ គ្រប់ថ្ងៃចាប់តាំងពីដើមរដូវកាល (កាន់តែច្រើនជាបន្តបន្ទាប់រាល់ថ្ងៃកន្លងទៅ)។ ថ្លៃថ្នូរនេះ
// មិនលេចធ្ងន់ធ្ងរ ពេលមានតែ SuperAdmin/Admin/PEC21/សង្កេតការណ៍ហៅ (Cache Key លទ្ធផលចុងក្រោយ ចែករំលែករួម scope
// 'ALL' — មើល getDailyEntries) ប៉ុន្តែក្រោយ "Fix112" អនុញ្ញាតឲ្យគ្រូប្រចាំក្រុងស្រុកទាំង១០ឃើញផងដែរ (Cache Key
// ដាច់ដោយឡែកម្នាក់ៗតាមស្រុក — cacheScope_ = myDistrict) Sheet ពេញលេញនេះត្រូវបានអាន+Parse JSON រាល់ជួរដេក
// ទាំងអស់ដដែលៗ ដល់ម្តងក្នុងមួយនាទីក្នុងមួយស្រុក (រហូតដល់១០ដងច្រើនជាងមុន សរុបគ្រប់ស្រុក) — នេះជាមូលហេតុចម្បងនៃ
// "មានអារម្មណ៍ថាដំណើរការយឺត" ដែលអ្នកប្រើប្រាស់រាយការណ៍។ ដំណោះស្រាយ (ប្រើវិធីដូចគ្នានឹង fastReindexDistrictDay_()
// ខាងលើ ដែលសរសេរ)៖ អាន "តែជួរឈរ ១-២ ប៉ុណ្ណោះ" (ស្រុក/ថ្ងៃ — អក្សរធម្មតា ថោក មិនមែន JSON) ដើម្បីរកលេខជួរដេកគោលដៅ
// សិន (Cache ទុកម្តងគត់ក្នុងមួយ Execution ដូច loadDailyIndexOnce_ ដែរ — មិនអានឡើងវិញរាល់ស្រុកក្នុង Loop
// districtsToRead ទេ) រួច Parse តែក្រឡា JSON តែមួយ (ជួរដេកគោលដៅ ជួរឈរទី៣) ជំនួសការអាន+Parse JSON គ្រប់ជួរដេក
// ទាំងអស់ក្នុង Sheet។ លទ្ធផលចុងក្រោយ (រូបរាង Object ដូចគ្នាបេះបិទ) នៅតែដូចមុនទាំងស្រុង — មានតែល្បឿនប្រសើរឡើងប៉ុណ្ណោះ ====
var DAILY_INDEX_KEYS_CACHE_ = null; // ម្តងគត់ក្នុងមួយ Execution — [[ស្រុក, ថ្ងៃ], ...] ស្របតាមលេខជួរដេកក្នុង Sheet (ចាប់ពីជួរដេកទី២)
function loadDailyIndexKeysOnce_() {
  if (DAILY_INDEX_KEYS_CACHE_) return DAILY_INDEX_KEYS_CACHE_;
  var keys = [];
  try {
    var ss = getSS_();
    var sheet = ss.getSheetByName(SHEET_DAILY_INDEX);
    var lastRow = sheet ? sheet.getLastRow() : 0;
    if (lastRow >= 2) keys = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // ស្រុក, កាលបរិច្ឆេទ ប៉ុណ្ណោះ (ថោក — មិនមែន JSON)
  } catch (err) {}
  DAILY_INDEX_KEYS_CACHE_ = keys;
  return keys;
}
function getDailyIndexEntries_(district, dateStr) {
  try {
    var keys = loadDailyIndexKeysOnce_();
    for (var i = keys.length - 1; i >= 0; i--) { // ពីក្រោយមកមុខ — ដូច fastReindexDistrictDay_() (ជួរដេកចុងក្រោយឈ្នះ បើមានស្ទួន)
      if (String(keys[i][0]) === district && String(keys[i][1]) === dateStr) {
        var ss = getSS_();
        var sheet = ss.getSheetByName(SHEET_DAILY_INDEX);
        if (!sheet) return null;
        var json = sheet.getRange(i + 2, 3, 1, 1).getValue(); // Parse តែក្រឡា JSON តែមួយគត់ (ជួរដេកគោលដៅ)
        try { return JSON.parse(json); } catch (e) { return null; }
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

// ==== ធ្វើបច្ចុប្បន្នភាព "សន្ទស្សន៍" ភ្លាមៗ សម្រាប់ស្រុក+ថ្ងៃជាក់លាក់មួយ (មិនចាំបាច់រង់ចាំ Trigger ១-៤ វិនាទីទេ) ====
// ប្រើភ្លាមៗបន្ទាប់ពីរក្សាទុក/លុប ដើម្បីធានាថាទិន្នន័យថ្មីៗបំផុតឃើញភ្លាមៗ បើអ្នកប្រើចុច 🔄 ភ្លាមៗ (មិនចាំបាច់រង់ចាំ
// rebuildAllDerivedData_ រត់ពីក្រោយឆាកទេ)។ ==== សំខាន់៖ ត្រូវ "ជំនួស" ជួរដេកចាស់ (បើមាន) មិនមែន "បន្ថែម" ជួរដេកថ្មីរាល់
// ដងទេ — បើកែ/រក្សាទុកថ្ងៃដដែលច្រើនដងជាប់ៗគ្នា (ឧ. កែកំហុសម្តងហើយម្តងទៀត) ដោយមិនទាន់ដល់ពេល Trigger សម្អាតរត់
// (rebuildAllDerivedData_ ២-៤ វិនាទីក្រោយ) Sheet សន្ទស្សន៍ នឹងកកកុញជួរដេកស្ទួនកាន់តែច្រើនឡើងៗ ធ្វើឲ្យរាល់ការអានក្រោយមក
// (getDailyEntries) យឺតជាងមុនៗ ជារៀងរាល់ការកែប្រែ — នេះជាមូលហេតុពិតដែលការធ្វើបច្ចុប្បន្នភាពតារាងយឺតជាបន្តបន្ទាប់
// បន្ទាប់ពីកែប្រែច្រើនដង ====
// ==== FIX (ល្បឿន — "រក្សាទុកទិន្នន័យ" កម្រិតបន្ថែម)៖ Cache ពេញលេញនៃ ការផ្គូផ្គង "ស្រុក|កាលបរិច្ឆេទ" → លេខជួរដេក
// ក្នុង SHEET_DAILY_INDEX (ដូចគ្នានឹង Pattern getPeriodSheetDateMap_ ក្នុង Utils.gs) ដើម្បីជៀសវាងការស្កេនគ្រប់ជួរដេក
// ទាំងអស់ (idxSheet.getRange(2,1,lastRow-1,2).getValues()) រាល់ពេលរក្សាទុក/លុប/កែផ្ទាល់ក្នុង Sheet — ការស្កេននេះ
// កាន់តែយឺតបន្តិចម្តងៗ តាមចំនួនការផ្សំ (ស្រុក × ថ្ងៃ) ដែលធ្លាប់មានទិន្នន័យចាប់តាំងពីដើមរដូវកាល។
// ==== សុវត្ថិភាព (Race Condition)៖ គ្រប់ Caller ទាំង៤ (saveDistrictDayEntries, onDistrictSheetEdit_,
// upsertDailyEntry_, deleteDailyEntry) សុទ្ធតែហៅ fastReindexDistrictDay_() ខណៈកាន់កាប់ LockService.getScriptLock()
// (Lock សកលតែមួយ) រួចហើយ ដូច្នេះមិនអាចមានការហៅព្រមគ្នា (concurrent) ២ដងលើ Cache នេះក្នុងពេលតែមួយឡើយ។
// ==== សំខាន់ (Invalidation)៖ rebuildAllDerivedData_() សរសេរជាន់ Sheet(សន្ទស្សន៍) ទាំងស្រុងឡើងវិញ (clearContent
// + setValues តាមលំដាប់ថ្មីទាំងស្រុង ដែលមិនចាំបាច់ដូចលំដាប់ចាស់ទេ) ដូច្នេះ Cache នេះ ត្រូវលុបចោលភ្លាមៗបន្ទាប់ពី
// rebuildAllDerivedData_ រត់ចប់ (មើលចំណុចហៅ clearDailyIndexRowMapCache_() ខាងក្រោមក្នុង Function នោះ) — បើមិន
// ដូច្នេះទេ Cache ចាស់នឹងចង្អុលទៅជួរដេកខុស ធ្វើឲ្យទិន្នន័យស្រុក/ថ្ងៃខុសគ្នា សរសេរជាន់គ្នាដោយចៃដន្យ ====
var CACHE_KEY_DAILY_INDEX_ROWMAP_ = 'dailyIdxRowMap_v1';
var CACHE_TTL_DAILY_INDEX_ROWMAP_ = 21600; // វិនាទី (៦ ម៉ោង) — ដូចគ្នានឹង CACHE_TTL_PERIOD_DATES_

function getDailyIndexRowMap_(idxSheet) {
  var cache = CacheService.getScriptCache();
  try {
    var cached = cache.get(CACHE_KEY_DAILY_INDEX_ROWMAP_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}
  var map = {};
  var lastRow_ = idxSheet.getLastRow();
  if (lastRow_ >= 2) {
    var keys_ = idxSheet.getRange(2, 1, lastRow_ - 1, 2).getValues();
    for (var i_ = 0; i_ < keys_.length; i_++) {
      map[String(keys_[i_][0]) + '|' + String(keys_[i_][1])] = i_ + 2;
    }
  }
  try { cache.put(CACHE_KEY_DAILY_INDEX_ROWMAP_, JSON.stringify(map), CACHE_TTL_DAILY_INDEX_ROWMAP_); } catch (err) {}
  return map;
}

function clearDailyIndexRowMapCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY_DAILY_INDEX_ROWMAP_); } catch (err) {}
}

function fastReindexDistrictDay_(district, dateStr, sheetData) {
  try {
    var n = DAILY_FIELDS.length;
    var fieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1;
    var noteColIdx = fieldsStart0 + n, byColIdx = fieldsStart0 + n + 1, atColIdx = fieldsStart0 + n + 2;
    // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ")៖ ជួរឈរទី៤ (index +3) ថ្មី — បន្ថែមទៅសន្ទស្សន៍
    // JSON ដូចគ្នានឹង note/enteredBy/enteredAt ខាងលើដែរ ====
    var editCountColIdx = fieldsStart0 + n + 3;
    var communeIdx0 = DAILY_META_PREFIX.length - 1;
    var entries = [];
    sheetData.forEach(function(row) {
      var commune = row[communeIdx0];
      if (!commune) return;
      var values = {};
      DAILY_FIELDS.forEach(function(f, i) { values[f.key] = row[fieldsStart0 + i]; });
      entries.push({
        id: row[0], commune: commune, values: values,
        note: row[noteColIdx], enteredBy: row[byColIdx], enteredAt: row[atColIdx],
        editCount: Number(row[editCountColIdx]) || 0,
        hasData: periodRowHasData_(row)
      });
    });
    var idxSheet = ensureDailyIndexSheet_(getSS_());
    var rowValues_ = [district, dateStr, JSON.stringify(entries), formatNow_()];
    // ==== ស្វែងរកជួរដេកចាស់ (បើមាន) សម្រាប់ស្រុក+ថ្ងៃដដែលនេះ ដើម្បីសរសេរជាន់ជំនួស (មិនបន្ថែមថ្មីស្ទួន) — ប្រើ Cache
    // ខាងលើ (O(1)) ជំនួសការស្កេនគ្រប់ជួរដេកទាំងអស់ (O(n)) រាល់ពេលហៅ ====
    var mapKey_ = district + '|' + dateStr;
    var rowMap_ = getDailyIndexRowMap_(idxSheet);
    var targetRow_ = rowMap_[mapKey_] || -1;
    if (targetRow_ > 0) {
      idxSheet.getRange(targetRow_, 1, 1, 4).setValues([rowValues_]);
    } else {
      idxSheet.appendRow(rowValues_);
      var newRow_ = idxSheet.getLastRow();
      rowMap_[mapKey_] = newRow_;
      try { CacheService.getScriptCache().put(CACHE_KEY_DAILY_INDEX_ROWMAP_, JSON.stringify(rowMap_), CACHE_TTL_DAILY_INDEX_ROWMAP_); } catch (errCachePut_) {}
    }
  } catch (err) {
    try { Logger.log("fastReindexDistrictDay_ error: " + err.message); } catch (e2) {}
  }
}

// ==== ទិដ្ឋភាពទូទៅ (Dashboard) — ស្កេន Period Sheet ម្តងគត់ ដើម្បីទាញទាំងចំនួនសរុប + ការបែងចែកតាមខែ (ឆ្នាំបច្ចុប្បន្ន) ====
function getDailyEntries(currentUsername, filterDate, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var viewAll = canViewAllDistricts_(currentUsername);
  var myDistrict = getUserDistrict_(currentUsername);
  var d = filterDate ? normalizeDateStr_(filterDate) : null;

  // ==== ល្បឿន៖ Cache លទ្ធផល ១ នាទី — ជៀសវាងស្កេន Period Sheet គ្រប់ស្រុកទាំង១០ (បើក Spreadsheet ដាច់ដោយឡែក
  // ១០ ដង) ឡើងវិញរាល់ពេលហៅ ជាពិសេសភ្លាមៗបន្ទាប់ពីរក្សាទុកទិន្នន័យមួយស្រុក ដែលទម្រង់ខាងលើហៅមុខងារនេះឡើងវិញភ្លាមៗ
  // ដើម្បីធ្វើបច្ចុប្បន្នភាពតារាងខាងក្រោម (ជាហេតុធ្វើឲ្យអេក្រង់គាំង "កំពុងផ្ទុកតារាង..." យូរ ព្រោះត្រូវស្កេនគ្រប់ស្រុកទាំងអស់
  // ម្តងទៀត ទោះបីជាមានតែស្រុកមួយប៉ុណ្ណោះទើបផ្លាស់ប្តូរក៏ដោយ)។ កំណែ Cache (getReportsCacheVersion_) ត្រូវបានតម្កើងឡើងវិញ
  // ភ្លាមៗ ពេលមានការកែប្រែទិន្នន័យ (មើល scheduleRebuild_) ដូច្នេះលទ្ធផលថ្មីៗនៅតែធានាបានក្នុងរយៈពេលមិនលើសពី ១ នាទី ====
  var cacheScope_ = viewAll ? 'ALL' : (myDistrict || '');
  var cacheKey_ = 'dailyEntries_v' + getReportsCacheVersion_() + '_' + cacheScope_ + '_' + role + '_' + (d || '');
  try {
    var cached_ = CacheService.getScriptCache().get(cacheKey_);
    if (cached_) return JSON.parse(cached_);
  } catch (err) {}

  var totalCols = periodSheetTotalCols_();
  var n = DAILY_FIELDS.length;
  var fieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1; // 0-indexed
  var noteColIdx = fieldsStart0 + n, byColIdx = fieldsStart0 + n + 1, atColIdx = fieldsStart0 + n + 2;
  // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ") ====
  var editCountColIdx = fieldsStart0 + n + 3;
  var districtIdx0 = DAILY_META_PREFIX.length - 2; // ស្រុក ជានិច្ចធាតុទីពីរចុងក្រោយនៃ Prefix
  var communeIdx0 = DAILY_META_PREFIX.length - 1;  // ឃុំ/សង្កាត់ ជានិច្ចធាតុចុងក្រោយនៃ Prefix
  var districtsToRead = viewAll ? DISTRICT_LIST : (isDistrictRole_(role) ? [myDistrict] : []);
  var list = [];

  // ==== ល្បឿន៖ បើមានកាលបរិច្ឆេទត្រូវច្រោះ (filterDate) អានតែ Sheet ដែលត្រូវនឹងថ្ងៃនោះប៉ុណ្ណោះ ====
  // (មិនស្កេនគ្រប់ Period Sheet ទាំងអស់ទេ — លឿនជាងច្រើន ជាពិសេសពេលមានទិន្នន័យជាច្រើនខែ)
  // ចំណាំល្បឿន៖ អានតែម្តងគត់ក្នុងមួយ Sheet (getValues() ឆៅ) ជំនួសការអានពីរដង (getDisplayValues()+getValues())
  // ដូចមុន — កាត់បន្ថយចំនួនការហៅ Google Sheets ពាក់កណ្តាល។ ID/ស្រុក/ឃុំសង្កាត់/ចំណាំ/អ្នកបញ្ចូល ជាអក្សរធម្មតា
  // ដូច្នេះមិនប៉ះពាល់ទេ ចំណែកឯកាលបរិច្ឆេទ ត្រូវការតម្លៃឆៅ (Date object) ស្រាប់ដើម្បីឲ្យ normalizeDateStr_ ត្រឹមត្រូវ។
  districtsToRead.forEach(function(dist) {
    if (d) {
      // ==== ច្រោះតាមកាលបរិច្ឆេទ៖ បង្ហាញគ្រប់ឃុំ/សង្កាត់ទាំងអស់ ទោះមិនទាន់មានទិន្នន័យក៏ដោយ (សម្គាល់ hasData) ====
      // ល្បឿន៖ ព្យាយាមអានពី "សន្ទស្សន៍" ជាមុនសិន (មិនបើក Spreadsheet ស្រុកនេះទាល់តែសោះ) — លឿនណាស់
      var idxEntries = getDailyIndexEntries_(dist, d);
      if (idxEntries) {
        var idxByCommune = {};
        idxEntries.forEach(function(e) { if (e.commune) idxByCommune[e.commune] = e; });
        var dmI = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        (COMMUNE_ORDER[dist] || []).forEach(function(c) {
          var e = idxByCommune[c.name];
          if (e) {
            list.push({
              id: e.id || '', date: d, dateDisplay: dmI ? (dmI[3] + '-' + dmI[2] + '-' + dmI[1]) : d,
              district: dist, commune: c.name,
              values: e.values || {}, note: e.note || '',
              enteredBy: e.enteredBy || '', enteredAt: e.enteredAt || '',
              editCount: e.editCount || 0,
              hasData: !!e.hasData
            });
          } else {
            var emptyValuesI = {};
            DAILY_FIELDS.forEach(function(f) { emptyValuesI[f.key] = 0; });
            list.push({
              id: '', date: d, dateDisplay: dmI ? (dmI[3] + '-' + dmI[2] + '-' + dmI[1]) : d,
              district: dist, commune: c.name,
              values: emptyValuesI, note: '', enteredBy: '', enteredAt: '',
              editCount: 0,
              hasData: false
            });
          }
        });
        return;
      }
      // ---- វិធីចាស់ (ត្រលប់ប្រើ ប្រសិនបើសន្ទស្សន៍មិនទាន់មានទិន្នន័យសម្រាប់ស្រុក+ថ្ងៃនេះ — ថ្មីទើបរក្សាទុក ឬកំពុងរង់ចាំ
      // Rebuild ត្រឡប់ក្រោយ ១-២ វិនាទី) ៖ បើក Spreadsheet ស្រុកនេះផ្ទាល់ (ស្រុកតែមួយប៉ុណ្ណោះ មិនមែនទាំង១០ទេ) ----
      var dSs = getDistrictSpreadsheet_(dist);
      var oneSheet = dSs.getSheetByName(periodSheetName_(d));
      var rowsByCommune = {};
      if (oneSheet) {
        // ==== FIX (សំខាន់ណាស់ — សុវត្ថិភាពទិន្នន័យ)៖ ផ្លូវនេះមិនឆ្លងកាត់ createDistrictPeriodSheet_ ទេ (អាន Sheet
        // ចាស់ដោយផ្ទាល់) — ត្រូវហៅ Migration ដោយផ្ទាល់ត្រង់នេះ មុននឹងអាន totalCols ទទឹងពេញ (បើមិនដូច្នេះទេ Sheet
        // ចាស់ៗ ដែលមិនទាន់ Migrate នឹងធ្វើឲ្យ getRange ខាងក្រោមបរាជ័យ — Range outside sheet bounds) ====
        try { ensurePeriodSheetHasEditCountColumn_(oneSheet); } catch (errEnsure5) {}
        var lastRow1 = oneSheet.getLastRow();
        if (lastRow1 >= 4) {
          var data1 = oneSheet.getRange(4, 1, lastRow1 - 3, totalCols).getValues();
          data1.forEach(function(row) {
            var c = row[communeIdx0];
            if (c) rowsByCommune[c] = row;
          });
        }
      }
      (COMMUNE_ORDER[dist] || []).forEach(function(c) {
        var row = rowsByCommune[c.name];
        if (row) {
          var values = {};
          DAILY_FIELDS.forEach(function(f, i) { values[f.key] = row[fieldsStart0 + i]; });
          var isoDate = normalizeDateStr_(row[1]);
          var dm = isoDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
          list.push({
            id: row[0], date: isoDate, dateDisplay: dm ? (dm[3] + '-' + dm[2] + '-' + dm[1]) : isoDate,
            district: dist, commune: c.name,
            values: values, note: row[noteColIdx],
            enteredBy: row[byColIdx], enteredAt: row[atColIdx],
            editCount: Number(row[editCountColIdx]) || 0,
            hasData: periodRowHasData_(row)
          });
        } else {
          // គ្មាន Sheet ថ្ងៃនេះទាល់តែសោះ ឬគ្មានជួរដេកឃុំ/សង្កាត់នេះ — បង្ហាញជាតម្លៃទទេ សម្គាល់ថាមិនទាន់មានទិន្នន័យ
          var emptyValues = {};
          DAILY_FIELDS.forEach(function(f) { emptyValues[f.key] = 0; });
          var dm0 = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
          list.push({
            id: '', date: d, dateDisplay: dm0 ? (dm0[3] + '-' + dm0[2] + '-' + dm0[1]) : d,
            district: dist, commune: c.name,
            values: emptyValues, note: '', enteredBy: '', enteredAt: '',
            editCount: 0,
            hasData: false
          });
        }
      });
      return;
    }

    // ==== គ្មានកាលបរិច្ឆេទច្រោះ (មើលប្រវត្តិសាស្ត្រទាំងអស់) — បង្ហាញតែជួរដែលមានទិន្នន័យប៉ុណ្ណោះ (ជៀសវាងបញ្ជីវែងពេក) ====
    // ល្បឿន៖ អានពី "សន្ទស្សន៍" ជាមុនសិន (មិនស្កេន Tab ថ្ងៃខែទាំងអស់ក្នុង Spreadsheet ស្រុកនេះទាល់តែសោះ)
    var idxMap = loadDailyIndexOnce_();
    var prefix = dist + '|';
    var hasIdxForDist = false;
    Object.keys(idxMap).forEach(function(key) {
      if (key.indexOf(prefix) !== 0) return;
      hasIdxForDist = true;
      var dateForKey = key.slice(prefix.length);
      (idxMap[key] || []).forEach(function(e) {
        if (!e.hasData) return;
        var dm = dateForKey.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        list.push({
          id: e.id || '', date: dateForKey, dateDisplay: dm ? (dm[3] + '-' + dm[2] + '-' + dm[1]) : dateForKey,
          district: dist, commune: e.commune,
          values: e.values || {}, note: e.note || '',
          enteredBy: e.enteredBy || '', enteredAt: e.enteredAt || '',
          editCount: e.editCount || 0,
          hasData: true
        });
      });
    });
    if (hasIdxForDist) return;

    // ---- វិធីចាស់ (ត្រលប់ប្រើ ប្រសិនបើសន្ទស្សន៍មិនទាន់មានទិន្នន័យសម្រាប់ស្រុកនេះទាល់តែសោះ) ----
    listDistrictPeriodSheets_(dist).forEach(function(sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow < 4) return;
      // ==== FIX (សំខាន់ណាស់ — សុវត្ថិភាពទិន្នន័យ)៖ ស្កេន Period Sheet ចាស់ៗ (រួមទាំង Sheet ខែចាស់ដែលបិទរួច មិនទាន់
      // ធ្លាប់ត្រូវបាន Migrate តាំងពីលក្ខណៈនេះមាន) ដោយផ្ទាល់ ដូច្នេះត្រូវហៅ Migration នៅទីនេះជាដាច់ខាត ====
      try { ensurePeriodSheetHasEditCountColumn_(sheet); } catch (errEnsure6) {}
      var data = sheet.getRange(4, 1, lastRow - 3, totalCols).getValues();
      data.forEach(function(row) {
        if (!row[0] || !periodRowHasData_(row)) return; // រំលងជួរឃុំ/សង្កាត់ដែលមិនទាន់មានអ្នកបញ្ចូល
        var values = {};
        DAILY_FIELDS.forEach(function(f, i) { values[f.key] = row[fieldsStart0 + i]; });
        var isoDate = normalizeDateStr_(row[1]);
        var dm = isoDate.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        list.push({
          id: row[0], date: isoDate, dateDisplay: dm ? (dm[3] + '-' + dm[2] + '-' + dm[1]) : isoDate,
          district: row[districtIdx0], commune: row[communeIdx0],
          values: values, note: row[noteColIdx],
          enteredBy: row[byColIdx], enteredAt: row[atColIdx],
          editCount: Number(row[editCountColIdx]) || 0,
          hasData: true
        });
      });
    });
  });

  // ==== ជួរឈរ "សរុប" ថ្មី (ជួរឈរនិម្មិត — មិនមែនក្នុង DAILY_FIELDS ទេ, សូមមើលមូលហេតុលម្អិតនៅ getReportDisplayFields_()
  // ក្នុង Utils.gs) — គណនាបន្ថែមចូល entry.values នីមួយៗ ដោយមិនគិតថា entry នោះមកពីផ្លូវអាន ៥ ផ្លូវផ្សេងគ្នាខាងលើ ====
  list.forEach(function(item) {
    var iv = item.values || (item.values = {});
    iv.correctionBiometricTotal_total = (Number(iv.correction_total) || 0) + (Number(iv.biometric_total) || 0);
    iv.correctionBiometricTotal_female = (Number(iv.correction_female) || 0) + (Number(iv.biometric_female) || 0);
    // ==== FIX (សំណើថ្មី "សុំប្តូរឈ្មោះគណនី(jav) មកជាដាក់ឈ្មោះ(គោត្តនាម-នាម) វិញ")៖ saveDistrictDayEntries()/
    // upsertDailyEntry_() ខាងលើ (ឥឡូវប្រើ formatEnteredBy_() រួមគ្នា — មើល Auth.gs) កត់ត្រា "Username - ឈ្មោះពេញ"
    // ចាប់ពីពេលរក្សាទុកលើកក្រោយតទៅ ប៉ុន្តែជួរដេកចាស់ៗ ណាដែលធ្លាប់រក្សាទុករួចហើយ (មុនកែ) អាចនៅតែមាន Username ឆៅ
    // (ឬឈ្មោះឆៅ) ជាប់ក្នុង Sheet ដដែល។ ដើម្បីឲ្យទិន្នន័យចាស់ៗទាំងនោះ ក៏បង្ហាញឈ្មោះត្រឹមត្រូវភ្លាមៗដែរ (មិនចាំបាច់រង់ចាំ
    // កែ/រក្សាទុកជាថ្មីម្តងទៀតជួរដេកនីមួយៗទេ) ត្រូវព្យាយាមបកប្រែ enteredBy នៅពេលអាន (Display-time Resolution)៖
    // បើ enteredBy ជា Username ត្រឹមត្រូវ (ចាស់ណាស់) → ប្តូរទៅ "Username - ឈ្មោះពេញ" ស្វ័យប្រវត្តិ (formatEnteredBy_)។
    // បើ enteredBy ជាទម្រង់ផ្សេងរួចហើយ (ឈ្មោះម្នាក់ឯង ឬ "Username - ឈ្មោះពេញ" រួចហើយ — ទម្រង់ថ្មីបំផុត) ឬគណនីត្រូវបាន
    // លុបរួច → getUserRow_() ខាងក្នុង getUserFullName_()/formatEnteredBy_() រកមិនឃើញគណនីត្រូវនឹងខ្សែអក្សរទាំងមូលនោះ
    // ទេ ដូច្នេះនៅតែបង្ហាញតម្លៃដើមដដែល (ត្រឹមត្រូវស្រាប់ ក្នុងករណីនេះ) ====
    if (item.enteredBy) item.enteredBy = getUserFullName_(item.enteredBy) ? formatEnteredBy_(item.enteredBy) : item.enteredBy;
  });

  var ref = getRollingReferenceMap_(filterDate || todayDateStr_());
  // ==== ការកំណត់បង្ហាញ/លាក់ជួរឈរបន្ថែម ដាច់ដោយឡែកសម្រាប់ Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) នេះ (Group "Daily") —
  // ដាច់ដោយឡែកពីទម្រង់បញ្ចូលទិន្នន័យ (Group A) និង Tab(របាយការណ៍សរុប) (Group C) — មើល resolveGroupDisplaySettings_() ====
  var dispDaily_ = resolveGroupDisplaySettings_(getSystemSettings_(), 'Daily');
  var result_ = {
    success: true, entries: list, districts: DISTRICT_LIST, fields: getReportDisplayFields_(), canEditAll: isAdmin_(currentUsername),
    communesByDistrict: getCommuneMap_(), baseline2025: getBaseline2025Map_(),
    showDuplicateFieldsDaily: dispDaily_.showDuplicateFields, showEstimateDaily: dispDaily_.showEstimate,
    showProvinceEstimateDaily: dispDaily_.showProvinceEstimate, showNewStationDaily: dispDaily_.showNewStation,
    showTotalStationDaily: dispDaily_.showTotalStation, showCommuneCodeDaily: dispDaily_.showCommuneCode,
    districtOrder: DISTRICT_LIST, districtCode: DISTRICT_CODE, districtType: DISTRICT_TYPE,
    referenceMap: ref.map, referenceIsFirstDay: ref.isFirstDay, referenceLabel: ref.label,
    operationStartDate: getOperationStartDate_(), todayStr: formatNow_().slice(0, 10)
  };
  // ==== ការពារ៖ ធានាថា Object ដែលត្រឡប់ទៅ Client គឺជា JSON "ស្អាត" ១០០% (Object/Array/String/Number/Boolean/null
  // ធម្មតាតែប៉ុណ្ណោះ) ព្រោះ google.script.run របស់ Google Apps Script ត្រូវការទម្រង់នេះជានិច្ច។ បើ Object ណាមួយ
  // ចេញក្រៅពីនេះ (ឧ. Date object ឆៅ ជំនួសជា String, undefined, NaN, ឬអ្វីផ្សេងទៀត) — Google អាចនឹងផ្ញើតម្លៃ null
  // ត្រឡប់ទៅ Client ដោយស្ងាត់ស្ងៀម (មិនបង្ហាញកំហុសអ្វីទាំងអស់ សូម្បីតែក្នុង Executions Log ក៏ដោយ ព្រោះ Execution ខ្លួនឯង
  // "Completed" ដោយជោគជ័យ តែជំហានបំប្លែងទិន្នន័យទៅ Client ទៀតទេដែលបរាជ័យ) ដែលធ្វើឲ្យអេក្រង់គាំង "កំពុងផ្ទុកតារាង..."
  // ជារៀងរហូត។ ការឆ្លងកាត់ JSON.parse(JSON.stringify(...)) ត្រង់នេះ បង្ខំឲ្យបញ្ហាបែបនេះលេចចេញជាកំហុសច្បាស់លាស់
  // (Failed ក្នុង Executions Log) ជំនួសឲ្យលទ្ធផល null ស្ងាត់ស្ងៀម ====
  var safeResult_;
  try {
    safeResult_ = JSON.parse(JSON.stringify(result_));
  } catch (sanitizeErr) {
    return { success: false, message: "កំហុសទិន្នន័យខាងក្នុង (មិនអាចបម្លែងជា JSON បាន)៖ " + sanitizeErr.message };
  }
  try { CacheService.getScriptCache().put(cacheKey_, JSON.stringify(safeResult_), REPORTS_CACHE_TTL_); } catch (err) {}
  return safeResult_;
}

// ==== ស្វែងរកធាតុតែមួយ ផ្ទាល់ពី Google Sheet (ស្រស់ជានិច្ច + លឿនបំផុត) ====
// មុននេះ ទម្រង់បញ្ចូលពឹងផ្អែកលើបញ្ជីធាតុទាំងអស់ ដែលបាន Cache ទុកនៅផ្នែក Client ចាប់តាំងពីពេលបើកទំព័រ។
// ដូច្នេះបើគ្រូប្រចាំស្រុកបញ្ចូលទិន្នន័យដោយផ្ទាល់ក្នុង Google Sheet ក្រោយពេលបើកទំព័ររួច — ទម្រង់នៅតែបង្ហាញ
// ទិន្នន័យចាស់ រហូតដល់បើកទំព័រឡើងវិញ។ មុខងារនេះអានផ្ទាល់ពី Period Sheet ត្រឹមតែ "១ Sheet + ១ ជួរដេក"
// ដូច្នេះទិន្នន័យថ្មីៗបំផុតបង្ហាញភ្លាមៗជានិច្ច ហើយលឿនជាងការស្កេនបញ្ជីទាំងមូលច្រើនដង។
// សំខាន់៖ មិនបង្កើត Period Sheet ថ្មីទេ (បើថ្ងៃនោះមិនទាន់មាន Sheet → ត្រឡប់ null) ព្រោះការគ្រាន់តែ
// "មើល" មិនគួរបង្កើត Tab ថ្មីដោយស្វ័យប្រវត្តិឡើយ។
// ==== ទម្រង់តារាងគ្រប់ឃុំ/សង្កាត់ក្នុងស្រុកតែមួយ (ជំនួសទម្រង់ជ្រើសរើសឃុំ/សង្កាត់តែមួយៗ) ====
// អានតម្លៃទាំងអស់ ក្នុង Period Sheet របស់ ស្រុក+កាលបរិច្ឆេទ ដែលបានផ្តល់មក — ត្រឡប់ជា Map គិតតាមឈ្មោះឃុំ/សង្កាត់
// មិនបង្កើត Period Sheet ថ្មីទេ (បើមិនទាន់មាន Sheet ថ្ងៃនោះ → ត្រឡប់ Map ទទេ ដើម្បីឲ្យ Client បង្ហាញ ០ គ្រប់ជួរ)
function getDistrictDayEntries(currentUsername, date, district, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  if (!canViewAllDistricts_(currentUsername) && getUserDistrict_(currentUsername) !== district) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិមើលទិន្នន័យស្រុកនេះទេ!" };
  }
  try {
    var d = normalizeDateStr_(date);
    var dSs = getDistrictSpreadsheet_(district);
    var sheet = dSs.getSheetByName(periodSheetName_(d));
    var n = DAILY_FIELDS.length;
    var communeIdx0 = DAILY_META_PREFIX.length - 1; // ឃុំ/សង្កាត់ ជានិច្ចជាធាតុចុងក្រោយនៃ Prefix
    var fieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1; // បម្លែងទៅជា 0-indexed
    var communes = {};
    if (sheet) {
      // ==== FIX (សំខាន់ណាស់ — សុវត្ថិភាពទិន្នន័យ)៖ Function នេះជាផ្លូវអាន "ស្រស់ជាងគេ" សម្រាប់ផ្ទុកទម្រង់បញ្ចូល
      // ទិន្នន័យប្រចាំថ្ងៃ (ចំណុចចូលមុនគេបំផុត ដែល Client ហៅមុននឹងមាន Save ណាមួយកើតឡើងផង) — មិនឆ្លងកាត់
      // createDistrictPeriodSheet_ ទេ ដូច្នេះត្រូវហៅ Migration ដោយផ្ទាល់ត្រង់នេះជាដាច់ខាត មុននឹងអាន totalCols ទទឹង
      // ពេញ (បើមិនដូច្នេះទេ Sheet ចាស់ៗ ដែលមិនទាន់ Migrate តាំងពីលក្ខណៈ "ចំនួនដងកែប្រែ" នេះមាន នឹងធ្វើឲ្យទម្រង់
      // បញ្ចូលទិន្នន័យ បើកមិនរួចទាល់តែសោះ — Range outside sheet bounds) ====
      try { ensurePeriodSheetHasEditCountColumn_(sheet); } catch (errEnsure7) {}
      var lastRow = sheet.getLastRow();
      if (lastRow >= 4) {
        var totalCols = periodSheetTotalCols_();
        var data = sheet.getRange(4, 1, lastRow - 3, totalCols).getValues();
        data.forEach(function(r) {
          // ==== FIX (រាយការណ៍ថ្មី "ចំនួនឃុំ/សង្កាត់មិនត្រឹមត្រូវ — ស្រុកអង្គរបូរីមាន៦ឃុំ តែបង្ហាញ៧")៖ ជួរដេក "សរុប"
          // ស្វ័យប្រវត្តិ (writePeriodSheetTotalRow_) មាន ID ទទេ ប៉ុន្តែជួរឈរ "ឃុំ/សង្កាត់" របស់វា ("សរុបទាំងអស់") មិន
          // ទទេទេ — ការត្រួតពិនិត្យ "if (!c) return;" ខាងក្រោម ត្រួតពិនិត្យតែឈ្មោះឃុំទទេ មិនចាប់ជួរដេកនេះទេ ធ្វើឲ្យវា
          // ត្រូវបានរាប់ជា "ឃុំ/សង្កាត់" បន្ថែម ១ ដោយអចេតនា ក្នុងចំនួន Object.keys(res.communes).length ដែល Client
          // ប្រើសម្រាប់សារ "ថ្ងៃនេះមានទិន្នន័យរួចហើយសម្រាប់ N ឃុំ/សង្កាត់" (Index.html) — FIX៖ រំលងជួរដេកនេះជាមុនសិន
          // (ដូចគ្នានឹងគោលការណ៍ blank-ID guard ដែលប្រើនៅគ្រប់ទីកន្លែងផ្សេងទៀតរួចហើយ សម្រាប់ជួរដេកនេះ) ====
          if (!r[0]) return; // ជួរដេក "សរុប" (ID ទទេ) — មិនមែនឃុំ/សង្កាត់ពិត
          var c = String(r[communeIdx0] || "").trim();
          if (!c) return;
          var values = {};
          DAILY_FIELDS.forEach(function(f, i) { values[f.key] = r[fieldsStart0 + i]; });
          // ==== FIX (សំណើថ្មី "រក្សាទុកការកែប្រែទិន្នន័យ")៖ ជួរដេកឃុំ/សង្កាត់ទាំងអស់មានស្រាប់ជានិច្ចតាំងពី Sheet
          // ត្រូវបានបង្កើត (តម្លៃ ០ លំនាំដើម) ដូច្នេះ `communes[c]` ខាងលើពិតជានិច្ចមានតម្លៃ ទោះបីជាឃុំ/សង្កាត់នោះមិនទាន់
          // មានអ្នកបញ្ចូលទិន្នន័យពិតប្រាកដក៏ដោយ — ត្រូវប្រើ periodRowHasData_(r) (មានស្រាប់ — ប្រើដដែលក្នុង
          // fastReindexDistrictDay_) ដើម្បីដឹងថាឃុំ/សង្កាត់នេះ "ធ្លាប់មានទិន្នន័យបញ្ចូលរួច" ពិតប្រាកដដែរឬអត់ — ប្រើដោយ
          // Client ដើម្បីសម្រេចថាតើត្រូវបង្ហាញប៊ូតុង "រក្សាទុកទិន្នន័យ" (ថ្មី) ឬ "រក្សាទុកការកែប្រែទិន្នន័យ" (កែពីមុន) ====
          communes[c] = { commune: c, values: values, note: r[fieldsStart0 + n] || "", hasData: periodRowHasData_(r) };
        });
      }
    }
    // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គណនីកម្រិតឃុំសង្កាត់ មិនត្រូវឃើញទិន្នន័យឃុំសង្កាត់ដទៃទៀតក្នុងស្រុកខ្លួនទេ —
    // ត្រង `communes` (ដែលមាន Key ជាឈ្មោះឃុំសង្កាត់ទាំងអស់ក្នុងស្រុក) សល់ត្រឹមតែឃុំសង្កាត់ខ្លួនផ្ទាល់ (Server-side —
    // Defense in depth, មិនទុកចិត្តតែ Client Filter ទេ) ====
    if (isCommuneRole_(role)) {
      var ownCommune_ = getUserCommune_(currentUsername);
      var filteredCommunes_ = {};
      if (communes[ownCommune_]) filteredCommunes_[ownCommune_] = communes[ownCommune_];
      communes = filteredCommunes_;
    }
    var ref = getRollingReferenceMap_(d, district);
    // ==== FIX (សំណើថ្មី "ថ្ងៃប្រជុំ")៖ ជូនដំណឹង Client ថាថ្ងៃនេះត្រូវបានសម្គាល់ថា "គ្មានទិន្នន័យ" ដើម្បីបង្ហាញសារ
    // ច្បាស់លាស់ (មិនមែនសារទូទៅ "ថ្ងៃនេះមិនទាន់មាន Sheet") ====
    var skipInfo_ = isDateSkipped_(d);
    return {
      success: true, communes: communes, sheetExists: !!sheet, pastLocked: isPastDateLocked_(currentUsername, district, d),
      // ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព")៖ ខុសពី pastLocked ខាងលើ (ស្វ័យប្រវត្តិ — ថ្ងៃចាស់
      // ប៉ុណ្ណោះ) — printLocked គឺជាការចាក់សោដោយដៃ (Admin ចុចប៊ូតុងផ្ទាល់) អាចកើតឡើងសម្រាប់ថ្ងៃណាក៏បាន (រួមទាំងថ្ងៃបច្ចុប្បន្ន)
      // មើលការពន្យល់ពេញលេញនៅ isPrintLockedForDistrict_() ====
      printLocked: isPrintLockedForDistrict_(currentUsername, district, d),
      referenceMap: ref.map, referenceIsFirstDay: ref.isFirstDay, referenceLabel: ref.label,
      isSkippedDate: !!skipInfo_, skipNote: skipInfo_ ? skipInfo_.note : '',
      // ==== FIX (សំណើថ្មី "ធ្វើបច្ចុប្បន្នភាពទិន្នន័យភ្លាមៗឆ្លងគណនី")៖ ផ្ញើ "កំណែ" បច្ចុប្បន្នមកជាមួយការទាញយកនេះតែម្តង
      // (មិនចាំបាច់ហៅ getDistrictDayVersion() ដាច់ដោយឡែកបន្ថែមទៀតទេ) — Client ប្រើតម្លៃនេះជាចំណុចចាប់ផ្តើមសម្រាប់
      // ការប្រៀបធៀប Poll លើកបន្ទាប់ៗ — មើលការពន្យល់ពេញលេញនៅ districtDayVersionCacheKey_() ខាងលើ ====
      version: readDistrictDayVersion_(district, d)
    };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== FIX (សំណើថ្មី "ធ្វើបច្ចុប្បន្នភាពទិន្នន័យភ្លាមៗឆ្លងគណនី — មិនបាច់ប្តូរថ្ងៃ/Refresh")៖ RPC ស្រាលៗសម្រាប់ Client
// Poll រៀងរាល់ ១០វិនាទី (មើល pollDistrictDayUpdates_ ក្នុង Index.html) ដើម្បីដឹងថាស្រុក+ថ្ងៃកំពុងមើល មានការផ្លាស់ប្តូរ
// ថ្មីឬអត់ ដោយមិនចាំបាច់ទាញយកទិន្នន័យពេញលេញ (Sheets API) រាល់ដងទេ — អាន CacheService តែមួយប៉ុណ្ណោះ (ថោកបំផុត)។
// ត្រួតពិនិត្យសិទ្ធិដូចគ្នាបេះបិទនឹង getDistrictDayEntries() (Defense in depth — មិនឲ្យព័ត៌មាន "មានអ្វីប្តូរ" សូម្បីតែ
// Timestamp ធម្មតា លេចធ្លាយទៅគណនីគ្មានសិទ្ធិមើលស្រុកនោះ) ====
function getDistrictDayVersion(currentUsername, date, district, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  if (!canViewAllDistricts_(currentUsername) && getUserDistrict_(currentUsername) !== district) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិមើលទិន្នន័យស្រុកនេះទេ!" };
  }
  try {
    var d = normalizeDateStr_(date);
    return { success: true, version: readDistrictDayVersion_(district, d) };
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// ==== FIX (សំណើថ្មី "កែលម្អ Lock — អនុញ្ញាតឲ្យស្រុកខុសគ្នារក្សាទុកដំណាលគ្នាបានពិតប្រាកដ")៖ ពីមុន saveDistrictDayEntries/
// upsertDailyEntry_/deleteDailyEntry ទាំង៣ ប្រើ LockService.getScriptLock() តែមួយ (Lock រួមទាំង Script — មិនញែក
// ដាច់ដោយឡែកតាមស្រុកទេ) ព័ទ្ធជុំវិញ "ការអាន Sheet ស្រុកទាំងមូល → កែក្នុងសតិ → សរសេរជាន់ត្រឡប់វិញ" (ដើម្បីការពារ
// Race Condition ដូចមានពន្យល់នៅចំណុចនីមួយៗ)។ ដោយសារស្រុកនីមួយៗពិតជាមាន Google Spreadsheet ដាច់ដោយឡែកគ្នា
// ១០០% ស្រាប់ (គ្មានហេតុផលពិតប្រាកដត្រូវឲ្យស្រុក A រង់ចាំស្រុក B ដោះសោជាមុនសិនទេ) — Lock រួមនេះធ្វើឲ្យរាល់ការ
// "រក្សាទុក"/"លុប" ទូទាំងគណនីទាំង ២០០ (រាប់បញ្ចូលទាំងស្រុកផ្សេងគ្នា) ត្រូវតម្រៀបជួរតែមួយជាប់គ្នាដោយអសារឥតការ
// ទោះជាមិនប៉ះទង្គិចគ្នាផ្ទាល់ក៏ដោយ។ FIX៖ បំបែក Lock ជា ២ ជាន់ដាច់ដោយឡែក៖
//   (១) acquireDistrictDayLock_/releaseDistrictDayLock_ ខាងក្រោម — Lock "តាម Logic" ដាច់ដោយឡែកគ្នាតាម ស្រុក+
//       កាលបរិច្ឆេទ (ប្រើ CacheService ជា Flag ដែលមានកំណត់ថ្ងៃផុតកំណត់ស្វ័យប្រវត្តិ ២៥វិនាទី ជា "សំណាញ់សុវត្ថិភាព"
//       ក្នុងករណី Execution គាំង/Timeout ខណៈកំពុងកាន់កាប់ — ការពារ Deadlock អចិន្ត្រៃយ៍) ព័ទ្ធជុំវិញតែផ្នែក "អាន Sheet
//       ស្រុកខ្លួនផ្ទាល់ → កែ → សរសេរជាន់ត្រឡប់វិញ" (ដែលជាផ្នែកយឺតបំផុត និងជាកម្មសិទ្ធិ ១០០% របស់ស្រុកនោះម្នាក់ឯង —
//       File ខុសគ្នាទាំងស្រុងទៅតាមស្រុក) — ស្រុកខុសគ្នា ប្រើ Key ខុសគ្នា ដូច្នេះអាចដំណើរការស្របគ្នាបានពិតប្រាកដ។ ការ
//       ត្រួតពិនិត្យ+កំណត់ Flag ខ្លួនឯង ត្រូវការពារដោយ LockService.getScriptLock() ខ្លីៗតែមួយភ្លែត (មិនមែនរយៈពេលពេញ
//       នៃការសរសេរ Sheet ទាំងមូលទេ) ដើម្បីធានាថា ២ សំណើមិនអាចឃើញ "មិនទាន់រវល់" ក្នុងពេលដំណាលគ្នាបានទាល់តែសោះ
//       (Atomic Check-and-Set)។
//   (២) fastReindexDistrictDay_() (ធ្វើបច្ចុប្បន្នភាព SHEET_DAILY_INDEX — Sheet សន្ទស្សន៍ដែលជា "កម្មសិទ្ធិរួម" សម្រាប់
//       គ្រប់ស្រុកទាំងអស់ រួមទាំង Cache Row-Map ដែលជា JSON តែមួយ គ្របដណ្ដប់គ្រប់ស្រុក) នៅតែត្រូវការពារដោយ
//       LockService.getScriptLock() ដដែល (មិនប្តូរទៅ Lock ដាច់ដោយឡែកតាមស្រុកទេ) ព្រោះប្រតិបត្តិការនេះប៉ះពាល់
//       ធនធានរួម (Shared Row-Map Cache + ការបន្ថែមជួរដេកថ្មីនៅចុង Sheet សន្ទស្សន៍) ដែលស្រុកផ្សេងគ្នាពិតជាអាចប៉ះទង្គិច
//       គ្នាបានពិតប្រាកដ (ឧ. ២ស្រុកខុសគ្នាទាំង២កំពុងបន្ថែមធាតុសន្ទស្សន៍ថ្មីក្នុងពេលជិតគ្នា អាចធ្វើឲ្យ Cache Row-Map
//       សរសេរជាន់គ្នា បាត់ធាតុថ្មីមួយ)។ ដំណើរការនេះលឿនណាស់ (អាន/សរសេរជួរដេកតែមួយ) ដូច្នេះការរង់ចាំគ្នាខ្លីៗនេះ
//       ស្ទើរតែមិនអាចមើលឃើញដោយអ្នកប្រើប្រាស់ឡើយ — មិនប៉ះពាល់ផលចំណេញនៃការបំបែក Lock ខាងលើឡើយ ====
var DISTRICT_DAY_LOGIC_LOCK_TTL_SEC_ = 25; // ថ្ងៃផុតកំណត់ស្វ័យប្រវត្តិរបស់ Flag (សំណាញ់សុវត្ថិភាព — ការពារ Deadlock អចិន្ត្រៃយ៍)
// ==== FIX (សំណើថ្មី "ដកចេញការរង់ចាំរវាងឃុំសង្កាត់ក្នុងស្រុក+ថ្ងៃដដែល")៖ បន្ថែម scopeSuffix ជា Parameter ជម្រើស
// (Optional) — បើមិនបានផ្តល់មក ឬទទេ Key នៅតែដដែលទាំងស្រុង (ស្រុក+កាលបរិច្ឆេទ តែប៉ុណ្ណោះ) ដូចមុន។ ត្រូវបានប្រើដោយ
// saveDistrictDayEntries សម្រាប់គណនីឃុំសង្កាត់ (isCallerCommune_) ប៉ុណ្ណោះ — ដាក់ឈ្មោះឃុំសង្កាត់ខ្លួនឯងបន្ថែមក្នុង Key
// ដើម្បីឲ្យឃុំសង្កាត់ផ្សេងគ្នា (ក្នុងស្រុក+ថ្ងៃដដែល) ទទួលបាន Key ខុសគ្នា — មិនចាំបាច់រង់ចាំគ្នាឡើយ ព្រោះជួរដេកដែលអាច
// ត្រូវបានប៉ះពាល់ (touchedIdxs_) តាមរចនាសម្ព័ន្ធកូដ តែងតែមានតែមួយ (ជួរដេកឃុំសង្កាត់ខ្លួនឯង) សម្រាប់អ្នកហៅជាគណនី
// ឃុំសង្កាត់ — ២ឃុំសង្កាត់ខុសគ្នា មិនដែលសរសេរជាន់លើជួរដេកគ្នាទៅវិញទៅមកទាល់តែសោះ។ គណនីស្រុក/Admin/PEC21 នៅតែប្រើ Key
// ដដែល (គ្មាន scopeSuffix — ស្រុក+កាលបរិច្ឆេទតែប៉ុណ្ណោះ) ព្រោះមួយ Batch របស់ពួកគេអាចប៉ះពាល់ច្រើនជួរដេកឃុំសង្កាត់
// ខុសៗគ្នាមិនប្រាកដប្រជា (មិនអាចធានាថាមានតែជួរដេកតែមួយបានទេ) — មើលការពន្យល់ពេញលេញនៅ saveDistrictDayEntries ====
function districtDayLockCacheKey_(district, date, scopeSuffix) {
  return 'ddlogiclock_' + district + '|' + date + (scopeSuffix ? '|' + scopeSuffix : '');
}

// ==== FIX (សំណើថ្មី "ធ្វើបច្ចុប្បន្នភាពទិន្នន័យភ្លាមៗឆ្លងគណនី — មិនបាច់ប្តូរថ្ងៃ/Refresh")៖ ពេលគណនីឃុំសង្កាត់ណាមួយ
// បញ្ចូល/កែប្រែទិន្នន័យ ត្រូវឲ្យអ្នកមើលផ្សេងទៀត (គ្រូប្រចាំក្រុងស្រុក/SuperAdmin/Admin/PEC21) ដែលកំពុងមើលស្រុក+ថ្ងៃ
// ដដែលឃើញភ្លាមៗដែរ ដោយមិនបាច់កែថ្ងៃ ឬ Refresh ទំព័រ។ ដោយសារ Google Apps Script HtmlService មិនគាំទ្រ WebSocket/
// Server Push ទេ (មានតែ Client→Server Request-Response ធម្មតា) ដំណោះស្រាយសមស្របតែមួយគត់គឺ Client Polling — ប៉ុន្តែ
// ការហៅ getDistrictDayEntries ពេញលេញ (អាន Sheet ទាំងមូល) រាល់ ១០វិនាទីពី Browser Tab ជាច្រើន (រាប់រយគណនី) នឹងធ្ងន់
// ពេកលើ Sheets API Quota ណាស់។ FIX៖ បំបែកជា ២ ជាន់៖ (១) "កំណែ" (Version) ស្រាលៗ — CacheService Key តែមួយក្នុងមួយ
// ស្រុក+ថ្ងៃ ផ្ទុកតែ Timestamp (Millis) ចុងក្រោយបំផុតដែលមានការសរសេរជោគជ័យ — Client Poll រកមើល Key នេះញឹកញាប់
// (រាល់ ១០វិនាទី) ព្រោះថោកបំផុត (CacheService.get() តែមួយ — គ្មាន Sheets API Call ទាល់តែសោះ) (២) មានតែពេល Version
// ប្តូរពិតប្រាកដ (មានការរក្សាទុកថ្មីកើតឡើងចាប់តាំងពី Client ដឹងចុងក្រោយ) ទើប Client ទាញយកទិន្នន័យពេញលេញម្តង (ដូចគ្នា
// នឹង getDistrictDayEntries ធម្មតា) — មើល getDistrictDayVersion()/bumpDistrictDayVersion_() ខាងក្រោម ព្រមទាំង
// applyDistrictDayLiveUpdate_()/pollDistrictDayUpdates_() ក្នុង Index.html សម្រាប់ភាគ Client។ Cache TTL 6ម៉ោង
// (អតិបរមារបស់ CacheService) — គ្រប់គ្រាន់ណាស់សម្រាប់រយៈពេលបើកមើលទំព័រធម្មតាមួយថ្ងៃ ====
function districtDayVersionCacheKey_(district, date) {
  return 'ddver_' + district + '|' + date;
}
// ហៅរាល់ពេលមានការសរសេរជោគជ័យ (ពី saveDistrictDayEntries) សម្រាប់ស្រុក+ថ្ងៃដែលបានផ្តល់មក — "បោះឲ្យដឹង" ថាមានការ
// ផ្លាស់ប្តូរថ្មី ដោយកំណត់ Cache ទៅជា Timestamp បច្ចុប្បន្ន (Millis, ជា String)។ បរាជ័យស្ងាត់ៗ (try/catch) ដូចទីកន្លែង
// ប្រើ CacheService ដទៃទៀតក្នុង File នេះ — មិនត្រូវឲ្យបញ្ហា Cache ធ្វើឲ្យការរក្សាទុកទិន្នន័យពិតបរាជ័យទាំងមូលឡើយ
// (Best-effort — បើ Cache បរាជ័យ Client គ្រាន់តែឃើញការធ្វើបច្ចុប្បន្នភាពយឺតជាងធម្មតាបន្តិច មិនមែនបាត់បង់ទិន្នន័យទេ)
function bumpDistrictDayVersion_(district, date) {
  try {
    CacheService.getScriptCache().put(districtDayVersionCacheKey_(district, date), String(Date.now()), 21600);
  } catch (eBumpVer_) {}
}
// អាន "កំណែ" បច្ចុប្បន្នសម្រាប់ស្រុក+ថ្ងៃដែលបានផ្តល់មក (String ទទេ '' បើមិនទាន់ធ្លាប់មានការសរសេរណាមួយ ចាប់តាំងពី Cache
// ចុងក្រោយផុតកំណត់/មិនទាន់ចាប់ផ្តើម — ស្មើនឹង "មិនស្គាល់" — Client នឹងចាត់ទុកថាមិនមានអ្វីប្តូរនៅឡើយ ទាល់តែឃើញការ
// សរសេរជោគជ័យពិតប្រាកដលើកដំបូងសិន)
function readDistrictDayVersion_(district, date) {
  try {
    return CacheService.getScriptCache().get(districtDayVersionCacheKey_(district, date)) || '';
  } catch (eReadVer_) { return ''; }
}
// ព្យាយាមទទួល "សោតាម Logic" សម្រាប់ស្រុក+កាលបរិច្ឆេទដែលបានផ្តល់មក — ត្រឡប់ Key ដែលបានចាក់សោ (String) បើជោគជ័យ
// ឬ null បើអស់ពេលរង់ចាំ (maxWaitMs) ដោយមិនអាចទទួលបាន — ត្រូវហៅ releaseDistrictDayLock_() ជានិច្ចនៅចុងបញ្ចប់
// (ក្នុង finally) បើទទួលបានជោគជ័យ។ scopeSuffix ជា Parameter ជម្រើស (Optional) — មើលការពន្យល់នៅ
// districtDayLockCacheKey_() ខាងលើ
function acquireDistrictDayLock_(district, date, maxWaitMs, scopeSuffix) {
  var cache = CacheService.getScriptCache();
  var cacheKey = districtDayLockCacheKey_(district, date, scopeSuffix);
  var deadline = Date.now() + maxWaitMs;
  while (true) {
    var scriptLock = LockService.getScriptLock();
    var gotShortLock = false;
    try {
      try { gotShortLock = scriptLock.tryLock(500); } catch (eShort) {}
      if (gotShortLock) {
        var busy = null;
        try { busy = cache.get(cacheKey); } catch (eGet) {}
        if (!busy) {
          try { cache.put(cacheKey, '1', DISTRICT_DAY_LOGIC_LOCK_TTL_SEC_); } catch (ePut) {}
          return cacheKey; // ជោគជ័យ — ទទួលបាន "សោ Logic" សម្រាប់ស្រុក+កាលបរិច្ឆេទនេះ
        }
      }
    } finally {
      if (gotShortLock) { try { scriptLock.releaseLock(); } catch (eRel) {} }
    }
    if (Date.now() >= deadline) return null; // អស់ពេលរង់ចាំ
    Utilities.sleep(200);
  }
}
function releaseDistrictDayLock_(cacheKey) {
  if (!cacheKey) return;
  try { CacheService.getScriptCache().remove(cacheKey); } catch (e) {}
}

// រក្សាទុកទិន្នន័យសម្រាប់គ្រប់ឃុំ/សង្កាត់ក្នុងស្រុកមួយ ក្នុងកាលបរិច្ឆេទតែមួយ ក្នុងការហៅតែម្តងគត់ (Batch — មិនមែនម្តងមួយឃុំសង្កាត់ទេ)
function saveDistrictDayEntries(currentUsername, date, district, rows, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ canEditDistrict_ ត្រឡប់ false ជានិច្ចសម្រាប់ ROLE_COMMUNE (ព្រោះមិនអាចកែប្រែ
  // ទាំងស្រុកបានទេ) — ត្រូវអនុញ្ញាតឲ្យគណនីឃុំសង្កាត់ហៅមុខងារនេះបានដែរ (ជាផ្លូវសំខាន់ដែល Client ប្រើសម្រាប់រក្សាទុក)
  // ប៉ុន្តែនឹងត្រូវបានត្រងឲ្យសល់តែជួរដេកឃុំសង្កាត់ខ្លួនផ្ទាល់ប៉ុណ្ណោះ ខាងក្រោម (មិនទុកចិត្តលើ Client ថាផ្ញើមកតែជួរដេក
  // ខ្លួនប៉ុណ្ណោះទេ — Server-side Filter ជា Defense in depth) ====
  var isCallerCommune_ = isCommuneRole_(getUserRole_(currentUsername));
  if (!canEditDistrict_(currentUsername, district) && !(isCallerCommune_ && getUserDistrict_(currentUsername) === district)) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិបញ្ចូលទិន្នន័យសម្រាប់ស្រុកនេះទេ!" };
  }
  try {
    var d = normalizeDateStr_(date);
    // ==== ការពារ៖ គ្រូប្រចាំស្រុក (District role) មិនអាចបង្កើត Period Sheet ដោយខ្លួនឯងបានទេ ====
    // ត្រូវឲ្យ Admin/SuperAdmin/PEC21 ចុច "បង្កើត Sheet ថ្មីបន្ទាប់" ជាមុនសិន ទើបគ្រូអាចបញ្ចូលទិន្នន័យបាន
    // (Admin/SuperAdmin/PEC21 ខ្លួនឯង នៅតែអាចបង្កើត Sheet ដោយស្វ័យប្រវត្តិតាមរយៈការរក្សាទុកផ្ទាល់បាន)
    // ==== FIX (ល្បឿន x4 — "Save All" នៅតែយឺត)៖ ពីមុន សម្រាប់អ្នកប្រើមិនមែន Admin (ករណីភាគច្រើនបំផុត — គ្រូប្រចាំ
    // ក្រុងស្រុក/ឃុំសង្កាត់ធម្មតា) Block ខាងក្រោមហៅ dSsCheck.getSheetByName() ដើម្បីត្រួតពិនិត្យថា Sheet មានស្រាប់ឬអត់
    // (existingSheet) រួច "ចោល" លទ្ធផលនោះចោល ហើយហៅ getOrCreatePeriodSheet_() ខាងក្រោម ដែលខាងក្នុងហៅ
    // getSheetByName() ម្តងទៀត (Sheets API ២ដងសរុប សម្រាប់ Sheet ដូចគ្នា ក្នុងសំណើតែមួយ)។ ជាក់ស្តែង សម្រាប់អ្នកប្រើ
    // មិនមែន Admin ផ្លូវ "បង្កើត Sheet ថ្មី" ក្នុង getOrCreatePeriodSheet_ មិនអាចកើតឡើងបានទេ (ព្រោះបើ existingSheet
    // គ្មាន យើងបាន return ខាងលើរួចហើយ) — ដូច្នេះ Sheet ដែលទទួលបានពី getOrCreatePeriodSheet_() តែងតែជា Sheet
    // Object ដូចគ្នានឹង existingSheet ជានិច្ច។ FIX៖ ប្រើ existingSheet ដោយផ្ទាល់ ជំនួសការហៅ getOrCreatePeriodSheet_()
    // ម្តងទៀត សម្រាប់ផ្លូវនេះ — កាត់បន្ថយ Sheets API Call មួយទៀត (getSheetByName) រាល់ពេល "រក្សាទុកទាំងអស់" ====
    var sheet;
    if (!isAdmin_(currentUsername)) {
      var dSsCheck = getDistrictSpreadsheet_(district);
      var existingSheet = dSsCheck.getSheetByName(periodSheetName_(d));
      if (!existingSheet) {
        return { success: false, message: "ថ្ងៃនេះមិនទាន់មាន Sheet ទេ! សូមទាក់ទង Admin ដើម្បីបង្កើត Sheet ថ្ងៃនេះជាមុនសិន (ចុច \"បង្កើត Sheet ថ្មីបន្ទាប់\")។" };
      }
      if (isPastDateLocked_(currentUsername, district, d)) {
        return { success: false, message: "ថ្ងៃនេះបានកន្លងផុតទៅហើយ! សូមទាក់ទង SuperAdmin/Admin ដើម្បីបើកសិទ្ធិកែប្រែថ្ងៃនេះជាមុនសិន។" };
      }
      // ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព") — មើលការពន្យល់ពេញលេញនៅ isPrintLockedForDistrict_() ====
      if (isPrintLockedForDistrict_(currentUsername, district, d)) {
        return { success: false, message: "ថ្ងៃនេះកំពុងត្រូវបានចាក់សោបណ្តោះអាសន្ន ដើម្បីការពារទិន្នន័យអំឡុងពេលបោះពុម្ពរបាយការណ៍! សូមទាក់ទង SuperAdmin/Admin ។" };
      }
      sheet = existingSheet;
    } else {
      // ==== FIX (Fix122)៖ អានចំណាំពេញលេញនៅ createDistrictPeriodSheet_() ខាងលើ — Function នេះនឹងសរសេរទិន្នន័យ
      // ថ្មីទាំងមូល និងហៅ applyPeriodSheetFormulasAndFormatting_() ដោយខ្លួនឯងផ្ទាល់ភ្លាមៗខាងក្រោម ដូច្នេះរំលងវដ្ត
      // រូបមន្ត+ការពារជួរឈរដែលដដែលៗ (ស្ទួន) ខាងក្នុង getOrCreatePeriodSheet_ ដើម្បីកាត់បន្ថយពេលវេលា "រក្សាទុកទាំងអស់"
      // (Admin អាចត្រូវការបង្កើត Sheet ថ្មីស្វ័យប្រវត្តិ ដូច្នេះនៅតែហៅ getOrCreatePeriodSheet_() ធម្មតា) ====
      sheet = getOrCreatePeriodSheet_(district, d, true);
    }
    // ==== FIX (សំខាន់ណាស់ — សុវត្ថិភាពទិន្នន័យ)៖ ផ្លូវអ្នកប្រើមិនមែន Admin ខាងលើ (ភាគច្រើនបំផុត — គ្រូប្រចាំក្រុងស្រុក/
    // ឃុំសង្កាត់ធម្មតា) ប្រើ `existingSheet` ដោយផ្ទាល់ រំលងការហៅ createDistrictPeriodSheet_() ទាំងស្រុង (មើលការពន្យល់
    // ពេញលេញនៅខាងលើ "FIX (ល្បឿន x4)") ដែលមានន័យថា ensurePeriodSheetHasEditCountColumn_() (ជាធម្មតាហៅពី
    // createDistrictPeriodSheet_) មិនដែលមានឱកាសរត់សម្រាប់ផ្លូវនេះទាល់តែសោះ — Sheet ចាស់ៗ (បង្កើតមុនលក្ខណៈ "ចំនួន
    // ដងកែប្រែ" នេះ) នឹងតូចជាង totalCols ថ្មីជានិច្ច ធ្វើឲ្យ sheet.getRange(..., totalCols) ខាងក្រោម បរាជ័យ (Range
    // outside sheet bounds) សម្រាប់រាល់ការរក្សាទុកដំបូងគេលើ Sheet ចាស់ណាមួយ។ ត្រូវហៅដោយផ្ទាល់ត្រង់នេះជាដាច់ខាត
    // (មិនអាស្រ័យលើ createDistrictPeriodSheet_ តែម្នាក់ឯងទេ) — Idempotent + Anchor-checked ដូច្នេះថ្លៃថ្នូរតូចតាច
    // (Read ១ក្រឡា) សម្រាប់ Sheet ដែល Migrate រួចស្រាប់ (ករណីភាគច្រើនក្រោយពេលដំបូង) ====
    try { ensurePeriodSheetHasEditCountColumn_(sheet); } catch (errEnsure3) { try { Logger.log('ensurePeriodSheetHasEditCountColumn_ error (saveDistrictDayEntries, ' + district + '/' + periodSheetName_(d) + '): ' + (errEnsure3 && errEnsure3.message)); } catch (eLog3) {} }
    var totalCols = periodSheetTotalCols_();

    // ==== ការពារការប៉ះទង្គិចគ្នា (Race Condition)៖ ការអាន (getRange().getValues()) → កែក្នុងសតិ → សរសេរជាន់ត្រឡប់
    // វិញ (setValues()) ខាងក្រោម មិនមែនជាប្រតិបត្តិការតែមួយឥតដាច់ (Atomic) ទេ — ប្រសិនបើ ២ សំណើកើតឡើងស្ទើរតែក្នុងពេល
    // តែមួយ សម្រាប់ស្រុក+ថ្ងៃដដែល (ឧ. ២ឧបករណ៍ ឬគ្រូ+ការកែដោយផ្ទាល់ក្នុង Sheet) ការសរសេរលើកចុងក្រោយអាចជាន់សរសេរ
    // ជាន់លើការផ្លាស់ប្តូររបស់លើកមុន ធ្វើឲ្យទិន្នន័យបាត់ដោយស្ងាត់ស្ងៀម។
    // ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ ប្តូរពី LockService.getScriptLock() រួម (ព័ទ្ធជុំវិញគ្រប់ស្រុកទាំងអស់ក្នុង
    // Script តែមួយ) ទៅជា acquireDistrictDayLock_() ដែលញែក Lock ដាច់ដោយឡែកតាមស្រុក+កាលបរិច្ឆេទ — មើលការពន្យល់
    // ពេញលេញនៅ acquireDistrictDayLock_()/releaseDistrictDayLock_() ខាងលើ។ ស្រុកខុសគ្នា ឥឡូវអាចរក្សាទុកដំណាលគ្នា
    // បានពិតប្រាកដ (មិនចាំបាច់រង់ចាំគ្នាទៀតទេ) — មានតែស្រុក+ថ្ងៃដដែលប៉ុណ្ណោះ ដែលនៅតែត្រូវតម្រៀបជួរ (រហូតដល់ ១៥វិនាទី
    // ដដែល) ====
    // ==== FIX (សំណើថ្មី "ដកចេញការរង់ចាំរវាងឃុំសង្កាត់ក្នុងស្រុក+ថ្ងៃដដែល")៖ បើអ្នកហៅជាគណនីឃុំសង្កាត់ (isCallerCommune_)
    // ពង្រីក Key បន្ថែមឈ្មោះឃុំសង្កាត់ខ្លួនឯង (callerCommuneForLock_) — ធានាថា ២ឃុំសង្កាត់ខុសគ្នា ក្នុងស្រុក+ថ្ងៃដដែល
    // ទទួលបាន Key ខុសគ្នា ដូច្នេះមិនចាំបាច់រង់ចាំគ្នាឡើយ (មិនរង់ចាំសោះ ក្នុងករណីធម្មតា)។ គណនីស្រុក/Admin/PEC21 មិនមែន
    // ឃុំសង្កាត់ (isCallerCommune_ === false) ចាកចេញ callerCommuneForLock_ ជាខ្សែអក្សរទទេ — Key នៅតែជាស្រុក+
    // កាលបរិច្ឆេទតែប៉ុណ្ណោះ ដូចមុន (មិនផ្លាស់ប្តូរអាកប្បកិរិយាទាល់តែសោះ) ព្រោះមួយ Batch របស់ពួកគេអាចប៉ះពាល់ច្រើនជួរដេក
    // ឃុំសង្កាត់ខុសៗគ្នាមិនប្រាកដប្រជា (មិនអាចធានាថាមានតែជួរដេកតែមួយបានទេ ដូចករណីគណនីឃុំសង្កាត់) ====
    var callerCommuneForLock_ = isCallerCommune_ ? getUserCommune_(currentUsername) : '';
    var lockKey = acquireDistrictDayLock_(district, d, 15000, callerCommuneForLock_);
    if (!lockKey) {
      return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានការរក្សាទុកផ្សេងទៀតកំពុងដំណើរការសម្រាប់ស្រុក+ថ្ងៃដដែល)! សូមព្យាយាមរក្សាទុកម្តងទៀត។" };
    }
    try {
      var lastRow = sheet.getLastRow();
      if (lastRow < 4) return { success: true, message: "រក្សាទុកទិន្នន័យជោគជ័យ!" };

      var data = sheet.getRange(4, 1, lastRow - 3, totalCols).getValues();
      var communeIdx0 = DAILY_META_PREFIX.length - 1;
      var fieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1;
      var rowIndexMap = {};
      // ==== FIX (សុវត្ថិភាពទិន្នន័យ — Defense in depth, ជាប់ពាក់ព័ន្ធនឹងរាយការណ៍ "ចំនួនឃុំ/សង្កាត់មិនត្រឹមត្រូវ")៖
      // ជួរដេក "សរុប" ស្វ័យប្រវត្តិ (ID ទទេ) មិនត្រូវអាចត្រូវបានរកឃើញ/សរសេរជាន់តាមរយៈ Map នេះបានទេ — ទោះ Client
      // ធម្មតាមិនដែលផ្ញើ commune:"សរុបទាំងអស់" មកទេ (renderDistrictDayTableBody ស្ថាបនាតែពី COMMUNE_ORDER) ក៏ដោយ
      // ត្រូវទប់ស្កាត់នៅ Server ជាទីបញ្ចប់ (Defense in depth) កុំឲ្យវាអាចសរសេរជាន់លើរូបមន្ត SUM ជាអចិន្ត្រៃយ៍ដោយចៃដន្យ ====
      data.forEach(function(r, idx) { if (r[0]) rowIndexMap[String(r[communeIdx0]).trim()] = idx; });

      var n = DAILY_FIELDS.length;
      var now = formatNow_();
      // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ត្រង Rows ដែលផ្ញើមក សល់ត្រឹមតែជួរដេកឃុំសង្កាត់ខ្លួនផ្ទាល់ប៉ុណ្ណោះ បើអ្នកហៅ
      // ជាគណនីឃុំសង្កាត់ (សូម្បីតែ Client ផ្ញើមកច្រើនជួរដេក ក៏មិនអាចសរសេរជាន់ឃុំសង្កាត់ដទៃទៀតបានដែរ) ====
      var effectiveRows_ = rows || [];
      if (isCallerCommune_) {
        // ==== FIX (សំណើថ្មី "ដកចេញការរង់ចាំរវាងឃុំសង្កាត់")៖ ប្រើ callerCommuneForLock_ ដដែល (គណនារួចហើយខាងលើ
        // សម្រាប់ Lock Scope) ជំនួសការហៅ getUserCommune_() ម្តងទៀត — តម្លៃដូចគ្នាបេះបិទ (គ្មានហានិភ័យអសង្គតិ) ====
        var ownCommune2_ = callerCommuneForLock_;
        effectiveRows_ = effectiveRows_.filter(function(item) { return String(item.commune || "").trim() === ownCommune2_; });
      }
      var touchedIdxs_ = [];
      effectiveRows_.forEach(function(item) {
        var commune = String(item.commune || "").trim();
        if (!commune) return;
        var idx = rowIndexMap[commune];
        if (idx === undefined) return; // ឃុំ/សង្កាត់មិនស្គាល់ (មិនស្ថិតក្នុង COMMUNE_ORDER) — រំលង
        var values = item.values || {};
        DAILY_FIELDS.forEach(function(f, i) { data[idx][fieldsStart0 + i] = Number(values[f.key]) || 0; });
        data[idx][fieldsStart0 + n] = sanitizeForSheetCell_(item.note || "");
        // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ចំណុចទី៥)៖ ជួរឈរ "អ្នកបញ្ចូល" ឥឡូវកត់ត្រាទាំង ឈ្មោះគណនី (Username)
        // + ឈ្មោះគោត្តនាម-នាម ព្រមគ្នា (formatEnteredBy_() ក្នុង Auth.gs — ចំណុចតែមួយប្រើរួមជាមួយ upsertDailyEntry_()
        // ខាងក្រោម ដើម្បីធានាទម្រង់ដូចគ្នាបេះបិទ) ។ ចំណាំ៖ នេះជាការថតកំណត់ត្រា ត្រឹមតែពេលរក្សាទុកប៉ុណ្ណោះ (Point-in-time
        // — ដូចគ្នានឹង "ពេលបញ្ចូល" ក្រោមនេះ) ដូច្នេះជួរដេកចាស់ៗ ដែលធ្លាប់រក្សាទុករួចហើយ នឹងបន្តបង្ហាញតម្លៃចាស់ រហូតដល់
        // មានការកែប្រែ/រក្សាទុកជាថ្មីម្តងទៀត ====
        data[idx][fieldsStart0 + n + 1] = formatEnteredBy_(currentUsername);
        data[idx][fieldsStart0 + n + 2] = now;
        // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ")៖ អាន "តម្លៃចាស់" (មុននឹងសរសេរជាន់ខាងលើ —
        // data[idx] នៅតែជា Snapshot ដើមដែលអានពី Sheet ដំបូង មុននឹង Loop នេះកែប្រែ) បូក ១ រាល់ពេលមានការរក្សាទុកជាក់
        // ស្តែងលើឃុំ/សង្កាត់នេះ (touchedIdxs_ — ឃុំ/សង្កាត់ដែលពិតជាមានការផ្លាស់ប្តូរតម្លៃណាមួយ មិនមែនគ្រប់ជួរដេកទាំងអស់
        // ក្នុងតារាងទេ — មើលការពន្យល់ពេញលេញនៅ touchedIdxs_ ខាងលើ) ។ ចាត់ទុកតម្លៃទទេ/មិនមែនលេខ ជា ០ (ជួរដេកថ្មីដែល
        // មិនទាន់ធ្លាប់មានការកែប្រែពីមុន ឬ Sheet ចាស់ដែលទើបត្រូវបាន Migrate បំពេញ ០ រួចស្រាប់) ====
        var oldEditCount_ = Number(data[idx][fieldsStart0 + n + 3]) || 0;
        data[idx][fieldsStart0 + n + 3] = oldEditCount_ + 1;
        touchedIdxs_.push(idx);
      });

      // ==== FIX (ល្បឿន x7 — "រក្សាទុកទិន្នន័យ" នៅតែយឺត)៖ ពីមុន setValues()/applyPeriodSheetFormulasAndFormatting_()
      // ខាងក្រោម សរសេរជាន់ត្រឡប់លើជួរដេកទាំងអស់ក្នុងតារាង (data.length — គ្រប់ឃុំ/សង្កាត់ក្នុងស្រុក) ជានិច្ច ទោះបីជា
      // ភាគច្រើនដងមានតែឃុំ/សង្កាត់ ១-២ ប៉ុណ្ណោះដែលពិតជាបានកែប្រែ (Client ខ្លួនឯងផ្ញើមកតែជួរដេកដែលបានកែ — dataset.edited
      // ស្រាប់ហើយ ដូចមានពន្យល់ខាងលើ) — ការសរសេរជាន់លើជួរដេកទាំងអស់ធ្វើឲ្យទំហំទិន្នន័យ (Payload) នៃការហៅ Sheets API
      // ធំជាងតម្រូវការជាច្រើនដង (ឧ. ស្រុកមាន ២០ឃុំ/សង្កាត់ តែកែតែ១ — ២០ដងច្រើនជាងតម្រូវការ)។ FIX៖ ដកព្រំដែន (Range)
      // ការសរសេរឲ្យតូចចុះ — សរសេរតែចន្លោះជួរដេក [ជួរដេកតូចបំផុត..ជួរដេកធំបំផុត] ក្នុងចំណោមឃុំ/សង្កាត់ដែលពិតជាបានកែ
      // ប៉ុណ្ណោះ (touchedIdxs_) មិនមែនគ្រប់ជួរដេកទាំងអស់ក្នុងតារាងទេ (`data` នៅតែអានពេញលេញដដែល ព្រោះត្រូវការសម្រាប់
      // fastReindexDistrictDay_ ខាងក្រោម ដែលត្រូវការស្ថានភាពគ្រប់ឃុំ/សង្កាត់ទាំងអស់ សម្រាប់សន្ទស្សន៍ថ្ងៃទាំងមូល — មិនមែន
      // ត្រឹមតែឃុំ/សង្កាត់ដែលទើបកែទេ) — បើគ្មានឃុំ/សង្កាត់ណាមួយត្រូវបានកែពិតប្រាកដទេ (ករណីកម្រ — ឧ. Server-side Filter
      // គណនីឃុំសង្កាត់ខាងលើ ច្រានចោលអស់) រំលងជំហានសរសេរ+រូបមន្តទាំងស្រុង ====
      if (touchedIdxs_.length) {
        var minIdx_ = Math.min.apply(null, touchedIdxs_);
        var maxIdx_ = Math.max.apply(null, touchedIdxs_);
        var writeFirstRow_ = 4 + minIdx_;
        var writeLastRow_ = 4 + maxIdx_;
        var writeData_ = data.slice(minIdx_, maxIdx_ + 1);
        sheet.getRange(writeFirstRow_, 1, writeData_.length, totalCols).setValues(writeData_);
        // ==== FIX (ល្បឿន — "រក្សាទុកទិន្នន័យ")៖ skipProtection=true — Function នេះសរសេរជាន់ត្រឡប់លើជួរដេកដែលមានស្រាប់
        // ស្រាប់ជានិច្ច (មិនដែលបន្ថែមជួរដេកថ្មីទេ) ដូច្នេះមិនចាំបាច់លុប+ការពារជួរឈររូបមន្តឡើងវិញ (~៣០ការហៅ Sheets API)
        // រាល់ពេលរក្សាទុកនោះទេ — មើលការពន្យល់លម្អិតក្នុង applyPeriodSheetFormulasAndFormatting_() ====
        applyPeriodSheetFormulasAndFormatting_(sheet, totalCols, writeFirstRow_, writeLastRow_, true);
        // ==== FIX (សំណើ "ពិនិត្យរូបមន្តបច្ចុប្បន្នភាពបញ្ជីឆ្នាំ២០២៦")៖ ទិន្នន័យចូល/ចេញថ្ងៃនេះទើបផ្លាស់ប្តូរ — ត្រូវ
        // ជួសជុលជួរឈរ "ទិន្នន័យយោង" របស់ Sheet ថ្ងៃបន្តបន្ទាប់ (បើមានស្រាប់រួចហើយ) ដែលអាចនៅតែយោងលើតម្លៃចាស់មុននឹងកែនេះ
        // (មើលការពន្យល់ពេញលេញនៅ cascadeRefreshRollingReferenceForward_ ក្នុង Function ខាងលើ) — ករណីធម្មតា (គ្មាន
        // Sheet អនាគត) ចាកចេញភ្លាមៗ មិនបន្ថែមថ្លៃថ្កាសអ្វីទាល់តែសោះ ====
        cascadeRefreshRollingReferenceForward_(district, d);
        // ==== FIX (សំណើថ្មី "ធ្វើបច្ចុប្បន្នភាពទិន្នន័យភ្លាមៗឆ្លងគណនី")៖ "បោះឲ្យដឹង" ថាស្រុក+ថ្ងៃនេះមានការផ្លាស់ប្តូរ
        // ថ្មី ដើម្បីឲ្យអ្នកមើលផ្សេងទៀត (Poll រៀងរាល់ ១០វិនាទីតាម getDistrictDayVersion — មើលការពន្យល់ពេញលេញនៅ
        // districtDayVersionCacheKey_ ខាងលើ) ដឹង ហើយទាញយកទិន្នន័យថ្មីដោយស្វ័យប្រវត្តិ ====
        bumpDistrictDayVersion_(district, d);
      }
      scheduleRebuild_();
      // ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ fastReindexDistrictDay_() ប៉ះពាល់ SHEET_DAILY_INDEX ដែលជា Sheet សន្ទស្សន៍
      // រួម (Shared Row-Map Cache តែមួយ គ្របដណ្ដប់ស្រុកទាំងអស់) — នៅតែត្រូវការពារដោយ LockService.getScriptLock()
      // សុទ្ធ (មិនប្តូរទៅ Lock ដាច់ដោយឡែកតាមស្រុកទេ) ព្រោះស្រុកខុសគ្នាអាចប៉ះទង្គិចគ្នាបានពិតប្រាកដត្រង់ចំណុចនេះ
      // (ការសរសេរជាន់គ្នាលើ Cache Row-Map រួម) — ប៉ុន្តែដំណើរការនេះលឿនណាស់ (អាន/សរសេរជួរដេកតែមួយ) ដូច្នេះការរង់ចាំ
      // គ្នាខ្លីៗនេះ ស្ទើរតែមិនអាចមើលឃើញឡើយ — មើលការពន្យល់ពេញលេញនៅ acquireDistrictDayLock_() ខាងលើ ====
      var idxLock_ = LockService.getScriptLock();
      var gotIdxLock_ = false;
      try { gotIdxLock_ = idxLock_.tryLock(10000); } catch (eIdx_) {}
      if (gotIdxLock_) {
        try {
          fastReindexDistrictDay_(district, d, data); // ==== ធានាថាតារាងខាងក្រោមឃើញទិន្នន័យថ្មីភ្លាមៗ បើចុច 🔄 ភ្លាមៗបន្ទាប់ពីរក្សាទុក ====
        } finally {
          try { idxLock_.releaseLock(); } catch (eRelIdx_) {}
        }
      } // ==== បើមិនអាចទទួល Lock បាន (កម្រណាស់) — scheduleRebuild_() ខាងលើ នៅតែនឹងកែសន្ទស្សន៍ជាមួយពីក្រោយឆាកនៅឡើយ ====
      return { success: true, message: "រក្សាទុកទិន្នន័យជោគជ័យ!" };
    } finally {
      releaseDistrictDayLock_(lockKey);
    }
  } catch (err) {
    return { success: false, message: "កំហុស៖ " + err.message };
  }
}

// សរសេរតម្លៃ Entry មួយ ទៅជួរដេកឃុំ/សង្កាត់ត្រូវគ្នា ក្នុង Period Sheet នៃ ស្រុក+កាលបរិច្ឆេទ ដែលបានផ្តល់មក
// (បង្កើត Period Sheet ថ្មីស្វ័យប្រវត្តិ បើមិនទាន់មាន)។ បើ oldId ត្រូវបានផ្តល់មក (ករណីកែប្រែ) ហើយទីតាំងថ្មី
// ខុសពីទីតាំងចាស់ (ប្តូរកាលបរិច្ឆេទ/ស្រុក/ឃុំសង្កាត់) — សម្អាតជួរដេកចាស់ត្រឡប់ទៅ ០ វិញ (មិនលុបជួរដេកចោលទេ ព្រោះ
// ជួរដេកនីមួយៗជាកម្មសិទ្ធិអចិន្ត្រៃយ៍របស់ឃុំ/សង្កាត់នោះ ក្នុង Period Sheet នោះ)។
function upsertDailyEntry_(currentUsername, entry, oldId) {
  var date = normalizeDateStr_(entry.date);
  // ==== ការពារការប៉ះទង្គិចគ្នា (Race Condition — FIX)៖ មុននេះ Function នេះគ្មានការចាក់សោទាល់តែសោះ ខុសពី
  // saveDistrictDayEntries (ជួរដេកច្រើន) ដែលមានចាក់សោរួចហើយ។ ការអាន (findCommuneRowIndex_) → សម្រេចថាត្រូវ Append
  // ជួរដេកថ្មីដែរឬអត់ → សរសេរ គឺមិនមែនប្រតិបត្តិការតែមួយឥតដាច់ទេ — បើសំណើ ២ កើតឡើងស្ទើរតែក្នុងពេលតែមួយសម្រាប់ឃុំ/
  // សង្កាត់ថ្មីមួយ (មិនទាន់ស្ថិតក្នុង COMMUNE_ORDER) ទាំងពីរអាចរកមិនឃើញជួរដេក ហើយទាំងពីរ Append ជួរដេកថ្មីដាច់ដោយ
  // ឡែក ធ្វើឲ្យឃុំ/សង្កាត់នោះមានជួរដេកស្ទួនគ្នា ២ (រាប់ទិន្នន័យស្ទួន)។ ដូចគ្នានេះដែរ ការធ្វើ Update/Delete ក្នុងពេល
  // ជិតគ្នាជាមួយ saveDistrictDayEntries (ដែលអាន Sheet ទាំងមូលចូលសតិ កែ រួចសរសេរជាន់ត្រឡប់ទាំងមូលវិញ) អាចជាន់
  // សរសេរគ្នាទៅវិញទៅមក ធ្វើឲ្យការកែ/លុបនោះបាត់ដោយស្ងាត់ស្ងៀម។
  // ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ ប្តូរទៅ acquireDistrictDayLock_() ដាច់ដោយឡែកតាមស្រុក+កាលបរិច្ឆេទ ដូចគ្នា
  // នឹង saveDistrictDayEntries — មើលការពន្យល់ពេញលេញនៅ Function នោះ ====
  var lockKey = acquireDistrictDayLock_(entry.district, date, 15000);
  if (!lockKey) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានការរក្សាទុកផ្សេងទៀតកំពុងដំណើរការសម្រាប់ស្រុក+ថ្ងៃដដែល)! សូមព្យាយាមរក្សាទុកម្តងទៀត។" };
  }
  try {
    var commune = String(entry.commune || '').trim();
    var sheet = getOrCreatePeriodSheet_(entry.district, date);
    var rIdx = findCommuneRowIndex_(sheet, commune);
    if (rIdx === -1) rIdx = appendCustomCommuneRow_(sheet, entry.district, date, commune);

    if (oldId) {
      var oldLoc = findDailyEntryLocation_(oldId);
      if (oldLoc && (oldLoc.sheet.getSheetId() !== sheet.getSheetId() || oldLoc.rowIndex !== rIdx)) {
        resetEntryRow_(oldLoc.sheet, oldLoc.rowIndex);
      }
    }

    var values = entry.values || {};
    var vals = DAILY_FIELDS.map(function(f) { return Number(values[f.key]) || 0; });
    sheet.getRange(rIdx, periodFieldColIndex_(DAILY_FIELDS[0].key), 1, DAILY_FIELDS.length).setValues([vals]);
    var noteCol = periodFieldColIndex_(DAILY_FIELDS[DAILY_FIELDS.length - 1].key) + 1;
    // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ចំណុចទី៥)៖ (១) ជួរឈរ "អ្នកបញ្ចូល" ឥឡូវប្រើ formatEnteredBy_()
    // (Auth.gs) ដូចគ្នាបេះបិទនឹង saveDistrictDayEntries() ខាងលើ (មុននេះ ២ផ្លូវនេះមិនស៊ីគ្នា៖ ទីនេះសរសេរ Username
    // ឆៅ ខណៈ saveDistrictDayEntries សរសេរឈ្មោះម្នាក់ឯង) — ឥឡូវទាំង២សរសេរ "Username - ឈ្មោះពេញ" ដូចគ្នា។
    // (២) ជួរឈរ "ចំនួនដងកែប្រែ" ថ្មី — អាន "តម្លៃចាស់" ជាមុនសិន (មុននឹងសរសេរជាន់) បូក ១ ====
    var oldEditCountU_ = Number(sheet.getRange(rIdx, noteCol + 3).getValue()) || 0;
    sheet.getRange(rIdx, noteCol, 1, 4).setValues([[sanitizeForSheetCell_(entry.note || ''), formatEnteredBy_(currentUsername), formatNow_(), oldEditCountU_ + 1]]);
    applyPeriodSheetFormulasAndFormatting_(sheet, periodSheetTotalCols_(), rIdx, rIdx);
    // ==== FIX (សំណើ "ពិនិត្យរូបមន្តបច្ចុប្បន្នភាពបញ្ជីឆ្នាំ២០២៦")៖ ដូចគ្នានឹង saveDistrictDayEntries — ធាតុតែមួយនេះ
    // ទើបផ្លាស់ប្តូរ ត្រូវជួសជុលខ្សែសង្វាក់ "ទិន្នន័យយោង" ថ្ងៃបន្តបន្ទាប់ (បើមាន) ផងដែរ ====
    cascadeRefreshRollingReferenceForward_(entry.district, date);
    scheduleRebuild_();
    // ==== ធានាថាតារាងខាងក្រោមឃើញការបញ្ចូល/កែប្រែធាតុតែមួយនេះភ្លាមៗ ដូចគ្នានឹង saveDistrictDayEntries/deleteDailyEntry
    // ដែរ (មិនចាំបាច់រង់ចាំ Trigger ១-៤ វិនាទីទេ) — មុននេះ addDailyEntry/updateDailyEntry (ហៅមុខងារនេះ) មិនបាន
    // ធ្វើបច្ចុប្បន្នភាព "សន្ទស្សន៍" ភ្លាមៗទេ ខុសពីមុខងារបញ្ចូលទិន្នន័យផ្សេងទៀត ធ្វើឲ្យតារាងអាចបង្ហាញទិន្នន័យចាស់
    // ក្នុងបង្អួចពេលខ្លីៗនោះ ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ ការហៅ fastReindexDistrictDay_() (ប៉ះពាល់សន្ទស្សន៍
    // រួម) នៅតែត្រូវការពារដោយ LockService.getScriptLock() ខ្លីៗសុទ្ធ ដូចគ្នានឹង saveDistrictDayEntries — មើលការ
    // ពន្យល់ពេញលេញនៅទីនោះ ====
    try {
      var lastRowU = sheet.getLastRow();
      if (lastRowU >= 4) {
        var freshDataU = sheet.getRange(4, 1, lastRowU - 3, periodSheetTotalCols_()).getValues();
        var idxLockU_ = LockService.getScriptLock();
        var gotIdxLockU_ = false;
        try { gotIdxLockU_ = idxLockU_.tryLock(10000); } catch (eIdxU_) {}
        if (gotIdxLockU_) {
          try { fastReindexDistrictDay_(entry.district, date, freshDataU); } finally { try { idxLockU_.releaseLock(); } catch (eRelIdxU_) {} }
        }
      }
    } catch (err) {}
    return { success: true };
  } finally {
    releaseDistrictDayLock_(lockKey);
  }
}

function addDailyEntry(currentUsername, entry, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!entry || !entry.date || !entry.district || !String(entry.commune || "").trim()) {
    return { success: false, message: "សូមបំពេញកាលបរិច្ឆេទ ស្រុក និងឃុំ/សង្កាត់ ឲ្យបានគ្រប់គ្រាន់!" };
  }
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ប្តូរពី canEditDistrict_ (កម្រិតស្រុកទាំងមូល) ទៅ canEditCommune_ (កម្រិត
  // ឃុំសង្កាត់ជាក់លាក់) — ត្រឹមត្រូវជាងសម្រាប់ Function នេះដែលធ្វើការលើធាតុតែមួយ (ឃុំសង្កាត់តែមួយ) ស្រាប់ ព្រមទាំង
  // ធានាថាគណនីឃុំសង្កាត់ បញ្ចូលបានតែឃុំសង្កាត់ខ្លួនផ្ទាល់ (Admin-tier/គ្រូប្រចាំក្រុងស្រុក នៅតែធ្វើការដូចមុនទាំងស្រុង) ====
  if (!canEditCommune_(currentUsername, entry.district, entry.commune)) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិបញ្ចូលទិន្នន័យសម្រាប់ស្រុក/ឃុំសង្កាត់នេះទេ!" };
  }
  if (isPastDateLocked_(currentUsername, entry.district, entry.date)) {
    return { success: false, message: "ថ្ងៃនេះបានកន្លងផុតទៅហើយ! សូមទាក់ទង SuperAdmin/Admin ដើម្បីបើកសិទ្ធិកែប្រែថ្ងៃនេះជាមុនសិន។" };
  }
  // ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព") — មើលការពន្យល់ពេញលេញនៅ isPrintLockedForDistrict_() ====
  if (isPrintLockedForDistrict_(currentUsername, entry.district, entry.date)) {
    return { success: false, message: "ថ្ងៃនេះកំពុងត្រូវបានចាក់សោបណ្តោះអាសន្ន ដើម្បីការពារទិន្នន័យអំឡុងពេលបោះពុម្ពរបាយការណ៍! សូមទាក់ទង SuperAdmin/Admin ។" };
  }
  // ==== FIX (សុវត្ថិភាព — ភាពមិនស៊ីសង្វាក់គ្នានៃច្បាប់អាជីវកម្ម)៖ saveDistrictDayEntries() (ផ្លូវសំខាន់ដែល Client
  // ពិតជាប្រើ) ការពាររួចហើយកុំឲ្យគ្រូប្រចាំស្រុក (District role) បង្កើត Period Sheet ថ្មីដោយខ្លួនឯង (ត្រូវឲ្យ Admin/
  // SuperAdmin/PEC21 ចុច "បង្កើត Sheet ថ្មីបន្ទាប់" ជាមុនសិន) ប៉ុន្តែ Function នេះ (addDailyEntry — មិនត្រូវបាន
  // Client ហៅតាមផ្លូវធម្មតាទេ ប៉ុន្តែនៅតែជា Top-level Function ហៅផ្ទាល់បានពី Browser ដូចគ្នា) ហៅ upsertDailyEntry_()
  // ដែលបង្កើត Sheet ថ្មីស្វ័យប្រវត្តិដដែល ដោយគ្មានការការពារដូចគ្នាទេ — ត្រូវឲ្យស្របគ្នា ====
  if (!isAdmin_(currentUsername)) {
    var dSsCheck_ = getDistrictSpreadsheet_(entry.district);
    if (!dSsCheck_.getSheetByName(periodSheetName_(normalizeDateStr_(entry.date)))) {
      return { success: false, message: "ថ្ងៃនេះមិនទាន់មាន Sheet ទេ! សូមទាក់ទង Admin ដើម្បីបង្កើត Sheet ថ្ងៃនេះជាមុនសិន (ចុច \"បង្កើត Sheet ថ្មីបន្ទាប់\")។" };
    }
  }
  var upRes = upsertDailyEntry_(currentUsername, entry, null);
  if (!upRes || !upRes.success) return upRes || { success: false, message: "កំហុសមិនស្គាល់មូលហេតុ!" };
  return { success: true, message: "បញ្ចូលទិន្នន័យជោគជ័យ!" };
}

function updateDailyEntry(currentUsername, entry, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (!entry || !entry.id) return { success: false, message: "ទិន្នន័យមិនត្រឹមត្រូវទេ!" };
  if (!entry.date || !entry.district || !String(entry.commune || "").trim()) {
    return { success: false, message: "សូមបំពេញកាលបរិច្ឆេទ ស្រុក និងឃុំ/សង្កាត់ ឲ្យបានគ្រប់គ្រាន់!" };
  }
  var loc = findDailyEntryLocation_(entry.id);
  if (!loc) return { success: false, message: "រកមិនឃើញទិន្នន័យនេះទេ!" };
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ប្តូរពី canEditDistrict_ ទៅ canEditCommune_ — ត្រូវត្រួតពិនិត្យទាំងទីតាំង
  // ដើម (loc — ឃុំសង្កាត់ចាស់) និងទីតាំងគោលដៅថ្មី (entry — ឃុំសង្កាត់ថ្មី បើកំពុងផ្លាស់ទី) ដើម្បីកុំឲ្យគណនីឃុំសង្កាត់
  // កែប្រែ/ផ្លាស់ទីទិន្នន័យទៅ/ចេញពីឃុំសង្កាត់ដទៃទៀតបានឡើយ ====
  // ==== FIX (ល្បឿន)៖ អាន "កាលបរិច្ឆេទ"+"ឃុំ/សង្កាត់" ក្នុងការហៅ getRange().getValues() តែមួយ (ជួរឈរជាប់គ្នា ២→
  // DAILY_META_PREFIX.length) ជំនួសការហៅដាច់ដោយឡែក ២ដង ====
  var locRow_ = loc.sheet.getRange(loc.rowIndex, 2, 1, DAILY_META_PREFIX.length - 1).getValues()[0];
  var locCommune_ = String(locRow_[DAILY_META_PREFIX.length - 2] || "").trim();
  if (!canEditCommune_(currentUsername, loc.district, locCommune_) || !canEditCommune_(currentUsername, entry.district, entry.commune)) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិកែប្រែទិន្នន័យនេះទេ!" };
  }
  var origDate_ = normalizeDateStr_(locRow_[0]);
  if (isPastDateLocked_(currentUsername, loc.district, origDate_) || isPastDateLocked_(currentUsername, entry.district, entry.date)) {
    return { success: false, message: "ថ្ងៃនេះបានកន្លងផុតទៅហើយ! សូមទាក់ទង SuperAdmin/Admin ដើម្បីបើកសិទ្ធិកែប្រែថ្ងៃនេះជាមុនសិន។" };
  }
  // ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព") — មើលការពន្យល់ពេញលេញនៅ isPrintLockedForDistrict_() ====
  if (isPrintLockedForDistrict_(currentUsername, loc.district, origDate_) || isPrintLockedForDistrict_(currentUsername, entry.district, entry.date)) {
    return { success: false, message: "ថ្ងៃនេះកំពុងត្រូវបានចាក់សោបណ្តោះអាសន្ន ដើម្បីការពារទិន្នន័យអំឡុងពេលបោះពុម្ពរបាយការណ៍! សូមទាក់ទង SuperAdmin/Admin ។" };
  }
  // ==== FIX (សុវត្ថិភាព — ភាពមិនស៊ីសង្វាក់គ្នានៃច្បាប់អាជីវកម្ម)៖ ដូចគ្នានឹង addDailyEntry() ខាងលើ — បើកំពុងផ្លាស់ទី
  // ធាតុនេះទៅ ស្រុក/ថ្ងៃ គោលដៅថ្មី (ខុសពី loc.district/origDate_ ដើម) ត្រូវឃាត់កុំឲ្យគ្រូប្រចាំស្រុកបង្កើត Period Sheet
  // ថ្មីដោយខ្លួនឯងតាមផ្លូវនេះដែរ (ត្រូវឲ្យ Admin បង្កើតជាមុនសិន ដូចគ្នានឹង saveDistrictDayEntries) ====
  if (!isAdmin_(currentUsername)) {
    var dSsCheck2_ = getDistrictSpreadsheet_(entry.district);
    if (!dSsCheck2_.getSheetByName(periodSheetName_(normalizeDateStr_(entry.date)))) {
      return { success: false, message: "ថ្ងៃនេះមិនទាន់មាន Sheet ទេ! សូមទាក់ទង Admin ដើម្បីបង្កើត Sheet ថ្ងៃនេះជាមុនសិន (ចុច \"បង្កើត Sheet ថ្មីបន្ទាប់\")។" };
    }
  }
  var upRes = upsertDailyEntry_(currentUsername, entry, entry.id);
  if (!upRes || !upRes.success) return upRes || { success: false, message: "កំហុសមិនស្គាល់មូលហេតុ!" };
  return { success: true, message: "កែប្រែទិន្នន័យជោគជ័យ!" };
}

function deleteDailyEntry(currentUsername, entryId, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var loc = findDailyEntryLocation_(entryId);
  if (!loc) return { success: false, message: "រកមិនឃើញទិន្នន័យនេះទេ!" };
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ ប្តូរពី canEditDistrict_ ទៅ canEditCommune_ (ដូច updateDailyEntry ខាងលើ) ====
  // ==== FIX (ល្បឿន)៖ អាន "កាលបរិច្ឆេទ"+"ឃុំ/សង្កាត់" ក្នុងការហៅ getRange().getValues() តែមួយ ដូច updateDailyEntry() ====
  var locRowDel_ = loc.sheet.getRange(loc.rowIndex, 2, 1, DAILY_META_PREFIX.length - 1).getValues()[0];
  var locCommuneDel_ = String(locRowDel_[DAILY_META_PREFIX.length - 2] || "").trim();
  if (!canEditCommune_(currentUsername, loc.district, locCommuneDel_)) {
    return { success: false, message: "អ្នកគ្មានសិទ្ធិលុបទិន្នន័យនេះទេ!" };
  }
  var entryDate = normalizeDateStr_(locRowDel_[0]);
  if (isPastDateLocked_(currentUsername, loc.district, entryDate)) {
    return { success: false, message: "ថ្ងៃនេះបានកន្លងផុតទៅហើយ! សូមទាក់ទង SuperAdmin/Admin ដើម្បីបើកសិទ្ធិកែប្រែថ្ងៃនេះជាមុនសិន។" };
  }
  // ==== FIX (សំណើថ្មី "ប៊ូតុងចាក់សោការពារទិន្នន័យមុនបោះពុម្ព") — មើលការពន្យល់ពេញលេញនៅ isPrintLockedForDistrict_() ====
  if (isPrintLockedForDistrict_(currentUsername, loc.district, entryDate)) {
    return { success: false, message: "ថ្ងៃនេះកំពុងត្រូវបានចាក់សោបណ្តោះអាសន្ន ដើម្បីការពារទិន្នន័យអំឡុងពេលបោះពុម្ពរបាយការណ៍! សូមទាក់ទង SuperAdmin/Admin ។" };
  }
  // ==== ការពារការប៉ះទង្គិចគ្នា (Race Condition — FIX)៖ ដូចគ្នានឹង upsertDailyEntry_ ខាងលើ — ចាក់សោដើម្បីកុំឲ្យការលុប
  // នេះជាន់គ្នាជាមួយ saveDistrictDayEntries (ដែលអាច "រស់ឡើងវិញ" ជួរដេកនេះ ដោយចៃដន្យ ប្រសិនបើវាកំពុងសរសេរជាន់
  // Snapshot ចាស់ត្រឡប់ទៅវិញក្នុងពេលជិតគ្នា)។ ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ ប្តូរទៅ acquireDistrictDayLock_()
  // ដាច់ដោយឡែកតាមស្រុក+កាលបរិច្ឆេទ ដូចគ្នានឹង saveDistrictDayEntries/upsertDailyEntry_ — មើលការពន្យល់ពេញលេញនៅ
  // acquireDistrictDayLock_() ខាងលើ ====
  var lockKey = acquireDistrictDayLock_(loc.district, entryDate, 15000);
  if (!lockKey) {
    return { success: false, message: "ប្រព័ន្ធកំពុងរវល់ (មានការរក្សាទុក/លុបផ្សេងទៀតកំពុងដំណើរការសម្រាប់ស្រុក+ថ្ងៃដដែល)! សូមព្យាយាមម្តងទៀត។" };
  }
  try {
    resetEntryRow_(loc.sheet, loc.rowIndex);
    // ==== FIX (សំណើ "ពិនិត្យរូបមន្តបច្ចុប្បន្នភាពបញ្ជីឆ្នាំ២០២៦")៖ ការលុបក៏ផ្លាស់ប្តូរចូល/ចេញរបស់ថ្ងៃនេះដែរ — ត្រូវ
    // ជួសជុលខ្សែសង្វាក់ "ទិន្នន័យយោង" ថ្ងៃបន្តបន្ទាប់ (បើមាន) ផងដែរ ====
    cascadeRefreshRollingReferenceForward_(loc.district, entryDate);
    scheduleRebuild_();
    // ==== ធានាថាតារាងខាងក្រោមឃើញការលុបភ្លាមៗ បើចុច 🔄 ភ្លាមៗបន្ទាប់ពីលុប (មិនចាំបាច់រង់ចាំ Trigger ពីក្រោយឆាកទេ)
    // ==== FIX (សំណើថ្មី "កែលម្អ Lock")៖ fastReindexDistrictDay_() នៅតែត្រូវការពារដោយ LockService.getScriptLock()
    // ខ្លីៗសុទ្ធ ដូចគ្នានឹង saveDistrictDayEntries/upsertDailyEntry_ — មើលការពន្យល់ពេញលេញនៅទីនោះ ====
    try {
      var lastRowD = loc.sheet.getLastRow();
      if (lastRowD >= 4) {
        var freshData = loc.sheet.getRange(4, 1, lastRowD - 3, periodSheetTotalCols_()).getValues();
        var idxLockD_ = LockService.getScriptLock();
        var gotIdxLockD_ = false;
        try { gotIdxLockD_ = idxLockD_.tryLock(10000); } catch (eIdxD_) {}
        if (gotIdxLockD_) {
          try { fastReindexDistrictDay_(loc.district, entryDate, freshData); } finally { try { idxLockD_.releaseLock(); } catch (eRelIdxD_) {} }
        }
      }
    } catch (err) {}
    return { success: true, message: "លុបទិន្នន័យជោគជ័យ!" };
  } finally {
    releaseDistrictDayLock_(lockKey);
  }
}

// ==================== ៧. ការបូកសរុប + សរសេរឡើងវិញនូវ Tab សរុបរួម ====================
// អាន Period Sheet គ្រប់ថ្ងៃខែ ក្នុង Spreadsheet ស្រុកទាំង១០ ជាប្រភពទិន្នន័យពិត
// បន្ទាប់មកសរសេរឡើងវិញ Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) សរុបរួម + Tab ស្ថិតិទាំង២
// ==== ល្បឿន៖ ហៅតាមរយៈ scheduleRebuild_() មិនមែនផ្ទាល់ទេ ====
// ការស្កេន Spreadsheet ស្រុកទាំង១០ (ដើម្បីធ្វើបច្ចុប្បន្នភាព Tab សរុប ដែលជា Tab សម្រាប់មនុស្សមើលប៉ុណ្ណោះ — App មិនបានអាន
// ពី Tab ទាំងនេះវិញទេ ព្រោះ getDailyEntries/computeSummaries_ អានផ្ទាល់ពី Period Sheet រួចហើយ) គឺជាការងារធ្ងន់បំផុត
// ក្នុងប្រព័ន្ធ។ បើហៅផ្ទាល់ (Synchronous) រាល់ពេលរក្សាទុក/លុប អ្នកប្រើនឹងត្រូវរង់ចាំយូរដោយអសារឥតការ។ ដូច្នេះ
// scheduleRebuild_() កំណត់ Trigger ឲ្យរត់វានៅពីក្រោយឆាកបន្តិចក្រោយមក (មិនធ្វើឲ្យអ្នកប្រើរង់ចាំ) ជំនួសវិញ។
function periodRowToMasterRow_(row) {
  var n = DAILY_FIELDS.length;
  var prefixLen = DAILY_META_PREFIX.length; // ID, កាលបរិច្ឆេទ, លេខកូដ, ស្រុក, ឃុំ/សង្កាត់
  var fieldsStart0 = prefixLen + PERIOD_EXTRA_FIELDS.length; // 0-indexed ចំណុចចាប់ផ្តើម DAILY_FIELDS ក្នុង Period Sheet (លំដាប់ពិតប្រាកដ — មិនប៉ះពាល់ទេ)
  var out = row.slice(0, prefixLen);
  // ==== សរសេរតម្លៃចេញតាមលំដាប់ "សម្រាប់បង្ហាញ" (getReportDisplayFields_) ដើម្បីឲ្យ Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ)
  // សរុបរួម ត្រូវនឹងក្បាលតារាងថ្មី (buildDailySheetHeader_) ខណៈពេលអានតម្លៃពិតប្រាកដពី Period Sheet (លំដាប់ដើម
  // មិនផ្លាស់ប្តូរ) ដោយប្រើ periodFieldColIndex_() ស្វែងរកទីតាំងពិតប្រាកដតាមឈ្មោះ Key ជំនួសលិបិក្រមលំដាប់ដើម ====
  getReportDisplayFields_().forEach(function(f) { out.push(row[periodFieldColIndex_(f.key) - 1]); });
  // ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាព..." — ជួរឈរ "ចំនួនដងកែប្រែ")៖ បន្ថែមតម្លៃទី៤ (fieldsStart0+n+3) ដើម្បីត្រូវនឹង
  // DAILY_META_SUFFIX ថ្មី (4ធាតុ ជំនួស 3) — buildDailySheetHeader_()/dailyTotalCols_() ខាងលើ ក៏ត្រូវបានពង្រីក
  // ដោយស្វ័យប្រវត្តិរួចហើយ (អាស្រ័យលើ DAILY_META_SUFFIX.length ដូចគ្នា) ====
  out.push(row[fieldsStart0 + n], row[fieldsStart0 + n + 1], row[fieldsStart0 + n + 2], row[fieldsStart0 + n + 3]);
  return out;
}
function rebuildAllDerivedData_() {
  // លុប Trigger បណ្តោះអាសន្នចោល (បើកំពុងហៅតាមរយៈ Trigger) — ជៀសវាងកុំឲ្យប្រមូលផ្តុំ Trigger ច្រើនលើសកំណត់
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'rebuildAllDerivedData_') ScriptApp.deleteTrigger(t);
    });
  } catch (err) {}

  var ss = getSS_();
  var totalCols = periodSheetTotalCols_();
  var masterCols = dailyTotalCols_();
  var communeMap = {};
  var districtMap = {};
  var masterRows = [];

  // ==== សន្ទស្សន៍ខាងក្នុង (SHEET_DAILY_INDEX) — ១ធាតុ ក្នុង Map ក្នុងមួយ "ស្រុក|កាលបរិច្ឆេទ" ====
  // ត្រូវបានសាងសង់ឡើងវិញទាំងស្រុងរាល់ពេលនេះរត់ (ដូចគ្នានឹង Tab សរុបរួមផ្សេងទៀតខាងក្រោម) ព្រោះការស្កេននេះ
  // ត្រូវការស្កេន Period Sheet គ្រប់ស្រុកទាំង១០ រួចហើយ (ការងារធ្ងន់បំផុតក្នុងប្រព័ន្ធ) ដូច្នេះមិនចាំបាច់ធ្វើម្តងទៀតទេ
  // — getDailyEntries() អានពី Sheet នេះជំនួស ជាជាងបើក Spreadsheet ស្រុកទាំង១០ រាល់ពេលអ្នកប្រើចុច 🔄
  var dailyIndexMap = {};
  var idxFieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1;
  var idxN = DAILY_FIELDS.length;
  var idxNoteCol = idxFieldsStart0 + idxN, idxByCol = idxFieldsStart0 + idxN + 1, idxAtCol = idxFieldsStart0 + idxN + 2;
  var idxEditCountCol = idxFieldsStart0 + idxN + 3;
  var idxCommuneCol0 = DAILY_META_PREFIX.length - 1;

  DISTRICT_LIST.forEach(function(district) {
    listDistrictPeriodSheets_(district).forEach(function(sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow < 4) return;
      // ==== FIX (សំខាន់ណាស់ — សុវត្ថិភាពទិន្នន័យ)៖ Function នេះស្កេន Period Sheet គ្រប់ស្រុក/គ្រប់ខែទាំងអស់ (រួមទាំង
      // Sheet ខែចាស់ដែលបិទរួច ដែលអាចមិនធ្លាប់ត្រូវបាន Migrate តាំងពីលក្ខណៈ "ចំនួនដងកែប្រែ" នេះមាន) ដោយផ្ទាល់ — ត្រូវ
      // ហៅ Migration នៅទីនេះជាដាច់ខាត (ដូចគ្នានឹងគ្រប់ចំណុចអានផ្ទាល់ផ្សេងទៀតទាំងអស់) មុននឹងអាន totalCols ទទឹងពេញ ====
      try { ensurePeriodSheetHasEditCountColumn_(sheet); } catch (errEnsure8) {}
      var data = sheet.getRange(4, 1, lastRow - 3, totalCols).getValues();
      data.forEach(function(row) {
        // ==== បន្ថែមទៅសន្ទស្សន៍ជាមុនសិន (មុននឹងរំលងជួរគ្មានទិន្នន័យ) ព្រោះ getDailyEntries() ត្រូវការឃើញ
        // ជួរឃុំ/សង្កាត់ទាំងអស់ (សូម្បីមិនទាន់មានទិន្នន័យ) ដើម្បីបង្ហាញជាតម្លៃទទេត្រឹមត្រូវក្នុងតារាង ====
        // ==== FIX (រាយការណ៍ថ្មី "ចំនួនឃុំ/សង្កាត់មិនត្រឹមត្រូវ")៖ ជួរដេក "សរុប" ស្វ័យប្រវត្តិ (ID ទទេ, ជួរឈរឃុំ/សង្កាត់ =
        // "សរុបទាំងអស់") មិនត្រូវត្រូវបានបញ្ចូលទៅក្នុងសន្ទស្សន៍ (dailyIndexMap) នេះជា "ធាតុឃុំ/សង្កាត់" មួយឡើយ — បើមិន
        // ដូច្នេះទេ getDailyEntries() ក្នុងផ្លូវ "គ្មានកាលបរិច្ឆេទច្រោះ" (ខាងក្រោម ~បន្ទាត់ ២០៦០) ដែលអាន idxMap[key]
        // ដោយផ្ទាល់ (មិនច្រោះតាម COMMUNE_ORDER ដូចផ្លូវផ្សេងទៀត) នឹងបង្ហាញជួរដេកនេះជា "ឃុំ/សង្កាត់" ក្លែងក្លាយមួយ
        // (ដែលមានលេខសរុបនៃឃុំ/សង្កាត់ទាំងអស់ — ច្រឡំធ្ងន់ធ្ងរ) ព្រោះ hasData របស់វាស្ទើរតែតែងតែ true (រូបមន្ត SUM
        // តម្លៃមិនមែន ០ ភាគច្រើន) ====
        var idxCommune = row[idxCommuneCol0];
        if (idxCommune && row[0]) {
          var idxDateStr = normalizeDateStr_(row[1]);
          var idxKey = district + '|' + idxDateStr;
          if (!dailyIndexMap[idxKey]) dailyIndexMap[idxKey] = [];
          var idxValues = {};
          DAILY_FIELDS.forEach(function(f, i) { idxValues[f.key] = row[idxFieldsStart0 + i]; });
          dailyIndexMap[idxKey].push({
            id: row[0], commune: idxCommune, values: idxValues,
            note: row[idxNoteCol], enteredBy: row[idxByCol], enteredAt: row[idxAtCol],
            editCount: Number(row[idxEditCountCol]) || 0,
            hasData: periodRowHasData_(row)
          });
        }

        if (!row[0] || !periodRowHasData_(row)) return; // រំលងជួរឃុំ/សង្កាត់ដែលមិនទាន់មានអ្នកបញ្ចូល
        var commune = row[DAILY_META_PREFIX.length - 1];
        if (!commune) return;
        masterRows.push(periodRowToMasterRow_(row));

        var cKey = district + "|" + commune;
        if (!communeMap[cKey]) {
          communeMap[cKey] = { district: district, commune: commune, sums: {}, days: 0 };
          DAILY_FIELDS.forEach(function(f) { communeMap[cKey].sums[f.key] = 0; });
        }
        if (!districtMap[district]) {
          districtMap[district] = { district: district, sums: {}, communes: {} };
          DAILY_FIELDS.forEach(function(f) { districtMap[district].sums[f.key] = 0; });
        }
        DAILY_FIELDS.forEach(function(f, i) {
          var v = Number(row[idxFieldsStart0 + i]) || 0;
          communeMap[cKey].sums[f.key] += v;
          districtMap[district].sums[f.key] += v;
        });
        communeMap[cKey].days += 1;
        districtMap[district].communes[commune] = true;
      });
    });
  });

  // ---- សរសេរឡើងវិញទាំងស្រុង Sheet(សន្ទស្សន៍ទិន្នន័យប្រចាំថ្ងៃ) ----
  try {
    var idxSheet = ensureDailyIndexSheet_(ss);
    var idxLastRow = idxSheet.getLastRow();
    if (idxLastRow > 1) idxSheet.getRange(2, 1, idxLastRow - 1, 4).clearContent();
    var idxNow = formatNow_();
    var idxRows = Object.keys(dailyIndexMap).map(function(key) {
      var parts = key.split('|');
      return [parts[0], parts.slice(1).join('|'), JSON.stringify(dailyIndexMap[key]), idxNow];
    });
    if (idxRows.length) idxSheet.getRange(2, 1, idxRows.length, 4).setValues(idxRows);
  } catch (err) {
    try { Logger.log("rebuildAllDerivedData_ index write error: " + err.message); } catch (e2) {}
  }
  // ==== FIX (ល្បឿន — "រក្សាទុកទិន្នន័យ" កម្រិតបន្ថែម)៖ Sheet(សន្ទស្សន៍) ខាងលើ ទើបត្រូវបានសរសេរជាន់ទាំងស្រុងឡើងវិញ
  // (លំដាប់ជួរដេកថ្មីទាំងស្រុង ខុសពីមុន — មិនអាស្រ័យលើលំដាប់ចាស់ទេ) ដូច្នេះ Cache ជួរដេក (getDailyIndexRowMap_)
  // ដែល fastReindexDistrictDay_() ប្រើ ត្រូវលុបចោលភ្លាមៗ បើមិនដូច្នេះទេ ការហៅបន្តបន្ទាប់ទៀតនឹងប្រើលេខជួរដេកចាស់
  // ខុស ហើយអាចសរសេរជាន់ខុសលើទិន្នន័យស្រុក/ថ្ងៃផ្សេង ====
  clearDailyIndexRowMapCache_();

  // ---- សរសេរឡើងវិញ Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) — សរុបរួមគ្រប់ស្រុក (Read-only សម្រាប់មនុស្ស) ----
  var masterSheet = ensureDailySheet_(ss);
  var lastRowMaster = masterSheet.getLastRow();
  if (lastRowMaster > 3) masterSheet.getRange(4, 1, lastRowMaster - 3, masterCols).clearContent();
  if (masterRows.length) masterSheet.getRange(4, 1, masterRows.length, masterCols).setValues(masterRows);

  // ---- សរសេរទៅ Sheet(របាយការណ៍សរុបតាមឃុំសង្កាត់) ----
  var communeSheet = ensureCommuneSumSheet_(ss);
  var communeCols = communeSumTotalCols_();
  var lastRowC = communeSheet.getLastRow();
  if (lastRowC > 3) communeSheet.getRange(4, 1, lastRowC - 3, communeCols).clearContent();
  var communeRows = [];
  for (var key in communeMap) {
    var c = communeMap[key];
    var row = [c.district, c.commune];
    // ==== ប្រើលំដាប់ "សម្រាប់បង្ហាញ" (getReportDisplayFields_) ដើម្បីឲ្យត្រូវនឹងក្បាលតារាងថ្មី (buildSumSheetHeader_)
    // — c.sums ស្វែងរកតម្លៃដោយ Key (មិនមែនលិបិក្រម) ដូច្នេះលទ្ធផលបូកសរុបខ្លួនឯង មិនប៉ះពាល់ដោយការប្តូរលំដាប់នេះឡើយ ====
    getReportDisplayFields_().forEach(function(f) { row.push(c.sums[f.key]); });
    row.push(c.days, formatNow_());
    communeRows.push(row);
  }
  if (communeRows.length) communeSheet.getRange(4, 1, communeRows.length, communeCols).setValues(communeRows);

  // ---- សរសេរទៅ Sheet(របាយការណ៍សរុបខេត្ត) ----
  var provinceSheet = ensureProvinceSumSheet_(ss);
  var provinceCols = provinceSumTotalCols_();
  var lastRowP = provinceSheet.getLastRow();
  if (lastRowP > 3) provinceSheet.getRange(4, 1, lastRowP - 3, provinceCols).clearContent();
  var provinceRows = [];
  for (var d in districtMap) {
    var p = districtMap[d];
    var row = [p.district];
    // ==== ប្រើលំដាប់ "សម្រាប់បង្ហាញ" (getReportDisplayFields_) — ដូចគ្នានឹង communeRows ខាងលើ ====
    getReportDisplayFields_().forEach(function(f) { row.push(p.sums[f.key]); });
    row.push(Object.keys(p.communes).length, formatNow_());
    provinceRows.push(row);
  }
  if (provinceRows.length) provinceSheet.getRange(4, 1, provinceRows.length, provinceCols).setValues(provinceRows);
}

// ==================== ៧.១ បូកសរុបតាមចន្លោះកាលបរិច្ឆេទ (យោងពី១ថ្ងៃទៅមួយថ្ងៃ) — គណនាផ្ទាល់ពី Period Sheet គ្រប់ថ្ងៃខែ ====================
// fromDate / toDate ជា string ទម្រង់ "YYYY-MM-DD" (ដូចគ្នានឹង <input type="date">) ។
// ទុកទទេ (falsy) មានន័យថា មិនកំណត់ព្រំដែននោះទេ (ឧ. ទុកទទេទាំង២ = បូកសរុបគ្រប់ថ្ងៃទាំងអស់ដែលធ្លាប់មាន)។
// រាល់ថ្ងៃដែលមិនទាន់មានទិន្នន័យបញ្ចូល នឹងមិនចូលរួមក្នុងការបូកសរុបទេ (ចាត់ទុកជា 0 ដោយស្វ័យប្រវត្តិ) ។
