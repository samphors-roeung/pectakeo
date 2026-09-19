// ==================== Utils.gs — ថេរ + Function ជំនួយរួម (District/Commune Reference, កាលបរិច្ឆេទ, Cache, Sheet មូលដ្ឋាន) ====================

/*******************************************************************
 * ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ (ខេត្តតាកែវ)
 * Code.gs — Server-side (Google Apps Script)
 * Sheet ID ត្រូវបានភ្ជាប់ដោយផ្ទាល់ (Bound) ទៅកាន់ Google Sheet ដែលបានផ្ដល់
 *******************************************************************/

var TARGET_SHEET_ID = '1M3H0vSTl9w9yfUZCR9ElXqXU4MXOKlcR-tpGERyCsdU';

// ==== ត្រឡប់ Spreadsheet ID "មេ" ដែលកំពុងប្រើប្រាស់ជាក់ស្តែង — ប្រើតំណភ្ជាប់ដែល Admin បិទភ្ជាប់ (Copy-Paste) ក្នុង
// Tab "ការកំណត់ប្រព័ន្ធ" (setting key: masterSheetLink) បើមាន បើមិនមាន ត្រឡប់ទៅ TARGET_SHEET_ID ដើមវិញ (លំនាំដើម)។
// ថត Cache ក្នុងការប្រតិបត្តិតែមួយ (Execution) ដើម្បីជៀសវាងការហៅ getSystemSettings_() ដដែលៗច្រើនដង ====
var EFFECTIVE_SHEET_ID_CACHE_ = null;
function getTargetSheetId_() {
  if (EFFECTIVE_SHEET_ID_CACHE_) return EFFECTIVE_SHEET_ID_CACHE_;
  var id = TARGET_SHEET_ID;
  try {
    var settings = getSystemSettings_();
    var customId = extractSpreadsheetId_(settings.masterSheetLink);
    if (customId) id = customId;
  } catch (err) {}
  EFFECTIVE_SHEET_ID_CACHE_ = id;
  return id;
}

// ==================== ០. តួនាទី និងបញ្ជីស្រុក ====================
var ROLE_SUPERADMIN = "SuperAdmin";
var ROLE_ADMIN = "Admin";
var ROLE_PEC21 = "PEC21";
// ==== "សង្កេតការណ៍" (Observer) — មើលទិន្នន័យបានគ្រប់ស្រុក (ដូច PEC21) ប៉ុន្តែមិនអាចបញ្ចូល/កែ/លុបទិន្នន័យ ឬបង្កើត
// Sheet ថ្មីបានឡើយ (ខុសពី PEC21 ដែលអាចបង្កើត Sheet ថ្មីបាន) — អាចសរសេរមតិយោបល់រាយការណ៍បញ្ហាទៅ SuperAdmin/គ្រូប្រចាំស្រុកបាន
// (មើល Comments.gs) ។ គណនីតួនាទីនេះត្រូវបង្កើតដោយ SuperAdmin/Admin ប៉ុណ្ណោះ (មិនអាចស្នើសុំដោយខ្លួនឯងបានទេ) ។
var ROLE_OBSERVER = "សង្កេតការណ៍";
var STATUS_ACTIVE = "សកម្ម";
var STATUS_PENDING = "រង់ចាំអនុម័ត";
var STATUS_SUSPENDED = "ផ្អាក";
var VISIT_COUNTER_PROP_ = "visitCounter_v1";

// ស្រុកទាំង១០ ក្នុងខេត្តតាកែវ — ត្រូវប្រើឈ្មោះទាំងនេះជា "Role" សម្រាប់គណនីគ្រូប្រចាំស្រុកនីមួយៗ
var DISTRICT_LIST = ["អង្គរបូរី", "បាទី", "បូរីជលសារ", "គិរីវង់", "កោះអណ្តែត", "ព្រៃកប្បាស", "សំរោង", "ដូនកែវ", "ត្រាំកក់", "ទ្រាំង"];

function isDistrictRole_(role) {
  return DISTRICT_LIST.indexOf(role) > -1;
}

// ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ តួនាទីថ្មី សម្រាប់គណនីទិន្នន័យកម្រិតឃុំសង្កាត់ (ក្រោមកម្រិតគ្រូប្រចាំក្រុងស្រុក) —
// ចំណាំសំខាន់៖ មិនអាចប្រើលេយបញ្ចូល "role = ឈ្មោះឃុំសង្កាត់" ដូចលំនាំគណនីស្រុក (role = ឈ្មោះស្រុក) បានទេ ព្រោះ
// ឈ្មោះឃុំសង្កាត់ មិនមែនតែមួយគត់ទូទាំងខេត្តទេ (ឧ. "កំពែង" មាននៅទាំងក្នុងស្រុកគិរីវង់ និងព្រៃកប្បាស, "អង្កាញ់" មាននៅ
// ព្រៃកប្បាស និងទ្រាំង) — ដូច្នេះត្រូវប្រើថេរ ROLE_COMMUNE តែមួយរួម ហើយផ្ទុកស្រុក+ឃុំសង្កាត់ដាច់ដោយឡែក (ជួរឈរផ្ទាល់ខ្លួន
// ក្នុង Sheet(អ្នកប្រើប្រាស់)) ដើម្បីកំណត់អត្តសញ្ញាណគណនីនីមួយៗ ====
var ROLE_COMMUNE = "អ្នកបញ្ចូលទិន្នន័យឃុំសង្កាត់";
function isCommuneRole_(role) {
  return role === ROLE_COMMUNE;
}

// ==== FIX (Fix129, "តួនាទីថ្មីទាំង៥ មានសិទ្ធិដូចសង្កេតការណ៍")៖ តួនាទីថ្មីទាំង៥ (មើលទិន្នន័យបានគ្រប់ស្រុកដូច
// ROLE_OBSERVER ខាងលើ ប៉ុន្តែមិនអាចបញ្ចូល/កែ/លុបទិន្នន័យបានឡើយ) — ប្រមូលផ្តុំជា Array តែមួយ (មិនប្រកាស Role
// Constant ដាច់ដោយឡែកនីមួយៗទេ) ដើម្បីងាយស្រួលបន្ថែម/គ្រប់គ្រងក្នុងកន្លែងតែមួយ (isExpiryApplicableRole_,
// canViewAllDistricts_, getUsersList, requestSignup, addUserAccount, syncSheetPermissions_, Index.html) —
// នីមួយៗមាន "settingKey" ផ្ទាល់ខ្លួន សម្រាប់ជម្រើសបើក/បិទការបង្កើតគណនីប្រភេទនោះដាច់ដោយឡែក (មើល
// isObserverTierRoleEnabled_ ក្នុង Settings.gs — លំនាំដើមទាំងអស់ "បិទ" ខុសពីគណនីឃុំសង្កាត់ Fix127/128) ====
// ==== FIX (Fix130, "Tab(គណនីភាគីពាក់ព័ន្ធ) ថ្មី")៖ Field "tab" ថ្មី — កំណត់ថាតួនាទីនេះគ្រប់គ្រង/បង្ហាញនៅ Tab
// ណា (មិនប៉ះពាល់សិទ្ធិ/ការគ្រប់គ្រងបើក-បិទខាងលើទេ — សុទ្ធសឹងជាចំណុចរៀបចំ UI តែប៉ុណ្ណោះ)៖ "users" = "មន្ត្រី ល.ខ.ប
// តាកែវ" ដែលនៅតែផ្ទុកក្នុង Tab(អ្នកប្រើប្រាស់) ដដែល, "stakeholder" = ៤ តួនាទីទៀត ដែលផ្ទុកក្នុង Tab
// "គណនីភាគីពាក់ព័ន្ធ" ថ្មី (រួមជាមួយ ROLE_OBSERVER ដែលមិនស្ថិតក្នុង Array នេះ — មើល getStakeholderUsersList) ====
var OBSERVER_TIER_ROLES_ = [
  { role: "មន្ត្រី ល.ខ.ប តាកែវ", settingKey: "enableRoleNecOfficer", accountType: "necofficer", optionId: "suAcctTypeNecOfficerOpt", label: "មន្ត្រី ល.ខ.ប តាកែវ", tab: "users" },
  { role: "រដ្ឋបាលខេត្ត", settingKey: "enableRoleProvinceAdmin", accountType: "provinceadmin", optionId: "suAcctTypeProvinceAdminOpt", label: "រដ្ឋបាលខេត្ត", tab: "stakeholder" },
  { role: "រដ្ឋបាលក្រុង ស្រុក", settingKey: "enableRoleDistrictAdmin", accountType: "districtadmin", optionId: "suAcctTypeDistrictAdminOpt", label: "រដ្ឋបាលក្រុង ស្រុក", tab: "stakeholder" },
  { role: "កងកម្លាំងទាំងបី", settingKey: "enableRoleArmedForces", accountType: "armedforces", optionId: "suAcctTypeArmedForcesOpt", label: "កងកម្លាំងទាំងបី", tab: "stakeholder" },
  { role: "គណបក្សនយោបាយ", settingKey: "enableRolePoliticalParty", accountType: "politicalparty", optionId: "suAcctTypePoliticalPartyOpt", label: "គណបក្សនយោបាយ", tab: "stakeholder" }
];
function getObserverTierRoleConfigByRole_(role) {
  for (var i = 0; i < OBSERVER_TIER_ROLES_.length; i++) {
    if (OBSERVER_TIER_ROLES_[i].role === role) return OBSERVER_TIER_ROLES_[i];
  }
  return null;
}
function getObserverTierRoleConfigByAccountType_(accountType) {
  for (var i = 0; i < OBSERVER_TIER_ROLES_.length; i++) {
    if (OBSERVER_TIER_ROLES_[i].accountType === accountType) return OBSERVER_TIER_ROLES_[i];
  }
  return null;
}
function isObserverTierRole_(role) {
  return getObserverTierRoleConfigByRole_(role) !== null;
}
// ==== ផ្ទៀងផ្ទាត់ថា (ស្រុក, ឃុំសង្កាត់) ជាគូដែលមានពិតប្រាកដក្នុង COMMUNE_ORDER (ការពារកុំឲ្យបញ្ចូលឈ្មោះខុស/មិនមាន) ====
function isValidCommune_(district, commune) {
  var arr = COMMUNE_ORDER[district];
  if (!arr) return false;
  for (var i = 0; i < arr.length; i++) { if (arr[i].name === commune) return true; }
  return false;
}

// ==================== ០.១ លំដាប់ផ្លូវការនៃ ស្រុក/ក្រុង និងឃុំ/សង្កាត់ (តាមលេខកូដផ្លូវការ) ====================
// ប្រើសម្រាប់តម្រៀបលំដាប់ និងបង្ហាញលេខកូដ ក្នុងរបាយការណ៍ (Print Preview + Excel Export)
// ជំនួសការតម្រៀបតាមអក្ខរក្រម ដែលអាចនឹងលំដាប់មិនត្រូវនឹងលេខកូដពិត។ DISTRICT_LIST ខាងលើ ស្រាប់តែជាលំដាប់ត្រឹមត្រូវរួចហើយ (21-01→21-10)។
var DISTRICT_TYPE = {
  "ដូនកែវ": "ក្រុង" // ស្រុកទាំងអស់ក្រៅពីនេះ ជា "ស្រុក" ដោយស្វ័យប្រវត្តិ (មើល getDistrictType_)
};

function getDistrictType_(district) {
  return DISTRICT_TYPE[district] || "ស្រុក";
}

// លេខកូដ ស្រុក/ក្រុង (21-01 ... 21-10)
var DISTRICT_CODE = {
  "អង្គរបូរី": "21-01", "បាទី": "21-02", "បូរីជលសារ": "21-03", "គិរីវង់": "21-04", "កោះអណ្តែត": "21-05",
  "ព្រៃកប្បាស": "21-06", "សំរោង": "21-07", "ដូនកែវ": "21-08", "ត្រាំកក់": "21-09", "ទ្រាំង": "21-10"
};
function districtCode_(district) {
  return DISTRICT_CODE[district] || '';
}

// ឈ្មោះ+លេខកូដឃុំ/សង្កាត់ តាមលំដាប់លេខកូដផ្លូវការ សម្រាប់ស្រុក/ក្រុងនីមួយៗ (21-001 ... 21-100)
var COMMUNE_ORDER = {
  "អង្គរបូរី": [
    { code: "21-001", name: "អង្គរបូរី" }, { code: "21-002", name: "បាស្រែ" }, { code: "21-003", name: "គោកធ្លក" },
    { code: "21-004", name: "ពន្លៃ" }, { code: "21-005", name: "ព្រែកផ្ទោល" }, { code: "21-006", name: "ព្រៃផ្គាំ" }
  ],
  "បាទី": [
    { code: "21-007", name: "ចំបក់" }, { code: "21-008", name: "ចំប៉ី" }, { code: "21-009", name: "ដូង" },
    { code: "21-010", name: "កណ្តឹង" }, { code: "21-011", name: "កុមាររាជា" }, { code: "21-012", name: "ក្រាំងលាវ" },
    { code: "21-013", name: "ក្រាំងធ្នង់" }, { code: "21-014", name: "លំពង់" }, { code: "21-015", name: "ពារាម" },
    { code: "21-016", name: "ពត់សរ" }, { code: "21-017", name: "សូរភី" }, { code: "21-018", name: "តាំងដូង" },
    { code: "21-019", name: "ត្នោត" }, { code: "21-020", name: "ត្រពាំងក្រសាំង" }, { code: "21-021", name: "ត្រពាំងសាប" }
  ],
  "បូរីជលសារ": [
    { code: "21-022", name: "បូរីជលសារ" }, { code: "21-023", name: "ជ័យជោគ" }, { code: "21-024", name: "ដួងខ្ពស់" },
    { code: "21-025", name: "កំពង់ក្រសាំង" }, { code: "21-026", name: "គោកពោធិ៍" }
  ],
  "គិរីវង់": [
    { code: "21-027", name: "អង្គប្រាសាទ" }, { code: "21-028", name: "ព្រះបាទជាន់ជុំ" }, { code: "21-029", name: "កំណប់" },
    { code: "21-030", name: "កំពែង" }, { code: "21-031", name: "គីរីចុងកោះ" }, { code: "21-032", name: "គោកព្រេច" },
    { code: "21-033", name: "ភ្នំដិន" }, { code: "21-034", name: "ព្រៃអំពក" }, { code: "21-035", name: "ព្រៃរំដេង" },
    { code: "21-036", name: "រាមអណ្តើក" }, { code: "21-037", name: "សោម" }, { code: "21-038", name: "តាអូរ" }
  ],
  "កោះអណ្តែត": [
    { code: "21-039", name: "ក្រពុំឈូក" }, { code: "21-040", name: "ពេជសារ" }, { code: "21-041", name: "ព្រៃខ្លា" },
    { code: "21-042", name: "ព្រៃយុថ្កា" }, { code: "21-043", name: "រមេញ" }, { code: "21-044", name: "ធ្លាប្រជុំ" }
  ],
  "ព្រៃកប្បាស": [
    { code: "21-045", name: "អង្កាញ់" }, { code: "21-046", name: "បានកាម" }, { code: "21-047", name: "ចំប៉ា" },
    { code: "21-048", name: "ចារ" }, { code: "21-049", name: "កំពែង" }, { code: "21-050", name: "កំពង់រាប" },
    { code: "21-051", name: "ក្តាញ់" }, { code: "21-052", name: "ពោធិ៍រំចាក" }, { code: "21-053", name: "ព្រៃកប្បាស" },
    { code: "21-054", name: "ព្រៃល្វា" }, { code: "21-055", name: "ព្រៃផ្តៅ" }, { code: "21-056", name: "ស្នោ" },
    { code: "21-057", name: "តាំងយ៉ាប" }
  ],
  "សំរោង": [
    { code: "21-058", name: "បឹងត្រាញ់ខាងជើង" }, { code: "21-059", name: "បឹងត្រាញ់ខាងត្បូង" }, { code: "21-060", name: "ជើងគួន" },
    { code: "21-061", name: "ជំរះពេន" }, { code: "21-062", name: "ខ្វាវ" }, { code: "21-063", name: "លំចង់" },
    { code: "21-064", name: "រវៀង" }, { code: "21-065", name: "សំរោង" }, { code: "21-066", name: "សឹង្ហ" },
    { code: "21-067", name: "ស្លា" }, { code: "21-068", name: "ទ្រា" }
  ],
  "ដូនកែវ": [
    { code: "21-069", name: "បារាយណ៍" }, { code: "21-070", name: "រកាក្នុង" }, { code: "21-071", name: "រកាក្រៅ" }
  ],
  "ត្រាំកក់": [
    { code: "21-072", name: "អង្គតាសោម" }, { code: "21-073", name: "ជាងទង" }, { code: "21-074", name: "គុស" },
    { code: "21-075", name: "លាយបូរ" }, { code: "21-076", name: "ញ៉ែងញ៉ង" }, { code: "21-077", name: "អូរសារាយ" },
    { code: "21-078", name: "ឧត្តមសុរិយា" }, { code: "21-079", name: "ពពេល" }, { code: "21-080", name: "សំរោង" },
    { code: "21-081", name: "ស្រែរនោង" }, { code: "21-082", name: "តាភេម" }, { code: "21-083", name: "ត្រាំកក់" },
    { code: "21-084", name: "ត្រពាំងធំខាងជើង" }, { code: "21-085", name: "ត្រពាំងធំខាងត្បូង" }, { code: "21-099", name: "ត្រពាំងក្រញូង" }
  ],
  "ទ្រាំង": [
    { code: "21-086", name: "អង្កាញ់" }, { code: "21-087", name: "អង្គខ្នុរ" }, { code: "21-088", name: "ជីខ្មា" },
    { code: "21-089", name: "ខ្វាវ" }, { code: "21-090", name: "ប្រាំបីមុំ" }, { code: "21-091", name: "ព្រៃស្លឹក" },
    { code: "21-092", name: "រនាម" }, { code: "21-093", name: "សំបួរ" }, { code: "21-094", name: "សន្លុង" },
    { code: "21-095", name: "ស្មោង" }, { code: "21-096", name: "ស្រង៉ែ" }, { code: "21-097", name: "ធ្លក" },
    { code: "21-098", name: "ត្រឡាច" }, { code: "21-100", name: "អង្គកែវ" }
  ]
};

function districtSortIndex_(district) {
  var idx = DISTRICT_LIST.indexOf(district);
  return idx === -1 ? 999 : idx;
}
function communeSortIndex_(district, commune) {
  var arr = COMMUNE_ORDER[district];
  if (!arr) return 999;
  for (var i = 0; i < arr.length; i++) { if (arr[i].name === commune) return i; }
  return 999;
}
function communeCode_(district, commune) {
  var arr = COMMUNE_ORDER[district];
  if (!arr) return '';
  for (var i = 0; i < arr.length; i++) { if (arr[i].name === commune) return arr[i].code; }
  return '';
}
function districtCommuneComparator_(a, b) {
  var da = districtSortIndex_(a.district), db = districtSortIndex_(b.district);
  if (da !== db) return da - db;
  var ca = communeSortIndex_(a.district, a.commune), cb = communeSortIndex_(b.district, b.commune);
  if (ca !== cb) return ca - cb;
  return String(a.commune).localeCompare(String(b.commune), 'km');
}
function districtOnlyComparator_(a, b) {
  var da = districtSortIndex_(a.district), db = districtSortIndex_(b.district);
  if (da !== db) return da - db;
  return String(a.district).localeCompare(String(b.district), 'km');
}

// បម្លែងកាលបរិច្ឆេទ "YYYY-MM-DD" ទៅជាអក្សរខ្មែរ ទម្រង់ "ថ្ងៃទី... ខែ... ឆ្នាំ..." (ប្រើសម្រាប់ចំណងជើងរបាយការណ៍ Excel)
var KHMER_DIGITS_ = ['០','១','២','៣','៤','៥','៦','៧','៨','៩'];
var KHMER_MONTHS_ = ['មករា','កុម្ភៈ','មីនា','មេសា','ឧសភា','មិថុនា','កក្កដា','សីហា','កញ្ញា','តុលា','វិច្ឆិកា','ធ្នូ'];
function toKhmerDigits_(n) {
  return String(n).replace(/[0-9]/g, function(d) { return KHMER_DIGITS_[Number(d)]; });
}
function formatKhmerDate_(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  var y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
  return 'ថ្ងៃទី' + toKhmerDigits_(d) + ' ខែ' + (KHMER_MONTHS_[m - 1] || '') + ' ឆ្នាំ' + toKhmerDigits_(y);
}

// ==================== ១. ឈ្មោះ Sheet ====================
var SHEET_USERS = "អ្នកប្រើប្រាស់";
var SHEET_DAILY = "របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ";
var SHEET_COMMUNE_SUM = "របាយការណ៍សរុបតាមឃុំសង្កាត់";
var SHEET_PROVINCE_SUM = "របាយការណ៍សរុបខេត្ត";
var SHEET_COMMUNE_LIST = "ឃុំ សង្កាត់";
var SHEET_SETTINGS = "ការកំណត់ប្រព័ន្ធ";
// ==== សន្ទស្សន៍ខាងក្នុង (កុំកែដោយដៃ) ៖ ថតចម្លងទិន្នន័យប្រចាំថ្ងៃរបស់ស្រុកទាំង១០ ត្រូវបានធ្វើបច្ចុប្បន្នភាពដោយស្វ័យប្រវត្តិ
// ជានិច្ចតាមរយៈ rebuildAllDerivedData_() — ប្រើសម្រាប់អានលឿន (ជំនួសការបើក Spreadsheet ដាច់ដោយឡែករបស់ស្រុកទាំង១០
// រាល់ពេល ដែលយឺតជាងគួរឲ្យកត់សម្គាល់ តាមចំនួន Tab កើនឡើងក្នុង Spreadsheet ស្រុកនីមួយៗ) ====
var SHEET_DAILY_INDEX = "សន្ទស្សន៍ទិន្នន័យប្រចាំថ្ងៃ (កុំកែដោយដៃ)";
var SETTINGS_KEY_START_DATE = "ថ្ងៃចាប់ផ្តើមដំណើរការ";
var SETTINGS_KEY_END_DATE = "ថ្ងៃចុងបញ្ចប់ដំណើរការ";

// ==== FIX (Fix119, "Retry Logic ស្វ័យប្រវត្តិ")៖ ពេលអ្នកប្រើប្រាស់ច្រើននាក់ Login/ស្នើសុំសកម្មភាពក្នុងពេលជិតគ្នា
// (ឧ. ដើមម៉ោងធ្វើការទាំងអស់គ្នា) ប្រតិបត្តិការសំខាន់ៗមួយចំនួន (ការបើក Spreadsheet, ការហៅ Telegram API) អាចជួប
// កំហុសបណ្តោះអាសន្ន (Transient — ឧ. កូតា/ដែនកំណត់ Google, បណ្តាញយឺតមួយភ្លែត) ដែលជាធម្មតាបើព្យាយាមម្តងទៀតភ្លាមៗ
// នឹងជោគជ័យវិញ។ Helper ទូទៅនេះព្យាយាមហៅ fn() ម្តងទៀតដោយស្វ័យប្រវត្តិ (មិនច្រើនដងហួសហេតុ — Backoff កើនឡើងជា
// លំដាប់ រវាងការព្យាយាមនីមួយៗ) មុននឹងបោះ Error ចុងក្រោយបំផុតចេញពិតប្រាកដ ====
function retryWithBackoff_(fn, maxAttempts, baseDelayMs) {
  maxAttempts = maxAttempts || 3;
  baseDelayMs = baseDelayMs || 300;
  var lastErr = null;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      if (attempt < maxAttempts) {
        try { Utilities.sleep(baseDelayMs * attempt); } catch (eSleep) {}
      }
    }
  }
  throw lastErr;
}

var MASTER_SS_CACHE_ = null; // ត្រូវកាចប្រើតែក្នុងមួយ Execution (មិនចែករំលែងរវាង Request ផ្សេងគ្នាទេ) — ជៀសវាងបើក Spreadsheet ដដែលៗ
function getSS_() {
  if (MASTER_SS_CACHE_) return MASTER_SS_CACHE_;
  // ==== FIX (ការពារកុំឲ្យប្រព័ន្ធទាំងមូលដួលរលំ)៖ saveSystemSettings() ឥឡូវផ្ទៀងផ្ទាត់ masterSheetLink មុននឹងទទួលយក
  // ប៉ុន្តែ Spreadsheet ដែលធ្លាប់ត្រឹមត្រូវ អាចនៅតែក្លាយជាចូលមិនកើតនៅពេលក្រោយ (ត្រូវបានលុប/ដកសិទ្ធិ/Trash ក្រោយពេលកំណត់
  // រួច) — ក្នុងករណីនោះ ត្រូវប្តូរទៅប្រើ Spreadsheet លំនាំដើម (TARGET_SHEET_ID) វិញដោយស្វ័យប្រវត្តិ ជំនួសឲ្យបោះ Error
  // ធ្វើឲ្យអ្នកប្រើប្រាស់ទាំងអស់ (រួមទាំង Admin ខ្លួនឯង ដែលជាមនុស្សតែម្នាក់គត់អាចកែតម្រូវការកំណត់នេះវិញបាន) ជាប់គាំង
  // ដោយគ្មានផ្លូវសម្រួលដោយខ្លួនឯង ====
  // ==== FIX (Fix119)៖ ព្យាយាមបើក ID ផ្ទាល់ខ្លួន (custom) ២ដងជាមុនសិន (Retry — ការពារកំហុសបណ្តោះអាសន្នច្រឡំថា
  // Spreadsheet នេះខូច ទាំងដែលពិតជាត្រឹមត្រូវ ធ្វើឲ្យប្តូរទៅ Spreadsheet លំនាំដើមខុសដោយអចេតនា) មុននឹងសន្មតថាខូចពិត
  // ហើយប្តូរទៅលំនាំដើមវិញ ====
  var targetId = getTargetSheetId_();
  try {
    MASTER_SS_CACHE_ = retryWithBackoff_(function() { return SpreadsheetApp.openById(targetId); }, 2, 300);
  } catch (err) {
    MASTER_SS_CACHE_ = SpreadsheetApp.openById(TARGET_SHEET_ID);
  }
  return MASTER_SS_CACHE_;
}

// ត្រឡប់តំណភ្ជាប់ (Link) ចង្អុលទៅ Tab ជាក់លាក់មួយ តាម Spreadsheet ID + sheet object (ដំណើរការជាមួយ Spreadsheet ណាមួយក៏បាន)
function buildSheetTabUrl_(spreadsheetId, sheet) {
  return 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit#gid=' + sheet.getSheetId();
}

// ត្រឡប់ព័ត៌មានប៊ូតុង "បើក Google Sheet" ដោយផ្អែកលើតួនាទីអ្នកប្រើប្រាស់៖
// - SuperAdmin/Admin/PEC21 → ចូលទៅ Spreadsheet មេ (Tab សរុបរួម មើលទិន្នន័យគ្រប់ស្រុក)
// - គ្រូប្រចាំស្រុក (District role) → ចូលទៅ Spreadsheet ដាច់ដោយឡែករបស់ស្រុកខ្លួន (File ខុសគ្នាទាំងស្រុងពី Spreadsheet មេ
//   និងស្រុកដទៃទៀត) ត្រង់ Tab ថ្ងៃខែបច្ចុប្បន្ន (Period Sheet ឧ. "25-08") ដែលជាប្រភពទិន្នន័យតែមួយគត់ឥឡូវនេះ
//   មានតែ Admin និងគ្រូនៃស្រុកនោះទេ ដែលមានសិទ្ធិចូល — ស្រុកដទៃទៀតគ្មានសិទ្ធិមើលទាល់តែសោះ
//   (សូមមើល syncSheetPermissions_ សម្រាប់ការចែករំលែកសិទ្ធិ)
function doGet() {
  // ==== FIX៖ setupAutoSheets() ត្រូវការពារកុំឲ្យបោះ Error ធ្វើឲ្យទំព័រទាំងមូលមិនអាចបើកបានសម្រាប់អ្នកប្រើប្រាស់
  // ទាំងអស់ (រួមទាំង Admin) — getSS_() ខាងលើឥឡូវមានផ្លូវសម្រួលដោយខ្លួនឯង (Fallback) រួចហើយ ប៉ុន្តែកំហុសផ្សេងទៀត
  // ដែលមិនបានគិតដល់ (ឧ. Drive Quota ពេលនោះ) មិនគួរធ្វើឲ្យទំព័រ Login ក៏មិនអាចបង្ហាញបានដែរ ====
  try { setupAutoSheets(); } catch (errSetup) {}
  // ==== FIX (Password Hashing)៖ ប្តូរលេខសម្ងាត់ចាស់ៗ (Plaintext) ទៅជា Hash ដោយស្វ័យប្រវត្តិ តែម្តងគត់ (មាន Flag
  // ការពារកុំឲ្យស្កេន Sheet ឡើងវិញរាល់ការស្នើសុំ — មើល migratePlaintextPasswordsIfNeeded_ ក្នុង Auth.gs) ====
  try { migratePlaintextPasswordsIfNeeded_(); } catch (errPwMig) {}
  try {
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty(VISIT_COUNTER_PROP_)) || 0;
    props.setProperty(VISIT_COUNTER_PROP_, String(n + 1));
  } catch (err) {}
  var tmpl = HtmlService.createTemplateFromFile('Index');
  return tmpl.evaluate()
      .setTitle('ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function newId_() {
  return Utilities.getUuid();
}
// ==== FIX (សុវត្ថិភាព — Formula/CSV Injection)៖ Google Sheets (ដូចគ្នានឹង Excel) បកប្រែតម្លៃណាមួយក្នុងក្រឡា ដែល
// ចាប់ផ្តើមដោយសញ្ញា =, +, -, @ ថាជារូបមន្ត (Formula) ដោយស្វ័យប្រវត្តិ។ អ្នកប្រើប្រាស់ (តួនាទីណាក៏ដោយ) អាចវាយបញ្ចូល
// អត្ថបទសេរី (ឧ. ឈ្មោះពេញ/មតិយោបល់/ចំណាំ) ចូល App ដោយត្រង់ៗ ដែលចុងក្រោយត្រូវបានរក្សាទុកក្នុងក្រឡា Google Sheet —
// បើតម្លៃនោះចាប់ផ្តើមដោយសញ្ញាទាំងនេះ ហើយ SuperAdmin/Admin/PEC21 បើកមើល Google Sheet ដោយផ្ទាល់ (ជំនួសការមើលតាម
// App ដែលមាន cmEsc_() ការពារ XSS រួចស្រាប់) រូបមន្តនោះអាចដំណើរការភ្លាមៗដោយស្វ័យប្រវត្តិ (ឧ. =IMPORTXML(...)/
// =HYPERLINK(...) ទាញយក/លាក់បាំងទិន្នន័យ) ។ ដោះស្រាយដោយបន្ថែម Apostrophe (') ខាងមុខ ដើម្បីបង្ខំ Google Sheets
// ចាត់ទុកតម្លៃនោះជា Text ធម្មតាជានិច្ច (មិនប៉ះពាល់តម្លៃដែលបង្ហាញ ឬតម្លៃដែល App ខ្លួនឯងអាន/ត្រួតពិនិត្យក្រោយមកទេ
// ព្រោះ getValues()/getDisplayValues() ទាំងអស់នៅតែត្រឡប់តម្លៃដើម ដោយគ្មាន Apostrophe នេះមកវិញ) ====
function sanitizeForSheetCell_(v) {
  var s = String(v === null || v === undefined ? "" : v);
  return /^[=+\-@]/.test(s) ? ("'" + s) : s;
}
function formatNow_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
  return Utilities.formatDate(new Date(), tz, 'dd-MM-yyyy HH:mm:ss');
}
// ធ្វើឲ្យតម្លៃកាលបរិច្ឆេទណាមួយ (Date object ដែល Google Sheet អាចនឹងបម្លែងស្វ័យប្រវត្តិ ឬ string ធម្មតា)
// ក្លាយទៅជា string ទម្រង់ "YYYY-MM-DD" ឲ្យដូចគ្នានឹងតម្លៃពី <input type="date"> ជានិច្ច
// (ត្រូវការសម្រាប់ប្រៀបធៀប/ស្វែងរកតាមកាលបរិច្ឆេទ ឲ្យបានត្រឹមត្រូវ ទោះបីជា Sheet រក្សាទុកជា Date object ក៏ដោយ)
function normalizeDateStr_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // ករណីទម្រង់ dd/MM/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return s;
}

// ==================== ២. រចនាសម្ព័ន្ធជួរឈរលម្អិត សម្រាប់ "របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ" ====================
// គម្រូតាម Google Sheets ដើម (ក្បាលជួរឈរ ក្រុមធំ/ក្រុមរង/សរុប-ស្រី)
var DAILY_FIELDS = [
  { key: "addNew_total",         col: "H",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ចុះឈ្មោះថ្មី",                    unit: "សរុប", type: "number" },
  { key: "addNew_female",        col: "I",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ចុះឈ្មោះថ្មី",                    unit: "ស្រី", type: "number" },
  { key: "addTransferIn_total",  col: "J",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ផ្ទេរចូល",                        unit: "សរុប", type: "number" },
  { key: "addTransferIn_female", col: "K",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ផ្ទេរចូល",                        unit: "ស្រី", type: "number" },
  { key: "addChangeStn_total",   col: "L",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ប្តូរការិ.ក្នុងឃុំសង្កាត់",       unit: "សរុប", type: "number" },
  { key: "addChangeStn_female",  col: "M",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "ប្តូរការិ.ក្នុងឃុំសង្កាត់",       unit: "ស្រី", type: "number" },
  { key: "addTotal_total",       col: "N",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "សរុប",                             unit: "សរុប", type: "number" },
  { key: "addTotal_female",      col: "O",  group: "ចំនួនអ្នកចុះឈ្មោះបន្ថែម", sub: "សរុប",                             unit: "ស្រី", type: "number" },

  { key: "delDeath_total",         col: "P", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ស្លាប់ ដកសិទ្ធិ ផ្លាស់ចេញ",       unit: "សរុប", type: "number" },
  { key: "delDeath_female",        col: "Q", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ស្លាប់ ដកសិទ្ធិ ផ្លាស់ចេញ",       unit: "ស្រី", type: "number" },
  { key: "delTransferOut_total",   col: "R", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ផ្ទេរចេញក្រៅឃុំ សង្កាត់",         unit: "សរុប", type: "number" },
  { key: "delTransferOut_female",  col: "S", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ផ្ទេរចេញក្រៅឃុំ សង្កាត់",         unit: "ស្រី", type: "number" },
  { key: "delDuplicate_total",     col: "T", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ស្ទួន",                          unit: "សរុប", type: "number" },
  { key: "delDuplicate_female",    col: "U", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ស្ទួន",                          unit: "ស្រី", type: "number" },
  { key: "delChangeStnOut_total",  col: "V", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ប្តូរការិ.ចេញក្នុងឃុំ សង្កាត់",   unit: "សរុប", type: "number" },
  { key: "delChangeStnOut_female", col: "W", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "ប្តូរការិ.ចេញក្នុងឃុំ សង្កាត់",   unit: "ស្រី", type: "number" },
  { key: "delTotal_total",         col: "X", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "លុបសរុប",                        unit: "សរុប", type: "number" },
  { key: "delTotal_female",        col: "Y", group: "ចំនួនលុបចេញពីបញ្ជីបោះឆ្នោត", sub: "លុបសរុប",                        unit: "ស្រី", type: "number" },

  { key: "update2026_total",  col: "Z",  group: "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦", sub: "", unit: "សរុប", type: "number" },
  { key: "update2026_female", col: "AA", group: "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦", sub: "", unit: "ស្រី", type: "number" },

  { key: "correction_total",  col: "AB", group: "ចំនួនករណីកែតម្រូវសរុប", sub: "", unit: "សរុប", type: "number" },
  { key: "correction_female", col: "AC", group: "ចំនួនករណីកែតម្រូវសរុប", sub: "", unit: "ស្រី", type: "number" },

  { key: "biometric_total",  col: "AD", group: "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ", sub: "", unit: "សរុប", type: "number" },
  { key: "biometric_female", col: "AE", group: "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ", sub: "", unit: "ស្រី", type: "number" },

  // ==== ជួរឈរ "សរុប" = ចំនួនករណីកែតម្រូវសរុប + ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ (រូបមន្តគណនាស្វ័យប្រវត្តិផ្ទាល់ក្នុង
  // Google Sheet ខ្លួនឯង — មើល applyPeriodSheetFormulasAndFormatting_ ក្នុង PeriodSheets.gs — មិនចាំបាច់សរសេរកូដ
  // បន្ថែមសម្រាប់រក្សាទុក/អានតម្លៃនេះឡើយ ដូចគ្នានឹង update2026_total ។ល។)។
  // ==== ប្រវត្តិ (សម្រាប់ជាឯកសារយោង)៖ ធាតុនេះកន្លងមកត្រូវបានដាក់ "នៅចុងគេ" ដោយចេតនា (មិនមែននៅចន្លោះកណ្តាលដូច
  // ទីតាំងបង្ហាញជាក់ស្តែងទេ) ព្រោះការបញ្ចូលនៅចន្លោះកណ្តាល នឹងកម្លាស់លេខសន្ទស្សន៍ជួរឈរជាក់ស្តែងនៃគ្រប់ធាតុទាំងអស់ដែល
  // មកក្រោយវា — ធ្វើឲ្យ Period Sheet ដែលមានទិន្នន័យស្រាប់ ត្រូវលែងផ្គូផ្គងជួរឈរត្រឹមត្រូវទៀតហើយ (Class បញ្ហាដូចគ្នា
  // នឹងកំហុសបាត់បង់ទិន្នន័យ fix90/91)។ ពេលនោះ getReportDisplayFields_() ខាងក្រោម ត្រូវធ្វើតួនាទីជា "ស្រទាប់បង្ហាញ"
  // ដាច់ដោយឡែក ដើម្បីតម្រៀបលំដាប់តែសម្រាប់បង្ហាញ (ដាក់ជាប់នឹង "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ") ដោយមិនប៉ះពាល់លំដាប់
  // ពិតក្នុង Sheet ឡើយ។ ==== FIX (ស្នើសុំដោយ SuperAdmin — កូដលំដាប់ថ្មីនេះ)៖ ដោយសារ SuperAdmin នឹង Reset ទិន្នន័យទាំង
  // អស់ (Google Sheet ស្រុកគ្រូប្រចាំក្រុងស្រុកទាំង១០ ថ្មីទាំងស្រុង) ភ្លាមៗបន្ទាប់ពីទទួលបានកំណែនេះ ហានិភ័យនៃការបំបែក
  // ជួរឈរទិន្នន័យចាស់ដែលធ្លាប់ព្រួយបារម្ភពីមុនលែងពាក់ព័ន្ធទៀតហើយ (គ្មានទិន្នន័យចាស់ណាមួយត្រូវរក្សាទុកបន្តទៀតទេ) —
  // ដូច្នេះឥឡូវផ្លាស់ទីធាតុនេះមកដាក់ត្រង់ទីតាំងពិតដែលចង់បាន (ជាប់នឹង "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ" ផ្ទាល់) ដើម្បីឲ្យ
  // Period Sheet ស្រុកនីមួយៗ (មិនមែនត្រឹមតែ Sheet សរុប៣ខាងក្រោមទេ) បង្ហាញលំដាប់ត្រឹមត្រូវដូចគ្នាដោយផ្ទាល់ តាំងពីកម្រិត
  // ទិន្នន័យ ដោយមិនចាំបាច់ពឹងផ្អែកលើ getReportDisplayFields_() ជា "ស្រទាប់បង្ហាញ" ដាច់ដោយឡែកទៀតឡើយ (ទោះ Function
  // នោះនៅតែរក្សាទុកជា Idempotent Safety Net — មើលមូលហេតុលម្អិតនៅខាងក្រោម)។ សម្រាប់ Period Sheet ចាស់ៗ ដែលអាចនៅតែ
  // ត្រូវប៉ះពាល់មុននឹង Reset ចប់ (ឧ. មានគេបញ្ចូលទិន្នន័យថ្ងៃនេះមុន SuperAdmin ចុច Reset) — មើល
  // ensurePeriodSheetHasTotalColumns_() ក្នុង PeriodSheets.gs ដែលបានកែសម្រួលឲ្យសុវត្ថិភាព (ប្រើ moveColumns() ផ្លាស់ទី
  // ជួរឈរនេះពីទីតាំងចាស់ (ចុងគេ) មករកទីតាំងថ្មីនេះ ដោយរក្សាទុកទិន្នន័យ/រូបមន្តទាំងអស់) ====
  { key: "correctionBiometricTotal_total",  col: "AF", group: "សរុប", sub: "", unit: "សរុប", type: "number" },
  { key: "correctionBiometricTotal_female", col: "AG", group: "សរុប", sub: "", unit: "ស្រី", type: "number" },

  { key: "idNew_total",          col: "AH", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "ចុះឈ្មោះថ្មី", unit: "សរុប", type: "number" },
  { key: "idNew_female",         col: "AI", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "ចុះឈ្មោះថ្មី", unit: "ស្រី", type: "number" },
  { key: "idTransferIn_total",   col: "AJ", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "ផ្ទេរចូល",     unit: "សរុប", type: "number" },
  { key: "idTransferIn_female",  col: "AK", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "ផ្ទេរចូល",     unit: "ស្រី", type: "number" },
  { key: "idTotal_total",        col: "AL", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "សរុប",         unit: "សរុប", type: "number" },
  { key: "idTotal_female",       col: "AM", group: "ចំនួន (ឯ.អ) បានចេញ", sub: "សរុប",         unit: "ស្រី", type: "number" },

  { key: "residenceCert_total",  col: "AN", group: "លិខិតបញ្ជាក់ទីលំនៅ", sub: "", unit: "សរុប", type: "number" },
  { key: "residenceCert_female", col: "AO", group: "លិខិតបញ្ជាក់ទីលំនៅ", sub: "", unit: "ស្រី", type: "number" }
];

// ==== ជួរឈរដែលជា "លទ្ធផលគណនា" ស្វ័យប្រវត្តិសុទ្ធសាធ (មានរូបមន្តក្នុង Google Sheet ផ្ទាល់) ====
// ==== FIX៖ ដកចេញ correctionBiometricTotal_total/_female ពី Map នេះ (ធ្វើឲ្យទទេ) — មុននេះប្រើសម្រាប់ដកជួរឈរនេះ
// ចេញពីទម្រង់បញ្ចូលទិន្នន័យទាំងស្រុង ប៉ុន្តែឥឡូវផ្លាស់ប្តូរតាមគំរូដូចគ្នានឹងជួរឈរគណនាស្វ័យប្រវត្តិដទៃទៀត (addTotal_total/
// delTotal_total/idTotal_total/update2026_total) — គឺបង្ហាញជាធម្មតាក្នុងទម្រង់បញ្ចូលទិន្នន័យ (មិនដកចេញទៀតទេ) ប៉ុន្តែជា
// ប្រអប់ "សម្រាប់មើលតែប៉ុណ្ណោះ" (readOnly) គណនាស្វ័យប្រវត្តិផ្ទាល់ក្នុង Browser (មើល AUTO_SUM_MAP ក្នុង Index.html) —
// ដូច្នេះគ្រូនឹងឃើញតម្លៃសរុបភ្លាមៗ ខណៈកំពុងបញ្ចូលទិន្នន័យ ដោយមិនចាំបាច់រង់ចាំរក្សាទុកសិន (ស្របតាមសំណើ)។ Map នេះ
// រក្សាទុកទទេ (មិនលុបចោលទាំងស្រុង) ដើម្បីរក្សា Interface ដដែល ក្នុងករណីត្រូវការដកជួរឈរណាមួយផ្សេងទៀតនាពេលអនាគត ====
var COMPUTED_FIELD_KEYS_ = {};

// ==== ត្រឡប់ច្បាប់ចម្លង DAILY_FIELDS "សម្រាប់បង្ហាញ" — ==== FIX៖ ចាប់ពីពេលដែលជួរឈរ "សរុប"
// (correctionBiometricTotal_total/_female) ត្រូវបានផ្លាស់ទីទីតាំងពិតប្រាកដក្នុង DAILY_FIELDS ខាងលើរួច (ជាប់នឹង
// "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ" ស្រាប់) Function នេះក្លាយជា Idempotent (ស្មើនឹង DAILY_FIELDS ជាក់ស្តែងរួចទៅហើយ —
// ការដក/បញ្ចូលឡើងវិញខាងក្រោម មិនផ្លាស់ប្តូរអ្វីឡើយ) ។ រក្សាទុក Function នេះជា "Safety Net" ដដែល (មិនលុបចោល) ដើម្បី
// ធានាថា renderDailyTable/renderCommuneSumTable/renderProvinceSumTable/renderPrintReport (Index.html) និង
// exportReportToExcel (Reports.gs) នៅតែបង្ហាញលំដាប់ត្រឹមត្រូវ ទោះ DAILY_FIELDS ខ្លួនឯងប្រែប្រួលយ៉ាងណាក៏ដោយនាពេលអនាគត ====
function getReportDisplayFields_() {
  var totalField = null, femaleField = null;
  for (var j = 0; j < DAILY_FIELDS.length; j++) {
    if (DAILY_FIELDS[j].key === "correctionBiometricTotal_total") totalField = DAILY_FIELDS[j];
    else if (DAILY_FIELDS[j].key === "correctionBiometricTotal_female") femaleField = DAILY_FIELDS[j];
  }
  var out = [];
  for (var i = 0; i < DAILY_FIELDS.length; i++) {
    var f = DAILY_FIELDS[i];
    if (f === totalField || f === femaleField) continue; // បញ្ចូលឡើងវិញនៅទីតាំងត្រឹមត្រូវខាងក្រោម មិនមែនទីតាំងដើមទេ
    out.push(f);
    if (f.key === "biometric_female" && totalField && femaleField) { out.push(totalField, femaleField); }
  }
  return out;
}

// ជួរឈរមូលដ្ឋាន មុន និង ក្រោយ ក្រុមវាលលម្អិត នៅក្នុង Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ)
var DAILY_META_PREFIX = ["ID", "កាលបរិច្ឆេទ", "លេខកូដ", "ស្រុក", "ឃុំ/សង្កាត់"];
// ==== FIX (សំណើថ្មី "កត់ត្រាសកម្មភាពគ្រប់ការបញ្ចូល/កែប្រែទិន្នន័យ" — ចំណុចទី៥, កំណែកែសម្រួលតាមសំណើកែតម្រូវផ្ទាល់
// របស់អ្នកប្រើប្រាស់)៖ អ្នកប្រើប្រាស់មិនចង់បាន Sheet កំណត់ត្រាសវនកម្មដាច់ដោយឡែកទេ (សូមមើលការពន្យល់ពេញលេញនៅ
// ensurePeriodSheetHasEditCountColumn_() ខាងក្រោមក្នុង PeriodSheets.gs) — ចង់បានត្រឹមតែជួរឈរបន្ថែម ១ ថ្មី
// ("ចំនួនដងកែប្រែ") ដាក់បញ្ចូលដោយផ្ទាល់នៅចុងគេនៃតារាងទិន្នន័យប្រចាំថ្ងៃដែលមានស្រាប់រួចហើយ (ដូចគ្នានឹង "ចំណាំ"/
// "អ្នកបញ្ចូល"/"ពេលបញ្ចូល" ដែលមានរួចស្រាប់) ។ ដោយសារ DAILY_META_SUFFIX នេះ ជាក្រុមចុងក្រោយគេនៃលំដាប់ជួរឈរទាំងមូល
// (មើល periodSheetTotalCols_/periodFieldColIndex_ ខាងក្រោម) ការបន្ថែមធាតុថ្មីនៅចុង Array នេះ មានន័យថា ជួរឈរថ្មី
// នឹងស្ថិតនៅចុងគេបំផុតនៃគ្រប់ Period Sheet ទាំងអស់ជានិច្ច — មិនដែលរុញ/ផ្លាស់ទីជួរឈរផ្សេងទៀតណាមួយឡើយ (ខុសពី
// ការផ្លាស់ទីជួរឈរ "សរុប" ពីចុងគេទៅកណ្តាល កាលពីមុន ដែលចាំបាច់ត្រូវការ moveColumns/insertColumnsBefore ស្មុគស្មាញ)
// — សុវត្ថិភាពជាង ងាយស្រួលជាង ក្នុងការ Migrate Sheet ចាស់ៗ ====
var DAILY_META_SUFFIX = ["ចំណាំ", "អ្នកបញ្ចូល", "ពេលបញ្ចូល", "ចំនួនដងកែប្រែ"];

// ==================== ៣. បង្កើត Sheet ស្វ័យប្រវត្តិ ====================
// កំណត់ចំណាំ៖ លែងបង្កើត Tab ស្រុកនីមួយៗ (ឈ្មោះតាមស្រុក) ដោយស្វ័យប្រវត្តិទៀតហើយ។
// Spreadsheet ដាច់ដោយឡែករបស់ស្រុកនីមួយៗ នៅតែបង្កើត (លើកដំបូង) ដដែល ប៉ុន្តែ Tab ខាងក្នុង
// នឹងជា Period Sheet ឈ្មោះតាមថ្ងៃខែ (ឧ. "25-08") ដែលបង្កើតតាមតម្រូវការ ពេលមានការបញ្ចូលទិន្នន័យ (មើល getOrCreatePeriodSheet_)។
function setupAutoSheets() {
  var ss = getSS_();
  ensureUsersSheet_(ss);
  ensureCommuneListSheet_(ss);
  ensureDailySheet_(ss);
  ensureCommuneSumSheet_(ss);
  ensureProvinceSumSheet_(ss);
  var dailyIndexSheet_ = ensureDailyIndexSheet_(ss);
  // ==== ល្បឿន៖ កក់ទុកសន្ទស្សន៍ជាមុនភ្លាមៗ បើមិនទាន់ធ្លាប់មានទិន្នន័យសោះ (ដំឡើងកូដថ្មីលើកដំបូង ឬក្រោយ Reset)
  // ដើម្បីកុំឲ្យអ្នកប្រើដំបូងគេ ត្រូវរង់ចាំយូរ (ការស្កេនពេញលេញលើកដំបូង) ពេលចុច "🔄 ធ្វើបច្ចុប្បន្នភាពតារាង" — សន្ទស្សន៍
  // នឹងសាងសង់ខាងក្រោយឆាកក្នុងរយៈពេលប៉ុន្មានវិនាទី ដោយមិនធ្វើឲ្យអ្នកប្រើត្រូវរង់ចាំ (មិនប៉ះពាល់ល្បឿនបើកទំព័រទេ) ====
  try {
    if (dailyIndexSheet_.getLastRow() <= 1) scheduleRebuild_();
  } catch (err) {}
  ensureDailyProtectionTrigger_(); // ធានាថា Trigger ប្រចាំថ្ងៃសម្រាប់ចាក់សោ Tab ថ្ងៃចាស់ ត្រូវបានដំឡើង (ធ្វើតែម្តងគត់)

  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getNumSheets() > 1) {
    ss.deleteSheet(defaultSheet);
  }
}

// ==== សន្ទស្សន៍ខាងក្នុង (កុំកែដោយដៃ) — ១ជួរដេក = ១ស្រុក+១ថ្ងៃ (JSON នៃទិន្នន័យគ្រប់ឃុំ/សង្កាត់ថ្ងៃនោះ) ====
// ត្រូវបានសរសេរឡើងវិញទាំងស្រុងរាល់ពេល rebuildAllDerivedData_() រត់ (បន្ទាប់ពីរក្សាទុក/លុប/កែផ្ទាល់ក្នុង Sheet ណាមួយ)
// ដូច្នេះមិនចាំបាច់ព្រួយបារម្ភរឿងទិន្នន័យចាស់ជាប់គាំងឡើយ — គ្រាន់តែយឺតជាង ១-២ វិនាទី (Trigger បណ្តោះអាសន្ន) ប៉ុណ្ណោះ
function ensureDailyIndexSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_DAILY_INDEX);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DAILY_INDEX);
    sheet.appendRow(["ស្រុក", "កាលបរិច្ឆេទ", "ទិន្នន័យ (JSON)", "ធ្វើបច្ចុប្បន្នភាពចុងក្រោយ"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f59e0b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.getRange(1, 2, 20000, 1).setNumberFormat('@'); // ជួរឈរកាលបរិច្ឆេទ រក្សាទុកជា Text ជានិច្ច (កុំឲ្យ Sheets បម្លែងទៅ Date)
    try { sheet.hideSheet(); } catch (err) {}
  }
  return sheet;
}

// ---- Sheet(អ្នកប្រើប្រាស់) ----
function ensureCommuneListSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_COMMUNE_LIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COMMUNE_LIST);
    var headers = ["ស្រុក/ក្រុង", "ឃុំ/សង្កាត់"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f59e0b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 2, 220);

    // បំពេញឈ្មោះឃុំ/សង្កាត់ពិតប្រាកដទាំងអស់ (ពី COMMUNE_ORDER) ជាមុនតែម្តង — មិនចាំបាច់ឲ្យ Admin វាយបញ្ចូលដោយដៃម្តងមួយៗទៀតទេ
    var sampleRows = [];
    DISTRICT_LIST.forEach(function(d) {
      (COMMUNE_ORDER[d] || []).forEach(function(c) { sampleRows.push([d, c.name]); });
    });
    if (sampleRows.length) sheet.getRange(2, 1, sampleRows.length, 2).setValues(sampleRows);

    // Dropdown សម្រាប់ជួរឈរ "ស្រុក/ក្រុង" ដើម្បីជៀសវាងការវាយបញ្ចូលខុស
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(DISTRICT_LIST, true).setAllowInvalid(false).build();
    sheet.getRange(2, 1, 500, 1).setDataValidation(rule);
  }
  return sheet;
}

// កំណត់ Sheet(ឃុំ សង្កាត់) ឡើងវិញទាំងស្រុង ដោយឈ្មោះឃុំ/សង្កាត់ត្រឹមត្រូវ (ពី COMMUNE_ORDER) — Admin ប៉ុណ្ណោះ
// ==== សំខាន់៖ លុបចោលទិន្នន័យចាស់ទាំងអស់ក្នុង Sheet នេះ ជំនួសដោយបញ្ជីត្រឹមត្រូវ — ប្រើនៅពេលទិន្នន័យក្នុង Sheet មានកំហុស/មិនពេញលេញ ====

// ==== ល្បឿន៖ លទ្ធផលត្រូវបាន Cache ទុក ១០នាទី ព្រោះបញ្ជីនេះកម្រប្តូរណាស់ (Cache ត្រូវលុបចោលស្វ័យប្រវត្តិ
// នៅពេលមានការកែប្រែ — មើល clearLookupCache_) ។ នេះកាត់បន្ថយការអាន Sheet ដដែលៗរាល់ពេលបើកទំព័រ។
var CACHE_KEY_COMMUNE_MAP_ = 'communeMap_v1';
var CACHE_KEY_BASELINE_MAP_ = 'baselineMap_v1';
var CACHE_TTL_ = 600; // វិនាទី (១០ នាទី)

function clearLookupCache_() {
  try { CacheService.getScriptCache().removeAll([CACHE_KEY_COMMUNE_MAP_, CACHE_KEY_BASELINE_MAP_]); } catch (err) {}
}

function getCommuneMap_() {
  try {
    var cached = CacheService.getScriptCache().get(CACHE_KEY_COMMUNE_MAP_);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var ss = getSS_();
  var sheet = ensureCommuneListSheet_(ss);
  var map = {};
  DISTRICT_LIST.forEach(function(d) { map[d] = []; });
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
    data.forEach(function(row) {
      var d = String(row[0] || "").trim();
      var c = String(row[1] || "").trim();
      if (!d || !c) return;
      if (!map[d]) map[d] = [];
      if (map[d].indexOf(c) === -1) map[d].push(c);
    });
  }

  try { CacheService.getScriptCache().put(CACHE_KEY_COMMUNE_MAP_, JSON.stringify(map), CACHE_TTL_); } catch (err) {}
  return map;
}

// ==================== ល្បឿន៖ Cache កាលបរិច្ឆេទពិតប្រាកដ (ជួរឈរ B) របស់ Period Sheet នីមួយៗ ក្នុងស្រុកមួយ ====================
// មុននេះ រាល់ពេលត្រូវការដឹងថា Tab ថ្ងៃខែមួយ ជាកាលបរិច្ឆេទណាពិតប្រាកដ (ឈ្មោះ Tab មានតែ "dd-MM" គ្មានឆ្នាំទេ)
// ត្រូវអាន sheet.getRange(4,2).getValue() ម្តងមួយៗ សម្រាប់ Tab ជារាប់សិប (ហៅញឹកញាប់ណាស់ — getRollingReferenceMap_,
// protectPastPeriodSheets_) ។ Map នេះកម្រប្តូរណាស់ (ប្តូរតែពេលមាន Tab ថ្ងៃថ្មីត្រូវបានបង្កើត) ដូច្នេះ Cache ទុក
// រហូតដល់ ៦ម៉ោង (អតិបរមាដែល CacheService អនុញ្ញាត) ហើយលុបចោលភ្លាមៗ ពេលមាន Tab ថ្មីត្រូវបានបង្កើត (createDistrictPeriodSheet_)។
var CACHE_KEY_PERIOD_DATES_PREFIX_ = 'periodDates_v1_';
var CACHE_TTL_PERIOD_DATES_ = 21600; // វិនាទី (៦ ម៉ោង — អតិបរមា)

function getPeriodSheetDateMap_(district) {
  var cache = CacheService.getScriptCache();
  var key = CACHE_KEY_PERIOD_DATES_PREFIX_ + district;
  try {
    var cached = cache.get(key);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var map = {};
  listDistrictPeriodSheets_(district).forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return;
    var d = normalizeDateStr_(sheet.getRange(4, 2).getValue());
    if (d) map[sheet.getName()] = d;
  });
  try { cache.put(key, JSON.stringify(map), CACHE_TTL_PERIOD_DATES_); } catch (err) {}
  return map;
}

function clearPeriodSheetDateCache_(district) {
  try { CacheService.getScriptCache().remove(CACHE_KEY_PERIOD_DATES_PREFIX_ + district); } catch (err) {}
}

// ==================== ល្បឿន៖ កំណែ (Version) Cache សម្រាប់លទ្ធផលរបាយការណ៍សរុប (computeSummaries_ ក្នុង Reports.gs) ====================
// computeSummaries_ ជាការគណនាធ្ងន់បំផុតមួយ (ស្កេន Period Sheet គ្រប់ស្រុកទាំង១០) ដូច្នេះ Cache លទ្ធផលទុកតាម
// ចន្លោះកាលបរិច្ឆេទ+ជម្រើសដែលបានស្នើសុំ។ ដើម្បីកុំឲ្យអ្នកប្រើឃើញរបាយការណ៍ចាស់ភ្លាមៗក្រោយពេលបញ្ចូល/កែ/លុបទិន្នន័យ
// ត្រូវ "តម្កើងកំណែ" (bump) រាល់ពេលមានការកែប្រែទិន្នន័យ (មើល scheduleRebuild_) — កំណែផ្លាស់ប្តូរ = Cache ចាស់លែងត្រូវប្រើ
// ដោយស្វ័យប្រវត្តិ (មិនចាំបាច់ដឹងជាមុនថា Key ណាខ្លះធ្លាប់ត្រូវ Cache ទុកទេ)។
var REPORTS_CACHE_VERSION_KEY_ = 'reportsCacheVer_v1';
var REPORTS_CACHE_TTL_ = 60; // វិនាទី (១ នាទី — ដូចគ្នានឹង Dashboard)

function getReportsCacheVersion_() {
  try {
    var v = CacheService.getScriptCache().get(REPORTS_CACHE_VERSION_KEY_);
    if (v) return v;
  } catch (err) {}
  return '0';
}

function bumpReportsCacheVersion_() {
  try {
    var v = (Number(getReportsCacheVersion_()) || 0) + 1;
    CacheService.getScriptCache().put(REPORTS_CACHE_VERSION_KEY_, String(v), CACHE_TTL_PERIOD_DATES_);
  } catch (err) {}
}

// ==================== ០.២ ទិន្នន័យមូលដ្ឋានឆ្នាំ២០២៥ + ការកំណត់ប្រព័ន្ធ (ថ្ងៃចាប់ផ្តើមដំណើរការ) ====================
// Sheet(ទិន្នន័យមូលដ្ឋានឆ្នាំ២០២៥) — មួយជួរ = ១ ឃុំ/សង្កាត់ (គ្រប់ឃុំសង្កាត់ទាំងអស់ បំពេញជាមុនតាម COMMUNE_ORDER)
// ប្រើសម្រាប់ជាមូលដ្ឋានគណនា "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦" និងបង្ហាញក្នុងជួរឈរ "យោង" នៃរបាយការណ៍បូកយោង
var SHEET_LINKED_DATA = "ទិន្នន័យភ្ជាប់";
// ==== v2៖ បន្ថែមជួរឈរ "ស្ថិតិប៉ាន់ស្មានខេត្ត" (ជួរឈរទី១១ ចុងក្រោយគេ) ដាច់ដោយឡែកពី "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" (ជួរឈរទី៣ ចាស់ —
// ប្តូរឈ្មោះបង្ហាញប៉ុណ្ណោះ មិនប្តូរទីតាំងជួរឈរទេ) — បន្ថែមនៅ "ចុង" ដូចធម្មតា (មិនមែនកន្លែងបង្ហាញឲ្យប្រើក្នុង UI ជាប់គ្នាទេ ព្រោះ
// លំដាប់បង្ហាញ (ភ្ជាប់ស្ថិតិប៉ាន់ស្មានខេត្ត នៅជាប់ក្រោយ ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប) ត្រូវបានគ្រប់គ្រងដោយ Client JS ដាច់ដោយឡែក
// (មិនអាស្រ័យលើលំដាប់ជួរឈរពិតប្រាកដក្នុង Google Sheet ទេ) ====
var LINKED_DATA_TOTAL_COLS = 11; // ស្រុក/ក្រុង, ឃុំ/សង្កាត់, ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប, ការិ.បង្កើតថ្មី, ចំនួនការិ.សរុប, ២០២៥-សរុប, ២០២៥-ស្រី, លេខកូដ, ជីវមាត្រ-សរុប, ជីវមាត្រ-ស្រី, ស្ថិតិប៉ាន់ស្មានខេត្ត
// ==== ចំណាំ៖ ២ថេរនេះ ធ្លាប់គណនាជា "LINKED_DATA_TOTAL_COLS - 1"/"LINKED_DATA_TOTAL_COLS" (សន្មតថាជីវមាត្រជាជួរឈរ
// ចុងក្រោយគេ) — ឥឡូវត្រូវកំណត់ជាតួលេខថេរដាច់ដោយឡែក ព្រោះជីវមាត្រលែងជាជួរឈរចុងក្រោយគេទៀតហើយ (មានស្ថិតិប៉ាន់ស្មានខេត្ត
// បន្ថែមនៅចុងក្រោយបំផុតវិញ) — បើនៅគណនាដូចមុន នឹងចង្អុលទៅជួរឈរខុស (ជួរឈរ ១០/១១ ជំនួសជួរឈរ ៩/១០ ត្រឹមត្រូវ) ====
var LINKED_DATA_BIO_TOTAL_COL_ = 9;  // ១-indexed — "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)"
var LINKED_DATA_BIO_FEMALE_COL_ = 10; // ១-indexed — "ទិន្នន័យធ្វើជីវមាត្រ (ស្រី)"
var LINKED_DATA_PROVINCE_EST_COL_ = 11; // ១-indexed — "ស្ថិតិប៉ាន់ស្មានខេត្ត" (ជួរឈរថ្មី ចុងក្រោយគេ)

function ensureLinkedDataSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_LINKED_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LINKED_DATA);
    var headers = ["ស្រុក/ក្រុង", "ឃុំ/សង្កាត់", "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប", "ការិ.បង្កើតថ្មី", "ចំនួនការិ.សរុប", "ទិន្នន័យ២០២៥ (សរុប)", "ទិន្នន័យ២០២៥ (ស្រី)", "លេខកូដ", "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)", "ទិន្នន័យធ្វើជីវមាត្រ (ស្រី)", "ស្ថិតិប៉ាន់ស្មានខេត្ត"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);

    var codeMap = getCommuneCodeMap_();
    var rows = [];
    DISTRICT_LIST.forEach(function(d) {
      (COMMUNE_ORDER[d] || []).forEach(function(c) { rows.push([d, c.name, 0, 0, 0, 0, 0, codeMap[d + "|" + c.name] || "", 0, 0, 0]); });
    });
    if (rows.length) sheet.getRange(2, 1, rows.length, LINKED_DATA_TOTAL_COLS).setValues(rows);
    return sheet;
  }

  // ==== ប្តូរឈ្មោះជួរឈរទី៣ ចាស់ "ស្ថិតិប៉ាន់ស្មាន" → "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" (ប្តូរឈ្មោះបង្ហាញប៉ុណ្ណោះ — ទិន្នន័យ/ទីតាំង
  // ជួរឈរដដែល) ដើម្បីបែងចែកឲ្យច្បាស់ពី "ស្ថិតិប៉ាន់ស្មានខេត្ត" ថ្មី ====
  if (sheet.getRange(1, 3).getDisplayValue() === "ស្ថិតិប៉ាន់ស្មាន") {
    sheet.getRange(1, 3).setValue("ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប")
        .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
  }

  // ==== បន្ថែមជួរឈរថ្មី "ស្ថិតិប៉ាន់ស្មានខេត្ត" (ជួរឈរទី១១ ចុងក្រោយគេ) សម្រាប់ Sheet ចាស់ដែលមិនទាន់មាន ====
  if (sheet.getRange(1, LINKED_DATA_PROVINCE_EST_COL_).getDisplayValue() !== "ស្ថិតិប៉ាន់ស្មានខេត្ត") {
    sheet.getRange(1, LINKED_DATA_PROVINCE_EST_COL_).setValue("ស្ថិតិប៉ាន់ស្មានខេត្ត")
        .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    var lastRowPE = sheet.getLastRow();
    if (lastRowPE >= 2) {
      var zerosPE = [];
      for (var zpi = 0; zpi < lastRowPE - 1; zpi++) zerosPE.push([0]);
      sheet.getRange(2, LINKED_DATA_PROVINCE_EST_COL_, zerosPE.length, 1).setValues(zerosPE);
    }
  }

  // ==== ធ្វើបច្ចុប្បន្នភាព Sheet ចាស់ (បង្កើតមុនពេលមានជួរឈរ "លេខកូដ" និង/ឬ "ទិន្នន័យធ្វើជីវមាត្រ") ====
  // ចំណាំ៖ ជួរឈរថ្មីៗ ត្រូវបានបន្ថែមនៅ ចុងជួរឈរ (មិនមែនដើម) ដើម្បីកុំឲ្យប៉ះពាល់/ផ្លាស់ទីលេខ ជួរឈរចាស់ៗ ដែលមានទិន្នន័យស្រាប់
  // ត្រួតពិនិត្យតាមខ្លឹមសារ Header ជាក់ស្តែង (មិនមែនគ្រាន់តែចំនួនជួរឈរ) ដើម្បីធានាថាដំណើរការត្រឹមត្រូវ ទោះ Sheet ស្ថិតក្នុងជំហានណាមួយក៏ដោយ
  if (sheet.getRange(1, 8).getDisplayValue() !== "លេខកូដ") {
    sheet.getRange(1, 8).setValue("លេខកូដ")
        .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    var lastRowOld = sheet.getLastRow();
    if (lastRowOld >= 2) {
      var codeMapOld = getCommuneCodeMap_();
      var oldKeys = sheet.getRange(2, 1, lastRowOld - 1, 2).getValues();
      var codeCol = oldKeys.map(function(r) {
        return [codeMapOld[String(r[0]).trim() + "|" + String(r[1]).trim()] || ""];
      });
      sheet.getRange(2, 8, codeCol.length, 1).setValues(codeCol);
    }
  }

  if (sheet.getRange(1, LINKED_DATA_BIO_TOTAL_COL_).getDisplayValue() !== "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)") {
    sheet.getRange(1, LINKED_DATA_BIO_TOTAL_COL_).setValue("ទិន្នន័យធ្វើជីវមាត្រ (សរុប)")
        .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    sheet.getRange(1, LINKED_DATA_BIO_FEMALE_COL_).setValue("ទិន្នន័យធ្វើជីវមាត្រ (ស្រី)")
        .setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    var lastRowBio = sheet.getLastRow();
    if (lastRowBio >= 2) {
      var zeros = [];
      for (var zi = 0; zi < lastRowBio - 1; zi++) zeros.push([0, 0]);
      sheet.getRange(2, LINKED_DATA_BIO_TOTAL_COL_, zeros.length, 2).setValues(zeros);
    }
  }

  // ធានាថាមានជួរដេកគ្រប់ឃុំ/សង្កាត់ (ករណី COMMUNE_ORDER ត្រូវបានកែប្រែថ្មី)
  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(function(r) { existing[String(r[0]).trim() + "|" + String(r[1]).trim()] = true; });
  }
  var toAdd = [];
  var codeMap2 = getCommuneCodeMap_();
  DISTRICT_LIST.forEach(function(d) {
    (COMMUNE_ORDER[d] || []).forEach(function(c) {
      var key = d + "|" + c.name;
      if (!existing[key]) toAdd.push([d, c.name, 0, 0, 0, 0, 0, codeMap2[key] || "", 0, 0, 0]);
    });
  });
  if (toAdd.length) sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, LINKED_DATA_TOTAL_COLS).setValues(toAdd);
  return sheet;
}

function dailyTotalCols_() {
  return DAILY_META_PREFIX.length + DAILY_FIELDS.length + DAILY_META_SUFFIX.length;
}

// ==== ពិនិត្យថាតើលំដាប់ជួរឈរ Header បច្ចុប្បន្នរបស់ Tab សរុប៣ (Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ)/Sheet(របាយការណ៍សរុប
// តាមឃុំសង្កាត់)/Sheet(របាយការណ៍សរុបខេត្ត)) ស្របតាមលំដាប់ "សម្រាប់បង្ហាញ" ថ្មី (getReportDisplayFields_ — ជួរឈរ
// "សរុប"/correctionBiometricTotal ដាក់ជាប់នឹង "ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ" វិញ) ដែរឬអត់ — ពិនិត្យត្រឹមតែទីតាំង
// ជាក់ស្តែងនៃជួរឈរ "សរុប" នេះគត់ (គ្រប់គ្រាន់ ព្រោះនេះជាតែមួយគត់ដែលមានលំដាប់ខុសពី DAILY_FIELDS ដើម)។ prefixLen =
// ចំនួនជួរឈរខាងមុខ DAILY_FIELDS ជាក់ស្តែងរបស់ Tab នីមួយៗ (ខុសគ្នាទៅតាម Tab) ====
function derivedSheetOrderMatches_(sheet, prefixLen) {
  var displayFields = getReportDisplayFields_();
  var idx = -1;
  for (var i = 0; i < displayFields.length; i++) {
    if (displayFields[i].key === 'correctionBiometricTotal_total') { idx = i; break; }
  }
  if (idx === -1) return true; // មិនគួរកើតឡើងទេ ប៉ុន្តែកុំបង្ខំបង្កើតឡើងវិញដោយអសារឥតការ បើពិតជារកមិនឃើញ
  var col = prefixLen + idx + 1;
  try { return sheet.getRange(1, col).getDisplayValue() === 'សរុប'; } catch (err) { return false; }
}

// ---- Sheet(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) — Tab សរុបរួម ត្រូវបានសរសេរដោយ Script ស្វ័យប្រវត្តិ (Read-only, កុំកែដោយដៃ) ----
// ទិន្នន័យពិតប្រាកដ ត្រូវបញ្ចូល/កែប្រែ តាមរយៈ Period Sheet (ឈ្មោះថ្ងៃខែ ឧ. "25-08") ក្នុង Spreadsheet ស្រុកនីមួយៗ (មើល getOrCreatePeriodSheet_)
function ensureDailySheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_DAILY);
  var totalCols = dailyTotalCols_();
  if (sheet) {
    // ការពារទិន្នន័យចាស់៖ បើទម្រង់ក្បាលជួរឈរខុសពីទម្រង់ថ្មី (ជួរឈរតិចជាង ឬលំដាប់ជួរឈរ "សរុប" ចាស់) សូមប្តូរឈ្មោះ
    // Sheet ចាស់ជា Backup ហើយបង្កើតថ្មី
    if (sheet.getLastColumn() < totalCols || sheet.getRange(3, 1).getDisplayValue() !== "ID" ||
        !derivedSheetOrderMatches_(sheet, DAILY_META_PREFIX.length)) {
      var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
      var backupName = SHEET_DAILY + " (ចាស់ " + Utilities.formatDate(new Date(), tz, 'ddMMyyyy_HHmmss') + ")";
      sheet.setName(backupName);
      sheet = null;
    }
  }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DAILY);
    buildDailySheetHeader_(sheet, totalCols);
  }
  return sheet;
}

// ---- ១ ស្រុក = ១ Google Spreadsheet ដាច់ដោយឡែក (មិនមែន Tab រួមក្នុង Spreadsheet មេទៀតទេ) ----
// ការពារត្រឹមត្រូវ១០០%៖ គណនីគ្រូប្រចាំស្រុក ត្រូវបានចែករំលែក (Share) ជូននូវ Spreadsheet របស់ខ្លួនប៉ុណ្ណោះ
// គាត់មិនត្រូវបានចែក ឬមានសិទ្ធិចូលមើល Spreadsheet មេ ឬស្រុកដទៃទៀតឡើយ ដូច្នេះពេលចុចបើក Google Sheet
// លទ្ធផលមាន "តែ" ទិន្នន័យ/Tab របស់ស្រុកខ្លួនប៉ុណ្ណោះ — ក្រៅពីនោះមិនអាចមើលឃើញ ឬចូលដល់បានឡើយ
// (សូមមើល syncSheetPermissions_ សម្រាប់ការចែករំលែកសិទ្ធិ)
var DISTRICT_SS_PROP_PREFIX = "districtSS_"; // Key ស្តុក Spreadsheet ID ក្នុង Script Properties (គង់វង្សរហូត)

function districtSheetName_(district) {
  return district;
}

// ត្រឡប់ Folder ក្នុង Drive ដែលផ្ទុក Spreadsheet មេ (ដើម្បីដាក់ Spreadsheet ស្រុកនីមួយៗនៅកន្លែងតែមួយ ងាយស្វែងរក)
// ==== អាទិភាព៖ (១) តំណភ្ជាប់ Google Drive ដែល Admin បិទភ្ជាប់ (Copy-Paste) ក្នុង Tab "ការកំណត់ប្រព័ន្ធ" (masterDriveLink)
// (២) Folder ជាក់ស្តែងបច្ចុប្បន្នរបស់ Spreadsheet មេ ដែលកំពុងប្រើប្រាស់ (getTargetSheetId_) ====
// ==== FIX៖ ប្តូរអាទិភាព តាមសំណើ — Google Sheet ស្រុករបស់គ្រូប្រចាំក្រុងស្រុក ត្រូវបង្កើតនៅក្នុង Google Drive
// ដដែលនឹង Google Sheet មេកំពុងតាំងនៅ "តែម្តង" ដោយស្វ័យប្រវត្តិ (មិនតម្រូវឲ្យកំណត់ Folder ដាច់ដោយឡែក ដែលងាយប្រឈម
// នឹងបញ្ហាសិទ្ធិចូលប្រើ Folder នោះ ដូចករណីមុន)។ អាទិភាពថ្មី៖
// (១) Folder ជាក់ស្តែងបច្ចុប្បន្នរបស់ Spreadsheet មេ ដែលកំពុងប្រើប្រាស់ (getTargetSheetId_) — ជាប្រភពត្រឹមត្រូវបំផុត
//     ព្រោះ Admin ត្រូវតែចូល Sheet មេនេះបានស្រាប់ (ដើម្បីកែទិន្នន័យ) ដូច្នេះ Folder របស់វា ក៏ត្រូវតែចូលបានដូចគ្នា
// (២) តំណភ្ជាប់ Google Drive ដែល Admin បិទភ្ជាប់ (Copy-Paste) ក្នុង Tab "ការកំណត់ប្រព័ន្ធ" (masterDriveLink) — ប្រើ
//     ជាជម្រើសបម្រុងតែក្នុងករណី Sheet មេ គ្មាន Folder មេជាក់ស្តែង (ឧ. ស្ថិតនៅ My Drive ដោយផ្ទាល់) ====
function getMasterParentFolder_() {
  try {
    var it = DriveApp.getFileById(getTargetSheetId_()).getParents();
    if (it.hasNext()) return it.next();
  } catch (err) {}
  try {
    var settings = getSystemSettings_();
    var customFolderId = extractDriveFolderId_(settings.masterDriveLink);
    if (customFolderId) {
      try { return DriveApp.getFolderById(customFolderId); } catch (err) {}
    }
  } catch (err) {}
  return null;
}

// ត្រឡប់ (ហើយបង្កើតលើកដំបូងបើមិនទាន់មាន) Spreadsheet ដាច់ដោយឡែក សម្រាប់ស្រុកមួយ
// ==== ការកាចទុកក្នុងមួយ Execution៖ ជៀសវាងអាន PropertiesService + បើក Spreadsheet ដដែលៗ ក្នុងសំណើតែមួយ ====
// (សុវត្ថិភាព ព្រោះ Spreadsheet Object ខ្លួនវានៅតែត្រឹមត្រូវ ទោះមាន Sheet Tab ថ្មីត្រូវបានបង្កើត/លុបនៅចន្លោះនោះក៏ដោយ)
var DISTRICT_SS_CACHE_ = {};
function getDistrictSpreadsheet_(district) {
  // ==== ត្រូវកាចដោយ ID Spreadsheet មេ "ជាក់ស្តែង" (effectiveId) ផងដែរ មិនមែនត្រឹមតែឈ្មោះស្រុកទេ — ដូច្នេះនៅពេល
  // Admin ប្តូរតំណភ្ជាប់ Google Sheet មេ (masterSheetLink) ក្នុង Tab "ការកំណត់ប្រព័ន្ធ" ប្រព័ន្ធនឹងបង្កើត Spreadsheet
  // ស្រុកថ្មីៗ ក្រោមការកំណត់ថ្មីនោះ ដោយមិនប៉ះពាល់/លុប Spreadsheet ស្រុកចាស់ៗ ដែលនៅភ្ជាប់ជាមួយ Sheet មេចាស់ឡើយ
  // (អាចត្រឡប់ទៅប្រើឡើងវិញបានគ្រប់ពេល ដោយគ្រាន់តែសម្អាត/ប្តូរតំណភ្ជាប់ masterSheetLink វិញ) ====
  var effectiveId = getTargetSheetId_();
  var cacheKey = effectiveId + '|' + district;
  if (DISTRICT_SS_CACHE_[cacheKey]) return DISTRICT_SS_CACHE_[cacheKey];
  var props = PropertiesService.getScriptProperties();
  var key = DISTRICT_SS_PROP_PREFIX + effectiveId + '_' + district;
  var ssId = props.getProperty(key);
  var dSs = null;
  if (ssId) {
    // ==== ព្យាយាមម្តងទៀត (retry) មុននឹងសន្មតថា Spreadsheet នេះបាត់ — ជៀសវាងបង្កើត Spreadsheet ថ្មីស្ទួន
    // ដោយច្រឡំ ដោយសារកំហុសបណ្ដោះអាសន្ន (ភ្ជាប់បណ្ដាញ/កូតា) ដែលនឹងធ្វើឲ្យទិន្នន័យចាស់បាត់ពីមុខអ្នកប្រើ ====
    try {
      dSs = SpreadsheetApp.openById(ssId);
    } catch (err1) {
      try {
        Utilities.sleep(400);
        dSs = SpreadsheetApp.openById(ssId);
      } catch (err2) {
        dSs = null;
      }
    }
  }
  if (!dSs && ssId) {
    // Spreadsheet ធ្លាប់មានស្រាប់ (មាន ID ចុះទុក) ប៉ុន្តែបើកមិនចេញសូម្បីព្យាយាមម្តងទៀត — កុំបង្កើតថ្មីស្ទួន
    // ដោយស្វ័យប្រវត្តិ (នឹងធ្វើឲ្យទិន្នន័យចាស់របស់ស្រុកនេះលេចមិនឃើញទៀត) ជំនួសមកបោះកំហុសច្បាស់លាស់វិញ
    throw new Error("មិនអាចបើក Spreadsheet របស់ស្រុក " + district + " បានទេ (ID: " + ssId + ")។ សូមពិនិត្យសិទ្ធិចូលប្រើ ឬទាក់ទង Admin។");
  }
  if (!dSs) {
    // ==== ការពារការប៉ះទង្គិចគ្នា (Race Condition — FIX)៖ មុននេះ "ពិនិត្យមុន-បង្កើតក្រោយ" (check-then-create) នេះ
    // គ្មានការចាក់សោទាល់តែសោះ — បើសំណើ ២ (ឧ. គ្រូប្រចាំស្រុកម្នាក់ បញ្ចូលទិន្នន័យ ខណៈ Admin កំពុងបើក Dashboard)
    // កើតឡើងស្ទើរតែក្នុងពេលតែមួយសម្រាប់ស្រុកមួយ ដែលមិនទាន់មាន Spreadsheet ស្រាប់ (ដំបូងគេផុត ឬបន្ទាប់ពី Admin
    // ប្តូរ masterSheetLink ថ្មី) ទាំងពីរអាចបង្កើត Spreadsheet ថ្មីដាច់ដោយឡែក ២ ក្បាល ហើយតែមួយប៉ុណ្ណោះដែលឈ្នះនៅសល់
    // ក្នុង Script Property ក្រោយគេ — ទិន្នន័យដែលបានសរសេរទៅ Spreadsheet ដែលចាញ់នឹងបាត់ដោយស្ងាត់ស្ងៀម (មិនមាន
    // Property ណាមួយចង្អុលទៅវិញទេ)។ ចាក់សោខ្លីមួយភ្លែត (តែក្នុងផ្លូវដ៏កម្រនេះប៉ុណ្ណោះ — ការហៅធម្មតាដែល Spreadsheet
    // មានស្រាប់រួច នៅតែលឿនដដែល មិនចាំបាច់ចាក់សោទេ) ព្រមទាំងពិនិត្យឡើងវិញ (double-check) បន្ទាប់ពីទទួលបានសោ ដើម្បី
    // ចៀសវាងបង្កើតស្ទួន ក្នុងករណីសំណើមួយទៀតបានបង្កើតរួចហើយ ខណៈកំពុងរង់ចាំសោ ====
    var lock = LockService.getScriptLock();
    var gotLock = false;
    try { gotLock = lock.tryLock(20000); } catch (eLock) {}
    try {
      var ssIdAfterLock = props.getProperty(key);
      if (ssIdAfterLock) {
        try {
          dSs = SpreadsheetApp.openById(ssIdAfterLock);
        } catch (errReopen) {
          dSs = null;
        }
      }
      if (!dSs) {
        dSs = SpreadsheetApp.create(districtSheetName_(district));
        var folder = getMasterParentFolder_();
        if (folder) {
          // ==== FIX (សំណើថ្មី "Reset យឺត")៖ File.moveTo() ធ្វើការផ្លាស់ទីក្នុងការហៅ Drive API តែម្តងគត់ (មុននេះ
          // addFile()+removeFile() ត្រូវការហៅ ២ដងដាច់ដោយឡែក — យឺតជាងទ្វេដង ជាពិសេសពេលបង្កើត Spreadsheet ១០ស្រុក
          // ជាប់គ្នា ដូចជាការ Reset ទិន្នន័យទាំងអស់) ====
          try {
            var file = DriveApp.getFileById(dSs.getId());
            file.moveTo(folder);
          } catch (err) {}
        }
        props.setProperty(key, dSs.getId());
        try { ensureDistrictEditTrigger_(dSs); } catch (err) {} // តាមដានអ្នកកែប្រែផ្ទាល់ក្នុង Google Sheet ចាប់ពីពេលនេះតទៅ
      }
    } finally {
      if (gotLock) { try { lock.releaseLock(); } catch (eRel) {} }
    }
  }
  DISTRICT_SS_CACHE_[cacheKey] = dSs;
  return dSs;
}

// ==================== ៦.១ Sheet ថ្មីតាមថ្ងៃខែ (Period Sheet) — ប្រភពទិន្នន័យតែមួយគត់ ====================
// បង្កើត Tab ថ្មីមួយ ក្នុង Spreadsheet ដាច់ដោយឡែករបស់ស្រុកនីមួយៗ ដាក់ឈ្មោះតាមកាលបរិច្ឆេទ "dd-MM" (ឧ. "25-08")។
// Tab នេះ គឺជាប្រភពទិន្នន័យ "តែមួយគត់" សម្រាប់ទិន្នន័យប្រចាំថ្ងៃ — ទាំងសម្រាប់បញ្ចូល/កែ/លុប ដោយផ្ទាល់តាមរយៈ App
// (មើល getOrCreatePeriodSheet_, addDailyEntry, updateDailyEntry, deleteDailyEntry) និងសម្រាប់គ្រូប្រចាំស្រុកកែផ្ទាល់ក្នុង
// Google Sheet ក៏បាន (មិនចាំបាច់ចូល App ទេ)។ លែងមាន Tab ដាច់ដោយឡែកឈ្មោះតាមស្រុកទៀតហើយ។
// Tab នេះ មានជួរឈរបន្ថែម (ស្ថិតិប៉ាន់ស្មាន, ការិ.បង្កើតថ្មី, ចំនួនការិ.សរុប, ទិន្នន័យ២០២៥ សរុប/ស្រី) ដែលបំពេញតម្លៃចាប់ផ្តើម
// ជាមុនម្តង ពេលបង្កើត Tab លើកដំបូង (ចម្លងពី Sheet ដែលមានស្រាប់) ។ rebuildAllDerivedData_ អាន/បូកសរុបផ្ទាល់ពី Tab នេះ។
var PERIOD_EXTRA_FIELDS = [
  { key: "estimate",            label: "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" },
  { key: "newStation",          label: "ការិ.បង្កើតថ្មី" },
  { key: "totalStation",        label: "ចំនួនការិ.សរុប" },
  { key: "baseline2025_total",  label: "ទិន្នន័យ២០២៥ សរុប" },
  { key: "baseline2025_female", label: "ទិន្នន័យ២០២៥ ស្រី" }
];

function todayDateStr_() {
  var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

// ==== FIX (សំណើថ្មី "ការបង្កើត Sheet ត្រូវរៀបតាមលំដាប់ថ្ងៃខែ")៖ បូក/ដក ចំនួនថ្ងៃទៅលើកាលបរិច្ឆេទ "yyyy-MM-dd"
// មួយ ត្រឡប់ជា "yyyy-MM-dd" ថ្មី — ប្រើ Date.UTC() គណនាសុទ្ធសាធ (មិនប្រើ Timezone ក្នុងស្រុករបស់ Server ទាល់តែសោះ)
// ដើម្បីជៀសវាងបញ្ហា DST/Timezone នៅពេលបូក/ដកថ្ងៃ (ត្រឹមត្រូវ ១០០% សម្រាប់ការគណនាកាលបរិច្ឆេទសុទ្ធសាធ ដោយមិនប៉ះពាល់
// ម៉ោងជាក់ស្តែងឡើយ) — ប្រើសម្រាប់គណនា "ថ្ងៃបន្ទាប់" ត្រឹមត្រូវ ពេលបង្កើត Period Sheet ថ្មី (getNextExpectedPeriodDate_) ====
function addDaysToDateStr_(dateStr, n) {
  var m = String(normalizeDateStr_(dateStr) || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return dateStr;
  var ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) + n * 86400000;
  var dt = new Date(ms);
  return dt.getUTCFullYear() + '-' + ('0' + (dt.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + dt.getUTCDate()).slice(-2);
}

// ដាក់ឈ្មោះ Sheet តាមកាលបរិច្ឆេទ (yyyy-MM-dd → dd-MM) ដូចគំរូ "25-08"
function periodSheetName_(dateStr) {
  var m = String(dateStr || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return String(dateStr || '');
  return m[3] + '-' + m[2];
}

function periodSheetTotalCols_() {
  return DAILY_META_PREFIX.length + PERIOD_EXTRA_FIELDS.length + DAILY_FIELDS.length + DAILY_META_SUFFIX.length;
}

// ==== ល្បឿន៖ អាន Sheet(ទិន្នន័យភ្ជាប់) តែម្តងគត់ក្នុងមួយ Execution (ចងចាំក្នុងសតិ) ជំនួសការអានដដែលៗ —
// createDistrictPeriodSheet_ ហៅ getSingleValueMapForDistrict_() ៣ដង (estimate/newStation/totalStation)
// បូក getLinkedDataCodeMap_() ១ដងទៀត សម្រាប់ស្រុកនីមួយៗ ដែលមុននេះនីមួយៗអាន Sheet ដដែលទាំងមូលដាច់ដោយឡែក ពី
// គ្នា (4 ការហៅ Sheets API ស្ទួនគ្នា × ស្រុកទាំង១០ = ~40 ការហៅ សម្រាប់ "បង្កើត Sheet ថ្មីបន្ទាប់" ១ដង) —
// ឥឡូវអានតែម្តងគត់ ហើយចែករំលែកគ្នា ====
var LINKED_DATA_ROWS_CACHE_ = null;
function getLinkedDataRowsOnce_() {
  if (LINKED_DATA_ROWS_CACHE_) return LINKED_DATA_ROWS_CACHE_;
  var ss = getSS_();
  var sheet = ensureLinkedDataSheet_(ss);
  var lastRow = sheet.getLastRow();
  LINKED_DATA_ROWS_CACHE_ = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, LINKED_DATA_TOTAL_COLS).getValues() : [];
  return LINKED_DATA_ROWS_CACHE_;
}
function getSingleValueMapForDistrict_(categoryKey, district) {
  var colIdx = categoryKey === 'estimate' ? 2 : (categoryKey === 'newStation' ? 3 : (categoryKey === 'totalStation' ? 4 : -1));
  var map = {};
  if (colIdx === -1) return map;
  getLinkedDataRowsOnce_().forEach(function(r) {
    if (String(r[0]).trim() === district) map[String(r[1]).trim()] = Number(r[colIdx]) || 0;
  });
  return map;
}

// បង្កើត Tab កាលបរិច្ឆេទថ្មី ១ ក្នុង Spreadsheet ស្រុកមួយ (បើមានរួចហើយ មិនបង្កើតឡើងវិញទេ)
// មួយជួរ = ១ ឃុំ/សង្កាត់ (បំពេញឈ្មោះជាមុន តាមលំដាប់ COMMUNE_ORDER) ដើម្បីឲ្យគ្រូបញ្ចូលទិន្នន័យបានភ្លាមៗ
// អានលេខកូដឃុំ/សង្កាត់ ពី Sheet(ទិន្នន័យភ្ជាប់) ដោយផ្ទាល់ — ប្រើសម្រាប់បំពេញជួរឈរ "លេខកូដ" ក្នុង Period Sheet
var LINKED_DATA_CODE_MAP_CACHE_ = null;
function getLinkedDataCodeMap_() {
  if (LINKED_DATA_CODE_MAP_CACHE_) return LINKED_DATA_CODE_MAP_CACHE_;
  var map = {};
  getLinkedDataRowsOnce_().forEach(function(r) {
    var d = String(r[0] || "").trim(), c = String(r[1] || "").trim();
    if (!d || !c) return;
    map[d + "|" + c] = r[7] || "";
  });
  LINKED_DATA_CODE_MAP_CACHE_ = map;
  return map;
}

// ធានាថា Sheet មួយមានជួរឈរ "លេខកូដ" ត្រឹមត្រូវ នៅជួរឈរទី៣ (ក្រោយ ID+កាលបរិច្ឆេទ មុនស្រុក+ឃុំសង្កាត់)
// ដោះស្រាយគ្រប់ករណីដែលអាចកើតមាន៖ (១) ត្រឹមត្រូវរួចហើយ (២) ធ្លាប់មាននៅចុងជួរឈរ (ការដាក់ពង្រាយពីមុន) (៣) មិនទាន់មានទាល់តែសោះ
function isPeriodSheetName_(name) {
  return /^\d{2}-\d{2}$/.test(String(name || ''));
}

// ត្រឡប់រាល់ Period Sheet ទាំងអស់ (គ្រប់ថ្ងៃខែ) ក្នុង Spreadsheet របស់ស្រុកមួយ
function listDistrictPeriodSheets_(district) {
  var dSs = getDistrictSpreadsheet_(district);
  return dSs.getSheets().filter(function(s) { return isPeriodSheetName_(s.getName()); });
}

// ធានាថា Period Sheet សម្រាប់ស្រុក+កាលបរិច្ឆេទមួយ មានស្រាប់ (បង្កើតបើមិនទាន់មាន) ហើយត្រឡប់ Sheet object ផ្ទាល់
// នេះជាចំណុចចូល (Entry point) ដែល App ប្រើ ដើម្បីទទួលបាន Tab ត្រឹមត្រូវសម្រាប់អាន/សរសេរទិន្នន័យប្រចាំថ្ងៃ
function periodFieldColIndex_(fieldKey) {
  var base = DAILY_META_PREFIX.length + PERIOD_EXTRA_FIELDS.length; // = 9
  for (var i = 0; i < DAILY_FIELDS.length; i++) {
    if (DAILY_FIELDS[i].key === fieldKey) return base + i + 1; // ទៅជា column index ១-based
  }
  return -1;
}
function periodExtraColIndex_(extraKey) {
  var base = DAILY_META_PREFIX.length; // = 4
  for (var i = 0; i < PERIOD_EXTRA_FIELDS.length; i++) {
    if (PERIOD_EXTRA_FIELDS[i].key === extraKey) return base + i + 1;
  }
  return -1;
}
function colLetter_(col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - m) / 26); }
  return s;
}

// កំណត់រូបមន្តគណនាស្វ័យប្រវត្តិ ក្នុងជួរដេកទិន្នន័យទាំងអស់ (firstRow..lastRow) នៃ Period Sheet មួយ
// ១. សរុប = ចុះឈ្មោះថ្មី + ផ្ទេរចូល + ប្តូរការិ.ក្នុងឃុំសង្កាត់
// ២. ប្តូរការិ.ចេញក្នុងឃុំសង្កាត់ = ស្មើនឹង ប្តូរការិ.ក្នុងឃុំសង្កាត់ (ចម្លងស្វ័យប្រវត្តិ)
// ៣. លុបសរុប = ស្លាប់ដកសិទ្ធិផ្លាស់ចេញ + ផ្ទេរចេញក្រៅឃុំសង្កាត់ + ស្ទួន + ប្តូរការិ.ចេញក្នុងឃុំសង្កាត់
// ៤. បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦ = ទិន្នន័យ២០២៥ + សរុប(១) − លុបសរុប(២)
function communeSumTotalCols_() {
  return 2 + DAILY_FIELDS.length + 2;
}
function ensureCommuneSumSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_COMMUNE_SUM);
  var totalCols = communeSumTotalCols_();
  // ==== ការពារទិន្នន័យចាស់៖ បើទម្រង់ក្បាលជួរឈរខុសពីទម្រង់ថ្មី (ជួរឈរតិចជាង ព្រោះ DAILY_FIELDS ទើបមានបន្ថែម, ឬលំដាប់
  // ជួរឈរ "សរុប" ចាស់) សូមប្តូរឈ្មោះ Sheet ចាស់ជា Backup ហើយបង្កើតថ្មី — Tab នេះជាទិន្នន័យសរុបស្វ័យប្រវត្តិសុទ្ធសាធ
  // (មិនមែនកន្លែងបញ្ចូលដោយដៃទេ) ដូច្នេះការបង្កើតឡើងវិញមិនប៉ះពាល់ទិន្នន័យពិតប្រាកដឡើយ (rebuildAllDerivedData_
  // សរសេរជាថ្មីភ្លាមៗ) — ដូចគ្នានឹងគំរូ ensureDailySheet_() ខាងលើ ====
  if (sheet && (sheet.getLastColumn() < totalCols || !derivedSheetOrderMatches_(sheet, 2))) { sheet.setName(SHEET_COMMUNE_SUM + " (ចាស់ " + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Phnom_Penh', 'ddMMyyyy_HHmmss') + ")"); sheet = null; }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COMMUNE_SUM);
    buildSumSheetHeader_(sheet, totalCols, ["ស្រុក", "ឃុំ/សង្កាត់"], ["ចំនួនថ្ងៃមានទិន្នន័យ", "ធ្វើបច្ចុប្បន្នភាពចុងក្រោយ"], "#16a34a");
  }
  return sheet;
}

// ---- Sheet(ស្ថិតិរួមខេត្ត(បូកយោង)) — បូកសរុបស្វ័យប្រវត្តិទាំងខេត្ត (គ្រប់ស្រុករួមគ្នា) ----
function provinceSumTotalCols_() {
  return 1 + DAILY_FIELDS.length + 2;
}
function ensureProvinceSumSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_PROVINCE_SUM);
  var totalCols = provinceSumTotalCols_();
  // ==== ការពារទិន្នន័យចាស់ — ដូចគ្នានឹង ensureCommuneSumSheet_()/ensureDailySheet_() ខាងលើ ====
  if (sheet && (sheet.getLastColumn() < totalCols || !derivedSheetOrderMatches_(sheet, 1))) { sheet.setName(SHEET_PROVINCE_SUM + " (ចាស់ " + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Phnom_Penh', 'ddMMyyyy_HHmmss') + ")"); sheet = null; }
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROVINCE_SUM);
    buildSumSheetHeader_(sheet, totalCols, ["ស្រុក"], ["ចំនួនឃុំ/សង្កាត់មានទិន្នន័យ", "ធ្វើបច្ចុប្បន្នភាពចុងក្រោយ"], "#dc2626");
  }
  return sheet;
}

function sendTelegramMessage_(token, chatId, text) {
  // ==== FIX (Fix119, "Retry Logic ស្វ័យប្រវត្តិ")៖ ការហៅ Telegram API (UrlFetchApp.fetch) ជាការហៅបណ្តាញ
  // ដែលអាចបរាជ័យបណ្តោះអាសន្នម្តងម្កាល (បណ្តាញយឺត/Telegram រវល់មួយភ្លែត) — ជាពិសេសសារ Key Number ដែលឥឡូវផ្ញើ
  // ភ្លាមៗ (notifyTelegramImmediate_, គ្មានជួរដេកការពារទៀតទេ) ត្រូវការភាពជឿជាក់បន្ថែម ព្យាយាមម្តងទៀតដោយស្វ័យ
  // ប្រវត្តិមុននឹងសន្មតថាបរាជ័យពិតប្រាកដ ====
  try {
    return retryWithBackoff_(function() {
      var url = "https://api.telegram.org/bot" + token + "/sendMessage";
      var res = UrlFetchApp.fetch(url, {
        method: "post", contentType: "application/json",
        payload: JSON.stringify({ chat_id: chatId, text: text }),
        muteHttpExceptions: true
      });
      var data = JSON.parse(res.getContentText());
      if (!data.ok) throw new Error('Telegram API ត្រឡប់ ok:false');
      return true;
    }, 2, 500);
  } catch (err) {
    return false;
  }
}

// ==================== ល្បឿន៖ ជូនដំណឹង Telegram ជា "ជួរដេក" (Queue) មិនផ្ញើភ្លាមៗពេលកំពុងព្យួរអ្នកប្រើទេ ====================
// UrlFetchApp.fetch (ការហៅ Telegram API) ជាការហៅបណ្តាញ (Network) ដែលអាចយឺតរាប់វិនាទី — បើហៅផ្ទាល់ (Synchronous)
// ខណៈកំពុង Login/Logout/ស្នើសុំគណនីថ្មី អ្នកប្រើនឹងត្រូវរង់ចាំបន្ថែម រហូតដល់ Telegram ឆ្លើយតបទើបឃើញលទ្ធផល។
// ដូច្នេះ សារត្រូវបានដាក់ក្នុងជួរដេកនេះសិន (រហ័សណាស់ — គ្រាន់តែសរសេរ PropertiesService) រួចកំណត់ Trigger តែម្តង
// ឲ្យផ្ញើនៅពីក្រោយឆាកបន្តិចក្រោយមក (មិនធ្វើឲ្យអ្នកប្រើរង់ចាំ) — ដូចគ្នានឹងគោលការណ៍ scheduleRebuild_ ។
var TG_QUEUE_PROP_KEY_ = 'tgNotifyQueue_v1';
// ==== FIX (សំណើថ្មី "Form Login ដើរយឺតបន្តិច")៖ មុននេះ queueTelegramNotify_() ហៅ ScriptApp.getProjectTriggers()
// រាល់ដងគ្មានលើកលែង ដើម្បីត្រួតពិនិត្យថាមាន Trigger flushTelegramQueue_ កំពុងកំណត់ទុករួចហើយឬអត់ (ជៀសវាងបង្កើត
// Trigger ស្ទួនច្រើនដងក្នុងករណីមានសារជាច្រើនក្នុងរយៈពេលខ្លីៗគ្នា)។ ប៉ុន្តែ ដូចការពន្យល់ត្រង់ scheduleRebuild_() ខាងក្រោម
// (ស្រដៀងគ្នាបេះបិទ)៖ ScriptApp.getProjectTriggers() ជា Apps Script API ដែលយឺតបំផុតមួយ (រាប់រយ Millisecond ឡើង)
// ហើយ Function នេះ (notifyLoginEvent_ → notifyTelegram_ → queueTelegramNotify_) ត្រូវបានហៅ "ដោយផ្ទាល់"
// (Synchronous — មិនទាន់ត្រឡប់លទ្ធផលទៅ Client ទេ) រាល់ពេល Login/Logout របស់គណនីគ្រូប្រចាំក្រុងស្រុក/ឃុំសង្កាត់
// (ភាគច្រើននៃអ្នកប្រើប្រាស់ប្រព័ន្ធ) — នេះជាមូលហេតុចម្បងមួយទៀតដែលធ្វើឲ្យ Form Login មានអារម្មណ៍ថាយឺត។ ជំនួសដូចគ្នា
// នឹង scheduleRebuild_()៖ ប្រើ CacheService ជា Flag ខ្លីមួយ (លឿនជាង ១០ដងឡើង) ជំនួសការស្កេន Trigger ទាំងអស់ក្នុង
// គម្រោងរាល់ដង — Flag ត្រូវបានកំណត់តែក្រោយពេលបង្កើត Trigger ជោគជ័យប៉ុណ្ណោះ (មិនមែនមុន) ដើម្បីជៀសវាង Flag "ភ្លើងខុស"
// ក្នុងករណី newTrigger() បរាជ័យ ====

function queueTelegramNotify_(text) {
  try {
    var lock = LockService.getScriptLock();
    try { lock.tryLock(2000); } catch (e) {}
    try {
      var props = PropertiesService.getScriptProperties();
      var raw = props.getProperty(TG_QUEUE_PROP_KEY_);
      var list = [];
      try { list = raw ? JSON.parse(raw) : []; } catch (e2) { list = []; }
      list.push(text);
      if (list.length > 50) list = list.slice(list.length - 50); // ការពារកុំឲ្យធំហួសទំហំកំណត់ PropertiesService (~9KB/Key)
      props.setProperty(TG_QUEUE_PROP_KEY_, JSON.stringify(list));
    } finally {
      try { lock.releaseLock(); } catch (e3) {}
    }
    var cache = CacheService.getScriptCache();
    if (cache.get('tgFlushTriggerScheduled_')) return; // Trigger ស្រាប់ត្រូវបានគ្រោងទុករួចហើយ (ក្នុងរយៈពេលថ្មីៗនេះ)
    ScriptApp.newTrigger('flushTelegramQueue_').timeBased().after(2000).create();
    cache.put('tgFlushTriggerScheduled_', '1', 6);
  } catch (err) {}
}

// ហៅដោយ Trigger ស្វ័យប្រវត្តិប៉ុណ្ណោះ (មិនត្រូវហៅផ្ទាល់ពីកន្លែងផ្សេងទេ) — ផ្ញើសារទាំងអស់ក្នុងជួរដេក ម្តងតែមួយ រួចលុបជួរដេកចោល
function flushTelegramQueue_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'flushTelegramQueue_') { try { ScriptApp.deleteTrigger(t); } catch (e) {} }
    });
  } catch (err) {}

  var lock = LockService.getScriptLock();
  try { lock.tryLock(5000); } catch (e) {}
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(TG_QUEUE_PROP_KEY_);
    if (!raw) return;
    props.deleteProperty(TG_QUEUE_PROP_KEY_);
    var list = [];
    try { list = JSON.parse(raw) || []; } catch (e2) { list = []; }
    if (!list.length) return;
    var s = getSystemSettings_();
    if (!s.tgBotToken || !s.tgChatId) return;
    list.forEach(function(text) { sendTelegramMessage_(s.tgBotToken, s.tgChatId, text); });
  } catch (err) {
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

// ==================== ល្បឿន៖ ធ្វើសមកាលកម្មសិទ្ធិ Google Sheet (Editor/Viewer) ជាផ្ទៃខាងក្រោយ ====================
// ==== FIX (សំណើ — Tab អ្នកប្រើប្រាស់យឺត)៖ syncSheetPermissions_() ស្កេន/កែប្រែសិទ្ធិ (Drive API — ការហៅបណ្តាញ
// ដែលអាចយឺតរាប់វិនាទីក្នុងមួយ Spreadsheet) លើ Spreadsheet ទាំង១១ (មេ + ស្រុកទាំង១០) ជារៀងរាល់ដង — មុននេះ ហៅផ្ទាល់
// (Synchronous) ពេលបង្កើត/អនុម័តគណនី ធ្វើឲ្យ Admin ត្រូវរង់ចាំយូរ (Browser "កក" រហូតដល់ចប់) មុននឹងឃើញលទ្ធផល។
// ដូច្នេះ ប្តូរឲ្យប្រើគោលការណ៍ Queue+Trigger ដូចគ្នានឹង queueTelegramNotify_ ខាងលើ (រង់ចាំបន្តិចទៀត រួចដំណើរការនៅ
// ពីក្រោយឆាក) សម្រាប់ប្រតិបត្តិការ "ផ្តល់សិទ្ធិ" (បង្កើត/កែគណនី/អនុម័ត/បង្កើត Sheet ថ្មី) ដែលមិនប៉ះពាល់សុវត្ថិភាពខ្លាំង
// ទោះយឺតបន្តិច។ ចំណែក "ផ្អាក"/"លុប" គណនី (ដកសិទ្ធិ) ចេតនាទុកឲ្យហៅផ្ទាល់ (Synchronous) ដដែល — ព្រោះការដកសិទ្ធិ
// Google Sheet ត្រូវការឲ្យកើតឡើងភ្លាមៗ សម្រាប់ហេតុផលសុវត្ថិភាព (កុំឲ្យគណនីដែលទើបផ្អាក/លុប នៅតែបើក Google Sheet
// ដោយផ្ទាល់ (មិនកាត់តាម App) កែប្រែទិន្នន័យបានទៀតរយៈពេលខ្លីមួយ ខណៈកំពុងរង់ចាំ Trigger ដំណើរការ) ====
function scheduleSyncSheetPermissions_() {
  try {
    var already = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'runScheduledSyncSheetPermissions_'; });
    if (!already) ScriptApp.newTrigger('runScheduledSyncSheetPermissions_').timeBased().after(2000).create();
  } catch (err) {
    // ==== FIX (ភាពរឹងមាំ)៖ ដូចគ្នានឹង scheduleRebuild_() ខាងក្រោម — ករណី Trigger បង្កើតមិនចេញ (ឧ. លើសកម្រិត Quota
    // ចំនួន Trigger អតិបរមាក្នុងមួយ Script) មុននេះគណនីដែលទើបបង្កើត/អនុម័ត នឹងគ្មានសិទ្ធិចូល Google Sheet ដោយស្ងាត់ស្ងៀម
    // ដោយគ្មានវិធីណាមួយស្វែងរឡើងវិញទេ (លើកលែងតែ Admin ចាំចុចប៊ូតុង "ធ្វើសមកាលកម្មសិទ្ធិ" ដោយដៃ) — ត្រូវ Fallback
    // ទៅហៅផ្ទាល់ (Synchronous) វិញភ្លាមៗ ជៀសវាងធ្លាក់ខាងក្រោយ ====
    try { syncSheetPermissions_(); } catch (errFallback) {}
  }
}

// ហៅដោយ Trigger ស្វ័យប្រវត្តិប៉ុណ្ណោះ (មិនត្រូវហៅផ្ទាល់ពីកន្លែងផ្សេងទេ — ប្រើ scheduleSyncSheetPermissions_() ជំនួសវិញ)
function runScheduledSyncSheetPermissions_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'runScheduledSyncSheetPermissions_') { try { ScriptApp.deleteTrigger(t); } catch (e) {} }
    });
  } catch (err) {}
  var lock = LockService.getScriptLock();
  try { lock.tryLock(10000); } catch (e) {}
  try {
    syncSheetPermissions_();
  } catch (errSync) {
    // ត្រូវការ Authorize Drive scope លើកទីមួយ (ករណីនេះ) ឬបញ្ហាបណ្តោះអាសន្នផ្សេងទៀត — សូមប្រើប៊ូតុង "ធ្វើសមកាលកម្មសិទ្ធិ" ដោយដៃ
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

// ហៅពីកន្លែងផ្សេងក្នុងកូដ (ឧ. ពេលមានការស្នើសុំគណនីថ្មី) — បរាជ័យដោយស្ងាត់ៗ បើមិនទាន់កំណត់ Token/Chat ID
function periodRowHasData_(row) {
  var n = DAILY_FIELDS.length;
  var fieldsStart0 = periodFieldColIndex_(DAILY_FIELDS[0].key) - 1;
  if (String(row[fieldsStart0 + n] || '').trim()) return true;      // ចំណាំ
  if (String(row[fieldsStart0 + n + 1] || '').trim()) return true;  // អ្នកបញ្ចូល
  for (var i = 0; i < n; i++) {
    var key = DAILY_FIELDS[i].key;
    // ==== "update2026_total"/"update2026_female" ជាជួរឈររូបមន្ត = ទិន្នន័យ២០២៥ (ទិន្នន័យយោង) + បន្ថែម - លុបនៃថ្ងៃនេះ
    // ព្រោះទិន្នន័យ២០២៥ ស្ទើរតែមិនដែលស្មើ ០ ទេ វាលនេះស្ទើរតែមិនដែលចេញជា ០ ទោះមិនទាន់មានអ្នកបញ្ចូលបញ្ចូលទិន្នន័យថ្ងៃនេះក៏ដោយ
    // ដូច្នេះមិនត្រូវរាប់វាលនេះថា "មានទិន្នន័យ" ទេ បើមិនដូច្នេះទេ ស្រុក/ឃុំ ដែលមិនទាន់បញ្ចូល (គ្រាន់តែបើក Tab) នឹងបង្ហាញខុសថាមានទិន្នន័យរួចហើយ
    if (key === 'update2026_total' || key === 'update2026_female') continue;
    if ((Number(row[fieldsStart0 + i]) || 0) !== 0) return true;
  }
  return false;
}

// ស្វែងរកថា ជួរទិន្នន័យ (ID) មួយ កំពុងស្ថិតនៅ Period Sheet (ថ្ងៃខែ) មួយណា របស់ស្រុកមួយណា
// ==== ល្បឿន៖ ប្រើ getLinkedDataRowsOnce_() ជាប្រភពទិន្នន័យរួម ជំនួសការអាន Sheet(ទិន្នន័យភ្ជាប់) ដោយឡែកម្តងទៀត —
// មុននេះ getStationStatsMap_/getBiometricTargetMap_ នីមួយៗអាន Sheet ដដែលទាំងមូលដាច់ដោយឡែកពីគ្នា ខណៈ Dashboard/
// Reports ជាទូទៅហៅមុខងារទាំងពីរនេះជាមួយគ្នាក្នុងការស្នើសុំតែមួយ (= អានស្ទួនគ្នា ២ដងលើសពីតម្រូវការក្នុងករណីនោះ) ====
var STATION_STATS_MAP_CACHE_ = null;
function getStationStatsMap_() {
  if (STATION_STATS_MAP_CACHE_) return STATION_STATS_MAP_CACHE_;
  var map = {};
  getLinkedDataRowsOnce_().forEach(function(r) {
    var d = String(r[0] || "").trim(), c = String(r[1] || "").trim();
    if (!d || !c) return;
    map[d + "|" + c] = { estimate: Number(r[2]) || 0, newStation: Number(r[3]) || 0, totalStation: Number(r[4]) || 0, provinceEstimate: Number(r[10]) || 0 };
  });
  STATION_STATS_MAP_CACHE_ = map;
  return map;
}

// ត្រឡប់ផែនទី { "ស្រុក|ឃុំ" -> { total, female } } "ទិន្នន័យធ្វើជីវមាត្រ" (ទិន្នន័យយោង/គោលដៅ) ពី Sheet(ទិន្នន័យភ្ជាប់)
// ប្រើសម្រាប់គណនាភាគរយ (ចំនួនធ្វើជីវមាត្រពិតប្រាកដ ÷ ទិន្នន័យយោងនេះ) ក្នុងទំព័រ ស្ថិតិទូទៅ (Dashboard)
var BIOMETRIC_TARGET_MAP_CACHE_ = null;
function getBiometricTargetMap_() {
  if (BIOMETRIC_TARGET_MAP_CACHE_) return BIOMETRIC_TARGET_MAP_CACHE_;
  var map = {};
  getLinkedDataRowsOnce_().forEach(function(r) {
    var d = String(r[0] || "").trim(), c = String(r[1] || "").trim();
    if (!d || !c) return;
    map[d + "|" + c] = { total: Number(r[8]) || 0, female: Number(r[9]) || 0 };
  });
  BIOMETRIC_TARGET_MAP_CACHE_ = map;
  return map;
}

// អានលេខកូដឃុំ/សង្កាត់ (ឧ. "21-001") ដែលមានស្រាប់ក្នុង COMMUNE_ORDER
function getCommuneCodeMap_() {
  var map = {};
  DISTRICT_LIST.forEach(function(d) {
    (COMMUNE_ORDER[d] || []).forEach(function(c) { map[d + "|" + c.name] = c.code; });
  });
  return map;
}

function scheduleRebuild_() {
  bumpReportsCacheVersion_(); // ==== ធានាថារបាយការណ៍សរុប (computeSummaries_) មិនបង្ហាញលទ្ធផលចាស់ បន្ទាប់ពីមានការកែប្រែ ====
  // ==== FIX (ល្បឿន x2 — "រក្សាទុកទាំងអស់")៖ ScriptApp.getProjectTriggers() ជា Script Service API ដែលយឺតបំផុតមួយ
  // (ត្រូវទាញយក Metadata គ្រប់ Trigger ទាំងអស់ក្នុងគម្រោង — ជាធម្មតារាប់រយ Millisecond ឡើង) ហើយ Function នេះត្រូវ
  // បានហៅ Synchronous ជានិច្ចជាផ្នែកមួយនៃពេលវេលាឆ្លើយតបចំពោះអ្នកប្រើប្រាស់ផ្ទាល់ (saveDistrictDayEntries ចុច
  // "រក្សាទុកទាំងអស់"/upsertDailyEntry_/onDistrictSheetEdit_ ។ល។) ដូច្នេះជះឥទ្ធិពលដល់ល្បឿនដែលអ្នកប្រើប្រាស់មានអារម្មណ៍
  // ដោយផ្ទាល់។ ជំនួសមកប្រើ CacheService (លឿនជាង ១០ដងឡើង ធៀបនឹង getProjectTriggers()) ជា Flag ខ្លីមួយ ជំនួសការស្កេន
  // Trigger ទាំងអស់ក្នុងគម្រោងរាល់ដង — Flag ត្រូវបានកំណត់តែក្រោយពេលបង្កើត Trigger ជោគជ័យប៉ុណ្ណោះ (មិនមែនមុន) ដើម្បីជៀស
  // វាង Flag "ភ្លើងខុស" ក្នុងករណី newTrigger() បរាជ័យ។ ការផុតកំណត់ ៦វិនាទី (លើសពេលវេលា Trigger ១ភ្លែត ៤វិនាទី បន្តិច
  // ដើម្បីទុកចន្លោះសុវត្ថិភាពសម្រាប់ភាពយឺតយ៉ាវជាក់ស្តែងរបស់ Trigger ដែលអាចលើសពេលកំណត់ដើមខ្លះ)។ ចំណាំសុវត្ថិភាព៖ បើ Cache
  // ផុតកំណត់/ត្រូវបានជម្រះមុនម៉ោង (កម្រណាស់) អាចនាំឲ្យបង្កើត Trigger ស្ទួនកម្រិតតិចតួច ដែលមិនប៉ះពាល់ដល់ភាពត្រឹមត្រូវ
  // នៃទិន្នន័យទេ (rebuildAllDerivedData_ ខ្លួនឯង Idempotent ជានិច្ច — គ្រាន់តែអាចធ្វើការងារដដែលស្ទួនម្តងបន្ថែម) ====
  try {
    var cache = CacheService.getScriptCache();
    if (cache.get('rebuildTriggerScheduled_')) return; // Trigger ស្រាប់ត្រូវបានគ្រោងទុករួចហើយ (ក្នុងរយៈពេលថ្មីៗនេះ)
    ScriptApp.newTrigger('rebuildAllDerivedData_').timeBased().after(4000).create();
    cache.put('rebuildTriggerScheduled_', '1', 6);
  } catch (err) {
    // ករណី Trigger បង្កើតមិនចេញ (ដែនកំណត់/សិទ្ធិ) — ធ្វើភ្លាមៗវិញ ដើម្បីធានាទិន្នន័យមិនធ្លាក់ខាងក្រោយ
    rebuildAllDerivedData_();
  }
}

// បម្លែងជួរដេក Period Sheet (មានជួរឈរ "ទិន្នន័យយោង" បន្ថែម) ទៅជាទម្រង់ Tab(របាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ) (មិនមានជួរឈរនោះ)
