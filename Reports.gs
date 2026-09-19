// ==================== Reports.gs — ការសរុប, របាយការណ៍សម្រាប់បោះពុម្ព, នាំចេញ Excel ====================

// ==== សរសេរតម្លៃទៅក្រឡាឆ្វេងគេបំផុត (១ក្រឡា) ជាមុនសិន រួចទើបបញ្ចូលគ្នា (Merge) ក្រោយ ====
// ជៀសវាងករណីកម្រដែល Google Sheets អាចសរសេរតម្លៃដដែលចូលទៅក្នុងក្រឡាច្រើនក្នុងជួរដេកតែមួយ ប្រសិនបើហៅ
// .merge().setValue() ភ្ជាប់គ្នាភ្លាមៗលើជួរដេកធំៗ (គ្របដណ្តប់ជួរឈរទាំងអស់) ភ្លាមៗបន្ទាប់ពីទើបបង្កើត Spreadsheet ថ្មី
// (ជាហេតុធ្វើឲ្យក្បាលលិខិត ដូចជា "ព្រះរាជាណាចក្រកម្ពុជា" លេចចេញច្រើនដងក្នុងជួរដេកតែមួយ ជំនួសតែម្តងនៅកណ្តាល)។
function mergeWideTitle_(sheet, row, totalCols, value) {
  sheet.getRange(row, 1).setValue(value); // សរសេរតម្លៃចូលក្រឡាតែមួយសិន (មុនមាន Merge ណាមួយប៉ះពាល់)
  var range = sheet.getRange(row, 1, 1, totalCols);
  range.merge();
  return range;
}

// ==== ធានាថា Sheet ថ្មី (ដែលលំនាំដើមមានតែ ២៦ជួរឈរ) មានចំនួនជួរឈរគ្រប់គ្រាន់ (totalCols) មុននឹងហៅ getRange/merge
// ណាមួយទៅលើជួរឈរធំជាងទំហំបច្ចុប្បន្ន — ជៀសវាងកំហុស "The coordinates of the range are outside the dimensions
// of the sheet" សម្រាប់របាយការណ៍ដែលបើកមុខងារបន្ថែមច្រើន (ស្ថិតិប៉ាន់ស្មាន/ភាគរយ/ស្ទួន...) ធ្វើឲ្យជួរឈរសរុបលើសពី ២៦ ====
function ensureSheetCols_(sheet, totalCols) {
  var cur = sheet.getMaxColumns();
  if (cur < totalCols) sheet.insertColumnsAfter(cur, totalCols - cur);
}

// ==== ល្បឿន៖ សរសេរច្រើនជួរដេកម្តងតែមួយកញ្ចប់ (Batch) ជំនួសការហៅ Sheets API ដាច់ដោយឡែក ៥-៧ដងក្នុងមួយជួរដេកៗ ====
// (setValues + setBorder + setFontFamily + setHorizontalAlignment + ជួរឈរតម្រឹមឆ្វេង + ទ្រង់ទ្រាយភាគរយ)។
// នេះជាកន្លែងយឺតបំផុតក្នុងការនាំចេញ Excel (ជួរដេកឃុំ/សង្កាត់រាប់រយ × ការហៅ Sheets API ច្រើនដងក្នុងមួយជួរដេក)។
// ដកយកតែជួរដេក "ក្បាលក្រុមស្រុក" (merge) ចេញ ព្រោះមុខងារនោះនៅតែត្រូវការ setValue+merge ដាច់ដោយឡែក ដូចមុន —
// ជួរដេកទិន្នន័យ/ជួរដេកសរុប (ដែលមិនប្រើ merge) ប៉ុណ្ណោះទើបផ្ទុកក្នុងកញ្ចប់នេះ។ startRow=ជួរដេកចាប់ផ្តើមកញ្ចប់,
// rows2d=អារេ២វិមាត្រនៃតម្លៃ, totalCols=ចំនួនជួរឈរសរុប, opts.leftAlignCol=ជួរឈរតម្រូវតម្រឹមឆ្វេង (2ជួរឈរជាប់គ្នា)
// ឬ null/0=មិនប្រើ, opts.pctCols=អារេជួរឈរភាគរយ (ត្រូវការទ្រង់ទ្រាយ "0.00%"), opts.boldRowOffsets=អារេលិបិក្រម
// (0-indexed ក្នុងកញ្ចប់) នៃជួរដេកសរុប/ជួរដេកចុងបញ្ចប់ ដែលត្រូវដិត + ផ្ទៃខាងក្រោយពណ៌ប្រផេះស្រាល (#f6f8fb)
// ត្រឡប់ជួរដេកបន្ទាប់ (startRow + n) ដូចគ្នានឹងលំនាំដើម "r++" ក្រោយសរសេររួច
function writeBulkRowBlock_(sheet, startRow, rows2d, totalCols, opts) {
  var n = rows2d.length;
  if (!n) return startRow;
  opts = opts || {};
  var range = sheet.getRange(startRow, 1, n, totalCols);
  range.setValues(rows2d);
  range.setBorder(true, true, true, true, true, true).setFontFamily('Kantumruy Pro').setHorizontalAlignment('center');
  if (opts.leftAlignCol) sheet.getRange(startRow, opts.leftAlignCol, n, 2).setHorizontalAlignment('left');
  (opts.pctCols || []).forEach(function(c) {
    if (c) sheet.getRange(startRow, c, n, 1).setNumberFormat('0.00"%"');
  });
  (opts.boldRowOffsets || []).forEach(function(off) {
    sheet.getRange(startRow + off, 1, 1, totalCols).setFontWeight('bold').setBackground('#f6f8fb');
  });
  return startRow + n;
}

function getDailyFormBootstrap(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var sysSettings = getSystemSettings_();
  // ==== FIX (ស្នើសុំ)៖ មុននេះជួរឈរ "សរុប" (correctionBiometricTotal_total/_female) ត្រូវបានដកចេញពីទម្រង់បញ្ចូល
  // ទិន្នន័យទាំងស្រុង (COMPUTED_FIELD_KEYS_) ព្រោះមានរូបមន្តគណនាស្វ័យប្រវត្តិផ្ទាល់ក្នុង Google Sheet។ ឥឡូវផ្លាស់ប្តូរ
  // តាមគំរូដូចគ្នានឹងជួរឈរគណនាស្វ័យប្រវត្តិដទៃទៀត (addTotal_total/delTotal_total/idTotal_total/update2026_total)៖
  // នៅតែបង្ហាញធម្មតាក្នុងទម្រង់បញ្ចូលទិន្នន័យ (មិនដកចេញទៀតទេ) ប៉ុន្តែជាប្រអប់ "សម្រាប់មើលតែប៉ុណ្ណោះ" (readOnly)
  // គណនាស្វ័យប្រវត្តិផ្ទាល់ក្នុង Browser ភ្លាមៗ (មើល AUTO_SUM_MAP ក្នុង Index.html) — ត្រូវការទាំង DAILY_FIELDS ទាំង
  // អស់ (COMPUTED_FIELD_KEYS_ ឥឡូវទទេ — មើល Utils.gs) ====
  var formFields = DAILY_FIELDS.filter(function(f) { return !COMPUTED_FIELD_KEYS_[f.key]; });
  return {
    success: true, districts: DISTRICT_LIST, fields: formFields, canEditAll: isAdmin_(currentUsername),
    communesByDistrict: getCommuneMap_(), baseline2025: getBaseline2025Map_(), stationStats: getStationStatsMap_(),
    communeCodes: getCommuneCodeMap_(),
    showDuplicateFields: sysSettings.showDuplicateFields === "true",
    showEstimate: sysSettings.showEstimate === "true",
    showNewStation: sysSettings.showNewStation === "true",
    showTotalStation: sysSettings.showTotalStation === "true",
    showCommuneCode: sysSettings.showCommuneCode === "true",
    showProvinceEstimate: sysSettings.showProvinceEstimate === "true",
    // ==== ៤ជួរឈរ "ភាគរយ" នេះ លំនាំដើមបង្ហាញ — លាក់លុះត្រាតែកំណត់ជា "false" ជាក់ស្តែងប៉ុណ្ណោះ (ភាគរយ៣/៤ ថ្មី —
    // ធៀបនឹង "ស្ថិតិប៉ាន់ស្មានខេត្ត" ជំនួស "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប") ====
    showAggPercent1: sysSettings.showAggPercent1 !== "false",
    showAggPercent2: sysSettings.showAggPercent2 !== "false",
    showAggPercent3: sysSettings.showAggPercent3 !== "false",
    showAggPercent4: sysSettings.showAggPercent4 !== "false",
    labelBaseline2025: sysSettings.labelBaseline2025 || "ទិន្នន័យ២០២៥",
    labelUpdate2026: sysSettings.labelUpdate2026 || "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦",
    operationStartDate: getOperationStartDate_(), operationEndDate: getOperationEndDate_(), todayStr: formatNow_().slice(0, 10),
    currentPeriodDate: getCurrentPeriodDate_()
  };
}

function computeSummaries_(fromDate, toDate, useRollingReference) {
  // ==== ល្បឿន៖ Cache លទ្ធផល ១ នាទី (ដូច Dashboard) — ការគណនានេះស្កេន Period Sheet គ្រប់ស្រុកទាំង១០ ធ្ងន់ជាងគេបំផុត
  // ក្នុងប្រព័ន្ធ ហើយត្រូវបានហៅរាល់ពេលបើកទំព័រ/របាយការណ៍/នាំចេញ Excel ។ កំណែ Cache (getReportsCacheVersion_) ត្រូវបាន
  // តម្កើងឡើងវិញភ្លាមៗ ពេលមានការកែប្រែទិន្នន័យ (មើល scheduleRebuild_) ដូច្នេះលទ្ធផលថ្មីៗនៅតែធានាបានជានិច្ច ====
  var cacheKey_ = 'summaries_v' + getReportsCacheVersion_() + '_' + (fromDate || '') + '_' + (toDate || '') + '_' + (useRollingReference ? 1 : 0);
  try {
    var cached_ = CacheService.getScriptCache().get(cacheKey_);
    if (cached_) return JSON.parse(cached_);
  } catch (err) {}

  var communeMap = {};
  var districtMap = {};
  fromDate = String(fromDate || '').trim();
  toDate = String(toDate || '').trim();

  // ចាប់ផ្តើមបំពេញជាមុននូវរាល់ ស្រុក/ក្រុង និងឃុំ/សង្កាត់ ទាំងអស់ (តម្លៃ ០) — ដើម្បីឲ្យបង្ហាញគ្រប់ឯកតាទាំងអស់ក្នុងរបាយការណ៍
  // សូម្បីតែឯកតាណាមួយមិនទាន់មានទិន្នន័យបញ្ចូលក៏ដោយ (មិនមែនបង្ហាញតែឯកតាដែលមានទិន្នន័យប៉ុណ្ណោះទេ)
  DISTRICT_LIST.forEach(function(district) {
    districtMap[district] = { district: district, sums: {}, communes: {} };
    DAILY_FIELDS.forEach(function(f) { districtMap[district].sums[f.key] = 0; });
    (COMMUNE_ORDER[district] || []).forEach(function(c) {
      var cKey0 = district + "|" + c.name;
      communeMap[cKey0] = { district: district, commune: c.name, sums: {}, days: {} };
      DAILY_FIELDS.forEach(function(f) { communeMap[cKey0].sums[f.key] = 0; });
    });
  });

  // ==== ល្បឿន៖ អាន DailyIndex (Sheet មួយក្នុង Spreadsheet មេ) ជំនួសការបើក Spreadsheet ស្រុកទាំង១០ ដោយផ្ទាល់ ដូចមុន
  // (ដូចគ្នានឹងវិធីដែលបានប្រើកែល្បឿន Dashboard — មើលមូលហេតុលម្អិតនៅ getDashboardStats() ក្នុង Dashboard.gs) —
  // computeSummaries_ នេះជាកន្លែងធ្ងន់បំផុតក្នុងប្រព័ន្ធ ព្រោះត្រូវបានហៅរាល់ពេលបើកទំព័រសង្ខេប/របាយការណ៍/នាំចេញ Excel ====
  var dailyIndexMap_ = loadDailyIndexOnce_();
  Object.keys(dailyIndexMap_).forEach(function(key) {
    var sep = key.indexOf('|');
    if (sep < 0) return;
    var district = key.slice(0, sep);
    var rowDate = key.slice(sep + 1);
    if (fromDate && rowDate < fromDate) return; // មុនចន្លោះកំណត់ — មិនរាប់បញ្ចូល
    if (toDate && rowDate > toDate) return;     // ក្រោយចន្លោះកំណត់ — មិនរាប់បញ្ចូល
    var entries_ = dailyIndexMap_[key] || [];
    entries_.forEach(function(entry) {
      if (!entry.hasData) return; // រំលងជួរឃុំ/សង្កាត់ដែលមិនទាន់មានអ្នកបញ្ចូល
      var commune = entry.commune;
      if (!commune) return;

      var cKey = district + "|" + commune;
      if (!communeMap[cKey]) {
        communeMap[cKey] = { district: district, commune: commune, sums: {}, days: {} };
        DAILY_FIELDS.forEach(function(f) { communeMap[cKey].sums[f.key] = 0; });
      }
      if (!districtMap[district]) {
        districtMap[district] = { district: district, sums: {}, communes: {} };
        DAILY_FIELDS.forEach(function(f) { districtMap[district].sums[f.key] = 0; });
      }
      var entryValues_ = entry.values || {};
      DAILY_FIELDS.forEach(function(f) {
        // "បច្ចុប្បន្នភាព២០២៦" ជាតម្លៃសរុបបូកបញ្ចូល (Cumulative Snapshot) ក្នុងមួយថ្ងៃៗ — មិនត្រូវបូកបញ្ចូលរួមតាមចន្លោះកាលបរិច្ឆេទទេ
        // (បើបូក នឹងធ្វើឲ្យលេខធំហួសហេតុ ព្រោះជាន់គ្នារាល់ថ្ងៃ) — គណនាដាច់ដោយឡែកខាងក្រោម ពី ទិន្នន័យ២០២៥ + បូក/ដកសរុប
        if (f.key === 'update2026_total' || f.key === 'update2026_female') return;
        var v = Number(entryValues_[f.key]) || 0;
        communeMap[cKey].sums[f.key] += v;
        districtMap[district].sums[f.key] += v;
      });
      communeMap[cKey].days[rowDate] = true; // ថ្ងៃប្លែកគ្នា (មិនរាប់ស្ទួន)
      districtMap[district].communes[commune] = true;
    });
  });

  // ==== ជម្រើសពីរបែប៖ (១) ទិន្នន័យយោងរមូរ (សម្រាប់របាយការណ៍ប្រចាំថ្ងៃតែមួយ ដូចគ្នានឹងទម្រង់បញ្ចូលទិន្នន័យប្រចាំថ្ងៃ)
  // (២) ទិន្នន័យ២០២៥ពិតប្រាកដ + បូក/ដកសរុបលើចន្លោះទាំងមូល (សម្រាប់របាយការណ៍បូកយោង/សង្ខេប — ជៀសវាងរាប់ស្ទួនប្រវត្តិសាស្ត្រ)
  var baseline2025Map, refIsFirstDayOut = true, refLabelOut = null;
  if (useRollingReference) {
    var rollingRef = getRollingReferenceMap_(toDate || todayDateStr_());
    baseline2025Map = rollingRef.map;
    refIsFirstDayOut = rollingRef.isFirstDay;
    refLabelOut = rollingRef.label;
  } else {
    baseline2025Map = getBaseline2025Map_();
  }

  // ==== គណនា "បច្ចុប្បន្នភាព២០២៦" ដាច់ដោយឡែក = ទិន្នន័យយោង (ខាងលើ) + ចំនួនអ្នកចុះឈ្មោះបន្ថែម(សរុប) - ចំនួនលុបចេញ(លុបសរុប)
  for (var ck in communeMap) {
    var cm = communeMap[ck];
    var cb = baseline2025Map[ck] || { total: 0, female: 0 };
    cm.sums.update2026_total = (Number(cb.total) || 0) + (Number(cm.sums.addTotal_total) || 0) - (Number(cm.sums.delTotal_total) || 0);
    cm.sums.update2026_female = (Number(cb.female) || 0) + (Number(cm.sums.addTotal_female) || 0) - (Number(cm.sums.delTotal_female) || 0);
  }
  for (var dk in districtMap) {
    var dm = districtMap[dk];
    var dSumBase = { total: 0, female: 0 };
    (COMMUNE_ORDER[dk] || []).forEach(function(cm2) {
      var b2 = baseline2025Map[dk + "|" + cm2.name];
      if (b2) { dSumBase.total += Number(b2.total) || 0; dSumBase.female += Number(b2.female) || 0; }
    });
    dm.sums.update2026_total = dSumBase.total + (Number(dm.sums.addTotal_total) || 0) - (Number(dm.sums.delTotal_total) || 0);
    dm.sums.update2026_female = dSumBase.female + (Number(dm.sums.addTotal_female) || 0) - (Number(dm.sums.delTotal_female) || 0);
  }

  // ==== ជួរឈរ "សរុប" ថ្មី = ចំនួនករណីកែតម្រូវសរុប + ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ (ជួរឈរនិម្មិត — មិនមាននៅក្នុង
  // DAILY_FIELDS ទេ, សូមមើលមូលហេតុលម្អិតនៅ getReportDisplayFields_() ក្នុង Utils.gs) ====
  for (var ck3 in communeMap) {
    var cmv3 = communeMap[ck3].sums;
    cmv3.correctionBiometricTotal_total = (Number(cmv3.correction_total) || 0) + (Number(cmv3.biometric_total) || 0);
    cmv3.correctionBiometricTotal_female = (Number(cmv3.correction_female) || 0) + (Number(cmv3.biometric_female) || 0);
  }
  for (var dk3 in districtMap) {
    var dmv3 = districtMap[dk3].sums;
    dmv3.correctionBiometricTotal_total = (Number(dmv3.correction_total) || 0) + (Number(dmv3.biometric_total) || 0);
    dmv3.correctionBiometricTotal_female = (Number(dmv3.correction_female) || 0) + (Number(dmv3.biometric_female) || 0);
  }

  var codeMap = getCommuneCodeMap_();
  var stationStatsMap = getStationStatsMap_();
  var bioTargetMap = getBiometricTargetMap_(); // ទិន្នន័យយោង "ត្រូវធ្វើជីវមាត្រ" ពី Tab "ទិន្នន័យភ្ជាប់" — សម្រាប់របាយការណ៍ធ្វើជីវមាត្រ
  var communeList = [];
  for (var cKey2 in communeMap) {
    var c = communeMap[cKey2];
    var cBaseline = baseline2025Map[cKey2] || { total: 0, female: 0 };
    var cStats = stationStatsMap[cKey2] || { estimate: 0, newStation: 0, totalStation: 0, provinceEstimate: 0 };
    var cBioTarget = bioTargetMap[cKey2] || { total: 0, female: 0 };
    communeList.push({
      district: c.district, commune: c.commune, values: c.sums, days: Object.keys(c.days).length,
      baseline: cBaseline, code: codeMap[cKey2] || "", stationStats: cStats, bioTarget: cBioTarget
    });
  }
  var provinceList = [];
  for (var d in districtMap) {
    var p = districtMap[d];
    var dBaseline = { total: 0, female: 0 };
    var dStats = { estimate: 0, newStation: 0, totalStation: 0, provinceEstimate: 0 };
    var dBioTarget = { total: 0, female: 0 };
    (COMMUNE_ORDER[d] || []).forEach(function(cm) {
      var b = baseline2025Map[d + "|" + cm.name];
      if (b) { dBaseline.total += b.total; dBaseline.female += b.female; }
      var s = stationStatsMap[d + "|" + cm.name];
      if (s) { dStats.estimate += s.estimate; dStats.newStation += s.newStation; dStats.totalStation += s.totalStation; dStats.provinceEstimate += Number(s.provinceEstimate) || 0; }
      var bt = bioTargetMap[d + "|" + cm.name];
      if (bt) { dBioTarget.total += Number(bt.total) || 0; dBioTarget.female += Number(bt.female) || 0; }
    });
    provinceList.push({
      district: p.district, values: p.sums, communeCount: Object.keys(p.communes).length,
      baseline: dBaseline, stationStats: dStats, bioTarget: dBioTarget
    });
  }
  var result_ = { communeList: communeList, provinceList: provinceList, referenceIsFirstDay: refIsFirstDayOut, referenceLabel: refLabelOut };
  try { CacheService.getScriptCache().put(cacheKey_, JSON.stringify(result_), REPORTS_CACHE_TTL_); } catch (err) {}
  return result_;
}

// ត្រឡប់ទិន្នន័យសង្ខេបទាំង២កម្រិត (របាយការណ៍សរុបតាមឃុំសង្កាត់ + របាយការណ៍សរុបខេត្ត) សម្រាប់បង្ហាញក្នុងកម្មវិធី
// fromDate/toDate (មិនចាំបាច់) — កំណត់ចន្លោះកាលបរិច្ឆេទសម្រាប់ការបូកសរុប (បូកយោងពី១ថ្ងៃទៅមួយថ្ងៃ) ។ ទុកទទេ = សរុបគ្រប់ថ្ងៃទាំងអស់។
function getSummaries(currentUsername, fromDate, toDate, useRollingReference, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var result = computeSummaries_(fromDate, toDate, useRollingReference);
  var communeList = result.communeList;
  var provinceList = result.provinceList;
  if (!canViewAllDistricts_(currentUsername)) {
    var myDistrict = getUserDistrict_(currentUsername);
    communeList = communeList.filter(function(r) { return r.district === myDistrict; });
    provinceList = provinceList.filter(function(r) { return r.district === myDistrict; });
  }
  // ==== FIX (គណនីឃុំសង្កាត់)៖ កម្រិតបន្ថែម — កម្រិតកន្លងមកខាងលើដាក់កម្រិតត្រឹមតែស្រុកខ្លួន ប៉ុន្តែគណនីឃុំសង្កាត់
  // ត្រូវឃើញតែទិន្នន័យឃុំសង្កាត់ខ្លួនប៉ុណ្ណោះ (មិនមែនទាំងស្រុក)។ provinceList (សរុបតាមស្រុក) ត្រូវលាក់ចោលទាំងស្រុង
  // ព្រោះលេខសរុបនោះនឹងបញ្ចេញឲ្យដឹងពីទិន្នន័យសរុបរបស់ឃុំសង្កាត់ដទៃទៀតក្នុងស្រុកតែម្តង (ដកចេញនូវឃុំសង្កាត់ខ្លួន) ====
  if (isCommuneRole_(role)) {
    var myCommune = getUserCommune_(currentUsername);
    communeList = communeList.filter(function(r) { return r.commune === myCommune; });
    provinceList = [];
  }
  var sysSettings = getSystemSettings_();
  // ==== ការកំណត់បង្ហាញ/លាក់ជួរឈរបន្ថែម ដាច់ដោយឡែកសម្រាប់ Group "Sum" — ចែករំលែករវាង Tab(របាយការណ៍សរុបតាមឃុំសង្កាត់/
  // សរុបខេត្ត) + របាយការណ៍សម្រាប់បោះពុម្ព ប្រភេទបូកយោង (លេខ៣.៤) (ទាំងអស់ហៅ getSummaries() នេះជាប្រភពទិន្នន័យតែមួយ) —
  // មើល resolveGroupDisplaySettings_() ====
  var dispSum_ = resolveGroupDisplaySettings_(sysSettings, 'Sum');
  // ==== FIX (សំណើថ្មី "១"/"៣.៤")៖ ក្រុមថ្មីទាំង២ សម្រាប់ "របាយការណ៍សម្រាប់បោះពុម្ព" ប្រភេទប្រចាំថ្ងៃ (លេខ១.២) និង
  // ធ្វើជីវមាត្របូកយោង (លេខ៥.៦) ដាច់ដោយឡែកពីក្រុម "Sum" ខាងលើ (ដែលឥឡូវសម្រាប់តែលេខ៣.៤ + Tab(របាយការណ៍សរុប)) —
  // Client ជ្រើសរើសក្រុមណាមួយអាស្រ័យលើប្រភេទរបាយការណ៍ដែលកំពុងបង្កើត (មើល applySumDisplaySettings_ ក្នុង Index.html) ====
  var dispPrintDaily_ = resolvePrintDailyDisplaySettings_(sysSettings);
  var dispBio_ = resolveBioDisplaySettings_(sysSettings);
  return {
    success: true,
    fields: getReportDisplayFields_(),
    communeSummary: communeList,
    provinceSummary: provinceList,
    fromDate: fromDate || '',
    toDate: toDate || '',
    districtOrder: DISTRICT_LIST,   // លំដាប់ផ្លូវការនៃ ស្រុក/ក្រុង (តាមលេខកូដ 21-01→21-10)
    communeOrder: COMMUNE_ORDER,    // លំដាប់ផ្លូវការនៃ ឃុំ/សង្កាត់ ក្នុងស្រុក/ក្រុងនីមួយៗ (តាមលេខកូដ) — [{code,name}, ...]
    districtType: DISTRICT_TYPE,    // "ក្រុង" សម្រាប់ដូនកែវ, ក្រៅពីនេះ "ស្រុក"
    districtCode: DISTRICT_CODE,    // លេខកូដ ស្រុក/ក្រុង (21-01 ... 21-10)
    referenceIsFirstDay: result.referenceIsFirstDay, referenceLabel: result.referenceLabel,
    operationStartDate: getOperationStartDate_(), operationEndDate: getOperationEndDate_(),
    todayStr: formatNow_().slice(0, 10),
    showEstimateSum: dispSum_.showEstimate, showProvinceEstimateSum: dispSum_.showProvinceEstimate,
    showNewStationSum: dispSum_.showNewStation, showTotalStationSum: dispSum_.showTotalStation,
    showCommuneCodeSum: dispSum_.showCommuneCode, showDuplicateFieldsSum: dispSum_.showDuplicateFields,
    showAggPercent1: dispSum_.showAggPercent1, showAggPercent2: dispSum_.showAggPercent2,
    showAggPercent3: dispSum_.showAggPercent3, showAggPercent4: dispSum_.showAggPercent4,
    // ==== FIX (សំណើថ្មី "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)/ភាគរយ")៖ ជួរឈរបន្ថែមថ្មី ២ ជាប់បន្ទាប់ពីឃុំ/សង្កាត់ (ក្រុង/ស្រុក)
    // — Tab(របាយការណ៍សរុប) + របាយការណ៍សម្រាប់បោះពុម្ព ប្រភេទបូកយោង (លេខ៣.៤) ប៉ុណ្ណោះ ====
    showBioTotalSum: dispSum_.showBioTotalSum,
    // ==== FIX (សំណើថ្មី "១")៖ ក្រុម "PrintDaily" — របាយការណ៍សម្រាប់បោះពុម្ព ប្រភេទប្រចាំថ្ងៃ (លេខ១.២) ====
    showEstimatePrintDaily: dispPrintDaily_.showEstimate, showProvinceEstimatePrintDaily: dispPrintDaily_.showProvinceEstimate,
    showNewStationPrintDaily: dispPrintDaily_.showNewStation, showTotalStationPrintDaily: dispPrintDaily_.showTotalStation,
    showCommuneCodePrintDaily: dispPrintDaily_.showCommuneCode, showDuplicateFieldsPrintDaily: dispPrintDaily_.showDuplicateFields,
    // ==== FIX (សំណើថ្មី "៣.៤")៖ ក្រុម "Bio" — របាយការណ៍សម្រាប់បោះពុម្ព ប្រភេទធ្វើជីវមាត្របូកយោង (លេខ៥.៦) ====
    showEstimateBio: dispBio_.showEstimate, showProvinceEstimateBio: dispBio_.showProvinceEstimate,
    showCommuneCodeBio: dispBio_.showCommuneCode,
    orgName: sysSettings.orgName, orgSlogan: sysSettings.orgSlogan, logoUrl: sysSettings.logoUrl,
    footerChairmanTitle: sysSettings.footerChairmanTitle, footerChairmanName: sysSettings.footerChairmanName,
    footerPreparerTitle: sysSettings.footerPreparerTitle, footerPreparerName: sysSettings.footerPreparerName,
    reportLogoUrl: sysSettings.reportLogoUrl, reportCommitteeName: sysSettings.reportCommitteeName,
    reportSecretariatName: sysSettings.reportSecretariatName
  };
}

// ==================== ៨. នាំចេញរបាយការណ៍ជា Excel (.xlsx) ====================
// reportType: 'dailyCommune' | 'dailyProvince' | 'cumCommune' | 'cumProvince'
// បង្កើត Google Sheet បណ្តោះអាសន្ន → បំពេញទិន្នន័យ+ទម្រង់ដូចរបាយការណ៍បោះពុម្ព → នាំចេញជា .xlsx (base64) → លុបចោល
function exportReportToExcel(currentUsername, reportType, fromDate, toDate, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };

  // ==== FIX (គណនីឃុំសង្កាត់)៖ របាយការណ៍ "...សរុបខេត្ត" (dailyProvince/cumProvince/bioProvince) បង្ហាញលេខសរុបតាមស្រុក
  // ដែលនឹងបញ្ចេញឲ្យដឹងពីទិន្នន័យសរុបរបស់ឃុំសង្កាត់ដទៃទៀតក្នុងស្រុកតែម្តង — ដូច្នេះបដិសេធភ្លាមៗ មិនឲ្យគណនីឃុំសង្កាត់
  // នាំចេញរបាយការណ៍ប្រភេទនេះបានឡើយ (ជំនួសការត្រឡប់ Excel ទទេស្រងាត់ដែលអាចមានការភាន់ច្រឡំ) ====
  if (isCommuneRole_(role) && (reportType === 'dailyProvince' || reportType === 'cumProvince' || reportType === 'bioProvince')) {
    return { success: false, message: "គណនីឃុំសង្កាត់មិនអាចនាំចេញរបាយការណ៍សរុបខេត្តបានទេ!" };
  }

  // ==== "របាយការណ៍ធ្វើជីវមាត្រ" (Tab ៥-៦) មានរចនាសម្ព័ន្ធជួរឈរខុសគ្នាទាំងស្រុង — ប្រគល់ទៅ Function ដាច់ដោយឡែក ====
  if (reportType === 'bioCommune' || reportType === 'bioProvince') {
    return exportBiometricReportToExcel_(currentUsername, reportType, fromDate, toDate, sessionToken);
  }

  var isDailyType_ = (reportType === 'dailyCommune' || reportType === 'dailyProvince');
  var result = computeSummaries_(fromDate, toDate, isDailyType_);
  var communeList = result.communeList;
  var provinceList = result.provinceList;
  if (!canViewAllDistricts_(currentUsername)) {
    var myDistrict = getUserDistrict_(currentUsername);
    communeList = communeList.filter(function(r) { return r.district === myDistrict; });
    provinceList = provinceList.filter(function(r) { return r.district === myDistrict; });
  }
  // ==== FIX (គណនីឃុំសង្កាត់)៖ កម្រិតបន្ថែម ដាក់កម្រិតត្រឹមតែឃុំសង្កាត់ខ្លួន (មិនមែនទាំងស្រុក) ====
  if (isCommuneRole_(role)) {
    var myCommune_ = getUserCommune_(currentUsername);
    communeList = communeList.filter(function(r) { return r.commune === myCommune_; });
    provinceList = [];
  }

  var isCommune = (reportType === 'dailyCommune' || reportType === 'cumCommune');
  var isCumulative = (reportType === 'cumCommune' || reportType === 'cumProvince');

  // ==== FIX (សំណើថ្មី)៖ របាយការណ៍ ២/៤ (...សរុបខេត្ត) ជាការពិតរាយបញ្ជីតាមឃុំ/សង្កាត់នីមួយៗ (មិនមែនជាតួលេខសរុប
  // តែមួយទេ) ដូច្នេះចំណងជើងតារាង ត្រូវប្តូរឲ្យដូចគ្នានឹងរបាយការណ៍ ១/៣ ("...តាមឃុំ សង្កាត់") ជំនួស "...សរុបខេត្ត" ====
  var titleMap = {
    dailyCommune: 'លទ្ធផលបណ្តោះអាសន្ន នៃការត្រួតពិនិត្យបញ្ជីឈ្មោះ និងការចុះឈ្មោះបោះឆ្នោត តាមឃុំ សង្កាត់',
    dailyProvince: 'លទ្ធផលបណ្តោះអាសន្ន នៃការត្រួតពិនិត្យបញ្ជីឈ្មោះ និងការចុះឈ្មោះបោះឆ្នោត តាមឃុំ សង្កាត់',
    cumCommune: 'លទ្ធផលបណ្តោះអាសន្ន នៃការត្រួតពិនិត្យបញ្ជីឈ្មោះ និងការចុះឈ្មោះបោះឆ្នោត តាមឃុំ សង្កាត់',
    cumProvince: 'លទ្ធផលបណ្តោះអាសន្ន នៃការត្រួតពិនិត្យបញ្ជីឈ្មោះ និងការចុះឈ្មោះបោះឆ្នោត តាមឃុំ សង្កាត់'
  };
  var title = titleMap[reportType] || 'របាយការណ៍';
  var isDailyType = !isCumulative;

  // ==== អានការកំណត់បង្ហាញ/លាក់បច្ចុប្បន្ន ដើម្បីឲ្យ Excel ដែលនាំចេញ ត្រូវនឹងរបាយការណ៍សម្រាប់បោះពុម្ពដែលកំពុងបង្ហាញលើ
  // អេក្រង់ជាក់ស្តែង — FIX (សំណើថ្មី "១")៖ របាយការណ៍ប្រចាំថ្ងៃ (លេខ១.២) ឥឡូវប្រើក្រុម "PrintDaily" ដាច់ដោយឡែក
  // ពីរបាយការណ៍បូកយោង (លេខ៣.៤ — ក្រុម "Sum" ចាស់ ចែករំលែកជាមួយ Tab(របាយការណ៍សរុប)) — មើល resolvePrintDailyDisplaySettings_()/
  // resolveGroupDisplaySettings_() ក្នុង Settings.gs ====
  var sysSettings_ = getSystemSettings_();
  var disp_ = isDailyType_ ? resolvePrintDailyDisplaySettings_(sysSettings_) : resolveGroupDisplaySettings_(sysSettings_, 'Sum');
  var showEstimate_ = disp_.showEstimate;
  var showNewStation_ = disp_.showNewStation;
  var showTotalStation_ = disp_.showTotalStation;
  var showCommuneCode_ = disp_.showCommuneCode;
  var showProvinceEstimate_ = disp_.showProvinceEstimate;
  var showDuplicateFields_ = disp_.showDuplicateFields;
  // ==== ៤ជួរឈរ "ភាគរយ" នេះ លំនាំដើមបង្ហាញ (ខុសពីធីចផ្សេងទៀត) — លាក់លុះត្រាតែកំណត់ជា "false" ជាក់ស្តែងប៉ុណ្ណោះ ====
  // ភាគរយ១/២ ធៀបនឹង "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" (ដូចមុន) ។ ភាគរយ៣/៤ ថ្មី ធៀបនឹង "ស្ថិតិប៉ាន់ស្មានខេត្ត" ====
  var showAggPercent1_ = disp_.showAggPercent1;
  var showAggPercent2_ = disp_.showAggPercent2;
  var showAggPercent3_ = disp_.showAggPercent3;
  var showAggPercent4_ = disp_.showAggPercent4;
  var HIDDEN_FIELD_KEYS_GS = { delDuplicate_total: true, delDuplicate_female: true };
  var fields = getReportDisplayFields_().filter(function(f) {
    if (HIDDEN_FIELD_KEYS_GS[f.key] && !showDuplicateFields_) return false;
    return true;
  });

  // ជួរឈរបន្ថែម "យោង" — សម្រាប់របាយការណ៍បូកយោង បង្ហាញទិន្នន័យមូលដ្ឋានឆ្នាំ២០២៥ជាក់ស្តែង (មិនមែនទទេទៀតទេ)។
  // ឈ្មោះជួរឈរនេះប្តូរស្វ័យប្រវត្តិ៖ មុនដល់ "ថ្ងៃចាប់ផ្តើមដំណើរការ" បង្ហាញ "ទិន្នន័យ២០២៥" បន្ទាប់ពីនោះប្តូរទៅ
  // "យោងចំនួនក្នុងបញ្ជីសរុបពីមុន" វិញ (ប៉ុន្តែទិន្នន័យឆ្នាំ២០២៥ ត្រូវបានប្រើក្នុងការគណនាជានិច្ច សម្រាប់របាយការណ៍បូកយោង)
  var REF_COL_LABEL_GS = result.referenceIsFirstDay ? "ទិន្នន័យ២០២៥" : (result.referenceLabel || "យោងចំនួនក្នុងបញ្ជីសរុបពីមុន");

  // ==== ភាពរឹងមាំ៖ ចន្លោះកាលបរិច្ឆេទបញ្ច្រាស (fromDate > toDate) ធ្វើឲ្យរបាយការណ៍ទទេ/ច្របូកច្របល់ដោយស្ងាត់ស្ងៀម
  // ជាជាងបញ្ហាធ្ងន់ធ្ងរ ប៉ុន្តែសមនឹងជូនដំណឹងច្បាស់លាស់ជាជាងឲ្យអ្នកប្រើឆ្ងល់ថាហេតុអ្វីរបាយការណ៍ទទេ ====
  if (fromDate && toDate && String(fromDate) > String(toDate)) {
    return { success: false, message: "\"ពីថ្ងៃ\" ត្រូវនៅមុន ឬស្មើ \"ដល់ថ្ងៃ\"!" };
  }

  // ==== ភាពរឹងមាំ៖ Spreadsheet បណ្តោះអាសន្នខាងក្រោម ត្រូវបានធានាថាលុបចោលជានិច្ច (try/finally) ទោះបីជាកើតកំហុស
  // ណាមួយក៏ដោយ ខណៈកំពុងសរសេរទិន្នន័យ/នាំចេញ — មុននេះ បើកើតកំហុសពាក់កណ្តាលដំណើរការ Spreadsheet បណ្តោះអាសន្ននេះនឹង
  // នៅសេសសល់អចិន្ត្រៃយ៍ក្នុង Drive (មិនដែលត្រូវលុបចោល) ជាហេតុឲ្យកកកុញកាន់តែច្រើនតាមពេលវេលា ====
  var tempSs = SpreadsheetApp.create('EXPORT_' + newId_());
  var tempFileId_ = tempSs.getId();
  try {
  var sheet = tempSs.getSheets()[0];
  sheet.setName('របាយការណ៍');
  sheet.setRightToLeft(false);
  SpreadsheetApp.flush();

  // ==== ជួរឈរខាងឆ្វេង ដូចគ្នានឹងអេក្រង់ជាក់ស្តែង៖ លេខកូដ/ស្ថិតិប៉ាន់ស្មាន អាចលាក់/បង្ហាញបាន តាមការកំណត់ ====
  // ភាគរយទី១ = (ចុះឈ្មោះថ្មី + ផ្ទេរចូល) ÷ ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប × ១០០ ។ ភាគរយទី២ = ចុះឈ្មោះថ្មី ÷ ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប × ១០០
  // ភាគរយទី៣ (ថ្មី) = ចុះឈ្មោះថ្មី ÷ ស្ថិតិប៉ាន់ស្មានខេត្ត × ១០០ ។ ភាគរយទី៤ (ថ្មី) = (ចុះឈ្មោះថ្មី + ផ្ទេរចូល) ÷ ស្ថិតិប៉ាន់ស្មានខេត្ត × ១០០
  // ត្រូវបានប្រើសម្រាប់តែរបាយការណ៍បូកយោង (cumCommune/cumProvince) ហើយអាចលាក់/បង្ហាញនីមួយៗដាច់ដោយឡែក
  var showPct1_ = isCumulative && showAggPercent1_;
  var showPct2_ = isCumulative && showAggPercent2_;
  var showPct3_ = isCumulative && showAggPercent3_;
  var showPct4_ = isCumulative && showAggPercent4_;
  // ==== FIX (សំណើថ្មី "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)/ភាគរយ")៖ ជួរឈរបន្ថែមថ្មី ២ ដាក់ភ្លាមៗបន្ទាប់ពីជួរឈរ ឃុំ/សង្កាត់
  // (ក្រុង/ស្រុក) — មានន័យសម្រាប់តែរបាយការណ៍បូកយោង (isCumulative — លេខ៣.៤) ដូចគ្នានឹងអេក្រង់ (Tab(របាយការណ៍សរុប) +
  // renderPrintReport()) ====
  var showBioTotal_ = isCumulative && !!disp_.showBioTotalSum;
  var bioTargetCol_ = 0, bioPctCol_ = 0;
  // ==== FIX (ធាតុទី៤-៥ ក្នុងសំណើ)៖ ជួរឈរ "ភាគរយ" ដាក់ភ្លាមៗបន្ទាប់ពីជួរឈរស្ថិតិប៉ាន់ស្មានពាក់ព័ន្ធរៀងៗខ្លួន (interleaved)
  // ជំនួសការដាក់ជាប្លុកដាច់ដោយឡែកនៅចុងតារាងដូចមុន — ភាគរយ១/២ បន្ទាប់ពី "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប", ភាគរយ៣/៤ បន្ទាប់ពី
  // "ស្ថិតិប៉ាន់ស្មានខេត្ត" ====
  var pctCol1_ = 0, pctCol2_ = 0, pctCol3_ = 0, pctCol4_ = 0;
  var leftCols = [];
  if (isCommune) {
    if (showCommuneCode_) leftCols.push('លេខកូដ');
    leftCols.push('ឃុំ/សង្កាត់');
  } else {
    leftCols.push('ល.រ'); // លេខរៀង ១→១០ (មិនមែនលេខកូដទេ) — បង្ហាញជានិច្ចសម្រាប់របាយការណ៍សរុបខេត្ត
    leftCols.push('ក្រុង/ស្រុក');
  }
  if (showBioTotal_) {
    bioTargetCol_ = leftCols.length + 1; leftCols.push('ទិន្នន័យធ្វើជីវមាត្រ (សរុប)');
    bioPctCol_ = leftCols.length + 1; leftCols.push('ភាគរយ');
  }
  if (showEstimate_) {
    leftCols.push('ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប');
    if (showPct1_) { pctCol1_ = leftCols.length + 1; leftCols.push('ភាគរយ'); }
    if (showPct2_) { pctCol2_ = leftCols.length + 1; leftCols.push('ភាគរយ'); }
  }
  if (showProvinceEstimate_) {
    leftCols.push('ស្ថិតិប៉ាន់ស្មានខេត្ត');
    if (showPct3_) { pctCol3_ = leftCols.length + 1; leftCols.push('ភាគរយ'); }
    if (showPct4_) { pctCol4_ = leftCols.length + 1; leftCols.push('ភាគរយ'); }
  }
  if (showNewStation_) leftCols.push('ការិ.បង្កើតថ្មី');
  if (showTotalStation_) leftCols.push('ចំនួនការិ.សរុប');
  var totalCols = leftCols.length + 2 + fields.length; // +2 = ជួរឈរយោង (សរុប+ស្រី) — ជួរឈរ "ភាគរយ" ឥឡូវរាប់បញ្ចូលក្នុង leftCols រួចហើយ
  ensureSheetCols_(sheet, totalCols);

  // ---- ក្បាលលិខិត + ចំណងជើង (Font "Moul") ----
  // ==== កំណត់ជាអនុគមន៍ ហៅពេលក្រោយ (បន្ទាប់ពី autoResizeColumns) ព្រោះ Google Sheets ផ្គុំទទឹងជួរឈរ A តាម
  // អក្សរដែលមាននៅក្នុងក្រឡា A ផ្ទាល់ (ទោះបីជាបានបញ្ចូលគ្នា/Merge ពេញមួយជួរដេកហើយក៏ដោយ) — បើសរសេរចំណងជើងវែងៗ
  // ទាំងនេះមុន នោះជួរឈរ "លេខកូដ"/"ល.រ" (ជួរឈរ A) នឹងធំហួសទំហំតាមចំណងជើង មិនមែនតាមខ្លឹមសារខ្លួនវាទេ។
  function writeTitleRows_() {
    mergeWideTitle_(sheet, 1, totalCols, 'ព្រះរាជាណាចក្រកម្ពុជា').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(13).setFontFamily('Moul');
    mergeWideTitle_(sheet, 2, totalCols, 'ជាតិ សាសនា ព្រះមហាក្សត្រ').setHorizontalAlignment('center').setFontFamily('Moul');
    mergeWideTitle_(sheet, 3, totalCols, title).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12).setFontFamily('Moul');
    if (isDailyType) {
      // របាយការណ៍ប្រចាំថ្ងៃ៖ បង្ហាញ "សម្រាប់ថ្ងៃទី..." ដូចជាបន្ត​ចំណងជើង (មិនបង្ហាញបន្ទាត់ ខេត្តតាកែវ/គិតត្រឹម ដាច់ដោយឡែកទៀតទេ)
      mergeWideTitle_(sheet, 4, totalCols, 'សម្រាប់' + formatKhmerDate_(fromDate)).setHorizontalAlignment('center').setFontWeight('bold').setFontFamily('Moul');
      mergeWideTitle_(sheet, 5, totalCols, '').setHorizontalAlignment('center').setFontFamily('Moul');
    } else {
      // ==== បន្ទាត់ "ចាប់ពី...ដល់..." យកតាមថ្ងៃចាប់ផ្តើម/ថ្ងៃចុងបញ្ចប់ដំណើរការ (Tab ការកំណត់ប្រព័ន្ធ) ជានិច្ច — មិនមែនតាមកាលបរិច្ឆេទបូកយោងដែលបានជ្រើសរើសទេ ====
      var opStart_ = getOperationStartDate_(), opEnd_ = getOperationEndDate_();
      var subtitle = opStart_ && opEnd_ ? ('ចាប់ពី' + formatKhmerDate_(opStart_) + ' ដល់' + formatKhmerDate_(opEnd_))
        : opStart_ ? ('ចាប់ពី' + formatKhmerDate_(opStart_))
        : opEnd_ ? ('រហូតដល់' + formatKhmerDate_(opEnd_))
        : '';
      if (subtitle) mergeWideTitle_(sheet, 4, totalCols, subtitle).setHorizontalAlignment('center').setFontWeight('bold').setFontFamily('Moul');
      mergeWideTitle_(sheet, 5, totalCols, 'គិតត្រឹមថ្ងៃទី ' + toDate).setHorizontalAlignment('center').setFontFamily('Moul');
    }
  }

  // ---- ក្បាលតារាង (ជាន់ទី៦-៨ បញ្ចូលគ្នាតាមក្រុម/ក្រុមរង/ឯកតា) ----
  var headerStartRow = 6;
  leftCols.forEach(function(label, idx) {
    sheet.getRange(headerStartRow, idx + 1, 3, 1).merge().setValue(label);
  });
  // ==== ជួរឈរយោង (សរុប+ស្រី) — colspan=2 នៅជាន់ទី១-២ រួច "សរុប"/"ស្រី" ដាច់ដោយឡែកនៅជាន់ទី៣ ====
  var refColIdx = leftCols.length + 1;
  sheet.getRange(headerStartRow, refColIdx, 2, 2).merge().setValue(REF_COL_LABEL_GS);
  sheet.getRange(headerStartRow + 2, refColIdx).setValue('សរុប');
  sheet.getRange(headerStartRow + 2, refColIdx + 1).setValue('ស្រី');
  var col = refColIdx + 2;
  // ==== ស្វែងរកជួរឈរពិតប្រាកដនៃក្រុម "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦" (update2026_total/update2026_female) —
  // គណនាតាមលិបិក្រមក្នុង fields (មិនមែនតួលេខថេរទេ ព្រោះជួរឈរខាងលើអាចត្រូវលាក់/បង្ហាញតាមការកំណត់) ====
  var u26Idx_ = -1;
  for (var fi_ = 0; fi_ < fields.length; fi_++) { if (fields[fi_].key === 'update2026_total') { u26Idx_ = fi_; break; } }
  var u26Col_ = u26Idx_ >= 0 ? (col + u26Idx_) : 0;
  var i = 0;
  while (i < fields.length) {
    var f = fields[i];
    var groupSpan = 1;
    while (i + groupSpan < fields.length && fields[i + groupSpan].group === f.group) groupSpan++;

    // ==== ពិនិត្យថាតើក្រុមនេះ មានក្រុមរងពិតប្រាកដដែរឬទេ (បើគ្មាន បញ្ចូលគ្នាជាន់ទី១+ទី២ ដើម្បីកុំឲ្យទុកទទេ) ====
    var isSingleSub = true;
    for (var si = i; si < i + groupSpan; si++) { if (fields[si].sub) { isSingleSub = false; break; } }
    if (isSingleSub) {
      sheet.getRange(headerStartRow, col, 2, groupSpan).merge().setValue(f.group);
    } else {
      sheet.getRange(headerStartRow, col, 1, groupSpan).merge().setValue(f.group);
    }

    var j = i, subCol = col;
    while (j < i + groupSpan) {
      var sf = fields[j];
      if (!sf.sub) {
        sheet.getRange(headerStartRow + 2, subCol, 1, 1).setValue(sf.unit); // ត្រឹមតែជាន់ទី៣ (ជាន់ទី១-២ បានបញ្ចូលគ្នារួច)
        subCol++; j++;
      } else {
        var subSpan = 1;
        while (j + subSpan < i + groupSpan && fields[j + subSpan].sub === sf.sub) subSpan++;
        sheet.getRange(headerStartRow + 1, subCol, 1, subSpan).merge().setValue(sf.sub);
        for (var k = 0; k < subSpan; k++) {
          sheet.getRange(headerStartRow + 2, subCol + k, 1, 1).setValue(fields[j + k].unit);
        }
        subCol += subSpan; j += subSpan;
      }
    }
    col += groupSpan;
    i += groupSpan;
  }
  // ==== ជួរឈរ "ភាគរយ" (បើមាន) ត្រូវបានសរសេរក្បាលរួចហើយ ដោយ leftCols.forEach() ខាងលើ (interleaved ភ្លាមៗបន្ទាប់ពី
  // ជួរឈរស្ថិតិប៉ាន់ស្មានពាក់ព័ន្ធ — មើលកន្លែងកំណត់ pctCol1_-pctCol4_ ខាងលើ ជិត leftCols) ====
  var headerRange = sheet.getRange(headerStartRow, 1, 3, totalCols);
  headerRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBackground('#eef2f9').setBorder(true, true, true, true, true, true).setFontFamily('Kantumruy Pro');

  // ---- ជួរទិន្នន័យ ----
  var dataStartRow = headerStartRow + 3;
  var r = dataStartRow;

  function pct_(num, den) { return den > 0 ? (Math.round((num / den) * 10000) / 100) : ''; } // ភាគរយ (២ខ្ទង់ក្រោយចំណុច) ឬទទេបើគ្មានភាគបែង
  // ==== ជួរឈរ "បច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦" (u26Col_/u26Col_+1) ត្រូវនៅតម្រឹមឆ្វេង — ទិន្នន័យផ្សេងទៀត
  // ទាំងអស់តម្រឹមកណ្តាល តាមការស្នើសុំ។ Font ទិន្នន័យទាំងអស់ប្តូរជា "Kantumruy Pro" ====
  function writeRow_(rowIdx, cells, bold) {
    var range = sheet.getRange(rowIdx, 1, 1, totalCols);
    range.setValues([cells]);
    range.setBorder(true, true, true, true, true, true).setFontFamily('Kantumruy Pro').setHorizontalAlignment('center');
    if (u26Col_) sheet.getRange(rowIdx, u26Col_, 1, 2).setHorizontalAlignment('left');
    if (bold) range.setFontWeight('bold').setBackground('#f6f8fb');
    if (showPct1_ && cells[pctCol1_ - 1] !== '') sheet.getRange(rowIdx, pctCol1_).setNumberFormat('0.00"%"');
    if (showPct2_ && cells[pctCol2_ - 1] !== '') sheet.getRange(rowIdx, pctCol2_).setNumberFormat('0.00"%"');
    if (showPct3_ && cells[pctCol3_ - 1] !== '') sheet.getRange(rowIdx, pctCol3_).setNumberFormat('0.00"%"');
    if (showPct4_ && cells[pctCol4_ - 1] !== '') sheet.getRange(rowIdx, pctCol4_).setNumberFormat('0.00"%"');
    if (showBioTotal_ && cells[bioPctCol_ - 1] !== '') sheet.getRange(rowIdx, bioPctCol_).setNumberFormat('0.00"%"');
  }
  // ==== ជួរឈរ "ភាគរយ" (បើមាន) ត្រូវបានបញ្ចូលដោយ leftCellsFor_() ខាងក្រោម ដាក់ភ្ជាប់ភ្លាមៗបន្ទាប់ពីជួរឈរ
  // ស្ថិតិប៉ាន់ស្មានពាក់ព័ន្ធ (ត្រូវនឹងទីតាំងក្បាលតារាង pctCol1_-pctCol4_ ដែលបានកំណត់រួចនៅខាងលើ ជិត leftCols) ====
  // ជួរក្បាលក្រុម (លេខកូដ+ឈ្មោះស្រុក/ក្រុង) — បញ្ចូលគ្នា (merge) ពេញមួយជួរ ដាក់នៅលើឈ្មោះឃុំ/សង្កាត់ជាក់ស្តែង
  function writeDistrictHead_(rowIdx, district) {
    var label = districtCode_(district) + '  ' + getDistrictType_(district) + district;
    mergeWideTitle_(sheet, rowIdx, totalCols, label)
      .setFontWeight('bold').setHorizontalAlignment('left').setBackground('#dfe6f5').setFontFamily('Kantumruy Pro')
      .setBorder(true, true, true, true, true, true);
  }
  function sumFields_(rows) {
    var sums = {};
    fields.forEach(function(f) { sums[f.key] = 0; });
    rows.forEach(function(row) { fields.forEach(function(f) { sums[f.key] += Number(row.values[f.key]) || 0; }); });
    return sums;
  }
  // ==== ក្រឡាខាងឆ្វេង (លេខកូដ/ឈ្មោះ/ស្ថិតិប៉ាន់ស្មាន/យោង) ដូចគ្នាសម្រាប់ជួរដេកទិន្នន័យ និងជួរដេកសរុប — គោរពការកំណត់បង្ហាញ/លាក់ ====
  function leftCellsFor_(codeVal, nameVal, stationStats, baseline, addNewT, addTrT, bioTargetTotal, bioActualTotal) {
    var out = [];
    if (isCommune ? showCommuneCode_ : true) out.push(codeVal || '');
    out.push(nameVal);
    if (showBioTotal_) {
      out.push(Number(bioTargetTotal) || 0);
      out.push(pct_(Number(bioActualTotal) || 0, Number(bioTargetTotal) || 0));
    }
    var s = stationStats || { estimate: 0, newStation: 0, totalStation: 0, provinceEstimate: 0 };
    var a = Number(addNewT) || 0, t = Number(addTrT) || 0;
    if (showEstimate_) {
      out.push(Number(s.estimate) || 0);
      if (showPct1_) out.push(pct_(a + t, Number(s.estimate) || 0));
      if (showPct2_) out.push(pct_(a, Number(s.estimate) || 0));
    }
    if (showProvinceEstimate_) {
      out.push(Number(s.provinceEstimate) || 0);
      if (showPct3_) out.push(pct_(a, Number(s.provinceEstimate) || 0));
      if (showPct4_) out.push(pct_(a + t, Number(s.provinceEstimate) || 0));
    }
    if (showNewStation_) out.push(Number(s.newStation) || 0);
    if (showTotalStation_) out.push(Number(s.totalStation) || 0);
    var b = baseline || { total: 0, female: 0 };
    out.push(Number(b.total) || 0, Number(b.female) || 0); // ==== បង្ហាញទិន្នន័យយោង (សរុប+ស្រី) ជានិច្ច (ទាំងប្រចាំថ្ងៃ និងបូកយោង) ====
    return out;
  }
  function sumBaseline_(rows) {
    return rows.reduce(function(s, row) {
      s.total += (row.baseline && row.baseline.total) || 0;
      s.female += (row.baseline && row.baseline.female) || 0;
      return s;
    }, { total: 0, female: 0 });
  }
  function sumStationStats_(rows) {
    return rows.reduce(function(s, row) {
      var st = row.stationStats || {};
      s.estimate += Number(st.estimate) || 0; s.newStation += Number(st.newStation) || 0; s.totalStation += Number(st.totalStation) || 0;
      s.provinceEstimate += Number(st.provinceEstimate) || 0;
      return s;
    }, { estimate: 0, newStation: 0, totalStation: 0, provinceEstimate: 0 });
  }
  // ==== "ទិន្នន័យធ្វើជីវមាត្រ (សរុប)" (row.bioTarget.total) + ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រពិតប្រាកដ
  // (row.values.biometric_total) — បូកសរុបលើជួរដេកកូនច្រើន សម្រាប់ជួរដេកសរុបស្រុក/សរុបខេត្ត ====
  function sumBioTarget_(rows) {
    return rows.reduce(function(s, row) { return s + ((row.bioTarget && Number(row.bioTarget.total)) || 0); }, 0);
  }
  function sumBioActual_(rows) {
    return rows.reduce(function(s, row) { return s + ((row.values && Number(row.values.biometric_total)) || 0); }, 0);
  }

  // ==== ល្បឿន៖ ជួរដេកឃុំ/សង្កាត់ក្នុងស្រុកតែមួយ (រហូតដល់ជួរដេកសរុបស្រុក) សរសេរជាកញ្ចប់តែម្តង (មិនមែនម្តងមួយជួរដេកទៀតទេ) —
  // ជួរដេក "ក្បាលក្រុមស្រុក" (writeDistrictHead_ ប្រើ merge) នៅតែសរសេរដាច់ដោយឡែកជាមុន ព្រោះ merge ត្រូវការដំណើរការ
  // ខុសពីគេ (setValue+merge) — ប៉ុន្តែជួរដេកនោះមានចំនួនតិច (១ក្នុងមួយស្រុក ១០ស្រុក) ដូច្នេះមិនប៉ះពាល់ល្បឿនច្រើនទេ ====
  if (isCommune) {
    var list = communeList.slice().sort(districtCommuneComparator_);
    var lastDistrict = null, pending = [], districtBlockRows_ = [];
    var flushDistrictBlock_ = function(district) {
      if (!pending.length) return;
      var sums = sumFields_(pending);
      var lbl = getDistrictType_(district) === 'ក្រុង' ? 'សរុបក្រុង' : 'សរុបស្រុក';
      var stSum = sumStationStats_(pending);
      var subtotalCells = leftCellsFor_('', lbl, stSum, sumBaseline_(pending), sums.addNew_total, sums.addTransferIn_total, sumBioTarget_(pending), sumBioActual_(pending)).concat(fields.map(function(f) { return sums[f.key]; }));
      districtBlockRows_.push(subtotalCells);
      r = writeBulkRowBlock_(sheet, r, districtBlockRows_, totalCols, {
        leftAlignCol: u26Col_, pctCols: [pctCol1_, pctCol2_, pctCol3_, pctCol4_, bioPctCol_], boldRowOffsets: [districtBlockRows_.length - 1]
      });
      districtBlockRows_ = [];
      pending = [];
    };
    list.forEach(function(row) {
      if (row.district !== lastDistrict) {
        if (lastDistrict !== null) flushDistrictBlock_(lastDistrict);
        writeDistrictHead_(r, row.district); r++;
        lastDistrict = row.district;
      }
      var cells = leftCellsFor_(communeCode_(row.district, row.commune), row.commune, row.stationStats, row.baseline, row.values.addNew_total, row.values.addTransferIn_total, (row.bioTarget && row.bioTarget.total) || 0, (row.values && row.values.biometric_total) || 0).concat(fields.map(function(f) { return Number(row.values[f.key]) || 0; }));
      districtBlockRows_.push(cells);
      pending.push(row);
    });
    if (lastDistrict !== null) flushDistrictBlock_(lastDistrict);
    if (list.length) {
      var grand = sumFields_(list);
      var stG = sumStationStats_(list);
      var cellsG = leftCellsFor_('', 'សរុបខេត្ត', stG, sumBaseline_(list), grand.addNew_total, grand.addTransferIn_total, sumBioTarget_(list), sumBioActual_(list)).concat(fields.map(function(f) { return grand[f.key]; }));
      writeRow_(r, cellsG, true); r++;
    } else {
      sheet.getRange(r, 1).setValue('គ្មានទិន្នន័យទេ'); r++;
    }
  } else {
    var plist = provinceList.slice().sort(districtOnlyComparator_);
    var provinceBlockRows_ = [];
    plist.forEach(function(row, idx) {
      var cells = leftCellsFor_(idx + 1, row.district, row.stationStats, row.baseline, row.values.addNew_total, row.values.addTransferIn_total, (row.bioTarget && row.bioTarget.total) || 0, (row.values && row.values.biometric_total) || 0).concat(fields.map(function(f) { return Number(row.values[f.key]) || 0; }));
      provinceBlockRows_.push(cells);
    });
    if (plist.length) {
      var grand2 = sumFields_(plist);
      var stG2 = sumStationStats_(plist);
      var cellsG2 = leftCellsFor_('-', 'សរុបខេត្ត', stG2, sumBaseline_(plist), grand2.addNew_total, grand2.addTransferIn_total, sumBioTarget_(plist), sumBioActual_(plist)).concat(fields.map(function(f) { return grand2[f.key]; }));
      provinceBlockRows_.push(cellsG2);
      r = writeBulkRowBlock_(sheet, r, provinceBlockRows_, totalCols, {
        leftAlignCol: u26Col_, pctCols: [pctCol1_, pctCol2_, pctCol3_, pctCol4_, bioPctCol_], boldRowOffsets: [provinceBlockRows_.length - 1]
      });
    } else {
      sheet.getRange(r, 1).setValue('គ្មានទិន្នន័យទេ'); r++;
    }
  }

  // ==== ធ្វើ Auto-resize ជួរឈរទាំងអស់ (រួមទាំង "លេខកូដ"/"ល.រ") តាមខ្លឹមសារក្បាលតារាង+ទិន្នន័យ ជាមុនសិន —
  // មុននឹងសរសេរចំណងជើងវែងៗចូលទៅជួរដេកទី១-៥ (ភ្ជាប់ជាមួយជួរឈរ A) ដើម្បីកុំឲ្យជួរឈរ A ត្រូវបានពង្រីកតាមចំណងជើង ====
  try { sheet.autoResizeColumns(1, totalCols); } catch (err) {}
  writeTitleRows_();
  SpreadsheetApp.flush();

  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempFileId_ + '/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    return { success: false, message: "ការនាំចេញ Excel បរាជ័យ (កូដ " + response.getResponseCode() + ")! សូមព្យាយាមម្តងទៀត។" };
  }
  var base64 = Utilities.base64Encode(response.getContent());

  var fileNameMap = {
    dailyCommune: 'របាយការណ៍ប្រចាំថ្ងៃ_គ្រប់ឃុំសង្កាត់',
    dailyProvince: 'របាយការណ៍ប្រចាំថ្ងៃ_សរុបខេត្ត',
    cumCommune: 'របាយការណ៍បូកយោង_គ្រប់ឃុំសង្កាត់',
    cumProvince: 'របាយការណ៍បូកយោង_សរុបខេត្ត'
  };
  var filename = (fileNameMap[reportType] || 'របាយការណ៍') + '_' + (toDate || formatNow_().slice(0, 10)) + '.xlsx';

  return { success: true, base64: base64, filename: filename };
  } catch (err) {
    return { success: false, message: "កំហុសពេលនាំចេញ Excel៖ " + err.message };
  } finally {
    try { DriveApp.getFileById(tempFileId_).setTrashed(true); } catch (e) {}
  }
}

// ==================== ៨.១ របាយការណ៍ធ្វើជីវមាត្រ (Tab ៥-៦) — រចនាសម្ព័ន្ធជួរឈរដាច់ដោយឡែក ====================
// reportType: 'bioCommune' | 'bioProvince' — តែងតែជាទម្រង់បូកយោង (មិនមានទម្រង់ប្រចាំថ្ងៃទេ)
// ជួរឈរ៖ លេខកូដ/ល.រ, ឈ្មោះ, ស្ថិតិប៉ាន់ស្មានក្រៅរបប, [ចុះឈ្មោះថ្មី|ផ្ទេរចូល|សរុបថ្មី](សរុប/ស្រី), ភាគរយ,
//        ចំនួនឈ្មោះក្នុងបញ្ជីដែលត្រូវមកធ្វើជីវមាត្រ (សរុប/ស្រី — មកពី Tab "ទិន្នន័យភ្ជាប់"),
//        ចំនួនអ្នកមកធ្វើបច្ចុប្បន្នភាពជីវមាត្រ (សរុប/ស្រី — មកពីការបញ្ចូលទិន្នន័យប្រចាំថ្ងៃ បូកសរុបលើចន្លោះកាលបរិច្ឆេទ), ភាគរយ
// ភាគរយទី១ = សរុបថ្មី ÷ ស្ថិតិប៉ាន់ស្មាន × ១០០ ។ ភាគរយទី២ = ចំនួនអ្នកមកធ្វើបច្ចុប្បន្នភាពជីវមាត្រ ÷ ចំនួនឈ្មោះក្នុងបញ្ជីដែលត្រូវមកធ្វើជីវមាត្រ × ១០០
var BIO_REPORT_K_LABEL_ = 'ចំនួនឈ្មោះក្នុងបញ្ជីដែលត្រូវមកធ្វើជីវមាត្រ';
var BIO_REPORT_M_LABEL_ = 'ចំនួនអ្នកមកធ្វើបច្ចុប្បន្នភាពជីវមាត្រ';

// ==== FIX (សុវត្ថិភាព — សំខាន់ណាស់, Authentication Bypass)៖ Function នេះ (ទោះមានសញ្ញា "_" ខាងចុងឈ្មោះ ដែលតាមទំនៀម
// ក្នុងគម្រោងនេះមានន័យថា "ខាងក្នុងតែប៉ុណ្ណោះ") ពិតជាអាចហៅផ្ទាល់ពី Browser បានដដែល តាមរយៈ google.script.run (ដូចគ្នា
// នឹងហេតុផលដែលពន្យល់លម្អិតក្នុង Auth.gs ស្តីពី requireValidSession_) ព្រោះ Google Apps Script អនុញ្ញាតឲ្យហៅ Top-level
// Function ណាមួយក៏បាន ដោយមិនគិតពីឈ្មោះ។ មុននេះ Function នេះគ្មានការផ្ទៀងផ្ទាត់ Session ទាល់តែសោះ (មិនដូច
// exportReportToExcel() ដែលហៅវា) ដូច្នេះនរណាម្នាក់ ដោយមិនចាំបាច់ចូល Login ទាល់តែសោះ អាចហៅ
// google.script.run.exportBiometricReportToExcel_("admin", "bioProvince", "", "") ដោយផ្ទាល់ (គ្រាន់តែទាយ/ស្គាល់
// Username របស់ SuperAdmin/Admin/PEC21/សង្កេតការណ៍ណាម្នាក់ ដោយមិនចាំបាច់ដឹងលេខសម្ងាត់ សូម្បីតែម្តង) ដើម្បីទាញយក
// របាយការណ៍ជីវមាត្រពេញខេត្តជា .xlsx បាន — ត្រូវបន្ថែម sessionToken + requireValidSession_ ដូចមុខងារផ្សេងទៀតទាំងអស់ ====
function exportBiometricReportToExcel_(currentUsername, reportType, fromDate, toDate, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  if (fromDate && toDate && String(fromDate) > String(toDate)) {
    return { success: false, message: "\"ពីថ្ងៃ\" ត្រូវនៅមុន ឬស្មើ \"ដល់ថ្ងៃ\"!" };
  }
  // ==== FIX (គណនីឃុំសង្កាត់ — សុវត្ថិភាព)៖ Function នេះអាចហៅផ្ទាល់ពី Browser បាន (មើលកំណត់ចំណាំខាងលើ) ដូច្នេះត្រូវ
  // ធ្វើការត្រួតពិនិត្យនេះនៅទីនេះដោយផ្ទាល់ដែរ មិនអាចទុកចិត្តតែលើ exportReportToExcel() ដែលហៅវាឡើយ ====
  var __roleForBioCheck_ = getUserRole_(currentUsername);
  if (isCommuneRole_(__roleForBioCheck_) && reportType === 'bioProvince') {
    return { success: false, message: "គណនីឃុំសង្កាត់មិនអាចនាំចេញរបាយការណ៍សរុបខេត្តបានទេ!" };
  }
  var result = computeSummaries_(fromDate, toDate, false); // តែងតែបូកយោង — ទិន្នន័យ២០២៥ ជាមូលដ្ឋានគណនា update2026
  var communeList = result.communeList;
  var provinceList = result.provinceList;
  if (!canViewAllDistricts_(currentUsername)) {
    var myDistrict = getUserDistrict_(currentUsername);
    communeList = communeList.filter(function(r) { return r.district === myDistrict; });
    provinceList = provinceList.filter(function(r) { return r.district === myDistrict; });
  }
  // ==== FIX (គណនីឃុំសង្កាត់)៖ កម្រិតបន្ថែម ដាក់កម្រិតត្រឹមតែឃុំសង្កាត់ខ្លួន (មិនមែនទាំងស្រុក) ====
  if (isCommuneRole_(__roleForBioCheck_)) {
    var myCommuneBio_ = getUserCommune_(currentUsername);
    communeList = communeList.filter(function(r) { return r.commune === myCommuneBio_; });
    provinceList = [];
  }
  var isCommune = (reportType === 'bioCommune');

  // ==== FIX (សំណើថ្មី)៖ របាយការណ៍ ៦ (...សរុបខេត្ត) ជាការពិតរាយបញ្ជីតាមឃុំ/សង្កាត់នីមួយៗ ដូច្នេះចំណងជើងតារាង
  // ត្រូវដូចគ្នានឹងរបាយការណ៍ ៥ ("...តាមឃុំ សង្កាត់") ជំនួស "...សរុបខេត្ត" ====
  var titleMap = {
    bioCommune: 'លទ្ធផលបណ្តោះអាសន្ន នៃការធ្វើបច្ចុប្បន្នភាពជីវមាត្រ តាមឃុំ សង្កាត់',
    bioProvince: 'លទ្ធផលបណ្តោះអាសន្ន នៃការធ្វើបច្ចុប្បន្នភាពជីវមាត្រ តាមឃុំ សង្កាត់'
  };
  var title = titleMap[reportType] || 'របាយការណ៍ធ្វើជីវមាត្រ';

  // ==== FIX (សំណើថ្មី "៣.៤")៖ ក្រុម "Bio" ដាច់ដោយឡែក — គាំទ្រការជ្រើសរើសប្រភពស្ថិតិប៉ាន់ស្មាន (គ.ជ.ប/ខេត្ត/ទាំងពីរ)
  // ជំនួសការ Hardcode បង្ហាញតែ "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" ជានិច្ចដូចមុន — មើល resolveBioDisplaySettings_() ក្នុង Settings.gs ====
  var sysSettings_ = getSystemSettings_();
  var dispBio_ = resolveBioDisplaySettings_(sysSettings_);
  var showCommuneCode_ = dispBio_.showCommuneCode;
  var showEstimateBio_ = dispBio_.showEstimate;
  var showProvinceEstimateBio_ = dispBio_.showProvinceEstimate;
  var pctCount_ = (showEstimateBio_ ? 1 : 0) + (showProvinceEstimateBio_ ? 1 : 0);

  // ==== ភាពរឹងមាំ៖ Spreadsheet បណ្តោះអាសន្នខាងក្រោម ត្រូវបានធានាថាលុបចោលជានិច្ច (try/finally) — មើលមូលហេតុលម្អិត
  // ក្នុង exportReportToExcel ខាងលើ (Pattern ដូចគ្នា) ====
  var tempSs = SpreadsheetApp.create('EXPORT_' + newId_());
  var tempFileId_ = tempSs.getId();
  try {
  var sheet = tempSs.getSheets()[0];
  sheet.setName('របាយការណ៍');
  sheet.setRightToLeft(false);
  SpreadsheetApp.flush();

  // ==== FIX (សំណើថ្មី "៣.៤")៖ ជួរឈរ "ស្ថិតិប៉ាន់ស្មាន" ឥឡូវអាចមាន ០/១/២ ជួរឈរ (គ.ជ.ប/ខេត្ត/ទាំងពីរ) អាស្រ័យលើការកំណត់
  // ក្រុម "Bio" — ជំនួសការ Hardcode ១ជួរឈរជានិច្ចដូចមុន ====
  var leftCols = [];
  if (isCommune) {
    if (showCommuneCode_) leftCols.push('លេខកូដ');
    leftCols.push('ឃុំ/សង្កាត់');
  } else {
    leftCols.push('ល.រ');
    leftCols.push('ក្រុង/ស្រុក');
  }
  if (showEstimateBio_) leftCols.push('ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប');
  if (showProvinceEstimateBio_) leftCols.push('ស្ថិតិប៉ាន់ស្មានខេត្ត');
  var totalCols = leftCols.length + 6 + pctCount_ + 2 + 2 + 1; // +6 ចុះឈ្មោះថ្មី/ផ្ទេរចូល/សរុបថ្មី, +pctCount_ %(តាមចំនួនស្ថិតិប៉ាន់ស្មានដែលបានជ្រើសរើស), +2 ២០២៦, +2 ជីវមាត្រ, +1 %
  ensureSheetCols_(sheet, totalCols);

  // ---- ក្បាលលិខិត + ចំណងជើង (បូកយោងជានិច្ច, Font "Moul") ----
  // ==== កំណត់ជាអនុគមន៍ ហៅពេលក្រោយ (បន្ទាប់ពី autoResizeColumns) ដូចគ្នានឹង exportReportToExcel ដើម្បីកុំឲ្យ
  // ជួរឈរ A ("លេខកូដ"/"ល.រ") ត្រូវបានពង្រីកតាមចំណងជើងវែងៗ ====
  function writeTitleRows_() {
    mergeWideTitle_(sheet, 1, totalCols, 'ព្រះរាជាណាចក្រកម្ពុជា').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(13).setFontFamily('Moul');
    mergeWideTitle_(sheet, 2, totalCols, 'ជាតិ សាសនា ព្រះមហាក្សត្រ').setHorizontalAlignment('center').setFontFamily('Moul');
    mergeWideTitle_(sheet, 3, totalCols, title).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12).setFontFamily('Moul');
    // ==== បន្ទាត់ "ចាប់ពី...ដល់..." យកតាមថ្ងៃចាប់ផ្តើម/ថ្ងៃចុងបញ្ចប់ដំណើរការ (Tab ការកំណត់ប្រព័ន្ធ) ជានិច្ច — មិនមែនតាមកាលបរិច្ឆេទបូកយោងដែលបានជ្រើសរើសទេ ====
    var opStart_ = getOperationStartDate_(), opEnd_ = getOperationEndDate_();
    var subtitle = opStart_ && opEnd_ ? ('ចាប់ពី' + formatKhmerDate_(opStart_) + ' ដល់' + formatKhmerDate_(opEnd_))
      : opStart_ ? ('ចាប់ពី' + formatKhmerDate_(opStart_))
      : opEnd_ ? ('រហូតដល់' + formatKhmerDate_(opEnd_))
      : '';
    if (subtitle) mergeWideTitle_(sheet, 4, totalCols, subtitle).setHorizontalAlignment('center').setFontWeight('bold').setFontFamily('Moul');
    mergeWideTitle_(sheet, 5, totalCols, 'គិតត្រឹមថ្ងៃទី ' + toDate).setHorizontalAlignment('center').setFontFamily('Moul');
  }

  // ---- ក្បាលតារាង (ជាន់ទី៦-៨) ----
  var headerStartRow = 6;
  leftCols.forEach(function(label, idx) {
    sheet.getRange(headerStartRow, idx + 1, 3, 1).merge().setValue(label);
  });
  var col = leftCols.length + 1;
  sheet.getRange(headerStartRow, col, 1, 6).merge().setValue('ចំនួនអ្នកចុះឈ្មោះបន្ថែម');
  ['ចុះឈ្មោះថ្មី', 'ផ្ទេរចូល', 'សរុបថ្មី'].forEach(function(lbl, si) {
    sheet.getRange(headerStartRow + 1, col + si * 2, 1, 2).merge().setValue(lbl);
    sheet.getRange(headerStartRow + 2, col + si * 2).setValue('សរុប');
    sheet.getRange(headerStartRow + 2, col + si * 2 + 1).setValue('ស្រី');
  });
  col += 6;
  // ==== FIX (សំណើថ្មី "៣.៤")៖ ០/១/២ ជួរឈរ "ភាគរយ" ត្រង់នេះ អាស្រ័យលើចំនួនស្ថិតិប៉ាន់ស្មានដែលបានជ្រើសរើស (pctCount_) —
  // លំដាប់ដូចគ្នានឹងលំដាប់ជួរឈរ "ស្ថិតិប៉ាន់ស្មាន" ក្នុង leftCols (គ.ជ.ប មុន, ខេត្តក្រោយ) ====
  var pctColsBio_ = [];
  for (var pbi_ = 0; pbi_ < pctCount_; pbi_++) {
    pctColsBio_.push(col);
    sheet.getRange(headerStartRow, col, 3, 1).merge().setValue('ភាគរយ');
    col += 1;
  }
  var u26Col_ = col; // ==== ជួរឈរ "ចំនួនឈ្មោះក្នុងបញ្ជីដែលត្រូវមកធ្វើជីវមាត្រ" (u26Col_/u26Col_+1) — ត្រូវនៅតម្រឹមឆ្វេង តាមការស្នើសុំ ====
  sheet.getRange(headerStartRow, col, 2, 2).merge().setValue(BIO_REPORT_K_LABEL_);
  sheet.getRange(headerStartRow + 2, col).setValue('សរុប');
  sheet.getRange(headerStartRow + 2, col + 1).setValue('ស្រី');
  col += 2;
  sheet.getRange(headerStartRow, col, 2, 2).merge().setValue(BIO_REPORT_M_LABEL_);
  sheet.getRange(headerStartRow + 2, col).setValue('សរុប');
  sheet.getRange(headerStartRow + 2, col + 1).setValue('ស្រី');
  col += 2;
  var percent2Col = col;
  sheet.getRange(headerStartRow, col, 3, 1).merge().setValue('ភាគរយ');

  var headerRange = sheet.getRange(headerStartRow, 1, 3, totalCols);
  headerRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBackground('#eef2f9').setBorder(true, true, true, true, true, true).setFontFamily('Kantumruy Pro');

  // ---- ជួរទិន្នន័យ ----
  var dataStartRow = headerStartRow + 3;
  var r = dataStartRow;

  function pct_(num, den) { return den > 0 ? (Math.round((num / den) * 10000) / 100) : ''; } // ភាគរយ (២ខ្ទង់ក្រោយចំណុច) ឬទទេបើគ្មានភាគបែង

  // ==== Font ទិន្នន័យទាំងអស់ប្តូរជា "Kantumruy Pro" តម្រឹមកណ្តាល លើកលែងតែជួរឈរ u26Col_/u26Col_+1 (ចំនួនឈ្មោះក្នុងបញ្ជីឆ្នាំ២០២៦) តម្រឹមឆ្វេង ====
  function writeRow_(rowIdx, cells, bold) {
    var range = sheet.getRange(rowIdx, 1, 1, totalCols);
    range.setValues([cells]);
    range.setBorder(true, true, true, true, true, true).setFontFamily('Kantumruy Pro').setHorizontalAlignment('center');
    sheet.getRange(rowIdx, u26Col_, 1, 2).setHorizontalAlignment('left');
    if (bold) range.setFontWeight('bold').setBackground('#f6f8fb');
    pctColsBio_.forEach(function(c) { if (cells[c - 1] !== '') sheet.getRange(rowIdx, c).setNumberFormat('0.00"%"'); });
    if (cells[percent2Col - 1] !== '') sheet.getRange(rowIdx, percent2Col).setNumberFormat('0.00"%"');
  }
  function writeDistrictHead_(rowIdx, district) {
    var label = districtCode_(district) + '  ' + getDistrictType_(district) + district;
    mergeWideTitle_(sheet, rowIdx, totalCols, label)
      .setFontWeight('bold').setHorizontalAlignment('left').setBackground('#dfe6f5').setFontFamily('Kantumruy Pro')
      .setBorder(true, true, true, true, true, true);
  }

  // ==== K = ចំនួនឈ្មោះក្នុងបញ្ជីដែលត្រូវមកធ្វើជីវមាត្រ (bioTarget — មកពី Tab "ទិន្នន័យភ្ជាប់") ====
  // ==== M = ចំនួនអ្នកមកធ្វើបច្ចុប្បន្នភាពជីវមាត្រ (v.biometric_total/female — ការបញ្ចូលទិន្នន័យប្រចាំថ្ងៃ បូកសរុប) ====
  // ==== FIX (សំណើថ្មី "៣.៤")៖ estimate/provinceEstimate ទាំង២ត្រូវបានហុចមកទាំងអស់ជានិច្ច ប៉ុន្តែសរសេរចូល Sheet
  // តែជួរឈរដែលបានបើកតាមការកំណត់ក្រុម "Bio" ប៉ុណ្ណោះ (ត្រូវនឹងលំដាប់ leftCols ខាងលើ — គ.ជ.ប មុន ខេត្តក្រោយ) ====
  function rowCells_(codeVal, nameVal, estimate, provinceEstimate, v, bioTarget) {
    var addNewT = Number(v.addNew_total) || 0, addNewF = Number(v.addNew_female) || 0;
    var addTrT = Number(v.addTransferIn_total) || 0, addTrF = Number(v.addTransferIn_female) || 0;
    var sumT = addNewT + addTrT, sumF = addNewF + addTrF;
    var est = Number(estimate) || 0;
    var provEst = Number(provinceEstimate) || 0;
    var bioTargetT = Number(bioTarget.total) || 0, bioTargetF = Number(bioTarget.female) || 0;
    var bioActualT = Number(v.biometric_total) || 0, bioActualF = Number(v.biometric_female) || 0;
    var out = [];
    if (isCommune ? showCommuneCode_ : true) out.push(codeVal || '');
    out.push(nameVal);
    if (showEstimateBio_) out.push(est);
    if (showProvinceEstimateBio_) out.push(provEst);
    out.push(addNewT, addNewF, addTrT, addTrF, sumT, sumF);
    if (showEstimateBio_) out.push(pct_(sumT, est));
    if (showProvinceEstimateBio_) out.push(pct_(sumT, provEst));
    out.push(bioTargetT, bioTargetF);
    out.push(bioActualT, bioActualF);
    out.push(pct_(bioActualT, bioTargetT));
    return out;
  }
  function sumOf_(rows, key) { return rows.reduce(function(s, row) { return s + (Number(row.values[key]) || 0); }, 0); }
  function sumEstimate_(rows) { return rows.reduce(function(s, row) { return s + (Number(row.stationStats && row.stationStats.estimate) || 0); }, 0); }
  function sumProvinceEstimate_(rows) { return rows.reduce(function(s, row) { return s + (Number(row.stationStats && row.stationStats.provinceEstimate) || 0); }, 0); }
  function sumBioTarget_(rows) {
    return rows.reduce(function(s, row) {
      s.total += (row.bioTarget && row.bioTarget.total) || 0;
      s.female += (row.bioTarget && row.bioTarget.female) || 0;
      return s;
    }, { total: 0, female: 0 });
  }
  function subtotalCells_(rows, label, codeVal) {
    var v = {
      addNew_total: sumOf_(rows, 'addNew_total'), addNew_female: sumOf_(rows, 'addNew_female'),
      addTransferIn_total: sumOf_(rows, 'addTransferIn_total'), addTransferIn_female: sumOf_(rows, 'addTransferIn_female'),
      biometric_total: sumOf_(rows, 'biometric_total'), biometric_female: sumOf_(rows, 'biometric_female')
    };
    return rowCells_(codeVal || '', label, sumEstimate_(rows), sumProvinceEstimate_(rows), v, sumBioTarget_(rows));
  }

  // ==== ល្បឿន៖ ដូចគ្នានឹង exportReportToExcel ខាងលើ — ជួរដេកទិន្នន័យសរសេរជាកញ្ចប់ ជំនួសម្តងមួយជួរដេក (មើលមូលហេតុលម្អិត
  // នៅ writeBulkRowBlock_) ។ ជួរដេក "ក្បាលក្រុមស្រុក" (merge) នៅតែសរសេរដាច់ដោយឡែកដដែល (ចំនួនតិច មិនប៉ះពាល់ល្បឿន) ====
  if (isCommune) {
    var list = communeList.slice().sort(districtCommuneComparator_);
    var lastDistrict = null, pending = [], districtBlockRows_ = [];
    var flushDistrictBlock_ = function(district) {
      if (!pending.length) return;
      var lbl = getDistrictType_(district) === 'ក្រុង' ? 'សរុបក្រុង' : 'សរុបស្រុក';
      districtBlockRows_.push(subtotalCells_(pending, lbl));
      r = writeBulkRowBlock_(sheet, r, districtBlockRows_, totalCols, {
        leftAlignCol: u26Col_, pctCols: pctColsBio_.concat([percent2Col]), boldRowOffsets: [districtBlockRows_.length - 1]
      });
      districtBlockRows_ = [];
      pending = [];
    };
    list.forEach(function(row) {
      if (row.district !== lastDistrict) {
        if (lastDistrict !== null) flushDistrictBlock_(lastDistrict);
        writeDistrictHead_(r, row.district); r++;
        lastDistrict = row.district;
      }
      var cells = rowCells_(
        communeCode_(row.district, row.commune), row.commune, row.stationStats && row.stationStats.estimate,
        row.stationStats && row.stationStats.provinceEstimate, row.values, row.bioTarget
      );
      districtBlockRows_.push(cells);
      pending.push(row);
    });
    if (lastDistrict !== null) flushDistrictBlock_(lastDistrict);
    if (list.length) {
      writeRow_(r, subtotalCells_(list, 'សរុបខេត្ត'), true); r++;
    } else {
      sheet.getRange(r, 1).setValue('គ្មានទិន្នន័យទេ'); r++;
    }
  } else {
    var plist = provinceList.slice().sort(districtOnlyComparator_);
    var provinceBlockRows_ = [];
    plist.forEach(function(row, idx) {
      var cells = rowCells_(
        idx + 1, row.district, row.stationStats && row.stationStats.estimate,
        row.stationStats && row.stationStats.provinceEstimate, row.values, row.bioTarget
      );
      provinceBlockRows_.push(cells);
    });
    if (plist.length) {
      provinceBlockRows_.push(subtotalCells_(plist, 'សរុបខេត្ត', '-'));
      r = writeBulkRowBlock_(sheet, r, provinceBlockRows_, totalCols, {
        leftAlignCol: u26Col_, pctCols: pctColsBio_.concat([percent2Col]), boldRowOffsets: [provinceBlockRows_.length - 1]
      });
    } else {
      sheet.getRange(r, 1).setValue('គ្មានទិន្នន័យទេ'); r++;
    }
  }

  // ==== Auto-resize មុនសរសេរចំណងជើងវែងៗ (ដូចហេតុផលក្នុង exportReportToExcel) ====
  try { sheet.autoResizeColumns(1, totalCols); } catch (err) {}
  writeTitleRows_();
  SpreadsheetApp.flush();

  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + tempFileId_ + '/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(exportUrl, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    return { success: false, message: "ការនាំចេញ Excel បរាជ័យ (កូដ " + response.getResponseCode() + ")! សូមព្យាយាមម្តងទៀត។" };
  }
  var base64 = Utilities.base64Encode(response.getContent());

  var fileNameMap = {
    bioCommune: 'របាយការណ៍ធ្វើជីវមាត្រ_បូកយោង_គ្រប់ឃុំសង្កាត់',
    bioProvince: 'របាយការណ៍ធ្វើជីវមាត្រ_បូកយោង_សរុបខេត្ត'
  };
  var filename = (fileNameMap[reportType] || 'របាយការណ៍ធ្វើជីវមាត្រ') + '_' + (toDate || formatNow_().slice(0, 10)) + '.xlsx';

  return { success: true, base64: base64, filename: filename };
  } catch (err) {
    return { success: false, message: "កំហុសពេលនាំចេញ Excel៖ " + err.message };
  } finally {
    try { DriveApp.getFileById(tempFileId_).setTrashed(true); } catch (e) {}
  }
}
