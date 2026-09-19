// ==================== Dashboard.gs — ស្ថិតិសម្រាប់ទិដ្ឋភាពទូទៅ ====================

function getDashboardStats(currentUsername, sessionToken) {
  var __sess = requireValidSession_(currentUsername, sessionToken);
  if (!__sess.valid) return { success: false, message: __sess.message, sessionInvalid: true };

  var role = getUserRole_(currentUsername);
  if (!role) return { success: false, message: "គណនីមិនត្រឹមត្រូវទេ!" };
  var viewAll = canViewAllDistricts_(currentUsername);
  var myDistrict = getUserDistrict_(currentUsername);
  // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ គណនីឃុំសង្កាត់ ក៏ត្រូវឲ្យស្កេនស្រុកខ្លួន (ដើម្បីរកទិន្នន័យឃុំសង្កាត់ខ្លួន)
  // ដូចគ្នានឹងគ្រូប្រចាំក្រុងស្រុកដែរ ប៉ុន្តែខាងក្រោម (ក្នុងចំណុចប្រមូលផ្តុំទិន្នន័យ) នឹងច្រោះឲ្យរាប់បញ្ចូលតែជួរដេកនៃ
  // ឃុំសង្កាត់ខ្លួនប៉ុណ្ណោះ (មិនមែនទាំងស្រុក) — ជៀសវាងលេខសរុបកាតទាំងអស់បញ្ចេញឲ្យដឹងពីទិន្នន័យសរុបរបស់ឃុំសង្កាត់ដទៃ ====
  var isCommuneCaller_ = isCommuneRole_(role);
  var myCommune_ = isCommuneCaller_ ? getUserCommune_(currentUsername) : null;
  var districtsToRead = viewAll ? DISTRICT_LIST : ((isDistrictRole_(role) || isCommuneCaller_) ? [myDistrict] : []);
  // ==== ជំនួយសម្រាប់រាល់កន្លែងដែលដួល COMMUNE_ORDER[d] ខាងក្រោម — គណនីឃុំសង្កាត់ ឃើញតែឃុំសង្កាត់ខ្លួនប៉ុណ្ណោះ ====
  function communesInScope_(d) {
    var all = COMMUNE_ORDER[d] || [];
    if (!isCommuneCaller_) return all;
    return all.filter(function(c) { return c.name === myCommune_; });
  }

  // ==== ល្បឿន៖ ចងចាំលទ្ធផលរយៈពេលខ្លី (៦០វិនាទី) សម្រាប់វិសាលភាពដូចគ្នា ជៀសវាងស្កេនគ្រប់ Sheet ឡើងវិញរាល់ពេលចូល Tab នេះ —
  // FIX (សំណើថ្មី "គណនីឃុំសង្កាត់" — សំខាន់ណាស់, ការពារ Cache ចម្រុះគ្នា)៖ ត្រូវបញ្ចូល myCommune_ ចូល Cache Key ដែរ
  // ព្រោះគណនីឃុំសង្កាត់ ឥឡូវត្រឡប់លទ្ធផលខុសពីគ្រូប្រចាំក្រុងស្រុក (ដែលឃើញទាំងស្រុក) ថ្វីត្បិតតែស្ថិតក្នុងស្រុកតែមួយ
  // (myDistrict) ដូចគ្នាក៏ដោយ — បើប្រើ Key ដូចគ្នា គណនីឃុំសង្កាត់អាចទទួលបានទិន្នន័យសរុបទាំងស្រុករបស់ស្រុកនោះ (Cache
  // ដែលគ្រូប្រចាំក្រុងស្រុកទុកមកមុន) ដោយច្រឡំ — ជាការលេចធ្លាយទិន្នន័យសំខាន់ ====
  var cacheKey = "dashStats_v7_" + (viewAll ? "all" : (myDistrict + (isCommuneCaller_ ? ("|" + myCommune_) : ""))); // v7៖ បន្ថែម provinceEstimateTotal + ២ភាគរយថ្មីធៀបនឹង "ស្ថិតិប៉ាន់ស្មានខេត្ត" (v6 មុន៖ កែតម្រូវ Card#7 ឲ្យរាប់បញ្ចូលទិន្នន័យយោងគ្រប់ឃុំ/សង្កាត់)
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (err) {}

  var tz = Session.getScriptTimeZone() || 'Asia/Phnom_Penh';
  var curYearMonth = Utilities.formatDate(new Date(), tz, 'yyyy-MM'); // ខែបច្ចុប្បន្ន (យោងតាមម៉ោង Server)
  var daysInMonth = new Date(Number(curYearMonth.slice(0, 4)), Number(curYearMonth.slice(5, 7)), 0).getDate();
  var dailyAddT = new Array(daysInMonth).fill(0), dailyAddF = new Array(daysInMonth).fill(0);
  var dailyAddNewT = new Array(daysInMonth).fill(0), dailyAddNewF = new Array(daysInMonth).fill(0);
  var dailyDelT = new Array(daysInMonth).fill(0), dailyDelF = new Array(daysInMonth).fill(0);
  var dailyUpd26T = new Array(daysInMonth).fill(0), dailyUpd26F = new Array(daysInMonth).fill(0);
  var dailyBioT = new Array(daysInMonth).fill(0), dailyBioF = new Array(daysInMonth).fill(0);
  var dailyNewPlusTrInT = new Array(daysInMonth).fill(0), dailyNewPlusTrInF = new Array(daysInMonth).fill(0);
  var addTSum = 0, addFSum = 0, addNewTSum = 0, addNewFSum = 0, delTSum = 0, delFSum = 0;
  var addTrInTSum = 0, addTrInFSum = 0; // ចំនួនផ្ទេរចូលសរុប (ដាច់ដោយឡែកពី "ថ្មីសរុប+ផ្ទេរចូលសរុប" newPlusTrInTSum ខាងក្រោម)
  var bioTSum = 0, bioFSum = 0, newPlusTrInTSum = 0, newPlusTrInFSum = 0;
  // ==== កាតថ្មី៥ (v5)៖ ក្រុមរងនៃ "លុបចេញពីបញ្ជីបោះឆ្នោត" (ស្លាប់ដកសិទ្ធិផ្លាស់ចេញ/ផ្ទេរចេញក្រៅឃុំសង្កាត់) បូក
  // "ករណីកែតម្រូវ"/"ឯ.អ បានចេញ"/"លិខិតបញ្ជាក់ទីលំនៅ" ដែលមានស្រាប់ក្នុង DAILY_FIELDS ស្រាប់ (correction_*/idTotal_*/residenceCert_*) ====
  var deathTSum = 0, deathFSum = 0, transferOutTSum = 0, transferOutFSum = 0;
  var correctionTSum = 0, correctionFSum = 0, idIssuedTSum = 0, idIssuedFSum = 0, residenceCertTSum = 0, residenceCertFSum = 0;
  var communeCumulative = {}; // "ស្រុក|ឃុំ" -> { addT, addF, delT, delF } — សម្រាប់គណនា "បច្ចុប្បន្នភាព២០២៦" ដោយមិនចាំបាច់ស្កេនម្តងទៀត
  // ==== FIX (v6)៖ ត្រូវចាប់ផ្តើមបំពេញជាមុន "ស្រុក|ឃុំ" ទាំងអស់ក្នុងវិសាលភាព (districtsToRead × COMMUNE_ORDER) មុនស្កេន
  // DailyIndex — ដូចគ្នានឹង computeSummaries_() ក្នុង Reports.gs (ដែលបំពេញ communeMap ជាមុនសម្រាប់ឃុំ/សង្កាត់ទាំងអស់
  // ក្នុងជួរទី ៨៦-៩៦) ។ បើមិនធ្វើដូច្នេះទេ ឃុំ/សង្កាត់ណាដែលមិនទាន់មានអ្នកបញ្ចូលទិន្នន័យប្រចាំថ្ងៃសោះ (hasData មិនដែលពិត)
  // នឹងមិនចូលរួមក្នុង communeCumulative ទាល់តែសោះ ធ្វើឲ្យទិន្នន័យយោង (baseline2025) របស់ឃុំ/សង្កាត់នោះទាំងមូល
  // ត្រូវបានលុបចោលពី "ចំនួនកាតបច្ចុប្បន្នភាពបញ្ជីបោះឆ្នោតឆ្នាំ២០២៦" (Card #7) ធ្វើឲ្យលេខតូចជាងការពិតយ៉ាងខ្លាំង
  // (ដូចករណីដែលរកឃើញ៖ Card បង្ហាញ ៧៥,៣៣៧ ខណៈរបាយការណ៍សរុបខេត្តពិតប្រាកដមាន ៦៩៥,៨៥៣) ====
  districtsToRead.forEach(function(d) {
    communesInScope_(d).forEach(function(c) {
      communeCumulative[d + "|" + c.name] = { addT: 0, addF: 0, delT: 0, delF: 0 };
    });
  });

  // ==== ល្បឿន៖ អាន DailyIndex (Sheet មួយក្នុង Spreadsheet មេ) ជំនួសការបើក Spreadsheet ស្រុកទាំង១០ ដោយផ្ទាល់ ដូចមុន
  // (ជាការផ្លាស់ប្តូរធំបំផុតសម្រាប់ល្បឿន Dashboard — ពីមុនត្រូវបើក Spreadsheet ១០ស្រុក ហើយអាន Sheet រាល់ថ្ងៃម្នាក់ៗ
  // ដែលអាចមានរាប់រយ Sheet សរុប ប៉ុន្តែឥឡូវអានតែ Sheet មួយក្នុង Spreadsheet មេ) — DailyIndex ត្រូវបានរក្សាទុកឲ្យ
  // ទាន់សម័យជានិច្ចរួចស្រាប់ (មើល fastReindexDistrictDay_/rebuildAllDerivedData_ ក្នុង PeriodSheets.gs) ដូច្នេះ
  // ទុកចិត្តបានពេញលេញ ដូចគ្នានឹង getDailyEntries() ប្រើសម្រាប់តារាងខាងក្រោមទំព័របញ្ចូលទិន្នន័យដែរ ====
  var districtSet_ = {};
  districtsToRead.forEach(function(d) { districtSet_[d] = true; });
  var dailyIndexMap = loadDailyIndexOnce_();
  Object.keys(dailyIndexMap).forEach(function(key) {
    var sep = key.indexOf('|');
    if (sep < 0) return;
    var d = key.slice(0, sep);
    if (!districtSet_[d]) return;
    var dateStr = key.slice(sep + 1);
    var entries = dailyIndexMap[key] || [];
    entries.forEach(function(entry) {
      if (!entry.hasData) return;
      // ==== FIX (សំណើថ្មី "គណនីឃុំសង្កាត់")៖ Server-side Scoping — Defense in depth, មិនទុកចិត្តតែ Client Filter ទេ —
      // គណនីឃុំសង្កាត់ រាប់បញ្ចូលតែជួរដេកនៃឃុំសង្កាត់ខ្លួនប៉ុណ្ណោះ ចូលទៅក្នុងលេខសរុបកាតទាំងអស់ ====
      if (isCommuneCaller_ && entry.commune !== myCommune_) return;
      var v = entry.values || {};
      var addT = Number(v.addTotal_total) || 0;
      var addF = Number(v.addTotal_female) || 0;
      var addNewT = Number(v.addNew_total) || 0;
      var addNewF = Number(v.addNew_female) || 0;
      var addTrInT = Number(v.addTransferIn_total) || 0;
      var addTrInF = Number(v.addTransferIn_female) || 0;
      var bioT = Number(v.biometric_total) || 0;
      var bioF = Number(v.biometric_female) || 0;
      var delT = Number(v.delTotal_total) || 0;
      var delF = Number(v.delTotal_female) || 0;
      var upd26T = Number(v.update2026_total) || 0;
      var upd26F = Number(v.update2026_female) || 0;
      var newPlusTrInT = addNewT + addTrInT, newPlusTrInF = addNewF + addTrInF;
      addTSum += addT; addFSum += addF; addNewTSum += addNewT; addNewFSum += addNewF; delTSum += delT; delFSum += delF;
      bioTSum += bioT; bioFSum += bioF; newPlusTrInTSum += newPlusTrInT; newPlusTrInFSum += newPlusTrInF;
      addTrInTSum += addTrInT; addTrInFSum += addTrInF;
      // ==== កាតថ្មី៥ (v5) ====
      deathTSum += Number(v.delDeath_total) || 0; deathFSum += Number(v.delDeath_female) || 0;
      transferOutTSum += Number(v.delTransferOut_total) || 0; transferOutFSum += Number(v.delTransferOut_female) || 0;
      correctionTSum += Number(v.correction_total) || 0; correctionFSum += Number(v.correction_female) || 0;
      idIssuedTSum += Number(v.idTotal_total) || 0; idIssuedFSum += Number(v.idTotal_female) || 0;
      residenceCertTSum += Number(v.residenceCert_total) || 0; residenceCertFSum += Number(v.residenceCert_female) || 0;

      var commune = entry.commune;
      var cKey = d + "|" + commune;
      if (!communeCumulative[cKey]) communeCumulative[cKey] = { addT: 0, addF: 0, delT: 0, delF: 0 };
      communeCumulative[cKey].addT += addT; communeCumulative[cKey].addF += addF;
      communeCumulative[cKey].delT += delT; communeCumulative[cKey].delF += delF;

      var isoDate = normalizeDateStr_(dateStr);
      var m = isoDate.match(/^(\d{4}-\d{2})-(\d{2})$/);
      if (m && m[1] === curYearMonth) {
        var di = Number(m[2]) - 1;
        dailyAddT[di] += addT; dailyAddF[di] += addF;
        dailyAddNewT[di] += addNewT; dailyAddNewF[di] += addNewF;
        dailyBioT[di] += bioT; dailyBioF[di] += bioF;
        dailyNewPlusTrInT[di] += newPlusTrInT; dailyNewPlusTrInF[di] += newPlusTrInF;
        dailyDelT[di] += delT; dailyDelF[di] += delF;
        dailyUpd26T[di] += upd26T; dailyUpd26F[di] += upd26F; // Snapshot ត្រឹមត្រូវរួចស្រាប់ (បូកឆ្លងកាត់ឃុំ/សង្កាត់ក្នុងថ្ងៃតែមួយ ជាការត្រឹមត្រូវ)
      }
    });
  });

  var dayLabels = []; for (var di2 = 1; di2 <= daysInMonth; di2++) dayLabels.push(String(di2));

  // ==== "បច្ចុប្បន្នភាព២០២៦" (កាតសរុប) = ទិន្នន័យ២០២៥ + បូក/ដកសរុបគ្រប់ថ្ងៃ ក្នុងមួយឃុំ/សង្កាត់ៗ (ដូចគ្នានឹង computeSummaries_)
  // ប្រើទិន្នន័យដែលបានស្កេនរួចស្រាប់ខាងលើ (per-commune) ជៀសវាងស្កេន Sheet ដដែលៗម្តងទៀត (លឿនជាងច្រើន)
  var baseline2025Map = getBaseline2025Map_();
  var upd26TSumFixed = 0, upd26FSumFixed = 0;
  Object.keys(communeCumulative).forEach(function(key) {
    var cum = communeCumulative[key];
    var base = baseline2025Map[key] || { total: 0, female: 0 };
    upd26TSumFixed += (Number(base.total) || 0) + cum.addT - cum.delT;
    upd26FSumFixed += (Number(base.female) || 0) + cum.addF - cum.delF;
  });

  // ==== ភាគរយ "ធ្វើជីវមាត្រ" = ចំនួនករណីជីវមាត្រពិតប្រាកដ (bioTSum/bioFSum ខាងលើ) ÷ ទិន្នន័យយោង (ពី Tab "ទិន្នន័យភ្ជាប់") ====
  // ទិន្នន័យយោង ជាតួលេខកំណត់ជាមុន (មិនប្រែប្រួលតាមកាលបរិច្ឆេទ) — បូកសរុបតាមវិសាលភាពគណនីអ្នកប្រើប្រាស់ (districtsToRead) ដូចគ្នា
  var biometricTargetMap = getBiometricTargetMap_();
  var bioTargetTSum = 0, bioTargetFSum = 0;
  districtsToRead.forEach(function(d) {
    communesInScope_(d).forEach(function(c) {
      var t = biometricTargetMap[d + "|" + c.name];
      if (!t) return;
      bioTargetTSum += Number(t.total) || 0;
      bioTargetFSum += Number(t.female) || 0;
    });
  });
  var bioPercentTotal = bioTargetTSum > 0 ? Math.round((bioTSum / bioTargetTSum) * 1000) / 10 : null; // null = គ្មានទិន្នន័យយោង
  var bioPercentFemale = bioTargetFSum > 0 ? Math.round((bioFSum / bioTargetFSum) * 1000) / 10 : null;

  // ==== ភាគរយ "ចុះឈ្មោះថ្មីសរុប + ផ្ទេរចូលសរុប" ធៀបនឹង "ស្ថិតិប៉ាន់ស្មាន" (Tab ទិន្នន័យភ្ជាប់) ====
  // រូបមន្ត = (ចុះឈ្មោះថ្មីសរុប + ផ្ទេរចូលសរុប) × ១០០ ÷ ស្ថិតិប៉ាន់ស្មាន — ដូចគ្នានឹងជួរឈរ "ភាគរយ" ក្នុង Tab របាយការណ៍សម្រាប់បោះពុម្ព (pct_ ក្នុង Reports.gs)
  var stationStatsMap = getStationStatsMap_();
  var estimateTSum = 0, provinceEstimateTSum = 0;
  districtsToRead.forEach(function(d) {
    communesInScope_(d).forEach(function(c) {
      var s = stationStatsMap[d + "|" + c.name];
      if (!s) return;
      estimateTSum += Number(s.estimate) || 0;
      provinceEstimateTSum += Number(s.provinceEstimate) || 0;
    });
  });
  var newPlusTrInPercentTotal = estimateTSum > 0 ? Math.round((newPlusTrInTSum / estimateTSum) * 10000) / 100 : null; // null = គ្មានស្ថិតិប៉ាន់ស្មាន គ.ជ.ប
  // ==== ភាគរយ "ចុះឈ្មោះថ្មីសរុប" (មិនរួមផ្ទេរចូល) ធៀបនឹង "ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប" — ដាច់ដោយឡែកពី newPlusTrInPercentTotal ខាងលើ ====
  var newOnlyPercentTotal = estimateTSum > 0 ? Math.round((addNewTSum / estimateTSum) * 10000) / 100 : null; // null = គ្មានស្ថិតិប៉ាន់ស្មាន គ.ជ.ប
  // ==== កាតថ្មី (v7)៖ ២ភាគរយថ្មី ធៀបនឹង "ស្ថិតិប៉ាន់ស្មានខេត្ត" (ដូចគ្នានឹង២ខាងលើ ប៉ុន្តែប្តូរភាគបែង) ====
  var newPlusTrInPercentProvinceTotal = provinceEstimateTSum > 0 ? Math.round((newPlusTrInTSum / provinceEstimateTSum) * 10000) / 100 : null; // null = គ្មានស្ថិតិប៉ាន់ស្មានខេត្ត
  var newOnlyPercentProvinceTotal = provinceEstimateTSum > 0 ? Math.round((addNewTSum / provinceEstimateTSum) * 10000) / 100 : null; // null = គ្មានស្ថិតិប៉ាន់ស្មានខេត្ត

  var result = {
    success: true,
    addTotal: addTSum, addFemale: addFSum,
    delTotal: delTSum, delFemale: delFSum,
    netAddTotal: addNewTSum, netAddFemale: addNewFSum, // ==== "ចំនួនអ្នកចុះឈ្មោះថ្មី" = ក្រុមរង "ចុះឈ្មោះថ្មី" ជាក់លាក់ (មិនមែនបូក-ដកទេ) ====
    transferInTotal: addTrInTSum, transferInFemale: addTrInFSum, // ចំនួនផ្ទេរចូលសរុប (ក្រុមរង "ផ្ទេរចូល" ជាក់លាក់)
    update2026Total: upd26TSumFixed, update2026Female: upd26FSumFixed,
    bioTotal: bioTSum, bioFemale: bioFSum, // ចំនួនករណីបច្ចុប្បន្នភាពជីវមាត្រ
    bioTargetTotal: bioTargetTSum, bioTargetFemale: bioTargetFSum, // ទិន្នន័យយោង (ត្រូវធ្វើជីវមាត្រ) ពី Tab "ទិន្នន័យភ្ជាប់"
    bioPercentTotal: bioPercentTotal, bioPercentFemale: bioPercentFemale, // ភាគរយ = ពិតប្រាកដ ÷ យោង × ១០០ (null បើគ្មានទិន្នន័យយោង)
    newPlusTrInTotal: newPlusTrInTSum, newPlusTrInFemale: newPlusTrInFSum, // ចុះឈ្មោះថ្មីសរុប + ផ្ទេរចូលសរុប
    estimateTotal: estimateTSum, // ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប (Tab ទិន្នន័យភ្ជាប់) — ភាគបែងសម្រាប់ newPlusTrInPercentTotal/newOnlyPercentTotal
    newOnlyPercentTotal: newOnlyPercentTotal, // ភាគរយ = ចុះឈ្មោះថ្មីសរុប × ១០០ ÷ ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប (null បើគ្មាន)
    newPlusTrInPercentTotal: newPlusTrInPercentTotal, // ភាគរយ = (ថ្មីសរុប+ផ្ទេរចូលសរុប) × ១០០ ÷ ស្ថិតិប៉ាន់ស្មាន គ.ជ.ប (null បើគ្មាន)
    // ==== កាតថ្មី (v7) ====
    provinceEstimateTotal: provinceEstimateTSum, // ស្ថិតិប៉ាន់ស្មានខេត្ត (Tab ទិន្នន័យភ្ជាប់) — ភាគបែងសម្រាប់ ២ភាគរយថ្មីខាងក្រោម
    newOnlyPercentProvinceTotal: newOnlyPercentProvinceTotal, // ភាគរយ = ចុះឈ្មោះថ្មីសរុប × ១០០ ÷ ស្ថិតិប៉ាន់ស្មានខេត្ត (null បើគ្មាន)
    newPlusTrInPercentProvinceTotal: newPlusTrInPercentProvinceTotal, // ភាគរយ = (ថ្មីសរុប+ផ្ទេរចូលសរុប) × ១០០ ÷ ស្ថិតិប៉ាន់ស្មានខេត្ត (null បើគ្មាន)
    // ==== កាតថ្មី៥ (v5) ====
    deathTotal: deathTSum, deathFemale: deathFSum, // ចំនួនស្លាប់ ដកសិទ្ធិ ផ្លាស់ចេញសរុប
    transferOutTotal: transferOutTSum, transferOutFemale: transferOutFSum, // ចំនួនផ្ទេរចេញក្រៅឃុំសង្កាត់សរុប
    correctionTotal: correctionTSum, correctionFemale: correctionFSum, // ចំនួនករណីកែតម្រូវសរុប
    idIssuedTotal: idIssuedTSum, idIssuedFemale: idIssuedFSum, // ចំនួន (ឯ.អ) បានចេញសរុប
    residenceCertTotal: residenceCertTSum, residenceCertFemale: residenceCertFSum, // ចំនួនលិខិតបញ្ជាក់ទីលំនៅសរុប
    days: dayLabels, curMonth: curYearMonth,
    dailyAdd: dailyAddT, dailyAddFemale: dailyAddF,
    dailyDel: dailyDelT, dailyDelFemale: dailyDelF,
    dailyNet: dailyAddNewT, dailyNetFemale: dailyAddNewF,
    dailyUpdate2026: dailyUpd26T, dailyUpdate2026Female: dailyUpd26F,
    dailyBio: dailyBioT, dailyBioFemale: dailyBioF,
    dailyNewPlusTrIn: dailyNewPlusTrInT, dailyNewPlusTrInFemale: dailyNewPlusTrInF
  };
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 60); } catch (err) {}
  return result;
}

// ==== ទិន្នន័យបម្រុងទុក (Backup) — ចម្លង Spreadsheet មេទាំងមូល ====
