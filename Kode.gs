/**
 * =========================================================================
 * BACKEND REPORT GOLIVE PT3 - GOOGLE APPS SCRIPT
 * =========================================================================
 * File ini siap dipakai langsung di Google Sheets:
 * 1. Buka spreadsheet sumber -> Extensions -> Apps Script
 * 2. Hapus kode default, tempel seluruh isi file ini ke 'Code.gs'
 * 3. Jalankan setupLopHistory sekali untuk membuat spreadsheet history terpisah.
 *    Lalu pilih fungsi 'getReport' pada toolbar atas, lalu klik 'Run' (Jalankan)
 *    untuk memberikan izin akses spreadsheet.
 * 4. Klik Deploy -> New deployment -> Web app:
 *    - Description : Report Golive PT3 Live
 *    - Execute as  : Me (email Anda)
 *    - Who has access: Anyone (Siapa saja yang memiliki link)
 * 5. Selesai! URL Web App dapat dibuka langsung di browser sebagai Dashboard,
 *    atau diakses sebagai API JSON (?api=1).
 * =========================================================================
 */

function getNodeBReport() {
  const ss = SpreadsheetApp.openById('1D6StHSC4cWCZLbAImb8EFjXwZ61-wwTQFkmMzstZJUY');
  const sheet = ss.getSheets().find(s => s.getSheetId() === 0);
  if (!sheet) throw new Error('Sheet NODE-B gid=0 tidak ditemukan.');
  const timezone = ss.getSpreadsheetTimeZone();
  // Read actual date cells, not ambiguous localized display strings.
  const values = sheet.getDataRange().getValues().map(row => row.map(v => v instanceof Date ? Utilities.formatDate(v, timezone, 'yyyy-MM-dd') : v));
  const report = nodeBReport_(values);
  report.curve = nodeBCurve_(values);
  return report;
}

function nodeBCurve_(values) {
  const clean = v => String(v == null ? '' : v).trim().toUpperCase();
  const h = values.findIndex(r => r.some(v => clean(v) === 'SITE ID'));
  if (h < 0) throw new Error('Kolom SITE ID tidak ditemukan.');
  const header = values[h].map(clean);
  const cols = ['SITE ID','KOMITMEN OA','REALISASI OA'].map(name => {
    const i = header.indexOf(name);
    if (i < 0) throw new Error('Kolom ' + name + ' tidak ditemukan.');
    return i;
  });
  function dateKey(value) {
    const s = clean(value);
    let m, y, month, day;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) {y=+m[1];month=+m[2];day=+m[3];}
    else if ((m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/.exec(s))) {day=+m[1];month=+m[2];y=+m[3];}
    else return '';
    const d = new Date(Date.UTC(y,month-1,day));
    if (d.getUTCFullYear() !== y || d.getUTCMonth() !== month-1 || d.getUTCDate() !== day) return '';
    return d.toISOString().slice(0,10);
  }
  const days = new Map();
  const invalid = [0,0];
  values.slice(h+1).forEach(row => {
    if (!clean(row[cols[0]])) return;
    cols.slice(1).forEach((col,i) => {
      if (!clean(row[col])) return;
      const key = dateKey(row[col]);
      if (!key) {invalid[i]++;return;}
      if (!days.has(key)) days.set(key,{planDaily:0,actualDaily:0});
      days.get(key)[i ? 'actualDaily' : 'planDaily']++;
    });
  });
  let plan=0, actual=0;
  const points = Array.from(days.keys()).sort().map(date => {
    const d=days.get(date); plan+=d.planDaily;actual+=d.actualDaily;
    return {date:date,plan:plan,actual:actual,planDaily:d.planDaily,actualDaily:d.actualDaily};
  });
  return {points:points,invalidPlan:invalid[0],invalidActual:invalid[1]};
}

function nodeBReport_(values) {
  const text = v => String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
  const header = values.findIndex(r => r.some(v => text(v).toUpperCase() === 'SITE ID'));
  if (header < 0) throw new Error('Kolom SITE ID tidak ditemukan.');
  const names = values[header].map(v => text(v).toUpperCase());
  const cols = ['SITE ID', 'STATUS', 'PROGRAM'].map(n => {
    const i = names.indexOf(n);
    if (i < 0) throw new Error('Kolom ' + n + ' tidak ditemukan.');
    return i;
  });
  const stages = ['Drop','Aanwijzing','Perizinan','Matdel','Instalasi','Finish Install','On Air','Uji Terima'];
  const aliases = {'drop':'Drop','hold':'Drop','plan drop':'Drop','aanwijzing':'Aanwijzing','perizinan':'Perizinan','matdel':'Matdel','material delivery':'Matdel','material preparation':'Matdel','instalasi':'Instalasi','finish install':'Finish Install','finish instalasi':'Finish Install','on air':'On Air','rfs':'On Air','uji terima':'Uji Terima'};
  const programs = new Map(), counts = new Map();
  let order = 0, deployment = 0, onAir = 0;
  values.slice(header + 1).forEach(row => {
    if (!text(row[cols[0]])) return;
    order++;
    const raw = text(row[cols[1]]).replace(/^\d+\s*[.\-:)]?\s*/, '').toLowerCase();
    const status = aliases[raw] || raw.toUpperCase() || 'Tanpa Status';
    counts.set(status, (counts.get(status) || 0) + 1);
    if (stages.slice(4).includes(status)) deployment++;
    if (status === 'On Air') onAir++;
    const name = text(row[cols[2]]);
    if (!name) return;
    const key = name.toLowerCase();
    if (!programs.has(key)) programs.set(key, {name:name,order:0,closed:0});
    const p = programs.get(key);
    p.order++;
    if (status === 'On Air') p.closed++;
  });
  const percent = n => order ? n / order * 100 : 0;
  return {
    ok:true, order:order, deployment:deployment, onAir:onAir, gap:order-onAir,
    updatedAt:new Date().toISOString(),
    programs:Array.from(programs.values()).map(p => Object.assign(p,{open:p.order-p.closed,progress:p.closed/p.order*100})).sort((a,b)=>a.name.localeCompare(b.name)),
    milestones:stages.concat(Array.from(counts.keys()).filter(s=>!stages.includes(s)).sort()).filter(s=>counts.has(s)).map(name=>({name:name,total:counts.get(name),percentage:percent(counts.get(name))}))
  };
}

const CONFIG = {
  sourceSpreadsheetId: '1FdDuiHUKvRU7VMLydb5cZ5FNwU7Mqxp1zb2mAsTShS0',
  sheetName: 'Sheet1',               // Nama sheet sumber (otomatis fallback jika tidak ditemukan)
  pt2SheetName: 'LIST FI PT2',
  timezone: 'Asia/Jakarta',          // Zona waktu WIB
  cacheSeconds: 7200,                // Cache dashboard selama 2 jam
  csvUrl: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRF9glw8zA3bP8YQ3P0LfB-mbh_oWc2ehESfI4sG4jkTrI68dv7uVNOzVrhlh35czWGuaLmUX52IcNC/pub?output=csv&gid=0',
  pt2CsvUrl: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRF9glw8zA3bP8YQ3P0LfB-mbh_oWc2ehESfI4sG4jkTrI68dv7uVNOzVrhlh35czWGuaLmUX52IcNC/pub?output=csv&gid=311076524',
  headerNames: {
    status: 'STATUS',
    port: 'REAL PORT',
    region: 'REGION FMC',
    goliveDate: 'TANGGAL GOLIVE',
    komitmenDate: 'KOMITMEN GOLIVE'
  }
};

const DASHBOARD_CACHE_CHUNK_SIZE = 70000;
const DASHBOARD_CACHE_BUILD_LEASE_SECONDS = 120;

/** Cache JSON bertingkat agar payload >100 KB tetap dapat disimpan dengan aman. */
function dashboardCacheRead_(key) {
  const cache = CacheService.getScriptCache();
  const rawMeta = cache.get(key + ':meta');
  if (!rawMeta) return null;
  try {
    const meta = JSON.parse(rawMeta);
    if (!meta || !meta.token || !meta.parts) return null;
    const chunks = [];
    for (let i = 0; i < meta.parts; i++) {
      const part = cache.get(key + ':' + meta.token + ':' + i);
      if (part === null) return null;
      chunks.push(part);
    }
    const json = chunks.join('');
    if (meta.length !== json.length) return null;
    return JSON.parse(json);
  } catch (err) {
    console.warn('Cache JSON tidak dapat dibaca: ' + err.message);
    return null;
  }
}

function dashboardCacheWrite_(key, value) {
  const cache = CacheService.getScriptCache();
  const json = JSON.stringify(value);
  const oldMetaRaw = cache.get(key + ':meta');
  const token = new Date().getTime().toString(36) + '-' + Math.floor(Math.random() * 1000000).toString(36);
  const parts = Math.max(1, Math.ceil(json.length / DASHBOARD_CACHE_CHUNK_SIZE));
  for (let i = 0; i < parts; i++) {
    cache.put(key + ':' + token + ':' + i, json.slice(i * DASHBOARD_CACHE_CHUNK_SIZE, (i + 1) * DASHBOARD_CACHE_CHUNK_SIZE), CONFIG.cacheSeconds);
  }
  cache.put(key + ':meta', JSON.stringify({token:token, parts:parts, length:json.length, cachedAt:new Date().toISOString()}), CONFIG.cacheSeconds);
  cache.remove(key + ':building');
  try {
    const oldMeta = oldMetaRaw ? JSON.parse(oldMetaRaw) : null;
    if (oldMeta && oldMeta.token && oldMeta.token !== token) {
      for (let j = 0; j < (oldMeta.parts || 0); j++) cache.remove(key + ':' + oldMeta.token + ':' + j);
    }
  } catch (ignore) {}
}

/**
 * Mencegah cache stampede tanpa menahan ScriptLock selama pembacaan Sheets.
 * Lock hanya dipakai singkat untuk menetapkan lease; proses history tetap bebas
 * memakai ScriptLock-nya sendiri.
 */
function dashboardCacheGate_(key, forceRefresh) {
  if (!forceRefresh) {
    const cached = dashboardCacheRead_(key);
    if (cached) return {value:cached, owner:false};
  }
  const cache = CacheService.getScriptCache();
  const leaseKey = key + ':building';
  let owner = false;
  let lock = null;
  try {
    lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
      if (!forceRefresh) {
        const secondRead = dashboardCacheRead_(key);
        if (secondRead) return {value:secondRead, owner:false};
      }
      if (!cache.get(leaseKey)) {
        cache.put(leaseKey, new Date().toISOString(), DASHBOARD_CACHE_BUILD_LEASE_SECONDS);
        owner = true;
      }
    }
  } finally {
    if (lock && lock.hasLock()) lock.releaseLock();
  }
  if (!owner) {
    for (let i = 0; i < 12; i++) {
      Utilities.sleep(250);
      const awaited = dashboardCacheRead_(key);
      if (awaited) return {value:awaited, owner:false};
    }
  }
  return {value:null, owner:true};
}

/** Jalankan sekali untuk memasang trigger pemanas cache setiap dua jam. */
function setupTwoHourlyPreload() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'refreshDashboardPreloadCache') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('refreshDashboardPreloadCache').timeBased().everyHours(2).create();
  return refreshDashboardPreloadCache();
}

/** Handler trigger: memperbarui cache default seluruh tab sebelum pengguna membuka dashboard. */
function refreshDashboardPreloadCache() {
  const result = {ok:true, refreshedAt:new Date().toISOString(), report:false, combined:false, errors:[]};
  try { getReport_({forceRefresh:true}); result.report = true; } catch (err) { result.errors.push('PT3: ' + err.message); }
  try { getCombinedGoliveReport_({forceRefresh:true}); result.combined = true; } catch (err) { result.errors.push('PT2+PT3: ' + err.message); }
  result.ok = result.report && result.combined;
  if (!result.ok) console.warn('Preload cache tidak lengkap: ' + result.errors.join(' | '));
  return result;
}

/**
 * Endpoint HTTP GET Web App
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  const isApi = params.api === '1' ||
                params.format === 'json' ||
                params.json === '1' ||
                params.view === 'json' ||
                Boolean(params.callback);

  // Jika request ditujukan sebagai API data (JSON / JSONP)
  if (isApi) {
    let payload;
    try {
      payload = getReport_();
    } catch (error) {
      payload = {
        ok: false,
        error: String((error && error.message) || error)
      };
    }
    const callback = safeCallback_(params.callback);
    const body = callback ? `${callback}(${JSON.stringify(payload)});` : JSON.stringify(payload);
    return ContentService.createTextOutput(body)
      .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }

  // Tampilkan antarmuka HTML Dashboard.
  // Index.html adalah HTML murni (tanpa scriptlet), sehingga createHtmlOutputFromFile
  // sudah tepat dan paling cepat. Bila file Index tidak ada, tampilkan pesan yang jelas
  // alih-alih gagal dengan error samar.
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Monitoring Golive PT-3 Sumatera')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(getFallbackHtml_(err))
      .setTitle('Monitoring Golive PT-3 Sumatera')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Halaman darurat bila file Index.html belum dibuat di project Apps Script.
 * Menjelaskan langkah perbaikannya, sekaligus tetap menyediakan tautan data mentah.
 */
function getFallbackHtml_(err) {
  const detail = String((err && err.message) || err || 'File Index.html tidak ditemukan');
  return `<!doctype html><html lang="id"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>body{font:14px/1.6 system-ui,Segoe UI,Arial,sans-serif;max-width:680px;margin:48px auto;padding:0 20px;color:#102a43}
h1{font-size:19px}code{background:#eef3f0;padding:2px 5px;border-radius:4px}
.box{border:1px solid #dbe5df;border-left:4px solid #c51f3d;border-radius:10px;padding:14px 16px;background:#f8faf9;margin-top:16px}
a{color:#087249}</style></head><body>
<h1>Dashboard belum siap ditampilkan</h1>
<p><b>Penyebab:</b> ${detail.replace(/[<>]/g, '')}</p>
<div class="box"><b>Cara memperbaiki:</b><ol>
<li>Buka project Apps Script ini, klik <b>+</b> &rarr; <b>HTML</b>, beri nama tepat <code>Index</code>.</li>
<li>Tempel seluruh isi file <code>appscript/Index.html</code> ke sana, lalu simpan (Ctrl+S).</li>
<li>Deploy &rarr; Manage deployments &rarr; Edit &rarr; Version: <b>New version</b> &rarr; Deploy.</li>
</ol></div>
<p>Data mentah tetap bisa diakses di
<a href="?api=1">?api=1</a> (JSON) setelah deployment diperbarui.</p>
</body></html>`;
}

/**
 * Fungsi publik untuk dipanggil oleh frontend (google.script.run.getReport())
 * dan untuk pengujian izin akses awal di editor toolbar Apps Script.
 */
function getReport(filters) {
  return getReport_(filters);
}

/**
 * Ringkasan lintas program untuk tab MONITORING TGT GOLIVE PT-2 & PT-3.
 * Mengikuti filter Vendor global; filter Region tidak diterapkan agar empat
 * kartu regional tetap tampil dan AREA-1 selalu tie-out ke tiga regional.
 */
function getCombinedGoliveReport(filters) {
  return getCombinedGoliveReport_(filters);
}

/** Detail on-demand agar payload awal tab gabungan tetap ringan. */
function getCombinedGoliveDetails(filters) {
  filters = filters || {};
  const wantedRegion = String(filters.region || 'AREA-1').trim().toUpperCase();
  const wantedProgram = String(filters.program || '').trim().toUpperCase();
  const wantedMetric = String(filters.metric || '').trim().toLowerCase();
  const wantedDay = Number(filters.day) || 0;
  const wantedDateKind = String(filters.dateKind || '').trim().toLowerCase();
  const wantedVendor = normalize_(filters.vendor);
  const ctx = combinedContext_();
  const records = [];
  let targetPorts = 0;
  let actualPorts = 0;

  combinedWalkRows_(ctx, function (record) {
    if (wantedVendor && normalize_(record.vendor) !== wantedVendor) return;
    if (wantedRegion !== 'AREA-1' && record.region !== wantedRegion) return;
    if (wantedProgram && record.program !== wantedProgram) return;
    let metricMatch = record.targetInMonth || record.actualInMonth;
    if (wantedMetric === 'achievement') metricMatch = (record.targetInMonth && record.targetDay <= ctx.todayDay) || record.actualInMonth;
    if (wantedMetric === 'ontime') metricMatch = record.actualInMonth && record.targetValid;
    if (wantedMetric === 'overdue') metricMatch = record.overdueOpen;
    if (wantedMetric === 'momentum') metricMatch = record.momentumTarget || record.momentumActual;
    if (wantedMetric === 'data') metricMatch = record.qualityEligible;
    if (wantedMetric === 'gap') metricMatch = record.targetInMonth && record.targetDay <= ctx.todayDay && !record.actualCompletedByToday;
    if (wantedDateKind === 'target') metricMatch = record.targetInMonth && (!wantedDay || record.targetDay === wantedDay);
    if (wantedDateKind === 'actual') metricMatch = record.actualInMonth && (!wantedDay || record.actualDay === wantedDay);
    if (!metricMatch) return;
    records.push([
      record.program, record.region, record.id, record.lop, record.branch,
      record.sto, record.status, record.port, record.targetDate || record.rawTarget,
      record.actualDate || record.rawActual, record.targetInMonth ? 1 : 0, record.actualInMonth ? 1 : 0
    ]);
    if (record.targetInMonth) targetPorts += record.port;
    if (record.actualInMonth) actualPorts += record.port;
  });

  records.sort(function (a, b) {
    const ad = a[9] || a[8] || '9999-12-31';
    const bd = b[9] || b[8] || '9999-12-31';
    if (ad !== bd) return ad < bd ? -1 : 1;
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    return String(a[3]).localeCompare(String(b[3]));
  });

  return {
    ok:true,
    region:wantedRegion,
    program:wantedProgram,
    metric:wantedMetric,
    day:wantedDay,
    dateKind:wantedDateKind,
    selectedVendor:String(filters.vendor || '').trim(),
    records:records,
    matchedRows:records.length,
    targetPorts:targetPorts,
    actualPorts:actualPorts
  };
}

function getCombinedGoliveReport_(filters) {
  filters = filters || {};
  const selectedVendor = String(filters.vendor || '').trim();
  const wantedVendor = normalize_(selectedVendor);
  const forceRefresh = filters.forceRefresh === true;
  const cacheDate = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const cacheVendor = (wantedVendor || 'ALL').replace(/[^A-Z0-9]/g, '_').slice(0, 60);
  const cacheKey = 'pt23-combined-v7-executive-toggle-' + cacheDate + '-' + cacheVendor;
  const cacheGate = dashboardCacheGate_(cacheKey, forceRefresh);
  if (cacheGate.value) return cacheGate.value;
  const ctx = combinedContext_();

  const regionNames = ['SUMBAGUT', 'SUMBAGTENG', 'SUMBAGSEL'];
  const regions = {};
  const daily = [];
  const quality = {
    pt2InvalidTargetDate:0, pt2BlankActualDate:0, pt2FutureActualDate:0,
    pt3InvalidTargetDate:0, pt3BlankActualDate:0, pt3FutureActualDate:0,
    ignoredRegionRows:0
  };

  function blankProgram_() {
    return {
      targetMTD:0,targetFM:0,actualMTD:0,
      actualEligiblePorts:0,onTimePorts:0,
      overdueDuePorts:0,overdueOpenPorts:0,
      momentumTarget:0,momentumActual:0,
      qualityEarned:0,qualityPossible:0
    };
  }
  function blankRegion_() { return {pt2:blankProgram_(), pt3:blankProgram_()}; }
  regionNames.forEach(function (name) { regions[name] = blankRegion_(); });
  function blankRegionPorts_() { return {SUMBAGUT:0,SUMBAGTENG:0,SUMBAGSEL:0}; }
  for (let day = 1; day <= ctx.totalDays; day++) {
    daily.push({
      day:day,
      target:0,actual:0,
      pt2Target:0,pt3Target:0,
      pt2Actual:0,pt3Actual:0,
      targetRegions:blankRegionPorts_(),
      regions:blankRegionPorts_(),
      pt2TargetRegions:blankRegionPorts_(),
      pt3TargetRegions:blankRegionPorts_(),
      pt2Regions:blankRegionPorts_(),
      pt3Regions:blankRegionPorts_()
    });
  }

  combinedWalkRows_(ctx, function (record) {
    if (wantedVendor && normalize_(record.vendor) !== wantedVendor) return;
    if (!regions[record.region]) { quality.ignoredRegionRows++; return; }
    const key = record.program === 'PT2' ? 'pt2' : 'pt3';
    const bucket = regions[record.region][key];
    const prefix = record.program === 'PT2' ? 'pt2' : 'pt3';

    if (record.targetInMonth) {
      bucket.targetFM += record.port;
      const targetDay = daily[record.targetDay - 1];
      targetDay.target += record.port;
      targetDay[record.program === 'PT2' ? 'pt2Target' : 'pt3Target'] += record.port;
      targetDay.targetRegions[record.region] += record.port;
      targetDay[record.program === 'PT2' ? 'pt2TargetRegions' : 'pt3TargetRegions'][record.region] += record.port;
      if (record.targetDay <= ctx.todayDay) bucket.targetMTD += record.port;
    } else if (record.rawTarget) {
      quality[prefix + 'InvalidTargetDate']++;
    }

    if (record.actualInMonth) {
      bucket.actualMTD += record.port;
      const d = daily[record.actualDay - 1];
      d.actual += record.port;
      d[record.program === 'PT2' ? 'pt2Actual' : 'pt3Actual'] += record.port;
      d.regions[record.region] += record.port;
      d[record.program === 'PT2' ? 'pt2Regions' : 'pt3Regions'][record.region] += record.port;
    } else if (!record.rawActual) {
      quality[prefix + 'BlankActualDate']++;
    } else if (record.actualIsFuture) {
      quality[prefix + 'FutureActualDate']++;
    }

    if (record.actualInMonth && record.targetValid) {
      bucket.actualEligiblePorts += record.port;
      if (record.onTime) bucket.onTimePorts += record.port;
    }
    if (record.overdueDue) {
      bucket.overdueDuePorts += record.port;
      if (record.overdueOpen) bucket.overdueOpenPorts += record.port;
    }
    if (record.momentumTarget) bucket.momentumTarget += record.port;
    if (record.momentumActual) bucket.momentumActual += record.port;
    if (record.qualityEligible) {
      bucket.qualityEarned += record.qualityEarned;
      bucket.qualityPossible += record.qualityPossible;
    }
  });

  function finishProgram_(p) {
    const achievement = p.targetMTD ? Math.min(100, p.actualMTD / p.targetMTD * 100) : 0;
    const onTime = p.actualEligiblePorts ? p.onTimePorts / p.actualEligiblePorts * 100 : 0;
    const overdue = p.overdueDuePorts ? Math.max(0, (1 - p.overdueOpenPorts / p.overdueDuePorts) * 100) : 100;
    const momentum = p.momentumTarget ? Math.min(100, p.momentumActual / p.momentumTarget * 100) : (p.momentumActual > 0 ? 100 : 0);
    const dataDiscipline = p.qualityPossible ? p.qualityEarned / p.qualityPossible * 100 : 0;
    const score = achievement * 0.35 + onTime * 0.25 + overdue * 0.20 + momentum * 0.10 + dataDiscipline * 0.10;
    return {
      targetMTD:p.targetMTD,
      targetFM:p.targetFM,
      actualMTD:p.actualMTD,
      actualEligiblePorts:p.actualEligiblePorts,
      onTimePorts:p.onTimePorts,
      fullPct:p.targetFM ? Number((p.actualMTD / p.targetFM * 100).toFixed(1)) : null,
      pacePct:p.targetMTD ? Number((p.actualMTD / p.targetMTD * 100).toFixed(1)) : null,
      remaining:p.targetFM - p.actualMTD,
      overduePorts:p.overdueOpenPorts,
      kpis:{
        achievement:Number(achievement.toFixed(1)),
        onTime:Number(onTime.toFixed(1)),
        overdueControl:Number(overdue.toFixed(1)),
        momentum7:Number(momentum.toFixed(1)),
        dataDiscipline:Number(dataDiscipline.toFixed(1)),
        score:Number(score.toFixed(1)),
        gapTarget:p.targetMTD-p.actualMTD
      }
    };
  }
  function addProgram_(a, b) {
    const out = blankProgram_();
    Object.keys(out).forEach(function(key){ out[key]=(Number(a[key])||0)+(Number(b[key])||0); });
    return out;
  }
  function finishRegion_(r) {
    const total = addProgram_(r.pt2, r.pt3);
    return {pt2:finishProgram_(r.pt2),pt3:finishProgram_(r.pt3),total:finishProgram_(total)};
  }

  const outRegions = {};
  let area = blankRegion_();
  regionNames.forEach(function (name) {
    outRegions[name] = finishRegion_(regions[name]);
    area.pt2 = addProgram_(area.pt2, regions[name].pt2);
    area.pt3 = addProgram_(area.pt3, regions[name].pt3);
  });
  const areaOut = finishRegion_(area);

  const ranking = regionNames.map(function(name){
    const total = outRegions[name].total;
    return {
      region:name,
      score:total.kpis.score,
      achievement:total.kpis.achievement,
      onTime:total.kpis.onTime,
      overdueControl:total.kpis.overdueControl,
      momentum7:total.kpis.momentum7,
      dataDiscipline:total.kpis.dataDiscipline,
      gapTarget:total.kpis.gapTarget,
      overduePorts:total.overduePorts
    };
  }).sort(function(a,b){
    return b.score-a.score || b.onTime-a.onTime || a.overduePorts-b.overduePorts || b.achievement-a.achievement || a.region.localeCompare(b.region);
  }).map(function(item,index){item.rank=index+1;return item;});

  const highestDays = daily.filter(function (d) { return d.actual > 0; }).map(function (d) {
    let topRegion = '';
    let topPorts = 0;
    regionNames.forEach(function (name) {
      if (d.regions[name] > topPorts) { topRegion = name; topPorts = d.regions[name]; }
    });
    return {
      date:ctx.year + '-' + String(ctx.month + 1).padStart(2,'0') + '-' + String(d.day).padStart(2,'0'),
      day:d.day, actual:d.actual, target:d.target, gap:d.actual-d.target,
      topRegion:topRegion, topPorts:topPorts
    };
  }).sort(function (a,b) { return b.actual-a.actual || a.day-b.day; }).slice(0,5);

  const result = {
    ok:true,
    schemaVersion:3,
    selectedVendor:selectedVendor,
    vendorFallbackApplied:true,
    reportDate:ctx.reportDate,
    todayDay:ctx.todayDay,
    totalDays:ctx.totalDays,
    monthLabel:Utilities.formatDate(ctx.today, CONFIG.timezone, 'MMMM yyyy'),
    activeDays:daily.filter(function(d){return d.actual>0;}).length,
    regions:outRegions,
    area1:areaOut,
    ranking:ranking,
    rankingBenchmark:{region:'AREA-1',score:areaOut.total.kpis.score,kpis:areaOut.total.kpis},
    rankingWeights:{achievement:35,onTime:25,overdueControl:20,momentum7:10,dataDiscipline:10},
    daily:daily,
    highestDays:highestDays,
    distribution:{pt2:areaOut.pt2.actualMTD,pt3:areaOut.pt3.actualMTD,total:areaOut.total.actualMTD},
    quality:quality,
    source:{pt2Rows:ctx.pt2.values.length-ctx.pt2.headerIndex-1,pt3Rows:ctx.pt3.values.length-ctx.pt3.headerIndex-1}
  };
  try { dashboardCacheWrite_(cacheKey, result); }
  catch (err) { console.warn('Cache gabungan dilewati: ' + err.message); }
  return result;
}

function combinedContext_() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, CONFIG.timezone, 'yyyy'));
  const month = Number(Utilities.formatDate(now, CONFIG.timezone, 'M')) - 1;
  const todayDay = Number(Utilities.formatDate(now, CONFIG.timezone, 'd'));
  const today = new Date(year, month, todayDay);
  const totalDays = new Date(year, month + 1, 0).getDate();
  let pt2Values = null;
  let pt3Values = null;
  try {
    const ss = SpreadsheetApp.openById(CONFIG.sourceSpreadsheetId);
    const pt2Sheet = ss.getSheetByName(CONFIG.pt2SheetName);
    const pt3Sheet = ss.getSheetByName(CONFIG.sheetName);
    if (!pt2Sheet) throw new Error('Sheet ' + CONFIG.pt2SheetName + ' tidak ditemukan.');
    if (!pt3Sheet) throw new Error('Sheet ' + CONFIG.sheetName + ' tidak ditemukan.');
    pt2Values = pt2Sheet.getDataRange().getValues();
    pt3Values = pt3Sheet.getDataRange().getValues();
  } catch (sheetErr) {
    console.warn('Sumber langsung PT2+PT3 gagal, memakai CSV fallback: ' + sheetErr.message);
    const responses = UrlFetchApp.fetchAll([
      {url:CONFIG.pt2CsvUrl,muteHttpExceptions:true},
      {url:CONFIG.csvUrl,muteHttpExceptions:true}
    ]);
    if (responses[0].getResponseCode() !== 200) throw new Error('Sumber LIST FI PT2 gagal dibaca (HTTP ' + responses[0].getResponseCode() + ').');
    if (responses[1].getResponseCode() !== 200) throw new Error('Sumber PT3 gagal dibaca (HTTP ' + responses[1].getResponseCode() + ').');
    pt2Values = Utilities.parseCsv(responses[0].getContentText());
    pt3Values = Utilities.parseCsv(responses[1].getContentText());
  }
  const pt2Header = combinedHeaderRow_(pt2Values, ['REGION','LOP','PORT ACTUAL','KOMITMEN GOLIVE']);
  const pt3Header = combinedHeaderRow_(pt3Values, ['REGION FMC','NAMA LOP','REAL PORT','KOMITMEN GOLIVE']);
  if (pt2Header < 0) throw new Error('Header LIST FI PT2 tidak lengkap.');
  if (pt3Header < 0) throw new Error('Header PT3 tidak lengkap.');
  return {
    now:now,year:year,month:month,todayDay:todayDay,today:today,totalDays:totalDays,
    reportDate:Utilities.formatDate(today,CONFIG.timezone,'yyyy-MM-dd'),
    pt2:{values:pt2Values,headerIndex:pt2Header},
    pt3:{values:pt3Values,headerIndex:pt3Header}
  };
}

function combinedHeaderRow_(values, required) {
  for (let i = 0; i < Math.min(values.length, 20); i++) {
    const headers = values[i].map(normalize_);
    if (required.every(function(name){return headers.indexOf(name)>=0;})) return i;
  }
  return -1;
}

function combinedWalkRows_(ctx, callback) {
  function walk_(program, source, names) {
    const headers = source.values[source.headerIndex].map(normalize_);
    const col = {};
    Object.keys(names).forEach(function(key){col[key]=headers.indexOf(names[key]);});
    for (let i = source.headerIndex + 1; i < source.values.length; i++) {
      const row = source.values[i];
      const region = String(row[col.region] || '').trim().toUpperCase();
      const port = toNumber_(row[col.port]);
      if (!region) continue;
      const rawTarget = String(row[col.target] || '').trim();
      const rawActual = String(row[col.actual] || '').trim();
      const rawVendor = col.vendor >= 0 ? String(row[col.vendor] || '').trim() : '';
      const vendor = program === 'PT2' && !rawVendor ? 'TELKOM AKSES' : rawVendor;
      const status = String(row[col.status] || '').trim();
      const target = parseDate_(row[col.target]);
      const actual = parseDate_(row[col.actual]);
      const targetInMonth = !!(target && target.getFullYear()===ctx.year && target.getMonth()===ctx.month);
      const actualCurrentMonth = !!(actual && actual.getFullYear()===ctx.year && actual.getMonth()===ctx.month);
      const actualInMonth = !!(actualCurrentMonth && actual.getDate()<=ctx.todayDay);
      const actualCompletedByToday = !!(actual && actual <= ctx.today);
      const momentumStart = Math.max(1,ctx.todayDay-6);
      const overdueDue = !!(targetInMonth && target.getDate()<ctx.todayDay);
      const overdueOpen = !!(overdueDue && !actualCompletedByToday);
      const monthTokens = [
        ['JAN','JANUARI'],['FEB','FEBRUARI'],['MAR','MARET'],['APR','APRIL'],['MEI','MAY'],['JUN','JUNI'],
        ['JUL','JULI'],['AGU','AGUSTUS','AUG'],['SEP','SEPTEMBER'],['OKT','OKTOBER','OCT'],['NOV','NOVEMBER'],['DES','DESEMBER','DEC']
      ][ctx.month];
      const rawTargetUpper = rawTarget.toUpperCase();
      const statusUpper = status.toUpperCase();
      const targetTextCurrentMonth = !target && monthTokens.some(function(token){return rawTargetUpper.indexOf(token)>=0;});
      const dropped = /DROP|CANCEL/.test(statusUpper+' '+rawTargetUpper);
      const qualityEligible = !dropped && (targetInMonth || actualCurrentMonth || targetTextCurrentMonth);
      const statusNeedsActual = statusUpper.indexOf('GOLIVE')>=0;
      const qualityEarned = qualityEligible ? ((region?1:0)+(port>0?1:0)+(targetInMonth?1:0)+((!statusNeedsActual||actual)?1:0)) : 0;
      callback({
        program:program,vendor:vendor,region:region,id:String(row[col.id]||'').trim(),lop:String(row[col.lop]||'').trim(),
        branch:String(row[col.branch]||'').trim(),sto:String(row[col.sto]||'').trim(),status:status,port:port,
        rawTarget:rawTarget,rawActual:rawActual,
        targetDate:target?Utilities.formatDate(target,CONFIG.timezone,'yyyy-MM-dd'):'',
        actualDate:actual?Utilities.formatDate(actual,CONFIG.timezone,'yyyy-MM-dd'):'',
        targetInMonth:targetInMonth,actualInMonth:actualInMonth,
        targetDay:targetInMonth?target.getDate():0,actualDay:actualInMonth?actual.getDate():0,
        targetValid:!!target,actualCompletedByToday:actualCompletedByToday,
        onTime:!!(actual && target && actual<=target),
        overdueDue:overdueDue,overdueOpen:overdueOpen,
        momentumTarget:!!(targetInMonth && target.getDate()>=momentumStart && target.getDate()<=ctx.todayDay),
        momentumActual:!!(actualInMonth && actual.getDate()>=momentumStart),
        qualityEligible:qualityEligible,qualityEarned:qualityEarned,qualityPossible:qualityEligible?4:0,
        actualIsFuture:!!(actualCurrentMonth && actual.getDate()>ctx.todayDay)
      });
    }
  }
  walk_('PT2',ctx.pt2,{vendor:'VENDOR',region:'REGION',id:'IHLD',lop:'LOP',branch:'BRANCH',sto:'STO',status:'STATUS',port:'PORT ACTUAL',target:'KOMITMEN GOLIVE',actual:'TANGGAL GOLIVE'});
  walk_('PT3',ctx.pt3,{vendor:'VENDOR',region:'REGION FMC',id:'ID-IHLD',lop:'NAMA LOP',branch:'BRANCH FMC',sto:'STO',status:'STATUS',port:'REAL PORT',target:'KOMITMEN GOLIVE',actual:'TANGGAL GOLIVE'});
}

/**
 * Detail LoP untuk sel Monitoring Progress & Prioritas.
 * Dipanggil saat pengguna mengklik angka agar payload awal tetap ringan.
 */
function getMonitoringDetails(filters) {
  filters = filters || {};
  let values = null;
  let headerIndex = -1;

  try {
    const ss = pt3Spreadsheet_();
    if (ss) {
      const sheet = getActiveOrTargetSheet_(ss, CONFIG.sheetName);
      if (sheet) {
        const raw = sheet.getDataRange().getValues();
        const idx = findHeaderRow_(raw);
        if (idx >= 0) { values = raw; headerIndex = idx; }
      }
    }
  } catch (err) {
    console.warn('Detail monitoring gagal membaca spreadsheet aktif: ' + err.message);
  }

  if ((!values || headerIndex < 0) && CONFIG.csvUrl) {
    const res = UrlFetchApp.fetch(CONFIG.csvUrl, {muteHttpExceptions:true});
    if (res.getResponseCode() === 200) {
      const parsed = Utilities.parseCsv(res.getContentText());
      const idx = findHeaderRow_(parsed);
      if (idx >= 0) { values = parsed; headerIndex = idx; }
    }
  }
  if (!values || headerIndex < 0) throw new Error('Sumber detail monitoring tidak ditemukan.');

  const headers = values[headerIndex].map(normalize_);
  function pos_(name, fallback) { const p = headers.indexOf(name); return p >= 0 ? p : fallback; }
  const col = {
    id:pos_('ID-IHLD',0), lop:pos_('NAMA LOP',1), regional:pos_('REGIONAL',2),
    area:pos_('AREA',3), sto:pos_('STO',4), region:pos_('REGION FMC',5),
    branch:pos_('BRANCH FMC',6), mitra:pos_('MITRA',7), status:pos_('STATUS',8),
    subStatus:pos_('SUB STATUS KONS',9), port:pos_('REAL PORT',10), boq:pos_('BOQ',11),
    goliveDate:pos_('TANGGAL GOLIVE',22), komitmenDate:pos_('KOMITMEN GOLIVE',23),
    flagging:headers.indexOf('FLAGGING PRIO')
  };
  const vendorPos = headers.indexOf('VENDOR');
  const wantedVendor = String(filters.vendor || '').trim();
  if (wantedVendor && vendorPos < 0) throw new Error('Kolom VENDOR tidak ditemukan.');
  const globalRegion = String(filters.globalRegion || '').trim().toUpperCase();
  const wantedRegion = String(filters.region || '').trim().toUpperCase();
  const wantedStatus = String(filters.status || '').trim().toUpperCase();
  const dimensionType = String(filters.dimensionType || '').trim().toLowerCase();
  const wantedDimension = String(filters.dimensionValue || '').trim().toUpperCase();
  const priorityOnly = filters.priorityOnly === true || String(filters.priorityOnly || '') === '1';
  const wantedPriorityGroup = dimensionType === 'prioritygroup' ? wantedDimension : '';
  const wantedSubStatus = dimensionType === 'substatus' ? wantedDimension : '';
  function priorityGroup_(value) {
    const normalized = String(value || '').trim().toUpperCase().replace(/\s+/g, ' ')
      .replace(/\(\s*/g, '(').replace(/\s*\)/g, ')');
    return normalized === 'TSEL PRIORITY AGS (CARRY OVER)' || normalized === 'TSEL PRIORITY SEP'
      ? 'PRIORITAS' : 'NON PRIORITAS';
  }
  function statusStage_(value) {
    return String(value || '').replace(/^\s*\d+\s*[.)\-:]?\s*/, '').trim().toUpperCase();
  }
  const records = [];

  for (let i = headerIndex + 1; i < values.length; i++) {
    const row = values[i];
    if (wantedVendor && String(row[vendorPos] || '').trim() !== wantedVendor) continue;
    const region = String(row[col.region] || '').trim().toUpperCase() || 'REGION TIDAK TERISI';
    const branch = String(row[col.branch] || '').trim().toUpperCase() || 'BRANCH TIDAK TERISI';
    const status = String(row[col.status] || '').trim().toUpperCase() || 'STATUS KOSONG';
    const flag = col.flagging >= 0 ? (String(row[col.flagging] || '').trim().toUpperCase() || 'NON PRIORITY') : 'NON PRIORITY';
    if (priorityOnly && flag === 'NON PRIORITY') continue;
    if (wantedPriorityGroup && priorityGroup_(flag) !== wantedPriorityGroup) continue;
    if (globalRegion && region !== globalRegion) continue;
    if (wantedRegion && region !== wantedRegion) continue;
    if (wantedStatus && status !== wantedStatus && statusStage_(status) !== statusStage_(wantedStatus)) continue;
    if (dimensionType === 'branch' && wantedDimension && branch !== wantedDimension) continue;
    if (dimensionType === 'flag' && wantedDimension && flag !== wantedDimension) continue;
    if (wantedSubStatus && (String(row[col.subStatus] || '').trim().toUpperCase() || 'SUB STATUS KOSONG') !== wantedSubStatus) continue;
    records.push([
      String(row[col.id] || '').trim(), String(row[col.lop] || '').trim(),
      String(row[col.regional] || '').trim(), String(row[col.area] || '').trim(),
      String(row[col.sto] || '').trim(), region, String(row[col.branch] || '').trim(),
      String(row[col.mitra] || '').trim(), String(row[col.status] || '').trim(),
      String(row[col.subStatus] || '').trim(), toNumber_(row[col.port]),
      String(row[col.boq] || '').trim(), 0, flag, pt3Date_(row[col.goliveDate]), pt3Date_(row[col.komitmenDate])
    ]);
  }
  return {ok:true, records:records, matchedRows:records.length};
}

/**
 * Logika utama pembacaan dan kalkulasi data Golive PT3
 */
function getReport_(filters) {
  filters = filters || {};
  const vendorFilter = String(filters.vendor || "").trim();
  const regionFilter = String(filters.region || "").trim().toUpperCase();
  const forceRefresh = filters.forceRefresh === true;
  const now = new Date();
  const year = Number(Utilities.formatDate(now, CONFIG.timezone, 'yyyy'));
  const month = Number(Utilities.formatDate(now, CONFIG.timezone, 'M')) - 1; // 0-indexed
  const days = Number(Utilities.formatDate(now, CONFIG.timezone, 'd'));
  const today = new Date(year, month, days);
  const totalDaysInMonth = new Date(year, month + 1, 0).getDate();

  const cacheKey = 'pt3-v57-stage-overdue-' + Utilities.formatDate(now, CONFIG.timezone, 'yyyyMMdd') + '-' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify([vendorFilter,regionFilter]))) + '-' + (PropertiesService.getScriptProperties().getProperty('PT3_HISTORY_REV') || '0');
  const cacheGate = dashboardCacheGate_(cacheKey, forceRefresh);
  if (cacheGate.value) return cacheGate.value;

  let values = null;
  let sourceName = 'Active Spreadsheet';
  let headerIndex = -1;

  // 1. Coba baca dari spreadsheet yang terhubung
  try {
    const ss = pt3Spreadsheet_();
    if (ss) {
      const sheet = getActiveOrTargetSheet_(ss, CONFIG.sheetName);
      if (sheet) {
        const rawValues = sheet.getDataRange().getValues();
        if (rawValues && rawValues.length >= 2) {
          const idx = findHeaderRow_(rawValues);
          if (idx >= 0) {
            values = rawValues;
            headerIndex = idx;
            sourceName = sheet.getName();
          }
        }
      }
    }
  } catch (sheetErr) {
    console.warn('Gagal membaca Active Spreadsheet: ' + sheetErr.message);
  }

  // 2. Fallback cerdas: jika spreadsheet lokal kosong / tidak ada header, baca dari CSV publik
  if ((!values || headerIndex < 0) && CONFIG.csvUrl) {
    try {
      const res = UrlFetchApp.fetch(CONFIG.csvUrl, {muteHttpExceptions: true});
      if (res.getResponseCode() === 200) {
        const parsed = Utilities.parseCsv(res.getContentText());
        const idx = findHeaderRow_(parsed);
        if (idx >= 0) {
          values = parsed;
          headerIndex = idx;
          sourceName = 'Google Sheets (Published CSV)';
        }
      }
    } catch (csvErr) {
      console.warn('Gagal membaca fallback CSV di server: ' + csvErr.message);
    }
  }

  if (!values || headerIndex < 0) {
    throw new Error('Data spreadsheet tidak ditemukan. Pastikan sheet memuat kolom STATUS dan REAL PORT.');
  }

  const headers = values[headerIndex].map(normalize_);

  // Deteksi kolom atribut
  const col = {
    id: headers.indexOf('ID-IHLD') >= 0 ? headers.indexOf('ID-IHLD') : 0,
    lop: headers.indexOf('NAMA LOP') >= 0 ? headers.indexOf('NAMA LOP') : 1,
    regional: headers.indexOf('REGIONAL') >= 0 ? headers.indexOf('REGIONAL') : 2,
    area: headers.indexOf('AREA') >= 0 ? headers.indexOf('AREA') : 3,
    sto: headers.indexOf('STO') >= 0 ? headers.indexOf('STO') : 4,
    region: headers.indexOf('REGION FMC') >= 0 ? headers.indexOf('REGION FMC') : 5,
    branch: headers.indexOf('BRANCH FMC') >= 0 ? headers.indexOf('BRANCH FMC') : 6,
    mitra: headers.indexOf('MITRA') >= 0 ? headers.indexOf('MITRA') : 7,
    status: headers.indexOf('STATUS') >= 0 ? headers.indexOf('STATUS') : 8,
    subStatus: headers.indexOf('SUB STATUS KONS') >= 0 ? headers.indexOf('SUB STATUS KONS') : 9,
    port: headers.indexOf('REAL PORT') >= 0 ? headers.indexOf('REAL PORT') : 10,
    boq: headers.indexOf('BOQ') >= 0 ? headers.indexOf('BOQ') : 11,
    vendor: headers.indexOf('VENDOR'),
    flagging: headers.indexOf('FLAGGING PRIO'),
    finishInstallCommit: 24,   // Kolom Y
    installCommit: 25,         // Kolom Z
    aanwijzingCommit: 28,      // Kolom AC
    goliveDate: headers.indexOf('TANGGAL GOLIVE') >= 0 ? headers.indexOf('TANGGAL GOLIVE') : 22,
    komitmenDate: headers.indexOf('KOMITMEN GOLIVE') >= 0 ? headers.indexOf('KOMITMEN GOLIVE') : 23
  };

  // Inisialisasi region default PT3
  const defaultRegions = regionFilter ? [regionFilter] : ['SUMBAGUT', 'SUMBAGTENG', 'SUMBAGSEL'];
  const byRegion = {};
  defaultRegions.forEach(reg => {
    byRegion[reg] = Array(days).fill(0);
  });

  const dailyTarget = Array(totalDaysInMonth).fill(0);
  const dailyReal = Array(totalDaysInMonth).fill(0);

  let sourceRows = 0;
  let matchedRows = 0;
  const recordsList = [];
  const targetRecordsList = [];

  const allRows = values.slice(headerIndex + 1);
  const vendors = [...new Set(allRows.map(row => col.vendor >= 0 ? String(row[col.vendor] || '').trim() : '').filter(Boolean))].sort();
  const regionOptions = [...new Set(allRows.map(row => String(row[col.region] || '').trim().toUpperCase()).filter(Boolean))].sort();
  if (regionFilter && headers.indexOf('REGION FMC') < 0) throw new Error('Kolom REGION FMC tidak ditemukan.');
  if (vendorFilter && col.vendor < 0) throw new Error('Kolom VENDOR tidak ditemukan.');
  let historyState = {meta:{}, available:false, message:'Jalankan setupLopHistory sekali untuk mengaktifkan history Plan B.'};
  try { historyState = pt3SyncHistory_(values, headerIndex); } catch (err) { historyState.message = err.message; }
  const dataRows = allRows.filter(row => (!vendorFilter || String(row[col.vendor] || '').trim() === vendorFilter) && (!regionFilter || String(row[col.region] || '').trim().toUpperCase() === regionFilter));
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    sourceRows++;
    const status = String(row[col.status] || '').trim();
    const port = toNumber_(row[col.port]);
    const rawRegion = String(row[col.region] || '').trim().toUpperCase() || 'LAINNYA';

    // 1. Hitung Target Komitmen Golive berdasarkan header KOMITMEN GOLIVE
    const komitmenDate = parseDate_(row[col.komitmenDate]);
    if (komitmenDate && komitmenDate.getFullYear() === year && komitmenDate.getMonth() === month) {
      const kDay = komitmenDate.getDate();
      if (kDay >= 1 && kDay <= totalDaysInMonth) {
        dailyTarget[kDay - 1] += port;

        // Simpan dalam format array ringkas agar payload ringan (< 40 KB)
        targetRecordsList.push([
          String(row[col.id] || '').trim(),
          String(row[col.lop] || '').trim(),
          String(row[col.regional] || '').trim(),
          String(row[col.area] || '').trim(),
          String(row[col.sto] || '').trim(),
          rawRegion,
          String(row[col.branch] || '').trim(),
          String(row[col.mitra] || '').trim(),
          status,
          String(row[col.subStatus] || '').trim(),
          port,
          String(row[col.boq] || '').trim(),
          kDay, '', pt3Date_(row[col.goliveDate]), pt3Date_(row[col.komitmenDate])
        ]);
      }
    }

    // 2. Hitung Realisasi Aktual Golive berdasarkan TANGGAL GOLIVE.
    // Sesuai definisi bisnis PT-3, tanggal valid adalah bukti realisasi;
    // teks STATUS tidak dijadikan syarat agar data tidak terlewat.
    const goliveDate = parseDate_(row[col.goliveDate]);
    if (!goliveDate) continue;
    if (goliveDate.getFullYear() !== year || goliveDate.getMonth() !== month) continue;

    const gDay = goliveDate.getDate();
    if (gDay >= 1 && gDay <= totalDaysInMonth) {
      dailyReal[gDay - 1] += port;
    }

    // Data s.d. hari ini untuk tabel per region dan detail modal
    if (goliveDate <= today) {
      if (!byRegion[rawRegion]) {
        byRegion[rawRegion] = Array(days).fill(0);
      }

      byRegion[rawRegion][goliveDate.getDate() - 1] += port;
      matchedRows++;

      recordsList.push([
        String(row[col.id] || '').trim(),
        String(row[col.lop] || '').trim(),
        String(row[col.regional] || '').trim(),
        String(row[col.area] || '').trim(),
        String(row[col.sto] || '').trim(),
        rawRegion,
        String(row[col.branch] || '').trim(),
        String(row[col.mitra] || '').trim(),
        status,
        String(row[col.subStatus] || '').trim(),
        port,
        String(row[col.boq] || '').trim(),
        goliveDate.getDate(), '', pt3Date_(row[col.goliveDate]), pt3Date_(row[col.komitmenDate])
      ]);
    }
  }

  const regions = Object.keys(byRegion).sort();
  const daily = Array.from({length: days}, (_, i) =>
    regions.reduce((total, reg) => total + (byRegion[reg][i] || 0), 0)
  );
  const cumulative = daily.reduce((total, val) => total + val, 0);
  const reportDate = Utilities.formatDate(today, CONFIG.timezone, 'yyyy-MM-dd');

  // Akumulasi Target vs Realisasi
  const targetCum = [];
  const realCum = [];
  const pctCum = [];
  let tSum = 0, rSum = 0;

  for (let i = 0; i < totalDaysInMonth; i++) {
    tSum += dailyTarget[i];
    targetCum.push(tSum);

    if (i < days) {
      rSum += dailyReal[i];
      realCum.push(rSum);
      const pct = tSum > 0 ? Number((rSum / tSum * 100).toFixed(1)) : 0;
      pctCum.push(pct);
    } else {
      realCum.push(null);
      pctCum.push(null);
    }
  }

  const totalTarget = targetCum[totalDaysInMonth - 1] || 0;
  const totalReal = cumulative;
  const totalPct = totalTarget > 0 ? Number((totalReal / totalTarget * 100).toFixed(1)) : null;

  const targetVsReal = {
    totalDays: totalDaysInMonth,
    dailyTarget,
    dailyReal,
    targetCum,
    realCum,
    pctCum,
    totalTarget,
    totalReal,
    totalPct
  };

  const insight = buildCalloutInsights_(dataRows, col, today, year, month);
  const monitoring = buildProgressMonitoring_(dataRows, col);
  const h1SnapshotDate = Utilities.formatDate(new Date(year, month, days - 1), CONFIG.timezone, 'yyyy-MM-dd');
  const previousMonitoring = (vendorFilter || regionFilter) ? null : loadMonitoringSnapshot_(h1SnapshotDate);
  monitoring.movement = previousMonitoring ? buildMonitoringMovement_(monitoring, previousMonitoring, h1SnapshotDate) : null;
  if (!vendorFilter && !regionFilter) saveMonitoringSnapshot_(reportDate, monitoring);

  const idCounts = {};
  allRows.forEach(row => { const key = 'ID:' + String(row[col.id] || '').trim(); idCounts[key] = (idCounts[key] || 0) + 1; });
  allRows.forEach(row => {
    const id = String(row[col.id] || '').trim(), key = 'ID:' + id;
    if (!id || idCounts[key] !== 1) { delete historyState.meta[key]; return; }
    historyState.meta[key] = Object.assign({}, historyState.meta[key] || {}, {
      vendor:col.vendor >= 0 ? String(row[col.vendor] || '').trim() : '', currentCommitment:pt3Date_(row[col.komitmenDate])
    });
  });
  const payload = {
    ok: true,
    schemaVersion: 13,
    regionOptions, selectedRegion:regionFilter, regionColumnFound:headers.indexOf('REGION FMC') >= 0,
    vendors, selectedVendor:vendorFilter, vendorColumnFound:col.vendor >= 0,
    lopMeta:historyState.meta, historyAvailable:historyState.available, historyMessage:historyState.message || "",
    reportDate,
    timezone: CONFIG.timezone,
    updatedAt: new Date().toISOString(),
    source: {
      sheet: sourceName,
      headerRow: headerIndex + 1,
      sourceRows,
      matchedRows
    },
    regions: byRegion,
    daily,
    cumulative,
    cols: ['id','lop','regional','area','sto','region','branch','mitra','status','subStatus','port','boq','day','flag','actualDate','commitmentDate'],
    records: recordsList,
    targetRecords: targetRecordsList,
    targetVsReal,
    callouts: insight.callouts,
    calloutRecords: insight.records,
    dataQuality: insight.dataQuality,
    monitoring
  };

  try {
    dashboardCacheWrite_(cacheKey, payload);
  } catch (err) {
    console.warn('Gagal menyimpan cache (diabaikan): ' + err.message);
  }

  return payload;
}

/**
 * Matriks snapshot untuk tab Monitoring Progress & Prioritas.
 * Agregasi dilakukan di server agar browser hanya menerima subtotal ringkas,
 * bukan seluruh baris mentah spreadsheet.
 */
function buildProgressMonitoring_(dataRows, col) {
  const PROCESS_STATUS_ORDER = [
    'DROP', 'AANWIJZING', 'DONE AANWIJZING', 'PERIZINAN', 'MATDEL',
    'INSTALASI', 'FINISH INSTALASI', 'GOLIVE', 'UJI TERIMA'
  ];
  const EMPTY_STATUS = 'STATUS KOSONG';
  const EMPTY_REGION = 'REGION TIDAK TERISI';
  const EMPTY_BRANCH = 'BRANCH TIDAK TERISI';
  const NON_PRIORITY = 'NON PRIORITY';
  const PRIORITY_GROUP = 'PRIORITAS';
  const NON_PRIORITY_GROUP = 'NON PRIORITAS';
  const statusesSeen = {};
  const statusOrder = [];
  const progressRegions = {};
  const priorityRegions = {};
  const progressGrand = {lop:0, port:0, boq:0};
  const priorityGrand = {lop:0, port:0, boq:0};
  const nonPriorityGrand = {lop:0, port:0, boq:0};
  const finishIssuesRegions = {};
  const finishIssuesGrand = {lop:0, port:0, boq:0};
  let includedRows = 0;
  let blankStatusRows = 0;
  let blankRegionRows = 0;
  let blankBranchRows = 0;
  let nonPriorityRows = 0;
  let invalidPortRows = 0;
  let invalidBoqRows = 0;

  function text_(value, fallback) {
    const s = String(value === null || value === undefined ? '' : value).trim().replace(/\s+/g, ' ');
    return s ? s.toUpperCase() : fallback;
  }
  function priorityGroup_(value) {
    const normalized = text_(value, NON_PRIORITY).replace(/\(\s*/g, '(').replace(/\s*\)/g, ')');
    return normalized === 'TSEL PRIORITY AGS (CARRY OVER)' || normalized === 'TSEL PRIORITY SEP'
      ? PRIORITY_GROUP : NON_PRIORITY_GROUP;
  }
  function addStatus_(status) {
    if (!statusesSeen[status]) {
      statusesSeen[status] = true;
      if (status !== EMPTY_STATUS) statusOrder.push(status);
    }
  }
  function statusStage_(status) {
    return String(status || '').replace(/^\s*\d+\s*[.)\-:]?\s*/, '').trim();
  }
  function statusRank_(status) {
    const rank = PROCESS_STATUS_ORDER.indexOf(statusStage_(status));
    return rank === -1 ? 999 : rank;
  }
  function addPair_(target, port, boq) {
    target.lop = (target.lop || 0) + 1;
    target.port = (target.port || 0) + port;
    target.boq = (target.boq || 0) + boq;
  }
  function ensureRegion_(map, region, childName) {
    if (!map[region]) map[region] = {total:{lop:0,port:0,boq:0}, cells:{}};
    if (!map[region][childName]) map[region][childName] = {};
    return map[region];
  }
  function addDimension_(map, region, childName, childValue, status, port, boq) {
    const regionNode = ensureRegion_(map, region, childName);
    if (!regionNode[childName][childValue]) regionNode[childName][childValue] = {total:{lop:0,port:0,boq:0}, cells:{}};
    const child = regionNode[childName][childValue];
    if (!child.cells[status]) child.cells[status] = {lop:0,port:0,boq:0};
    if (!regionNode.cells[status]) regionNode.cells[status] = {lop:0,port:0,boq:0};
    addPair_(child.cells[status], port, boq);
    addPair_(regionNode.cells[status], port, boq);
    addPair_(child.total, port, boq);
    addPair_(regionNode.total, port, boq);
  }
  function addFinishIssue_(region, subStatus, port, boq) {
    if (!finishIssuesRegions[region]) finishIssuesRegions[region] = {total:{lop:0,port:0,boq:0}, subStatuses:{}};
    const node = finishIssuesRegions[region];
    if (!node.subStatuses[subStatus]) node.subStatuses[subStatus] = {lop:0,port:0,boq:0};
    addPair_(node.subStatuses[subStatus], port, boq);
    addPair_(node.total, port, boq);
    addPair_(finishIssuesGrand, port, boq);
  }

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const rawRegion = String(row[col.region] || '').trim();
    const rawBranch = String(row[col.branch] || '').trim();
    const rawStatus = String(row[col.status] || '').trim();
    const rawFlag = col.flagging >= 0 ? String(row[col.flagging] || '').trim() : '';
    const rawPort = row[col.port];
    const rawBoq = row[col.boq];

    // Abaikan baris yang benar-benar kosong, tetapi pertahankan baris operasional bernilai nol.
    if (!rawRegion && !rawBranch && !rawStatus && !rawFlag && String(rawPort || '').trim() === '' && String(rawBoq || '').trim() === '') continue;

    const region = text_(rawRegion, EMPTY_REGION);
    const branch = text_(rawBranch, EMPTY_BRANCH);
    const status = text_(rawStatus, EMPTY_STATUS);
    const flag = text_(rawFlag, NON_PRIORITY);
    const priorityGroup = priorityGroup_(flag);
    const port = toNumber_(rawPort);
    const boq = toNumber_(rawBoq);

    includedRows++;
    if (!rawStatus) blankStatusRows++;
    if (!rawRegion) blankRegionRows++;
    if (!rawBranch) blankBranchRows++;
    if (priorityGroup === NON_PRIORITY_GROUP) nonPriorityRows++;
    if (String(rawPort === null || rawPort === undefined ? '' : rawPort).trim() && port === 0 && !/^[-+]?0*(?:[.,]0+)?$/.test(String(rawPort).trim())) invalidPortRows++;
    if (String(rawBoq === null || rawBoq === undefined ? '' : rawBoq).trim() && boq === 0 && !/^[-+]?0*(?:[.,]0+)?$/.test(String(rawBoq).trim())) invalidBoqRows++;
    addStatus_(status);
    addDimension_(progressRegions, region, 'branches', branch, status, port, boq);
    addPair_(progressGrand, port, boq);
    addDimension_(priorityRegions, region, 'groups', priorityGroup, status, port, boq);
    if (statusStage_(status) === 'FINISH INSTALASI') {
      addFinishIssue_(region, text_(row[col.subStatus], 'SUB STATUS KOSONG'), port, boq);
    }
    if (priorityGroup === NON_PRIORITY_GROUP) {
      addPair_(nonPriorityGrand, port, boq);
    } else {
      addPair_(priorityGrand, port, boq);
    }
  }

  // Gunakan hanya nilai aktual dari kolom STATUS. Nomor pada teks dipertahankan,
  // tetapi pengurutan mengikuti alur proses yang disepakati.
  const orderedStatuses = statusOrder.slice().sort((a, b) => {
    const diff = statusRank_(a) - statusRank_(b);
    return diff || a.localeCompare(b);
  });
  if (statusesSeen[EMPTY_STATUS]) orderedStatuses.push(EMPTY_STATUS);

  return {
    statuses: orderedStatuses,
    progress: {regions:progressRegions, grand:progressGrand},
    priority: {regions:priorityRegions, grand:progressGrand},
    finishIssues: {regions:finishIssuesRegions, grand:finishIssuesGrand},
    reconciliation: {
      includedRows,
      progressLop:progressGrand.lop,
      progressPort:progressGrand.port,
      progressBoq:progressGrand.boq,
      priorityLop:priorityGrand.lop,
      priorityPort:priorityGrand.port,
      priorityBoq:priorityGrand.boq,
      excludedNonPriorityLop:nonPriorityGrand.lop,
      excludedNonPriorityPort:nonPriorityGrand.port,
      excludedNonPriorityBoq:nonPriorityGrand.boq,
      lopDifference:progressGrand.lop - priorityGrand.lop - nonPriorityGrand.lop,
      portDifference:progressGrand.port - priorityGrand.port - nonPriorityGrand.port,
      boqDifference:progressGrand.boq - priorityGrand.boq - nonPriorityGrand.boq,
      blankStatusRows,
      blankRegionRows,
      blankBranchRows,
      nonPriorityRows,
      invalidPortRows,
      invalidBoqRows,
      missingFlaggingHeader:col.flagging < 0
    }
  };
}

/**
 * Delta H-1 untuk seluruh matriks monitoring. Nilai positif = naik,
 * negatif = turun, dan nol tidak ditampilkan di UI.
 */
function buildMonitoringMovement_(current, previous, baselineDate) {
  function pairDelta_(a, b) {
    a = a || {}; b = b || {};
    return {
      lop:(Number(a.lop) || 0) - (Number(b.lop) || 0),
      port:(Number(a.port) || 0) - (Number(b.port) || 0),
      boq:(Number(a.boq) || 0) - (Number(b.boq) || 0)
    };
  }
  function matrixDelta_(nowMatrix, oldMatrix, childKey) {
    nowMatrix = nowMatrix || {regions:{},grand:{}};
    oldMatrix = oldMatrix || {regions:{},grand:{}};
    const out = {grand:pairDelta_(nowMatrix.grand, oldMatrix.grand), regions:{}};
    Object.keys(nowMatrix.regions || {}).forEach(region => {
      const nowNode = nowMatrix.regions[region] || {};
      const oldNode = (oldMatrix.regions || {})[region] || {};
      const node = {total:pairDelta_(nowNode.total, oldNode.total), cells:{}};
      (current.statuses || Object.keys(nowNode.cells || {})).forEach(status => {
        node.cells[status] = pairDelta_(nowNode.cells[status], (oldNode.cells || {})[status]);
      });
      node[childKey] = {};
      Object.keys(nowNode[childKey] || {}).forEach(childName => {
        const nowChild = nowNode[childKey][childName] || {};
        const oldChild = (oldNode[childKey] || {})[childName] || {};
        const child = {total:pairDelta_(nowChild.total, oldChild.total), cells:{}};
        (current.statuses || Object.keys(nowChild.cells || {})).forEach(status => {
          child.cells[status] = pairDelta_(nowChild.cells[status], (oldChild.cells || {})[status]);
        });
        node[childKey][childName] = child;
      });
      out.regions[region] = node;
    });
    return out;
  }
  return {
    baselineDate,
    progress:matrixDelta_(current.progress, previous.progress, 'branches'),
    priority:matrixDelta_(current.priority, previous.priority, 'groups')
  };
}

function monitoringSnapshotPayload_(monitoring) {
  return {statuses:monitoring.statuses || [], progress:monitoring.progress || {}, priority:monitoring.priority || {}};
}

function saveMonitoringSnapshot_(dateKey, monitoring) {
  try {
    const props = PropertiesService.getScriptProperties();
    const prefix = 'PT3_MONITORING_' + String(dateKey).replace(/-/g, '') + '_';
    const packed = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(monitoringSnapshotPayload_(monitoring)))).getBytes());
    const chunkSize = 7000;
    const count = Math.ceil(packed.length / chunkSize);
    const oldCount = Number(props.getProperty(prefix + 'COUNT') || 0);
    const updates = {};
    updates[prefix + 'COUNT'] = String(count);
    for (let i = 0; i < count; i++) updates[prefix + i] = packed.slice(i * chunkSize, (i + 1) * chunkSize);
    props.setProperties(updates, false);
    for (let j = count; j < oldCount; j++) props.deleteProperty(prefix + j);

    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 8);
    const cutoffKey = Utilities.formatDate(cutoff, CONFIG.timezone, 'yyyyMMdd');
    Object.keys(props.getProperties()).forEach(key => {
      const match = key.match(/^PT3_MONITORING_(\d{8})_/);
      if (match && match[1] < cutoffKey) props.deleteProperty(key);
    });
  } catch (err) {
    console.warn('Snapshot monitoring gagal disimpan: ' + err.message);
  }
}

function loadMonitoringSnapshot_(dateKey) {
  try {
    const props = PropertiesService.getScriptProperties();
    const prefix = 'PT3_MONITORING_' + String(dateKey).replace(/-/g, '') + '_';
    const count = Number(props.getProperty(prefix + 'COUNT') || 0);
    if (!count) return null;
    let packed = '';
    for (let i = 0; i < count; i++) {
      const chunk = props.getProperty(prefix + i);
      if (chunk === null) return null;
      packed += chunk;
    }
    const json = Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(packed))).getDataAsString();
    return JSON.parse(json);
  } catch (err) {
    console.warn('Snapshot monitoring H-1 gagal dibaca: ' + err.message);
    return null;
  }
}

/**
 * Enam callout eksekutif PT-3.
 * Semua nilai utama berbobot REAL PORT; jumlah LoP memakai ID-IHLD unik
 * (baris tanpa ID diberi kunci baris agar tidak hilang dari kontrol).
 */
function buildCalloutInsights_(dataRows, col, today, year, month) {
  const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  const horizon = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7);
  const monthStart = new Date(year, month, 1);
  const monthEnd = new Date(year, month + 1, 0);
  const mondayOffset = (today.getDay() + 6) % 7;
  const calendarWeekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate() - mondayOffset);
  const calendarWeekEnd = new Date(calendarWeekStart.getFullYear(), calendarWeekStart.getMonth(), calendarWeekStart.getDate() + 6);
  const weekStart = calendarWeekStart < monthStart ? monthStart : calendarWeekStart;
  const weekEnd = calendarWeekEnd > monthEnd ? monthEnd : calendarWeekEnd;
  const firstOfMonth = new Date(year, month, 1);
  const firstMonthOffset = (firstOfMonth.getDay() + 6) % 7;
  const weekNo = Math.floor((today.getDate() + firstMonthOffset - 1) / 7) + 1;
  const regionsFound = {SUMBAGUT:true, SUMBAGTENG:true, SUMBAGSEL:true};

  const achievement = {target:0, actual:0, targetIds:{}, actualIds:{}, regions:{}};
  const h1 = {ports:0, ids:{}, regions:{}};
  const onTime = {ports:0, eligiblePorts:0, ids:{}, eligibleIds:{}, regions:{}};
  const overdue = {ports:0, ids:{}, regions:{}};
  const overdueAanwijzing = {ports:0, ids:{}, regions:{}};
  const overdueInstalasi = {ports:0, ids:{}, regions:{}};
  const overdueFinishInstalasi = {ports:0, ids:{}, regions:{}};
  const upcoming = {ports:0, ids:{}, regions:{}};
  const weekly = {target:0, actual:0, targetIds:{}, actualIds:{}, regions:{}};
  const records = {achievement:[], h1:[], onTime:[], overdue:[], overdueAanwijzing:[], overdueInstalasi:[], overdueFinishInstalasi:[], upcoming7:[], weekly:[]};
  const seenIdRows = {};
  let duplicateIdRows = 0;
  let missingRegionRows = 0;
  let invalidPortRows = 0;
  const stageCommitmentGaps = {aanwijzing:0, instalasi:0, finishInstalasi:0};
  const WORKFLOW = ['AANWIJZING','DONE AANWIJZING','PERIZINAN','MATDEL','INSTALASI','FINISH INSTALASI','GOLIVE','UJI TERIMA'];

  function sameDate_(a, b) {
    return a && b && a.getFullYear() === b.getFullYear() &&
      a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }
  function ensure_(map, region, template) {
    if (!map[region]) map[region] = Object.assign({}, template, {ids:{}, eligibleIds:{}, targetIds:{}, actualIds:{}});
    return map[region];
  }
  function mark_(bag, key) { bag['K:' + key] = true; }
  function count_(bag) { return Object.keys(bag || {}).length; }
  function workflowRank_(value) {
    const stage = String(value || '').replace(/^\s*\d+\s*[.)\-:]?\s*/, '').trim().toUpperCase();
    return WORKFLOW.indexOf(stage);
  }
  function pack_(row, dateValue, commitmentOverride) {
    return [
      String(row[col.id] || '').trim(), String(row[col.lop] || '').trim(),
      String(row[col.regional] || '').trim(), String(row[col.area] || '').trim(),
      String(row[col.sto] || '').trim(), String(row[col.region] || '').trim().toUpperCase() || 'LAINNYA',
      String(row[col.branch] || '').trim(), String(row[col.mitra] || '').trim(),
      String(row[col.status] || '').trim(), String(row[col.subStatus] || '').trim(),
      toNumber_(row[col.port]), String(row[col.boq] || '').trim(), dateValue ? dateValue.getDate() : 0, '', pt3Date_(row[col.goliveDate]), pt3Date_(commitmentOverride || row[col.komitmenDate])
    ];
  }

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const id = String(row[col.id] || '').trim();
    const lopKey = id || ('ROW-' + (i + 1));
    const rawRegion = String(row[col.region] || '').trim().toUpperCase();
    const region = rawRegion || 'LAINNYA';
    const port = toNumber_(row[col.port]);
    const actualDate = parseDate_(row[col.goliveDate]);
    const commitDate = parseDate_(row[col.komitmenDate]);
    const statusRank = workflowRank_(row[col.status]);
    regionsFound[region] = true;

    if (!rawRegion) missingRegionRows++;
    if (row[col.port] !== '' && row[col.port] !== null && row[col.port] !== undefined && port === 0 && !/^[-+]?0*(?:[.,]0+)?$/.test(String(row[col.port]).trim())) invalidPortRows++;
    if (id) {
      if (seenIdRows[id]) duplicateIdRows++;
      seenIdRows[id] = true;
    }

    const targetDue = commitDate && commitDate >= monthStart && commitDate <= today;
    const actualThisMonth = actualDate && actualDate >= monthStart && actualDate <= today;

    if (commitDate && commitDate >= weekStart && commitDate <= weekEnd) {
      weekly.target += port;
      mark_(weekly.targetIds, lopKey);
      const wReg = ensure_(weekly.regions, region, {target:0, actual:0});
      wReg.target += port;
      mark_(wReg.targetIds, lopKey);
    }
    if (actualDate && actualDate >= weekStart && actualDate <= today && actualDate <= weekEnd) {
      weekly.actual += port;
      mark_(weekly.actualIds, lopKey);
      const wReg = ensure_(weekly.regions, region, {target:0, actual:0});
      wReg.actual += port;
      mark_(wReg.actualIds, lopKey);
      records.weekly.push(pack_(row, actualDate));
    }

    if (targetDue) {
      achievement.target += port;
      mark_(achievement.targetIds, lopKey);
      const aReg = ensure_(achievement.regions, region, {target:0, actual:0});
      aReg.target += port;
      mark_(aReg.targetIds, lopKey);
    }
    if (actualThisMonth) {
      achievement.actual += port;
      mark_(achievement.actualIds, lopKey);
      const aReg = ensure_(achievement.regions, region, {target:0, actual:0});
      aReg.actual += port;
      mark_(aReg.actualIds, lopKey);
      records.achievement.push(pack_(row, actualDate));
    }

    if (sameDate_(actualDate, yesterday)) {
      h1.ports += port;
      mark_(h1.ids, lopKey);
      const r = ensure_(h1.regions, region, {ports:0});
      r.ports += port;
      mark_(r.ids, lopKey);
      records.h1.push(pack_(row, actualDate));
    }

    if (actualThisMonth && commitDate) {
      onTime.eligiblePorts += port;
      mark_(onTime.eligibleIds, lopKey);
      const r = ensure_(onTime.regions, region, {ports:0, eligiblePorts:0});
      r.eligiblePorts += port;
      mark_(r.eligibleIds, lopKey);
      if (actualDate <= commitDate) {
        onTime.ports += port;
        mark_(onTime.ids, lopKey);
        r.ports += port;
        mark_(r.ids, lopKey);
        records.onTime.push(pack_(row, actualDate));
      }
    }

    if (commitDate && commitDate < today && !actualDate) {
      overdue.ports += port;
      mark_(overdue.ids, lopKey);
      const r = ensure_(overdue.regions, region, {ports:0});
      r.ports += port;
      mark_(r.ids, lopKey);
      records.overdue.push(pack_(row, commitDate));
    }

    function stageOverdue_(bucket, recordKey, gapKey, columnIndex, completionRank) {
      // DROP, status kosong, dan status di luar delapan tahap tidak dinilai.
      if (statusRank < 0) return;
      const rawCommitment = row[columnIndex];
      const stageCommitment = parseDate_(rawCommitment);
      if (!stageCommitment) {
        stageCommitmentGaps[gapKey]++;
        return;
      }
      if (stageCommitment < today && statusRank < completionRank) {
        bucket.ports += port;
        mark_(bucket.ids, lopKey);
        const r = ensure_(bucket.regions, region, {ports:0});
        r.ports += port;
        mark_(r.ids, lopKey);
        records[recordKey].push(pack_(row, stageCommitment, stageCommitment));
      }
    }
    stageOverdue_(overdueAanwijzing, 'overdueAanwijzing', 'aanwijzing', col.aanwijzingCommit, 1);
    stageOverdue_(overdueInstalasi, 'overdueInstalasi', 'instalasi', col.installCommit, 4);
    stageOverdue_(overdueFinishInstalasi, 'overdueFinishInstalasi', 'finishInstalasi', col.finishInstallCommit, 5);

    if (commitDate && commitDate >= today && commitDate <= horizon && !actualDate) {
      upcoming.ports += port;
      mark_(upcoming.ids, lopKey);
      const r = ensure_(upcoming.regions, region, {ports:0});
      r.ports += port;
      mark_(r.ids, lopKey);
      records.upcoming7.push(pack_(row, commitDate));
    }
  }

  const regionNames = Object.keys(regionsFound).sort();
  const out = {achievement:{}, h1:{}, onTime:{}, overdue:{}, overdueAanwijzing:{}, overdueInstalasi:{}, overdueFinishInstalasi:{}, upcoming7:{}, weekly:{}};
  regionNames.forEach(region => {
    const a = achievement.regions[region] || {target:0, actual:0, targetIds:{}, actualIds:{}};
    const d = h1.regions[region] || {ports:0, ids:{}};
    const o = onTime.regions[region] || {ports:0, eligiblePorts:0, ids:{}, eligibleIds:{}};
    const l = overdue.regions[region] || {ports:0, ids:{}};
    const oa = overdueAanwijzing.regions[region] || {ports:0, ids:{}};
    const oi = overdueInstalasi.regions[region] || {ports:0, ids:{}};
    const ofi = overdueFinishInstalasi.regions[region] || {ports:0, ids:{}};
    const u = upcoming.regions[region] || {ports:0, ids:{}};
    const w = weekly.regions[region] || {target:0, actual:0, targetIds:{}, actualIds:{}};
    out.achievement[region] = {target:a.target, actual:a.actual, pct:a.target > 0 ? Number((a.actual / a.target * 100).toFixed(1)) : null, lops:count_(a.actualIds)};
    out.h1[region] = {ports:d.ports, lops:count_(d.ids)};
    out.onTime[region] = {ports:o.ports, eligiblePorts:o.eligiblePorts, pct:o.eligiblePorts > 0 ? Number((o.ports / o.eligiblePorts * 100).toFixed(1)) : null, lops:count_(o.ids)};
    out.overdue[region] = {ports:l.ports, lops:count_(l.ids)};
    out.overdueAanwijzing[region] = {ports:oa.ports, lops:count_(oa.ids)};
    out.overdueInstalasi[region] = {ports:oi.ports, lops:count_(oi.ids)};
    out.overdueFinishInstalasi[region] = {ports:ofi.ports, lops:count_(ofi.ids)};
    out.upcoming7[region] = {ports:u.ports, lops:count_(u.ids)};
    out.weekly[region] = {target:w.target, actual:w.actual, pct:w.target > 0 ? Number((w.actual / w.target * 100).toFixed(1)) : null, lops:count_(w.actualIds)};
  });

  return {
    callouts: {
      achievement: {target:achievement.target, actual:achievement.actual, pct:achievement.target > 0 ? Number((achievement.actual / achievement.target * 100).toFixed(1)) : null, lops:count_(achievement.actualIds), regions:out.achievement},
      h1: {ports:h1.ports, lops:count_(h1.ids), date:Utilities.formatDate(yesterday, CONFIG.timezone, 'yyyy-MM-dd'), regions:out.h1},
      onTime: {ports:onTime.ports, eligiblePorts:onTime.eligiblePorts, pct:onTime.eligiblePorts > 0 ? Number((onTime.ports / onTime.eligiblePorts * 100).toFixed(1)) : null, lops:count_(onTime.ids), regions:out.onTime},
      overdue: {ports:overdue.ports, lops:count_(overdue.ids), regions:out.overdue},
      overdueAanwijzing: {ports:overdueAanwijzing.ports, lops:count_(overdueAanwijzing.ids), regions:out.overdueAanwijzing},
      overdueInstalasi: {ports:overdueInstalasi.ports, lops:count_(overdueInstalasi.ids), regions:out.overdueInstalasi},
      overdueFinishInstalasi: {ports:overdueFinishInstalasi.ports, lops:count_(overdueFinishInstalasi.ids), regions:out.overdueFinishInstalasi},
      upcoming7: {ports:upcoming.ports, lops:count_(upcoming.ids), from:Utilities.formatDate(today, CONFIG.timezone, 'yyyy-MM-dd'), to:Utilities.formatDate(horizon, CONFIG.timezone, 'yyyy-MM-dd'), regions:out.upcoming7},
      weekly: {
        target:weekly.target,
        actual:weekly.actual,
        pct:weekly.target > 0 ? Number((weekly.actual / weekly.target * 100).toFixed(1)) : null,
        lops:count_(weekly.actualIds),
        weekNo,
        from:Utilities.formatDate(weekStart, CONFIG.timezone, 'yyyy-MM-dd'),
        to:Utilities.formatDate(weekEnd, CONFIG.timezone, 'yyyy-MM-dd'),
        regions:out.weekly
      }
    },
    records,
    dataQuality: {duplicateIdRows, missingRegionRows, invalidPortRows, stageCommitmentGaps}
  };
}

/**
 * Mencari sheet yang sesuai dengan preferensi nama, atau mencari sheet yang memiliki data yang sesuai
 */
function getActiveOrTargetSheet_(ss, preferredName) {
  if (preferredName) {
    const s = ss.getSheetByName(preferredName);
    if (s) return s;
  }
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const s = allSheets[i];
    const range = s.getDataRange();
    if (range.getNumRows() > 1) {
      const sample = s.getRange(1, 1, Math.min(10, s.getLastRow()), Math.max(1, s.getLastColumn())).getValues();
      if (findHeaderRow_(sample) >= 0) return s;
    }
  }
  return ss.getActiveSheet() || allSheets[0];
}

/**
 * Mencari baris index tempat header berada
 */
function findHeaderRow_(values) {
  return values.findIndex(row => {
    const headers = row.map(normalize_);
    return headers.some(h => h.includes('STATUS')) && headers.some(h => h.includes('REAL PORT'));
  });
}

function normalize_(value) {
  return String(value || '').trim().toUpperCase();
}

/**
 * Parsing tanggal serbaguna: mendukung Date object Sheets, ISO, D-Mon-YYYY (Indonesia/Inggris), DD/MM/YYYY
 */
function parseDate_(value) {
  if (!value && value !== 0) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    const y = Number(Utilities.formatDate(value, CONFIG.timezone, 'yyyy'));
    const m = Number(Utilities.formatDate(value, CONFIG.timezone, 'M')) - 1;
    const d = Number(Utilities.formatDate(value, CONFIG.timezone, 'd'));
    return new Date(y, m, d);
  }
  const raw = String(value || '').trim();
  if (!raw || raw === '#N/A' || raw === '-' || raw.toUpperCase() === 'NULL') return null;

  // Format ISO: YYYY-MM-DD
  let match = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  // Format Teks: DD-Mon-YYYY / DD Mon YYYY (e.g. 12-Sep-2026, 30-Agu-2026)
  match = raw.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{4})$/);
  if (match) {
    const months = {
      jan:0, januari:0, january:0,
      feb:1, februari:1, february:1,
      mar:2, maret:2, march:2,
      apr:3, april:3,
      mei:4, may:4,
      jun:5, juni:5, june:5,
      jul:6, juli:6, july:6,
      agu:7, agustus:7, aug:7, august:7,
      sep:8, september:8,
      okt:9, oktober:9, oct:9, october:9,
      nov:10, november:10,
      des:11, desember:11, dec:11, december:11
    };
    const m = months[match[2].toLowerCase()];
    if (m !== undefined) return new Date(Number(match[3]), m, Number(match[1]));
  }

  // Format Angka Pemisah: DD/MM/YYYY atau MM/DD/YYYY
  match = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (match) {
    const p1 = Number(match[1]);
    const p2 = Number(match[2]);
    const y = Number(match[3]);
    if (p1 > 12) return new Date(y, p2 - 1, p1);
    if (p2 > 12) return new Date(y, p1 - 1, p2);
    return new Date(y, p2 - 1, p1); // Default Indonesia: DD/MM/YYYY
  }
  return null;
}

/**
 * Konversi nilai port ke angka murni
 */
function toNumber_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  const str = String(value || '').trim();
  if (!str) return 0;
  const normalized = str.replace(/\./g, '').replace(/,/g, '').replace(/[^\d-]/g, '');
  const num = Number(normalized);
  return isNaN(num) ? 0 : num;
}

/**
 * Validasi nama callback JSONP agar aman
 */
function safeCallback_(callback) {
  return callback && /^[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*$/.test(callback) ? callback : '';
}

/**
 * Jalankan fungsi ini dari editor Apps Script jika ingin membersihkan cache data hari ini secara manual.
 */
function clearReportCache() {
  PropertiesService.getScriptProperties().setProperty('PT3_HISTORY_REV', Utilities.getUuid());
  console.log('Cache semua vendor diperbarui.');
}

/**
 * =========================================================================
 * DIAGNOSTIK DATA - Jalankan fungsi ini untuk MEMBUKTIKAN data terbaca atau tidak.
 * =========================================================================
 * Pilih 'testReport' pada dropdown toolbar Apps Script, klik Run, lalu buka
 * menu View > Logs (atau panel Execution log) untuk melihat ringkasannya.
 *
 * getReport() sendiri TIDAK mencetak apa pun karena ia mengembalikan nilai
 * (return). Itu sebabnya log hanya menampilkan "Eksekusi dimulai/selesai".
 * Fungsi ini mencetak hasilnya agar Anda bisa memastikan angkanya benar.
 */
function testReport() {
  const t0 = new Date();
  let report;
  try {
    report = getReport_({forceRefresh:true});
  } catch (err) {
    console.log('GAGAL: ' + err.message);
    return { ok: false, error: err.message };
  }
  const ms = new Date() - t0;

  console.log('========================================');
  console.log('HASIL PEMBACAAN DATA GOLIVE PT3');
  console.log('========================================');
  console.log('Status        : ' + (report.ok ? 'BERHASIL' : 'GAGAL'));
  console.log('Sumber data   : ' + report.source.sheet);
  console.log('Baris sumber  : ' + report.source.sourceRows);
  console.log('Baris Golive  : ' + report.source.matchedRows);
  console.log('Tanggal lapor : ' + report.reportDate + ' (' + report.timezone + ')');
  console.log('Durasi        : ' + ms + ' ms');
  console.log('Ukuran payload: ' + Math.round(JSON.stringify(report).length / 1024) + ' KB');
  console.log('----------------------------------------');
  console.log('PER REGION (akumulasi s.d. hari ini):');
  Object.keys(report.regions).forEach(reg => {
    const arr = report.regions[reg];
    const total = arr.reduce((a, b) => a + b, 0);
    console.log('  ' + reg + ' : ' + total + ' port');
  });
  console.log('  TOTAL   : ' + report.cumulative + ' port');
  console.log('----------------------------------------');
  const t = report.targetVsReal || {};
  console.log('TARGET VS REALISASI:');
  console.log('  Target bulan ini  : ' + (t.totalTarget || 0) + ' port');
  console.log('  Realisasi         : ' + (t.totalReal || 0) + ' port');
  console.log('  Capaian           : ' + (t.totalPct === null ? '—' : t.totalPct + ' %'));
  console.log('  Jumlah hari       : ' + (t.totalDays || 0));
  console.log('----------------------------------------');
  console.log('DETAIL UNTUK MODAL:');
  console.log('  Baris realisasi   : ' + report.records.length);
  console.log('  Baris target      : ' + report.targetRecords.length);
  console.log('----------------------------------------');
  const c = report.callouts || {};
  console.log('LIMA CALLOUT:');
  console.log('  Achievement s.d. H: ' + (c.achievement && c.achievement.pct !== null ? c.achievement.pct + '%' : '-'));
  console.log('  Golive H-1         : ' + ((c.h1 && c.h1.ports) || 0) + ' port');
  console.log('  On-time rate       : ' + (c.onTime && c.onTime.pct !== null ? c.onTime.pct + '%' : '-'));
  console.log('  Overdue            : ' + ((c.overdue && c.overdue.ports) || 0) + ' port');
  console.log('  Komitmen H s.d H+7 : ' + ((c.upcoming7 && c.upcoming7.ports) || 0) + ' port');
  console.log('  Duplikat ID (baris): ' + ((report.dataQuality && report.dataQuality.duplicateIdRows) || 0));
  const m = report.monitoring || {};
  const mr = m.reconciliation || {};
  console.log('----------------------------------------');
  console.log('MONITORING PROGRESS & PRIORITAS:');
  console.log('  Jumlah status      : ' + ((m.statuses && m.statuses.length) || 0));
  console.log('  Total progress     : ' + (mr.progressPort || 0) + ' port | Rp ' + (mr.progressBoq || 0));
  console.log('  Total prioritas    : ' + (mr.priorityPort || 0) + ' port | Rp ' + (mr.priorityBoq || 0));
  console.log('  Selisih port       : ' + (mr.portDifference || 0));
  console.log('  Selisih BOQ        : ' + (mr.boqDifference || 0));
  console.log('  Flag kosong        : ' + (mr.nonPriorityRows || 0) + ' baris (NON PRIORITY)');
  console.log('  Header flag tersedia: ' + (!mr.missingFlaggingHeader));
  console.log('========================================');
  console.log('Jika angka di atas sesuai, dashboard siap dipakai.');

  return report;
}


// PT3 v37: persistent, append-only per-LOP notes and observed commitment changes.
function pt3Spreadsheet_() {
  // Spreadsheet operasional dikunci eksplisit agar dashboard tidak bergantung
  // pada container aktif maupun cache published CSV.
  if (!CONFIG.sourceSpreadsheetId) return null;
  return SpreadsheetApp.openById(CONFIG.sourceSpreadsheetId);
}
function pt3HistorySpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('PT3_HISTORY_SPREADSHEET');
  if (!id) throw new Error('Jalankan setupLopHistory sekali dari project Apps Script ini untuk membuat spreadsheet history.');
  try { return SpreadsheetApp.openById(id); }
  catch (err) { throw new Error('Spreadsheet history tidak dapat dibuka. Pastikan akun deployment memiliki akses. ID history tetap dipertahankan: ' + id); }
}

function pt3Date_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const date = parseDate_(value);
  return date ? Utilities.formatDate(date, CONFIG.timezone, 'yyyy-MM-dd') : 'Tanggal tidak valid: ' + String(value);
}
function pt3Source_(suppliedValues, suppliedIndex) {
  let values = suppliedValues, index = suppliedIndex;
  if (!values) {
    if (!CONFIG.csvUrl) throw new Error('URL CSV sumber belum diisi.');
    const response = UrlFetchApp.fetch(CONFIG.csvUrl, {muteHttpExceptions:true});
    if (response.getResponseCode() !== 200) throw new Error('CSV sumber gagal dibaca (HTTP ' + response.getResponseCode() + '). Catatan belum disimpan.');
    values = Utilities.parseCsv(response.getContentText());
    index = findHeaderRow_(values);
  }
  if (index < 0 || !values[index]) throw new Error('Header sumber CSV tidak ditemukan.');
  const headers = values[index].map(normalize_);
  const idCol = headers.indexOf('ID-IHLD'), dateCol = headers.indexOf('KOMITMEN GOLIVE');
  if (idCol < 0 || dateCol < 0) throw new Error('History memerlukan ID-IHLD dan KOMITMEN GOLIVE.');
  const lopCol = headers.indexOf('NAMA LOP'), vendorCol = headers.indexOf('VENDOR');
  const records = {}, counts = {};
  values.slice(index + 1).forEach(row => {
    const id = String(row[idCol] || '').trim();
    if (!id) return;
    const key = 'ID:' + id;
    counts[key] = (counts[key] || 0) + 1;
    records[key] = {id:id, lop:String(row[lopCol] || '').trim(), vendor:vendorCol < 0 ? '' : String(row[vendorCol] || '').trim(), currentCommitment:pt3Date_(row[dateCol])};
  });
  Object.keys(counts).forEach(key => { if (counts[key] !== 1) delete records[key]; });
  return {key:'CSV:' + CONFIG.csvUrl, records:records};
}
function pt3LogSheet_(ss) {
  let sheet = ss.getSheetByName('_PT3_LOP_HISTORY');
  if (!sheet) {
    sheet = ss.insertSheet('_PT3_LOP_HISTORY');
    sheet.getRange(1,1,1,3).setValues([['EVENT_ID','RECORDED_AT','DATA_JSON']]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  const header = sheet.getRange(1,1,1,3).getValues()[0];
  if (header.join('|') !== 'EVENT_ID|RECORDED_AT|DATA_JSON') throw new Error('Struktur sheet history tidak sesuai; data tidak diubah.');
  return sheet;
}
function pt3Events_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2,1,sheet.getLastRow()-1,3).getValues().map(row => {
    let event;
    try { event = JSON.parse(row[2]); } catch (err) { throw new Error('Ada catatan history tidak valid; periksa sheet history.'); }
    event.eventId = String(row[0]); event.recordedAt = String(row[1]);
    return event;
  });
}
function pt3Append_(sheet, events) {
  if (!events.length) return;
  const now = new Date().toISOString();
  const rows = events.map(event => [event.eventId || Utilities.getUuid(), now, JSON.stringify(event)]);
  sheet.getRange(sheet.getLastRow()+1,1,rows.length,3).setValues(rows);
  SpreadsheetApp.flush();
  PropertiesService.getScriptProperties().setProperty('PT3_HISTORY_REV', Utilities.getUuid());
}
function pt3SyncLocked_(source, sheet, events) {
  const states = {};
  events.forEach(event => {
    if (event.sourceSheet !== source.key || (event.type !== 'BASELINE' && event.type !== 'COMMITMENT')) return;
    const key = 'ID:' + event.id;
    const prior = states[key] || {};
    states[key] = {currentCommitment:event.current, previousCommitment:event.type === 'COMMITMENT' ? event.previous : prior.previousCommitment,
      hasPrevious:event.type === 'COMMITMENT' || !!prior.hasPrevious};
  });
  const additions = [], meta = {};
  Object.keys(source.records).forEach(key => {
    const rec = source.records[key], old = states[key];
    if (!old || old.currentCommitment !== rec.currentCommitment) {
      const entry = {sourceSheet:source.key, id:rec.id, lop:rec.lop,
        type:old ? 'COMMITMENT' : 'BASELINE', current:rec.currentCommitment, previous:old ? old.currentCommitment : null};
      additions.push(entry);
      states[key] = {currentCommitment:rec.currentCommitment, previousCommitment:old ? old.currentCommitment : null, hasPrevious:!!old};
    }
    meta[key] = Object.assign({}, rec, states[key], {editable:true});
  });
  pt3Append_(sheet, additions);
  return {meta:meta, available:true, message:''};
}
function pt3SyncHistory_(values, headerIndex) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const source = pt3Source_(values, headerIndex), sheet = pt3LogSheet_(pt3HistorySpreadsheet_());
    return pt3SyncLocked_(source, sheet, pt3Events_(sheet));
  } finally { lock.releaseLock(); }
}
function setupLopHistory() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    // Validate the source before creating anything. Repeated setup reuses the same file.
    const source = pt3Source_();
    const props = PropertiesService.getScriptProperties();
    let id = props.getProperty('PT3_HISTORY_SPREADSHEET');
    let ss;
    if (id) {
      ss = pt3HistorySpreadsheet_();
    } else {
      ss = SpreadsheetApp.create('PT3 - History Progress LOP');
      id = ss.getId();
      // Persist immediately so a later initialization failure can be retried safely.
      props.setProperty('PT3_HISTORY_SPREADSHEET', id);
    }
    const sheet = pt3LogSheet_(ss);
    pt3SyncLocked_(source, sheet, pt3Events_(sheet));
    props.setProperty('PT3_HISTORY_REV', Utilities.getUuid());
    const result = 'Plan B aktif. History tersimpan di: https://docs.google.com/spreadsheets/d/' + id + '/edit';
    console.log(result);
    return result;
  } finally { lock.releaseLock(); }
}
function pt3LopSourceEdited(e) {
  // Legacy entry point retained for old deployments. Plan B observes published CSV on reads.
}

function getLopHistory(id) {
  const key = 'ID:' + String(id || '').trim();
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const source = pt3Source_();
    if (!source.records[key]) throw new Error('ID-IHLD kosong, duplikat, atau sudah tidak ada. History tidak dapat dipasangkan dengan aman.');
    const sheet = pt3LogSheet_(pt3HistorySpreadsheet_());
    const state = pt3SyncLocked_(source, sheet, pt3Events_(sheet));
    const events = pt3Events_(sheet).filter(event => event.sourceSheet === source.key && 'ID:' + event.id === key).reverse();
    return {ok:true, meta:state.meta[key], events:events};
  } finally { lock.releaseLock(); }
}
function saveLopProgress(input) {
  input = input || {};
  const id = String(input.id || '').trim(), note = String(input.note || '').trim(), requestId = String(input.requestId || '');
  if (!id || !note || note.length > 4000) throw new Error('Isi catatan 1–4.000 karakter dan pilih LOP yang valid.');
  if (!/^[a-zA-Z0-9-]{16,80}$/.test(requestId)) throw new Error('Identitas penyimpanan tidak valid. Buka ulang detail LOP.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const source = pt3Source_(), rec = source.records['ID:' + id];
    if (!rec) throw new Error('ID-IHLD kosong, duplikat, atau tidak ditemukan. Catatan tidak disimpan.');
    const sheet = pt3LogSheet_(pt3HistorySpreadsheet_()), events = pt3Events_(sheet);
    const existing = events.find(event => event.eventId === requestId);
    if (existing) {
      if (existing.id !== id || existing.note !== note) throw new Error('Identitas penyimpanan sudah dipakai. Buka ulang detail.');
      return {ok:true, duplicate:true};
    }
    pt3SyncLocked_(source, sheet, events);
    // Never substitute the deployment owner's identity for an unidentified viewer.
    const actor = Session.getActiveUser().getEmail() || 'Identitas pengguna tidak tersedia';
    pt3Append_(sheet, [{eventId:requestId, sourceSheet:source.key, type:'NOTE', id:id, lop:rec.lop, note:note, actor:actor}]);
    return {ok:true};
  } finally { lock.releaseLock(); }
}
