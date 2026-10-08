/**
 * TimesheetOcr.gs — AI อ่านรูปใบลงเวลา (ฟอร์ม PAY-TS A4 ของเทมเพลต)
 *
 * พอร์ตจาก Nano Bot src/timesheet.js (prompt + normDate/normTime + resolveRowDate + todayBangkok
 * + responseSchema/thinking fallback) มาเป็น Apps Script V8 ที่ใช้ UrlFetchApp
 *
 * ★ แชร์ global scope กับ Code.gs — ชื่อ global ทุกตัวในไฟล์นี้ขึ้นต้นด้วย "tsocr"
 *   ยกเว้นทางเข้า ocrTimesheetImage_ (สัญญากับ Code.gs apiOcrTimesheet) · ไม่มี top-level const/let
 * ★ คีย์อยู่ใน Script Properties "GEMINI_API_KEY" เท่านั้น (ใส่หลายคีย์คั่นด้วย , ได้ — เจอโควต้าเต็ม 429 จะสลับคีย์ถัดไป)
 *   ส่งคีย์ทาง header x-goog-api-key ไม่ใส่ใน URL · ข้อความ error/log ถูกล้างคีย์ออกทุกครั้ง
 * ★ ไม่อ่าน OT / พัก / ชม.ปกติ — ระบบเงินเดือนคำนวณเองจากเวลาเข้า-ออก
 * ★ อ่าน 1 รอบ temperature 0 (ไม่โหวตหลายรอบเหมือนต้นทาง — ประหยัดโควต้าและเวลา 6 นาทีของ Apps Script)
 */

/** base64: string ไม่มี prefix data:, mimeType: 'image/jpeg'|'image/png', opts: {period_from:'YYYY-MM-DD', period_to:'YYYY-MM-DD'} (ไม่บังคับ)
 *  return object ธรรมดา (ไม่มี Date) :
 *  {ok:true, model:'...', header:{name:'',employee_id:'',period_from:'',period_to:''},
 *   rows:[{work_date:'YYYY-MM-DD'|'', check_in:'HH:MM'|'', check_out:'HH:MM'|'', raw_date:'', raw_in:'', raw_out:'', unreadable:boolean, future:boolean, note:''}],
 *   warnings:['...']}
 *  หรือ {ok:false, error:'ข้อความไทย'} */
function ocrTimesheetImage_(base64, mimeType, opts) {
  opts = opts || {};
  var props = PropertiesService.getScriptProperties();
  var keys = tsocrSplitKeys_(props.getProperty('GEMINI_API_KEY'));
  if (!keys.length) return { ok: false, error: 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน Script Properties' };

  var data = String(base64 == null ? '' : base64).replace(/^data:[^,]*,/, '').replace(/\s+/g, '');
  if (!data) return { ok: false, error: 'ยังไม่ได้แนบรูปใบลงเวลา' };
  var mime = String(mimeType || '').toLowerCase().trim();
  if (mime === 'image/jpg') mime = 'image/jpeg';
  if (!/^(image\/(png|jpeg|webp|heic|heif)|application\/pdf)$/.test(mime)) {
    return { ok: false, error: 'รองรับเฉพาะรูป JPG / PNG (หรือ WEBP / HEIC / PDF)' };
  }
  if (Math.floor(data.length * 3 / 4) > 15 * 1024 * 1024) {
    return { ok: false, error: 'ไฟล์ใหญ่เกิน 15 MB — ถ่ายใหม่ให้เล็กลง หรือสแกนเป็น JPG' };
  }

  var today = tsocrTodayBangkok_();
  var models = tsocrModels_(props);
  var res;
  try {
    res = tsocrCallGemini_(keys, models, data, mime, today);
  } catch (e) {
    tsocrLog_('exception: ' + tsocrScrub_(String(e && e.message || e), keys));
    return { ok: false, error: 'AI อ่านรูปไม่สำเร็จ ลองใหม่อีกครั้ง หรือกรอกตารางเอง' };
  }
  if (!res.ok) {
    tsocrLog_('failed: ' + tsocrScrub_(res.errors.join(' | '), keys));
    return { ok: false, error: res.userError || 'AI อ่านรูปไม่สำเร็จ ลองใหม่อีกครั้ง หรือกรอกตารางเอง' };
  }
  return tsocrBuildResult_(res.data, res.model, opts, today);
}

/* ───────────────────────── แปลงผลของ Gemini → สัญญา ───────────────────────── */

function tsocrBuildResult_(d, model, opts, today) {
  d = d || {};
  if (d.is_timesheet === false) {
    var kind = tsocrStr_(d.doc_kind);
    return { ok: false, error: 'รูปนี้ไม่ใช่ใบลงเวลา' + (kind ? ' (ดูเหมือน ' + kind + ')' : '') + ' — ถ่ายใบลงเวลาใหม่ให้เห็นทั้งแผ่น' };
  }
  var warnings = [];
  var header = {
    name: tsocrStr_(d.employee_name).replace(/\s+/g, ' '),
    employee_id: tsocrThaiDigits_(tsocrStr_(d.employee_id)).replace(/[\s\-_.|\[\]]+/g, '').toUpperCase(),
    period_from: tsocrNormDate_(d.period_start),
    period_to: tsocrNormDate_(d.period_end)
  };
  if (header.period_from && !header.period_to) header.period_to = header.period_from;
  if (!header.name && !header.employee_id) warnings.push('อ่านชื่อ/รหัสพนักงานที่หัวกระดาษไม่ออก — กรอกเองก่อนบันทึก');

  // งวดที่ใช้เติมปีให้วันที่ในตาราง: งวดที่ผู้ใช้เลือก (opts) มาก่อน ไม่งั้นใช้งวดที่หัวกระดาษ
  var oFrom = tsocrNormDate_(opts.period_from), oTo = tsocrNormDate_(opts.period_to);
  var period = oFrom ? { start: oFrom, end: oTo || oFrom } : { start: header.period_from, end: header.period_to || header.period_from };
  if (oFrom && header.period_from && (header.period_from !== oFrom || (oTo && header.period_to && header.period_to !== oTo))) {
    warnings.push('งวดบนกระดาษ (' + header.period_from + ' ถึง ' + header.period_to + ') ไม่ตรงกับงวดที่เลือก (' + oFrom + ' ถึง ' + (oTo || oFrom) + ')');
  }

  var src = Array.isArray(d.rows) ? d.rows : [];
  var rows = [];
  for (var i = 0; i < src.length; i++) {
    var r = src[i] || {};
    var rawIn = tsocrStr_(r.check_in), rawOut = tsocrStr_(r.check_out);
    if (!rawIn && !rawOut) continue;                       // แถวที่ไม่มีเวลาเลย = วันหยุด/ไม่ได้มา
    var rawDate = tsocrStr_(r.date_raw);
    var ci = tsocrNormTime_(rawIn), co = tsocrNormTime_(rawOut);
    var wd = tsocrResolveRowDate_(rawDate, tsocrStr_(r.work_date), period, today);
    var notes = [];
    if (!wd) notes.push('อ่านวันที่ไม่ออก' + (rawDate ? ' (' + rawDate + ')' : ''));
    if (rawIn && !ci) notes.push('เวลาเข้าไม่ถูกต้อง (' + rawIn + ')');
    if (rawOut && !co) notes.push('เวลาออกไม่ถูกต้อง (' + rawOut + ')');
    if (!rawIn) notes.push('ไม่มีเวลาเข้า');
    if (!rawOut) notes.push('ไม่มีเวลาออก');
    var unreadable = !wd || (!!rawIn && !ci) || (!!rawOut && !co);
    var future = !!wd && wd > today;
    if (future) notes.push('วันที่ล่วงหน้า (วันนี้ ' + today + ')');
    rows.push({
      work_date: wd, check_in: ci, check_out: co,
      raw_date: rawDate, raw_in: rawIn, raw_out: rawOut,
      unreadable: unreadable, future: future, note: notes.join(' · ')
    });
  }

  var nUnread = rows.filter(function (x) { return x.unreadable; }).length;
  var fut = rows.filter(function (x) { return x.future; }).map(function (x) { return x.work_date; });
  if (!rows.length) warnings.push('ไม่พบแถวที่มีเวลาเข้า-ออก' + (tsocrStr_(d.summary) ? ' — ' + tsocrStr_(d.summary) : ''));
  if (nUnread) warnings.push('มี ' + nUnread + ' แถวที่อ่านวันที่/เวลาไม่ออก — ตรวจแก้ก่อนบันทึก');
  if (fut.length) warnings.push('มีวันที่ล่วงหน้า ' + fut.length + ' แถว (' + fut.join(', ') + ') — วันนี้ ' + today + ' ตรวจว่าอ่านวัน/เดือนถูกไหม');
  var seen = {}, dup = [];
  rows.forEach(function (x) { if (x.work_date) { if (seen[x.work_date] && dup.indexOf(x.work_date) < 0) dup.push(x.work_date); seen[x.work_date] = 1; } });
  if (dup.length) warnings.push('วันที่ซ้ำกันในใบเดียว: ' + dup.join(', '));
  if (period.start) {
    var out = rows.filter(function (x) { return x.work_date && (x.work_date < period.start || x.work_date > period.end); }).map(function (x) { return x.work_date; });
    if (out.length) warnings.push('วันที่อยู่นอกงวด ' + period.start + ' ถึง ' + period.end + ': ' + out.join(', '));
  }
  return { ok: true, model: model, header: header, rows: rows, warnings: warnings };
}

/* ───────────────────────── เรียก Gemini (UrlFetchApp) ───────────────────────── */

/** ลองรุ่นตามลำดับ · 400 (ไม่รับ schema/thinking) → ถอดทีละชั้น · 429 → คีย์ถัดไป
 *  · คีย์ผิด/โปรเจกต์ถูกระงับ → คีย์ถัดไป (หมดคีย์ = หยุด) · 401/403 อื่น → รุ่นถัดไป */
function tsocrCallGemini_(keys, models, data, mime, today) {
  var prompt = tsocrPrompt_(today);
  var errors = [];
  var ki = 0, lastMsg = '', pendingAuth = '';
  for (var mi = 0; mi < models.length; mi++) {
    var model = models[mi];
    var think = tsocrThinking_(model);
    var variants = [{ schema: true, think: think }];
    if (think) variants.push({ schema: true, think: null });
    variants.push({ schema: false, think: null });
    for (var vi = 0; vi < variants.length; vi++) {
      var v = variants[vi];
      var gc = { responseMimeType: 'application/json', temperature: 0, maxOutputTokens: 8192 };
      if (v.schema) gc.responseSchema = tsocrSchema_();
      if (v.think) gc.thinkingConfig = v.think;
      var body = {
        contents: [{ role: 'user', parts: [{ text: prompt }, { inline_data: { mime_type: mime, data: data } }] }],
        generationConfig: gc
      };
      var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
      var resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(body),
        headers: { 'x-goog-api-key': keys[ki] },
        muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      var txt = resp.getContentText();
      if (code === 429 && ki + 1 < keys.length) { ki++; vi--; errors.push(model + ' 429 → คีย์ถัดไป'); continue; }
      var authErr = tsocrAuthError_(code, txt, keys);
      if (authErr) {
        errors.push(model + ' ' + code + ': ' + txt.slice(0, 160));
        if (tsocrKeyLevel_(code, txt)) {                  // คีย์ผิด/โปรเจกต์ถูกระงับ → คีย์ถัดไป · หมดคีย์ = หยุด
          if (ki + 1 < keys.length) { ki++; vi--; continue; }
          return { ok: false, errors: errors, userError: authErr };
        }
        pendingAuth = authErr; break;                     // 403 อื่น (อาจเฉพาะรุ่น) → ลองรุ่นถัดไปก่อน
      }
      if (code < 200 || code >= 300) lastMsg = model + ' ' + code + ': ' + tsocrGoogleMsg_(txt, keys);
      if (code === 400) { errors.push(model + ' 400: ' + txt.slice(0, 160)); continue; }   // ถอด schema/thinking
      if (code < 200 || code >= 300) { errors.push(model + ' ' + code + ': ' + txt.slice(0, 160)); break; }   // รุ่นถัดไป
      var dj;
      try { dj = JSON.parse(txt); } catch (e) { lastMsg = model + ': ตอบไม่ใช่ JSON'; errors.push(lastMsg); break; }
      var cand = dj && dj.candidates && dj.candidates[0];
      var parts = (cand && cand.content && cand.content.parts) || [];
      var text = parts.filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
      if (!text) {
        lastMsg = model + ': ไม่มีผลลัพธ์ (' + ((cand && cand.finishReason) || (dj && dj.promptFeedback && dj.promptFeedback.blockReason) || '?') + ')';
        errors.push(lastMsg);
        break;
      }
      try {
        var s = text.indexOf('{'), e2 = text.lastIndexOf('}');
        return { ok: true, model: model, data: JSON.parse(s >= 0 && e2 > s ? text.slice(s, e2 + 1) : text) };
      } catch (e3) { lastMsg = model + ': แปลง JSON ไม่ได้'; errors.push(lastMsg); break; }
    }
  }
  var all429 = errors.length && errors.every(function (x) { return / 429/.test(x); });
  return { ok: false, errors: errors, userError: all429 ? 'โควต้า AI ของวันนี้เต็มแล้ว — ลองใหม่ภายหลัง หรือกรอกตารางเอง'
    : pendingAuth ? pendingAuth
    : lastMsg ? 'AI อ่านรูปไม่สำเร็จ ลองใหม่อีกครั้ง หรือกรอกตารางเอง (Google ตอบ: ' + lastMsg + ')' : '' };
}

/** error ระดับคีย์ (ใช้คีย์นี้ต่อกับรุ่นไหนก็ไม่ผ่าน): 400 คีย์ผิด · 401/403 โปรเจกต์ถูกระงับ */
function tsocrKeyLevel_(code, txt) {
  return (code === 400 && /API[_ ]?key/i.test(txt)) || ((code === 401 || code === 403) && /denied access/i.test(txt));
}

/** ข้อความจาก Google (error.message) ล้างคีย์แล้ว ตัดสั้น — ปลอดภัยที่จะส่งให้หน้าเว็บ */
function tsocrGoogleMsg_(txt, keys) {
  var m = '';
  try { var j = JSON.parse(txt); m = j && j.error && j.error.message || ''; } catch (e) { /* ไม่ใช่ JSON */ }
  return tsocrScrub_(String(m || txt || '').replace(/\s+/g, ' ').trim(), keys).slice(0, 160);
}

/** 400 คีย์ผิด / 401 / 403 → ข้อความไทยบอกทางแก้ · อย่างอื่นคืน '' */
function tsocrAuthError_(code, txt, keys) {
  var why = ' (Google ตอบ: ' + tsocrGoogleMsg_(txt, keys) + ')';
  if (code === 400 && /API[_ ]?key/i.test(txt)) return 'คีย์ GEMINI_API_KEY ใน Script Properties ใช้ไม่ได้ — ตรวจคีย์อีกครั้ง' + why;
  if (code !== 401 && code !== 403) return '';
  if (/denied access/i.test(txt)) {
    return 'Google ระงับโปรเจกต์ที่ออกคีย์นี้ — สร้างคีย์ใหม่ที่ aistudio.google.com/apikey โดยเลือก "สร้างในโปรเจกต์ใหม่" แล้วใส่ทับ GEMINI_API_KEY' + why;
  }
  return 'คีย์ GEMINI_API_KEY ไม่มีสิทธิ์เรียก AI — ตรวจคีย์/โปรเจกต์ใน Google AI Studio' + why;
}

/** ปุ่ม "ทดสอบคีย์" ในหน้าตั้งค่า: ส่งข้อความสั้น 1 ครั้ง (ไม่มีรูป) ตามลำดับรุ่นเดียวกับตอนอ่านจริง
 *  return {ok:true, model} หรือ {ok:false, error:'ข้อความไทย'} */
function tsocrPing_() {
  var props = PropertiesService.getScriptProperties();
  var keys = tsocrSplitKeys_(props.getProperty('GEMINI_API_KEY'));
  if (!keys.length) return { ok: false, error: 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน Script Properties' };
  // เว็บเปิดไม่ต้องล็อกอิน → กันคนกดรัวจนโควต้าหมด: 1 ครั้งต่อ 30 วิ ทั้งระบบ
  var cache = CacheService.getScriptCache();
  if (cache.get('tsocr_ping')) return { ok: false, error: 'เพิ่งทดสอบไปเมื่อครู่ — รอ 30 วินาทีแล้วกดใหม่' };
  cache.put('tsocr_ping', '1', 30);
  var models = tsocrModels_(props), ki = 0, last = '', all429 = true, pendingAuth = '';
  try {
    for (var mi = 0; mi < models.length; mi++) {
      var resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(models[mi]) + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: 'ตอบว่า OK' }] }], generationConfig: { maxOutputTokens: 256 } }),
        headers: { 'x-goog-api-key': keys[ki] }
      });
      var code = resp.getResponseCode(), txt = resp.getContentText();
      if (code === 429 && ki + 1 < keys.length) { ki++; mi--; continue; }
      var authErr = tsocrAuthError_(code, txt, keys);
      if (authErr) {
        all429 = false;
        tsocrLog_('ping ' + models[mi] + ' ' + code + ': ' + tsocrGoogleMsg_(txt, keys));
        if (tsocrKeyLevel_(code, txt)) {
          if (ki + 1 < keys.length) { ki++; mi--; continue; }
          return { ok: false, error: authErr };
        }
        pendingAuth = authErr; continue;
      }
      if (code >= 200 && code < 300) return { ok: true, model: models[mi] };
      if (code !== 429) all429 = false;
      last = models[mi] + ' ' + code + ': ' + tsocrGoogleMsg_(txt, keys);
      tsocrLog_('ping ' + last);
    }
    if (pendingAuth) return { ok: false, error: pendingAuth };
  } catch (e) {
    var m = tsocrScrub_(String(e && e.message || e), keys);
    tsocrLog_('ping exception: ' + m);
    return { ok: false, error: 'เรียก AI ไม่สำเร็จ: ' + m.slice(0, 160) + ' — ถ้าเพิ่งติดตั้ง ให้ Run authorizeUrlFetch ในเอดิเตอร์ก่อน' };
  }
  return { ok: false, error: all429 ? 'โควต้า AI ของวันนี้เต็มแล้ว — ลองใหม่ภายหลัง' : 'เรียก AI ไม่สำเร็จทุกรุ่น (' + last + ')' };
}

function tsocrModels_(props) {
  // 2.5 จำกัดเฉพาะผู้ใช้เดิมตั้งแต่ ต.ค. 69 (คีย์ใหม่เรียกไม่ได้) → ตั้งต้นด้วย 3.5 แล้วค่อยถอยไป 2.5
  var arr = [props.getProperty('GEMINI_MODEL'), props.getProperty('GEMINI_FALLBACK_MODEL'), 'gemini-3.5-flash', 'gemini-2.5-flash'];
  return arr.map(function (m) { return String(m || '').trim(); })
    .filter(function (m, i, a) { return m && a.indexOf(m) === i; });
}

/** 2.5 ใช้ thinkingBudget, 3.x ใช้ thinkingLevel (รุ่นไหนไม่รับ → ถอยไปแบบไม่ส่ง) */
function tsocrThinking_(model) {
  if (/^gemini-2\.5/.test(model)) return { thinkingBudget: 2048 };
  if (/^gemini-3/.test(model)) return { thinkingLevel: 'low' };
  return null;
}

function tsocrSplitKeys_(v) {
  return String(v || '').split(',').map(function (s) { return s.trim(); })
    .filter(function (k, i, a) { return k && a.indexOf(k) === i; });
}

/** ล้างคีย์ออกจากข้อความก่อน log เสมอ */
function tsocrScrub_(s, keys) {
  s = String(s || '');
  (keys || []).forEach(function (k) { if (k) s = s.split(k).join('***'); });
  return s.replace(/AIza[0-9A-Za-z_\-]{20,}/g, '***');
}

function tsocrLog_(msg) {
  try { console.log('[TimesheetOcr] ' + msg); } catch (e) { /* ignore */ }
}

/* ───────────────────────── prompt + schema ───────────────────────── */

function tsocrPrompt_(today) {
  return 'วันนี้คือ ' + today + ' (เขตเวลา Asia/Bangkok)\n' +
'รูป/ไฟล์ที่ส่งมา "อาจจะใช่หรือไม่ใช่" ใบลงเวลาทำงานก็ได้ — ดูก่อนว่าใช่ไหม แล้วตอบเป็น JSON ออบเจ็กต์เดียวตาม schema เท่านั้น ห้ามมีข้อความอื่น\n' +
'\n' +
'ขั้นที่ 1 — ตัดสินก่อนว่าใช่ใบลงเวลาไหม\n' +
'- ใบลงเวลา = เอกสารที่มี "ตารางรายวัน + ช่องเวลาเข้า/เวลาออก" ให้พนักงานกรอก\n' +
'- ใบเสร็จ ใบกำกับภาษี ใบแจ้งหนี้ ใบเสนอราคา สัญญา สลิปโอนเงิน บัตรประชาชน รูปถ่ายทั่วไป สกรีนช็อต = ไม่ใช่ใบลงเวลา\n' +
'- อย่าฝืนตีความเอกสารอื่นให้เป็นใบลงเวลา — ถ้าไม่ใช่: is_timesheet=false ใส่ doc_kind (ประเภทเอกสารสั้น ๆ) + summary (1-2 ประโยค) และ rows=[]\n' +
'\n' +
'ขั้นที่ 2 — ถ้าใช่ อ่านตามโครงฟอร์ม "ใบลงเวลาทำงาน" A4 (หัวกระดาษเป็นโลโก้และชื่อบริษัท)\n' +
'- หัวกระดาษ: โลโก้ + ช่อง "ชื่อ-สกุล" (กล่องยาว) · "รหัสพนักงาน" (ช่องสี่เหลี่ยม 8 ช่อง ช่องละ 1 ตัวอักษร) · "งวดวันที่ ... ถึง ..." — อาจพิมพ์มาแล้วหรือเขียนด้วยลายมือ\n' +
'  → employee_name / employee_id อ่านจากหัวกระดาษ (ไม่ใช่จากในตาราง) · employee_id ให้ต่อตัวอักษรจากช่องที่มีค่าเรียงซ้าย→ขวา ไม่เว้นวรรค ช่องว่างที่เหลือไม่ต้องใส่อะไร\n' +
'  → งวดเขียนเป็นช่อง วว / ดด / 25ปป โดยเลข "25" พิมพ์ไว้แล้ว = ปี พ.ศ. 25xx (เช่น 01/09/2569)\n' +
'  → "งวด" คือช่วงของใบ ไม่ใช่วันทำงาน ห้ามใส่ใน rows: ใส่ข้อความตามที่เห็นใน period_text และแปลงเป็น period_start / period_end รูปแบบ YYYY-MM-DD (ค.ศ.) ถ้าอ่านได้\n' +
'  → มุมล่างซ้ายมีรหัสฟอร์ม เช่น "PAY-TS · ใบที่ 1/1" (ฉบับเก่าเขียน "แผ่นที่ 1/1") ใส่ใน form_code\n' +
'- ช่อง "ลายมือชื่อ" ข้างรหัสพนักงาน กล่อง "วิธีกรอก" และช่อง "หัวหน้างานตรวจ" ท้ายกระดาษ ไม่ใช่ข้อมูลลงเวลา\n' +
'- ตารางเรียงซ้าย→ขวา: วันที่ (วัน / เดือน) | เวลาเข้า | เวลาออก | รวม (ชั่วโมง) · ไม่เกิน 16 แถวต่อแผ่น (จำนวนแถวอาจน้อยกว่า)\n' +
'  → ใบรุ่นแรกไม่มีคอลัมน์ "รวม" (มีแค่ 3 คอลัมน์) — อ่านได้ทั้งสองรุ่น 3 คอลัมน์แรกเรียงเหมือนกัน\n' +
'  → คอลัมน์ "รวม" เป็นจำนวนชั่วโมงที่พนักงานเขียนเอง (เช่น 8 หรือ 8.5) **ไม่ต้องอ่าน และห้ามเอาไปใส่เป็นเวลาออกเด็ดขาด** — เวลาออกอยู่ในช่องสี่เหลี่ยม [ ][ ]:[ ][ ] เท่านั้น\n' +
'  → ใบอาจเป็นครึ่งหน้า A4 และตารางอาจแบ่ง 2 บล็อกซ้าย-ขวา (แต่ละบล็อกมีหัว วันที่ | เวลาเข้า | เวลาออก | รวม ของตัวเอง) — อ่านบล็อกซ้ายจากบนลงล่างก่อน แล้วต่อด้วยบล็อกขวาจากบนลงล่าง\n' +
'  → เวลาเป็นแบบ 24 ชั่วโมง เขียนเป็น "ตัวเลขทีละหลักในช่องสี่เหลี่ยม" 2 ช่อง : 2 ช่อง โดยจุด ":" พิมพ์คั่นไว้แล้ว เช่น [0][8]:[3][0] = 08:30, [1][7]:[4][5] = 17:45\n' +
'  → อ่านตัวเลขทีละช่อง ห้ามอ่านจุด ":" ที่พิมพ์ไว้เป็นตัวเลข ห้ามเอาเลขจากแถวอื่นมาปน\n' +
'- ฟอร์มมี 2 แบบ:\n' +
'  (ก) แบบเปล่า — ช่องวันที่เป็นช่องสี่เหลี่ยม [ ][ ] / [ ][ ] พนักงานเขียน วว/ดด เอง\n' +
'  (ข) แบบเติมให้ — วันที่พิมพ์มาแล้วทุกแถว (มีชื่อวันย่อ เช่น "อ. 01/09") แถวที่ไม่มีเวลาเข้า-ออก = วันหยุด/ไม่ได้มาทำงาน → ข้ามแถวนั้น\n' +
'\n' +
'กฎของ rows:\n' +
'- 1 วัน = 1 แถว เรียงจากบนลงล่างตามที่เห็นในรูป ห้ามยุบหรือรวมหลายวันเป็นแถวเดียว แม้เวลาจะเหมือนกัน\n' +
'- ข้ามแถวที่ไม่มีทั้งเวลาเข้าและเวลาออก (ช่องเวลาว่างทั้งหมด = ไม่ต้องใส่แถวนั้น)\n' +
'- date_raw = วันที่ของแถวตามที่เห็นจริงบนกระดาษ เช่น "03/09" (ไม่ต้องแปลง ไม่ต้องใส่ชื่อวัน)\n' +
'- work_date = วันที่เต็ม YYYY-MM-DD (ค.ศ.) — เดือน/ปีอ้างอิงจากงวดที่หัวกระดาษ; ปี พ.ศ. ให้ลบ 543 (พ.ศ. 2 หลัก เช่น 69 = 2569 = ค.ศ. 2026)\n' +
'- check_in / check_out = "HH:MM" แบบ 24 ชั่วโมง; ช่องที่ว่างจริงให้เป็น ""\n' +
'- ใส่เฉพาะค่าที่อ่านได้จริง ห้ามเดาค่าที่มองไม่เห็น — แต่ถ้าเลขกำกวม ให้ใช้บริบทช่วย (ปกติเวลาออกมากกว่าเวลาเข้า และเวลาเข้า-ออกของแต่ละวันมักใกล้เคียงกัน)\n' +
'- ตัวเลขลายมือที่สับสนบ่อย: 0↔8, 1↔7, 4↔9, 3↔8, 5↔6, 2↔7 — ดูรูปทรงให้ดี\n' +
'- ห้ามอ่านหรือใส่ OT / พัก / ชั่วโมงทำงาน / ช่อง "รวม" (ระบบเงินเดือนคำนวณชั่วโมงเองจากเวลาเข้า-ออก)\n' +
'- ถ้าอ่านไม่ได้เลยสักแถว หรือไม่มีแถวไหนกรอกเวลา ให้ rows=[] แล้วอธิบายเหตุผลใน summary\n' +
'ตอบ JSON เท่านั้น';
}

/** โครง JSON ที่บังคับให้ Gemini ตอบ — ถ้ารุ่นไหนไม่รับ จะถอยไปใช้ JSON mode ธรรมดา */
function tsocrSchema_() {
  return {
    type: 'OBJECT',
    properties: {
      is_timesheet: { type: 'BOOLEAN' },
      doc_kind: { type: 'STRING' },
      summary: { type: 'STRING' },
      form_code: { type: 'STRING' },
      employee_id: { type: 'STRING' },
      employee_name: { type: 'STRING' },
      period_text: { type: 'STRING' },
      period_start: { type: 'STRING' },
      period_end: { type: 'STRING' },
      rows: {
        type: 'ARRAY',
        items: {
          type: 'OBJECT',
          properties: {
            date_raw: { type: 'STRING' },
            work_date: { type: 'STRING' },
            check_in: { type: 'STRING' },
            check_out: { type: 'STRING' }
          },
          required: ['date_raw', 'work_date', 'check_in', 'check_out']
        }
      }
    },
    required: ['is_timesheet', 'summary', 'employee_id', 'employee_name', 'period_text', 'rows']
  };
}

/* ───────────────────────── วันที่ / เวลา ───────────────────────── */

/** วันนี้ตามเวลาไทย (YYYY-MM-DD) — ไทยไม่มี DST จึงบวก 7 ชม. จาก UTC ได้ตรง ๆ ไม่ขึ้นกับ time zone ของสคริปต์ */
function tsocrTodayBangkok_(ms) {
  var t = (typeof ms === 'number' ? ms : Date.now()) + 7 * 3600 * 1000;
  return new Date(t).toISOString().slice(0, 10);
}

function tsocrStr_(v) { return v == null ? '' : String(v).trim(); }

function tsocrThaiDigits_(s) {
  return String(s == null ? '' : s).replace(/[๐-๙]/g, function (c) { return String('๐๑๒๓๔๕๖๗๘๙'.indexOf(c)); });
}

function tsocrPad2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

function tsocrValidYmd_(y, m, d) {
  if (!(y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return false;
  var dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

/** ปี 2 หลัก: >= 40 ถือเป็น พ.ศ. (69 = 2569) ไม่งั้น ค.ศ. (26 = 2026) · 4 หลัก > 2400 = พ.ศ. */
function tsocrFullYear_(y) {
  var n = Number(y);
  if (!isFinite(n)) return null;
  if (n < 100) n = n >= 40 ? 2500 + n : 2000 + n;
  if (n > 2400) n -= 543;
  return n;
}

/** สตริงวันที่เต็ม → YYYY-MM-DD (ค.ศ.) รับ 2026-09-01 / 2569-09-01 / 01/09/2569 / 1.9.69 / เลขไทย · ไม่ได้ = "" */
function tsocrNormDate_(s) {
  var t = tsocrThaiDigits_(s).trim();
  if (!t) return '';
  var m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/), y;
  if (m) { y = tsocrFullYear_(m[1]); return tsocrValidYmd_(y, +m[2], +m[3]) ? y + '-' + tsocrPad2_(+m[2]) + '-' + tsocrPad2_(+m[3]) : ''; }
  m = t.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2}|\d{4})$/);
  if (m) { y = tsocrFullYear_(m[3]); return tsocrValidYmd_(y, +m[2], +m[1]) ? y + '-' + tsocrPad2_(+m[2]) + '-' + tsocrPad2_(+m[1]) : ''; }
  return '';
}

/** เวลา → "HH:MM" 24 ชม. (รับ 8.30 / 8:30 / 0830 / 830 / ๐๘๓๐ / "08 30") · ไม่ได้ = "" */
function tsocrNormTime_(s) {
  var t = tsocrThaiDigits_(s).trim().replace(/\s+/g, ' ');
  if (!t) return '';
  var h, mi, m = t.match(/^(\d{1,2})\s*[:.,;\s]\s*(\d{2})$/);
  if (m) { h = +m[1]; mi = +m[2]; }
  else if ((m = t.match(/^(\d{3,4})$/))) { var v = ('0000' + m[1]).slice(-4); h = +v.slice(0, 2); mi = +v.slice(2); }
  else if ((m = t.match(/^(\d{1,2})$/))) { h = +m[1]; mi = 0; }
  else return '';
  if (h > 23 || mi > 59) return '';
  return tsocrPad2_(h) + ':' + tsocrPad2_(mi);
}

function tsocrDaysBetween_(a, b) {
  return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);
}

/** วันที่ของแถว: ถ้า raw มี วว/ดด(/ปป) → ใช้ตัวเลขที่เห็นจริง เติมปีจากงวด (เลือกปีที่ใกล้ช่วงงวดที่สุด ข้ามปีได้)
 *  ไม่งั้นใช้ work_date ที่โมเดลแปลงให้ · คืน "YYYY-MM-DD" หรือ "" */
function tsocrResolveRowDate_(raw, modelDate, period, today) {
  var r = tsocrThaiDigits_(raw).trim();
  var full = tsocrNormDate_(r);                      // raw เป็นวันที่เต็มอยู่แล้ว (มีปี) → ใช้เลย
  if (full) return full;
  // ตัดชื่อวันนำหน้าได้ เช่น "อ. 01/09" · (?:^|\D) กันจับกลางตัวเลขปี
  var m = r.match(/(?:^|\D)(\d{1,2})\s*[-\/.]\s*(\d{1,2})(?:\s*[-\/.]\s*(\d{4}|\d{2}))?(?!\d)/);
  if (m) {
    var d = +m[1], mo = +m[2];
    if (m[3]) {
      var y = tsocrFullYear_(m[3]);
      if (tsocrValidYmd_(y, mo, d)) return y + '-' + tsocrPad2_(mo) + '-' + tsocrPad2_(d);
    } else {
      var years = [];
      [period.start, period.end, modelDate, today].forEach(function (s) {
        var yy = Number(String(s || '').slice(0, 4));
        if (yy > 1900 && years.indexOf(yy) < 0) years.push(yy);
      });
      var best = '', bestDist = Infinity;
      var anchorA = period.start || today, anchorB = period.end || period.start || today;
      years.forEach(function (yy) {
        if (!tsocrValidYmd_(yy, mo, d)) return;
        var cand = yy + '-' + tsocrPad2_(mo) + '-' + tsocrPad2_(d);
        var dist = cand < anchorA ? tsocrDaysBetween_(cand, anchorA) : cand > anchorB ? tsocrDaysBetween_(anchorB, cand) : 0;
        if (dist < bestDist) { best = cand; bestDist = dist; }
      });
      if (best) return best;
    }
  }
  return tsocrNormDate_(modelDate);
}
