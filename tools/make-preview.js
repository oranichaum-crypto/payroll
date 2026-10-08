#!/usr/bin/env node
/**
 * make-preview.js — สร้าง preview.html จาก Index.html ให้เปิดดูหน้าเว็บในเบราว์เซอร์ปกติได้ โดยไม่ต้อง deploy
 * ฉีด mock ของ google.script.run (ข้อมูลสมมติใน preview-mock.json) เข้าไปก่อน script ของแอป
 *
 * รัน (จากโฟลเดอร์ template):  node tools/make-preview.js   → ได้ preview.html ข้าง Index.html
 * เพิ่ม API ใหม่ใน Code.gs ต้องเติม key ใน preview-mock.json ด้วย ไม่งั้นพรีวิวเรียกแล้ว TypeError หน้าขาว
 * ห้ามอัปไฟล์ preview.html ขึ้น Apps Script (.claspignore กันไว้แล้ว)
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'Index.html');
const OUT = path.join(ROOT, 'preview.html');
const MOCK = path.join(__dirname, 'preview-mock.json');
const ANCHOR = '<div class="toast" id="toast"></div>';
// api ที่ตัวจริง return JSON.stringify(...) — mock ต้องคืน string เหมือนกัน
const STRING_APIS = ['apiBootstrap', 'apiPayrollReport', 'apiListPayslips', 'apiOcrTimesheet', 'apiSaveTimesheetRows'];

function buildHtml(html, R) {
  if (!html.includes(ANCHOR)) throw new Error('หาจุดฉีดไม่เจอใน Index.html: ' + ANCHOR);
  R = JSON.parse(JSON.stringify(R));
  STRING_APIS.forEach((k) => { if (R[k] && typeof R[k] !== 'string') R[k] = JSON.stringify(R[k]); });
  const data = JSON.stringify(R).replace(/<\//g, '<\\/');
  const mock = `${ANCHOR}
<!-- ===== พรีวิวเท่านั้น: mock google.script.run (สร้างโดย tools/make-preview.js) ===== -->
<script>
(function(){
 var R=${data};
 // api ที่ต้องสะท้อนค่าที่ส่งไป (พรีวิวโลโก้/ตั้งค่า) — ที่เหลือคืนค่าคงที่จาก preview-mock.json
 var ECHO={
   apiSaveLogo:function(a){return {ok:true,logo:a[1]};},
   apiSaveSettings:function(a){var b=JSON.parse(R.apiBootstrap);Object.assign(b.settings,a[1]||{});R.apiBootstrap=JSON.stringify(b);return {ok:true};}
 };
 function runner(){var self={};self.withSuccessHandler=function(f){self._s=f;return self;};self.withFailureHandler=function(f){self._f=f;return self;};
   Object.keys(Object.assign({},R,ECHO)).forEach(function(n){self[n]=function(){var s=self._s,f=self._f,a=[].slice.call(arguments);
     setTimeout(function(){try{s&&s(ECHO[n]?ECHO[n](a):R[n]);}catch(e){f&&f(e);}},60);};});return self;}
 window.google={script:{run:runner()}};
 window.addEventListener('load',function(){var p=(location.hash||'').slice(1);if(p)setTimeout(function(){try{go(p);}catch(e){}},400);});
})();
</script>`;
  let out = html.replace(ANCHOR, mock);
  if (!/<meta[^>]+name=["']viewport/i.test(out)) {
    out = out.replace('<meta charset="utf-8">', '<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">');
  }
  return out;
}

if (require.main === module) {
  const out = buildHtml(fs.readFileSync(SRC, 'utf8'), JSON.parse(fs.readFileSync(MOCK, 'utf8')));
  fs.writeFileSync(OUT, out, 'utf8');
  console.log('สร้าง ' + OUT + ' แล้ว — เปิดด้วยเบราว์เซอร์ได้เลย');
}
module.exports = { buildHtml };
