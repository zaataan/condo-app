// ════════════════════════════════════════════════════════
// Google Apps Script — Condo Manager Backend
// รองรับทุกไฟล์: the-base-manage-final, the-base-sheets, property-manager
//
// วิธีใช้:
//   1. Extensions → Apps Script → วาง code นี้แทน code เดิม
//   2. Deploy → Manage deployments → กด Edit (✏️) → Version: New version → Deploy
//   3. URL เดิมยังใช้ได้ ไม่ต้องเปลี่ยน
// ════════════════════════════════════════════════════════

const TH_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
                   'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

// ── GET: ?path=thebase | propmanager | thebase_sheets ──
function doGet(e) {
  try {
    const path  = (e.parameter && e.parameter.path) || 'thebase';
    const sheet = getOrCreate('Data_' + path);
    const raw   = sheet.getRange('A1').getValue();
    return out(raw || '{}');
  } catch(err) {
    return out(JSON.stringify({ error: err.message }));
  }
}

// ── POST: body JSON must include { path: '...', ...data } ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const path    = payload.path || 'thebase';
    const ss      = SpreadsheetApp.getActiveSpreadsheet();

    // บันทึก current state
    getOrCreate('Data_' + path).getRange('A1').setValue(e.postData.contents);

    // แยก tab รายเดือน (เฉพาะไฟล์ที่มี extras/pool)
    if(payload.extras || payload.pool) {
      writeMonthlyTabs(ss, payload, path);
    }

    return out(JSON.stringify({ status: 'ok' }));
  } catch(err) {
    return out(JSON.stringify({ error: err.message }));
  }
}

// ────────────────────────────────────────────────────────
// แยกรายการตามเดือน → เขียนลง tab
// ────────────────────────────────────────────────────────
function writeMonthlyTabs(ss, data, path) {
  const prefix = path === 'propmanager' ? '[PM] ' : '[TB] ';
  const items  = [];

  // extras (room income/expenses)
  const extras = data.extras || {};
  Object.keys(extras).forEach(roomId => {
    const room = extras[roomId];
    (room.income || []).forEach(r => {
      if(!r.amt || !r.date) return;
      items.push({ date:r.date, month:monthKey(r.date), room:'ห้อง '+roomId,
        type:'รายรับ', desc:r.desc||'', income:parseFloat(r.amt)||0, expense:0 });
    });
    (room.expenses || []).forEach(r => {
      if(!r.amt || !r.date) return;
      items.push({ date:r.date, month:monthKey(r.date), room:'ห้อง '+roomId,
        type:'รายจ่าย', desc:r.desc||'', income:0, expense:parseFloat(r.amt)||0 });
    });
  });

  // pool
  const pool = data.pool || {};
  (pool.income || []).forEach(r => {
    if(!r.amt || !r.date) return;
    items.push({ date:r.date, month:monthKey(r.date), room:'กองกลาง',
      type:'รายรับ', desc:r.desc||'', income:parseFloat(r.amt)||0, expense:0 });
  });
  (pool.expenses || []).forEach(r => {
    if(!r.amt || !r.date) return;
    items.push({ date:r.date, month:monthKey(r.date), room:'กองกลาง',
      type:'รายจ่าย', desc:r.desc||'', income:0, expense:parseFloat(r.amt)||0 });
  });

  // ค่าเช่าจาก rooms (property-manager structure)
  if(data.props) {
    (data.props || []).forEach(prop => {
      (prop.rooms || []).forEach(room => {
        if(room.rented && room.rent) {
          const today = new Date().toISOString().slice(0,10);
          items.push({ date:today, month:monthKey(today),
            room: prop.condoName || prop.name,
            type:'รายรับ', desc:'ค่าเช่า '+(room.name||''),
            income:parseFloat(room.rent)||0, expense:0 });
        }
      });
    });
  }

  // จัดกลุ่มตามเดือน
  const byMonth = {};
  items.forEach(item => {
    if(!byMonth[item.month]) byMonth[item.month] = [];
    byMonth[item.month].push(item);
  });

  Object.keys(byMonth).forEach(mk => {
    const tabName = prefix + monthTabName(mk);
    writeMonthSheet(getOrCreate(tabName), byMonth[mk], tabName);
  });
}

function writeMonthSheet(sheet, items, tabName) {
  sheet.clearContents();
  sheet.getRange(1,1,1,6).setValues([['📅 ' + tabName,'','','','','']]);
  sheet.getRange(1,1,1,6).setFontWeight('bold').setFontSize(13);

  const cols = ['วันที่','ห้อง/หมวด','ประเภท','รายการ','รายรับ (บาท)','รายจ่าย (บาท)'];
  sheet.getRange(2,1,1,6).setValues([cols]).setFontWeight('bold')
    .setBackground('#1a6fc4').setFontColor('#ffffff');

  items.sort((a,b) => a.date.localeCompare(b.date));
  const rows = items.map(it => [
    it.date, it.room, it.type, it.desc,
    it.income > 0 ? it.income : '',
    it.expense > 0 ? it.expense : ''
  ]);

  if(rows.length > 0) {
    sheet.getRange(3,1,rows.length,6).setValues(rows);
    rows.forEach((r,i) => {
      sheet.getRange(3+i,1,1,6).setBackground(i%2===0?'#f9fafb':'#ffffff');
      if(r[2]==='รายรับ')  sheet.getRange(3+i,5).setFontColor('#16a34a');
      if(r[2]==='รายจ่าย') sheet.getRange(3+i,6).setFontColor('#dc2626');
    });
  }

  const sr       = rows.length + 3;
  const totalInc = items.reduce((s,it)=>s+it.income, 0);
  const totalExp = items.reduce((s,it)=>s+it.expense, 0);
  const net      = totalInc - totalExp;
  sheet.getRange(sr,1,1,6).setValues([['','','','รวม',totalInc,totalExp]])
    .setFontWeight('bold').setBackground('#e8f1fb');
  sheet.getRange(sr+1,1,1,6)
    .setValues([['','','','Net Cash Flow', net>=0?net:'', net<0?net:'']])
    .setFontWeight('bold')
    .setBackground(net>=0?'#dcfce7':'#fee2e2')
    .setFontColor(net>=0?'#16a34a':'#dc2626');

  [100,120,80,200,120,120].forEach((w,i) => sheet.setColumnWidth(i+1,w));
}

function monthKey(d)     { return d ? d.slice(0,7) : ''; }
function monthTabName(mk) {
  if(!mk) return 'ไม่ระบุ';
  const [y,m] = mk.split('-');
  return TH_MONTHS[parseInt(m)-1] + ' ' + ((parseInt(y)+543)%100);
}
function getOrCreate(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function out(json) {
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
