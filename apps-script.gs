// 重慶之旅記帳同步：貼到 Google 試算表的「擴充功能 → Apps Script」
const SHEET = '花費';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(['id', '日期代號', '項目', '分類', '金額', '幣別', '記錄時間']);
    s.setFrozenRows(1);
  }
  return s;
}

function list_() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  return {
    ok: true,
    exp: rows.filter(r => r[0] !== '').map(r => ({
      id: String(r[0]), day: String(r[1]), name: String(r[2]).replace(/^'/, ''),
      cat: String(r[3]), amt: Number(r[4]) || 0, cur: String(r[5])
    }))
  };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function safe_(v) {
  v = String(v == null ? '' : v).slice(0, 200);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doGet() {
  return out_(list_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);
    const ops = body.action === 'batch' ? body.ops : [body];
    const s = sheet_();
    ops.forEach(op => {
      const ids = s.getRange(1, 1, Math.max(s.getLastRow(), 1), 1).getValues().map(r => String(r[0]));
      if (op.action === 'add' && op.item && ids.indexOf(String(op.item.id)) < 0) {
        const it = op.item;
        s.appendRow([safe_(it.id), safe_(it.day), safe_(it.name), safe_(it.cat), Number(it.amt) || 0,
          it.cur === 'TWD' ? 'TWD' : 'CNY', new Date()]);
      }
      if (op.action === 'del') {
        const i = ids.indexOf(String(op.id));
        if (i > 0) s.deleteRow(i + 1);
      }
    });
    return out_(list_());
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
