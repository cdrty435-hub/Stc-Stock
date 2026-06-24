// ============================================================
// STC Die Casting Spare Parts – Google Apps Script
// PARTS_MASTER: PART_ID / GROUP / PIC / TYPE / DESCRIPTION / 초기STOCK / MIN / MAX / CURRENT / STATUS / QR_DATA
// IN:  TX_ID / DATE / PART_ID / DESCRIPTION / QTY / WORKER / REMARK
// OUT: TX_ID / DATE / PART_ID / DESCRIPTION / QTY / WORKER / MACHINE / REMARK
// ============================================================

const SS_NAME = 'DC_Parts_Management';

function getSpreadsheet() {
  const files = DriveApp.getFilesByName(SS_NAME);
  if (!files.hasNext()) throw new Error('스프레드시트를 찾을 수 없습니다: ' + SS_NAME);
  return SpreadsheetApp.open(files.next());
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = (data.type || '').toUpperCase();
    if (type === 'IN') recordIN(data);
    else if (type === 'OUT') recordOUT(data);
    else throw new Error('type은 IN 또는 OUT이어야 합니다');
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', type }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const partId = e.parameter.partId;
    if (!partId) throw new Error('partId가 필요합니다');
    const stock = getCurrentStock(partId);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', partId, stock }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function recordIN(data) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName('IN');
  const now = new Date();
  const txId = 'IN-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMddHHmmss');
  sh.appendRow([
    txId,
    Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
    data.partId || '',
    data.description || '',
    Number(data.qty) || 0,
    data.by || 'App',
    data.remark || ''
  ]);
}

function recordOUT(data) {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName('OUT');
  const now = new Date();
  const txId = 'OUT-' + Utilities.formatDate(now, 'Asia/Bangkok', 'yyyyMMddHHmmss');
  sh.appendRow([
    txId,
    Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
    data.partId || '',
    data.description || '',
    Number(data.qty) || 0,
    data.by || 'App',
    data.machine || '',
    data.remark || ''
  ]);
}

function getCurrentStock(partId) {
  const ss = getSpreadsheet();
  const master = ss.getSheetByName('PARTS_MASTER');
  const inSh   = ss.getSheetByName('IN');
  const outSh  = ss.getSheetByName('OUT');

  // 초기STOCK
  const mData = master.getDataRange().getValues();
  const mH = mData[0];
  const pidCol  = mH.indexOf('PART_ID');
  const iniCol  = mH.indexOf('초기STOCK');
  let initial = 0;
  for (let i = 1; i < mData.length; i++) {
    if (String(mData[i][pidCol]) === String(partId)) {
      initial = Number(mData[i][iniCol]) || 0;
      break;
    }
  }

  // IN 합계
  const inData = inSh.getDataRange().getValues();
  const inPid = inData[0].indexOf('PART_ID');
  const inQty = inData[0].indexOf('QTY');
  let inTotal = 0;
  for (let i = 1; i < inData.length; i++) {
    if (String(inData[i][inPid]) === String(partId))
      inTotal += Number(inData[i][inQty]) || 0;
  }

  // OUT 합계
  const outData = outSh.getDataRange().getValues();
  const outPid = outData[0].indexOf('PART_ID');
  const outQty = outData[0].indexOf('QTY');
  let outTotal = 0;
  for (let i = 1; i < outData.length; i++) {
    if (String(outData[i][outPid]) === String(partId))
      outTotal += Number(outData[i][outQty]) || 0;
  }

  return initial + inTotal - outTotal;
}
