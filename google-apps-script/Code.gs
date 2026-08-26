const SPREADSHEET_ID = '19Fg7u6M37RK3snZOtXt12mTZRwXzhiimPvju763Ihbg';
const PRODUCT_SHEET = 'สินค้า';
const ORDER_SHEET = 'ออร์เดอร์';
const ITEM_SHEET = 'รายการในออร์เดอร์';
const SETTINGS_SHEET = 'ตั้งค่าร้าน';
const SHIPPING_RATE_SHEET = 'เรทค่าจัดส่ง';
const MAX_SLIP_BYTES = 8 * 1024 * 1024;

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'store');
    if (action === 'store') return json_({ ok: true, products: getProducts_(), settings: getSettings_() });
    if (action === 'track') return json_(getOrderTracking_(e.parameter.orderNumber, e.parameter.phone));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (error) {
    return json_({ ok: false, error: safeError_(error) });
  }
}

function getOrderTracking_(orderNumber, phone) {
  const wantedOrder = String(orderNumber || '').trim().toUpperCase();
  const wantedPhone = String(phone || '').replace(/\D/g, '');
  if (!/^SS-\d{8}-\d{4}$/.test(wantedOrder) || wantedPhone.length < 9) return { ok: false, error: 'กรุณาตรวจสอบเลขออร์เดอร์และเบอร์โทรศัพท์' };
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ORDER_SHEET);
  if (!sheet) throw new Error('ไม่พบแท็บออร์เดอร์');
  const values = sheet.getDataRange().getDisplayValues();
  const row = values.slice(1).find(item => String(item[0] || '').trim().toUpperCase() === wantedOrder && String(item[3] || '').replace(/\D/g, '') === wantedPhone);
  if (!row) return { ok: false, error: 'ไม่พบคำสั่งซื้อ กรุณาตรวจสอบข้อมูลอีกครั้ง' };
  return {
    ok: true,
    orderNumber: row[0],
    paymentStatus: row[12] || 'รอตรวจสอบ',
    fulfillmentStatus: row[14] || 'รอจัดเตรียม',
    trackingNumber: row[15] || ''
  };
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (payload.action !== 'createOrder') return json_({ ok: false, error: 'Unknown action' });
    return json_(createOrder_(payload));
  } catch (error) {
    return json_({ ok: false, error: safeError_(error) });
  }
}

function getProducts_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(PRODUCT_SHEET);
  if (!sheet) throw new Error('ไม่พบแท็บสินค้า');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  return values.slice(1).filter(row => row[0] && row[5] === 'พร้อมขาย').map(row => ({
    sku: row[0], name: row[1], category: row[2], price: number_(row[3]), stock: number_(row[4]),
    status: row[5], mediaType: row[6], imageUrl: cleanUrl_(row[7]), videoUrl: cleanUrl_(row[8]),
    size: row[9], color: row[10], description: row[11]
  })).filter(product => product.stock > 0);
}

function getSettings_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SETTINGS_SHEET);
  if (!sheet) throw new Error('ไม่พบแท็บตั้งค่าร้าน');
  const values = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), 2).getDisplayValues();
  const map = {};
  values.forEach(row => { if (row[0]) map[String(row[0]).trim()] = String(row[1] || '').trim(); });
  return {
    storeName: map['ชื่อร้าน'] || 'Smile Shirt',
    depositRate: 0.5,
    storeAddress: map['ที่อยู่รับหน้าร้าน'] || '',
    lineId: map['LINE ID'] || '',
    bankName: map['ธนาคาร'] || 'รอเพิ่มข้อมูลจริง',
    accountName: map['ชื่อบัญชี'] || 'รอเพิ่มข้อมูลจริง',
    accountNumber: map['เลขบัญชี'] || 'รอเพิ่มข้อมูลจริง',
    qrUrl: cleanUrl_(map['QR Code'] || map['ลิงก์ QR Code'] || ''),
    routeVideoUrl: cleanUrl_(map['วิดีโอแนะนำทางมารับหน้าร้าน'] || ''),
    promoImageUrl: cleanUrl_(map['รูปภาพโปรโมชั่นหน้าแรก'] || ''),
    promoQualityText: map['ข้อความป้ายคุณภาพ'] || 'คัดคุณภาพทุกตัว',
    promoPriceText: map['ข้อความป้ายราคาโปรโมชั่น'] || 'เริ่มต้น ฿120',
    shippingMinimum: number_(map['ค่าขนส่งขั้นต่ำ']) || 60,
    codRate: number_(map['อัตราค่าบริการเก็บเงินปลายทาง']) || 0.04,
    shippingRates: getShippingRates_()
  };
}

function getShippingRates_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHIPPING_RATE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [{ minItems: 1, maxItems: 999, privateFee: 60, postFee: 60 }];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().filter(row => row[4] === 'ใช้งาน').map(row => ({
    minItems: Number(row[0]) || 1,
    maxItems: Number(row[1]) || 999,
    privateFee: Number(row[2]) || 0,
    postFee: Number(row[3]) || 0
  }));
}

function createOrder_(payload) {
  validateText_(payload.customerName, 'กรุณากรอกชื่อผู้รับ', 120);
  validateText_(payload.phone, 'กรุณากรอกเบอร์โทร', 30);
  validateText_(payload.shippingMethod, 'กรุณาเลือกวิธีรับสินค้า', 50);
  validateText_(payload.address, 'กรุณากรอกที่อยู่หรือหมายเหตุรับหน้าร้าน', 500);
  if (!Array.isArray(payload.items) || !payload.items.length || payload.items.length > 30) throw new Error('ไม่มีสินค้าในออร์เดอร์');
  if (!payload.slip || !payload.slip.data) throw new Error('กรุณาแนบสลิป');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const productSheet = ss.getSheetByName(PRODUCT_SHEET);
  const orderSheet = ss.getSheetByName(ORDER_SHEET);
  const itemSheet = ss.getSheetByName(ITEM_SHEET);
  if (!productSheet || !orderSheet || !itemSheet) throw new Error('โครงสร้าง Google Sheets ไม่ครบ');

  const productRows = productSheet.getDataRange().getValues();
  const catalog = {};
  productRows.slice(1).forEach((row, index) => {
    if (row[0]) catalog[String(row[0])] = { row: index + 2, name: row[1], price: Number(row[3]) || 0, stock: Number(row[4]) || 0, status: row[5] };
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let subtotal = 0;
    const cleanItems = payload.items.map(item => {
      const sku = String(item.sku || '').trim();
      const quantity = Math.max(1, Math.min(10, Number(item.quantity) || 1));
      const product = catalog[sku];
      if (!product || product.status !== 'พร้อมขาย') throw new Error('สินค้าบางรายการไม่พร้อมขาย');
      if (quantity > product.stock) throw new Error('สินค้า ' + product.name + ' มีจำนวนไม่เพียงพอ');
      subtotal += product.price * quantity;
      return { sku, name: product.name, price: product.price, quantity, row: product.row, stock: product.stock };
    });

    const settings = getSettings_();
    const itemCount = cleanItems.reduce((sum, item) => sum + item.quantity, 0);
    const shippingFee = shippingFee_(payload.shippingMethod, itemCount, settings);
    const codFee = payload.shippingMethod === 'รับหน้าร้าน' ? 0 : moneyRound_((subtotal + shippingFee) * settings.codRate);
    const total = moneyRound_(subtotal + shippingFee + codFee);
    const deposit = roundUpHalf_(total * 0.5);
    const balance = moneyRound_(total - deposit);
    const orderNumber = nextOrderNumber_();
    const slipUrl = saveSlip_(payload.slip, orderNumber);
    const now = new Date();

    orderSheet.appendRow([
      orderNumber, now, clean_(payload.customerName), clean_(payload.phone), clean_(payload.lineId),
      clean_(payload.shippingMethod), clean_(payload.address), subtotal, shippingFee, total, deposit, balance,
      'รอตรวจสอบ', slipUrl, 'รอจัดเตรียม', '', 'ชำระผ่าน ' + clean_(payload.paymentMethod || 'ไม่ระบุ'), codFee
    ]);
    cleanItems.forEach(item => {
      itemSheet.appendRow([orderNumber, item.sku, item.name, item.quantity, item.price, item.price * item.quantity, '']);
      productSheet.getRange(item.row, 5).setValue(item.stock - item.quantity);
      if (item.stock - item.quantity <= 0) productSheet.getRange(item.row, 6).setValue('หมด');
    });
    SpreadsheetApp.flush();
    return { ok: true, orderNumber, subtotal, shippingFee, codFee, total, deposit, balance };
  } finally {
    lock.releaseLock();
  }
}

function saveSlip_(slip, orderNumber) {
  const mime = String(slip.type || 'application/octet-stream');
  const allowed = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
  if (allowed.indexOf(mime) === -1) throw new Error('รองรับสลิปเฉพาะ JPG, PNG, WEBP หรือ PDF');
  const bytes = Utilities.base64Decode(String(slip.data || ''));
  if (!bytes.length || bytes.length > MAX_SLIP_BYTES) throw new Error('ไฟล์สลิปต้องไม่เกิน 8 MB');
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('SLIP_FOLDER_ID');
  let folder;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (_) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder('Smile Shirt - หลักฐานการโอน');
    props.setProperty('SLIP_FOLDER_ID', folder.getId());
  }
  const extension = mime === 'application/pdf' ? '.pdf' : mime === 'image/png' ? '.png' : mime === 'image/webp' ? '.webp' : '.jpg';
  const file = folder.createFile(Utilities.newBlob(bytes, mime, orderNumber + extension));
  return file.getUrl();
}

function nextOrderNumber_() {
  const tz = 'Asia/Bangkok';
  const date = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'ORDER_COUNTER_' + date;
  const next = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(next));
  return 'SS-' + date + '-' + String(next).padStart(4, '0');
}

function cleanUrl_(value) {
  const text = String(value || '').trim();
  if (!/^https:\/\//i.test(text)) return '';
  const driveMatch = text.match(/drive\.google\.com\/(?:file\/d\/|open\?id=)([a-zA-Z0-9_-]+)/);
  return driveMatch ? 'https://drive.google.com/uc?export=view&id=' + driveMatch[1] : text;
}
function clean_(value) { return String(value || '').trim().replace(/[<>]/g, ''); }
function validateText_(value, message, max) { const text = clean_(value); if (!text || text.length > max) throw new Error(message); }
function number_(value) { return Number(String(value || '').replace(/[^0-9.-]/g, '')) || 0; }
function moneyRound_(value) { return Math.round((Number(value) || 0) * 100) / 100; }
function roundUpHalf_(value) { return Math.ceil((Number(value) || 0) * 2) / 2; }
function shippingFee_(method, itemCount, settings) {
  if (method === 'รับหน้าร้าน') return 0;
  const rates = settings.shippingRates || [];
  const rate = rates.find(item => itemCount >= item.minItems && itemCount <= item.maxItems) || {};
  const configured = method === 'ไปรษณีย์ไทย' ? Number(rate.postFee) : Number(rate.privateFee);
  return Math.max(Number(settings.shippingMinimum) || 60, configured || 0);
}
function safeError_(error) { return error && error.message ? String(error.message).slice(0, 250) : 'เกิดข้อผิดพลาด'; }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
