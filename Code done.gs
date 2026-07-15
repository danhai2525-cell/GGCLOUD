/**
 * ============================================================
 * GG CLOUD — BACKEND CRM & AFFILIATE & CDK (Google Apps Script)
 * ============================================================
 * CÁCH DÙNG:
 * 1. Tạo Google Spreadsheet mới → Extensions → Apps Script → dán toàn bộ file này.
 * 2. Chạy hàm setupDatabase() 1 lần (menu Run) để tạo các tab + dữ liệu mẫu.
 * 3. Deploy → New deployment → Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy URL /exec dán vào Tweaks "apiUrl" trên trang Frontend & CRM.
 * 5. Luồng đơn: Khách đặt → KTV kích hoạt → Khách kiểm tra → KTV gửi bill → Admin duyệt → paid_success.
 *    Mọi thanh toán qua KTV, KHÔNG QR/CDK tự động.
 * ============================================================
 */

// ================== CẤU HÌNH ==================
var WEBHOOK_URL = ''; // TODO: dán URL Bot Telegram/n8n khi có. Để trống = bỏ qua webhook.
// ===== THÔNG BÁO ĐƠN HÀNG QUA ZALO BOT =====
// 1. Tạo bot tại https://zalo.me/s/botcreator/ (tên bắt buộc dạng "Bot ...") → nhận token dạng {bot_id}:{access_token}
// 2. Thêm bot vào NHÓM ZALO (admin + CTV) → nhắn 1 tin trong nhóm → lấy chat_id nhóm qua getUpdates
//    (xem HƯỚNG DẪN DEPLOY.md mục Zalo Bot)
// 3. Dán token + danh sách chat_id dưới. Có thể gửi nhiều nơi cùng lúc (nhóm + cá nhân).
//    Để trống = không gửi Zalo.
var TELEGRAM_BOT_TOKEN = '8916450324:AAHgiI4mIN5WYpSdN-sISD3Ofr6SUfRdXpE';
var TELEGRAM_CHAT_IDS = ['872658011'];
var DRIVE_FOLDER_NAME = 'GGCLOUD_PAYMENT_PROOFS';
var FOMO_MINUTES = 30;
var COMMISSION_RATE = 0.20;      // Hoa hồng CTV 20% — CHỈ tính cho gói có ctv:true (199K/299K/399K)
var CTV_DISCOUNT_RATE = 0.05;    // Khách nhập mã CTV: giảm 5% — CHỈ gói có ctv:true (199K/299K/399K)

// ===== RÚT TIỀN HOA HỒNG (CTV) =====
var MAX_WITHDRAWALS_PER_DAY = 2;   // Giới hạn số lệnh rút / CTV / ngày
var MIN_WITHDRAW_AMOUNT = 50000;   // Số tiền rút tối thiểu mỗi lệnh
var ADMIN_CHAT_URL = 'https://t.me/+DGh5eC0Ce2EwOTc9';  // Link chat Admin (CTV bấm khi đơn rút quá 4 giờ).

// flow: 'manual'  = Nâng cấp trước – Kiểm tra – Thanh toán sau (KTV kích hoạt, KTV tự gửi thông tin thanh toán)
// ctv: true = gói được tính hoa hồng CTV 20% + khách nhập mã được giảm 5% (chỉ áp dụng gói 12 tháng)
// CHIẾN LƯỢC MỚI: 3 gói, mọi khâu qua KTV, KHÔNG QR/CDK tự động.
var PACKAGES = {
  '1THANG':  { amount: 29000,  cogs: 5000,  flow: 'manual', ctv: false }, // 1 tháng dùng thử — trả sau, KHÔNG hoa hồng CTV
  '12THANG': { amount: 499000, cogs: 20000, flow: 'manual', ctv: true },  // 12 tháng toàn diện — trả sau, hoa hồng CTV 20%
  '18THANG': { amount: 699000, cogs: 25000, flow: 'manual', ctv: true }   // 18 tháng ultimate — trả sau, hoa hồng CTV 20%
};

// ================== KHỞI TẠO DATABASE ==================
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {
    sheet_users: ['id', 'username', 'password_hash', 'full_name', 'role', 'status'],
    sheet_orders: ['id', 'customer_gmail', 'customer_phone', 'package_type', 'original_amount', 'discount_amount', 'amount', 'status', 'ktv_id', 'commission_ktv', 'applied_ctv_code', 'commission_ctv', 'cogs_amount', 'created_at', 'activated_at', 'payment_proof_url', 'cdk_code'],
    sheet_ctvs: ['id', 'gmail', 'ctv_code', 'balance_accumulated', 'status'],
    sheet_cdk_pool: ['id', 'cdk_code', 'status', 'order_id', 'distributed_at'],
    sheet_transactions: ['id', 'type', 'ctv_code', 'order_id', 'amount', 'note', 'created_at'],
    sheet_withdrawals: ['id', 'ctv_code', 'gmail', 'amount', 'bank_name', 'bank_account', 'account_holder', 'status', 'created_at', 'processed_at']
  };
  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, defs[name].length).setValues([defs[name]]).setFontWeight('bold');
    }
  });
  // Bổ sung cột cdk_code cho sheet_orders cũ (nếu nâng cấp từ bản trước)
  var os = ss.getSheetByName('sheet_orders');
  if (os.getRange(1, 17).getValue() !== 'cdk_code') os.getRange(1, 17).setValue('cdk_code').setFontWeight('bold');

  // Tài khoản mẫu (đổi mật khẩu sau khi chạy!): admin/admin123, tho01/tho123
  var users = ss.getSheetByName('sheet_users');
  if (users.getLastRow() === 1) {
    users.appendRow(['TKAdmin01', 'admin', sha256_('admin123'), 'Danh', 'ADMIN', 'Hoạt động']);
    users.appendRow(['TKTho01', 'tho01', sha256_('tho123'), 'Hương', 'KTV', 'Hoạt động']);
  }
  // CTV mẫu để test đăng nhập Ví CTV (mật khẩu cố định trên Frontend là cvt123)
  var ctvs = ss.getSheetByName('sheet_ctvs');
  if (ctvs.getLastRow() === 1) {
    ctvs.appendRow(['CTV001', 'demo@gmail.com', 'CTV-DEMO01', 0, 1]);
  }

}

// ================== ROUTER ==================
function doGet(e) {
  return route_(e, null);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  // Webhook đã bỏ (không còn QR/Seepay)
  return route_(e, body);
}

function route_(e, body) {
  var action = (e.parameter && e.parameter.action) || (body && body.action) || '';
  var p = body || e.parameter || {};
  var out;
  try {
    switch (action) {
      case 'createOrder':   out = createOrder(p); break;
      case 'getStatus':     out = getOrderStatus(e.parameter.order_id || p.order_id); break;
      case 'activateOrder': out = activateOrder(p); break;
      case 'completeOrder': out = completeOrder(p); break;
      case 'registerCTV':   out = registerCTV(p); break;
      case 'loginCTV':      out = loginCTV(p); break;
      case 'checkCTV':      out = checkCTV(e.parameter.ctv_code || p.ctv_code, e.parameter.package_type || p.package_type); break;
      case 'login':         out = login(p); break;
      case 'getOrders':     out = getOrders(e.parameter || p); break;
      case 'getCTVs':       out = getCTVs(); break;
      case 'getDashboard':  out = getDashboard(p); break;
      case 'ctvPortal':     out = ctvPortal(p); break;
      case 'requestWithdraw': out = requestWithdraw(p); break;
      case 'getWithdrawals':  out = getWithdrawals(e.parameter || p); break;
      case 'approveWithdraw': out = approveWithdraw(p); break;
      case 'rejectWithdraw':  out = rejectWithdraw(p); break;
      default: out = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    out = { success: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ================== HELPERS ==================
function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function sha256_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function nextId_(sheet, prefix, pad) {
  var n = sheet.getLastRow(); // dòng 1 là header
  if (prefix) return prefix + ('000' + n).slice(-(pad || 3));
  return 1000 + n; // đơn hàng: 1001, 1002...
}
function logTransaction_(type, ctvCode, orderId, amount, note) {
  try {
    var sh = sheet_('sheet_transactions');
    if (!sh) return;
    sh.appendRow(['GD' + ('00000' + sh.getLastRow()).slice(-5), type, ctvCode || '', orderId || '', Number(amount || 0), note || '', nowStr_()]);
  } catch (err) {}
}
function notify_(payload) {
  // Gửi thông báo qua Telegram Bot
  sendTelegram_(formatTelegramMsg_(payload));
  // Gửi JSON thô tới webhook tùy ý (n8n/…)
  if (!WEBHOOK_URL) return;
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });
  } catch (err) {}
}

// Gửi tin nhắn qua Telegram Bot API
function sendTelegram_(text) {
  if (!TELEGRAM_BOT_TOKEN || !text) return;
  var ids = (TELEGRAM_CHAT_IDS || []).filter(function (x) { return x && String(x).trim(); });
  ids.forEach(function (cid) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ chat_id: String(cid).trim(), text: text })
      });
    } catch (err) {}
  });
}

// ===== HÀM TEST TELEGRAM — CHẠY TRỰC TIẾP TRONG APPS SCRIPT =====
// Chọn hàm "testTelegram" ở thanh trên → bấm Run → xem Execution log (Ctrl+Enter).
// Nếu Telegram nhận được tin "✅ GG CLOUD test..." nghĩa là token + chat_id ĐÚNG.
// Nếu log in ra {"ok":false,...} thì đọc lỗi để biết sai gì.
function testTelegram() {
  var ids = (TELEGRAM_CHAT_IDS || []);
  Logger.log('TOKEN: ' + (TELEGRAM_BOT_TOKEN ? TELEGRAM_BOT_TOKEN.slice(0, 12) + '…' : '(TRỐNG)'));
  Logger.log('CHAT_IDS: ' + JSON.stringify(ids));
  ids.forEach(function (cid) {
    var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_BOT_TOKEN + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: String(cid).trim(), text: '✅ GG CLOUD test gửi tới ' + cid + ' lúc ' + nowStr_() })
    });
    Logger.log('→ chat_id ' + cid + ' | HTTP ' + res.getResponseCode() + ' | ' + res.getContentText());
  });
}

// Soạn nội dung Telegram theo sự kiện
function formatTelegramMsg_(p) {
  if (!p || !p.event) return '';
  var vnd = function (n) { return Number(n || 0).toLocaleString('vi-VN') + 'đ'; };
  if (p.event === 'new_order') {
    return '🛒 ĐƠN HÀNG MỚI #' + p.order_id + '\n📦 Gói: ' + p.package + '\n📧 Gmail: ' + (p.gmail || '(không có)') + '\n📞 SĐT: ' + p.phone + '\n💰 Số tiền: ' + vnd(p.amount) + (p.ctv_code ? '\n🎟 Mã CTV: ' + p.ctv_code : '') + '\n⏰ ' + p.created_at;
  }
  if (p.event === 'withdraw_request') {
    return '💸 LỆNH RÚT TIỀN MỚI #' + p.withdraw_id + '\n🎟 CTV: ' + p.ctv_code + '\n📧 Gmail: ' + (p.gmail || '') + '\n💰 Số tiền: ' + vnd(p.amount) + '\n🏦 Ngân hàng: ' + p.bank_name + '\n#️⃣ STK: ' + p.bank_account + '\n👤 Chủ TK: ' + p.account_holder + '\n⏰ ' + p.created_at + '\n\n➡️ Vào CRM → tab "Duyệt rút tiền" để xác nhận.';
  }
  if (p.event === 'withdraw_done') {
    return '✅ ĐÃ CHI TRẢ LỆNH RÚT #' + p.withdraw_id + '\n🎟 CTV: ' + p.ctv_code + '\n💰 Số tiền: ' + vnd(p.amount) + '\n⏰ ' + p.at;
  }
  if (p.event === 'withdraw_rejected') {
    return '⛔ TỪ CHỐI LỆNH RÚT #' + p.withdraw_id + '\n🎟 CTV: ' + p.ctv_code + '\n💰 Số tiền hoàn lại: ' + vnd(p.amount) + '\n⏰ ' + p.at;
  }
  return '🔔 ' + JSON.stringify(p);
}

// Soạn nội dung thông báo tiếng Việt theo từng sự kiện
function formatZaloMsg_(p) {
  if (!p || !p.event) return '';
  var vnd = function (n) { return Number(n || 0).toLocaleString('vi-VN') + 'đ'; };
  if (p.event === 'new_order') {
    return '🛒 ĐƠN HÀNG MỚI #' + p.order_id +
      '\n📦 Gói: ' + p.package + ' (nâng cấp trước — trả sau)' +
      '\n📧 Gmail: ' + (p.gmail || '(không có)') +
      '\n📞 SĐT/Zalo: ' + p.phone +
      '\n💰 Số tiền thanh toán: ' + vnd(p.amount) +
      (p.ctv_code ? '\n🎟 Mã CTV: ' + p.ctv_code : '') +
      '\n⏰ ' + p.created_at;
  }
  if (p.event === 'order_paid') {
    return '✅ ĐÃ THANH TOÁN ĐƠN #' + p.order_id +
      '\n📦 Gói: ' + p.package +
      '\n💰 Số tiền: ' + vnd(p.amount) +
      '\n🔍 Nguồn xác thực: ' + p.source +
      (p.ctv_code ? '\n🎟 CTV ' + p.ctv_code + ' nhận hoa hồng ' + vnd(p.commission_ctv) : '') +
      '\n⏰ ' + p.at;
  }
  if (p.event === 'underpaid') {
    return '⚠️ THIẾU TIỀN ĐƠN #' + p.order_id + '\nĐã nhận: ' + vnd(p.received) + ' / Cần: ' + vnd(p.due);
  }
  if (p.event === 'unmatched_payment') {
    return '❓ TIỀN VÀO KHÔNG KHỚP ĐƠN NÀO\n💰 ' + vnd(p.received) + '\n📝 Nội dung: ' + (p.content || '(trống)');
  }
  return '🔔 ' + JSON.stringify(p);
}

// ================== 1. VALIDATE & ÁP MÃ CTV ==================
// Hoa hồng 20% + giảm 5% cho khách: CHỈ áp dụng gói có ctv:true (12T/18T).
function validateAndApplyCTV(packageType, ctvCode, originalAmount) {
  var discountAmount = 0, amountToPay = originalAmount, commissionCTV = 0, isValidCTV = false;
  var pkg = PACKAGES[packageType];

  if (pkg && pkg.ctv && ctvCode && String(ctvCode).trim() !== '') {
    var data = sheet_('sheet_ctvs').getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).toUpperCase() === String(ctvCode).trim().toUpperCase() && Number(data[i][4]) === 1) {
        isValidCTV = true;
        ctvCode = String(data[i][2]).toUpperCase();
        break;
      }
    }
    if (isValidCTV) {
      discountAmount = Math.round(originalAmount * CTV_DISCOUNT_RATE);
      amountToPay = originalAmount - discountAmount;
      commissionCTV = Math.round(amountToPay * COMMISSION_RATE);
    } else {
      ctvCode = ''; // mã sai/khóa → bỏ qua
    }
  } else {
    ctvCode = '';
  }

  return {
    applied_ctv_code: ctvCode,
    discount_amount: discountAmount,
    amount: amountToPay,
    commission_ctv: commissionCTV,
    is_ctv_code_accepted: isValidCTV
  };
}

// API cho Frontend hiển thị giá giảm ngay khi khách nhập mã (trước khi tạo đơn)
function checkCTV(ctvCode, packageType) {
  var pkg = PACKAGES[packageType];
  if (!pkg) return { success: false, error: 'Gói không hợp lệ' };
  var r = validateAndApplyCTV(packageType, ctvCode, pkg.amount);
  var msg;
  if (r.is_ctv_code_accepted) {
    msg = 'Áp dụng mã CTV thành công! Bạn được giảm 5% — người giới thiệu nhận hoa hồng 20%.';
  } else {
    msg = pkg.ctv ? 'Mã CTV không tồn tại hoặc đã bị khóa.' : 'Mã CTV chỉ áp dụng cho gói 12 tháng &amp; 18 tháng.';
  }
  return {
    success: true,
    is_ctv_code_accepted: r.is_ctv_code_accepted,
    financials: { original_price: pkg.amount, discount_applied: r.discount_amount, final_amount_to_pay: r.amount },
    message: msg
  };
}

// ================== 2. TẠO ĐƠN HÀNG ==================
// Luồng đơn: pending_activation → (KTV) activated → (KTV up bill) paid_success
function createOrder(p) {
  var gmail = String(p.customer_gmail || '').trim().toLowerCase();
  var phone = String(p.customer_phone || '').trim();
  var packageType = String(p.package_type || '').trim();
  var pkg = PACKAGES[packageType];
  if (!pkg) return { success: false, error: 'Gói không hợp lệ (1THANG/12THANG/18THANG)' };
  if (!phone) return { success: false, error: 'Thiếu số điện thoại' };
  if (!gmail || !/^\S+@\S+\.\S+$/.test(gmail)) return { success: false, error: 'Gmail không hợp lệ' };

  var ctv = validateAndApplyCTV(packageType, p.applied_ctv_code, pkg.amount);
  var sh = sheet_('sheet_orders');
  var id = nextId_(sh);
  var commissionKTV = Math.round(ctv.amount * 0.05);
  var initialStatus = 'pending_activation';
  var finalAmount = ctv.amount;

  sh.appendRow([id, gmail, phone, packageType, pkg.amount, ctv.discount_amount, finalAmount,
    initialStatus, '', commissionKTV, ctv.applied_ctv_code, ctv.commission_ctv,
    pkg.cogs, nowStr_(), '', '', '']);

  notify_({
    event: 'new_order', order_id: id, gmail: gmail, phone: phone,
    package: packageType, amount: finalAmount, ctv_code: ctv.applied_ctv_code,
    flow: pkg.flow, created_at: nowStr_()
  });

  return {
    success: true, order_id: id, package_type: packageType, flow: 'manual',
    transfer_note: 'GGCLOUD' + id,
    financials: { original_price: pkg.amount, discount_applied: ctv.discount_amount, final_amount_to_pay: finalAmount },
    is_ctv_code_accepted: ctv.is_ctv_code_accepted,
    status: initialStatus,
    message: 'Tạo đơn thành công. KTV sẽ liên hệ qua Zalo để nâng cấp.'
  };
}

// ================== 3. TRẠNG THÁI + ĐẾM NGƯỢC FOMO ==================
function getOrderStatus(orderId) {
  expireOverdue_();
  var data = sheet_('sheet_orders').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      var status = data[i][7];
      var res = { success: true, order_id: data[i][0], status: status, amount: data[i][6], countdown_seconds: 0 };
      if (status === 'activated' && data[i][14]) {
        var activatedAt = new Date(data[i][14]);
        var elapsed = Math.floor((Date.now() - activatedAt.getTime()) / 1000);
        res.countdown_seconds = Math.max(0, FOMO_MINUTES * 60 - elapsed);
      }
      return res;
    }
  }
  return { success: false, error: 'Không tìm thấy đơn hàng #' + orderId };
}

// ================== 4. KTV KÍCH HOẠT ĐƠN (gói manual) ==================
function activateOrder(p) {
  var sh = sheet_('sheet_orders');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.order_id)) {
      if (data[i][7] === 'paid_success') return { success: false, error: 'Đơn đã hoàn tất trước đó' };
      var t = nowStr_();
      sh.getRange(i + 1, 8).setValue('activated');       // H: status
      sh.getRange(i + 1, 9).setValue(p.ktv_id || '');    // I: ktv_id
      sh.getRange(i + 1, 15).setValue(t);                // O: activated_at
      return { success: true, activated_at: t, message: 'Đã chuyển sang activated, kích hoạt QR và đếm ngược tại client.' };
    }
  }
  return { success: false, error: 'Không tìm thấy đơn hàng' };
}

// ================== 5. HOÀN TẤT ĐƠN (chung cho mọi luồng) ==================
// markPaid_: đổi trạng thái + giải ngân hoa hồng CTV + nhả CDK (nếu gói cdk). IDEMPOTENT.
function markPaid_(rowIndex, data, source, proofUrl) {
  var sh = sheet_('sheet_orders');
  var r = data[rowIndex];
  var id = r[0];
  if (r[7] === 'paid_success') {
    return { success: true, already: true, order_id: id, cdk_code: String(r[16] || '') || undefined };
  }
  sh.getRange(rowIndex + 1, 8).setValue('paid_success');
  if (proofUrl) sh.getRange(rowIndex + 1, 16).setValue(proofUrl);

  // 1) Giải ngân hoa hồng CTV (gói 12T/18T — commission_ctv đã gate lúc tạo đơn)
  var ctvCode = String(r[10] || '');
  var commission = Number(r[11] || 0);
  if (ctvCode && commission > 0) {
    var cs = sheet_('sheet_ctvs');
    var cd = cs.getDataRange().getValues();
    for (var j = 1; j < cd.length; j++) {
      if (String(cd[j][2]).toUpperCase() === ctvCode.toUpperCase()) {
        cs.getRange(j + 1, 4).setValue(Number(cd[j][3] || 0) + commission);
        logTransaction_('commission', ctvCode, id, commission, 'Hoàn tất - Nhận hoa hồng đơn hàng #' + id);
        break;
      }
    }
  }

  notify_({
    event: 'order_paid', order_id: id, source: source, package: String(r[3]),
    amount: Number(r[6] || 0), ctv_code: ctvCode, commission_ctv: commission, at: nowStr_()
  });

  var out = { success: true, order_id: id, status: 'paid_success' };
  if (proofUrl) out.payment_proof_url = proofUrl;
  return out;
}

// completeOrder: khách bấm "đã thanh toán" / tải bill lên (phương án dự phòng duyệt tay)
function completeOrder(p) {
  var sh = sheet_('sheet_orders');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.order_id)) {
      // Lưu ảnh hóa đơn lên Drive
      var proofUrl = '';
      if (p.payment_proof_base64) {
        try {
          var m = String(p.payment_proof_base64).match(/^data:(image\/\w+);base64,(.+)$/);
          var mime = m ? m[1] : 'image/png';
          var b64 = m ? m[2] : p.payment_proof_base64;
          var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, 'proof_' + p.order_id + '.png');
          var folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
          var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          proofUrl = file.getUrl();
        } catch (err) { proofUrl = 'UPLOAD_ERROR: ' + err.message; }
      }
      return markPaid_(i, data, p.payment_proof_base64 ? 'bill_upload' : 'manual_confirm', proofUrl);
    }
  }
  return { success: false, error: 'Không tìm thấy đơn hàng' };
}



// ================== 6. ĐĂNG KÝ CTV (CHỐNG SPAM) ==================
function registerCTV(p) {
  var gmail = String(p.gmail || '').trim().toLowerCase();
  if (!gmail || !/^\S+@\S+\.\S+$/.test(gmail)) return { success: false, error: 'Gmail không hợp lệ' };

  // Điều kiện: đã từng mua hàng thành công
  var orders = sheet_('sheet_orders').getDataRange().getValues();
  var bought = false;
  for (var i = 1; i < orders.length; i++) {
    if (String(orders[i][1]).toLowerCase() === gmail && orders[i][7] === 'paid_success') { bought = true; break; }
  }
  if (!bought) return { success: false, error: 'Bạn chưa từng mua hàng tại shop, không đủ điều kiện làm CTV' };

  var cs = sheet_('sheet_ctvs');
  var data = cs.getDataRange().getValues();
  // Nếu gmail đã là CTV → trả lại mã cũ
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][1]).toLowerCase() === gmail) {
      return { success: true, ctv_code: data[j][2], existed: true, message: 'Gmail này đã là CTV, đây là mã của bạn.' };
    }
  }
  // Tự sinh mã duy nhất: 3 ký tự đầu gmail + 4 ký tự ngẫu nhiên
  var existing = {};
  for (var k = 1; k < data.length; k++) existing[String(data[k][2]).toUpperCase()] = true;
  var code;
  do {
    code = 'CTV-' + gmail.replace(/[^a-z0-9]/g, '').slice(0, 3).toUpperCase() +
      Math.random().toString(36).slice(2, 6).toUpperCase();
  } while (existing[code]);

  cs.appendRow(['CTV' + ('000' + cs.getLastRow()).slice(-3), gmail, code, 0, 1]);
  return { success: true, ctv_code: code, existed: false, message: 'Đăng ký CTV thành công!' };
}

// ================== 6B. ĐĂNG NHẬP VÍ CTV (Frontend) ==================
// Frontend gửi { gmail, password }. Mật khẩu CTV cố định = 'cvt123'.
// Trả về số dư khả dụng / đóng băng / tổng thu nhập + số khách đã giới thiệu.
function loginCTV(p) {
  var gmail = String(p.gmail || '').trim().toLowerCase();
  var pass = String(p.password || '');
  if (!gmail || !/^\S+@\S+\.\S+$/.test(gmail)) return { success: false, error: 'Gmail không hợp lệ' };
  if (pass !== 'cvt123') return { success: false, error: 'Mật khẩu CTV không đúng.' };
  var ctv = findCTV_(null, gmail);
  if (!ctv) return { success: false, error: 'Gmail này chưa đăng ký CTV. Vui lòng đăng ký CTV trước.' };
  if (Number(ctv.status) !== 1) return { success: false, error: 'Tài khoản CTV đang bị khóa.' };
  var orders = sheet_('sheet_orders').getDataRange().getValues();
  var refCount = 0;
  for (var i = 1; i < orders.length; i++) {
    if (String(orders[i][10]).toUpperCase() === ctv.ctv_code) refCount++;
  }
  var w = walletSummary_(ctv.ctv_code, ctv.balance_accumulated);
  return {
    success: true, name: String(ctv.gmail).split('@')[0], ctv_code: ctv.ctv_code,
    balance: w.available, frozen: w.frozen, total_income: w.total, ref_count: refCount
  };
}

// ================== 7. CRM: LOGIN ==================
function login(p) {
  var data = sheet_('sheet_users').getDataRange().getValues();
  var hash = sha256_(String(p.password || ''));
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(p.username) && String(data[i][2]) === hash) {
      if (data[i][5] !== 'Hoạt động') return { success: false, error: 'Tài khoản đã bị khóa' };
      return { success: true, user: { id: data[i][0], username: data[i][1], full_name: data[i][3], role: data[i][4] } };
    }
  }
  return { success: false, error: 'Sai tên đăng nhập hoặc mật khẩu' };
}

// ================== HẾT HẠN 30 PHÚT → "QUÁ HẠN BÙNG TIỀN" ==================
function expireOverdue_() {
  var sh = sheet_('sheet_orders');
  var data = sh.getDataRange().getValues();
  var now = Date.now();
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'activated' && data[i][14]) {
      var t = new Date(data[i][14]).getTime();
      if (t && now - t > FOMO_MINUTES * 60000) sh.getRange(i + 1, 8).setValue('expired_unpaid');
    }
  }
}

// ================== CỔNG CTV: VÍ + KHÁCH ĐÃ GIỚI THIỆU ==================
function ctvPortal(p) {
  var gmail = String(p.gmail || '').trim().toLowerCase();
  var code = String(p.ctv_code || '').trim().toUpperCase();
  var data = sheet_('sheet_ctvs').getDataRange().getValues();
  var ctv = null;
  for (var i = 1; i < data.length; i++) {
    if ((gmail && String(data[i][1]).toLowerCase() === gmail) || (code && String(data[i][2]).toUpperCase() === code)) {
      ctv = { id: data[i][0], gmail: data[i][1], ctv_code: data[i][2], balance_accumulated: Number(data[i][3] || 0), status: data[i][4] };
      break;
    }
  }
  if (!ctv) return { success: false, error: 'Không tìm thấy CTV với gmail/mã này' };
  expireOverdue_();
  var orders = sheet_('sheet_orders').getDataRange().getValues();
  var rows = [];
  for (var j = orders.length - 1; j >= 1; j--) {
    var r = orders[j];
    if (String(r[10]).toUpperCase() !== String(ctv.ctv_code).toUpperCase()) continue;
    rows.push({ id: r[0], customer_gmail: r[1], package_type: r[3], amount: r[6], status: r[7], commission_ctv: r[11], created_at: String(r[13]) });
  }
  var wl = getWithdrawals({ ctv_code: ctv.ctv_code });
  return { success: true, ctv: ctv, orders: rows, wallet: walletSummary_(ctv.ctv_code, ctv.balance_accumulated), withdrawals: wl.withdrawals };
}

// ================== VÍ HOA HỒNG: TỔNG / ĐÓNG BĂNG / KHẢ DỤNG ==================
// total    = tổng hoa hồng tích lũy (balance_accumulated trong sheet_ctvs)
// frozen   = tổng các lệnh rút đang 'pending' (chờ Admin duyệt)
// paid_out = tổng các lệnh rút đã 'success'
// available = total - frozen - paid_out
function walletSummary_(ctvCode, totalAccumulated) {
  var total = Number(totalAccumulated || 0);
  var frozen = 0, paidOut = 0;
  var sh = sheet_('sheet_withdrawals');
  if (sh) {
    var d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]).toUpperCase() !== String(ctvCode).toUpperCase()) continue;
      var amt = Number(d[i][3] || 0);
      var st = String(d[i][7] || '');
      if (st === 'pending') frozen += amt;
      else if (st === 'success') paidOut += amt;
    }
  }
  return { total: total, frozen: frozen, paid_out: paidOut, available: Math.max(0, total - frozen - paidOut) };
}

// Lấy CTV theo mã hoặc gmail
function findCTV_(ctvCode, gmail) {
  var data = sheet_('sheet_ctvs').getDataRange().getValues();
  var code = String(ctvCode || '').trim().toUpperCase();
  var mail = String(gmail || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if ((code && String(data[i][2]).toUpperCase() === code) || (mail && String(data[i][1]).toLowerCase() === mail)) {
      return { row: i + 1, id: data[i][0], gmail: data[i][1], ctv_code: String(data[i][2]).toUpperCase(), balance_accumulated: Number(data[i][3] || 0), status: Number(data[i][4]) };
    }
  }
  return null;
}

// ================== RÚT TIỀN: CTV GỬI LỆNH ==================
function requestWithdraw(p) {
  var ctv = findCTV_(p.ctv_code, p.gmail);
  if (!ctv) return { success: false, error: 'Không tìm thấy CTV.' };
  if (Number(ctv.status) !== 1) return { success: false, error: 'Tài khoản CTV đang bị khóa.' };

  var amount = Math.round(Number(p.amount || 0));
  var bankName = String(p.bank_name || '').trim();
  var bankAcc = String(p.bank_account || '').trim();
  var holder = String(p.account_holder || '').trim();
  if (!amount || amount <= 0) return { success: false, error: 'Số tiền rút không hợp lệ.' };
  if (amount < MIN_WITHDRAW_AMOUNT) return { success: false, error: 'Số tiền rút tối thiểu là ' + MIN_WITHDRAW_AMOUNT.toLocaleString('vi-VN') + 'đ.' };
  if (!bankName || !bankAcc || !holder) return { success: false, error: 'Vui lòng nhập đủ ngân hàng, số tài khoản và tên chủ tài khoản.' };

  // Kiểm tra giới hạn số lệnh trong ngày
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var sh = sheet_('sheet_withdrawals');
  var d = sh.getDataRange().getValues();
  var todayCount = 0;
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][1]).toUpperCase() !== ctv.ctv_code) continue;
    if (String(d[i][8] || '').indexOf(today) === 0) todayCount++;
  }
  if (todayCount >= MAX_WITHDRAWALS_PER_DAY) {
    return { success: false, error: 'Bạn đã tạo đủ ' + MAX_WITHDRAWALS_PER_DAY + ' lệnh rút trong hôm nay. Vui lòng thử lại vào ngày mai.' };
  }

  // Kiểm tra số dư khả dụng
  var sum = walletSummary_(ctv.ctv_code, ctv.balance_accumulated);
  if (amount > sum.available) {
    return { success: false, error: 'Số tiền rút vượt quá số dư khả dụng (' + sum.available.toLocaleString('vi-VN') + 'đ).' };
  }

  var id = 'RUT' + ('00000' + sh.getLastRow()).slice(-5);
  var now = nowStr_();
  sh.appendRow([id, ctv.ctv_code, ctv.gmail, amount, bankName, bankAcc, holder, 'pending', now, '']);
  logTransaction_('withdraw_request', ctv.ctv_code, id, -amount, 'Yêu cầu rút tiền');

  notify_({
    event: 'withdraw_request', withdraw_id: id, ctv_code: ctv.ctv_code, gmail: ctv.gmail,
    amount: amount, bank_name: bankName, bank_account: bankAcc, account_holder: holder, created_at: now
  });

  var newSum = walletSummary_(ctv.ctv_code, ctv.balance_accumulated);
  return { success: true, withdraw_id: id, wallet: newSum, message: 'Lệnh rút #' + id + ' đã gửi. Chờ Admin duyệt (5 phút – tối đa 4 giờ).' };
}

// ================== RÚT TIỀN: DANH SÁCH ==================
// p.ctv_code: lọc theo CTV; không có → trả toàn bộ (Admin)
function getWithdrawals(p) {
  var sh = sheet_('sheet_withdrawals');
  if (!sh) return { success: true, withdrawals: [] };
  var d = sh.getDataRange().getValues();
  var code = String((p && p.ctv_code) || '').trim().toUpperCase();
  var rows = [];
  for (var i = d.length - 1; i >= 1; i--) {
    if (code && String(d[i][1]).toUpperCase() !== code) continue;
    rows.push({
      id: d[i][0], ctv_code: d[i][1], gmail: d[i][2], amount: Number(d[i][3] || 0),
      bank_name: d[i][4], bank_account: String(d[i][5]), account_holder: d[i][6],
      status: String(d[i][7] || ''), created_at: String(d[i][8] || ''), processed_at: String(d[i][9] || '')
    });
    if (rows.length >= 300) break;
  }
  return { success: true, withdrawals: rows };
}

// ================== RÚT TIỀN: ADMIN DUYỆT / TỪ CHỐI ==================
function findWithdrawRow_(id) {
  var sh = sheet_('sheet_withdrawals');
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]) === String(id)) return { sh: sh, row: i + 1, rec: d[i] };
  }
  return null;
}

function approveWithdraw(p) {
  var f = findWithdrawRow_(p.withdraw_id);
  if (!f) return { success: false, error: 'Không tìm thấy lệnh rút.' };
  if (String(f.rec[7]) !== 'pending') return { success: false, error: 'Lệnh này đã được xử lý.' };
  var now = nowStr_();
  f.sh.getRange(f.row, 8).setValue('success');   // status
  f.sh.getRange(f.row, 10).setValue(now);         // processed_at
  logTransaction_('withdraw_paid', String(f.rec[1]), String(f.rec[0]), -Number(f.rec[3] || 0), 'Admin duyệt chi trả');
  notify_({ event: 'withdraw_done', withdraw_id: f.rec[0], ctv_code: f.rec[1], amount: Number(f.rec[3] || 0), at: now });
  return { success: true, message: 'Đã duyệt chi trả lệnh #' + f.rec[0] + '.' };
}

function rejectWithdraw(p) {
  var f = findWithdrawRow_(p.withdraw_id);
  if (!f) return { success: false, error: 'Không tìm thấy lệnh rút.' };
  if (String(f.rec[7]) !== 'pending') return { success: false, error: 'Lệnh này đã được xử lý.' };
  var now = nowStr_();
  f.sh.getRange(f.row, 8).setValue('rejected');   // hoàn lại số dư (không còn đóng băng)
  f.sh.getRange(f.row, 10).setValue(now);
  logTransaction_('withdraw_rejected', String(f.rec[1]), String(f.rec[0]), Number(f.rec[3] || 0), String(p.reason || 'Admin từ chối'));
  notify_({ event: 'withdraw_rejected', withdraw_id: f.rec[0], ctv_code: f.rec[1], amount: Number(f.rec[3] || 0), at: now });
  return { success: true, message: 'Đã từ chối lệnh #' + f.rec[0] + ' — số dư hoàn lại cho CTV.' };
}

// ================== 8. CRM: DANH SÁCH ĐƠN ==================
function getOrders(p) {
  expireOverdue_();
  var data = sheet_('sheet_orders').getDataRange().getValues();
  var rows = [];
  for (var i = data.length - 1; i >= 1; i--) { // mới nhất trước
    var r = data[i];
    if (p && p.status && r[7] !== p.status) continue;
    rows.push({
      id: r[0], customer_gmail: r[1], customer_phone: r[2], package_type: r[3],
      original_amount: r[4], discount_amount: r[5], amount: r[6], status: r[7],
      ktv_id: r[8], commission_ktv: r[9], applied_ctv_code: r[10], commission_ctv: r[11],
      cogs_amount: r[12], created_at: String(r[13]), activated_at: String(r[14]), payment_proof_url: r[15],
      cdk_code: String(r[16] || '')
    });
    if (rows.length >= 200) break;
  }
  return { success: true, orders: rows };
}

function getCTVs() {
  var data = sheet_('sheet_ctvs').getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    rows.push({ id: data[i][0], gmail: data[i][1], ctv_code: data[i][2], balance_accumulated: data[i][3], status: data[i][4] });
  }
  return { success: true, ctvs: rows };
}





// ================== 10. CRM: DASHBOARD TÀI CHÍNH ==================
// p.period: 'day' | 'month' | 'year' | 'all' (mặc định 'all')
// ĐÃ BỮ: mục tiêu doanh thu & chi phí vận hành.
function getDashboard(p) {
  expireOverdue_();
  var orders = sheet_('sheet_orders').getDataRange().getValues();
  var now = new Date();
  var dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  var weekStart = dayStart - ((now.getDay() + 6) % 7) * 86400000; // Thứ 2 đầu tuần
  var monthStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  var yearStart = new Date(now.getFullYear(), 0, 1).getTime();
  var period = (p && p.period) || 'all';
  var periodStart = period === 'day' ? dayStart : period === 'month' ? monthStart : period === 'year' ? yearStart : 0;

  var rev = { day: 0, week: 0, month: 0, year: 0, total: 0, period: 0 };
  var cogs = 0, comKtv = 0, comCtv = 0;
  var counts = { pending_activation: 0, activated: 0, paid_success: 0, expired_unpaid: 0 };
  var ctvSales = {}; // ctv_code -> {revenue, orders} trong chu kỳ lọc

  for (var i = 1; i < orders.length; i++) {
    var r = orders[i];
    if (counts[r[7]] !== undefined) counts[r[7]]++;
    if (r[7] !== 'paid_success') continue;
    var amt = Number(r[6] || 0);
    var t = new Date(r[13]).getTime();
    rev.total += amt;
    if (t >= dayStart) rev.day += amt;
    if (t >= weekStart) rev.week += amt;
    if (t >= monthStart) rev.month += amt;
    if (t >= yearStart) rev.year += amt;
    if (t >= periodStart) rev.period += amt;
    var code = String(r[10] || '').toUpperCase();
    if (code && t >= periodStart) {
      if (!ctvSales[code]) ctvSales[code] = { ctv_code: code, revenue: 0, orders: 0 };
      ctvSales[code].revenue += amt;
      ctvSales[code].orders += 1;
    }
    cogs += Number(r[12] || 0);
    comKtv += Number(r[9] || 0);
    comCtv += Number(r[11] || 0);
  }

  // Top 3 CTV theo doanh số trong chu kỳ lọc (kèm gmail)
  var ctvData = sheet_('sheet_ctvs').getDataRange().getValues();
  var gmailByCode = {};
  for (var k = 1; k < ctvData.length; k++) gmailByCode[String(ctvData[k][2]).toUpperCase()] = ctvData[k][1];
  var top = Object.keys(ctvSales).map(function (c) {
    var s = ctvSales[c];
    s.gmail = gmailByCode[c] || '';
    return s;
  }).sort(function (a, b) { return b.revenue - a.revenue; }).slice(0, 3);

  return {
    success: true,
    period: period,
    revenue: rev,
    order_counts: counts,
    expenses: { cogs: cogs, commission_ktv: comKtv, commission_ctv: comCtv },
    net_profit: rev.total - (cogs + comKtv + comCtv),
    top_ctvs: top
  };
}
