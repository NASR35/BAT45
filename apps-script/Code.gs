// ==========================================
// نظام إدارة الحي / العمارة (v2 - Backend API)
// ==========================================

function doGet(e) {
  var action = e ? e.parameter.action : '';
  if (action === 'public') {
    return handlePublicData();
  }
  return responseJSON({ status: 'error', message: 'Action not allowed' }, 400);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    var action = data.action;

    if (action === 'login') {
      return handleLogin(data.payload);
    }

    // التحقق من التوكن للعمليات الإدارية
    if (!verifyToken(data.token)) {
      return responseJSON({ status: 'error', message: 'Unauthorized / Expired Session' }, 401);
    }

    switch (action) {
      case 'dash': return handleDash();
      case 'addHouse': return handleAddHouse(data.payload);
      case 'pay': return handlePay(data.payload);
      case 'cancelPayment': return handleCancelPayment(data.payload);
      case 'addExpense': return handleAddExpense(data.payload);
      default: return responseJSON({ status: 'error', message: 'Unknown action' }, 400);
    }
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() }, 500);
  }
}

// ------------------------------------------
// دالة الإعداد المبدئي (تشغل مرة واحدة فقط)
// ------------------------------------------
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheets = ['بيوت', 'أنواع', 'مدفوعات', 'مصروفات', 'إعدادات', 'مستخدمون', 'سجل'];
  sheets.forEach(function(s) {
    if (!ss.getSheetByName(s)) ss.insertSheet(s);
  });

  // إعداد أنواع الاشتراكات الإفتراضية
  var sheetTypes = ss.getSheetByName('أنواع');
  if (sheetTypes.getLastRow() === 0) {
    sheetTypes.appendRow(['id', 'النوع', 'الاسم', 'المبلغ', 'شهر_البداية', 'نشط', 'رصيد_افتتاحي', 'وصف']);
    sheetTypes.appendRow(['T1', 'شهري', 'الاشتراك الشهري العام', 1000, '2026-01', true, 0, 'الاشتراك الدوري للخدمات الأساسية']);
  }

  // إعداد جدول الإعدادات
  var sheetSettings = ss.getSheetByName('إعدادات');
  if (sheetSettings.getLastRow() === 0) {
    sheetSettings.appendRow(['المفتاح', 'القيمة']);
    sheetSettings.appendRow(['monthly_target', 50000]);
  }

  // إعداد جدول المستخدمين المبدئي (admin / admin123)
  var sheetUsers = ss.getSheetByName('مستخدمون');
  if (sheetUsers.getLastRow() === 0) {
    sheetUsers.appendRow(['username', 'salt', 'hash']);
    // كلمة السر المبدئية: admin123
    var salt = 'rand_salt_123';
    var hash = computeHash('admin123', salt);
    sheetUsers.appendRow(['admin', salt, hash]);
  }

  logEvent('INFO', 'system', 'تم تنفيذ setup() بنجاح وتجهيز التبويبات');
  return 'تم إعداد التبويبات بنجاح!';
}

// ------------------------------------------
// الدوال المساعدة للبيانات والمعالجة
// ------------------------------------------
function handlePublicData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHouses = ss.getSheetByName('بيوت');
  var sheetPayments = ss.getSheetByName('مدفوعات');
  var sheetExpenses = ss.getSheetByName('مصروفات');
  var sheetTypes = ss.getSheetByName('أنواع');

  var housesCount = Math.max(0, sheetHouses.getLastRow() - 1);
  
  // إجمالي المقبوضات غير الملغاة
  var totalPaid = 0;
  if (sheetPayments.getLastRow() > 1) {
    var pData = sheetPayments.getRange(2, 1, sheetPayments.getLastRow() - 1, 7).getValues();
    pData.forEach(function(r) {
      if (!r[6]) totalPaid += Number(r[5] || 0); // r[6] = ملغاة
    });
  }

  // إجمالي المصروفات
  var totalExpenses = 0;
  if (sheetExpenses.getLastRow() > 1) {
    var eData = sheetExpenses.getRange(2, 1, sheetExpenses.getLastRow() - 1, 3).getValues();
    eData.forEach(function(r) {
      totalExpenses += Number(r[2] || 0);
    });
  }

  var balance = totalPaid - totalExpenses;

  var output = {
    schema_version: 2,
    generated_at: new Date().toISOString(),
    summary: {
      total_houses: housesCount,
      total_income: totalPaid,
      total_expenses: totalExpenses,
      current_balance: balance
    }
  };

  return responseJSON(output, 200);
}

function handleLogin(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsers = ss.getSheetByName('مستخدمون');
  var uData = sheetUsers.getRange(2, 1, Math.max(1, sheetUsers.getLastRow() - 1), 3).getValues();

  for (var i = 0; i < uData.length; i++) {
    if (uData[i][0] === payload.username) {
      var calcHash = computeHash(payload.password, uData[i][1]);
      if (calcHash === uData[i][2]) {
        var token = Utilities.getUuid();
        CacheService.getScriptCache().put('token_' + token, payload.username, 21600); // 6 ساعات
        logEvent('INFO', payload.username, 'تسجيل دخول ناجح');
        return responseJSON({ status: 'success', token: token });
      }
    }
  }
  logEvent('WARN', payload.username || 'unknown', 'محاولة دخول فاشلة');
  return responseJSON({ status: 'error', message: 'بيانات الدخول غير صحيحة' }, 401);
}

function verifyToken(token) {
  if (!token) return false;
  var user = CacheService.getScriptCache().get('token_' + token);
  return !!user;
}

function computeHash(pass, salt) {
  var raw = pass + salt;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return digest.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function logEvent(level, user, msg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName('سجل');
  if (s) s.appendRow([new Date(), level, user, msg]);
}

function responseJSON(data, code) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}