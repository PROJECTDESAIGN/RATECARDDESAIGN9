/**
 * ============================================================================
 * GOOGLE APPS SCRIPT BACKEND — CREATOR PORTFOLIO & RATE CARD
 * Database Utama: Google Spreadsheet (Auto-Provisioning 10+ Sheets on Deploy)
 * Theme: Deep Plum (#481B3A)
 * Template: index9.html
 * ============================================================================
 * 
 * STRUKTUR DATABASE GOOGLE SPREADSHEET (OTOMATIS DIBUAT SAAT DEPLOY):
 * 1. Portfolio_Data : Master JSON & Kunci Konfigurasi Global
 * 2. Profile        : Informasi profil creator, kontak utama, WhatsApp, status update, dan PIN Admin
 * 3. ContactInfo    : Daftar channel sosial media & tautan kontak
 * 4. SocialStats    : Statistik performa akun (Followers, Reach, ER, Audience)
 * 5. AudienceInfo   : Niche, gender, demografi, dan geografi audiens
 * 6. BrandCollabs   : Daftar brand partner, logo, dan fallback teks
 * 7. Videos         : Daftar video portfolio, embed player URL, durasi, thumbnail, dan platform
 * 8. RateCards      : Paket rate card, harga, badge terpopuler, dan rincian fitur/layanan
 * 9. CollabImages   : Galeri foto kolaborasi dan caption
 * 10. Invoices      : Data invoice tersimpan, nomor invoice, tanggal, data pelanggan, dan tanda tangan digital
 * 11. Inquiries     : Pencatatan otomatis form inquiry/booking yang masuk dari website
 * ============================================================================
 */

// KONFIGURASI SPREADSHEET & GOOGLE DRIVE
// Jika script ini bound ke Spreadsheet (Extensions > Apps Script), biarkan SPREADSHEET_ID kosong.
// Jika script standalone, masukkan ID Google Spreadsheet Anda di bawah ini:
var SPREADSHEET_ID = '';

// Folder ID Google Drive untuk upload video & gambar (opsional).
// Jika kosong, file disimpan di root My Drive Anda.
var VIDEO_FOLDER_ID = '';

// Kunci Cache & Session Admin
var PORTFOLIO_CACHE_KEY = 'PORTFOLIO_CACHE_DATA_V5';
var PORTFOLIO_DATA_KEY = 'PORTFOLIO_DATA';
var PORTFOLIO_CHUNK_COUNT_KEY = 'PORTFOLIO_DATA_CHUNK_COUNT';
var PORTFOLIO_CHUNK_PREFIX = 'PORTFOLIO_DATA_CHUNK_';
var ADMIN_SESSION_PREFIX = 'portfolio_admin_session_';
var ADMIN_ATTEMPT_KEY = 'portfolio_admin_attempts';
var ADMIN_SESSION_TTL_SECONDS = 21600; // 6 jam

// Batasan Upload
var MAX_ASSET_BYTES = 25 * 1024 * 1024;   // 25 MB
var MAX_VIDEO_DURATION_SECONDS = 300;     // 5 menit
var ALLOWED_VIDEO_MIME = ['video/mp4', 'video/webm', 'video/quicktime', 'video/x-m4v', 'video/ogg'];
var ALLOWED_IMAGE_MIME = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/gif'];
var THEME_ACCENT_COLOR = '#481B3A';

/* ================================================================
   1. ROUTING: doGet & doPost (DUAL MODE: Web App & Universal API)
================================================================ */

function apiResponse_(data, callback) {
  var json = JSON.stringify(data);
  if (callback && String(callback).trim()) {
    var cb = String(callback).replace(/[^\w$.]/g, '');
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Mode API: dipanggil via GET dengan query parameter (?action=getData dll)
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var cb = e.parameter.callback || e.parameter.prefix || '';

    if (action === 'getData' || action === 'getPortfolioData') {
      return apiResponse_({ success: true, data: getData() }, cb);
    }
    if (action === 'verifyPin' || action === 'verifyAdminPin') {
      return apiResponse_(verifyAdminPin(e.parameter.pin), cb);
    }
    if (action === 'getInvoices') {
      return apiResponse_(getInvoices(e.parameter.token), cb);
    }
    if (action === 'detectVideo') {
      return apiResponse_(detectVideoUrl(e.parameter.url), cb);
    }
    if (action === 'resetData') {
      return apiResponse_(resetData(), cb);
    }
    if (action === 'initDatabase' || action === 'syncDatabase') {
      return apiResponse_(initSpreadsheetDatabase(), cb);
    }
  }

  // Saat Web App dibuka pertama kali saat di-deploy, pastikan struktur database otomatis ter-create
  try {
    getData();
  } catch (errInit) {
    Logger.log('Auto init saat doGet: ' + errInit);
  }

  // Mode Web App: render file HTML
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('Creator Portfolio & Rate Card – Deep Plum')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (pe) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    var action = payload.action || (e && e.parameter && e.parameter.action) || '';
    var args = payload.args || [];
    var cb = payload.callback || (e && e.parameter && e.parameter.callback) || '';

    if (action === 'getData' || action === 'getPortfolioData') {
      return apiResponse_({ success: true, data: getData() }, cb);
    }
    if (action === 'verifyAdminPin' || action === 'verifyPin') {
      var pin = payload.pin || args[0] || '';
      return apiResponse_(verifyAdminPin(pin), cb);
    }
    if (action === 'saveData' || action === 'savePortfolioData') {
      var data = payload.data || args[0];
      var token = payload.token || args[1] || '';
      return apiResponse_(saveData(data, token), cb);
    }
    if (action === 'uploadAsset') {
      var dataUri = payload.dataUri || args[0];
      var fileName = payload.fileName || args[1];
      var mimeType = payload.mimeType || args[2];
      var authToken = payload.token || args[3];
      var durationSec = payload.durationSeconds || args[4];
      return apiResponse_(uploadAsset(dataUri, fileName, mimeType, authToken, durationSec), cb);
    }
    if (action === 'submitInquiry') {
      var formData = payload.formData || payload.data || args[0] || payload;
      return apiResponse_(submitInquiry(formData), cb);
    }
    if (action === 'getInvoices') {
      var invTok = payload.token || args[0] || '';
      return apiResponse_(getInvoices(invTok), cb);
    }
    if (action === 'saveInvoice') {
      var invData = payload.invoiceData || payload.data || args[0];
      var saveTok = payload.token || args[1] || '';
      return apiResponse_(saveInvoice(invData, saveTok), cb);
    }
    if (action === 'deleteInvoice') {
      var invId = payload.id || args[0] || '';
      var delTok = payload.token || args[1] || '';
      return apiResponse_(deleteInvoice(invId, delTok), cb);
    }
    if (action === 'generateInvoicePdf') {
      var pdfId = payload.id || args[0] || '';
      var pdfTok = payload.token || args[1] || '';
      var pdfSig = payload.signatureDataUrl || args[2] || '';
      return apiResponse_(generateInvoicePdf(pdfId, pdfTok, pdfSig), cb);
    }
    if (action === 'detectVideo') {
      var vidUrl = payload.url || args[0] || '';
      return apiResponse_(detectVideoUrl(vidUrl), cb);
    }
    if (action === 'resetData') {
      return apiResponse_(resetData(), cb);
    }
    if (action === 'initDatabase' || action === 'syncDatabase') {
      return apiResponse_(initSpreadsheetDatabase(), cb);
    }

    return apiResponse_({ success: false, message: 'Action tidak dikenal: ' + action }, cb);
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return apiResponse_({ success: false, message: 'Server error: ' + err.toString() }, '');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Creator Portfolio')
      .addItem('Inisialisasi / Sinkronkan Sheet Database', 'initSpreadsheetDatabase')
      .addItem('Reset data portfolio ke default', 'resetData')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen dilewati: ' + err);
  }
}

function onInstall(e) {
  onOpen(e);
}

/* ================================================================
   2. GOOGLE SPREADSHEET DATABASE ENGINE (AUTO PROVISIONING)
================================================================ */

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var configuredId = SPREADSHEET_ID || props.getProperty('SPREADSHEET_ID') || props.getProperty('INQUIRIES_SPREADSHEET_ID');
  
  if (configuredId && String(configuredId).trim() && String(configuredId).trim() !== 'YOUR_SPREADSHEET_ID_HERE') {
    try {
      return SpreadsheetApp.openById(String(configuredId).trim());
    } catch (e) {
      Logger.log('Gagal membuka spreadsheet dengan ID ' + configuredId + ': ' + e);
    }
  }

  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {
    Logger.log('SpreadsheetApp.getActiveSpreadsheet() tidak tersedia: ' + e);
  }

  return null;
}

function getOrCreateSheet_(ss, sheetName, headers, headerColor) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground(headerColor || THEME_ACCENT_COLOR || '#0090A8')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Inisialisasi Database: Memastikan seluruh 10+ sheet otomatis terbuat saat pertama deploy
 */
function initSpreadsheetDatabase() {
  var ss = getSpreadsheet_();
  if (!ss) {
    return { success: false, message: 'Spreadsheet tidak terhubung. Masukkan SPREADSHEET_ID atau pasang script pada Google Spreadsheet.' };
  }
  var data = getStoredData_() || getDefaultData_();
  writePortfolioToSpreadsheet_(ss, data);
  return { success: true, message: '10+ Sheet Database Google Spreadsheet berhasil disinkronkan & dibuat otomatis.' };
}

/**
 * Memeriksa dan membuat seluruh 10+ sheet jika belum ada di Spreadsheet
 */
function ensureAllSheetsExist_(ss, data) {
  if (!ss) return;
  var color = THEME_ACCENT_COLOR || '#0090A8';
  var def = data || getDefaultData_();
  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd HH:mm:ss');

  // 1. Portfolio_Data
  var master = getOrCreateSheet_(ss, 'Portfolio_Data', ['Key', 'Value / Detail', 'Last Updated'], color);
  if (master.getLastRow() < 2) {
    master.getRange(2, 1, 9, 3).setValues([
      ['MASTER_JSON', JSON.stringify(def), nowStr],
      ['ADMIN_PIN', def.adminPin || '1234', nowStr],
      ['WHATSAPP_NUMBER', def.whatsappNumber || '', nowStr],
      ['PROFILE_NAME', (def.profile && def.profile.name) || '', nowStr],
      ['PROFILE_TAGLINE', (def.profile && def.profile.tagline) || '', nowStr],
      ['PROFILE_BIO', (def.profile && def.profile.bio) || '', nowStr],
      ['PROFILE_LOCATION', (def.profile && def.profile.location) || '', nowStr],
      ['PROFILE_EMAIL', (def.profile && def.profile.email) || '', nowStr],
      ['PROFILE_AVATAR_URL', (def.profile && def.profile.avatarUrl) || '', nowStr]
    ]);
  }

  // 2. Profile
  var pSheet = getOrCreateSheet_(ss, 'Profile', ['Field', 'Value', 'Keterangan / Status', 'Last Updated'], color);
  if (pSheet.getLastRow() < 2) {
    pSheet.getRange(2, 1, 9, 4).setValues([
      ['Name', (def.profile && def.profile.name) || '', 'Nama Lengkap Creator', nowStr],
      ['Tagline', (def.profile && def.profile.tagline) || '', 'Tagline / Deskripsi Singkat', nowStr],
      ['Bio', (def.profile && def.profile.bio) || '', 'Biografi Lengkap', nowStr],
      ['Avatar URL', (def.profile && def.profile.avatarUrl) || '', 'Link Foto Profil', nowStr],
      ['Location', (def.profile && def.profile.location) || '', 'Domisili / Lokasi', nowStr],
      ['Email', (def.profile && def.profile.email) || '', 'Email Kontak Utama', nowStr],
      ['WhatsApp Number', def.whatsappNumber || '', 'Nomor WhatsApp Utama', nowStr],
      ['Admin PIN', def.adminPin || '1234', 'PIN Autentikasi Admin', nowStr],
      ['Status Update', 'Aktif & Terverifikasi', 'Status Profil', nowStr]
    ]);
  }

  // 3. ContactInfo
  var cSheet = getOrCreateSheet_(ss, 'ContactInfo', ['Icon', 'Label', 'Platform / Handle', 'Link URL', 'Last Updated'], color);
  if (cSheet.getLastRow() < 2 && def.contactInfo && def.contactInfo.length) {
    var cRows = def.contactInfo.map(function(c){
      return [c.icon || 'info', c.label || '', c.value || '', c.href || '', nowStr];
    });
    cSheet.getRange(2, 1, cRows.length, 5).setValues(cRows);
  }

  // 4. SocialStats
  var sSheet = getOrCreateSheet_(ss, 'SocialStats', ['Platform / Handle', 'Value', 'Label', 'Icon', 'Last Updated'], color);
  if (sSheet.getLastRow() < 2 && def.socialStats && def.socialStats.length) {
    var sRows = def.socialStats.map(function(s){
      return [s.handle || s.platform || '', s.value || '', s.label || '', s.icon || 'bar-chart-2', nowStr];
    });
    sSheet.getRange(2, 1, sRows.length, 5).setValues(sRows);
  }

  // 5. AudienceInfo
  var aSheet = getOrCreateSheet_(ss, 'AudienceInfo', ['Label', 'Value', 'Icon', 'Last Updated'], color);
  if (aSheet.getLastRow() < 2 && def.audienceInfo && def.audienceInfo.length) {
    var aRows = def.audienceInfo.map(function(a){
      return [a.label || '', a.value || '', a.icon || 'info', nowStr];
    });
    aSheet.getRange(2, 1, aRows.length, 4).setValues(aRows);
  }

  // 6. BrandCollabs
  var bSheet = getOrCreateSheet_(ss, 'BrandCollabs', ['Nama Brand', 'Logo URL', 'Fallback Text', 'Last Updated'], color);
  if (bSheet.getLastRow() < 2 && def.brandCollabs && def.brandCollabs.length) {
    var bRows = def.brandCollabs.map(function(b){
      return [b.name || '', b.logoUrl || '', b.textFallback || b.name || '', nowStr];
    });
    bSheet.getRange(2, 1, bRows.length, 4).setValues(bRows);
  }

  // 7. Videos
  var vSheet = getOrCreateSheet_(ss, 'Videos', ['ID', 'Judul Video', 'Platform', 'Durasi', 'Thumbnail URL', 'Video URL', 'Embed URL', 'Source Type', 'Durasi Detik', 'Last Updated'], color);
  if (vSheet.getLastRow() < 2 && def.videos && def.videos.length) {
    var vRows = def.videos.map(function(v){
      return [
        v.id || '',
        v.title || '',
        v.platform || '',
        v.duration || '',
        v.thumbnailUrl || '',
        v.videoUrl || '',
        v.embedUrl || '',
        v.sourceType || '',
        Number(v.durationSeconds || 0),
        nowStr
      ];
    });
    vSheet.getRange(2, 1, vRows.length, 10).setValues(vRows);
  }

  // 8. RateCards
  var rSheet = getOrCreateSheet_(ss, 'RateCards', ['ID', 'Nama Paket', 'Harga', 'Popular', 'Ikon', 'Fitur (Dipisah Koma)', 'Last Updated'], color);
  if (rSheet.getLastRow() < 2 && def.rateCards && def.rateCards.length) {
    var rRows = def.rateCards.map(function(r){
      var featStr = Array.isArray(r.features) ? r.features.join(', ') : String(r.features || '');
      return [
        r.id || '',
        r.title || '',
        r.price || '',
        r.popular ? 'TRUE' : 'FALSE',
        r.icon || 'package',
        featStr,
        nowStr
      ];
    });
    rSheet.getRange(2, 1, rRows.length, 7).setValues(rRows);
  }

  // 9. CollabImages
  var ciSheet = getOrCreateSheet_(ss, 'CollabImages', ['ID', 'Judul / Brand', 'Image URL', 'Caption / Keterangan', 'Kategori', 'Last Updated'], color);
  if (ciSheet.getLastRow() < 2 && def.collabImages && def.collabImages.length) {
    var ciRows = def.collabImages.map(function(img){
      return [
        img.id || '',
        img.title || '',
        img.imageUrl || '',
        img.caption || '',
        img.category || '',
        nowStr
      ];
    });
    ciSheet.getRange(2, 1, ciRows.length, 6).setValues(ciRows);
  }

  // 10. Invoices
  getOrCreateSheet_(ss, 'Invoices', ['ID', 'Invoice Number', 'Nama Pelanggan', 'Tanggal Kerja Sama', 'Paket Rate Card', 'Nomor Transfer', 'Syarat Ketentuan', 'Signature Data', 'Created At', 'Updated At'], color);

  // 11. Inquiries
  getOrCreateSheet_(ss, 'Inquiries', ['Tanggal & Waktu', 'Nama', 'Brand', 'Email', 'Paket', 'Pesan', 'Status'], color);
}

/**
 * Menyimpan seluruh data portfolio ke 10+ sheet Google Spreadsheet
 */
function writePortfolioToSpreadsheet_(ss, data) {
  if (!ss || !data) return false;

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
    var color = THEME_ACCENT_COLOR || '#0090A8';

    // 1. SHEET 'Portfolio_Data' (Master Record)
    var masterSheet = getOrCreateSheet_(ss, 'Portfolio_Data', ['Key', 'Value / Detail', 'Last Updated'], color);
    masterSheet.clearContents();
    masterSheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value / Detail', 'Last Updated']]);
    masterSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    masterSheet.setFrozenRows(1);

    var masterRows = [
      ['MASTER_JSON', JSON.stringify(data), nowStr],
      ['ADMIN_PIN', data.adminPin || '1234', nowStr],
      ['WHATSAPP_NUMBER', data.whatsappNumber || '', nowStr],
      ['PROFILE_NAME', (data.profile && data.profile.name) || '', nowStr],
      ['PROFILE_TAGLINE', (data.profile && data.profile.tagline) || '', nowStr],
      ['PROFILE_BIO', (data.profile && data.profile.bio) || '', nowStr],
      ['PROFILE_LOCATION', (data.profile && data.profile.location) || '', nowStr],
      ['PROFILE_EMAIL', (data.profile && data.profile.email) || '', nowStr],
      ['PROFILE_AVATAR_URL', (data.profile && data.profile.avatarUrl) || '', nowStr]
    ];
    masterSheet.getRange(2, 1, masterRows.length, 3).setValues(masterRows);

    // 2. SHEET 'Profile'
    var pSheet = getOrCreateSheet_(ss, 'Profile', ['Field', 'Value', 'Keterangan / Status', 'Last Updated'], color);
    pSheet.clearContents();
    pSheet.getRange(1, 1, 1, 4).setValues([['Field', 'Value', 'Keterangan / Status', 'Last Updated']]);
    pSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    pSheet.setFrozenRows(1);
    var pRows = [
      ['Name', (data.profile && data.profile.name) || '', 'Nama Lengkap Creator', nowStr],
      ['Tagline', (data.profile && data.profile.tagline) || '', 'Tagline / Deskripsi Singkat', nowStr],
      ['Bio', (data.profile && data.profile.bio) || '', 'Biografi Lengkap', nowStr],
      ['Avatar URL', (data.profile && data.profile.avatarUrl) || '', 'Link Foto Profil', nowStr],
      ['Location', (data.profile && data.profile.location) || '', 'Domisili / Lokasi', nowStr],
      ['Email', (data.profile && data.profile.email) || '', 'Email Kontak Utama', nowStr],
      ['WhatsApp Number', data.whatsappNumber || '', 'Nomor WhatsApp Utama', nowStr],
      ['Admin PIN', data.adminPin || '1234', 'PIN Autentikasi Admin', nowStr],
      ['Status Update', 'Aktif & Terverifikasi', 'Status Profil', nowStr]
    ];
    pSheet.getRange(2, 1, pRows.length, 4).setValues(pRows);

    // 3. SHEET 'ContactInfo'
    var cSheet = getOrCreateSheet_(ss, 'ContactInfo', ['Icon', 'Label', 'Platform / Handle', 'Link URL', 'Last Updated'], color);
    cSheet.clearContents();
    cSheet.getRange(1, 1, 1, 5).setValues([['Icon', 'Label', 'Platform / Handle', 'Link URL', 'Last Updated']]);
    cSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    cSheet.setFrozenRows(1);
    if (data.contactInfo && data.contactInfo.length) {
      var cRows = data.contactInfo.map(function(c){
        return [c.icon || 'info', c.label || '', c.value || '', c.href || '', nowStr];
      });
      cSheet.getRange(2, 1, cRows.length, 5).setValues(cRows);
    }

    // 4. SHEET 'SocialStats'
    var sSheet = getOrCreateSheet_(ss, 'SocialStats', ['Platform / Handle', 'Value', 'Label', 'Icon', 'Last Updated'], color);
    sSheet.clearContents();
    sSheet.getRange(1, 1, 1, 5).setValues([['Platform / Handle', 'Value', 'Label', 'Icon', 'Last Updated']]);
    sSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    sSheet.setFrozenRows(1);
    if (data.socialStats && data.socialStats.length) {
      var sRows = data.socialStats.map(function(s){
        return [s.handle || s.platform || '', s.value || '', s.label || '', s.icon || 'bar-chart-2', nowStr];
      });
      sSheet.getRange(2, 1, sRows.length, 5).setValues(sRows);
    }

    // 5. SHEET 'AudienceInfo'
    var aSheet = getOrCreateSheet_(ss, 'AudienceInfo', ['Label', 'Value', 'Icon', 'Last Updated'], color);
    aSheet.clearContents();
    aSheet.getRange(1, 1, 1, 4).setValues([['Label', 'Value', 'Icon', 'Last Updated']]);
    aSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    aSheet.setFrozenRows(1);
    if (data.audienceInfo && data.audienceInfo.length) {
      var aRows = data.audienceInfo.map(function(a){
        return [a.label || '', a.value || '', a.icon || 'info', nowStr];
      });
      aSheet.getRange(2, 1, aRows.length, 4).setValues(aRows);
    }

    // 6. SHEET 'BrandCollabs'
    var bSheet = getOrCreateSheet_(ss, 'BrandCollabs', ['Nama Brand', 'Logo URL', 'Fallback Text', 'Last Updated'], color);
    bSheet.clearContents();
    bSheet.getRange(1, 1, 1, 4).setValues([['Nama Brand', 'Logo URL', 'Fallback Text', 'Last Updated']]);
    bSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    bSheet.setFrozenRows(1);
    if (data.brandCollabs && data.brandCollabs.length) {
      var bRows = data.brandCollabs.map(function(b){
        return [b.name || '', b.logoUrl || '', b.textFallback || b.name || '', nowStr];
      });
      bSheet.getRange(2, 1, bRows.length, 4).setValues(bRows);
    }

    // 7. SHEET 'Videos'
    var vSheet = getOrCreateSheet_(ss, 'Videos', ['ID', 'Judul Video', 'Platform', 'Durasi', 'Thumbnail URL', 'Video URL', 'Embed URL', 'Source Type', 'Durasi Detik', 'Last Updated'], color);
    vSheet.clearContents();
    vSheet.getRange(1, 1, 1, 10).setValues([['ID', 'Judul Video', 'Platform', 'Durasi', 'Thumbnail URL', 'Video URL', 'Embed URL', 'Source Type', 'Durasi Detik', 'Last Updated']]);
    vSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    vSheet.setFrozenRows(1);
    if (data.videos && data.videos.length) {
      var vRows = data.videos.map(function(v){
        return [
          v.id || '',
          v.title || '',
          v.platform || '',
          v.duration || '',
          v.thumbnailUrl || '',
          v.videoUrl || '',
          v.embedUrl || '',
          v.sourceType || '',
          Number(v.durationSeconds || 0),
          nowStr
        ];
      });
      vSheet.getRange(2, 1, vRows.length, 10).setValues(vRows);
    }

    // 8. SHEET 'RateCards'
    var rSheet = getOrCreateSheet_(ss, 'RateCards', ['ID', 'Nama Paket', 'Harga', 'Popular', 'Ikon', 'Fitur (Dipisah Koma)', 'Last Updated'], color);
    rSheet.clearContents();
    rSheet.getRange(1, 1, 1, 7).setValues([['ID', 'Nama Paket', 'Harga', 'Popular', 'Ikon', 'Fitur (Dipisah Koma)', 'Last Updated']]);
    rSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    rSheet.setFrozenRows(1);
    if (data.rateCards && data.rateCards.length) {
      var rRows = data.rateCards.map(function(r){
        var featStr = Array.isArray(r.features) ? r.features.join(', ') : String(r.features || '');
        return [
          r.id || '',
          r.title || '',
          r.price || '',
          r.popular ? 'TRUE' : 'FALSE',
          r.icon || 'package',
          featStr,
          nowStr
        ];
      });
      rSheet.getRange(2, 1, rRows.length, 7).setValues(rRows);
    }

    // 9. SHEET 'CollabImages'
    var ciSheet = getOrCreateSheet_(ss, 'CollabImages', ['ID', 'Judul / Brand', 'Image URL', 'Caption / Keterangan', 'Kategori', 'Last Updated'], color);
    ciSheet.clearContents();
    ciSheet.getRange(1, 1, 1, 6).setValues([['ID', 'Judul / Brand', 'Image URL', 'Caption / Keterangan', 'Kategori', 'Last Updated']]);
    ciSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground(color).setFontColor('#ffffff');
    ciSheet.setFrozenRows(1);
    if (data.collabImages && data.collabImages.length) {
      var ciRows = data.collabImages.map(function(img){
        return [
          img.id || '',
          img.title || '',
          img.imageUrl || '',
          img.caption || '',
          img.category || '',
          nowStr
        ];
      });
      ciSheet.getRange(2, 1, ciRows.length, 6).setValues(ciRows);
    }

    // 10. SHEET 'Invoices' & 11. SHEET 'Inquiries'
    getOrCreateSheet_(ss, 'Invoices', ['ID', 'Invoice Number', 'Nama Pelanggan', 'Tanggal Kerja Sama', 'Paket Rate Card', 'Nomor Transfer', 'Syarat Ketentuan', 'Signature Data', 'Created At', 'Updated At'], color);
    getOrCreateSheet_(ss, 'Inquiries', ['Tanggal & Waktu', 'Nama', 'Brand', 'Email', 'Paket', 'Pesan', 'Status'], color);

    return true;
  } catch (err) {
    Logger.log('writePortfolioToSpreadsheet_ error: ' + err);
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membaca data portfolio dari Google Spreadsheet.
 */
function readPortfolioFromSpreadsheet_(ss) {
  if (!ss) return null;

  try {
    var masterSheet = ss.getSheetByName('Portfolio_Data');
    if (masterSheet && masterSheet.getLastRow() >= 2) {
      var values = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        if (values[i][0] === 'MASTER_JSON' && values[i][1]) {
          try {
            var parsed = JSON.parse(values[i][1]);
            if (parsed && typeof parsed === 'object') return parsed;
          } catch (e) {}
        }
      }
    }

    // Rekonstruksi dari Sheet individual jika Master JSON tidak ada
    var result = getDefaultData_();
    var pSheet = ss.getSheetByName('Profile');
    if (pSheet && pSheet.getLastRow() >= 2) {
      var pVals = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 2).getValues();
      pVals.forEach(function(row){
        var key = String(row[0] || '').toLowerCase().trim();
        var val = String(row[1] || '').trim();
        if (key === 'name') result.profile.name = val;
        if (key === 'tagline') result.profile.tagline = val;
        if (key === 'bio') result.profile.bio = val;
        if (key === 'avatar url') result.profile.avatarUrl = val;
        if (key === 'location') result.profile.location = val;
        if (key === 'email') result.profile.email = val;
        if (key === 'whatsapp number') result.whatsappNumber = val;
        if (key === 'admin pin' && val) result.adminPin = val;
      });
    }

    var sSheet = ss.getSheetByName('SocialStats') || ss.getSheetByName('Social_Stats');
    if (sSheet && sSheet.getLastRow() >= 2) {
      var sVals = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 4).getValues();
      result.socialStats = sVals.map(function(r){
        return { platform: String(r[0]||''), handle: String(r[0]||''), value: String(r[1]||''), label: String(r[2]||''), icon: String(r[3]||'bar-chart-2') };
      });
    }

    var aSheet = ss.getSheetByName('AudienceInfo') || ss.getSheetByName('Audience');
    if (aSheet && aSheet.getLastRow() >= 2) {
      var aVals = aSheet.getRange(2, 1, aSheet.getLastRow() - 1, 3).getValues();
      result.audienceInfo = aVals.map(function(r){
        return { label: String(r[0]||''), value: String(r[1]||''), icon: String(r[2]||'info') };
      });
    }

    var bSheet = ss.getSheetByName('BrandCollabs') || ss.getSheetByName('Brands');
    if (bSheet && bSheet.getLastRow() >= 2) {
      var bVals = bSheet.getRange(2, 1, bSheet.getLastRow() - 1, 3).getValues();
      result.brandCollabs = bVals.map(function(r){
        return { name: String(r[0]||''), logoUrl: String(r[1]||''), textFallback: String(r[2]||r[0]||'') };
      });
    }

    var vSheet = ss.getSheetByName('Videos');
    if (vSheet && vSheet.getLastRow() >= 2) {
      var vVals = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 9).getValues();
      result.videos = vVals.map(function(r, idx){
        return {
          id: String(r[0] || ('v' + (idx + 1))),
          title: String(r[1] || ''),
          platform: String(r[2] || ''),
          duration: String(r[3] || ''),
          thumbnailUrl: String(r[4] || ''),
          videoUrl: String(r[5] || ''),
          embedUrl: String(r[6] || ''),
          sourceType: String(r[7] || ''),
          durationSeconds: Number(r[8] || 0)
        };
      });
    }

    var rSheet = ss.getSheetByName('RateCards') || ss.getSheetByName('Rate_Cards');
    if (rSheet && rSheet.getLastRow() >= 2) {
      var rVals = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 6).getValues();
      result.rateCards = rVals.map(function(r, idx){
        var feats = String(r[5] || '').split(',').map(function(x){ return x.trim(); }).filter(Boolean);
        return {
          id: String(r[0] || ('r' + (idx + 1))),
          title: String(r[1] || ''),
          price: String(r[2] || ''),
          popular: String(r[3]).toUpperCase() === 'TRUE',
          icon: String(r[4] || 'package'),
          features: feats
        };
      });
    }

    var cSheet = ss.getSheetByName('ContactInfo') || ss.getSheetByName('Contact_Info');
    if (cSheet && cSheet.getLastRow() >= 2) {
      var cVals = cSheet.getRange(2, 1, cSheet.getLastRow() - 1, 4).getValues();
      result.contactInfo = cVals.map(function(r){
        return { icon: String(r[0] || 'info'), label: String(r[1] || ''), value: String(r[2] || ''), href: String(r[3] || '') };
      });
    }

    var ciSheet = ss.getSheetByName('CollabImages');
    if (ciSheet && ciSheet.getLastRow() >= 2) {
      var ciVals = ciSheet.getRange(2, 1, ciSheet.getLastRow() - 1, 5).getValues();
      result.collabImages = ciVals.map(function(r, idx){
        return {
          id: String(r[0] || ('img' + (idx + 1))),
          title: String(r[1] || ''),
          imageUrl: String(r[2] || ''),
          caption: String(r[3] || ''),
          category: String(r[4] || '')
        };
      });
    }

    return result;
  } catch (err) {
    Logger.log('readPortfolioFromSpreadsheet_ error: ' + err);
    return null;
  }
}

/* ================================================================
   3. DATA DEFAULT & NORMALISASI
================================================================ */

function getDefaultData_() {
  return {
    adminPin: '1234',
    whatsappNumber: '6281234567890',

    profile: {
      name: 'Safwa Sakilla',
      tagline: 'Lifestyle & Beauty Content Creator',
      bio: 'Saya adalah content creator yang berfokus pada lifestyle, beauty, self development dan daily life. Saya suka membuat konten yang autentik, aesthetic, dan memberikan value untuk audiens saya.',
      avatarUrl: 'https://images.unsplash.com/photo-1544005313-94ddf0286df2?auto=format&fit=crop&w=600&q=80',
      location: 'Jakarta, Indonesia',
      email: 'safwa@email.com'
    },

    contactInfo: [
      { icon: 'youtube', label: 'YOUTUBE', value: '@safwasakilla', href: 'https://youtube.com/@safwasakilla' },
      { icon: 'twitter', label: 'TWITTER / X', value: '@safwasakilla', href: 'https://x.com/safwasakilla' },
      { icon: 'instagram', label: 'INSTAGRAM', value: '@safwasakilla', href: 'https://instagram.com/safwasakilla' },
      { icon: 'music', label: 'TIKTOK', value: '@safwasakilla', href: 'https://tiktok.com/@safwasakilla' }
    ],

    socialStats: [
      { platform: 'Instagram', handle: '@safwasakilla', value: '250K', label: 'Followers', icon: 'instagram' },
      { platform: 'TikTok', handle: '@safwasakilla', value: '180K', label: 'Followers', icon: 'music' },
      { platform: 'YouTube', handle: 'Safwa Sakilla', value: '75K', label: 'Subscribers', icon: 'youtube' },
      { platform: 'Engagement Rate', handle: '', value: '4.8%', label: 'Rata-rata', icon: 'trending-up' },
      { platform: 'Reach Bulanan', handle: '', value: '1.2M+', label: 'Accounts', icon: 'eye' },
      { platform: 'Audience Dominan', handle: '', value: '18 - 34', label: 'Tahun', icon: 'users' }
    ],

    audienceInfo: [
      { icon: 'sparkles', label: 'Niche', value: 'Lifestyle, Beauty, Fashion' },
      { icon: 'users', label: 'Gender Audience', value: '78% Wanita' },
      { icon: 'map-pin', label: 'Lokasi Audience', value: '92% Indonesia' },
      { icon: 'globe', label: 'Bahasa', value: 'Indonesia' }
    ],

    brandCollabs: [
      { name: 'Somethinc', logoUrl: '', textFallback: 'Somethinc' },
      { name: 'Avoskin', logoUrl: 'https://images.unsplash.com/photo-1556742049-0cfed4f6a45d?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Scarlett', logoUrl: 'https://images.unsplash.com/photo-1542744094-3a31f272c490?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Erha', logoUrl: 'https://images.unsplash.com/photo-1493119508027-2b584f234d6c?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Make Over', logoUrl: 'https://images.unsplash.com/photo-1512436991641-6745cdb1723f?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' }
    ],

    videos: [
      { id: 'v1', title: 'Makeup Tutorial Collab', platform: 'Reels', duration: '0:45', thumbnailUrl: 'https://images.unsplash.com/photo-1487412912498-0447578fcca8?auto=format&fit=crop&w=600&q=70', videoUrl: 'https://www.w3schools.com/html/mov_bbb.mp4', embedUrl: '', sourceType: 'direct', durationSeconds: 45 },
      { id: 'v2', title: 'Skincare Routine Review', platform: 'TikTok', duration: '0:52', thumbnailUrl: 'https://images.unsplash.com/photo-1515688594390-b649af70d282?auto=format&fit=crop&w=600&q=70', videoUrl: 'https://www.w3schools.com/html/mov_bbb.mp4', embedUrl: '', sourceType: 'direct', durationSeconds: 52 },
      { id: 'v3', title: 'Beauty Brand Campaign', platform: 'Reels', duration: '0:48', thumbnailUrl: 'https://images.unsplash.com/photo-1583241801015-607ccbda4920?auto=format&fit=crop&w=600&q=70', videoUrl: 'https://www.w3schools.com/html/mov_bbb.mp4', embedUrl: '', sourceType: 'direct', durationSeconds: 48 },
      { id: 'v4', title: 'Lifestyle Vlog x Brand', platform: 'YouTube', duration: '1:05', thumbnailUrl: 'https://images.unsplash.com/photo-1522335789203-aabd1fc54bc9?auto=format&fit=crop&w=600&q=70', videoUrl: 'https://www.w3schools.com/html/mov_bbb.mp4', embedUrl: '', sourceType: 'direct', durationSeconds: 65 }
    ],

    rateCards: [
      { id: 'r1', title: 'Story 1 Frame', price: 'Rp500.000', popular: false, icon: 'smartphone', features: ['1 Story (1 Frame)', 'Tag @brand', 'Include Link'] },
      { id: 'r2', title: 'Story 3 Frame', price: 'Rp1.200.000', popular: false, icon: 'smartphone', features: ['3 Story (3 Frame)', 'Tag @brand', 'Include Link'] },
      { id: 'r3', title: 'Feed Post', price: 'Rp2.500.000', popular: true, icon: 'layers', features: ['1 Feed Instagram', 'Foto / Carousel', 'Caption Review', 'Tag @brand'] },
      { id: 'r4', title: 'Reel Video', price: 'Rp3.500.000', popular: false, icon: 'video', features: ['1 Video Reel (Max 60s)', 'Tag @brand', 'Include Link', 'Creative Editing'] },
      { id: 'r5', title: 'Campaign Package', price: 'Custom Price', popular: false, icon: 'film', features: ['Paket campaign', 'Content plan', 'Multiple platform', 'Harga menyesuaikan'] }
    ],

    collabImages: [
      { id: 'img1', title: 'Somethinc Campaign Shoot', imageUrl: 'https://images.unsplash.com/photo-1522335789203-aabd1fc54bc9?auto=format&fit=crop&w=600&q=70', caption: 'Kolaborasi photoshoot kampanye skincare', category: 'Beauty' },
      { id: 'img2', title: 'Avoskin Serum Editorial', imageUrl: 'https://images.unsplash.com/photo-1487412912498-0447578fcca8?auto=format&fit=crop&w=600&q=70', caption: 'Review dan photoshoot brightening serum', category: 'Skincare' },
      { id: 'img3', title: 'Scarlett Whitening Flatlay', imageUrl: 'https://images.unsplash.com/photo-1515688594390-b649af70d282?auto=format&fit=crop&w=600&q=70', caption: 'Product review dan aesthetic flatlay feed', category: 'Lifestyle' },
      { id: 'img4', title: 'Make Over Editorial Look', imageUrl: 'https://images.unsplash.com/photo-1583241801015-607ccbda4920?auto=format&fit=crop&w=600&q=70', caption: 'Makeup tutorial dan editorial photoshoot', category: 'Makeup' }
    ]
  };
}

function cloneObject_(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeContactInfo_(items) {
  if (items === null || items === undefined) {
    return getDefaultData_().contactInfo;
  }
  var source = Array.isArray(items) ? items : [];
  return source.map(function(item) {
    if (!item || typeof item !== 'object') return {};
    if (item.icon === undefined) item.icon = 'info';
    if (item.label === undefined) item.label = 'LABEL';
    if (item.value === undefined) item.value = '';
    if (item.href === undefined) item.href = '';
    return item;
  });
}

function normalizeVideoUrl_(url) {
  var u = String(url || '').trim();
  var empty = { sourceType: '', embedUrl: '', videoUrl: '', platform: '' };
  if (!u) return empty;

  var m;

  // YouTube
  m = u.match(/(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/)|youtu\.be\/)([A-Za-z0-9_-]{11})/);
  if (m) {
    return {
      sourceType: 'youtube',
      embedUrl: 'https://www.youtube.com/embed/' + m[1],
      videoUrl: u,
      platform: 'YouTube'
    };
  }

  // TikTok
  m = u.match(/tiktok\.com\/(?:@[^\/]+\/video\/|v\/)(\d+)/);
  if (m) {
    return {
      sourceType: 'tiktok',
      embedUrl: 'https://www.tiktok.com/embed/v2/' + m[1],
      videoUrl: u,
      platform: 'TikTok'
    };
  }

  // Instagram
  m = u.match(/instagram\.com\/(?:reel|p|tv)\/([A-Za-z0-9_-]+)/);
  if (m) {
    return {
      sourceType: 'instagram',
      embedUrl: 'https://www.instagram.com/reel/' + m[1] + '/embed',
      videoUrl: u,
      platform: 'Instagram'
    };
  }

  // Twitter / X
  m = u.match(/(?:twitter|x)\.com\/[^\/]+\/status\/(\d+)/);
  if (m) {
    return {
      sourceType: 'twitter',
      embedUrl: 'https://twitframe.com/show?url=' + encodeURIComponent(u),
      videoUrl: u,
      platform: 'Twitter'
    };
  }

  // Facebook
  m = u.match(/facebook\.com\/.*\/videos\/(\d+)/);
  if (m) {
    return {
      sourceType: 'facebook',
      embedUrl: 'https://www.facebook.com/plugins/video.php?href=' + encodeURIComponent(u) + '&show_text=0',
      videoUrl: u,
      platform: 'Facebook'
    };
  }

  // Direct video file
  if (/\.(mp4|webm|mov|m4v|ogg|ogv)(\?|$)/i.test(u) || /drive\.google\.com\/uc\?/i.test(u)) {
    return { sourceType: 'direct', embedUrl: '', videoUrl: u, platform: 'Video' };
  }

  return { sourceType: 'direct', embedUrl: '', videoUrl: u, platform: 'Video' };
}

function normalizeVideos_(items) {
  var list = Array.isArray(items) ? items : [];
  return list.map(function(v, idx) {
    var video = v && typeof v === 'object' ? v : {};
    if (!video.id) video.id = 'v' + (idx + 1) + '_' + new Date().getTime();
    if (video.title == null) video.title = '';
    if (video.platform == null) video.platform = '';
    if (video.duration == null) video.duration = '';
    if (video.thumbnailUrl == null) video.thumbnailUrl = '';
    if (video.videoUrl == null) video.videoUrl = '';
    if (video.embedUrl == null) video.embedUrl = '';
    if (video.sourceType == null) video.sourceType = '';
    if (video.durationSeconds == null) video.durationSeconds = 0;

    if (!video.sourceType && video.videoUrl && String(video.videoUrl).indexOf('data:') !== 0) {
      var info = normalizeVideoUrl_(video.videoUrl);
      if (info.sourceType) {
        video.sourceType = info.sourceType;
        if (!video.embedUrl) video.embedUrl = info.embedUrl;
        if (!video.platform && info.platform) video.platform = info.platform;
      }
    }
    return video;
  });
}

function getStoredData_() {
  // 1. Coba baca dari Google Spreadsheet terlebih dahulu (Primary Database)
  var ss = getSpreadsheet_();
  if (ss) {
    var ssData = readPortfolioFromSpreadsheet_(ss);
    if (ssData) return ssData;
  }

  // 2. Fallback baca dari Script Properties (jika spreadsheet belum dikonfigurasi)
  var props = PropertiesService.getScriptProperties();
  var json = '';
  var chunkCount = Number(props.getProperty(PORTFOLIO_CHUNK_COUNT_KEY) || 0);

  if (chunkCount > 0 && chunkCount <= 300) {
    var chunks = [];
    for (var i = 0; i < chunkCount; i++) {
      chunks.push(props.getProperty(PORTFOLIO_CHUNK_PREFIX + i) || '');
    }
    json = chunks.join('');
  } else {
    json = props.getProperty(PORTFOLIO_DATA_KEY) || '';
  }

  if (!json) return null;

  try {
    var parsed = JSON.parse(json);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (err) {
    Logger.log('Parsing fallback ScriptProperties gagal: ' + err);
    return null;
  }
}

function removePrivateFields_(data) {
  var copy = cloneObject_(data || {});
  delete copy.adminPin;
  (copy.videos || []).forEach(function(video) {
    delete video._localVideoFile;
  });
  return copy;
}

/**
 * Mengambil data portfolio untuk website (Frontend).
 * Kompatibel 100% dengan getPortfolioData() & getData()
 * Otomatis memastikan 10+ sheet terbuat saat website pertama kali diakses/dideploy.
 */
function getData() {
  try {
    var cache = CacheService.getScriptCache();
    var cachedJson = cache.get(PORTFOLIO_CACHE_KEY);
    if (cachedJson) {
      try {
        var cData = JSON.parse(cachedJson);
        if (cData) return removePrivateFields_(cData);
      } catch (ce) {}
    }

    var data = getStoredData_();
    var ss = getSpreadsheet_();

    if (!data) {
      data = getDefaultData_();
      if (ss) {
        writePortfolioToSpreadsheet_(ss, data);
      }
      writeFallbackScriptProperties_(data);
      Logger.log('Data default diinisialisasi.');
    } else if (ss) {
      // Pastikan seluruh 10+ sheet selalu ada & lengkap
      ensureAllSheetsExist_(ss, data);
    }

    data.contactInfo = normalizeContactInfo_(data.contactInfo);
    data.videos = normalizeVideos_(data.videos);

    try {
      cache.put(PORTFOLIO_CACHE_KEY, JSON.stringify(data), 180); // cache 3 menit
    } catch(errCache) {}

    return removePrivateFields_(data);
  } catch (err) {
    Logger.log('getData error: ' + err);
    return removePrivateFields_(getDefaultData_());
  }
}

// Alias function untuk kompatibilitas penuh
function getPortfolioData() {
  return getData();
}

/* ================================================================
   4. AUTENTIKASI ADMIN & PIN
================================================================ */

function constantTimeEquals_(left, right) {
  left = String(left || '');
  right = String(right || '');
  var different = left.length ^ right.length;
  var maxLength = Math.max(left.length, right.length);
  for (var i = 0; i < maxLength; i++) {
    different |= (left.charCodeAt(i) || 0) ^ (right.charCodeAt(i) || 0);
  }
  return different === 0;
}

function getAdminPin_() {
  var data = getStoredData_();
  return String((data && data.adminPin) || getDefaultData_().adminPin);
}

function verifyAdminPin(pin) {
  try {
    var cache = CacheService.getScriptCache();
    var attempts = Number(cache.get(ADMIN_ATTEMPT_KEY) || 0);
    if (attempts >= 10) {
      return { success: false, message: 'Terlalu banyak percobaan. Coba lagi dalam beberapa menit.' };
    }

    if (!constantTimeEquals_(pin, getAdminPin_())) {
      cache.put(ADMIN_ATTEMPT_KEY, String(attempts + 1), 300);
      Logger.log('Verifikasi PIN gagal.');
      return { success: false, message: 'PIN salah. Coba lagi.' };
    }

    cache.remove(ADMIN_ATTEMPT_KEY);
    var token = Utilities.getUuid();
    cache.put(ADMIN_SESSION_PREFIX + token, '1', ADMIN_SESSION_TTL_SECONDS);
    if (CacheService.getUserCache()) {
      CacheService.getUserCache().put(ADMIN_SESSION_PREFIX + token, '1', ADMIN_SESSION_TTL_SECONDS);
    }
    Logger.log('Sesi admin berhasil dibuat: ' + token);
    return { success: true, token: token };
  } catch (err) {
    Logger.log('verifyAdminPin error: ' + err);
    return { success: false, message: 'Verifikasi PIN gagal. Silakan coba lagi.' };
  }
}

function verifyPin(pin) {
  return verifyAdminPin(pin);
}

function isAdminSessionValid_(token) {
  if (!token) return false;
  var tok = String(token);
  var sCache = CacheService.getScriptCache();
  var uCache = CacheService.getUserCache();
  if (sCache && sCache.get(ADMIN_SESSION_PREFIX + tok) === '1') return true;
  if (uCache && uCache.get(ADMIN_SESSION_PREFIX + tok) === '1') return true;
  return false;
}

/* ================================================================
   5. GOOGLE DRIVE ASSETS UPLOADER
================================================================ */

function isDataUri_(value) {
  return typeof value === 'string' && /^data:[^;,]+;base64,/i.test(value);
}

function getTargetFolder_() {
  if (VIDEO_FOLDER_ID && String(VIDEO_FOLDER_ID).trim()) {
    try {
      return DriveApp.getFolderById(String(VIDEO_FOLDER_ID).trim());
    } catch (folderErr) {
      Logger.log('VIDEO_FOLDER_ID tidak valid, memakai root Drive: ' + folderErr);
    }
  }
  return null;
}

function saveDataUriToDrive_(dataUri, fileName, mimeType) {
  var match = String(dataUri).match(/^data:([^;,]+);base64,(.*)$/i);
  if (!match) throw new Error('Format data file tidak valid.');

  var detectedMime = (match[1] || mimeType || 'application/octet-stream').toLowerCase();
  var bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > MAX_ASSET_BYTES) {
    throw new Error('Ukuran file melebihi batas 25 MB.');
  }

  var isVideo = detectedMime.indexOf('video/') === 0;

  var safeName = String(fileName || 'portfolio-asset')
    .replace(/[^\w.\- ]+/g, '_')
    .substring(0, 100);

  var blob = Utilities.newBlob(bytes, detectedMime, safeName || 'portfolio-asset');
  var folder = getTargetFolder_();
  var file = folder ? folder.createFile(blob) : DriveApp.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    Logger.log('Sharing publik tidak tersedia: ' + sharingError);
  }

  var id = file.getId();
  Logger.log('Asset Drive dibuat: ' + id + ' (' + detectedMime + ')');

  if (isVideo) {
    return 'https://drive.google.com/uc?export=download&id=' + id;
  }
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w800';
}

function persistDataAssets_(data) {
  var warnings = [];
  var stamp = new Date().getTime();

  // Avatar
  if (data.profile && isDataUri_(data.profile.avatarUrl)) {
    data.profile.avatarUrl = saveDataUriToDrive_(data.profile.avatarUrl, 'portfolio-avatar-' + stamp, 'image/jpeg');
  }

  // Brand logos
  (data.brandCollabs || []).forEach(function(brand, index) {
    if (isDataUri_(brand.logoUrl)) {
      brand.logoUrl = saveDataUriToDrive_(brand.logoUrl, 'portfolio-brand-' + index + '-' + stamp, 'image/png');
    }
  });

  // Videos
  (data.videos || []).forEach(function(video, index) {
    if (isDataUri_(video.thumbnailUrl)) {
      video.thumbnailUrl = saveDataUriToDrive_(video.thumbnailUrl, 'portfolio-thumb-' + index + '-' + stamp, 'image/jpeg');
    }

    if (isDataUri_(video.videoUrl)) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" belum di-upload — mohon klik tombol upload lagi.');
      video.videoUrl = '';
    }

    if (video.videoUrl && String(video.videoUrl).indexOf('blob:') === 0) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" belum di-upload — mohon klik tombol upload lagi.');
      video.videoUrl = '';
    }

    var dur = Number(video.durationSeconds || 0);
    if (video.sourceType === 'upload' && dur > 0 && dur > MAX_VIDEO_DURATION_SECONDS) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" melebihi 5 menit dan tidak disimpan.');
      video.videoUrl = '';
      video.sourceType = '';
    }

    if (video.videoUrl && video.sourceType !== 'upload') {
      var info = normalizeVideoUrl_(video.videoUrl);
      if (info.sourceType) {
        if (!video.sourceType) video.sourceType = info.sourceType;
        if (!video.embedUrl) video.embedUrl = info.embedUrl;
        if (!video.platform && info.platform) video.platform = info.platform;
      }
    }
  });

  // Rate icons
  (data.rateCards || []).forEach(function(rate, index) {
    if (isDataUri_(rate.iconUrl)) {
      rate.iconUrl = saveDataUriToDrive_(rate.iconUrl, 'portfolio-rate-icon-' + index + '-' + stamp, 'image/png');
    }
  });

  // Collab gallery images
  (data.collabImages || []).forEach(function(img, index) {
    if (isDataUri_(img.imageUrl)) {
      img.imageUrl = saveDataUriToDrive_(img.imageUrl, 'portfolio-collab-img-' + index + '-' + stamp, 'image/jpeg');
    }
  });

  return warnings;
}

function writeFallbackScriptProperties_(data) {
  try {
    var json = JSON.stringify(data);
    var props = PropertiesService.getScriptProperties();
    var oldCount = Number(props.getProperty(PORTFOLIO_CHUNK_COUNT_KEY) || 0);
    props.deleteProperty(PORTFOLIO_DATA_KEY);
    props.deleteProperty(PORTFOLIO_CHUNK_COUNT_KEY);

    for (var i = 0; i < oldCount; i++) {
      props.deleteProperty(PORTFOLIO_CHUNK_PREFIX + i);
    }

    var chunkSize = 2000;
    var count = Math.max(1, Math.ceil(json.length / chunkSize));
    for (var chunkIndex = 0; chunkIndex < count; chunkIndex++) {
      props.setProperty(
        PORTFOLIO_CHUNK_PREFIX + chunkIndex,
        json.substring(chunkIndex * chunkSize, (chunkIndex + 1) * chunkSize)
      );
    }
    props.setProperty(PORTFOLIO_CHUNK_COUNT_KEY, String(count));
  } catch(e) {
    Logger.log('writeFallbackScriptProperties_ error: ' + e);
  }
}

/**
 * Menyimpan data portfolio dari Admin Panel ke Google Spreadsheet.
 */
function saveData(newData, authToken) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      Logger.log('saveData ditolak: sesi admin tidak valid.');
      return { success: false, message: 'Sesi admin tidak valid atau sudah berakhir. Silakan login kembali.' };
    }
    if (!newData || typeof newData !== 'object' || Array.isArray(newData)) {
      return { success: false, message: 'Data tidak valid.' };
    }

    var oldData = getStoredData_() || getDefaultData_();
    var cleanData = cloneObject_(newData);
    cleanData.contactInfo = normalizeContactInfo_(cleanData.contactInfo);
    cleanData.videos = normalizeVideos_(cleanData.videos);

    var requestedPin = cleanData.adminPin;
    cleanData.adminPin = requestedPin && String(requestedPin).length >= 4
      ? String(requestedPin)
      : String(oldData.adminPin || getDefaultData_().adminPin);

    var warnings = persistDataAssets_(cleanData);
    (cleanData.videos || []).forEach(function(video) {
      delete video._localVideoFile;
    });

    // 1. Simpan ke Google Spreadsheet sebagai Database Utama (10+ Sheets)
    var ss = getSpreadsheet_();
    var savedToSheets = false;
    if (ss) {
      savedToSheets = writePortfolioToSpreadsheet_(ss, cleanData);
    }

    // 2. Simpan juga ke Script Properties sebagai backup sinkronisasi
    writeFallbackScriptProperties_(cleanData);

    // 3. Hapus cache agar pemanggilan getPortfolioData berikutnya langsung fresh
    try {
      CacheService.getScriptCache().remove(PORTFOLIO_CACHE_KEY);
    } catch(errCache) {}

    Logger.log('Data disimpan ke Google Spreadsheet: ' + (savedToSheets ? 'Berhasil' : 'Fallback ScriptProperties'));

    return {
      success: true,
      message: savedToSheets 
        ? '✓ Seluruh data berhasil disimpan ke 10+ Sheet Google Spreadsheet!' 
        : 'Data tersimpan ke Google Apps Script (Hubungkan Spreadsheet untuk database penuh).',
      warning: warnings.length ? warnings.join(' ') : '',
      savedAt: new Date().toISOString(),
      data: removePrivateFields_(cleanData)
    };
  } catch (err) {
    Logger.log('saveData error: ' + err);
    return { success: false, message: 'Gagal menyimpan: ' + err.message };
  }
}

// Alias function
function savePortfolioData(newData, authToken) {
  return saveData(newData, authToken);
}

function uploadAsset(dataUri, fileName, mimeType, authToken, durationSeconds) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    if (!isDataUri_(dataUri)) {
      return { success: false, message: 'Data file tidak valid.' };
    }

    var mime = String(mimeType || '').toLowerCase();
    if (!mime) {
      var m = String(dataUri).match(/^data:([^;,]+);base64,/i);
      if (m) mime = m[1].toLowerCase();
    }

    var isVideo = mime.indexOf('video/') === 0;
    var isImage = mime.indexOf('image/') === 0;

    if (!isVideo && !isImage) {
      return { success: false, message: 'Tipe file tidak didukung. Hanya gambar atau video.' };
    }
    if (isVideo && ALLOWED_VIDEO_MIME.indexOf(mime) === -1) {
      return { success: false, message: 'Format video tidak didukung. Pakai MP4, WebM, atau MOV.' };
    }

    if (isVideo) {
      var dur = Number(durationSeconds || 0);
      if (!dur || dur <= 0) {
        return { success: false, message: 'Durasi video tidak terbaca. Coba upload ulang.' };
      }
      if (dur > MAX_VIDEO_DURATION_SECONDS) {
        return { success: false, message: 'Durasi maksimal 5 menit (300 detik).' };
      }
    }

    var url = saveDataUriToDrive_(dataUri, fileName, mime);
    return {
      success: true,
      url: url,
      sourceType: isVideo ? 'upload' : 'image',
      durationSeconds: isVideo ? Number(durationSeconds || 0) : 0,
      mimeType: mime
    };
  } catch (err) {
    Logger.log('uploadAsset error: ' + err);
    var msg = err && err.message ? err.message : String(err);
    if (msg.indexOf('exceed') !== -1 || msg.indexOf('too large') !== -1 || msg.indexOf('25 MB') !== -1) {
      msg = 'Ukuran file melebihi 25 MB. Pakai URL sosmed untuk file besar.';
    }
    return { success: false, message: 'Upload gagal: ' + msg };
  }
}

function detectVideoUrl(url) {
  try {
    return { success: true, info: normalizeVideoUrl_(url) };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

/* ================================================================
   6. INQUIRY KE GOOGLE SPREADSHEET
================================================================ */

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().substring(0, maxLength || 2000);
}

function submitInquiry(formData) {
  try {
    if (!formData || typeof formData !== 'object') {
      return { success: false, message: 'Data inquiry tidak valid.' };
    }

    var name = cleanText_(formData.name, 120);
    var brand = cleanText_(formData.brand, 160);
    var email = cleanText_(formData.email, 200);
    var service = cleanText_(formData.service, 160);
    var message = cleanText_(formData.message, 4000);

    if (!name || !brand || !email || !message) {
      return { success: false, message: 'Nama, brand, email, dan pesan wajib diisi.' };
    }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return { success: false, message: 'Format email tidak valid.' };
    }

    var ss = getSpreadsheet_();
    if (!ss) {
      Logger.log('Inquiry diterima tanpa spreadsheet: ' + name);
      return {
        success: true,
        message: 'Inquiry siap diteruskan.',
        warning: 'Spreadsheet belum terhubung. Pesan WhatsApp tetap dapat dikirim.'
      };
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sheet = getOrCreateSheet_(ss, 'Inquiries', ['Tanggal & Waktu', 'Nama', 'Brand', 'Email', 'Paket', 'Pesan', 'Status'], THEME_ACCENT_COLOR || '#0090A8');
      var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT+7', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([nowStr, name, brand, email, service, message, 'Baru']);
    } finally {
      lock.releaseLock();
    }

    Logger.log('Inquiry tercatat dari: ' + name + ' (' + brand + ')');
    return { success: true, message: 'Inquiry berhasil dicatat ke Google Spreadsheet.' };
  } catch (err) {
    Logger.log('submitInquiry error: ' + err);
    return { success: false, message: 'Inquiry tidak dapat dicatat ke Spreadsheet: ' + err.message };
  }
}

/* ================================================================
   7. INVOICE ENGINE (GOOGLE SPREADSHEET + PDF PRINT)
================================================================ */

var INVOICE_CHUNK_PREFIX = 'INVOICE_DATA_CHUNK_';
var INVOICE_CHUNK_COUNT_KEY = 'INVOICE_DATA_CHUNK_COUNT';

function readInvoicesFromSpreadsheet_(ss) {
  if (!ss) return null;
  try {
    var sheet = ss.getSheetByName('Invoices');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    return data.map(function(row){
      return {
        id: String(row[0] || ''),
        invoiceNumber: String(row[1] || ''),
        namaPelanggan: String(row[2] || ''),
        tanggalKerjaSama: String(row[3] || ''),
        paketRateCard: String(row[4] || ''),
        nomorTransfer: String(row[5] || ''),
        syaratKetentuan: String(row[6] || ''),
        signatureDataUrl: String(row[7] || ''),
        createdAt: String(row[8] || ''),
        updatedAt: String(row[9] || '')
      };
    }).filter(function(inv){ return inv.id !== ''; });
  } catch(e) {
    Logger.log('readInvoicesFromSpreadsheet_ error: ' + e);
    return null;
  }
}

function writeInvoicesToSpreadsheet_(ss, list) {
  if (!ss) return false;
  try {
    var sheet = getOrCreateSheet_(ss, 'Invoices', ['ID', 'Invoice Number', 'Nama Pelanggan', 'Tanggal Kerja Sama', 'Paket Rate Card', 'Nomor Transfer', 'Syarat Ketentuan', 'Signature Data', 'Created At', 'Updated At'], THEME_ACCENT_COLOR || '#0090A8');
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 10).setValues([['ID', 'Invoice Number', 'Nama Pelanggan', 'Tanggal Kerja Sama', 'Paket Rate Card', 'Nomor Transfer', 'Syarat Ketentuan', 'Signature Data', 'Created At', 'Updated At']]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground(THEME_ACCENT_COLOR || '#0090A8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    if (list && list.length) {
      var rows = list.map(function(inv){
        return [
          inv.id || '',
          inv.invoiceNumber || '',
          inv.namaPelanggan || '',
          inv.tanggalKerjaSama || '',
          inv.paketRateCard || '',
          inv.nomorTransfer || '',
          inv.syaratKetentuan || '',
          inv.signatureDataUrl || '',
          inv.createdAt || '',
          inv.updatedAt || ''
        ];
      });
      sheet.getRange(2, 1, rows.length, 10).setValues(rows);
    }
    return true;
  } catch(e) {
    Logger.log('writeInvoicesToSpreadsheet_ error: ' + e);
    return false;
  }
}

function readInvoices_() {
  // 1. Coba baca dari Spreadsheet
  var ss = getSpreadsheet_();
  if (ss) {
    var ssInvoices = readInvoicesFromSpreadsheet_(ss);
    if (ssInvoices !== null) return ssInvoices;
  }

  // 2. Fallback baca dari Script Properties
  var props = PropertiesService.getScriptProperties();
  var count = Number(props.getProperty(INVOICE_CHUNK_COUNT_KEY) || 0);
  if (count <= 0) return [];
  var chunks = [];
  for (var i = 0; i < count; i++) {
    chunks.push(props.getProperty(INVOICE_CHUNK_PREFIX + i) || '');
  }
  var json = chunks.join('');
  if (!json) return [];
  try {
    var parsed = JSON.parse(json);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    Logger.log('readInvoices_ parse error: ' + e);
    return [];
  }
}

function writeInvoices_(list) {
  var safeList = Array.isArray(list) ? list : [];
  
  // 1. Simpan ke Spreadsheet
  var ss = getSpreadsheet_();
  if (ss) {
    writeInvoicesToSpreadsheet_(ss, safeList);
  }

  // 2. Simpan juga ke Script Properties (Backup)
  try {
    var json = JSON.stringify(safeList);
    var props = PropertiesService.getScriptProperties();
    var oldCount = Number(props.getProperty(INVOICE_CHUNK_COUNT_KEY) || 0);
    props.deleteProperty(INVOICE_CHUNK_COUNT_KEY);
    for (var i = 0; i < oldCount; i++) {
      props.deleteProperty(INVOICE_CHUNK_PREFIX + i);
    }
    var chunkSize = 2000;
    var count = Math.max(1, Math.ceil(json.length / chunkSize));
    for (var c = 0; c < count; c++) {
      props.setProperty(INVOICE_CHUNK_PREFIX + c, json.substring(c * chunkSize, (c + 1) * chunkSize));
    }
    props.setProperty(INVOICE_CHUNK_COUNT_KEY, String(count));
  } catch(e) {
    Logger.log('writeInvoices_ properties backup error: ' + e);
  }
}

function getInvoices(authToken) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    return { success: true, invoices: readInvoices_() };
  } catch (e) {
    Logger.log('getInvoices error: ' + e);
    return { success: false, message: 'Gagal memuat invoice: ' + e.message };
  }
}

function saveInvoice(invoiceData, authToken) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    if (!invoiceData || typeof invoiceData !== 'object') {
      return { success: false, message: 'Data invoice tidak valid.' };
    }

    if (invoiceData.signatureDataUrl) {
      var sig = String(invoiceData.signatureDataUrl);
      if (!/^data:image\//i.test(sig)) {
        invoiceData.signatureDataUrl = '';
      }
      if (sig.length > 200000) {
        invoiceData.signatureDataUrl = '';
        Logger.log('signatureDataUrl terlalu besar, dibuang.');
      }
    }

    var list = readInvoices_();
    var now = new Date();

    if (invoiceData.id) {
      var found = false;
      for (var i = 0; i < list.length; i++) {
        if (list[i].id === invoiceData.id) {
          list[i] = Object.assign({}, list[i], invoiceData, { updatedAt: now.toISOString() });
          found = true;
          break;
        }
      }
      if (!found) return { success: false, message: 'Invoice tidak ditemukan.' };
    } else {
      var invNumber = 'INV-' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'GMT+7', 'yyyyMMdd') + '-' + String(list.length + 1).padStart(3, '0');
      invoiceData.id = Utilities.getUuid();
      invoiceData.invoiceNumber = invNumber;
      invoiceData.createdAt = now.toISOString();
      invoiceData.updatedAt = now.toISOString();
      list.push(invoiceData);
    }

    writeInvoices_(list);
    return { success: true, message: 'Invoice berhasil disimpan ke Google Spreadsheet.', invoices: list };
  } catch (e) {
    Logger.log('saveInvoice error: ' + e);
    return { success: false, message: 'Gagal menyimpan invoice: ' + e.message };
  }
}

function deleteInvoice(id, authToken) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    var list = readInvoices_();
    var newList = list.filter(function(inv) { return inv.id !== id; });
    if (newList.length === list.length) {
      return { success: false, message: 'Invoice tidak ditemukan.' };
    }
    writeInvoices_(newList);
    return { success: true, message: 'Invoice berhasil dihapus dari Google Spreadsheet.', invoices: newList };
  } catch (e) {
    Logger.log('deleteInvoice error: ' + e);
    return { success: false, message: 'Gagal menghapus invoice: ' + e.message };
  }
}

function generateInvoicePdf(id, authToken, signatureDataUrl) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    var list = readInvoices_();
    var inv = null;
    for (var i = 0; i < list.length; i++) {
      if (list[i].id === id) { inv = list[i]; break; }
    }
    if (!inv) return { success: false, message: 'Invoice tidak ditemukan.' };

    var data = getStoredData_() || getDefaultData_();
    var creatorName = (data.profile && data.profile.name) ? data.profile.name : 'Creator';
    var creatorLocation = (data.profile && data.profile.location) ? data.profile.location : '';
    var creatorEmail = (data.profile && data.profile.email) ? data.profile.email : '';

    var tanggal = inv.tanggalKerjaSama || inv.createdAt || '';
    var tanggalFormatted = '';
    try {
      var d = new Date(tanggal);
      tanggalFormatted = Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+7', 'dd MMMM yyyy');
    } catch(fe) { tanggalFormatted = tanggal; }

    var createdFormatted = '';
    try {
      var dc = new Date(inv.createdAt);
      createdFormatted = Utilities.formatDate(dc, Session.getScriptTimeZone() || 'GMT+7', 'dd MMMM yyyy');
    } catch(fe2) { createdFormatted = inv.createdAt || ''; }

    var syarat = String(inv.syaratKetentuan || '').replace(/\n/g, '<br>');

    var sigUrl = '';
    if (signatureDataUrl && /^data:image\//i.test(String(signatureDataUrl))) {
      sigUrl = signatureDataUrl;
    } else if (inv.signatureDataUrl && /^data:image\//i.test(String(inv.signatureDataUrl))) {
      sigUrl = inv.signatureDataUrl;
    }

    var creatorSignHtml = sigUrl
      ? '<img src="' + sigUrl + '" alt="Tanda Tangan" style="max-width:500px;max-height:290px;width:auto;height:auto;display:block;margin:0 auto 4px;object-fit:contain">'
      : '<div style="height:290px"></div>';

    var themeCol = THEME_ACCENT_COLOR || '#0090A8';

    var htmlContent = '<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8">' +
      '<title>Invoice ' + inv.invoiceNumber + '</title>' +
      '<style>' +
        'body{margin:0;padding:0;font-family:Arial,sans-serif;font-size:13px;color:#222}' +
        '.page{max-width:740px;margin:0 auto;padding:40px 44px}' +
        '.header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;padding-bottom:20px;border-bottom:3px solid ' + themeCol + '}' +
        '.brand-col h1{margin:0 0 4px;font-size:24px;color:' + themeCol + ';letter-spacing:.02em}' +
        '.brand-col p{margin:2px 0;font-size:11px;color:#666}' +
        '.inv-meta{text-align:right}' +
        '.inv-meta .inv-num{font-size:18px;font-weight:700;color:' + themeCol + ';margin-bottom:4px}' +
        '.inv-meta p{margin:2px 0;font-size:11px;color:#666}' +
        '.section-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:' + themeCol + ';margin:22px 0 8px}' +
        '.info-box{background:#F4F9FA;border:1px solid #D0E8EE;border-radius:8px;padding:14px 18px}' +
        '.info-row{display:flex;gap:8px;margin-bottom:6px;font-size:13px}' +
        '.info-row:last-child{margin-bottom:0}' +
        '.info-label{flex:0 0 150px;color:#666;font-size:12px}' +
        '.info-value{flex:1;font-weight:600;color:#222}' +
        '.paket-box{background:' + themeCol + ';color:#fff;border-radius:8px;padding:16px 18px;margin:12px 0}' +
        '.paket-box .paket-name{font-size:16px;font-weight:700;margin-bottom:4px}' +
        '.syarat-box{background:#FFFDE7;border:1px solid #FFF59D;border-radius:8px;padding:14px 18px;line-height:1.7;font-size:12px;color:#444}' +
        '.transfer-box{display:flex;align-items:center;gap:16px;background:#EBF7FA;border:1px solid #D0E8EE;border-radius:8px;padding:14px 18px}' +
        '.transfer-icon{font-size:28px}' +
        '.transfer-label{font-size:10px;color:#555;font-weight:800;text-transform:uppercase;letter-spacing:.08em}' +
        '.transfer-value{font-size:15px;font-weight:700;color:#111;letter-spacing:.04em}' +
        '.footer{margin-top:40px;padding-top:16px;border-top:1px solid #D0E8EE;display:flex;justify-content:space-between;font-size:11px;color:#888}' +
        '.sign-area{margin-top:48px;display:flex;justify-content:space-between;gap:24px}' +
        '.sign-block{text-align:center;flex:0 0 200px}' +
        '.sign-block .sign-img-wrap{min-height:290px;display:flex;align-items:flex-end;justify-content:center;margin-bottom:0}' +
        '.sign-line{width:180px;border-top:1px solid #aaa;margin:6px auto 6px}' +
        '.sign-label{font-size:11px;color:#666}' +
        '@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
      '</style></head><body><div class="page">' +
        '<div class="header">' +
          '<div class="brand-col"><h1>' + creatorName + '</h1>' +
            '<p>Creator Portfolio &amp; Media Kit</p>' +
            (creatorLocation ? '<p>' + creatorLocation + '</p>' : '') +
            (creatorEmail ? '<p>' + creatorEmail + '</p>' : '') +
          '</div>' +
          '<div class="inv-meta">' +
            '<div class="inv-num">' + inv.invoiceNumber + '</div>' +
            '<p>Tanggal Dibuat: ' + createdFormatted + '</p>' +
          '</div>' +
        '</div>' +
        '<div class="section-title">Data Pelanggan</div>' +
        '<div class="info-box">' +
          '<div class="info-row"><span class="info-label">Nama Pelanggan</span><span class="info-value">' + (inv.namaPelanggan || '-') + '</span></div>' +
          '<div class="info-row"><span class="info-label">Tanggal Kerja Sama</span><span class="info-value">' + tanggalFormatted + '</span></div>' +
        '</div>' +
        '<div class="section-title">Paket Rate Card</div>' +
        '<div class="paket-box">' +
          '<div class="paket-name">' + (inv.paketRateCard || '-') + '</div>' +
        '</div>' +
        '<div class="section-title">Nomor Tujuan Transfer</div>' +
        '<div class="transfer-box">' +
          '<div class="transfer-icon">🏦</div>' +
          '<div><div class="transfer-label">Transfer ke</div><div class="transfer-value">' + (inv.nomorTransfer || '-') + '</div></div>' +
        '</div>' +
        '<div class="section-title">Syarat &amp; Ketentuan</div>' +
        '<div class="syarat-box">' + (syarat || 'Tidak ada syarat &amp; ketentuan.') + '</div>' +
        '<div class="sign-area">' +
          '<div class="sign-block">' +
            '<div class="sign-img-wrap"><div style="height:70px"></div></div>' +
            '<div class="sign-line"></div>' +
            '<div class="sign-label">Pelanggan<br>' + (inv.namaPelanggan || '') + '</div>' +
          '</div>' +
          '<div class="sign-block">' +
            '<div class="sign-img-wrap">' + creatorSignHtml + '</div>' +
            '<div class="sign-line"></div>' +
            '<div class="sign-label">Creator<br>' + creatorName + '</div>' +
          '</div>' +
        '</div>' +
        '<div class="footer"><span>' + inv.invoiceNumber + ' · Digenerate otomatis</span><span>' + creatorName + ' &copy; 2026</span></div>' +
      '</div></body></html>';

    return { success: true, html: htmlContent, invoiceNumber: inv.invoiceNumber };
  } catch (e) {
    Logger.log('generateInvoicePdf error: ' + e);
    return { success: false, message: 'Gagal membuat invoice: ' + e.message };
  }
}

/* ================================================================
   8. RESET DATA
================================================================ */

function resetData() {
  try {
    var def = getDefaultData_();
    var ss = getSpreadsheet_();
    if (ss) {
      writePortfolioToSpreadsheet_(ss, def);
    }
    writeFallbackScriptProperties_(def);
    try {
      CacheService.getScriptCache().remove(PORTFOLIO_CACHE_KEY);
    } catch(errCache) {}

    Logger.log('Data direset ke default.');
    return { success: true, message: 'Data berhasil direset ke nilai default.' };
  } catch (err) {
    Logger.log('resetData error: ' + err);
    return { success: false, message: 'Gagal mereset data: ' + err.message };
  }
}
