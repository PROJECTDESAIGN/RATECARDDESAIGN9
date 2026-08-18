/**
 * Google Apps Script Backend — Creator Portfolio & Rate Card v3
 *
 * Perubahan v3 (Video Uploader):
 *   - Skema video baru: videoUrl, embedUrl, sourceType, durationSeconds
 *     (backward-compatible dengan data lama)
 *   - normalizeVideoUrl_() : deteksi platform sosmed (YouTube / TikTok /
 *     Instagram / X / Facebook / direct file) dan bangun embedUrl.
 *   - uploadAsset() sekarang mendukung MIME video (mp4, webm, mov, quicktime).
 *   - Video Drive dikembalikan sebagai URL uc?export=download&id=... agar
 *     bisa dimainkan tag <video> di browser.
 *   - Validasi server-side durasi <= 300 detik (5 menit) untuk video upload.
 *   - Batas ukuran per-file 25 MB (batas keras Apps Script).
 *   - Baris yang meng-clear blob: dihapus — video lokal wajib sudah di-upload
 *     via uploadAsset() sebelum saveData().
 *
 * File HTML harus diberi nama "index.html" (atau "Index.html" — GAS
 * tidak case-sensitive untuk createTemplateFromFile) agar doGet() dapat
 * menemukannya. Pastikan file HTML ini disertakan saat upload ke GAS.
 */

var PORTFOLIO_DATA_KEY = 'PORTFOLIO_DATA';
var PORTFOLIO_CHUNK_COUNT_KEY = 'PORTFOLIO_DATA_CHUNK_COUNT';
var PORTFOLIO_CHUNK_PREFIX = 'PORTFOLIO_DATA_CHUNK_';
var ADMIN_SESSION_PREFIX = 'portfolio_admin_session_';
var ADMIN_ATTEMPT_KEY = 'portfolio_admin_attempts';
var ADMIN_SESSION_TTL_SECONDS = 21600; // 6 jam

// ISI DENGAN FOLDER ID GOOGLE DRIVE ANDA (opsional).
// Contoh: '1AbCdEfGhIjKlMnOpQrStUvWxYz'
// Jika kosong, file akan disimpan di root My Drive milik pemilik script.
var VIDEO_FOLDER_ID = '';

// Batas ukuran keras Apps Script.
var MAX_ASSET_BYTES = 25 * 1024 * 1024;   // 25 MB
var MAX_VIDEO_DURATION_SECONDS = 300;     // 5 menit
var ALLOWED_VIDEO_MIME = ['video/mp4', 'video/webm', 'video/quicktime', 'video/x-m4v', 'video/ogg'];
var ALLOWED_IMAGE_MIME = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/gif'];

// JSON dapat berisi karakter Unicode yang mengambil lebih dari 1 byte.
var PROPERTY_CHUNK_SIZE = 2000;
var PROPERTY_TOTAL_LIMIT = 450000; // di bawah batas total 500 KB

function doGet(e) {
  // Catatan: addMetaTag hanya mendukung 'keywords' dan 'description'.
  // Meta viewport, theme-color, apple-* sudah ditulis langsung di Index.html.
  var output = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Creator Portfolio & Rate Card')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Creator Portfolio')
      .addItem('Reset data portfolio', 'resetData')
      .addToUi();
  } catch (err) {
    Logger.log('onOpen dilewati (script bukan spreadsheet-bound): ' + err);
  }
}

function onInstall(e) {
  onOpen(e);
}

/* ================================================================
   DATA DEFAULT DAN PEMBACAAN
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
      { name: 'Brand 2', logoUrl: 'https://images.unsplash.com/photo-1556742049-0cfed4f6a45d?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Brand 3', logoUrl: 'https://images.unsplash.com/photo-1542744094-3a31f272c490?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Brand 4', logoUrl: 'https://images.unsplash.com/photo-1493119508027-2b584f234d6c?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Brand 5', logoUrl: 'https://images.unsplash.com/photo-1512436991641-6745cdb1723f?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Brand 6', logoUrl: 'https://images.unsplash.com/photo-1522335789203-aabd1fc54bc9?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' },
      { name: 'Brand 7', logoUrl: 'https://images.unsplash.com/photo-1583241800698-e8ab01830a22?auto=format&fit=crop&w=120&h=60&q=60', textFallback: '' }
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
    ]
  };
}

function cloneObject_(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeContactInfo_(items) {
  var source = Array.isArray(items) ? items : [];
  var result = source.map(function(item) {
    if (!item || typeof item !== 'object') return {};
    // Pastikan setiap item memiliki field yang dibutuhkan tanpa menimpa data user.
    if (item.icon === undefined) item.icon = 'info';
    if (item.label === undefined) item.label = 'LABEL';
    if (item.value === undefined) item.value = '';
    if (item.href === undefined) item.href = '';
    return item;
  });

  // Jika benar-benar kosong, isi dengan contoh default agar UI tidak rusak.
  if (result.length === 0) {
    result = [
      { icon: 'youtube', label: 'YOUTUBE', value: '@safwasakilla', href: 'https://youtube.com/@safwasakilla' },
      { icon: 'twitter', label: 'TWITTER / X', value: '@safwasakilla', href: 'https://x.com/safwasakilla' },
      { icon: 'instagram', label: 'INSTAGRAM', value: '@safwasakilla', href: 'https://instagram.com/safwasakilla' },
      { icon: 'music', label: 'TIKTOK', value: '@safwasakilla', href: 'https://tiktok.com/@safwasakilla' }
    ];
  }

  return result;
}

/* ================================================================
   NORMALIZE VIDEO URL (deteksi platform sosmed)
================================================================ */

/**
 * Deteksi platform video dari URL dan bangun embedUrl-nya.
 * Return: { sourceType, embedUrl, videoUrl, platform }
 * sourceType: 'youtube' | 'tiktok' | 'instagram' | 'twitter' | 'facebook' | 'direct' | ''
 */
function normalizeVideoUrl_(url) {
  var u = String(url || '').trim();
  var empty = { sourceType: '', embedUrl: '', videoUrl: '', platform: '' };
  if (!u) return empty;

  var m;

  // YouTube (watch, shorts, embed, youtu.be)
  m = u.match(/(?:youtube\.com\/(?:watch\?v=|shorts\/|embed\/)|youtu\.be\/)([A-Za-z0-9_-]{11})/);
  if (m) {
    return {
      sourceType: 'youtube',
      embedUrl: 'https://www.youtube.com/embed/' + m[1],
      videoUrl: u,
      platform: 'YouTube'
    };
  }

  // TikTok (@user/video/ID atau v/ID)
  m = u.match(/tiktok\.com\/(?:@[^\/]+\/video\/|v\/)(\d+)/);
  if (m) {
    return {
      sourceType: 'tiktok',
      embedUrl: 'https://www.tiktok.com/embed/v2/' + m[1],
      videoUrl: u,
      platform: 'TikTok'
    };
  }

  // Instagram (reel, p, tv)
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

  // Direct video file (mp4, webm, mov, m4v, ogg) atau Google Drive uc?export=download
  if (/\.(mp4|webm|mov|m4v|ogg|ogv)(\?|$)/i.test(u) || /drive\.google\.com\/uc\?/i.test(u)) {
    return { sourceType: 'direct', embedUrl: '', videoUrl: u, platform: 'Video' };
  }

  // Fallback — anggap sebagai direct URL apa adanya
  return { sourceType: 'direct', embedUrl: '', videoUrl: u, platform: 'Video' };
}

/**
 * Normalisasi struktur video agar backward-compatible dengan data lama
 * (yang mungkin tidak punya field embedUrl / sourceType / durationSeconds).
 */
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

    // Deteksi otomatis sourceType kalau kosong tapi ada videoUrl
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
    Logger.log('PORTFOLIO_DATA tidak valid: ' + err);
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

function getData() {
  try {
    var data = getStoredData_();
    if (!data) {
      data = getDefaultData_();
      writeStoredData_(data);
      Logger.log('Data default dibuat.');
    }
    data.contactInfo = normalizeContactInfo_(data.contactInfo);
    data.videos = normalizeVideos_(data.videos);
    return removePrivateFields_(data);
  } catch (err) {
    Logger.log('getData error: ' + err);
    return removePrivateFields_(getDefaultData_());
  }
}

/* ================================================================
   AUTENTIKASI ADMIN
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
    var cache = CacheService.getUserCache();
    var attempts = Number(cache.get(ADMIN_ATTEMPT_KEY) || 0);
    if (attempts >= 5) {
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
    Logger.log('Sesi admin berhasil dibuat.');
    return { success: true, token: token };
  } catch (err) {
    Logger.log('verifyAdminPin error: ' + err);
    return { success: false, message: 'Verifikasi PIN gagal. Silakan coba lagi.' };
  }
}

function isAdminSessionValid_(token) {
  return Boolean(token) &&
    CacheService.getUserCache().get(ADMIN_SESSION_PREFIX + String(token)) === '1';
}

/* ================================================================
   PENYIMPANAN DAN ASSET DRIVE
================================================================ */

function isDataUri_(value) {
  return typeof value === 'string' && /^data:[^;,]+;base64,/i.test(value);
}

function getTargetFolder_() {
  if (VIDEO_FOLDER_ID && String(VIDEO_FOLDER_ID).trim()) {
    try {
      return DriveApp.getFolderById(String(VIDEO_FOLDER_ID).trim());
    } catch (folderErr) {
      Logger.log('VIDEO_FOLDER_ID tidak valid, pakai root: ' + folderErr);
    }
  }
  return null;
}

/**
 * Simpan data URI ke Drive. Return URL yang bisa langsung dipakai:
 *   - Gambar → https://drive.google.com/thumbnail?id=<ID>&sz=w800  (untuk <img>)
 *   - Video  → https://drive.google.com/uc?export=download&id=<ID> (untuk <video>)
 */
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
    // Format ini bisa langsung dimainkan oleh tag <video> di browser.
    return 'https://drive.google.com/uc?export=download&id=' + id;
  }
  // Gambar tetap pakai thumbnail agar tampil via <img>.
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
    // Thumbnail (gambar) – seperti sebelumnya.
    if (isDataUri_(video.thumbnailUrl)) {
      video.thumbnailUrl = saveDataUriToDrive_(video.thumbnailUrl, 'portfolio-thumb-' + index + '-' + stamp, 'image/jpeg');
    }

    // Kalau videoUrl masih data URI (belum di-upload via uploadAsset), tolak.
    if (isDataUri_(video.videoUrl)) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" belum di-upload — mohon klik tombol upload lagi.');
      video.videoUrl = '';
    }

    // Kalau videoUrl blob: (belum di-upload) → beri warning, kosongkan.
    if (video.videoUrl && String(video.videoUrl).indexOf('blob:') === 0) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" belum di-upload — mohon klik tombol upload lagi.');
      video.videoUrl = '';
    }

    // Validasi durasi (upload lokal) — hanya jika ada info durasi.
    var dur = Number(video.durationSeconds || 0);
    if (video.sourceType === 'upload' && dur > 0 && dur > MAX_VIDEO_DURATION_SECONDS) {
      warnings.push('Video "' + (video.title || ('#' + (index + 1))) + '" melebihi 5 menit dan tidak disimpan.');
      video.videoUrl = '';
      video.sourceType = '';
    }

    // Deteksi ulang sourceType untuk URL sosmed (jika kosong).
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

  return warnings;
}

function writeStoredData_(data) {
  var json = JSON.stringify(data);
  var jsonBytes = Utilities.newBlob(json, 'application/json').getBytes().length;
  if (jsonBytes > PROPERTY_TOTAL_LIMIT) {
    throw new Error('Data terlalu besar. Kurangi jumlah item atau ukuran file.');
  }

  var props = PropertiesService.getScriptProperties();
  var oldCount = Number(props.getProperty(PORTFOLIO_CHUNK_COUNT_KEY) || 0);
  props.deleteProperty(PORTFOLIO_DATA_KEY);
  props.deleteProperty(PORTFOLIO_CHUNK_COUNT_KEY);

  for (var i = 0; i < oldCount; i++) {
    props.deleteProperty(PORTFOLIO_CHUNK_PREFIX + i);
  }

  var count = Math.max(1, Math.ceil(json.length / PROPERTY_CHUNK_SIZE));
  for (var chunkIndex = 0; chunkIndex < count; chunkIndex++) {
    props.setProperty(
      PORTFOLIO_CHUNK_PREFIX + chunkIndex,
      json.substring(chunkIndex * PROPERTY_CHUNK_SIZE, (chunkIndex + 1) * PROPERTY_CHUNK_SIZE)
    );
  }
  props.setProperty(PORTFOLIO_CHUNK_COUNT_KEY, String(count));
}

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

    writeStoredData_(cleanData);
    Logger.log('Data disimpan: ' + new Date().toISOString());
    return {
      success: true,
      message: 'Data berhasil disimpan!',
      warning: warnings.length ? warnings.join(' ') : '',
      savedAt: new Date().toISOString(),
      data: removePrivateFields_(cleanData)
    };
  } catch (err) {
    Logger.log('saveData error: ' + err);
    return { success: false, message: 'Gagal menyimpan: ' + err.message };
  }
}

/**
 * Upload asset terpisah — dipakai frontend untuk mengunggah:
 *   • Gambar (image/*)
 *   • Video  (video/mp4, video/webm, video/quicktime, ...)
 * Untuk video, wajib menyertakan durationSeconds (dihitung di client).
 */
function uploadAsset(dataUri, fileName, mimeType, authToken, durationSeconds) {
  try {
    if (!isAdminSessionValid_(authToken)) {
      return { success: false, message: 'Sesi admin tidak valid.' };
    }
    if (!isDataUri_(dataUri)) {
      return { success: false, message: 'Data file tidak valid.' };
    }

    var mime = String(mimeType || '').toLowerCase();
    // Coba deteksi dari data URI jika mimeType tidak dikirim.
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

/**
 * Endpoint kecil untuk mendeteksi platform dari URL yang di-paste user.
 * Dipakai tombol "Deteksi & Preview" di frontend.
 */
function detectVideoUrl(url) {
  try {
    return { success: true, info: normalizeVideoUrl_(url) };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

/* ================================================================
   INQUIRY KE GOOGLE SHEETS
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

    var props = PropertiesService.getScriptProperties();
    var spreadsheetId = props.getProperty('INQUIRIES_SPREADSHEET_ID');
    var spreadsheet = null;

    try {
      spreadsheet = spreadsheetId
        ? SpreadsheetApp.openById(spreadsheetId)
        : SpreadsheetApp.getActiveSpreadsheet();
    } catch (spreadsheetError) {
      Logger.log('Spreadsheet inquiry tidak tersedia: ' + spreadsheetError);
    }

    if (!spreadsheet) {
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
      var sheet = spreadsheet.getSheetByName('Inquiries') || spreadsheet.insertSheet('Inquiries');
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, 6).setValues([[
          'Tanggal & Waktu', 'Nama', 'Brand', 'Email', 'Paket', 'Pesan'
        ]]);
        sheet.getRange(1, 1, 1, 6)
          .setFontWeight('bold')
          .setBackground('#4a6741')
          .setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
      sheet.appendRow([new Date(), name, brand, email, service, message]);
    } finally {
      lock.releaseLock();
    }

    Logger.log('Inquiry dari: ' + name + ' (' + brand + ')');
    return { success: true, message: 'Inquiry tercatat.' };
  } catch (err) {
    Logger.log('submitInquiry error: ' + err);
    return { success: false, message: 'Inquiry tidak dapat dicatat ke Spreadsheet: ' + err.message };
  }
}

/* ================================================================
   INVOICE — CRUD + GENERATE PDF (via HtmlService)
================================================================ */

var INVOICE_CHUNK_PREFIX = 'INVOICE_DATA_CHUNK_';
var INVOICE_CHUNK_COUNT_KEY = 'INVOICE_DATA_CHUNK_COUNT';

function readInvoices_() {
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
  var json = JSON.stringify(Array.isArray(list) ? list : []);
  var props = PropertiesService.getScriptProperties();
  var oldCount = Number(props.getProperty(INVOICE_CHUNK_COUNT_KEY) || 0);
  props.deleteProperty(INVOICE_CHUNK_COUNT_KEY);
  for (var i = 0; i < oldCount; i++) {
    props.deleteProperty(INVOICE_CHUNK_PREFIX + i);
  }
  var count = Math.max(1, Math.ceil(json.length / PROPERTY_CHUNK_SIZE));
  for (var c = 0; c < count; c++) {
    props.setProperty(INVOICE_CHUNK_PREFIX + c, json.substring(c * PROPERTY_CHUNK_SIZE, (c + 1) * PROPERTY_CHUNK_SIZE));
  }
  props.setProperty(INVOICE_CHUNK_COUNT_KEY, String(count));
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

    // Validasi signatureDataUrl — hanya terima data URI gambar yang valid
    if (invoiceData.signatureDataUrl) {
      var sig = String(invoiceData.signatureDataUrl);
      if (!/^data:image\//i.test(sig)) {
        invoiceData.signatureDataUrl = '';
      }
      // Batasi ukuran tanda tangan agar tidak membebani Properties (maks ~200 KB base64)
      if (sig.length > 200000) {
        invoiceData.signatureDataUrl = '';
        Logger.log('signatureDataUrl terlalu besar (' + Math.round(sig.length/1024) + ' KB), dibuang.');
      }
    }

    var list = readInvoices_();
    var now = new Date();

    if (invoiceData.id) {
      // Update invoice yang sudah ada
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
      // Invoice baru
      var invNumber = 'INV-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + String(list.length + 1).padStart(3, '0');
      invoiceData.id = Utilities.getUuid();
      invoiceData.invoiceNumber = invNumber;
      invoiceData.createdAt = now.toISOString();
      invoiceData.updatedAt = now.toISOString();
      list.push(invoiceData);
    }
    writeInvoices_(list);
    return { success: true, message: 'Invoice berhasil disimpan.', invoices: list };
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
    return { success: true, message: 'Invoice berhasil dihapus.', invoices: newList };
  } catch (e) {
    Logger.log('deleteInvoice error: ' + e);
    return { success: false, message: 'Gagal menghapus invoice: ' + e.message };
  }
}

/**
 * Menghasilkan konten HTML invoice sebagai string,
 * lalu dikembalikan ke frontend untuk di-print/save PDF via browser print dialog.
 *
 * @param {string} id          - ID invoice
 * @param {string} authToken   - Token sesi admin
 * @param {string} [signatureDataUrl] - Data URI PNG tanda tangan (opsional, override tanda tangan tersimpan)
 */
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

    // Ambil data portfolio untuk header invoice
    var data = getStoredData_() || getDefaultData_();
    var creatorName = (data.profile && data.profile.name) ? data.profile.name : 'Creator';
    var creatorLocation = (data.profile && data.profile.location) ? data.profile.location : '';
    var creatorEmail = (data.profile && data.profile.email) ? data.profile.email : '';

    var tanggal = inv.tanggalKerjaSama || inv.createdAt || '';
    var tanggalFormatted = '';
    try {
      var d = new Date(tanggal);
      tanggalFormatted = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMMM yyyy');
    } catch(fe) { tanggalFormatted = tanggal; }

    var createdFormatted = '';
    try {
      var dc = new Date(inv.createdAt);
      createdFormatted = Utilities.formatDate(dc, Session.getScriptTimeZone(), 'dd MMMM yyyy');
    } catch(fe2) { createdFormatted = inv.createdAt || ''; }

    var syarat = String(inv.syaratKetentuan || '').replace(/\n/g, '<br>');

    // Tentukan tanda tangan: parameter langsung lebih prioritas dari yang tersimpan di invoice
    var sigUrl = '';
    if (signatureDataUrl && /^data:image\//i.test(String(signatureDataUrl))) {
      sigUrl = signatureDataUrl;
    } else if (inv.signatureDataUrl && /^data:image\//i.test(String(inv.signatureDataUrl))) {
      sigUrl = inv.signatureDataUrl;
    }

    // Blok HTML tanda tangan creator — ukuran diperbesar agar standar dan jelas di PDF
var creatorSignHtml = sigUrl
  ? '<img src="' + sigUrl + '" alt="Tanda Tangan" style="max-width:500px;max-height:290px;width:auto;height:auto;display:block;margin:0 auto 4px;object-fit:contain">'
  : '<div style="height:290px"></div>';

var htmlContent = '<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8">' +
  '<title>Invoice ' + inv.invoiceNumber + '</title>' +
  '<style>' +
    'body{margin:0;padding:0;font-family:Arial,sans-serif;font-size:13px;color:#222}' +
    '.page{max-width:740px;margin:0 auto;padding:40px 44px}' +
    /* Border & aksen utama menggunakan warna Hijau Telur Asin (#5B9A8B) */
    '.header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:32px;padding-bottom:20px;border-bottom:3px solid #5B9A8B}' +
    '.brand-col h1{margin:0 0 4px;font-size:24px;color:#5B9A8B;letter-spacing:.02em}' +
    '.brand-col p{margin:2px 0;font-size:11px;color:#666}' +
    '.inv-meta{text-align:right}' +
    '.inv-meta .inv-num{font-size:18px;font-weight:700;color:#5B9A8B;margin-bottom:4px}' +
    '.inv-meta p{margin:2px 0;font-size:11px;color:#666}' +
    '.section-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:#5B9A8B;margin:22px 0 8px}' +
    /* Background info box disesuaikan dengan tone hijau telur asin soft (#F2F8F6) */
    '.info-box{background:#F2F8F6;border:1px solid #D2E7E2;border-radius:8px;padding:14px 18px}' +
    '.info-row{display:flex;gap:8px;margin-bottom:6px;font-size:13px}' +
    '.info-row:last-child{margin-bottom:0}' +
    '.info-label{flex:0 0 150px;color:#666;font-size:12px}' +
    '.info-value{flex:1;font-weight:600;color:#222}' +
    /* Box paket utama dengan warna Hijau Telur Asin (#5B9A8B) */
    '.paket-box{background:#5B9A8B;color:#fff;border-radius:8px;padding:16px 18px;margin:12px 0}' +
    '.paket-box .paket-name{font-size:16px;font-weight:700;margin-bottom:4px}' +
    /* Box syarat & ketentuan */
    '.syarat-box{background:#F8FCFA;border:1px solid #E1EFF3;border-radius:8px;padding:14px 18px;line-height:1.7;font-size:12px;color:#444}' +
    /* Box transfer disesuaikan ke nuansa soft teal/mint (#E8F3F1) */
    '.transfer-box{display:flex;align-items:center;gap:16px;background:#E8F3F1;border:1px solid #B8DDD6;border-radius:8px;padding:14px 18px}' +
    '.transfer-icon{font-size:28px}' +
    '.transfer-label{font-size:10px;color:#666;font-weight:800;text-transform:uppercase;letter-spacing:.08em}' +
    '.transfer-value{font-size:15px;font-weight:700;color:#396A5E;letter-spacing:.04em}' +
    '.footer{margin-top:40px;padding-top:16px;border-top:1px solid #D2E7E2;display:flex;justify-content:space-between;font-size:11px;color:#888}' +
    '.sign-area{margin-top:48px;display:flex;justify-content:space-between;gap:24px}' +
    '.sign-block{text-align:center;flex:0 0 200px}' +
    '.sign-block .sign-img-wrap{min-height:290px;display:flex;align-items:flex-end;justify-content:center;margin-bottom:0}' +
    '.sign-block .sign-line{width:180px;border-top:1px solid #aaa;margin:6px auto 6px}' +
    '.sign-block .sign-label{font-size:11px;color:#666}' +
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
   RESET
================================================================ */

function resetData() {
  try {
    var props = PropertiesService.getScriptProperties();
    var count = Number(props.getProperty(PORTFOLIO_CHUNK_COUNT_KEY) || 0);
    props.deleteProperty(PORTFOLIO_DATA_KEY);
    props.deleteProperty(PORTFOLIO_CHUNK_COUNT_KEY);
    for (var i = 0; i < count; i++) {
      props.deleteProperty(PORTFOLIO_CHUNK_PREFIX + i);
    }
    Logger.log('Data direset.');
    return { success: true, message: 'Data berhasil direset.' };
  } catch (err) {
    Logger.log('resetData error: ' + err);
    return { success: false, message: 'Gagal mereset data: ' + err.message };
  }
}
