function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Créateur d\'affiches Médiathèque A3');
}

function createPoster(payload) {
  try {
    var data = JSON.parse(payload);

    var entryId = data.entryId || Utilities.getUuid();
    var folder = ensurePosterFolder();

    var footerLogo1Id = '';
    var footerLogo2Id = '';

    if (data.footerLogo1) {
      if (data.footerLogo1Id) maybeTrashFile_(data.footerLogo1Id);
      footerLogo1Id = saveAssetFromDataUrl_(folder, entryId, 'FOOTER1', data.footerLogo1);
    } else if (data.footerLogo1Id) {
      footerLogo1Id = String(data.footerLogo1Id);
    }

    if (data.footerLogo2) {
      if (data.footerLogo2Id) maybeTrashFile_(data.footerLogo2Id);
      footerLogo2Id = saveAssetFromDataUrl_(folder, entryId, 'FOOTER2', data.footerLogo2);
    } else if (data.footerLogo2Id) {
      footerLogo2Id = String(data.footerLogo2Id);
    }

    data.footerLogo1 = data.footerLogo1 || (footerLogo1Id ? driveFileIdToDataUrl_(footerLogo1Id) : '');
    data.footerLogo2 = data.footerLogo2 || (footerLogo2Id ? driveFileIdToDataUrl_(footerLogo2Id) : '');

    data.footerLogo1Id = footerLogo1Id;
    data.footerLogo2Id = footerLogo2Id;

    var html = buildPosterHtml(data);

    var pdfBlob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);

    var dateForName = formatDateForFilename(data.dates || '');
    var titleForName = sanitizeFilename(data.title || 'AFFICHE');
    var fileName = 'MED-' + (dateForName || 'DATE') + '-' + (titleForName || 'AFFICHE') + '.pdf';

    var file = folder.createFile(pdfBlob).setName(fileName);

    savePosterEntry(entryId, data, file.getUrl());

    return JSON.stringify({
      ok: true,
      entryId: entryId,
      pdfUrl: file.getUrl(),
      fileName: fileName,
      footerLogo1Id: footerLogo1Id,
      footerLogo2Id: footerLogo2Id
    });
  } catch (e) {
    return JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) });
  }
}

function listPosterEntries() {
  var sheet = ensurePosterSheet();
  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) return JSON.stringify([]);

  var headers = values[0];
  var rows = values.slice(1).map(function (r) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = r[i]; });
    return obj;
  });

  rows.sort(function (a, b) {
    var da = new Date(a.updatedAt || a.createdAt || 0).getTime();
    var db = new Date(b.updatedAt || b.createdAt || 0).getTime();
    return db - da;
  });

  return JSON.stringify(rows);
}

function extractDriveFileId_(url) {
  var s = String(url || '');
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return '';
}

function dataUrlToBlob_(dataUrl, nameForBlob) {
  var s = String(dataUrl || '');
  var m = s.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) return null;
  var contentType = m[1];
  var bytes = Utilities.base64Decode(m[2]);
  return Utilities.newBlob(bytes, contentType, nameForBlob || 'asset');
}

function maybeTrashFile_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) {}
}

function saveAssetFromDataUrl_(folder, entryId, label, dataUrl) {
  var blob = dataUrlToBlob_(dataUrl, 'ASSET-' + entryId + '-' + label);
  if (!blob) return '';
  var ct = String(blob.getContentType() || '');
  var ext = 'bin';
  if (ct.indexOf('png') !== -1) ext = 'png';
  else if (ct.indexOf('jpeg') !== -1 || ct.indexOf('jpg') !== -1) ext = 'jpg';
  else if (ct.indexOf('gif') !== -1) ext = 'gif';
  else if (ct.indexOf('webp') !== -1) ext = 'webp';
  var file = folder.createFile(blob).setName('ASSET-' + entryId + '-' + label + '.' + ext);
  return file.getId();
}

function driveFileIdToDataUrl_(fileId) {
  if (!fileId) return '';
  var file = DriveApp.getFileById(String(fileId));
  var blob = file.getBlob();
  var ct = blob.getContentType() || 'application/octet-stream';
  var b64 = Utilities.base64Encode(blob.getBytes());
  return 'data:' + ct + ';base64,' + b64;
}

function deletePosterEntry(entryId) {
  try {
    var sheet = ensurePosterSheet();
    var rowIndex = findEntryRow(sheet, entryId);
    if (!rowIndex) return JSON.stringify({ ok: false, error: 'NOT_FOUND' });

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var pdfCol = headers.indexOf('pdfUrl') + 1;
    var pdfUrl = (pdfCol > 0) ? sheet.getRange(rowIndex, pdfCol).getValue() : '';

    var fileId = extractDriveFileId_(pdfUrl);
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
    }

    sheet.deleteRow(rowIndex);
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) });
  }
}

function ensurePosterSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('POSTER_SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);

  var ss = SpreadsheetApp.create('AFFICHE MEDIATHEQUE A3 - Historique');
  props.setProperty('POSTER_SHEET_ID', ss.getId());
  return ss;
}

function ensurePosterSheet() {
  var ss = ensurePosterSpreadsheet();
  var name = 'Historique';
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureSheetHeaders(sheet);
  return sheet;
}

function ensureSheetHeaders(sheet) {
  var headers = ['id','createdAt','updatedAt','title','titleLine2','subtitle','description','dates','infos','publicCible','siteUrl','pdfUrl','contactLine','showContactLine','footerLogo1Id','footerLogo2Id'];
  var existing = sheet.getLastColumn() >= headers.length
    ? sheet.getRange(1, 1, 1, headers.length).getValues()[0] : [];
  if (existing.length !== headers.length || existing.some(function (v, i) { return v !== headers[i]; })) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function ensurePosterFolder() {
  var folderName = 'AFFICHE MEDIATHEQUE A3';
  var folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function savePosterEntry(entryId, data, pdfUrl) {
  var sheet = ensurePosterSheet();
  var now = new Date();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var payload = {
    id: entryId || Utilities.getUuid(),
    createdAt: now,
    updatedAt: now,
    title: data.title || '',
    titleLine2: data.titleLine2 || '',
    subtitle: data.subtitle || '',
    description: data.description || '',
    dates: data.dates || '',
    infos: data.infos || '',
    publicCible: data.publicCible || '',
    siteUrl: data.siteUrl || '',
    pdfUrl: pdfUrl || '',
    contactLine: data.contactLine || '',
    showContactLine: (data.showContactLine === true) || (String(data.showContactLine).toLowerCase() === 'true')
  };

  var rowIndex = findEntryRow(sheet, payload.id);
  if (rowIndex) {
    payload.createdAt = sheet.getRange(rowIndex, 2).getValue();
    payload.updatedAt = now;
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([headers.map(function (h) { return payload[h]; })]);
  } else {
    sheet.appendRow(headers.map(function (h) { return payload[h]; }));
  }
}

function findEntryRow(sheet, id) {
  if (!id) return null;
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return null;
}

function buildPosterHtml(data) {
  var topLogo = data.topLogo || '';
  var bottomLogo = data.bottomLogo || '';
  var mainImage = data.mainImage || '';

  var footerLogo1 = data.footerLogo1 || '';
  var footerLogo2 = data.footerLogo2 || '';

  var title1 = (data.title || '').trim();
  var title2 = (data.titleLine2 || '').trim();
  var titleText = [title1, title2].filter(function (v) { return v; }).join('\n');

  var subtitle = (data.subtitle || '').trim();
  var description = (data.description || '').trim();
  var subtitleText = [subtitle, description].filter(function (v) { return v; }).join('\n');

  var datesText = (data.dates || '').trim();

  var infosText = (data.infos || '').trim();
  var contactLine = (data.contactLine || '').trim();
  var showContactLine = (data.showContactLine === true) || (String(data.showContactLine).toLowerCase() === 'true');

  var infosRender = infosText;
  if (showContactLine && contactLine) {
    infosRender = [infosText, contactLine].filter(function (v) { return v && String(v).trim(); }).join('\n');
  }

  var siteUrl = (data.siteUrl || '').trim();

  var publicLines = (data.publicCible || '')
    .split(/\r?\n/)
    .map(function (l) { return (l || '').trim(); })
    .filter(function (l) { return l; })
    .slice(0, 2);

  var hasTitleLine2 = !!title2;
  var isPublicSingleLine = publicLines.length <= 1;

  var ff = '';
  if (data.fontRias) ff += '@font-face{font-family:"Rias";src:url("' + data.fontRias + '");font-weight:700;font-style:normal}';
  if (data.fontRobotoBold) ff += '@font-face{font-family:"Roboto";src:url("' + data.fontRobotoBold + '");font-weight:700;font-style:normal}';
  if (data.fontVagRounded) ff += '@font-face{font-family:"VAG Rounded Std";src:url("' + data.fontVagRounded + '");font-weight:700;font-style:normal}';
  if (data.fontUniNeueBold) ff += '@font-face{font-family:"UniNeueBold";src:url("' + data.fontUniNeueBold + '");font-weight:700;font-style:normal}';

  var PURPLE = '#786fb2';
  var BLACK = '#231F20';

  var FRAME_LEFT = '12.855mm';
  var FRAME_TOP = '9.083mm';
  var FRAME_WIDTH = '274.290mm';
  var FRAME_HEIGHT = '379.447mm';

  var BAND_LEFT = '22.893mm';
  var BAND_TOP = hasTitleLine2 ? '239.600mm' : '249.766mm';
  var BAND_WIDTH = '236.345mm';
  var BAND_HEIGHT = hasTitleLine2 ? '108.174mm' : '98.007mm';

  var TEXT_LEFT = '29.014mm';
  var TEXT_TOP = hasTitleLine2 ? '264.834mm' : '275.768mm';
  var TEXT_WIDTH = '210.000mm';

  // Triangle blanc — positions originales, même direction que le petit triangle du cartouche
  var TRI_LEFT = isPublicSingleLine ? '203.849mm' : '209.327mm';
  var TRI_TOP  = isPublicSingleLine ? '213.074mm' : '225.292mm';
  var TRI_W    = isPublicSingleLine ? '68.387mm'  : '60.484mm';
  var TRI_H    = isPublicSingleLine ? '64.918mm'  : '57.416mm';
  TRI_TOP = (parseFloat(TRI_TOP) + 8).toFixed(3) + 'mm';

  var URLBAR_LEFT = '22.974mm';
  var URLBAR_W = '120.300mm';
  var URLBAR_H = '13.678mm';
  var URLTRI_LEFT = '49.492mm';
  var URLTRI_W = '8.227mm';
  var URLTRI_H = '8.585mm';
  var URLTEXT_LEFT = '32.838mm';

  var BOTTOM_WEDGE_LEFT = '12.855mm';
  var BOTTOM_WEDGE_TOP = '388.437mm';
  var BOTTOM_WEDGE_W = '274.290mm';
  var BOTTOM_WEDGE_H = '52.563mm';
  var BOTTOM_LOGO_LEFT = '209.500mm';
  var BOTTOM_LOGO_TOP = '368.500mm';
  var BOTTOM_LOGO_W = '77.645mm';
  var BOTTOM_LOGO_H = '50.000mm';

  var FOOTER_BULLET_LEFT = '14.15mm';
  var FOOTER_BULLET_TOP = '394.40mm';
  var FOOTER_TEXT_LEFT = '18.10mm';
  var FOOTER_TEXT_TOP = '393.60mm';

  var FOOTER_LOGO1_LEFT_MM = 18.10;
  var FOOTER_LOGO_TOP_MM = 401.20;
  var FOOTER_LOGO_W_MM = 16.00;
  var FOOTER_LOGO_H_MM = 16.00;
  var FOOTER_LOGO_GAP_MM = 2.50;

  var FOOTER_LOGO1_LEFT = FOOTER_LOGO1_LEFT_MM.toFixed(2) + 'mm';
  var FOOTER_LOGO2_LEFT = (FOOTER_LOGO1_LEFT_MM + FOOTER_LOGO_W_MM + FOOTER_LOGO_GAP_MM).toFixed(2) + 'mm';
  var FOOTER_LOGO_TOP = FOOTER_LOGO_TOP_MM.toFixed(2) + 'mm';
  var FOOTER_LOGO_W = FOOTER_LOGO_W_MM.toFixed(2) + 'mm';
  var FOOTER_LOGO_H = FOOTER_LOGO_H_MM.toFixed(2) + 'mm';

  var BAND_BOTTOM_MM = parseFloat(BAND_TOP) + parseFloat(BAND_HEIGHT);

  var URLBAR_TOP  = BAND_BOTTOM_MM.toFixed(3) + 'mm';
  var URLTEXT_TOP = (BAND_BOTTOM_MM + 1.567).toFixed(3) + 'mm';
  var URLTRI_TOP  = (BAND_BOTTOM_MM + 11.601).toFixed(3) + 'mm';

  var ORIG_BAND_TOP_MM   = parseFloat(BAND_TOP);
  var ORIG_TEXT_TOP_MM   = parseFloat(TEXT_TOP);
  var ORIG_TRI_TOP_MM    = parseFloat(TRI_TOP);
  var BASE_HEADROOM_MM   = ORIG_TEXT_TOP_MM - ORIG_BAND_TOP_MM;
  var BASE_TRI_OFFSET_MM = ORIG_TRI_TOP_MM  - ORIG_BAND_TOP_MM;

  var SUBTITLE_MT = 3.970;
  var DATES_MT    = 6.310;
  var INFOS_MT    = 1.000;

  var URL_SAFE_GAP_MM    = 2.0;
  var MAX_BOTTOM_GAP_MM  = 0.0;
  var PREF_BOTTOM_GAP_MM = 0.0;

  function estLines_(txt, charsPerLine) {
    if (!txt) return 0;
    return String(txt).split(/\r?\n/).reduce(function (sum, p) {
      p = (p || '').trim();
      return sum + Math.max(1, Math.ceil(((p.length || 1)) / charsPerLine));
    }, 0);
  }

  function computeNeedMM_() {
    var titleLines = estLines_(titleText, 24);
    var subLines   = estLines_(subtitleText, 60);
    var datesLines = estLines_(datesText, 34);
    var infosLines = estLines_(infosRender, 60);
    var needMM = 0;
    needMM += titleLines * 11.0;
    if (subtitleText) needMM += SUBTITLE_MT + subLines * 7.9;
    if (datesText)    needMM += DATES_MT    + datesLines * 8.7;
    if (infosRender)  needMM += INFOS_MT    + infosLines * 7.9;
    needMM += 2.0;
    return needMM;
  }

  var safeBottom   = parseFloat(URLBAR_TOP) - URL_SAFE_GAP_MM;
  var need         = computeNeedMM_();
  var targetBottom = Math.min(BAND_BOTTOM_MM - PREF_BOTTOM_GAP_MM, safeBottom);

  if ((targetBottom - need) < ORIG_TEXT_TOP_MM - 0.1) {
    SUBTITLE_MT = 2.000; DATES_MT = 0.000; INFOS_MT = 0.000;
    need         = computeNeedMM_();
    targetBottom = Math.min(BAND_BOTTOM_MM - PREF_BOTTOM_GAP_MM, safeBottom);
  }

  var minBottomAllowed = Math.min(BAND_BOTTOM_MM - MAX_BOTTOM_GAP_MM, safeBottom);
  if (targetBottom < minBottomAllowed) targetBottom = minBottomAllowed;

  var newTextTopMM = Math.min(ORIG_TEXT_TOP_MM, targetBottom - need);
  TEXT_TOP    = newTextTopMM.toFixed(3) + 'mm';

  var newBandTopMM = Math.min(ORIG_BAND_TOP_MM, newTextTopMM - BASE_HEADROOM_MM);
  BAND_TOP    = newBandTopMM.toFixed(3) + 'mm';
  BAND_HEIGHT = (BAND_BOTTOM_MM - newBandTopMM).toFixed(3) + 'mm';
  TRI_TOP     = (newBandTopMM + BASE_TRI_OFFSET_MM).toFixed(3) + 'mm';

  var SUBTITLE_MT_S = SUBTITLE_MT.toFixed(3) + 'mm';
  var DATES_MT_S    = DATES_MT.toFixed(3)    + 'mm';
  var INFOS_MT_S    = INFOS_MT.toFixed(3)    + 'mm';

  var audienceHtml = '';
  if (publicLines.length) {
    if (isPublicSingleLine) {
      audienceHtml = '<div class="audience-text single">' + escapeHtml(publicLines[0]) + '</div>';
    } else {
      audienceHtml =
        '<div class="audience-text double">' +
          '<div class="audience-small">' + escapeHtml(publicLines[0]) + '</div>' +
          '<div class="audience-big">'   + escapeHtml(publicLines[1]) + '</div>' +
        '</div>';
    }
  }

  var css =
    '@page{size:A3 portrait;margin:0;}' +
    'html,body{margin:0;padding:0;width:297mm;height:420mm;}' +
    'body{-webkit-print-color-adjust:exact;print-color-adjust:exact;background:#ffffff;}' +
    '.poster{position:relative;width:297mm;height:420mm;overflow:hidden;background:#ffffff;}' +
    '.frame{position:absolute;left:' + FRAME_LEFT + ';top:' + FRAME_TOP + ';width:' + FRAME_WIDTH + ';height:' + FRAME_HEIGHT + ';overflow:hidden;border:none;background:#fff;z-index:1;}' +
    '.frame img{width:100%;height:100%;object-fit:cover;}' +
    '.top-logo{position:absolute;left:38.779mm;top:10.562mm;width:54.379mm;height:86.729mm;object-fit:contain;z-index:5;}' +
    '.violet-band{position:absolute;left:' + BAND_LEFT + ';top:' + BAND_TOP + ';width:' + BAND_WIDTH + ';height:' + BAND_HEIGHT + ';background:' + PURPLE + ';z-index:3;' +
      'clip-path:polygon(0 0,100% 21%,100% 100%,0 100%);-webkit-clip-path:polygon(0 0,100% 21%,100% 100%,0 100%)}' +
    '.band-content{position:absolute;left:' + TEXT_LEFT + ';top:' + TEXT_TOP + ';width:' + TEXT_WIDTH + ';z-index:6;}' +
    '.title1{font-family:"Rias",cursive;font-size:37.716pt;font-weight:700;line-height:0.83;color:#ffffff;white-space:pre-line;}' +
    '.subtitle{margin-top:' + SUBTITLE_MT_S + ';font-family:"VAG Rounded Std",sans-serif;font-size:18.858pt;font-weight:700;line-height:1.18;color:' + BLACK + ';white-space:pre-line;}' +
    '.dates{margin-top:' + DATES_MT_S + ';font-family:"Roboto",sans-serif;font-size:23.131pt;font-weight:700;line-height:1.06;color:#ffffff;text-transform:uppercase;white-space:pre-line;}' +
    '.infos{margin-top:' + INFOS_MT_S + ';font-family:"VAG Rounded Std",sans-serif;font-size:18.858pt;font-weight:700;line-height:1.18;color:' + BLACK + ';white-space:pre-line;}' +
    // Grand triangle blanc : même direction que le petit triangle violet du cartouche
    '.audience-triangle{position:absolute;left:' + TRI_LEFT + ';top:' + TRI_TOP + ';width:' + TRI_W + ';height:' + TRI_H + ';background:#ffffff;z-index:7;' +
      'clip-path:polygon(0 50%,100% 0,100% 100%);-webkit-clip-path:polygon(0 50%,100% 0,100% 100%)}' +
    '.audience-text{position:absolute;left:' + TRI_LEFT + ';top:' + TRI_TOP + ';width:' + TRI_W + ';height:' + TRI_H + ';z-index:8;' +
      'display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;' +
      'font-family:"Roboto",sans-serif;font-weight:700;color:' + PURPLE + ';' +
      'transform:translateX(' + (parseFloat(TRI_W) / 6).toFixed(3) + 'mm);}' +
    '.audience-text.single{font-size:24pt;line-height:1.05;}' +
    '.audience-text.double{line-height:1.05;}' +
    '.audience-small{font-size:15.531pt;line-height:1.05;}' +
    '.audience-big{font-size:24pt;line-height:1.05;}' +
    '.url-bar{position:absolute;left:' + URLBAR_LEFT + ';top:' + URLBAR_TOP + ';width:' + URLBAR_W + ';height:' + URLBAR_H + ';background:' + BLACK + ';z-index:7;}' +
    '.small-triangle{position:absolute;left:' + URLTRI_LEFT + ';top:' + URLTRI_TOP + ';width:' + URLTRI_W + ';height:' + URLTRI_H + ';background:' + PURPLE + ';z-index:8;' +
      'clip-path:polygon(0 50%,100% 0,100% 100%);-webkit-clip-path:polygon(0 50%,100% 0,100% 100%)}' +
    '.url-text{position:absolute;left:' + URLTEXT_LEFT + ';top:' + URLTEXT_TOP + ';z-index:9;font-family:"UniNeueBold","Roboto",sans-serif;font-size:23.369pt;font-weight:700;line-height:1;color:' + PURPLE + ';white-space:nowrap;}' +
    '.bottom-wedge{position:absolute;left:' + BOTTOM_WEDGE_LEFT + ';top:' + BOTTOM_WEDGE_TOP + ';width:' + BOTTOM_WEDGE_W + ';height:' + BOTTOM_WEDGE_H + ';background:#fff;z-index:2;' +
      'clip-path:polygon(0 40.1%,71.2% 40.1%,71.2% 0,100% 0,100% 100%,0 100%);' +
      '-webkit-clip-path:polygon(0 40.1%,71.2% 40.1%,71.2% 0,100% 0,100% 100%,0 100%)}' +
    '.bottom-logo{position:absolute;left:' + BOTTOM_LOGO_LEFT + ';top:' + BOTTOM_LOGO_TOP + ';width:' + BOTTOM_LOGO_W + ';height:' + BOTTOM_LOGO_H + ';object-fit:contain;object-position:center;z-index:3;}' +
    '.footer-bullet{position:absolute;left:' + FOOTER_BULLET_LEFT + ';top:' + FOOTER_BULLET_TOP + ';width:0;height:0;' +
      'border-top:1.6mm solid transparent;border-bottom:1.6mm solid transparent;border-left:2.2mm solid ' + BLACK + ';z-index:6;}' +
    '.footer-text{position:absolute;left:' + FOOTER_TEXT_LEFT + ';top:' + FOOTER_TEXT_TOP + ';font-family:"Roboto",sans-serif;font-size:13.2pt;font-weight:700;color:' + BLACK + ';z-index:6;}' +
    '.footer-logo{position:absolute;z-index:6;object-fit:contain;object-position:left top;}' +
    '.footer-logo1{left:' + FOOTER_LOGO1_LEFT + ';top:' + FOOTER_LOGO_TOP + ';width:' + FOOTER_LOGO_W + ';height:' + FOOTER_LOGO_H + ';}' +
    '.footer-logo2{left:' + FOOTER_LOGO2_LEFT + ';top:' + FOOTER_LOGO_TOP + ';width:' + FOOTER_LOGO_W + ';height:' + FOOTER_LOGO_H + ';}';

  var html =
    '<!doctype html>' +
    '<html><head><meta charset="utf-8"><style>' + ff + css + '</style></head><body>' +
    '<div class="poster">' +
      '<div class="frame">' +
        (mainImage ? '<img src="' + mainImage + '">' : '') +
      '</div>' +
      (topLogo ? '<img class="top-logo" src="' + topLogo + '">' : '') +
      '<div class="violet-band"></div>' +
      '<div class="band-content">' +
        (titleText    ? '<div class="title1">'   + escapeHtml(titleText).replace(/\n/g,    '<br>') + '</div>' : '') +
        (subtitleText ? '<div class="subtitle">'  + escapeHtml(subtitleText).replace(/\n/g, '<br>') + '</div>' : '') +
        (datesText    ? '<div class="dates">'     + escapeHtml(datesText).replace(/\n/g,    '<br>') + '</div>' : '') +
        (infosRender  ? '<div class="infos">'     + escapeHtml(infosRender).replace(/\n/g,  '<br>') + '</div>' : '') +
      '</div>' +
      '<div class="audience-triangle"></div>' +
      audienceHtml +
      (siteUrl ? '<div class="url-bar"></div><div class="url-text">' + escapeHtml(siteUrl) + '</div>' : '') +
      '<div class="small-triangle"></div>' +
      '<div class="bottom-wedge"></div>' +
      (bottomLogo ? '<img class="bottom-logo" src="' + bottomLogo + '">' : '') +
      '<div class="footer-bullet"></div>' +
      '<div class="footer-text">golfedumorbihan-vannesagglomeration.bzh</div>' +
      (footerLogo1 ? '<img class="footer-logo footer-logo1" src="' + footerLogo1 + '">' : '') +
      (footerLogo2 ? '<img class="footer-logo footer-logo2" src="' + footerLogo2 + '">' : '') +
    '</div></body></html>';

  return html;
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#39;');
}

function formatDateForFilename(datesText) {
  var s = String(datesText || '').trim();
  if (!s) return '';

  var m = s.match(/(20\d{2})[-\/\.](\d{1,2})[-\/\.](\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);

  m = s.match(/(\d{1,2})[-\/\.](\d{1,2})[-\/\.](20\d{2})/);
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);

  var norm = s.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  m = norm.match(/(\d{1,2})\s*(JANVIER|FEVRIER|MARS|AVRIL|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE)(?:\s*(20\d{2}))?/);
  if (!m) return '';

  var day       = pad2_(m[1]);
  var monthName = m[2];
  var year      = m[3] || String(new Date().getFullYear());

  var months = {
    JANVIER:'01',FEVRIER:'02',MARS:'03',AVRIL:'04',MAI:'05',JUIN:'06',
    JUILLET:'07',AOUT:'08',SEPTEMBRE:'09',OCTOBRE:'10',NOVEMBRE:'11',DECEMBRE:'12'
  };
  var mm = months[monthName];
  if (!mm) return '';
  return year + '-' + mm + '-' + day;

  function pad2_(n) { n = String(n); return (n.length === 1) ? ('0' + n) : n; }
}

function sanitizeFilename(name) {
  return String(name || '')
    .trim()
    .toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 40);
}