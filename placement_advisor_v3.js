/**
 * ============================================================
 *  DEMAND GEN PLACEMENT ADVISOR
 *
 *  Author  : Nadim Mahmud Sizan
 *  LinkedIn: https://www.linkedin.com/in/nmsizan98/
 *
 *  ✅ Read-Only — never modifies Google Ads
 *  ✅ Uses built-in YouTube Advanced API
 *     (tick YouTube under Advanced APIs in Script Editor)
 *  ✅ Writes to your existing Google Sheet via URL
 *  ✅ Auto-pulls ALL active Demand Gen campaigns — no manual filter
 *
 * ============================================================
 *
 *  SETUP (2 steps only):
 *  1. Script Editor → Advanced APIs → tick "YouTube" → Save
 *  2. Paste your Google Sheet URL in SPREADSHEET_URL below
 *  3. Run → done!
 *
 * ============================================================
 */

/// ── CONFIGURATION — only edit this section ────────────────

var SPREADSHEET_URL = 'PASTE_YOUR_GOOGLE_SHEET_URL_HERE';

var LOOKBACK_WINDOW = 30; // days of data to pull

var MIN_IMPRESSIONS = 50; // ignore placements below this — keeps sheet clean and script fast

// Languages you allow ads to appear next to (video language check)
// Use 2-letter codes. Leave empty [] to skip language check.
var ALLOWED_LANGUAGES = ['en', 'EN'];

// TLDs you consider safe for website placements
var SAFE_TLDS = ['.com', '.edu', '.org', '.net', '.co.nz', '.nz', '.au', '.co.au', '.co.uk', '.uk'];

// Keywords that flag a channel/video as kids content (checked against title/name)
var KIDS_KEYWORDS = [
  'kids', 'children', 'nursery', 'cartoon', 'toddler', 'baby', 'preschool',
  'kindergarten', 'toy', 'peppa', 'cocomelon', 'sesame', 'bluey', 'paw patrol',
  'disney junior', 'nick jr', 'lullaby', 'rhymes', 'abc kids'
];

// Clickbait phrases in video titles
var CLICKBAIT_PHRASES = [
  'shocking', 'bombshell', 'you won\'t believe', 'must see', 'secret', 'revealed',
  'exposed', 'caught on camera', 'nightmare', 'scandal', 'jaw-dropping', 'insane',
  'unbelievable', 'must-watch', 'gone wrong', 'wtf', 'omg', 'don\'t miss'
];

// Political keywords in video titles
var POLITICAL_KEYWORDS = [
  'election', 'government', 'politics', 'president', 'congress', 'senate',
  'democracy', 'voting', 'votes', 'political', 'protest', 'liberal', 'conservative',
  'Trump', 'Biden', 'Harris', 'Putin', 'Pelosi'
];

// Spammy keywords in domain names
var SPAMMY_KEYWORDS = [
  'free', 'cheap', 'earn', 'win', 'cash', 'deal', 'bonus', 'prizes',
  'xxx', 'adult', 'sex', 'porn', 'escort',
  'bitcoin', 'crypto', 'forex', 'loan', 'payday', 'profit', 'trading'
];

/// ── END OF CONFIGURATION ──────────────────────────────────


// ── COLUMN HEADERS ─────────────────────────────────────────
var HEADERS = [
  'Placement',
  'Channel / Placement URL',
  'Type',
  'Campaign',
  'Ad Group',
  'Impressions',
  'Clicks',
  'Cost',
  'Conversions',
  'Cost / Conv.',
  'Conv. Rate',
  'CTR',
  'Avg. CPC',
  'Engagements',
  '👶 Kids / Made For Kids',
  '🌐 Language Issue',
  '🔞 Unsafe TLD/Domain',
  '📢 Clickbait / Political',
  '📱 Mobile App',
  '💡 Recommendation',
  '📝 Reasons',
];

// ── COLOURS ────────────────────────────────────────────────
var C = {
  headerBg    : '#1a1a2e',
  headerFont  : '#ffffff',
  subHeaderBg : '#2d2d44',
  exclude     : '#f4cccc',
  review      : '#fff2cc',
  keep        : '#d9ead3',
  altRow      : '#f8f8f8',
  totalRowBg  : '#e8e8e8',
  goodText    : '#38761d',
  reviewText  : '#7f6000',
  excludeText : '#990000',
};


// ══════════════════════════════════════════════════════════════
//  MAIN
// ══════════════════════════════════════════════════════════════
function main() {

  if (!SPREADSHEET_URL || SPREADSHEET_URL === 'PASTE_YOUR_GOOGLE_SHEET_URL_HERE') {
    throw new Error('❌ Please paste your Google Sheet URL into SPREADSHEET_URL at the top of the script.');
  }

  var account  = AdsApp.currentAccount();
  var accountName = account.getName();
  var tz       = account.getTimeZone();
  var currency = account.getCurrencyCode();
  var dateRange = getDateRange_(LOOKBACK_WINDOW, tz);

  Logger.log('▶ Demand Gen Placement Advisor starting…');
  Logger.log('🏢 Account: ' + accountName);
  Logger.log('📅 Period: ' + dateRange.startDate + ' → ' + dateRange.endDate);

  // Step 1: Fetch all placements
  var placements = fetchPlacements_(dateRange);
  Logger.log('📦 Placements fetched (>= ' + MIN_IMPRESSIONS + ' impressions): ' + placements.length);

  if (placements.length === 0) {
    Logger.log('⚠️ No placements found. Check your campaign has data in the date range.');
    return;
  }

  // Step 2: Analyse each placement
  var results = analysePlacements_(placements);
  Logger.log('✅ Analysis complete.');
  Logger.log('   → Exclude: ' + results.filter(function(r){return r.recommendation==='Exclude';}).length);
  Logger.log('   → Review:  ' + results.filter(function(r){return r.recommendation==='Review';}).length);
  Logger.log('   → Keep:    ' + results.filter(function(r){return r.recommendation==='Keep';}).length);

  // Step 3: Write to sheet
  writeToSheet_(results, accountName, currency, dateRange);

  Logger.log('📊 Done! Sheet: ' + SPREADSHEET_URL);
}


// ══════════════════════════════════════════════════════════════
//  STEP 1 — FETCH PLACEMENTS
//  Uses AdsApp.search() with GAQL for Demand Gen placements
//  but queries detail_placement_view for Demand Gen (DEMAND_GEN enum)
// ══════════════════════════════════════════════════════════════
function fetchPlacements_(dateRange) {
  var rows = [];

  var query =
    'SELECT ' +
      'detail_placement_view.display_name, ' +
      'detail_placement_view.placement, ' +
      'detail_placement_view.placement_type, ' +
      'campaign.name, ' +
      'ad_group.name, ' +
      'metrics.impressions, ' +
      'metrics.clicks, ' +
      'metrics.cost_micros, ' +
      'metrics.conversions, ' +
      'metrics.cost_per_conversion, ' +
      'metrics.ctr, ' +
      'metrics.average_cpc, ' +
      'metrics.engagements ' +
    'FROM detail_placement_view ' +
    'WHERE campaign.advertising_channel_type = DEMAND_GEN ' +
    'AND campaign.status = ENABLED ' +
    'AND metrics.impressions > ' + MIN_IMPRESSIONS + ' ' +
    'AND segments.date BETWEEN \'' + dateRange.startDate + '\' AND \'' + dateRange.endDate + '\' ' +
    'ORDER BY metrics.impressions DESC';

  Logger.log('🔍 Running GAQL query...');

  try {
    var result = AdsApp.search(query);

    while (result.hasNext()) {
      var row = result.next();

      // ── Access row fields from AdsApp.search() result ──
      var dpv     = row.detailPlacementView;
      var metrics = row.metrics;
      var campaign = row.campaign;
      var adGroup  = row.adGroup;

      var costMicros    = parseFloat(metrics.costMicros        || 0);
      var costPerConvMicros = parseFloat(metrics.costPerConversion || 0);
      var avgCpcMicros  = parseFloat(metrics.averageCpc          || 0);
      var clicks        = parseInt(metrics.clicks                || 0, 10);
      var conversions   = parseFloat(metrics.conversions         || 0);

      // Build full URL from placement ID and type
      var rawPlacement = dpv.placement || '';
      var pType = (dpv.placementType || '').toUpperCase();
      var fullUrl = rawPlacement;
      if (pType === 'YOUTUBE_VIDEO' && rawPlacement && rawPlacement.indexOf('http') === -1) {
        fullUrl = 'https://www.youtube.com/watch?v=' + rawPlacement;
      } else if (pType === 'YOUTUBE_CHANNEL' && rawPlacement && rawPlacement.indexOf('http') === -1) {
        fullUrl = 'https://www.youtube.com/channel/' + rawPlacement;
      }

      rows.push({
        placement     : dpv.displayName  || dpv.placement || '',
        placementUrl  : fullUrl,
        placementType : dpv.placementType || '',
        campaign      : campaign.name    || '',
        adGroup       : adGroup.name     || '',
        impressions   : parseInt(metrics.impressions || 0, 10),
        clicks        : clicks,
        cost          : costMicros / 1000000,
        conversions   : conversions,
        costPerConv   : costPerConvMicros > 0 ? costPerConvMicros / 1000000 : 0,
        convRate      : clicks > 0 ? (conversions / clicks) * 100 : 0,
        ctr           : parseFloat(metrics.ctr || 0) * 100,
        avgCpc        : avgCpcMicros / 1000000,
        engagements   : parseInt(metrics.engagements || 0, 10),
      });
    }

    Logger.log('✅ Query returned ' + rows.length + ' rows');

  } catch (e) {
    Logger.log('❌ Query error: ' + e.message);
    Logger.log('Query was: ' + query);
  }

  return rows;
}


// ══════════════════════════════════════════════════════════════
//  STEP 2 — ANALYSE PLACEMENTS
//  Analyse each placement for exclusion signals
// ══════════════════════════════════════════════════════════════
function analysePlacements_(placements) {
  var results = [];
  var total = placements.length;

  placements.forEach(function(p, idx) {

    if (idx % 200 === 0) {
      Logger.log('  Analysing... ' + idx + ' / ' + total);
    }

    var flags = {
      isKids      : false,
      langIssue   : false,
      unsafeDomain: false,
      clickbait   : false,
      mobileApp   : false,
    };
    var reasons = [];

    var type = (p.placementType || '').toUpperCase();

    // ── MOBILE APP — flag all mobile app placements ────────
    if (type === 'MOBILE_APPLICATION' || type === 'MOBILE_APP') {
      flags.mobileApp = true;
      reasons.push('Mobile app placement — low conversion value');
    }

    // ── YOUTUBE CHANNEL / VIDEO ──────────────────────────
    if (type === 'YOUTUBE_CHANNEL' || type === 'YOUTUBE_VIDEO') {

      var name = (p.placement || '').toLowerCase();

      // Kids keyword check on channel name
      KIDS_KEYWORDS.forEach(function(kw) {
        if (name.indexOf(kw.toLowerCase()) !== -1) {
          flags.isKids = true;
          reasons.push('Kids keyword in name: "' + kw + '"');
        }
      });

      // YouTube API enrichment for videos (only if API available)
      if (type === 'YOUTUBE_VIDEO' && typeof YouTube !== 'undefined') {
        var videoId = p.placementUrl;
        if (videoId) {
          var videoCheck = checkYouTubeVideo_(videoId, p.placement);
          if (videoCheck.madeForKids) {
            flags.isKids = true;
            reasons.push('YouTube API: madeForKids = true');
          }
          if (videoCheck.langIssue) {
            flags.langIssue = true;
            reasons.push('Language not in allowed list: ' + videoCheck.lang);
          }
          if (videoCheck.clickbait) {
            flags.clickbait = true;
            reasons.push('Title: ' + videoCheck.clickbaitReason);
          }
        }
      }

      // Clickbait check on display name (fast, no API needed)
      var displayLower = p.placement.toLowerCase();
      CLICKBAIT_PHRASES.forEach(function(phrase) {
        if (displayLower.indexOf(phrase.toLowerCase()) !== -1) {
          flags.clickbait = true;
          reasons.push('Clickbait phrase in name: "' + phrase + '"');
        }
      });
      POLITICAL_KEYWORDS.forEach(function(kw) {
        if (displayLower.indexOf(kw.toLowerCase()) !== -1) {
          flags.clickbait = true;
          reasons.push('Political keyword in name: "' + kw + '"');
        }
      });
    }

    // ── WEBSITE ─────────────────────────────────────────
    if (type === 'WEBSITE') {
      var targetUrl = p.placementUrl || p.placement;

      // TLD check
      var tldOk = SAFE_TLDS.some(function(tld) {
        return targetUrl.toLowerCase().indexOf(tld) !== -1;
      });
      if (!tldOk) {
        flags.unsafeDomain = true;
        reasons.push('Unsafe/unknown TLD: ' + targetUrl);
      }

      // Domain quality checks
      var sldCheck = checkSLD_(targetUrl);
      if (sldCheck.isQuestionable) {
        flags.unsafeDomain = true;
        reasons.push(sldCheck.reason);
      }

      // Spammy keyword in domain
      SPAMMY_KEYWORDS.forEach(function(kw) {
        if (targetUrl.toLowerCase().indexOf(kw.toLowerCase()) !== -1) {
          flags.unsafeDomain = true;
          reasons.push('Spammy keyword in domain: "' + kw + '"');
        }
      });
    }

    // ── RECOMMENDATION LOGIC ─────────────────────────────
    var rec;
    if (flags.isKids || flags.mobileApp || flags.unsafeDomain) {
      rec = 'Exclude';
    } else if (flags.langIssue || flags.clickbait) {
      rec = 'Review';
    } else if (p.impressions >= MIN_IMPRESSIONS && p.conversions === 0 && p.cost > 1) {
      rec = 'Review';
      reasons.push('Spend with no conversions: ' + fmtMoney_(p.cost));
    } else {
      rec = 'Keep';
      if (p.conversions > 0) {
        reasons.push('Has conversions');
      } else {
        reasons.push('No negative signals');
      }
    }

    results.push({
      placement     : p.placement,
      placementUrl  : p.placementUrl || '',
      placementType : formatType_(p.placementType),
      campaign      : p.campaign,
      adGroup       : p.adGroup,
      impressions   : p.impressions,
      clicks        : p.clicks,
      cost          : p.cost,
      conversions   : p.conversions,
      costPerConv   : p.costPerConv,
      convRate      : p.convRate,
      ctr           : p.ctr,
      avgCpc        : p.avgCpc,
      engagements   : p.engagements,
      isKids        : flags.isKids        ? 'YES ⚠️' : 'No',
      langIssue     : flags.langIssue     ? 'YES ⚠️' : 'No',
      unsafeDomain  : flags.unsafeDomain  ? 'YES ⚠️' : 'No',
      clickbait     : flags.clickbait     ? 'YES ⚠️' : 'No',
      mobileApp     : flags.mobileApp     ? 'YES ⚠️' : 'No',
      recommendation: rec,
      reasons       : reasons.join(' | '),
    });
  });

  return results;
}


// ══════════════════════════════════════════════════════════════
//  YOUTUBE API CHECK — video-level signals
// ══════════════════════════════════════════════════════════════
function checkYouTubeVideo_(videoId, displayName) {
  var result = { madeForKids: false, langIssue: false, lang: '', clickbait: false, clickbaitReason: '' };

  try {
    var response = YouTube.Videos.list('snippet,status', { id: videoId });

    if (!response.items || response.items.length === 0) return result;

    var item    = response.items[0];
    var snippet = item.snippet || {};
    var status  = item.status  || {};

    // Kids check
    if (status.madeForKids === true) {
      result.madeForKids = true;
    }

    // Language check
    if (ALLOWED_LANGUAGES.length > 0) {
      var lang = snippet.defaultLanguage || snippet.defaultAudioLanguage || '';
      if (lang && lang !== '') {
        var langAllowed = ALLOWED_LANGUAGES.some(function(l) {
          return lang.toLowerCase().indexOf(l.toLowerCase()) !== -1;
        });
        if (!langAllowed) {
          result.langIssue = true;
          result.lang = lang;
        }
      }
    }

    // Clickbait title check
    var title = (snippet.title || '').toLowerCase();
    CLICKBAIT_PHRASES.forEach(function(phrase) {
      if (title.indexOf(phrase.toLowerCase()) !== -1) {
        result.clickbait = true;
        result.clickbaitReason = 'Clickbait: "' + phrase + '"';
      }
    });
    POLITICAL_KEYWORDS.forEach(function(kw) {
      if (title.indexOf(kw.toLowerCase()) !== -1) {
        result.clickbait = true;
        result.clickbaitReason = 'Political: "' + kw + '"';
      }
    });

    // Excessive punctuation
    if (/[!?]{2,}/.test(snippet.title || '')) {
      result.clickbait = true;
      result.clickbaitReason = 'Excessive punctuation in title';
    }

    // Monetary amounts in title
    if (/[\$\€\£]\d+/.test(snippet.title || '')) {
      result.clickbait = true;
      result.clickbaitReason = 'Monetary amount in title';
    }

  } catch (e) {
    // YouTube API failed for this video — skip silently
  }

  return result;
}


// ══════════════════════════════════════════════════════════════
//  DOMAIN QUALITY CHECK — SLD and TLD signals
// ══════════════════════════════════════════════════════════════
function checkSLD_(targetUrl) {
  try {
    var parts = targetUrl.replace(/https?:\/\//, '').split('.');
    var sld = parts.length > 1 ? parts[parts.length - 2] : parts[0];
    sld = sld.split('/')[0];

    if (sld.length < 2 || sld.length > 30) {
      return { isQuestionable: true, reason: 'Suspicious domain length: ' + sld };
    }
    if (sld.includes('-')) {
      return { isQuestionable: true, reason: 'Domain contains hyphen: ' + sld };
    }
    if (/[0-9]{4,}/.test(sld)) {
      return { isQuestionable: true, reason: 'Domain has 4+ numbers in a row: ' + sld };
    }
    if (/(.)\1{3,}/.test(sld)) {
      return { isQuestionable: true, reason: 'Domain has repeated characters: ' + sld };
    }
    if (sld.indexOf('xn--') !== -1) {
      return { isQuestionable: true, reason: 'Domain uses punycode (internationalized): ' + sld };
    }
  } catch (e) { /* skip */ }

  return { isQuestionable: false };
}


// ══════════════════════════════════════════════════════════════
//  WRITE TO GOOGLE SHEET
// ══════════════════════════════════════════════════════════════
function writeToSheet_(results, accountName, currency, dateRange) {

  var ss;
  try {
    ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  } catch (e) {
    throw new Error('❌ Cannot open sheet. Check SPREADSHEET_URL. Error: ' + e.message);
  }

  var tabName = accountName + ' - Placement Advisor';
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    Logger.log('📄 Created new tab: ' + tabName);
  } else {
    sheet.setName(tabName); // ensure name is correct even if it was renamed
    Logger.log('📄 Using existing tab: ' + tabName);
  }
  sheet.clearContents();
  sheet.clearFormats();
  // Also rename the spreadsheet itself
  try { ss.rename(tabName); } catch(e) { /* ignore if no permission */ }

  var col = HEADERS.length;
  var now = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'dd MMM yyyy HH:mm');

  // Row 1 — Title
  var r1 = sheet.getRange(1, 1, 1, col).merge();
  r1.setValue('📊  DEMAND GEN PLACEMENT ADVISOR   |   ' + accountName)
    .setBackground(C.headerBg).setFontColor(C.headerFont)
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  // Row 2 — Subtitle
  var r2 = sheet.getRange(2, 1, 1, col).merge();
  r2.setValue(
    '🗓  Last ' + LOOKBACK_WINDOW + ' days  (' + dateRange.startDate + '  →  ' + dateRange.endDate + ')' +
    '      🕐  Generated: ' + now + '      ⚠️ Min impressions: ' + MIN_IMPRESSIONS
  ).setBackground(C.subHeaderBg).setFontColor('#cccccc')
   .setFontSize(10).setHorizontalAlignment('center').setFontStyle('italic');
  sheet.setRowHeight(2, 26);

  // Row 3 — Legend
  sheet.setRowHeight(3, 26);
  legendCell_(sheet, 3, 1,  5, '✅ Keep — no issues detected',          C.keep,    C.goodText);
  legendCell_(sheet, 3, 6,  5, '⚠️ Review — investigate further',       C.review,  C.reviewText);
  legendCell_(sheet, 3, 11, col - 10, '❌ Exclude — apply manually in Google Ads UI → Placements → Exclusions', C.exclude, C.excludeText);

  // Row 4 — Headers
  var hRange = sheet.getRange(4, 1, 1, col);
  hRange.setValues([HEADERS])
    .setBackground(C.headerBg).setFontColor(C.headerFont)
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(4, 46);

  // Row 5 — Totals
  var tot = calcTotals_(results);
  sheet.getRange(5, 1, 1, col).setValues([[
    'TOTAL  (' + results.length + ' placements)',
    '', '', '', '',
    tot.impressions, tot.clicks, tot.cost, tot.conversions,
    tot.costPerConv, tot.convRate, tot.ctr, tot.avgCpc, tot.engagements,
    '', '', '', '', '', '', ''
  ]]).setBackground(C.totalRowBg).setFontWeight('bold');

  // Rows 6+ — Data
  var dataStart = 6;
  if (results.length > 0) {

    var dataValues = results.map(function(r) {
      return [
        r.placement, r.placementUrl, r.placementType, r.campaign, r.adGroup,
        r.impressions, r.clicks, r.cost, r.conversions,
        r.costPerConv, r.convRate, r.ctr, r.avgCpc, r.engagements,
        r.isKids, r.langIssue, r.unsafeDomain, r.clickbait, r.mobileApp,
        r.recommendation, r.reasons,
      ];
    });

    sheet.getRange(dataStart, 1, dataValues.length, col).setValues(dataValues);

    // ── Batch number formats ──────────────────────────────
    var nr = dataValues.length;
    var curr = '"$"#,##0.00';
    var pct  = '0.00"%"';
    var num  = '#,##0';

    sheet.getRange(dataStart, 6,  nr).setNumberFormat(num);   // Impressions
    sheet.getRange(dataStart, 7,  nr).setNumberFormat(num);   // Clicks
    sheet.getRange(dataStart, 8,  nr).setNumberFormat(curr);  // Cost
    sheet.getRange(dataStart, 9,  nr).setNumberFormat('0.00');// Conversions
    sheet.getRange(dataStart, 10, nr).setNumberFormat(curr);  // Cost/Conv
    sheet.getRange(dataStart, 11, nr).setNumberFormat(pct);   // Conv Rate
    sheet.getRange(dataStart, 12, nr).setNumberFormat(pct);   // CTR
    sheet.getRange(dataStart, 13, nr).setNumberFormat(curr);  // Avg CPC
    sheet.getRange(dataStart, 14, nr).setNumberFormat(num);   // Engagements

    // Also format totals row
    sheet.getRange(5, 6).setNumberFormat(num);
    sheet.getRange(5, 7).setNumberFormat(num);
    sheet.getRange(5, 8).setNumberFormat(curr);
    sheet.getRange(5, 9).setNumberFormat('0.00');
    sheet.getRange(5, 10).setNumberFormat(curr);
    sheet.getRange(5, 11).setNumberFormat(pct);
    sheet.getRange(5, 12).setNumberFormat(pct);
    sheet.getRange(5, 13).setNumberFormat(curr);
    sheet.getRange(5, 14).setNumberFormat(num);

    // ── Batch row colouring ───────────────────────────────
    for (var i = 0; i < dataValues.length; i++) {
      var rec = dataValues[i][19]; // Recommendation column (shifted +1 for URL col)
      var bg  = rec === 'Exclude' ? C.exclude :
                rec === 'Review'  ? C.review  :
                (i % 2 === 0)     ? '#ffffff'  : C.altRow;
      sheet.getRange(dataStart + i, 1, 1, col).setBackground(bg);

      // Recommendation cell colour
      var fc = rec === 'Exclude' ? C.excludeText :
               rec === 'Review'  ? C.reviewText  : C.goodText;
      sheet.getRange(dataStart + i, 20).setFontColor(fc).setFontWeight('bold');
    }
  }

  // Freeze & filter
  sheet.setFrozenRows(4);
  // Note: setFrozenColumns removed — conflicts with merged header cells
  // Remove existing filter before creating new one
  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  // Remove existing filter before creating a new one
  var existingFilter = sheet.getFilter();
  if (existingFilter) { existingFilter.remove(); }
  sheet.getRange(4, 1, results.length + 2, col).createFilter();

  // Column widths
  var widths = [200, 280, 110, 180, 130, 90, 70, 80, 90, 90, 85, 70, 80, 90, 110, 100, 110, 120, 90, 120, 320];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  // Footer instructions
  writeFooter_(sheet, results, dataStart + results.length + 2, col);

  Logger.log('📝 Written ' + results.length + ' rows to: ' + tabName);
}


// ══════════════════════════════════════════════════════════════
//  FOOTER — SUMMARY + HOW TO EXCLUDE
// ══════════════════════════════════════════════════════════════
function writeFooter_(sheet, results, startRow, col) {
  var keep    = results.filter(function(r) { return r.recommendation === 'Keep';    }).length;
  var review  = results.filter(function(r) { return r.recommendation === 'Review';  }).length;
  var exclude = results.filter(function(r) { return r.recommendation === 'Exclude'; }).length;
  var kids    = results.filter(function(r) { return r.isKids === 'YES ⚠️';           }).length;
  var wasted  = results.filter(function(r) { return r.recommendation !== 'Keep';    })
                       .reduce(function(s, r) { return s + r.cost; }, 0);

  // Build empty row with exact col count
  function emptyRow() {
    var r = [];
    for (var i = 0; i < col; i++) r.push('');
    return r;
  }

  // Build a data row: first cell = label, second = value, rest empty
  function dataRow(label, value) {
    var r = emptyRow();
    r[0] = label;
    r[1] = value;
    return r;
  }

  // Build a full-width label row
  function fullRow(label) {
    var r = emptyRow();
    r[0] = label;
    return r;
  }

  var summaryData = [
    emptyRow(),
    fullRow('📊  SUMMARY'),
    dataRow('Total Placements',    results.length),
    dataRow('✅ Keep',             keep),
    dataRow('⚠️ Review',          review),
    dataRow('❌ Exclude',         exclude),
    dataRow('👶 Kids Detected',    kids),
    dataRow('💰 Spend on Flagged', fmtMoney_(wasted)),
    emptyRow(),
    fullRow('📌  HOW TO APPLY EXCLUSIONS IN GOOGLE ADS (this script is READ-ONLY — it never changes Google Ads)'),
    fullRow('Step 1 →  Filter column T (Recommendation) → select "Exclude" or "Review"'),
    fullRow('Step 2 →  Copy placement URLs from column B'),
    fullRow('Step 3 →  Google Ads → your Demand Gen campaign → Placements tab → Exclusions → Add'),
    fullRow('Step 4 →  Paste placement URLs → Save'),
  ];

  sheet.getRange(startRow, 1, summaryData.length, col).setValues(summaryData);

  // Style summary header
  sheet.getRange(startRow + 1, 1, 1, 4).merge()
    .setBackground(C.headerBg).setFontColor('#ffffff').setFontWeight('bold').setFontSize(12);

  // Style stat rows
  [[2, C.goodText], [3, C.reviewText], [4, C.excludeText], [5, C.excludeText]].forEach(function(pair) {
    sheet.getRange(startRow + pair[0], 2).setFontColor(pair[1]).setFontWeight('bold');
  });

  // Style instructions header
  sheet.getRange(startRow + 9, 1, 1, col).merge()
    .setBackground('#efefef').setFontWeight('bold').setFontSize(11);
}


// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════
function getDateRange_(lookbackWindow, tz) {
  var today = new Date();
  var start = new Date(today.getTime() - lookbackWindow * 24 * 60 * 60 * 1000);
  var yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
  return {
    startDate : Utilities.formatDate(start,     tz, 'yyyy-MM-dd'),
    endDate   : Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd'),
  };
}

function formatType_(type) {
  var map = {
    'YOUTUBE_CHANNEL'   : 'YouTube channel',
    'YOUTUBE_VIDEO'     : 'YouTube video',
    'WEBSITE'           : 'Website',
    'MOBILE_APPLICATION': 'Mobile app',
    'MOBILE_APP'        : 'Mobile app',
  };
  return map[(type || '').toUpperCase()] || type || 'Unknown';
}

function fmtMoney_(val) {
  return '$' + (val || 0).toFixed(2);
}

function calcTotals_(results) {
  var t = { impressions:0, clicks:0, cost:0, conversions:0, costPerConv:0,
            convRate:0, ctr:0, avgCpc:0, engagements:0 };
  results.forEach(function(r) {
    t.impressions += r.impressions;
    t.clicks      += r.clicks;
    t.cost        += r.cost;
    t.conversions += r.conversions;
    t.engagements += r.engagements;
  });
  t.costPerConv = t.conversions > 0  ? t.cost / t.conversions : 0;
  t.convRate    = t.clicks      > 0  ? (t.conversions / t.clicks) * 100 : 0;
  t.ctr         = t.impressions > 0  ? (t.clicks / t.impressions) * 100 : 0;
  t.avgCpc      = t.clicks      > 0  ? t.cost / t.clicks : 0;
  return t;
}

function legendCell_(sheet, row, startCol, span, text, bg, fc) {
  sheet.getRange(row, startCol, 1, span).merge()
    .setValue(text).setBackground(bg).setFontColor(fc)
    .setFontWeight('bold').setFontSize(9)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

