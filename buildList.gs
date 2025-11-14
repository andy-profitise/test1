/***** CONFIG *****/
const SHEET_BUYERS_L1M       = 'Buyers L1M';
const SHEET_BUYERS_L6M       = 'Buyers L6M';
const SHEET_AFFILIATES_L1M   = 'Affiliates L1M';
const SHEET_AFFILIATES_L6M   = 'Affiliates L6M';
const SHEET_MON_BUYERS       = 'buyers monday.com';
const SHEET_MON_AFFILIATES   = 'affiliates monday.com';

const SHEET_OUT              = 'List';
const SHEET_SETTINGS         = 'Settings';

// Header candidates to auto-detect vendor and TTL
const VENDOR_HEADER_CANDIDATES = ['Buyer', 'Affiliate', 'Publisher', 'Name', 'Vendor'];
const TTL_HEADER_CANDIDATES    = ['TTL , USD', 'TTL, USD', 'TTL USD'];

// monday.com special columns
const ALIAS_COL_INDEX = 15;       // Column P (1-based 16) -> 0-based 15
const BUYERS_STATUS_INDEX = 17;   // Column R (1-based 18) -> 0-based 17
const AFFILIATES_STATUS_INDEX = 18; // Column S (1-based 19) -> 0-based 18

// Status priority used when TTL = 0
const STATUS_PRIORITY = [
  'live',
  'onboarding',
  'paused',
  'preonboarding',
  'early talks',
  'top 500 remodelers',  // placed after early talks
  'other',
  'dead'
];
const STATUS_RANK = STATUS_PRIORITY.reduce((m, s, i) => (m[s] = i, m), {});

/**
 * NOTE: Menu is defined in BattleStation.gs onOpen() function
 * to avoid collision with multiple onOpen() declarations
 */

/**
 * Build unified "List" with:
 * - HOT ZONE (top): Vendors with recent emails (label:00.received last 7 days)
 * - NORMAL ZONE: All other vendors
 * - Within each zone: Primary sort by TTL DESC, then type (Buyers > Affiliates), then A-Z
 * - TTL = 0: metrics in order (L6M buyers, L6M aff, L1M buyers, L1M aff), then monday.com by status
 * Output columns: Vendor | TTL , USD | Source | Status
 */
function buildVendorList() {
  const ss = SpreadsheetApp.getActive();

  console.log('=== BUILD VENDOR LIST START ===');

  // Metric sheets required
  const shBuyL1 = mustGetSheet_(ss, SHEET_BUYERS_L1M);
  const shBuyL6 = mustGetSheet_(ss, SHEET_BUYERS_L6M);
  const shAffL1 = mustGetSheet_(ss, SHEET_AFFILIATES_L1M);
  const shAffL6 = mustGetSheet_(ss, SHEET_AFFILIATES_L6M);

  // monday.com sheets optional
  const shMonB  = ss.getSheetByName(SHEET_MON_BUYERS);
  const shMonA  = ss.getSheetByName(SHEET_MON_AFFILIATES);

  console.log('Monday.com sheets found:', {
    buyers: shMonB ? 'YES' : 'NO',
    affiliates: shMonA ? 'YES' : 'NO'
  });

  // Blacklist from Settings column E
  const blacklist = readBlacklist_(ss);
  console.log('Blacklist size:', blacklist.size);

  // Buyers L1M
  const buyersL1M = readMetricSheet_(shBuyL1, 'Buyer', 'Buyers L1M', blacklist);
  const buyersL1MSet = new Set(buyersL1M.map(r => r.name.toLowerCase()));
  console.log('Buyers L1M:', buyersL1M.length, 'items');

  // Buyers L6M (only names not already in Buyers L1M)
  const buyersL6MAll = readMetricSheet_(shBuyL6, 'Buyer', 'Buyers L6M', blacklist);
  const buyersL6M = buyersL6MAll.filter(r => !buyersL1MSet.has(r.name.toLowerCase()));
  console.log('Buyers L6M:', buyersL6M.length, 'items (after L1M dedup)');

  // buyers monday.com (only new after Buyers L1M + Buyers L6M)
  const buyersExisting = new Set([...buyersL1MSet, ...buyersL6M.map(r => r.name.toLowerCase())]);
  const buyersMon = shMonB ? readMondaySheet_(shMonB, 'Buyer', 'buyers monday.com', blacklist, buyersExisting) : [];
  console.log('Buyers monday.com:', buyersMon.length, 'items');

  // Affiliates L1M
  const affL1M = readMetricSheet_(shAffL1, 'Affiliate', 'Affiliates L1M', blacklist);
  const affL1MSet = new Set(affL1M.map(r => r.name.toLowerCase()));
  console.log('Affiliates L1M:', affL1M.length, 'items');

  // Affiliates L6M (only names not already in Affiliates L1M)
  const affL6MAll = readMetricSheet_(shAffL6, 'Affiliate', 'Affiliates L6M', blacklist);
  const affL6M = affL6MAll.filter(r => !affL1MSet.has(r.name.toLowerCase()));
  console.log('Affiliates L6M:', affL6M.length, 'items (after L1M dedup)');

  // affiliates monday.com (only new after Affiliates L1M + Affiliates L6M)
  const affExisting = new Set([...affL1MSet, ...affL6M.map(r => r.name.toLowerCase())]);
  const affMon = shMonA ? readMondaySheet_(shMonA, 'Affiliate', 'affiliates monday.com', blacklist, affExisting) : [];
  console.log('Affiliates monday.com:', affMon.length, 'items');

  // Split by TTL > 0 vs TTL = 0
  const gt0 = [];
  const z_buyL6 = [], z_affL6 = [], z_buyL1 = [], z_affL1 = [];
  const z_mon = [];

  const pushByTtl = (arr, zeroTarget) => {
    for (const r of arr) {
      if ((r.ttl || 0) > 0) gt0.push(r);
      else zeroTarget.push(r);
    }
  };

  pushByTtl(buyersL6M, z_buyL6);
  pushByTtl(affL6M,   z_affL6);
  pushByTtl(buyersL1M, z_buyL1);
  pushByTtl(affL1M,   z_affL1);

  for (const r of buyersMon) ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);
  for (const r of affMon)    ((r.ttl || 0) > 0 ? gt0 : z_mon).push(r);

  console.log('Split results:', {
    'gt0': gt0.length,
    'z_buyL6': z_buyL6.length,
    'z_affL6': z_affL6.length,
    'z_buyL1': z_buyL1.length,
    'z_affL1': z_affL1.length,
    'z_mon': z_mon.length
  });

  // Sort TTL > 0
  gt0.sort((a, b) => {
    const ttlDiff = (b.ttl || 0) - (a.ttl || 0);
    if (ttlDiff !== 0) return ttlDiff;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Zero TTL metric groups sort alphabetically within each group
  const alpha = (a, b) => String(a.name).localeCompare(String(b.name));
  z_buyL6.sort(alpha);
  z_affL6.sort(alpha);
  z_buyL1.sort(alpha);
  z_affL1.sort(alpha);

  // Zero TTL monday.com sorted by status priority, then Buyer>Affiliate, then A→Z
  z_mon.sort((a, b) => {
    const sA = STATUS_RANK[String(a.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    const sB = STATUS_RANK[String(b.status || '').toLowerCase()] ?? STATUS_RANK['other'];
    if (sA !== sB) return sA - sB;
    const rankA = (a.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    const rankB = (b.type || '').toLowerCase().startsWith('buyer') ? 0 : 1;
    if (rankA !== rankB) return rankA - rankB;
    return String(a.name).localeCompare(String(b.name));
  });

  // Assembly before hot zone detection
  const all = [
    ...gt0,
    ...z_buyL6,
    ...z_affL6,
    ...z_buyL1,
    ...z_affL1,
    ...z_mon
  ];

  console.log('Total items before hot zone:', all.length);

  // Build status lookup maps
  const statusMaps = buildStatusMaps_(shMonB, shMonA);
  console.log('Status maps built:', {
    buyers: statusMaps.buyers.size,
    affiliates: statusMaps.affiliates.size
  });

  // Lookup and assign status to each vendor
  for (const r of all) {
    r.status = lookupStatus_(r.name, r.type, statusMaps);
  }

  // HOT ZONE: Detect vendors with recent emails
  console.log('Detecting hot vendors with recent emails...');
  const hotVendorSet = getHotVendorsFromGmail_(all);
  console.log('Hot vendors found:', hotVendorSet.size);

  // Split into hot and normal zones
  const hotZone = [];
  const normalZone = [];

  for (const r of all) {
    if (hotVendorSet.has(r.name.toLowerCase())) {
      hotZone.push(r);
    } else {
      normalZone.push(r);
    }
  }

  console.log('Hot zone:', hotZone.length, 'Normal zone:', normalZone.length);

  // Final assembly: HOT at top, then NORMAL
  const finalList = [...hotZone, ...normalZone];

  console.log('Total items for output:', finalList.length);

  // Write to List (clear A-E to prevent hanging data from previous runs)
  const shOut = ensureSheet_(ss, SHEET_OUT);
  const lastRow = shOut.getLastRow();
  if (lastRow > 0) shOut.getRange(1, 1, lastRow, 5).clearContent();

  shOut.getRange(1, 1, 1, 4).setValues([['Vendor', 'TTL , USD', 'Source', 'Status']]);
  if (finalList.length) {
    const data = finalList.map(r => [r.name, r.ttl || 0, r.source, r.status || '']);
    shOut.getRange(2, 1, data.length, 4).setValues(data);
    shOut.getRange(2, 2, data.length, 1).setNumberFormat('$#,##0.00;($#,##0.00)');
  }
  if (lastRow > finalList.length + 1) {
    shOut.getRange(finalList.length + 2, 1, lastRow - (finalList.length + 1), 5).clearContent();
  }
  shOut.autoResizeColumns(1, 4);

  // Diagnostics
  const counts = {
    '🔥 HOT (recent emails)': hotZone.length,
    'Buyers L1M': buyersL1M.length,
    'Buyers L6M': buyersL6M.length,
    'buyers monday.com': buyersMon.length,
    'Affiliates L1M': affL1M.length,
    'Affiliates L6M': affL6M.length,
    'affiliates monday.com': affMon.length,
    '>0 TTL total': gt0.length,
    '0 TTL monday.com': z_mon.length
  };

  console.log('Final counts:', counts);
  console.log('=== BUILD VENDOR LIST END ===');

  SpreadsheetApp.getActive().toast(
    Object.entries(counts).map(([k,v]) => `${k}: ${v}`).join(' • '),
    'Build List Summary',
    8
  );
}

/**
 * Search Gmail for threads with label:00.received from last 7 days
 * Return a Set of vendor names (lowercased) that have recent activity
 *
 * Detection methods (in priority order):
 * 1. Gmail label "zzzVendors/<vendor_name>" (most accurate)
 * 2. Exact vendor name match in subject/sender/recipient
 * 3. Token-based fuzzy matching (fallback)
 */
function getHotVendorsFromGmail_(allVendors) {
  const hotSet = new Set();

  try {
    // Search for emails with label "00.received" from last 7 days
    const threads = GmailApp.search('label:00.received newer_than:7d', 0, 100);

    console.log(`Found ${threads.length} hot threads`);

    if (threads.length === 0) return hotSet;

    // Build vendor name lookup for fast matching
    const vendorMap = new Map();
    const vendorNames = allVendors.map(v => {
      const nameLower = v.name.toLowerCase();
      vendorMap.set(nameLower, v.name);
      return {
        name: v.name,
        nameLower: nameLower,
        // Also create searchable tokens (words) from vendor name
        tokens: nameLower.split(/\s+/).filter(t => t.length > 2)
      };
    });

    let labelMatches = 0;
    let exactMatches = 0;
    let tokenMatches = 0;

    for (const thread of threads) {
      try {
        let matched = false;

        // METHOD 1 (BEST): Check for zzzVendors/<vendor_name> label
        const labels = thread.getLabels();
        for (const label of labels) {
          const labelName = label.getName();

          // Check if this is a vendor label (zzzVendors/<vendor_name>)
          if (labelName.startsWith('zzzVendors/')) {
            const vendorNameFromLabel = labelName.substring('zzzVendors/'.length).toLowerCase();

            // Try exact match first
            if (vendorMap.has(vendorNameFromLabel)) {
              hotSet.add(vendorNameFromLabel);
              labelMatches++;
              console.log(`HOT: ${vendorMap.get(vendorNameFromLabel)} (label: ${labelName})`);
              matched = true;
              break;
            }

            // Try partial match (label contains vendor name or vice versa)
            for (const vendor of vendorNames) {
              if (vendorNameFromLabel.includes(vendor.nameLower) ||
                  vendor.nameLower.includes(vendorNameFromLabel)) {
                hotSet.add(vendor.nameLower);
                labelMatches++;
                console.log(`HOT: ${vendor.name} (label partial match: ${labelName})`);
                matched = true;
                break;
              }
            }

            if (matched) break;
          }
        }

        // If matched by label, skip other methods
        if (matched) continue;

        // METHOD 2: Exact name match in subject/sender/recipient
        const subject = thread.getFirstMessageSubject().toLowerCase();
        const messages = thread.getMessages();

        let emailText = subject;
        if (messages.length > 0) {
          const firstMsg = messages[0];
          emailText += ' ' + firstMsg.getFrom().toLowerCase();
          emailText += ' ' + firstMsg.getTo().toLowerCase();
        }

        for (const vendor of vendorNames) {
          // Try exact match
          if (emailText.includes(vendor.nameLower)) {
            hotSet.add(vendor.nameLower);
            exactMatches++;
            console.log(`HOT: ${vendor.name} (exact match in: "${subject.substring(0, 50)}...")`);
            matched = true;
            break;
          }
        }

        // If matched by exact text, skip token matching
        if (matched) continue;

        // METHOD 3 (FALLBACK): Token-based matching (at least 2 significant words match)
        for (const vendor of vendorNames) {
          if (vendor.tokens.length >= 2) {
            const matchCount = vendor.tokens.filter(token => emailText.includes(token)).length;
            if (matchCount >= Math.min(2, vendor.tokens.length)) {
              hotSet.add(vendor.nameLower);
              tokenMatches++;
              console.log(`HOT: ${vendor.name} (token match: ${matchCount} tokens in: "${subject.substring(0, 50)}...")`);
              break;
            }
          }
        }
      } catch (e) {
        console.log(`Error processing thread: ${e.message}`);
      }
    }

    console.log(`Hot vendor detection summary: ${labelMatches} label matches, ${exactMatches} exact matches, ${tokenMatches} token matches`);

  } catch (e) {
    console.log(`Error searching Gmail: ${e.message}`);
    // Non-fatal - continue with empty hot set
  }

  return hotSet;
}

/**
 * Build status lookup maps from monday.com sheets
 * Returns { buyers: Map, affiliates: Map }
 */
function buildStatusMaps_(shMonB, shMonA) {
  const buyersMap = new Map();
  const affiliatesMap = new Map();

  if (shMonB) {
    const buyersData = shMonB.getDataRange().getValues();
    for (let i = 1; i < buyersData.length; i++) { // Skip header
      const row = buyersData[i];
      const vendor = normalizeName_(row[0]);
      const status = row[BUYERS_STATUS_INDEX]; // Column R
      if (vendor) {
        buyersMap.set(vendor.toLowerCase(), String(status || '').trim());
      }
    }
  }

  if (shMonA) {
    const affiliatesData = shMonA.getDataRange().getValues();
    for (let i = 1; i < affiliatesData.length; i++) { // Skip header
      const row = affiliatesData[i];
      const vendor = normalizeName_(row[0]);
      const status = row[AFFILIATES_STATUS_INDEX]; // Column S
      if (vendor) {
        affiliatesMap.set(vendor.toLowerCase(), String(status || '').trim());
      }
    }
  }

  return { buyers: buyersMap, affiliates: affiliatesMap };
}

/**
 * Lookup status for a vendor from monday.com status maps
 */
function lookupStatus_(vendorName, vendorType, statusMaps) {
  const key = vendorName.toLowerCase();
  const isBuyer = (vendorType || '').toLowerCase().includes('buyer');

  if (isBuyer && statusMaps.buyers.has(key)) {
    return statusMaps.buyers.get(key);
  } else if (!isBuyer && statusMaps.affiliates.has(key)) {
    return statusMaps.affiliates.get(key);
  }

  return '';
}

/**
 * Standalone function to lookup and update statuses
 * (Kept for backward compatibility - now integrated into buildVendorList)
 */
function lookupVendorStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = ss.getSheetByName(SHEET_OUT);
  const buyersSheet = ss.getSheetByName(SHEET_MON_BUYERS);
  const affiliatesSheet = ss.getSheetByName(SHEET_MON_AFFILIATES);

  if (!mainSheet) {
    throw new Error(`Sheet "${SHEET_OUT}" not found. Run "Build List" first.`);
  }

  if (!buyersSheet || !affiliatesSheet) {
    throw new Error("Missing required sheets: 'buyers monday.com' or 'affiliates monday.com'");
  }

  // Build status maps
  const statusMaps = buildStatusMaps_(buyersSheet, affiliatesSheet);

  // Get data from List sheet
  const mainData = mainSheet.getDataRange().getValues();

  // Update status for each vendor (skip header row)
  for (let i = 1; i < mainData.length; i++) {
    const vendorName = String(mainData[i][0]).trim();
    const source = String(mainData[i][2]).toLowerCase(); // Source column

    if (!vendorName) continue;

    const isBuyer = source.includes("buyer");
    const status = lookupStatus_(vendorName, isBuyer ? 'Buyer' : 'Affiliate', statusMaps);

    if (status) {
      mainSheet.getRange(i + 1, 4).setValue(status); // Column D (Status)
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Vendor status lookup complete.", "✅ Done", 3);
}

/** ---------- Readers & Helpers ---------- **/

/**
 * Read a metric sheet (L1M/L6M) and return [{name, ttl, type, source}, ...]
 * Uses header index detection for vendor and TTL columns.
 */
function readMetricSheet_(sh, type, sourceLabel, blacklist) {
  const { columns, rows } = readObjectsFromSheet_(sh);
  console.log(`[${sourceLabel}] Total rows:`, rows.length);

  if (!rows.length) return [];
  const vHeader = firstExistingHeader_(columns, VENDOR_HEADER_CANDIDATES, sh.getName(), 'vendor');
  const tHeader = firstExistingHeader_(columns, TTL_HEADER_CANDIDATES,    sh.getName(), 'TTL , USD');

  const out = [];
  for (const r of rows) {
    const name = normalizeName_(r[vHeader]);
    if (!name) continue;

    const key = name.toLowerCase();
    if (blacklist.has(key)) {
      recordSkipReason_({ name, source: sourceLabel, reason: 'blacklist' });
      continue;
    }

    out.push({
      name,
      ttl: toNumber_(r[tHeader]),
      type,
      source: sourceLabel
    });
  }
  return dedupeKeepMaxTTL_(out);
}

/**
 * Read a monday.com sheet:
 * - Name from Column A (index 0)
 * - Aliases in Column P (index 15)
 * - TTL optional via recognized TTL header (or specific column)
 * - Status from Column R (buyers, index 17) or Column S (affiliates, index 18)
 * - Dedupes vs existingSet and blacklist
 */
function readMondaySheet_(sh, type, sourceLabel, blacklist, existingSet) {
  console.log(`\n[${sourceLabel}] START READ`);

  // Get raw array data instead of objects for monday.com sheets
  const allValues = sh.getDataRange().getValues();
  if (allValues.length < 2) return []; // Need at least header + 1 row

  const headers = allValues[0].map(h => String(h || '').trim());
  console.log(`[${sourceLabel}] Total rows:`, allValues.length - 1);

  // For monday.com exports: name is always in column A (index 0)
  const nameIdx = 0; // Column A

  const ttlIdx = headers.findIndex(h =>
    eq_(h, 'TTL , USD') || eq_(h, 'TTL, USD') || eq_(h, 'TTL USD')
  );
  const typeIdx = headers.findIndex(h => eq_(h, 'Type'));

  // Status is in column R (index 17) for buyers, column S (index 18) for affiliates
  let statusIdx = null;
  const shNameLower = sh.getName().toLowerCase();
  if (shNameLower === SHEET_MON_BUYERS.toLowerCase()) {
    statusIdx = BUYERS_STATUS_INDEX; // 17 (Column R)
  } else if (shNameLower === SHEET_MON_AFFILIATES.toLowerCase()) {
    statusIdx = AFFILIATES_STATUS_INDEX; // 18 (Column S)
  } else {
    statusIdx = (type.toLowerCase().startsWith('buyer')) ? BUYERS_STATUS_INDEX : AFFILIATES_STATUS_INDEX;
  }

  const out = [];
  let skippedBlacklist = 0;
  let skippedExisting = 0;
  let skippedAlias = 0;
  let skippedNoName = 0;

  // Process rows starting from index 1 (skip header)
  for (let i = 1; i < allValues.length; i++) {
    const row = allValues[i];

    // Name is always in column A
    const rawName = row[nameIdx] || '';
    const name = normalizeName_(rawName);

    if (!name) {
      skippedNoName++;
      continue;
    }

    const key = name.toLowerCase();

    if (blacklist.has(key)) {
      skippedBlacklist++;
      continue;
    }

    if (existingSet.has(key)) {
      skippedExisting++;
      continue;
    }

    // Aliases in Column P (index 15)
    const aliasRaw = (ALIAS_COL_INDEX < row.length) ? (row[ALIAS_COL_INDEX] || '') : '';
    const aliases = String(aliasRaw).split(',')
      .map(s => normalizeName_(s).toLowerCase())
      .filter(Boolean);

    if (aliases.some(a => existingSet.has(a) || blacklist.has(a))) {
      skippedAlias++;
      recordSkipReason_({ name, source: sourceLabel, reason: 'alias-duplicate' });
      continue;
    }

    const ttl = (ttlIdx !== -1 && ttlIdx < row.length) ? toNumber_(row[ttlIdx]) : 0;

    const explicitType = (typeIdx !== -1 && typeIdx < row.length) ? String(row[typeIdx] || '').trim() : '';
    const finalType = explicitType || type;

    const statusRaw = (statusIdx != null && statusIdx < row.length) ? String(row[statusIdx] || '').trim() : '';
    const status = normalizeStatus_(statusRaw);

    out.push({ name, ttl, type: finalType, source: sourceLabel, status });
    existingSet.add(key);
  }

  console.log(`[${sourceLabel}] Skip summary:`, {
    noName: skippedNoName,
    blacklist: skippedBlacklist,
    existing: skippedExisting,
    alias: skippedAlias,
    added: out.length
  });
  console.log(`[${sourceLabel}] END READ\n`);

  return dedupeKeepMaxTTL_(out);
}

/** Normalize buyer or affiliate status to our buckets */
function normalizeStatus_(s) {
  const v = String(s || '').trim().toLowerCase();
  if (!v) return 'other';

  if (v.includes('live')) return 'live';
  if (v.includes('onboard')) return 'onboarding';
  if (v.includes('pause')) return 'paused';
  if (v.includes('pre')) return 'preonboarding';
  if (v.includes('early')) return 'early talks';
  if (v.includes('top') && v.includes('500')) return 'top 500 remodelers';
  if (v.includes('dead') || (v.includes('no') && v.includes('go')) || v.includes('closed')) return 'dead';

  return 'other';
}

/** Reads Settings column E "blacklist" and returns a Set of normalized, lowercased names */
function readBlacklist_(ss) {
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const set = new Set();
  if (!sh) return set;

  const last = sh.getLastRow();
  if (last < 2) return set;

  const header = String(sh.getRange(1, 5).getValue() || '').trim().toLowerCase(); // E1
  if (header !== 'blacklist' && header !== 'vendor blacklist') return set;

  const vals = sh.getRange(2, 5, last - 1, 1).getValues().flat(); // E2:E
  for (const v of vals) {
    const norm = normalizeName_(v);
    if (norm) set.add(norm.toLowerCase());
  }
  return set;
}

/** Data readers */
function readObjectsFromSheet_(sh) {
  const values = sh.getDataRange().getValues();
  if (!values.length) return { columns: [], rows: [] };
  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj;
  });
  return { columns: headers, rows };
}

/** De-dupe by lowercased name keeping max TTL */
function dedupeKeepMaxTTL_(rows) {
  const map = new Map();
  for (const r of rows) {
    const key = r.name.toLowerCase();
    const prev = map.get(key);
    if (!prev || (r.ttl || 0) > (prev.ttl || 0)) {
      map.set(key, r);
    }
  }
  return Array.from(map.values());
}

/** Normalize display name by removing leading "[123]" token if present */
function normalizeName_(v) {
  if (v == null) return '';
  let s = String(v).trim();
  s = s.replace(/^\[\s*\d+\s*\]\s*/, '');
  return s.trim();
}

/** Robust currency parsing */
function toNumber_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  let neg = false;
  if (s.startsWith('(') && s.endsWith(')')) {
    neg = true; s = s.slice(1, -1);
  }
  s = s.replace(/[$,\s]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}

/** Sheet helpers */
function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function mustGetSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet "${name}".`);
  return sh;
}

/** Header helpers */
function eq_(a, b) {
  return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
}

/** Helper: return the first header name from candidates that exists (not an index) */
function firstExistingHeader_(columns, candidates, sheetName, label) {
  for (const c of candidates) {
    if (columns.some(h => eq_(h, c))) return c;
  }
  throw new Error(`Sheet "${sheetName}" is missing required ${label} column (looked for: ${candidates.join(', ')})`);
}

/** Optional: log why a row was skipped (creates/append to a "Why Not In List" sheet) */
function recordSkipReason_({ name, source, reason, detail }) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('Why Not In List') || ss.insertSheet('Why Not In List');
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Source', 'Vendor', 'Reason']]);
    }
    sh.appendRow([new Date(), source || '', name || '', `${reason || ''}${detail ? `: ${detail}` : ''}`]);
  } catch (e) {
    // non-fatal
  }
}
