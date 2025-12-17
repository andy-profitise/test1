/**
 * Label Audit & Repair System
 *
 * Addresses Gmail's label indexing issue where labels are applied to emails
 * but don't appear in search results. This module:
 * 1. Collects all emails under zzzVendors/ sublabels
 * 2. Checks actual labels via API vs. what search returns
 * 3. Reports discrepancies
 * 4. Optionally repairs by removing/re-adding labels
 */

/***** CONFIGURATION *****/

// Labels to audit - these are labels that sometimes don't appear in search
// Add any labels you want to check here
const IMPORTANT_LABELS = [
  '00.received',
  '01.sent',
  '02.waiting/customer',
  '02.waiting/partner',
  '03.action-required'
];

// Parent label for vendor sublabels
const VENDOR_LABEL_PREFIX = 'zzzVendors/';

// Output sheet for audit results
const SHEET_LABEL_AUDIT = 'Label Audit Results';

// Maximum threads to process per vendor (to avoid timeouts)
const MAX_THREADS_PER_VENDOR = 50;

// Maximum vendors to process in one run
const MAX_VENDORS_PER_RUN = 20;

/***** MAIN FUNCTIONS *****/

/**
 * Run a full label audit across all zzzVendors/ sublabels
 * Compares API-reported labels vs. search results for important labels
 */
function auditImportantLabels() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  console.log('=== LABEL AUDIT START ===');

  // Get all vendor labels
  const vendorLabels = getVendorLabels_();
  console.log(`Found ${vendorLabels.length} vendor labels`);

  if (vendorLabels.length === 0) {
    ui.alert('No vendor labels found',
      'Could not find any labels starting with "' + VENDOR_LABEL_PREFIX + '".\n\n' +
      'Make sure you have Gmail labels like "zzzVendors/VendorName".',
      ui.ButtonSet.OK);
    return;
  }

  // Limit vendors per run to avoid timeout
  const vendorsToProcess = vendorLabels.slice(0, MAX_VENDORS_PER_RUN);
  if (vendorLabels.length > MAX_VENDORS_PER_RUN) {
    console.log(`Processing first ${MAX_VENDORS_PER_RUN} of ${vendorLabels.length} vendors`);
  }

  // Build search index for important labels
  const importantLabelSearchIndex = buildImportantLabelSearchIndex_();
  console.log(`Built search index for ${Object.keys(importantLabelSearchIndex).length} important labels`);

  // Audit results
  const discrepancies = [];
  let totalThreadsChecked = 0;
  let totalDiscrepancies = 0;

  for (const vendorLabel of vendorsToProcess) {
    const vendorName = vendorLabel.getName().substring(VENDOR_LABEL_PREFIX.length);
    console.log(`\nAuditing: ${vendorName}`);

    try {
      // Get all threads with this vendor label
      const threads = GmailApp.search('label:' + formatLabelForSearch_(vendorLabel.getName()), 0, MAX_THREADS_PER_VENDOR);
      console.log(`  Found ${threads.length} threads`);

      for (const thread of threads) {
        totalThreadsChecked++;
        const threadId = thread.getId();
        const subject = thread.getFirstMessageSubject() || '(no subject)';
        const date = thread.getLastMessageDate();

        // Get actual labels from API
        const actualLabels = thread.getLabels().map(l => l.getName());

        // Check each important label
        for (const importantLabel of IMPORTANT_LABELS) {
          const hasLabelViaApi = actualLabels.some(l => l.toLowerCase() === importantLabel.toLowerCase());
          const appearsInSearch = importantLabelSearchIndex[importantLabel]?.has(threadId) || false;

          if (hasLabelViaApi && !appearsInSearch) {
            // DISCREPANCY: Has label but doesn't appear in search
            totalDiscrepancies++;
            discrepancies.push({
              vendor: vendorName,
              threadId: threadId,
              subject: subject.substring(0, 80),
              date: date,
              label: importantLabel,
              issue: 'Has label but NOT in search',
              gmailLink: 'https://mail.google.com/mail/u/0/#inbox/' + threadId
            });
            console.log(`  DISCREPANCY: "${subject.substring(0, 40)}..." has ${importantLabel} but not in search`);
          }
        }
      }
    } catch (e) {
      console.log(`  Error processing ${vendorName}: ${e.message}`);
    }
  }

  console.log(`\n=== AUDIT COMPLETE ===`);
  console.log(`Threads checked: ${totalThreadsChecked}`);
  console.log(`Discrepancies found: ${totalDiscrepancies}`);

  // Write results to sheet
  writeAuditResults_(ss, discrepancies, totalThreadsChecked, vendorsToProcess.length);

  // Show summary
  if (discrepancies.length > 0) {
    ui.alert('Label Audit Complete',
      `Checked ${totalThreadsChecked} threads across ${vendorsToProcess.length} vendors.\n\n` +
      `Found ${discrepancies.length} label discrepancies!\n\n` +
      `Results written to "${SHEET_LABEL_AUDIT}" sheet.\n\n` +
      `Run "Repair Label Discrepancies" to fix these issues.`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('Label Audit Complete',
      `Checked ${totalThreadsChecked} threads across ${vendorsToProcess.length} vendors.\n\n` +
      `No discrepancies found! All labels are properly indexed.`,
      ui.ButtonSet.OK);
  }
}

/**
 * Repair label discrepancies by removing and re-adding the label
 * This forces Gmail to re-index the label
 */
function repairLabelDiscrepancies() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Get audit results sheet
  const auditSheet = ss.getSheetByName(SHEET_LABEL_AUDIT);
  if (!auditSheet) {
    ui.alert('No Audit Results',
      'Please run "Audit Important Labels" first to identify discrepancies.',
      ui.ButtonSet.OK);
    return;
  }

  // Get discrepancies from sheet (skip header and summary rows)
  const data = auditSheet.getDataRange().getValues();
  const dataStartRow = findDataStartRow_(data);

  if (dataStartRow === -1) {
    ui.alert('No Discrepancies',
      'No label discrepancies found in the audit results.',
      ui.ButtonSet.OK);
    return;
  }

  // Count discrepancies
  const discrepancies = [];
  for (let i = dataStartRow; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[2]) { // Has vendor and threadId
      discrepancies.push({
        row: i + 1,
        vendor: row[0],
        threadId: row[2],
        subject: row[3],
        label: row[5],
        status: row[7] || ''
      });
    }
  }

  // Filter to only unrepaired items
  const unrepairedItems = discrepancies.filter(d => d.status !== 'REPAIRED' && d.status !== 'SKIPPED');

  if (unrepairedItems.length === 0) {
    ui.alert('All Repaired',
      'All discrepancies have already been processed.',
      ui.ButtonSet.OK);
    return;
  }

  // Confirm repair
  const response = ui.alert('Confirm Repair',
    `Found ${unrepairedItems.length} unrepaired label discrepancies.\n\n` +
    `This will remove and re-add the labels to force Gmail to re-index them.\n\n` +
    `Proceed with repair?`,
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    return;
  }

  console.log('=== LABEL REPAIR START ===');

  let repaired = 0;
  let failed = 0;

  for (const item of unrepairedItems) {
    try {
      const thread = GmailApp.getThreadById(item.threadId);
      if (!thread) {
        console.log(`Thread not found: ${item.threadId}`);
        auditSheet.getRange(item.row, 8).setValue('SKIPPED - Thread not found');
        failed++;
        continue;
      }

      // Get the label object
      const label = GmailApp.getUserLabelByName(item.label);
      if (!label) {
        console.log(`Label not found: ${item.label}`);
        auditSheet.getRange(item.row, 8).setValue('SKIPPED - Label not found');
        failed++;
        continue;
      }

      // Check if thread actually has this label
      const currentLabels = thread.getLabels();
      const hasLabel = currentLabels.some(l => l.getName().toLowerCase() === item.label.toLowerCase());

      if (hasLabel) {
        // Remove and re-add the label
        thread.removeLabel(label);
        Utilities.sleep(100); // Small delay to ensure removal is processed
        thread.addLabel(label);

        console.log(`Repaired: ${item.subject} - ${item.label}`);
        auditSheet.getRange(item.row, 8).setValue('REPAIRED');
        repaired++;
      } else {
        // Label was already removed somehow
        console.log(`Label already missing: ${item.subject} - ${item.label}`);
        auditSheet.getRange(item.row, 8).setValue('SKIPPED - Label already gone');
        failed++;
      }

      // Avoid rate limiting
      if (repaired % 10 === 0) {
        Utilities.sleep(500);
      }

    } catch (e) {
      console.log(`Error repairing ${item.threadId}: ${e.message}`);
      auditSheet.getRange(item.row, 8).setValue('ERROR: ' + e.message);
      failed++;
    }
  }

  console.log('=== LABEL REPAIR COMPLETE ===');
  console.log(`Repaired: ${repaired}, Failed: ${failed}`);

  ui.alert('Repair Complete',
    `Repair process finished.\n\n` +
    `Successfully repaired: ${repaired}\n` +
    `Failed/Skipped: ${failed}\n\n` +
    `Note: It may take a few minutes for Gmail to re-index the labels.`,
    ui.ButtonSet.OK);
}

/**
 * Quick audit for a single important label
 * Useful for checking just "00.received" without checking all vendors
 */
function auditSingleLabel() {
  const ui = SpreadsheetApp.getUi();

  // Build selection prompt
  let labelOptions = IMPORTANT_LABELS.map((l, i) => `${i + 1}. ${l}`).join('\n');

  const response = ui.prompt('Audit Single Label',
    'Enter the number of the label to audit:\n\n' + labelOptions,
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const selection = parseInt(response.getResponseText().trim());
  if (isNaN(selection) || selection < 1 || selection > IMPORTANT_LABELS.length) {
    ui.alert('Invalid selection');
    return;
  }

  const labelToAudit = IMPORTANT_LABELS[selection - 1];
  auditSpecificLabel_(labelToAudit);
}

/**
 * Audit a specific label by comparing all threads that have it (via API scan)
 * against threads returned by search
 */
function auditSpecificLabel_(labelName) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  console.log(`=== AUDITING LABEL: ${labelName} ===`);

  // Get the label object
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    ui.alert('Label Not Found',
      `Could not find label "${labelName}".`,
      ui.ButtonSet.OK);
    return;
  }

  // Method 1: Get threads via label.getThreads() - this gets threads that ACTUALLY have the label
  console.log('Getting threads via API...');
  const threadsViaApi = label.getThreads(0, 200);
  const apiThreadIds = new Set(threadsViaApi.map(t => t.getId()));
  console.log(`Found ${apiThreadIds.size} threads via API`);

  // Method 2: Get threads via search - this is what we're comparing against
  console.log('Getting threads via search...');
  const searchQuery = 'label:' + formatLabelForSearch_(labelName);
  const threadsViaSearch = GmailApp.search(searchQuery, 0, 200);
  const searchThreadIds = new Set(threadsViaSearch.map(t => t.getId()));
  console.log(`Found ${searchThreadIds.size} threads via search`);

  // Find discrepancies
  const inApiNotSearch = [];
  for (const threadId of apiThreadIds) {
    if (!searchThreadIds.has(threadId)) {
      const thread = GmailApp.getThreadById(threadId);
      inApiNotSearch.push({
        threadId: threadId,
        subject: thread.getFirstMessageSubject() || '(no subject)',
        date: thread.getLastMessageDate(),
        gmailLink: 'https://mail.google.com/mail/u/0/#inbox/' + threadId
      });
    }
  }

  const inSearchNotApi = [];
  for (const threadId of searchThreadIds) {
    if (!apiThreadIds.has(threadId)) {
      const thread = GmailApp.getThreadById(threadId);
      inSearchNotApi.push({
        threadId: threadId,
        subject: thread.getFirstMessageSubject() || '(no subject)',
        date: thread.getLastMessageDate()
      });
    }
  }

  console.log(`In API but not search: ${inApiNotSearch.length}`);
  console.log(`In search but not API: ${inSearchNotApi.length}`);

  // Write results
  writeSingleLabelAuditResults_(ss, labelName, apiThreadIds.size, searchThreadIds.size, inApiNotSearch, inSearchNotApi);

  // Show summary
  if (inApiNotSearch.length > 0 || inSearchNotApi.length > 0) {
    ui.alert('Audit Complete',
      `Label: ${labelName}\n\n` +
      `Threads with label (API): ${apiThreadIds.size}\n` +
      `Threads in search results: ${searchThreadIds.size}\n\n` +
      `Has label but NOT in search: ${inApiNotSearch.length}\n` +
      `In search but NOT via API: ${inSearchNotApi.length}\n\n` +
      `Results written to "${SHEET_LABEL_AUDIT}" sheet.`,
      ui.ButtonSet.OK);
  } else {
    ui.alert('Audit Complete',
      `Label: ${labelName}\n\n` +
      `All ${apiThreadIds.size} threads are properly indexed!\n` +
      `No discrepancies found.`,
      ui.ButtonSet.OK);
  }
}

/**
 * Show current important labels configuration
 */
function showImportantLabelsConfig() {
  const ui = SpreadsheetApp.getUi();

  const labelList = IMPORTANT_LABELS.map((l, i) => `${i + 1}. ${l}`).join('\n');

  ui.alert('Important Labels Configuration',
    'The following labels are being audited:\n\n' +
    labelList + '\n\n' +
    'To modify this list, edit the IMPORTANT_LABELS array in labelAudit.gs',
    ui.ButtonSet.OK);
}

/***** HELPER FUNCTIONS *****/

/**
 * Get all Gmail labels that start with zzzVendors/
 * @returns {GmailLabel[]} Array of vendor label objects
 */
function getVendorLabels_() {
  const allLabels = GmailApp.getUserLabels();
  return allLabels.filter(label => label.getName().startsWith(VENDOR_LABEL_PREFIX));
}

/**
 * Build an index of which threads appear in search for each important label
 * @returns {Object} Map of labelName -> Set of threadIds
 */
function buildImportantLabelSearchIndex_() {
  const index = {};

  for (const labelName of IMPORTANT_LABELS) {
    index[labelName] = new Set();

    try {
      const searchQuery = 'label:' + formatLabelForSearch_(labelName);
      const threads = GmailApp.search(searchQuery, 0, 200);

      for (const thread of threads) {
        index[labelName].add(thread.getId());
      }

      console.log(`Search index for ${labelName}: ${index[labelName].size} threads`);
    } catch (e) {
      console.log(`Error building search index for ${labelName}: ${e.message}`);
    }
  }

  return index;
}

/**
 * Format a label name for use in Gmail search
 * Handles special characters and spaces
 * @param {string} labelName - The label name
 * @returns {string} Formatted label for search query
 */
function formatLabelForSearch_(labelName) {
  // Replace / with - for search (Gmail search syntax)
  // Also handle spaces by wrapping in quotes or using hyphens
  let formatted = labelName.replace(/\//g, '-');

  // If label contains special chars or spaces, it might need different handling
  // Gmail's search uses hyphens instead of slashes
  return formatted;
}

/**
 * Write audit results to a sheet
 */
function writeAuditResults_(ss, discrepancies, totalThreads, vendorCount) {
  const sheet = ensureLabelAuditSheet_(ss);
  sheet.clearContents();

  // Write header/summary
  const summaryData = [
    ['LABEL AUDIT RESULTS', '', '', '', '', '', '', ''],
    ['Run Date:', new Date(), '', 'Vendors Checked:', vendorCount, '', '', ''],
    ['Threads Scanned:', totalThreads, '', 'Discrepancies Found:', discrepancies.length, '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['Vendor', 'Subject', 'Thread ID', 'Date', 'Gmail Link', 'Missing Label', 'Issue', 'Repair Status']
  ];

  sheet.getRange(1, 1, summaryData.length, 8).setValues(summaryData);

  // Style header
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setFontSize(14);
  sheet.getRange(5, 1, 1, 8).setFontWeight('bold').setBackground('#e0e0e0');

  // Write discrepancies
  if (discrepancies.length > 0) {
    const discData = discrepancies.map(d => [
      d.vendor,
      d.subject,
      d.threadId,
      d.date,
      d.gmailLink,
      d.label,
      d.issue,
      ''
    ]);
    sheet.getRange(6, 1, discData.length, 8).setValues(discData);

    // Format date column
    sheet.getRange(6, 4, discData.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 8);

  // Make Gmail links clickable
  if (discrepancies.length > 0) {
    for (let i = 0; i < discrepancies.length; i++) {
      const cell = sheet.getRange(6 + i, 5);
      const url = discrepancies[i].gmailLink;
      cell.setFormula(`=HYPERLINK("${url}", "Open Thread")`);
    }
  }
}

/**
 * Write single label audit results
 */
function writeSingleLabelAuditResults_(ss, labelName, apiCount, searchCount, inApiNotSearch, inSearchNotApi) {
  const sheet = ensureLabelAuditSheet_(ss);
  sheet.clearContents();

  // Write header/summary
  const summaryData = [
    [`SINGLE LABEL AUDIT: ${labelName}`, '', '', '', '', ''],
    ['Run Date:', new Date(), '', '', '', ''],
    ['Threads via API:', apiCount, '', 'Threads via Search:', searchCount, ''],
    ['In API but NOT search:', inApiNotSearch.length, '', 'In search but NOT API:', inSearchNotApi.length, ''],
    ['', '', '', '', '', ''],
    ['DISCREPANCIES: Has label but NOT in search', '', '', '', '', ''],
    ['Thread ID', 'Subject', 'Date', 'Gmail Link', 'Repair Status', '']
  ];

  sheet.getRange(1, 1, summaryData.length, 6).setValues(summaryData);

  // Style headers
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setFontSize(14);
  sheet.getRange(6, 1, 1, 6).setFontWeight('bold').setBackground('#ffcccc');
  sheet.getRange(7, 1, 1, 6).setFontWeight('bold').setBackground('#e0e0e0');

  // Write discrepancies (in API but not search - these need repair)
  let currentRow = 8;
  if (inApiNotSearch.length > 0) {
    const discData = inApiNotSearch.map(d => [
      d.threadId,
      d.subject,
      d.date,
      d.gmailLink,
      '',
      ''
    ]);
    sheet.getRange(currentRow, 1, discData.length, 6).setValues(discData);

    // Make Gmail links clickable
    for (let i = 0; i < inApiNotSearch.length; i++) {
      const cell = sheet.getRange(currentRow + i, 4);
      const url = inApiNotSearch[i].gmailLink;
      cell.setFormula(`=HYPERLINK("${url}", "Open Thread")`);
    }

    currentRow += inApiNotSearch.length + 2;
  }

  // Write reverse discrepancies (if any)
  if (inSearchNotApi.length > 0) {
    sheet.getRange(currentRow, 1, 1, 6).setValues([['ANOMALY: In search but NOT via API (may indicate stale index)', '', '', '', '', '']]);
    sheet.getRange(currentRow, 1, 1, 6).setFontWeight('bold').setBackground('#ffffcc');
    currentRow++;

    sheet.getRange(currentRow, 1, 1, 6).setValues([['Thread ID', 'Subject', 'Date', '', '', '']]);
    sheet.getRange(currentRow, 1, 1, 6).setFontWeight('bold').setBackground('#e0e0e0');
    currentRow++;

    const anomalyData = inSearchNotApi.map(d => [d.threadId, d.subject, d.date, '', '', '']);
    sheet.getRange(currentRow, 1, anomalyData.length, 6).setValues(anomalyData);
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);
}

/**
 * Find where the data rows start (after headers/summary)
 * @param {Array[]} data - Sheet data
 * @returns {number} Row index where data starts, or -1 if not found
 */
function findDataStartRow_(data) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Look for a row that looks like data (has thread ID pattern)
    if (row[2] && String(row[2]).match(/^[a-f0-9]+$/i)) {
      return i;
    }
  }
  return -1;
}

/**
 * Ensure the Label Audit sheet exists
 */
function ensureLabelAuditSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_LABEL_AUDIT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LABEL_AUDIT);
  }
  return sheet;
}

/**
 * Repair a specific label across all threads that have it
 * Finds threads with the label via API and re-syncs them
 */
function repairSpecificLabel() {
  const ui = SpreadsheetApp.getUi();

  // Build selection prompt
  let labelOptions = IMPORTANT_LABELS.map((l, i) => `${i + 1}. ${l}`).join('\n');

  const response = ui.prompt('Repair Specific Label',
    'This will remove and re-add the label on ALL threads that have it.\n' +
    'Use this to force Gmail to fully re-index the label.\n\n' +
    'Enter the number of the label to repair:\n\n' + labelOptions,
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const selection = parseInt(response.getResponseText().trim());
  if (isNaN(selection) || selection < 1 || selection > IMPORTANT_LABELS.length) {
    ui.alert('Invalid selection');
    return;
  }

  const labelToRepair = IMPORTANT_LABELS[selection - 1];

  // Confirm
  const confirmResponse = ui.alert('Confirm Full Repair',
    `This will remove and re-add "${labelToRepair}" on ALL threads that have it.\n\n` +
    `This could affect many emails. Are you sure?`,
    ui.ButtonSet.YES_NO);

  if (confirmResponse !== ui.Button.YES) return;

  // Get the label
  const label = GmailApp.getUserLabelByName(labelToRepair);
  if (!label) {
    ui.alert('Label Not Found', `Could not find label "${labelToRepair}".`, ui.ButtonSet.OK);
    return;
  }

  console.log(`=== REPAIRING ALL: ${labelToRepair} ===`);

  // Get threads with this label
  const threads = label.getThreads(0, 500);
  console.log(`Found ${threads.length} threads with ${labelToRepair}`);

  let repaired = 0;
  let failed = 0;

  for (const thread of threads) {
    try {
      thread.removeLabel(label);
      Utilities.sleep(50);
      thread.addLabel(label);
      repaired++;

      if (repaired % 20 === 0) {
        console.log(`Repaired ${repaired} threads...`);
        Utilities.sleep(500);
      }
    } catch (e) {
      console.log(`Failed: ${e.message}`);
      failed++;
    }
  }

  console.log(`=== REPAIR COMPLETE: ${repaired} repaired, ${failed} failed ===`);

  ui.alert('Repair Complete',
    `Label: ${labelToRepair}\n\n` +
    `Successfully repaired: ${repaired}\n` +
    `Failed: ${failed}\n\n` +
    `Note: It may take a few minutes for Gmail to re-index.`,
    ui.ButtonSet.OK);
}
