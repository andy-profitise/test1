/************************************************************
 * BATTLE STATION - One-by-one vendor review dashboard
 *
 * Features:
 * - Navigate through vendors sequentially via menu
 * - View vendor details, notes, status
 * - See all related emails and monday.com tasks (live search)
 * - Update monday.com notes directly
 * - Mark vendors as reviewed/complete
 *
 * UPDATED: Menu-driven, no checkboxes, live Gmail search
 ************************************************************/

const BS_CFG = {
  // Sheet names
  LIST_SHEET: 'List',
  BATTLE_SHEET: 'Battle Station',
  GMAIL_OUTPUT_SHEET: 'Gmail Review Output',
  TASKS_SHEET: 'monday.com tasks',

  // List sheet columns (0-based)
  L_VENDOR: 0,
  L_TTL_USD: 1,
  L_SOURCE: 2,
  L_STATUS: 3,
  L_NOTES: 4,
  L_GMAIL_LINK: 5,
  L_NO_SNOOZE: 6,
  L_PROCESSED: 7,

  // Battle Station layout
  HEADER_ROWS: 3,
  DATA_START_ROW: 5,

  // Colors
  COLOR_HEADER: '#4a86e8',
  COLOR_SUBHEADER: '#6d9eeb',
  COLOR_EMAIL: '#fff2cc',
  COLOR_TASK: '#d9ead3',
  COLOR_BUTTON: '#e8f0fe',
  COLOR_WARNING: '#f4cccc',
  COLOR_SUCCESS: '#d9ead3',
  COLOR_SNOOZED: '#d0e8f2'
};

/**
 * Helper function: Get current vendor index from the display row
 */
function getCurrentVendorIndex_() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) return null;

  // Look for the navigation bar (should be row 3)
  const cellValue = String(bsSh.getRange(3, 1).getValue() || '');

  // Extract number from "◀ PREVIOUS | Viewing: 1 of 361 | Vendor Name | NEXT ▶"
  const match = cellValue.match(/Viewing:\s*(\d+)\s*of/);

  if (!match) {
    Logger.log(`Could not parse index from navigation: "${cellValue}"`);
    return null;
  }

  return parseInt(match[1]);
}

/**
 * Create or reset the Battle Station sheet
 */
function setupBattleStation() {
  const ss = SpreadsheetApp.getActive();
  let bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  // Create sheet if it doesn't exist
  if (!bsSh) {
    bsSh = ss.insertSheet(BS_CFG.BATTLE_SHEET);
  } else {
    bsSh.clear();
  }

  // Set up the control panel
  bsSh.setColumnWidth(1, 200);
  bsSh.setColumnWidth(2, 250);
  bsSh.setColumnWidth(3, 150);
  bsSh.setColumnWidth(4, 300);

  // Initialize with first vendor
  loadVendorData(1);

  SpreadsheetApp.getUi().alert('Battle Station initialized!\n\nUse the ⚡ Battle Station menu to navigate:\n- ▶ Next Vendor\n- ◀ Previous Vendor\n- 💾 Update monday.com Notes\n- ✓ Mark as Reviewed');
}

/**
 * Load and display data for a specific vendor by index
 */
function loadVendorData(vendorIndex) {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!bsSh || !listSh) {
    throw new Error('Required sheets not found');
  }

  const totalVendors = listSh.getLastRow() - 1; // Exclude header

  // Bounds check
  if (vendorIndex < 1) vendorIndex = 1;
  if (vendorIndex > totalVendors) vendorIndex = totalVendors;

  // Get vendor data
  const listRow = vendorIndex + 1; // +1 for header row
  const vendorData = listSh.getRange(listRow, 1, 1, 8).getValues()[0];

  const vendor = vendorData[BS_CFG.L_VENDOR] || '';
  const ttlUsd = vendorData[BS_CFG.L_TTL_USD] || 0;
  const source = vendorData[BS_CFG.L_SOURCE] || '';
  const status = vendorData[BS_CFG.L_STATUS] || '';
  const notes = vendorData[BS_CFG.L_NOTES] || '';
  const processed = vendorData[BS_CFG.L_PROCESSED] || false;

  // Calculate display values
  const mondayBoardId = source.toLowerCase().includes('buyer') ? '9007735194 (Buyers)' :
                        source.toLowerCase().includes('affiliate') ? '9007716156 (Affiliates)' :
                        '9007735194 (Buyers - default)';

  const processedDisplay = processed ? '✅ Yes (Reviewed)' : '⚠️ No (Needs Review)';

  // Clear content below row 1
  const lastRow = bsSh.getMaxRows();
  if (lastRow > 1) {
    bsSh.getRange(2, 1, lastRow - 1, 4).clearContent().clearFormat();
  }

  let currentRow = 1;

  // Title
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue('⚡ BATTLE STATION - Vendor Command Center')
    .setFontSize(16).setFontWeight('bold')
    .setBackground(BS_CFG.COLOR_HEADER)
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 40);
  currentRow++;

  // Instructions
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue('Use ⚡ Battle Station menu to navigate and perform actions')
    .setFontSize(9)
    .setBackground('#f3f3f3')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 25);
  currentRow++;

  // Navigation bar
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`◀ PREVIOUS | Viewing: ${vendorIndex} of ${totalVendors} | ${vendor} | NEXT ▶`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground('#e8f0fe');

  bsSh.setRowHeight(currentRow, 35);
  currentRow++;

  // Freeze the navigation area
  bsSh.setFrozenRows(currentRow - 1);

  // VENDOR INFO SECTION
  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📊 VENDOR ${vendorIndex} of ${totalVendors}`)
    .setBackground(BS_CFG.COLOR_SUBHEADER)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 30);
  currentRow++;

  // Vendor details - each row is label in A-B, value in C-D merged
  const detailRows = [
    ['Vendor:', vendor],
    ['Total USD:', `$${Number(ttlUsd).toLocaleString()}`],
    ['Source:', source],
    ['monday.com Board:', mondayBoardId],
    ['Status:', status + (status === 'Dead' ? ' ⚠️' : ' ✅')],
    ['Processed:', processedDisplay]
  ];

  for (const detail of detailRows) {
    bsSh.getRange(currentRow, 1, 1, 2).merge().setValue(detail[0]).setFontWeight('bold');
    bsSh.getRange(currentRow, 3, 1, 2).merge().setValue(detail[1]);
    currentRow++;
  }

  // Color coding
  const statusRow = currentRow - 2;
  const processedRow = currentRow - 1;

  if (status === 'Dead') {
    bsSh.getRange(statusRow, 3, 1, 2).setBackground(BS_CFG.COLOR_WARNING);
  }

  if (processed) {
    bsSh.getRange(processedRow, 3, 1, 2).setBackground(BS_CFG.COLOR_SUCCESS);
  } else {
    bsSh.getRange(processedRow, 3, 1, 2).setBackground('#fff4e5');
  }

  // Empty row
  currentRow++;

  // Notes section
  bsSh.getRange(currentRow, 1, 1, 2).merge().setValue('Notes:').setFontWeight('bold');
  currentRow++;
  bsSh.getRange(currentRow, 1, 2, 4).merge()
    .setValue(notes || '(no notes)')
    .setWrap(true)
    .setVerticalAlignment('top')
    .setBackground('#fafafa');
  currentRow += 2;

  // Empty row
  currentRow++;

  // ACTIONS REMINDER
  bsSh.setRowHeight(currentRow, 15);
  currentRow++;

  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue('⚡ ACTIONS - Use menu: 💾 Update monday.com Notes | ✓ Mark as Reviewed | 📧 Open Gmail Search')
    .setBackground(BS_CFG.COLOR_BUTTON)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 40);
  currentRow++;

  currentRow++;
  bsSh.setRowHeight(currentRow, 10);
  currentRow++;

  // EMAILS SECTION
  ss.toast('Searching Gmail...', '📧 Loading', 2);
  const emails = getEmailsForVendor_(vendor, listRow);

  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📧 EMAILS (${emails.length}) - Yellow=Waiting/Customer, Blue=Snoozed, White=Active`)
    .setBackground(BS_CFG.COLOR_EMAIL)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 30);
  currentRow++;

  if (emails.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No emails found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Email headers
    bsSh.getRange(currentRow, 1).setValue('Subject').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Date').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Count').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Labels').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;

    // Email data
    for (const email of emails.slice(0, 20)) {
      bsSh.getRange(currentRow, 1).setValue(email.subject);
      bsSh.getRange(currentRow, 2).setValue(email.date);
      bsSh.getRange(currentRow, 3).setValue(email.count);
      bsSh.getRange(currentRow, 4).setValue(email.labels);

      // Make subject a clickable link
      if (email.link) {
        bsSh.getRange(currentRow, 1)
          .setFormula(`=HYPERLINK("${email.link}", "${email.subject.replace(/"/g, '""')}")`);
      }

      // Color code emails - priority order: waiting/customer > snoozed > active
      if (email.labels.includes('02.waiting/customer')) {
        // Yellow for waiting on customer (highest priority)
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#fff44f');
      } else if (email.isSnoozed) {
        // Light blue for snoozed
        bsSh.getRange(currentRow, 1, 1, 4).setBackground(BS_CFG.COLOR_SNOOZED);
      } else {
        // White for active (not snoozed, not waiting)
        bsSh.getRange(currentRow, 1, 1, 4).setBackground('#ffffff');
      }

      currentRow++;
    }

    if (emails.length > 20) {
      bsSh.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`... and ${emails.length - 20} more emails (showing first 20)`)
        .setFontStyle('italic')
        .setHorizontalAlignment('center');
      currentRow++;
    }
  }

  // Empty row for spacing
  bsSh.setRowHeight(currentRow, 10);
  currentRow++;

  // TASKS SECTION
  const tasks = getTasksForVendor_(vendor, listRow);

  bsSh.getRange(currentRow, 1, 1, 4).merge()
    .setValue(`📋 MONDAY.COM UPDATES (${tasks.length})`)
    .setBackground(BS_CFG.COLOR_TASK)
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  bsSh.setRowHeight(currentRow, 30);
  currentRow++;

  if (tasks.length === 0) {
    bsSh.getRange(currentRow, 1, 1, 4).merge()
      .setValue('No tasks found')
      .setFontStyle('italic')
      .setBackground('#fafafa')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    bsSh.setRowHeight(currentRow, 25);
    currentRow++;
  } else {
    // Task headers
    bsSh.getRange(currentRow, 1).setValue('Task').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 2).setValue('Status').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 3).setValue('Created').setFontWeight('bold').setBackground('#f3f3f3');
    bsSh.getRange(currentRow, 4).setValue('Details').setFontWeight('bold').setBackground('#f3f3f3');
    currentRow++;

    // Task data
    for (const task of tasks) {
      bsSh.getRange(currentRow, 1).setValue(task.subject).setWrap(true);
      bsSh.getRange(currentRow, 2).setValue(task.status).setWrap(true);
      bsSh.getRange(currentRow, 3).setValue(task.created).setWrap(true);
      bsSh.getRange(currentRow, 4).setValue(task.snippet).setWrap(true);

      // Apply strikethrough if status is Done
      if (task.isDone) {
        bsSh.getRange(currentRow, 1, 1, 4)
          .setFontLine('line-through')
          .setFontColor('#999999'); // Gray out the text too
      }

      currentRow++;
    }
  }

  // Auto-resize rows for content
  if (currentRow > 5) {
    bsSh.autoResizeRows(5, currentRow - 5);
  }

  ss.toast(`Loaded vendor ${vendorIndex} of ${totalVendors}`, '✅ Ready', 2);
}

/**
 * Get emails for a specific vendor by searching Gmail live
 * Searches both "All" and "No Snooze" queries and marks snoozed emails
 */
function getEmailsForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) {
    Logger.log('List sheet not found');
    return [];
  }

  try {
    // Get both Gmail search links from the List sheet
    const gmailLinkAll = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();
    const gmailLinkNoSnooze = listSh.getRange(listRow, BS_CFG.L_NO_SNOOZE + 1).getValue();

    Logger.log(`=== GMAIL SEARCH DEBUG ===`);
    Logger.log(`Vendor: ${vendor}`);
    Logger.log(`List Row: ${listRow}`);
    Logger.log(`Gmail Link (All): ${gmailLinkAll}`);
    Logger.log(`Gmail Link (No Snooze): ${gmailLinkNoSnooze}`);

    const allEmails = [];

    // Search with "All" query
    if (gmailLinkAll && gmailLinkAll.toString().includes('#search')) {
      const threadsAll = searchGmailFromLink_(gmailLinkAll, 'All');
      Logger.log(`Found ${threadsAll.length} threads in "All" query`);
      allEmails.push(...threadsAll);
    }

    // Search with "No Snooze" query
    const noSnoozeThreadIds = new Set();
    if (gmailLinkNoSnooze && gmailLinkNoSnooze.toString().includes('#search')) {
      const threadsNoSnooze = searchGmailFromLink_(gmailLinkNoSnooze, 'No Snooze');
      Logger.log(`Found ${threadsNoSnooze.length} threads in "No Snooze" query`);

      // Track which thread IDs are in the "No Snooze" results
      for (const email of threadsNoSnooze) {
        noSnoozeThreadIds.add(email.threadId);
        Logger.log(`No Snooze thread ID: ${email.threadId} - ${email.subject}`);
      }
    }

    // Remove duplicates - keep only unique thread IDs
    const uniqueEmails = [];
    const seenThreadIds = new Set();

    for (const email of allEmails) {
      if (!seenThreadIds.has(email.threadId)) {
        seenThreadIds.add(email.threadId);

        // Mark as snoozed if NOT in the noSnoozeThreadIds set
        email.isSnoozed = !noSnoozeThreadIds.has(email.threadId);

        Logger.log(`Email: ${email.subject}`);
        Logger.log(`  Thread ID: ${email.threadId}`);
        Logger.log(`  In NoSnooze set: ${noSnoozeThreadIds.has(email.threadId)}`);
        Logger.log(`  Is Snoozed: ${email.isSnoozed}`);

        uniqueEmails.push(email);
      }
    }

    // Sort by date descending (most recent first)
    uniqueEmails.sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });

    Logger.log(`Returning ${uniqueEmails.length} total emails (${uniqueEmails.filter(e => e.isSnoozed).length} snoozed)`);

    return uniqueEmails;

  } catch (e) {
    Logger.log(`ERROR in getEmailsForVendor_: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return [];
  }
}

/**
 * Helper function to search Gmail from a link and return email objects
 */
function searchGmailFromLink_(gmailLink, querySetName) {
  try {
    const gmailLinkStr = gmailLink.toString();

    // Extract the search query from the Gmail URL
    const urlParts = gmailLinkStr.split('#search/');
    if (urlParts.length < 2) {
      Logger.log(`Could not parse Gmail search URL for ${querySetName}`);
      return [];
    }

    // Decode the search query
    let searchQuery = decodeURIComponent(urlParts[1]);

    // Clean up the query - remove any trailing fragments
    if (searchQuery.includes('?')) {
      searchQuery = searchQuery.split('?')[0];
    }

    // Replace + with spaces for Gmail search API
    searchQuery = searchQuery.replace(/\+/g, ' ');

    Logger.log(`${querySetName} search query: ${searchQuery}`);

    // Search Gmail with the query (limit to 50 most recent)
    const threads = GmailApp.search(searchQuery, 0, 50);

    Logger.log(`${querySetName} found ${threads.length} threads`);

    const emails = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length === 0) continue;

      const firstMessage = messages[0];

      // Get thread metadata
      const subject = thread.getFirstMessageSubject();
      const date = firstMessage.getDate();
      const labels = thread.getLabels().map(label => label.getName()).join(', ');
      const messageCount = messages.length;
      const threadId = thread.getId();

      // Create Gmail link to this thread
      const threadLink = `https://mail.google.com/mail/u/0/#inbox/${threadId}`;

      // Get snippet safely
      let snippet = '';
      try {
        snippet = firstMessage.getPlainBody().substring(0, 200) + '...';
      } catch (e) {
        snippet = '(unable to load snippet)';
      }

      emails.push({
        threadId: threadId,
        subject: subject || '(no subject)',
        date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        count: messageCount,
        labels: labels,
        link: threadLink,
        querySet: querySetName,
        snippet: snippet,
        isSnoozed: false
      });
    }

    return emails;

  } catch (e) {
    Logger.log(`Error searching Gmail for ${querySetName}: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return [];
  }
}

/**
 * Get tasks for a specific vendor by searching the Tasks board (9007661294)
 * Searches for items containing the vendor name and sorts by status (Done last)
 */
function getTasksForVendor_(vendor, listRow) {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) return [];

  const apiToken = 'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQ1NDc2OTk0NywiYWFpIjoxMSwidWlkIjo1MzkyOTA3OCwiaWFkIjoiMjAyNS0wMS0wN1QxODoyNzo1My4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjA1NzI0NjIsInJnbiI6InVzZTEifQ.h8_RIEP9thz-UwIT5SSbkf73n4mzRwyu7ALSZkkTDE8';

  const tasksBoardId = '9007661294'; // Tasks board

  Logger.log(`=== MONDAY.COM TASKS SEARCH ===`);
  Logger.log(`Vendor: ${vendor}`);
  Logger.log(`Tasks Board ID: ${tasksBoardId}`);

  // Escape the vendor name for GraphQL
  const escapedVendor = vendor.replace(/"/g, '\\"');

  // Query to search for items where the name contains the vendor name
  const query = `
    query {
      boards (ids: [${tasksBoardId}]) {
        items_page (
          limit: 100
          query_params: {
            rules: [
              {
                column_id: "name"
                compare_value: ["${escapedVendor}"]
                operator: contains_text
              }
            ]
          }
        ) {
          items {
            id
            name
            group {
              title
            }
            column_values {
              id
              text
              type
            }
            created_at
            updated_at
          }
        }
      }
    }
  `;

  Logger.log(`Searching Tasks board for items containing: "${vendor}"`);

  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': apiToken
      },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());

    if (result.errors && result.errors.length > 0) {
      Logger.log(`API Error: ${result.errors[0].message}`);
      return [];
    }

    if (!result.data?.boards?.[0]?.items_page?.items) {
      Logger.log('No items found in Tasks board');
      return [];
    }

    const items = result.data.boards[0].items_page.items;
    Logger.log(`Found ${items.length} tasks matching vendor`);

    const tasks = [];

    for (const item of items) {
      const itemName = String(item.name || '');
      const groupTitle = item.group?.title || '';

      // Extract status, notes, and dates from column_values
      let status = '';
      let notes = '';
      let date = '';

      for (const col of item.column_values) {
        // Look for status column by ID (should be "status" or contain "status")
        if (col.id === 'status' || col.id === 'status4' || col.id === 'status_1') {
          status = col.text || '';
          Logger.log(`Found status for "${itemName}": ${status}`);
        }
        // Text columns (notes)
        else if ((col.type === 'text' || col.type === 'long-text') && col.text && !notes) {
          notes = col.text.substring(0, 150);
        }
        // Date column
        else if (col.type === 'date' && col.text) {
          date = col.text;
        }
      }

      // Use created_at if no date column found
      if (!date) {
        date = Utilities.formatDate(new Date(item.created_at), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      tasks.push({
        subject: itemName,
        status: status || 'No Status',
        created: Utilities.formatDate(new Date(item.created_at), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        due: date,
        snippet: notes || `Group: ${groupTitle}`,
        groupTitle: groupTitle,
        isDone: (status.toLowerCase() === 'done'),
        statusSort: (status.toLowerCase() === 'done') ? 'zzz' : status // Done goes last alphabetically
      });
    }

    // Sort: Status ASC (Done last), then Created DESC (newest first)
    tasks.sort((a, b) => {
      // First sort by status (alphabetically, with Done forced to end)
      const statusCompare = a.statusSort.localeCompare(b.statusSort);
      if (statusCompare !== 0) {
        return statusCompare;
      }
      // Within same status, sort by created date (newest first)
      return b.created.localeCompare(a.created);
    });

    Logger.log(`Returning ${tasks.length} tasks (${tasks.filter(t => !t.isDone).length} active, ${tasks.filter(t => t.isDone).length} done)`);

    return tasks.slice(0, 30); // Show up to 30 tasks

  } catch (e) {
    Logger.log(`Error fetching monday.com tasks: ${e.message}`);
    Logger.log(`Stack: ${e.stack}`);
    return [];
  }
}

/**
 * Navigation: Go to next vendor (and mark current as reviewed)
 */
function battleStationNext() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  Logger.log(`=== NEXT NAVIGATION ===`);
  Logger.log(`Current index: ${currentIndex}`);

  if (!currentIndex) {
    Logger.log('Could not get index, defaulting to 1');
    loadVendorData(1);
    return;
  }

  const totalVendors = listSh.getLastRow() - 1;

  // Check if we're at the end
  if (currentIndex >= totalVendors) {
    ss.toast('Already at the last vendor!', '⚠️ End of List', 3);
    return;
  }

  // Mark current as reviewed
  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();
  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);

  Logger.log(`Marking vendor ${currentIndex} as reviewed: ${vendor}`);
  Logger.log(`Moving to vendor ${currentIndex + 1}`);

  ss.toast(`Marked "${vendor}" as reviewed`, '▶️ Next', 2);

  // Move to next
  loadVendorData(currentIndex + 1);
}

/**
 * Navigation: Go to previous vendor
 */
function battleStationPrevious() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  Logger.log(`=== PREVIOUS NAVIGATION ===`);
  Logger.log(`Current index: ${currentIndex}`);

  if (!currentIndex) {
    Logger.log('Could not get index, defaulting to 1');
    loadVendorData(1);
    return;
  }

  // Check if we're at the beginning
  if (currentIndex <= 1) {
    ss.toast('Already at the first vendor!', '⚠️ Start of List', 3);
    return;
  }

  Logger.log(`Moving to vendor ${currentIndex - 1}`);

  loadVendorData(currentIndex - 1);
}

/**
 * Refresh current vendor data
 */
function battleStationRefresh() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);

  if (!bsSh) {
    SpreadsheetApp.getUi().alert('Battle Station not found. Run setupBattleStation() first.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  if (!currentIndex) {
    loadVendorData(1);
    return;
  }

  loadVendorData(currentIndex);
}

/**
 * Update monday.com notes for current vendor from Battle Station notes field
 */
function battleStationUpdateMondayNotes() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  ss.toast('Processing update request...', '⚙️ Update monday.com Notes', 3);

  if (!bsSh || !listSh) {
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  if (!currentIndex || isNaN(currentIndex)) {
    ss.toast('Could not determine current vendor', '❌ Error', 5);
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }

  const listRow = currentIndex + 1;

  // Get vendor name directly from List sheet
  const vendor = String(listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue() || '').trim();

  if (!vendor) {
    ss.toast('Could not find vendor name', '❌ Error', 5);
    SpreadsheetApp.getUi().alert('Error: Could not determine vendor name.');
    return;
  }

  Logger.log(`=== UPDATE NOTES START ===`);
  Logger.log(`Current Index: ${currentIndex}`);
  Logger.log(`List Row: ${listRow}`);
  Logger.log(`Vendor Name: "${vendor}"`);

  ss.toast('Reading notes from Battle Station...', '📝 ' + vendor, 2);

  // Find the notes by looking for the "Notes:" label row
  let notesRow = -1;
  for (let i = 5; i < 30; i++) {
    const label = String(bsSh.getRange(i, 1).getValue() || '');
    if (label.indexOf('Notes:') !== -1) {
      notesRow = i + 1;
      break;
    }
  }

  if (notesRow === -1) {
    ss.toast('Could not find notes field', '❌ Error', 5);
    SpreadsheetApp.getUi().alert('Could not find notes field in Battle Station.');
    return;
  }

  // Get notes from the merged cell
  const notes = String(bsSh.getRange(notesRow, 1).getValue() || '').trim();

  Logger.log(`Notes Row: ${notesRow}`);
  Logger.log(`Notes Content: "${notes}"`);

  if (!notes || notes === '(no notes)') {
    ss.toast('No notes to sync - add notes first', '⚠️ Empty Notes', 5);
    SpreadsheetApp.getUi().alert('No notes to sync.\n\nEdit the notes field in Battle Station first, then use this menu item.');
    return;
  }

  // Get user confirmation
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Update monday.com Notes',
    `Update notes for ${vendor}?\n\nNew notes:\n${notes.substring(0, 200)}${notes.length > 200 ? '...' : ''}`,
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ss.toast('Update cancelled', 'ℹ️ Cancelled', 2);
    return;
  }

  try {
    ss.toast('Updating monday.com...', '⚡ Syncing', 3);

    // Update the List sheet with these notes
    listSh.getRange(listRow, BS_CFG.L_NOTES + 1).setValue(notes);

    // Call monday.com API to update notes
    const result = updateMondayComNotesForVendor_(vendor, notes, listRow);

    if (result.success) {
      ss.toast('Notes updated successfully!', '✅ Success - ' + vendor, 5);
      ui.alert(`✓ Successfully updated monday.com notes for ${vendor}!\n\nNotes also saved to List sheet.`);

      // Refresh the display
      battleStationRefresh();
    } else {
      ss.toast('Update failed: ' + result.error, '❌ Failed', 5);
      ui.alert(`⚠️ Failed to update notes: ${result.error}`);
    }
  } catch (e) {
    ss.toast('Error: ' + e.message, '❌ Error', 5);
    ui.alert(`❌ Error: ${e.message}`);
    Logger.log(`Error updating monday.com notes: ${e}`);
  }
}

/**
 * Helper function to update monday.com notes via API
 */
function updateMondayComNotesForVendor_(vendor, notes, listRow) {
  const apiToken = 'eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjQ1NDc2OTk0NywiYWFpIjoxMSwidWlkIjo1MzkyOTA3OCwiaWFkIjoiMjAyNS0wMS0wN1QxODoyNzo1My4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjA1NzI0NjIsInJnbiI6InVzZTEifQ.h8_RIEP9thz-UwIT5SSbkf73n4mzRwyu7ALSZkkTDE8';

  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listRow) {
    const currentIndex = getCurrentVendorIndex_();
    if (!currentIndex) {
      return {
        success: false,
        error: 'Could not determine current vendor index'
      };
    }
    listRow = currentIndex + 1;
  }

  const source = String(listSh.getRange(listRow, BS_CFG.L_SOURCE + 1).getValue() || '');

  Logger.log(`=== UPDATE MONDAY API CALL ===`);
  Logger.log(`Vendor: "${vendor}"`);
  Logger.log(`Source: "${source}"`);
  Logger.log(`List Row: ${listRow}`);

  // Auto-select board based on source
  let boardId;
  if (source.toLowerCase().includes('buyer')) {
    boardId = '9007735194';
  } else if (source.toLowerCase().includes('affiliate')) {
    boardId = '9007716156';
  } else {
    boardId = '9007735194';
  }

  Logger.log(`Board ID selected: ${boardId}`);

  if (!apiToken) {
    return {
      success: false,
      error: 'monday.com API token is missing from the script.'
    };
  }

  const itemId = findMondayItemIdByVendor_(vendor, boardId, apiToken);

  Logger.log(`Item ID found: ${itemId}`);

  if (!itemId) {
    return {
      success: false,
      error: `Could not find monday.com item for vendor: ${vendor}`
    };
  }

  const notesColumnId = (boardId === '9007735194') ? 'text_mkqnvsqh' : 'text_mkrdahqz';

  Logger.log(`Column ID: ${notesColumnId}`);
  Logger.log(`Notes to update: "${notes}"`);

  const valueJson = JSON.stringify(notes);
  Logger.log(`Value JSON: ${valueJson}`);

  const escapedValue = valueJson.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
  Logger.log(`Escaped value: ${escapedValue}`);

  const mutation = `
    mutation {
      change_column_value (
        board_id: ${boardId},
        item_id: ${itemId},
        column_id: "${notesColumnId}",
        value: "${escapedValue}"
      ) {
        id
      }
    }
  `;

  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': apiToken
      },
      payload: JSON.stringify({ query: mutation }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const responseText = response.getContentText();
    const result = JSON.parse(responseText);

    Logger.log('=== API RESPONSE ===');
    Logger.log(responseText);

    if (result.errors && result.errors.length > 0) {
      Logger.log('❌ API returned errors');
      return {
        success: false,
        error: result.errors[0].message
      };
    }

    if (result.data && result.data.change_column_value && result.data.change_column_value.id) {
      Logger.log('✓ API SUCCESS - Item updated successfully');
      return {
        success: true,
        itemId: result.data.change_column_value.id
      };
    }

    Logger.log('⚠ Unexpected API response format');
    return {
      success: false,
      error: 'Unexpected API response format: ' + responseText
    };

  } catch (e) {
    Logger.log(`❌ Exception during API call: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Find a monday.com item ID by searching for a vendor name
 */
function findMondayItemIdByVendor_(vendor, boardId, apiToken) {
  Logger.log(`=== SEARCHING FOR VENDOR ===`);
  Logger.log(`Search term: "${vendor}"`);
  Logger.log(`Board ID: ${boardId}`);

  // Try 1: Exact match
  let query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${vendor.replace(/"/g, '\\"')}"]}]}) {
          items {
            id
            name
          }
        }
      }
    }
  `;

  let itemId = tryFindItem_(query, apiToken, 'Exact match');
  if (itemId) return itemId;

  // Try 2: Remove parentheses
  const withoutParens = vendor.replace(/\s*\([^)]*\)/g, '').trim();
  if (withoutParens !== vendor) {
    Logger.log(`Trying without parentheses: "${withoutParens}"`);
    query = `
      query {
        boards (ids: [${boardId}]) {
          items_page (limit: 100, query_params: {rules: [{column_id: "name", compare_value: ["${withoutParens.replace(/"/g, '\\"')}"]}]}) {
            items {
              id
              name
            }
          }
        }
      }
    `;

    itemId = tryFindItem_(query, apiToken, 'Without parentheses');
    if (itemId) return itemId;
  }

  // Try 3: Contains search
  Logger.log(`Trying contains search...`);
  query = `
    query {
      boards (ids: [${boardId}]) {
        items_page (limit: 500) {
          items {
            id
            name
          }
        }
      }
    }
  `;

  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': apiToken
      },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());

    if (result.data?.boards?.[0]?.items_page?.items) {
      const items = result.data.boards[0].items_page.items;
      Logger.log(`Got ${items.length} total items, searching for matches...`);

      for (const item of items) {
        const itemName = String(item.name || '').toLowerCase();
        const searchTerm = vendor.toLowerCase();
        const searchTermNoParens = withoutParens.toLowerCase();

        if (itemName.includes(searchTerm) || searchTerm.includes(itemName) ||
            itemName.includes(searchTermNoParens) || searchTermNoParens.includes(itemName)) {
          Logger.log(`✓ FOUND MATCH: "${item.name}" (ID: ${item.id})`);
          return item.id;
        }
      }

      Logger.log('No match found. First 10 items:');
      for (let i = 0; i < Math.min(10, items.length); i++) {
        Logger.log(`  - "${items[i].name}" (ID: ${items[i].id})`);
      }
    }

    return null;
  } catch (e) {
    Logger.log(`Error in contains search: ${e}`);
    return null;
  }
}

/**
 * Helper to try a specific query
 */
function tryFindItem_(query, apiToken, attemptName) {
  try {
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': apiToken
      },
      payload: JSON.stringify({ query: query }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch('https://api.monday.com/v2', options);
    const result = JSON.parse(response.getContentText());

    if (result.data?.boards?.[0]?.items_page?.items?.length > 0) {
      const item = result.data.boards[0].items_page.items[0];
      Logger.log(`✓ ${attemptName} SUCCESS: Found "${item.name}" (ID: ${item.id})`);
      return item.id;
    } else {
      Logger.log(`✗ ${attemptName} failed: No items found`);
    }

    return null;
  } catch (e) {
    Logger.log(`✗ ${attemptName} error: ${e}`);
    return null;
  }
}

/**
 * Mark current vendor as reviewed/processed
 */
function battleStationMarkReviewed() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  ss.toast('Marking as reviewed...', '✓ Mark as Reviewed', 2);

  if (!bsSh || !listSh) {
    ss.toast('Required sheets not found', '❌ Error', 3);
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  if (!currentIndex) {
    ss.toast('Could not determine current vendor', '❌ Error', 3);
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }

  const listRow = currentIndex + 1;
  const vendor = listSh.getRange(listRow, BS_CFG.L_VENDOR + 1).getValue();

  listSh.getRange(listRow, BS_CFG.L_PROCESSED + 1).setValue(true);

  battleStationRefresh();

  ss.toast('Marked as reviewed!', '✅ ' + vendor, 3);
}

/**
 * Open Gmail search for current vendor
 */
function battleStationOpenGmail() {
  const ss = SpreadsheetApp.getActive();
  const bsSh = ss.getSheetByName(BS_CFG.BATTLE_SHEET);
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  ss.toast('Opening Gmail...', '📧 Gmail Search', 2);

  if (!bsSh || !listSh) {
    ss.toast('Required sheets not found', '❌ Error', 3);
    SpreadsheetApp.getUi().alert('Required sheets not found.');
    return;
  }

  const currentIndex = getCurrentVendorIndex_();

  if (!currentIndex) {
    ss.toast('Could not determine current vendor', '❌ Error', 3);
    SpreadsheetApp.getUi().alert('Error: Could not determine current vendor index.');
    return;
  }

  const listRow = currentIndex + 1;
  const gmailLink = listSh.getRange(listRow, BS_CFG.L_GMAIL_LINK + 1).getValue();

  if (!gmailLink || gmailLink.toString().indexOf('#search') === -1) {
    SpreadsheetApp.getUi().alert('No valid Gmail search link found.');
    return;
  }

  const html = `<html><body><script>
    window.open('${gmailLink}', '_blank');
    google.script.host.close();
  </script></body></html>`;

  const ui = HtmlService.createHtmlOutput(html).setWidth(200).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Opening Gmail...');
}

/**
 * Go to a specific vendor by index
 */
function battleStationGoTo() {
  const ss = SpreadsheetApp.getActive();
  const listSh = ss.getSheetByName(BS_CFG.LIST_SHEET);

  if (!listSh) return;

  const totalVendors = listSh.getLastRow() - 1;
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Go to Vendor',
    `Enter vendor index (1-${totalVendors}):`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const index = parseInt(response.getResponseText());
    if (!isNaN(index) && index >= 1 && index <= totalVendors) {
      loadVendorData(index);
    } else {
      ui.alert('Invalid index. Please enter a number between 1 and ' + totalVendors);
    }
  }
}

/**
 * Add menu to Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Battle Station')
    .addItem('🔧 Setup Battle Station', 'setupBattleStation')
    .addSeparator()
    .addItem('▶ Next Vendor', 'battleStationNext')
    .addItem('◀ Previous Vendor', 'battleStationPrevious')
    .addItem('🔄 Refresh', 'battleStationRefresh')
    .addSeparator()
    .addItem('💾 Update monday.com Notes', 'battleStationUpdateMondayNotes')
    .addItem('✓ Mark as Reviewed', 'battleStationMarkReviewed')
    .addItem('📧 Open Gmail Search', 'battleStationOpenGmail')
    .addSeparator()
    .addItem('🔍 Go to Specific Vendor...', 'battleStationGoTo')
    .addToUi();
}
