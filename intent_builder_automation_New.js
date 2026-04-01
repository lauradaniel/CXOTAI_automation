/**
 * Intent Builder UI Automation Script
 * =====================================
 * Run this script in the browser console while on the Intent Builder page.
 *
 * Usage:
 *   1. Open the Intent Builder application in your browser.
 *   2. Open DevTools (F12) → Console tab.
 *   3. Paste this entire script and press Enter.
 *   4. A file picker will open — select your CSV file.
 *   5. The script will process each row automatically.
 *
 * CSV Format (required columns):
 *   Category, Topic, Intent, Intent Percentage, Volume, Examples, Tag, Action, Required Change
 *
 * Supported Actions:
 *   Rename  — renames the item to the value in "Required Change"
 *   Remove  — removes the item (confirms the popup)
 *   Move    — moves to destination in "Required Change" (format: "Category" or "Category > Topic" or "Category > Topic > Intent")
 *   Merge   — merges with target in "Required Change" (same format as Move)
 */

(async function intentBuilderAutomation() {

  // ─────────────────────────────────────────────
  // CONFIGURATION
  // ─────────────────────────────────────────────
  const CFG = {
    shortDelay:  400,
    actionDelay: 700,
    dialogDelay: 1200,
    waitTimeout: 8000,
    scrollDelay: 300,
  };


  // ─────────────────────────────────────────────
  // LOGGER
  // ─────────────────────────────────────────────
  const LOG_STYLES = {
    info:    'color:#2196F3;font-weight:bold',
    success: 'color:#4CAF50;font-weight:bold',
    warn:    'color:#FF9800;font-weight:bold',
    error:   'color:#F44336;font-weight:bold',
  };
  function log(msg, type = 'info') {
    console.log(`%c[IB-Auto] ${msg}`, LOG_STYLES[type]);
  }


  // ─────────────────────────────────────────────
  // UTILITIES
  // ─────────────────────────────────────────────
  const sleep = ms => new Promise(r => setTimeout(r, ms));

  function waitFor(selectorFn, timeout = CFG.waitTimeout) {
    return new Promise((resolve, reject) => {
      const deadline = Date.now() + timeout;
      const iv = setInterval(() => {
        const el = selectorFn();
        if (el) { clearInterval(iv); resolve(el); }
        else if (Date.now() > deadline) {
          clearInterval(iv);
          reject(new Error('waitFor timeout'));
        }
      }, 100);
    });
  }

  function setInputValue(input, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input',  { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  async function hoverOver(el) {
    el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mouseover',  { bubbles: true }));
    await sleep(CFG.shortDelay);
  }

  async function scrollIntoView(el) {
    el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    await sleep(CFG.scrollDelay);
  }

  async function tryCloseDialog() {
    // Covers both the main tree dialogs and the Move/Merge cxone-modal
    const btn = document.querySelector([
      'i.close-button',
      'button.cancel-btn',
      'p-dialog .p-dialog-header-close',
      '.p-dialog .p-dialog-header-close',
      '[role="dialog"] button[aria-label="Close"]',
    ].join(','));
    if (btn) { btn.click(); await sleep(CFG.dialogDelay); }
  }


  // ─────────────────────────────────────────────
  // CSV PARSING  (RFC 4180 — handles multi-line quoted fields)
  // ─────────────────────────────────────────────
  function parseCSV(text) {
    const records = [];
    let field = '', fields = [], inQuotes = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i], next = text[i + 1];
      if (inQuotes) {
        if (ch === '"' && next === '"') { field += '"'; i++; }
        else if (ch === '"')            { inQuotes = false; }
        else                            { field += ch; }
      } else {
        if      (ch === '"')                   { inQuotes = true; }
        else if (ch === ',')                   { fields.push(field.trim()); field = ''; }
        else if (ch === '\r' && next === '\n') {
          fields.push(field.trim()); field = '';
          if (fields.some(f => f !== '')) records.push(fields);
          fields = []; i++;
        } else if (ch === '\n' || ch === '\r') {
          fields.push(field.trim()); field = '';
          if (fields.some(f => f !== '')) records.push(fields);
          fields = [];
        } else { field += ch; }
      }
    }
    if (field || fields.length > 0) {
      fields.push(field.trim());
      if (fields.some(f => f !== '')) records.push(fields);
    }
    if (records.length < 2) throw new Error('CSV has no data rows');
    const headers = records[0].map(h => h.replace(/^"|"$/g, '').trim());
    return records.slice(1).map(cols => {
      const row = {};
      headers.forEach((h, i) => { row[h] = (cols[i] || '').replace(/^"|"$/g, '').trim(); });
      return row;
    });
  }


  // ─────────────────────────────────────────────
  // FILE PICKER
  // ─────────────────────────────────────────────
  function pickCSVFile() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file'; input.accept = '.csv,text/csv'; input.style.display = 'none';
      document.body.appendChild(input);
      input.onchange = e => {
        const file = e.target.files[0];
        if (!file) { reject(new Error('No file selected')); return; }
        const reader = new FileReader();
        reader.onload  = ev => { document.body.removeChild(input); resolve(ev.target.result); };
        reader.onerror = ()  => { document.body.removeChild(input); reject(new Error('FileReader error')); };
        reader.readAsText(file);
      };
      input.click();
    });
  }


  // ─────────────────────────────────────────────
  // MAIN TREE NAVIGATION  (PrimeNG p-tree)
  // ─────────────────────────────────────────────

  const normalize = s => (s || '').trim().toLowerCase();

  // --- 1. Get the exact name using the built-in accessibility label ---
  function getItemName(li) {
    // PrimeNG stamps the exact text into the aria-label
    let name = li.getAttribute('aria-label');
    if (name) return name;

    // Fallback just in case
    const nameDiv = li.querySelector('.kanban-tree-node-name');
    if (nameDiv) return nameDiv.textContent;

    return '';
  }

  // --- 2. Find the item using robust string matching ---
  function findItem(name, level, parentLi = null) {
    const cleanString = (str) => str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const targetName = cleanString(name);

    let scope = parentLi ? parentLi : document;

    // Look for PrimeNG tree items at the correct depth
    const elements = scope.querySelectorAll(`li[role="treeitem"][aria-level="${level}"]`);

    for (const li of elements) {
      const rawName = getItemName(li);
      if (rawName) {
          const currentName = cleanString(rawName);
          if (currentName === targetName || currentName.startsWith(targetName)) {
              return li;
          }
      }
    }
    return null;
  }

  // --- 3. Locate the folder where children are kept ---
  function getChildrenContainer(parentLi) {
    // PrimeNG stores children in a <ul> immediately inside the <li>
    return parentLi.querySelector('ul.p-treenode-children, [role="group"]');
  }

  // --- 4. Expand the folder if it's closed ---
  async function expandItem(li) {
    // Check if PrimeNG already marks it as expanded
    const isExpanded = li.getAttribute('aria-expanded') === 'true';
    if (isExpanded) return;

    // Find the specific PrimeNG toggler arrow
    const toggler = li.querySelector('.p-tree-toggler');
    if (toggler) {
      toggler.click();
      await sleep(CFG.actionDelay);
    }
  }

  async function locateTarget(row) {
    const category = row['Category'];
    const topic    = row['Topic'];
    const intent   = row['Intent'];

    const catItem = await waitFor(() => findItem(category, 1), CFG.waitTimeout).catch(() => null);
    if (!catItem) { log(`Category not found: "${category}"`, 'error'); return null; }
    if (!topic) return catItem;

    await expandItem(catItem);
    const topicItem = await waitFor(() => findItem(topic, 2, catItem), CFG.waitTimeout).catch(() => null);
    if (!topicItem) { log(`Topic not found: "${topic}" under "${category}"`, 'error'); return null; }
    if (!intent) return topicItem;

    await expandItem(topicItem);
    const intentItem = await waitFor(() => findItem(intent, 3, topicItem), CFG.waitTimeout).catch(() => null);
    if (!intentItem) { log(`Intent not found: "${intent}" under "${topic}" > "${category}"`, 'error'); return null; }

    await scrollIntoView(intentItem);
    return intentItem;
  }


  // ─────────────────────────────────────────────
  // MAIN TREE MENU HELPERS
  // ─────────────────────────────────────────────

  async function openMoreMenu(li) {
    await hoverOver(li);
    const btn = await waitFor(() => li.querySelector('button.kanban-more-button'));
    btn.click();
    await sleep(CFG.actionDelay);
  }

  async function clickMenuOption(label) {
    const target = normalize(label);
    const el = await waitFor(() => {
      const containers = document.querySelectorAll([
        'p-overlaypanel', '.p-overlaypanel', '.p-menu', '.p-contextmenu',
        '[role="menu"]', '.cxone-menu', '.dropdown-menu', '.context-menu',
      ].join(','));
      for (const c of containers) {
        if (!c.offsetParent && c.style.display === 'none') continue;
        for (const item of c.querySelectorAll('li, [role="menuitem"], button, a')) {
          if (normalize(item.textContent).includes(target)) return item;
        }
      }
      for (const el of document.querySelectorAll('li, [role="menuitem"], button.p-menuitem-link')) {
        if (normalize(el.textContent).trim() === target && el.offsetParent !== null) return el;
      }
      return null;
    });
    el.click();
    await sleep(CFG.actionDelay);
  }

  async function clickDialogButton(...labels) {
    const targets = labels.map(normalize);
    const btn = await waitFor(() => {
      for (const d of document.querySelectorAll('p-dialog,.p-dialog,[role="dialog"],mat-dialog-container')) {
        for (const b of d.querySelectorAll('button')) {
          if (targets.includes(normalize(b.textContent).trim()) && !b.disabled) return b;
        }
      }
      return null;
    });
    btn.click();
    await sleep(CFG.dialogDelay);
  }


  // ─────────────────────────────────────────────
  // MOVE / MERGE — AG-GRID POPUP NAVIGATION
  // ─────────────────────────────────────────────

  /** Returns the display name of an AG-Grid row. */
  function getGridRowName(row) {
    const el = row.querySelector('span[data-aid="ellipsis-sliced-text"]');
    return el ? normalize(el.textContent) : '';
  }

  /** True if the row has children (is a group node). */
  function isGridRowExpandable(row) {
    return row.classList.contains('ag-row-group');
  }

  /** True if the row is currently expanded. */
  function isGridRowExpanded(row) {
    return row.getAttribute('aria-expanded') === 'true';
  }

  /** Expands an AG-Grid row by clicking its contracted icon. */
  async function expandGridRow(row) {
    if (!isGridRowExpandable(row) || isGridRowExpanded(row)) return;
    const icon = row.querySelector('.ag-group-contracted:not(.ag-hidden)');
    if (!icon) return;
    icon.click();
    await waitFor(
      () => row.getAttribute('aria-expanded') === 'true' ? row : null,
      5000
    ).catch(() => null);
    await sleep(300);
  }

  /** Returns all currently-rendered AG-Grid rows at a specific level. */
  function getGridRows(level) {
    return Array.from(
      document.querySelectorAll(`[role="treegrid"] [role="row"].ag-row-level-${level}`)
    );
  }

  /** Finds a visible (DOM-rendered) AG-Grid row by name and level. */
  function findGridRow(name, level) {
    const n = normalize(name);
    for (const row of getGridRows(level)) {
      if (getGridRowName(row) === n) return row;
    }
    return null;
  }

  async function scrollGridToFind(name, level, maxScrollSteps = 40) {
    let row = findGridRow(name, level);
    if (row) return row;

    const viewport =
      document.querySelector('[role="treegrid"] .ag-body-viewport') ||
      document.querySelector('[role="treegrid"] .ag-center-cols-viewport') ||
      document.querySelector('.ag-body-viewport');

    if (!viewport) return null;

    viewport.scrollTop = 0;
    await sleep(150);

    const step = Math.max(80, Math.floor(viewport.clientHeight * 0.4));

    for (let i = 1; i <= maxScrollSteps; i++) {
      viewport.scrollTop = step * i;
      await sleep(120);
      row = findGridRow(name, level);
      if (row) return row;
      if (viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 5) break;
    }

    viewport.scrollTop = 0;
    await sleep(200);
    return findGridRow(name, level);
  }

  /**
   * Opens the Move/Merge popup, navigates to the destination in the AG-Grid tree,
   * and clicks Save. Used by both performMove and performMerge.
   *
   * destination formats:
   *   "Category"                   → click the Level-0 category row
   *   "Category > Topic"           → expand category, click Level-1 topic row
   *   "Category > Topic > Intent"  → expand both, click Level-2 intent row
   */
  async function executeMoveMerge(targetLi, changePath) {
    // 1. Open the More Actions menu for the selected item
    await hoverOver(targetLi);
    const moreBtn = targetLi.querySelector('.kanban-more-button, .actions-renderer-container .more-icon-container, use[href="#icon-more"]');
    if (!moreBtn) throw new Error('More Actions button not found on target item');

    const clickTarget = moreBtn.closest('button') || moreBtn.closest('.more-icon-container') || moreBtn;
    clickTarget.click();
    await sleep(CFG.actionDelay);

    // 2. Select Move/Merge option in the dropdown
    const allTextElements = Array.from(document.querySelectorAll('*')).filter(
      el => el.children.length === 0 && el.textContent.trim() === 'Move/Merge'
    );
    const moveMergeOpt = allTextElements.find(el => el.offsetWidth > 0 && el.offsetHeight > 0);

    if (moveMergeOpt) {
      moveMergeOpt.click();
      if (document.activeElement) document.activeElement.blur();
    } else {
      throw new Error('Move/Merge option not found in the dropdown menu');
    }

    await sleep(CFG.dialogDelay);

    // 3. Wait for the AG-Grid/Tree popup modal
    const modal = await waitFor(
      () => document.querySelector('move-to-modal, cxone-modal, .p-dialog'),
      CFG.waitTimeout
    );
    if (!modal) throw new Error('Move/Merge popup did not open');

    // 4. Parse the target destination
    const parts = changePath.split('>').map(s => s.trim());
    const targetNodeName = parts[parts.length - 1];

    // 5. Helper: find the name span for a node
    const findTargetSpan = (name) => {
      let spans = Array.from(modal.querySelectorAll(
        'span[data-aid="ellipsis-sliced-text"], .kanban-tree-node-name, .p-treenode-label'
      ));
      let match = spans.find(el => el.textContent.trim().toLowerCase() === name.toLowerCase());
      if (!match) {
        match = spans.find(el =>
          el.textContent.trim().toLowerCase().includes(name.toLowerCase().substring(0, 15))
        );
      }
      return match;
    };

    // 6. Navigate the AG-Grid tree: expand rows until target is found
    let nodeSpan = null;
    let hitEnd = false;
    let attempts = 0;
    const maxAttempts = 200;

    let viewport =
      modal.querySelector('.ag-body-viewport, .ag-center-cols-viewport, .p-dialog-content') ||
      modal;
    viewport.scrollTop = 0;
    await sleep(300);

    while (!nodeSpan && !hitEnd && attempts < maxAttempts) {
      attempts++;

      nodeSpan = findTargetSpan(targetNodeName);
      if (nodeSpan) break;

      let expanders = Array.from(modal.querySelectorAll(
        '.ag-group-contracted:not(.ag-hidden), .p-tree-toggler[aria-expanded="false"]'
      ));
      if (expanders.length > 0) {
        for (let expander of expanders) {
          expander.click();
          await sleep(150);
        }
        await sleep(400);
        continue;
      }

      let beforeScroll = viewport.scrollTop;
      let scrollAmount = viewport.clientHeight > 0 ? (viewport.clientHeight - 40) : 300;
      viewport.scrollTop += scrollAmount;
      await sleep(400);

      if (viewport.scrollTop > beforeScroll) continue;

      let nextBtn = modal.querySelector('.ag-paging-button[aria-label="Next Page"]');
      if (
        nextBtn &&
        !nextBtn.classList.contains('ag-disabled') &&
        nextBtn.getAttribute('aria-disabled') !== 'true' &&
        !nextBtn.disabled
      ) {
        nextBtn.click();
        await sleep(1500);
        viewport.scrollTop = 0;
        continue;
      }

      hitEnd = true;
    }

    if (!nodeSpan) {
      throw new Error(
        `Destination "${targetNodeName}" not found in the Move/Merge tree after exhaustive search.`
      );
    }

    // 7. Select the target row with a SINGLE click.
    //    Multiple clicks (row + cell + span) can toggle the selection off in AG-Grid.
    const row = nodeSpan.closest('[role="row"], .ag-row, .p-treenode-content');
    if (row) {
      row.scrollIntoView({ behavior: 'instant', block: 'center' });
      await sleep(200);
      row.click();
    } else {
      // Fallback: click the span itself
      nodeSpan.scrollIntoView({ behavior: 'instant', block: 'center' });
      await sleep(200);
      nodeSpan.click();
    }

    // 8. Wait for the Save button to become enabled (Angular removes `disabled`
    //    attribute only after processing the row-selection change event).
    log('  Waiting for Save button to become enabled...');
    const saveBtn = await waitFor(() => {
      const btn =
        modal.querySelector('button.save-btn') ||
        modal.querySelector('.modal-footer-wrapper button:not(.cancel-btn)');
      return (btn && !btn.disabled) ? btn : null;
    }, 6000).catch(() => null);

    if (!saveBtn) {
      throw new Error(
        'Save button did not become enabled. Row selection may not have registered — try again.'
      );
    }

    saveBtn.scrollIntoView({ block: 'center' });
    await sleep(150);
    saveBtn.click();
    log('  Save clicked.');
    await sleep(CFG.dialogDelay);
  }


  // ─────────────────────────────────────────────
  // ACTION HANDLERS
  // ─────────────────────────────────────────────

  async function performRename(li, newName) {
    await openMoreMenu(li);
    await clickMenuOption('Rename');
    const input = await waitFor(() => document.querySelector([
      'p-dialog input[type="text"]', '.p-dialog input[type="text"]',
      '[role="dialog"] input[type="text"]', 'mat-dialog-container input[type="text"]',
    ].join(',')));
    input.focus(); input.select();
    setInputValue(input, newName);
    await sleep(CFG.shortDelay);
    await clickDialogButton('rename', 'save', 'confirm', 'ok');
    log(`  Renamed → "${newName}"`, 'success');
  }

  async function performRemove(li) {
    await openMoreMenu(li);
    await clickMenuOption('Remove');
    await clickDialogButton('remove', 'delete', 'confirm', 'yes');
    log('  Removed', 'success');
  }

  async function performMove(targetLi, change) {
    await executeMoveMerge(targetLi, change);
    log(`  Moved → "${change}"`, 'success');
  }

  async function performMerge(targetLi, change, newName) {
    // Step 1: open popup, navigate to destination, click Save
    await executeMoveMerge(targetLi, change);

    // Step 2: wait for the second Merge confirmation/rename popup
    await sleep(CFG.dialogDelay);

    const secondModal = await waitFor(() => {
      const modals = Array.from(document.querySelectorAll(
        'cxone-modal, merge-modal, .p-dialog, [role="dialog"]'
      ));
      return modals.find(m => {
        const hasInput = m.querySelector('input:not([type="hidden"]), textarea');
        const isVisible = m.offsetParent !== null || m.style.display !== 'none';
        const isNotTreeModal = !m.querySelector('.p-treenode, [role="treeitem"]');
        return hasInput && isVisible && isNotTreeModal;
      });
    }, CFG.waitTimeout).catch(() => null);

    if (!secondModal) {
      throw new Error('Second Merge popup (rename step) did not open or is unrecognizable');
    }

    // Step 3: fill in the new merged intent name
    if (newName) {
      const nameInput = secondModal.querySelector(
        'textarea, input[type="text"], input:not([type="hidden"])'
      );
      if (!nameInput) throw new Error('Text field not found in the second Merge popup');

      nameInput.focus();
      const nativeSetter =
        Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set ||
        Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set;

      if (nativeSetter) {
        nativeSetter.call(nameInput, newName);
      } else {
        nameInput.value = newName;
      }
      nameInput.dispatchEvent(new Event('input',  { bubbles: true }));
      nameInput.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(CFG.shortDelay);
    }

    // Step 4: click the final Merge/Confirm button
    let mergeBtn = Array.from(secondModal.querySelectorAll('button')).find(btn => {
      const text = btn.textContent.trim().toLowerCase();
      return text.match(/^(merge|save|confirm|submit|yes)$/i) &&
             !btn.classList.contains('cancel-btn') &&
             !btn.disabled;
    });

    if (!mergeBtn) {
      const footerBtns = Array.from(
        secondModal.querySelectorAll('.modal-footer-wrapper button, .p-dialog-footer button')
      );
      mergeBtn = footerBtns.find(b => !b.classList.contains('cancel-btn') && !b.disabled);
    }

    if (!mergeBtn) throw new Error('Confirm button not found in the second Merge popup');

    mergeBtn.click();
    await sleep(CFG.dialogDelay);
    log(`  Merged → "${change}"`, 'success');
  }


  // ─────────────────────────────────────────────
  // MAIN RUNNER
  // ─────────────────────────────────────────────
  async function run() {
    log('=== Intent Builder Automation Started ===');

    let csvText;
    try { csvText = await pickCSVFile(); }
    catch (e) { log('File selection failed: ' + e.message, 'error'); return; }

    let rows;
    try { rows = parseCSV(csvText); }
    catch (e) { log('CSV parse error: ' + e.message, 'error'); return; }
    log(`CSV loaded — ${rows.length} data rows`);

    let success = 0, skipped = 0, errors = 0;

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const rowNum = i + 2;

      // Helper to safely extract column values — ignores hidden chars/spaces from Excel
      const getCol = (colName) => {
        for (let key in row) {
          let cleanKey = key.replace(/[^a-zA-Z0-9 ]/g, '').trim().toLowerCase();
          if (cleanKey.includes(colName.toLowerCase())) {
            return row[key] ? row[key].trim() : '';
          }
        }
        return '';
      };

      const action = getCol('recommended change') || row['Action'] || '';
      const label  = row['Intent'] || row['Topic'] || row['Category'] || '(unknown)';

      if (!action) {
        log(`Row ${rowNum}: No action — skipping "${label}"`, 'warn');
        skipped++; continue;
      }

      log(`Row ${rowNum}: [${action}] "${label}"`);

      try {
        const target = await locateTarget(row);
        if (!target) { errors++; continue; }

        const renameTarget = getCol('updated intent name');
        const moveTarget   = getCol('move to topic');
        const mergeTarget  = getCol('merge to intent');

        switch (action.toLowerCase()) {
          case 'rename':
            if (!renameTarget) throw new Error('Column "Updated Intent Name" is empty');
            await performRename(target, renameTarget); break;

          case 'deactivate':
          case 'remove':
            await performRemove(target); break;

          case 'move':
            if (!moveTarget) throw new Error('Column "Move to Topic" is empty');
            await performMove(target, moveTarget); break;

          case 'merge':
            if (!mergeTarget) throw new Error('Column "Merge to Intent" is empty');
            if (!renameTarget) log(`Row ${rowNum}: Warning — "Updated Intent Name" is empty for merge rename step.`, 'warn');
            await performMerge(target, mergeTarget, renameTarget); break;

          default:
            log(`Row ${rowNum}: Unknown action "${action}" — skipping`, 'warn');
            skipped++; continue;
        }
        success++;

      } catch (e) {
        log(`Row ${rowNum}: ERROR — ${e.message}`, 'error');
        errors++;
        await tryCloseDialog();
      }

      await sleep(CFG.shortDelay);
    }

    log(
      `=== Done — Success: ${success} | Skipped: ${skipped} | Errors: ${errors} ===`,
      errors === 0 ? 'success' : 'warn'
    );
  }

  await run();

})();
