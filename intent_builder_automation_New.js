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

 // --- 1. Get the display name of a tree node ---
  function getItemName(li) {
    // PrimeNG stamps the exact text into aria-label
    const ariaLabel = li.getAttribute('aria-label');
    if (ariaLabel) return ariaLabel.trim();

    // Known class names used by this app
    for (const sel of ['.kanban-tree-node-name', '.p-treenode-label', '[class*="node-name"]', '[class*="tree-label"]']) {
      const el = li.querySelector(sel);
      if (el) return el.textContent.trim();
    }

    // Last resort: text inside .p-treenode-content but strip child-list text
    const content = li.querySelector('.p-treenode-content');
    if (content) {
      const clone = content.cloneNode(true);
      clone.querySelectorAll('ul').forEach(u => u.remove());
      const t = clone.textContent.trim();
      if (t) return t;
    }

    return '';
  }

  // --- 1b. Determine the nesting level of a tree node ---
  function getItemLevel(li) {
    // Prefer the explicit attribute when present
    const ariaLevel = parseInt(li.getAttribute('aria-level'), 10);
    if (!isNaN(ariaLevel) && ariaLevel > 0) return ariaLevel;

    // Fall back to counting treeitem ancestors
    let level = 1;
    let el = li.parentElement;
    while (el) {
      if (el.matches('li[role="treeitem"]')) level++;
      el = el.parentElement;
    }
    return level;
  }

  // --- 2. Find the item using robust string matching ---
  function findItem(name, level, parentLi = null) {
    const cleanString = (str) => str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const targetName = cleanString(name);

    const scope = parentLi || document;

    // Try the fast aria-level selector first
    let elements = Array.from(scope.querySelectorAll(`li[role="treeitem"][aria-level="${level}"]`));

    // If nothing found with aria-level (e.g. fresh page load where attribute is absent),
    // fall back to all treeitems and filter by computed depth
    if (elements.length === 0) {
      elements = Array.from(scope.querySelectorAll('li[role="treeitem"]'))
                      .filter(li => getItemLevel(li) === level);
    }

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
  //
  // The Move/Merge dialog uses AG-Grid (not p-tree).
  // Rows: div[role="row"] with class ag-row-level-0 / ag-row-level-1 / ag-row-level-2
  // Name: span[data-aid="ellipsis-sliced-text"] inside each row
  // Expand: click .ag-group-contracted span inside the row
  // Save:   button.save-btn  (disabled until a row is clicked)
  //
  // AG-Grid uses virtual scrolling: rows outside the viewport are NOT in the DOM.
  // scrollGridToFind() handles this by scrolling the AG-Grid body viewport.
  //
  // "Required Change" format:
  //   "Category"                    — move/merge into a category
  //   "Category > Topic"            — move/merge into a topic
  //   "Category > Topic > Intent"   — move/merge into an intent
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

  /**
   * Scrolls the AG-Grid body viewport incrementally to find a row.
   * AG-Grid virtual scrolling only renders rows in the visible area;
   * this function scrolls through the entire list to expose all rows.
   *
   * Returns the matching row element, or null if not found.
   */
  async function scrollGridToFind(name, level, maxScrollSteps = 40) {
    // Check already-rendered rows first (fast path)
    let row = findGridRow(name, level);
    if (row) return row;

    // Locate the AG-Grid scrollable viewport inside the dialog
    const viewport =
      document.querySelector('[role="treegrid"] .ag-body-viewport') ||
      document.querySelector('[role="treegrid"] .ag-center-cols-viewport') ||
      document.querySelector('.ag-body-viewport');

    if (!viewport) {
      log('  AG-Grid viewport not found, cannot scroll', 'warn');
      return null;
    }

    // Reset to top before scanning
    viewport.scrollTop = 0;
    await sleep(150);

    const step = Math.max(80, Math.floor(viewport.clientHeight * 0.4));

    for (let i = 1; i <= maxScrollSteps; i++) {
      viewport.scrollTop = step * i;
      await sleep(120); // Let virtual scroll render new rows

      row = findGridRow(name, level);
      if (row) return row;

      // Stop if we've reached the bottom
      if (viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 5) break;
    }

    // Scroll back to top and do one final check
    viewport.scrollTop = 0;
    await sleep(200);
    return findGridRow(name, level);
  }

  /**
   * Navigates the AG-Grid in the Move/Merge dialog and clicks the destination row.
   *
   * destination formats:
   *   "Category"                   → click the Level-0 category row
   *   "Category > Topic"           → expand category, click Level-1 topic row
   *   "Category > Topic > Intent"  → expand both, click Level-2 intent row
   */
  async function selectDestination(destination) {
    await sleep(CFG.dialogDelay);

    // Wait for the AG-Grid treegrid to appear inside the modal
    await waitFor(
      () => document.querySelector('[role="treegrid"]'),
      CFG.waitTimeout
    );
    await sleep(CFG.shortDelay);

    const parts = destination.split('>').map(s => s.trim()).filter(Boolean);
    log(`  Navigating dialog to: ${parts.join(' > ')}`);

    let target;

    if (parts.length === 1) {
      // ── Single name: scan level 0, then 1, then 2 ──────────────────────
      target = await scrollGridToFind(parts[0], 0);
      if (!target) target = await scrollGridToFind(parts[0], 1);
      if (!target) target = await scrollGridToFind(parts[0], 2);
      if (!target) throw new Error(`Move/Merge dialog: "${destination}" not found`);

      log(`  Found "${parts[0]}", clicking...`);
      await scrollIntoView(target);
      target.click();

    } else if (parts.length === 2) {
      // ── Category > Topic ─────────────────────────────────────────────
      const l0 = await scrollGridToFind(parts[0], 0);
      if (!l0) throw new Error(`Move/Merge dialog: Category "${parts[0]}" not found`);
      log(`  Found Category "${parts[0]}", expanding...`);
      await scrollIntoView(l0);
      await expandGridRow(l0);

      // Children should now be rendered near l0; scroll to find if needed
      const l1 =
        await waitFor(() => findGridRow(parts[1], 1), 3000).catch(() => null) ||
        await scrollGridToFind(parts[1], 1);
      if (!l1) throw new Error(`Move/Merge dialog: Topic "${parts[1]}" not found under "${parts[0]}"`);

      log(`  Found Topic "${parts[1]}", clicking...`);
      await scrollIntoView(l1);
      l1.click();

    } else {
      // ── Category > Topic > Intent ─────────────────────────────────────
      const l0 = await scrollGridToFind(parts[0], 0);
      if (!l0) throw new Error(`Move/Merge dialog: Category "${parts[0]}" not found`);
      log(`  Found Category "${parts[0]}", expanding...`);
      await scrollIntoView(l0);
      await expandGridRow(l0);

      const l1 =
        await waitFor(() => findGridRow(parts[1], 1), 3000).catch(() => null) ||
        await scrollGridToFind(parts[1], 1);
      if (!l1) throw new Error(`Move/Merge dialog: Topic "${parts[1]}" not found`);
      log(`  Found Topic "${parts[1]}", expanding...`);
      await scrollIntoView(l1);
      await expandGridRow(l1);

      const l2 =
        await waitFor(() => findGridRow(parts[2], 2), 3000).catch(() => null) ||
        await scrollGridToFind(parts[2], 2);
      if (!l2) throw new Error(`Move/Merge dialog: Intent "${parts[2]}" not found`);
      log(`  Found Intent "${parts[2]}", clicking...`);
      await scrollIntoView(l2);
      l2.click();
    }

    await sleep(CFG.shortDelay);

    // Save button becomes enabled once a row is selected
    const saveBtn = await waitFor(
      () => {
        const btn = document.querySelector('button.save-btn');
        return (btn && !btn.disabled) ? btn : null;
      },
      5000
    );
    saveBtn.click();
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
  }

async function performMerge(targetLi, change, newName) {
    // 1. Execute the first step (finding the branch and clicking Save)
    await executeMoveMerge(targetLi, change);

    // 2. Wait for the second Merge confirmation popup
    await sleep(CFG.dialogDelay);
    
    const secondModal = await waitFor(() => {
        // Look for any dialog that has a text input and is NOT the tree menu
        const modals = Array.from(document.querySelectorAll('cxone-modal, merge-modal, .p-dialog, [role="dialog"]'));
        return modals.find(m => {
            const hasInput = m.querySelector('input:not([type="hidden"]), textarea');
            const isVisible = m.offsetParent !== null || m.style.display !== 'none';
            const isNotTreeModal = !m.querySelector('.p-treenode, [role="treeitem"]'); // Ensure we moved past the first modal
            return hasInput && isVisible && isNotTreeModal;
        });
    }, CFG.waitTimeout);

    if (!secondModal) throw new Error('Second Merge popup (for renaming) did not open or is unrecognizable');

    // 3. Enter the New Name
    if (newName) {
        const nameInput = secondModal.querySelector('textarea, input[type="text"], input:not([type="hidden"])');
        if (!nameInput) throw new Error('Could not find text area for New Name in the second Merge popup');

        nameInput.focus();
        nameInput.value = '';
        nameInput.dispatchEvent(new Event('input', { bubbles: true }));
        await sleep(200);

        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set || 
                                       Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value")?.set;
        
        if (nativeInputValueSetter) {
            nativeInputValueSetter.call(nameInput, newName);
        } else {
            nameInput.value = newName;
        }

        nameInput.dispatchEvent(new Event('input', { bubbles: true }));
        nameInput.dispatchEvent(new Event('change', { bubbles: true }));
        nameInput.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Enter' }));
        
        await sleep(CFG.shortDelay);
    }

    // 4. Click the Final Merge/Confirm button
    let mergeBtn = Array.from(secondModal.querySelectorAll('button')).find(btn => {
        const text = btn.textContent.trim().toLowerCase();
        return text.match(/^(merge|save|confirm|submit|yes)$/i) && !btn.classList.contains('cancel-btn');
    });

    if (!mergeBtn) {
        const footerBtns = Array.from(secondModal.querySelectorAll('.modal-footer-wrapper button, .p-dialog-footer button'));
        if (footerBtns.length > 0) mergeBtn = footerBtns[footerBtns.length - 1];
    }

    if (!mergeBtn) throw new Error('Confirmation button not found in the second popup');

    if (!mergeBtn.disabled && !mergeBtn.classList.contains('p-disabled') && mergeBtn.getAttribute('aria-disabled') !== 'true') {
        mergeBtn.focus();
        mergeBtn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        mergeBtn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        mergeBtn.click();
    } else {
        throw new Error('Merge button in second popup is disabled.');
    }

    await sleep(CFG.dialogDelay);
  }

async function executeMoveMerge(targetLi, changePath) {
    // 1. Open the More Actions menu for the selected item
    await hoverOver(targetLi);
    const moreBtn = targetLi.querySelector('.kanban-more-button, .actions-renderer-container .more-icon-container, use[href="#icon-more"]');
    if (!moreBtn) throw new Error('More Actions button not found on target item');
    
    const clickTarget = moreBtn.closest('button') || moreBtn.closest('.more-icon-container') || moreBtn;
    clickTarget.click();
    await sleep(CFG.actionDelay);

    // 2. Select Move/Merge option in the dropdown
    const allTextElements = Array.from(document.querySelectorAll('*')).filter(el => el.children.length === 0 && el.textContent.trim() === 'Move/Merge');
    const moveMergeOpt = allTextElements.find(el => el.offsetWidth > 0 && el.offsetHeight > 0);
    
    if (moveMergeOpt) {
        moveMergeOpt.click();
        if (document.activeElement) document.activeElement.blur(); // Clears accessibility warning
    } else {
        throw new Error('Move/Merge option not found in the dropdown menu');
    }
    
    await sleep(CFG.dialogDelay);

    // 3. Wait for the AG-Grid/Tree popup modal
    const modal = await waitFor(() => document.querySelector('move-to-modal, cxone-modal, .p-dialog'), CFG.waitTimeout);
    if (!modal) throw new Error('Move/Merge popup did not open');

    // 4. Parse the target destination from CSV
    const parts = changePath.split('>').map(s => s.trim());
    const targetNodeName = parts[parts.length - 1];

    // Helper to find the specific text span in the grid
    const findTargetSpan = (name) => {
        let spans = Array.from(modal.querySelectorAll('span[data-aid="ellipsis-sliced-text"], .kanban-tree-node-name, .p-treenode-label'));
        let match = spans.find(el => el.textContent.trim().toLowerCase() === name.toLowerCase());
        if (!match) {
            match = spans.find(el => el.textContent.trim().toLowerCase().includes(name.toLowerCase().substring(0, 15)));
        }
        return match;
    };

    // 5. Navigate the AG-Grid tree by scanning, expanding, scrolling, and paginating
    let nodeSpan = null;
    let hitEnd = false;
    let attempts = 0;
    const maxAttempts = 200;

    let viewport = modal.querySelector('.ag-body-viewport, .ag-center-cols-viewport, .p-dialog-content') || modal;
    viewport.scrollTop = 0; 
    await sleep(300);

    while (!nodeSpan && !hitEnd && attempts < maxAttempts) {
        attempts++;

        nodeSpan = findTargetSpan(targetNodeName);
        if (nodeSpan) break;

        let expanders = Array.from(modal.querySelectorAll('.ag-group-contracted:not(.ag-hidden), .p-tree-toggler[aria-expanded="false"]'));
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

        if (viewport.scrollTop > beforeScroll) {
            continue;
        }

        let nextBtn = modal.querySelector('.ag-paging-button[aria-label="Next Page"]');
        if (nextBtn && 
            !nextBtn.classList.contains('ag-disabled') && 
            nextBtn.getAttribute('aria-disabled') !== 'true' &&
            !nextBtn.disabled) {
            
            nextBtn.click();
            await sleep(1500); 
            viewport.scrollTop = 0; 
            continue;
        }

        hitEnd = true;
    }

    if (!nodeSpan) {
      throw new Error(`Destination branch "${targetNodeName}" not found in the tree after exhaustive scrolling.`);
    }

    // 6. Select the target node
    const row = nodeSpan.closest('[role="row"], .ag-row, .p-treenode-content');
    if (row) {
      row.scrollIntoView({ behavior: 'instant', block: 'center' });
      await sleep(CFG.shortDelay);

      row.click();
      await sleep(CFG.shortDelay);
      
      const cell = nodeSpan.closest('.ag-cell, .p-treenode-label');
      if (cell) cell.click();
      
      nodeSpan.click();
      await sleep(CFG.shortDelay);
    }

    // 7. Wait for the Primary Action Button (Save / Next / Move / Merge) to become enabled
    log('  Waiting for Save/Next button to become enabled...');
    const actionBtn = await waitFor(
      () => {
        let btn = modal.querySelector('.modal-footer-wrapper .save-btn') ||
                  modal.querySelector('.modal-footer-wrapper .btn-primary') ||
                  modal.querySelector('.p-dialog-footer .p-button-primary');

        if (!btn) {
          const buttons = Array.from(modal.querySelectorAll('button'));
          btn = buttons.find(b => {
            const text = b.textContent.trim().toLowerCase();
            const isCancel = b.classList.contains('cancel-btn') || text === 'cancel' || text === 'close';
            if (isCancel) return false;
            return b.classList.contains('save-btn') || b.classList.contains('btn-primary') ||
                   text.match(/^(save|merge|move|next|confirm|submit)$/i);
          });
        }

        if (!btn) return null;

        // Only return the button once it is actually enabled
        const isDisabled = btn.disabled ||
                           btn.classList.contains('p-disabled') ||
                           btn.classList.contains('disabled') ||
                           btn.getAttribute('aria-disabled') === 'true';
        return isDisabled ? null : btn;
      },
      8000
    ).catch(() => null);

    if (!actionBtn) throw new Error('Primary action button (Save/Next) not found or never became enabled after selecting the destination node.');

    actionBtn.scrollIntoView({ block: 'center' });
    await sleep(100);

    // Fire a full suite of pointer and mouse events on the inner span and the button itself
    const innerSpan = actionBtn.querySelector('span');
    const btnTarget = innerSpan || actionBtn;

    btnTarget.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
    btnTarget.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await sleep(50);
    btnTarget.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    btnTarget.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
    btnTarget.click();

    // Fallback native click on the button itself just in case the framework bound it there
    if (innerSpan) actionBtn.click();
    
    await sleep(CFG.dialogDelay);
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
    // Helper to safely extract column values ignoring hidden characters/spaces from Excel
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

        // Extract values from the newly named columns
        const renameTarget = getCol('updated intent name'); // Catches 'Updated Intent Name (L3)'
        const moveTarget   = getCol('move to topic');       // Catches 'Move to Topic (L2)'
        const mergeTarget  = getCol('merge to intent');     // Catches 'Merge to Intent (L3)'

        switch (action.toLowerCase()) {
          case 'rename':
            if (!renameTarget) throw new Error('Column "Updated Intent Name (L3)" is empty');
            await performRename(target, renameTarget); break;
            
          case 'deactivate':
          case 'remove': // Kept as a fallback just in case old files are used
            await performRemove(target); break;
            
          case 'move':
            if (!moveTarget) throw new Error('Column "Move to Topic (L2)" is empty');
            await performMove(target, moveTarget); break;
            
          case 'merge':
            if (!mergeTarget) throw new Error('Column "Merge to Intent (L3)" is empty');
            if (!renameTarget) log(`Row ${rowNum}: Warning — Could not find "Updated Intent Name (L3)" for the rename step.`, 'warn');
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