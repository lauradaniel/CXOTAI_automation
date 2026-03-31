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
 *   Move    — moves the item to the destination in "Required Change"
 *   Merge   — merges the item with the target in "Required Change"
 *
 * "Required Change" for Move/Merge supports two formats:
 *   Plain name:          "Payment Processing"          (searched across all levels)
 *   Path (recommended):  "BILLING & PAYMENT > Payment Processing"
 */

(async function intentBuilderAutomation() {

  // ─────────────────────────────────────────────
  // CONFIGURATION
  // ─────────────────────────────────────────────
  const CFG = {
    shortDelay:  400,
    actionDelay: 700,
    dialogDelay: 1200,
    waitTimeout: 10000,  // slightly longer for dialog tree
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
    const btn = document.querySelector([
      'p-dialog .p-dialog-header-close',
      '.p-dialog .p-dialog-header-close',
      '[role="dialog"] button[aria-label="Close"]',
      '[role="dialog"] button[aria-label="close"]',
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
  // TREE NAVIGATION  (shared by main tree and dialog tree)
  // ─────────────────────────────────────────────

  const normalize = s => (s || '').trim().toLowerCase();

  function getItemName(li) {
    const el = li.querySelector('.kanban-tree-node-name');
    return el ? normalize(el.textContent) : normalize(li.getAttribute('aria-label'));
  }

  /**
   * Returns the <ul> that holds the direct children of a treeitem.
   * Handles both standard HTML (ul inside li) and PrimeNG's sibling pattern
   * (ul is a sibling of li inside a <p-treenode> wrapper).
   */
  function getChildrenContainer(li) {
    const inner = li.querySelector(':scope > ul[role="group"], :scope > ul.p-treenode-children');
    if (inner) return inner;
    const wrapper = li.parentElement;
    if (wrapper) {
      const sibling = wrapper.querySelector(':scope > ul[role="group"], :scope > ul.p-treenode-children');
      if (sibling) return sibling;
    }
    return null;
  }

  /**
   * Finds a treeitem by name at the given aria-level.
   * - parentLi: scope search to that node's children container.
   * - rootScope: when parentLi is null, use this element instead of document
   *   (useful for scoping to a dialog).
   */
  function findItem(name, level, parentLi = null, rootScope = document) {
    const n = normalize(name);
    let scope;
    if (parentLi) {
      scope = getChildrenContainer(parentLi);
      if (!scope) return null;
    } else {
      scope = rootScope;
    }
    for (const li of scope.querySelectorAll(`li[role="treeitem"][aria-level="${level}"]`)) {
      if (getItemName(li) === n) return li;
    }
    return null;
  }

  /**
   * Expands a treeitem if not already open.
   * Uses `=== 'true'` so nodes with aria-expanded=null (never interacted with
   * on a fresh page load) are also expanded.
   */
  async function expandItem(li) {
    await scrollIntoView(li);
    if (li.getAttribute('aria-expanded') === 'true') return; // confirmed open
    const toggler = li.querySelector('button.p-tree-toggler');
    if (!toggler) return; // leaf node
    toggler.click();
    await waitFor(() => {
      const open        = li.getAttribute('aria-expanded') === 'true';
      const hasChildren = !!getChildrenContainer(li);
      return (open && hasChildren) ? li : null;
    }, CFG.waitTimeout).catch(() => null);
    await sleep(CFG.shortDelay);
  }

  /**
   * Navigates the main tree to find the target item for a CSV row.
   * Expands Category → Topic → Intent as needed.
   */
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
  // DIALOG TREE NAVIGATION  (Move / Merge popup)
  // ─────────────────────────────────────────────

  /**
   * Locates and clicks the destination node inside the Move/Merge dialog.
   *
   * The dialog contains the same PrimeNG tree (fully collapsed by default).
   * We expand it the same way we expand the main tree.
   *
   * destination format ("Required Change" column):
   *   Path (preferred): "CATEGORY > Topic"  or  "CATEGORY > Topic > Intent"
   *   Plain name:       "Topic Name"   (scans all levels until found)
   */
  async function selectDestinationInDialog(destination) {
    // Wait for the dialog and its tree to be ready
    const dialog = await waitFor(() =>
      document.querySelector('p-dialog, .p-dialog, [role="dialog"], mat-dialog-container'),
      CFG.waitTimeout
    );
    await sleep(CFG.shortDelay); // small buffer for tree to render

    const parts = destination.split('>').map(s => s.trim()).filter(Boolean);

    if (parts.length > 1) {
      // ── Path-based navigation ──
      await navigateDialogByPath(dialog, parts);
    } else {
      // ── Single name: expand tree level by level to find it ──
      await navigateDialogByName(dialog, destination);
    }
  }

  /**
   * Navigate using an explicit path array, e.g. ["CATEGORY", "Topic", "Intent"].
   * Clicks the last element in the path.
   */
  async function navigateDialogByPath(dialog, parts) {
    // Level 1
    const l1 = await waitFor(
      () => findItem(parts[0], 1, null, dialog), CFG.waitTimeout
    ).catch(() => null);
    if (!l1) throw new Error(`Destination L1 not found: "${parts[0]}"`);
    if (parts.length === 1) { await scrollIntoView(l1); l1.click(); return; }

    // Level 2
    await expandItem(l1);
    const l2 = await waitFor(
      () => findItem(parts[1], 2, l1), CFG.waitTimeout
    ).catch(() => null);
    if (!l2) throw new Error(`Destination L2 not found: "${parts[1]}"`);
    if (parts.length === 2) { await scrollIntoView(l2); l2.click(); return; }

    // Level 3
    await expandItem(l2);
    const l3 = await waitFor(
      () => findItem(parts[2], 3, l2), CFG.waitTimeout
    ).catch(() => null);
    if (!l3) throw new Error(`Destination L3 not found: "${parts[2]}"`);
    await scrollIntoView(l3); l3.click();
  }

  /**
   * Scan the dialog tree by expanding each level until the named item is found.
   * Checks L1 first, then expands each L1 to check L2, then L2 to check L3.
   */
  async function navigateDialogByName(dialog, name) {
    const n = normalize(name);

    // Wait for at least one Level-1 item to be present
    await waitFor(
      () => dialog.querySelector('li[role="treeitem"][aria-level="1"]'),
      CFG.waitTimeout
    );

    const l1Items = Array.from(dialog.querySelectorAll('li[role="treeitem"][aria-level="1"]'));

    for (const l1 of l1Items) {
      // Check Level 1
      if (getItemName(l1) === n) {
        await scrollIntoView(l1); l1.click(); return;
      }

      // Expand and check Level 2
      await expandItem(l1);
      const l1Container = getChildrenContainer(l1);
      if (!l1Container) continue;

      const l2Items = Array.from(l1Container.querySelectorAll('li[role="treeitem"][aria-level="2"]'));
      for (const l2 of l2Items) {
        if (getItemName(l2) === n) {
          await scrollIntoView(l2); l2.click(); return;
        }

        // Expand and check Level 3
        await expandItem(l2);
        const l2Container = getChildrenContainer(l2);
        if (!l2Container) continue;

        for (const l3 of l2Container.querySelectorAll('li[role="treeitem"][aria-level="3"]')) {
          if (getItemName(l3) === n) {
            await scrollIntoView(l3); l3.click(); return;
          }
        }
      }
    }

    throw new Error(`Destination "${name}" not found in Move/Merge dialog tree`);
  }


  // ─────────────────────────────────────────────
  // MENU HELPERS
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

  async function performMove(li, destination) {
    await openMoreMenu(li);
    await clickMenuOption('Move');
    await selectDestinationInDialog(destination);
    await clickDialogButton('save', 'move', 'confirm', 'ok');
    log(`  Moved → "${destination}"`, 'success');
  }

  async function performMerge(li, target) {
    await openMoreMenu(li);
    await clickMenuOption('Merge');
    await selectDestinationInDialog(target);
    await clickDialogButton('save', 'merge', 'confirm', 'ok');
    log(`  Merged → "${target}"`, 'success');
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
      const action = (row['Action'] || '').trim();
      const change = (row['Required Change'] || '').trim();
      const label  = row['Intent'] || row['Topic'] || row['Category'] || '(unknown)';

      if (!action) {
        log(`Row ${rowNum}: No action — skipping "${label}"`, 'warn');
        skipped++; continue;
      }

      log(`Row ${rowNum}: [${action}] "${label}"`);

      try {
        const target = await locateTarget(row);
        if (!target) { errors++; continue; }

        switch (action.toLowerCase()) {
          case 'rename':
            if (!change) throw new Error('"Required Change" is empty');
            await performRename(target, change); break;
          case 'remove':
            await performRemove(target); break;
          case 'move':
            if (!change) throw new Error('"Required Change" is empty');
            await performMove(target, change); break;
          case 'merge':
            if (!change) throw new Error('"Required Change" is empty');
            await performMerge(target, change); break;
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
