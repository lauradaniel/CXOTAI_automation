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
  // TREE NAVIGATION
  // ─────────────────────────────────────────────

  const normalize = s => (s || '').trim().toLowerCase();

  function getItemName(li) {
    const el = li.querySelector('.kanban-tree-node-name');
    return el ? normalize(el.textContent) : normalize(li.getAttribute('aria-label'));
  }

  /**
   * Returns the <ul> that holds the direct children of a treeitem.
   *
   * PrimeNG renders its tree like this:
   *
   *   <p-treenode>                         ← Angular wrapper
   *     <li role="treeitem">               ← the item we hold a ref to
   *       <div class="p-treenode-content">
   *     </li>
   *     <ul class="p-treenode-children">   ← SIBLING of li, NOT inside it
   *       <p-treenode>...
   *     </ul>
   *   </p-treenode>
   *
   * We check inside the li first (standard HTML), then the sibling pattern.
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
   * Finds a treeitem by name at the given level.
   * Scopes the search to the children container of parentLi when provided.
   */
  function findItem(name, level, parentLi = null) {
    const n = normalize(name);
    let scope;
    if (parentLi) {
      scope = getChildrenContainer(parentLi);
      if (!scope) return null;
    } else {
      scope = document;
    }
    for (const li of scope.querySelectorAll(`li[role="treeitem"][aria-level="${level}"]`)) {
      if (getItemName(li) === n) return li;
    }
    return null;
  }

  /**
   * Scrolls a treeitem into view, then expands it if it is not already open.
   *
   * KEY FIX: the skip condition is now `=== 'true'` (explicitly expanded).
   * Previously it was `!== 'false'`, which also returned early when the
   * attribute was null — i.e. nodes that had never been interacted with
   * on a fresh page load, causing all expansions to be silently skipped.
   */
  async function expandItem(li) {
    await scrollIntoView(li);

    // Only skip if the node is confirmed open. null / 'false' / absent → expand.
    if (li.getAttribute('aria-expanded') === 'true') return;

    const toggler = li.querySelector('button.p-tree-toggler');
    if (!toggler) return; // leaf node — nothing to expand

    toggler.click();

    // Wait until the attribute confirms expansion AND the children <ul>
    // exists in the DOM (PrimeNG adds it via *ngIf, not display:none).
    await waitFor(() => {
      const open        = li.getAttribute('aria-expanded') === 'true';
      const hasChildren = !!getChildrenContainer(li);
      return (open && hasChildren) ? li : null;
    }, CFG.waitTimeout).catch(() => null);

    await sleep(CFG.shortDelay);
  }

  /**
   * Resolves the correct treeitem for a CSV row by navigating
   * Category → Topic → Intent, expanding collapsed nodes as needed.
   */
  async function locateTarget(row) {
    const category = row['Category'];
    const topic    = row['Topic'];
    const intent   = row['Intent'];

    // ── Level 1: Category (always visible) ──
    const catItem = await waitFor(() => findItem(category, 1), CFG.waitTimeout).catch(() => null);
    if (!catItem) { log(`Category not found: "${category}"`, 'error'); return null; }
    if (!topic) return catItem;

    // ── Level 2: expand category, wait for topic ──
    await expandItem(catItem);
    const topicItem = await waitFor(
      () => findItem(topic, 2, catItem), CFG.waitTimeout
    ).catch(() => null);
    if (!topicItem) { log(`Topic not found: "${topic}" under "${category}"`, 'error'); return null; }
    if (!intent) return topicItem;

    // ── Level 3: expand topic, wait for intent ──
    await expandItem(topicItem);
    const intentItem = await waitFor(
      () => findItem(intent, 3, topicItem), CFG.waitTimeout
    ).catch(() => null);
    if (!intentItem) { log(`Intent not found: "${intent}" under "${topic}" > "${category}"`, 'error'); return null; }

    await scrollIntoView(intentItem);
    return intentItem;
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

  async function selectDestination(destination) {
    await sleep(CFG.dialogDelay);
    const searchInput = await waitFor(() => document.querySelector([
      'p-dialog input[type="search"]', '.p-dialog input[type="search"]',
      '[role="dialog"] input[type="search"]',
      'p-dialog input[placeholder*="Search" i]', '.p-dialog input[placeholder*="Search" i]',
      '[role="dialog"] input[placeholder*="Search" i]',
      '[role="dialog"] input[type="text"]',
    ].join(','))).catch(() => null);
    if (searchInput) { setInputValue(searchInput, destination); await sleep(CFG.dialogDelay); }

    const destEl = await waitFor(() => {
      for (const d of document.querySelectorAll('p-dialog,.p-dialog,[role="dialog"],mat-dialog-container')) {
        for (const el of d.querySelectorAll('.kanban-tree-node-name')) {
          if (normalize(el.textContent) === normalize(destination)) return el;
        }
        for (const el of d.querySelectorAll('li[role="treeitem"]')) {
          if (normalize(el.getAttribute('aria-label')) === normalize(destination)) return el;
        }
        for (const el of d.querySelectorAll('li, span, div')) {
          if (normalize(el.textContent).trim() === normalize(destination) &&
              el.offsetParent !== null && !el.querySelector('li, span, div')) return el;
        }
      }
      return null;
    });
    destEl.click();
    await sleep(CFG.shortDelay);
    await clickDialogButton('save', 'move', 'merge', 'confirm', 'ok');
  }

  async function performMove(li, destination) {
    await openMoreMenu(li);
    await clickMenuOption('Move');
    await selectDestination(destination);
    log(`  Moved → "${destination}"`, 'success');
  }

  async function performMerge(li, target) {
    await openMoreMenu(li);
    await clickMenuOption('Merge');
    await selectDestination(target);
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
