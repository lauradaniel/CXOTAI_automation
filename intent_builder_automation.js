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
    shortDelay:  400,   // ms – between micro-steps
    actionDelay: 700,   // ms – after opening a menu / typing
    dialogDelay: 1200,  // ms – after a dialog opens or closes
    waitTimeout: 8000,  // ms – max time to wait for a DOM element
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

  /**
   * Polls until selectorFn() returns a truthy value or timeout is reached.
   */
  function waitFor(selectorFn, timeout = CFG.waitTimeout) {
    return new Promise((resolve, reject) => {
      const deadline = Date.now() + timeout;
      const iv = setInterval(() => {
        const el = selectorFn();
        if (el) { clearInterval(iv); resolve(el); }
        else if (Date.now() > deadline) {
          clearInterval(iv);
          reject(new Error('waitFor timeout: element not found'));
        }
      }, 100);
    });
  }

  /**
   * Sets an input's value in a way Angular's change detection picks up.
   */
  function setInputValue(input, value) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(input, value);
    input.dispatchEvent(new Event('input',  { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  /**
   * Fires mouseenter + mouseover on an element (reveals hover-only buttons).
   */
  async function hoverOver(el) {
    el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mouseover',  { bubbles: true }));
    await sleep(CFG.shortDelay);
  }

  /**
   * Closes any open dialog by clicking its close / cancel button.
   * Used for error recovery.
   */
  async function tryCloseDialog() {
    const sel = [
      'p-dialog .p-dialog-header-close',
      '.p-dialog .p-dialog-header-close',
      '[role="dialog"] button[aria-label="Close"]',
      '[role="dialog"] button[aria-label="close"]',
      '[role="dialog"] .p-dialog-header-close',
    ].join(', ');
    const btn = document.querySelector(sel);
    if (btn) { btn.click(); await sleep(CFG.dialogDelay); }
  }


  // ─────────────────────────────────────────────
  // CSV PARSING
  // ─────────────────────────────────────────────
  function parseCSV(text) {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) throw new Error('CSV has no data rows');

    // Parse a single line respecting quoted fields
    function parseLine(line) {
      const fields = [];
      let cur = '', inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          // Handle escaped quotes ("")
          if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
          else inQ = !inQ;
        } else if (ch === ',' && !inQ) {
          fields.push(cur.trim()); cur = '';
        } else {
          cur += ch;
        }
      }
      fields.push(cur.trim());
      return fields;
    }

    const headers = parseLine(lines[0]);
    return lines.slice(1).map(line => {
      const vals = parseLine(line);
      const row = {};
      headers.forEach((h, i) => row[h] = (vals[i] || '').replace(/^"|"$/g, '').trim());
      return row;
    });
  }


  // ─────────────────────────────────────────────
  // FILE PICKER
  // ─────────────────────────────────────────────
  function pickCSVFile() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.csv,text/csv';
      input.style.display = 'none';
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

  /**
   * Returns the display name of a treeitem element.
   * Prefers .kanban-tree-node-name text; falls back to aria-label.
   */
  function getItemName(li) {
    const nameEl = li.querySelector('.kanban-tree-node-name');
    if (nameEl) return normalize(nameEl.textContent);
    return normalize(li.getAttribute('aria-label'));
  }

  /**
   * Finds a treeitem at the given aria-level whose display name matches.
   * Optionally restricts search to within a parent element.
   */
  function findItem(name, level, parent = document) {
    const n = normalize(name);
    const items = parent.querySelectorAll(`li[role="treeitem"][aria-level="${level}"]`);
    for (const li of items) {
      if (getItemName(li) === n) return li;
    }
    return null;
  }

  /**
   * Expands a treeitem if it is currently collapsed.
   */
  async function ensureExpanded(li) {
    if (li.getAttribute('aria-expanded') === 'false') {
      const toggler = li.querySelector('button.p-tree-toggler');
      if (toggler) { toggler.click(); await sleep(CFG.actionDelay); }
    }
  }

  /**
   * Locates the target treeitem described by a CSV row.
   * Handles all three levels (Category → Topic → Intent).
   */
  async function locateTarget(row) {
    const category = row['Category'];
    const topic    = row['Topic'];
    const intent   = row['Intent'];

    // ── Level 1: Category ──
    const catItem = findItem(category, 1);
    if (!catItem) {
      log(`Category not found: "${category}"`, 'error');
      return null;
    }

    if (!topic) return catItem; // action targets the category itself

    // ── Level 2: Topic ──
    await ensureExpanded(catItem);
    const topicItem = findItem(topic, 2, catItem);
    if (!topicItem) {
      log(`Topic not found: "${topic}" under "${category}"`, 'error');
      return null;
    }

    if (!intent) return topicItem; // action targets the topic itself

    // ── Level 3: Intent ──
    await ensureExpanded(topicItem);
    const intentItem = findItem(intent, 3, topicItem);
    if (!intentItem) {
      log(`Intent not found: "${intent}" under "${topic}" > "${category}"`, 'error');
      return null;
    }

    return intentItem;
  }


  // ─────────────────────────────────────────────
  // MENU HELPERS
  // ─────────────────────────────────────────────

  /**
   * Hovers over a treeitem and clicks its three-dot (kanban-more-button).
   */
  async function openMoreMenu(li) {
    await hoverOver(li);
    const btn = await waitFor(() => li.querySelector('button.kanban-more-button'));
    btn.click();
    await sleep(CFG.actionDelay);
  }

  /**
   * Finds and clicks a menu option whose visible text contains `label`.
   * Searches inside any currently-visible overlay / context-menu.
   */
  async function clickMenuOption(label) {
    const target = normalize(label);
    const el = await waitFor(() => {
      // Candidate containers for the floating menu
      const containers = document.querySelectorAll([
        'p-overlaypanel',
        '.p-overlaypanel',
        '.p-menu',
        '.p-contextmenu',
        '[role="menu"]',
        '.cxone-menu',
        '.dropdown-menu',
        '.context-menu',
      ].join(','));

      for (const c of containers) {
        if (!c.offsetParent && c.style.display === 'none') continue; // hidden
        const items = c.querySelectorAll('li, [role="menuitem"], button, a');
        for (const item of items) {
          if (normalize(item.textContent).includes(target)) return item;
        }
      }

      // Fallback: any visible element exactly matching the label
      for (const el of document.querySelectorAll('li, [role="menuitem"], button.p-menuitem-link')) {
        if (normalize(el.textContent).trim() === target && el.offsetParent !== null) return el;
      }
      return null;
    });

    el.click();
    await sleep(CFG.actionDelay);
  }

  /**
   * Finds a button inside an open dialog whose text matches one of the provided labels.
   */
  async function clickDialogButton(...labels) {
    const targets = labels.map(normalize);
    const btn = await waitFor(() => {
      const dialogs = document.querySelectorAll('p-dialog,.p-dialog,[role="dialog"],mat-dialog-container');
      for (const d of dialogs) {
        for (const b of d.querySelectorAll('button')) {
          const t = normalize(b.textContent).trim();
          if (targets.includes(t) && !b.disabled) return b;
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

  // ── RENAME ──────────────────────────────────
  async function performRename(li, newName) {
    await openMoreMenu(li);
    await clickMenuOption('Rename');

    // Locate text input in the rename dialog
    const input = await waitFor(() =>
      document.querySelector([
        'p-dialog input[type="text"]',
        '.p-dialog input[type="text"]',
        '[role="dialog"] input[type="text"]',
        'mat-dialog-container input[type="text"]',
      ].join(','))
    );

    input.focus();
    input.select();
    setInputValue(input, newName);
    await sleep(CFG.shortDelay);

    await clickDialogButton('rename', 'save', 'confirm', 'ok');
    log(`  Renamed → "${newName}"`, 'success');
  }


  // ── REMOVE ──────────────────────────────────
  async function performRemove(li) {
    await openMoreMenu(li);
    await clickMenuOption('Remove');

    await clickDialogButton('remove', 'delete', 'confirm', 'yes');
    log('  Removed', 'success');
  }


  // ── SHARED: select destination in Move/Merge dialog ──
  async function selectDestination(destination) {
    await sleep(CFG.dialogDelay); // wait for dialog to fully render

    // Try to use the search box inside the dialog
    const searchInput = await waitFor(() =>
      document.querySelector([
        'p-dialog input[type="search"]',
        '.p-dialog input[type="search"]',
        '[role="dialog"] input[type="search"]',
        'p-dialog input[placeholder*="Search" i]',
        '.p-dialog input[placeholder*="Search" i]',
        '[role="dialog"] input[placeholder*="Search" i]',
        'p-dialog input[placeholder*="search" i]',
        '[role="dialog"] input[type="text"]',
      ].join(','))
    ).catch(() => null);

    if (searchInput) {
      setInputValue(searchInput, destination);
      await sleep(CFG.dialogDelay);
    }

    // Find and click the destination item inside the dialog tree
    const destEl = await waitFor(() => {
      const dialogs = document.querySelectorAll('p-dialog,.p-dialog,[role="dialog"],mat-dialog-container');
      for (const d of dialogs) {
        // Prefer .kanban-tree-node-name elements
        for (const el of d.querySelectorAll('.kanban-tree-node-name')) {
          if (normalize(el.textContent) === normalize(destination)) return el;
        }
        // Fall back to treeitem aria-label
        for (const el of d.querySelectorAll('li[role="treeitem"]')) {
          if (normalize(el.getAttribute('aria-label')) === normalize(destination)) return el;
        }
        // Generic text match on any visible element
        for (const el of d.querySelectorAll('li, span, div')) {
          if (
            normalize(el.textContent).trim() === normalize(destination) &&
            el.offsetParent !== null &&
            !el.querySelector('li, span, div') // leaf node only
          ) return el;
        }
      }
      return null;
    });

    destEl.click();
    await sleep(CFG.shortDelay);

    await clickDialogButton('save', 'move', 'merge', 'confirm', 'ok');
  }


  // ── MOVE ────────────────────────────────────
  async function performMove(li, destination) {
    await openMoreMenu(li);
    // The menu item may say "Move", "Move/Merge", or similar
    await clickMenuOption('Move');
    await selectDestination(destination);
    log(`  Moved → "${destination}"`, 'success');
  }


  // ── MERGE ───────────────────────────────────
  async function performMerge(li, target) {
    await openMoreMenu(li);
    // If the menu has a combined "Move/Merge" option, "Merge" text will match it
    await clickMenuOption('Merge');
    await selectDestination(target);
    log(`  Merged → "${target}"`, 'success');
  }


  // ─────────────────────────────────────────────
  // MAIN RUNNER
  // ─────────────────────────────────────────────
  async function run() {
    log('=== Intent Builder Automation Started ===');

    // 1. Pick the CSV file
    let csvText;
    try {
      csvText = await pickCSVFile();
    } catch (e) {
      log('File selection cancelled or failed: ' + e.message, 'error');
      return;
    }

    // 2. Parse CSV
    let rows;
    try {
      rows = parseCSV(csvText);
    } catch (e) {
      log('CSV parse error: ' + e.message, 'error');
      return;
    }
    log(`CSV loaded — ${rows.length} data rows`);

    // 3. Process rows
    let success = 0, skipped = 0, errors = 0;

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const rowNum = i + 2; // 1-based + header
      const action = (row['Action'] || '').trim();
      const change = (row['Required Change'] || '').trim();
      const label  = row['Intent'] || row['Topic'] || row['Category'] || '(unknown)';

      if (!action) {
        log(`Row ${rowNum}: No action — skipping "${label}"`, 'warn');
        skipped++;
        continue;
      }

      log(`Row ${rowNum}: [${action}] "${label}"`);

      try {
        const target = await locateTarget(row);
        if (!target) { errors++; continue; }

        switch (action.toLowerCase()) {

          case 'rename':
            if (!change) throw new Error('"Required Change" is empty for Rename');
            await performRename(target, change);
            break;

          case 'remove':
            await performRemove(target);
            break;

          case 'move':
            if (!change) throw new Error('"Required Change" is empty for Move');
            await performMove(target, change);
            break;

          case 'merge':
            if (!change) throw new Error('"Required Change" is empty for Merge');
            await performMerge(target, change);
            break;

          default:
            log(`Row ${rowNum}: Unknown action "${action}" — skipping`, 'warn');
            skipped++;
            continue;
        }

        success++;

      } catch (e) {
        log(`Row ${rowNum}: ERROR — ${e.message}`, 'error');
        errors++;
        await tryCloseDialog();
      }

      await sleep(CFG.shortDelay); // brief pause between rows
    }

    log(
      `=== Done — Success: ${success} | Skipped: ${skipped} | Errors: ${errors} ===`,
      errors === 0 ? 'success' : 'warn'
    );
  }

  await run();

})();
