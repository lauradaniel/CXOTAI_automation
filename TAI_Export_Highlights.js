/**
 * ============================================================================
 *  CXOne Intent Scraper — Browser Console Version
 * ============================================================================
 *
 *  HOW TO USE:
 *    1. Open CXOne in Chrome and navigate to the Intent Builder page
 *       (the page with the Category > Topic > Intent kanban tree).
 *    2. Open Chrome DevTools (F12 or Ctrl+Shift+I).
 *    3. Go to the "Console" tab.
 *    4. Copy-paste this ENTIRE script into the console and press Enter.
 *    5. The script will:
 *       - Expand every collapsed tree node
 *       - Click each intent to open its detail panel
 *       - Read phrases from .phrases-snippets-container
 *       - Download the result as .xlsx and .csv files
 *
 *  NOTE: Zero external dependencies. Just paste and run.
 * ============================================================================
 */

(async function CXOneIntentScraper() {
  'use strict';

  // ── Configuration ──────────────────────────────────────────────────────
  const EXPAND_DELAY   = 1000;  // ms to wait after expanding a tree node
  const CLICK_DELAY    = 2500;  // ms to wait after clicking an intent
  const BETWEEN_CLICKS = 300;   // ms between sequential intent clicks

  // ── Helpers ────────────────────────────────────────────────────────────
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function log(msg) {
    console.log('%c[CXOne Scraper] ' + msg, 'color: #2196F3; font-weight: bold;');
  }
  function logProgress(cur, tot, text) {
    console.log('%c[CXOne Scraper] [' + cur + '/' + tot + ' - ' + ((cur/tot)*100).toFixed(1) + '%] ' + text, 'color: #4CAF50;');
  }
  function logWarn(msg) {
    console.log('%c[CXOne Scraper] WARNING: ' + msg, 'color: #FF9800; font-weight: bold;');
  }

  // ── Dynamic Filename Builder ──────────────────────────────────────────
  function buildFileName() {
    var parts = [];

    // 1. Company name from aptrinsic('identify', ...) script
    var companyName = '';
    var scripts = document.querySelectorAll('script');
    for (var s = 0; s < scripts.length; s++) {
      var text = scripts[s].textContent || '';
      var match = text.match(/aptrinsic\s*\(\s*['"]identify['"]\s*,[\s\S]*?,\s*\{[^}]*"name"\s*:\s*"([^"]+)"/);
      if (match) { companyName = match[1].trim(); break; }
    }
    if (companyName) parts.push(companyName.replace(/\s+/g, '_'));

    // 2. Model name from .model-selection-dropdown-wrapper span[data-aid="ellipsis-sliced-text"]
    var modelEl = document.querySelector('.model-selection-dropdown-wrapper span[data-aid="ellipsis-sliced-text"]');
    var modelName = modelEl ? modelEl.textContent.trim() : '';
    if (modelName) parts.push(modelName.replace(/\s+/g, '_'));

    // 3. Date from span.version-creation-date
    var dateEl = document.querySelector('span.version-creation-date');
    var dateStr = dateEl ? dateEl.textContent.trim() : '';
    if (dateStr) parts.push(dateStr.replace(/\s+/g, '_'));

    // Combine with underscores; fallback if nothing was found
    var name = parts.length > 0 ? parts.join('_') : 'CXOne_Intents_Output';
    // Remove any characters that are unsafe for filenames
    name = name.replace(/[\/\\:*?"<>|]/g, '_');
    return name;
  }

  // ── Step 1: Verify page ────────────────────────────────────────────────
  log('Starting CXOne Intent Scraper...');

  const kanbanPanel = document.querySelector('.kanban-view-panel');
  if (!kanbanPanel) {
    logWarn('Could not find .kanban-view-panel. Are you on the Intent Builder page?');
    return;
  }

  // ── Step 2: Expand all collapsed tree nodes ────────────────────────────
  log('Step 1/4: Expanding all collapsed tree nodes...');

  let expandedTotal = 0;
  let passNum = 0;

  while (true) {
    passNum++;
    const collapsedTogglers = kanbanPanel.querySelectorAll(
      'p-treenode .p-tree-toggler:has(chevronrighticon)'
    );
    const visible = Array.from(collapsedTogglers).filter(el => el.offsetParent !== null);
    if (visible.length === 0) break;

    log('  Pass ' + passNum + ': expanding ' + visible.length + ' collapsed node(s)...');
    for (const toggler of visible) {
      toggler.scrollIntoView({ block: 'center', behavior: 'instant' });
      toggler.click();
      await sleep(EXPAND_DELAY);
      expandedTotal++;
    }
  }
  log('  Done - expanded ' + expandedTotal + ' node(s).');

  // ── Step 3: Collect tree hierarchy ─────────────────────────────────────
  log('Step 2/4: Collecting Category > Topic > Intent hierarchy...');

  const allNodes = kanbanPanel.querySelectorAll('.kanban-tree-node');
  const intentList = [];
  let currentCategory = '';
  let currentTopic = '';

  for (const node of allNodes) {
    const cls = node.className;
    const nameEl = node.querySelector('.kanban-tree-node-name');
    const pctEl = node.querySelector('.kanban-tree-node-statistics .percentage');
    const name = nameEl ? nameEl.textContent.trim() : '';
    const pct = pctEl ? pctEl.textContent.trim() : '0%';

    if (cls.includes('node-level-1')) {
      currentCategory = name;
      currentTopic = '';
    } else if (cls.includes('node-level-2')) {
      currentTopic = name;
    } else if (cls.includes('node-level-3')) {
      // Read the Active label from .new-item-label inside the node
      const activeEl = node.querySelector('.new-item-label');
      const activeText = activeEl ? activeEl.textContent.trim() : '';
      intentList.push({
        category: currentCategory,
        topic: currentTopic,
        intent: name,
        intentPercentage: pct,
        volume: '',
        examples: '',
        tag: activeText,
        _nodeEl: node,
      });
    }
  }

  log('  Found ' + intentList.length + ' Level-3 intents.');
  if (intentList.length === 0) {
    logWarn('No Level-3 intents found. Is the kanban tree fully expanded?');
    return;
  }

  // ── Step 4: Click each intent and read phrases ─────────────────────────
  log('Step 3/4: Clicking each intent to extract phrases...');
  var totalTime = Math.ceil((intentList.length * (CLICK_DELAY + BETWEEN_CLICKS)) / 1000 / 60);
  log('  Estimated time: ~' + totalTime + ' minutes for ' + intentList.length + ' intents.');

  // The first intent may already be selected/highlighted on page load,
  // so clicking it won't trigger the detail panel to load. To fix this,
  // click the second intent first (to deselect the first), then proceed
  // normally starting from the first intent.
  if (intentList.length > 1) {
    const secondContent = intentList[1]._nodeEl.closest('.p-treenode-content') || intentList[1]._nodeEl;
    secondContent.scrollIntoView({ block: 'center', behavior: 'instant' });
    secondContent.click();
    await sleep(CLICK_DELAY);
    log('  Deselected first intent to ensure click registers.');
  }

  for (let i = 0; i < intentList.length; i++) {
    const item = intentList[i];

    // Click the intent node to open its detail panel on the left
    const treeContent = item._nodeEl.closest('.p-treenode-content') || item._nodeEl;
    treeContent.scrollIntoView({ block: 'center', behavior: 'instant' });
    treeContent.click();

    await sleep(CLICK_DELAY);

    // Read phrases from the detail panel
	const phrasesContainer = document.querySelector('.phrases-snippets-container');
	if (phrasesContainer) {
	  // Target specific phrase containers or rows if possible. 
	  // If the structure is generic, we use a more inclusive text-grabbing method:
	  const phraseElements = phrasesContainer.querySelectorAll('.phrase-text, .snippet-item, li, p');
	  
	  let phrases = [];
	  if (phraseElements.length > 0) {
		phrases = Array.from(phraseElements)
		  .map(el => el.textContent.trim())
		  .filter(txt => txt.length > 0);
	  } else {
		// Fallback: Get all text but split by lines to preserve separation
		phrases = phrasesContainer.innerText.split('\n')
		  .map(t => t.trim())
		  .filter(t => t.length > 0);
	  }
	  
	  // Remove duplicates and join
	  item.examples = [...new Set(phrases)].join('\n'); // [cite: 40]
	}

    // If Active wasn't found during tree collection, try again now
    if (!item.tag) {
      const activeEl = item._nodeEl.querySelector('.new-item-label');
      if (activeEl) {
        item.tag = activeEl.textContent.trim();
      }
    }

    const exCount = item.examples ? item.examples.split('\n').length : 0;
    logProgress(i + 1, intentList.length, item.category + ' > ' + item.topic + ' > ' + item.intent + ' (' + exCount + ' phrases)');
    await sleep(BETWEEN_CLICKS);
  }

  // Clean up DOM references
  const rows = intentList.map(function(item) {
    return {
      category: item.category,
      topic: item.topic,
      intent: item.intent,
      intentPercentage: item.intentPercentage,
      volume: item.volume,
      examples: item.examples,
      tag: item.tag
    };
  });

  // ── Step 5: Download ───────────────────────────────────────────────────
  log('Step 4/4: Generating Excel file...');
  var baseFileName = buildFileName();
  log('  Filename: ' + baseFileName);
  downloadExcel(rows, baseFileName);

  const categories = [];
  const seen = {};
  for (const r of rows) { if (!seen[r.category]) { seen[r.category] = 1; categories.push(r.category); } }
  const withExamples = rows.filter(function(r) { return r.examples; }).length;

  log('');
  log('=== COMPLETE ===');
  log('  Categories:      ' + categories.length);
  log('  Total intents:   ' + rows.length);
  log('  With examples:   ' + withExamples);
  log('  File downloaded: ' + baseFileName + '.xlsx');

  console.table(rows.slice(0, 10));
  if (rows.length > 10) {
    log('  ... and ' + (rows.length - 10) + ' more rows (see the downloaded file).');
  }

  // ────────────────────────────────────────────────────────────────────────
  //  XLSX Generator + CSV + ZIP (zero dependencies)
  // ────────────────────────────────────────────────────────────────────────

  function downloadExcel(data, baseFileName) {
    var headers = ['Category','Topic','Intent','Intent Percentage','Volume','Examples','Tag'];
    var keys = ['category','topic','intent','intentPercentage','volume','examples','tag'];

    function esc(s) {
      if (s == null) return '';
      return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/\n/g,'&#10;');
    }
    function col(i) {
      var s = ''; i++;
      while (i > 0) { i--; s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26); }
      return s;
    }

    var sr = '<row r="1">';
    for (var c = 0; c < headers.length; c++) {
      sr += '<c r="' + col(c) + '1" t="inlineStr" s="1"><is><t>' + esc(headers[c]) + '</t></is></c>';
    }
    sr += '</row>';
    for (var r = 0; r < data.length; r++) {
      var rn = r + 2;
      sr += '<row r="' + rn + '">';
      for (var c2 = 0; c2 < keys.length; c2++) {
        // Apply wrap-text style (s="2") to Examples column (index 5)
        var style = c2 === 5 ? ' s="2"' : '';
        var rawVal = data[r][keys[c2]] || '';

        // Intent Percentage column (index 3) — store as number
        if (c2 === 3) {
          var num = parseFloat(String(rawVal).replace('%', ''));
          if (!isNaN(num)) {
            sr += '<c r="' + col(c2) + rn + '" s="3"><v>' + num + '</v></c>';
          } else {
            sr += '<c r="' + col(c2) + rn + '" t="inlineStr"><is><t>' + esc(rawVal) + '</t></is></c>';
          }
        } else {
          var val = esc(rawVal);
          var preserveSpace = val.indexOf('&#10;') !== -1 ? ' xml:space="preserve"' : '';
          sr += '<c r="' + col(c2) + rn + '" t="inlineStr"' + style + '><is><t' + preserveSpace + '>' + val + '</t></is></c>';
        }
      }
      sr += '</row>';
    }

    var lc = col(headers.length - 1);
    var lr = data.length + 1;

    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="45" customWidth="1"/><col min="4" max="4" width="18" customWidth="1"/><col min="5" max="5" width="12" customWidth="1"/><col min="6" max="6" width="60" customWidth="1"/><col min="7" max="7" width="10" customWidth="1"/></cols>' +
      '<sheetData>' + sr + '</sheetData>' +
      '<autoFilter ref="A1:' + lc + lr + '"/></worksheet>';

    // 
    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>' + // Moved to top
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>' +
      '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellXfs count="4"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf><xf applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>';

    var wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Intents" sheetId="1" r:id="rId1"/></sheets></workbook>';

    var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';

    var blob = buildZip([
      {name:'[Content_Types].xml', content:ct},
      {name:'_rels/.rels', content:rels},
      {name:'xl/workbook.xml', content:wbXml},
      {name:'xl/_rels/workbook.xml.rels', content:wbRels},
      {name:'xl/styles.xml', content:stylesXml},
      {name:'xl/worksheets/sheet1.xml', content:sheetXml}
    ]);

    dl(blob, baseFileName + '.xlsx');
    log('  Excel file download triggered.');

    // CSV backup
    var csv = [headers.join(',')];
    for (var ri = 0; ri < data.length; ri++) {
      csv.push(keys.map(function(k) { return '"' + (data[ri][k]||'').replace(/"/g,'""') + '"'; }).join(','));
    }
    dl(new Blob([csv.join('\n')], {type:'text/csv'}), baseFileName + '.csv');
    log('  CSV backup also downloaded.');
  }

  function dl(blob, name) {
    var u = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = u; a.download = name;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(u);
  }

  function buildZip(files) {
    var enc = new TextEncoder();
    var parts = [], cd = [], off = 0;
    for (var i = 0; i < files.length; i++) {
      var nb = enc.encode(files[i].name);
      var cb = enc.encode(files[i].content);
      var cr = crc32(cb), sz = cb.length;

      var lh = new Uint8Array(30 + nb.length);
      var lv = new DataView(lh.buffer);
      lv.setUint32(0,0x04034b50,true); lv.setUint16(4,20,true);
      lv.setUint16(8,0,true); lv.setUint32(14,cr,true);
      lv.setUint32(18,sz,true); lv.setUint32(22,sz,true);
      lv.setUint16(26,nb.length,true); lh.set(nb,30);
      parts.push(lh, cb);

      var ce = new Uint8Array(46 + nb.length);
      var cv = new DataView(ce.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint32(16,cr,true); cv.setUint32(20,sz,true); cv.setUint32(24,sz,true);
      cv.setUint16(28,nb.length,true); cv.setUint32(42,off,true);
      ce.set(nb,46); cd.push(ce);
      off += lh.length + cb.length;
    }
    var cdOff = off, cdSz = 0;
    for (var j = 0; j < cd.length; j++) { parts.push(cd[j]); cdSz += cd[j].length; }
    var eo = new Uint8Array(22);
    var ev = new DataView(eo.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true);
    ev.setUint32(12,cdSz,true); ev.setUint32(16,cdOff,true);
    parts.push(eo);
    return new Blob(parts,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }

  function crc32(b) {
    var c = 0xffffffff;
    for (var i = 0; i < b.length; i++) {
      c ^= b[i];
      for (var j = 0; j < 8; j++) c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0);
    }
    return (c ^ 0xffffffff) >>> 0;
  }
})();