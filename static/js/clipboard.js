'use strict';

// Clipboard integration for ep_docx_html_customizer – simplified version.
// We now rely on Etherpad part ordering (our part loads before ep_hyperlinked_text)
// so we use the same pattern (jQuery .on('paste')) instead of capture-phase hacks.

const {customizeDocument} = require('../../transform_common');

// ADD: Constants matching ep_tables5 for table detection & delimiter cleanup
const ATTR_TABLE_JSON = 'tbljson';
const DELIMITER = '\u241F'; // same invisible delimiter used by ep_tables5

exports.postAceInit = (hook, context) => {
  console.log('[docx_customizer] postAceInit invoked – clipboard customization ready');
  const DEBUG = true;

  // Defer until ace_inner iframe is available, same pattern as hyperlinked_text
  context.ace.callWithAce(() => {
    const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
    if (!$innerIframe.length) {
      if (DEBUG) console.warn('[docx_customizer] inner iframe not found at postAceInit');
      return;
    }
    if (DEBUG) console.log('[docx_customizer] inner iframe located');

    const $innerBody = $innerIframe.contents().find('body');
    if (!$innerBody.length) {
      if (DEBUG) console.warn('[docx_customizer] inner body not found inside iframe');
      return;
    }

    if (DEBUG) console.log('[docx_customizer] attaching paste listener');

    $innerBody.on('paste', (evt) => {
      const clipboardData = evt.originalEvent.clipboardData;
      if (DEBUG) console.log('[docx_customizer] paste event captured');
      if (!clipboardData) {
        if (DEBUG) console.log('[docx_customizer] no clipboardData');
        return;
      }
      if (!clipboardData.types.includes('text/html')) return; // let core handle plain text

      const html = clipboardData.getData('text/html');
      if (!html) {
        if (DEBUG) console.log('[docx_customizer] clipboard has no HTML data');
        return;
      }

      if (DEBUG) console.log('[docx_customizer] transforming HTML');
      evt.preventDefault();
      evt.stopImmediatePropagation();

      // Transform clipboard HTML
      const doc = new DOMParser().parseFromString(html, 'text/html');
      customizeDocument(doc, {env: 'browser'});
      const cleanedHtml = doc.body.innerHTML;

      if (DEBUG) console.log('[docx_customizer] cleanedHtml length', cleanedHtml.length);

      // Inline remote images (e.g., Google Drive) to data URIs if CORS allows
      const inlineImages = async (html) => {
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        const spans = tmp.querySelectorAll('span[class*="image:"]');
        await Promise.all(Array.from(spans).map(async (sp) => {
          const m = sp.className.match(/image:([^ ]+)/);
          if (!m) return;
          let url = decodeURIComponent(m[1]);
          if (!/^https?:/.test(url) || url.startsWith('data:')) return;
          try {
            // First attempt direct fetch (may fail due to CORS)
            let resp;
            try {
              resp = await fetch(url, {mode: 'cors'});
              if (!resp.ok) throw new Error(`status ${resp.status}`);
            } catch (corsErr) {
              if (DEBUG) console.warn('[docx_customizer] direct fetch failed, retry via proxy', url, corsErr);
              resp = await fetch(`/ep_docx_image_proxy?url=${encodeURIComponent(url)}`);
              if (resp.status === 404) {
                // Might be served behind a reverse-proxy with a path prefix (e.g. /etherpad)
                const prefix = window.location.pathname.split('/p/')[0] || '';
                try {
                  resp = await fetch(`${prefix}/ep_docx_image_proxy?url=${encodeURIComponent(url)}`);
                } catch (_) { /* ignore */ }
              }
              if (!resp.ok) throw new Error(`proxy status ${resp.status}`);
            }

            const blob = await resp.blob();
            const dataUrl = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(blob); });
            // Determine intrinsic size to add width/height classes for correct aspect ratio
            try {
              const dim = await new Promise((res, rej) => {
                const imgObj = new Image();
                imgObj.onload = () => res({w: imgObj.naturalWidth, h: imgObj.naturalHeight});
                imgObj.onerror = rej;
                imgObj.src = URL.createObjectURL(blob);
              });
              if (dim && dim.w && dim.h) {
                const ratio = (dim.w / dim.h).toFixed(4);
                sp.classList.add(`image-width:${dim.w}px`);
                sp.classList.add(`image-height:${dim.h}px`);
                sp.classList.add(`imageCssAspectRatio:${ratio}`);
              }
            } catch (_) { /* ignore failures */ }
            const encoded = encodeURIComponent(dataUrl);
            sp.className = sp.className.replace(m[1], encoded);
            // `customizeDocument()` already wrapped the placeholder with
            // a single ZWSP on each side.  No extra normalisation needed.
            if (DEBUG) console.log('[docx_customizer] inlined remote image', url);
          } catch (e) {
            if (DEBUG) console.warn('[docx_customizer] failed to inline', url, e);
          }
        }));
        return tmp.innerHTML;
      };

      inlineImages(cleanedHtml).then((finalHtml) => {
        context.ace.callWithAce((ace) => {
          // Determine whether caret is inside an ep_tables5 table cell
          let insideTableCell = false;
          try {
            const rep = ace.ace_getRep && ace.ace_getRep();
            const docMgr = ace.documentAttributeManager || ace.editorInfo?.documentAttributeManager;
            if (rep && rep.selStart && docMgr && docMgr.getAttributeOnLine) {
              const lineNum = rep.selStart[0];
              const attr = docMgr.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
              if (attr) insideTableCell = true;
            }
          } catch (_) {/* ignore */}

          // Fallback DOM-based detection (for block-styled rows)
          if (!insideTableCell) {
            try {
              const innerWin = $innerIframe[0].contentWindow;
              const sel = innerWin.getSelection && innerWin.getSelection();
              if (sel && sel.rangeCount) {
                let n = sel.getRangeAt(0).startContainer;
                while (n) {
                  if (n.nodeType === 1 && n.matches && n.matches('table.dataTable')) { insideTableCell = true; break; }
                  n = n.parentNode;
                }
              }
            } catch (_) {/* ignore */}
          }

          // NEW: Ensure the paste selection is fully inside a single cell. Abort if it spans multiple cells/rows.
          if (insideTableCell) {
            if (DEBUG) console.log('[docx_customizer] insideTableCell confirmed – using safe PLAINTEXT insertion');

            try {
              const repNow = ace.ace_getRep();
              if (!repNow || !repNow.selStart) {
                console.warn('[docx_customizer] rep missing at paste-time, aborting plain-text insert');
                return;
              }

              // Validate that selection is within the current cell (same logic as before)
              const selStart = repNow.selStart.slice();
              const selEnd   = repNow.selEnd.slice();
              const lineNum  = selStart[0];

              const lineText = repNow.lines.atIndex(lineNum)?.text || '';
              const cells    = lineText.split(DELIMITER);
              const docMgrPre = ace.documentAttributeManager || ace.editorInfo?.documentAttributeManager;
              const attrStrBefore = docMgrPre?.getAttributeOnLine ? docMgrPre.getAttributeOnLine(lineNum, ATTR_TABLE_JSON) : null;
              let tableMetaBefore = null;
              try { if (attrStrBefore) tableMetaBefore = JSON.parse(attrStrBefore); } catch {}
              let currentOffset = 0;
              let targetCellIndex = -1;
              let cellEndCol = 0;
              for (let i = 0; i < cells.length; i++) {
                const cellLength = cells[i]?.length ?? 0;
                const cellEndThis = currentOffset + cellLength;
                if (selStart[1] >= currentOffset && selStart[1] <= cellEndThis) {
                  targetCellIndex = i;
                  cellEndCol   = cellEndThis;
                  break;
                }
                currentOffset += cellLength + DELIMITER.length;
              }

              if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
                console.warn('[docx_customizer] selection extends outside target cell – abort');
                return;
              }

              // --- Convert HTML to plain text (strip tags) ---
              const tmp = document.createElement('div');
              tmp.innerHTML = finalHtml;
              let plainText = tmp.textContent || tmp.innerText || '';

              // Sanitize similar to ep_tables5
              plainText = plainText
                .replace(/(\r\n|\n|\r)/gm, ' ')
                .replace(new RegExp(DELIMITER, 'g'), ' ') // Strip delimiter char
                .replace(/\t/g, ' ') // Tabs to space
                .replace(/\s+/g, ' ') // Collapse whitespace
                .trim();

              if (!plainText) {
                if (DEBUG) console.log('[docx_customizer] Plaintext empty after sanitization – abort paste');
                return;
              }

              // Length cap same as tables plugin
              const currentCellText = cells[targetCellIndex] || '';
              const selectionLength = selEnd[1] - selStart[1];
              const MAX_CELL_LENGTH = 8000;
              const newCellLength = currentCellText.length - selectionLength + plainText.length;
              if (newCellLength > MAX_CELL_LENGTH) {
                plainText = plainText.substring(0, MAX_CELL_LENGTH - (currentCellText.length - selectionLength));
              }

              // Perform replacement
              ace.ace_performDocumentReplaceRange(selStart, selEnd, plainText);

              // Reapply tbljson metadata
              if (ace.ep_tables5_applyMeta && tableMetaBefore && typeof tableMetaBefore.cols === 'number') {
                const repAfter = ace.ace_getRep();
                const docMgrSafe = docMgrPre;
                ace.ep_tables5_applyMeta(
                  lineNum,
                  tableMetaBefore.tblId,
                  tableMetaBefore.row,
                  tableMetaBefore.cols,
                  repAfter,
                  ace,
                  null,
                  docMgrSafe,
                );
              }

              // Move caret to end of inserted text
              const newCaretCol = selStart[1] + plainText.length;
              ace.ace_performSelectionChange([lineNum, newCaretCol], [lineNum, newCaretCol], false);

              ace.ace_fastIncorp && ace.ace_fastIncorp(10);

              if (DEBUG) console.log('[docx_customizer] Plain-text table-cell paste completed');

            } catch (errPaste) {
              console.error('[docx_customizer] error during plain-text table paste', errPaste);
            }
          } else {
            // Normal (non-table) insertion – keep HTML & formatting
            ace.ace_inCallStackIfNecessary('docxPaste', () => {
              const innerWin = $innerIframe[0].contentWindow;
              innerWin.document.execCommand('insertHTML', false, finalHtml);
            });
          }
        }, 'docxPaste', true);
      });

      if (DEBUG) console.log('[docx_customizer] insertion done, event propagation stopped');
    });

    if (DEBUG) console.log('[docx_customizer] paste listener attached');
  }, 'setupDocxCustomizerPaste', true);
}; 