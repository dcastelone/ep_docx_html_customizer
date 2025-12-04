'use strict';

// Clipboard integration for ep_docx_html_customizer – simplified version.
// We now rely on Etherpad part ordering (our part loads before ep_hyperlinked_text)
// so we use the same pattern (jQuery .on('paste')) instead of capture-phase hacks.

const {customizeDocument, uploadImageToS3Browser} = require('../../transform_common');

// ADD: Constants matching ep_tables5 for table detection & delimiter cleanup
const ATTR_TABLE_JSON = 'tbljson';
const DELIMITER = '\u241F'; // same invisible delimiter used by ep_tables5
const ATTR_CELL = 'td';
const DEBUG = true;

// Base64 decode helper (URL-safe) reused by table logic
const dec = (s) => {
  try {
    return atob(s.replace(/-/g, '+').replace(/_/g, '/'));
  } catch (e) {
    if (DEBUG) console.error('[docx_customizer] Base64 decode failed', s, e);
    return null;
  }
};

// Base64 encode helper (URL-safe) reused by table logic
const enc = (s) => {
  if (typeof btoa === 'function') {
    return btoa(s).replace(/\+/g, '-').replace(/_/g, '_');
  } else if (typeof Buffer === 'function') {
    return Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/_/g, '_');
  }
  return s;
};

exports.postAceInit = (hook, context) => {
  if (DEBUG) console.log('[docx_customizer] postAceInit invoked');
  context.ace.callWithAce(() => {
    const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
    if (!$innerIframe.length) return;
    const $innerBody = $innerIframe.contents().find('body');
    if (!$innerBody.length) return;

    $innerBody.on('paste', (evt) => {
      const clipboardData = evt.originalEvent.clipboardData;
      if (!clipboardData || !clipboardData.types.includes('text/html')) return;
      const html = clipboardData.getData('text/html');
      if (!html) return;

      if (DEBUG) console.log('[docx_customizer] paste event captured with HTML');
      evt.preventDefault();
      evt.stopImmediatePropagation();

      const doc = new DOMParser().parseFromString(html, 'text/html');
      customizeDocument(doc, {env: 'browser'});
      
      // CRITICAL: Regenerate tblId values to prevent conflicts with existing tables
      // When copying a table from Etherpad and pasting back, the original tblId is preserved
      // This causes orphan detection to incorrectly merge pasted rows into the original table
      const regenerateTableIds = (docBody) => {
        const rand = () => Math.random().toString(36).slice(2, 8);
        const tblIdMap = new Map(); // oldId -> newId mapping
        
        // Find all elements with tbljson-* classes
        const tbljsonElements = docBody.querySelectorAll('[class*="tbljson-"]');
        if (tbljsonElements.length > 0 && DEBUG) {
          console.log('[docx_customizer] regenerating tblIds for', tbljsonElements.length, 'elements');
        }
        
        tbljsonElements.forEach((el) => {
          const newClasses = [];
          const oldClasses = el.className.split(/\s+/);
          
          oldClasses.forEach((cls) => {
            if (cls.startsWith('tbljson-')) {
              try {
                // Decode the existing metadata
                const encoded = cls.substring(8);
                const decoded = dec(encoded);
                if (decoded) {
                  const meta = JSON.parse(decoded);
                  if (meta && meta.tblId) {
                    // Get or create a new tblId for this old tblId
                    if (!tblIdMap.has(meta.tblId)) {
                      tblIdMap.set(meta.tblId, rand() + rand());
                    }
                    // Update the tblId
                    meta.tblId = tblIdMap.get(meta.tblId);
                    // Re-encode with new tblId
                    const newEncoded = enc(JSON.stringify(meta));
                    newClasses.push('tbljson-' + newEncoded);
                    if (DEBUG) console.debug('[docx_customizer] regenerated tblId', { old: cls.substring(8, 20) + '...', newTblId: meta.tblId });
                    return;
                  }
                }
              } catch (e) {
                if (DEBUG) console.warn('[docx_customizer] failed to decode/re-encode tbljson', e);
              }
            }
            // Keep non-tbljson classes as-is
            newClasses.push(cls);
          });
          
          el.className = newClasses.join(' ');
        });
        
        if (tblIdMap.size > 0 && DEBUG) {
          console.log('[docx_customizer] regenerated', tblIdMap.size, 'unique tblIds');
        }
      };
      
      regenerateTableIds(doc.body);
      
      const cleanedHtml = doc.body.innerHTML;
      if (DEBUG) console.log('[docx_customizer] cleanedHtml length', cleanedHtml.length);

      const inlineImages = async (html) => {
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        const spans = tmp.querySelectorAll('span[class*="image:"]');
        const imageSpans = Array.from(spans);
        const totalImages = imageSpans.length;
        if (totalImages > 0) {
          if (DEBUG) console.log(`[docx_customizer] Starting to process ${totalImages} images`);
        }
        await Promise.all(imageSpans.map(async (sp) => {
          const m = sp.className.match(/image:([^ ]+)/);
          if (!m) return;
          let url = decodeURIComponent(m[1]);
          try {
            let blob;
            const padId = (typeof clientVars !== 'undefined') ? clientVars.padId : 'clipboard';
            let filename = `image-${Date.now()}`;

            if (url.startsWith('data:')) {
              // Convert data URL to Blob so we can upload via S3 presign
              const commaIdx = url.indexOf(',');
              if (commaIdx === -1) return;
              const header = url.substring(0, commaIdx);
              const b64 = url.substring(commaIdx + 1);
              const mimeMatch = /data:([^;]+);base64/i.exec(header);
              const mimeType = (mimeMatch && mimeMatch[1]) || 'application/octet-stream';
              const binary = atob(b64);
              const len = binary.length;
              const bytes = new Uint8Array(len);
              for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
              blob = new Blob([bytes], {type: mimeType});
              const ext = (mimeType.split('/')[1] || 'png');
              filename += `.${ext}`;
            } else if (/^https?:/.test(url)) {
              // Fetch remote and fall back to same-origin proxy on CORS failure
              let resp;
              try {
                resp = await fetch(url, {mode: 'cors'});
                if (!resp.ok) throw new Error(`status ${resp.status}`);
              } catch (corsErr) {
                const basePath = window.location.pathname.split('/p/')[0] || '';
                const proxyVariants = [
                  `${basePath}/ep_docx_image_proxy?url=${encodeURIComponent(url)}`,
                  `/ep_docx_image_proxy?url=${encodeURIComponent(url)}`,
                  `${window.location.origin}${basePath}/ep_docx_image_proxy?url=${encodeURIComponent(url)}`,
                ];
                let proxySuccess = false;
                for (const proxyUrl of proxyVariants) {
                  try {
                    resp = await fetch(proxyUrl);
                    if (resp.ok) { proxySuccess = true; break; }
                    if (DEBUG) console.warn(`[docx_customizer] proxy ${proxyUrl} returned ${resp.status}`);
                  } catch (proxyErr) {
                    if (DEBUG) console.warn(`[docx_customizer] proxy attempt failed ${proxyUrl}`, proxyErr);
                  }
                }
                if (!proxySuccess) throw new Error('All proxy attempts failed');
              }
              blob = await resp.blob();
              const urlName = new URL(url).pathname.split('/').pop() || filename;
              filename = urlName;
              if (!/\.[A-Za-z0-9]+$/.test(filename)) {
                const mimeExt = (blob.type && blob.type.split('/')[1]) || 'png';
                filename += `.${mimeExt}`;
              }
            } else {
              return; // unsupported scheme
            }

            const finalUrl = await uploadImageToS3Browser(blob, filename, padId);
            if (!finalUrl) throw new Error('S3 upload failed');

            const dim = await new Promise((res, rej) => {
              const imgObj = new Image();
              imgObj.onload = () => res({w: imgObj.naturalWidth, h: imgObj.naturalHeight});
              imgObj.onerror = rej;
              imgObj.src = URL.createObjectURL(blob);
            });
            if (dim && dim.w && dim.h) {
              const ratio = (dim.w / dim.h).toFixed(4);
              const hasWidthCls  = Array.from(sp.classList).some(c => c.startsWith('image-width:'));
              const hasHeightCls = Array.from(sp.classList).some(c => c.startsWith('image-height:'));
              const hasRatioCls  = Array.from(sp.classList).some(c => c.startsWith('imageCssAspectRatio:'));
              if (!hasWidthCls)  sp.classList.add(`image-width:${dim.w}px`);
              if (!hasHeightCls) sp.classList.add(`image-height:${dim.h}px`);
              if (!hasRatioCls)  sp.classList.add(`imageCssAspectRatio:${ratio}`);
            }
            sp.className = sp.className.replace(m[1], encodeURIComponent(finalUrl));
          } catch (e) {
            // S3 upload failed - replace image span with warning emoji (never fall back to base64)
            console.warn('[docx_customizer] S3 upload failed for image, replacing with warning:', url, e);
            const warning = document.createTextNode('⚠️');
            sp.parentNode.replaceChild(warning, sp);
          }
        }));
        return tmp.innerHTML;
      };

      inlineImages(cleanedHtml).then((finalHtml) => {
        context.ace.callWithAce((ace) => {
          let insideTableCell = false;
          try {
            const rep = ace.ace_getRep && ace.ace_getRep();
            const docMgr = ace.documentAttributeManager || ace.editorInfo?.documentAttributeManager;
            if (rep && rep.selStart && docMgr && docMgr.getAttributeOnLine) {
              const lineNum = rep.selStart[0];
              if (docMgr.getAttributeOnLine(lineNum, ATTR_TABLE_JSON)) insideTableCell = true;
            }
          } catch (_) {}
          if (!insideTableCell) {
            try {
              const innerWin = $innerIframe[0].contentWindow;
              const sel = innerWin.getSelection && innerWin.getSelection();
              if (sel && sel.rangeCount) {
                let n = sel.getRangeAt(0).startContainer;
                while (n) {
                  if (n.nodeType === 1 && n.matches && n.matches('table.dataTable')) {
                    insideTableCell = true;
                    break;
                  }
                  n = n.parentNode;
                }
              }
            } catch (_) {}
          }

          if (insideTableCell) {
            if (DEBUG) console.log('[docx_customizer] insideTableCell confirmed');
            try {
              const repNow = ace.ace_getRep();
              if (!repNow || !repNow.selStart) return;

              const selStart = repNow.selStart.slice();
              const selEnd = repNow.selEnd.slice();
              const lineNum = selStart[0];
              const lineText = repNow.lines.atIndex(lineNum)?.text || '';
              const cells = lineText.split(DELIMITER);
              let currentOffset = 0;
              let targetCellIndex = -1, cellEndCol = 0;
              for (let i = 0; i < cells.length; i++) {
                const cellLength = cells[i]?.length ?? 0;
                const cellEndThis = currentOffset + cellLength;
                if (selStart[1] >= currentOffset && selStart[1] <= cellEndThis) {
                  targetCellIndex = i;
                  cellEndCol = cellEndThis;
                  break;
                }
                currentOffset += cellLength + DELIMITER.length;
              }
              if (targetCellIndex === -1 || selEnd[1] > cellEndCol) return;

              const tempDivSeg = document.createElement('div');
              tempDivSeg.innerHTML = finalHtml;
              const segments = [];
              const extractSegmentsRecursive = (node, inheritedAttributes = []) => {
                if (node.nodeType === Node.TEXT_NODE) {
                  segments.push({text: node.textContent || '', attributes: inheritedAttributes});
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                  let currentAttributes = [...inheritedAttributes];
                  let isImageSpan = false;
                  
                  if (node.classList) {
                    for (const cls of node.classList) {
                      if (cls.startsWith('hyperlink-')) {
                        try { currentAttributes.push(['hyperlink', decodeURIComponent(cls.slice('hyperlink-'.length))]); } catch {}
                      } else if (cls.startsWith('color:')) {
                        currentAttributes.push(['color', cls.slice('color:'.length)]);
                      } else if (cls.startsWith('font-size:')) {
                        currentAttributes.push(['font-size', cls.slice('font-size:'.length)]);
                      } else if (cls === 'sup') {
                        currentAttributes.push(['sup', 'true']);
                        if (DEBUG) console.log('[docx_customizer] detected sup class');
                      } else if (cls === 'sub') {
                        currentAttributes.push(['sub', 'true']);
                        if (DEBUG) console.log('[docx_customizer] detected sub class');
                      }
                    }
                  }
                  
                  if (node.nodeName === 'A' && node.getAttribute('href')) {
                    currentAttributes.push(['hyperlink', node.getAttribute('href')]);
                  }
                  
                  // Direct tag-based detection for superscript/subscript
                  if (node.nodeName === 'SUP') {
                    currentAttributes.push(['sup', 'true']);
                    if (DEBUG) console.log('[docx_customizer] detected <SUP> tag');
                  } else if (node.nodeName === 'SUB') {
                    currentAttributes.push(['sub', 'true']);
                    if (DEBUG) console.log('[docx_customizer] detected <SUB> tag');
                  }
                  
                  if (node.classList && node.classList.contains('inline-image')) {
                    isImageSpan = true;
                    const imageMatch = Array.from(node.classList).find(c => c.startsWith('image:'));
                    if (imageMatch) {
                      try {
                        segments.push({
                          text: '\u200B', isImage: true,
                          imageUrl: decodeURIComponent(imageMatch.slice('image:'.length)),
                          imageClasses: node.className,
                          attributes: [],
                        });
                      } catch (e) {}
                    }
                  }
                  if (!isImageSpan) {
                    for (let i = 0; i < node.childNodes.length; i++) {
                      extractSegmentsRecursive(node.childNodes[i], currentAttributes);
                    }
                  }

                  // Detect style-based superscript/subscript (e.g., Google Docs uses vertical-align)
                  if (!currentAttributes.some(a => a[0]==='sup' || a[0]==='sub')) {
                    const vAlign = (node.style && node.style.verticalAlign || '').toLowerCase();
                    if (vAlign === 'super') {
                      currentAttributes.push(['sup','true']);
                      if (DEBUG) console.log('[docx_customizer] detected vertical-align:super');
                    } else if (vAlign === 'sub') {
                      currentAttributes.push(['sub','true']);
                      if (DEBUG) console.log('[docx_customizer] detected vertical-align:sub');
                    }
                  }
                }
              };
              for (let i = 0; i < tempDivSeg.childNodes.length; i++) {
                extractSegmentsRecursive(tempDivSeg.childNodes[i], []);
              }

              const cleanedSegments = segments.map(s => {
                s.text = s.isImage ? '\u200B' : (s.text || '').replace(/(\r\n|\n|\r)/gm, ' ').replace(new RegExp(DELIMITER, 'g'), ' ').replace(/\t/g, ' ').trim();
                return s;
              }).filter(s => s.text.length > 0 || s.isImage);
              if (cleanedSegments.length === 0) return;

              ace.ace_performDocumentReplaceRange(selStart, selEnd, '');
              let currentPos = ace.ace_getRep().selStart;

              for (let i = 0; i < cleanedSegments.length; i++) {
                const seg = cleanedSegments[i];
                
                // Add a space between segments if needed
                if (i > 0) {
                  const prevSeg = cleanedSegments[i - 1];
                  if (!seg.isImage && !prevSeg.isImage && seg.text.length > 0 && prevSeg.text.length > 0 &&
                      !/\s$/.test(prevSeg.text) && !/^\s/.test(seg.text)) {
                    ace.ace_performDocumentReplaceRange(currentPos, currentPos, ' ');
                    currentPos = ace.ace_getRep().selEnd;
                  }
                }

                ace.ace_performDocumentReplaceRange(currentPos, currentPos, seg.text);
                let endPos = ace.ace_getRep().selEnd;
                
                if (seg.attributes && seg.attributes.length > 0) {
                  if (DEBUG) {
                    const supAttr = seg.attributes.find(a => a[0]==='sup');
                    const subAttr = seg.attributes.find(a => a[0]==='sub');
                    if (supAttr) console.log('[docx_customizer] applying sup attribute to segment', seg.text);
                    if (subAttr) console.log('[docx_customizer] applying sub attribute to segment', seg.text);
                  }
                  ace.ace_performDocumentApplyAttributesToRange(currentPos, endPos, seg.attributes);
                }

                if (seg.isImage && seg.imageUrl) {
                  const imageAttrs = (seg.imageClasses || '').split(' ').map(cls => {
                      if (cls.startsWith('image-width:')) return ['image-width', cls.slice('image-width:'.length)];
                      if (cls.startsWith('image-height:')) return ['image-height', cls.slice('image-height:'.length)];
                      if (cls.startsWith('imageCssAspectRatio:')) return ['imageCssAspectRatio', cls.slice('imageCssAspectRatio:'.length)];
                      if (cls.startsWith('image-id-')) return ['image-id', cls.slice('image-id-'.length)];
                      return null;
                  }).filter(Boolean);
                  imageAttrs.push(['image', encodeURIComponent(seg.imageUrl)]);
                  ace.ace_performDocumentApplyAttributesToRange(currentPos, endPos, imageAttrs);
                }
                currentPos = ace.ace_getRep().selEnd;
              }
              ace.ace_fastIncorp && ace.ace_fastIncorp(10);
            } catch (errPaste) {
              console.error('[docx_customizer] error during rich-text table paste', errPaste);
            }
          } else {
            ace.ace_inCallStackIfNecessary('docxPaste', () => {
              const innerWin = $innerIframe[0].contentWindow;
              innerWin.document.execCommand('insertHTML', false, finalHtml);
            });
          }
        }, 'docxPaste', true);
      });
    });
  }, 'setupDocxCustomizerPaste', true);
};

exports.collectContentPre = (hookName, context) => {
  const {cls, cc, state} = context;
  if (!cls) return;

  const tblJsonClass = /(?:^| )(tbljson-[^ ]*)/.exec(cls);
  if (tblJsonClass) {
    const encoded = tblJsonClass[1].substring(8); // "tbljson-".length
    const decoded = dec(encoded);
    if (decoded) {
      cc.doAttrib(state, `${ATTR_TABLE_JSON}::${decoded}`);
    }
  }

  const tblCellClass = /(?:^| )(tblCell-[^ ]*)/.exec(cls);
  if (tblCellClass) {
    const cellIdx = tblCellClass[1].substring(8); // "tblCell-".length
    cc.doAttrib(state, `td::${cellIdx}`);
  }

  /* ─────────────────────────── Font-size normalisation ───────────────────────────
   * ep_font_size's collectContentPre wrongly encodes the value in the key
   * ("font-size:12") which prevents later edits from overwriting the old size.
   * If we spot a class that matches that legacy pattern we:
   *   1. Add the canonical attribute "font-size::N" so Etherpad stores the size
   *      in the value position (same shape the toolbar uses).
   *   2. Remove the legacy class from the class list before ep_font_size gets
   *      its turn, thereby preventing it from re-adding the broken attribute.
   *
   * This keeps the visual appearance identical but avoids duplicate
   * font-size attributes once the user changes the size via the toolbar.
   */
  const fsMatch = /(?:^| )font-size:([0-9]+)(?= |$)/.exec(cls);
  if (fsMatch && fsMatch[1]) {
    const sizeVal = fsMatch[1];
    // Inject correct attribute (key "font-size", value = <size>)
    cc.doAttrib(state, `font-size::${sizeVal}`);

    // Strip *all* occurrences of the legacy class so the downstream plugin
    // doesn't create an extra attribute with the value baked into the key.
    const clsClean = cls.replace(new RegExp(`(?:^| )font-size:${sizeVal}(?= |$)`, 'g'), ' ').trim();
    context.cls = clsClean; // propagate the mutation to later hooks
  }
};

exports.aceAttribsToClasses = (hook, context) => {
  if (context.key === ATTR_TABLE_JSON) {
    return [`tbljson-${enc(context.value)}`];
  }
  if (context.key === ATTR_CELL) {
    return [`tblCell-${context.value}`];
  }
  return [];
}; 