'use strict';

// Clipboard integration for ep_docx_html_customizer – simplified version.
// We now rely on Etherpad part ordering (our part loads before ep_hyperlinked_text)
// so we use the same pattern (jQuery .on('paste')) instead of capture-phase hacks.

const {customizeDocument, uploadImageToS3Browser} = require('../../transform_common');

// ADD: Constants matching ep_tables5 for table detection & delimiter cleanup
const ATTR_TABLE_JSON = 'tbljson';
const DELIMITER = '\u241F'; // same invisible delimiter used by ep_tables5
const ATTR_CELL = 'td';

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
      
      // Check if this is just raw image data - let ep_images_extended handle it
      const hasImageFiles = clipboardData.types.includes('Files') && clipboardData.files && clipboardData.files.length > 0;
      const hasHTML = clipboardData.types.includes('text/html');
      
      if (hasImageFiles && !hasHTML) {
        if (DEBUG) console.log('[docx_customizer] clipboard contains only image files, letting ep_images_extended handle it');
        return; // Let ep_images_extended handle raw image pastes
      }
      
      if (!hasHTML) {
        if (DEBUG) console.log('[docx_customizer] no HTML content, letting core handle plain text');
        return; // let core handle plain text
      }

      const html = clipboardData.getData('text/html');
      if (!html) {
        if (DEBUG) console.log('[docx_customizer] clipboard has no HTML data');
        return;
      }

      // Check if the HTML content actually needs our transformations
      // If it's just a simple image or doesn't contain our target elements, let other plugins handle it
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = html;
      
      const needsTransformation = (
        tempDiv.querySelector('h1, h2, h3, h4, h5, h6') ||          // headings
        tempDiv.querySelector('[style*="text-align"], [align]') ||   // alignment
        tempDiv.querySelector('ol') ||                               // ordered lists
        tempDiv.querySelector('a[href]') ||                          // hyperlinks
        tempDiv.querySelector('table') ||                            // tables
        tempDiv.querySelector('[style*="color"], font[color], [style*="font-size"]') || // styling
        (tempDiv.querySelector('img') && tempDiv.textContent.trim()) // images with text content
      );
      
      if (!needsTransformation) {
        if (DEBUG) console.log('[docx_customizer] HTML content does not need our transformations, skipping');
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

      // Inline remote images with S3 upload or data URI fallback
      const inlineImages = async (html) => {
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        const spans = tmp.querySelectorAll('span[class*="image:"]');
        
        // Show progress toast if there are images to process
        const imageSpans = Array.from(spans);
        const totalImages = imageSpans.length;
        let processedCount = 0;
        let successCount = 0;
        let errorCount = 0;
        
        if (totalImages > 0) {
          if (DEBUG) console.log(`[docx_customizer] Starting to process ${totalImages} images`);
          if (typeof window.docxToast !== 'undefined') {
            docxToast.showImageUploadProgress('clipboard-images', 'Processing Clipboard Images');
            docxToast.updateImageUploadProgress('clipboard-images', 0, totalImages);
          }
        }
        
        await Promise.all(imageSpans.map(async (sp) => {
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
              
              // Try multiple proxy URL variations to handle different deployment scenarios
              const basePath = window.location.pathname.split('/p/')[0];
              const proxyUrls = [
                `${basePath}/ep_docx_image_proxy?url=${encodeURIComponent(url)}`,
                `/ep_docx_image_proxy?url=${encodeURIComponent(url)}`,
                `${window.location.origin}${basePath}/ep_docx_image_proxy?url=${encodeURIComponent(url)}`
              ];
              
              let proxyWorked = false;
              for (const proxyUrl of proxyUrls) {
                if (DEBUG) console.log('[docx_customizer] trying proxy URL:', proxyUrl);
                
                try {
                  resp = await fetch(proxyUrl);
                  if (DEBUG) console.log('[docx_customizer] proxy response status:', resp.status);
                  
                  if (resp.ok) {
                    proxyWorked = true;
                    break;
                  } else {
                    if (DEBUG) console.warn(`[docx_customizer] proxy ${proxyUrl} returned ${resp.status}: ${resp.statusText}`);
                  }
                } catch (proxyErr) {
                  if (DEBUG) console.warn(`[docx_customizer] proxy ${proxyUrl} failed:`, proxyErr);
                }
              }
              
              if (!proxyWorked) {
                throw new Error(`All fetch methods failed. Direct CORS: ${corsErr.message}, Proxy attempts all failed`);
              }
            }

            const blob = await resp.blob();
            
            if (DEBUG) console.log('[docx_customizer] got blob:', blob.size, 'bytes, type:', blob.type);
            
            // Validate blob
            if (!blob || blob.size === 0) {
              throw new Error('Empty or invalid blob received from proxy');
            }
            
            // Try to upload to S3 - error out if it fails
            let finalUrl = null;
            try {
              // Generate a filename based on the original URL
              const urlPath = new URL(url).pathname;
              let filename = urlPath.split('/').pop() || `image-${Date.now()}`;
              
              // Ensure proper file extension
              if (!filename.includes('.')) {
                const ext = blob.type.split('/')[1] || 'png';
                filename += `.${ext}`;
              }
              
              if (DEBUG) console.log('[docx_customizer] uploading to S3 with filename:', filename, 'type:', blob.type);
              
              // Use the helper function to upload to S3
              const padId = (typeof clientVars !== 'undefined') ? clientVars.padId : 'clipboard';
              finalUrl = await uploadImageToS3Browser(blob, filename, padId);
              
              if (finalUrl && DEBUG) {
                console.log('[docx_customizer] uploaded remote image to S3', url, '->', finalUrl);
              } else {
                // Error out if S3 upload fails - no data URI fallback
                throw new Error('S3 upload returned null - upload failed or not configured');
              }
            } catch (s3Err) {
              if (DEBUG) console.error('[docx_customizer] S3 upload failed, skipping image', url, s3Err);
              errorCount++;
              processedCount++;
              if (totalImages > 0 && typeof window.docxToast !== 'undefined') {
                docxToast.updateImageUploadProgress('clipboard-images', processedCount, totalImages);
              }
              return; // Skip this image entirely
            }
            
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
                const hasWidthCls  = Array.from(sp.classList).some(c => c.startsWith('image-width:'));
                const hasHeightCls = Array.from(sp.classList).some(c => c.startsWith('image-height:'));
                const hasRatioCls  = Array.from(sp.classList).some(c => c.startsWith('imageCssAspectRatio:'));

                if (!hasWidthCls)  sp.classList.add(`image-width:${dim.w}px`);
                if (!hasHeightCls) sp.classList.add(`image-height:${dim.h}px`);
                if (!hasRatioCls)  sp.classList.add(`imageCssAspectRatio:${ratio}`);
              }
            } catch (_) { /* ignore failures */ }
            
            const encoded = encodeURIComponent(finalUrl);
            sp.className = sp.className.replace(m[1], encoded);
            // `customizeDocument()` already wrapped the placeholder with
            // a single ZWSP on each side.  No extra normalisation needed.
            
            successCount++;
            processedCount++;
            if (totalImages > 0 && typeof window.docxToast !== 'undefined') {
              docxToast.updateImageUploadProgress('clipboard-images', processedCount, totalImages);
            }
            
          } catch (e) {
            if (DEBUG) console.warn('[docx_customizer] failed to inline', url, e);
            errorCount++;
            processedCount++;
            if (totalImages > 0 && typeof window.docxToast !== 'undefined') {
              docxToast.updateImageUploadProgress('clipboard-images', processedCount, totalImages);
            }
          }
        }));
        
        // Show completion notification
        if (totalImages > 0) {
          if (typeof window.docxToast !== 'undefined') {
            docxToast.completeImageUpload('clipboard-images', successCount, totalImages, errorCount);
          }
          if (DEBUG) console.log(`[docx_customizer] Image processing complete: ${successCount} uploaded, ${errorCount} failed`);
        }
        
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
              let selEnd   = repNow.selEnd.slice();
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

              // === Extract hyperlink-aware segments from HTML (borrowed from ep_hyperlinked_text) ===
              const tempDivSeg = document.createElement('div');
              tempDivSeg.innerHTML = finalHtml;

              const segments = [];
              const extractSegmentsRecursive = (node, inheritedUrl) => {
                if (node.nodeType === Node.TEXT_NODE) {
                  segments.push({text: node.textContent || '', url: inheritedUrl});
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                  let currentUrl = inheritedUrl;
                  let isImageSpan = false;
                  
                  // Check for hyperlinks
                  if (node.nodeName === 'A' && node.getAttribute('href')) {
                    let href = node.getAttribute('href');
                    if (href && href.trim() !== '' && !href.trim().toLowerCase().startsWith('javascript:')) {
                      if (!/^(https?:\/\/|mailto:|ftp:|file:|#|\/)/i.test(href)) {
                        href = `http://${href}`;
                      }
                      currentUrl = href;
                    }
                  }
                  
                  // Detect span with hyperlink classes (produced by customizeDocument)
                  if (!currentUrl && node.classList) {
                    const m = Array.from(node.classList).find(c => c.startsWith('hyperlink-'));
                    if (m) {
                      try { currentUrl = decodeURIComponent(m.slice('hyperlink-'.length)); } catch {}
                    }
                  }
                  
                  // Detect image spans (produced by customizeDocument and inlineImages)
                  if (node.classList && node.classList.contains('inline-image')) {
                    isImageSpan = true;
                    // Extract image URL from class
                    const imageMatch = Array.from(node.classList).find(c => c.startsWith('image:'));
                    if (imageMatch) {
                      try {
                        const imageUrl = decodeURIComponent(imageMatch.slice('image:'.length));
                        // Add image as a special text segment that will be converted to image span
                        segments.push({
                          text: '\u200B', // ZWSP placeholder
                          url: null,
                          isImage: true,
                          imageUrl: imageUrl,
                          imageClasses: node.className
                        });
                      } catch (e) {
                        if (DEBUG) console.warn('[docx_customizer] failed to extract image URL from class', e);
                      }
                    }
                  }
                  
                  // Don't recurse into image spans (they just contain ZWSP)
                  if (!isImageSpan) {
                    for (let i = 0; i < node.childNodes.length; i++) {
                      extractSegmentsRecursive(node.childNodes[i], currentUrl);
                    }
                  }
                }
              };

              for (let i = 0; i < tempDivSeg.childNodes.length; i++) {
                extractSegmentsRecursive(tempDivSeg.childNodes[i], null);
              }

              if (segments.length === 0 && tempDivSeg.textContent) {
                segments.push({text: tempDivSeg.textContent, url: null});
              }

              // Sanitize each segment and filter empties
              segments.forEach((seg) => {
                if (seg.isImage) {
                  // For image segments, preserve the ZWSP character
                  seg.text = '\u200B';
                } else {
                  seg.text = (seg.text || '')
                    .replace(/(\r\n|\n|\r)/gm, ' ')
                    .replace(new RegExp(DELIMITER, 'g'), ' ')
                    .replace(/\t/g, ' ')
                    .replace(/\s+/g, ' ')
                    .trim();
                }
              });
              const cleanedSegments = segments.filter((s) => s.text.length > 0 || s.isImage);
              if (cleanedSegments.length === 0) {
                if (DEBUG) console.log('[docx_customizer] No text after sanitization – abort paste');
                return;
              }

              // Enforce maximum cell length like ep_tables5
              const selectionLength = selEnd[1] - selStart[1];
              const currentCellText = cells[targetCellIndex] || '';
              const MAX_CELL_LENGTH = 8000;
              let remaining = MAX_CELL_LENGTH - (currentCellText.length - selectionLength);
              if (remaining <= 0) {
                if (DEBUG) console.log('[docx_customizer] Cell at max length – abort paste');
                return;
              }

              // ================= EXACT insertion loop from ep_hyperlinked_text (isTableLine forced true) =================
              let selStartMod = selStart.slice();
              let selEndMod   = selEnd.slice();

              if (selStartMod[0] !== selEndMod[0] || selStartMod[1] !== selEndMod[1]) {
                ace.ace_performDocumentReplaceRange(selStartMod, selEndMod, '');
                selEndMod = selStartMod.slice();
              }

              let currentLine = selStartMod[0];
              let currentCol  = selStartMod[1];

              if (DEBUG) console.log('[docx_customizer] cleanedSegments', cleanedSegments.length, cleanedSegments);

              for (let segIdx = 0; segIdx < cleanedSegments.length; segIdx++) {
                const segment = cleanedSegments[segIdx];
                let textToInsert = segment.text;
                textToInsert = textToInsert.replace(/\n+/g, ' '); // table line flatten

                // Conditional leading space (same as hyperlinked plugin)
                if (segIdx > 0) {
                  const previousSegment = cleanedSegments[segIdx - 1];
                  /* Insert a literal space only when BOTH neighbouring segments are plain text.
                     Image placeholders must keep their exact ZWSP-…-ZWSP pattern or ep_images_extended
                     cannot detect them for resize/float handling. */
                  if (!segment.isImage && !previousSegment.isImage &&
                      previousSegment.text.length > 0 && textToInsert.length > 0 &&
                      !/\s$/.test(previousSegment.text) && !/^\s/.test(textToInsert)) {
                      ace.ace_performDocumentReplaceRange([currentLine, currentCol],
                                                          [currentLine, currentCol], ' ');
                    const repAfterSpace = ace.ace_getRep();
                    currentLine = repAfterSpace.selEnd[0];
                    currentCol  = repAfterSpace.selEnd[1];
                  }
                }

                if (textToInsert.length > 0) {
                  if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: inserting ${segment.isImage ? 'image' : 'text'} "${textToInsert.slice(0,80)}" (len ${textToInsert.length}) url=${segment.url || 'none'} at L${currentLine}C${currentCol}`);
                  const insertStart = [currentLine, currentCol];
                  ace.ace_performDocumentReplaceRange(insertStart, insertStart, textToInsert);

                  const repAfterTxt = ace.ace_getRep();
                  const insertEndLine = repAfterTxt.selEnd[0];
                  const insertEndCol  = repAfterTxt.selEnd[1];

                  // Apply hyperlink attribute
                  if (segment.url) {
                    if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: applying hyperlink attr to range L${insertStart[0]}C${insertStart[1]} - L${insertEndLine}C${insertEndCol}`);
                    try {
                      ace.ace_performDocumentApplyAttributesToRange(insertStart, [insertEndLine, insertEndCol], [['hyperlink', segment.url]]);
                    } catch (attrErr) {
                      console.warn('[docx_customizer] Failed to apply hyperlink attribute', attrErr);
                    }
                  }

                  // Apply image attributes if this is an image segment
                  if (segment.isImage && segment.imageUrl) {
                    if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: applying image attributes for ${segment.imageUrl}`);
                    if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: imageClasses = "${segment.imageClasses}"`);
                    try {
                      // Parse image classes to extract individual attributes
                      const classes = (segment.imageClasses || '').split(' ');
                      const imageAttrs = [];
                      
                      // Add the main image URL – ensure it is percent-encoded so the resulting
                      // `image:` class contains a valid CSS identifier.  This matters especially
                      // for images pasted into ep_tables5 cells where the URL was previously left
                      // unencoded ("image:https://…") and therefore broke resize handling.
                      const encodedUrl = encodeURIComponent(segment.imageUrl);
                      imageAttrs.push(['image', encodedUrl]);
                      
                      // Extract other image attributes from classes
                      for (const cls of classes) {
                        if (cls.startsWith('image-width:')) {
                          imageAttrs.push(['image-width', cls.slice('image-width:'.length)]);
                        } else if (cls.startsWith('image-height:')) {
                          imageAttrs.push(['image-height', cls.slice('image-height:'.length)]);
                        } else if (cls.startsWith('imageCssAspectRatio:')) {
                          imageAttrs.push(['imageCssAspectRatio', cls.slice('imageCssAspectRatio:'.length)]);
                        } else if (cls.startsWith('image-id-')) {
                          imageAttrs.push(['image-id', cls.slice('image-id-'.length)]);
                        }
                      }
                      
                      if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: extracted image attributes:`, imageAttrs);
                      
                      // Apply all image attributes
                      if (imageAttrs.length > 0) {
                        ace.ace_performDocumentApplyAttributesToRange(insertStart, [insertEndLine, insertEndCol], imageAttrs);
                        if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: successfully applied ${imageAttrs.length} image attributes`);
                      } else {
                        if (DEBUG) console.warn(`[docx_customizer] seg${segIdx}: no image attributes to apply`);
                      }
                    } catch (attrErr) {
                      console.warn('[docx_customizer] Failed to apply image attributes', attrErr);
                    }
                  } else if (segment.isImage) {
                    if (DEBUG) console.warn(`[docx_customizer] seg${segIdx}: image segment missing required data - imageUrl: ${!!segment.imageUrl}, imageClasses: ${!!segment.imageClasses}`);
                  }

                  if (DEBUG) console.log(`[docx_customizer] seg${segIdx}: after insertion caret at L${insertEndLine}C${insertEndCol}`);
                  currentLine = insertEndLine;
                  currentCol  = insertEndCol;
                }
              }
              if (DEBUG) console.log('[docx_customizer] Finished segment loop. Final caret', currentLine, currentCol);

              // Reapply tbljson metadata after paste
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

              ace.ace_performSelectionChange([currentLine, currentCol], [currentLine, currentCol], false);
              ace.ace_fastIncorp && ace.ace_fastIncorp(10);

              if (DEBUG) console.log('[docx_customizer] Hyperlink-aware table-cell paste completed');

            } catch (errPaste) {
              console.error('[docx_customizer] error during plain-text table paste', errPaste);
            }
          } else {
            // Normal (non-table) insertion – keep HTML & formatting
            ace.ace_inCallStackIfNecessary('docxPaste', () => {
              // 1) Let the browser paste the HTML so spans with tbljson-* classes
              //    land in the DOM and Ace converts them into atext.
              const innerWin = $innerIframe[0].contentWindow;
              innerWin.document.execCommand('insertHTML', false, finalHtml);

              // 2) Sync to Ace's internal document model.
              ace.ace_fastIncorp && ace.ace_fastIncorp(10);

              const rep     = ace.ace_getRep();
              const docMgr  = ace.documentAttributeManager || ace.editorInfo?.documentAttributeManager;
              if (!rep || !docMgr) return;

              // Determine a conservative line range that likely contains the newly
              // inserted block – from the first line of the current selection to
              // the current end of the document.
              const firstLine = (rep.selStart && Array.isArray(rep.selStart)) ? rep.selStart[0] : 0;
              const lastLine  = rep.lines.length() - 1;

              for (let ln = firstLine; ln <= lastLine; ln++) {
                const attrStr = docMgr.getAttributeOnLine(ln, ATTR_TABLE_JSON);
                if (!attrStr) continue; // not a table row

                let meta;
                try { meta = JSON.parse(attrStr); } catch (_) {/* ignore */}

                const lineText = rep.lines.atIndex(ln).text || '';
                const cells    = lineText.split(DELIMITER);

                // 3) Apply the per-cell td attribute so Etherpad produces author spans.
                let offset = 0;
                cells.forEach((cellTxt, idx) => {
                  if (cellTxt.length > 0) {
                    ace.ace_performDocumentApplyAttributesToRange(
                      [ln, offset], [ln, offset + cellTxt.length], [[ATTR_CELL, String(idx)]]);
                  }
                  offset += cellTxt.length;
                  if (idx < cells.length - 1) offset += DELIMITER.length;
                });

                // 4) Re-assert tbljson line attribute via official helper to make
                //    sure column-width metadata is stored.
                if (ace.ep_tables5_applyMeta && meta && typeof meta.cols === 'number') {
                  ace.ep_tables5_applyMeta(
                    ln,
                    meta.tblId,
                    meta.row,
                    meta.cols,
                    rep,
                    ace,
                    null,
                    docMgr,
                  );
                }
              }

              // Final sync so the attribute changes are flushed.
              ace.ace_fastIncorp && ace.ace_fastIncorp(10);
            });
          }
        }, 'docxPaste', true);
      });

      if (DEBUG) console.log('[docx_customizer] insertion done, event propagation stopped');
    });

    if (DEBUG) console.log('[docx_customizer] paste listener attached');
  }, 'setupDocxCustomizerPaste', true);
}; 