'use strict';

const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const { exec } = require('child_process');
const os = require('os');
const log4js = require('log4js');
const { JSDOM } = require('jsdom');
const settings = require('ep_etherpad-lite/node/utils/Settings');
const util = require('util');
const execPromise = util.promisify(exec);
const mime = require('mime');

const logger = log4js.getLogger('ep_docx_html_customizer');

// Helper for stable random ids
const rand = () => Math.random().toString(36).slice(2, 8);

// encode/decode so JSON can survive as a CSS class token if ever needed
const enc = (s) => {
  if (typeof btoa === 'function') {
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_'); // Browser
  } else if (typeof Buffer === 'function') {
    return Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'); // Node.js
  }
  logger.warn('[ep_docx_html_customizer] Base64 encoding function (btoa or Buffer) not found.');
  return s; // Fallback
};

// Use the same rare delimiter as ep_tables5 so the client can split cells reliably
const DELIMITER = '\u241F'; // ␟ Unit Separator – internal column delimiter

/**
 * Import hook
 * Handles the import of DOCX, DOC, ODT, and ODF files, customizing their HTML output.
 *
 * @param {string} hookName Hook name ("import").
 * @param {object} context Object containing the arguments passed to hook {srcFile, fileEnding, destFile, padId, ImportError}.
 * @param {function} cb Callback
 *
 * @returns {boolean} true if hook handled the import, false otherwise.
 */
exports.import = async (hookName, context) => {
  const { srcFile, fileEnding, destFile } = context;
  logger.info(`[ep_docx_html_customizer] Import hook called for file: ${srcFile}, type: ${fileEnding}`);

  const convertibleTypes = ['.docx', '.doc', '.odt', '.odf'];
  const ZWSP = '\u200B';

  if (!convertibleTypes.includes(fileEnding.toLowerCase())) {
    logger.info(`[ep_docx_html_customizer] File type ${fileEnding} is not supported. Passing to core or other plugins.`);
    return false; // Let Etherpad core or other plugins handle it
  }

  try {
    // Phase 1: LibreOffice Conversion
    logger.info(`[ep_docx_html_customizer] Attempting to convert ${srcFile} to HTML using LibreOffice.`);
    if (!settings.soffice) {
      logger.warn('[ep_docx_html_customizer] soffice path not configured in settings.json. Cannot convert document.');
      return false;
    }

    const converterPath = settings.soffice;
    const outDir = path.dirname(destFile);
    const tempConvertedBaseName = `${path.basename(srcFile, fileEnding)}.html`;
    const actualConvertedTmpFile = path.join(outDir, tempConvertedBaseName);

    const conversionCommand = `"${converterPath}" --headless --invisible --nologo --nolockcheck --writer --convert-to html "${srcFile}" --outdir "${outDir}"`;
    logger.debug(`[ep_docx_html_customizer] Executing soffice command: ${conversionCommand}`);

    await execPromise(conversionCommand);
    logger.info(`[ep_docx_html_customizer] LibreOffice conversion successful. HTML output at: ${actualConvertedTmpFile}`);

    if (!fs.existsSync(actualConvertedTmpFile)) {
      logger.error(`[ep_docx_html_customizer] Conversion failed: ${actualConvertedTmpFile} not found after soffice execution.`);
      return false;
    }

    if (actualConvertedTmpFile !== destFile) {
      logger.debug(`[ep_docx_html_customizer] Renaming ${actualConvertedTmpFile} to ${destFile}`);
      await fsp.rename(actualConvertedTmpFile, destFile);
    }

    logger.info(`[ep_docx_html_customizer] Processing converted HTML at ${destFile} for images.`);
    let htmlContent = await fsp.readFile(destFile, 'utf8');
    const dom = new JSDOM(htmlContent);
    const document = dom.window.document;
    const images = document.querySelectorAll('img');
    let modified = false;

    logger.debug(`[ep_docx_html_customizer] Found ${images.length} image(s) in converted HTML: ${destFile}`);

    /* ─────────────────────────────── Headings processing ─────────────────────────────── */
    logger.info(`[ep_docx_html_customizer] Processing converted HTML for headings.`);
    
    const headingElements = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, h6'));
    logger.debug(`[ep_docx_html_customizer] Found ${headingElements.length} heading element(s).`);
    
    headingElements.forEach((heading, idx) => {
      // Ensure line break after each heading to prevent content merging
      const parent = heading.parentNode;
      const br = document.createElement('br');
      
      if (heading.nextSibling) {
        parent.insertBefore(br, heading.nextSibling);
      } else {
        parent.appendChild(br);
      }
      
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Added line break after ${heading.tagName.toLowerCase()} element ${idx + 1}.`);
    });

    /* ───────────────────────────── Alignment processing ───────────────────────────── */
    logger.info(`[ep_docx_html_customizer] Processing converted HTML for paragraph alignment.`);

    const ALIGN_MAP = {
      'center': 'center',
      'right': 'right', 
      'justify': 'justify',
      'left': 'left',
    };

    // Process alignment BEFORE image processing to avoid disrupting ZWSP structure
    const alignedBlocks = Array.from(document.querySelectorAll('[style*="text-align"], [align]'));
    logger.debug(`[ep_docx_html_customizer] Found ${alignedBlocks.length} elements with alignment style/attr.`);

    alignedBlocks.forEach((blk, idx) => {
      // Skip if already processed
      const tagNameLc = blk.tagName.toLowerCase();
      if (['center','left','right','justify'].includes(tagNameLc)) return;

      // Skip if inside a table (we ignore alignment inside tables)
      if (blk.closest('table')) return;

      // Determine align value
      let alignVal = (blk.getAttribute('align') || '').toLowerCase();
      if (!alignVal) {
        const styleAlign = /text-align\s*:\s*(left|right|center|justify)/i.exec(blk.getAttribute('style') || '');
        if (styleAlign) alignVal = styleAlign[1].toLowerCase();
      }

      if (!ALIGN_MAP[alignVal] || alignVal === 'left') return; // left is default → skip

      const wrapperTag = ALIGN_MAP[alignVal];
      const newEl = document.createElement(wrapperTag);

      // Preserve the HTML content instead of just moving nodes
      newEl.innerHTML = blk.innerHTML;

      // Replace the block and ensure line break after
      const parent = blk.parentNode;
      parent.replaceChild(newEl, blk);
      
      // Add explicit line break after aligned block to prevent merging
      const br = document.createElement('br');
      if (newEl.nextSibling) {
        parent.insertBefore(br, newEl.nextSibling);
      } else {
        parent.appendChild(br);
      }
      
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Wrapped element ${idx + 1} with <${wrapperTag}> and added line break.`);
    });

    /* ─────────────────────────────── Image processing ─────────────────────────────── */
    for (const [index, img] of images.entries()) {
      let imgSrc = img.getAttribute('src');
      logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1}/${images.length}: Original src="${imgSrc}"`);

      if (!imgSrc) {
        logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} has no src, skipping.`);
        continue;
      }

      if (imgSrc && !imgSrc.startsWith('http') && !imgSrc.startsWith('data:') && !imgSrc.startsWith('/')) {
        const imagePath = path.resolve(path.dirname(destFile), imgSrc);
        logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} is relative. Attempting to read: ${imgSrc} at ${imagePath}`);
        try {
          if (fs.existsSync(imagePath)) {
            const imageBuffer = await fsp.readFile(imagePath);
            const mimeType = mime.getType(imagePath) || 'application/octet-stream';
            imgSrc = `data:${mimeType};base64,${imageBuffer.toString('base64')}`;
            logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} successfully converted to data URI.`);
          } else {
            logger.warn(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} relative path not found: ${imagePath}.`);
            continue;
          }
        } catch (e) {
          logger.error(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} error reading/converting ${imagePath}: ${e.message}`);
          continue;
        }
      } else {
        logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} src is not relative or already data/http: "${imgSrc ? imgSrc.substring(0,50)+'...' : 'EMPTY'}"`);
      }

      const outerSpan = document.createElement('span');
      outerSpan.textContent = ZWSP;
      let outerClasses = 'inline-image character image-placeholder';
      
      outerClasses += ` image:${encodeURIComponent(imgSrc)}`;

      const imgWidth = img.getAttribute('width') || img.style.width;
      const imgHeight = img.getAttribute('height') || img.style.height;
      if (imgWidth) outerClasses += ` image-width:${/^[0-9]+(\.\d+)?$/.test(imgWidth) ? `${imgWidth}px` : imgWidth}`;
      if (imgHeight) outerClasses += ` image-height:${/^[0-9]+(\.\d+)?$/.test(imgHeight) ? `${imgHeight}px` : imgHeight}`;
      
      const numWidth = parseFloat(imgWidth);
      const numHeight = parseFloat(imgHeight);
      if (!isNaN(numWidth) && numWidth > 0 && !isNaN(numHeight) && numHeight > 0) {
        outerClasses += ` imageCssAspectRatio:${(numWidth / numHeight).toFixed(4)}`;
      }

      // NEW – add persistent image identifier expected by ep_image_insert (must be >10 chars)
      const imageId = typeof crypto !== 'undefined' && crypto.randomUUID ? crypto.randomUUID() :
                      (Math.random().toString(36).slice(2) + '-' + Math.random().toString(36).slice(2, 10));
      outerClasses += ` image-id-${imageId}`;
      outerSpan.setAttribute('data-image-id', imageId);

      outerSpan.className = outerClasses.trim();
      
      const fragment = document.createDocumentFragment();
      fragment.appendChild(document.createTextNode(ZWSP));
      fragment.appendChild(outerSpan);
      fragment.appendChild(document.createTextNode(ZWSP));

      img.parentNode.replaceChild(fragment, img);
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} replaced with ZWSP-span-ZWSP structure.`);
    }

    /* ───────────────────────────────── Hyperlink processing ───────────────────────────────── */
    logger.info(`[ep_docx_html_customizer] Processing converted HTML at ${destFile} for hyperlinks.`);
    const anchors = document.querySelectorAll('a[href]');
    logger.debug(`[ep_docx_html_customizer] Found ${anchors.length} <a> tag(s) in converted HTML.`);

    anchors.forEach((a, idx) => {
      const rawHref = a.getAttribute('href') || '';
      if (!rawHref.trim()) return; // Skip empty href

      // Sanitise href – add protocol if missing (same heuristic as client plugin)
      let href = rawHref.trim();
      if (!/^(https?:\/\/|mailto:|ftp:|file:|#|\/)/i.test(href)) href = `http://${href}`;

      const encodedHref = encodeURIComponent(href);

      const span = document.createElement('span');
      span.className = `hyperlink hyperlink-${encodedHref}`;
      span.textContent = a.textContent || href;

      // Wrap with ZWSP before & after, mirroring images & plugin expectation
      const frag = document.createDocumentFragment();
      frag.appendChild(document.createTextNode(ZWSP));
      frag.appendChild(span);
      frag.appendChild(document.createTextNode(ZWSP));

      a.parentNode.replaceChild(frag, a);
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Replaced <a> tag ${idx + 1} with hyperlink span.`);
    });

    /* ────────────────────────── Font color & size processing ────────────────────────── */
    logger.info(`[ep_docx_html_customizer] Processing converted HTML at ${destFile} for font colors and sizes.`);
    
    const COLOR_CLASSES = ['black', 'red', 'green', 'blue', 'yellow', 'orange'];
    
    // Helper to parse CSS color values to RGB
    const parseColorToRgb = (colorStr) => {
      if (!colorStr) return null;
      const str = colorStr.toLowerCase().trim();
      
      // Handle named colors
      const namedColors = {
        black: [0, 0, 0], white: [255, 255, 255], red: [255, 0, 0],
        green: [0, 128, 0], blue: [0, 0, 255], yellow: [255, 255, 0],
        orange: [255, 165, 0], purple: [128, 0, 128], gray: [128, 128, 128],
        grey: [128, 128, 128], lime: [0, 255, 0], cyan: [0, 255, 255],
        magenta: [255, 0, 255], maroon: [128, 0, 0], navy: [0, 0, 128]
      };
      if (namedColors[str]) return namedColors[str];
      
      // Handle hex colors
      let hexMatch = str.match(/^#?([0-9a-f]{6}|[0-9a-f]{3})$/);
      if (hexMatch) {
        let hex = hexMatch[1];
        if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
        return [
          parseInt(hex.slice(0, 2), 16),
          parseInt(hex.slice(2, 4), 16),
          parseInt(hex.slice(4, 6), 16)
        ];
      }
      
      // Handle rgb() colors
      const rgbMatch = str.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
      if (rgbMatch) {
        return [parseInt(rgbMatch[1]), parseInt(rgbMatch[2]), parseInt(rgbMatch[3])];
      }
      
      return null;
    };
    
    // Helper to find nearest color class by RGB distance
    const findNearestColor = (rgb) => {
      if (!rgb) return null;
      const colorPalette = {
        black: [0, 0, 0], red: [255, 0, 0], green: [0, 128, 0],
        blue: [0, 0, 255], yellow: [255, 255, 0], orange: [255, 165, 0]
      };
      
      let bestColor = 'black';
      let bestDistance = Infinity;
      
      for (const [colorName, paletteRgb] of Object.entries(colorPalette)) {
        const distance = Math.sqrt(
          Math.pow(rgb[0] - paletteRgb[0], 2) +
          Math.pow(rgb[1] - paletteRgb[1], 2) +
          Math.pow(rgb[2] - paletteRgb[2], 2)
        );
        if (distance < bestDistance) {
          bestDistance = distance;
          bestColor = colorName;
        }
      }
      
      return bestColor;
    };

    const parsePx = (val) => {
      if (!val) return null;
      const m = /([0-9.]+)(px|pt)?/i.exec(val);
      if (!m) return null;
      const num = parseFloat(m[1]);
      if (isNaN(num)) return null;
      if (m[2] && m[2].toLowerCase() === 'pt') return Math.round(num * 1.333); // rough pt→px
      return Math.round(num);
    };

    const allStyledNodes = Array.from(document.querySelectorAll('[style*="color"], font[color], [style*="font-size"]'));
    logger.debug(`[ep_docx_html_customizer] Found ${allStyledNodes.length} nodes with inline color/size.`);

    allStyledNodes.forEach((el, idx) => {
      // Get existing classes to preserve them
      const existingClasses = (el.className || '').split(/\s+/).filter(cls => cls.trim());
      const newClasses = [...existingClasses];
      let modified = false;

      // Process color
      const colorAttr = el.getAttribute('color') || el.style.color || '';
      if (colorAttr) {
        // Remove any existing color: classes
        const filteredClasses = newClasses.filter(cls => !cls.startsWith('color:'));
        const rgb = parseColorToRgb(colorAttr);
        const mappedColor = findNearestColor(rgb);
        
        if (mappedColor) {
          filteredClasses.push(`color:${mappedColor}`);
          newClasses.length = 0;
          newClasses.push(...filteredClasses);
          modified = true;
          logger.debug(`[ep_docx_html_customizer] Node ${idx + 1}: mapped color "${colorAttr}" to "${mappedColor}"`);
        }
      }

      // Process font size
      const sizeStyle = el.style.fontSize || '';
      if (sizeStyle) {
        // Remove any existing font-size: classes
        const filteredClasses = newClasses.filter(cls => !cls.startsWith('font-size:'));
        const px = parsePx(sizeStyle);
        
        if (px) {
          // Map to nearest from font size palette
          const sizePalette = [8,9,10,11,12,13,14,15,16,17,18,19,20,22,24,26,28,30,35,40,45,50,60];
          let best = sizePalette[0];
          let bestDiff = Math.abs(px - best);
          for (const s of sizePalette) {
            const diff = Math.abs(px - s);
            if (diff < bestDiff) { best = s; bestDiff = diff; }
          }
          
          filteredClasses.push(`font-size:${best}`);
          newClasses.length = 0;
          newClasses.push(...filteredClasses);
          modified = true;
          logger.debug(`[ep_docx_html_customizer] Node ${idx + 1}: mapped font-size "${sizeStyle}" (${px}px) to "${best}"`);
        }
      }

      if (modified) {
        // Update the element's classes directly, preserving other classes like hyperlink, image-id, etc.
        el.className = newClasses.join(' ');
        
        // Remove the inline styles that we've converted to classes
        if (colorAttr && el.style.color) el.style.removeProperty('color');
        if (sizeStyle && el.style.fontSize) el.style.removeProperty('font-size');
        if (el.hasAttribute('color')) el.removeAttribute('color');
        
        logger.debug(`[ep_docx_html_customizer] Node ${idx + 1}: updated classes to "${el.className}"`);
      }
    });

    if (allStyledNodes.length > 0) {
      modified = true;
      logger.info(`[ep_docx_html_customizer] Processed ${allStyledNodes.length} styled nodes for color/size.`);
    }

    // Phase 3: Table Processing (mimicking ep_tables5/tableImport.js)
    logger.info(`[ep_docx_html_customizer] Processing converted HTML at ${destFile} for tables.`);
    const tables = document.querySelectorAll('table');
    logger.debug(`[ep_docx_html_customizer] Found ${tables.length} table(s) in converted HTML.`);

    if (tables.length > 0) {
      const tblId = rand(); // Generate one ID for all tables in this import for simplicity, or per table
      let tableCount = 0;

      tables.forEach((tableNode, tableIndex) => {
        logger.debug(`[ep_docx_html_customizer] Processing table ${tableIndex + 1}/${tables.length}`);
        const rows = tableNode.querySelectorAll('tr');
        if (rows.length === 0) {
          logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1} has no rows, skipping.`);
          return;
        }

        const numCols = rows[0] ? Array.from(rows[0].querySelectorAll('td, th')).length : 0;
        if (numCols === 0) {
            logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1} has no columns in the first row, skipping.`);
            return;
        }

        const tableLines = [];

        rows.forEach((rowNode, rowIndex) => {
          const cells = Array.from(rowNode.querySelectorAll('td, th'));
          // Ensure all rows have the same number of columns as the first row, pad if necessary
          const cellContents = Array.from({ length: numCols }, (_, cellIdx) => {
            const cell = cells[cellIdx];
            // Obtain the raw HTML inside the cell and normalise whitespace/newlines.
            let cellHTML = cell ? cell.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim() : '';

            // Remove any <p> tags that might be wrapping content within cells after LibreOffice conversion.
            // This is a common artifact we want to flatten.
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = cellHTML;
            const pTags = tempDiv.querySelectorAll('p');
            if (pTags.length === 1 && tempDiv.textContent.trim() === pTags[0].textContent.trim()) {
              cellHTML = pTags[0].innerHTML.replace(/\r\n|\r|\n/g, ' ').trim();
            } else if (pTags.length > 0) {
              // If multiple p tags exist, join their content as a simple heuristic.
              cellHTML = Array.from(pTags).map(p => p.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim()).join(' ');
            }

            // After flattening <p> wrappers, strip lone <br> tags (LibreOffice's way
            // of representing empty cells). Treat cells with only <br> as truly empty.
            if (cellHTML === '<br>' || cellHTML === '<br/>' || cellHTML === '<br />') {
              cellHTML = '';
              logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1}, Row ${rowIndex + 1}, Col ${cellIdx + 1}: stripped lone <br> tag, treating as empty`);
            }

            // If the cell is empty after all normalisation, inject a non-breaking space in a span.
            // Using &nbsp; prevents the browser from collapsing whitespace and closely matches the
            // markup generated by ep_tables5 for blank cells.
            if (!cellHTML) {
              cellHTML = '<span>&nbsp;</span>';
              logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1}, Row ${rowIndex + 1}, Col ${cellIdx + 1}: inserted &nbsp; placeholder span in empty cell`);
            }
            return cellHTML;
          });

          const lineText = cellContents.join(DELIMITER);
          const metadata = {
            tblId: `${tblId}-${tableIndex}`, // Unique ID per table
            row: rowIndex,
            cols: numCols,
          };
          const attributeString = JSON.stringify(metadata);
          const encodedJson = enc(attributeString);

          // Create a new div element for the line
          const lineDiv = document.createElement('div');
          // Add the tbljson attribute as a class.
          // This class will be picked up by aceAttribsToClasses in client_hooks.js
          lineDiv.className = `tbljson-${encodedJson}`;
          // Set the innerHTML of the div to the delimited text
          // This text will be used by acePostWriteDomLineHTML to render the table
          lineDiv.innerHTML = lineText;
          tableLines.push(lineDiv);
          logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1}, Row ${rowIndex + 1}: Created div with class tbljson-${encodedJson} and text "${lineText.substring(0,50)}..."`);
        });

        if (tableLines.length > 0) {
          const fragment = document.createDocumentFragment();
          tableLines.forEach(lineDiv => fragment.appendChild(lineDiv));
          // Replace the original table with the new divs
          tableNode.parentNode.replaceChild(fragment, tableNode);
          modified = true;
          tableCount++;
          logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1} replaced with ${tableLines.length} div elements.`);
        }
      });
      if (tableCount > 0) {
        logger.info(`[ep_docx_html_customizer] Successfully processed and replaced ${tableCount} table(s).`);
      }
    }

    if (modified) {
      logger.info(`[ep_docx_html_customizer] Converted HTML (${destFile}) was modified. Writing changes.`);
      await fsp.writeFile(destFile, dom.serialize());
    } else {
      logger.info(`[ep_docx_html_customizer] Converted HTML (${destFile}) was not modified.`);
    }

    return true; // Signal that the import was handled

  } catch (err) {
    logger.error(`[ep_docx_html_customizer] Error during document processing for ${srcFile}:`, err);
    return false; // Signal that the import failed or was not fully handled
  }
}; 