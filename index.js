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
    let actualConvertedTmpFile = path.join(outDir, tempConvertedBaseName);

    const conversionCommand = `"${converterPath}" --headless --invisible --nologo --nolockcheck --writer --convert-to html "${srcFile}" --outdir "${outDir}"`;
    logger.debug(`[ep_docx_html_customizer] Executing soffice command: ${conversionCommand}`);

    await execPromise(conversionCommand);
    logger.info(`[ep_docx_html_customizer] LibreOffice conversion successful. HTML output at: ${actualConvertedTmpFile}`);

    // ─── Fallback lookup ────────────────────────────────────────────────────────
    if (!fs.existsSync(actualConvertedTmpFile)) {
      /*
       * Some temporary upload names no longer end with the real extension
       * (e.g. `/tmp/abcd.odt.5`). LibreOffice will therefore drop the extra
       * suffix and produce `/tmp/abcd.html`, not `/tmp/abcd.5.html`.
       * If our first guess is missing, scan the outDir for any fresh *.html
       * file produced by this conversion run and pick the most recent one.
       */
      try {
        const htmlFiles = fs.readdirSync(outDir)
          .filter(f => f.toLowerCase().endsWith('.html'))
          .map(f => path.join(outDir, f))
          // sort newest first by mtime
          .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);

        if (htmlFiles.length > 0) {
          const candidate = htmlFiles[0];
          logger.warn(`[ep_docx_html_customizer] Expected converted file not found; using fallback candidate ${candidate}`);
          actualConvertedTmpFile = candidate;
        }
      } catch (e) {
        logger.error('[ep_docx_html_customizer] Error while searching for fallback HTML file:', e);
      }

      if (!fs.existsSync(actualConvertedTmpFile)) {
        logger.error(`[ep_docx_html_customizer] Conversion failed: could not locate any converted HTML file (looked for ${actualConvertedTmpFile}).`);
        return false;
      }
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
      // Ensure line breaks before and after each heading to prevent content merging
      const parent = heading.parentNode;
      
      // Add line break before the heading
      const brBefore = document.createElement('br');
      parent.insertBefore(brBefore, heading);
      
      // Add line break after the heading
      const brAfter = document.createElement('br');
      if (heading.nextSibling) {
        parent.insertBefore(brAfter, heading.nextSibling);
      } else {
        parent.appendChild(brAfter);
      }
      
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Added line breaks before and after ${heading.tagName.toLowerCase()} element ${idx + 1}.`);
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
      const parent = blk.parentNode;

      const isHeading = /^h[1-6]$/i.test(blk.tagName);

      if (isHeading) {
        // Keep heading tag intact by nesting it
        parent.replaceChild(newEl, blk);
        newEl.appendChild(blk);

        // Remove redundant alignment from heading itself
        blk.removeAttribute('align');
        if (blk.style) blk.style.removeProperty('text-align');
      } else {
        // For regular paragraphs or other blocks, drop the original tag and just keep content
        newEl.innerHTML = blk.innerHTML;
        parent.replaceChild(newEl, blk);
      }
      
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Wrapped element ${idx + 1} with <${wrapperTag}> and added line break.`);

      // Add explicit line break after the aligned wrapper to keep lines separate
      const br = document.createElement('br');
      if (newEl.nextSibling) {
        parent.insertBefore(br, newEl.nextSibling);
      } else {
        parent.appendChild(br);
      }
    });

    /* ─────────────────────────── Ordered list flattening ─────────────────────────── */
    logger.info(`[ep_docx_html_customizer] Converting ordered lists (<ol>) to plain numbered paragraphs.`);

    const processOrderedList = (olNode, depth = 0) => {
      const startAttr = parseInt(olNode.getAttribute('start') || '1', 10);
      let counter = isNaN(startAttr) ? 1 : startAttr;

      // Collect new paragraph nodes
      const frag = document.createDocumentFragment();

      Array.from(olNode.children).forEach((child) => {
        if (child.tagName && child.tagName.toLowerCase() === 'li') {
          // Build prefix (e.g., "1. ")
          const prefixSpan = document.createElement('span');
          prefixSpan.textContent = `${counter}. `;

          // Create wrapper paragraph <div>
          const p = document.createElement('div');
          // Preserve nested content inside <li>
          const liClone = child.cloneNode(true);

          // Remove any nested ordered lists before cloning (they will be handled recursively)
          const nestedOLs = liClone.querySelectorAll('ol');
          nestedOLs.forEach((nested) => {
            const replacement = processOrderedList(nested, depth + 1);
            nested.parentNode.replaceChild(replacement, nested);
          });

          // Move all children from liClone into wrapper, but unwrap a single wrapping <p>
          if (liClone.childNodes.length === 1 && liClone.firstChild.tagName && liClone.firstChild.tagName.toLowerCase() === 'p') {
            const innerP = liClone.firstChild;
            while (innerP.firstChild) {
              p.appendChild(innerP.firstChild);
            }
          } else {
            while (liClone.firstChild) {
              p.appendChild(liClone.firstChild);
            }
          }

          // Prepend number prefix
          p.insertBefore(prefixSpan, p.firstChild);

          // Optionally indent child paragraphs depending on depth
          if (depth > 0) p.style.marginLeft = `${depth * 1.5}em`;

          frag.appendChild(p);

          // No explicit <br>; <div> itself creates a new line in Etherpad

          counter += 1;
        }
      });

      modified = true;
      return frag;
    };

    const allOrderedLists = document.querySelectorAll('ol');
    logger.debug(`[ep_docx_html_customizer] Found ${allOrderedLists.length} ordered list(s) to flatten.`);

    Array.from(allOrderedLists).forEach((ol) => {
      const replacementFrag = processOrderedList(ol);
      ol.parentNode.replaceChild(replacementFrag, ol);
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

        // Determine logical column count: honour colspan and consider every row, then take the max.
        const getLogicalColCount = (tr) => {
          return Array.from(tr.querySelectorAll('td, th')).reduce((count, cell) => {
            const span = parseInt(cell.getAttribute('colspan') || '1', 10);
            return count + (isNaN(span) || span < 1 ? 1 : span);
          }, 0);
        };

        const numCols = Array.from(rows).reduce((max, tr) => {
          const cols = getLogicalColCount(tr);
          return cols > max ? cols : max;
        }, 0);

        if (numCols === 0) {
            logger.debug(`[ep_docx_html_customizer] Table ${tableIndex + 1} appears to have zero columns after analysis, skipping.`);
            return;
        }

        const tableLines = [];

        // Track vertical merges that must propagate blank cells downward
        const pendingRowspan = Array(numCols).fill(0);

        rows.forEach((rowNode, rowIndex) => {
          const rawCells = Array.from(rowNode.querySelectorAll('td, th'));
          let rawPtr = 0;
          const cellContents = new Array(numCols);

          for (let col = 0; col < numCols; col++) {
            // 1) Handle continuation of a rowspan merge
            if (pendingRowspan[col] > 0) {
              pendingRowspan[col]--;
              cellContents[col] = ' <span>&nbsp;</span>';
              continue;
            }

            const cell = rawCells[rawPtr++];
            if (!cell) {
              cellContents[col] = ' <span>&nbsp;</span>';
              continue;
            }

            // ───── Extract colspan/rowspan safely ─────
            let colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
            let rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
            if (isNaN(colspan) || colspan < 1) colspan = 1;
            if (isNaN(rowspan) || rowspan < 1) rowspan = 1;

            // ───── Normalise inner HTML (logic from previous version) ─────
            let cellHTML = cell.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim();

            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = cellHTML;
            const headingEl = tempDiv.querySelector('h1, h2, h3, h4, h5, h6');
            if (headingEl) {
              // Collect existing classes that colour / size mapping added earlier.
              const existingClasses = (headingEl.className || '').split(/\s+/).filter(Boolean);
              if (!existingClasses.includes('bold')) existingClasses.push('bold');

              // Build a <span> with the combined classes and the heading's inner HTML.
              const span = document.createElement('span');
              span.className = existingClasses.join(' ');
              span.innerHTML = headingEl.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim();

              // Replace the heading element inside the tempDiv.
              headingEl.parentNode.replaceChild(span, headingEl);
            }

            /* ── Paragraph flattening ─────────────────────────────────────────── */
            const flattenParagraphs = (div) => {
              const parts = [];
              div.childNodes.forEach((node) => {
                if (node.nodeType === 3) { // text
                  const txt = node.textContent.replace(/\r\n|\r|\n/g, ' ').trim();
                  if (txt) parts.push(txt);
                } else if (node.tagName && node.tagName.toLowerCase() === 'p') {
                  const inner = node.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim();
                  if (inner) parts.push(inner);
                } else if (node.outerHTML) {
                  parts.push(node.outerHTML.replace(/\r\n|\r|\n/g, ' ').trim());
                }
              });
              return parts.join(' ').trim();
            };

            cellHTML = flattenParagraphs(tempDiv);

            if (cellHTML === '<br>' || cellHTML === '<br/>' || cellHTML === '<br />') {
              cellHTML = '';
            }

            if (cellHTML) {
              cellHTML = cellHTML.replace(/(<br\s*\/?>\s*)+$/gi, '').trim();
              if (cellHTML.includes('<br')) {
                cellHTML = cellHTML.replace(/<br\s*\/?>/gi, ' ');
              }
            }

            let isVisiblyEmpty = false;
            if (!cellHTML) {
              isVisiblyEmpty = true;
            } else {
              const probeDiv = document.createElement('div');
              probeDiv.innerHTML = cellHTML;
              const text = (probeDiv.textContent || '').replace(/\u00A0/g, '').trim();
              isVisiblyEmpty = text === '';
            }

            if (isVisiblyEmpty) {
              cellHTML = ' <span>&nbsp;</span>';
            }

            // Place main cell content
            cellContents[col] = cellHTML;

            // Register vertical merge
            if (rowspan > 1) {
              pendingRowspan[col] = rowspan - 1;
            }

            // Handle colspan by filling following columns with blanks and also tracking rowspan
            for (let extra = 1; extra < colspan && col + extra < numCols; extra++) {
              cellContents[col + extra] = ' <span>&nbsp;</span>';
              if (rowspan > 1) pendingRowspan[col + extra] = rowspan - 1;
            }

            // Skip over columns we just filled via colspan
            col += colspan - 1;
          }
           
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

// ============================================================================
// expressCreateServer – install same-origin image proxy to bypass CORS
// ============================================================================

let _fetchImpl; // lazy-loaded fetch replacement when global fetch is absent
const _getFetch = () => {
  if (typeof global.fetch === 'function') return global.fetch;
  if (!_fetchImpl) {
    try {
      // Dynamically import node-fetch (v3 is ESM-only)
      _fetchImpl = (...args) => import('node-fetch').then(m => (m.default || m)(...args));
    } catch (e) {
      logger.warn('[ep_docx_html_customizer] node-fetch is not available and global fetch is missing.');
    }
  }
  return _fetchImpl;
};

/**
 * expressCreateServer hook – adds /ep_docx_image_proxy?url=…
 * This endpoint streams remote images back to the browser with permissive
 * CORS headers so that client-side code can convert them into data URIs.
 */
exports.expressCreateServer = (hookName, {app}) => {
  console.log('[docx_customizer] expressCreateServer – setting up /ep_docx_image_proxy');
  logger.info('[ep_docx_html_customizer] expressCreateServer hook: registering /ep_docx_image_proxy route');
  const FETCH_TIMEOUT_MS = 10000;
  const RATE_LIMIT_WINDOW_MS = 60 * 1000;          // 1 minute sliding window
  const RATE_LIMIT_MAX_REQUESTS = 30;               // max requests per IP within the window
  const MAX_CONTENT_LENGTH = 10 * 1024 * 1024;      // 10 MB hard size limit
  // Simple in-memory store {ip: [timestamp,…]}. Good enough for single-node deployments.
  const _rateLimitStore = new Map();

  // Helper to test whether a host looks like a private / loopback address we should refuse.
  const _isForbiddenHost = (host) => {
    if (!host) return true;
    const lc = host.toLowerCase();
    // block localhost & obvious loopback labels
    if (lc === 'localhost' || lc === '::1') return true;
    // IPv4 private / link-local ranges & loopback
    return /^(127\.|10\.|0\.|169\.254\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.)/.test(lc);
  };

  app.get('/ep_docx_image_proxy', async (req, res) => {
    // Require Etherpad session (authenticated user or guest author).
    if (!req.session || (!req.session.user && !req.session.authorId)) {
      res.status(401).send('Authentication required');
      return;
    }
    const url = req.query.url;
    if (!url || !/^https?:\/\//i.test(url)) {
      res.status(400).send('Bad url');
      return;
    }

    // Block requests targeting private or loopback hosts to mitigate SSRF.
    try {
      const {hostname} = new URL(url);
      if (_isForbiddenHost(hostname)) {
        res.status(400).send('Forbidden host');
        return;
      }
    } catch (_) {
      res.status(400).send('Malformed url');
      return;
    }

    // Primitive per-IP rate limiter (in-memory).
    const ip = req.ip || req.headers['x-forwarded-for'] || req.connection?.remoteAddress || 'unknown';
    const now = Date.now();
    let timestamps = _rateLimitStore.get(ip) || [];
    timestamps = timestamps.filter((t) => t > now - RATE_LIMIT_WINDOW_MS);
    if (timestamps.length >= RATE_LIMIT_MAX_REQUESTS) {
      res.status(429).send('Too many requests, slow down');
      return;
    }
    timestamps.push(now);
    _rateLimitStore.set(ip, timestamps);

    const fetch = _getFetch();
    if (!fetch) {
      res.status(500).send('fetch unavailable');
      return;
    }

    try {
      let controller;
      let timer;
      if (typeof AbortController === 'function') {
        controller = new AbortController();
        timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
      }
      const resp = await fetch(url, controller ? {signal: controller.signal} : {});
      if (timer) clearTimeout(timer);

      if (!resp.ok) {
        res.status(resp.status).send('Upstream error');
        return;
      }

      // Reject very large payloads to avoid memory/bandwidth abuse.
      const respLen = parseInt(resp.headers.get('content-length') || '0', 10);
      if (respLen && respLen > MAX_CONTENT_LENGTH) {
        res.status(413).send('Payload too large');
        return;
      }

      res.set({
        'Content-Type': resp.headers.get('content-type') || 'application/octet-stream',
        'Cache-Control': 'public, max-age=86400',
        'Access-Control-Allow-Origin': '*',
      });

      if (resp.body && typeof resp.body.pipe === 'function') {
        resp.body.pipe(res);
      } else {
        const buf = Buffer.from(await resp.arrayBuffer());
        res.end(buf);
      }
    } catch (e) {
      logger.warn('[ep_docx_html_customizer] Proxy error', e);
      res.status(e.name === 'AbortError' ? 504 : 502).send('Proxy error');
    }
  });
}; 