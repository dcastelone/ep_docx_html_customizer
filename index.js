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

const DELIMITER = '|'; // Simplified delimiter for internal text representation

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

      outerSpan.className = outerClasses.trim();
      
      const fragment = document.createDocumentFragment();
      fragment.appendChild(document.createTextNode(ZWSP));
      fragment.appendChild(outerSpan);
      fragment.appendChild(document.createTextNode(ZWSP));

      img.parentNode.replaceChild(fragment, img);
      modified = true;
      logger.debug(`[ep_docx_html_customizer] Converted HTML - Image ${index + 1} replaced with ZWSP-span-ZWSP structure.`);
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
            // Preserve inner HTML, strip outer tags of cell, replace newlines with spaces
            let cellHTML = cell ? cell.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim() : ' '; // space for empty cell
            // Remove any <p> tags that might be wrapping content within cells after LibreOffice conversion
            // and just get their innerHTML. This is a common artifact.
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = cellHTML;
            const pTags = tempDiv.querySelectorAll('p');
            if (pTags.length === 1 && tempDiv.textContent.trim() === pTags[0].textContent.trim()) {
                cellHTML = pTags[0].innerHTML.replace(/\r\n|\r|\n/g, ' ').trim();
            } else if (pTags.length > 0) {
                // If multiple p tags, join their content. This is a simple heuristic.
                cellHTML = Array.from(pTags).map(p => p.innerHTML.replace(/\r\n|\r|\n/g, ' ').trim()).join(' ');
            }
            return cellHTML || ' '; // Ensure space for truly empty cells
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