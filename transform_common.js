'use strict';

/**
 * Shared transformation utilities for ep_docx_html_customizer.
 *
 * The function exported here receives a DOM `document` (real browser DOM or jsdom) and mutates it
 * in-place, applying the same set of conversions we perform for DOCX import:
 *   • Line-break insertion around headings.
 *   • Alignment wrapper (<center>, <right>, etc.).
 *   • Ordered-list flattening.
 *   • Image span replacement compatible with ep_image_insert.
 *   • Hyperlink <a> → <span class="hyperlink-…"> conversion.
 *   • Basic colour & size mapping.
 *   • Table → tbljson-… lines (compatible with ep_tables5).
 *
 * The implementation is a trimmed copy of the logic that already existed in index.js, but stripped
 * of Node-only APIs so that it can run in the browser as well.  Where file-system access was
 * required (e.g., converting relative image src into data URIs) the behaviour now depends on the
 * options.env flag – it is executed only on the server side.
 */

// We intentionally avoid unconditional `require()` of Node modules so that this file can be
// bundled for the browser without pulling in heavy stubs.
let fs, path, mime;
if (typeof window === 'undefined') {
  const _req = eval('require'); // avoid static analysis by bundlers
  fs = _req('fs');
  path = _req('path');
  mime = _req('mime');
}

const logger = (() => {
  // Use log4js in Node, console in browser.
  if (typeof window === 'undefined') {
    try {
      const _req = eval('require');
      return _req('log4js').getLogger('ep_docx_html_customizer');
    } catch (_) { /* fall through */ }
  }
  return {
    debug: console.debug.bind(console),
    info: console.info.bind(console),
    warn: console.warn.bind(console),
    error: console.error.bind(console),
  };
})();

// Helper for stable random ids (client and server)
const rand = () => Math.random().toString(36).slice(2, 8);

// Base64 encode helper (URL-safe) reused by table logic
const enc = (s) => {
  if (typeof btoa === 'function') {
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
  } else if (typeof Buffer === 'function') {
    return Buffer.from(s).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  }
  return s;
};

// Same delimiter used by ep_tables5
const DELIMITER = '\u241F';
const ZWSP = '\u200B';

/**
 * Upload an image to S3 using the existing ep_images_extended endpoint (browser environment)
 * @param {Blob} blob - The image blob
 * @param {string} filename - The filename to use
 * @param {string} padId - The pad ID
 * @returns {Promise<string|null>} - The public URL or null if failed
 */
async function uploadImageToS3Browser(blob, filename, padId) {
  try {
    // Check if we're in browser environment and S3 is configured
    if (typeof window === 'undefined') {
      throw new Error('Not in browser environment');
    }
    
    if (typeof clientVars === 'undefined') {
      throw new Error('clientVars not available');
    }
    
    console.log('[transform_common] clientVars.ep_images_extended:', clientVars.ep_images_extended);
    
    if (!clientVars.ep_images_extended) {
      throw new Error('ep_images_extended not configured in clientVars');
    }
    
    if (clientVars.ep_images_extended.storageType !== 's3_presigned') {
      throw new Error(`Storage type is "${clientVars.ep_images_extended.storageType}", expected "s3_presigned"`);
    }

    // Get presigned URL from ep_images_extended
    const queryParams = new URLSearchParams({ name: filename, type: blob.type });
    const basePath = window.location.pathname.split('/p/')[0] || '';
    const presignUrl = `${basePath}/p/${padId}/pluginfw/ep_images_extended/s3_presign?${queryParams}`;
    
    console.log('[transform_common] requesting presign URL:', presignUrl);
    console.log('[transform_common] parameters:', { name: filename, type: blob.type, padId });
    
    const presignResp = await fetch(presignUrl);

    if (!presignResp.ok) {
      const errorText = await presignResp.text();
      console.error('[transform_common] presign error response:', errorText);
      throw new Error(`Presign request failed: ${presignResp.status} - ${errorText}`);
    }

    const presignData = await presignResp.json();
    if (!presignData.signedUrl || !presignData.publicUrl) {
      throw new Error('Invalid presign response');
    }

    // Upload to S3
    const uploadResp = await fetch(presignData.signedUrl, {
      method: 'PUT',
      headers: { 'Content-Type': blob.type },
      body: blob,
    });

    if (!uploadResp.ok) {
      throw new Error(`S3 upload failed: ${uploadResp.status}`);
    }

    return presignData.publicUrl;
  } catch (error) {
    console.warn('[transform_common] S3 upload failed:', error);
    return null;
  }
}

function customizeDocument(document, options = {}) {
  let modified = false;

  /* ───────────── Headings ───────────── */
  const headingElements = Array.from(document.querySelectorAll('h1, h2, h3, h4, h5, h6'));
  headingElements.forEach((heading) => {
    const parent = heading.parentNode;
    const brBefore = document.createElement('br');
    parent.insertBefore(brBefore, heading);
    const brAfter = document.createElement('br');
    if (heading.nextSibling) parent.insertBefore(brAfter, heading.nextSibling);
    else parent.appendChild(brAfter);
    modified = true;
  });

  /* ───────────── Alignment ───────────── */
  const ALIGN_MAP = {center: 'center', right: 'right', justify: 'justify', left: 'left'};
  const alignedBlocks = Array.from(document.querySelectorAll('[style*="text-align"], [align]'));
  alignedBlocks.forEach((blk) => {
    const tagName = blk.tagName.toLowerCase();
    if (['center', 'left', 'right', 'justify'].includes(tagName)) return;
    if (blk.closest('table')) return; // ignore inside tables

    let alignVal = (blk.getAttribute('align') || '').toLowerCase();
    if (!alignVal) {
      const m = /text-align\s*:\s*(left|right|center|justify)/i.exec(blk.getAttribute('style') || '');
      if (m) alignVal = m[1].toLowerCase();
    }
    if (!ALIGN_MAP[alignVal] || alignVal === 'left') return;

    const wrapper = document.createElement(ALIGN_MAP[alignVal]);
    const parent = blk.parentNode;
    const isHeading = /^h[1-6]$/i.test(blk.tagName);
    if (isHeading) {
      parent.replaceChild(wrapper, blk);
      wrapper.appendChild(blk);
      blk.removeAttribute('align');
      blk.style.removeProperty && blk.style.removeProperty('text-align');
    } else {
      wrapper.innerHTML = blk.innerHTML;
      parent.replaceChild(wrapper, blk);
    }
    const br = document.createElement('br');
    if (wrapper.nextSibling) parent.insertBefore(br, wrapper.nextSibling);
    else parent.appendChild(br);
    modified = true;
  });

  /* ───────────── Ordered-list flattening ───────────── */
  const processOrderedList = (olNode, depth = 0) => {
    const startAttr = parseInt(olNode.getAttribute('start') || '1', 10);
    let counter = isNaN(startAttr) ? 1 : startAttr;
    const frag = document.createDocumentFragment();
    Array.from(olNode.children).forEach((child) => {
      if (child.tagName && child.tagName.toLowerCase() === 'li') {
        const prefixSpan = document.createElement('span');
        prefixSpan.textContent = `${counter}. `;
        const p = document.createElement('div');
        const liClone = child.cloneNode(true);
        const nestedOLs = liClone.querySelectorAll('ol');
        nestedOLs.forEach((nested) => {
          const replacement = processOrderedList(nested, depth + 1);
          nested.parentNode.replaceChild(replacement, nested);
        });
        if (liClone.childNodes.length === 1 && liClone.firstChild.tagName && liClone.firstChild.tagName.toLowerCase() === 'p') {
          const innerP = liClone.firstChild;
          while (innerP.firstChild) p.appendChild(innerP.firstChild);
        } else {
          while (liClone.firstChild) p.appendChild(liClone.firstChild);
        }
        p.insertBefore(prefixSpan, p.firstChild);
        if (depth > 0) p.style.marginLeft = `${depth * 1.5}em`;
        frag.appendChild(p);
        counter += 1;
      }
    });
    modified = true;
    return frag;
  };
  Array.from(document.querySelectorAll('ol')).forEach((ol) => {
    const replacement = processOrderedList(ol);
    ol.parentNode.replaceChild(replacement, ol);
  });

  /* ───────────── Images ───────────── */
  const cssImageSizeMap = {};
  try {
    document.querySelectorAll('style').forEach((styleNode) => {
      const cssText = styleNode.textContent || '';
      cssText.split('}').forEach((rule) => {
        const m = rule.match(/\.([^,{\s]+)\s*\{([^}]*)/); // .classname { ... }
        if (!m) return;
        const cls = m[1].trim();
        const decls = m[2] || '';
        const wMatch = /width\s*:\s*([^;]+)/i.exec(decls);
        const hMatch = /height\s*:\s*([^;]+)/i.exec(decls);
        if (wMatch || hMatch) {
          cssImageSizeMap[cls] = {
            w: wMatch ? wMatch[1].trim() : null,
            h: hMatch ? hMatch[1].trim() : null,
          };
        }
      });
    });
  } catch (e) { /* swallow but log */ console.warn('[transform_common] CSS parse error', e); }

  if (typeof console !== 'undefined' && console.debug) {
    console.debug('[transform_common] cssImageSizeMap keys', Object.keys(cssImageSizeMap).length);
  }

  const images = document.querySelectorAll('img');
  images.forEach((img) => {
    let src = img.getAttribute('src');
    if (!src) return;

    if (typeof window === 'undefined' && fs && !src.startsWith('http') && !src.startsWith('data:') && !src.startsWith('/')) {
      // Node env: try to inline relative image as data URI
      try {
        const imagePath = path.resolve(options.destDir || process.cwd(), src);
        if (fs.existsSync(imagePath)) {
          const buffer = fs.readFileSync(imagePath);
          const mimeType = mime.getType(imagePath) || 'application/octet-stream';
          src = `data:${mimeType};base64,${buffer.toString('base64')}`;
        }
      } catch (e) {
        logger.warn('Image inlining failed', e);
      }
    }

    const outerSpan = document.createElement('span');
    outerSpan.textContent = ZWSP;
    let classes = 'inline-image character image-placeholder';
    classes += ` image:${encodeURIComponent(src)}`;

    // Preserve explicit width/height from the source HTML so that ep_images_extended can
    // recreate the original size immediately after the paste (even before remote inlining
    // occurs).  This mirrors the logic already present in the server-side import path
    // (index.js) so both clipboard and document import behave identically.

    let imgWidthAttr  = img.getAttribute('width')  || (img.style && img.style.width)  || '';
    let imgHeightAttr = img.getAttribute('height') || (img.style && img.style.height) || '';

    if (typeof console !== 'undefined' && console.debug) {
      console.debug('[transform_common] <img> initial dims', {src, width: imgWidthAttr, height: imgHeightAttr, classList: img.className});
    }

    // Fallback to CSS rule–declared sizes if direct attrs missing
    if (!imgWidthAttr || !imgHeightAttr) {
      const classesImg = (img.className || '').split(/\s+/).filter(Boolean);
      for (const c of classesImg) {
        const entry = cssImageSizeMap[c];
        if (entry) {
          if (!imgWidthAttr && entry.w)  imgWidthAttr  = entry.w;
          if (!imgHeightAttr && entry.h) imgHeightAttr = entry.h;
          if (typeof console !== 'undefined' && console.debug) {
            console.debug('[transform_common] CSS fallback found', {cls: c, w: entry.w, h: entry.h});
          }
          if (imgWidthAttr && imgHeightAttr) break;
        }
      }
    }

    // Helper: returns a pixel value string with "px" suffix when the input is a plain number.
    const _normaliseDim = (val) => {
      if (!val) return null;
      const trimmed = String(val).trim();
      // Accept formats like "123", "123px", "123.4", "123.4px" (case-insensitive)
      const m = /^([0-9]+(?:\.[0-9]+)?)(px)?$/i.exec(trimmed);
      if (m) return `${m[1]}px`;
      return trimmed; // keep original (e.g. "%", "pt", etc.)
    };

    const imgWidthPx  = _normaliseDim(imgWidthAttr);
    const imgHeightPx = _normaliseDim(imgHeightAttr);

    if (imgWidthPx)  classes += ` image-width:${imgWidthPx}`;
    if (imgHeightPx) classes += ` image-height:${imgHeightPx}`;

    // Add CSS aspect ratio helper if both dimensions are numeric.
    const numW = parseFloat(imgWidthPx);
    const numH = parseFloat(imgHeightPx);
    if (!isNaN(numW) && numW > 0 && !isNaN(numH) && numH > 0) {
      classes += ` imageCssAspectRatio:${(numW / numH).toFixed(4)}`;
    }

    const imageId = rand() + rand();
    classes += ` image-id-${imageId}`;
    outerSpan.setAttribute('data-image-id', imageId);
    outerSpan.className = classes;
    const frag = document.createDocumentFragment();
    frag.appendChild(document.createTextNode(ZWSP));
    frag.appendChild(outerSpan);
    frag.appendChild(document.createTextNode(ZWSP));
    img.parentNode.replaceChild(frag, img);
    modified = true;

    if (typeof console !== 'undefined' && console.debug) {
      console.debug('[transform_common] <img> final dims to encode', {width: imgWidthAttr, height: imgHeightAttr});
    }
  });

  /* ───────────── Hyperlinks ───────────── */
  const anchors = document.querySelectorAll('a[href]');
  anchors.forEach((a) => {
    let href = (a.getAttribute('href') || '').trim();
    if (!href) return;
    if (!/^(https?:\/\/|mailto:|ftp:|file:|#|\/)/i.test(href)) href = `http://${href}`;
    const encodedHref = encodeURIComponent(href);
    const span = document.createElement('span');
    span.className = `hyperlink hyperlink-${encodedHref}`;
    span.textContent = a.textContent || href;
    const frag = document.createDocumentFragment();
    frag.appendChild(document.createTextNode(ZWSP));
    frag.appendChild(span);
    frag.appendChild(document.createTextNode(ZWSP));
    a.parentNode.replaceChild(frag, a);
    modified = true;
  });

  /* ───────────── Font color & size ───────────── */
  const parseColorToRgb = (str) => {
    if (!str) return null;
    const s = str.toLowerCase().trim();
    const named = {black:[0,0,0], red:[255,0,0], green:[0,128,0], blue:[0,0,255], yellow:[255,255,0], orange:[255,165,0]};
    if (named[s]) return named[s];
    const hex = s.match(/^#?([0-9a-f]{3}|[0-9a-f]{6})$/i);
    if (hex) {
      let h = hex[1];
      if (h.length===3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
      return [parseInt(h.slice(0,2),16),parseInt(h.slice(2,4),16),parseInt(h.slice(4,6),16)];
    }
    const rgb = s.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
    if (rgb) return [parseInt(rgb[1]),parseInt(rgb[2]),parseInt(rgb[3])];
    return null;
  };

  const findNearestColor = (rgb) => {
    if (!rgb) return null;
    const palette = {black:[0,0,0], red:[255,0,0], green:[0,128,0], blue:[0,0,255], yellow:[255,255,0], orange:[255,165,0]};
    let best='black',bestDist=1e9;
    for(const [name,p] of Object.entries(palette)){
      const d=Math.sqrt((rgb[0]-p[0])**2+(rgb[1]-p[1])**2+(rgb[2]-p[2])**2);
      if(d<bestDist){bestDist=d;best=name;}
    }
    return best;
  };

  const parsePx = (val)=>{
    if(!val) return null;
    const m=/([0-9.]+)(px|pt)?/i.exec(val);
    if(!m) return null;
    let num=parseFloat(m[1]);
    if(isNaN(num)) return null;
    if(m[2]&&m[2].toLowerCase()==='pt') num=Math.round(num*1.333);
    return Math.round(num);
  };

  const styledNodes = Array.from(document.querySelectorAll('[style*="color"], font[color], [style*="font-size"]'));
  styledNodes.forEach((el)=>{
    const classes=(el.className||'').split(/\s+/).filter(Boolean);
    let changed=false;
    const colorAttr=el.getAttribute('color')||el.style.color;
    if(colorAttr){
      const rgb=parseColorToRgb(colorAttr);
      const mapped=findNearestColor(rgb);
      if(mapped && mapped!=='black'){
        const cls=`color:${mapped}`;
        if(!classes.includes(cls)){classes.push(cls);changed=true;}
      }
    }
    // Skip font-size mapping for <sup>/<sub> to preserve normal baseline sizing
    const isSuperOrSub = el.tagName && (el.tagName.toUpperCase()==='SUP' || el.tagName.toUpperCase()==='SUB');
    const sizeStyle=isSuperOrSub ? '' : el.style.fontSize;
    if(sizeStyle){
      const px=parsePx(sizeStyle);
      if(px){
        const palette=[8,9,10,11,12,14,16,18,20,22,24,26,28,30,35,40];
        let best=palette[0],bestDiff=Math.abs(px-best);
        for(const s of palette){ const d=Math.abs(px-s); if(d<bestDiff){best=s;bestDiff=d;} }
        if(best!==14){
          const cls=`font-size:${best}`;
          if(!classes.includes(cls)){classes.push(cls);changed=true;}
        }
      }
    }
    if(changed){
      el.className=classes.join(' ');
      el.removeAttribute('color');
      el.style.color='';
      el.style.fontSize='';
      modified=true;
    }
  });

  /* ───────────── Table → tbljson conversion ───────────── */
  const tables = document.querySelectorAll('table');
  if (tables.length) {
    const tblIdBase = rand();
    tables.forEach((tableNode, tableIdx) => {
      const rows = tableNode.querySelectorAll('tr');
      if (!rows.length) return;

      // helper to get logical column count taking colspan into account
      const getLogicalCols = (tr) => Array.from(tr.querySelectorAll('td, th')).reduce((cnt, c) => {
        const span = parseInt(c.getAttribute('colspan') || '1', 10);
        return cnt + (isNaN(span) || span < 1 ? 1 : span);
      }, 0);

      const numCols = Array.from(rows).reduce((max, tr) => {
        const cols = getLogicalCols(tr);
        return cols > max ? cols : max;
      }, 0);
      if (!numCols) return;

      const pendingRowspan = Array(numCols).fill(0);
      const newLines = [];

      rows.forEach((rowNode, rowIdx) => {
        const rawCells = Array.from(rowNode.querySelectorAll('td, th'));
        let rawPtr = 0;
        const cellHTMLs = [];

        for (let col = 0; col < numCols; col++) {
          if (pendingRowspan[col] > 0) {
            pendingRowspan[col]--;
            cellHTMLs.push('<span>&nbsp;</span>');
            continue;
          }

          const cell = rawCells[rawPtr++];
          if (!cell) {
            cellHTMLs.push('<span>&nbsp;</span>');
            continue;
          }

          let colspan = parseInt(cell.getAttribute('colspan') || '1', 10);
          let rowspan = parseInt(cell.getAttribute('rowspan') || '1', 10);
          if (isNaN(colspan) || colspan < 1) colspan = 1;
          if (isNaN(rowspan) || rowspan < 1) rowspan = 1;

          let html = cell.innerHTML.replace(/\r?\n/g, ' ').trim();
          const tmp = document.createElement('div');
          tmp.innerHTML = html;
          const heading = tmp.querySelector('h1, h2, h3, h4, h5, h6');
          if (heading) {
            const cls = (heading.className || '').split(/\s+/).filter(Boolean);
            if (!cls.includes('bold')) cls.push('bold');
            const span = document.createElement('span');
            span.className = cls.join(' ');
            span.innerHTML = heading.innerHTML.replace(/\r?\n/g, ' ').trim();
            heading.parentNode.replaceChild(span, heading);
          }

          const flatten = (div) => {
            const parts = [];
            div.childNodes.forEach((n) => {
              if (n.nodeType === 3) {
                const t = n.textContent.replace(/\r?\n/g, ' ').trim(); if (t) parts.push(t);
              } else if (n.tagName && n.tagName.toLowerCase() === 'p') {
                const inner = n.innerHTML.replace(/\r?\n/g, ' ').trim(); if (inner) parts.push(inner);
              } else if (n.outerHTML) { parts.push(n.outerHTML.replace(/\r?\n/g, ' ').trim()); }
            });
            return parts.join(' ').trim();
          };

          html = flatten(tmp);
          if (/^<br\/?>(\s*)?$/i.test(html)) html = '';
          if (html) html = html.replace(/(<br\s*\/?>\s*)+$/gi, '').replace(/<br\s*\/?>/gi, ' ').trim();

          let visiblyEmpty = false;
          if (!html) {
            visiblyEmpty = true;
          } else {
            const probeDiv = document.createElement('div');
            probeDiv.innerHTML = html;
            const txt = (probeDiv.textContent || '').replace(/\u00A0/g, '').trim();
            visiblyEmpty = txt === '';
          }

          if (visiblyEmpty) html = '<span>&nbsp;</span>';

          cellHTMLs.push(html);

          if (rowspan > 1) pendingRowspan[col] = rowspan - 1;
          for (let extra = 1; extra < colspan && col + extra < numCols; extra++) {
            cellHTMLs.push('<span>&nbsp;</span>');
            if (rowspan > 1) pendingRowspan[col + extra] = rowspan - 1;
          }
          col += colspan - 1;
        }

        const meta = {tblId: `${tblIdBase}-${tableIdx}`, row: rowIdx, cols: numCols};
        const encodedMeta = enc(JSON.stringify(meta));
        const lineDiv = document.createElement('div');
        
        cellHTMLs.forEach((cellHTML, cellIdx) => {
            const span = document.createElement('span');
            span.className = `tbljson-${encodedMeta} tblCell-${cellIdx}`;
            span.innerHTML = cellHTML;
            lineDiv.appendChild(span);
            if (cellIdx < cellHTMLs.length - 1) {
                lineDiv.appendChild(document.createTextNode(DELIMITER));
            }
        });
        
        // Add the final, empty span that ep_tables5 expects to exist at the end
        // of the line. This span will hold line-level attributes and eventually the caret anchor.
        const finalSpan = document.createElement('span');
        finalSpan.className = `tbljson-${encodedMeta}`;
        lineDiv.appendChild(finalSpan);

        newLines.push(lineDiv);
      });

      if (newLines.length) {
        const frag = document.createDocumentFragment();
        newLines.forEach(d=>frag.appendChild(d));
        tableNode.parentNode.replaceChild(frag, tableNode);
        modified = true;
      }
    });
  }

  /* ───────────── Superscript / Subscript (vertical-align) ───────────── */
  const vaElements = Array.from(document.querySelectorAll('[style*="vertical-align" ]'));
  vaElements.forEach((el) => {
    const vAlign = (el.style && el.style.verticalAlign || '').toLowerCase();
    if (vAlign !== 'super' && vAlign !== 'sub') return;

    const tagNameTarget = vAlign === 'super' ? 'sup' : 'sub';
    const newTag = document.createElement(tagNameTarget);

    // Preserve existing classes (excluding sup/sub to avoid duplication later)
    newTag.className = (el.className || '').replace(/\b(sup|sub)\b/g, '').trim();
    newTag.innerHTML = el.innerHTML;

    el.parentNode.replaceChild(newTag, el);
    modified = true;
  });

  return modified;
}

module.exports = {customizeDocument, uploadImageToS3Browser, DELIMITER, ZWSP}; 