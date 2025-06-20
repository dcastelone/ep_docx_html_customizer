'use strict';

// Clipboard integration for ep_docx_html_customizer – simplified version.
// We now rely on Etherpad part ordering (our part loads before ep_hyperlinked_text)
// so we use the same pattern (jQuery .on('paste')) instead of capture-phase hacks.

const {customizeDocument} = require('../../transform_common');

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
          ace.ace_inCallStackIfNecessary('docxPaste', () => {
            const innerWin = $innerIframe[0].contentWindow;
            innerWin.document.execCommand('insertHTML', false, finalHtml);
          });
        }, 'docxPaste', true);
      });

      if (DEBUG) console.log('[docx_customizer] insertion done, event propagation stopped');
    });

    if (DEBUG) console.log('[docx_customizer] paste listener attached');
  }, 'setupDocxCustomizerPaste', true);
}; 