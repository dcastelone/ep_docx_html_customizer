// Client-side hooks to support Shift+Enter soft-breaks inside ordered lists
// – aceKeyEvent intercepts Shift+Enter, inserts a newline, and tags the
//   newly created blank line with the custom attribute `listbreak:true`.
// – aceAttribsToClasses maps that attribute to the CSS class `listbreak` so
//   we can hide its list marker.
// – aceEditorCSS injects a tiny style-sheet that hides the marker.

/* global $, _, window */

// 1. Intercept Shift+Enter
exports.aceKeyEvent = (hookName, ctx) => {
  const {evt, rep, editorInfo, documentAttributeManager: docAttrMgr} = ctx;
  if (!(evt.type === 'keydown' && evt.key === 'Enter' && evt.shiftKey)) return false;
  console.debug('[listSoftBreak] Shift+Enter pressed');

  // Only when the caret is inside a numbered list line
  const lineNum = rep.selStart[0];
  const listAttr = docAttrMgr.getAttributeOnLine(lineNum, 'list');
  console.debug('[listSoftBreak] caret line', lineNum, 'listAttr', listAttr);
  if (!listAttr || listAttr.indexOf('number') === -1) return false;

  // Prevent Etherpad's default Return handling
  evt.preventDefault();
  console.debug('[listSoftBreak] inserting soft break');

  // Replace the current selection with a line-feed to create a new blank line
  editorInfo.ace_performDocumentReplaceRange(rep.selStart, rep.selEnd, '\n');

  // After the new line exists, mark it with listbreak:true so CSS can hide its marker
  // We must schedule this for the next tick because the line isn't created until
  // the replaceRange has been processed.
  setTimeout(() => {
    const newLine = lineNum + 1;
    const repNow = editorInfo.ace_getRep();
    const lineLen = repNow.lines.atIndex(newLine).text.length;
    const endCol = Math.min(1, lineLen); // if empty line len=0, use 0
    try {
      editorInfo.ace_performDocumentApplyAttributesToRange(
        [newLine, 0], [newLine, endCol], [['listbreak', 'true']],
      );
      console.debug('[listSoftBreak] applied listbreak attr on line', newLine);
    } catch (e) {
      console.warn('[listSoftBreak] Could not apply attr:', e);
    }
  }, 0);

  return true; // We handled the key event
};

// 2. Map attribute → CSS class so the frontend can style it.
exports.aceAttribsToClasses = (hookName, ctx) => {
  if (ctx.key === 'listbreak' && ctx.value === 'true') return ['listbreak'];
  return [];
};

// 3. Provide a tiny style-sheet that hides the list marker for those lines
exports.aceEditorCSS = () => [
  'ep_docx_html_customizer/static/css/list-softbreak.css',
];

console.debug('[ep_docx_html_customizer:listSoftBreak] Loaded client hooks'); 