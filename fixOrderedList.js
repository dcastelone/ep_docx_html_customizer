// Server-side hook to reset ordered-list numbering between separate <ol> blocks
// This works around Etherpad core behaviour where `state.start` is never
// cleared once an ordered list ends, causing subsequent lists to continue the
// previous numbering (eg. 11,12,13 …).  By deleting `state.start` before the
// ContentCollector enters a new <ol> element we ensure that the counter starts
// from 1 again for every distinct ordered list.
//
// Place this file alongside index.js and register it via ep.json with the
// `collectContentPre` server hook.

exports.collectContentPre = (hookName, context) => {
  const {state, tname} = context;

  // We only want to act on ordered lists (<ol>).  Unordered lists (<ul>) are
  // unaffected because they do not rely on the `start` counter.
  if (tname !== 'ol') return;

  // If we are about to enter a fresh top-level ordered list (not nested inside
  // another list) we nuke the persistent counter so that ContentCollector will
  // rebuild it from 0 → 1 for this list only.
  // `state.listNesting` is undefined or 0 when we are not already inside a list.
  if (!state.listNesting) {
    delete state.start; // let _enterList() initialise it to 1
  }
};

// NEW: advance the counter once the current <li> has been collected
exports.collectContentPost = (hookName, context) => {
  const {state, tname} = context;
  if (tname !== 'li') return;                                   // only list items
  const listType = state.lineAttributes && state.lineAttributes.list;
  if (!listType || listType.indexOf('number') === -1) return;   // only OLs
  // Skip soft-break (no-number) items
  if (context.cls && context.cls.indexOf('listbreak') !== -1) return;
  state.start = (state.start || 1) + 1;                         // next item number
}; 