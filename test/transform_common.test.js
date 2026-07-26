'use strict';

const assert = require('node:assert/strict');
const Module = require('node:module');
const test = require('node:test');
const {JSDOM} = require('jsdom');

const originalLoad = Module._load;
Module._load = function(request, parent, isMain) {
  if (request === 'log4js') return {getLogger: () => console};
  if (request === 'mime') return {getType: () => 'application/octet-stream'};
  return originalLoad.call(this, request, parent, isMain);
};
const {customizeDocument, DELIMITER, ZWSP} = require('../transform_common');
const {collectContentPre} = require('../static/js/clipboard');
Module._load = originalLoad;

const transform = (html) => {
  const dom = new JSDOM(`<body>${html}</body>`);
  const modified = customizeDocument(dom.window.document);
  return {document: dom.window.document, html: dom.window.document.body.innerHTML, modified};
};

test('one pass transforms a mixed image, link, color, and table document', () => {
  const result = transform([
    '<p><a href="https://example.com/a?x=1">Example</a></p>',
    '<p><img src="https://cdn.example/image.png" width="320" height="160"></p>',
    '<p><span style="color: rgb(51, 102, 153); font-size: 16px">Styled</span></p>',
    '<table><tr><td>One</td><td><strong>Two</strong></td></tr></table>',
  ].join(''));

  assert.equal(result.modified, true);
  assert.match(result.html, /hyperlink-https%3A%2F%2Fexample\.com%2Fa%3Fx%3D1/);
  assert.match(result.html, /image:https%3A%2F%2Fcdn\.example%2Fimage\.png/);
  assert.match(result.html, /image-width:320px/);
  assert.match(result.html, /image-height:160px/);
  assert.match(result.html, /imageCssAspectRatio:2\.0000/);
  assert.match(result.html, /color:#336699/);
  assert.match(result.html, /tbljson-/);
  assert.match(result.document.body.textContent, new RegExp(`One${DELIMITER}Two`));
  assert.match(result.document.body.textContent, new RegExp(ZWSP));
  assert.equal(result.document.querySelector('script'), null);
});

test('normalizes headings, alignment, ordered lists, and super/subscript', () => {
  const result = transform([
    '<h2 style="text-align:center"><span style="color:green;font-size:29px">Title</span></h2>',
    '<ol start="3"><li>Third</li><li>Fourth</li></ol>',
    '<span style="vertical-align:super">2</span><span style="vertical-align:sub">n</span>',
  ].join(''));
  assert.ok(result.document.querySelector('center h2'));
  assert.match(result.document.body.textContent, /3\. Third/);
  assert.match(result.document.body.textContent, /4\. Fourth/);
  assert.equal(result.document.querySelector('h2 span').className.includes('font-size:29'), false);
  assert.ok(result.document.querySelector('sup'));
  assert.ok(result.document.querySelector('sub'));
});

test('returns false for a document requiring no transformation', () => {
  assert.equal(transform('<p>Plain text</p>').modified, false);
});

test('collects Etherpad inline style classes preserved by table clipboard HTML', () => {
  const calls = [];
  const state = {};
  collectContentPre('collectContentPre', {
    cls: 'author-a.test b i u s sub sup',
    state,
    cc: {doAttrib: (...args) => calls.push(args)},
  });
  assert.deepEqual(calls, [
    [state, 'bold'],
    [state, 'italic'],
    [state, 'underline'],
    [state, 'strikethrough'],
    [state, 'sub'],
    [state, 'sup'],
  ]);
});

test('emits copied Etherpad table rows as sibling lines and preserves nested styles', () => {
  const oldMeta = (row) => Buffer.from(JSON.stringify({
    tblId: 'source-table', row, cols: 2,
  })).toString('base64');
  const result = transform([
    '<div class="ace-line" id="magicdomid1">Before table</div>',
    '<div class="ace-line" id="magicdomid2">',
    '<table class="dataTable" data-tblid="source-table" data-row="0"><tbody><tr>',
    `<td><span class="tbljson-${oldMeta(0)} tblCell-0 b i u">`,
    '<b><i><u>nested style</u></i></b></span></td>',
    `<td><span class="tbljson-${oldMeta(0)} tblCell-1">`,
    '<a href="https://example.com/nested"><b><i>linked style</i></b></a></span></td>',
    '</tr></tbody></table>',
    '</div>',
    '<div class="ace-line" id="magicdomid3">',
    '<table class="dataTable" data-tblid="source-table" data-row="1"><tbody><tr>',
    `<td><span class="tbljson-${oldMeta(1)} tblCell-0 sub"><sub>2</sub></span></td>`,
    `<td><span class="tbljson-${oldMeta(1)} tblCell-1 sup"><sup>3</sup></span></td>`,
    '</tr></tbody></table>',
    '</div>',
    '<div class="ace-line" id="magicdomid4">After table</div>',
  ].join(''));

  const children = Array.from(result.document.body.children);
  assert.equal(children.length, 4);
  assert.equal(children[0].textContent, 'Before table');
  assert.equal(children[3].textContent, 'After table');
  assert.equal(result.document.querySelector('.ace-line > div'), null);
  assert.equal(result.document.querySelector('table'), null);
  assert.ok(children[1].querySelectorAll('[class*="tbljson-"]').length >= 3);
  assert.equal(children[2].querySelectorAll('[class*="tbljson-"]').length, 3);
  assert.ok(children[1].querySelector('b > i > u'));
  assert.ok(children[1].querySelector('.hyperlink b > i'));
  assert.ok(children[2].querySelector('sub'));
  assert.ok(children[2].querySelector('sup'));
});
