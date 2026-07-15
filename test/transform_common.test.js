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
