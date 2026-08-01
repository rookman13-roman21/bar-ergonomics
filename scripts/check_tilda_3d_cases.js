#!/usr/bin/env node
'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '..', 'tilda-blocks', '05-portfolio.html'), 'utf8');

assert.match(source, /https:\/\/drawings\.barista-school\.ru\/api\/public\/cases/);
assert.match(source, /\^\\\/api\\\/public\\\/case-assets/);
assert.match(source, /\^\[A-Za-z0-9_-\]\{20,64\}\$/);
assert.match(source, /credentials: 'omit'/);
assert.match(source, /textContent = value/);
assert.match(source, /appendText\(heading, 'h2', '', item\.name\)/);
assert.match(source, /scroll-snap-type: x mandatory/);
assert.match(source, /touch-action: pan-y/);
assert.match(source, /mbs-3d-cases-dialog__counter/);
assert.match(source, /utm_source', 'baristaschool'/);
assert.match(source, /utm_medium', 'case_popup'/);
assert.match(source, /createProjectLink\([^\n]+, item, 'Смотреть 3D-проект'\)/);
assert.match(source, /lead\.textContent = 'Оставить заявку'/);
assert.match(source, /document\.body\.style\.position = 'fixed'/);
assert.doesNotMatch(source, /rutube\.ru/);
assert.doesNotMatch(source, /iframe/);

const script = source.match(/<script>\s*([\s\S]*?)\s*<\/script>/u)?.[1];
assert.ok(script, 'inline script is missing');
new Function(script); // eslint-disable-line no-new-func

console.log('check_tilda_3d_cases: OK');
