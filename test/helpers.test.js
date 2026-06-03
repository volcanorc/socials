import assert from 'node:assert/strict';
import test from 'node:test';

import { columnNumberToName, parseSheetTable } from '../lib/google.js';
import { buildSessionSetCookie, createRandomState, readSessionCookie } from '../lib/session.js';

test('columnNumberToName converts spreadsheet columns', () => {
  assert.equal(columnNumberToName(1), 'A');
  assert.equal(columnNumberToName(26), 'Z');
  assert.equal(columnNumberToName(27), 'AA');
  assert.equal(columnNumberToName(52), 'AZ');
});

test('parseSheetTable extracts headers and rows', () => {
  const table = parseSheetTable([
    ['Name', 'Email'],
    ['Ada', 'ada@example.com'],
    ['Linus', 'linus@example.com'],
  ]);
  assert.deepEqual(table.headers, ['Name', 'Email']);
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[0].rowNumber, 2);
  assert.equal(table.rows[0].values.Name, 'Ada');
});

test('session cookies round trip', () => {
  const secret = 'unit-test-secret';
  const cookie = buildSessionSetCookie(secret, { hello: 'world' });
  const session = readSessionCookie(secret, cookie);
  assert.equal(session.hello, 'world');
});

test('createRandomState is non-empty', () => {
  assert.ok(createRandomState().length > 10);
});
