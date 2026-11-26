import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resolveHistoricalReplyAuthor, isActiveReplyRecipient } from './reply-recipient-safety.js';
const staff = [
  {
    "id": "1",
    "name": "Sam",
    "active": true
  },
  {
    "id": "2",
    "name": "Sam Patel",
    "active": true
  },
  {
    "id": "3",
    "name": "Lee Jones",
    "active": false
  },
  {
    "id": "4",
    "name": "Lee",
    "active": true
  },
  {
    "id": "5",
    "name": "Chris Patel",
    "active": true
  },
  {
    "id": "6",
    "name": "Chris Jones",
    "active": true
  },
  {
    "id": "7",
    "name": "Jai Sharma"
  },
  {
    "id": "8",
    "name": "Dup"
  },
  {
    "id": "9",
    "name": "Dup"
  }
];
const cases = [["Sam","1"],["SAM PATEL","2"],["Lee Jones",null],["Lee","4"],["Chris",null],["Jai","7"],["Dup",null],["",null],["unknown",null],["  jai sharma ","7"]] as const;
for (const [name,id] of cases) test(`historical author ${name || '(empty)'}`,()=>assert.equal(resolveHistoricalReplyAuthor(staff,name)?.id??null,id));
test('inactive recipients are excluded and legacy active state remains allowed',()=>{assert.equal(isActiveReplyRecipient(staff[2]),false);assert.equal(isActiveReplyRecipient(staff[6]),true);});
