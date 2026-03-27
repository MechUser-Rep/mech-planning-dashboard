// tests/data.test.js
// Require data.js from parent directory
const DataLoader = require('../data.js');

// ── parseMechanismsLookup ───────────────────────────────────────────────────

describe('parseMechanismsLookup', () => {
  const rows = [
    ['Col A', 'Col B', 'Col C', 'New Code', 'Sage code', 'Sage Code 2', 'Seminar Code', 'Seminar Code 2', 'Lead-time'],
    ['',      '',      '',      'MECH-01',  'SAGE-A',    'SAGE-A2',     'SEM-A',        'SEM-A2',         '14'],
    ['',      '',      '',      'MECH-02',  'SAGE-B',    '',            'SEM-B',        '',               '21'],
  ];

  test('builds lookup by New Code', () => {
    const { lookup } = DataLoader.parseMechanismsLookup(rows);
    expect(lookup['MECH-01']).toMatchObject({ newCode: 'MECH-01', leadTimeDays: 14 });
    expect(lookup['MECH-02']).toMatchObject({ newCode: 'MECH-02', leadTimeDays: 21 });
  });

  test('builds bySageCode reverse map including Code 2', () => {
    const { bySageCode } = DataLoader.parseMechanismsLookup(rows);
    expect(bySageCode['SAGE-A']).toBe('MECH-01');
    expect(bySageCode['SAGE-A2']).toBe('MECH-01');
    expect(bySageCode['SAGE-B']).toBe('MECH-02');
  });

  test('builds bySeminarCode reverse map', () => {
    const { bySeminarCode } = DataLoader.parseMechanismsLookup(rows);
    expect(bySeminarCode['SEM-A']).toBe('MECH-01');
    expect(bySeminarCode['SEM-A2']).toBe('MECH-01');
  });

  test('skips rows with empty New Code', () => {
    const withBlank = [...rows, ['', '', '', '', 'SAGE-C', '', '', '', '7']];
    const { lookup } = DataLoader.parseMechanismsLookup(withBlank);
    expect(Object.keys(lookup)).toHaveLength(2);
  });
});

// ── parseSortlyStock ────────────────────────────────────────────────────────

describe('parseSortlyStock', () => {
  // Col A=EntryName, J=Qty(idx9), L=MinLevel(idx11), M=Price(idx12)
  function makeRow(entryName, qty, minLevel, price) {
    const row = new Array(13).fill('');
    row[0]  = entryName;
    row[9]  = qty;
    row[11] = minLevel;
    row[12] = price;
    return row;
  }

  const rows = [
    makeRow('Entry Name', 'Quantity', 'Min Level', 'Price'), // header
    makeRow('MECH-01', 50, 20, 12.50),
    makeRow('MECH-02', 10, 15, 8.00),
    makeRow('',        0,  0,  0),   // blank — should be skipped
  ];

  test('parses quantity, minLevel and price', () => {
    const stock = DataLoader.parseSortlyStock(rows);
    expect(stock['MECH-01']).toEqual({ entryName: 'MECH-01', quantity: 50, minLevel: 20, price: 12.5 });
  });

  test('skips blank entry names', () => {
    const stock = DataLoader.parseSortlyStock(rows);
    expect(Object.keys(stock)).toHaveLength(2);
  });
});
