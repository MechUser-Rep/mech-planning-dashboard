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

// ── parseProductionDemand ───────────────────────────────────────────────────

describe('parseProductionDemand', () => {
  // Simulate two sheets. L2 = week commencing date at row index 1, col index 11.
  // Headers in row 3 (index 3). Data starts row 4 (index 4).
  function makeSheet(weekDate, jobs) {
    const rows = [
      new Array(20).fill(''),  // row 1 (index 0) — unused
      [...new Array(11).fill(''), weekDate, ...new Array(8).fill('')],  // L2 = index 1,11
      new Array(20).fill(''),  // row 3 — unused
      // row 4 = headers (index 3)
      ['Job', 'Ref', 'Chair', 'ITEMS', ...new Array(6).fill(''), 'Mechanism \u2013 1', 'Mechanism \u2013 2'],
    ];
    for (const [items, m1, m2] of jobs) {
      const row = new Array(20).fill('');
      row[3] = items;
      row[10] = m1;
      row[11] = m2;
      rows.push(row);
    }
    return rows;
  }

  const mechLookup = {
    lookup: { 'MECH-01': {}, 'MECH-02': {}, 'MECH-03': {} },
    bySageCode: {}, bySeminarCode: {},
  };

  const weekKeys = DataLoader.get12WeekKeys(new Date('2026-03-23')); // Monday

  test('counts 1 per populated ITEMS row for each mechanism', () => {
    const sheet = makeSheet('2026-03-23', [
      [1,      'MECH-01', 'MECH-02'],
      ['X',    'MECH-01', ''],
      ['yes',  'MECH-03', ''],
    ]);
    const allSheets = { 'WK 13': sheet };
    const { demand } = DataLoader.parseProductionDemand(allSheets, weekKeys, mechLookup);
    expect(demand['MECH-01']['2026-03-23']).toBe(2);
    expect(demand['MECH-02']['2026-03-23']).toBe(1);
    expect(demand['MECH-03']['2026-03-23']).toBe(1);
  });

  test('skips rows with empty ITEMS', () => {
    const sheet = makeSheet('2026-03-23', [
      ['', 'MECH-01', ''],
      [1,  'MECH-01', ''],
    ]);
    const allSheets = { 'WK 13': sheet };
    const { demand } = DataLoader.parseProductionDemand(allSheets, weekKeys, mechLookup);
    expect(demand['MECH-01']['2026-03-23']).toBe(1);
  });

  test('records unmatched mechanism codes', () => {
    const sheet = makeSheet('2026-03-23', [
      [1, 'UNKNOWN-99', ''],
    ]);
    const allSheets = { 'WK 13': sheet };
    const { unmatched } = DataLoader.parseProductionDemand(allSheets, weekKeys, mechLookup);
    expect(unmatched.some(u => u.includes('UNKNOWN-99'))).toBe(true);
  });
});

// ── parsePOListingIncoming ──────────────────────────────────────────────────

describe('parsePOListingIncoming', () => {
  // Headers in row index 1 (Excel row 2).
  const headers = new Array(15).fill('');
  headers[2]  = 'PurchaseOrder.AccountReference';
  headers[3]  = 'PurchaseOrder.Date';
  headers[7]  = 'PurchaseOrderItem.Description';
  headers[9]  = 'PurchaseOrderItem.Quantity';

  function makePoRow(accountRef, date, desc, qty) {
    const row = new Array(15).fill('');
    row[2] = accountRef;
    row[3] = date;
    row[7] = desc;
    row[9] = qty;
    return row;
  }

  // Lead-time of 7 days for MECH-01, so PO date 2026-03-23 → delivery 2026-03-30
  const mechLookup = {
    lookup: { 'MECH-01': { leadTimeDays: 7 }, 'MECH-02': { leadTimeDays: 0 } },
    bySageCode: { 'SAGE-A': 'MECH-01', 'SAGE-B': 'MECH-02' },
    bySeminarCode: {},
  };

  const weekKeys = DataLoader.get12WeekKeys(new Date('2026-03-23'));

  const rows = [
    new Array(15).fill(''),  // row index 0 — unused
    headers,                 // row index 1 — headers
    makePoRow('SUPP1',   '2026-03-23', 'SAGE-A', 10),  // delivery 2026-03-30 (week 2)
    makePoRow('SEMINARC2','2026-03-23', 'SAGE-B', 5),   // excluded — SEMINARC2
    makePoRow('SUPP1',   '2026-03-23', 'SAGE-B', 20),  // delivery 2026-03-23 (week 1, 0 lead)
    makePoRow('SUPP1',   '2025-01-01', 'SAGE-A', 3),   // outside 12-week window — excluded
  ];

  test('buckets delivery into correct week', () => {
    const { incoming } = DataLoader.parsePOListingIncoming(rows, weekKeys, mechLookup);
    expect(incoming['MECH-01']['2026-03-30']).toBe(10);
  });

  test('excludes SEMINARC2 rows', () => {
    const { incoming } = DataLoader.parsePOListingIncoming(rows, weekKeys, mechLookup);
    expect(incoming['MECH-02'] && incoming['MECH-02']['2026-03-23']).toBe(20);
  });

  test('excludes POs outside the 12-week window', () => {
    const { incoming } = DataLoader.parsePOListingIncoming(rows, weekKeys, mechLookup);
    // Only 10 for MECH-01, not 13
    expect(incoming['MECH-01']['2026-03-30']).toBe(10);
    expect(Object.values(incoming['MECH-01'] || {}).reduce((a, b) => a + b, 0)).toBe(10);
  });
});

// ── parseSeminarIncoming ────────────────────────────────────────────────────

describe('parseSeminarIncoming', () => {
  const headers = new Array(15).fill('');
  headers[5] = 'Description';
  headers[7] = 'Due Date';
  headers[9] = 'Quantity';

  function makeSemRow(desc, dueDate, qty) {
    const row = new Array(15).fill('');
    row[5] = desc; row[7] = dueDate; row[9] = qty;
    return row;
  }

  const mechLookup = {
    lookup: { 'MECH-01': { leadTimeDays: 14 } },
    bySageCode: {},
    bySeminarCode: { 'SEM-A': 'MECH-01' },
  };

  const weekKeys = DataLoader.get12WeekKeys(new Date('2026-03-23'));

  const rows = [
    headers,
    makeSemRow('SEM-A', '2026-03-25', 8),
    makeSemRow('UNKNOWN', '2026-03-25', 2),
  ];

  test('uses Due Date directly without adding lead time', () => {
    const { incoming } = DataLoader.parseSeminarIncoming(rows, weekKeys, mechLookup);
    expect(incoming['MECH-01']['2026-03-23']).toBe(8); // 2026-03-25 buckets to Monday 2026-03-23
  });

  test('records unmatched Seminar descriptions', () => {
    const { unmatched } = DataLoader.parseSeminarIncoming(rows, weekKeys, mechLookup);
    expect(unmatched.some(u => u.includes('UNKNOWN'))).toBe(true);
  });
});
