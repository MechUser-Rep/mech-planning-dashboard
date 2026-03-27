// tests/calculations.test.js
const Calculations = require('../calculations.js');

const weekKeys = ['2026-03-23', '2026-03-30', '2026-04-06',
                  '2026-04-13', '2026-04-20', '2026-04-27',
                  '2026-05-04', '2026-05-11', '2026-05-18',
                  '2026-05-25', '2026-06-01', '2026-06-08'];

// ── buildMechRows ───────────────────────────────────────────────────────────

describe('buildMechRows', () => {
  const mechLookup = {
    lookup: {
      'MECH-01': { newCode: 'MECH-01', leadTimeDays: 14 },
    },
    bySageCode: {}, bySeminarCode: {},
  };
  const stockMap = { 'MECH-01': { quantity: 30, minLevel: 20, price: 10 } };

  test('projects stock correctly week by week', () => {
    const demand   = { 'MECH-01': { '2026-03-23': 5, '2026-03-30': 5 } };
    const incoming = { 'MECH-01': { '2026-03-30': 10 } };
    const rows = Calculations.buildMechRows(mechLookup, stockMap, demand, incoming, weekKeys);
    const row = rows[0];
    // Week 1: 30 + 0 - 5 = 25
    expect(row.weeks[0].projected).toBe(25);
    // Week 2: 25 + 10 - 5 = 30
    expect(row.weeks[1].projected).toBe(30);
    // Week 3 onwards: no demand, no incoming — stays at 30
    expect(row.weeks[2].projected).toBe(30);
  });

  test('sets needsOrder true when projected drops below minLevel', () => {
    const demand   = { 'MECH-01': { '2026-03-23': 25 } }; // drops to 5, below min 20
    const incoming = {};
    const rows = Calculations.buildMechRows(mechLookup, stockMap, demand, incoming, weekKeys);
    expect(rows[0].needsOrder).toBe(true);
  });

  test('sets needsOrder false when always at or above minLevel', () => {
    const demand   = {};
    const incoming = {};
    const rows = Calculations.buildMechRows(mechLookup, stockMap, demand, incoming, weekKeys);
    expect(rows[0].needsOrder).toBe(false);
  });
});

// ── calcSuggestedOrderQty ───────────────────────────────────────────────────

describe('calcSuggestedOrderQty', () => {
  const mechLookup = {
    lookup: { 'MECH-01': { newCode: 'MECH-01', leadTimeDays: 14 } },
    bySageCode: {}, bySeminarCode: {},
  };

  test('returns 0 when no order needed', () => {
    const row = {
      newCode: 'MECH-01', needsOrder: false, minLevel: 20, minProjected: 25,
      weeks: weekKeys.map(wk => ({ weekKey: wk, demand: 0, incoming: 0, projected: 25 })),
    };
    expect(Calculations.calcSuggestedOrderQty(row, weekKeys, mechLookup)).toBe(0);
  });

  test('calculates qty to restore min stock and cover lead-time demand', () => {
    // minLevel=20, minProjected=5 (shortfall=15), demand 3/week, lead=14d=2 weeks demand=6
    // suggested = 20 - 5 + 6 = 21
    const row = {
      newCode: 'MECH-01', needsOrder: true, minLevel: 20, minProjected: 5,
      weeks: weekKeys.map(wk => ({ weekKey: wk, demand: 3, incoming: 0, projected: 5 })),
    };
    expect(Calculations.calcSuggestedOrderQty(row, weekKeys, mechLookup)).toBe(21);
  });
});

// ── getStockColour ──────────────────────────────────────────────────────────

describe('getStockColour', () => {
  test('red when projected < 0', () => {
    expect(Calculations.getStockColour(-1, 20)).toBe('red');
  });
  test('light-pink when 0 <= projected < minLevel', () => {
    expect(Calculations.getStockColour(10, 20)).toBe('light-pink');
    expect(Calculations.getStockColour(0,  20)).toBe('light-pink');
  });
  test('green when projected between minLevel and minLevel*1.05', () => {
    expect(Calculations.getStockColour(20, 20)).toBe('green');
    expect(Calculations.getStockColour(21, 20)).toBe('green');
  });
  test('orange when projected > minLevel*1.05', () => {
    expect(Calculations.getStockColour(22, 20)).toBe('orange');
  });
});
