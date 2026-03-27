// data.js
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.DataLoader = factory();
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {

  // ── Utility ──────────────────────────────────────────────────────────────

  function parseExcelDate(value) {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') {
      // Excel serial date: days since 1900-01-00 (with 1900 leap-year bug offset)
      return new Date((value - 25569) * 86400 * 1000);
    }
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
  }

  function getMondayOfWeek(date) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    const day = d.getDay(); // 0=Sun
    d.setDate(d.getDate() + (day === 0 ? -6 : 1 - day));
    return d;
  }

  function toWeekKey(date) {
    return getMondayOfWeek(date).toISOString().slice(0, 10);
  }

  function get12WeekKeys(fromDate) {
    const monday = getMondayOfWeek(fromDate);
    return Array.from({ length: 12 }, (_, i) => {
      const d = new Date(monday);
      d.setDate(d.getDate() + i * 7);
      return d.toISOString().slice(0, 10);
    });
  }

  // ── parseMechanismsLookup ─────────────────────────────────────────────────
  // rows: 2D array from usedRange. Row 0 = headers.
  // Returns: { lookup: { newCode → mechObj }, bySageCode: {}, bySeminarCode: {} }

  function parseMechanismsLookup(rows) {
    const headers = rows[0].map(h => String(h).trim());
    const col = name => headers.indexOf(name);

    const newCodeIdx    = col('New Code');
    const sageIdx       = col('Sage code');
    const sage2Idx      = col('Sage Code 2');
    const seminarIdx    = col('Seminar Code');
    const seminar2Idx   = col('Seminar Code 2');
    const leadTimeIdx   = col('Lead-time');

    const lookup        = {};
    const bySageCode    = {};
    const bySeminarCode = {};

    for (let i = 1; i < rows.length; i++) {
      const row     = rows[i];
      const newCode = String(row[newCodeIdx] || '').trim();
      if (!newCode) continue;

      const mech = {
        newCode,
        sageCode:     String(row[sageIdx]     || '').trim(),
        sageCode2:    String(row[sage2Idx]    || '').trim(),
        seminarCode:  String(row[seminarIdx]  || '').trim(),
        seminarCode2: String(row[seminar2Idx] || '').trim(),
        leadTimeDays: parseInt(row[leadTimeIdx], 10) || 0,
      };
      lookup[newCode] = mech;

      if (mech.sageCode)     bySageCode[mech.sageCode]     = newCode;
      if (mech.sageCode2)    bySageCode[mech.sageCode2]    = newCode;
      if (mech.seminarCode)  bySeminarCode[mech.seminarCode]  = newCode;
      if (mech.seminarCode2) bySeminarCode[mech.seminarCode2] = newCode;
    }

    return { lookup, bySageCode, bySeminarCode };
  }

  // ── parseSortlyStock ──────────────────────────────────────────────────────
  // Fixed columns: A(0)=EntryName, J(9)=Qty, L(11)=MinLevel, M(12)=Price
  // Returns: { entryName → { entryName, quantity, minLevel, price } }

  function parseSortlyStock(rows) {
    const stock = {};
    for (let i = 1; i < rows.length; i++) {
      const row       = rows[i];
      const entryName = String(row[0] || '').trim();
      if (!entryName) continue;
      stock[entryName] = {
        entryName,
        quantity: parseFloat(row[9])  || 0,
        minLevel: parseFloat(row[11]) || 0,
        price:    parseFloat(row[12]) || 0,
      };
    }
    return stock;
  }

  // ── parseProductionDemand ─────────────────────────────────────────────────
  // allSheets: { sheetName: 2D array }  — all sheets from the production workbook
  // weekKeys:  array of 12 'YYYY-MM-DD' strings (Mondays)
  // mechLookup: result of parseMechanismsLookup
  // Returns: { demand: { newCode: { weekKey: count } }, unmatched: string[] }

  function parseProductionDemand(allSheets, weekKeys, mechLookup) {
    const demand    = {};
    const unmatched = new Set();

    for (const weekKey of weekKeys) {
      let matchingRows = null;

      for (const rows of Object.values(allSheets)) {
        if (!rows || rows.length < 2) continue;
        const wcDate = parseExcelDate(rows[1][11]); // L2
        if (wcDate && toWeekKey(wcDate) === weekKey) {
          matchingRows = rows;
          break;
        }
      }

      if (!matchingRows) continue; // no sheet for this week — zero demand

      const headers  = (matchingRows[3] || []).map(h => String(h).trim());
      const itemsIdx = headers.indexOf('ITEMS');
      const mech1Idx = headers.indexOf('Mechanism \u2013 1');
      const mech2Idx = headers.indexOf('Mechanism \u2013 2');

      if (itemsIdx === -1 || mech1Idx === -1) continue;

      for (let i = 4; i < matchingRows.length; i++) {
        const row      = matchingRows[i];
        const itemsVal = row[itemsIdx];
        if (itemsVal === null || itemsVal === undefined || itemsVal === '') continue;

        const codes = [
          String(row[mech1Idx] || '').trim(),
          mech2Idx !== -1 ? String(row[mech2Idx] || '').trim() : '',
        ];

        for (const code of codes) {
          if (!code) continue;
          if (!mechLookup.lookup[code]) {
            unmatched.add(`Production plan: "${code}"`);
            continue;
          }
          if (!demand[code]) demand[code] = {};
          demand[code][weekKey] = (demand[code][weekKey] || 0) + 1;
        }
      }
    }

    return { demand, unmatched: [...unmatched] };
  }

  // ── parsePOListingIncoming ────────────────────────────────────────────────
  // rows: 2D array, headers in row index 1 (Excel row 2).
  // Excludes SEMINARC2 rows. Delivery date = PO date + lead-time days.

  function parsePOListingIncoming(rows, weekKeys, mechLookup) {
    const headers = (rows[1] || []).map(h => String(h).trim());
    const col = name => headers.indexOf(name);

    const accountRefIdx = col('PurchaseOrder.AccountReference');
    const dateIdx       = col('PurchaseOrder.Date');
    const descIdx       = col('PurchaseOrderItem.Description');
    const qtyIdx        = col('PurchaseOrderItem.Quantity');

    const windowStart = new Date(weekKeys[0]);
    const windowEnd   = new Date(weekKeys[weekKeys.length - 1]);
    windowEnd.setDate(windowEnd.getDate() + 6); // end of last week Sunday

    const incoming  = {};
    const unmatched = new Set();

    for (let i = 2; i < rows.length; i++) {
      const row        = rows[i];
      const accountRef = String(row[accountRefIdx] || '').trim();
      if (accountRef === 'SEMINARC2') continue;

      const desc    = String(row[descIdx] || '').trim();
      const newCode = mechLookup.bySageCode[desc];
      if (!newCode) {
        if (desc) unmatched.add(`PO Listing: "${desc}"`);
        continue;
      }

      const poDate = parseExcelDate(row[dateIdx]);
      if (!poDate) continue;

      const leadDays    = mechLookup.lookup[newCode]?.leadTimeDays || 0;
      const deliveryDate = new Date(poDate);
      deliveryDate.setDate(deliveryDate.getDate() + leadDays);

      if (deliveryDate < windowStart || deliveryDate > windowEnd) continue;

      const weekKey = toWeekKey(deliveryDate);
      const qty     = parseFloat(row[qtyIdx]) || 0;

      if (!incoming[newCode]) incoming[newCode] = {};
      incoming[newCode][weekKey] = (incoming[newCode][weekKey] || 0) + qty;
    }

    return { incoming, unmatched: [...unmatched] };
  }

  // ── parseSeminarIncoming ──────────────────────────────────────────────────
  // rows: 2D array, headers in row 0. Uses Due Date (col H, idx 7) directly.

  function parseSeminarIncoming(rows, weekKeys, mechLookup) {
    const headers  = (rows[0] || []).map(h => String(h).trim());
    const descIdx  = 5; // col F
    const dateIdx  = 7; // col H — Due Date
    const qtyIdx   = headers.indexOf('Quantity') !== -1
      ? headers.indexOf('Quantity')
      : headers.findIndex(h => /qty/i.test(h));

    const windowStart = new Date(weekKeys[0]);
    const windowEnd   = new Date(weekKeys[weekKeys.length - 1]);
    windowEnd.setDate(windowEnd.getDate() + 6);

    const incoming  = {};
    const unmatched = new Set();

    for (let i = 1; i < rows.length; i++) {
      const row     = rows[i];
      const desc    = String(row[descIdx] || '').trim();
      const newCode = mechLookup.bySeminarCode[desc];
      if (!newCode) {
        if (desc) unmatched.add(`Seminar orders: "${desc}"`);
        continue;
      }

      const dueDate = parseExcelDate(row[dateIdx]);
      if (!dueDate || dueDate < windowStart || dueDate > windowEnd) continue;

      const weekKey = toWeekKey(dueDate);
      const qty     = qtyIdx !== -1 ? (parseFloat(row[qtyIdx]) || 1) : 1;

      if (!incoming[newCode]) incoming[newCode] = {};
      incoming[newCode][weekKey] = (incoming[newCode][weekKey] || 0) + qty;
    }

    return { incoming, unmatched: [...unmatched] };
  }

  // ── mergeIncoming ─────────────────────────────────────────────────────────
  // Combines two incoming maps (PO Listing + Seminar) by summing quantities.

  function mergeIncoming(a, b) {
    const merged = {};
    for (const [code, weeks] of Object.entries(a)) {
      merged[code] = { ...weeks };
    }
    for (const [code, weeks] of Object.entries(b)) {
      if (!merged[code]) merged[code] = {};
      for (const [wk, qty] of Object.entries(weeks)) {
        merged[code][wk] = (merged[code][wk] || 0) + qty;
      }
    }
    return merged;
  }

  // Exports
  return {
    parseExcelDate,
    getMondayOfWeek,
    toWeekKey,
    get12WeekKeys,
    parseMechanismsLookup,
    parseSortlyStock,
    parseProductionDemand,
    parsePOListingIncoming,
    parseSeminarIncoming,
    mergeIncoming,
  };
}));
