// calculations.js
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.Calculations = factory();
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {

  // Builds a row object per mechanism with week-by-week demand, incoming, projected stock.
  function buildMechRows(mechLookup, stockMap, demand, incoming, weekKeys) {
    return Object.values(mechLookup.lookup).map(mech => {
      const { newCode } = mech;
      const stock = stockMap[newCode] || { quantity: 0, minLevel: 0, price: 0 };

      let projected = stock.quantity;
      const weeks = weekKeys.map(weekKey => {
        const wDemand   = (demand[newCode]   && demand[newCode][weekKey])   || 0;
        const wIncoming = (incoming[newCode] && incoming[newCode][weekKey]) || 0;
        projected = projected + wIncoming - wDemand;
        return { weekKey, demand: wDemand, incoming: wIncoming, projected };
      });

      const minProjected = Math.min(...weeks.map(w => w.projected));
      const needsOrder   = minProjected < stock.minLevel;

      return {
        newCode,
        currentStock: stock.quantity,
        minLevel:     stock.minLevel,
        price:        stock.price,
        weeks,
        needsOrder,
        minProjected,
      };
    });
  }

  // Suggested order qty: enough to restore min stock and cover demand through lead time.
  function calcSuggestedOrderQty(row, weekKeys, mechLookup) {
    if (!row.needsOrder) return 0;
    const leadDays  = mechLookup.lookup[row.newCode]?.leadTimeDays || 0;
    const leadWeeks = Math.ceil(leadDays / 7);
    const demandDuringLead = row.weeks
      .slice(0, leadWeeks)
      .reduce((sum, w) => sum + w.demand, 0);
    return Math.max(0, row.minLevel - row.minProjected + demandDuringLead);
  }

  // Returns a CSS class string used for projected-stock cell colouring.
  function getStockColour(projected, minLevel) {
    if (projected < 0)              return 'red';
    if (projected < minLevel)       return 'light-pink';
    if (projected <= minLevel * 1.05) return 'green';
    return 'orange';
  }

  return { buildMechRows, calcSuggestedOrderQty, getStockColour };
}));
