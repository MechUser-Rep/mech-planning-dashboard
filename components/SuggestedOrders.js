// components/SuggestedOrders.js
const SuggestedOrders = {
  name: 'SuggestedOrders',
  props: {
    rows:       { type: Array,  required: true },
    weekKeys:   { type: Array,  required: true },
    mechLookup: { type: Object, required: true },
  },
  computed: {
    ordersNeeded() {
      return this.rows
        .filter(r => r.needsOrder)
        .map(r => ({
          newCode:      r.newCode,
          currentStock: r.currentStock,
          minLevel:     r.minLevel,
          price:        r.price,
          suggestedQty: Calculations.calcSuggestedOrderQty(r, this.weekKeys, this.mechLookup),
          totalCost:    r.price * Calculations.calcSuggestedOrderQty(r, this.weekKeys, this.mechLookup),
        }));
    },
    grandTotal() {
      return this.ordersNeeded.reduce((sum, o) => sum + o.totalCost, 0);
    },
  },
  methods: {
    fmt(n) { return typeof n === 'number' ? n.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : n; },
  },
  template: `
    <section class="suggested-orders" v-if="ordersNeeded.length">
      <h2>Suggested Orders <span class="order-count">({{ ordersNeeded.length }})</span></h2>
      <table class="orders-table">
        <thead>
          <tr>
            <th>Code</th>
            <th>Stock now</th>
            <th>Min level</th>
            <th>Suggested qty</th>
            <th>Unit price (£)</th>
            <th>Total cost (£)</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="o in ordersNeeded" :key="o.newCode">
            <td>{{ o.newCode }}</td>
            <td>{{ o.currentStock }}</td>
            <td>{{ o.minLevel }}</td>
            <td class="suggested-qty">{{ o.suggestedQty }}</td>
            <td>{{ fmt(o.price) }}</td>
            <td>{{ fmt(o.totalCost) }}</td>
          </tr>
        </tbody>
        <tfoot>
          <tr>
            <td colspan="5" class="grand-total-label">Total suggested spend</td>
            <td class="grand-total">£{{ fmt(grandTotal) }}</td>
          </tr>
        </tfoot>
      </table>
    </section>
    <section class="suggested-orders" v-else>
      <p class="all-clear">All mechanisms projected above minimum stock for the next 12 weeks.</p>
    </section>
  `,
};
