// components/PlanningGrid.js
// Requires Vue 3 and Calculations loaded globally before this script.

const PlanningGrid = {
  name: 'PlanningGrid',
  props: {
    rows:       { type: Array,   required: true },
    weekKeys:   { type: Array,   required: true },
    showAtRisk: { type: Boolean, default: false },
    jobDemand:  { type: Object,  default: () => ({}) }, // { newCode: { weekKey: count } }
  },
  emits: ['select-row'],
  data() {
    return { expandedRow: null };
  },
  computed: {
    visibleRows() {
      const sorted = [...this.rows].sort((a, b) => {
        const aAtRisk = a.needsOrder ? 0 : 1;
        const bAtRisk = b.needsOrder ? 0 : 1;
        return aAtRisk - bAtRisk;
      });
      return this.showAtRisk ? sorted.filter(r => r.needsOrder) : sorted;
    },
  },
  methods: {
    toggleRow(newCode) {
      this.expandedRow = this.expandedRow === newCode ? null : newCode;
    },
    colourClass(projected, minLevel) {
      return 'stock-' + Calculations.getStockColour(projected, minLevel);
    },
    formatWeekLabel(weekKey) {
      const d = new Date(weekKey);
      return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
    },
    // Returns [{ weekKey, count }] for the expanded row's job breakdown.
    // jobDemand prop: { newCode: { weekKey: count } } — passed from app.js
    getJobsForRow(newCode) {
      if (!this.jobDemand || !this.jobDemand[newCode]) return [];
      return this.weekKeys
        .map(wk => ({ weekKey: wk, count: this.jobDemand[newCode][wk] || 0 }))
        .filter(j => j.count > 0);
    },
  },
  template: `
    <div class="planning-grid-wrapper">
      <table class="planning-grid">
        <thead>
          <tr>
            <th class="col-code sticky-left">Code</th>
            <th class="col-stock sticky-left-2">Stock / Min</th>
            <th v-for="wk in weekKeys" :key="wk" class="col-week">
              {{ formatWeekLabel(wk) }}
            </th>
          </tr>
        </thead>
        <tbody>
          <template v-for="row in visibleRows" :key="row.newCode">
            <tr
              class="mech-row"
              :class="{ 'at-risk': row.needsOrder, 'expanded': expandedRow === row.newCode }"
              @click="toggleRow(row.newCode)"
            >
              <td class="col-code sticky-left">{{ row.newCode }}</td>
              <td class="col-stock sticky-left-2">{{ row.currentStock }} / {{ row.minLevel }}</td>
              <td
                v-for="w in row.weeks"
                :key="w.weekKey"
                class="col-week week-cell"
                :class="colourClass(w.projected, row.minLevel)"
              >
                <span class="cell-demand">{{ w.demand || '' }}</span>
                <span class="cell-incoming">{{ w.incoming ? '+' + w.incoming : '' }}</span>
                <span class="cell-projected">{{ w.projected }}</span>
              </td>
            </tr>
            <tr v-if="expandedRow === row.newCode" class="detail-row">
              <td :colspan="2 + weekKeys.length" class="detail-cell">
                <table class="job-detail-table" v-if="getJobsForRow(row.newCode).length">
                  <thead><tr><th>Week</th><th>Job count</th></tr></thead>
                  <tbody>
                    <tr v-for="j in getJobsForRow(row.newCode)" :key="j.weekKey">
                      <td>{{ formatWeekLabel(j.weekKey) }}</td>
                      <td>{{ j.count }}</td>
                    </tr>
                  </tbody>
                </table>
                <em v-else>No jobs found for this mechanism in the 12-week window.</em>
              </td>
            </tr>
          </template>
        </tbody>
      </table>
    </div>
  `,
};
