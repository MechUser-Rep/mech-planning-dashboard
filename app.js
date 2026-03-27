// app.js
// Requires all other scripts loaded before this one.

const App = {
  components: { PlanningGrid, SuggestedOrders, WarningsPanel },

  data() {
    return {
      user:         null,
      loadingStep:  null,   // string like "Loading production plan (3/5)..."
      error:        null,
      lastRefresh:  null,
      showAtRisk:   false,
      rows:         [],
      weekKeys:     [],
      mechLookup:   null,
      warnings:     [],
      jobDemand:    {}, // { newCode: { weekKey: count } } — raw demand for drill-down
    };
  },

  async mounted() {
    await this.load();
  },

  methods: {
    async load() {
      this.error       = null;
      this.loadingStep = 'Signing in...';
      try {
        const token = await getToken();
        this.user   = getLoggedInUser();

        this.loadingStep = 'Loading mechanism codes (1/5)...';
        const lookupSheet = (await getWorksheetNames(token, CONFIG.files.lookup))[0];
        const lookupRows  = await getExcelUsedRange(token, CONFIG.files.lookup, lookupSheet);
        const mechLookup  = DataLoader.parseMechanismsLookup(lookupRows);

        this.loadingStep = 'Loading stock levels (2/5)...';
        const sortlySheet = (await getWorksheetNames(token, CONFIG.files.sortly))[0];
        const sortlyRows  = await getExcelUsedRange(token, CONFIG.files.sortly, sortlySheet);
        const stockMap   = DataLoader.parseSortlyStock(sortlyRows);

        // Merge stock: Sortly uses Entry Name = New Code
        const stockByNewCode = {};
        for (const [key, val] of Object.entries(stockMap)) {
          if (mechLookup.lookup[key]) stockByNewCode[key] = val;
        }

        this.loadingStep = 'Loading production plan (3/5)...';
        const sheetNames    = await getWorksheetNames(token, CONFIG.files.production);
        const weekKeys       = DataLoader.get12WeekKeys(new Date());
        const allSheets      = {};
        for (const name of sheetNames) {
          allSheets[name] = await getExcelUsedRange(token, CONFIG.files.production, name);
        }
        const { demand, unmatched: demandUnmatched } =
          DataLoader.parseProductionDemand(allSheets, weekKeys, mechLookup);

        this.loadingStep = 'Loading purchase orders (4/5)...';
        const poSheetName = await getWorksheetNames(token, CONFIG.files.poListing)
          .then(names => names[0]);
        const poRows = await getExcelUsedRange(token, CONFIG.files.poListing, poSheetName);
        const { incoming: poIncoming, unmatched: poUnmatched } =
          DataLoader.parsePOListingIncoming(poRows, weekKeys, mechLookup);

        this.loadingStep = 'Loading Seminar orders (5/5)...';
        const semRows = await getExcelUsedRange(token, CONFIG.files.seminar, 'Seminar open orders');
        const { incoming: semIncoming, unmatched: semUnmatched } =
          DataLoader.parseSeminarIncoming(semRows, weekKeys, mechLookup);

        const allIncoming = DataLoader.mergeIncoming(poIncoming, semIncoming);

        this.weekKeys   = weekKeys;
        this.mechLookup = mechLookup;
        this.jobDemand  = demand; // pass raw demand for row drill-down
        this.rows       = Calculations.buildMechRows(mechLookup, stockByNewCode, demand, allIncoming, weekKeys);
        this.warnings   = [...demandUnmatched, ...poUnmatched, ...semUnmatched];
        this.lastRefresh = new Date().toLocaleTimeString('en-GB');
        this.loadingStep = null;

      } catch (e) {
        this.error       = e.message;
        this.loadingStep = null;
        console.error(e);
      }
    },
  },

  template: `
    <header class="app-header">
      <h1>Mechanism Planning Dashboard</h1>
      <span class="user-info" v-if="user">{{ user }}</span>
      <span class="last-refresh" v-if="lastRefresh">Last refreshed: {{ lastRefresh }}</span>
      <button class="btn btn-toggle" @click="showAtRisk = !showAtRisk">
        {{ showAtRisk ? 'Show all' : 'Show at-risk only' }}
      </button>
      <button class="btn btn-refresh" :disabled="!!loadingStep" @click="load">
        {{ loadingStep ? 'Loading...' : 'Refresh' }}
      </button>
    </header>

    <div class="loading-bar" v-if="loadingStep">{{ loadingStep }}</div>

    <div class="error-bar" v-if="error" style="padding:0.5rem 1rem;background:#fee2e2;color:#991b1b;">
      Error: {{ error }}
    </div>

    <template v-if="!loadingStep && !error">
      <WarningsPanel :warnings="warnings" />
      <PlanningGrid
        :rows="rows"
        :week-keys="weekKeys"
        :show-at-risk="showAtRisk"
        :job-demand="jobDemand"
      />
      <SuggestedOrders
        :rows="rows"
        :week-keys="weekKeys"
        :mech-lookup="mechLookup || { lookup: {} }"
      />
    </template>
  `,
};
