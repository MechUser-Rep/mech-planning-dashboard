// components/WarningsPanel.js
const WarningsPanel = {
  name: 'WarningsPanel',
  props: {
    warnings: { type: Array, required: true },
  },
  data() { return { expanded: false }; },
  template: `
    <div class="warnings-panel" v-if="warnings.length">
      <button class="warnings-toggle" @click="expanded = !expanded">
        ⚠ {{ warnings.length }} unmatched code{{ warnings.length > 1 ? 's' : '' }} — click to {{ expanded ? 'hide' : 'view' }}
      </button>
      <ul class="warnings-list" v-if="expanded">
        <li v-for="w in warnings" :key="w">{{ w }}</li>
      </ul>
    </div>
  `,
};
