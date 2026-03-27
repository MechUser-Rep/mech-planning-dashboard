// config.js
// Fill in clientId and tenantId after IT registers the Azure AD app.
const CONFIG = {
  clientId:  'YOUR_CLIENT_ID',
  tenantId:  'YOUR_TENANT_ID',
  redirectUri: window.location.href.split('?')[0].split('#')[0],
  scopes: ['Files.Read.All'],
  sharePointHostname: 'reposefurniturelimited.sharepoint.com',
  sitePath: '/sites/reposefurniture-planningrepose',
  files: {
    lookup:     'Mech Forecast/Mechanism Codes/Mechanisms Lookup.xlsx',
    sortly:     'Mech Forecast/Sortly Reports/Latest Sortly Mech Report.xlsx',
    production: 'Mech Forecast/Production 2026 Dec-Nov.xlsx',
    poListing:  'Mech Forecast/Purchase Orders/PO Listing \u2013 with unit costs LATEST.xlsx',
    seminar:    'Mech Forecast/Seminar/Latest Repose Order Summary.xlsx',
  }
};
