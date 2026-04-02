// config.js
// Fill in clientId and tenantId after IT registers the Azure AD app.
const CONFIG = {
  clientId:  'YOUR_CLIENT_ID',   // TODO: fill in after IT registers Azure AD app
  tenantId:  'YOUR_TENANT_ID',   // TODO: fill in after IT registers Azure AD app
  redirectUri: 'http://localhost:8000',
  scopes: ['Files.Read.All'],
  sharePointHostname: 'reposefurniturelimited.sharepoint.com',
  sitePath: '/sites/ReposeFurniture-PlanningRepose',
  basePath: 'Shared Documents/Planning Repose/',
  files: {
    lookup:     'Mech Forecast/Mechanism Codes/Mechanisms Lookup.xlsx',
    sortly:     'Mech Forecast/Sortly Reports/Latest Sortly Mech Report.xlsx',
    production: 'Production 2026 Dec-Nov.xlsx',
    poListing:  'Mech Forecast/Purchase Orders/PO Listing - with unit costs LATEST.xlsx',
    seminar:    'Mech Forecast/Seminar/Latest Repose Order Summary.xlsx',
  }
};
