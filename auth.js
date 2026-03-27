// auth.js
// Requires MSAL.js loaded from CDN before this script.

let _msalInstance = null;

function getMsalInstance() {
  if (_msalInstance) return _msalInstance;
  const msalConfig = {
    auth: {
      clientId:   CONFIG.clientId,
      authority:  `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: CONFIG.redirectUri,
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  };
  _msalInstance = new msal.PublicClientApplication(msalConfig);
  return _msalInstance;
}

async function signIn() {
  const instance = getMsalInstance();
  await instance.loginPopup({ scopes: CONFIG.scopes });
}

async function getToken() {
  const instance = getMsalInstance();
  const accounts = instance.getAllAccounts();
  if (accounts.length === 0) {
    await instance.loginPopup({ scopes: CONFIG.scopes });
  }
  const account = instance.getAllAccounts()[0];
  try {
    const result = await instance.acquireTokenSilent({ scopes: CONFIG.scopes, account });
    return result.accessToken;
  } catch {
    const result = await instance.acquireTokenPopup({ scopes: CONFIG.scopes, account });
    return result.accessToken;
  }
}

function getLoggedInUser() {
  const instance = getMsalInstance();
  const accounts = instance.getAllAccounts();
  return accounts.length > 0 ? accounts[0].name : null;
}

function signOut() {
  const instance = getMsalInstance();
  const account = instance.getAllAccounts()[0];
  if (account) instance.logoutPopup({ account });
}
