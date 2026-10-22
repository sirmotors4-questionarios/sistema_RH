import {RH_CONFIG} from './rh-config.js';

// Autenticação Microsoft Entra (MSAL). A biblioteca é carregada na página como global `msal`.
let app = null;
let ready = null;

function lib() {
  const m = globalThis.msal;
  if (!m || !m.PublicClientApplication) throw Error('A biblioteca de autenticação Microsoft não foi carregada.');
  return m;
}

export function authConfigured() {
  return Boolean(RH_CONFIG.clientId);
}

export function initAuth() {
  if (!authConfigured()) return Promise.resolve(null);
  if (!ready) {
    ready = (async () => {
      const m = lib();
      app = new m.PublicClientApplication({
        auth: {
          clientId: RH_CONFIG.clientId,
          authority: 'https://login.microsoftonline.com/' + RH_CONFIG.tenantId,
          redirectUri: location.origin + location.pathname
        },
        cache: {cacheLocation: 'sessionStorage'}
      });
      await app.initialize();
      const result = await app.handleRedirectPromise();
      const account = result?.account || app.getAllAccounts()[0];
      if (account) app.setActiveAccount(account);
      return app;
    })();
  }
  return ready;
}

export function currentAccount() {
  return app ? app.getActiveAccount() : null;
}

export async function signIn() {
  const client = await initAuth();
  if (!client) throw Error('A autenticação ainda não foi configurada (clientId em rh-config.js).');
  await client.loginRedirect({scopes: RH_CONFIG.scopes, prompt: 'select_account'});
}

export async function signOut() {
  const client = await initAuth();
  if (client) await client.logoutRedirect({account: client.getActiveAccount()});
}

export async function getToken() {
  const client = await initAuth();
  const account = client?.getActiveAccount();
  if (!account) {
    const error = Error('Inicie sessão com a sua conta Microsoft da SIR.');
    error.code = 'RH_AUTH';
    throw error;
  }
  try {
    return (await client.acquireTokenSilent({scopes: RH_CONFIG.scopes, account})).accessToken;
  } catch (error) {
    if (error instanceof lib().InteractionRequiredAuthError) {
      await client.acquireTokenRedirect({scopes: RH_CONFIG.scopes, account});
    }
    throw error;
  }
}
