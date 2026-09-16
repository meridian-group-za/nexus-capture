// Shared MSAL + Graph setup for login.html and index.html, mirroring the Call
// Cycle Coverage portal so there is one auth pattern across the Meridian
// dashboards. Reuses the same Entra ID app registration -- the only thing that
// needs adding in Entra is this dashboard's redirect URI on that existing app.
const MSAL_CONFIG = {
  clientId: "f7e6dfb5-798c-4feb-8d7c-d1f6c53dc37f",
  tenantId: "cbe83df2-b350-4dab-8d5f-f78d21fe7d27", // Meridian Group tenant
  // Must exactly match a redirect URI registered on the app above.
  redirectUri: location.origin + location.pathname.replace(/[^/]*$/, "") + "login.html",
};

// Files.ReadWrite is already consented on this app registration (the Call Cycle
// portal relies on it), so reusing it avoids a fresh admin-consent round trip.
const GRAPH_SCOPES = ["Files.ReadWrite"];

// Where the dashboard's data lives in SharePoint. This is the same library the
// Capture folder syncs from -- resolved from the OneDrive sync mount point
// C:\Users\...\OneDrive - Meridian Group\Meridian Nexus - Documents.
const SHAREPOINT_HOSTNAME = "meridiangroupza.sharepoint.com";
const SHAREPOINT_SITE_PATH = "sites/MeridianNexus";
// Folder inside "Shared Documents" that extract_compliance_data.py writes to.
const DATA_FOLDER = "Capture/Capture/Dashboard";

// ---------------------------------------------------------------------------
// AUTH_DISABLED: set true only to review the UI without signing in. While true
// the dashboard reads compliance_data.json.gz from next to index.html instead
// of pulling it from SharePoint with the user's token -- which means the data
// would have to be published with the page. Never deploy it true.
//
// False (live): every visitor signs in with their Meridian account and the
// payload is fetched from SharePoint with their own token, so no store-level
// data sits in the repo. localhost still bypasses sign-in for development.
// ---------------------------------------------------------------------------
const AUTH_DISABLED = false;

// Served from localhost/127.0.0.1 -> skip real auth and read the local JSON,
// exactly as the Call Cycle portal does. Only a deployed origin requires a real
// Microsoft sign-in and pulls data from SharePoint.
function isLocalDev() {
  return AUTH_DISABLED || ["localhost", "127.0.0.1", ""].includes(location.hostname);
}

let msalInstance = null;
function getMsalInstance() {
  if (!msalInstance) {
    if (typeof msal === "undefined") throw new Error("MSAL library did not load");
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: MSAL_CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + MSAL_CONFIG.tenantId,
        redirectUri: MSAL_CONFIG.redirectUri,
      },
      // localStorage, not sessionStorage: the portal opens every dashboard in a
  // new tab, and sessionStorage is empty in a new tab by definition - so the
  // sign-in never carried over and each dashboard asked again.
      cache: { cacheLocation: "localStorage" },
    });
  }
  return msalInstance;
}

function getAccount() {
  if (isLocalDev()) return { name: "Local dev", username: "localhost" };
  try { return getMsalInstance().getAllAccounts()[0] || null; }
  catch (e) { return null; }
}

async function signIn() {
  const inst = getMsalInstance();
  await inst.handleRedirectPromise();
  return inst.loginRedirect({ scopes: GRAPH_SCOPES });
}

async function signOut() {
  // With auth off there is no session to end, so the honest behaviour is to
  // return to the login screen rather than pretend a sign-out happened.
  if (AUTH_DISABLED) {
    location.href = location.pathname.replace(/[^/]*$/, "") + "login.html";
    return;
  }
  if (isLocalDev()) { location.reload(); return; }
  const inst = getMsalInstance();
  const account = inst.getAllAccounts()[0];
  return inst.logoutRedirect({ account: account, postLogoutRedirectUri: MSAL_CONFIG.redirectUri });
}

// Gate for index.html. Resolves once there is a signed-in account; otherwise
// sends the browser to login.html and never resolves (the page is leaving).
async function requireAuth() {
  if (isLocalDev()) return getAccount();
  const inst = getMsalInstance();
  await inst.handleRedirectPromise();
  let account = inst.getAllAccounts()[0];

  // Nothing cached for THIS origin does not mean the person is signed
  // out. MSAL caches per app registration and per origin, so someone who
  // signed into another Meridian dashboard a minute ago still lands here
  // with an empty cache and gets a login screen they do not need.
  //
  // ssoSilent asks Entra to reuse the session it already has, in a hidden
  // iframe, with no prompt. It fails in ordinary circumstances -- no
  // session, more than one account, or a browser blocking third-party
  // cookies -- so the login page stays exactly as the fallback rather
  // than being replaced by it.
  if (!account) {
    try {
      const sso = await inst.ssoSilent({ scopes: GRAPH_SCOPES });
      if (sso && sso.account) account = sso.account;
    } catch (e) {
      account = null;
    }
  }

  if (!account) {
    location.replace(location.pathname.replace(/[^/]*$/, "") + "login.html");
    return new Promise(function () {});
  }
  return account;
}

async function getGraphToken() {
  const inst = getMsalInstance();
  const account = inst.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  try {
    const result = await inst.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: account });
    return result.accessToken;
  } catch (e) {
    await inst.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account: account });
    return null; // page is redirecting; nothing after this runs
  }
}

// Site id is stable for the session, so resolve it once rather than on every
// file read (the per-form question files mean many reads per session).
let _siteId = null;
async function getSiteId(token) {
  if (_siteId) return _siteId;
  const res = await fetch(
    "https://graph.microsoft.com/v1.0/sites/" + SHAREPOINT_HOSTNAME + ":/" + SHAREPOINT_SITE_PATH,
    { headers: { Authorization: "Bearer " + token } }
  );
  if (!res.ok) throw new Error("Could not resolve SharePoint site (status " + res.status + ")");
  _siteId = (await res.json()).id;
  return _siteId;
}

// Payloads are stored gzipped (.json.gz): the main file is 17.9MB raw but
// 1.3MB gzipped, and neither python's http.server nor SharePoint/Graph applies
// transport compression to a .json response -- so compressing the file itself
// is the only thing that helps on both paths. Inflate with DecompressionStream,
// which needs no library.
async function inflateJson(response, label) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("This browser cannot decompress the data file (needs a current Edge, Chrome, Firefox or Safari).");
  }
  const stream = response.body.pipeThrough(new DecompressionStream("gzip"));
  const text = await new Response(stream).text();
  try { return JSON.parse(text); }
  catch (e) { throw new Error("Malformed data in " + label); }
}

// Single entry point the dashboard uses. Local dev reads the file sitting next
// to index.html; a deployed origin pulls the same file from SharePoint with the
// signed-in user's token, so the data is never public alongside the page.
// `filename` is the logical name -- ".gz" is appended here so callers do not
// have to know how the file is stored.
async function loadDataFile(filename, optional) {
  const stored = filename + ".gz";
  if (isLocalDev()) {
    const res = await fetch(stored + "?t=" + Date.now());
    if (res.status === 404 && optional) return null;
    if (!res.ok) throw new Error("HTTP " + res.status + " loading " + stored);
    return inflateJson(res, stored);
  }
  const token = await getGraphToken();
  if (!token) return null;
  const siteId = await getSiteId(token);
  const path = encodeURI(DATA_FOLDER + "/" + stored);
  const res = await fetch(
    "https://graph.microsoft.com/v1.0/sites/" + siteId + "/drive/root:/" + path + ":/content",
    { headers: { Authorization: "Bearer " + token } }
  );
  if (res.status === 404 && optional) return null;
  if (!res.ok) throw new Error("Could not load " + stored + " from SharePoint (status " + res.status + ")");
  return inflateJson(res, stored);
}
