import { PublicClientApplication } from "@azure/msal-browser";

// ── AZURE APP CONFIG ─────────────────────────────────────────────────────
// These values are NOT secret — they identify the app, not authenticate it.
export const msalConfig = {
  auth: {
    clientId: "682c009f-633d-4073-9f9e-d4bcd6cd54b0",
    authority: "https://login.microsoftonline.com/6e03983e-0e96-41f7-806b-bb21f023fb19",
    // MSAL v5 requires popup-based logins to redirect through a dedicated
    // "bridge" page (redirect.html) rather than the main app page itself.
    redirectUri: window.location.origin + "/redirect.html",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    // Default popup/iframe timeouts can be too short on slower network paths
    // (e.g. long round-trips to Microsoft's auth servers). Extend generously
    // so a real human entering a password doesn't trigger a false "timed_out".
    windowHashTimeout: 120000, // 2 minutes for the popup window
    iframeHashTimeout: 15000,
    loadFrameTimeout: 15000,
    asyncPopups: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Files.Read.All"],
};

export const msalInstance = new PublicClientApplication(msalConfig);
