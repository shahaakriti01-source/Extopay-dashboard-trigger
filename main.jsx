import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import { msalInstance } from './authConfig.js'

async function start() {
  // MSAL must be initialized before any other MSAL API is called.
  await msalInstance.initialize();

  // Consume any leftover redirect-response (e.g. "#code=..." in the URL) so it
  // never gets stuck and confuses subsequent sign-in attempts.
  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response) {
      console.log("MSAL: processed a redirect response for", response.account?.username);
    }
  } catch (err) {
    console.error("MSAL redirect handling error:", err);
  } finally {
    // Whether or not there was a response to process, strip any leftover
    // "#code=...", "#error=...", etc. fragment from the visible URL so it
    // can never be misread as a new/stuck interaction on a future load.
    if (window.location.hash && /^#(code|error|state)=/.test(window.location.hash)) {
      window.history.replaceState(null, "", window.location.pathname + window.location.search);
    }
  }

  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
}

start();
