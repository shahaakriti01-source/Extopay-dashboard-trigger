import { msalInstance, loginRequest } from "./authConfig";

// ── OneDrive/SharePoint sharing URL for the live pilot Excel file ───────────
const SHARE_URL = "https://mergencompass-my.sharepoint.com/:x:/p/sandeep/IQBxf7rS-YmgRJG5clEQAPwgAcK13A5u8XVoU5QTRZB796c";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function encodeShareUrl(url) {
  const base64 = btoa(unescape(encodeURIComponent(url)))
    .replace(/=/g, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  return "u!" + base64;
}

// Never open a second popup from this module — MicrosoftSignIn (in App.jsx) is the
// ONLY place allowed to call loginPopup/acquireTokenPopup. If silent acquisition
// keeps failing here, it means the user truly isn't signed in yet, so we retry
// silently with short delays (covers the brief window right after a fresh login
// where MSAL's cache hasn't fully settled) rather than racing a second popup.
const SILENT_RETRY_ATTEMPTS = 5;
const SILENT_RETRY_DELAY_MS = 600;

async function getAccessToken() {
  for (let attempt = 0; attempt < SILENT_RETRY_ATTEMPTS; attempt++) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) {
      await new Promise(r => setTimeout(r, SILENT_RETRY_DELAY_MS));
      continue;
    }
    try {
      const response = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      return response.accessToken;
    } catch (err) {
      if (attempt === SILENT_RETRY_ATTEMPTS - 1) {
        throw new Error(
          "Could not get a Microsoft access token after signing in. Please click 'Refresh' to try again, or sign out and back in."
        );
      }
      await new Promise(r => setTimeout(r, SILENT_RETRY_DELAY_MS));
    }
  }
  throw new Error("No signed-in Microsoft account found.");
}

async function graphFetch(path, token) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph API error ${res.status}: ${body}`);
  }
  return res.json();
}

async function resolveShareLink(token) {
  const encoded = encodeShareUrl(SHARE_URL);
  const data = await graphFetch(`/shares/${encoded}/driveItem`, token);
  return { driveId: data.parentReference.driveId, itemId: data.id };
}

async function getSheetValues(driveId, itemId, sheetName, token) {
  const path = `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange(valuesOnly=true)`;
  const data = await graphFetch(path, token);
  return data.values || [];
}

// Finds the header row (the first row containing at least 2 of the expected
// column names) and returns { headerRowIndex, colIndex(name) }. This makes
// parsing resilient to the sheet's columns being reordered, inserted, or
// removed over time — we look up "Sender's name" etc. by NAME, not by a
// fixed numeric position, which previously broke silently when a column
// (e.g. "Created at") was removed from Payments Transaction.
function buildHeaderLookup(rows, expectedNames) {
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const row = rows[r] || [];
    const matches = row.filter(cell => expectedNames.includes(String(cell || "").trim())).length;
    if (matches >= 2) {
      const colIndex = {};
      row.forEach((cell, i) => {
        const name = String(cell || "").trim();
        if (name) colIndex[name] = i;
      });
      return { headerRowIndex: r, colIndex };
    }
  }
  return null; // caller should fall back to a sensible default or throw
}

function col(colIndex, ...candidateNames) {
  for (const name of candidateNames) {
    if (colIndex[name] !== undefined) return colIndex[name];
  }
  return -1;
}

// Builds an Account ID -> resolved Name lookup from the "Person name mapped to
// Acc ID" sheet — a clean two-column reference (Account ID -> Merchant/User name)
// that explicitly lists the 5 special accounts plus every real merchant. This is
// the authoritative source for names; the Onboarding sheet's own "Agent name"
// column is frequently "#N/A" so we resolve names ourselves.
function buildAccountIdToNameMap(mappingRows) {
  const header = buildHeaderLookup(mappingRows, ["Account ID", "Merchant/User name"]);
  const accountCol = header ? col(header.colIndex, "Account ID", "Account") : 1;
  const nameCol = header ? col(header.colIndex, "Merchant/User name") : 2;
  const startRow = header ? header.headerRowIndex + 1 : 4;

  const map = new Map();
  mappingRows.slice(startRow).forEach(row => {
    const accountId = row[accountCol] ? String(row[accountCol]).trim() : null;
    const name = row[nameCol] ? String(row[nameCol]).trim() : null;
    if (accountId && name && name !== "#N/A") {
      map.set(accountId, name);
    }
  });
  return map;
}

// ── MAIN EXPORT: fetch and parse all sheets we need ──────────────────────
export async function fetchLivePilotData() {
  const token = await getAccessToken();
  const { driveId, itemId } = await resolveShareLink(token);

  const [mappingRows, onboardingRows, paymentsRows] = await Promise.all([
    getSheetValues(driveId, itemId, "Person name mapped to Acc ID", token),
    getSheetValues(driveId, itemId, "Accounts Trans - Onboarding(DD)", token),
    getSheetValues(driveId, itemId, "Payments Transaction", token),
  ]);

  const accountIdToName = buildAccountIdToNameMap(mappingRows);

  // ── Onboarding(DD) — resolve columns by header name ──
  const onbHeader = buildHeaderLookup(onboardingRows, [
    "Transaction ID", "Owner", "Destination", "Authorize", "Submission", "DLT Close Time", "User name", "Agent name",
  ]);
  const onbCol = onbHeader ? onbHeader.colIndex : {};
  const onbStartRow = onbHeader ? onbHeader.headerRowIndex + 1 : 4;
  const c_txnId = col(onbCol, "Transaction ID");
  const c_owner = col(onbCol, "Owner");
  const c_destination = col(onbCol, "Destination");
  const c_authorize = col(onbCol, "Authorize");
  const c_submission = col(onbCol, "Submission");
  const c_dltClose = col(onbCol, "DLT Close Time");
  const c_userName = col(onbCol, "User name");

  const onboardingEvents = [];
  onboardingRows.slice(onbStartRow).forEach(row => {
    if (!row[c_txnId] || row[c_txnId] === "Transaction ID") return;
    if (row[c_submission] !== "tesSUCCESS") return; // only successful submissions count
    let authorizeId = row[c_authorize] ? String(row[c_authorize]).trim() : "";
    // The sheet sometimes shows a placeholder dash ("—") in the Authorize column
    // instead of a real ID or leaving it blank. Treat that as "no authorize ID" —
    // these rows represent a self/system account-creation, not a real onboarding
    // of one person by another, so they should not be counted at all.
    if (authorizeId === "—" || authorizeId === "-" || authorizeId === "0") authorizeId = "";
    if (!authorizeId) return; // no real onboarder — exclude entirely, per spec

    // Resolve the Authorize column's name ourselves from the mapping sheet —
    // the sheet's own "Agent name" column is frequently "#N/A" and can't be
    // trusted, especially exactly when Exto/a coordinator is the authorizer.
    const resolvedAgentName = accountIdToName.get(authorizeId) || null;

    onboardingEvents.push({
      txn_id: String(row[c_txnId]).slice(0, 16),
      owner: row[c_owner] ? String(row[c_owner]) : "",
      destination: row[c_destination] ? String(row[c_destination]) : "",
      authorize: authorizeId,
      dlt_close: row[c_dltClose] ? String(row[c_dltClose]) : "",
      user_name: row[c_userName] ? String(row[c_userName]) : "Unknown",
      agent_name: resolvedAgentName,
    });
  });

  // ── Payments Transaction — resolve columns by header name ──
  const payHeader = buildHeaderLookup(paymentsRows, [
    "Transaction ID", "Transaction type", "Account", "Destination", "Amount", "Submission",
    "DLT Close Time", "Amount (BWP)", "Sender's name", "Receiver's name",
  ]);
  const payCol = payHeader ? payHeader.colIndex : {};
  const payStartRow = payHeader ? payHeader.headerRowIndex + 1 : 3;
  const c_pTxnId = col(payCol, "Transaction ID");
  const c_pType = col(payCol, "Transaction type", "Transaction Type");
  const c_pAccount = col(payCol, "Account");
  const c_pDestination = col(payCol, "Destination");
  const c_pSubmission = col(payCol, "Submission");
  const c_pDltClose = col(payCol, "DLT Close Time");
  const c_pCreatedAt = col(payCol, "Created at"); // may not exist — col() returns -1, handled below
  const c_pAmountBWP = col(payCol, "Amount (BWP)", "Amount(BWP)");
  const c_pSender = col(payCol, "Sender's name", "Sender name");
  const c_pReceiver = col(payCol, "Receiver's name", "Receiver name");

  const payments = [];
  paymentsRows.slice(payStartRow).forEach(row => {
    if (!row[c_pTxnId] || row[c_pTxnId] === "Transaction ID") return;
    // NOTE: unlike onboarding, payments are NOT filtered by Submission status —
    // every payment row is considered regardless of its submission tag.
    const amtRaw = row[c_pAmountBWP] !== undefined && row[c_pAmountBWP] !== null ? String(row[c_pAmountBWP]).trim() : "0";
    const amt = parseFloat(amtRaw.replace("BWP", "").trim()) || 0;
    const sender = row[c_pSender] ? String(row[c_pSender]).trim() : "";
    const receiver = row[c_pReceiver] ? String(row[c_pReceiver]).trim() : "";

    payments.push({
      txn_id: String(row[c_pTxnId]).slice(0, 16),
      txn_type: row[c_pType] ? String(row[c_pType]).trim() : "",
      account: row[c_pAccount] ? String(row[c_pAccount]) : "",
      destination: row[c_pDestination] ? String(row[c_pDestination]) : "",
      amount: Math.round(amt * 100) / 100,
      dlt_close: row[c_pDltClose] ? String(row[c_pDltClose]) : "",
      created_at: c_pCreatedAt >= 0 && row[c_pCreatedAt] ? String(row[c_pCreatedAt]) : "",
      sender: sender && sender !== "#N/A" ? sender : "",
      receiver: receiver && receiver !== "#N/A" ? receiver : "",
    });
  });

  return { onboardingEvents, payments, fetchedAt: new Date().toISOString() };
}
