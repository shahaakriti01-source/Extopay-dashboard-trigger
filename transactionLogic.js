// ─── TRANSACTION CLASSIFICATION & PAIRING LOGIC ─────────────────────────────
// Based on the 12 transaction types in the Payments Transaction sheet.

export const EXTO_BACKEND_ID = "rPezy2QDaNVDJ5LpABXpqAMiWPrricuKwW"; // "Exto Backend ID - Pay"
export const EXTO_BACKEND_ONBOARD_ID = "rHBFTaCH4Pcm3D5bSMH8nq6iibanr8vkKE"; // "Exto Backend ID - Onboard"
export const EXTO_BACKEND_NAME = "Exto Backend ID";

// Always-counted types (no pairing needed — always successful, standalone)
export const ALWAYS_COUNT_TYPES = new Set(["Online Remote", "Offline Final"]);

// Types that require pairing/reconciliation against a possible failure counterpart.
// Key = provisional type, Value = the failure type that, if matched, excludes the transaction.
export const PAIR_TYPES = {
  "Online Revokable": "Online Revoked",
  "Online Returnable": "Online Returned",
};

// Types never counted for processing/activation/milestone incentives at all
// (internal transfers, top-ups, cash-outs, and precursor/already-counted states)
export const EXCLUDED_TYPES = new Set([
  "Update Limit",
  "Payment",
  "Redeem",
  "Offline Revokable", // precursor only; resolves into Offline Final
  "Online Final", // precursor already counted as Online Revokable
]);

const MATCH_WINDOW_MINUTES = 15;

function parseDLTDate(dateStr) {
  // Expected format: "D/M/YYYY, H:MM:SS am/pm" or "D/M/YYYY"
  if (!dateStr) return null;
  const [datePart, timePart] = dateStr.split(",").map(s => s.trim());
  const [day, month, year] = datePart.split("/").map(Number);
  if (!timePart) return new Date(year, month - 1, day);
  const match = timePart.match(/(\d+):(\d+)(?::(\d+))?\s*(am|pm)?/i);
  if (!match) return new Date(year, month - 1, day);
  let [, h, m, s, ampm] = match;
  h = parseInt(h, 10);
  if (ampm && ampm.toLowerCase() === "pm" && h !== 12) h += 12;
  if (ampm && ampm.toLowerCase() === "am" && h === 12) h = 0;
  return new Date(year, month - 1, day, h, parseInt(m, 10), s ? parseInt(s, 10) : 0);
}

export function isExtoTransaction(p) {
  return p.account === EXTO_BACKEND_ID || p.destination === EXTO_BACKEND_ID ||
    p.sender === EXTO_BACKEND_NAME || p.receiver === EXTO_BACKEND_NAME;
}

/**
 * Returns the set of "qualifying" transactions for a merchant — i.e. transactions
 * that count toward transaction-processing incentive, activation bonus, and milestones.
 *
 * Logic:
 *  - Excludes anything touching the Exto Backend ID/account.
 *  - Always-count types (Online Remote, Offline Final) are included directly.
 *  - Paired types (Online Revokable, Online Returnable) are included UNLESS a matching
 *    failure transaction (same sender, same receiver, same exact amount, within 15 min after)
 *    is found — in which case BOTH are excluded entirely.
 *  - All other types (Update Limit, Payment, Redeem, Offline Revokable, Online Final) are excluded.
 */
export function getQualifyingTransactions(merchantName, allPayments) {
  const involved = allPayments.filter(
    p => (p.sender === merchantName || p.receiver === merchantName) && !isExtoTransaction(p)
  );

  const usedAsFailureMatch = new Set();
  const qualifying = [];

  involved.forEach((p, i) => {
    if (ALWAYS_COUNT_TYPES.has(p.txn_type)) {
      qualifying.push(p);
      return;
    }
    const failureType = PAIR_TYPES[p.txn_type];
    if (!failureType) return; // excluded type entirely

    if (usedAsFailureMatch.has(i)) return; // this row was itself consumed as someone else's match

    const pTime = parseDLTDate(p.dlt_close || p.created_at);
    const matchIndex = involved.findIndex((q, j) => {
      if (j === i || usedAsFailureMatch.has(j)) return false;
      if (q.txn_type !== failureType) return false;
      if (q.sender !== p.sender || q.receiver !== p.receiver) return false;
      if (q.amount !== p.amount) return false;
      const qTime = parseDLTDate(q.dlt_close || q.created_at);
      if (!pTime || !qTime) return false;
      const diffMinutes = (qTime - pTime) / 60000;
      return diffMinutes >= 0 && diffMinutes <= MATCH_WINDOW_MINUTES;
    });

    if (matchIndex !== -1) {
      usedAsFailureMatch.add(matchIndex);
      // this transaction failed — excluded entirely, not added to qualifying
    } else {
      qualifying.push(p);
    }
  });

  return qualifying;
}

/**
 * Audits EVERY payment transaction and returns each one tagged with whether it
 * counts toward any merchant's incentive, and a plain-language reason. Used by
 * the Transaction Audit tab for full transparency.
 *
 * `merchantNameSet` is a Set of all real merchant names (so we can tell whether
 * either side of a transaction is a real merchant).
 */
export function auditTransactions(allPayments, merchantNameSet) {
  // First, work out which paired transactions get cancelled by a matching failure,
  // globally (not per-merchant), so the audit reason is consistent.
  const failureMatched = new Set();
  allPayments.forEach((p, i) => {
    const failureType = PAIR_TYPES[p.txn_type];
    if (!failureType || failureMatched.has(i)) return;
    const pTime = parseDLTDate(p.dlt_close || p.created_at);
    const matchIndex = allPayments.findIndex((q, j) => {
      if (j === i || failureMatched.has(j)) return false;
      if (q.txn_type !== failureType) return false;
      if (q.sender !== p.sender || q.receiver !== p.receiver) return false;
      if (q.amount !== p.amount) return false;
      const qTime = parseDLTDate(q.dlt_close || q.created_at);
      if (!pTime || !qTime) return false;
      const diff = (qTime - pTime) / 60000;
      return diff >= 0 && diff <= MATCH_WINDOW_MINUTES;
    });
    if (matchIndex !== -1) {
      failureMatched.add(i);
      failureMatched.add(matchIndex);
    }
  });

  return allPayments.map((p, i) => {
    const senderIsMerchant = merchantNameSet.has(p.sender);
    const receiverIsMerchant = merchantNameSet.has(p.receiver);
    const merchant = senderIsMerchant ? p.sender : (receiverIsMerchant ? p.receiver : null);

    let counted = false;
    let reason = "";

    if (isExtoTransaction(p)) {
      reason = "Excluded — involves Exto Backend (internal/activation transfer)";
    } else if (!senderIsMerchant && !receiverIsMerchant) {
      reason = "Excluded — no real merchant on either side (e.g. coordinator-to-coordinator)";
    } else if (ALWAYS_COUNT_TYPES.has(p.txn_type)) {
      counted = true;
      reason = `Counted — "${p.txn_type}" always qualifies, for ${merchant}`;
    } else if (PAIR_TYPES[p.txn_type]) {
      if (failureMatched.has(i)) {
        reason = `Excluded — "${p.txn_type}" was reversed by a matching ${PAIR_TYPES[p.txn_type]}`;
      } else {
        counted = true;
        reason = `Counted — "${p.txn_type}" completed (no matching failure), for ${merchant}`;
      }
    } else if (["Online Revoked", "Online Returned"].includes(p.txn_type)) {
      reason = `Excluded — "${p.txn_type}" is a failure/reversal record`;
    } else {
      reason = `Excluded — "${p.txn_type}" is not an incentive-qualifying type`;
    }

    return { ...p, counted, reason, merchant };
  });
}
