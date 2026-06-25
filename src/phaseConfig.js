// ─── PHASE CONFIGURATION ─────────────────────────────────────────────────────
// All incentive variables are configurable per phase and persist via localStorage
// (this is a standalone deployed app, not a Claude artifact, so we use the browser's
// own storage rather than the artifact-only window.storage API).

export const DEFAULT_INCENTIVE_CONFIG = {
  consumerRate: 13,
  consumerCap: 50,
  activationBonus: 95,
  activationThreshold: 2,
  fixedRate: 1,
  variableRate: 0.5, // percent
  variableCap: 27,
  txnCap: 150,
  milestones: [
    { count: 50, bonus: 100 },
    { count: 100, bonus: 125 },
    { count: 150, bonus: 150 },
  ],
};

// Dashboard-tab targets, from the Exto Pay / PulaConnect KPI deck.
// Independent of the Incentives-tab config above.
export const DEFAULT_DASHBOARD_TARGETS = {
  0: { merchants: 20, consumers: 200, p2mTransactions: 500 },
  1: { merchants: 30, consumers: 300, p2mTransactions: 800 },
  2: { merchants: 150, consumers: 5000, p2mTransactions: 7500 },
  3: { merchants: 1000, consumers: 50000, p2mTransactions: 50000 },
};

const STORAGE_KEY_INCENTIVE = "exto_incentive_config_by_phase";
const STORAGE_KEY_TARGETS = "exto_dashboard_targets_by_phase";
const STORAGE_KEY_PAID_TRIGGERS = "exto_paid_triggers";

function safeGet(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (raw) return JSON.parse(raw);
  } catch (e) {
    console.error(`Failed to read ${key} from localStorage:`, e);
  }
  return fallback;
}

function safeSet(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch (e) {
    console.error(`Failed to write ${key} to localStorage:`, e);
  }
}

export function loadIncentiveConfigs() {
  return safeGet(STORAGE_KEY_INCENTIVE, {
    0: { ...DEFAULT_INCENTIVE_CONFIG },
    1: { ...DEFAULT_INCENTIVE_CONFIG },
    2: { ...DEFAULT_INCENTIVE_CONFIG },
    3: { ...DEFAULT_INCENTIVE_CONFIG },
  });
}

export function saveIncentiveConfigs(configs) {
  safeSet(STORAGE_KEY_INCENTIVE, configs);
}

export function loadDashboardTargets() {
  return safeGet(STORAGE_KEY_TARGETS, { ...DEFAULT_DASHBOARD_TARGETS });
}

export function saveDashboardTargets(targets) {
  safeSet(STORAGE_KEY_TARGETS, targets);
}

export function loadPaidTriggers() {
  const arr = safeGet(STORAGE_KEY_PAID_TRIGGERS, []);
  return new Set(arr);
}

export function savePaidTriggers(paidSet) {
  safeSet(STORAGE_KEY_PAID_TRIGGERS, [...paidSet]);
}
