(() => {
  "use strict";

  // LocalStorage key for app-created state, settings, backups, and visible snapshots.
  const APP_STORAGE_KEY = "AM_CRM_APP_STATE_V2";

  // Human-readable app version stored in backups and LLM exports.
  const APP_VERSION = "1.0.0-local-command-center";

  // Optional ISO date override for repeatable calculations; blank means use the browser's current date.
  const CURRENT_DATE_OVERRIDE = "";

  // Number of days after a workbook's last modified timestamp before Settings flags it as stale.
  const FILE_STALE_DAYS = 14;

  // Warning shown when the sidebar refresh button is clicked before any source folder has been granted.
  const SOURCE_FOLDER_NOT_GRANTED_WARNING = "No source folder has been granted yet. Use Settings > Data Management > Grant Folder Access once.";

  // Warning prefix shown when a typed source folder path exists but the browser has no saved folder handle for it.
  const CONFIGURED_SOURCE_FOLDER_NOT_GRANTED_PREFIX = "Cannot refresh ";

  // Warning suffix shown after the configured source folder path when browser folder access is missing.
  const CONFIGURED_SOURCE_FOLDER_NOT_GRANTED_SUFFIX = " until folder access has been granted once in Settings > Data Management.";

  // Warning shown when Chrome denies access to the previously saved source folder handle.
  const SOURCE_FOLDER_PERMISSION_DENIED_WARNING = "Browser permission for the saved source folder was not granted.";

  // Minimum character count before the global search filters the Home grid.
  const SEARCH_MIN_CHARS = 3;

  // Default maximum number of recent activity log entries shown in Settings.
  const ACTIVITY_LOG_VISIBLE_LIMIT = 300;

  // Default global font stack used everywhere in the app unless changed in Settings.
  const DEFAULT_GLOBAL_FONT_FAMILY = 'Consolas, "Cascadia Mono", "Segoe UI", Arial, sans-serif';

  // Default app-wide font size in pixels for dense table, popup, and settings text.
  const DEFAULT_GLOBAL_FONT_SIZE_PX = 12;

  // Default smaller app-wide font size in pixels for footer, timestamp, and SVG relationship labels.
  const DEFAULT_SMALL_FONT_SIZE_PX = 11;

  // Default number of Outlook messages exported per client email address.
  const DEFAULT_OUTLOOK_MAX_EMAILS_PER_ADDRESS = 20;

  // Default number of calendar days the Outlook helper scans backward for matching email history.
  const DEFAULT_OUTLOOK_LOOKBACK_DAYS = 365;

  // Default maximum number of characters saved from each Outlook email body preview.
  const DEFAULT_OUTLOOK_BODY_PREVIEW_CHARS = 750;

  // Default downloaded JSON filename for the client email list passed into the Outlook VBA helper.
  const DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME = "AM_CRM_Outlook_Client_Emails.json";

  // Default imported JSON filename expected from the helper recent-email export.
  const DEFAULT_OUTLOOK_RECENT_EMAILS_FILENAME = "AM_CRM_Outlook_Recent_Emails.json";

  // Default imported JSON filename expected from the helper sent-email log export.
  const DEFAULT_OUTLOOK_SENT_LOG_FILENAME = "AM_CRM_Outlook_Sent_Email_Log.json";

  // Default imported JSON filename expected from the helper email-to-task export.
  const DEFAULT_OUTLOOK_TASK_CANDIDATE_FILENAME = "AM_CRM_Outlook_Task_Candidates.json";

  // Default prefix used when this app downloads a draft JSON file for the Outlook VBA helper.
  const DEFAULT_OUTLOOK_DRAFT_FILE_PREFIX = "AM_CRM_Outlook_Draft";

  // Default contact type label used when imported sent emails are counted as actual client contact.
  const DEFAULT_OUTLOOK_SENT_CONTACT_TYPE = "Sent email";

  // IndexedDB database name used only to remember the browser-granted source folder handle.
  const FOLDER_HANDLE_DB_NAME = "AM_CRM_FOLDER_HANDLES";

  // IndexedDB object store name for local File System Access API handles.
  const FOLDER_HANDLE_STORE_NAME = "handles";

  // IndexedDB key for the primary source workbook folder handle.
  const SOURCE_FOLDER_HANDLE_KEY = "sourceWorkbookFolder";

  // Default one-based worksheet row that contains column headers when a workbook has no specific override.
  const DEFAULT_WORKBOOK_HEADER_ROW = 1;

  // One-based worksheet header rows for known workbooks that include title/filter rows before the real table.
  const DEFAULT_WORKBOOK_HEADER_ROW_OVERRIDES = {
    "ref_account.xlsx": 6,
    "data_activity.xlsx": 6
  };

  // Required and optional workbook files that the local import engine understands.
  const EXPECTED_WORKBOOKS = [
    { fileName: "ref_ALG.xlsx", tableName: "ref_ALG", required: true },
    { fileName: "ref_ACCOUNT.xlsx", tableName: "ref_ACCOUNT", required: true },
    { fileName: "data_POSITIONS.xlsx", tableName: "data_POSITIONS", required: true },
    { fileName: "data_ACTIVITY.xlsx", tableName: "data_ACTIVITY", required: true },
    { fileName: "data_REVENUE.xlsx", tableName: "data_REVENUE", required: true },
    { fileName: "ref_PRODUCT.xlsx", tableName: "ref_PRODUCT", required: true },
    { fileName: "ref_SECURITY.xlsx", tableName: "ref_SECURITY", required: true },
    { fileName: "ref_WEALTHDESK.xlsx", tableName: "ref_WEALTHDESK", required: true },
    { fileName: "data_GPS.xlsx", tableName: "data_GPS", required: false },
    { fileName: "data_CLIENTCONTACT.xlsx", tableName: "data_CLIENTCONTACT", required: false },
    { fileName: "data_POSITIONS_EXTERNAL.xlsx", tableName: "data_POSITIONS_EXTERNAL", required: false },
    { fileName: "ref_ALLOCATION_TARGET.xlsx", tableName: "ref_ALLOCATION_TARGET", required: false }
  ];

  // Home grid column definitions, labels, default widths, and sorting hints.
  const COLUMN_DEFS = [
    { key: "status", label: "Status", width: 64, sticky: 1, align: "center" },
    { key: "clientName", label: "Client Name", width: 220, sticky: 2 },
    { key: "email", label: "Email", width: 230 },
    { key: "nextContact", label: "Next Contact", width: 104 },
    { key: "t12Revenue", label: "T12 Revenue", width: 116, numeric: true },
    { key: "nna", label: "NNA", width: 96, numeric: true },
    { key: "availableCash", label: "Available Cash", width: 138, numeric: true },
    { key: "totalAssets", label: "Total Assets", width: 122, numeric: true },
    { key: "assetAllocation", label: "Asset Allocation", width: 126 },
    { key: "financialPlan", label: "Financial Plan", width: 120 },
    { key: "onlineActivity", label: "Online Activity", width: 118, align: "center" },
    { key: "tasks", label: "Tasks", width: 74, numeric: true },
    { key: "notes", label: "Notes", width: 70, align: "center" }
  ];

  // Default sidebar and grid utility icons used when no PNG override is uploaded.
  const DEFAULT_APP_ICONS = {
    sidebar_settings: "&#9881;",
    sidebar_home: "&#8962;",
    sidebar_refresh: "&#8635;",
    sidebar_backup: "&#8681;",
    sidebar_export: "&#8681;",
    sidebar_add_task: "&#9745;",
    sidebar_add_note: "&#9998;",
    grid_email: "",
    grid_notes: ""
  };

  // Default top-to-bottom order for the draggable left sidebar buttons.
  const DEFAULT_SIDEBAR_BUTTON_ORDER = [
    "sidebar_settings",
    "sidebar_refresh",
    "sidebar_backup",
    "sidebar_export",
    "sidebar_add_task",
    "sidebar_add_note",
    "sidebar_home"
  ];

  // DOM button ids keyed by the persisted sidebar button order values.
  const SIDEBAR_BUTTON_IDS = {
    sidebar_settings: "btnSettingsNav",
    sidebar_home: "btnHomeNav",
    sidebar_refresh: "btnRefresh",
    sidebar_backup: "btnBackup",
    sidebar_export: "btnExport",
    sidebar_add_task: "btnAddTask",
    sidebar_add_note: "btnAddNote"
  };

  // Icon override labels shown in Settings > Icons.
  const APP_ICON_LABELS = [
    { key: "sidebar_settings", label: "Sidebar Settings" },
    { key: "sidebar_home", label: "Sidebar Home" },
    { key: "sidebar_refresh", label: "Sidebar Refresh Data" },
    { key: "sidebar_backup", label: "Sidebar Backup" },
    { key: "sidebar_export", label: "Sidebar Export Grid" },
    { key: "sidebar_add_task", label: "Sidebar Add Task" },
    { key: "sidebar_add_note", label: "Sidebar Add Note" },
    { key: "grid_email", label: "Main Table Email" },
    { key: "grid_notes", label: "Main Table Notes" }
  ];

  // Settings submenu order for the only non-Home app view.
  const SETTINGS_SECTIONS = [
    "Main Table",
    "Data Management",
    "Status",
    "Client Profile",
    "Email",
    "VBA Outlook Helper",
    "Next Contact",
    "T12 Revenue",
    "NNA",
    "Available Cash",
    "Total Assets",
    "Advisory Assets",
    "Asset Allocation",
    "Financial Plan",
    "Online Activity",
    "Tasks",
    "Notes",
    "Icons",
    "Backups",
    "Activity Log",
    "LLM Export / Import"
  ];

  // Valid status tiers in best-to-worst order.
  const STATUS_ORDER = ["A", "B", "C", "D"];

  // Color names used by CSS for status dots when no local PNG override is configured.
  const STATUS_LABELS = {
    A: "Platinum",
    B: "Gold",
    C: "Silver",
    D: "Bronze"
  };

  // Default status tier rules. Settings can override the numeric thresholds.
  const STATUS_TIER_RULES = {
    A: { revenue: 30000, assets: 5000000 },
    B: { revenue: 15000, assets: 2500000 },
    C: { revenue: 5000, assets: 1000000 },
    D: { revenue: 0, assets: 0 }
  };

  // Default profile cadence values driven by displayed status.
  const STATUS_PROFILE_DEFAULTS = {
    A: { contactFrequency: 12, numberOfReviews: 2 },
    B: { contactFrequency: 6, numberOfReviews: 2 },
    C: { contactFrequency: 1, numberOfReviews: 0 },
    D: { contactFrequency: 0, numberOfReviews: 0 }
  };

  // Contact frequency code to day cadence mapping used for Next Contact.
  const DEFAULT_CONTACT_CADENCE_DAYS = {
    0: 0,
    1: 365,
    4: 91,
    6: 61,
    12: 30
  };

  // Profile choices shown in Client Profile popups and Settings.
  const RISK_PROFILES = [
    "1 - Wealth Conservation",
    "2 - Income Oriented",
    "3 - Balanced Growth",
    "4 - Market Growth",
    "5 - Opportunistic Growth"
  ];

  // Month choices used by Client Profile review scheduling.
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  // Contact types that count as interactions for Next Contact.
  const DEFAULT_CONTACT_TYPES = [
    "Sent email",
    "Sent text message",
    "Zoom meeting",
    "In-person meeting",
    "Logged call",
    "Voicemail"
  ];

  // Task statuses managed entirely inside this app.
  const DEFAULT_TASK_STATUSES = ["Open", "In Progress", "Waiting", "Completed", "Deferred", "Cancelled"];

  // Task priorities managed entirely inside this app.
  const DEFAULT_TASK_PRIORITIES = ["Low", "Normal", "High", "Urgent"];

  // Default email templates and trigger names used by the Email popup.
  const DEFAULT_EMAIL_TEMPLATES = [
    {
      id: "template_overdue_contact",
      name: "Overdue contact",
      trigger: "overdue_contact",
      subject: "Quick check-in",
      body: "Hi {{salutation}},\n\nI wanted to check in and find a good time for us to connect.\n\nBest,\nAdam"
    },
    {
      id: "template_excess_cash",
      name: "Excess cash",
      trigger: "excess_cash",
      subject: "Reviewing available cash",
      body: "Hi {{salutation}},\n\nI noticed available cash that may be worth reviewing. Let me know a good time to talk through options.\n\nBest,\nAdam"
    },
    {
      id: "template_upcoming_review",
      name: "Upcoming review",
      trigger: "upcoming_review",
      subject: "Upcoming review",
      body: "Hi {{salutation}},\n\nI wanted to schedule our next review and update any planning priorities.\n\nBest,\nAdam"
    },
    {
      id: "template_default",
      name: "Default",
      trigger: "default",
      subject: "Following up",
      body: "Hi {{salutation}},\n\nI wanted to follow up with you.\n\nBest,\nAdam"
    }
  ];

  // Default app settings. Every setting here is included in JSON backups.
  const DEFAULT_SETTINGS = {
    popupMode: "floating",
    cornerRadiusPx: 0,
    columnHeaderBgColor: "#2d2d30",
    sidebarButtonOrder: DEFAULT_SIDEBAR_BUTTON_ORDER,
    sidebarButtonVisibility: DEFAULT_SIDEBAR_BUTTON_ORDER.reduce((visibility, key) => ({ ...visibility, [key]: true }), {}),
    globalFontFamily: DEFAULT_GLOBAL_FONT_FAMILY,
    globalFontSizePx: DEFAULT_GLOBAL_FONT_SIZE_PX,
    smallFontSizePx: DEFAULT_SMALL_FONT_SIZE_PX,
    showTooltips: true,
    showFieldSource: false,
    includeHeldAwayCash: true,
    autoBackups: false,
    sourceFolderPath: "",
    autoRefreshSavedFolder: true,
    revenueAmountField: "Amt Cmp Grs",
    revenuePivotRowField: "GROUP",
    revenuePivotColumnField: "Month",
    nnaPivotRowField: "Activity Type",
    nnaPivotColumnField: "Month",
    nnaExcludedKeywords: ["dividend"],
    defaultMinAvailableCash: 0,
    defaultMaxAvailableCash: 100000,
    cashKeywords: ["cash", "money market", "savings deposit", "bdps", "bank deposit"],
    totalAssetDisabledCategories: [],
    driftThresholds: { Cash: 5, "Fixed Income": 5, Equity: 5, Alternatives: 5, UNKNOWN: 0 },
    statusTierRules: STATUS_TIER_RULES,
    statusProfileDefaults: STATUS_PROFILE_DEFAULTS,
    contactCadenceDays: DEFAULT_CONTACT_CADENCE_DAYS,
    contactTypes: DEFAULT_CONTACT_TYPES,
    outlookMaxEmailsPerAddress: DEFAULT_OUTLOOK_MAX_EMAILS_PER_ADDRESS,
    outlookLookbackDays: DEFAULT_OUTLOOK_LOOKBACK_DAYS,
    outlookBodyPreviewChars: DEFAULT_OUTLOOK_BODY_PREVIEW_CHARS,
    outlookClientExportFilename: DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME,
    outlookRecentEmailsFilename: DEFAULT_OUTLOOK_RECENT_EMAILS_FILENAME,
    outlookSentLogFilename: DEFAULT_OUTLOOK_SENT_LOG_FILENAME,
    outlookTaskCandidateFilename: DEFAULT_OUTLOOK_TASK_CANDIDATE_FILENAME,
    outlookDraftFilePrefix: DEFAULT_OUTLOOK_DRAFT_FILE_PREFIX,
    outlookSentContactType: DEFAULT_OUTLOOK_SENT_CONTACT_TYPE,
    taskStatuses: DEFAULT_TASK_STATUSES,
    taskPriorities: DEFAULT_TASK_PRIORITIES,
    defaultRiskProfile: "3 - Balanced Growth",
    noteShowArchived: false,
    recurrencePlaceholderEnabled: true,
    financialPlanMapping: {
      clientGroupId: "",
      likelihood: "",
      status: "",
      lastUpdateDate: "",
      nextReviewDate: ""
    },
    onlineActivityMapping: {
      tableName: "ref_ACCOUNT",
      clientGroupId: "Sort Name",
      accountNumber: "Account Number",
      username: "",
      lastLoginDate: ""
    },
    allocationTargetMapping: {
      profile: "Risk Profile",
      includeAlts: "Include Alts",
      includeMunis: "Include Munis",
      assetClass: "Asset Class",
      targetPercent: "Target Percent"
    },
    assetClassMappings: {}
  };

  // Fallback target allocations used until ref_ALLOCATION_TARGET.xlsx is imported and mapped.
  const DEFAULT_TARGET_ALLOCATIONS = {
    "1 - Wealth Conservation|No|No": { Cash: 5, "Fixed Income": 70, Equity: 25, Alternatives: 0 },
    "2 - Income Oriented|No|No": { Cash: 4, "Fixed Income": 60, Equity: 36, Alternatives: 0 },
    "3 - Balanced Growth|No|No": { Cash: 3, "Fixed Income": 40, Equity: 57, Alternatives: 0 },
    "4 - Market Growth|No|No": { Cash: 2, "Fixed Income": 25, Equity: 73, Alternatives: 0 },
    "5 - Opportunistic Growth|No|No": { Cash: 2, "Fixed Income": 15, Equity: 83, Alternatives: 0 },
    "1 - Wealth Conservation|Yes|No": { Cash: 5, "Fixed Income": 65, Equity: 20, Alternatives: 10 },
    "2 - Income Oriented|Yes|No": { Cash: 4, "Fixed Income": 55, Equity: 31, Alternatives: 10 },
    "3 - Balanced Growth|Yes|No": { Cash: 3, "Fixed Income": 35, Equity: 52, Alternatives: 10 },
    "4 - Market Growth|Yes|No": { Cash: 2, "Fixed Income": 20, Equity: 68, Alternatives: 10 },
    "5 - Opportunistic Growth|Yes|No": { Cash: 2, "Fixed Income": 10, Equity: 78, Alternatives: 10 }
  };

  const schemaCatalog = window.AM_CRM_SCHEMA_CATALOG || { tables: {}, metadata: {}, relationshipMapping: { relationships: [] } };
  const seedClients = window.AM_CRM_SEED_CLIENTS || [];

  let appState = loadState();
  let runtime = emptyRuntime();
  let gridRows = [];
  let activeSort = { key: "t12Revenue", direction: "desc" };
  let activeSettingsSection = "Main Table";
  let activePopup = null;
  let autoRefreshAttempted = false;

  function emptyRuntime() {
    return {
      importedTables: {},
      importedHeaders: {},
      fileStatuses: [],
      warnings: [],
      validation: [],
      aggregates: null,
      loadedAt: null,
      sourceMode: null
    };
  }

  function defaultState() {
    return {
      appVersion: APP_VERSION,
      lastRefreshAt: null,
      restoredAt: null,
      settings: clone(DEFAULT_SETTINGS),
      tableLayout: {
        columns: COLUMN_DEFS.map(col => ({ key: col.key, width: col.width }))
      },
      dataManagement: {
        sourceFolderName: null,
        expectedFiles: normalizeExpectedFiles(),
        joins: defaultRelationshipMappings(),
        userJoinCounter: 1
      },
      clientOverrides: {},
      manualClients: [],
      notes: [],
      tasks: [],
      activityLog: [],
      outlookEmailHistory: [],
      outlookImportHistory: [],
      emailTemplates: clone(DEFAULT_EMAIL_TEMPLATES),
      iconOverrides: {},
      snapshots: {
        mainGridRows: [],
        popupSummaries: {},
        fileStatuses: [],
        validation: [],
        importedTableHeaders: {},
        generatedAt: null
      },
      priorRefreshSnapshot: null
    };
  }

  function defaultRelationshipMappings() {
    const rels = schemaCatalog.relationshipMapping?.relationships || [];
    return rels.map(rel => ({ ...rel, approved: true, enabled: true }));
  }

  function defaultHeaderRowForFile(fileName) {
    return DEFAULT_WORKBOOK_HEADER_ROW_OVERRIDES[String(fileName || "").toLowerCase()] || DEFAULT_WORKBOOK_HEADER_ROW;
  }

  function tableNameFromFileName(fileName) {
    return cleanText(fileName).replace(/\.xlsx$/i, "");
  }

  function normalizeWorkbookConfig(item, fallback = {}) {
    const fileName = cleanText(item?.fileName || fallback.fileName || "");
    const defaultHeaderRow = defaultHeaderRowForFile(fileName);
    return {
      fileName,
      tableName: cleanText(item?.tableName || fallback.tableName || tableNameFromFileName(fileName)),
      required: Boolean(item?.required ?? fallback.required ?? false),
      headerRow: Math.round(clampNumber(item?.headerRow ?? fallback.headerRow ?? defaultHeaderRow, 1, 500, defaultHeaderRow)),
      builtIn: Boolean(item?.builtIn ?? fallback.builtIn ?? false)
    };
  }

  function normalizeExpectedFiles(files = []) {
    const saved = Array.isArray(files) ? files : [];
    const byFileName = new Map(saved.map(item => [cleanText(item.fileName).toLowerCase(), item]));
    const usedFileNames = new Set();
    const builtIns = EXPECTED_WORKBOOKS.map(expected => {
      const defaultItem = normalizeWorkbookConfig({
        ...expected,
        headerRow: defaultHeaderRowForFile(expected.fileName),
        builtIn: true
      });
      const savedItem = byFileName.get(expected.fileName.toLowerCase());
      usedFileNames.add(expected.fileName.toLowerCase());
      return normalizeWorkbookConfig({
        ...defaultItem,
        ...(savedItem || {}),
        fileName: expected.fileName,
        tableName: expected.tableName,
        builtIn: true
      }, defaultItem);
    });
    const custom = saved
      .filter(item => {
        const fileKey = cleanText(item.fileName).toLowerCase();
        return fileKey && !usedFileNames.has(fileKey);
      })
      .map(item => normalizeWorkbookConfig(item, { required: false, builtIn: false }))
      .filter(item => item.fileName && item.tableName);
    return [...builtIns, ...custom];
  }

  function ensureExpectedFileConfig() {
    appState.dataManagement.expectedFiles = normalizeExpectedFiles(appState.dataManagement.expectedFiles);
    return appState.dataManagement.expectedFiles;
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(APP_STORAGE_KEY);
      if (!raw) return defaultState();
      return mergeState(defaultState(), JSON.parse(raw));
    } catch (err) {
      console.warn("AM_CRM state load failed.", err);
      return defaultState();
    }
  }

  function mergeState(base, override) {
    const merged = { ...base, ...(override || {}) };
    merged.settings = deepMerge(base.settings, override?.settings || {});
    merged.tableLayout = deepMerge(base.tableLayout, override?.tableLayout || {});
    merged.dataManagement = deepMerge(base.dataManagement, override?.dataManagement || {});
    merged.snapshots = deepMerge(base.snapshots, override?.snapshots || {});
    for (const key of ["clientOverrides", "manualClients", "notes", "tasks", "activityLog", "outlookEmailHistory", "outlookImportHistory", "emailTemplates", "iconOverrides"]) {
      merged[key] = Array.isArray(base[key]) ? (override?.[key] || base[key]) : { ...(base[key] || {}), ...(override?.[key] || {}) };
    }
    if (!merged.emailTemplates.length) merged.emailTemplates = clone(DEFAULT_EMAIL_TEMPLATES);
    if (!merged.dataManagement.joins?.length) merged.dataManagement.joins = defaultRelationshipMappings();
    merged.dataManagement.expectedFiles = normalizeExpectedFiles(merged.dataManagement.expectedFiles);
    merged.appVersion = APP_VERSION;
    return merged;
  }

  function deepMerge(base, override) {
    const out = Array.isArray(base) ? [...base] : { ...base };
    Object.entries(override || {}).forEach(([key, value]) => {
      if (value && typeof value === "object" && !Array.isArray(value) && base && typeof base[key] === "object" && !Array.isArray(base[key])) {
        out[key] = deepMerge(base[key], value);
      } else {
        out[key] = value;
      }
    });
    return out;
  }

  function saveState(activity) {
    if (activity) addActivity(activity, false);
    localStorage.setItem(APP_STORAGE_KEY, JSON.stringify(appState));
  }

  function uid(prefix = "id") {
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
  }

  function todayDate() {
    const date = CURRENT_DATE_OVERRIDE ? new Date(`${CURRENT_DATE_OVERRIDE}T00:00:00`) : new Date();
    date.setHours(0, 0, 0, 0);
    return date;
  }

  function normalizeTextKey(value) {
    return String(value ?? "").trim().toUpperCase();
  }

  function leftAccountKey(value) {
    return String(value ?? "").trim().slice(0, 10);
  }

  function buildAccountKey(branchOrOffice, accountNumber) {
    const branch = digitsOnly(branchOrOffice).padStart(3, "0").slice(-3);
    const account = digitsOnly(accountNumber).padStart(6, "0").slice(-6);
    return `${branch}-${account}`;
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const cleaned = String(value).replace(/\(([^)]+)\)/, "-$1").replace(/[$,%\s,]/g, "");
    if (!cleaned || cleaned === "-" || cleaned.toUpperCase() === "N/A") return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function toDate(value) {
    if (!value && value !== 0) return null;
    if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
    if (typeof value === "number" && Number.isFinite(value)) return excelSerialToDate(value);
    const raw = String(value).trim();
    if (!raw || raw.toUpperCase() === "N/A") return null;
    if (/^\d+(\.\d+)?$/.test(raw)) return excelSerialToDate(Number(raw));
    const mdy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (mdy) {
      const year = Number(mdy[3].length === 2 ? `20${mdy[3]}` : mdy[3]);
      const date = new Date(year, Number(mdy[1]) - 1, Number(mdy[2]));
      return Number.isNaN(date.valueOf()) ? null : date;
    }
    const parsed = new Date(raw);
    return Number.isNaN(parsed.valueOf()) ? null : parsed;
  }

  function excelSerialToDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    if (Number.isNaN(dateInfo.valueOf())) return null;
    return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
  }

  function digitsOnly(value) {
    return String(value ?? "").trim().replace(/\D/g, "");
  }

  function cleanText(value) {
    return String(value ?? "").trim();
  }

  function safeText(value) {
    const text = cleanText(value);
    return text ? text : "N/A";
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function stripTags(value) {
    return String(value ?? "").replace(/<[^>]*>/g, "");
  }

  function formatMoney(value, scale = "auto") {
    const n = toNumber(value);
    if (n === null) return "N/A";
    if (scale === "k") return `$${(n / 1000).toFixed(1)}K`;
    if (scale === "mm") return `$${(n / 1000000).toFixed(1)}MM`;
    return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);
  }

  function formatPercent(value, digits = 1) {
    const n = toNumber(value);
    if (n === null) return "N/A";
    return `${n.toFixed(digits)}%`;
  }

  function formatDate(value) {
    const date = toDate(value);
    return date ? date.toISOString().slice(0, 10) : "N/A";
  }

  function clampNumber(value, min, max, fallback) {
    const n = Number(value);
    if (!Number.isFinite(n)) return fallback;
    return Math.min(max, Math.max(min, n));
  }

  function sanitizeFontFamily(value) {
    return String(value || DEFAULT_GLOBAL_FONT_FAMILY).replace(/[;\n\r{}]/g, "").trim() || DEFAULT_GLOBAL_FONT_FAMILY;
  }

  function sanitizeCssColor(value, fallback) {
    const color = String(value || "").trim();
    return /^#[0-9a-fA-F]{6}$/.test(color) ? color : fallback;
  }

  function splitEmails(value) {
    if (Array.isArray(value)) return value.map(cleanText).filter(Boolean);
    return String(value ?? "")
      .split(/[;,]/)
      .map(cleanText)
      .filter(Boolean);
  }

  function isYes(value) {
    return ["Y", "YES", "TRUE", "1"].includes(normalizeTextKey(value));
  }

  function render() {
    document.documentElement.style.setProperty("--radius", `${Number(appState.settings.cornerRadiusPx || 0)}px`);
    document.documentElement.style.setProperty("--grid-header-bg", sanitizeCssColor(appState.settings.columnHeaderBgColor, "#2d2d30"));
    document.documentElement.style.setProperty("--font", sanitizeFontFamily(appState.settings.globalFontFamily));
    document.documentElement.style.setProperty("--font-size", `${clampNumber(appState.settings.globalFontSizePx, 9, 24, DEFAULT_GLOBAL_FONT_SIZE_PX)}px`);
    document.documentElement.style.setProperty("--font-size-small", `${clampNumber(appState.settings.smallFontSizePx, 8, 20, DEFAULT_SMALL_FONT_SIZE_PX)}px`);
    renderSidebarButtons();
    renderAppIcons();
    gridRows = buildGridRows();
    snapshotVisibleState();
    renderGrid();
    renderLastRefresh();
    if (isSettingsVisible()) renderSettings();
  }

  function isSettingsVisible() {
    return document.getElementById("settingsView")?.classList.contains("active");
  }

  function snapshotVisibleState() {
    appState.snapshots.mainGridRows = gridRows.map(row => rowToSnapshot(row));
    appState.snapshots.popupSummaries = {};
    gridRows.forEach(row => {
      appState.snapshots.popupSummaries[row.id] = row.summaries || {};
    });
    appState.snapshots.fileStatuses = runtime.fileStatuses.length ? clone(runtime.fileStatuses) : appState.snapshots.fileStatuses;
    appState.snapshots.validation = runtime.validation.length ? clone(runtime.validation) : appState.snapshots.validation;
    appState.snapshots.importedTableHeaders = runtime.importedHeaders && Object.keys(runtime.importedHeaders).length ? clone(runtime.importedHeaders) : appState.snapshots.importedTableHeaders;
    appState.snapshots.generatedAt = new Date().toISOString();
  }

  function rowToSnapshot(row) {
    return {
      id: row.id,
      clientName: row.clientName,
      status: row.status,
      calculatedStatus: row.calculatedStatus,
      email: row.email,
      emails: row.emails,
      phone: row.phone,
      salutation: row.salutation,
      reviewLocation: row.reviewLocation,
      source: row.source,
      lastContactDate: row.lastContactDate,
      nextContact: row.nextContact,
      contactFrequencyCode: row.contactFrequencyCode,
      reviewFrequencyCode: row.reviewFrequencyCode,
      reviewMonth: row.reviewMonth,
      riskProfile: row.riskProfile,
      includeAlts: row.includeAlts,
      includeMunis: row.includeMunis,
      totalNetWorth: row.totalNetWorth,
      annualIncome: row.annualIncome,
      minAvailableCash: row.minAvailableCash,
      maxAvailableCash: row.maxAvailableCash,
      t12Revenue: row.t12Revenue,
      nna: row.nna,
      availableCash: row.availableCash,
      availableCashRaw: row.availableCashRaw,
      availableCashPercent: row.availableCashPercent,
      availableCashAlert: row.availableCashAlert,
      totalAssets: row.totalAssets,
      advisoryAssets: row.advisoryAssets,
      advisoryPercent: row.advisoryPercent,
      assetAllocation: row.assetAllocation,
      financialPlan: row.financialPlan,
      onlineActivity: row.onlineActivity,
      tasks: row.tasks,
      notes: row.notes,
      alerts: row.alerts,
      summaries: row.summaries
    };
  }

  function buildGridRows() {
    if (hasRuntimeTables()) return buildImportedGridRows();
    if (appState.snapshots.mainGridRows?.length) return rowsFromSnapshot();
    return buildSeedGridRows();
  }

  function hasRuntimeTables() {
    return Object.values(runtime.importedTables || {}).some(rows => Array.isArray(rows) && rows.length);
  }

  function rowsFromSnapshot() {
    return appState.snapshots.mainGridRows.map(snapshot => {
      const row = mergeStatefulCounts({ ...snapshot });
      row.nextContact = normalizeNextContact(row.nextContact, row.lastContactDate, row.contactFrequencyCode);
      row.summaries = row.summaries || {};
      row.alerts = row.alerts || [];
      return row;
    });
  }

  function buildSeedGridRows() {
    return seedClients.concat(appState.manualClients || []).map(client => {
      const id = client.idNameAlg || client.clientGroupId || client.id;
      const status = validStatus(client.status) || "N/A";
      const profile = profileForClient(id, status, client, [], null);
      const lastContact = getLastContactDate(id, client.lastReviewRaw);
      const row = {
        id,
        clientName: id,
        status,
        calculatedStatus: "D",
        email: (client.emails || splitEmails(client.emailRaw)).join("; "),
        emails: client.emails || splitEmails(client.emailRaw),
        phone: client.phone || "",
        salutation: client.salutation || "",
        reviewLocation: client.reviewLocation || "",
        source: client.seedRowNumber ? "seed" : "manual",
        lastContactDate: lastContact ? lastContact.toISOString() : null,
        contactFrequencyCode: profile.contactFrequency,
        reviewFrequencyCode: profile.numberOfReviews,
        reviewMonth: profile.reviewMonth,
        riskProfile: profile.riskProfile,
        includeAlts: profile.includeAlts,
        includeMunis: profile.includeMunis,
        totalNetWorth: null,
        annualIncome: null,
        minAvailableCash: profile.minAvailableCash,
        maxAvailableCash: profile.maxAvailableCash,
        t12Revenue: null,
        nna: null,
        availableCash: null,
        availableCashRaw: null,
        availableCashPercent: null,
        availableCashAlert: false,
        totalAssets: null,
        advisoryAssets: null,
        advisoryPercent: null,
        assetAllocation: "N/A",
        financialPlan: "N/A",
        onlineActivity: null,
        alerts: [],
        summaries: defaultSummaries()
      };
      row.nextContact = computeNextContact(lastContact, profile.contactFrequency);
      row.calculatedStatus = calculateStatusTier(row).tier;
      return mergeStatefulCounts(row);
    });
  }

  function buildImportedGridRows() {
    runtime.aggregates = computeAggregates();
    runtime.validation = validateRelationships();
    const aggregates = runtime.aggregates;
    const sources = buildClientSources(aggregates);
    return sources.map(source => {
      const clientKey = normalizeTextKey(source.id);
      const accountSummary = aggregates.accountSummaryByClient.get(clientKey) || {};
      const revenue = aggregates.revenueByClient.get(clientKey) || null;
      const nna = aggregates.nnaByClient.get(clientKey) || null;
      const positions = aggregates.positionByClient.get(clientKey) || null;
      const financial = aggregates.financialByClient.get(clientKey) || null;
      const online = aggregates.onlineByClient.get(clientKey) || null;
      const status = statusForSource(source, accountSummary);
      const profile = profileForClient(source.id, status, source.seed || source.alg || {}, accountSummary.accounts || [], source.alg);
      const lastContact = getLastContactDate(source.id, source.seed?.lastReviewRaw || source.alg?.["LAST REVIEW"] || null);
      const totalAssets = positions?.totalAssets ?? null;
      const row = {
        id: source.id,
        clientName: source.id,
        status,
        calculatedStatus: "D",
        email: source.emails.join("; "),
        emails: source.emails,
        phone: source.phone,
        salutation: source.salutation,
        reviewLocation: source.reviewLocation,
        source: source.source,
        lastContactDate: lastContact ? lastContact.toISOString() : null,
        contactFrequencyCode: profile.contactFrequency,
        reviewFrequencyCode: profile.numberOfReviews,
        reviewMonth: profile.reviewMonth,
        riskProfile: profile.riskProfile,
        includeAlts: profile.includeAlts,
        includeMunis: profile.includeMunis,
        totalNetWorth: accountSummary.totalNetWorth ?? null,
        annualIncome: accountSummary.annualIncome ?? null,
        minAvailableCash: profile.minAvailableCash,
        maxAvailableCash: profile.maxAvailableCash,
        t12Revenue: revenue?.total ?? null,
        nna: nna?.total ?? null,
        availableCash: positions?.availableCashExcess ?? null,
        availableCashRaw: positions?.availableCashRaw ?? null,
        availableCashPercent: positions?.availableCashPercent ?? null,
        availableCashAlert: positions?.availableCashAlert ?? false,
        totalAssets,
        advisoryAssets: positions?.advisoryAssets ?? null,
        advisoryPercent: positions?.advisoryPercent ?? null,
        assetAllocation: "N/A",
        financialPlan: financial?.display ?? "N/A",
        onlineActivity: online || null,
        alerts: [],
        summaries: {
          accounts: accountSummary,
          revenue: revenue || emptyRevenueSummary(),
          nna: nna || emptyNnaSummary(),
          positions: positions || emptyPositionSummary(),
          financial: financial || {},
          online: online || {}
        }
      };
      row.nextContact = computeNextContact(lastContact, profile.contactFrequency);
      row.calculatedStatus = calculateStatusTier(row).tier;
      row.status = applyParentStatusOverride(row, row.calculatedStatus, row.status);
      row.assetAllocation = computeAllocationStatus(row, positions);
      row.alerts = buildRowAlerts(row);
      return mergeStatefulCounts(row);
    });
  }

  function defaultSummaries() {
    return {
      accounts: {},
      revenue: emptyRevenueSummary(),
      nna: emptyNnaSummary(),
      positions: emptyPositionSummary(),
      financial: {},
      online: {}
    };
  }

  function emptyRevenueSummary() {
    return { total: null, rows: [], pivot: {}, byProduct: {}, byMonth: {} };
  }

  function emptyNnaSummary() {
    return { total: null, rows: [], pivot: {}, byType: {}, byMonth: {} };
  }

  function emptyPositionSummary() {
    return {
      totalAssets: null,
      msAssets: null,
      externalAssets: null,
      advisoryAssets: null,
      advisoryPercent: null,
      availableCashRaw: null,
      availableCashExcess: null,
      availableCashPercent: null,
      availableCashAlert: false,
      cashBreakdown: {},
      maturityList: [],
      strategyBreakdown: {},
      custodianBreakdown: {},
      categoryBreakdown: {},
      allocationRows: [],
      singleStockAlerts: []
    };
  }

  function mergeStatefulCounts(row) {
    row.tasks = appState.tasks.filter(task => task.clientGroupId === row.id && !task.archivedDate && !["Completed", "Cancelled"].includes(task.status)).length;
    row.notes = appState.notes.filter(note => note.clientGroupId === row.id && !note.archivedDate).length;
    return row;
  }

  function buildClientSources(aggregates) {
    const byKey = new Map();
    const algRows = runtime.importedTables.ref_ALG || [];
    algRows.forEach(alg => {
      const id = cleanText(alg.ID_NAME_ALG);
      if (!id) return;
      byKey.set(normalizeTextKey(id), {
        id,
        source: "ref_ALG",
        alg,
        seed: null,
        emails: splitEmails(alg.EMAIL),
        phone: cleanText(alg.PHONE),
        salutation: cleanText(alg.SALUTATION),
        reviewLocation: cleanText(alg["REVIEW LOCATION"])
      });
    });
    if (algRows.length) {
      seedClients.forEach(seed => {
        const existing = byKey.get(normalizeTextKey(seed.idNameAlg));
        if (!existing) return;
        existing.seed = seed;
        if (!existing.emails.length) existing.emails = seed.emails || splitEmails(seed.emailRaw);
        existing.phone ||= seed.phone || "";
        existing.salutation ||= seed.salutation || "";
        existing.reviewLocation ||= seed.reviewLocation || "";
      });
      return [...byKey.values()].sort((a, b) => a.id.localeCompare(b.id));
    }
    seedClients.forEach(seed => {
      const id = seed.idNameAlg;
      const key = normalizeTextKey(id);
      if (byKey.has(key)) {
        const existing = byKey.get(key);
        existing.seed = seed;
        if (!existing.emails.length) existing.emails = seed.emails || splitEmails(seed.emailRaw);
        existing.phone ||= seed.phone || "";
        existing.salutation ||= seed.salutation || "";
        existing.reviewLocation ||= seed.reviewLocation || "";
        return;
      }
      byKey.set(key, {
        id,
        source: "seed",
        alg: null,
        seed,
        emails: seed.emails || splitEmails(seed.emailRaw),
        phone: seed.phone || "",
        salutation: seed.salutation || "",
        reviewLocation: seed.reviewLocation || ""
      });
    });
    (appState.manualClients || []).forEach(manual => {
      const id = manual.idNameAlg || manual.clientGroupId || manual.id;
      const key = normalizeTextKey(id);
      if (!id || byKey.has(key)) return;
      byKey.set(key, {
        id,
        source: "manual",
        alg: null,
        seed: manual,
        emails: manual.emails || splitEmails(manual.emailRaw),
        phone: manual.phone || "",
        salutation: manual.salutation || "",
        reviewLocation: manual.reviewLocation || ""
      });
    });
    aggregates.clientKeys.forEach((display, key) => {
      if (byKey.has(key)) return;
      byKey.set(key, {
        id: display,
        source: "imported_account",
        alg: null,
        seed: null,
        emails: [],
        phone: "",
        salutation: "",
        reviewLocation: ""
      });
    });
    return [...byKey.values()].sort((a, b) => a.id.localeCompare(b.id));
  }

  function statusForSource(source, accountSummary) {
    const algStatus = validStatus(source.alg?.STATUS);
    if (algStatus) return algStatus;
    const seedStatus = validStatus(source.seed?.status);
    if (seedStatus) return seedStatus;
    const accountStatus = bestStatus(accountSummary.statuses || []);
    return accountStatus || "N/A";
  }

  function validStatus(value) {
    const status = normalizeTextKey(value);
    return STATUS_ORDER.includes(status) ? status : null;
  }

  function bestStatus(statuses) {
    let best = null;
    statuses.forEach(status => {
      const s = validStatus(status);
      if (!s) return;
      if (!best || STATUS_ORDER.indexOf(s) < STATUS_ORDER.indexOf(best)) best = s;
    });
    return best;
  }

  function profileForClient(clientId, status, source, accounts, alg) {
    const override = appState.clientOverrides[clientId] || {};
    const statusDefaults = appState.settings.statusProfileDefaults?.[validStatus(status) || "D"] || STATUS_PROFILE_DEFAULTS.D;
    const algContactFrequency = parseContactFrequency(alg?.["CONTACT FREQUENCY"]);
    const algReviewFrequency = parseReviewFrequency(alg?.["REVIEW FREQUENCY"]);
    const accountReviewMonth = firstNonBlank(accounts.map(a => a["Review Month"]));
    return {
      parentClientGroupId: override.parentClientGroupId || "",
      riskProfile: override.riskProfile || normalizeRiskProfile(alg?.["GIC ALLOCATION"] || source.gicAllocation) || appState.settings.defaultRiskProfile,
      includeAlts: override.includeAlts || "No",
      includeMunis: override.includeMunis || "No",
      minAvailableCash: toNumber(override.minAvailableCash) ?? Number(appState.settings.defaultMinAvailableCash || 0),
      maxAvailableCash: toNumber(override.maxAvailableCash) ?? Number(appState.settings.defaultMaxAvailableCash || 100000),
      numberOfReviews: toNumber(override.numberOfReviews) ?? algReviewFrequency ?? source.reviewFrequencyCode ?? statusDefaults.numberOfReviews,
      reviewMonth: override.reviewMonth || accountReviewMonth || "",
      contactFrequency: toNumber(override.contactFrequency) ?? algContactFrequency ?? source.contactFrequencyCode ?? statusDefaults.contactFrequency,
      relationshipNotes: override.relationshipNotes || "",
      familyManager: override.familyManager || ""
    };
  }

  function normalizeRiskProfile(value) {
    const text = normalizeTextKey(value);
    if (!text) return "";
    const model = text.match(/MODEL\s*([1-5])/);
    if (model) return RISK_PROFILES[Number(model[1]) - 1];
    return RISK_PROFILES.find(profile => normalizeTextKey(profile) === text) || "";
  }

  function parseContactFrequency(value) {
    const text = normalizeTextKey(value);
    if (!text) return null;
    if (text.includes("MONTHLY")) return 12;
    if (text.includes("BI-MONTHLY") || text.includes("BIMONTHLY")) return 6;
    if (text.includes("QUARTER")) return 4;
    if (text.includes("ANNUAL") || text.includes("YEAR")) return 1;
    if (text.includes("NONE") || text.includes("REQUEST")) return 0;
    return toNumber(value);
  }

  function parseReviewFrequency(value) {
    const text = normalizeTextKey(value);
    if (!text) return null;
    if (text.includes("SEMI")) return 2;
    if (text.includes("ANNUAL")) return 1;
    if (text.includes("NONE")) return 0;
    return toNumber(value);
  }

  function firstNonBlank(values) {
    return values.map(cleanText).find(Boolean) || "";
  }

  function getLastContactDate(clientId, fallbackDate) {
    const appContacts = appState.activityLog
      .filter(entry => entry.clientGroupId === clientId && ["contact_log", "sent_email_log"].includes(entry.type))
      .map(entry => toDate(entry.detail?.contactDate || entry.timestamp))
      .filter(Boolean)
      .sort((a, b) => b - a);
    if (appContacts.length) return appContacts[0];

    const imported = getImportedClientContactDate(clientId);
    if (imported) return imported;
    return toDate(fallbackDate);
  }

  function getImportedClientContactDate(clientId) {
    const rows = runtime.importedTables.data_CLIENTCONTACT || [];
    if (!rows.length) return null;
    const clientKey = normalizeTextKey(clientId);
    const dateFields = detectDateFields(rows, ["contact", "date"]);
    const clientFields = detectClientFields(rows);
    const dates = rows
      .filter(row => clientFields.some(field => normalizeTextKey(row[field]) === clientKey))
      .flatMap(row => dateFields.map(field => toDate(row[field])).filter(Boolean))
      .sort((a, b) => b - a);
    return dates[0] || null;
  }

  function normalizeNextContact(value, lastContactDate, contactFrequency) {
    if (value && typeof value === "object" && value.display) return value;
    return computeNextContact(toDate(lastContactDate), contactFrequency);
  }

  function computeNextContact(lastContactDate, contactFrequency) {
    const freq = Number(contactFrequency || 0);
    if (!freq) return { display: "N/A", days: null, dueDate: null, className: "cellMuted" };
    const cadenceDays = Number(appState.settings.contactCadenceDays?.[freq] ?? DEFAULT_CONTACT_CADENCE_DAYS[freq] ?? 365);
    const basis = lastContactDate || todayDate();
    const due = new Date(basis.getTime());
    due.setDate(due.getDate() + cadenceDays);
    due.setHours(0, 0, 0, 0);
    const days = Math.round((due - todayDate()) / 86400000);
    return {
      display: days === 0 ? "Due" : days > 0 ? `+${days}d` : `${days}d`,
      days,
      dueDate: due.toISOString().slice(0, 10),
      className: days < 0 ? "cellDanger" : days === 0 ? "cellWarn" : ""
    };
  }

  function calculateStatusTier(row) {
    const rules = appState.settings.statusTierRules || STATUS_TIER_RULES;
    const revenue = Number(row.t12Revenue || 0);
    const assets = Number(row.totalAssets || 0);
    for (const tier of ["A", "B", "C"]) {
      const rule = rules[tier] || STATUS_TIER_RULES[tier];
      if (revenue >= Number(rule.revenue || 0) || assets >= Number(rule.assets || 0)) {
        return { tier, reason: `T12 Revenue >= ${formatMoney(rule.revenue)} or Total Assets >= ${formatMoney(rule.assets)}` };
      }
    }
    return { tier: "D", reason: "Revenue and assets are below C-tier thresholds." };
  }

  function applyParentStatusOverride(row, calculatedStatus, displayedStatus) {
    const override = appState.clientOverrides[row.id] || {};
    const parentId = override.parentClientGroupId;
    if (!parentId) return displayedStatus;
    const parent = gridRows.find(r => normalizeTextKey(r.id) === normalizeTextKey(parentId));
    if (!parent) return displayedStatus;
    const parentIndex = STATUS_ORDER.indexOf(validStatus(parent.calculatedStatus || parent.status) || "D");
    const childFloor = STATUS_ORDER[Math.min(parentIndex + 1, STATUS_ORDER.length - 1)];
    const standardIndex = STATUS_ORDER.indexOf(validStatus(calculatedStatus) || "D");
    const floorIndex = STATUS_ORDER.indexOf(childFloor);
    return STATUS_ORDER[Math.min(standardIndex, floorIndex)];
  }

  function buildRowAlerts(row) {
    const alerts = [];
    if (row.nextContact?.days !== null && row.nextContact.days < 0) alerts.push({ columnKey: "nextContact", text: `Contact overdue by ${Math.abs(row.nextContact.days)} days.` });
    if (row.availableCashAlert) alerts.push({ columnKey: "availableCash", text: "Available cash is outside profile thresholds." });
    const positions = row.summaries?.positions || {};
    (positions.maturityList || []).filter(item => item.days <= 10).forEach(item => alerts.push({ columnKey: "availableCash", text: `${item.securityName} matures in ${item.days} days.` }));
    (positions.singleStockAlerts || []).forEach(item => alerts.push({ columnKey: "assetAllocation", text: `${item.securityName} is ${formatPercent(item.percent)} of assets.` }));
    if (row.assetAllocation === "DRIFT") alerts.push({ columnKey: "assetAllocation", text: "Asset allocation drift exceeds configured threshold." });
    return alerts;
  }

  function computeAggregates() {
    const tables = runtime.importedTables;
    const aggregates = {
      clientKeys: new Map(),
      accountByKey: new Map(),
      accountSummaryByClient: new Map(),
      productByCode: new Map(),
      securityById: new Map(),
      wealthDeskByAccount: new Map(),
      revenueByClient: new Map(),
      nnaByClient: new Map(),
      positionByClient: new Map(),
      financialByClient: new Map(),
      onlineByClient: new Map()
    };
    indexProducts(aggregates, tables.ref_PRODUCT || []);
    indexSecurities(aggregates, tables.ref_SECURITY || []);
    indexWealthDesk(aggregates, tables.ref_WEALTHDESK || []);
    indexAccounts(aggregates, tables.ref_ACCOUNT || []);
    aggregateRevenue(aggregates, tables.data_REVENUE || []);
    aggregateNna(aggregates, tables.data_ACTIVITY || []);
    aggregatePositions(aggregates, tables.data_POSITIONS || [], false);
    if (appState.settings.includeHeldAwayCash) aggregatePositions(aggregates, tables.data_POSITIONS_EXTERNAL || [], true);
    aggregateFinancialPlans(aggregates, tables.data_GPS || []);
    aggregateOnlineActivity(aggregates);
    finalizePositionAggregates(aggregates);
    return aggregates;
  }

  function indexProducts(aggregates, rows) {
    rows.forEach(row => {
      const key = cleanText(row["Cde Prd"]);
      if (key) aggregates.productByCode.set(key, row);
    });
  }

  function indexSecurities(aggregates, rows) {
    rows.forEach(row => {
      const key = cleanText(row.ID_MSDW_SECURITY);
      if (key) aggregates.securityById.set(key, row);
    });
  }

  function indexWealthDesk(aggregates, rows) {
    rows.forEach(row => {
      const key = leftAccountKey(row["Account Number"]);
      if (key) aggregates.wealthDeskByAccount.set(key, row);
    });
  }

  function indexAccounts(aggregates, rows) {
    rows.forEach(row => {
      const accountKey = leftAccountKey(row["Account Number"]);
      const clientName = cleanText(row["Sort Name"]);
      const clientKey = normalizeTextKey(clientName);
      if (!accountKey || !clientKey) return;
      aggregates.accountByKey.set(accountKey, row);
      aggregates.clientKeys.set(clientKey, clientName);
      const summary = getOrCreate(aggregates.accountSummaryByClient, clientKey, () => ({
        accounts: [],
        statuses: [],
        totalNetWorth: null,
        annualIncome: null,
        t12RevenueFromAccount: null
      }));
      row.__accountKey = accountKey;
      row.__clientGroupId = clientName;
      summary.accounts.push(row);
      summary.statuses.push(row.Status);
      summary.totalNetWorth = sumNullable(summary.totalNetWorth, toNumber(row["Total Net Worth"]));
      summary.annualIncome = sumNullable(summary.annualIncome, toNumber(row["Total Annual Income"]));
      summary.t12RevenueFromAccount = sumNullable(summary.t12RevenueFromAccount, toNumber(row["T12 Revenue($)"]));
    });
  }

  function aggregateRevenue(aggregates, rows) {
    const cutoff = new Date(todayDate().getTime() - 365 * 86400000);
    const amountField = appState.settings.revenueAmountField || "Amt Cmp Grs";
    rows.forEach(row => {
      const date = toDate(row["Dt2 Trd Entr"]);
      if (!date || date < cutoff) return;
      const accountKey = buildAccountKey(row["Num Alt Ofc"], row["Num Acc"]);
      const account = aggregates.accountByKey.get(accountKey);
      if (!account) return;
      const clientKey = normalizeTextKey(account.__clientGroupId || account["Sort Name"]);
      const product = aggregates.productByCode.get(cleanText(row["Cde Prd"])) || {};
      const amount = toNumber(row[amountField]);
      if (amount === null) return;
      const summary = getOrCreate(aggregates.revenueByClient, clientKey, () => emptyRevenueSummary());
      const enriched = {
        ...row,
        __date: date.toISOString().slice(0, 10),
        __month: date.toISOString().slice(0, 7),
        __amount: amount,
        __accountKey: accountKey,
        __productGroup: cleanText(product.GROUP || product.PRODUCT_L1 || "UNKNOWN"),
        __productL1: cleanText(product.PRODUCT_L1 || "UNKNOWN"),
        __productL2: cleanText(product.PRODUCT_L2 || "UNKNOWN"),
        __productL3: cleanText(product.PRODUCT_L3 || "UNKNOWN")
      };
      summary.total = sumNullable(summary.total, amount);
      summary.rows.push(enriched);
      addToObjectSum(summary.byProduct, enriched.__productGroup, amount);
      addToObjectSum(summary.byMonth, enriched.__month, amount);
    });
  }

  function aggregateNna(aggregates, rows) {
    const cutoff = new Date(todayDate().getTime() - 365 * 86400000);
    const excluded = (appState.settings.nnaExcludedKeywords || []).map(normalizeTextKey);
    rows.forEach(row => {
      const date = toDate(row["Activity Date"]);
      if (!date || date < cutoff) return;
      const type = cleanText(row["Activity Type"]);
      const desc = cleanText(row["Activity Description"]);
      const combined = normalizeTextKey(`${type} ${desc}`);
      if (excluded.some(word => word && combined.includes(word))) return;
      const accountKey = leftAccountKey(row["Account Number"]);
      const account = aggregates.accountByKey.get(accountKey);
      const clientName = account?.__clientGroupId || row["Sort Name"];
      const clientKey = normalizeTextKey(clientName);
      if (!clientKey) return;
      const amount = toNumber(row["Amount($)"]);
      if (amount === null) return;
      const summary = getOrCreate(aggregates.nnaByClient, clientKey, () => emptyNnaSummary());
      const enriched = {
        ...row,
        __date: date.toISOString().slice(0, 10),
        __month: date.toISOString().slice(0, 7),
        __amount: amount,
        __activityType: type || "UNKNOWN",
        __accountKey: accountKey
      };
      summary.total = sumNullable(summary.total, amount);
      summary.rows.push(enriched);
      addToObjectSum(summary.byType, enriched.__activityType, amount);
      addToObjectSum(summary.byMonth, enriched.__month, amount);
    });
  }

  function aggregatePositions(aggregates, rows, external) {
    rows.forEach(row => {
      const accountKey = row["Account Number"] && row["Branch Number"] ? buildAccountKey(row["Branch Number"], row["Account Number"]) : leftAccountKey(row["Account Number"]);
      const account = aggregates.accountByKey.get(accountKey);
      const clientName = account?.__clientGroupId || getOptionalClientId(row);
      const clientKey = normalizeTextKey(clientName);
      if (!clientKey) return;
      aggregates.clientKeys.set(clientKey, cleanText(clientName));
      const summary = getOrCreate(aggregates.positionByClient, clientKey, () => emptyPositionSummary());
      const marketValue = toNumber(row["Amt Mkt Val"] ?? row["Market Value"] ?? row["MarketValue"]);
      if (marketValue === null) return;
      const security = aggregates.securityById.get(cleanText(row["Cde Msdw Sec"])) || {};
      const wealthDesk = aggregates.wealthDeskByAccount.get(accountKey) || {};
      const managed = isManaged(row, account, wealthDesk);
      const category = classifyProductCategory(row, security);
      if (isCategoryDisabled(category)) return;
      const assetClass = classifyAssetClass(row, security);
      const custodian = external ? cleanText(row.Custodian || row.Platform || "External") : "Morgan Stanley";
      const enriched = { ...row, __accountKey: accountKey, __marketValue: marketValue, __assetClass: assetClass, __category: category, __custodian: custodian, __managed: managed, __external: external };
      if (!summary.rows) summary.rows = [];
      summary.rows.push(enriched);
      summary.totalAssets = sumNullable(summary.totalAssets, marketValue);
      if (external) summary.externalAssets = sumNullable(summary.externalAssets, marketValue);
      else summary.msAssets = sumNullable(summary.msAssets, marketValue);
      addToObjectSum(summary.categoryBreakdown, category, marketValue);
      addToObjectSum(summary.custodianBreakdown, custodian, marketValue);
      addToObjectSum(summary.allocationByClass || (summary.allocationByClass = {}), assetClass, marketValue);
      if (managed) {
        summary.advisoryAssets = sumNullable(summary.advisoryAssets, marketValue);
        addToObjectSum(summary.strategyBreakdown, cleanText(wealthDesk["Investment Type"] || security.STRATEGY || "UNKNOWN"), marketValue);
      }
      if (isAvailableCashPosition(row, account, security, managed)) {
        const cashType = cashTypeForPosition(row, security);
        summary.availableCashRaw = sumNullable(summary.availableCashRaw, marketValue);
        if (!summary.cashBreakdown[cashType]) summary.cashBreakdown[cashType] = { taxable: 0, ira: 0, total: 0 };
        const bucket = isIraAccount(account) ? "ira" : "taxable";
        summary.cashBreakdown[cashType][bucket] += marketValue;
        summary.cashBreakdown[cashType].total += marketValue;
      }
      const maturity = toDate(row["Mat Exp Date"] || row["Dt2 Call"]);
      if (maturity) {
        const days = Math.round((maturity - todayDate()) / 86400000);
        if (days >= 0 && days <= 30) {
          summary.maturityList.push({
            securityName: cleanText(row["Pref Iss Nme"] || security["Pref Iss Nme"] || row["Nme Sym"] || "Security"),
            accountKey,
            date: maturity.toISOString().slice(0, 10),
            days,
            marketValue
          });
        }
      }
    });
  }

  function finalizePositionAggregates(aggregates) {
    aggregates.positionByClient.forEach(summary => {
      const total = summary.totalAssets || 0;
      const rawCash = summary.availableCashRaw;
      const excess = rawCash === null ? null : rawCash;
      summary.advisoryPercent = total && summary.advisoryAssets !== null ? (summary.advisoryAssets / total) * 100 : null;
      summary.availableCashExcess = excess;
      summary.availableCashPercent = total && excess !== null ? (excess / total) * 100 : null;
      summary.allocationRows = Object.entries(summary.allocationByClass || {}).map(([assetClass, value]) => ({
        assetClass,
        value,
        currentPercent: total ? (value / total) * 100 : 0,
        targetPercent: null,
        driftPercent: null,
        thresholdPercent: Number(appState.settings.driftThresholds?.[assetClass] ?? appState.settings.driftThresholds?.UNKNOWN ?? 0),
        status: "N/A"
      }));
      summary.singleStockAlerts = (summary.rows || [])
        .filter(row => total > 0 && looksLikeSingleStock(row) && row.__marketValue / total > 0.10)
        .map(row => ({
          securityName: cleanText(row["Pref Iss Nme"] || row["Nme Sym"] || row["Cde Msdw Sec"] || "Security"),
          value: row.__marketValue,
          percent: (row.__marketValue / total) * 100
        }));
    });
  }

  function aggregateFinancialPlans(aggregates, rows) {
    if (!rows.length) return;
    const mapping = appState.settings.financialPlanMapping || {};
    const clientField = mapping.clientGroupId || detectClientFields(rows)[0];
    const likelihoodField = mapping.likelihood || detectField(rows, ["likelihood", "success"]);
    const statusField = mapping.status || detectField(rows, ["status"]);
    const lastUpdateField = mapping.lastUpdateDate || detectField(rows, ["last", "update"]);
    const nextReviewField = mapping.nextReviewDate || detectField(rows, ["next", "review"]);
    rows.forEach(row => {
      const clientKey = normalizeTextKey(row[clientField]);
      if (!clientKey) return;
      const likelihood = toNumber(row[likelihoodField]);
      aggregates.financialByClient.set(clientKey, {
        likelihood,
        display: likelihood === null ? "N/A" : `${likelihood.toFixed(0)}%`,
        status: safeText(row[statusField]),
        lastUpdateDate: formatDate(row[lastUpdateField]),
        nextReviewDate: formatDate(row[nextReviewField]),
        sourceFields: { clientField, likelihoodField, statusField, lastUpdateField, nextReviewField }
      });
    });
  }

  function aggregateOnlineActivity(aggregates) {
    const mapping = appState.settings.onlineActivityMapping || {};
    const rows = runtime.importedTables[mapping.tableName || "ref_ACCOUNT"] || [];
    if (!rows.length) return;
    const clientField = mapping.clientGroupId || "Sort Name";
    const usernameField = mapping.username;
    const loginField = mapping.lastLoginDate;
    if (!loginField) return;
    const byClient = new Map();
    rows.forEach(row => {
      const clientKey = normalizeTextKey(row[clientField]);
      if (!clientKey) return;
      const lastLogin = toDate(row[loginField]);
      const current = byClient.get(clientKey);
      if (!current || (lastLogin && (!current.lastLoginDateRaw || lastLogin > current.lastLoginDateRaw))) {
        byClient.set(clientKey, {
          username: cleanText(row[usernameField]),
          lastLoginDateRaw: lastLogin,
          lastLoginDate: lastLogin ? lastLogin.toISOString().slice(0, 10) : null
        });
      }
    });
    byClient.forEach((value, key) => {
      if (!value.lastLoginDateRaw) {
        aggregates.onlineByClient.set(key, { display: "N/A", className: "cellMuted", ...value });
        return;
      }
      const days = Math.round((todayDate() - value.lastLoginDateRaw) / 86400000);
      const className = days < 31 ? "dotGreen" : days < 180 ? "dotYellow" : "dotRed";
      aggregates.onlineByClient.set(key, { display: `${days}d`, daysSinceLogin: days, className, ...value });
    });
  }

  function getOptionalClientId(row) {
    const fields = ["ID_NAME_ALG", "Client Group ID", "clientGroupId", "Sort Name", "Client", "Client Name", "Household"];
    return firstNonBlank(fields.map(field => row[field]));
  }

  function isManaged(position, account, wealthDesk) {
    return isYes(position.Managed) || Boolean(wealthDesk && Object.keys(wealthDesk).length) || isYes(account?.Managed);
  }

  function isIraAccount(account) {
    const text = normalizeTextKey(`${account?.["Account Type"] || ""} ${account?.["Account Taxable / Non-Taxable"] || ""}`);
    return text.includes("IRA") || text.includes("ROTH") || text.includes("SEP");
  }

  function isTaxableAccount(account) {
    const tax = normalizeTextKey(account?.["Account Taxable / Non-Taxable"]);
    if (!tax) return false;
    return tax.includes("TAXABLE") || tax.includes("NON-QUALIFIED");
  }

  function isAvailableCashPosition(row, account, security, managed) {
    if (managed || !account || !isTaxableAccount(account) || isIraAccount(account)) return false;
    if (isYes(security.isCASH) || isYes(security.isBDPS)) return true;
    const haystack = normalizeTextKey([
      row["Nme Iss Typ"],
      row["Nme Iss Sub Typ"],
      row["Nme Asset Typ"],
      row["Txt Lvl1 Ast Desc"],
      row["Txt Lvl2 Ast Desc"],
      row["Txt Lvl3 Ast Desc"],
      row["Pref Iss Nme"],
      row["Nme Sym"]
    ].join(" "));
    return (appState.settings.cashKeywords || DEFAULT_SETTINGS.cashKeywords).some(keyword => haystack.includes(normalizeTextKey(keyword)));
  }

  function cashTypeForPosition(row, security) {
    const text = normalizeTextKey(`${row["Pref Iss Nme"] || ""} ${row["Nme Iss Typ"] || ""} ${row["Nme Iss Sub Typ"] || ""} ${row["Txt Lvl1 Ast Desc"] || ""}`);
    if (isYes(security.isBDPS) || text.includes("BDPS") || text.includes("BANK DEPOSIT")) return "Savings Deposits";
    if (text.includes("MONEY MARKET")) return "Money Market";
    return "Cash";
  }

  function classifyProductCategory(row, security) {
    return cleanText(row["Txt Lvl3 Ast Desc"] || row["Nme Asset Typ"] || security.STRATEGY || "UNKNOWN");
  }

  function isCategoryDisabled(category) {
    return (appState.settings.totalAssetDisabledCategories || []).map(normalizeTextKey).includes(normalizeTextKey(category));
  }

  function classifyAssetClass(row, security) {
    const manual = appState.settings.assetClassMappings || {};
    const candidates = [
      security.STRATEGY,
      row["Txt Lvl3 Ast Desc"],
      row["Txt Lvl2 Ast Desc"],
      row["Txt Lvl1 Ast Desc"],
      row["Nme Asset Typ"],
      row["Nme Iss Typ"]
    ].map(cleanText).filter(Boolean);
    for (const item of candidates) {
      const mapped = manual[item] || manual[normalizeTextKey(item)];
      if (mapped) return mapped;
    }
    const haystack = normalizeTextKey(candidates.join(" "));
    if (isYes(security.isCASH) || isYes(security.isBDPS) || haystack.includes("CASH") || haystack.includes("MONEY MARKET")) return "Cash";
    if (isYes(security.isALT) || haystack.includes("ALTERNATIVE") || haystack.includes("ALT")) return "Alternatives";
    if (haystack.includes("FIXED") || haystack.includes("BOND") || haystack.includes("MUNI") || haystack.includes("TREASURY")) return "Fixed Income";
    if (haystack.includes("EQUITY") || haystack.includes("STOCK") || haystack.includes("ETF") || haystack.includes("MUTUAL FUND")) return "Equity";
    return candidates[0] || "UNKNOWN";
  }

  function looksLikeSingleStock(row) {
    const text = normalizeTextKey(`${row["Txt Lvl1 Ast Desc"] || ""} ${row["Nme Iss Typ"] || ""} ${row["Nme Asset Typ"] || ""}`);
    return text.includes("COMMON STOCK") || text === "EQUITY" || text.includes("STOCK");
  }

  function computeAllocationStatus(row, positions) {
    if (!positions || !positions.allocationRows?.length || !row.totalAssets) return "N/A";
    const target = targetAllocationFor(row);
    let drifted = false;
    positions.allocationRows.forEach(item => {
      const targetPercent = Number(target[item.assetClass] ?? target[normalizeTextKey(item.assetClass)] ?? 0);
      item.targetPercent = targetPercent;
      item.driftPercent = item.currentPercent - targetPercent;
      item.thresholdPercent = Number(appState.settings.driftThresholds?.[item.assetClass] ?? appState.settings.driftThresholds?.UNKNOWN ?? 0);
      item.status = Math.abs(item.driftPercent) > item.thresholdPercent ? "DRIFT" : "OK";
      if (item.status === "DRIFT") drifted = true;
    });
    return drifted ? "DRIFT" : "OK";
  }

  function targetAllocationFor(row) {
    const imported = importedTargetAllocationFor(row);
    if (Object.keys(imported).length) return imported;
    const key = `${row.riskProfile}|${row.includeAlts}|${row.includeMunis}`;
    return DEFAULT_TARGET_ALLOCATIONS[key] || DEFAULT_TARGET_ALLOCATIONS[`${row.riskProfile}|No|No`] || {};
  }

  function importedTargetAllocationFor(row) {
    const rows = runtime.importedTables.ref_ALLOCATION_TARGET || [];
    if (!rows.length) return {};
    const mapping = appState.settings.allocationTargetMapping || {};
    const target = {};
    rows.forEach(item => {
      const profile = cleanText(item[mapping.profile]);
      const includeAlts = cleanText(item[mapping.includeAlts] || "No");
      const includeMunis = cleanText(item[mapping.includeMunis] || "No");
      if (profile !== row.riskProfile || includeAlts !== row.includeAlts || includeMunis !== row.includeMunis) return;
      const assetClass = cleanText(item[mapping.assetClass]);
      const pct = toNumber(item[mapping.targetPercent]);
      if (assetClass && pct !== null) target[assetClass] = pct;
    });
    return target;
  }

  function sumNullable(current, value) {
    if (value === null || value === undefined || Number.isNaN(value)) return current;
    return (current ?? 0) + value;
  }

  function addToObjectSum(obj, key, value) {
    const name = cleanText(key) || "UNKNOWN";
    obj[name] = (obj[name] || 0) + value;
  }

  function getOrCreate(map, key, factory) {
    if (!map.has(key)) map.set(key, factory());
    return map.get(key);
  }

  async function importSelectedFiles(fileList, sourceMode) {
    const expectedFiles = ensureExpectedFileConfig();
    if (!window.XLSX) {
      runtime.warnings = ["SheetJS is missing. Place xlsx.full.min.js at lib/xlsx.full.min.js and reopen index.html."];
      runtime.fileStatuses = expectedFiles.map(item => ({ ...item, status: "xlsx-library-missing", message: "SheetJS missing" }));
      render();
      return;
    }
    const files = [...fileList].filter(file => /\.xlsx$/i.test(file.name));
    const byLowerName = new Map(files.map(file => [file.name.toLowerCase(), file]));
    const nextRuntime = emptyRuntime();
    nextRuntime.sourceMode = sourceMode;
    appState.priorRefreshSnapshot = clone(appState.snapshots);
    appState.dataManagement.sourceFolderName = inferFolderName(files);

    for (const expected of expectedFiles) {
      const file = byLowerName.get(expected.fileName.toLowerCase());
      if (!file) {
        nextRuntime.fileStatuses.push({
          ...expected,
          status: expected.required ? "missing" : "optional-missing",
          rowCount: 0,
          message: expected.required ? "Required workbook missing" : "Optional workbook missing"
        });
        continue;
      }
      await readWorkbookIntoRuntime(file, expected, nextRuntime);
    }

    runtime = nextRuntime;
    runtime.loadedAt = new Date().toISOString();
    appState.lastRefreshAt = runtime.loadedAt;
    appState.snapshots.importedTableHeaders = clone(runtime.importedHeaders);
    render();
    saveState({ type: "data_refresh", summary: `Imported ${Object.keys(runtime.importedTables).length} local workbook tables.` });
    if (appState.settings.autoBackups) downloadBackup();
  }

  async function grantSourceFolderAccess() {
    if (!window.showDirectoryPicker) {
      document.getElementById("xlsxFolderInput").click();
      return;
    }
    const handle = await window.showDirectoryPicker({ id: "am-crm-source-workbooks", mode: "read" });
    await storeDirectoryHandle(handle);
    if (!appState.settings.sourceFolderPath) appState.settings.sourceFolderPath = handle.name;
    appState.dataManagement.sourceFolderName = handle.name;
    saveState({ type: "source_folder_grant", summary: `Granted folder access: ${handle.name}` });
    await refreshFromDirectoryHandle(handle, "granted-folder");
  }

  async function autoRefreshSavedFolder(force = false) {
    if (!force && autoRefreshAttempted) return false;
    autoRefreshAttempted = true;
    if (!appState.settings.autoRefreshSavedFolder && !force) return false;
    const handle = await loadDirectoryHandle();
    if (!handle) return false;
    const permission = await verifyDirectoryPermission(handle, force);
    if (!permission) return false;
    await refreshFromDirectoryHandle(handle, "saved-folder");
    return true;
  }

  async function refreshConfiguredSourceFolder() {
    runtime.warnings = [];
    const handle = await loadDirectoryHandle();
    if (!handle) {
      const configured = cleanText(appState.settings.sourceFolderPath || appState.dataManagement.sourceFolderName || "");
      const message = configured
        ? `${CONFIGURED_SOURCE_FOLDER_NOT_GRANTED_PREFIX}${configured}${CONFIGURED_SOURCE_FOLDER_NOT_GRANTED_SUFFIX}`
        : SOURCE_FOLDER_NOT_GRANTED_WARNING;
      runtime.warnings = [message];
      render();
      alert(message);
      return false;
    }

    const permission = await verifyDirectoryPermission(handle, true);
    if (!permission) {
      const message = SOURCE_FOLDER_PERMISSION_DENIED_WARNING;
      runtime.warnings = [message];
      render();
      alert(message);
      return false;
    }

    await refreshFromDirectoryHandle(handle, "saved-folder");
    return true;
  }

  async function refreshFromDirectoryHandle(handle, sourceMode) {
    const files = [];
    for await (const entry of handle.values()) {
      if (entry.kind !== "file" || !/\.xlsx$/i.test(entry.name)) continue;
      files.push(await entry.getFile());
    }
    appState.dataManagement.sourceFolderName = handle.name;
    await importSelectedFiles(files, sourceMode);
  }

  async function verifyDirectoryPermission(handle, requestIfNeeded) {
    if (!handle.queryPermission || !handle.requestPermission) return true;
    const query = await handle.queryPermission({ mode: "read" });
    if (query === "granted") return true;
    if (!requestIfNeeded) return false;
    return (await handle.requestPermission({ mode: "read" })) === "granted";
  }

  function openFolderHandleDb() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(FOLDER_HANDLE_DB_NAME, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(FOLDER_HANDLE_STORE_NAME);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function storeDirectoryHandle(handle) {
    const db = await openFolderHandleDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(FOLDER_HANDLE_STORE_NAME, "readwrite");
      tx.objectStore(FOLDER_HANDLE_STORE_NAME).put(handle, SOURCE_FOLDER_HANDLE_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  async function loadDirectoryHandle() {
    if (!window.indexedDB) return null;
    try {
      const db = await openFolderHandleDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(FOLDER_HANDLE_STORE_NAME, "readonly");
        const request = tx.objectStore(FOLDER_HANDLE_STORE_NAME).get(SOURCE_FOLDER_HANDLE_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
      });
    } catch (err) {
      console.warn("Could not load saved folder handle.", err);
      return null;
    }
  }

  async function readWorkbookIntoRuntime(file, expected, targetRuntime) {
    const workbookConfig = normalizeWorkbookConfig(expected);
    try {
      const workbook = window.XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false, range: workbookConfig.headerRow - 1 });
      const normalizedRows = rows.map(normalizeImportedRow);
      const headers = Object.keys(normalizedRows[0] || {}).length ? Object.keys(normalizedRows[0] || {}) : worksheetHeaderFields(sheet, workbookConfig.headerRow);
      targetRuntime.importedTables[workbookConfig.tableName] = normalizedRows;
      targetRuntime.importedHeaders[workbookConfig.tableName] = headers;
      const headerValidation = validateHeaders(workbookConfig.tableName, headers);
      const stale = isFileStale(file);
      targetRuntime.fileStatuses.push({
        ...workbookConfig,
        status: headerValidation.status === "ok" && !stale ? "loaded" : stale ? "stale" : "changed-columns",
        rowCount: normalizedRows.length,
        sheetName,
        lastModified: file.lastModified ? new Date(file.lastModified).toISOString() : null,
        missingColumns: headerValidation.missingColumns,
        extraColumns: headerValidation.extraColumns,
        message: stale ? `Workbook last modified more than ${FILE_STALE_DAYS} days ago` : headerValidation.message
      });
    } catch (err) {
      targetRuntime.fileStatuses.push({ ...workbookConfig, status: "error", rowCount: 0, message: err.message });
    }
  }

  function worksheetHeaderFields(sheet, headerRow) {
    const previewRows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false, range: headerRow - 1, blankrows: false });
    return (previewRows[0] || []).map(cleanText).filter(Boolean);
  }

  function normalizeImportedRow(row) {
    const out = {};
    Object.entries(row).forEach(([key, value]) => {
      const cleanKey = cleanText(key);
      out[cleanKey] = typeof value === "string" ? value.trim() : value;
    });
    return out;
  }

  function validateHeaders(tableName, headers) {
    const expectedFields = (schemaCatalog.tables?.[tableName]?.fields || []).map(field => field.fieldName);
    if (!expectedFields.length) return { status: "ok", missingColumns: [], extraColumns: [], message: "Schema not attached; using workbook headers." };
    const headerSet = new Set(headers);
    const expectedSet = new Set(expectedFields);
    const missingColumns = expectedFields.filter(field => !headerSet.has(field));
    const extraColumns = headers.filter(header => !expectedSet.has(header));
    return {
      status: missingColumns.length ? "changed" : "ok",
      missingColumns,
      extraColumns,
      message: missingColumns.length ? "Expected columns missing" : "Headers match expected schema"
    };
  }

  function isFileStale(file) {
    if (!file.lastModified) return false;
    return todayDate() - new Date(file.lastModified) > FILE_STALE_DAYS * 86400000;
  }

  function inferFolderName(files) {
    const path = files.find(file => file.webkitRelativePath)?.webkitRelativePath || "";
    return path.includes("/") ? path.split("/")[0] : files.length ? "Selected files" : null;
  }

  function validateRelationships() {
    const messages = [];
    (appState.dataManagement.joins || []).filter(rel => rel.enabled !== false).forEach(rel => {
      const lookupRows = runtime.importedTables[rel.lookupParentTable] || [];
      const detailRows = runtime.importedTables[rel.detailRelatedTable] || [];
      if (!lookupRows.length || !detailRows.length) return;
      const lookupCounts = new Map();
      lookupRows.forEach(row => {
        const key = relationshipKey(row, rel, "lookup");
        if (!key) return;
        lookupCounts.set(key, (lookupCounts.get(key) || 0) + 1);
      });
      const duplicates = [...lookupCounts.entries()].filter(([, count]) => count > 1).length;
      if (duplicates) messages.push({ severity: "warn", relationshipId: rel.relationshipId, message: `${duplicates} duplicate lookup keys in ${rel.lookupParentTable}.` });
      let blankDetails = 0;
      let unmatched = 0;
      detailRows.forEach(row => {
        const key = relationshipKey(row, rel, "detail");
        if (!key) {
          blankDetails += 1;
          return;
        }
        if (!lookupCounts.has(key)) unmatched += 1;
      });
      if (blankDetails) messages.push({ severity: "warn", relationshipId: rel.relationshipId, message: `${blankDetails} blank detail keys in ${rel.detailRelatedTable}.` });
      if (unmatched) messages.push({ severity: "warn", relationshipId: rel.relationshipId, message: `${unmatched} ${rel.detailRelatedTable} rows have no ${rel.lookupParentTable} match.` });
    });
    runtime.fileStatuses.forEach(status => {
      if (["missing", "optional-missing", "changed-columns", "stale", "error"].includes(status.status)) {
        messages.push({ severity: status.required ? "error" : "warn", relationshipId: "FILE", message: `${status.fileName}: ${status.message}` });
      }
    });
    return messages;
  }

  function relationshipKey(row, rel, side) {
    const id = rel.relationshipId;
    if (id === "R001") return cleanText(row["Cde Prd"]);
    if (id === "R002" && side === "lookup") return leftAccountKey(row["Account Number"]);
    if (id === "R002" && side === "detail") return buildAccountKey(row["Num Alt Ofc"], row["Num Acc"]);
    if (id === "R003" && side === "lookup") return leftAccountKey(row["Account Number"]);
    if (id === "R003" && side === "detail") return buildAccountKey(row["Branch Number"], row["Account Number"]);
    if (id === "R004" && side === "lookup") return leftAccountKey(row["Account Number"]);
    if (id === "R004" && side === "detail") return buildAccountKey(row["Branch Number"], row["Account Number"]);
    if (id === "R005") return leftAccountKey(row["Account Number"]);
    if (id === "R006" && side === "lookup") return cleanText(row.ID_MSDW_SECURITY);
    if (id === "R006" && side === "detail") return cleanText(row["Cde Msdw Sec"]);
    if (id === "R007" && side === "lookup") return normalizeTextKey(row.ID_NAME_ALG);
    if (id === "R007" && side === "detail") return normalizeTextKey(row["Sort Name"]);
    const expr = side === "lookup" ? rel.lookupKeyColumnOrExpression : rel.detailKeyColumnOrExpression;
    return expressionKey(row, expr);
  }

  function expressionKey(row, expression) {
    const expr = cleanText(expression);
    const fields = [...expr.matchAll(/\[([^\]]+)\]/g)].map(match => match[1]);
    if (expr.toUpperCase().includes("LEFT") && fields[0]) return leftAccountKey(row[fields[0]]);
    if (expr.toUpperCase().includes("UPPER") && fields[0]) return normalizeTextKey(row[fields[0]]);
    if (expr.toUpperCase().includes("TEXT") && fields.length >= 2) return buildAccountKey(row[fields[0]], row[fields[1]]);
    if (fields[0]) return cleanText(row[fields[0]]);
    return cleanText(row[expr]);
  }

  function renderGrid() {
    const table = document.getElementById("clientGrid");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    const tfoot = table.querySelector("tfoot");
    thead.innerHTML = `<tr>${COLUMN_DEFS.map(col => {
      const stickyClass = col.sticky ? ` stickyCol${col.sticky}` : "";
      const sortIcon = activeSort.key === col.key ? (activeSort.direction === "asc" ? " ^" : " v") : "";
      return `<th class="sortable${stickyClass}" data-key="${col.key}" style="width:${columnWidth(col.key)}px" title="${escapeHtml(col.label)}"><span class="headerLabel">${escapeHtml(col.label.toUpperCase())}${sortIcon}</span><span class="columnResizeHandle" data-key="${col.key}" title="Resize column"></span></th>`;
    }).join("")}</tr>`;

    const rows = applySort(applySearch(gridRows));
    tbody.innerHTML = rows.map(row => `<tr data-client-id="${escapeHtml(row.id)}">${COLUMN_DEFS.map(col => renderCell(row, col)).join("")}</tr>`).join("");
    tfoot.innerHTML = `<tr>${COLUMN_DEFS.map(col => footerCell(rows, col)).join("")}</tr>`;

    thead.querySelectorAll("th").forEach(th => {
      th.addEventListener("click", () => {
        const key = th.dataset.key;
        if (activeSort.key === key) activeSort.direction = activeSort.direction === "asc" ? "desc" : "asc";
        else activeSort = { key, direction: COLUMN_DEFS.find(col => col.key === key)?.numeric ? "desc" : "asc" };
        renderGrid();
      });
    });
    thead.querySelectorAll(".columnResizeHandle").forEach(handle => {
      handle.addEventListener("click", event => event.stopPropagation());
      handle.addEventListener("pointerdown", startColumnResize);
    });
    tbody.querySelectorAll("td").forEach(td => {
      td.addEventListener("click", event => {
        const row = gridRows.find(item => item.id === td.parentElement.dataset.clientId);
        const col = COLUMN_DEFS.find(item => item.key === td.dataset.key);
        if (row && col) openPopup(row, col, event);
      });
    });
  }

  function footerCell(rows, col) {
    const stickyClass = col.sticky ? ` stickyCol${col.sticky}` : "";
    let value = "";
    if (col.key === "clientName") value = `${rows.length} clients`;
    if (col.key === "t12Revenue") value = formatMoney(sumRows(rows, "t12Revenue"));
    if (col.key === "nna") value = formatMoney(sumRows(rows, "nna"));
    if (col.key === "totalAssets") value = formatMoney(sumRows(rows, "totalAssets"));
    if (col.key === "tasks") value = rows.reduce((sum, row) => sum + Number(row.tasks || 0), 0);
    return `<td class="${stickyClass}">${escapeHtml(value)}</td>`;
  }

  function sumRows(rows, key) {
    const values = rows.map(row => toNumber(row[key])).filter(value => value !== null);
    return values.length ? values.reduce((sum, value) => sum + value, 0) : null;
  }

  function columnWidth(key) {
    return appState.tableLayout.columns.find(col => col.key === key)?.width || COLUMN_DEFS.find(col => col.key === key)?.width || 100;
  }

  function setColumnWidth(key, width) {
    const clamped = clampNumber(width, 48, 520, COLUMN_DEFS.find(col => col.key === key)?.width || 100);
    let layout = appState.tableLayout.columns.find(col => col.key === key);
    if (!layout) {
      layout = { key, width: clamped };
      appState.tableLayout.columns.push(layout);
    }
    layout.width = clamped;
  }

  function startColumnResize(event) {
    event.preventDefault();
    event.stopPropagation();
    const key = event.currentTarget.dataset.key;
    const startX = event.clientX;
    const startWidth = columnWidth(key);
    document.body.classList.add("isResizingColumn");
    const onMove = moveEvent => {
      setColumnWidth(key, startWidth + moveEvent.clientX - startX);
      renderGrid();
    };
    const onUp = () => {
      document.body.classList.remove("isResizingColumn");
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      saveState({ type: "layout_resize", summary: `Changed ${key} column width to ${columnWidth(key)}px` });
      if (isSettingsVisible() && activeSettingsSection === "Main Table") renderSettings();
    };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
  }

  function applySearch(rows) {
    const q = document.getElementById("globalSearch").value.trim().toLowerCase();
    if (q.length < SEARCH_MIN_CHARS) return rows;
    return rows.filter(row => {
      const text = [
        row.clientName,
        row.email,
        row.phone,
        row.salutation,
        row.status,
        row.financialPlan,
        ...appState.notes.filter(note => note.clientGroupId === row.id).map(note => note.body),
        ...appState.tasks.filter(task => task.clientGroupId === row.id).map(task => `${task.title} ${task.notes || ""}`)
      ].join(" ").toLowerCase();
      return text.includes(q);
    });
  }

  function applySort(rows) {
    const sorted = [...rows];
    sorted.sort((a, b) => compareRows(a, b, activeSort.key, activeSort.direction));
    return sorted;
  }

  function compareRows(a, b, key, direction) {
    const col = COLUMN_DEFS.find(item => item.key === key);
    const dir = direction === "asc" ? 1 : -1;
    if (key === "nextContact") return compareNullableNumber(a.nextContact?.days, b.nextContact?.days, dir);
    if (col?.numeric) return compareNullableNumber(a[key], b[key], dir);
    const av = String(a[key]?.display || a[key] || "").toLowerCase();
    const bv = String(b[key]?.display || b[key] || "").toLowerCase();
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return String(a.clientName).localeCompare(String(b.clientName));
  }

  function compareNullableNumber(a, b, dir) {
    const av = toNumber(a);
    const bv = toNumber(b);
    if (av === null && bv === null) return 0;
    if (av === null) return 1;
    if (bv === null) return -1;
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return 0;
  }

  function renderCell(row, col) {
    const stickyClass = col.sticky ? ` stickyCol${col.sticky}` : "";
    const align = `text-align:${col.key === "clientName" ? "left" : "center"};`;
    let className = stickyClass;
    let html = "";
    if (col.key === "status") html = statusIcon(row.status, row.id);
    else if (col.key === "clientName") {
      html = escapeHtml(row.clientName);
      className += " clientNameCell";
    } else if (col.key === "email") {
      html = gridIcon("grid_email", "emailIcon", row.email || "N/A");
      if (!row.email) className += " cellMuted";
    }
    else if (col.key === "nextContact") {
      html = escapeHtml(row.nextContact?.display || "N/A");
      className += ` ${row.nextContact?.className || ""}`;
    } else if (col.key === "t12Revenue" || col.key === "nna" || col.key === "totalAssets") {
      html = formatMoney(row[col.key]);
      if (html === "N/A") className += " cellMuted";
    } else if (col.key === "availableCash") {
      html = row.availableCash === null || row.availableCash === undefined ? "N/A" : `${formatMoney(row.availableCash, "k")} / ${formatPercent(row.availableCashPercent)}`;
      if (html === "N/A") className += " cellMuted";
      if (row.availableCashAlert) className += " cellDanger";
    } else if (col.key === "assetAllocation") {
      html = escapeHtml(row.assetAllocation || "N/A");
      if (row.assetAllocation === "OK") className += " cellSuccess";
      else if (row.assetAllocation === "DRIFT") className += " cellDanger";
      else className += " cellMuted";
    } else if (col.key === "financialPlan") {
      html = escapeHtml(row.financialPlan || "N/A");
      const likelihood = toNumber(row.financialPlan);
      if (likelihood === null) className += " cellMuted";
      else className += likelihood >= 85 ? " cellSuccess" : " cellDanger";
    } else if (col.key === "onlineActivity") {
      html = row.onlineActivity?.className ? `<span class="activityDot ${escapeHtml(row.onlineActivity.className)}"></span>` : `<span class="cellMuted">N/A</span>`;
    } else if (col.key === "tasks") {
      html = escapeHtml(row.tasks || 0);
    } else if (col.key === "notes") {
      html = gridIcon("grid_notes", "noteIcon", `${row.notes || 0} notes`);
    } else {
      html = escapeHtml(safeText(row[col.key]));
      if (stripTags(html) === "N/A") className += " cellMuted";
    }
    return `<td data-key="${col.key}" class="${className.trim()}" style="${align}" title="${escapeHtml(stripTags(html))}">${html}</td>`;
  }

  function statusIcon(status, clientId) {
    const tier = validStatus(status);
    const override = tier ? appState.iconOverrides[`status_${tier}`] : null;
    if (override) return `<img class="statusImg" alt="${escapeHtml(tier)}" src="${escapeHtml(override)}">`;
    if (!tier) return `<span class="cellMuted">N/A</span>`;
    return `<span class="statusDot status${escapeHtml(tier)}" title="${escapeHtml(`${tier} - ${STATUS_LABELS[tier]}`)}"></span>`;
  }

  function gridIcon(key, fallbackClass, title) {
    const override = appState.iconOverrides[key];
    if (override) return `<img class="gridIconImg" alt="${escapeHtml(title)}" src="${escapeHtml(override)}">`;
    return `<span class="${escapeHtml(fallbackClass)}" title="${escapeHtml(title)}"></span>`;
  }

  function renderAppIcons() {
    Object.entries(SIDEBAR_BUTTON_IDS).forEach(([key, id]) => {
      const button = document.getElementById(id);
      if (!button) return;
      const override = appState.iconOverrides[key];
      button.innerHTML = override ? `<img class="buttonIconImg" alt="" src="${escapeHtml(override)}">` : DEFAULT_APP_ICONS[key];
    });
  }

  function renderSidebarButtons() {
    const toolbar = document.getElementById("iconToolbar");
    if (!toolbar) return;
    const order = normalizedSidebarOrder().filter(key => sidebarButtonVisible(key));
    Object.entries(SIDEBAR_BUTTON_IDS).forEach(([key, id]) => {
      const button = document.getElementById(id);
      if (button && !order.includes(key)) button.remove();
    });
    order.forEach(key => {
      const button = document.getElementById(SIDEBAR_BUTTON_IDS[key]);
      if (!button) return;
      button.dataset.sidebarKey = key;
      button.draggable = true;
      toolbar.appendChild(button);
    });
    wireSidebarDrag(toolbar);
  }

  function normalizedSidebarOrder() {
    const saved = Array.isArray(appState.settings.sidebarButtonOrder) ? appState.settings.sidebarButtonOrder : [];
    const validSaved = saved.filter(key => DEFAULT_SIDEBAR_BUTTON_ORDER.includes(key));
    return [...validSaved, ...DEFAULT_SIDEBAR_BUTTON_ORDER.filter(key => !validSaved.includes(key))];
  }

  function sidebarButtonVisible(key) {
    const visibility = appState.settings.sidebarButtonVisibility || {};
    return visibility[key] !== false;
  }

  function wireSidebarDrag(toolbar) {
    toolbar.querySelectorAll(".iconButton[data-sidebar-key]").forEach(button => {
      button.ondragstart = event => {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", button.dataset.sidebarKey);
        button.classList.add("draggingSidebarButton");
      };
      button.ondragend = () => {
        button.classList.remove("draggingSidebarButton");
      };
      button.ondragover = event => {
        event.preventDefault();
        const dragging = toolbar.querySelector(".draggingSidebarButton");
        if (!dragging || dragging === button) return;
        const rect = button.getBoundingClientRect();
        if (event.clientY > rect.top + rect.height / 2) toolbar.insertBefore(dragging, button.nextSibling);
        else toolbar.insertBefore(dragging, button);
      };
      button.ondrop = event => {
        event.preventDefault();
        saveSidebarOrderFromDom();
      };
    });
    toolbar.ondragover = event => event.preventDefault();
    toolbar.ondrop = event => {
      event.preventDefault();
      saveSidebarOrderFromDom();
    };
  }

  function saveSidebarOrderFromDom() {
    const toolbar = document.getElementById("iconToolbar");
    const visibleOrder = [...toolbar.querySelectorAll(".iconButton[data-sidebar-key]")].map(button => button.dataset.sidebarKey);
    appState.settings.sidebarButtonOrder = [
      ...visibleOrder,
      ...normalizedSidebarOrder().filter(key => !visibleOrder.includes(key))
    ];
    saveState({ type: "sidebar_reorder", summary: "Reordered sidebar buttons" });
  }

  function appIconPreview(key) {
    const override = appState.iconOverrides[key];
    if (override) return `<img class="settingsIconPreview" alt="${escapeHtml(key)}" src="${escapeHtml(override)}">`;
    if (key === "grid_email") return `<span class="emailIcon"></span>`;
    if (key === "grid_notes") return `<span class="noteIcon"></span>`;
    return `<span class="settingsIconFallback">${DEFAULT_APP_ICONS[key] || ""}</span>`;
  }

  function openPopup(row, col, event) {
    closePopup();
    activePopup = { rowId: row.id, colKey: col.key };
    const layer = document.getElementById("popupLayer");
    const mode = appState.settings.popupMode || "floating";
    const target = event?.target?.getBoundingClientRect ? event.target : document.querySelector("#clientGrid tbody td");
    const rect = target?.getBoundingClientRect ? target.getBoundingClientRect() : { left: 160, bottom: 60 };
    const cardClass = mode === "drawer" ? "drawer" : mode === "modal" ? "modal" : "";
    const style = mode === "floating" ? `left:${Math.max(136, Math.min(rect.left, window.innerWidth - 460))}px;top:${Math.max(48, Math.min(rect.bottom + 4, window.innerHeight - 480))}px;` : "";
    layer.innerHTML = `
      <div class="popupBackdrop"></div>
      <div class="popupCard ${cardClass}" style="${style}">
        <div class="popupHeader">
          <div class="popupTitle">${escapeHtml(col.label)} - ${escapeHtml(row.clientName)}</div>
          <button id="popupClose" class="smallButton">Close</button>
        </div>
        <div class="popupBody">${popupBody(row, col)}</div>
      </div>`;
    layer.querySelector(".popupBackdrop").addEventListener("click", closePopup);
    layer.querySelector("#popupClose").addEventListener("click", closePopup);
    wirePopup(row, col);
  }

  function closePopup() {
    activePopup = null;
    const layer = document.getElementById("popupLayer");
    if (layer) layer.innerHTML = "";
  }

  function reopenActivePopup() {
    if (!activePopup) return;
    const row = gridRows.find(item => item.id === activePopup.rowId);
    const col = COLUMN_DEFS.find(item => item.key === activePopup.colKey);
    if (row && col) openPopup(row, col, { target: document.querySelector(`[data-client-id="${cssEscape(row.id)}"] td[data-key="${col.key}"]`) || document.querySelector("#clientGrid tbody td") });
  }

  function cssEscape(value) {
    return window.CSS?.escape ? CSS.escape(value) : String(value).replace(/"/g, '\\"');
  }

  function popupBody(row, col) {
    if (col.key === "status") return statusBody(row);
    if (col.key === "clientName") return clientProfileBody(row);
    if (col.key === "email") return emailBody(row);
    if (col.key === "nextContact") return nextContactBody(row);
    if (col.key === "t12Revenue") return pivotBody(row, "revenue");
    if (col.key === "nna") return pivotBody(row, "nna");
    if (col.key === "availableCash") return availableCashBody(row);
    if (col.key === "totalAssets") return totalAssetsBody(row);
    if (col.key === "assetAllocation") return allocationBody(row);
    if (col.key === "financialPlan") return financialPlanBody(row);
    if (col.key === "onlineActivity") return onlineActivityBody(row);
    if (col.key === "tasks") return tasksBody(row);
    if (col.key === "notes") return notesBody(row, "notes");
    return genericBody(row, col);
  }

  function statusBody(row) {
    const calculated = calculateStatusTier(row);
    return `
      <div class="popupGrid">
        <label>Displayed Status</label><div>${statusIcon(row.status, row.id)} ${escapeHtml(row.status)}</div>
        <label>Calculated Status</label><div>${statusIcon(calculated.tier, row.id)} ${escapeHtml(calculated.tier)}</div>
        <label>Rule</label><div>${escapeHtml(calculated.reason)}</div>
      </div>
      ${notesEditor(row, "status")}`;
  }

  function clientProfileBody(row) {
    const override = appState.clientOverrides[row.id] || {};
    const family = override.familyMembers || [];
    const activity = appState.activityLog.filter(entry => entry.clientGroupId === row.id).slice(0, 20);
    return `
      <div class="popupGrid">
        <label>Client Group ID</label><div>${escapeHtml(row.id)}</div>
        <label>Parent Client Group ID</label><input data-profile="parentClientGroupId" value="${escapeHtml(override.parentClientGroupId || "")}">
        <label>Risk Profile</label>${selectHtml("riskProfile", override.riskProfile || row.riskProfile || "", ["", ...RISK_PROFILES])}
        <label>Include Alts?</label>${selectHtml("includeAlts", override.includeAlts || row.includeAlts || "No", ["Yes", "No"])}
        <label>Include Munis?</label>${selectHtml("includeMunis", override.includeMunis || row.includeMunis || "No", ["Yes", "No"])}
        <label>Total Net Worth</label><div>${escapeHtml(formatMoney(row.totalNetWorth))}</div>
        <label>Annual Income</label><div>${escapeHtml(formatMoney(row.annualIncome))}</div>
        <label>Minimum Available Cash</label><input data-profile="minAvailableCash" type="number" value="${Number(row.minAvailableCash || 0)}">
        <label>Maximum Available Cash</label><input data-profile="maxAvailableCash" type="number" value="${Number(row.maxAvailableCash || 100000)}">
        <label>Number of Reviews</label>${selectHtml("numberOfReviews", String(row.reviewFrequencyCode ?? 0), ["0", "1", "2"])}
        <label>Review Month</label>${selectHtml("reviewMonth", row.reviewMonth || "", ["", ...MONTHS])}
        <label>Contact Frequency</label>${selectHtml("contactFrequency", String(row.contactFrequencyCode ?? 0), ["0", "1", "4", "6", "12"], ["0 - Request only", "1 - Annual", "4 - Quarterly", "6 - Bi-monthly", "12 - Monthly"])}
        <label>Relationship Notes</label><textarea data-profile="relationshipNotes">${escapeHtml(override.relationshipNotes || "")}</textarea>
        <label>Family / Manager</label><input data-profile="familyManager" value="${escapeHtml(override.familyManager || "")}">
      </div>
      <div class="sectionTitle">Family / Relationships</div>
      <div class="inlineRow"><input id="familyNameInput" placeholder="Name or relationship"><button id="addFamilyMember" class="smallButton">Add</button></div>
      ${family.map(member => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(member.name)}</span><button class="smallButton deleteFamily" data-family-id="${escapeHtml(member.id)}">Archive</button></div></div>`).join("")}
      <div class="sectionTitle">Activity History</div>
      ${activity.map(entry => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(new Date(entry.timestamp).toLocaleString())}</span><span>${escapeHtml(entry.type)}</span></div><div>${escapeHtml(entry.summary)}</div></div>`).join("") || `<div class="cellMuted">No activity yet.</div>`}
      ${notesEditor(row, "clientName")}`;
  }

  function selectHtml(key, selected, values, labels = null) {
    return `<select data-profile="${escapeHtml(key)}">${values.map((value, idx) => `<option value="${escapeHtml(value)}" ${String(selected) === String(value) ? "selected" : ""}>${escapeHtml(labels ? labels[idx] : value || "N/A")}</option>`).join("")}</select>`;
  }

  function emailBody(row) {
    const template = chooseEmailTemplate(row);
    const rendered = renderTemplate(template, row);
    return `
      <div class="popupGrid">
        <label>To</label><div>${escapeHtml(row.email || "N/A")}</div>
        <label>Template</label><select id="emailTemplateSelect">${appState.emailTemplates.map(t => `<option value="${escapeHtml(t.id)}" ${t.id === template.id ? "selected" : ""}>${escapeHtml(t.name)}</option>`).join("")}</select>
        <label>Subject</label><input id="emailSubject" value="${escapeHtml(rendered.subject)}">
        <label>Body</label><textarea id="emailBodyText">${escapeHtml(rendered.body)}</textarea>
      </div>
      <div class="sectionTitle"></div>
      <button id="openEmailDraft" class="smallButton">Open Mailto Draft</button>
      <button id="downloadOutlookDraft" class="smallButton">Download Outlook Draft JSON</button>
      <div class="cellMuted sectionHint">Opening a draft does not log contact activity. Sent email logging comes from the Outlook VBA helper export.</div>
      ${notesEditor(row, "email")}`;
  }

  function chooseEmailTemplate(row) {
    let trigger = "default";
    if (row.nextContact?.days !== null && row.nextContact.days < 0) trigger = "overdue_contact";
    else if (row.availableCashAlert && Number(row.availableCash || 0) > 0) trigger = "excess_cash";
    else if (row.nextContact?.days !== null && row.nextContact.days <= 14) trigger = "upcoming_review";
    return appState.emailTemplates.find(template => template.trigger === trigger) || appState.emailTemplates.find(template => template.trigger === "default") || appState.emailTemplates[0];
  }

  function renderTemplate(template, row) {
    const vars = {
      clientGroupId: row.id,
      clientName: row.clientName,
      salutation: row.salutation || row.clientName,
      t12Revenue: formatMoney(row.t12Revenue),
      totalAssets: formatMoney(row.totalAssets),
      availableCash: formatMoney(row.availableCash),
      nextContact: row.nextContact?.display || "N/A"
    };
    const applyVars = text => String(text || "").replace(/\{\{([^}]+)\}\}/g, (_, key) => vars[cleanText(key)] ?? "");
    return { subject: applyVars(template.subject), body: applyVars(template.body) };
  }

  function nextContactBody(row) {
    return `
      <div class="popupGrid">
        <label>Last Contact</label><div>${escapeHtml(formatDate(row.lastContactDate))}</div>
        <label>Next Contact</label><div class="${escapeHtml(row.nextContact?.className || "")}">${escapeHtml(row.nextContact?.display || "N/A")} (${escapeHtml(row.nextContact?.dueDate || "N/A")})</div>
        <label>Contact Frequency</label><div>${escapeHtml(String(row.contactFrequencyCode))}</div>
        <label>Contact Type</label><select id="contactType">${(appState.settings.contactTypes || DEFAULT_CONTACT_TYPES).map(type => `<option>${escapeHtml(type)}</option>`).join("")}</select>
        <label>Date / Time</label><input id="contactDateTime" type="datetime-local" value="${escapeHtml(new Date().toISOString().slice(0, 16))}">
        <label>Notes</label><textarea id="contactNotes"></textarea>
      </div>
      <div class="sectionTitle"></div>
      <button id="saveContactLog" class="smallButton">Save Interaction</button>
      ${row.nextContact?.days < 0 ? `<div class="miniCard cellDanger">Overdue contact alert is active for this client.</div>` : ""}
      ${notesEditor(row, "nextContact")}`;
  }

  function pivotBody(row, type) {
    const summary = type === "revenue" ? row.summaries?.revenue : row.summaries?.nna;
    const rows = summary?.rows || [];
    const fieldOptions = type === "revenue"
      ? ["__productGroup", "__productL1", "__productL2", "__productL3", "Cde Prd", "__accountKey"]
      : ["__activityType", "Activity Type", "Account Type", "Symbol", "__accountKey"];
    const defaultRow = type === "revenue" ? appState.settings.revenuePivotRowField : appState.settings.nnaPivotRowField;
    const rowField = defaultRow === "GROUP" ? "__productGroup" : defaultRow || fieldOptions[0];
    const colField = "__month";
    return `
      <div class="popupGrid">
        <label>Total</label><div>${escapeHtml(formatMoney(summary?.total))}</div>
        <label>Rows</label><select id="pivotRowField">${fieldOptions.map(field => `<option value="${escapeHtml(field)}" ${field === rowField ? "selected" : ""}>${escapeHtml(field.replace(/^__/, ""))}</option>`).join("")}</select>
        <label>Columns</label><select id="pivotColField"><option value="__month">Month</option></select>
        <label>Value</label><div>${type === "revenue" ? escapeHtml(appState.settings.revenueAmountField) : "Amount($)"}</div>
      </div>
      <div id="pivotTableTarget">${pivotTable(rows, rowField, colField)}</div>
      ${rows.length ? "" : `<div class="miniCard cellMuted">No imported ${type === "revenue" ? "revenue" : "activity"} rows matched the T12 rules.</div>`}
      ${notesEditor(row, type === "revenue" ? "t12Revenue" : "nna")}`;
  }

  function pivotTable(rows, rowField, colField) {
    if (!rows.length) return "";
    const rowKeys = [...new Set(rows.map(row => safeText(row[rowField])))].sort();
    const colKeys = [...new Set(rows.map(row => safeText(row[colField])))].sort();
    const matrix = new Map();
    rows.forEach(row => {
      const r = safeText(row[rowField]);
      const c = safeText(row[colField]);
      const key = `${r}||${c}`;
      matrix.set(key, (matrix.get(key) || 0) + Number(row.__amount || 0));
    });
    return `<table class="popupTable"><thead><tr><th>${escapeHtml(rowField.replace(/^__/, ""))}</th>${colKeys.map(col => `<th>${escapeHtml(col)}</th>`).join("")}<th>Total</th></tr></thead><tbody>${rowKeys.map(r => {
      const total = colKeys.reduce((sum, c) => sum + (matrix.get(`${r}||${c}`) || 0), 0);
      return `<tr><td>${escapeHtml(r)}</td>${colKeys.map(c => `<td>${escapeHtml(formatMoney(matrix.get(`${r}||${c}`) || 0))}</td>`).join("")}<td>${escapeHtml(formatMoney(total))}</td></tr>`;
    }).join("")}</tbody></table>`;
  }

  function availableCashBody(row) {
    const summary = row.summaries?.positions || emptyPositionSummary();
    return `
      <div class="popupGrid">
        <label>Raw Total Cash</label><div>${escapeHtml(formatMoney(row.availableCashRaw))}</div>
        <label>Minimum Threshold</label><div>${escapeHtml(formatMoney(row.minAvailableCash))}</div>
        <label>Maximum Threshold</label><div>${escapeHtml(formatMoney(row.maxAvailableCash))}</div>
        <label>Excess Cash</label><div class="${row.availableCashAlert ? "cellDanger" : ""}">${escapeHtml(formatMoney(row.availableCash))}</div>
        <label>Percent of Assets</label><div>${escapeHtml(formatPercent(row.availableCashPercent))}</div>
      </div>
      <div class="sectionTitle">Cash Breakdown</div>
      ${objectBreakdownTable(summary.cashBreakdown, ["taxable", "ira", "total"])}
      <div class="sectionTitle">Maturing Securities Next 30 Days</div>
      ${simpleRowsTable(summary.maturityList, ["securityName", "accountKey", "date", "days", "marketValue"])}
      ${notesEditor(row, "availableCash")}`;
  }

  function totalAssetsBody(row) {
    const summary = row.summaries?.positions || emptyPositionSummary();
    return `
      <div class="popupGrid">
        <label>Morgan Stanley Total</label><div>${escapeHtml(formatMoney(summary.msAssets))}</div>
        <label>External Total</label><div>${escapeHtml(formatMoney(summary.externalAssets))}</div>
        <label>Total Assets</label><div>${escapeHtml(formatMoney(row.totalAssets))}</div>
        <label>Advisory Assets</label><div>${escapeHtml(formatMoney(row.advisoryAssets))}</div>
        <label>Advisory Percent</label><div>${escapeHtml(formatPercent(row.advisoryPercent))}</div>
      </div>
      <div class="sectionTitle">Advisory Strategy Breakdown</div>
      ${numberBreakdownTable(summary.strategyBreakdown)}
      <div class="sectionTitle">Custodian / Platform</div>
      ${numberBreakdownTable(summary.custodianBreakdown)}
      <div class="sectionTitle">Product Category</div>
      ${numberBreakdownTable(summary.categoryBreakdown)}
      ${notesEditor(row, "totalAssets")}`;
  }

  function advisoryAssetsBody(row) {
    const summary = row.summaries?.positions || emptyPositionSummary();
    return `
      <div class="popupGrid">
        <label>Advisory Assets</label><div>${escapeHtml(formatMoney(row.advisoryAssets))}</div>
        <label>Percent of Assets</label><div>${escapeHtml(formatPercent(row.advisoryPercent))}</div>
        <label>Conversion Opportunity</label><div class="cellMuted">Placeholder for unmanaged-to-advisory opportunity rules.</div>
      </div>
      <div class="sectionTitle">Strategy Breakdown</div>
      ${numberBreakdownTable(summary.strategyBreakdown)}
      ${notesEditor(row, "advisoryAssets")}`;
  }

  function allocationBody(row) {
    const rows = row.summaries?.positions?.allocationRows || [];
    return `
      <div class="popupGrid">
        <label>Status</label><div class="${row.assetAllocation === "DRIFT" ? "cellDanger" : row.assetAllocation === "OK" ? "cellSuccess" : "cellMuted"}">${escapeHtml(row.assetAllocation)}</div>
        <label>Risk Profile</label><div>${escapeHtml(row.riskProfile || "N/A")}</div>
        <label>Target Set</label><div>${escapeHtml(row.includeAlts === "Yes" ? "With alternatives" : "Without alternatives")} / Munis ${escapeHtml(row.includeMunis || "No")}</div>
      </div>
      <table class="popupTable"><thead><tr><th>Asset Class</th><th>Current</th><th>Target</th><th>Drift</th><th>Threshold</th><th>Status</th></tr></thead><tbody>${rows.map(item => `<tr><td>${escapeHtml(item.assetClass)}</td><td>${escapeHtml(formatPercent(item.currentPercent))}</td><td>${escapeHtml(formatPercent(item.targetPercent))}</td><td>${escapeHtml(formatPercent(item.driftPercent))}</td><td>${escapeHtml(formatPercent(item.thresholdPercent))}</td><td class="${item.status === "DRIFT" ? "cellDanger" : "cellSuccess"}">${escapeHtml(item.status)}</td></tr>`).join("")}</tbody></table>
      <div class="sectionTitle">Single Stock Alerts</div>
      ${simpleRowsTable(row.summaries?.positions?.singleStockAlerts || [], ["securityName", "value", "percent"])}
      ${notesEditor(row, "assetAllocation")}`;
  }

  function financialPlanBody(row) {
    const summary = row.summaries?.financial || {};
    return `
      <div class="popupGrid">
        <label>Likelihood of Success</label><div>${escapeHtml(row.financialPlan || "N/A")}</div>
        <label>GPS Status</label><div>${escapeHtml(summary.status || "N/A")}</div>
        <label>Last Plan Update</label><div>${escapeHtml(summary.lastUpdateDate || "N/A")}</div>
        <label>Next Review Date</label><div>${escapeHtml(summary.nextReviewDate || "N/A")}</div>
      </div>
      ${row.financialPlan === "N/A" ? `<div class="miniCard cellMuted">No mapped GPS likelihood available.</div>` : ""}
      ${notesEditor(row, "financialPlan")}`;
  }

  function onlineActivityBody(row) {
    const summary = row.onlineActivity || {};
    return `
      <div class="popupGrid">
        <label>Status</label><div>${summary.className ? `<span class="activityDot ${escapeHtml(summary.className)}"></span>` : `<span class="cellMuted">N/A</span>`}</div>
        <label>Last Login</label><div>${escapeHtml(summary.lastLoginDate || "N/A")}</div>
        <label>Username</label><div>${escapeHtml(summary.username || "N/A")}</div>
      </div>
      ${!summary.lastLoginDate ? `<div class="miniCard cellMuted">Online Activity remains N/A until Settings maps a last-login field.</div>` : ""}
      ${notesEditor(row, "onlineActivity")}`;
  }

  function tasksBody(row) {
    const tasks = appState.tasks.filter(task => task.clientGroupId === row.id && !task.archivedDate);
    return `
      <div class="popupGrid">
        <label>Title</label><input id="taskTitle">
        <label>Due Date</label><input id="taskDueDate" type="date">
        <label>Priority</label><select id="taskPriority">${(appState.settings.taskPriorities || DEFAULT_TASK_PRIORITIES).map(p => `<option ${p === "Normal" ? "selected" : ""}>${escapeHtml(p)}</option>`).join("")}</select>
        <label>Status</label><select id="taskStatus">${(appState.settings.taskStatuses || DEFAULT_TASK_STATUSES).map(s => `<option>${escapeHtml(s)}</option>`).join("")}</select>
        <label>Recurrence</label><input id="taskRecurrence" placeholder="Placeholder">
        <label>Notes</label><textarea id="taskNotes"></textarea>
      </div>
      <div class="sectionTitle"></div>
      <button id="addTaskForClient" class="smallButton">Add Task</button>
      <div class="sectionTitle">Tasks</div>
      ${tasks.map(taskCard).join("") || `<div class="cellMuted">No tasks.</div>`}`;
  }

  function taskCard(task) {
    return `
      <div class="miniCard">
        <div class="miniCardHeader"><span>${escapeHtml(task.priority)} / ${escapeHtml(task.status)} / ${escapeHtml(task.dueDate || "No due date")}</span><span>${escapeHtml(task.source || "manual")}</span></div>
        <strong>${escapeHtml(task.title)}</strong>
        <div>${escapeHtml(task.notes || "")}</div>
        <div class="cardActions">
          <button class="smallButton editTask" data-task-id="${escapeHtml(task.id)}">Edit</button>
          <button class="smallButton completeTask" data-task-id="${escapeHtml(task.id)}">Complete</button>
          <button class="smallButton archiveTask" data-task-id="${escapeHtml(task.id)}">Archive</button>
          <button class="smallButton deleteTask" data-task-id="${escapeHtml(task.id)}">Delete</button>
        </div>
      </div>`;
  }

  function notesBody(row, columnKey) {
    return `${notesEditor(row, columnKey, true)}`;
  }

  function genericBody(row, col) {
    return `
      <div class="popupGrid">
        <label>Value</label><div>${escapeHtml(safeText(row[col.key]))}</div>
        <label>Source</label><div>${escapeHtml(row.source || "snapshot")}</div>
      </div>
      ${notesEditor(row, col.key)}`;
  }

  function notesEditor(row, columnKey, includeAllColumns = false) {
    const notes = appState.notes.filter(note => note.clientGroupId === row.id && (includeAllColumns || note.columnKey === columnKey) && (appState.settings.noteShowArchived || !note.archivedDate));
    return `
      <div class="sectionTitle">Notes</div>
      <div class="inlineRow"><textarea id="newNoteBody" placeholder="Add note"></textarea><button id="addNote" class="smallButton" data-column-key="${escapeHtml(columnKey)}">Add</button></div>
      ${notes.map(note => `
        <div class="miniCard ${note.archivedDate ? "archived" : ""}">
          <div class="miniCardHeader"><span>${note.starred ? "*" : ""} ${escapeHtml(note.columnKey)} / ${escapeHtml(new Date(note.createdDate).toLocaleString())}</span><span>${escapeHtml(note.source || "manual")}</span></div>
          <div>${escapeHtml(note.body)}</div>
          <div class="cardActions">
            <button class="smallButton starNote" data-note-id="${escapeHtml(note.id)}">${note.starred ? "Unstar" : "Star"}</button>
            <button class="smallButton editNote" data-note-id="${escapeHtml(note.id)}">Edit</button>
            <button class="smallButton archiveNote" data-note-id="${escapeHtml(note.id)}">Archive</button>
            <button class="smallButton deleteNote" data-note-id="${escapeHtml(note.id)}">Delete</button>
          </div>
        </div>`).join("") || `<div class="cellMuted">No notes.</div>`}`;
  }

  function objectBreakdownTable(obj, keys) {
    const entries = Object.entries(obj || {});
    if (!entries.length) return `<div class="cellMuted">N/A</div>`;
    return `<table class="popupTable"><thead><tr><th>Type</th>${keys.map(key => `<th>${escapeHtml(key)}</th>`).join("")}</tr></thead><tbody>${entries.map(([name, values]) => `<tr><td>${escapeHtml(name)}</td>${keys.map(key => `<td>${escapeHtml(formatMoney(values[key] || 0))}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  }

  function numberBreakdownTable(obj) {
    const entries = Object.entries(obj || {}).sort((a, b) => b[1] - a[1]);
    if (!entries.length) return `<div class="cellMuted">N/A</div>`;
    return `<table class="popupTable"><thead><tr><th>Name</th><th>Value</th></tr></thead><tbody>${entries.map(([name, value]) => `<tr><td>${escapeHtml(name)}</td><td>${escapeHtml(formatMoney(value))}</td></tr>`).join("")}</tbody></table>`;
  }

  function simpleRowsTable(rows, fields) {
    if (!rows?.length) return `<div class="cellMuted">N/A</div>`;
    return `<table class="popupTable"><thead><tr>${fields.map(field => `<th>${escapeHtml(field)}</th>`).join("")}</tr></thead><tbody>${rows.map(row => `<tr>${fields.map(field => {
      const value = field.toLowerCase().includes("value") ? formatMoney(row[field]) : field.toLowerCase().includes("percent") ? formatPercent(row[field]) : row[field];
      return `<td>${escapeHtml(value ?? "N/A")}</td>`;
    }).join("")}</tr>`).join("")}</tbody></table>`;
  }

  function wirePopup(row, col) {
    document.querySelectorAll("[data-profile]").forEach(input => {
      input.addEventListener("change", () => {
        if (!appState.clientOverrides[row.id]) appState.clientOverrides[row.id] = {};
        const key = input.dataset.profile;
        appState.clientOverrides[row.id][key] = input.value;
        saveState({ type: "profile_edit", clientGroupId: row.id, columnKey: "clientName", summary: `Changed ${key}` });
        render();
      });
    });

    const familyButton = document.getElementById("addFamilyMember");
    if (familyButton) familyButton.addEventListener("click", () => {
      const input = document.getElementById("familyNameInput");
      const name = input.value.trim();
      if (!name) return;
      const override = appState.clientOverrides[row.id] || (appState.clientOverrides[row.id] = {});
      override.familyMembers = override.familyMembers || [];
      override.familyMembers.push({ id: uid("family"), name, createdDate: new Date().toISOString() });
      saveState({ type: "family_add", clientGroupId: row.id, columnKey: "clientName", summary: `Added family/relationship: ${name}` });
      render();
      reopenActivePopup();
    });

    document.querySelectorAll(".deleteFamily").forEach(button => button.addEventListener("click", () => {
      const override = appState.clientOverrides[row.id] || {};
      override.familyMembers = (override.familyMembers || []).filter(member => member.id !== button.dataset.familyId);
      saveState({ type: "family_archive", clientGroupId: row.id, columnKey: "clientName", summary: "Archived family/relationship entry" });
      render();
      reopenActivePopup();
    }));

    const addNote = document.getElementById("addNote");
    if (addNote) addNote.addEventListener("click", () => {
      const body = document.getElementById("newNoteBody").value.trim();
      if (!body) return;
      addNoteRecord(row.id, addNote.dataset.columnKey, body, "manual");
      render();
      reopenActivePopup();
    });

    document.querySelectorAll(".starNote").forEach(button => button.addEventListener("click", () => mutateNote(button.dataset.noteId, note => {
      note.starred = !note.starred;
      note.editedDate = new Date().toISOString();
    }, "note_star", "Toggled note star")));

    document.querySelectorAll(".editNote").forEach(button => button.addEventListener("click", () => {
      const note = appState.notes.find(item => item.id === button.dataset.noteId);
      if (!note) return;
      const next = prompt("Edit note", note.body);
      if (next === null) return;
      mutateNote(note.id, item => {
        item.body = next.trim();
        item.editedDate = new Date().toISOString();
      }, "note_edit", "Edited note");
    }));

    document.querySelectorAll(".archiveNote").forEach(button => button.addEventListener("click", () => mutateNote(button.dataset.noteId, note => {
      note.archivedDate = new Date().toISOString();
    }, "note_archive", "Archived note")));

    document.querySelectorAll(".deleteNote").forEach(button => button.addEventListener("click", () => {
      const note = appState.notes.find(item => item.id === button.dataset.noteId);
      appState.notes = appState.notes.filter(item => item.id !== button.dataset.noteId);
      saveState({ type: "note_delete", clientGroupId: note?.clientGroupId, columnKey: note?.columnKey, summary: "Deleted note" });
      render();
      reopenActivePopup();
    }));

    const templateSelect = document.getElementById("emailTemplateSelect");
    if (templateSelect) templateSelect.addEventListener("change", () => {
      const template = appState.emailTemplates.find(item => item.id === templateSelect.value);
      const rendered = renderTemplate(template, row);
      document.getElementById("emailSubject").value = rendered.subject;
      document.getElementById("emailBodyText").value = rendered.body;
    });

    const openEmail = document.getElementById("openEmailDraft");
    if (openEmail) openEmail.addEventListener("click", () => {
      const href = `mailto:${encodeURIComponent(row.emails.join(";"))}?subject=${encodeURIComponent(document.getElementById("emailSubject").value)}&body=${encodeURIComponent(document.getElementById("emailBodyText").value)}`;
      window.location.href = href;
    });

    const draftDownload = document.getElementById("downloadOutlookDraft");
    if (draftDownload) draftDownload.addEventListener("click", () => {
      const subject = document.getElementById("emailSubject").value;
      const body = document.getElementById("emailBodyText").value;
      downloadJson(`${safeFileStem(appState.settings.outlookDraftFilePrefix || DEFAULT_OUTLOOK_DRAFT_FILE_PREFIX)}_${timestampForFile()}.json`, outlookDraftPayload(row, subject, body));
    });

    const saveContact = document.getElementById("saveContactLog");
    if (saveContact) saveContact.addEventListener("click", () => {
      const type = document.getElementById("contactType").value;
      const dt = document.getElementById("contactDateTime").value || new Date().toISOString();
      const notes = document.getElementById("contactNotes").value.trim();
      const summary = `${type} logged for ${dt}${notes ? `: ${notes}` : ""}`;
      addNoteRecord(row.id, "nextContact", summary, "contact_log", false);
      addActivity({ type: "contact_log", clientGroupId: row.id, columnKey: "nextContact", summary, detail: { contactType: type, contactDate: dt, notes } }, false);
      saveState();
      render();
      closePopup();
    });

    const pivotRow = document.getElementById("pivotRowField");
    const pivotCol = document.getElementById("pivotColField");
    if (pivotRow && pivotCol) {
      const rows = col.key === "t12Revenue" ? row.summaries?.revenue?.rows || [] : row.summaries?.nna?.rows || [];
      const updatePivot = () => document.getElementById("pivotTableTarget").innerHTML = pivotTable(rows, pivotRow.value, pivotCol.value);
      pivotRow.addEventListener("change", updatePivot);
      pivotCol.addEventListener("change", updatePivot);
    }

    const addTask = document.getElementById("addTaskForClient");
    if (addTask) addTask.addEventListener("click", () => {
      const title = document.getElementById("taskTitle").value.trim();
      if (!title) return;
      appState.tasks.unshift({
        id: uid("task"),
        clientGroupId: row.id,
        title,
        dueDate: document.getElementById("taskDueDate").value || null,
        priority: document.getElementById("taskPriority").value,
        status: document.getElementById("taskStatus").value,
        notes: document.getElementById("taskNotes").value.trim(),
        createdDate: new Date().toISOString(),
        completedDate: null,
        recurrence: document.getElementById("taskRecurrence").value.trim() || null,
        archivedDate: null,
        source: "manual"
      });
      saveState({ type: "task_add", clientGroupId: row.id, columnKey: "tasks", summary: `Added task: ${title}` });
      render();
      reopenActivePopup();
    });

    document.querySelectorAll(".completeTask").forEach(button => button.addEventListener("click", () => mutateTask(button.dataset.taskId, task => {
      task.status = "Completed";
      task.completedDate = new Date().toISOString();
      addNoteRecord(task.clientGroupId, "tasks", `Completed task: ${task.title}`, "task_complete", false);
    }, "task_complete", "Completed task")));

    document.querySelectorAll(".archiveTask").forEach(button => button.addEventListener("click", () => mutateTask(button.dataset.taskId, task => {
      task.archivedDate = new Date().toISOString();
    }, "task_archive", "Archived task")));

    document.querySelectorAll(".deleteTask").forEach(button => button.addEventListener("click", () => {
      const task = appState.tasks.find(item => item.id === button.dataset.taskId);
      appState.tasks = appState.tasks.filter(item => item.id !== button.dataset.taskId);
      saveState({ type: "task_delete", clientGroupId: task?.clientGroupId, columnKey: "tasks", summary: `Deleted task: ${task?.title || ""}` });
      render();
      reopenActivePopup();
    }));

    document.querySelectorAll(".editTask").forEach(button => button.addEventListener("click", () => {
      const task = appState.tasks.find(item => item.id === button.dataset.taskId);
      if (!task) return;
      const title = prompt("Edit task title", task.title);
      if (title === null) return;
      mutateTask(task.id, item => {
        item.title = title.trim();
        item.notes = prompt("Edit task notes", item.notes || "") ?? item.notes;
      }, "task_edit", "Edited task");
    }));
  }

  function addNoteRecord(clientGroupId, columnKey, body, source, persist = true) {
    appState.notes.unshift({
      id: uid("note"),
      clientGroupId,
      columnKey,
      body,
      starred: false,
      createdDate: new Date().toISOString(),
      editedDate: null,
      archivedDate: null,
      source
    });
    if (persist) saveState({ type: "note_add", clientGroupId, columnKey, summary: `Added note to ${columnKey}` });
  }

  function mutateNote(noteId, mutator, type, summary) {
    const note = appState.notes.find(item => item.id === noteId);
    if (!note) return;
    mutator(note);
    saveState({ type, clientGroupId: note.clientGroupId, columnKey: note.columnKey, summary });
    render();
    reopenActivePopup();
  }

  function mutateTask(taskId, mutator, type, summary) {
    const task = appState.tasks.find(item => item.id === taskId);
    if (!task) return;
    mutator(task);
    saveState({ type, clientGroupId: task.clientGroupId, columnKey: "tasks", summary: `${summary}: ${task.title}` });
    render();
    reopenActivePopup();
  }

  function addActivity(entry, persist = true) {
    appState.activityLog.unshift({
      id: uid("activity"),
      timestamp: new Date().toISOString(),
      type: entry.type || "event",
      clientGroupId: entry.clientGroupId || null,
      columnKey: entry.columnKey || null,
      summary: entry.summary || "",
      detail: entry.detail || null
    });
    if (persist) saveState();
  }

  function renderSettings() {
    const nav = document.getElementById("settingsNav");
    nav.innerHTML = SETTINGS_SECTIONS.map(section => `<button class="settingsNavButton ${section === activeSettingsSection ? "active" : ""}" data-section="${escapeHtml(section)}">${escapeHtml(section)}</button>`).join("");
    nav.querySelectorAll("button").forEach(button => button.addEventListener("click", () => {
      activeSettingsSection = button.dataset.section;
      renderSettings();
    }));
    document.getElementById("settingsPanel").innerHTML = settingsPanel(activeSettingsSection);
    wireSettings();
  }

  function settingsPanel(section) {
    if (section === "Main Table") return mainTableSettings();
    if (section === "Data Management") return dataManagementSettings();
    if (section === "Status") return statusSettings();
    if (section === "Client Profile") return clientProfileSettings();
    if (section === "Email") return emailSettings();
    if (section === "VBA Outlook Helper") return outlookHelperSettings();
    if (section === "Next Contact") return nextContactSettings();
    if (section === "T12 Revenue") return t12Settings();
    if (section === "NNA") return nnaSettings();
    if (section === "Available Cash") return availableCashSettings();
    if (section === "Total Assets") return totalAssetsSettings();
    if (section === "Advisory Assets") return advisorySettings();
    if (section === "Asset Allocation") return allocationSettings();
    if (section === "Financial Plan") return financialSettings();
    if (section === "Online Activity") return onlineSettings();
    if (section === "Tasks") return taskSettings();
    if (section === "Notes") return noteSettings();
    if (section === "Icons") return iconSettings();
    if (section === "Backups") return backupSettings();
    if (section === "Activity Log") return activityLogSettings();
    if (section === "LLM Export / Import") return llmSettings();
    return "";
  }

  function mainTableSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Main Table</div><div class="popupGrid">
      <label>Popup Mode</label>${settingsSelect("popupMode", appState.settings.popupMode, ["floating", "drawer", "modal"])}
      <label>Corner Radius</label><input data-setting="cornerRadiusPx" type="number" min="0" value="${Number(appState.settings.cornerRadiusPx || 0)}">
      <label>Column Header Background</label><input data-setting="columnHeaderBgColor" type="color" value="${escapeHtml(sanitizeCssColor(appState.settings.columnHeaderBgColor, "#2d2d30"))}">
      <label>Global Font Family</label><input data-setting="globalFontFamily" value="${escapeHtml(appState.settings.globalFontFamily || DEFAULT_GLOBAL_FONT_FAMILY)}">
      <label>Global Font Size</label><input data-setting="globalFontSizePx" type="number" min="9" max="24" value="${clampNumber(appState.settings.globalFontSizePx, 9, 24, DEFAULT_GLOBAL_FONT_SIZE_PX)}">
      <label>Small Font Size</label><input data-setting="smallFontSizePx" type="number" min="8" max="20" value="${clampNumber(appState.settings.smallFontSizePx, 8, 20, DEFAULT_SMALL_FONT_SIZE_PX)}">
      <label>Tooltips</label>${settingsSelect("showTooltips", String(appState.settings.showTooltips), ["true", "false"], ["On", "Off"])}
      <label>Field Source</label>${settingsSelect("showFieldSource", String(appState.settings.showFieldSource), ["true", "false"], ["On", "Off"])}
    </div><div class="sectionTitle">Column Widths</div><div class="popupGrid">${COLUMN_DEFS.map(col => `<label>${escapeHtml(col.label)}</label><input data-column-width="${escapeHtml(col.key)}" type="number" min="48" max="520" value="${columnWidth(col.key)}">`).join("")}</div><div class="sectionTitle"></div><button id="resetLayout" class="smallButton">Reset Layout</button></div>`;
  }

  function dataManagementSettings() {
    const expectedFiles = ensureExpectedFileConfig();
    const statuses = runtime.fileStatuses.length ? runtime.fileStatuses : appState.snapshots.fileStatuses || [];
    const validation = runtime.validation.length ? runtime.validation : appState.snapshots.validation || [];
    const warnings = runtime.warnings || [];
    return `
      <div class="settingsCard">
        <div class="sectionTitle">Data Import</div>
        <div class="popupGrid">
          <label>Source Folder Path</label><input id="sourceFolderPathInput" value="${escapeHtml(appState.settings.sourceFolderPath || "")}" placeholder="Example: C:\\Users\\adamm\\Desktop\\AM_CRM_Data">
          <label>Auto Refresh Saved Folder</label>${settingsSelect("autoRefreshSavedFolder", String(appState.settings.autoRefreshSavedFolder), ["true", "false"], ["On", "Off"])}
        </div>
        <div class="sectionTitle"></div>
        <button id="grantSourceFolder" class="smallButton">Grant Folder Access</button>
        <button id="refreshSavedFolder" class="smallButton">Refresh Saved Folder</button>
        <button id="importFolder" class="smallButton">Import Folder</button>
        <button id="importFiles" class="smallButton">Import Files</button>
        <button id="validateMappings" class="smallButton">Validate Mappings</button>
        <div class="cellMuted sectionHint">${window.XLSX ? "SheetJS loaded locally." : "SheetJS missing. Place lib/xlsx.full.min.js next to index.html."} Browser security requires one folder grant before automatic refresh can read the saved folder.</div>
        ${warnings.map(message => `<div class="miniCard cellWarn">${escapeHtml(message)}</div>`).join("")}
        <table class="popupTable"><thead><tr><th>File</th><th>Header Row</th><th>Status</th><th>Rows</th><th>Message</th></tr></thead><tbody>${statuses.map(s => `<tr><td>${escapeHtml(s.fileName)}</td><td>${escapeHtml(s.headerRow || DEFAULT_WORKBOOK_HEADER_ROW)}</td><td class="${s.status === "loaded" ? "cellSuccess" : s.required ? "cellDanger" : "cellWarn"}">${escapeHtml(s.status)}</td><td>${escapeHtml(s.rowCount ?? "")}</td><td>${escapeHtml(s.message || "")}</td></tr>`).join("")}</tbody></table>
      </div>
      <div class="settingsCard">
        <div class="sectionTitle">Configured Workbook Tables</div>
        <table class="popupTable workbookConfigTable">
          <thead><tr><th>File</th><th>Table Name</th><th>Header Row</th><th>Required</th><th></th></tr></thead>
          <tbody>${expectedFiles.map((item, idx) => `<tr>
            <td><input data-workbook-index="${idx}" data-workbook-field="fileName" value="${escapeHtml(item.fileName)}" ${item.builtIn ? "disabled" : ""}></td>
            <td><input data-workbook-index="${idx}" data-workbook-field="tableName" value="${escapeHtml(item.tableName)}" ${item.builtIn ? "disabled" : ""}></td>
            <td><input data-workbook-index="${idx}" data-workbook-field="headerRow" type="number" min="1" max="500" value="${Number(item.headerRow || DEFAULT_WORKBOOK_HEADER_ROW)}"></td>
            <td><input data-workbook-index="${idx}" data-workbook-field="required" type="checkbox" ${item.required ? "checked" : ""}></td>
            <td>${item.builtIn ? `<span class="cellMuted">Built in</span>` : `<button class="smallButton deleteWorkbookConfig" data-workbook-index="${idx}">Remove</button>`}</td>
          </tr>`).join("")}</tbody>
        </table>
        <div class="sectionTitle">Add Table</div>
        <div class="inlineRow workbookAddRow">
          <input id="newWorkbookFileName" placeholder="data_EXTRA.xlsx">
          <input id="newWorkbookTableName" placeholder="data_EXTRA">
          <input id="newWorkbookHeaderRow" type="number" min="1" max="500" value="${DEFAULT_WORKBOOK_HEADER_ROW}" title="Header row">
          <label class="checkRow"><input id="newWorkbookRequired" type="checkbox"> Required</label>
          <button id="addWorkbookConfig" class="smallButton">Add</button>
        </div>
      </div>
      <div class="settingsCard">
        <div class="sectionTitle">Visual Relationships</div>
        ${relationshipDiagram()}
        ${relationshipCards()}
        <div class="sectionTitle">Add Mapping</div>
        <div class="inlineRow"><input id="newRelLookupTable" placeholder="Lookup table"><input id="newRelLookupField" placeholder="Lookup field/expression"><input id="newRelDetailTable" placeholder="Detail table"><input id="newRelDetailField" placeholder="Detail field/expression"><button id="addRelationship" class="smallButton">Add</button></div>
      </div>
      <div class="settingsCard">
        <div class="sectionTitle">Validation</div>
        ${validation.map(item => `<div class="miniCard ${item.severity === "error" ? "cellDanger" : "cellWarn"}"><strong>${escapeHtml(item.relationshipId)}</strong>: ${escapeHtml(item.message)}</div>`).join("") || `<div class="cellMuted">No validation warnings yet.</div>`}
      </div>
      <div class="settingsCard">
        <div class="sectionTitle">Tables</div>
        ${tableCards()}
      </div>`;
  }

  function relationshipDiagram() {
    const rels = appState.dataManagement.joins || [];
    const height = Math.max(90, rels.length * 22 + 20);
    return `<svg class="relationshipSvg" viewBox="0 0 760 ${height}" role="img" aria-label="Relationship connector overview">${rels.map((rel, idx) => {
      const y = 20 + idx * 22;
      return `<line x1="170" y1="${y}" x2="590" y2="${y}" class="relLine"></line><text x="8" y="${y + 4}" class="relText">${escapeHtml(rel.lookupParentTable)}</text><text x="350" y="${y - 4}" class="relText">${escapeHtml(rel.relationshipId)}</text><text x="604" y="${y + 4}" class="relText">${escapeHtml(rel.detailRelatedTable)}</text>`;
    }).join("")}</svg>`;
  }

  function relationshipCards() {
    return (appState.dataManagement.joins || []).map(rel => `
      <div class="miniCard">
        <div class="miniCardHeader"><strong>${escapeHtml(rel.relationshipId)}</strong><span>${rel.approved ? "approved" : "manual"}</span></div>
        <div>${escapeHtml(rel.lookupParentTable)}: ${escapeHtml(rel.lookupKeyColumnOrExpression)}</div>
        <div>${escapeHtml(rel.detailRelatedTable)}: ${escapeHtml(rel.detailKeyColumnOrExpression)}</div>
        <pre>${escapeHtml(rel.canonicalJoinPredicate || "")}</pre>
        <div class="cardActions">
          <button class="smallButton editRelationship" data-rel-id="${escapeHtml(rel.relationshipId)}">Edit</button>
          <button class="smallButton toggleRelationship" data-rel-id="${escapeHtml(rel.relationshipId)}">${rel.enabled === false ? "Enable" : "Disable"}</button>
          <button class="smallButton deleteRelationship" data-rel-id="${escapeHtml(rel.relationshipId)}">Delete</button>
        </div>
      </div>`).join("");
  }

  function tableCards() {
    const tableNames = new Set([
      ...Object.keys(schemaCatalog.tables || {}),
      ...Object.keys(runtime.importedHeaders || {}),
      ...Object.keys(appState.snapshots.importedTableHeaders || {})
    ]);
    return [...tableNames].sort().map(tableName => {
      const fields = runtime.importedHeaders[tableName] || appState.snapshots.importedTableHeaders?.[tableName] || (schemaCatalog.tables?.[tableName]?.fields || []).map(f => f.fieldName);
      return `<div class="schemaTableCard"><strong>${escapeHtml(tableName)}</strong><div class="cellMuted">${escapeHtml(schemaCatalog.metadata?.[tableName]?.grain || "Workbook/generated table")}</div><div class="schemaFields">${fields.map(field => `<span class="fieldPill">${escapeHtml(field)}</span>`).join("")}</div></div>`;
    }).join("");
  }

  function statusSettings() {
    const rules = appState.settings.statusTierRules || STATUS_TIER_RULES;
    return `<div class="settingsCard"><div class="sectionTitle">Status Tier Rules</div><div class="popupGrid">${["A", "B", "C"].map(tier => `
      <label>${tier} Revenue</label><input data-rule-tier="${tier}" data-rule-field="revenue" type="number" value="${Number(rules[tier]?.revenue || 0)}">
      <label>${tier} Assets</label><input data-rule-tier="${tier}" data-rule-field="assets" type="number" value="${Number(rules[tier]?.assets || 0)}">`).join("")}
    </div><div class="sectionTitle">Calculated Preview</div>${gridRows.slice(0, 12).map(row => `<div class="miniCard">${statusIcon(row.status, row.id)} ${escapeHtml(row.clientName)} -> calculated ${escapeHtml(calculateStatusTier(row).tier)}</div>`).join("")}</div>`;
  }

  function clientProfileSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Client Profile Defaults</div><div class="popupGrid">
      <label>Default Risk Profile</label>${settingsSelect("defaultRiskProfile", appState.settings.defaultRiskProfile, RISK_PROFILES)}
      <label>Default Min Cash</label><input data-setting="defaultMinAvailableCash" type="number" value="${Number(appState.settings.defaultMinAvailableCash || 0)}">
      <label>Default Max Cash</label><input data-setting="defaultMaxAvailableCash" type="number" value="${Number(appState.settings.defaultMaxAvailableCash || 100000)}">
      ${STATUS_ORDER.map(tier => `<label>${tier} Contact Frequency</label><input data-status-default-tier="${tier}" data-status-default-field="contactFrequency" type="number" value="${Number(appState.settings.statusProfileDefaults?.[tier]?.contactFrequency ?? STATUS_PROFILE_DEFAULTS[tier].contactFrequency)}"><label>${tier} Reviews</label><input data-status-default-tier="${tier}" data-status-default-field="numberOfReviews" type="number" value="${Number(appState.settings.statusProfileDefaults?.[tier]?.numberOfReviews ?? STATUS_PROFILE_DEFAULTS[tier].numberOfReviews)}">`).join("")}
    </div></div>`;
  }

  function emailSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Email Templates</div>
      ${appState.emailTemplates.map(template => `<div class="miniCard"><div class="popupGrid"><label>Name</label><input data-template-id="${escapeHtml(template.id)}" data-template-field="name" value="${escapeHtml(template.name)}"><label>Trigger</label><input data-template-id="${escapeHtml(template.id)}" data-template-field="trigger" value="${escapeHtml(template.trigger)}"><label>Subject</label><input data-template-id="${escapeHtml(template.id)}" data-template-field="subject" value="${escapeHtml(template.subject)}"><label>Body</label><textarea data-template-id="${escapeHtml(template.id)}" data-template-field="body">${escapeHtml(template.body)}</textarea></div><button class="smallButton deleteTemplate" data-template-id="${escapeHtml(template.id)}">Delete</button></div>`).join("")}
      <button id="addTemplate" class="smallButton">Add Template</button>
      <pre>Variables: {{clientGroupId}}, {{clientName}}, {{salutation}}, {{t12Revenue}}, {{totalAssets}}, {{availableCash}}, {{nextContact}}</pre>
    </div>`;
  }

  function outlookHelperSettings() {
    const contacts = clientEmailContacts();
    const sentCount = appState.activityLog.filter(entry => entry.type === "sent_email_log").length;
    return `
      <div class="settingsCard">
        <div class="sectionTitle">VBA Outlook Helper</div>
        <div class="popupGrid">
          <label>Max Emails Per Address</label><input data-setting="outlookMaxEmailsPerAddress" type="number" min="1" max="500" value="${Number(appState.settings.outlookMaxEmailsPerAddress || DEFAULT_OUTLOOK_MAX_EMAILS_PER_ADDRESS)}">
          <label>Lookback Days</label><input data-setting="outlookLookbackDays" type="number" min="1" max="3650" value="${Number(appState.settings.outlookLookbackDays || DEFAULT_OUTLOOK_LOOKBACK_DAYS)}">
          <label>Body Preview Characters</label><input data-setting="outlookBodyPreviewChars" type="number" min="0" max="5000" value="${Number(appState.settings.outlookBodyPreviewChars ?? DEFAULT_OUTLOOK_BODY_PREVIEW_CHARS)}">
          <label>Draft File Prefix</label><input data-setting="outlookDraftFilePrefix" value="${escapeHtml(appState.settings.outlookDraftFilePrefix || DEFAULT_OUTLOOK_DRAFT_FILE_PREFIX)}">
          <label>Client Export Filename</label><input data-setting="outlookClientExportFilename" value="${escapeHtml(appState.settings.outlookClientExportFilename || DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME)}">
          <label>Recent Email Filename</label><input data-setting="outlookRecentEmailsFilename" value="${escapeHtml(appState.settings.outlookRecentEmailsFilename || DEFAULT_OUTLOOK_RECENT_EMAILS_FILENAME)}">
          <label>Sent Log Filename</label><input data-setting="outlookSentLogFilename" value="${escapeHtml(appState.settings.outlookSentLogFilename || DEFAULT_OUTLOOK_SENT_LOG_FILENAME)}">
          <label>Task Candidate Filename</label><input data-setting="outlookTaskCandidateFilename" value="${escapeHtml(appState.settings.outlookTaskCandidateFilename || DEFAULT_OUTLOOK_TASK_CANDIDATE_FILENAME)}">
          <label>Sent Contact Type</label><input data-setting="outlookSentContactType" value="${escapeHtml(appState.settings.outlookSentContactType || DEFAULT_OUTLOOK_SENT_CONTACT_TYPE)}">
        </div>
        <div class="sectionTitle">Files</div>
        <div class="cardActions">
          <button id="exportOutlookClients" class="smallButton">Export Client Email JSON</button>
          <button id="exportOutlookConfig" class="smallButton">Export Helper Config JSON</button>
          <a class="smallButton linkButton" href="vba/AM_CRM_Outlook_Helper.bas" download>VBA Module</a>
        </div>
        <div class="sectionTitle">Imports</div>
        <div class="cardActions">
          <button id="importOutlookRecentEmails" class="smallButton">Import Recent Email JSON</button>
          <button id="importOutlookSentLog" class="smallButton">Import Sent Log JSON</button>
          <button id="importOutlookTasks" class="smallButton">Import Task JSON</button>
          <button id="clearOutlookEmailHistory" class="smallButton">Clear Recent History</button>
        </div>
        <input id="outlookRecentEmailsInput" class="hidden" type="file" accept=".json,application/json">
        <input id="outlookSentLogInput" class="hidden" type="file" accept=".json,application/json">
        <input id="outlookTasksInput" class="hidden" type="file" accept=".json,application/json">
        <div class="sectionTitle">Status</div>
        <div class="miniCard"><strong>${contacts.length}</strong> client email addresses ready for export.</div>
        <div class="miniCard"><strong>${appState.outlookEmailHistory.length}</strong> recent Outlook emails stored.</div>
        <div class="miniCard"><strong>${sentCount}</strong> sent Outlook emails counted as client contact.</div>
        ${outlookImportHistoryCards()}
        ${outlookEmailHistoryCards()}
      </div>`;
  }

  function nextContactSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Cadence Days</div><div class="popupGrid">${Object.keys(DEFAULT_CONTACT_CADENCE_DAYS).map(code => `<label>${code}</label><input data-cadence-code="${code}" type="number" value="${Number(appState.settings.contactCadenceDays?.[code] ?? DEFAULT_CONTACT_CADENCE_DAYS[code])}">`).join("")}</div><div class="sectionTitle">Contact Types</div>${listEditor("contactTypes", appState.settings.contactTypes || DEFAULT_CONTACT_TYPES)}</div>`;
  }

  function t12Settings() {
    return `<div class="settingsCard"><div class="sectionTitle">T12 Revenue</div><div class="popupGrid"><label>Amount Field</label>${settingsSelect("revenueAmountField", appState.settings.revenueAmountField, ["Amt Cmp Grs", "Amt Pyout"])}<label>Default Row Field</label>${settingsSelect("revenuePivotRowField", appState.settings.revenuePivotRowField, ["GROUP", "PRODUCT_L1", "PRODUCT_L2", "PRODUCT_L3", "Cde Prd"])}<label>Default Column Field</label>${settingsSelect("revenuePivotColumnField", appState.settings.revenuePivotColumnField, ["Month"])}</div></div>`;
  }

  function nnaSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">NNA</div><div class="popupGrid"><label>Default Row Field</label>${settingsSelect("nnaPivotRowField", appState.settings.nnaPivotRowField, ["Activity Type", "Account Type", "Symbol"])}<label>Default Column Field</label>${settingsSelect("nnaPivotColumnField", appState.settings.nnaPivotColumnField, ["Month"])}</div><div class="sectionTitle">Excluded Keywords</div>${listEditor("nnaExcludedKeywords", appState.settings.nnaExcludedKeywords || ["dividend"])}</div>`;
  }

  function availableCashSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Available Cash</div><div class="popupGrid"><label>Include Held-Away Cash</label>${settingsSelect("includeHeldAwayCash", String(appState.settings.includeHeldAwayCash), ["true", "false"], ["Yes", "No"])}<label>Default Min</label><input data-setting="defaultMinAvailableCash" type="number" value="${Number(appState.settings.defaultMinAvailableCash || 0)}"><label>Default Max</label><input data-setting="defaultMaxAvailableCash" type="number" value="${Number(appState.settings.defaultMaxAvailableCash || 100000)}"></div><div class="sectionTitle">Cash Keywords</div>${listEditor("cashKeywords", appState.settings.cashKeywords || DEFAULT_SETTINGS.cashKeywords)}</div>`;
  }

  function totalAssetsSettings() {
    const categories = uniquePositionCategories();
    return `<div class="settingsCard"><div class="sectionTitle">Product Category Toggles</div>${categories.map(cat => `<label class="checkRow"><input type="checkbox" data-category-toggle="${escapeHtml(cat)}" ${isCategoryDisabled(cat) ? "" : "checked"}> ${escapeHtml(cat)}</label>`).join("") || `<div class="cellMuted">Import positions to see categories.</div>`}</div>`;
  }

  function advisorySettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Advisory Assets</div><div class="popupGrid"><label>Strategy Mapping</label><textarea data-json-setting="strategyMapping">${escapeHtml(JSON.stringify(appState.settings.strategyMapping || {}, null, 2))}</textarea><label>Conversion Logic</label><div class="cellMuted">Placeholder enabled in Advisory Assets popup.</div></div></div>`;
  }

  function allocationSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Asset Allocation</div><div class="popupGrid">${Object.keys(appState.settings.driftThresholds || {}).map(key => `<label>${escapeHtml(key)} Threshold</label><input data-drift-class="${escapeHtml(key)}" type="number" value="${Number(appState.settings.driftThresholds[key])}">`).join("")}<label>Asset Class Mapping JSON</label><textarea data-json-setting="assetClassMappings">${escapeHtml(JSON.stringify(appState.settings.assetClassMappings || {}, null, 2))}</textarea></div><div class="sectionTitle">Target File Mapping</div><div class="popupGrid">${Object.entries(appState.settings.allocationTargetMapping || {}).map(([key, value]) => `<label>${escapeHtml(key)}</label><input data-allocation-map="${escapeHtml(key)}" value="${escapeHtml(value)}">`).join("")}</div></div>`;
  }

  function financialSettings() {
    return mappingSettings("Financial Plan", "financialPlanMapping", appState.settings.financialPlanMapping);
  }

  function onlineSettings() {
    return mappingSettings("Online Activity", "onlineActivityMapping", appState.settings.onlineActivityMapping);
  }

  function mappingSettings(title, key, mapping) {
    return `<div class="settingsCard"><div class="sectionTitle">${escapeHtml(title)} Mapping</div><div class="popupGrid">${Object.entries(mapping || {}).map(([field, value]) => `<label>${escapeHtml(field)}</label><input data-map-setting="${escapeHtml(key)}" data-map-field="${escapeHtml(field)}" value="${escapeHtml(value)}">`).join("")}</div></div>`;
  }

  function taskSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Task Statuses</div>${listEditor("taskStatuses", appState.settings.taskStatuses || DEFAULT_TASK_STATUSES)}<div class="sectionTitle">Task Priorities</div>${listEditor("taskPriorities", appState.settings.taskPriorities || DEFAULT_TASK_PRIORITIES)}<div class="popupGrid"><label>Recurring Task Placeholder</label>${settingsSelect("recurrencePlaceholderEnabled", String(appState.settings.recurrencePlaceholderEnabled), ["true", "false"], ["Enabled", "Disabled"])}</div></div>`;
  }

  function noteSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Notes</div><div class="popupGrid"><label>Show Archived Notes</label>${settingsSelect("noteShowArchived", String(appState.settings.noteShowArchived), ["true", "false"], ["Yes", "No"])}</div></div>`;
  }

  function iconSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Status PNG Overrides</div>${STATUS_ORDER.map(tier => `<div class="miniCard"><div class="miniCardHeader"><span>${tier} ${STATUS_LABELS[tier]}</span>${statusIcon(tier)}</div><input type="file" accept=".png,image/png" data-icon-key="status_${tier}"><button class="smallButton clearIcon" data-icon-key="status_${tier}">Clear</button></div>`).join("")}</div>
      <div class="settingsCard"><div class="sectionTitle">Button and Table Icon PNG Overrides</div>${APP_ICON_LABELS.map(item => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(item.label)}</span>${appIconPreview(item.key)}</div>${SIDEBAR_BUTTON_IDS[item.key] ? `<label class="checkRow"><input type="checkbox" data-sidebar-visible="${escapeHtml(item.key)}" ${sidebarButtonVisible(item.key) ? "checked" : ""}> Show on sidebar</label>` : ""}<input type="file" accept=".png,image/png" data-icon-key="${escapeHtml(item.key)}"><button class="smallButton clearIcon" data-icon-key="${escapeHtml(item.key)}">Clear</button></div>`).join("")}</div>`;
  }

  function backupSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Backups</div><button id="backupNow" class="smallButton">Create JSON Backup</button><button id="restoreBackup" class="smallButton">Restore JSON Backup</button><button id="exportGrid" class="smallButton">Export Grid</button><input id="restoreBackupInput" class="hidden" type="file" accept=".json"><div class="popupGrid"><label>Automatic Backups</label>${settingsSelect("autoBackups", String(appState.settings.autoBackups), ["true", "false"], ["On", "Off"])}</div></div>`;
  }

  function activityLogSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">Activity Log</div>${appState.activityLog.slice(0, ACTIVITY_LOG_VISIBLE_LIMIT).map(entry => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(new Date(entry.timestamp).toLocaleString())} / ${escapeHtml(entry.type)}</span><button class="smallButton deleteActivity" data-activity-id="${escapeHtml(entry.id)}">Delete</button></div><div>${escapeHtml(entry.clientGroupId || "Global")} / ${escapeHtml(entry.columnKey || "")}</div><div>${escapeHtml(entry.summary)}</div></div>`).join("") || `<div class="cellMuted">No activity yet.</div>`}</div>`;
  }

  function llmSettings() {
    return `<div class="settingsCard"><div class="sectionTitle">LLM Export / Import</div><button id="exportLlmJson" class="smallButton">Export LLM JSON</button><button id="importLlmJson" class="smallButton">Import LLM Task JSON</button><input id="importLlmInput" class="hidden" type="file" accept=".json"><pre>${escapeHtml(JSON.stringify(llmExportPayload().clients[0] || {}, null, 2))}</pre></div>`;
  }

  function settingsSelect(key, selected, values, labels = null) {
    return `<select data-setting="${escapeHtml(key)}">${values.map((value, idx) => `<option value="${escapeHtml(value)}" ${String(selected) === String(value) ? "selected" : ""}>${escapeHtml(labels ? labels[idx] : value)}</option>`).join("")}</select>`;
  }

  function listEditor(key, values) {
    return `<div data-list-editor="${escapeHtml(key)}">${values.map((value, idx) => `<div class="inlineRow"><input data-list-key="${escapeHtml(key)}" data-list-index="${idx}" value="${escapeHtml(value)}"><button class="smallButton deleteListItem" data-list-key="${escapeHtml(key)}" data-list-index="${idx}">Delete</button></div>`).join("")}<div class="inlineRow"><input id="new_${escapeHtml(key)}" placeholder="Add value"><button class="smallButton addListItem" data-list-key="${escapeHtml(key)}">Add</button></div></div>`;
  }

  function uniquePositionCategories() {
    const categories = new Set();
    (runtime.importedTables.data_POSITIONS || []).forEach(row => categories.add(classifyProductCategory(row, {})));
    gridRows.forEach(row => Object.keys(row.summaries?.positions?.categoryBreakdown || {}).forEach(cat => categories.add(cat)));
    return [...categories].filter(Boolean).sort();
  }

  function wireSettings() {
    document.querySelectorAll("[data-setting]").forEach(input => input.addEventListener("change", () => {
      let value = input.value;
      if (value === "true") value = true;
      if (value === "false") value = false;
      if (input.type === "number") value = Number(value);
      appState.settings[input.dataset.setting] = value;
      saveState({ type: "settings_edit", summary: `Changed ${input.dataset.setting}` });
      render();
    }));

    document.getElementById("sourceFolderPathInput")?.addEventListener("change", event => {
      appState.settings.sourceFolderPath = event.target.value.trim();
      saveState({ type: "source_folder_path", summary: `Changed source folder path to ${appState.settings.sourceFolderPath}` });
      autoRefreshSavedFolder(true);
    });

    document.getElementById("resetLayout")?.addEventListener("click", () => {
      appState.tableLayout.columns = COLUMN_DEFS.map(col => ({ key: col.key, width: col.width }));
      saveState({ type: "layout_reset", summary: "Reset table layout" });
      render();
    });

    document.querySelectorAll("[data-sidebar-visible]").forEach(input => input.addEventListener("change", () => {
      const key = input.dataset.sidebarVisible;
      appState.settings.sidebarButtonVisibility = {
        ...(appState.settings.sidebarButtonVisibility || {}),
        [key]: input.checked
      };
      saveState({ type: "sidebar_visibility", summary: `${input.checked ? "Showed" : "Hid"} ${key}` });
      render();
    }));

    document.querySelectorAll("[data-column-width]").forEach(input => input.addEventListener("change", () => {
      setColumnWidth(input.dataset.columnWidth, input.value);
      saveState({ type: "layout_width_edit", summary: `Changed ${input.dataset.columnWidth} column width to ${columnWidth(input.dataset.columnWidth)}px` });
      render();
    }));

    document.getElementById("grantSourceFolder")?.addEventListener("click", () => grantSourceFolderAccess());
    document.getElementById("refreshSavedFolder")?.addEventListener("click", () => autoRefreshSavedFolder(true));
    document.getElementById("importFolder")?.addEventListener("click", () => document.getElementById("xlsxFolderInput").click());
    document.getElementById("importFiles")?.addEventListener("click", () => document.getElementById("xlsxFileInput").click());
    document.getElementById("validateMappings")?.addEventListener("click", () => {
      runtime.validation = validateRelationships();
      appState.snapshots.validation = clone(runtime.validation);
      saveState({ type: "validation_run", summary: "Validated data mappings" });
      renderSettings();
    });

    document.querySelectorAll("[data-workbook-field]").forEach(input => input.addEventListener("change", () => {
      const configs = ensureExpectedFileConfig();
      const item = configs[Number(input.dataset.workbookIndex)];
      if (!item) return;
      const field = input.dataset.workbookField;
      if (field === "required") {
        item.required = input.checked;
      } else if (field === "headerRow") {
        item.headerRow = Math.round(clampNumber(input.value, 1, 500, defaultHeaderRowForFile(item.fileName)));
      } else {
        item[field] = input.value.trim();
      }
      if (field === "fileName" && !item.tableName) item.tableName = tableNameFromFileName(item.fileName);
      appState.dataManagement.expectedFiles = normalizeExpectedFiles(configs);
      saveState({ type: "workbook_config_edit", summary: `Changed ${item.fileName || "workbook"} ${field}` });
      renderSettings();
    }));

    document.querySelectorAll(".deleteWorkbookConfig").forEach(button => button.addEventListener("click", () => {
      const configs = ensureExpectedFileConfig();
      const item = configs[Number(button.dataset.workbookIndex)];
      if (!item || item.builtIn) return;
      appState.dataManagement.expectedFiles = configs.filter(config => config !== item);
      saveState({ type: "workbook_config_delete", summary: `Removed workbook table ${item.tableName}` });
      renderSettings();
    }));

    document.getElementById("addWorkbookConfig")?.addEventListener("click", () => {
      let fileName = document.getElementById("newWorkbookFileName").value.trim();
      if (!fileName) return;
      if (!/\.xlsx$/i.test(fileName)) fileName = `${fileName}.xlsx`;
      const tableName = document.getElementById("newWorkbookTableName").value.trim() || tableNameFromFileName(fileName);
      const configs = ensureExpectedFileConfig();
      const duplicate = configs.some(item => item.fileName.toLowerCase() === fileName.toLowerCase() || item.tableName.toLowerCase() === tableName.toLowerCase());
      if (duplicate) {
        alert("That file or table name is already configured.");
        return;
      }
      configs.push(normalizeWorkbookConfig({
        fileName,
        tableName,
        headerRow: document.getElementById("newWorkbookHeaderRow").value,
        required: document.getElementById("newWorkbookRequired").checked,
        builtIn: false
      }));
      appState.dataManagement.expectedFiles = normalizeExpectedFiles(configs);
      saveState({ type: "workbook_config_add", summary: `Added workbook table ${tableName}` });
      renderSettings();
    });

    document.getElementById("addRelationship")?.addEventListener("click", () => {
      const id = `U${String(appState.dataManagement.userJoinCounter || 1).padStart(3, "0")}`;
      appState.dataManagement.userJoinCounter = (appState.dataManagement.userJoinCounter || 1) + 1;
      const lookupTable = document.getElementById("newRelLookupTable").value.trim();
      const lookupField = document.getElementById("newRelLookupField").value.trim();
      const detailTable = document.getElementById("newRelDetailTable").value.trim();
      const detailField = document.getElementById("newRelDetailField").value.trim();
      if (!lookupTable || !lookupField || !detailTable || !detailField) return;
      appState.dataManagement.joins.push({
        relationshipId: id,
        lookupParentTable: lookupTable,
        lookupKeyColumnOrExpression: lookupField,
        detailRelatedTable: detailTable,
        detailKeyColumnOrExpression: detailField,
        canonicalJoinPredicate: `${lookupTable}.${lookupField} = ${detailTable}.${detailField}`,
        approved: false,
        enabled: true
      });
      saveState({ type: "relationship_add", summary: `Added relationship ${id}` });
      renderSettings();
    });

    document.querySelectorAll(".toggleRelationship").forEach(button => button.addEventListener("click", () => {
      const rel = appState.dataManagement.joins.find(item => item.relationshipId === button.dataset.relId);
      if (!rel) return;
      rel.enabled = rel.enabled === false;
      saveState({ type: "relationship_toggle", summary: `Toggled relationship ${rel.relationshipId}` });
      renderSettings();
    }));

    document.querySelectorAll(".deleteRelationship").forEach(button => button.addEventListener("click", () => {
      appState.dataManagement.joins = appState.dataManagement.joins.filter(item => item.relationshipId !== button.dataset.relId);
      saveState({ type: "relationship_delete", summary: `Deleted relationship ${button.dataset.relId}` });
      renderSettings();
    }));

    document.querySelectorAll(".editRelationship").forEach(button => button.addEventListener("click", () => {
      const rel = appState.dataManagement.joins.find(item => item.relationshipId === button.dataset.relId);
      if (!rel) return;
      const next = prompt("Canonical predicate", rel.canonicalJoinPredicate || "");
      if (next === null) return;
      rel.canonicalJoinPredicate = next.trim();
      saveState({ type: "relationship_edit", summary: `Edited relationship ${rel.relationshipId}` });
      renderSettings();
    }));

    document.querySelectorAll("[data-rule-tier]").forEach(input => input.addEventListener("change", () => {
      const tier = input.dataset.ruleTier;
      const field = input.dataset.ruleField;
      if (!appState.settings.statusTierRules[tier]) appState.settings.statusTierRules[tier] = {};
      appState.settings.statusTierRules[tier][field] = Number(input.value || 0);
      saveState({ type: "settings_edit", summary: `Changed ${tier} ${field} rule` });
      render();
    }));

    document.querySelectorAll("[data-status-default-tier]").forEach(input => input.addEventListener("change", () => {
      const tier = input.dataset.statusDefaultTier;
      const field = input.dataset.statusDefaultField;
      if (!appState.settings.statusProfileDefaults[tier]) appState.settings.statusProfileDefaults[tier] = {};
      appState.settings.statusProfileDefaults[tier][field] = Number(input.value || 0);
      saveState({ type: "settings_edit", summary: `Changed ${tier} ${field} default` });
      render();
    }));

    document.querySelectorAll("[data-template-id]").forEach(input => input.addEventListener("change", () => {
      const template = appState.emailTemplates.find(item => item.id === input.dataset.templateId);
      if (!template) return;
      template[input.dataset.templateField] = input.value;
      saveState({ type: "email_template_edit", summary: `Edited template ${template.name}` });
    }));

    document.querySelectorAll(".deleteTemplate").forEach(button => button.addEventListener("click", () => {
      appState.emailTemplates = appState.emailTemplates.filter(item => item.id !== button.dataset.templateId);
      saveState({ type: "email_template_delete", summary: "Deleted email template" });
      renderSettings();
    }));

    document.getElementById("addTemplate")?.addEventListener("click", () => {
      appState.emailTemplates.push({ id: uid("template"), name: "New template", trigger: "default", subject: "", body: "" });
      saveState({ type: "email_template_add", summary: "Added email template" });
      renderSettings();
    });

    document.querySelectorAll("[data-cadence-code]").forEach(input => input.addEventListener("change", () => {
      appState.settings.contactCadenceDays[input.dataset.cadenceCode] = Number(input.value || 0);
      saveState({ type: "settings_edit", summary: `Changed cadence ${input.dataset.cadenceCode}` });
      render();
    }));

    document.querySelectorAll("[data-list-key]").forEach(input => {
      if (input.tagName !== "INPUT") return;
      input.addEventListener("change", () => {
        appState.settings[input.dataset.listKey][Number(input.dataset.listIndex)] = input.value;
        saveState({ type: "settings_edit", summary: `Changed ${input.dataset.listKey}` });
      });
    });

    document.querySelectorAll(".addListItem").forEach(button => button.addEventListener("click", () => {
      const key = button.dataset.listKey;
      const input = document.getElementById(`new_${key}`);
      const value = input.value.trim();
      if (!value) return;
      appState.settings[key] = appState.settings[key] || [];
      appState.settings[key].push(value);
      saveState({ type: "settings_edit", summary: `Added ${key} value` });
      renderSettings();
    }));

    document.querySelectorAll(".deleteListItem").forEach(button => button.addEventListener("click", () => {
      const key = button.dataset.listKey;
      appState.settings[key].splice(Number(button.dataset.listIndex), 1);
      saveState({ type: "settings_edit", summary: `Deleted ${key} value` });
      renderSettings();
    }));

    document.querySelectorAll("[data-category-toggle]").forEach(input => input.addEventListener("change", () => {
      const category = input.dataset.categoryToggle;
      const disabled = new Set(appState.settings.totalAssetDisabledCategories || []);
      if (input.checked) disabled.delete(category);
      else disabled.add(category);
      appState.settings.totalAssetDisabledCategories = [...disabled];
      saveState({ type: "settings_edit", summary: `Changed category toggle ${category}` });
      render();
    }));

    document.querySelectorAll("[data-json-setting]").forEach(input => input.addEventListener("change", () => {
      try {
        appState.settings[input.dataset.jsonSetting] = JSON.parse(input.value || "{}");
        saveState({ type: "settings_edit", summary: `Changed ${input.dataset.jsonSetting}` });
        render();
      } catch (err) {
        alert(`Invalid JSON: ${err.message}`);
      }
    }));

    document.querySelectorAll("[data-drift-class]").forEach(input => input.addEventListener("change", () => {
      appState.settings.driftThresholds[input.dataset.driftClass] = Number(input.value || 0);
      saveState({ type: "settings_edit", summary: `Changed ${input.dataset.driftClass} drift threshold` });
      render();
    }));

    document.querySelectorAll("[data-allocation-map]").forEach(input => input.addEventListener("change", () => {
      appState.settings.allocationTargetMapping[input.dataset.allocationMap] = input.value;
      saveState({ type: "settings_edit", summary: `Changed allocation mapping ${input.dataset.allocationMap}` });
    }));

    document.querySelectorAll("[data-map-setting]").forEach(input => input.addEventListener("change", () => {
      appState.settings[input.dataset.mapSetting][input.dataset.mapField] = input.value;
      saveState({ type: "settings_edit", summary: `Changed ${input.dataset.mapSetting}.${input.dataset.mapField}` });
      render();
    }));

    document.querySelectorAll("[data-icon-key]").forEach(input => {
      if (input.type !== "file") return;
      input.addEventListener("change", async () => {
        const file = input.files[0];
        if (!file) return;
        appState.iconOverrides[input.dataset.iconKey] = await fileToDataUrl(file);
        saveState({ type: "icon_override", summary: `Updated ${input.dataset.iconKey}` });
        render();
      });
    });

    document.querySelectorAll(".clearIcon").forEach(button => button.addEventListener("click", () => {
      delete appState.iconOverrides[button.dataset.iconKey];
      saveState({ type: "icon_clear", summary: `Cleared ${button.dataset.iconKey}` });
      render();
    }));

    document.getElementById("backupNow")?.addEventListener("click", downloadBackup);
    document.getElementById("exportGrid")?.addEventListener("click", exportGrid);
    document.getElementById("restoreBackup")?.addEventListener("click", () => document.getElementById("restoreBackupInput").click());
    document.getElementById("restoreBackupInput")?.addEventListener("change", restoreBackupFromInput);
    document.getElementById("exportOutlookClients")?.addEventListener("click", () => downloadJson(safeDownloadFilename(appState.settings.outlookClientExportFilename || DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME, DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME), outlookClientExportPayload()));
    document.getElementById("exportOutlookConfig")?.addEventListener("click", () => downloadJson(`AM_CRM_Outlook_Helper_Config_${timestampForFile()}.json`, outlookHelperConfigPayload()));
    document.getElementById("importOutlookRecentEmails")?.addEventListener("click", () => document.getElementById("outlookRecentEmailsInput").click());
    document.getElementById("importOutlookSentLog")?.addEventListener("click", () => document.getElementById("outlookSentLogInput").click());
    document.getElementById("importOutlookTasks")?.addEventListener("click", () => document.getElementById("outlookTasksInput").click());
    document.getElementById("clearOutlookEmailHistory")?.addEventListener("click", () => {
      appState.outlookEmailHistory = [];
      saveState({ type: "outlook_recent_clear", summary: "Cleared Outlook recent email history" });
      renderSettings();
    });
    document.getElementById("outlookRecentEmailsInput")?.addEventListener("change", importOutlookRecentEmailsFromInput);
    document.getElementById("outlookSentLogInput")?.addEventListener("change", importOutlookSentLogFromInput);
    document.getElementById("outlookTasksInput")?.addEventListener("change", importOutlookTasksFromInput);
    document.getElementById("exportLlmJson")?.addEventListener("click", () => downloadJson(`AM_CRM_LLM_EXPORT_${timestampForFile()}.json`, llmExportPayload()));
    document.getElementById("importLlmJson")?.addEventListener("click", () => document.getElementById("importLlmInput").click());
    document.getElementById("importLlmInput")?.addEventListener("change", importLlmFromInput);

    document.querySelectorAll(".deleteActivity").forEach(button => button.addEventListener("click", () => {
      appState.activityLog = appState.activityLog.filter(item => item.id !== button.dataset.activityId);
      saveState();
      renderSettings();
    }));
  }

  function fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }

  async function restoreBackupFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const payload = JSON.parse(await file.text());
    appState = mergeState(defaultState(), payload.appState || payload);
    appState.restoredAt = new Date().toISOString();
    runtime = emptyRuntime();
    saveState({ type: "backup_restore", summary: `Restored backup ${file.name}` });
    render();
  }

  async function importOutlookRecentEmailsFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const payload = JSON.parse(await file.text());
    const imported = mergeOutlookEmailHistory(outlookEmailItemsFromPayload(payload));
    rememberOutlookImport("recent_email_history", file.name, imported);
    saveState({ type: "outlook_recent_import", summary: `Imported ${imported} Outlook email history items from ${file.name}` });
    render();
  }

  async function importOutlookSentLogFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const payload = JSON.parse(await file.text());
    const imported = importOutlookSentEmailLog(payload);
    rememberOutlookImport("sent_email_log", file.name, imported);
    saveState({ type: "outlook_sent_log_import", summary: `Imported ${imported} sent Outlook email contacts from ${file.name}` });
    render();
  }

  async function importOutlookTasksFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const payload = JSON.parse(await file.text());
    const count = importLlmTasks({ ...payload, source: payload.source || "outlook_vba_helper" });
    rememberOutlookImport("task_candidates", file.name, count);
    saveState({ type: "outlook_task_import", summary: `Imported ${count} Outlook task candidates from ${file.name}` });
    render();
  }

  function outlookDraftPayload(row, subject, body) {
    return {
      exportType: "AM_CRM_OUTLOOK_DRAFT",
      generatedAt: new Date().toISOString(),
      appVersion: APP_VERSION,
      source: "AM_CRM local browser app",
      clientGroupId: row.id,
      clientName: row.clientName,
      to: row.emails || [],
      subject,
      body
    };
  }

  function outlookClientExportPayload() {
    return {
      exportType: "AM_CRM_OUTLOOK_CLIENT_EMAILS",
      generatedAt: new Date().toISOString(),
      appVersion: APP_VERSION,
      source: "AM_CRM local browser app",
      settings: outlookHelperConfig(),
      contacts: clientEmailContacts()
    };
  }

  function outlookHelperConfigPayload() {
    return {
      exportType: "AM_CRM_OUTLOOK_HELPER_CONFIG",
      generatedAt: new Date().toISOString(),
      appVersion: APP_VERSION,
      settings: outlookHelperConfig(),
      macros: {
        openDraftFromJson: "AMCRM_OpenDraftFromJsonFile",
        exportRecentEmails: "AMCRM_ExportRecentEmailsFromClientJson",
        exportSentEmailLog: "AMCRM_ExportSentEmailLogFromClientJson",
        exportTaskCandidates: "AMCRM_ExportTaskCandidatesFromClientJson"
      }
    };
  }

  function outlookHelperConfig() {
    return {
      maxEmailsPerAddress: Number(appState.settings.outlookMaxEmailsPerAddress || DEFAULT_OUTLOOK_MAX_EMAILS_PER_ADDRESS),
      lookbackDays: Number(appState.settings.outlookLookbackDays || DEFAULT_OUTLOOK_LOOKBACK_DAYS),
      bodyPreviewChars: Number(appState.settings.outlookBodyPreviewChars ?? DEFAULT_OUTLOOK_BODY_PREVIEW_CHARS),
      clientExportFilename: appState.settings.outlookClientExportFilename || DEFAULT_OUTLOOK_CLIENT_EXPORT_FILENAME,
      recentEmailsFilename: appState.settings.outlookRecentEmailsFilename || DEFAULT_OUTLOOK_RECENT_EMAILS_FILENAME,
      sentLogFilename: appState.settings.outlookSentLogFilename || DEFAULT_OUTLOOK_SENT_LOG_FILENAME,
      taskCandidateFilename: appState.settings.outlookTaskCandidateFilename || DEFAULT_OUTLOOK_TASK_CANDIDATE_FILENAME,
      draftFilePrefix: appState.settings.outlookDraftFilePrefix || DEFAULT_OUTLOOK_DRAFT_FILE_PREFIX,
      sentContactType: appState.settings.outlookSentContactType || DEFAULT_OUTLOOK_SENT_CONTACT_TYPE
    };
  }

  function clientEmailContacts() {
    const contacts = [];
    const seen = new Set();
    gridRows.forEach(row => {
      (row.emails || splitEmails(row.email)).forEach(email => {
        const key = normalizeEmail(email);
        if (!key || seen.has(key)) return;
        seen.add(key);
        contacts.push({
          clientGroupId: row.id,
          clientName: row.clientName,
          email: cleanText(email),
          salutation: row.salutation || "",
          status: row.status || ""
        });
      });
    });
    return contacts.sort((a, b) => a.clientName.localeCompare(b.clientName) || a.email.localeCompare(b.email));
  }

  function outlookEmailItemsFromPayload(payload) {
    return (payload.sentEmails || payload.emails || payload.recentEmails || [])
      .map(item => {
        const matchedAddress = cleanText(item.matchedAddress || item.email || item.clientEmail || "");
        const clientGroupId = cleanText(item.clientGroupId || findClientIdForEmail(matchedAddress));
        return {
          ...item,
          clientGroupId,
          matchedAddress,
          timestamp: cleanText(item.timestamp || item.sentOn || item.receivedTime || item.contactDate || ""),
          folder: cleanText(item.folder || ""),
          subject: cleanText(item.subject || ""),
          bodyPreview: cleanText(item.bodyPreview || item.preview || "")
        };
      })
      .filter(item => item.clientGroupId || item.matchedAddress || item.subject);
  }

  function mergeOutlookEmailHistory(items) {
    const existingKeys = new Set((appState.outlookEmailHistory || []).map(outlookItemKey));
    let imported = 0;
    items.forEach(item => {
      const key = outlookItemKey(item);
      if (existingKeys.has(key)) return;
      existingKeys.add(key);
      appState.outlookEmailHistory.unshift({
        importedAt: new Date().toISOString(),
        ...item
      });
      imported += 1;
    });
    appState.outlookEmailHistory = (appState.outlookEmailHistory || []).slice(0, 5000);
    return imported;
  }

  function importOutlookSentEmailLog(payload) {
    const items = outlookEmailItemsFromPayload(payload).filter(item => normalizeTextKey(item.folder) === "SENT" || payload.sentEmails);
    const existingKeys = new Set(appState.activityLog.filter(entry => entry.type === "sent_email_log").map(entry => entry.detail?.outlookKey || entry.detail?.outlookEntryId || `${entry.clientGroupId}|${entry.detail?.contactDate}|${entry.summary}`));
    let imported = 0;
    items.forEach(item => {
      const clientGroupId = item.clientGroupId || findClientIdForEmail(item.matchedAddress);
      if (!clientGroupId) return;
      const contactDate = cleanText(item.timestamp) || new Date().toISOString();
      const outlookKey = outlookItemKey({ ...item, clientGroupId, timestamp: contactDate });
      if (existingKeys.has(outlookKey)) return;
      existingKeys.add(outlookKey);
      appState.activityLog.unshift({
        id: uid("activity"),
        timestamp: contactDate,
        type: "sent_email_log",
        clientGroupId,
        columnKey: "email",
        summary: `Sent email: ${item.subject || "(no subject)"}`,
        detail: {
          outlookKey,
          outlookEntryId: item.outlookEntryId || null,
          contactDate,
          contactType: appState.settings.outlookSentContactType || DEFAULT_OUTLOOK_SENT_CONTACT_TYPE,
          matchedAddress: item.matchedAddress || "",
          folder: item.folder || "Sent",
          sender: item.sender || "",
          senderEmail: item.senderEmail || "",
          to: item.to || "",
          subject: item.subject || "",
          bodyPreview: item.bodyPreview || ""
        }
      });
      imported += 1;
    });
    return imported;
  }

  function rememberOutlookImport(type, fileName, count) {
    appState.outlookImportHistory.unshift({
      id: uid("outlook_import"),
      type,
      fileName,
      count,
      importedAt: new Date().toISOString()
    });
    appState.outlookImportHistory = appState.outlookImportHistory.slice(0, 20);
  }

  function outlookImportHistoryCards() {
    const entries = appState.outlookImportHistory || [];
    if (!entries.length) return `<div class="cellMuted">No Outlook helper imports yet.</div>`;
    return entries.slice(0, 6).map(entry => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(new Date(entry.importedAt).toLocaleString())}</span><span>${escapeHtml(entry.type)}</span></div><div>${escapeHtml(entry.fileName)} / ${Number(entry.count || 0)} items</div></div>`).join("");
  }

  function outlookEmailHistoryCards() {
    const entries = appState.outlookEmailHistory || [];
    if (!entries.length) return "";
    return `<div class="sectionTitle">Recent Email History</div>${entries.slice(0, 12).map(entry => `<div class="miniCard"><div class="miniCardHeader"><span>${escapeHtml(entry.clientGroupId || findClientIdForEmail(entry.matchedAddress) || "Unmatched")}</span><span>${escapeHtml(entry.folder || "")}</span></div><div>${escapeHtml(formatDate(entry.timestamp))} / ${escapeHtml(entry.subject || "(no subject)")}</div><div class="cellMuted">${escapeHtml(entry.matchedAddress || "")}</div></div>`).join("")}`;
  }

  function findClientIdForEmail(email) {
    const key = normalizeEmail(email);
    if (!key) return "";
    const row = gridRows.find(item => (item.emails || []).some(rowEmail => normalizeEmail(rowEmail) === key));
    return row?.id || "";
  }

  function outlookItemKey(item) {
    return cleanText(item.outlookEntryId || item.entryId) || [
      item.clientGroupId || findClientIdForEmail(item.matchedAddress),
      normalizeEmail(item.matchedAddress || item.email),
      cleanText(item.timestamp || item.sentOn || item.receivedTime),
      normalizeTextKey(item.folder),
      normalizeTextKey(item.subject)
    ].join("|");
  }

  function normalizeEmail(value) {
    return cleanText(value).toLowerCase();
  }

  async function importLlmFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const payload = JSON.parse(await file.text());
    const count = importLlmTasks(payload);
    saveState({ type: "llm_import", summary: `Imported ${count} LLM tasks from ${file.name}` });
    render();
  }

  function importLlmTasks(payload) {
    let count = 0;
    (payload.tasks || payload.nextBestActions || []).forEach(task => {
      if (!task.clientGroupId || !task.title) return;
      appState.tasks.unshift({
        id: uid("task"),
        clientGroupId: task.clientGroupId,
        title: task.title,
        dueDate: task.dueDate || null,
        priority: task.priority || "Normal",
        status: task.status || "Open",
        notes: task.notes || task.reasoning || "",
        createdDate: new Date().toISOString(),
        completedDate: null,
        recurrence: task.recurrence || null,
        archivedDate: null,
        source: task.source || payload.source || "llm_import",
        sourceColumn: task.sourceColumn || null
      });
      addActivity({ type: "llm_task_import", clientGroupId: task.clientGroupId, columnKey: task.sourceColumn || "tasks", summary: `Imported LLM task: ${task.title}` }, false);
      count += 1;
    });
    return count;
  }

  function llmExportPayload() {
    return {
      exportType: "AM_CRM_LLM_CLIENT_ANALYSIS",
      generatedAt: new Date().toISOString(),
      appVersion: APP_VERSION,
      source: "AM_CRM local browser app",
      clients: gridRows.map(row => ({
        clientGroupId: row.id,
        clientProfile: {
          clientName: row.clientName,
          salutation: row.salutation,
          emails: row.emails,
          riskProfile: row.riskProfile,
          includeAlts: row.includeAlts,
          includeMunis: row.includeMunis,
          contactFrequency: row.contactFrequencyCode,
          reviewFrequency: row.reviewFrequencyCode
        },
        status: {
          importedOrDisplayed: row.status,
          calculated: calculateStatusTier(row).tier
        },
        t12Revenue: row.t12Revenue,
        nna: row.nna,
        availableCash: {
          rawTotal: row.availableCashRaw,
          excess: row.availableCash,
          percentOfAssets: row.availableCashPercent,
          minimum: row.minAvailableCash,
          maximum: row.maxAvailableCash,
          alert: row.availableCashAlert
        },
        totalAssets: row.totalAssets,
        advisoryAssets: {
          value: row.advisoryAssets,
          percentOfAssets: row.advisoryPercent
        },
        assetAllocation: row.summaries?.positions?.allocationRows || [],
        financialPlan: row.summaries?.financial || { display: row.financialPlan },
        onlineActivity: row.onlineActivity || {},
        lastContact: row.lastContactDate,
        nextContact: row.nextContact,
        openTasks: appState.tasks.filter(task => task.clientGroupId === row.id && !task.archivedDate && !["Completed", "Cancelled"].includes(task.status)),
        importantNotes: appState.notes.filter(note => note.clientGroupId === row.id && !note.archivedDate && note.starred),
        notes: appState.notes.filter(note => note.clientGroupId === row.id && !note.archivedDate),
        recentActivityLog: appState.activityLog.filter(entry => entry.clientGroupId === row.id).slice(0, 20),
        columnSpecificAlerts: row.alerts || [],
        suggestedRuleBasedNextBestActions: nextBestActions(row)
      }))
    };
  }

  function nextBestActions(row) {
    const actions = [];
    if (row.nextContact?.days !== null && row.nextContact.days < 0) actions.push({ sourceColumn: "Next Contact", priority: "High", title: "Schedule overdue client contact" });
    if (row.availableCashAlert && Number(row.availableCash || 0) > 0) actions.push({ sourceColumn: "Available Cash", priority: "High", title: "Review excess available cash" });
    if (row.assetAllocation === "DRIFT") actions.push({ sourceColumn: "Asset Allocation", priority: "Normal", title: "Review allocation drift" });
    if (row.financialPlan === "N/A") actions.push({ sourceColumn: "Financial Plan", priority: "Normal", title: "Check financial plan status" });
    return actions;
  }

  function downloadBackup() {
    downloadJson(`AM_CRM_backup_${timestampForFile()}.json`, {
      backupType: "AM_CRM_BACKUP",
      generatedAt: new Date().toISOString(),
      appState: { ...appState, snapshots: clone(appState.snapshots) },
      schemaCatalog
    });
  }

  function exportGrid() {
    const rows = applySort(applySearch(gridRows));
    const data = rows.map(row => {
      const out = {};
      COLUMN_DEFS.forEach(col => {
        out[col.label] = col.key === "nextContact" ? row.nextContact?.display : col.key === "onlineActivity" ? row.onlineActivity?.display || "N/A" : row[col.key] ?? "";
      });
      return out;
    });
    if (window.XLSX) {
      const wb = window.XLSX.utils.book_new();
      const ws = window.XLSX.utils.json_to_sheet(data);
      window.XLSX.utils.book_append_sheet(wb, ws, "AM_CRM");
      window.XLSX.writeFile(wb, `AM_CRM_GRID_EXPORT_${timestampForFile()}.xlsx`);
      return;
    }
    const csvRows = [COLUMN_DEFS.map(col => col.label), ...data.map(obj => COLUMN_DEFS.map(col => obj[col.label]))];
    downloadBlob(`AM_CRM_GRID_EXPORT_${timestampForFile()}.csv`, csvRows.map(row => row.map(csvEscape).join(",")).join("\n"), "text/csv");
  }

  function csvEscape(value) {
    return `"${String(value ?? "").replace(/"/g, '""')}"`;
  }

  function downloadJson(filename, obj) {
    downloadBlob(filename, JSON.stringify(obj, null, 2), "application/json");
  }

  function safeDownloadFilename(value, fallback) {
    const name = cleanText(value || fallback).replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
    return name.toLowerCase().endsWith(".json") ? name : `${name || safeFileStem(fallback)}.json`;
  }

  function safeFileStem(value) {
    return cleanText(value || "download").replace(/[<>:"/\\|?*\x00-\x1F.]/g, "_") || "download";
  }

  function downloadBlob(filename, content, type) {
    const blob = new Blob([content], { type });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function timestampForFile() {
    const date = new Date();
    const pad = value => String(value).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
  }

  function renderLastRefresh() {
    const el = document.getElementById("lastRefresh");
    if (!el) return;
    const refresh = appState.lastRefreshAt ? new Date(appState.lastRefreshAt).toLocaleString() : "never";
    const restored = appState.restoredAt ? ` / restored ${new Date(appState.restoredAt).toLocaleString()}` : "";
    el.textContent = `Last data refresh: ${refresh}${restored}`;
  }

  function switchView(view) {
    document.querySelectorAll(".navButton").forEach(button => button.classList.toggle("active", button.dataset.view === view));
    document.querySelectorAll(".view").forEach(section => section.classList.toggle("active", section.id === `${view}View`));
    if (view === "settings") renderSettings();
  }

  function detectField(rows, words) {
    const headers = Object.keys(rows[0] || {});
    return headers.find(header => words.every(word => normalizeTextKey(header).includes(normalizeTextKey(word)))) || "";
  }

  function detectDateFields(rows, words) {
    const headers = Object.keys(rows[0] || {});
    return headers.filter(header => words.every(word => normalizeTextKey(header).includes(normalizeTextKey(word))) || normalizeTextKey(header).includes("DATE"));
  }

  function detectClientFields(rows) {
    const headers = Object.keys(rows[0] || {});
    return headers.filter(header => ["CLIENT", "SORT NAME", "ID_NAME_ALG", "HOUSEHOLD", "GROUP"].some(word => normalizeTextKey(header).includes(word)));
  }

  function createHiddenInputs() {
    const fileInput = document.createElement("input");
    fileInput.id = "xlsxFileInput";
    fileInput.type = "file";
    fileInput.accept = ".xlsx";
    fileInput.multiple = true;
    fileInput.className = "hidden";
    fileInput.addEventListener("change", () => importSelectedFiles(fileInput.files, "files"));
    document.body.appendChild(fileInput);

    const folderInput = document.createElement("input");
    folderInput.id = "xlsxFolderInput";
    folderInput.type = "file";
    folderInput.accept = ".xlsx";
    folderInput.multiple = true;
    folderInput.setAttribute("webkitdirectory", "");
    folderInput.setAttribute("directory", "");
    folderInput.className = "hidden";
    folderInput.addEventListener("change", () => importSelectedFiles(folderInput.files, "folder"));
    document.body.appendChild(folderInput);
  }

  function init() {
    window.AM_CRM_UTILS = { normalizeTextKey, leftAccountKey, buildAccountKey, toNumber, toDate };
    createHiddenInputs();
    document.querySelectorAll(".navButton").forEach(button => button.addEventListener("click", () => switchView(button.dataset.view)));
    document.getElementById("globalSearch").addEventListener("input", renderGrid);
    document.getElementById("btnBackup").addEventListener("click", downloadBackup);
    document.getElementById("btnExport").addEventListener("click", exportGrid);
    document.getElementById("btnRefresh").addEventListener("click", () => refreshConfiguredSourceFolder());
    document.getElementById("btnAddTask").addEventListener("click", () => {
      const row = gridRows[0];
      if (row) openPopup(row, COLUMN_DEFS.find(col => col.key === "tasks"), { target: document.querySelector("#clientGrid tbody td") });
    });
    document.getElementById("btnAddNote").addEventListener("click", () => {
      const row = gridRows[0];
      if (row) openPopup(row, COLUMN_DEFS.find(col => col.key === "notes"), { target: document.querySelector("#clientGrid tbody td") });
    });
    render();
    autoRefreshSavedFolder();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
