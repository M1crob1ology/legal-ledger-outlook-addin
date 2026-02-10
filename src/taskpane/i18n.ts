// src/taskpane/i18n.ts

export type Lang = "en" | "sv";
export const LANG_KEY = "ll:addin:lang";

function isLang(x: any): x is Lang {
  return x === "en" || x === "sv";
}

export function getInitialLang(): Lang {
  const saved = localStorage.getItem(LANG_KEY);
  if (isLang(saved)) return saved;

  // If nothing saved yet, try Outlook display language
  const dl = (globalThis as any).Office?.context?.displayLanguage as string | undefined;
  if (dl && dl.toLowerCase().startsWith("sv")) return "sv";

  return "en";
}

// Helper for English plural
function enPlural(n: number, one: string, many: string) {
  return n === 1 ? one : many;
}

export const I18N = {
  en: {
    // Top / status
    appTitle: "Legal Ledger Outlook Add-in",
    connected: "Connected",
    notConnected: "Not connected",

    // Connection card
    legalLedger: "Legal Ledger",
    loggedInAs: "Logged in as:",
    orgLabel: "Org:",

    // Settings
    settings: "Settings",
    settingsTitle: "Legal Ledger settings",
    language: "Language",
    english: "English",
    swedish: "Svenska",

    // Auth / org loading
    authEnterEmailPassword: "Enter email and password.",
    authLoggingIn: "Logging in…",
    authLoggedIn: "✅ Logged in.",
    authLoggedOut: "Logged out.",
    loadingOrganizations: "Loading organizations…",
    loadedOrgs: (n: number) => `Loaded ${n} org(s).`,
    noOrganizationsFound: "No organizations found.",
    errorLoadingOrgsPrefix: "Error loading orgs: ",
    notLoggedIn: "You are not logged in.",
    login: "Log in",
    logout: "Log out",
    done: "Done",

    // Destination
    type: "Type",
    case: "Case",
    client: "Client",
    reloadRecent: "Reload recent",
    loadingRecent: "Loading recent…",
    selectOrgFirstShort: "Select an org first.",
    loadedRecentCases: (n: number) => `Loaded ${n} recent ${enPlural(n, "case", "cases")}.`,
    loadedRecentClients: (n: number) => `Loaded ${n} recent ${enPlural(n, "client", "clients")}.`,
    filter: "Filter",
    filterCasesPlaceholder: "Type to filter cases…",
    filterClientsPlaceholder: "Type to filter clients…",
    select: "Select",
    selectPlaceholder: "Select…",

    // Folders
    loadingFolders: "Loading folders…",
    loadedFolders: (n: number) => `Loaded ${n} ${enPlural(n, "folder", "folders")}.`,
    loadedFoldersRootHint: "Loaded 0 folder(s). You can upload to the root folder.",
    filterFolders: "Filter folders",
    filterFoldersPlaceholder: "Type to filter folders…",
    folderOptional: "Folder (optional)",
    root: "(Root)",

    // Bundle/upload
    emailBundle: "Email bundle",
    emailBundleDesc: "Prepare a bundle (email.eml + attachments). Upload comes next.",
    emailEml: "Email (.eml)",
    attachments: "Attachments",
    email: "Email",
    password: "Password",
    organization: "Organization",
    emailPlaceholder: "you@company.com",
    passwordPlaceholder: "Password",
    selectOrg: "Select org…",
    uploadToLegalLedger: "Upload to Legal Ledger",
    uploadHelp:
      'Choose options and click "Upload to Legal Ledger" to upload the email and/or attachments to the selected case or client.',
    download: "Download",
    debug: "Debug",
    noMailboxItem: "No mailbox item available. Open a mail message (Read mode) and try again.",
    preparingBundle: "Preparing bundle…",
    bundleReady: (emlName: string, sizeKb: number, attachmentsCount: number) =>
      `Ready: ${emlName} (${sizeKb} KB), ${attachmentsCount} attachment(s)`,

    // Common messages
    errorPrefix: "Error: ",
    selectOrgFirst: "Select an org in settings first.",
    selectOrganizationFirst: "Select an organization first.",
    selectCaseOrClientFirst: "Select a case/client first.",
    preparingFiles: "Preparing files…",
    uploadingN: (i: number, total: number, name: string) => `Uploading ${i}/${total}: ${name}`,
    uploadingTotal: (n: number) => `Uploading ${n} file(s)…`,
    uploadedOk: (n: number) => `✅ Uploaded ${n} file(s) to Legal Ledger.`,
    chooseAtLeastOne: "Choose at least one: Email (.eml) and/or Attachments.",
    pleaseLogInFirst: "Please log in first.",
  },

  sv: {
    // Top / status
    appTitle: "Legal Ledger Outlook-tillägg",
    connected: "Ansluten",
    notConnected: "Inte ansluten",

    // Connection card
    legalLedger: "Legal Ledger",
    loggedInAs: "Inloggad som:",
    orgLabel: "Org:",

    // Settings
    settings: "Inställningar",
    settingsTitle: "Inställningar för Legal Ledger",
    language: "Språk",
    english: "English",
    swedish: "Svenska",

    // Auth / org loading
    authEnterEmailPassword: "Ange e-post och lösenord.",
    authLoggingIn: "Loggar in…",
    authLoggedIn: "✅ Inloggad.",
    authLoggedOut: "Utloggad.",
    loadingOrganizations: "Laddar organisationer…",
    loadedOrgs: (n: number) => `Laddade ${n} organisation${n === 1 ? "" : "er"}.`,
    noOrganizationsFound: "Inga organisationer hittades.",
    errorLoadingOrgsPrefix: "Fel vid laddning av organisationer: ",
    notLoggedIn: "Du är inte inloggad.",
    login: "Logga in",
    logout: "Logga ut",
    done: "Klar",

    // Destination
    type: "Typ",
    case: "Ärende",
    client: "Klient",
    reloadRecent: "Ladda om senaste",
    loadingRecent: "Laddar senaste…",
    selectOrgFirstShort: "Välj en org först.",
    loadedRecentCases: (n: number) => `Laddade ${n} senaste ${n === 1 ? "ärendet" : "ärendena"}.`,
    loadedRecentClients: (n: number) => `Laddade ${n} senaste ${n === 1 ? "klienten" : "klienterna"}.`,
    filter: "Filtrera",
    filterCasesPlaceholder: "Skriv för att filtrera ärenden…",
    filterClientsPlaceholder: "Skriv för att filtrera klienter…",
    select: "Välj",
    selectPlaceholder: "Välj…",

    // Folders
    loadingFolders: "Laddar mappar…",
    loadedFolders: (n: number) => `Laddade ${n} ${n === 1 ? "mapp" : "mappar"}.`,
    loadedFoldersRootHint: "Laddade 0 mappar. Du kan ladda upp till rotmappen.",
    filterFolders: "Filtrera mappar",
    filterFoldersPlaceholder: "Skriv för att filtrera mappar…",
    folderOptional: "Mapp (valfritt)",
    root: "(Rot)",

    // Bundle/upload
    emailBundle: "E-postpaket",
    emailBundleDesc: "Förbered ett paket (email.eml + bilagor). Uppladdning kommer härnäst.",
    emailEml: "E-post (.eml)",
    attachments: "Bilagor",
    email: "E-post",
    password: "Lösenord",
    organization: "Organisation",
    emailPlaceholder: "you@company.com",
    passwordPlaceholder: "Lösenord",
    selectOrg: "Välj org…",
    uploadToLegalLedger: "Ladda upp till Legal Ledger",
    uploadHelp:
      'Välj alternativ och klicka "Ladda upp till Legal Ledger" för att ladda upp e-post och/eller bilagor till valt ärende eller klient.',
    download: "Ladda ner",
    debug: "Felsökning",
    noMailboxItem: "Ingen mailbox-post tillgänglig. Öppna ett mejl (Läsläge) och försök igen.",
    preparingBundle: "Förbereder paket…",
    bundleReady: (emlName: string, sizeKb: number, attachmentsCount: number) =>
      `Redo: ${emlName} (${sizeKb} KB), ${attachmentsCount} ${attachmentsCount === 1 ? "bilaga" : "bilagor"}`,

    // Common messages
    errorPrefix: "Fel: ",
    selectOrgFirst: "Välj en organisation i inställningar först.",
    selectOrganizationFirst: "Välj en organisation först.",
    selectCaseOrClientFirst: "Välj ett ärende/klient först.",
    preparingFiles: "Förbereder filer…",
    uploadingN: (i: number, total: number, name: string) => `Laddar upp ${i}/${total}: ${name}`,
    uploadingTotal: (n: number) => `Laddar upp ${n} fil(er)…`,
    uploadedOk: (n: number) => `✅ Laddade upp ${n} fil(er) till Legal Ledger.`,
    chooseAtLeastOne: "Välj minst en: E-post (.eml) och/eller Bilagor.",
    pleaseLogInFirst: "Logga in först.",
  },
} as const;

