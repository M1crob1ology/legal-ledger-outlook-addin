import * as React from "react";
import {
  Badge,
  Button,
  Card,
  CardHeader,
  Checkbox,
  Divider,
  Dropdown,
  Field,
  Input,
  Option,
  Switch,
  Text,
  makeStyles,
  tokens,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
} from "@fluentui/react-components";

import { Settings24Regular } from "@fluentui/react-icons";

import { llSupabase } from "../services/legalLedgerSupabase";
import { getCurrentEmailBundle } from "../services/outlook";
import { loadAttachmentTree, ScopeType, filterResults, listRecentCases, listRecentClients } from "../services/legalLedgerData";
import { uploadAttachmentFile } from "../services/legalLedgerUpload";
import { I18N, Lang, LANG_KEY, getInitialLang } from "../i18n";



const ORG_ID_KEY = "ll:addin:orgId";
const ORG_NAME_KEY = "ll:addin:orgName";



interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: tokens.spacingHorizontalL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: tokens.spacingHorizontalM,
  },
  headerLeft: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  statusRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  cardBody: {
    padding: tokens.spacingHorizontalL,
    paddingTop: tokens.spacingVerticalM,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  row: {
    display: "flex",
    gap: tokens.spacingHorizontalM,
    flexWrap: "wrap",
    alignItems: "center",
  },
  grow: {
    flexGrow: 1,
  },
  downloadsList: {
    margin: 0,
    paddingLeft: "1.2rem",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },

  subtle: {
    color: tokens.colorNeutralForeground3,
  },
  codeBox: {
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
    overflowX: "auto",
    whiteSpace: "pre-wrap",
    fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
    fontSize: "12px",
  },
});

function isUuid(v: unknown): v is string {
  return typeof v === "string" && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
}

function uniqueUuids(values: unknown[]): string[] {
  const out: string[] = [];
  for (const v of values) {
    if (isUuid(v) && !out.includes(v)) out.push(v);
  }
  return out;
}

const isFolderNode = (n: any) => (n?.type ?? n?.node_type ?? n?.kind) === "folder";



const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  // -----------------------
  // Legal Ledger auth state
  // -----------------------
  const [llEmail, setLlEmail] = React.useState("");
  const [llPassword, setLlPassword] = React.useState("");
  const [llAuthStatus, setLlAuthStatus] = React.useState<string>("");

  const [llUserEmail, setLlUserEmail] = React.useState<string | null>(null);
  const [authChecked, setAuthChecked] = React.useState(false);
  const [orgs, setOrgs] = React.useState<any[]>([]);
  const [selectedOrgId, setSelectedOrgId] = React.useState<string>(() => {
    return localStorage.getItem(ORG_ID_KEY) ?? "";
  });
  const [settingsOpen, setSettingsOpen] = React.useState(false);
  const [selectedOrgName, setSelectedOrgName] = React.useState<string>(() => {
    return localStorage.getItem(ORG_NAME_KEY) ?? "";
  });
  const [folderFilter, setFolderFilter] = React.useState("");
  const [selectedFolderId, setSelectedFolderId] = React.useState<string>("");



  const [lang, setLang] = React.useState<Lang>(() => getInitialLang());
  const s = I18N[lang] ?? I18N.en;


  React.useEffect(() => {
    localStorage.setItem(LANG_KEY, lang);
    document.title = s.appTitle; // updates the tab title too
  }, [lang, s.appTitle]);


  React.useEffect(() => {
    let alive = true;

    (async () => {
      try {
        const { data } = await llSupabase.auth.getSession();
        if (!alive) return;
        setLlUserEmail(data.session?.user?.email ?? null);
      } finally {
        if (alive) setAuthChecked(true);
      }
    })();

    const { data: sub } = llSupabase.auth.onAuthStateChange((_event, session) => {
      setLlUserEmail(session?.user?.email ?? null);
      setAuthChecked(true);
    });

    return () => {
      alive = false;
      sub.subscription.unsubscribe();
    };
  }, []);


  const didInitOrgsRef = React.useRef(false);

  React.useEffect(() => {
    if (!llUserEmail) {
      didInitOrgsRef.current = false;
      return;
    }

    if (didInitOrgsRef.current) return;
    didInitOrgsRef.current = true;

    // Load org list + validate saved org as soon as the add-in opens
    void onLoadMyOrgs({ silent: true });
  }, [llUserEmail]);

  const [sendEml, setSendEml] = React.useState(true);
  const [sendAttachments, setSendAttachments] = React.useState(true);



  async function onLlLogin() {
    try {
      const email = llEmail.trim();
      if (!email || !llPassword) {
        setLlAuthStatus(s.authEnterEmailPassword);
        return;
      }

      setLlAuthStatus(s.authLoggingIn);
      const { data, error } = await llSupabase.auth.signInWithPassword({
        email,
        password: llPassword,
      });

      if (error) throw error;

      setLlUserEmail(data.user?.email ?? email);
      setLlPassword(""); // don't keep password in memory

      // Load orgs right away so the add-in can work immediately
      await onLoadMyOrgs({ silent: true });

      setLlAuthStatus(s.authLoggedIn);
    } catch (e: any) {
      setLlAuthStatus(`${s.errorPrefix}${e?.message ?? String(e)}`);
    }
  }


  async function onLlLogout() {
    await llSupabase.auth.signOut();
    setLlUserEmail(null);
    setOrgs([]);
    setSelectedOrgId("");
    setLlAuthStatus(s.authLoggedOut);
  }

  async function onLoadMyOrgs(opts: { silent?: boolean } = {}) {
    try {
      if (!opts.silent) setLlAuthStatus(s.loadingOrganizations);

      const { data, error } = await llSupabase.rpc("list_my_orgs");
      if (error) throw error;

      const list = Array.isArray(data) ? data : [];
      setOrgs(list);

      // Decide which org to keep:
      // 1) saved org from localStorage
      // 2) current selected org in state
      // 3) fallback to first org in list
      const savedOrgId = localStorage.getItem(ORG_ID_KEY) ?? "";
      const preferredOrgId = savedOrgId || selectedOrgId || "";

      let chosen = list.find((o) => getOrgId(o) === preferredOrgId) ?? null;
      if (!chosen && list.length > 0) chosen = list[0];

      if (chosen) {
        const id = getOrgId(chosen);
        const name = getOrgName(chosen);

        // Only update if changed (prevents loops + flicker)
        if (id !== selectedOrgId) setSelectedOrgId(id);
        if (name !== selectedOrgName) setSelectedOrgName(name);

        // Keep localStorage consistent
        if (localStorage.getItem(ORG_ID_KEY) !== id) localStorage.setItem(ORG_ID_KEY, id);
        if (localStorage.getItem(ORG_NAME_KEY) !== name) localStorage.setItem(ORG_NAME_KEY, name);
      } else {
        // No orgs available
        if (selectedOrgId) setSelectedOrgId("");
        if (selectedOrgName) setSelectedOrgName("");
        localStorage.removeItem(ORG_ID_KEY);
        localStorage.removeItem(ORG_NAME_KEY);
      }

      if (!opts.silent) {
        setLlAuthStatus(list.length > 0 ? s.loadedOrgs(list.length) : s.noOrganizationsFound);
      }
    } catch (e: any) {
      console.error(e);
      if (!opts.silent) setLlAuthStatus(`${s.errorLoadingOrgsPrefix}${e?.message ?? String(e)}`);
      throw e;
    }
  }


  // -----------------------
  // Destination
  // -----------------------
  const [scopeType, setScopeType] = React.useState<ScopeType>("case");
  const [destQuery, setDestQuery] = React.useState("");
  const [destStatus, setDestStatus] = React.useState("");
  const [destAllResults, setDestAllResults] = React.useState<Array<{ id: string; label: string; raw: any }>>([]);
  const [selectedScopeId, setSelectedScopeId] = React.useState<string>("");

  const [treeStatus, setTreeStatus] = React.useState("");
  const [treeNodes, setTreeNodes] = React.useState<any[]>([]);

  React.useEffect(() => {
    if (!selectedScopeId) return;
    void onLoadFolders(); // auto-load whenever selection changes
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedScopeId, scopeType]);


  const destFiltered = React.useMemo(() => filterResults(destAllResults, destQuery), [destAllResults, destQuery]);

  // Re-translate already shown "Loaded ..." texts when language changes
  React.useEffect(() => {
    // Destination status
    if (destStatus && (destStatus.startsWith("Loaded") || destStatus.startsWith("Laddade"))) {
      const next =
        scopeType === "case"
          ? s.loadedRecentCases(destAllResults.length)
          : s.loadedRecentClients(destAllResults.length);

      if (next !== destStatus) setDestStatus(next);
    }

    // Folder status
    if (treeStatus && (treeStatus.startsWith("Loaded") || treeStatus.startsWith("Laddade"))) {
      const folderCount = (treeNodes || []).filter((n: any) => n?.type === "folder").length;
      const next = s.loadedFolders(folderCount);

      if (next !== treeStatus) setTreeStatus(next);
    }
  }, [lang, destStatus, treeStatus, scopeType, destAllResults.length, treeNodes]);


  function persistSelectedOrg(orgId: string, orgName: string) {
    setSelectedOrgId(orgId);
    setSelectedOrgName(orgName);
    localStorage.setItem(ORG_ID_KEY, orgId);
    localStorage.setItem(ORG_NAME_KEY, orgName);
  }


  const selectedScopeLabel = React.useMemo(() => {
    const found = destAllResults.find((r) => r.id === selectedScopeId);
    return found ? found.label : "";
  }, [destAllResults, selectedScopeId]);

  React.useEffect(() => {
    // When you change destination, reset folder/tree UI
    setTreeNodes([]);
    setSelectedFolderId("");
    setTreeStatus("");
  }, [scopeType, selectedScopeId]);

  const selectedFolderLabel = React.useMemo(() => {
    if (!selectedFolderId) return s.root;
    const node = (treeNodes ?? []).find((n: any) => String(n.id) === String(selectedFolderId));
    return node?.name ?? s.root;
  }, [selectedFolderId, treeNodes, lang]);


  async function onLoadRecentDestination() {
    try {
      setDestStatus(s.loadingRecent);
      setDestAllResults([]);
      setSelectedScopeId("");
      setTreeNodes([]);
      setSelectedFolderId("");

      if (!selectedOrgId) throw new Error(s.selectOrganizationFirst);

      const results =
        scopeType === "case"
          ? await listRecentCases({ orgId: selectedOrgId, limit: 200 })
          : await listRecentClients({ orgId: selectedOrgId, limit: 200 });


      setDestAllResults(results);
      setDestStatus(scopeType === "case" ? s.loadedRecentCases(results.length) : s.loadedRecentClients(results.length));
    } catch (e: any) {
      setDestStatus(`${s.errorPrefix}${e?.message ?? String(e)}`);
    }
  }

  function onChangeScopeType(next: ScopeType) {
    // Update the type
    setScopeType(next);

    // Reset destination selection + lists
    setDestQuery("");
    setDestStatus("");
    setDestAllResults([]);
    setSelectedScopeId("");

    // Reset folders UI (this is what fixes the “folder part remains” bug)
    setTreeStatus("");
    setTreeNodes([]);
    setSelectedFolderId("");
    setFolderFilter("");
  }


  const folderNodes = React.useMemo(() => {
    const folders = (treeNodes ?? []).filter((n: any) => (n.kind ?? n.type) === "folder");
    const q = folderFilter.trim().toLowerCase();
    if (!q) return folders;
    return folders.filter((f: any) => String(f.name ?? "").toLowerCase().includes(q));
  }, [treeNodes, folderFilter]);


  const filteredFolderNodes = React.useMemo(() => {
    const q = folderFilter.trim().toLowerCase();
    if (!q) return folderNodes;
    return folderNodes.filter((n: any) => String(n.name ?? "").toLowerCase().includes(q));
  }, [folderNodes, folderFilter]);


  async function onLoadFolders() {
    try {
      setTreeStatus(s.loadingFolders);
      setTreeNodes([]);
      setSelectedFolderId("");

      if (!selectedScopeId) throw new Error(s.selectCaseOrClientFirst);

      // Find the selected row so we can access raw fields (party/client record)
      const selected = destAllResults.find((r) => r.id === selectedScopeId);
      const raw = selected?.raw ?? {};

      // Candidate IDs to try for CLIENT scope.
      // (We try selectedScopeId + common “party/client” id fields that refactors usually leave behind.)
      const candidateScopeIds =
        scopeType === "client"
          ? uniqueUuids([
            selectedScopeId,
            raw.id,
            raw.party_id,
            raw.partyId,
            raw.client_id,
            raw.clientId,
            raw.client_scope_id,
            raw.clientScopeId,
            raw.legacy_client_id,
            raw.legacyClientId,
          ])
          : uniqueUuids([selectedScopeId]);

      let nodes: any[] = [];
      let usedScopeId = candidateScopeIds[0] ?? selectedScopeId;

      for (const id of candidateScopeIds) {
        const attempt = await loadAttachmentTree({ scopeType, scopeId: id });
        // If your RPC returns Root as a node, you may want `attempt.length > 1` here instead of > 0.
        if (Array.isArray(attempt) && attempt.length > 0) {
          nodes = attempt;
          usedScopeId = id;
          break;
        }
      }

      setTreeNodes(nodes);

      setTreeStatus("");


    } catch (e: any) {
      setTreeStatus(`${s.errorPrefix}${e?.message ?? String(e)}`);
    }
  }

  // Auto-load recent when org or type changes (and user is logged in)
  React.useEffect(() => {
    if (!llUserEmail || !selectedOrgId) return;
    setDestQuery("");
    void onLoadRecentDestination();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [llUserEmail, selectedOrgId, scopeType]);



  // -----------------------
  // Bundle
  // -----------------------
  const [bundleStatus, setBundleStatus] = React.useState<string>("");
  const [bundleEmlDownload, setBundleEmlDownload] = React.useState<{ url: string; name: string } | null>(null);
  const [bundleAttachmentDownloads, setBundleAttachmentDownloads] = React.useState<Array<{ url: string; name: string }>>(
    []
  );
  const [includeEml, setIncludeEml] = React.useState(true);
  const [includeAttachments, setIncludeAttachments] = React.useState(true);

  const [uploading, setUploading] = React.useState(false);
  const [uploadStatus, setUploadStatus] = React.useState("");
  const [preparedBundle, setPreparedBundle] = React.useState<{ eml: File; attachments: File[] } | null>(null);

  const [isUploading, setIsUploading] = React.useState(false);

  // Auto-clear success message after 7 seconds (only on success)
  React.useEffect(() => {
    const isSuccess =
      !!uploadStatus &&
      (uploadStatus.startsWith("✅") || uploadStatus.toLowerCase().includes("uploaded"));

    let t: number | undefined;

    if (isSuccess) {
      t = window.setTimeout(() => {
        setUploadStatus("");
      }, 7000);
    }

    // Always return a cleanup function (prevents TS7030)
    return () => {
      if (t !== undefined) window.clearTimeout(t);
    };
  }, [uploadStatus]);


  // If you are not logged in, don't keep showing a remembered org
  React.useEffect(() => {
    // IMPORTANT: wait until we *know* whether the user is logged in
    if (!authChecked) return;

    // If logged in, keep remembered org
    if (llUserEmail) return;

    // Confirmed logged out -> clear remembered org
    localStorage.removeItem(ORG_ID_KEY);
    localStorage.removeItem(ORG_NAME_KEY);
    setSelectedOrgId("");
    setSelectedOrgName("");
    setOrgs([]);
  }, [authChecked, llUserEmail]);





  function revokeIfAny(url: string | undefined | null) {
    if (url) URL.revokeObjectURL(url);
  }

  async function onPrepareBundle() {
    try {
      setBundleStatus(s.preparingBundle);

      // clean up old links
      revokeIfAny(bundleEmlDownload?.url);
      bundleAttachmentDownloads.forEach((d) => revokeIfAny(d.url));

      setBundleEmlDownload(null);
      setBundleAttachmentDownloads([]);

      const bundle = await getCurrentEmailBundle();



      const emlUrl = URL.createObjectURL(bundle.eml);
      setBundleEmlDownload({ url: emlUrl, name: bundle.eml.name });

      const attachmentLinks = bundle.attachments.map((f) => ({
        url: URL.createObjectURL(f),
        name: f.name,
      }));
      setBundleAttachmentDownloads(attachmentLinks);

      setBundleStatus(s.bundleReady(bundle.eml.name, Math.round(bundle.eml.size / 1024), bundle.attachments.length));
    } catch (e: any) {
      setBundleStatus(`${s.errorPrefix}${e?.message ?? String(e)}`);
    }
  }

  async function onUploadBundle() {
    try {
      // Basic guards
      if (!llUserEmail) {
        setUploadStatus(s.pleaseLogInFirst);
        return;
      }
      if (!selectedOrgId) {
        setUploadStatus(s.selectOrganizationFirst);
        return;
      }
      if (!selectedScopeId) {
        setUploadStatus(s.selectCaseOrClientFirst);
        return;
      }
      if (!includeEml && !includeAttachments) {
        setUploadStatus(s.chooseAtLeastOne);
        return;
      }

      setUploading(true);
      setUploadStatus(s.preparingFiles);

      // Ensure we have a bundle
      const b = preparedBundle ?? (await getCurrentEmailBundle());
      if (!preparedBundle) setPreparedBundle(b);

      const filesToSend: File[] = [];
      if (includeEml) filesToSend.push(b.eml);
      if (includeAttachments) filesToSend.push(...b.attachments);

      // IMPORTANT: clients are Parties in Legal Ledger now
      const normalizedScopeType = scopeType === "client" ? "party" : scopeType;

      // Folder: empty string means Root in your UI

      setUploadStatus(s.uploadingTotal(filesToSend.length));

      const uploadScopeType: "case" | "party" = scopeType === "client" ? "party" : "case";

      // Folder: empty string means Root in your UI
      const folderIdOrNull: string | null = selectedFolderId ? selectedFolderId : null;

      let uploaded = 0;

      for (let i = 0; i < filesToSend.length; i++) {
        const f = filesToSend[i];
        setUploadStatus(s.uploadingN(i + 1, filesToSend.length, f.name));

        await uploadAttachmentFile({
          supabase: llSupabase,
          orgId: selectedOrgId,
          scopeType: uploadScopeType,
          scopeId: selectedScopeId,
          parentId: folderIdOrNull,
          file: f,
          customFileName: f.name,
        });

        uploaded++;
      }

      setUploadStatus(s.uploadedOk(uploaded));

      // Refresh attachment folders so the new files show up
      const refreshed = await loadAttachmentTree({ scopeType, scopeId: selectedScopeId });
      setTreeNodes(refreshed);

      setTreeStatus("");




      // If/when upload succeeds:
      // setUploadStatus("Upload complete.");
    } catch (e: any) {
      setUploadStatus(`${s.errorPrefix}${e?.message ?? String(e)}`);
    } finally {
      setUploading(false);
    }
  }


  // -----------------------
  // Debug (hidden by default)
  // -----------------------
  const [showDebug, setShowDebug] = React.useState(false);
  const [mailInfo, setMailInfo] = React.useState<any>(null);
  const [mailError, setMailError] = React.useState<string | null>(null);

  React.useEffect(() => {
    Office.onReady(() => {
      try {
        const item = Office.context?.mailbox?.item as any;
        if (!item) {
          setMailError(s.noMailboxItem);
          return;
        }

        setMailInfo({
          subject: item.subject ?? null,
          from: item.from?.emailAddress ?? null,
          to: item.to?.map((r: any) => r.emailAddress) ?? [],
          cc: item.cc?.map((r: any) => r.emailAddress) ?? [],
          attachmentsCount: item.attachments?.length ?? 0,
          itemId: item.itemId ?? null,
          itemType: item.itemType ?? null,
        });
      } catch (e: any) {
        setMailError(e?.message ?? String(e));
      }
    });
  }, []);

  const isLoggedIn = !!llUserEmail;

  function getOrgId(o: any): string {
    return String(o.org_id ?? o.id ?? o.organization_id ?? "");
  }

  function getOrgName(o: any): string {
    return String(
      o.name ??
      o.org_name ??
      o.organization_name ??
      o.display_name ??
      o.title ??
      getOrgId(o)
    );
  }

  const selectedOrgLabel = React.useMemo(() => {
    const found = orgs.find((o) => getOrgId(o) === selectedOrgId);
    return found ? getOrgName(found) : "";
  }, [orgs, selectedOrgId]);


  async function openSettings() {
    setSettingsOpen(true);

    // Always restore saved org immediately (so UI + queries are consistent)
    const savedOrgId = localStorage.getItem("ll:addin:orgId") ?? "";
    const savedOrgName = localStorage.getItem("ll:addin:orgName") ?? "";

    if (savedOrgId) setSelectedOrgId(savedOrgId);
    if (savedOrgName) setSelectedOrgName(savedOrgName);

    // If logged in, auto-load orgs when opening settings
    if (llUserEmail) {
      try {
        await onLoadMyOrgs({ silent: true });
      } catch (e) {
        console.error(e);
      }
    }
  }

  const canUpload =
    !!llUserEmail &&
    !!selectedOrgId &&
    !!selectedScopeId &&
    (includeEml || includeAttachments) &&
    !uploading;

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.headerRow}>
        <div className={styles.headerLeft}>
          <Text size={600} weight="semibold">
            {props.title || s.appTitle}
          </Text>

          <div className={styles.statusRow}>
            <Badge appearance={isLoggedIn ? "filled" : "outline"} color={isLoggedIn ? "success" : "important"}>
              {isLoggedIn ? s.connected : s.notConnected}
            </Badge>

            {isLoggedIn && (
              <Text size={200} className={styles.subtle}>
                {llUserEmail}
              </Text>
            )}
          </div>
        </div>


      </div>

      {/* Legal Ledger Connection */}
      <Card>
        <CardHeader
          header={<Text weight="semibold">{s.legalLedger}</Text>}
          description={
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <Text size={200}>{s.loggedInAs} {llUserEmail || "-"}</Text>
              <Text size={200}>{s.orgLabel} {selectedOrgName || "-"}</Text>
            </div>
          }

          action={
            <Button
              appearance="subtle"
              icon={<Settings24Regular />}
              onClick={openSettings}
              aria-label={s.settings}
            />
          }
        />
      </Card>


      <>
        <div className={styles.row}>
          <Field label={s.type} className={styles.grow}>
            <Dropdown
              value={scopeType === "case" ? s.case : s.client}
              selectedOptions={[scopeType]}
              onOptionSelect={(_, data) => onChangeScopeType(data.optionValue as ScopeType)}
            >
              <Option value="case">{s.case}</Option>
              <Option value="client">{s.client}</Option>
            </Dropdown>
          </Field>

          <Button onClick={onLoadRecentDestination} disabled={!selectedOrgId}>
            {s.reloadRecent}
          </Button>
          {!selectedOrgId && <Text size={200}>{s.selectOrgFirst}</Text>}


        </div>

        {destStatus && <Text size={200}>{destStatus}</Text>}

        <Field label={s.filter}>
          <Input
            value={destQuery}
            onChange={(e) => setDestQuery((e.target as HTMLInputElement).value)}
            placeholder={scopeType === "case" ? s.filterCasesPlaceholder : s.filterClientsPlaceholder}
          />
        </Field>



        {destFiltered.length > 0 && (
          <Field label={s.select}>
            <Dropdown
              value={selectedScopeLabel}
              selectedOptions={selectedScopeId ? [selectedScopeId] : []}
              onOptionSelect={(_, data) => setSelectedScopeId(String(data.optionValue ?? ""))}
              placeholder={s.selectPlaceholder}
            >
              {destFiltered.slice(0, 25).map((r) => (
                <Option key={r.id} value={r.id}>
                  {r.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
        )}

        {selectedScopeId && (
          <>
            <div className={styles.row}>
              <Text size={200}>
                {treeStatus || (folderNodes.length > 0 ? s.loadedFolders(folderNodes.length) : s.loadedFoldersRootHint)}
              </Text>

            </div>

            <Field label={s.filterFolders}>
              <Input
                value={folderFilter}
                onChange={(e) => setFolderFilter((e.target as HTMLInputElement).value)}
                placeholder={s.filterFoldersPlaceholder}
              />
            </Field>

            <Field label={s.folderOptional}>
              <Dropdown
                value={selectedFolderLabel}
                selectedOptions={selectedFolderId ? [selectedFolderId] : [""]}
                onOptionSelect={(_, data) => setSelectedFolderId(String(data.optionValue ?? ""))}
              >
                <Option value="">{s.root}</Option>

                {filteredFolderNodes.map((n: any) => (
                  <Option key={n.id} value={String(n.id)}>
                    {n.name}
                  </Option>
                ))}
              </Dropdown>
            </Field>

          </>
        )}


      </>



      {/* Bundle */}
      <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 8 }}>
        <Checkbox
          label={s.emailEml}
          checked={includeEml}
          onChange={(_, d) => setIncludeEml(!!d.checked)}
        />
        <Checkbox
          label={s.attachments}
          checked={includeAttachments}
          onChange={(_, d) => setIncludeAttachments(!!d.checked)}
        />
      </div>

      <Card>
        <CardHeader
          header={<Text weight="semibold">{s.uploadToLegalLedger}</Text>}
          description={
            <Text size={200} className={styles.subtle}>
              {s.uploadHelp}
            </Text>
          }
        />


        <Divider />
        <div className={styles.cardBody}>
          <div className={styles.row}>
            <Button appearance="primary" onClick={onUploadBundle} disabled={!canUpload}>
              {s.uploadToLegalLedger}
            </Button>

            {uploadStatus ? <Text size={200}>{uploadStatus}</Text> : null}

          </div>

          {bundleStatus && <Text size={200}>{bundleStatus}</Text>}

          {bundleEmlDownload && (
            <div>
              <Text size={200} weight="semibold">
                {s.email}
              </Text>
              <div>
                <a href={bundleEmlDownload.url} download={bundleEmlDownload.name}>
                  {s.download} {bundleEmlDownload.name}
                </a>
              </div>
            </div>
          )}

          {bundleAttachmentDownloads.length > 0 && (
            <div>
              <Text size={200} weight="semibold">
                {s.attachments}
              </Text>
              <ul className={styles.downloadsList}>
                {bundleAttachmentDownloads.map((d) => (
                  <li key={d.name}>
                    <a href={d.url} download={d.name}>
                      {s.download} {d.name}
                    </a>
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>
      </Card>

      {/* Debug (toggle) */}
      {showDebug && (
        <Card>
          <CardHeader header={<Text weight="semibold">{s.debug}</Text>} />
          <Divider />
          <div className={styles.cardBody}>
            {mailError && <div className={styles.codeBox}>{mailError}</div>}
            {mailInfo && <div className={styles.codeBox}>{JSON.stringify(mailInfo, null, 2)}</div>}
          </div>
        </Card>
      )}

      <Dialog open={settingsOpen} onOpenChange={(_, data) => setSettingsOpen(data.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>{s.settingsTitle}</DialogTitle>

            <DialogContent>
              {!llUserEmail ? (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <>
                    <Text size={200}>{s.notLoggedIn}</Text>

                    <Field label={s.email}>
                      <Input
                        value={llEmail}
                        onChange={(e) => setLlEmail((e.target as HTMLInputElement).value)}
                        placeholder={s.emailPlaceholder}
                      />
                    </Field>

                    <Field label={s.password}>
                      <Input
                        type="password"
                        value={llPassword}
                        onChange={(e) => setLlPassword((e.target as HTMLInputElement).value)}
                        placeholder={s.passwordPlaceholder}
                      />
                    </Field>

                    <Button
                      appearance="primary"
                      onClick={onLlLogin}
                      disabled={!llEmail.trim() || !llPassword}
                    >
                      {s.login}
                    </Button>

                    {llAuthStatus && <Text size={200}>{llAuthStatus}</Text>}
                  </>

                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <Field label={s.organization}>
                    {orgs.length === 0 ? (
                      <Text size={200}>{s.loadingOrganizations}</Text>
                    ) : (
                      <Dropdown
                        value={selectedOrgName || s.selectOrg}
                        selectedOptions={selectedOrgId ? [selectedOrgId] : []}
                        onOptionSelect={(_, data) => {
                          const orgId = String(data.optionValue ?? "");

                          const org = orgs.find((o: any) => String(o.org_id ?? o.id) === orgId);
                          const orgName = String(org?.name ?? org?.org_name ?? org?.title ?? orgId);

                          persistSelectedOrg(orgId, orgName);
                        }}
                        placeholder={s.selectPlaceholder}
                      >
                        {orgs.map((o: any) => {
                          const id = String(o.org_id ?? o.id);
                          const name = String(o.name ?? o.org_name ?? o.title ?? id);
                          return (
                            <Option key={id} value={id}>
                              {name}
                            </Option>
                          );
                        })}
                      </Dropdown>
                    )}
                  </Field>

                  <Button onClick={onLlLogout /* rename if needed */}>{s.logout}</Button>
                </div>
              )}
            </DialogContent>

            <Field label={s.language}>
              <Dropdown
                selectedOptions={[lang]}
                value={lang === "en" ? s.english : s.swedish}
                onOptionSelect={(_, data) => {
                  const v = String(data.optionValue ?? "");
                  if (v === "en" || v === "sv") setLang(v);
                }}
              >
                <Option value="en">{s.english}</Option>
                <Option value="sv">{s.swedish}</Option>
              </Dropdown>
            </Field>


            <DialogActions>
              <Button appearance="primary" onClick={() => setSettingsOpen(false)}>
                {s.done}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>


    </div>
  );
};

export default App;
