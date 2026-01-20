import * as React from "react";
import {
  Badge,
  Button,
  Card,
  CardHeader,
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
import { ScopeType, listRecentCases, listRecentClients, filterResults, loadAttachmentTree } from "../services/legalLedgerData";

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

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  // -----------------------
  // Legal Ledger auth state
  // -----------------------
  const [llEmail, setLlEmail] = React.useState("");
  const [llPassword, setLlPassword] = React.useState("");
  const [llAuthStatus, setLlAuthStatus] = React.useState<string>("");

  const [llUserEmail, setLlUserEmail] = React.useState<string | null>(null);
  const [orgs, setOrgs] = React.useState<any[]>([]);
  const [selectedOrgId, setSelectedOrgId] = React.useState<string>(() => {
    return localStorage.getItem(ORG_ID_KEY) ?? "";
  });
  const [settingsOpen, setSettingsOpen] = React.useState(false);
  const [selectedOrgName, setSelectedOrgName] = React.useState<string>(() => {
    return localStorage.getItem(ORG_NAME_KEY) ?? "";
  });


  React.useEffect(() => {
    llSupabase.auth.getSession().then(({ data }) => {
      const email = data.session?.user?.email ?? null;
      setLlUserEmail(email);
    });

    const { data: sub } = llSupabase.auth.onAuthStateChange((_event, session) => {
      setLlUserEmail(session?.user?.email ?? null);
    });

    return () => sub.subscription.unsubscribe();
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


  async function onLlLogin() {
    try {
      setLlAuthStatus("Logging in…");
      const { data, error } = await llSupabase.auth.signInWithPassword({
        email: llEmail,
        password: llPassword,
      });
      if (error) throw error;

      setLlUserEmail(data.user?.email ?? null);
      setLlAuthStatus("Logged in.");
    } catch (e: any) {
      setLlAuthStatus(`Error: ${e?.message ?? String(e)}`);
    }
  }

  async function onLlLogout() {
    await llSupabase.auth.signOut();
    setLlUserEmail(null);
    setOrgs([]);
    setSelectedOrgId("");
    setLlAuthStatus("Logged out.");
  }

  async function onLoadMyOrgs(opts: { silent?: boolean } = {}) {
    try {
      if (!opts.silent) setLlAuthStatus("Loading organizations...");

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
        setLlAuthStatus(list.length > 0 ? `Loaded ${list.length} org(s).` : "No organizations found.");
      }
    } catch (e: any) {
      console.error(e);
      if (!opts.silent) setLlAuthStatus(`Error loading orgs: ${e?.message ?? String(e)}`);
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
  const [selectedFolderId, setSelectedFolderId] = React.useState<string>(""); // optional

  const destFiltered = React.useMemo(() => filterResults(destAllResults, destQuery), [destAllResults, destQuery]);

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

  async function onLoadRecentDestination() {
    try {
      setDestStatus("Loading recent…");
      setDestAllResults([]);
      setSelectedScopeId("");
      setTreeNodes([]);
      setSelectedFolderId("");

      if (!selectedOrgId) throw new Error("Select an org first.");

      const results =
        scopeType === "case"
          ? await listRecentCases({ orgId: selectedOrgId, limit: 200 })
          : await listRecentClients({ orgId: selectedOrgId, limit: 200 });


      setDestAllResults(results);
      setDestStatus(`Loaded ${results.length} recent ${scopeType}s.`);
    } catch (e: any) {
      setDestStatus(`Error: ${e?.message ?? String(e)}`);
    }
  }

  async function onLoadFolders() {
    try {
      setTreeStatus("Loading folders…");
      setTreeNodes([]);
      setSelectedFolderId("");

      if (!selectedScopeId) throw new Error("Select a case/client first.");

      const nodes = await loadAttachmentTree({ scopeType, scopeId: selectedScopeId });
      setTreeNodes(nodes);
      setTreeStatus(`Loaded ${nodes.length} node(s).`);
    } catch (e: any) {
      setTreeStatus(`Error: ${e?.message ?? String(e)}`);
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

  function revokeIfAny(url: string | undefined | null) {
    if (url) URL.revokeObjectURL(url);
  }

  async function onPrepareBundle() {
    try {
      setBundleStatus("Preparing bundle…");

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

      setBundleStatus(
        `Ready: ${bundle.eml.name} (${Math.round(bundle.eml.size / 1024)} KB), ${bundle.attachments.length} attachment(s)`
      );
    } catch (e: any) {
      setBundleStatus(`Error: ${e?.message ?? String(e)}`);
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
          setMailError("No mailbox item available. Open a mail message (Read mode) and try again.");
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

  return (
    <div className={styles.root}>
      {/* Header */}
      <div className={styles.headerRow}>
        <div className={styles.headerLeft}>
          <Text size={600} weight="semibold">
            {props.title || "Legal Ledger"}
          </Text>

          <div className={styles.statusRow}>
            <Badge appearance={isLoggedIn ? "filled" : "outline"} color={isLoggedIn ? "success" : "important"}>
              {isLoggedIn ? "Connected" : "Not connected"}
            </Badge>

            {isLoggedIn && (
              <Text size={200} className={styles.subtle}>
                {llUserEmail}
              </Text>
            )}
          </div>
        </div>

        <Switch checked={showDebug} onChange={(_, data) => setShowDebug(!!data.checked)} label="Show debug" />
      </div>

      {/* Legal Ledger Connection */}
      <Card>
        <CardHeader
          header={<Text weight="semibold">Legal Ledger</Text>}
          description={
            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
              <Text size={200}>Logged in as: {llUserEmail || "-"}</Text>
              <Text size={200}>Org: {selectedOrgName || "-"}</Text>
            </div>
          }

          action={
            <Button
              appearance="subtle"
              icon={<Settings24Regular />}
              onClick={openSettings}
              aria-label="Settings"
            />
          }
        />
      </Card>


      <>
        <div className={styles.row}>
          <Field label="Type" className={styles.grow}>
            <Dropdown
              value={scopeType === "case" ? "Case" : "Client"}
              selectedOptions={[scopeType]}
              onOptionSelect={(_, data) => setScopeType(data.optionValue as ScopeType)}
            >
              <Option value="case">Case</Option>
              <Option value="client">Client</Option>
            </Dropdown>
          </Field>

          <Button onClick={onLoadRecentDestination} disabled={!selectedOrgId}>
            Reload recent
          </Button>
          {!selectedOrgId && <Text size={200}>Select an org in settings first.</Text>}


        </div>

        <Field label="Filter">
          <Input
            value={destQuery}
            onChange={(e) => setDestQuery((e.target as HTMLInputElement).value)}
            placeholder={scopeType === "case" ? "Type to filter cases…" : "Type to filter clients…"}
          />
        </Field>

        {destStatus && <Text size={200}>{destStatus}</Text>}

        {destFiltered.length > 0 && (
          <Field label="Select">
            <Dropdown
              value={selectedScopeLabel}
              selectedOptions={selectedScopeId ? [selectedScopeId] : []}
              onOptionSelect={(_, data) => setSelectedScopeId(String(data.optionValue ?? ""))}
              placeholder="Select…"
            >
              {destFiltered.slice(0, 25).map((r) => (
                <Option key={r.id} value={r.id}>
                  {r.label}
                </Option>
              ))}
            </Dropdown>
          </Field>
        )}

        <div className={styles.row}>
          <Button onClick={onLoadFolders} disabled={!selectedScopeId}>
            Load attachment folders
          </Button>
          {treeStatus && <Text size={200}>{treeStatus}</Text>}
        </div>

        {treeNodes.length > 0 && (
          <Field label="Folder (optional)">
            <Dropdown
              value={
                selectedFolderId
                  ? String(treeNodes.find((n) => String(n.id) === String(selectedFolderId))?.name ?? selectedFolderId)
                  : "(Root)"
              }
              selectedOptions={selectedFolderId ? [selectedFolderId] : []}
              onOptionSelect={(_, data) => setSelectedFolderId(String(data.optionValue ?? ""))}
            >
              <Option value="">(Root)</Option>
              {treeNodes
                .filter((n) => (n.node_type ?? n.type) === "folder")
                .map((n) => (
                  <Option key={String(n.id)} value={String(n.id)}>
                    {String(n.name ?? n.title ?? n.id)}
                  </Option>
                ))}
            </Dropdown>
          </Field>
        )}
      </>



      {/* Bundle */}
      <Card>
        <CardHeader
          header={<Text weight="semibold">Email bundle</Text>}
          description={
            <Text size={200} className={styles.subtle}>
              Prepare a bundle (email.eml + attachments). Upload comes next.
            </Text>
          }
        />
        <Divider />
        <div className={styles.cardBody}>
          <div className={styles.row}>
            <Button appearance="primary" onClick={onPrepareBundle}>
              Prepare email bundle
            </Button>
          </div>

          {bundleStatus && <Text size={200}>{bundleStatus}</Text>}

          {bundleEmlDownload && (
            <div>
              <Text size={200} weight="semibold">
                Email
              </Text>
              <div>
                <a href={bundleEmlDownload.url} download={bundleEmlDownload.name}>
                  Download {bundleEmlDownload.name}
                </a>
              </div>
            </div>
          )}

          {bundleAttachmentDownloads.length > 0 && (
            <div>
              <Text size={200} weight="semibold">
                Attachments
              </Text>
              <ul className={styles.downloadsList}>
                {bundleAttachmentDownloads.map((d) => (
                  <li key={d.name}>
                    <a href={d.url} download={d.name}>
                      Download {d.name}
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
          <CardHeader header={<Text weight="semibold">Debug</Text>} />
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
            <DialogTitle>Legal Ledger settings</DialogTitle>

            <DialogContent>
              {!llUserEmail ? (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <Text>You’re not logged in.</Text>
                  <Button appearance="primary" onClick={onLlLogin /* rename if needed */}>
                    Log in
                  </Button>
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                  <Field label="Organization">
                    {orgs.length === 0 ? (
                      <Text size={200}>Loading organizations…</Text>
                    ) : (
                      <Dropdown
                        value={selectedOrgName || "Select org…"}
                        selectedOptions={selectedOrgId ? [selectedOrgId] : []}
                        onOptionSelect={(_, data) => {
                          const orgId = String(data.optionValue ?? "");

                          const org = orgs.find((o: any) => String(o.org_id ?? o.id) === orgId);
                          const orgName = String(org?.name ?? org?.org_name ?? org?.title ?? orgId);

                          persistSelectedOrg(orgId, orgName);
                        }}
                        placeholder="Select…"
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

                  <Button onClick={onLlLogout /* rename if needed */}>Log out</Button>
                </div>
              )}
            </DialogContent>

            <DialogActions>
              <Button appearance="primary" onClick={() => setSettingsOpen(false)}>
                Done
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>


    </div>
  );
};

export default App;
