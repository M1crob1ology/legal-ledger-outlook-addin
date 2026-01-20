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
} from "@fluentui/react-components";

import { llSupabase } from "../services/legalLedgerSupabase";
import { getCurrentEmailBundle } from "../services/outlook";
import { ScopeType, listRecentCases, listRecentClients, filterResults, loadAttachmentTree } from "../services/legalLedgerData";



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
  const [selectedOrgId, setSelectedOrgId] = React.useState<string>("");

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

  async function onLoadMyOrgs() {
    try {
      setLlAuthStatus("Loading orgs…");
      const { data, error } = await llSupabase.rpc("list_my_orgs");
      if (error) throw error;

      const list = Array.isArray(data) ? data : [];
      setOrgs(list);

      const firstId = list?.[0]?.org_id || list?.[0]?.id || "";
      if (firstId) setSelectedOrgId(firstId);

      setLlAuthStatus(`Loaded ${list.length} org(s).`);
    } catch (e: any) {
      setLlAuthStatus(`Error: ${e?.message ?? String(e)}`);
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
          header={<Text weight="semibold">Legal Ledger connection</Text>}
          description={<Text size={200} className={styles.subtle}>Log in and select an organization.</Text>}
        />
        <Divider />
        <div className={styles.cardBody}>
          {!isLoggedIn ? (
            <>
              <Field label="Email">
                <Input value={llEmail} onChange={(e) => setLlEmail((e.target as HTMLInputElement).value)} />
              </Field>

              <Field label="Password">
                <Input
                  type="password"
                  value={llPassword}
                  onChange={(e) => setLlPassword((e.target as HTMLInputElement).value)}
                />
              </Field>

              <div className={styles.row}>
                <Button appearance="primary" onClick={onLlLogin}>
                  Log in
                </Button>
              </div>
            </>
          ) : (
            <>
              <div className={styles.row}>
                <Button onClick={onLoadMyOrgs}>Load my orgs</Button>
                <Button appearance="secondary" onClick={onLlLogout}>
                  Log out
                </Button>
              </div>

              <Field label="Organization">
                <Dropdown
                  value={selectedOrgLabel} // <-- show name, not id
                  selectedOptions={selectedOrgId ? [selectedOrgId] : []}
                  onOptionSelect={(_, data) => setSelectedOrgId(String(data.optionValue ?? ""))}
                  placeholder={orgs.length ? "Select an org…" : "Click “Load my orgs”"}
                >
                  {orgs.map((o) => {
                    const id = getOrgId(o);
                    const name = getOrgName(o);
                    return (
                      <Option key={id} value={id}>
                        {name}
                      </Option>
                    );
                  })}
                </Dropdown>

              </Field>
            </>
          )}

          {llAuthStatus && <Text size={200}>{llAuthStatus}</Text>}
        </div>
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

          <Button onClick={onLoadRecentDestination}>Reload recent</Button>
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
    </div>
  );
};

export default App;
