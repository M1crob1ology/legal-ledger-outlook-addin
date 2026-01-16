import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import { getCurrentMessageAsEmlFile, getCurrentMessageAttachmentsAsFiles } from "../services/outlook";
import { getCurrentEmailBundle } from "../services/outlook";


interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

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

      // Clean up old links so we don't leak memory
      revokeIfAny(bundleEmlDownload?.url);
      bundleAttachmentDownloads.forEach((d) => revokeIfAny(d.url));

      setBundleEmlDownload(null);
      setBundleAttachmentDownloads([]);

      const bundle = await getCurrentEmailBundle();

      // Create browser download links
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


  const [emlStatus, setEmlStatus] = React.useState<string>("");
  const [emlDownload, setEmlDownload] = React.useState<{ url: string; name: string } | null>(null);

  const [attStatus, setAttStatus] = React.useState<string>("");
  const [attDownloads, setAttDownloads] = React.useState<Array<{ url: string; name: string }>>([]);

  const [mailInfo, setMailInfo] = React.useState<any>(null);
  const [mailError, setMailError] = React.useState<string | null>(null);
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  async function onExportEml() {
    try {
      setEmlStatus("Exporting .eml…");
      const file = await getCurrentMessageAsEmlFile();
      const url = URL.createObjectURL(file);
      setEmlDownload({ url, name: file.name });
      setEmlStatus(`Ready: ${file.name} (${Math.round(file.size / 1024)} KB)`);
    } catch (e: any) {
      setEmlStatus(`Error: ${e?.message ?? String(e)}`);
    }
  }

  async function onExportAttachments() {
    try {
      setAttStatus("Exporting attachments…");
      const files = await getCurrentMessageAttachmentsAsFiles();

      const dl = files.map((f) => ({ url: URL.createObjectURL(f), name: f.name }));
      setAttDownloads(dl);
      setAttStatus(`Ready: ${files.length} attachment(s)`);
    } catch (e: any) {
      setAttStatus(`Error: ${e?.message ?? String(e)}`);
    }
  }


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

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <TextInsertion insertText={insertText} />
      <div style={{ marginTop: 16, padding: 12, border: "1px solid #ddd", borderRadius: 8 }}>
        <h3 style={{ margin: 0, marginBottom: 8 }}>Bundle</h3>

        <button onClick={onPrepareBundle}>Prepare email bundle</button>

        {bundleStatus && <div style={{ marginTop: 8 }}>{bundleStatus}</div>}

        {bundleEmlDownload && (
          <div style={{ marginTop: 8 }}>
            <a href={bundleEmlDownload.url} download={bundleEmlDownload.name}>
              Download {bundleEmlDownload.name}
            </a>
          </div>
        )}

        {bundleAttachmentDownloads.length > 0 && (
          <div style={{ marginTop: 8 }}>
            <div style={{ marginBottom: 4 }}>Attachments:</div>
            <ul style={{ margin: 0, paddingLeft: 16 }}>
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


      <div style={{ marginTop: 16, padding: 12, border: "1px solid #ddd", borderRadius: 8 }}>
        <h3 style={{ margin: 0, marginBottom: 8 }}>Debug: current mail</h3>
        {mailError && <pre style={{ whiteSpace: "pre-wrap" }}>{mailError}</pre>}
        {mailInfo && <pre style={{ whiteSpace: "pre-wrap" }}>{JSON.stringify(mailInfo, null, 2)}</pre>}
      </div>

    </div>
  );
};

export default App;
