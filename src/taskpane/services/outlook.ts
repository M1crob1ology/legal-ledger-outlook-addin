// src/taskpane/services/outlook.ts

function safeFilename(name: string) {
    // Windows + general filesystem-safe
    return name
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 120);
}

function getSliceCount(file: any): number {
    const raw = file?.sliceCount ?? file?._sliceCount;
    const n = Number(raw);

    if (Number.isFinite(n) && n > 0) return n;

    // Fallback: compute from size / sliceSize if present
    const size = Number(file?.size);
    const sliceSize = Number(file?.sliceSize ?? file?._sliceSize);

    if (Number.isFinite(size) && Number.isFinite(sliceSize) && size > 0 && sliceSize > 0) {
        return Math.ceil(size / sliceSize);
    }

    return 0;
}

function looksLikeBase64(s: string): boolean {
    const t = s.replace(/\s/g, "");
    // very lightweight heuristic
    return t.length > 0 && t.length % 4 === 0 && /^[A-Za-z0-9+/]+=*$/.test(t);
}

function sliceDataToArrayBuffer(data: any): ArrayBuffer {
    // Outlook can return slice data as base64 string, plain string, array-like, etc.
    if (typeof data === "string") {
        if (looksLikeBase64(data)) return base64ToArrayBuffer(data);
        // treat as plain text
        return new TextEncoder().encode(data).buffer;
    }

    if (data instanceof ArrayBuffer) return data;

    if (Array.isArray(data)) {
        return Uint8Array.from(data).buffer as ArrayBuffer;
    }

    if (data?.buffer && data?.byteLength != null) {
        // ArrayBufferView (Uint8Array etc)
        return data.buffer as ArrayBuffer;
    }

    // Last resort: stringify
    return new TextEncoder().encode(String(data)).buffer;
}


function base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    // Force the return type to ArrayBuffer (avoids ArrayBufferLike/SharedArrayBuffer typing issues)
    return bytes.buffer as ArrayBuffer;
}


function closeOfficeFile(file: any): Promise<void> {
    return new Promise((resolve) => {
        try {
            if (file && typeof file.closeAsync === "function") {
                file.closeAsync(() => resolve());
            } else {
                // Some Outlook builds return a file object without closeAsync
                resolve();
            }
        } catch {
            resolve();
        }
    });
}


export async function getCurrentMessageAsEmlFile(): Promise<File> {
  const item: any = Office.context?.mailbox?.item;

  if (!item) {
    throw new Error("No mailbox item found. Open an email in Read mode.");
  }

  if (typeof item.getAsFileAsync !== "function") {
    throw new Error("getAsFileAsync is not available in this Outlook client.");
  }

  const resultValue = await new Promise<any>((resolve, reject) => {
    item.getAsFileAsync((result: Office.AsyncResult<any>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error?.message || "getAsFileAsync failed"));
    });
  });

  const subject = safeFilename(item.subject || "email");

  // ✅ Case 1: Outlook returns a base64 string (common)
  if (typeof resultValue === "string") {
    const base64 = resultValue.replace(/\s/g, "");
    if (!base64) throw new Error("Outlook returned an empty EML payload.");

    // You already have base64ToArrayBuffer(...)
    const buffer = base64ToArrayBuffer(base64);
    const blob = new Blob([buffer], { type: "message/rfc822" });
    return new File([blob], `${subject}.eml`, { type: "message/rfc822" });
  }

  // ✅ Case 2: Outlook returns a file-like object with slices (less common in your client)
  const officeFile = resultValue;
  if (!officeFile || typeof officeFile.getSliceAsync !== "function") {
    throw new Error("Unexpected getAsFileAsync return type (neither base64 string nor slice file).");
  }

  try {
    const chunks: ArrayBuffer[] = [];
    const sliceCount = getSliceCount(officeFile);

    if (!sliceCount) {
      throw new Error("Outlook returned an empty file (sliceCount=0).");
    }

    for (let i = 0; i < sliceCount; i++) {
      const slice = await new Promise<Office.Slice>((resolve, reject) => {
        officeFile.getSliceAsync(i, (res: Office.AsyncResult<Office.Slice>) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
          else reject(new Error(res.error?.message || `getSliceAsync failed at ${i}`));
        });
      });

      chunks.push(sliceDataToArrayBuffer((slice as any).data));
    }

    const blob = new Blob(chunks, { type: "message/rfc822" });
    return new File([blob], `${subject}.eml`, { type: "message/rfc822" });
  } finally {
    // keep your safe closeOfficeFile(...) if you still have it; otherwise remove this line
    if (typeof closeOfficeFile === "function") {
      await closeOfficeFile(officeFile);
    }
  }
}


export interface AttachmentsResult {
    files: File[];
    skipped: string[];
}

export async function getCurrentMessageAttachmentsAsFiles(): Promise<AttachmentsResult> {
    const item: any = Office.context?.mailbox?.item;

    if (!item) {
        throw new Error("No mailbox item found. Open an email in Read mode.");
    }

    const attachments: any[] = item.attachments || [];
    const files: File[] = [];
    const skipped: string[] = [];

    for (const att of attachments) {
        const name = safeFilename(att.name || `attachment-${att.id}`);

        // Inline attachments are typically signature/embedded images — skip them.
        if (att.isInline === true) {
            skipped.push(name);
            continue;
        }

        const content = await new Promise<Office.AttachmentContent>((resolve, reject) => {
            item.getAttachmentContentAsync(att.id, (res: Office.AsyncResult<Office.AttachmentContent>) => {
                if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
                else reject(new Error(res.error?.message || `getAttachmentContentAsync failed for ${att.name}`));
            });
        });

        // content.format: "base64" | "url" | "eml" | "iCal"
        if (content.format === "base64") {
            const buffer = base64ToArrayBuffer(content.content as any);
            const type = att.contentType || "application/octet-stream";
            files.push(new File([buffer], name, { type }));

            continue;
        }

        if (content.format === "eml") {
            // message attachment (attached email) comes as EML content
            const blob = new Blob([content.content as any], { type: "message/rfc822" });
            files.push(new File([blob], name.endsWith(".eml") ? name : `${name}.eml`, { type: "message/rfc822" }));
            continue;
        }

        if (content.format === "url") {
            // Cloud attachments / links: no inline bytes to upload. Record and skip
            // so the rest of the bundle still uploads.
            skipped.push(name);
            continue;
        }

        // Unknown/unsupported format: skip rather than aborting the whole bundle.
        skipped.push(name);
    }

    return { files, skipped };
}

export async function getCurrentEmailBundle(): Promise<{
  eml: File;
  attachments: File[];
  skippedAttachments: string[];
  meta: { subject: string; from?: string; received?: string };
}> {
  const item: any = Office.context?.mailbox?.item;

  const eml = await getCurrentMessageAsEmlFile();
  const { files: attachments, skipped: skippedAttachments } = await getCurrentMessageAttachmentsAsFiles();

  return {
    eml,
    attachments,
    skippedAttachments,
    meta: {
      subject: item?.subject ?? "",
      from: item?.from?.emailAddress ?? undefined,
      received: item?.dateTimeCreated ?? undefined,
    },
  };
}

