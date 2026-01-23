import { llSupabase } from "./legalLedgerSupabase";

export type ScopeType = "case" | "client";
type AttachmentScopeType = "case" | "party";

function toAttachmentScopeType(scopeType: ScopeType): AttachmentScopeType {
    return scopeType === "client" ? "party" : "case";
}

export type SearchResult = { id: string; label: string; raw: any };

function norm(v: any): string {
    return (v ?? "").toString().toLowerCase();
}

function firstNonEmptyString(row: any, keys: string[]): string {
    for (const k of keys) {
        const v = row?.[k];
        if (typeof v === "string" && v.trim()) return v.trim();
    }
    return "";
}

function looksLikeUuid(s: string): boolean {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(s);
}

function getClientScopeId(row: any): string {
    // If "parties" is a wrapper around legacy client records, it often contains a client id
    const candidates = [
        row?.client_id,
        row?.clientId,
        row?.client_uuid,
        row?.clientUuid,
        row?.legacy_client_id,
        row?.legacyClientId,
    ]
        .map((v) => (v == null ? "" : String(v)))
        .filter(Boolean);

    const hit = candidates.find(looksLikeUuid);
    return hit ?? String(row?.id ?? "");
}


function shortId(id: any): string {
    const s = String(id ?? "");
    return s.length >= 8 ? s.slice(0, 8) : s;
}

function heuristicLabel(row: any, preferredKeyRegex: RegExp): string {
    if (!row || typeof row !== "object") return "";

    // Prefer keys that “sound” like a title/name/reference
    const keys = Object.keys(row).filter((k) => preferredKeyRegex.test(k));
    for (const k of keys) {
        const v = row[k];
        if (typeof v === "string" && v.trim() && !looksLikeUuid(v.trim())) return v.trim();
    }
    return "";
}

function buildCaseLabel(r: any): string {
    // Add a lot more candidates than before
    const label =
        firstNonEmptyString(r, [
            "case_name",
            "name",
            "title",
            "case_title",
            "display_name",
            "subject",
            "reference",
            "ref",
            "case_reference",
            "case_number",
            "matter_number",
            "matter",
            "description",
        ]) ||
        heuristicLabel(r, /(case|matter|title|name|subject|reference|ref|number)/i);

    if (label && !looksLikeUuid(label)) return label;

    // Fallback: don’t show full UUID
    return `Case ${shortId(r?.id)}`;
}

function buildClientLabel(r: any): string {
    const name =
        firstNonEmptyString(r, ["name", "display_name", "company_name", "legal_name", "full_name"]) ||
        heuristicLabel(r, /(name|company|display)/i) ||
        `Client ${shortId(r?.id)}`;

    const orgNo = firstNonEmptyString(r, ["organization_number", "org_number", "org_nr", "vat_number", "registration_number"]);
    return orgNo ? `${name} (${orgNo})` : name;
}


async function listWithFallbacks(table: "cases" | "clients" | "parties", orgId: string, limit: number): Promise<any[]> {
    const orgFields = ["org_id", "organization_id"]; // common in LL projects
    const orderFields = ["updated_at", "created_at", "inserted_at"]; // try in order

    let lastErr: any = null;

    // Try org-field + ordering combos
    for (const orgField of orgFields) {
        for (const orderField of orderFields) {
            try {
                const { data, error } = await (llSupabase as any)
                    .from(table)
                    .select("*")
                    .eq(orgField, orgId)
                    .order(orderField, { ascending: false })
                    .limit(limit);

                if (error) throw error;
                return Array.isArray(data) ? data : [];
            } catch (e) {
                lastErr = e;
            }
        }
    }

    // Last resort: no ordering
    for (const orgField of orgFields) {
        try {
            const { data, error } = await (llSupabase as any)
                .from(table)
                .select("*")
                .eq(orgField, orgId)
                .limit(limit);

            if (error) throw error;
            return Array.isArray(data) ? data : [];
        } catch (e) {
            lastErr = e;
        }
    }

    throw lastErr ?? new Error(`Failed to load from ${table}.`);
}

export async function listRecentCases(p: { orgId: string; limit?: number }): Promise<SearchResult[]> {
    const rows = await listWithFallbacks("cases", p.orgId, p.limit ?? 50);
    return rows.map((r) => ({ id: String(r.id), label: buildCaseLabel(r), raw: r }));
}

function isClientParty(r: any): boolean {
    if (r?.is_client === true) return true;

    const t = String(r.party_type ?? r.type ?? r.kind ?? r.role ?? r.category ?? "").toLowerCase();
    return t === "client" || t === "klient";
}

export async function listRecentClients(p: { orgId: string; limit?: number }): Promise<SearchResult[]> {
  const limit = p.limit ?? 200;

  // Clients are Parties now. Attachments use scope_type='party' and scope_id=parties.id.
  // So this function MUST return party.id as SearchResult.id.
  const allParties = await listWithFallbacks("parties", p.orgId, limit);

  if (!Array.isArray(allParties) || allParties.length === 0) {
    // IMPORTANT: Do NOT fall back to the legacy `clients` table, because those IDs
    // won't match attachment_nodes.scope_id (which is party.id). Returning legacy IDs
    // would break attachments again.
    return [];
  }

  const clients = allParties.filter(isClientParty);

  const rowsToUse = clients.length > 0 ? clients : allParties;

  return rowsToUse.map((r: any) => ({
    id: String(r.id), // party.id (UUID) — required for attachments
    label: buildClientLabel(r),
    raw: r,
  }));
}



export function filterResults(results: SearchResult[], q: string): SearchResult[] {
    const qq = q.trim().toLowerCase();
    if (!qq) return results;

    return results.filter((r) => {
        if (norm(r.label).includes(qq)) return true;

        // Fallback: search a few common fields if they exist
        const raw = r.raw ?? {};
        const hay = [
            raw.title,
            raw.case_title,
            raw.reference,
            raw.case_number,
            raw.matter_number,
            raw.name,
            raw.company_name,
            raw.display_name,
            raw.organization_number,
            raw.email,
        ]
            .filter(Boolean)
            .map((x) => x.toString())
            .join(" ");

        return norm(hay).includes(qq);
    });
}

export type AttachmentTreeNode = {
  id: string;
  name: string;
  kind: "folder" | "file";
  parent_id: string | null;
  raw: any;
};

function normalizeAttachmentNode(n: any): AttachmentTreeNode {
  const kindRaw =
    n.kind ??
    n.type ?? // <-- what Supabase returns from attachment_nodes
    n.node_type ??
    n.nodeType ??
    n.item_type ??
    n.itemType;

  const kind = String(kindRaw ?? "").toLowerCase();

  return {
    id: String(n.id),
    name: String(n.name ?? n.filename ?? n.title ?? "(unnamed)"),
    kind: kind === "folder" ? "folder" : "file",
    parent_id: n.parent_id ?? n.parentId ?? null,
    raw: n,
  };
}

export async function loadAttachmentTree(p: { scopeType: "case" | "client"; scopeId: string }): Promise<AttachmentTreeNode[]> {
  const scopeType = toAttachmentScopeType(p.scopeType); // "client" -> "party", "case" -> "case"

  const { data, error } = await (llSupabase as any).rpc("list_attachment_tree", {
    p_scope_type: scopeType,
    p_scope_id: p.scopeId,
    p_limit: 10000,
    p_offset: 0,
  });

  if (error) throw error;

  const rows = Array.isArray(data) ? data : [];
  return rows.map(normalizeAttachmentNode);
}


