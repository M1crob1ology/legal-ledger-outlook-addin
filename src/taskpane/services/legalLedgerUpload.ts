import type { SupabaseClient } from "@supabase/supabase-js";

export async function uploadAttachmentFile(p: {
  supabase: SupabaseClient;
  scopeType: "case" | "party";
  orgId: string;
  scopeId: string;
  parentId: string | null;
  file: File;
  customFileName?: string;
}) {
  const { supabase, scopeType, orgId, scopeId, parentId, file, customFileName } = p;

  const bucket = scopeType === "case" ? "case-attachments" : "client-attachments";
  const fileExt = (file.name.split(".").pop() || "bin").toLowerCase();
  const storagePath = `${orgId}/${scopeId}/${Date.now()}-${Math.random().toString(36).substring(7)}.${fileExt}`;

  const { error: uploadError } = await supabase.storage.from(bucket).upload(storagePath, file);
  if (uploadError) throw uploadError;

  const { error: dbError } = await supabase.from("attachment_nodes").insert({
    org_id: orgId,
    scope_type: scopeType,         // 'case' or 'party'
    scope_id: scopeId,             // case.id or party.id
    type: "file",
    name: customFileName || file.name,
    parent_id: parentId,           // folder id or null for root
    storage_path: storagePath,
    mime_type: file.type || null,
    file_size: file.size || null,
  });

  if (dbError) throw dbError;

  return { bucket, storagePath };
}
