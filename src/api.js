import { supabase } from "./supabaseClient";

// ============================================================
// Data layer. Every function here talks to Supabase.
// The RLS policies on the server decide which rows come back,
// so the app doesn't have to re-check permissions — it just asks.
//
// The DB uses snake_case (owner_id, close_date...) while the UI
// uses camelCase (ownerId, closeDate...). We translate at this
// boundary so the components stay unchanged.
// ============================================================

const toCamelProfile = (p) => p && ({ id: p.id, name: p.name, email: p.email, role: p.role, managerId: p.manager_id });
const toCamelEntry = (e) => e && ({ id: e.id, userId: e.user_id, date: e.date, calls: e.calls, emails: e.emails, appts: e.appts, notes: e.notes, fromDeal: e.from_deal, taggedRepId: e.tagged_rep_id || "", company: e.company || "", ban: e.ban || "", fan: e.fan || "", contact: e.contact || "", phone: e.phone || "", email: e.email || "", carrierRep: e.carrier_rep || "", companyId: e.company_id || "" });
const toCamelDeal = (d) => d && ({ id: d.id, ownerId: d.owner_id, company: d.company, contact: d.contact, value: Number(d.value), stage: d.stage, closeDate: d.close_date || "", notes: d.notes, apptCredited: d.appt_credited, createdAt: d.created_at, taggedRepId: d.tagged_rep_id || "", companyId: d.company_id || "" });

const fromCamelEntry = (e) => ({ ...(e.id ? { id: e.id } : {}), user_id: e.userId, date: e.date, calls: e.calls, emails: e.emails, appts: e.appts, notes: e.notes, from_deal: e.fromDeal ?? null, tagged_rep_id: e.taggedRepId || null, company: e.company || null, ban: e.ban || null, fan: e.fan || null, contact: e.contact || null, phone: e.phone || null, email: e.email || null, carrier_rep: e.carrierRep || null, company_id: e.companyId || null });
const fromCamelDeal = (d) => ({ ...(d.id ? { id: d.id } : {}), owner_id: d.ownerId, company: d.company, contact: d.contact, value: d.value, stage: d.stage, close_date: d.closeDate || null, notes: d.notes, appt_credited: d.apptCredited ?? false, tagged_rep_id: d.taggedRepId || null, company_id: d.companyId || null });

// ---- Auth ----
export async function signIn(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data.user;
}

export async function signOut() {
  await supabase.auth.signOut();
}

export async function getSession() {
  const { data } = await supabase.auth.getSession();
  return data.session;
}

// Fetch the signed-in user's own profile (role, name, manager, etc.)
export async function getMyProfile() {
  const { data: auth } = await supabase.auth.getUser();
  if (!auth?.user) return null;
  const { data, error } = await supabase
    .from("profiles")
    .select("*")
    .eq("id", auth.user.id)
    .single();
  if (error) throw error;
  return toCamelProfile(data);
}

// ---- Profiles (users) ----
// Returns every profile the caller is allowed to see (RLS-filtered).
export async function listProfiles() {
  const { data, error } = await supabase.from("profiles").select("*").order("name");
  if (error) throw error;
  return data.map(toCamelProfile);
}

// Admin: create a user. This calls a secure Edge Function (see supabase/functions)
// because creating an auth account requires the service role, which must never
// live in the browser. The function verifies the caller is an admin.
export async function adminCreateUser({ name, email, password, role, managerId }) {
  const { data, error } = await supabase.functions.invoke("admin-create-user", {
    body: { name, email, password, role, managerId },
  });
  if (error) throw error;
  return data;
}

export async function updateProfile(id, patch) {
  const dbPatch = {};
  if (patch.name !== undefined) dbPatch.name = patch.name;
  if (patch.email !== undefined) dbPatch.email = patch.email;
  if (patch.role !== undefined) dbPatch.role = patch.role;
  if (patch.managerId !== undefined) dbPatch.manager_id = patch.managerId;
  const { error } = await supabase.from("profiles").update(dbPatch).eq("id", id);
  if (error) throw error;
}

export async function deleteProfile(id) {
  // Deleting the auth user cascades to the profile; do it via Edge Function (admin only).
  const { error } = await supabase.functions.invoke("admin-delete-user", { body: { id } });
  if (error) throw error;
}

// ---- Entries ----
export async function listEntries() {
  const { data, error } = await supabase.from("entries").select("*").order("date", { ascending: false });
  if (error) throw error;
  return data.map(toCamelEntry);
}

export async function addEntry(entry) {
  // Tie the activity to a company record (create one if this name is new).
  if (entry.company && !entry.companyId) {
    entry = { ...entry, companyId: await findOrCreateCompany(entry.company) };
  }
  const { data, error } = await supabase.from("entries").insert(fromCamelEntry(entry)).select().single();
  if (error) throw error;
  return toCamelEntry(data);
}

// Insert many entries at once (used by bulk upload). Returns the created rows.
export async function addEntriesBulk(entries) {
  // Resolve each distinct company name to an id once, then link rows.
  const names = [...new Set(entries.map((e) => (e.company || "").trim()).filter(Boolean))];
  const idByName = {};
  for (const n of names) idByName[n.toLowerCase()] = await findOrCreateCompany(n);
  const linked = entries.map((e) => ({ ...e, companyId: e.companyId || idByName[(e.company || "").trim().toLowerCase()] || null }));
  const rows = linked.map(fromCamelEntry);
  const { data, error } = await supabase.from("entries").insert(rows).select();
  if (error) throw error;
  return (data || []).map(toCamelEntry);
}

export async function deleteEntry(id) {
  const { error } = await supabase.from("entries").delete().eq("id", id);
  if (error) throw error;
}

// ---- Deals ----
export async function listDeals() {
  const { data, error } = await supabase.from("deals").select("*").order("created_at", { ascending: false });
  if (error) throw error;
  return data.map(toCamelDeal);
}

export async function upsertDeal(deal) {
  if (deal.company && !deal.companyId) {
    deal = { ...deal, companyId: await findOrCreateCompany(deal.company) };
  }
  const { data, error } = await supabase.from("deals").upsert(fromCamelDeal(deal)).select().single();
  if (error) throw error;
  return toCamelDeal(data);
}

export async function updateDeal(id, patch) {
  const dbPatch = fromCamelDeal({ ...patch, id });
  delete dbPatch.id;
  const { data, error } = await supabase.from("deals").update(dbPatch).eq("id", id).select().single();
  if (error) throw error;
  return toCamelDeal(data);
}

export async function deleteDeal(id) {
  const { error } = await supabase.from("deals").delete().eq("id", id);
  if (error) throw error;
}

// ---- Goals ----
export async function getGoals() {
  const { data, error } = await supabase.from("goals").select("*").eq("id", 1).single();
  if (error) throw error;
  return data;
}

export async function saveGoals(patch) {
  const { error } = await supabase.from("goals").update(patch).eq("id", 1);
  if (error) throw error;
}

// ---- Per-person goal overrides ----
// Returns a map { userId: {calls, emails, appts} } of everyone with a custom goal.
export async function listUserGoals() {
  const { data, error } = await supabase.from("user_goals").select("*");
  if (error) throw error;
  const map = {};
  (data || []).forEach((g) => { map[g.user_id] = { calls: g.calls, emails: g.emails, appts: g.appts }; });
  return map;
}

// Set (create or update) an individual's goal override.
export async function setUserGoal(userId, goal) {
  const { error } = await supabase.from("user_goals").upsert({
    user_id: userId, calls: goal.calls, emails: goal.emails, appts: goal.appts, updated_at: new Date().toISOString(),
  });
  if (error) throw error;
}

// Remove an override so the person falls back to the team default.
export async function clearUserGoal(userId) {
  const { error } = await supabase.from("user_goals").delete().eq("user_id", userId);
  if (error) throw error;
}

// ============================================================
// CRM: Companies, Contacts, Notes, Attachments
// ============================================================
const toCamelCompany = (c) => c && ({
  id: c.id, name: c.name, nameKey: c.name_key, industry: c.industry || "",
  website: c.website || "", phone: c.phone || "", address: c.address || "",
  ban: c.ban || "", fan: c.fan || "", notes: c.notes || "", createdBy: c.created_by, createdAt: c.created_at, updatedAt: c.updated_at,
});

export async function listCompanies() {
  const { data, error } = await supabase.from("companies").select("*").order("name", { ascending: true });
  if (error) throw error;
  return (data || []).map(toCamelCompany);
}

export async function getCompany(id) {
  const { data, error } = await supabase.from("companies").select("*").eq("id", id).single();
  if (error) throw error;
  return toCamelCompany(data);
}

// Update editable company fields.
export async function updateCompany(id, patch) {
  const db = {};
  ["name", "industry", "website", "phone", "address", "ban", "fan", "notes"].forEach((k) => {
    if (patch[k] !== undefined) db[k] = patch[k];
  });
  if (patch.name !== undefined) db.name_key = patch.name.trim().toLowerCase();
  const { data, error } = await supabase.from("companies").update(db).eq("id", id).select().single();
  if (error) throw error;
  return toCamelCompany(data);
}

// Find an existing company by name, or create it. Returns the company id.
export async function findOrCreateCompany(name) {
  const clean = (name || "").trim();
  if (!clean) return null;
  // Use the SECURITY DEFINER DB function: it finds an existing company by
  // normalized name (regardless of row visibility) or creates it atomically.
  // This avoids the "can't see it -> try to insert -> duplicate key" failure
  // that blocked saves for BDRs. If anything goes wrong, return null so the
  // activity still saves (linked by company name, just without company_id).
  try {
    const { data, error } = await supabase.rpc("find_or_create_company", { p_name: clean });
    if (error) return null;
    return data || null;
  } catch {
    return null;
  }
}

// ---- Contacts ----
const toCamelContact = (c) => c && ({ id: c.id, companyId: c.company_id, name: c.name, title: c.title || "", phone: c.phone || "", email: c.email || "", notes: c.notes || "", createdAt: c.created_at });

export async function listContacts(companyId) {
  const { data, error } = await supabase.from("company_contacts").select("*").eq("company_id", companyId).order("name");
  if (error) throw error;
  return (data || []).map(toCamelContact);
}
export async function saveContact(companyId, contact) {
  const row = { company_id: companyId, name: contact.name, title: contact.title || "", phone: contact.phone || "", email: contact.email || "", notes: contact.notes || "" };
  if (contact.id) {
    const { error } = await supabase.from("company_contacts").update(row).eq("id", contact.id);
    if (error) throw error;
  } else {
    const { error } = await supabase.from("company_contacts").insert(row);
    if (error) throw error;
  }
}
export async function deleteContact(id) {
  const { error } = await supabase.from("company_contacts").delete().eq("id", id);
  if (error) throw error;
}

// ---- Notes ----
const toCamelNote = (n) => n && ({ id: n.id, companyId: n.company_id, authorId: n.author_id, body: n.body, createdAt: n.created_at });

export async function listCompanyNotes(companyId) {
  const { data, error } = await supabase.from("company_notes").select("*").eq("company_id", companyId).order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(toCamelNote);
}
export async function addCompanyNote(companyId, body) {
  const { data: me } = await supabase.auth.getUser();
  const { error } = await supabase.from("company_notes").insert({ company_id: companyId, author_id: me?.user?.id || null, body });
  if (error) throw error;
}
export async function deleteCompanyNote(id) {
  const { error } = await supabase.from("company_notes").delete().eq("id", id);
  if (error) throw error;
}

// ---- Attachments (Storage bucket: company-files) ----
const toCamelAttachment = (a) => a && ({ id: a.id, companyId: a.company_id, uploaderId: a.uploader_id, fileName: a.file_name, storagePath: a.storage_path, sizeBytes: a.size_bytes || 0, createdAt: a.created_at });

export async function listAttachments(companyId) {
  const { data, error } = await supabase.from("company_attachments").select("*").eq("company_id", companyId).order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(toCamelAttachment);
}
export async function uploadAttachment(companyId, file) {
  const { data: me } = await supabase.auth.getUser();
  const safe = file.name.replace(/[^\w.\-]+/g, "_");
  const path = `${companyId}/${Date.now()}_${safe}`;
  const up = await supabase.storage.from("company-files").upload(path, file, { upsert: false });
  if (up.error) throw up.error;
  const { error } = await supabase.from("company_attachments").insert({
    company_id: companyId, uploader_id: me?.user?.id || null,
    file_name: file.name, storage_path: path, size_bytes: file.size || 0,
  });
  if (error) throw error;
}
export async function attachmentUrl(storagePath) {
  const { data, error } = await supabase.storage.from("company-files").createSignedUrl(storagePath, 60 * 10);
  if (error) throw error;
  return data.signedUrl;
}
export async function deleteAttachment(att) {
  await supabase.storage.from("company-files").remove([att.storagePath]);
  const { error } = await supabase.from("company_attachments").delete().eq("id", att.id);
  if (error) throw error;
}
// ============================================================
const toCamelLead = (l) => l && ({
  id: l.id, company: l.company, contact: l.contact || "", title: l.title || "",
  phone: l.phone || "", email: l.email || "", ban: l.ban || "", fan: l.fan || "",
  source: l.source || "", notes: l.notes || "", status: l.status || "New",
  assignedTo: l.assigned_to || "", createdBy: l.created_by || "", companyId: l.company_id || "",
  createdAt: l.created_at, updatedAt: l.updated_at,
});
const fromCamelLead = (l) => ({
  ...(l.id ? { id: l.id } : {}),
  company: l.company, contact: l.contact || "", title: l.title || "",
  phone: l.phone || "", email: l.email || "", ban: l.ban || "", fan: l.fan || "",
  source: l.source || "", notes: l.notes || "", status: l.status || "New",
  assigned_to: l.assignedTo || null, created_by: l.createdBy || null,
});

export async function listLeads() {
  const { data, error } = await supabase.from("leads").select("*").order("created_at", { ascending: false });
  if (error) throw error;
  return (data || []).map(toCamelLead);
}

// Create one lead. created_by is set to the current user.
export async function addLead(lead) {
  const { data: me } = await supabase.auth.getUser();
  const row = { ...fromCamelLead(lead), created_by: me?.user?.id || null };
  const { data, error } = await supabase.from("leads").insert(row).select().single();
  if (error) throw error;
  return toCamelLead(data);
}

// Bulk insert leads (from an uploaded list).
export async function addLeadsBulk(leads) {
  const { data: me } = await supabase.auth.getUser();
  const uid = me?.user?.id || null;
  const rows = leads.map((l) => ({ ...fromCamelLead(l), created_by: uid }));
  const { data, error } = await supabase.from("leads").insert(rows).select();
  if (error) throw error;
  return (data || []).map(toCamelLead);
}

// Update a lead (status, assignment, notes, any field).
export async function updateLead(id, patch) {
  const db = {};
  const map = { company:"company", contact:"contact", title:"title", phone:"phone", email:"email",
    ban:"ban", fan:"fan", source:"source", notes:"notes", status:"status" };
  Object.keys(map).forEach((k) => { if (patch[k] !== undefined) db[map[k]] = patch[k]; });
  if (patch.assignedTo !== undefined) db.assigned_to = patch.assignedTo || null;
  const { data, error } = await supabase.from("leads").update(db).eq("id", id).select().single();
  if (error) throw error;
  return toCamelLead(data);
}

export async function deleteLead(id) {
  const { error } = await supabase.from("leads").delete().eq("id", id);
  if (error) throw error;
}
