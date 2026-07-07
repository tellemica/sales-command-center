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
const toCamelEntry = (e) => e && ({ id: e.id, userId: e.user_id, date: e.date, calls: e.calls, emails: e.emails, appts: e.appts, notes: e.notes, fromDeal: e.from_deal, taggedRepId: e.tagged_rep_id || "", company: e.company || "", ban: e.ban || "", contact: e.contact || "", phone: e.phone || "", email: e.email || "" });
const toCamelDeal = (d) => d && ({ id: d.id, ownerId: d.owner_id, company: d.company, contact: d.contact, value: Number(d.value), stage: d.stage, closeDate: d.close_date || "", notes: d.notes, apptCredited: d.appt_credited, createdAt: d.created_at, taggedRepId: d.tagged_rep_id || "" });

const fromCamelEntry = (e) => ({ ...(e.id ? { id: e.id } : {}), user_id: e.userId, date: e.date, calls: e.calls, emails: e.emails, appts: e.appts, notes: e.notes, from_deal: e.fromDeal ?? null, tagged_rep_id: e.taggedRepId || null, company: e.company || null, ban: e.ban || null, contact: e.contact || null, phone: e.phone || null, email: e.email || null });
const fromCamelDeal = (d) => ({ ...(d.id ? { id: d.id } : {}), owner_id: d.ownerId, company: d.company, contact: d.contact, value: d.value, stage: d.stage, close_date: d.closeDate || null, notes: d.notes, appt_credited: d.apptCredited ?? false, tagged_rep_id: d.taggedRepId || null });

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
  const { data, error } = await supabase.from("entries").insert(fromCamelEntry(entry)).select().single();
  if (error) throw error;
  return toCamelEntry(data);
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
