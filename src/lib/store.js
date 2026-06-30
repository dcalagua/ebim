import { supabase } from "./supabaseClient";

// Mismo contrato que el storage local del artifact original (set/get/del/list con claves "est:<code>"),
// pero respaldado por la tabla `estimations` de Supabase para que el equipo comparta el historial.
export const store = {
  async set(k, v) {
    const payload = JSON.parse(v);
    const code = k.replace(/^est:/, "");
    const { data: userData } = await supabase.auth.getUser();
    const { error } = await supabase.from("estimations").upsert({
      code,
      client: payload.client || "",
      project: payload.project || "",
      country: payload.country || "",
      data: payload,
      created_by: userData?.user?.id,
      saved_at: new Date(payload.savedAt || Date.now()).toISOString(),
    });
    if (error) throw error;
  },
  async get(k) {
    const code = k.replace(/^est:/, "");
    const { data, error } = await supabase.from("estimations").select("data").eq("code", code).maybeSingle();
    if (error) throw error;
    return data ? JSON.stringify(data.data) : null;
  },
  async del(k) {
    const code = k.replace(/^est:/, "");
    const { error } = await supabase.from("estimations").delete().eq("code", code);
    if (error) throw error;
  },
  async list() {
    const { data, error } = await supabase.from("estimations").select("code").order("saved_at", { ascending: false });
    if (error) throw error;
    return (data || []).map((r) => "est:" + r.code);
  },
};
