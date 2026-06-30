import { useEffect, useState } from "react";
import { supabase } from "./lib/supabaseClient";

export default function Auth({ children }) {
  const [session, setSession] = useState(undefined);
  const [email, setEmail] = useState("");
  const [sent, setSent] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  async function sendLink(e) {
    e.preventDefault();
    setError(""); setLoading(true);
    const { error: err } = await supabase.auth.signInWithOtp({
      email: email.trim(),
      options: { emailRedirectTo: window.location.origin },
    });
    setLoading(false);
    if (err) setError(err.message);
    else setSent(true);
  }

  if (session === undefined) return null;

  if (!session) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#EBEEF2", fontFamily: "Inter, system-ui, sans-serif" }}>
        <form onSubmit={sendLink} style={{ background: "#fff", border: "1px solid #D9E0E8", borderRadius: 14, padding: 32, width: 360 }}>
          <div style={{ width: 40, height: 40, borderRadius: 10, background: "#0B5563", color: "#fff", display: "grid", placeItems: "center", fontWeight: 700, marginBottom: 16 }}>E</div>
          <h2 style={{ margin: "0 0 6px", fontSize: 18 }}>EBIM · Estimador de Costos</h2>
          <p style={{ margin: "0 0 18px", fontSize: 13, color: "#5E6E81" }}>Acceso exclusivo para cuentas @ebim.pe / @grupoebim.com</p>
          {sent ? (
            <p style={{ fontSize: 13, color: "#0F8A5F" }}>Revisa tu correo y haz clic en el enlace de acceso.</p>
          ) : (
            <>
              <input
                type="email"
                required
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="tu.nombre@ebim.pe"
                style={{ width: "100%", border: "1px solid #D9E0E8", borderRadius: 8, padding: "9px 11px", fontSize: 14, marginBottom: 12, boxSizing: "border-box" }}
              />
              <button type="submit" disabled={loading} style={{ width: "100%", background: "#0B5563", color: "#fff", border: "none", borderRadius: 8, padding: "10px", fontWeight: 600, cursor: "pointer" }}>
                {loading ? "Enviando…" : "Enviar enlace de acceso"}
              </button>
            </>
          )}
          {error && <p style={{ fontSize: 12.5, color: "#C13B3B", marginTop: 12 }}>{error}</p>}
        </form>
      </div>
    );
  }

  return children;
}
