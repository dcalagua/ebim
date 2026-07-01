import { useEffect, useState } from "react";
import { Mail, CheckCircle2 } from "lucide-react";
import { supabase } from "./lib/supabaseClient";
import ebimLogo from "./assets/brand/ebim-logo-completo.png";

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
        <style>{`button:focus-visible, input:focus-visible { outline: 2px solid #056769; outline-offset: 2px; }`}</style>
        <form onSubmit={sendLink} style={{ background: "#fff", border: "1px solid #D9E0E8", borderRadius: 14, padding: 32, width: 360 }}>
          <img src={ebimLogo} alt="EBIM" style={{ height: 30, marginBottom: 18, display: "block" }} />
          <h2 style={{ margin: "0 0 6px", fontSize: 16, fontWeight: 600, color: "#15202E" }}>Estimador de Costos</h2>
          <p style={{ margin: "0 0 18px", fontSize: 13, color: "#5E6E81" }}>
            Sin contraseña: te enviamos un enlace de un solo uso a tu correo @ebim.pe / @grupoebim.com
          </p>
          {sent ? (
            <p style={{ fontSize: 13, color: "#0F8A5F", display: "flex", gap: 8, alignItems: "flex-start" }}>
              <CheckCircle2 size={16} style={{ flexShrink: 0, marginTop: 1 }} />
              <span>Enviamos un enlace a <b>{email}</b>. Ábrelo desde tu correo para entrar.</span>
            </p>
          ) : (
            <>
              <div style={{ position: "relative", marginBottom: 12 }}>
                <Mail size={16} style={{ position: "absolute", left: 11, top: 11, color: "#5E6E81" }} />
                <input
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="tu.nombre@ebim.pe"
                  style={{ width: "100%", border: "1px solid #D9E0E8", borderRadius: 8, padding: "9px 11px 9px 34px", fontSize: 14, boxSizing: "border-box" }}
                />
              </div>
              <button type="submit" disabled={loading} style={{ width: "100%", background: "#056769", color: "#fff", border: "none", borderRadius: 8, padding: "10px", fontWeight: 600, cursor: "pointer" }}>
                {loading ? "Enviando…" : "Enviarme el enlace de acceso"}
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
