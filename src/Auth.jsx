import { useEffect, useState } from "react";
import { supabase } from "./lib/supabaseClient";
import ebimLogo from "./assets/brand/ebim-logo-completo.png";

// Ícono oficial de Microsoft (4 cuadros de color) — se usa tal cual en botones "Sign in with Microsoft".
function MicrosoftLogo() {
  return (
    <svg width="18" height="18" viewBox="0 0 21 21" aria-hidden="true">
      <rect x="1" y="1" width="9" height="9" fill="#F25022" />
      <rect x="11" y="1" width="9" height="9" fill="#7FBA00" />
      <rect x="1" y="11" width="9" height="9" fill="#00A4EF" />
      <rect x="11" y="11" width="9" height="9" fill="#FFB900" />
    </svg>
  );
}

export default function Auth({ children }) {
  const [session, setSession] = useState(undefined);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  async function signInWithMicrosoft() {
    setError(""); setLoading(true);
    const { error: err } = await supabase.auth.signInWithOAuth({
      provider: "azure",
      options: { scopes: "email openid profile", redirectTo: window.location.origin },
    });
    if (err) { setError(err.message); setLoading(false); }
    // En éxito, Azure redirige fuera de la página — no hay más que hacer aquí.
  }

  if (session === undefined) return null;

  if (!session) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#EBEEF2", fontFamily: "Inter, system-ui, sans-serif" }}>
        <style>{`button:focus-visible { outline: 2px solid #056769; outline-offset: 2px; }`}</style>
        <div style={{ background: "#fff", border: "1px solid #D9E0E8", borderRadius: 14, padding: 32, width: 360 }}>
          <img src={ebimLogo} alt="EBIM" style={{ height: 30, marginBottom: 18, display: "block" }} />
          <h2 style={{ margin: "0 0 6px", fontSize: 16, fontWeight: 600, color: "#15202E" }}>Estimador de Costos</h2>
          <p style={{ margin: "0 0 20px", fontSize: 13, color: "#5E6E81" }}>
            Acceso exclusivo para cuentas de Grupo EBIM (Microsoft 365).
          </p>
          <button
            type="button"
            onClick={signInWithMicrosoft}
            disabled={loading}
            style={{ width: "100%", display: "flex", alignItems: "center", justifyContent: "center", gap: 10, background: "#fff", color: "#15202E", border: "1px solid #8C8C8C", borderRadius: 4, padding: "10px", fontWeight: 600, fontSize: 14, cursor: "pointer" }}
          >
            <MicrosoftLogo />
            {loading ? "Redirigiendo…" : "Iniciar sesión con Microsoft 365"}
          </button>
          {error && <p style={{ fontSize: 12.5, color: "#C13B3B", marginTop: 12 }}>{error}</p>}
        </div>
      </div>
    );
  }

  return children;
}
