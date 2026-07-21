// Envía cotizaciones (Excel + Propuesta + Inteligencia Comercial) por correo vía SMTP (Microsoft 365 / Outlook),
// autenticando con la contraseña guardada como secreto en Supabase. El front nunca ve esa contraseña;
// solo usuarios autenticados (verify_jwt) pueden invocar esta función. Mismo patrón que claude-proxy.
import { createClient } from "jsr:@supabase/supabase-js@2";
import { SMTPClient } from "https://deno.land/x/denomailer@1.6.0/mod.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: corsHeaders });

  try {
    const authHeader = req.headers.get("Authorization") ?? "";
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_ANON_KEY")!,
      { global: { headers: { Authorization: authHeader } } },
    );
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return new Response(JSON.stringify({ error: "No autorizado" }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const body = await req.json();
    const { to, subject, bodyText, attachments } = body;
    if (!to || !subject || !Array.isArray(attachments)) {
      return new Response(JSON.stringify({ error: "Faltan 'to', 'subject' o 'attachments'" }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const smtpUser = Deno.env.get("OFFICE365_USER") || "dcalagua@grupoebim.com";
    const smtpPassword = Deno.env.get("OFFICE365_PASSWORD");
    if (!smtpPassword) {
      return new Response(JSON.stringify({ error: "OFFICE365_PASSWORD no configurada en el servidor" }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const client = new SMTPClient({
      connection: {
        hostname: "smtp.office365.com",
        port: 587,
        tls: false, // STARTTLS: Office 365 exige upgrade sobre 587, no TLS implícito
        auth: { username: smtpUser, password: smtpPassword },
      },
    });

    await client.send({
      from: smtpUser,
      to,
      subject,
      content: bodyText || "",
      attachments: attachments.map((a: { filename: string; contentBase64: string; contentType: string }) => ({
        filename: a.filename,
        content: a.contentBase64,
        contentType: a.contentType,
        encoding: "base64" as const,
      })),
    });
    await client.close();

    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: String(err?.message || err) }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
