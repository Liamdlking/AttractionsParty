import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState(null);
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function onSubmit(e) {
    e.preventDefault();
    setError("");
    if (!file) return;

    setBusy(true);
    try {
      const fd = new FormData();
      fd.append("book1", file);
      fd.append("password", password);

      const res = await fetch("/api/generate", { method: "POST", body: fd });
      if (!res.ok) {
        const txt = await res.text();
        throw new Error(txt || "Generation failed");
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "TagX_Output.zip";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err?.message || "Something went wrong");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{minHeight:"100vh", background:"#0b1020", color:"#e8ecff", fontFamily:"system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif"}}>
      <div style={{maxWidth:820, margin:"0 auto", padding:32}}>
        <div style={{background:"#121a33", border:"1px solid #23305e", borderRadius:16, padding:22, boxShadow:"0 10px 30px rgba(0,0,0,.25)"}}>
          <h1 style={{fontSize:22, margin:"0 0 8px"}}>Upload Book1 and download the completed files</h1>
          <p style={{margin:"0 0 14px", color:"#c9d2ff", lineHeight:1.4}}>
            Generates party sheets (per date), Tag X signs (4 per page), and Stompers signs (2 per page).
          </p>

          {error ? <p style={{color:"#ffb4b4"}}><b>Error:</b> {error}</p> : null}

          <form onSubmit={onSubmit}>
            <label style={{display:"block", margin:"14px 0 6px", color:"#c9d2ff"}}>Book1 file (.xlsx)</label>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setFile(e.target.files?.[0] || null)}
              required
              style={{width:"100%", padding:12, borderRadius:12, border:"1px solid #2a3970", background:"#0e1630", color:"#e8ecff"}}
            />

            <label style={{display:"block", margin:"14px 0 6px", color:"#c9d2ff"}}>Password (optional)</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Leave blank if not enabled"
              style={{width:"100%", padding:12, borderRadius:12, border:"1px solid #2a3970", background:"#0e1630", color:"#e8ecff"}}
            />

            <button
              type="submit"
              disabled={busy}
              style={{marginTop:16, width:"100%", padding:"12px 14px", border:0, borderRadius:12, background:"#6d5efc", color:"white", fontWeight:700, cursor:"pointer", opacity: busy ? 0.7 : 1}}
            >
              {busy ? "Generating…" : "Generate ZIP"}
            </button>
          </form>

          <div style={{fontSize:12, color:"#9aa5db", marginTop:14}}>
            Admin: set <code style={{background:"#0e1630", padding:"2px 6px", borderRadius:8, border:"1px solid #2a3970"}}>APP_PASSWORD</code> env var to enable password protection.
          </div>
        </div>
      </div>
    </div>
  );
}
