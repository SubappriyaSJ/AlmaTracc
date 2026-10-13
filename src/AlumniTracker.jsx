import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  GraduationCap, LayoutDashboard, Upload, Users, LogOut, FileSpreadsheet, CheckCircle2,
  AlertCircle, X, Search, ArrowLeft, Eye, MapPin, Building2, Briefcase, Check, Circle,
  ShieldCheck, Database, Link2, Lock,
} from "lucide-react";

/* ───────────────────────── constants & helpers ───────────────────────── */

const FIELDS = [
  { key: "company", label: "Current company", icon: Building2 },
  { key: "title", label: "Job title", icon: Briefcase },
  { key: "industry", label: "Industry", icon: Database },
  { key: "location", label: "Location", icon: MapPin },
  { key: "linkedin", label: "LinkedIn URL", icon: Link2 },
];
const REQUIRED_COLS = ["name", "email", "graduation_year"];
const ADMIN_INVITE = "ALUMNI-ADMIN"; // mock invite code for admin sign-up
const ADMIN_ROUTES = ["admin", "import", "directory", "alumnus"];
const INDUSTRIES = ["Technology", "Finance", "Healthcare", "Education", "Consulting", "Manufacturing", "Public sector", "Other"];

const daysAgo = (n) => Date.now() - n * 864e5;
const completeness = (a) => Math.round((FIELDS.filter((f) => a[f.key]?.trim()).length / FIELDS.length) * 100);
const initials = (n = "") => n.split(" ").map((s) => s[0]).slice(0, 2).join("").toUpperCase();
const ago = (t) => { const d = Math.floor((Date.now() - t) / 864e5); return d <= 0 ? "Today" : d === 1 ? "Yesterday" : `${d} days ago`; };

const SEED_ALUMNI = [
  { id: 1, name: "Priya Raman", email: "priya@alumni.edu", year: 2016, company: "Stripe", title: "Staff Engineer", industry: "Technology", location: "Singapore", linkedin: "linkedin.com/in/priyaraman", updatedAt: daysAgo(6) },
  { id: 2, name: "Daniel Okafor", email: "d.okafor@mail.com", year: 2012, company: "Goldman Sachs", title: "Vice President", industry: "Finance", location: "London, UK", linkedin: "", updatedAt: daysAgo(48) },
  { id: 3, name: "Mei Tanaka", email: "mei.tanaka@mail.com", year: 2019, company: "", title: "", industry: "", location: "Osaka, Japan", linkedin: "", updatedAt: daysAgo(210) },
  { id: 4, name: "Arjun Mehta", email: "arjun.m@mail.com", year: 2008, company: "Apollo Hospitals", title: "Head of Operations", industry: "Healthcare", location: "Chennai, India", linkedin: "linkedin.com/in/arjunmehta", updatedAt: daysAgo(12) },
  { id: 5, name: "Sofia Lindqvist", email: "sofia.l@mail.com", year: 2014, company: "McKinsey & Company", title: "Associate Partner", industry: "Consulting", location: "Stockholm, Sweden", linkedin: "linkedin.com/in/sofial", updatedAt: daysAgo(95) },
  { id: 6, name: "Tomás Herrera", email: "tomas.h@mail.com", year: 2021, company: "Rappi", title: "Product Analyst", industry: "", location: "", linkedin: "", updatedAt: daysAgo(3) },
  { id: 7, name: "Aisha Bello", email: "aisha.b@mail.com", year: 2010, company: "University of Lagos", title: "Lecturer", industry: "Education", location: "Lagos, Nigeria", linkedin: "", updatedAt: daysAgo(160) },
];
const SEED_ACCOUNTS = [
  { email: "admin@alumni.edu", password: "admin123", role: "admin", name: "Records Office" },
  { email: "priya@alumni.edu", password: "alumni123", role: "user", name: "Priya Raman", alumniId: 1 },
];

/* Single source of truth for route permissions. */
function guardRoute(session, route) {
  if (!session) return { name: "auth" };
  if (session.role !== "admin" && ADMIN_ROUTES.includes(route.name)) return { name: "profile" };
  if (route.name === "auth") return session.role === "admin" ? { name: "admin" } : { name: "profile" };
  return route;
}

const font = { body: { fontFamily: "'Hanken Grotesk', system-ui, sans-serif" }, serif: { fontFamily: "'Newsreader', Georgia, serif" } };

/* ───────────────────────── small UI primitives ───────────────────────── */

const Button = ({ variant = "primary", className = "", ...p }) => {
  const styles = {
    primary: "bg-zinc-900 text-white hover:bg-zinc-700 border-zinc-900",
    accent: "bg-teal-700 text-white hover:bg-teal-800 border-teal-700",
    outline: "bg-white text-zinc-800 hover:bg-zinc-50 border-zinc-300",
    ghost: "bg-transparent text-zinc-600 hover:bg-zinc-100 border-transparent",
  };
  return <button {...p} className={`inline-flex items-center justify-center gap-2 h-9 px-3.5 text-sm font-medium border rounded-md transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-teal-600 focus-visible:ring-offset-1 disabled:opacity-40 disabled:pointer-events-none ${styles[variant]} ${className}`} />;
};
const Input = ({ label, hint, ...p }) => (
  <label className="block">
    <span className="block text-sm font-medium text-zinc-800 mb-1.5">{label}</span>
    <input {...p} className="w-full h-10 px-3 text-sm bg-white border border-zinc-300 rounded-md placeholder:text-zinc-400 focus:outline-none focus:border-teal-700 focus:ring-1 focus:ring-teal-700 disabled:bg-zinc-50 disabled:text-zinc-500" />
    {hint && <span className="block text-xs text-zinc-500 mt-1">{hint}</span>}
  </label>
);
const Panel = ({ title, aside, children, className = "" }) => (
  <section className={`bg-white border border-zinc-200 rounded-lg ${className}`}>
    {(title || aside) && (
      <header className="flex items-center justify-between px-5 h-12 border-b border-zinc-200">
        <h2 className="text-sm font-semibold text-zinc-900">{title}</h2>{aside}
      </header>
    )}
    {children}
  </section>
);
const Bar = ({ value }) => (
  <div className="h-1.5 w-full bg-zinc-100 rounded-full overflow-hidden">
    <div className="h-full bg-teal-700 transition-all duration-500" style={{ width: `${value}%` }} />
  </div>
);

/* ───────────────────────── authentication ───────────────────────── */

function AuthPage({ accounts, alumni, onSignIn, onSignUp }) {
  const [mode, setMode] = useState("signin");
  const [role, setRole] = useState("user");
  const [f, setF] = useState({ name: "", email: "", password: "", code: "" });
  const [err, setErr] = useState("");
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  const submit = (e) => {
    e.preventDefault(); setErr("");
    if (mode === "signin") {
      const acc = accounts.find((a) => a.email === f.email.trim().toLowerCase() && a.password === f.password);
      if (!acc) return setErr("Email or password is incorrect. Check both and try again.");
      if (acc.role !== role) return setErr(`This account is registered as ${acc.role === "admin" ? "an admin" : "an alumnus"}. Switch the role above.`);
      return onSignIn(acc);
    }
    if (!f.name.trim() || !f.email.includes("@") || f.password.length < 6) return setErr("Enter your name, a valid email, and a password of at least 6 characters.");
    if (accounts.some((a) => a.email === f.email.trim().toLowerCase())) return setErr("An account with this email already exists. Sign in instead.");
    if (role === "admin" && f.code !== ADMIN_INVITE) return setErr("That invite code isn't valid. Ask the records office for a new one.");
    onSignUp({ name: f.name.trim(), email: f.email.trim().toLowerCase(), password: f.password, role });
  };

  return (
    <div className="min-h-screen grid lg:grid-cols-[minmax(0,5fr)_minmax(0,6fr)] bg-zinc-50" style={font.body}>
      <aside className="hidden lg:flex flex-col justify-between bg-zinc-900 text-zinc-300 p-10 border-r border-zinc-800">
        <div className="flex items-center gap-2.5 text-white"><GraduationCap size={22} /><span className="font-semibold">Alumni Register</span></div>
        <div>
          <p className="text-3xl leading-snug text-white max-w-md" style={font.serif}>Careers move. The register should keep up.</p>
          <p className="mt-4 text-sm text-zinc-400 max-w-sm">Alumni confirm where they work now; the records office keeps one accurate, current directory.</p>
        </div>
        <dl className="grid grid-cols-3 border-t border-zinc-700 pt-6 text-sm">
          {[["Records", String(alumni.length)], ["Fields tracked", "5"], ["Backend", "PostgreSQL"]].map(([k, v]) => (
            <div key={k}><dd className="text-white font-medium">{v}</dd><dt className="text-zinc-500 text-xs mt-0.5">{k}</dt></div>
          ))}
        </dl>
      </aside>

      <main className="flex items-center justify-center p-6">
        <form onSubmit={submit} className="w-full max-w-sm space-y-5">
          <div>
            <h1 className="text-2xl text-zinc-900" style={font.serif}>{mode === "signin" ? "Sign in" : "Create your account"}</h1>
            <p className="text-sm text-zinc-500 mt-1">{mode === "signin" ? "Welcome back." : "Takes under a minute."}</p>
          </div>
          <div className="grid grid-cols-2 p-1 bg-zinc-100 border border-zinc-200 rounded-md" role="tablist">
            {[["user", "Alumnus"], ["admin", "Admin"]].map(([v, l]) => (
              <button type="button" key={v} onClick={() => setRole(v)} role="tab" aria-selected={role === v}
                className={`h-8 text-sm rounded transition-colors ${role === v ? "bg-white text-zinc-900 font-medium border border-zinc-200" : "text-zinc-500 hover:text-zinc-800"}`}>{l}</button>
            ))}
          </div>
          {mode === "signup" && <Input label="Full name" value={f.name} onChange={set("name")} placeholder="As on your degree certificate" />}
          <Input label="Email" type="email" value={f.email} onChange={set("email")} placeholder="you@example.com" hint={mode === "signup" && role === "user" ? "Use the email on file and we'll link your existing record." : undefined} />
          <Input label="Password" type="password" value={f.password} onChange={set("password")} />
          {mode === "signup" && role === "admin" && <Input label="Admin invite code" value={f.code} onChange={set("code")} hint={`Demo code: ${ADMIN_INVITE}`} />}
          {err && <p className="flex gap-2 text-sm text-red-700 bg-red-50 border border-red-200 rounded-md p-3"><AlertCircle size={16} className="mt-0.5 shrink-0" />{err}</p>}
          <Button variant="accent" className="w-full h-10" type="submit">{mode === "signin" ? "Sign in" : "Create account"}</Button>
          <p className="text-sm text-zinc-500 text-center">
            {mode === "signin" ? "New here?" : "Already registered?"}{" "}
            <button type="button" className="text-teal-700 font-medium hover:underline" onClick={() => { setMode(mode === "signin" ? "signup" : "signin"); setErr(""); }}>
              {mode === "signin" ? "Create an account" : "Sign in"}
            </button>
          </p>
          {mode === "signin" && (
            <div className="border border-dashed border-zinc-300 rounded-md p-3 text-xs text-zinc-500 space-y-1">
              <p className="font-medium text-zinc-700">Demo accounts</p>
              <p>Admin: admin@alumni.edu / admin123</p>
              <p>Alumnus: priya@alumni.edu / alumni123</p>
            </div>
          )}
        </form>
      </main>
    </div>
  );
}

/* ───────────────────────── app shell ───────────────────────── */

function Shell({ session, route, go, onSignOut, viewAs, children }) {
  const isAdmin = session.role === "admin";
  const nav = isAdmin
    ? [["admin", "Overview", LayoutDashboard], ["import", "Import data", Upload], ["directory", "Directory", Users], ["profile", "User view", Eye]]
    : [["profile", "My profile", Users]];
  const active = route.name === "alumnus" ? "directory" : route.name;
  return (
    <div className="min-h-screen flex bg-zinc-50 text-zinc-900" style={font.body}>
      <aside className="w-56 shrink-0 bg-white border-r border-zinc-200 flex flex-col sticky top-0 h-screen">
        <div className="h-14 px-4 flex items-center gap-2 border-b border-zinc-200 font-semibold"><GraduationCap size={18} className="text-teal-700" />Alumni Register</div>
        <nav className="p-2 flex-1 space-y-0.5">
          {nav.map(([k, label, Icon]) => (
            <button key={k} onClick={() => go({ name: k })}
              className={`w-full flex items-center gap-2.5 h-9 px-3 text-sm rounded-md border transition-colors ${active === k ? "bg-zinc-100 border-zinc-200 text-zinc-900 font-medium" : "border-transparent text-zinc-600 hover:bg-zinc-50"}`}>
              <Icon size={16} />{label}
            </button>
          ))}
        </nav>
        <div className="p-3 border-t border-zinc-200 space-y-3">
          <div className="flex items-center gap-2.5">
            <div className="h-8 w-8 rounded-full bg-zinc-900 text-white text-xs font-medium grid place-items-center">{initials(session.name)}</div>
            <div className="min-w-0"><p className="text-sm font-medium truncate">{session.name}</p>
              <p className="text-xs text-zinc-500 flex items-center gap-1">{isAdmin && <ShieldCheck size={12} />}{isAdmin ? "Administrator" : "Alumnus"}</p></div>
          </div>
          <Button variant="outline" className="w-full" onClick={onSignOut}><LogOut size={14} />Sign out</Button>
        </div>
      </aside>
      <div className="flex-1 min-w-0">
        {isAdmin && viewAs && route.name === "profile" && (
          <div className="bg-amber-50 border-b border-amber-200 px-8 h-10 flex items-center justify-between text-sm text-amber-900">
            <span className="flex items-center gap-2"><Eye size={14} />Viewing as {viewAs.name}. Changes are disabled.</span>
            <button className="font-medium underline" onClick={() => go({ name: "directory" })}>Exit user view</button>
          </div>
        )}
        <main className="max-w-5xl mx-auto px-8 py-8">{children}</main>
      </div>
    </div>
  );
}

const PageHead = ({ title, sub, action }) => (
  <div className="flex items-end justify-between mb-6">
    <div><h1 className="text-3xl tracking-tight" style={font.serif}>{title}</h1>{sub && <p className="text-sm text-zinc-500 mt-1">{sub}</p>}</div>{action}
  </div>
);

/* ───────────────────────── admin: overview ───────────────────────── */

function Overview({ alumni, go, setViewAs }) {
  const m = useMemo(() => {
    const missing = alumni.filter((a) => completeness(a) < 100);
    const recent = alumni.filter((a) => Date.now() - a.updatedAt < 30 * 864e5);
    const avg = alumni.length ? Math.round(alumni.reduce((s, a) => s + completeness(a), 0) / alumni.length) : 0;
    return { missing, recent, avg };
  }, [alumni]);
  const stats = [
    ["Total alumni tracked", alumni.length, "in the register"],
    ["Missing details", m.missing.length, "records below 100%"],
    ["Updated in 30 days", m.recent.length, "confirmed recently"],
    ["Average completeness", `${m.avg}%`, "across all records"],
  ];
  const queue = [...m.missing].sort((a, b) => completeness(a) - completeness(b)).slice(0, 5);
  return (
    <>
      <PageHead title="Overview" sub="Where the register stands today." action={<Button variant="accent" onClick={() => go({ name: "import" })}><Upload size={14} />Import records</Button>} />
      <div className="grid grid-cols-4 border border-zinc-200 rounded-lg bg-white divide-x divide-zinc-200 mb-6">
        {stats.map(([l, v, s]) => (
          <div key={l} className="p-5"><p className="text-sm text-zinc-500">{l}</p><p className="text-3xl mt-2 tabular-nums" style={font.serif}>{v}</p><p className="text-xs text-zinc-400 mt-1">{s}</p></div>
        ))}
      </div>
      <Panel title="Needs follow-up" aside={<Button variant="ghost" onClick={() => go({ name: "directory" })}>View directory</Button>}>
        {queue.length === 0 ? <p className="p-5 text-sm text-zinc-500">Every record is complete.</p> : (
          <ul className="divide-y divide-zinc-200">
            {queue.map((a) => (
              <li key={a.id} className="flex items-center gap-4 px-5 py-3">
                <div className="h-8 w-8 rounded-full bg-zinc-100 border border-zinc-200 text-xs font-medium grid place-items-center">{initials(a.name)}</div>
                <div className="flex-1 min-w-0"><p className="text-sm font-medium">{a.name}</p><p className="text-xs text-zinc-500">Class of {a.year} · missing {FIELDS.filter((f) => !a[f.key]).length} of 5 fields</p></div>
                <div className="w-32"><Bar value={completeness(a)} /></div>
                <span className="w-10 text-right text-sm tabular-nums text-zinc-600">{completeness(a)}%</span>
                <Button variant="outline" onClick={() => go({ name: "alumnus", id: a.id })}>Inspect</Button>
              </li>
            ))}
          </ul>
        )}
      </Panel>
    </>
  );
}

/* ───────────────────────── admin: import ───────────────────────── */

function ImportView({ onCommit }) {
  const [file, setFile] = useState(null);
  const [phase, setPhase] = useState("idle"); // idle | uploading | ready | error | done
  const [progress, setProgress] = useState(0);
  const [rows, setRows] = useState([]);
  const [error, setError] = useState("");
  const [drag, setDrag] = useState(false);
  const inputRef = useRef(null);

  const reset = () => { setFile(null); setPhase("idle"); setProgress(0); setRows([]); setError(""); };

  const handleFile = async (f) => {
    if (!f) return;
    reset(); setFile(f);
    if (!/\.(xlsx|csv)$/i.test(f.name)) { setPhase("error"); return setError("Unsupported file type. Upload an .xlsx or .csv file."); }
    if (f.size > 10 * 1024 * 1024) { setPhase("error"); return setError("File is larger than 10 MB. Split it and upload in parts."); }
    setPhase("uploading");
    // mock upload progress
    for (let p = 0; p <= 100; p += 10) { setProgress(p); await new Promise((r) => setTimeout(r, 90)); }
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: "array" });
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
      const norm = raw.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k.trim().toLowerCase().replace(/\s+/g, "_"), String(v).trim()])));
      const cols = Object.keys(norm[0] || {});
      const missing = REQUIRED_COLS.filter((c) => !cols.includes(c));
      if (!norm.length) throw new Error("The file has no rows.");
      if (missing.length) throw new Error(`Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}. Add ${missing.length > 1 ? "them" : "it"} to the header row.`);
      setRows(norm.map((r) => ({ ...r, _valid: /\S+@\S+\.\S+/.test(r.email) && !!r.name && /^\d{4}$/.test(r.graduation_year) })));
      setPhase("ready");
    } catch (e) { setPhase("error"); setError(e.message || "This file couldn't be read."); }
  };

  const valid = rows.filter((r) => r._valid);
  const commit = () => { onCommit(valid); setPhase("done"); };

  return (
    <>
      <PageHead title="Import data" sub="Upload alumni records as .xlsx or .csv. Required columns: name, email, graduation_year." />
      <div className="space-y-6">
        {phase !== "done" && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
            onDrop={(e) => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files[0]); }}
            className={`border border-dashed rounded-lg bg-white transition-colors ${drag ? "border-teal-700 bg-teal-50/40" : "border-zinc-300"}`}>
            {!file ? (
              <div className="py-12 flex flex-col items-center text-center">
                <div className="h-10 w-10 rounded-md border border-zinc-200 grid place-items-center text-zinc-500 mb-3"><Upload size={18} /></div>
                <p className="text-sm font-medium">Drop a file here, or <button className="text-teal-700 underline" onClick={() => inputRef.current.click()}>browse</button></p>
                <p className="text-xs text-zinc-500 mt-1">.xlsx or .csv, up to 10 MB</p>
              </div>
            ) : (
              <div className="p-5 flex items-center gap-4">
                <FileSpreadsheet size={28} className="text-zinc-400 shrink-0" />
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between"><p className="text-sm font-medium truncate">{file.name}</p><span className="text-xs text-zinc-500">{(file.size / 1024).toFixed(1)} KB</span></div>
                  {phase === "uploading" && <div className="mt-2"><Bar value={progress} /><p className="text-xs text-zinc-500 mt-1.5">Uploading and validating… {progress}%</p></div>}
                  {phase === "ready" && <p className="text-xs text-teal-700 mt-1 flex items-center gap-1"><CheckCircle2 size={13} />Structure valid. {valid.length} of {rows.length} rows are ready to import.</p>}
                  {phase === "error" && <p className="text-xs text-red-700 mt-1 flex items-center gap-1"><AlertCircle size={13} />{error}</p>}
                </div>
                <Button variant="ghost" onClick={reset} aria-label="Remove file"><X size={16} /></Button>
              </div>
            )}
            <input ref={inputRef} type="file" accept=".xlsx,.csv" className="hidden" onChange={(e) => { handleFile(e.target.files[0]); e.target.value = ""; }} />
          </div>
        )}

        {phase === "done" && (
          <Panel className="p-8 text-center">
            <CheckCircle2 className="mx-auto text-teal-700" size={28} />
            <p className="mt-3 font-medium">{valid.length} records sent to the database</p>
            <p className="text-sm text-zinc-500 mt-1">New alumni appear in the directory; existing emails were updated.</p>
            <Button className="mt-4" variant="outline" onClick={reset}>Import another file</Button>
          </Panel>
        )}

        {phase === "ready" && (
          <Panel title="Preview" aside={
            <div className="flex items-center gap-3">
              <span className="text-xs text-zinc-500">Target: <code className="bg-zinc-100 px-1.5 py-0.5 rounded">public.alumni</code></span>
              <Button variant="accent" onClick={commit} disabled={!valid.length}><Database size={14} />Import {valid.length} records</Button>
            </div>}>
            <div className="overflow-x-auto max-h-96">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-zinc-50 border-b border-zinc-200 text-left text-zinc-500">
                  <tr>{["Status", "Name", "Email", "Year", "Company", "Title", "Location"].map((h) => <th key={h} className="px-4 h-9 font-medium whitespace-nowrap">{h}</th>)}</tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {rows.slice(0, 50).map((r, i) => (
                    <tr key={i} className={r._valid ? "" : "bg-red-50/60"}>
                      <td className="px-4 h-10">{r._valid ? <Check size={15} className="text-teal-700" /> : <span className="text-xs text-red-700 flex items-center gap-1"><AlertCircle size={13} />Invalid</span>}</td>
                      <td className="px-4 font-medium whitespace-nowrap">{r.name || "—"}</td><td className="px-4 text-zinc-600">{r.email || "—"}</td>
                      <td className="px-4 tabular-nums">{r.graduation_year || "—"}</td><td className="px-4">{r.company || "—"}</td>
                      <td className="px-4">{r.title || "—"}</td><td className="px-4">{r.location || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {rows.length > 50 && <p className="px-5 py-3 text-xs text-zinc-500 border-t border-zinc-200">Showing the first 50 of {rows.length} rows.</p>}
          </Panel>
        )}
      </div>
    </>
  );
}

/* ───────────────────────── admin: directory & detail ───────────────────────── */

function Directory({ alumni, go }) {
  const [q, setQ] = useState("");
  const list = alumni.filter((a) => `${a.name} ${a.email} ${a.company}`.toLowerCase().includes(q.toLowerCase()));
  return (
    <>
      <PageHead title="Directory" sub={`${alumni.length} alumni in the register.`} />
      <Panel title="All alumni" aside={
        <div className="relative"><Search size={14} className="absolute left-2.5 top-2.5 text-zinc-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search name, email, company" className="h-8 w-64 pl-8 pr-3 text-sm border border-zinc-300 rounded-md focus:outline-none focus:border-teal-700" /></div>}>
        <table className="w-full text-sm">
          <thead className="bg-zinc-50 border-b border-zinc-200 text-left text-zinc-500"><tr>{["Alumnus", "Class", "Current role", "Completeness", "Last update", ""].map((h) => <th key={h} className="px-5 h-9 font-medium">{h}</th>)}</tr></thead>
          <tbody className="divide-y divide-zinc-100">
            {list.map((a) => (
              <tr key={a.id} className="hover:bg-zinc-50 cursor-pointer" onClick={() => go({ name: "alumnus", id: a.id })}>
                <td className="px-5 py-3"><p className="font-medium">{a.name}</p><p className="text-xs text-zinc-500">{a.email}</p></td>
                <td className="px-5 tabular-nums">{a.year}</td>
                <td className="px-5">{a.title ? `${a.title}, ${a.company}` : <span className="text-zinc-400">Not provided</span>}</td>
                <td className="px-5 w-44"><div className="flex items-center gap-2"><Bar value={completeness(a)} /><span className="text-xs tabular-nums w-8 text-zinc-600">{completeness(a)}%</span></div></td>
                <td className="px-5 text-zinc-500 whitespace-nowrap">{ago(a.updatedAt)}</td>
                <td className="px-5 text-right"><Eye size={15} className="inline text-zinc-400" /></td>
              </tr>
            ))}
            {!list.length && <tr><td colSpan={6} className="px-5 py-10 text-center text-zinc-500">No alumni match “{q}”.</td></tr>}
          </tbody>
        </table>
      </Panel>
    </>
  );
}

function AlumnusDetail({ alumnus, go, openAsUser }) {
  if (!alumnus) return <p className="text-sm text-zinc-500">That record no longer exists.</p>;
  return (
    <>
      <button onClick={() => go({ name: "directory" })} className="flex items-center gap-1.5 text-sm text-zinc-500 hover:text-zinc-900 mb-4"><ArrowLeft size={14} />Directory</button>
      <PageHead title={alumnus.name} sub={`${alumnus.email} · Class of ${alumnus.year}`} action={<Button variant="outline" onClick={() => openAsUser(alumnus)}><Eye size={14} />Open in user view</Button>} />
      <div className="grid grid-cols-3 gap-6">
        <Panel title="Professional details" className="col-span-2">
          <dl className="divide-y divide-zinc-100">
            {FIELDS.map(({ key, label, icon: I }) => (
              <div key={key} className="flex items-center px-5 h-12"><dt className="w-48 flex items-center gap-2 text-sm text-zinc-500"><I size={14} />{label}</dt>
                <dd className="text-sm">{alumnus[key] || <span className="text-zinc-400">Not provided</span>}</dd></div>
            ))}
          </dl>
        </Panel>
        <Panel title="Record status" className="self-start">
          <div className="p-5 space-y-3"><p className="text-3xl tabular-nums" style={font.serif}>{completeness(alumnus)}%</p><Bar value={completeness(alumnus)} /><p className="text-xs text-zinc-500">Last updated {ago(alumnus.updatedAt).toLowerCase()}.</p></div>
        </Panel>
      </div>
    </>
  );
}

/* ───────────────────────── user (alumni) dashboard ───────────────────────── */

function UserDashboard({ alumnus, readOnly, onSave }) {
  const [form, setForm] = useState({ company: alumnus.company, title: alumnus.title, industry: alumnus.industry, location: alumnus.location, linkedin: alumnus.linkedin });
  const [saved, setSaved] = useState(false);
  const set = (k) => (e) => { setForm({ ...form, [k]: e.target.value }); setSaved(false); };
  const live = { ...alumnus, ...form };
  const pct = completeness(live);
  const dirty = FIELDS.some((f) => form[f.key] !== alumnus[f.key]);
  const submit = (e) => { e.preventDefault(); onSave(form); setSaved(true); };

  return (
    <>
      <PageHead title={readOnly ? alumnus.name : `Hello, ${alumnus.name.split(" ")[0]}`} sub="Keep your professional details current so classmates and the university can reach you." />
      <div className="grid grid-cols-[minmax(0,3fr)_minmax(0,2fr)] gap-6 items-start">
        <Panel title="Professional details" aside={readOnly && <span className="text-xs text-zinc-500 flex items-center gap-1"><Lock size={12} />Read only</span>}>
          <form onSubmit={submit} className="p-5 space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <Input label="Current company" value={form.company} onChange={set("company")} disabled={readOnly} placeholder="e.g. Acme Corp" />
              <Input label="Job title" value={form.title} onChange={set("title")} disabled={readOnly} placeholder="e.g. Product Manager" />
            </div>
            <label className="block"><span className="block text-sm font-medium mb-1.5">Industry</span>
              <select value={form.industry} onChange={set("industry")} disabled={readOnly} className="w-full h-10 px-3 text-sm bg-white border border-zinc-300 rounded-md focus:outline-none focus:border-teal-700 disabled:bg-zinc-50 disabled:text-zinc-500">
                <option value="">Select an industry</option>{INDUSTRIES.map((i) => <option key={i}>{i}</option>)}
              </select></label>
            <Input label="Location" value={form.location} onChange={set("location")} disabled={readOnly} placeholder="City, Country" />
            <Input label="LinkedIn URL" value={form.linkedin} onChange={set("linkedin")} disabled={readOnly} placeholder="linkedin.com/in/your-name" />
            {!readOnly && (
              <div className="flex items-center gap-3 pt-1">
                <Button variant="accent" type="submit" disabled={!dirty}>Save changes</Button>
                {saved && <span className="text-sm text-teal-700 flex items-center gap-1"><CheckCircle2 size={14} />Saved</span>}
              </div>
            )}
          </form>
        </Panel>

        <Panel title="Profile completeness">
          <div className="p-5">
            <div className="flex items-baseline justify-between mb-2"><span className="text-3xl tabular-nums" style={font.serif}>{pct}%</span><span className="text-xs text-zinc-500">{FIELDS.filter((f) => live[f.key]?.trim()).length} of 5 complete</span></div>
            <Bar value={pct} />
            <ol className="mt-5 relative">
              <span className="absolute left-[9px] top-2 bottom-2 w-px bg-zinc-200" aria-hidden />
              <li className="relative flex gap-3 pb-4"><CheckCircle2 size={19} className="text-teal-700 bg-white z-10" /><div><p className="text-sm font-medium">Account created</p><p className="text-xs text-zinc-500">Class of {alumnus.year}</p></div></li>
              {FIELDS.map(({ key, label }, i) => {
                const done = !!live[key]?.trim();
                return (
                  <li key={key} className={`relative flex gap-3 ${i < FIELDS.length - 1 ? "pb-4" : ""}`}>
                    {done ? <CheckCircle2 size={19} className="text-teal-700 bg-white z-10" /> : <Circle size={19} className="text-zinc-300 bg-white z-10" />}
                    <div><p className={`text-sm ${done ? "font-medium" : "text-zinc-500"}`}>{label}</p>
                      <p className="text-xs text-zinc-500">{done ? live[key] : "Not added yet"}</p></div>
                  </li>
                );
              })}
            </ol>
          </div>
        </Panel>
      </div>
    </>
  );
}

/* ───────────────────────── root: state + conditional routing ───────────────────────── */

export default function App() {
  const [accounts, setAccounts] = useState(SEED_ACCOUNTS);
  const [alumni, setAlumni] = useState(SEED_ALUMNI);
  const [session, setSession] = useState(null);
  const [requested, setRequested] = useState({ name: "auth" });
  const [viewAs, setViewAs] = useState(null); // alumnus an admin is inspecting in user view

  const route = guardRoute(session, requested); // every render passes through the guard
  const go = (r) => { if (r.name === "profile" && session?.role === "admin" && !viewAs) r = { name: "directory" }; setRequested(r); };

  const signIn = (acc) => { setSession(acc); setRequested({ name: acc.role === "admin" ? "admin" : "profile" }); };
  const signUp = (acc) => {
    if (acc.role === "user") {
      const existing = alumni.find((a) => a.email === acc.email);
      if (existing) acc.alumniId = existing.id;
      else { const id = Date.now(); setAlumni((l) => [...l, { id, name: acc.name, email: acc.email, year: new Date().getFullYear(), company: "", title: "", industry: "", location: "", linkedin: "", updatedAt: Date.now() }]); acc.alumniId = id; }
    }
    setAccounts((l) => [...l, acc]); signIn(acc);
  };
  const signOut = () => { setSession(null); setViewAs(null); setRequested({ name: "auth" }); };

  const importRows = (rows) => setAlumni((cur) => {
    const next = [...cur];
    rows.forEach((r) => {
      const rec = { name: r.name, email: r.email.toLowerCase(), year: Number(r.graduation_year), company: r.company || "", title: r.title || r.job_title || "", industry: r.industry || "", location: r.location || "", linkedin: r.linkedin || r.linkedin_url || "", updatedAt: Date.now() };
      const i = next.findIndex((a) => a.email === rec.email);
      if (i >= 0) next[i] = { ...next[i], ...Object.fromEntries(Object.entries(rec).filter(([, v]) => v !== "")) };
      else next.push({ id: Date.now() + next.length, ...rec });
    });
    return next;
  });
  const saveProfile = (id, patch) => setAlumni((l) => l.map((a) => (a.id === id ? { ...a, ...patch, updatedAt: Date.now() } : a)));
  const openAsUser = (a) => { setViewAs(a); setRequested({ name: "profile" }); };

  const css = <style>{`@import url('https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;500;600&family=Newsreader:opsz,wght@6..72,400;6..72,500&display=swap');`}</style>;

  if (route.name === "auth") return <>{css}<AuthPage accounts={accounts} alumni={alumni} onSignIn={signIn} onSignUp={signUp} /></>;

  const isAdmin = session.role === "admin";
  const me = isAdmin ? viewAs : alumni.find((a) => a.id === session.alumniId);
  const inspected = route.name === "alumnus" ? alumni.find((a) => a.id === route.id) : null;

  return (
    <>
      {css}
      <Shell session={session} route={route} go={go} onSignOut={signOut} viewAs={viewAs}>
        {route.name === "admin" && <Overview alumni={alumni} go={go} setViewAs={setViewAs} />}
        {route.name === "import" && <ImportView onCommit={importRows} />}
        {route.name === "directory" && <Directory alumni={alumni} go={go} />}
        {route.name === "alumnus" && <AlumnusDetail alumnus={inspected} go={go} openAsUser={openAsUser} />}
        {route.name === "profile" && me && (
          <UserDashboard key={me.id} alumnus={me} readOnly={isAdmin} onSave={(patch) => saveProfile(me.id, patch)} />
        )}
      </Shell>
    </>
  );
}
