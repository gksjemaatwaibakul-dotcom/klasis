import "@/App.css";
import { FileCode2, FileText, BookOpen, Download, CheckCircle2 } from "lucide-react";

const FILES = [
  {
    key: "code-gs",
    name: "Code.gs",
    desc: "Backend Apps Script (REV 10) — peringatan kelas 11 + diagnosa, dashboard publik, chart.",
    href: "/download/Code_REV10_2026-10-06.gs?v=rev10",
    downloadName: "Code.gs",
    size: "173 KB",
    hash: "ab9c4e27b2e6e261",
    icon: FileCode2,
    accent: "accent-amber",
  },
  {
    key: "index-html",
    name: "Index.html",
    desc: "Frontend Web App (REV 10) — pembaca PDF anti-stuck, dashboard berisi nama kelas & guru.",
    href: "/download/Index_REV10_2026-10-06.html.txt?v=rev10",
    downloadName: "Index.html",
    size: "294 KB",
    hash: "9d399a7b46905147",
    icon: FileText,
    accent: "accent-sky",
  },
  {
    key: "panduan-md",
    name: "PANDUAN.md",
    desc: "Langkah menerapkan & deploy ulang ke project Apps Script Anda.",
    href: "/download/PANDUAN_REV10.md?v=rev10",
    downloadName: "PANDUAN.md",
    size: "4 KB",
    hash: "7ed525a84506df52",
    icon: BookOpen,
    accent: "accent-emerald",
  },
];

const STEPS = [
  "Buka project Apps Script (dari spreadsheet: Extensions \u2192 Apps Script).",
  "Timpa isi file lama dengan Code.gs dan Index.html yang baru diunduh.",
  "Simpan, lalu jalankan fungsi setupDatabase sekali dari editor.",
  "Deploy \u2192 Manage deployments \u2192 Edit \u2192 New version \u2192 Deploy.",
  "Set \u201cWho has access\u201d = Anyone (execute as: Me) agar mode Pengunjung aktif.",
];

export default function App() {
  return (
    <div className="dl-page" data-testid="download-page">
      <div className="dl-glow" aria-hidden="true" />
      <main className="dl-shell">
        <header className="dl-head">
          <span className="dl-eyebrow" data-testid="download-eyebrow">
            Google Apps Script · Siap Deploy
          </span>
          <h1 className="dl-title">
            Aplikasi Katekisasi
            <br />
            <span className="dl-title-sub">SMA Kristen Waibakul</span>
          </h1>
          <p className="dl-lead">
            Versi <strong style={{ color: "var(--amber)" }}>REV 10 · Build 2026-10-06</strong> —
            nama file berversi agar tidak tertukar cache unduhan lama. Setelah
            deploy, pastikan teks versi ini muncul di footer halaman login.
          </p>
        </header>

        <section className="dl-grid" data-testid="download-grid">
          {FILES.map((f) => {
            const Icon = f.icon;
            return (
              <a
                key={f.key}
                href={f.href}
                download={f.downloadName || true}
                className={`dl-card ${f.accent}`}
                data-testid={`download-${f.key}`}
              >
                <div className="dl-card-top">
                  <span className="dl-icon">
                    <Icon size={26} strokeWidth={1.75} />
                  </span>
                  <span className="dl-size">{f.size}</span>
                </div>
                <h2 className="dl-card-name">{f.name}</h2>
                <p className="dl-card-desc">{f.desc}</p>
                <p className="dl-card-hash" data-testid={`hash-${f.key}`}>
                  SHA-256: {f.hash}…
                </p>
                <span className="dl-card-cta">
                  <Download size={16} strokeWidth={2} />
                  Unduh
                </span>
              </a>
            );
          })}
        </section>

        <section className="dl-steps" data-testid="deploy-steps">
          <h3 className="dl-steps-title">Cara Deploy</h3>
          <ol className="dl-steps-list">
            {STEPS.map((s, i) => (
              <li key={i} className="dl-step" data-testid={`deploy-step-${i + 1}`}>
                <CheckCircle2 size={18} strokeWidth={2} className="dl-step-ic" />
                <span>{s}</span>
              </li>
            ))}
          </ol>
        </section>

        <footer className="dl-foot">
          Status baru tertulis <strong>“Alpa”</strong>, data lama{" "}
          <strong>“Alfa”</strong> tetap terbaca otomatis. Peringatan kini
          berbasis Jadwal Mengajar.
        </footer>
      </main>
    </div>
  );
}
