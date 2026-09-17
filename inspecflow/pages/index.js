import React, { useState, useMemo, useRef, useEffect } from "react";
import {
  Building2, Users, ClipboardCheck, Table2, Plus, Trash2, Pencil, X, Check,
  Upload, Download, LogOut, ImageIcon, ShieldCheck, CircleAlert, Maximize2,
  Search, BarChart3, LayoutGrid, Eye, EyeOff,
} from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "../lib/supabaseClient";

const REGIONAIS = ["Centro", "Oeste"];
const MESES_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

/* ---------------------------------------------------------------
   Mapeamento banco (snake_case) <-> app (camelCase)
---------------------------------------------------------------- */

function mapEmpresa(r) {
  return { id: r.id, nome: r.nome, pontoFocal: r.ponto_focal, login: r.login, equipes: r.equipes, userId: r.user_id };
}
function mapInspetor(r) {
  return { id: r.id, empresaId: r.empresa_id, nome: r.nome, funcao: r.funcao, regional: r.regional, metaCiclo1: r.meta_ciclo1, metaCiclo2: r.meta_ciclo2, metaCiclo3: r.meta_ciclo3 };
}
function mapInspecao(r) {
  return {
    id: r.id, empresaId: r.empresa_id, inspetorId: r.inspetor_id, regional: r.regional, tipo: r.tipo,
    dataInspecao: r.data_inspecao, horaRegistro: (r.hora_registro || "").slice(0, 5),
    evidencia: r.evidencia_url, status: r.status, justificativa: r.justificativa || "",
    criadoEm: new Date(r.created_at).getTime(),
  };
}

/* ---------------------------------------------------------------
   Helpers gerais
---------------------------------------------------------------- */

function fmtDate(iso) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}
function getMes(iso) { return iso ? iso.slice(0, 7) : null; }
function getCiclo(iso) {
  if (!iso) return null;
  const day = Number(iso.split("-")[2]);
  if (day <= 10) return 1;
  if (day <= 20) return 2;
  return 3;
}
function mesLabel(mesStr) {
  if (!mesStr) return "—";
  const [y, m] = mesStr.split("-");
  return `${MESES_PT[Number(m) - 1]}/${y}`;
}
function distinctMeses(inspecoes) {
  const set = new Set(inspecoes.map((i) => getMes(i.dataInspecao)).filter(Boolean));
  const list = [...set].sort().reverse();
  return list.length > 0 ? list : [new Date().toISOString().slice(0, 7)];
}
function tipoLabel(tipo) { return tipo === "abono" ? "Abono" : "Inspeção"; }

function StatusBadge({ status }) {
  const label = status === "aprovado" ? "Aprovado" : status === "reprovado" ? "Reprovado" : status === "abonado" ? "Abonado" : "Pendente";
  return <span className={`insp-badge ${status}`}>{label}</span>;
}

function cycleSums(inspetoresList) {
  return inspetoresList.reduce(
    (acc, i) => {
      acc.c1 += Number(i.metaCiclo1) || 0;
      acc.c2 += Number(i.metaCiclo2) || 0;
      acc.c3 += Number(i.metaCiclo3) || 0;
      return acc;
    },
    { c1: 0, c2: 0, c3: 0 }
  );
}

function faixaCiclo(pct) {
  if (pct >= 80) return { label: "EFICIENTE", color: "var(--good)" };
  if (pct >= 50) return { label: "ATENÇÃO", color: "var(--warn)" };
  return { label: "CRÍTICO", color: "var(--bad)" };
}
function faixaRanking(pct) {
  if (pct >= 80) return { label: "No alvo", color: "var(--good)" };
  if (pct >= 50) return { label: "Atenção", color: "var(--warn)" };
  return { label: "Crítico", color: "var(--bad)" };
}

function computeEmpresaStats(empresaId, mes, inspetoresAll, inspecoesAll, regionalFiltro = "todas") {
  const inspetores = inspetoresAll.filter((i) => i.empresaId === empresaId && (regionalFiltro === "todas" || i.regional === regionalFiltro));
  const inspecoesDoMes = inspecoesAll.filter((i) => i.empresaId === empresaId && getMes(i.dataInspecao) === mes);

  const porInspetor = inspetores.map((insp) => {
    const ciclos = [1, 2, 3].map((c) => {
      const metaOriginal = Number(insp[`metaCiclo${c}`]) || 0;
      const doCiclo = inspecoesDoMes.filter((x) => x.inspetorId === insp.id && getCiclo(x.dataInspecao) === c);
      const abonos = doCiclo.filter((x) => x.tipo === "abono" && x.status === "abonado").length;
      const metaEfetiva = Math.max(0, metaOriginal - abonos);
      const realizado = doCiclo.filter((x) => x.tipo === "inspecao" && x.status === "aprovado").length;
      const valido = Math.min(realizado, metaEfetiva);
      return { ciclo: c, metaOriginal, metaEfetiva, realizado, valido };
    });
    const metaSum = ciclos.reduce((a, c) => a + c.metaEfetiva, 0);
    const validoSum = ciclos.reduce((a, c) => a + c.valido, 0);
    const eficiencia = metaSum > 0 ? (validoSum / metaSum) * 100 : 100;
    const temInspecao = ciclos.some((c) => c.realizado > 0);
    return { inspetor: insp, ciclos, metaSum, validoSum, eficiencia, temInspecao };
  });

  const cicloTotais = [1, 2, 3].map((c) => {
    const metaTotal = porInspetor.reduce((a, p) => a + p.ciclos[c - 1].metaEfetiva, 0);
    const realizadoTotal = porInspetor.reduce((a, p) => a + p.ciclos[c - 1].realizado, 0);
    const validoTotal = porInspetor.reduce((a, p) => a + p.ciclos[c - 1].valido, 0);
    const pct = metaTotal > 0 ? (validoTotal / metaTotal) * 100 : 100;
    return { ciclo: c, metaTotal, realizadoTotal, validoTotal, pct };
  });

  const metaGeral = cicloTotais.reduce((a, c) => a + c.metaTotal, 0);
  const validoGeral = cicloTotais.reduce((a, c) => a + c.validoTotal, 0);
  const realizadoGeral = cicloTotais.reduce((a, c) => a + c.realizadoTotal, 0);
  const excedentes = Math.max(0, realizadoGeral - validoGeral);
  const eficienciaGeral = metaGeral > 0 ? (validoGeral / metaGeral) * 100 : 100;

  const semInspecao = porInspetor.filter((p) => !p.temInspecao).length;
  const comInspecao = porInspetor.length - semInspecao;

  const statusCounts = { aprovado: 0, abonado: 0, reprovado: 0 };
  inspecoesDoMes.forEach((x) => { if (statusCounts[x.status] !== undefined) statusCounts[x.status]++; });

  return { inspetores: porInspetor, cicloTotais, metaGeral, validoGeral, realizadoGeral, excedentes, eficienciaGeral, semInspecao, comInspecao, statusCounts };
}

/* ---------------------------------------------------------------
   Estilos
---------------------------------------------------------------- */

const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');
    .insp-root {
      --ink:#171b1f; --ink-soft:#565f68; --line:#dde1e4; --line-soft:#eceff1; --bg:#f5f6f4; --panel:#ffffff;
      --rail:#171b1f; --rail-soft:#2a2f35; --accent:#2c5f7c; --accent-soft:#e4edf1;
      --good:#2f7d54; --good-soft:#e5f2ea; --bad:#b5432f; --bad-soft:#fbe9e5; --warn:#b8801f; --warn-soft:#faf0dd;
      font-family:'Inter',sans-serif; color:var(--ink); background:var(--bg); min-height:100vh; width:100%;
      display:flex; font-size:14px; line-height:1.5;
    }
    .insp-root * { box-sizing: border-box; }
    .insp-h { font-family:'Barlow Condensed',sans-serif; font-weight:600; letter-spacing:0.01em; }
    .insp-rail { width:240px; flex-shrink:0; background:var(--rail); color:#eef0f1; display:flex; flex-direction:column; padding:22px 14px; min-height:100vh; }
    .insp-rail-brand { display:flex; align-items:center; gap:9px; padding:4px 8px 20px 8px; border-bottom:1px solid #383e45; margin-bottom:16px; }
    .insp-rail-brand span { font-family:'Barlow Condensed',sans-serif; font-weight:700; font-size:19px; }
    .insp-rail-session { background:var(--rail-soft); border-radius:8px; padding:10px 12px; margin-bottom:18px; }
    .insp-rail-session .role { font-size:10.5px; text-transform:uppercase; letter-spacing:0.08em; color:#9aa4ad; margin-bottom:3px; }
    .insp-rail-session .name { font-size:13.5px; font-weight:600; color:#fff; }
    .insp-nav-btn { display:flex; align-items:center; gap:10px; width:100%; text-align:left; padding:9px 10px; border-radius:7px; border:none; background:transparent; color:#c4cad0; font-size:13.5px; font-weight:500; cursor:pointer; margin-bottom:2px; }
    .insp-nav-btn:hover { background:#23272d; color:#fff; }
    .insp-nav-btn.active { background:var(--accent); color:#fff; }
    .insp-logout { margin-top:auto; display:flex; align-items:center; gap:8px; color:#9aa4ad; background:none; border:none; font-size:12.5px; padding:8px 10px; cursor:pointer; border-radius:7px; }
    .insp-logout:hover { color:#fff; background:#23272d; }
    .insp-rail-footer { text-align:center; font-size:10.5px; color:#565d64; padding-top:12px; margin-top:10px; border-top:1px solid #2c3138; }
    .insp-main { flex:1; padding:34px 40px; max-width:1280px; }
    .insp-page-title { font-size:27px; margin:0 0 3px 0; }
    .insp-page-sub { color:var(--ink-soft); font-size:13.5px; margin:0 0 26px 0; }
    .insp-card { background:var(--panel); border:1px solid var(--line-soft); border-radius:10px; }
    .insp-toolbar { display:flex; align-items:center; justify-content:space-between; margin-bottom:14px; gap:12px; flex-wrap:wrap; }
    .insp-btn { display:inline-flex; align-items:center; gap:7px; font-size:13px; font-weight:600; padding:9px 14px; border-radius:7px; border:1px solid transparent; cursor:pointer; background:var(--accent); color:#fff; }
    .insp-btn:hover { background:#234c63; }
    .insp-btn.secondary { background:#fff; color:var(--ink); border-color:var(--line); }
    .insp-btn.secondary:hover { background:#f4f5f5; }
    .insp-btn.good { background:var(--good); } .insp-btn.good:hover { background:#266644; }
    .insp-btn.bad { background:var(--bad); } .insp-btn.bad:hover { background:#973a29; }
    .insp-btn.ghost { background:transparent; color:var(--ink-soft); border-color:var(--line); }
    .insp-btn:disabled { opacity:0.45; cursor:not-allowed; }
    .insp-table { width:100%; border-collapse:collapse; font-size:13px; }
    .insp-table th { text-align:left; font-weight:600; color:var(--ink-soft); font-size:11.5px; text-transform:uppercase; letter-spacing:0.04em; padding:11px 14px; border-bottom:1px solid var(--line); background:#fafafa; }
    .insp-table td { padding:12px 14px; border-bottom:1px solid var(--line-soft); vertical-align:middle; }
    .insp-table tr:last-child td { border-bottom:none; }
    .insp-table tr:hover td { background:#fbfbfa; }
    .insp-badge { display:inline-flex; align-items:center; gap:5px; padding:3px 10px; border-radius:100px; font-size:12px; font-weight:600; }
    .insp-badge.aprovado { background:var(--good-soft); color:var(--good); }
    .insp-badge.reprovado { background:var(--bad-soft); color:var(--bad); }
    .insp-badge.pendente { background:var(--warn-soft); color:var(--warn); }
    .insp-badge.abonado { background:var(--accent-soft); color:var(--accent); }
    .insp-icon-btn { border:none; background:none; cursor:pointer; color:var(--ink-soft); padding:5px; border-radius:5px; display:inline-flex; }
    .insp-icon-btn:hover { background:var(--line-soft); color:var(--ink); }
    .insp-field { margin-bottom:14px; }
    .insp-field label { display:block; font-size:12.5px; font-weight:600; color:var(--ink-soft); margin-bottom:5px; }
    .insp-input, .insp-select, .insp-textarea { width:100%; padding:9px 11px; border:1px solid var(--line); border-radius:7px; font-size:13.5px; background:#fff; color:var(--ink); }
    .insp-input:focus, .insp-select:focus, .insp-textarea:focus { outline:2px solid var(--accent); border-color:var(--accent); }
    .insp-textarea { resize:vertical; min-height:70px; }
    .insp-hint { font-size:12px; color:var(--ink-soft); margin-top:5px; }
    .insp-grid3 { display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px; }
    .insp-modal-overlay { position:fixed; inset:0; background:rgba(23,27,31,0.45); display:flex; align-items:center; justify-content:center; z-index:50; padding:20px; }
    .insp-modal { background:#fff; border-radius:12px; width:100%; max-width:440px; padding:24px; max-height:90vh; overflow-y:auto; }
    .insp-modal-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:18px; }
    .insp-modal-head h3 { margin:0; font-size:19px; }
    .insp-empty { padding:50px 20px; text-align:center; color:var(--ink-soft); }
    .insp-empty svg { margin:0 auto 10px auto; display:block; opacity:0.5; }
    .insp-login-wrap { min-height:100vh; width:100%; display:flex; align-items:center; justify-content:center; background:var(--rail); padding:20px; }
    .insp-login-card { background:#fff; border-radius:14px; padding:36px; width:100%; max-width:400px; }
    .insp-login-footer { text-align:center; font-size:11px; color:var(--ink-soft); margin-top:22px; opacity:0.7; }
    .insp-val-card { border:1px solid var(--line-soft); border-radius:10px; padding:16px; background:#fff; margin-bottom:12px; }
    .insp-val-top { display:flex; gap:14px; }
    .insp-val-evidence-wrap { position:relative; flex-shrink:0; width:96px; height:96px; }
    .insp-val-evidence { width:96px; height:96px; border-radius:8px; object-fit:cover; border:1px solid var(--line); cursor:zoom-in; display:block; }
    .insp-val-evidence-expand { position:absolute; bottom:4px; right:4px; background:rgba(255,255,255,0.92); border-radius:6px; padding:4px; border:1px solid var(--line); cursor:pointer; display:flex; }
    .insp-val-evidence-empty { width:96px; height:96px; border-radius:8px; flex-shrink:0; border:1px dashed var(--line); display:flex; align-items:center; justify-content:center; color:var(--ink-soft); background:#fafafa; }
    .insp-val-meta { flex:1; min-width:0; }
    .insp-val-meta .name { font-size:15.5px; font-weight:600; margin-bottom:2px; }
    .insp-val-meta .sub { font-size:12.5px; color:var(--ink-soft); }
    .insp-val-tags { display:flex; gap:6px; margin-top:8px; flex-wrap:wrap; }
    .insp-tag { font-size:11.5px; padding:2px 8px; border-radius:100px; background:var(--line-soft); color:var(--ink-soft); font-weight:600; }
    .insp-val-actions { display:flex; gap:8px; margin-top:14px; }
    .insp-upload-box { border:1.5px dashed var(--line); border-radius:9px; padding:18px; text-align:center; cursor:pointer; background:#fafafa; }
    .insp-upload-box:hover { border-color:var(--accent); background:var(--accent-soft); }
    .insp-upload-box img { max-height:110px; border-radius:7px; margin-bottom:8px; }
    .insp-alert { display:flex; align-items:flex-start; gap:8px; background:var(--warn-soft); color:#7a5514; padding:10px 12px; border-radius:8px; font-size:12.5px; margin-bottom:8px; border:1px solid #f0dba8; }
    .insp-alert svg { flex-shrink:0; margin-top:1px; color:#b8801f; }
    .insp-filters { display:flex; gap:8px; flex-wrap:wrap; }
    .insp-lightbox-inner { position:relative; max-width:92vw; max-height:92vh; }
    .insp-lightbox-close { position:absolute; top:-14px; right:-14px; background:#fff; border-radius:50%; box-shadow:0 2px 10px rgba(0,0,0,0.28); padding:6px; border:none; cursor:pointer; display:flex; }
    .insp-lightbox-img { max-width:92vw; max-height:92vh; border-radius:10px; display:block; }
    .insp-search-input-wrap { position:relative; max-width:400px; margin-bottom:20px; }
    .insp-search-input-wrap svg { position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--ink-soft); }
    .insp-search-input-wrap input { padding-left:34px; }
    .insp-detail-grid { display:grid; grid-template-columns:1fr 1fr; gap:10px; font-size:12.5px; margin-bottom:12px; }
    .insp-detail-grid .lbl { color:var(--ink-soft); margin-bottom:2px; }
    .insp-detail-grid .val { font-weight:600; }
    .insp-footer { text-align:center; font-size:11px; color:var(--ink-soft); opacity:0.55; padding:30px 0 10px 0; }
    .insp-dash-head { background:var(--rail); color:#fff; border-radius:10px; padding:18px 22px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:14px; margin-bottom:20px; }
    .insp-tabbar { display:flex; background:#2a2f35; border-radius:7px; padding:3px; }
    .insp-tabbar button { border:none; padding:6px 12px; border-radius:5px; font-size:12.5px; font-weight:600; cursor:pointer; background:transparent; color:#c4cad0; }
    .insp-tabbar button.active { background:#fff; color:var(--ink); }
    .insp-stat-card { padding:18px; text-align:center; }
    .insp-stat-card .lbl { font-size:11.5px; font-weight:700; color:var(--ink-soft); text-transform:uppercase; letter-spacing:0.05em; margin-bottom:10px; }
    .insp-stat-card .big { font-size:32px; font-weight:700; }
    .insp-loading-wrap { min-height:100vh; width:100%; display:flex; align-items:center; justify-content:center; background:var(--bg); color:var(--ink-soft); font-size:13.5px; }
  `}</style>
);

/* ---------------------------------------------------------------
   Mini componentes visuais
---------------------------------------------------------------- */

function ImageLightbox({ src, onClose }) {
  if (!src) return null;
  return (
    <div className="insp-modal-overlay" onClick={onClose} style={{ zIndex: 80 }}>
      <div className="insp-lightbox-inner" onClick={(e) => e.stopPropagation()}>
        <button className="insp-lightbox-close" onClick={onClose}><X size={18} /></button>
        <img className="insp-lightbox-img" src={src} alt="evidência ampliada" />
      </div>
    </div>
  );
}

function Donut({ segments, size = 118, thickness = 16, centerTop, centerBottom }) {
  const total = segments.reduce((a, s) => a + s.value, 0);
  const radius = (size - thickness) / 2;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;
  return (
    <div style={{ position: "relative", width: size, height: size }}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="var(--line-soft)" strokeWidth={thickness} />
        <g transform={`rotate(-90 ${size / 2} ${size / 2})`}>
          {total > 0 && segments.map((s, idx) => {
            if (s.value <= 0) return null;
            const frac = s.value / total;
            const dash = frac * circumference;
            const el = (
              <circle key={idx} cx={size / 2} cy={size / 2} r={radius} fill="none" stroke={s.color}
                strokeWidth={thickness} strokeDasharray={`${dash} ${circumference - dash}`} strokeDashoffset={-offset} />
            );
            offset += dash;
            return el;
          })}
        </g>
      </svg>
      <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center" }}>
        {centerTop}{centerBottom}
      </div>
    </div>
  );
}

function Legend({ items }) {
  return (
    <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 11.5, color: "var(--ink-soft)", justifyContent: "center" }}>
      {items.map((it, idx) => (
        <div key={idx} style={{ display: "flex", alignItems: "center", gap: 5 }}>
          <span style={{ width: 9, height: 9, borderRadius: "50%", background: it.color, display: "inline-block" }} />
          {it.label}
        </div>
      ))}
    </div>
  );
}

function CicloCard({ titulo, faixaLabel, pct, meta, realizado, valido, color }) {
  return (
    <div className="insp-card" style={{ padding: 18 }}>
      <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>{titulo}</div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 10 }}>
        <span style={{ fontSize: 26, fontWeight: 700 }}>{Math.round(pct)}%</span>
        <span style={{ fontSize: 11.5, fontWeight: 700, color, textTransform: "uppercase", letterSpacing: "0.04em" }}>{faixaLabel}</span>
      </div>
      <div style={{ height: 8, borderRadius: 100, background: "var(--line-soft)", overflow: "hidden", marginBottom: 14 }}>
        <div style={{ height: "100%", width: `${Math.min(100, pct)}%`, background: color, borderRadius: 100 }} />
      </div>
      <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>
        Meta: <strong style={{ color: "var(--ink)" }}>{meta}</strong>{"  "}Realizado: <strong style={{ color: "var(--ink)" }}>{realizado}</strong>{"  "}Válido: <strong style={{ color: "var(--ink)" }}>{valido}</strong>
      </div>
    </div>
  );
}

function BarRow({ nome, pct }) {
  const faixa = faixaRanking(pct);
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 9 }}>
      <div style={{ width: 84, fontSize: 11.5, fontWeight: 600, color: "var(--ink-soft)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{nome}</div>
      <div style={{ flex: 1, height: 14, borderRadius: 4, background: "var(--line-soft)", overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${Math.min(100, pct)}%`, background: faixa.color, borderRadius: 4 }} />
      </div>
      <div style={{ width: 38, fontSize: 11.5, fontWeight: 700, textAlign: "right" }}>{Math.round(pct)}%</div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Login (Supabase Auth de verdade)
---------------------------------------------------------------- */

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [loading, setLoading] = useState(false);

  async function entrar(e) {
    e.preventDefault();
    setErro(""); setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password: senha });
    if (error) setErro("E-mail ou senha inválidos.");
    setLoading(false);
  }

  return (
    <div className="insp-root insp-login-wrap">
      <GlobalStyle />
      <div>
        <form className="insp-login-card" onSubmit={entrar}>
          <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 22 }}>
            <ShieldCheck size={22} color="#2c5f7c" />
            <span className="insp-h" style={{ fontSize: 21, fontWeight: 700 }}>Cilco+10Flow</span>
          </div>
          <h1 className="insp-h" style={{ fontSize: 22, margin: "0 0 4px 0" }}>Acessar sistema</h1>
          <p style={{ color: "var(--ink-soft)", fontSize: 13, margin: "0 0 20px 0" }}>
            Entre com o e-mail e senha cadastrados pelo administrador.
          </p>
          <div className="insp-field">
            <label>E-mail</label>
            <input className="insp-input" type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="seuemail@empresa.com" />
          </div>
          <div className="insp-field">
            <label>Senha</label>
            <input className="insp-input" type="password" required value={senha} onChange={(e) => setSenha(e.target.value)} placeholder="••••••••" />
          </div>
          {erro && <div style={{ color: "var(--bad)", fontSize: 12.5, marginBottom: 12 }}>{erro}</div>}
          <button className="insp-btn" style={{ width: "100%", justifyContent: "center" }} disabled={loading} type="submit">
            {loading ? "Entrando..." : "Entrar"}
          </button>
        </form>
        <div className="insp-login-footer">feito por matheus manske</div>
      </div>
    </div>
  );
}

function SemAcessoScreen({ onSair }) {
  return (
    <div className="insp-root insp-login-wrap">
      <GlobalStyle />
      <div className="insp-login-card" style={{ textAlign: "center" }}>
        <CircleAlert size={26} color="#b5432f" style={{ margin: "0 auto 12px auto" }} />
        <h1 className="insp-h" style={{ fontSize: 19, margin: "0 0 8px 0" }}>Usuário sem acesso vinculado</h1>
        <p style={{ color: "var(--ink-soft)", fontSize: 13, marginBottom: 20 }}>
          Seu login funcionou, mas ainda não está vinculado a nenhuma empresa nem marcado como administrador.
          Peça para o ADM configurar seu acesso no Supabase.
        </p>
        <button className="insp-btn secondary" onClick={onSair}>Sair</button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Sidebar
---------------------------------------------------------------- */

function Sidebar({ session, empresas, screen, setScreen, onLogout }) {
  const isAdm = session.tipo === "adm";
  const empresa = !isAdm ? empresas.find((e) => e.id === session.empresaId) : null;

  const admNav = [
    { key: "validacao", label: "Validação", icon: ClipboardCheck },
    { key: "base", label: "Base de inspeções", icon: Table2 },
    { key: "pesquisa", label: "Pesquisar", icon: Search },
    { key: "painel-empresa", label: "Painel da empresa", icon: LayoutGrid },
    { key: "painel-geral", label: "Painel geral", icon: BarChart3 },
    { key: "empresas", label: "Empresas", icon: Building2 },
    { key: "inspetores", label: "Inspetores", icon: Users },
  ];
  const empresaNav = [
    { key: "nova-inspecao", label: "Nova inspeção", icon: Plus },
    { key: "minhas-inspecoes", label: "Minhas inspeções", icon: Table2 },
    { key: "inspetores", label: "Meus inspetores", icon: Users },
    { key: "pesquisa", label: "Pesquisar", icon: Search },
    { key: "painel-empresa", label: "Meu painel", icon: LayoutGrid },
    { key: "painel-geral", label: "Painel geral", icon: BarChart3 },
  ];
  const nav = isAdm ? admNav : empresaNav;

  return (
    <div className="insp-rail">
      <div className="insp-rail-brand"><ShieldCheck size={20} /><span>Ciclo+10Flow</span></div>
      <div className="insp-rail-session">
        <div className="role">{isAdm ? "Administrador" : "Acesso empresa"}</div>
        <div className="name">{isAdm ? "Você" : empresa?.nome}</div>
      </div>
      {nav.map((item) => (
        <button key={item.key} className={`insp-nav-btn ${screen === item.key ? "active" : ""}`} onClick={() => setScreen(item.key)}>
          <item.icon size={16} />{item.label}
        </button>
      ))}
      <button className="insp-logout" onClick={onLogout}><LogOut size={14} /> Sair</button>
      <div className="insp-rail-footer">feito por matheus manske</div>
    </div>
  );
}

/* ---------------------------------------------------------------
   ADM — Empresas
---------------------------------------------------------------- */

function EmpresasScreen({ empresas, inspetores, reload }) {
  const [modal, setModal] = useState(null);
  const [salvando, setSalvando] = useState(false);

  function openNovo() { setModal({ mode: "novo", data: { nome: "", pontoFocal: "", login: "", equipes: "", userId: "" } }); }
  function openEditar(emp) { setModal({ mode: "editar", data: { id: emp.id, nome: emp.nome, pontoFocal: emp.pontoFocal, login: emp.login, equipes: String(emp.equipes ?? 0), userId: emp.userId || "" } }); }

  async function salvar(data) {
    setSalvando(true);
    const payload = { nome: data.nome, ponto_focal: data.pontoFocal, login: data.login, equipes: Number(data.equipes) || 0, user_id: data.userId.trim() || null };
    if (modal.mode === "novo") await supabase.from("empresas").insert(payload);
    else await supabase.from("empresas").update(payload).eq("id", data.id);
    setSalvando(false); setModal(null); reload();
  }
  async function remover(id) {
    if (inspetores.some((i) => i.empresaId === id)) { alert("Essa empresa tem inspetores cadastrados. Remova os inspetores primeiro."); return; }
    if (confirm("Remover esta empresa?")) { await supabase.from("empresas").delete().eq("id", id); reload(); }
  }

  return (
    <div>
      <h1 className="insp-h insp-page-title">Empresas</h1>
      <p className="insp-page-sub">Cadastro das empresas com acesso ao sistema. Apenas o ADM edita esta tela.</p>
      <div className="insp-toolbar">
        <div />
        <button className="insp-btn" onClick={openNovo}><Plus size={15} /> Nova empresa</button>
      </div>
      <div className="insp-card">
        <table className="insp-table">
          <thead><tr><th>Empresa</th><th>Ponto focal</th><th>Login</th><th>Inspetores</th><th>Equipes a inspecionar</th><th style={{ width: 90 }}></th></tr></thead>
          <tbody>
            {empresas.map((e) => (
              <tr key={e.id}>
                <td style={{ fontWeight: 600 }}>{e.nome}</td>
                <td>{e.pontoFocal}</td>
                <td>{e.login}</td>
                <td>{inspetores.filter((i) => i.empresaId === e.id).length}</td>
                <td>{e.equipes ?? 0}</td>
                <td>
                  <div style={{ display: "flex", gap: 2, justifyContent: "flex-end" }}>
                    <button className="insp-icon-btn" onClick={() => openEditar(e)}><Pencil size={15} /></button>
                    <button className="insp-icon-btn" onClick={() => remover(e.id)}><Trash2 size={15} /></button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {modal && <EmpresaModal modal={modal} salvando={salvando} onClose={() => setModal(null)} onSave={salvar} />}
    </div>
  );
}

function EmpresaModal({ modal, salvando, onClose, onSave }) {
  const [form, setForm] = useState(modal.data);
  const valid = form.nome.trim() && form.pontoFocal.trim() && form.login.trim() && form.equipes !== "";
  return (
    <div className="insp-modal-overlay" onClick={onClose}>
      <div className="insp-modal" onClick={(e) => e.stopPropagation()}>
        <div className="insp-modal-head"><h3 className="insp-h">{modal.mode === "novo" ? "Nova empresa" : "Editar empresa"}</h3><button className="insp-icon-btn" onClick={onClose}><X size={18} /></button></div>
        <div className="insp-field"><label>Nome da empresa</label><input className="insp-input" value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} /></div>
        <div className="insp-field"><label>Ponto focal do processo</label><input className="insp-input" value={form.pontoFocal} onChange={(e) => setForm({ ...form, pontoFocal: e.target.value })} /></div>
        <div className="insp-field"><label>Login de referência</label><input className="insp-input" value={form.login} onChange={(e) => setForm({ ...form, login: e.target.value })} placeholder="usuario.empresa" /></div>
        <div className="insp-field">
          <label>Quantidade de equipes</label>
          <input className="insp-input" type="number" min="0" value={form.equipes} onChange={(e) => setForm({ ...form, equipes: e.target.value })} />
        </div>
        <div className="insp-field">
          <label>ID do usuário (Supabase Auth)</label>
          <input className="insp-input" value={form.userId} onChange={(e) => setForm({ ...form, userId: e.target.value })} placeholder="cole aqui o UID do usuário desta empresa" />
          <div className="insp-hint">Crie o login em Authentication → Users no Supabase e cole o UID (identificador) dele aqui para vincular.</div>
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 20 }}>
          <button className="insp-btn secondary" onClick={onClose}>Cancelar</button>
          <button className="insp-btn" disabled={!valid || salvando} onClick={() => onSave(form)}>{salvando ? "Salvando..." : "Salvar"}</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Inspetores
---------------------------------------------------------------- */

function InspetoresScreen({ inspetores, empresas, restrictedEmpresaId, reload }) {
  const [modal, setModal] = useState(null);
  const [filtroEmpresa, setFiltroEmpresa] = useState(restrictedEmpresaId || "todas");

  const lista = inspetores.filter((i) => restrictedEmpresaId ? i.empresaId === restrictedEmpresaId : filtroEmpresa === "todas" || i.empresaId === filtroEmpresa);

  function openNovo() {
    setModal({ mode: "novo", data: { empresaId: restrictedEmpresaId || empresas[0]?.id || "", nome: "", funcao: "", regional: REGIONAIS[0], metaCiclo1: "", metaCiclo2: "", metaCiclo3: "" } });
  }
  function openEditar(insp) { setModal({ mode: "editar", data: { ...insp } }); }

  async function salvar(data) {
    const payload = { empresa_id: data.empresaId, nome: data.nome, funcao: data.funcao, regional: data.regional, meta_ciclo1: Number(data.metaCiclo1) || 0, meta_ciclo2: Number(data.metaCiclo2) || 0, meta_ciclo3: Number(data.metaCiclo3) || 0 };
    if (modal.mode === "novo") await supabase.from("inspetores").insert(payload);
    else await supabase.from("inspetores").update(payload).eq("id", data.id);
    setModal(null); reload();
  }
  async function remover(id) { if (confirm("Remover este inspetor?")) { await supabase.from("inspetores").delete().eq("id", id); reload(); } }

  const empresasParaAlerta = restrictedEmpresaId ? empresas.filter((e) => e.id === restrictedEmpresaId) : filtroEmpresa === "todas" ? empresas : empresas.filter((e) => e.id === filtroEmpresa);
  const alertas = empresasParaAlerta.map((emp) => {
    const seus = inspetores.filter((i) => i.empresaId === emp.id);
    const soma = cycleSums(seus);
    const equipes = Number(emp.equipes) || 0;
    const deficits = [];
    if (soma.c1 < equipes) deficits.push({ ciclo: 1, atual: soma.c1 });
    if (soma.c2 < equipes) deficits.push({ ciclo: 2, atual: soma.c2 });
    if (soma.c3 < equipes) deficits.push({ ciclo: 3, atual: soma.c3 });
    return { emp, equipes, deficits };
  }).filter((a) => a.deficits.length > 0 && a.equipes > 0);

  return (
    <div>
      <h1 className="insp-h insp-page-title">{restrictedEmpresaId ? "Meus inspetores" : "Inspetores"}</h1>
      <p className="insp-page-sub">{restrictedEmpresaId ? "Cadastre e gerencie os inspetores da sua empresa, com a meta de cada ciclo." : "Cadastro de inspetores por empresa, com função, regional e meta por ciclo."}</p>
      <div className="insp-toolbar">
        {restrictedEmpresaId ? <div /> : (
          <select className="insp-select" style={{ width: 220 }} value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)}>
            <option value="todas">Todas as empresas</option>
            {empresas.map((e) => <option key={e.id} value={e.id}>{e.nome}</option>)}
          </select>
        )}
        <button className="insp-btn" onClick={openNovo}><Plus size={15} /> Novo inspetor</button>
      </div>
      {alertas.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          {alertas.map((a) => (
            <div className="insp-alert" key={a.emp.id}>
              <CircleAlert size={15} />
              <div><strong>{a.emp.nome}:</strong> é necessário que a soma das metas dos inspetores seja igual ou maior que o número de equipes ({a.equipes}). Abaixo do mínimo em: {a.deficits.map((d) => `Ciclo ${d.ciclo} (${d.atual}/${a.equipes})`).join(", ")}.</div>
            </div>
          ))}
        </div>
      )}
      <div className="insp-card">
        <table className="insp-table">
          <thead><tr><th>Nome</th>{!restrictedEmpresaId && <th>Empresa</th>}<th>Função</th><th>Regional</th><th>Ciclo 1</th><th>Ciclo 2</th><th>Ciclo 3</th><th>Total</th><th style={{ width: 90 }}></th></tr></thead>
          <tbody>
            {lista.length === 0 && <tr><td colSpan={restrictedEmpresaId ? 8 : 9}><div className="insp-empty">Nenhum inspetor cadastrado.</div></td></tr>}
            {lista.map((i) => {
              const total = (Number(i.metaCiclo1) || 0) + (Number(i.metaCiclo2) || 0) + (Number(i.metaCiclo3) || 0);
              return (
                <tr key={i.id}>
                  <td style={{ fontWeight: 600 }}>{i.nome}</td>
                  {!restrictedEmpresaId && <td>{empresas.find((e) => e.id === i.empresaId)?.nome || "—"}</td>}
                  <td>{i.funcao}</td><td>{i.regional}</td>
                  <td>{i.metaCiclo1 ?? 0}</td><td>{i.metaCiclo2 ?? 0}</td><td>{i.metaCiclo3 ?? 0}</td>
                  <td style={{ fontWeight: 700 }}>{total}</td>
                  <td><div style={{ display: "flex", gap: 2, justifyContent: "flex-end" }}>
                    <button className="insp-icon-btn" onClick={() => openEditar(i)}><Pencil size={15} /></button>
                    <button className="insp-icon-btn" onClick={() => remover(i.id)}><Trash2 size={15} /></button>
                  </div></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {modal && <InspetorModal modal={modal} empresas={empresas} restrictedEmpresaId={restrictedEmpresaId} onClose={() => setModal(null)} onSave={salvar} />}
    </div>
  );
}

function InspetorModal({ modal, empresas, restrictedEmpresaId, onClose, onSave }) {
  const [form, setForm] = useState(modal.data);
  const valid = form.empresaId && form.nome.trim() && form.funcao.trim() && form.regional && form.metaCiclo1 !== "" && form.metaCiclo2 !== "" && form.metaCiclo3 !== "";
  return (
    <div className="insp-modal-overlay" onClick={onClose}>
      <div className="insp-modal" onClick={(e) => e.stopPropagation()}>
        <div className="insp-modal-head"><h3 className="insp-h">{modal.mode === "novo" ? "Novo inspetor" : "Editar inspetor"}</h3><button className="insp-icon-btn" onClick={onClose}><X size={18} /></button></div>
        {!restrictedEmpresaId && (
          <div className="insp-field"><label>Empresa</label>
            <select className="insp-select" value={form.empresaId} onChange={(e) => setForm({ ...form, empresaId: e.target.value })}>
              {empresas.map((e) => <option key={e.id} value={e.id}>{e.nome}</option>)}
            </select>
          </div>
        )}
        <div className="insp-field"><label>Nome do inspetor</label><input className="insp-input" value={form.nome} onChange={(e) => setForm({ ...form, nome: e.target.value })} /></div>
        <div className="insp-field"><label>Função</label><input className="insp-input" value={form.funcao} onChange={(e) => setForm({ ...form, funcao: e.target.value })} /></div>
        <div className="insp-field"><label>Regional</label>
          <select className="insp-select" value={form.regional} onChange={(e) => setForm({ ...form, regional: e.target.value })}>
            {REGIONAIS.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <div className="insp-field"><label>Meta por ciclo</label>
          <div className="insp-grid3">
            <input className="insp-input" type="number" min="0" placeholder="Ciclo 1" value={form.metaCiclo1} onChange={(e) => setForm({ ...form, metaCiclo1: e.target.value })} />
            <input className="insp-input" type="number" min="0" placeholder="Ciclo 2" value={form.metaCiclo2} onChange={(e) => setForm({ ...form, metaCiclo2: e.target.value })} />
            <input className="insp-input" type="number" min="0" placeholder="Ciclo 3" value={form.metaCiclo3} onChange={(e) => setForm({ ...form, metaCiclo3: e.target.value })} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 20 }}>
          <button className="insp-btn secondary" onClick={onClose}>Cancelar</button>
          <button className="insp-btn" disabled={!valid} onClick={() => onSave(form)}>Salvar</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Empresa — Nova inspeção
---------------------------------------------------------------- */

function NovaInspecaoScreen({ empresaId, empresas, inspetores, inspecoes, reload }) {
  const empresa = empresas.find((e) => e.id === empresaId);
  const meusInspetores = inspetores.filter((i) => i.empresaId === empresaId);
  const regionais = [...new Set(meusInspetores.map((i) => i.regional))];

  const [tipo, setTipo] = useState("inspecao");
  const [regional, setRegional] = useState("");
  const [inspetorId, setInspetorId] = useState("");
  const [data, setData] = useState("");
  const [arquivo, setArquivo] = useState(null);
  const [preview, setPreview] = useState(null);
  const [enviando, setEnviando] = useState(false);
  const fileRef = useRef(null);

  const inspetoresDaRegional = meusInspetores.filter((i) => i.regional === regional);
  const valid = regional && inspetorId && data && arquivo;

  function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setArquivo(file);
    setPreview(URL.createObjectURL(file));
  }

  async function registrar() {
    setEnviando(true);
    let evidenciaUrl = null;
    if (arquivo) {
      const path = `${empresaId}/${Date.now()}-${arquivo.name.replace(/\s+/g, "_")}`;
      const { error: upErr } = await supabase.storage.from("evidencias").upload(path, arquivo);
      if (upErr) { alert("Erro ao enviar a evidência: " + upErr.message); setEnviando(false); return; }
      const { data: pub } = supabase.storage.from("evidencias").getPublicUrl(path);
      evidenciaUrl = pub.publicUrl;
    }
    const now = new Date();
    const { error } = await supabase.from("inspecoes").insert({
      empresa_id: empresaId, inspetor_id: inspetorId, regional, tipo,
      data_inspecao: data, hora_registro: now.toTimeString().slice(0, 8),
      evidencia_url: evidenciaUrl, status: "pendente",
    });
    if (error) { alert("Erro ao registrar: " + error.message); setEnviando(false); return; }

    setTipo("inspecao"); setRegional(""); setInspetorId(""); setData(""); setArquivo(null); setPreview(null);
    if (fileRef.current) fileRef.current.value = "";
    setEnviando(false);
    reload();
  }

  const minhas = inspecoes.filter((i) => i.empresaId === empresaId).sort((a, b) => b.criadoEm - a.criadoEm).slice(0, 6);

  return (
    <div>
      <h1 className="insp-h insp-page-title">Nova inspeção</h1>
      <p className="insp-page-sub">{empresa?.nome} — registre a inspeção e envie a evidência para validação.</p>
      <div className="insp-card" style={{ padding: 22, maxWidth: 460, marginBottom: 28 }}>
        <div className="insp-field"><label>Tipo</label>
          <select className="insp-select" value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option value="inspecao">Inspeção</option><option value="abono">Abono</option>
          </select>
          <div className="insp-hint">Abono aprovado reduz a meta do ciclo do inspetor (não conta como realizado).</div>
        </div>
        <div className="insp-field"><label>Regional</label>
          <select className="insp-select" value={regional} onChange={(e) => { setRegional(e.target.value); setInspetorId(""); }}>
            <option value="">Selecione a regional</option>
            {regionais.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>
        <div className="insp-field"><label>Inspetor</label>
          <select className="insp-select" value={inspetorId} onChange={(e) => setInspetorId(e.target.value)} disabled={!regional}>
            <option value="">{regional ? "Selecione o inspetor" : "Escolha a regional primeiro"}</option>
            {inspetoresDaRegional.map((i) => <option key={i.id} value={i.id}>{i.nome} — {i.funcao}</option>)}
          </select>
        </div>
        <div className="insp-field"><label>Dia da inspeção</label><input className="insp-input" type="date" value={data} onChange={(e) => setData(e.target.value)} /></div>
        <div className="insp-field"><label>Evidência</label>
          <div className="insp-upload-box" onClick={() => fileRef.current?.click()}>
            {preview ? (<><img src={preview} alt="evidência selecionada" /><div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>{arquivo?.name} — clique para trocar</div></>) : (
              <><Upload size={20} style={{ marginBottom: 6 }} color="#565f68" /><div style={{ fontSize: 13, fontWeight: 600 }}>Enviar foto da evidência</div><div className="insp-hint">JPG ou PNG</div></>
            )}
          </div>
          <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handleFile} />
        </div>
        <button className="insp-btn" style={{ width: "100%", justifyContent: "center", marginTop: 4 }} disabled={!valid || enviando} onClick={registrar}>
          <Check size={15} /> {enviando ? "Enviando..." : `Registrar ${tipo === "abono" ? "abono" : "inspeção"}`}
        </button>
      </div>
      <h2 className="insp-h" style={{ fontSize: 18, marginBottom: 10 }}>Últimos envios</h2>
      <div className="insp-card">
        <table className="insp-table">
          <thead><tr><th>Inspetor</th><th>Tipo</th><th>Regional</th><th>Data</th><th>Status</th></tr></thead>
          <tbody>
            {minhas.length === 0 && <tr><td colSpan={5}><div className="insp-empty">Nenhuma inspeção enviada ainda.</div></td></tr>}
            {minhas.map((ins) => {
              const insp = inspetores.find((i) => i.id === ins.inspetorId);
              return (
                <tr key={ins.id}>
                  <td style={{ fontWeight: 600 }}>{insp?.nome || "—"}</td>
                  <td>{tipoLabel(ins.tipo)}</td><td>{ins.regional}</td><td>{fmtDate(ins.dataInspecao)}</td>
                  <td><StatusBadge status={ins.status} /></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Empresa — Minhas inspeções
---------------------------------------------------------------- */

function MinhasInspecoesScreen({ empresaId, empresas, inspetores, inspecoes, getCodigo }) {
  const empresa = empresas.find((e) => e.id === empresaId);
  const minhas = inspecoes.filter((i) => i.empresaId === empresaId).sort((a, b) => b.criadoEm - a.criadoEm);

  function exportarExcel() {
    const dados = minhas.map((ins) => {
      const insp = inspetores.find((i) => i.id === ins.inspetorId);
      return {
        ID: getCodigo(ins.id),
        Inspetor: insp?.nome || "",
        Função: insp?.funcao || "",
        Tipo: tipoLabel(ins.tipo),
        Regional: ins.regional,
        Data: fmtDate(ins.dataInspecao),
        Hora: ins.horaRegistro,
        Status: ins.status === "aprovado" ? "Aprovado" : ins.status === "reprovado" ? "Reprovado" : ins.status === "abonado" ? "Abonado" : "Pendente",
        Justificativa: ins.status === "reprovado" ? ins.justificativa || "" : "—",
      };
    });
    const ws = XLSX.utils.json_to_sheet(dados);
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Minhas Inspeções");
    XLSX.writeFile(wb, `minhas_inspecoes_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div>
      <h1 className="insp-h insp-page-title">Minhas inspeções</h1>
      <p className="insp-page-sub">{empresa?.nome} — histórico completo e status de validação.</p>
      <div className="insp-toolbar">
        <div />
        <button className="insp-btn secondary" onClick={exportarExcel}><Download size={15} /> Exportar Excel</button>
      </div>
      <div className="insp-card">
        <table className="insp-table">
          <thead><tr><th>ID</th><th>Inspetor</th><th>Função</th><th>Tipo</th><th>Regional</th><th>Data</th><th>Hora</th><th>Status</th><th>Justificativa</th></tr></thead>
          <tbody>
            {minhas.length === 0 && <tr><td colSpan={9}><div className="insp-empty">Nenhuma inspeção registrada.</div></td></tr>}
            {minhas.map((ins) => {
              const insp = inspetores.find((i) => i.id === ins.inspetorId);
              return (
                <tr key={ins.id}>
                  <td style={{ fontFamily: "monospace", fontSize: 12, color: "var(--ink-soft)" }}>{getCodigo(ins.id)}</td>
                  <td style={{ fontWeight: 600 }}>{insp?.nome || "—"}</td>
                  <td>{insp?.funcao || "—"}</td><td>{tipoLabel(ins.tipo)}</td><td>{ins.regional}</td>
                  <td>{fmtDate(ins.dataInspecao)}</td><td>{ins.horaRegistro}</td>
                  <td><StatusBadge status={ins.status} /></td>
                  <td style={{ color: "var(--ink-soft)" }}>{ins.status === "reprovado" ? ins.justificativa : "—"}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   ADM — Validação
---------------------------------------------------------------- */

function ValidacaoScreen({ inspecoes, inspetores, empresas, reload }) {
  const [justificativaAberta, setJustificativaAberta] = useState(null);
  const [textoJustificativa, setTextoJustificativa] = useState("");
  const [lightbox, setLightbox] = useState(null);

  const pendentes = inspecoes.filter((i) => i.status === "pendente").sort((a, b) => a.criadoEm - b.criadoEm);

  async function aprovar(ins) {
    await supabase.from("inspecoes").update({ status: ins.tipo === "abono" ? "abonado" : "aprovado", justificativa: null }).eq("id", ins.id);
    reload();
  }
  function abrirReprovar(id) { setJustificativaAberta(id); setTextoJustificativa(""); }
  async function confirmarReprovar(id) {
    if (!textoJustificativa.trim()) return;
    await supabase.from("inspecoes").update({ status: "reprovado", justificativa: textoJustificativa.trim() }).eq("id", id);
    setJustificativaAberta(null); setTextoJustificativa(""); reload();
  }

  return (
    <div>
      <h1 className="insp-h insp-page-title">Validação de inspeções</h1>
      <p className="insp-page-sub">{pendentes.length} inspeção(ões) aguardando aprovação.</p>
      {pendentes.length === 0 && <div className="insp-card"><div className="insp-empty"><ClipboardCheck size={30} />Nenhuma inspeção pendente no momento.</div></div>}
      {pendentes.map((ins) => {
        const insp = inspetores.find((i) => i.id === ins.inspetorId);
        const emp = empresas.find((e) => e.id === ins.empresaId);
        const aberta = justificativaAberta === ins.id;
        return (
          <div className="insp-val-card" key={ins.id}>
            <div className="insp-val-top">
              {ins.evidencia ? (
                <div className="insp-val-evidence-wrap">
                  <img className="insp-val-evidence" src={ins.evidencia} alt="evidência" onClick={() => setLightbox(ins.evidencia)} />
                  <button className="insp-val-evidence-expand" onClick={() => setLightbox(ins.evidencia)} title="Ampliar imagem"><Maximize2 size={13} /></button>
                </div>
              ) : <div className="insp-val-evidence-empty"><ImageIcon size={20} /></div>}
              <div className="insp-val-meta">
                <div className="name">{insp?.nome || "Inspetor removido"}</div>
                <div className="sub">{emp?.nome} · {insp?.funcao}</div>
                <div className="insp-val-tags">
                  <span className="insp-tag">Tipo: {tipoLabel(ins.tipo)}</span>
                  <span className="insp-tag">Regional: {ins.regional}</span>
                  <span className="insp-tag">Data: {fmtDate(ins.dataInspecao)}</span>
                  <span className="insp-tag">Registrado às {ins.horaRegistro}</span>
                </div>
              </div>
            </div>
            {!aberta ? (
              <div className="insp-val-actions">
                <button className="insp-btn good" onClick={() => aprovar(ins)}><Check size={15} /> Aprovar</button>
                <button className="insp-btn bad" onClick={() => abrirReprovar(ins.id)}><X size={15} /> Reprovar</button>
              </div>
            ) : (
              <div style={{ marginTop: 14 }}>
                <div className="insp-field" style={{ marginBottom: 10 }}>
                  <label>Justificativa da reprovação</label>
                  <textarea className="insp-textarea" value={textoJustificativa} onChange={(e) => setTextoJustificativa(e.target.value)} placeholder="Descreva o motivo da reprovação para a empresa" autoFocus />
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button className="insp-btn bad" disabled={!textoJustificativa.trim()} onClick={() => confirmarReprovar(ins.id)}>Confirmar reprovação</button>
                  <button className="insp-btn ghost" onClick={() => setJustificativaAberta(null)}>Cancelar</button>
                </div>
              </div>
            )}
          </div>
        );
      })}
      <ImageLightbox src={lightbox} onClose={() => setLightbox(null)} />
    </div>
  );
}

/* ---------------------------------------------------------------
   ADM — Base de inspeções
---------------------------------------------------------------- */

function BaseScreen({ inspecoes, inspetores, empresas, getCodigo, mesesDisponiveis }) {
  const [filtroEmpresa, setFiltroEmpresa] = useState("todas");
  const [filtroStatus, setFiltroStatus] = useState("todos");
  const [filtroMes, setFiltroMes] = useState("todos");

  const linhas = useMemo(() => {
    return inspecoes
      .filter((i) => filtroEmpresa === "todas" || i.empresaId === filtroEmpresa)
      .filter((i) => filtroStatus === "todos" || i.status === filtroStatus)
      .filter((i) => filtroMes === "todos" || getMes(i.dataInspecao) === filtroMes)
      .sort((a, b) => b.criadoEm - a.criadoEm)
      .map((ins) => ({ ins, insp: inspetores.find((i) => i.id === ins.inspetorId), emp: empresas.find((e) => e.id === ins.empresaId) }));
  }, [inspecoes, inspetores, empresas, filtroEmpresa, filtroStatus, filtroMes]);

  function exportarExcel() {
    const dados = linhas.map(({ ins, insp, emp }) => ({
      ID: getCodigo(ins.id), Empresa: emp?.nome || "", Inspetor: insp?.nome || "", Função: insp?.funcao || "",
      Tipo: tipoLabel(ins.tipo), Regional: ins.regional, Data: fmtDate(ins.dataInspecao), Hora: ins.horaRegistro,
      Status: ins.status === "aprovado" ? "Aprovado" : ins.status === "reprovado" ? "Reprovado" : ins.status === "abonado" ? "Abonado" : "Pendente",
      Justificativa: ins.status === "reprovado" ? ins.justificativa || "" : "—",
    }));
    const ws = XLSX.utils.json_to_sheet(dados);
    ws["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 22 }, { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inspeções");
    XLSX.writeFile(wb, `base_inspecoes_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div>
      <h1 className="insp-h insp-page-title">Base de inspeções</h1>
      <p className="insp-page-sub">Todas as inspeções registradas, com filtros e exportação para Excel.</p>
      <div className="insp-toolbar">
        <div className="insp-filters">
          <select className="insp-select" style={{ width: 180 }} value={filtroEmpresa} onChange={(e) => setFiltroEmpresa(e.target.value)}>
            <option value="todas">Todas as empresas</option>{empresas.map((e) => <option key={e.id} value={e.id}>{e.nome}</option>)}
          </select>
          <select className="insp-select" style={{ width: 150 }} value={filtroStatus} onChange={(e) => setFiltroStatus(e.target.value)}>
            <option value="todos">Todos os status</option><option value="pendente">Pendente</option><option value="aprovado">Aprovado</option><option value="abonado">Abonado</option><option value="reprovado">Reprovado</option>
          </select>
          <select className="insp-select" style={{ width: 150 }} value={filtroMes} onChange={(e) => setFiltroMes(e.target.value)}>
            <option value="todos">Todos os meses</option>{mesesDisponiveis.map((m) => <option key={m} value={m}>{mesLabel(m)}</option>)}
          </select>
        </div>
        <button className="insp-btn secondary" onClick={exportarExcel}><Download size={15} /> Exportar Excel</button>
      </div>
      <div className="insp-card">
        <table className="insp-table">
          <thead><tr><th>ID</th><th>Inspetor</th><th>Empresa</th><th>Função</th><th>Tipo</th><th>Regional</th><th>Data</th><th>Hora</th><th>Status</th><th>Justificativa</th></tr></thead>
          <tbody>
            {linhas.length === 0 && <tr><td colSpan={10}><div className="insp-empty">Nenhum registro encontrado para esse filtro.</div></td></tr>}
            {linhas.map(({ ins, insp, emp }) => (
              <tr key={ins.id}>
                <td style={{ fontFamily: "monospace", fontSize: 12, color: "var(--ink-soft)" }}>{getCodigo(ins.id)}</td>
                <td style={{ fontWeight: 600 }}>{insp?.nome || "—"}</td><td>{emp?.nome || "—"}</td><td>{insp?.funcao || "—"}</td>
                <td>{tipoLabel(ins.tipo)}</td><td>{ins.regional}</td><td>{fmtDate(ins.dataInspecao)}</td><td>{ins.horaRegistro}</td>
                <td><StatusBadge status={ins.status} /></td>
                <td style={{ color: "var(--ink-soft)" }}>{ins.status === "reprovado" ? ins.justificativa : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Pesquisar
---------------------------------------------------------------- */

function PesquisaScreen({ inspecoes, inspetores, empresas, restrictedEmpresaId, getCodigo }) {
  const [query, setQuery] = useState("");
  const [selecionadaId, setSelecionadaId] = useState(null);
  const [lightbox, setLightbox] = useState(null);
  const base = restrictedEmpresaId ? inspecoes.filter((i) => i.empresaId === restrictedEmpresaId) : inspecoes;

  const resultados = useMemo(() => {
    const q = query.trim().toLowerCase();
    let list = base;
    if (q) {
      list = list.filter((ins) => {
        const insp = inspetores.find((i) => i.id === ins.inspetorId);
        const emp = empresas.find((e) => e.id === ins.empresaId);
        const codigo = getCodigo(ins.id).toLowerCase();
        return codigo.includes(q) || (insp?.nome || "").toLowerCase().includes(q) || (emp?.nome || "").toLowerCase().includes(q) || ins.regional.toLowerCase().includes(q);
      });
    }
    return [...list].sort((a, b) => b.criadoEm - a.criadoEm).slice(0, 40);
  }, [query, base, inspetores, empresas, getCodigo]);

  const insSel = selecionadaId ? inspecoes.find((i) => i.id === selecionadaId) : null;
  const inspSel = insSel ? inspetores.find((i) => i.id === insSel.inspetorId) : null;
  const empSel = insSel ? empresas.find((e) => e.id === insSel.empresaId) : null;

  return (
    <div>
      <h1 className="insp-h insp-page-title">Pesquisar inspeção</h1>
      <p className="insp-page-sub">Busque por ID, inspetor{!restrictedEmpresaId && ", empresa"} ou regional, e veja a evidência e os dados completos.</p>
      <div className="insp-search-input-wrap">
        <Search size={15} />
        <input className="insp-input" placeholder="Ex: INSP-0002, nome do inspetor..." value={query} onChange={(e) => setQuery(e.target.value)} />
      </div>
      <div style={{ display: "flex", gap: 20, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div className="insp-card" style={{ flex: "1 1 340px", minWidth: 300 }}>
          <table className="insp-table">
            <thead><tr><th>ID</th><th>Inspetor</th>{!restrictedEmpresaId && <th>Empresa</th>}<th>Data</th><th>Status</th></tr></thead>
            <tbody>
              {resultados.length === 0 && <tr><td colSpan={restrictedEmpresaId ? 4 : 5}><div className="insp-empty">Nenhum resultado encontrado.</div></td></tr>}
              {resultados.map((ins) => {
                const insp = inspetores.find((i) => i.id === ins.inspetorId);
                const emp = empresas.find((e) => e.id === ins.empresaId);
                const ativa = selecionadaId === ins.id;
                return (
                  <tr key={ins.id} onClick={() => setSelecionadaId(ins.id)} style={{ cursor: "pointer", background: ativa ? "var(--accent-soft)" : undefined }}>
                    <td style={{ fontFamily: "monospace", fontSize: 12 }}>{getCodigo(ins.id)}</td>
                    <td style={{ fontWeight: 600 }}>{insp?.nome || "—"}</td>
                    {!restrictedEmpresaId && <td>{emp?.nome || "—"}</td>}
                    <td>{fmtDate(ins.dataInspecao)}</td><td><StatusBadge status={ins.status} /></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div className="insp-card" style={{ flex: "1 1 280px", minWidth: 260, padding: 20 }}>
          {!insSel ? <div className="insp-empty"><Search size={24} />Selecione um resultado para ver os detalhes.</div> : (
            <div>
              {insSel.evidencia ? (
                <img src={insSel.evidencia} alt="evidência" style={{ width: "100%", borderRadius: 9, marginBottom: 14, cursor: "zoom-in", border: "1px solid var(--line)", display: "block" }} onClick={() => setLightbox(insSel.evidencia)} />
              ) : <div className="insp-val-evidence-empty" style={{ width: "100%", height: 160, marginBottom: 14 }}><ImageIcon size={22} /></div>}
              <div style={{ fontSize: 12, fontFamily: "monospace", color: "var(--ink-soft)", marginBottom: 6 }}>{getCodigo(insSel.id)}</div>
              <div style={{ fontSize: 17, fontWeight: 600, marginBottom: 2 }}>{inspSel?.nome || "Inspetor removido"}</div>
              <div style={{ fontSize: 12.5, color: "var(--ink-soft)", marginBottom: 14 }}>{empSel?.nome} · {inspSel?.funcao}</div>
              <div className="insp-detail-grid">
                <div><div className="lbl">Tipo</div><div className="val">{tipoLabel(insSel.tipo)}</div></div>
                <div><div className="lbl">Regional</div><div className="val">{insSel.regional}</div></div>
                <div><div className="lbl">Data</div><div className="val">{fmtDate(insSel.dataInspecao)}</div></div>
                <div><div className="lbl">Hora</div><div className="val">{insSel.horaRegistro}</div></div>
                <div><div className="lbl">Status</div><StatusBadge status={insSel.status} /></div>
              </div>
              {insSel.status === "reprovado" && <div style={{ fontSize: 12.5 }}><div style={{ color: "var(--ink-soft)", marginBottom: 3 }}>Justificativa</div><div>{insSel.justificativa}</div></div>}
            </div>
          )}
        </div>
      </div>
      <ImageLightbox src={lightbox} onClose={() => setLightbox(null)} />
    </div>
  );
}

/* ---------------------------------------------------------------
   Painel da empresa
---------------------------------------------------------------- */

function PainelEmpresaScreen({ empresas, inspetores, inspecoes, session, mesesDisponiveis }) {
  const isAdm = session.tipo === "adm";
  const [empresaId, setEmpresaId] = useState(isAdm ? (empresas[0]?.id || "") : session.empresaId);
  const [regional, setRegional] = useState("todas");
  const [mes, setMes] = useState(mesesDisponiveis[0]);
  const [mostrarNomes, setMostrarNomes] = useState(true);

  const empresa = empresas.find((e) => e.id === empresaId);
  const stats = useMemo(() => computeEmpresaStats(empresaId, mes, inspetores, inspecoes, regional), [empresaId, mes, regional, inspetores, inspecoes]);

  const donutInspetores = [{ value: stats.comInspecao, color: "var(--rail)" }, { value: stats.semInspecao, color: "var(--bad)" }];
  const pctCom = stats.inspetores.length > 0 ? Math.round((stats.comInspecao / stats.inspetores.length) * 100) : 0;
  const donutStatus = [
    { value: stats.statusCounts.aprovado, color: "var(--good)", label: "Aprovado" },
    { value: stats.statusCounts.abonado, color: "var(--accent)", label: "Abono" },
    { value: stats.statusCounts.reprovado, color: "var(--bad)", label: "Reprovado" },
  ];
  const totalStatus = donutStatus.reduce((a, s) => a + s.value, 0);
  const geralFaixa = faixaCiclo(stats.eficienciaGeral);

  return (
    <div>
      <div className="insp-dash-head">
        <div>
          <div className="insp-h" style={{ fontSize: 22, fontWeight: 700 }}>{empresa?.nome || "Empresa"}</div>
          <div style={{ fontSize: 12.5, color: "#c4cad0" }}>{stats.inspetores.length} inspetor(es)</div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <select className="insp-select" style={{ width: 140 }} value={mes} onChange={(e) => setMes(e.target.value)}>
            {mesesDisponiveis.map((m) => <option key={m} value={m}>{mesLabel(m)}</option>)}
          </select>
          {isAdm && (
            <select className="insp-select" style={{ width: 190 }} value={empresaId} onChange={(e) => setEmpresaId(e.target.value)}>
              {empresas.map((e) => <option key={e.id} value={e.id}>{e.nome}</option>)}
            </select>
          )}
          <div className="insp-tabbar">
            {["todas", "Centro", "Oeste"].map((r) => (
              <button key={r} className={regional === r ? "active" : ""} onClick={() => setRegional(r)}>{r === "todas" ? "TODAS" : r.toUpperCase()}</button>
            ))}
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 16 }}>
        <div className="insp-card insp-stat-card">
          <div className="lbl">Eficiência geral</div>
          <div className="big" style={{ color: geralFaixa.color }}>{Math.round(stats.eficienciaGeral)}%</div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>dos 3 ciclos</div>
        </div>
        <div className="insp-card insp-stat-card">
          <div className="lbl">Total realizado</div>
          <div className="big">{stats.realizadoGeral} <span style={{ fontSize: 13, fontWeight: 600, color: "var(--ink-soft)" }}>+{stats.excedentes} excedentes</span></div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>Meta: {stats.metaGeral}</div>
        </div>
        <div className="insp-card" style={{ padding: 18, display: "flex", flexDirection: "column", alignItems: "center" }}>
          <Legend items={[{ label: "Total inspetores", color: "var(--rail)" }, { label: "Sem inspeções", color: "var(--bad)" }]} />
          <div style={{ marginTop: 8 }}>
            <Donut segments={donutInspetores} centerTop={<span style={{ fontSize: 20, fontWeight: 700 }}>{stats.inspetores.length}</span>} centerBottom={<span style={{ fontSize: 11, color: "var(--ink-soft)" }}>({pctCom}%)</span>} />
          </div>
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16, marginBottom: 16 }}>
        {stats.cicloTotais.map((c) => {
          const faixa = faixaCiclo(c.pct);
          return <CicloCard key={c.ciclo} titulo={`Ciclo ${c.ciclo} — ${c.ciclo === 1 ? "01 a 10" : c.ciclo === 2 ? "11 a 20" : "21 a 31"}`} faixaLabel={faixa.label} pct={c.pct} meta={c.metaTotal} realizado={c.realizadoTotal} valido={c.validoTotal} color={faixa.color} />;
        })}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 16, marginBottom: 10 }}>
        <div className="insp-card">
          <div className="insp-toolbar" style={{ padding: "14px 14px 0 14px", marginBottom: 0 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: "var(--ink-soft)" }}>Desempenho por inspetor</div>
            <button className="insp-btn secondary" onClick={() => setMostrarNomes((v) => !v)}>
              {mostrarNomes ? <EyeOff size={15} /> : <Eye size={15} />} {mostrarNomes ? "Ocultar nomes" : "Mostrar nomes"}
            </button>
          </div>
          <table className="insp-table">
            <thead><tr><th>Nome</th><th>Função</th><th>Regional</th><th>Ciclo 1 Feito/Meta</th><th>Ciclo 2 Feito/Meta</th><th>Ciclo 3 Feito/Meta</th><th>Eficiência</th></tr></thead>
            <tbody>
              {stats.inspetores.length === 0 && <tr><td colSpan={7}><div className="insp-empty">Nenhum inspetor para esse filtro.</div></td></tr>}
              {stats.inspetores.map((p) => (
                <tr key={p.inspetor.id}>
                  <td style={{ fontWeight: 600 }}>{mostrarNomes ? p.inspetor.nome : "•••••••••"}</td>
                  <td>{p.inspetor.funcao}</td><td>{p.inspetor.regional}</td>
                  {p.ciclos.map((c) => {
                    const ok = c.realizado >= c.metaEfetiva;
                    return <td key={c.ciclo}><span style={{ display: "inline-flex", alignItems: "center", gap: 5, fontWeight: 600, color: ok ? "var(--good)" : "var(--bad)" }}>{ok ? <Check size={13} /> : <X size={13} />} {c.realizado}/{c.metaEfetiva}</span></td>;
                  })}
                  <td style={{ minWidth: 130 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ flex: 1, height: 7, borderRadius: 100, background: "var(--line-soft)", overflow: "hidden" }}>
                        <div style={{ height: "100%", width: `${Math.min(100, p.eficiencia)}%`, background: faixaCiclo(p.eficiencia).color }} />
                      </div>
                      <span style={{ fontSize: 12.5, fontWeight: 700 }}>{Math.round(p.eficiencia)}%</span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="insp-card" style={{ padding: 18, display: "flex", flexDirection: "column", alignItems: "center" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 12, textAlign: "center" }}>Abono · Aprovado · Reprovado</div>
          <Donut segments={donutStatus} centerTop={<span style={{ fontSize: 18, fontWeight: 700 }}>{totalStatus}</span>} centerBottom={<span style={{ fontSize: 10.5, color: "var(--ink-soft)" }}>registros</span>} />
          <div style={{ marginTop: 14 }}>
            <Legend items={donutStatus.map((s) => ({ label: `${s.label} (${totalStatus > 0 ? Math.round((s.value / totalStatus) * 100) : 0}%)`, color: s.color }))} />
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   Painel geral
---------------------------------------------------------------- */

function PainelGeralScreen({ empresas, inspetores, inspecoes, mesesDisponiveis }) {
  const [mes, setMes] = useState(mesesDisponiveis[0]);
  const dados = useMemo(() => empresas.map((emp) => ({ empresa: emp, stats: computeEmpresaStats(emp.id, mes, inspetores, inspecoes, "todas") })), [empresas, mes, inspetores, inspecoes]);

  const eficientes = dados.filter((d) => d.stats.eficienciaGeral >= 80).length;
  const criticas = dados.filter((d) => d.stats.eficienciaGeral <= 50).length;
  const metaTotalGeral = dados.reduce((a, d) => a + d.stats.metaGeral, 0);
  const validoTotalGeral = dados.reduce((a, d) => a + d.stats.validoGeral, 0);
  const performanceGeral = metaTotalGeral > 0 ? (validoTotalGeral / metaTotalGeral) * 100 : 100;
  const realizadoPorCiclo = [1, 2, 3].map((c) => {
    const metaC = dados.reduce((a, d) => a + d.stats.cicloTotais[c - 1].metaTotal, 0);
    const validoC = dados.reduce((a, d) => a + d.stats.cicloTotais[c - 1].validoTotal, 0);
    return metaC > 0 ? (validoC / metaC) * 100 : 100;
  });
  const ranking = [...dados].sort((a, b) => b.stats.eficienciaGeral - a.stats.eficienciaGeral);
  const valoresDistintos = [...new Set(ranking.map((d) => Math.round(d.stats.eficienciaGeral)))].sort((a, b) => b - a);
  const getRank = (pct) => valoresDistintos.indexOf(Math.round(pct)) + 1;

  return (
    <div>
      <div className="insp-dash-head">
        <div className="insp-h" style={{ fontSize: 21, fontWeight: 700 }}>Painel geral</div>
        <select className="insp-select" style={{ width: 170 }} value={mes} onChange={(e) => setMes(e.target.value)}>
          {mesesDisponiveis.map((m) => <option key={m} value={m}>{mesLabel(m)}</option>)}
        </select>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1.3fr", gap: 14, marginBottom: 20 }}>
        <div className="insp-card insp-stat-card" style={{ padding: 16 }}>
          <div className="lbl" style={{ fontSize: 10.5 }}>Performance geral</div>
          <div style={{ fontSize: 24, fontWeight: 700, color: faixaRanking(performanceGeral).color, marginTop: 4 }}>{Math.round(performanceGeral)}%</div>
        </div>
        <div className="insp-card insp-stat-card" style={{ padding: 16 }}>
          <div className="lbl" style={{ fontSize: 10.5 }}>Empresas eficientes (≥80%)</div>
          <div style={{ fontSize: 24, fontWeight: 700, marginTop: 4 }}>{eficientes}<span style={{ fontSize: 12, color: "var(--ink-soft)", fontWeight: 600 }}> de {dados.length}</span></div>
        </div>
        <div className="insp-card insp-stat-card" style={{ padding: 16 }}>
          <div className="lbl" style={{ fontSize: 10.5 }}>Empresas (≤50%)</div>
          <div style={{ fontSize: 24, fontWeight: 700, marginTop: 4, color: criticas > 0 ? "var(--bad)" : "var(--ink)" }}>{criticas > 0 ? criticas : "--"}</div>
        </div>
        <div className="insp-card insp-stat-card" style={{ padding: 16 }}>
          <div className="lbl" style={{ fontSize: 10.5 }}>Empresas monitoradas</div>
          <div style={{ fontSize: 24, fontWeight: 700, marginTop: 4 }}>{dados.length}</div>
        </div>
        <div className="insp-card" style={{ padding: 16 }}>
          <div className="lbl" style={{ fontSize: 10.5, marginBottom: 8 }}>Realizado por ciclo</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5 }}>
            {realizadoPorCiclo.map((p, idx) => (
              <div key={idx} style={{ textAlign: "center" }}>
                <div style={{ color: "var(--ink-soft)", fontSize: 10.5 }}>{idx === 0 ? "1 a 10" : idx === 1 ? "11 a 20" : "21 a 31"}</div>
                <div style={{ fontWeight: 700, color: faixaRanking(p).color }}>{Math.round(p)}%</div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1.15fr", gap: 16, alignItems: "start" }}>
        {[1, 2, 3].map((c) => (
          <div className="insp-card" key={c} style={{ padding: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Performance por empresa — Ciclo {c}</div>
            {[...dados].sort((a, b) => b.stats.cicloTotais[c - 1].pct - a.stats.cicloTotais[c - 1].pct).map((d) => <BarRow key={d.empresa.id} nome={d.empresa.nome} pct={d.stats.cicloTotais[c - 1].pct} />)}
            <div style={{ marginTop: 10 }}>
              <Legend items={[{ label: "≥80% no alvo", color: "var(--good)" }, { label: "50–79% atenção", color: "var(--warn)" }, { label: "<50% crítico", color: "var(--bad)" }]} />
            </div>
          </div>
        ))}
        <div className="insp-card" style={{ padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12 }}>Ranking geral</div>
          <table className="insp-table">
            <thead><tr><th></th><th>Empresa</th><th>Performance</th><th>Status</th></tr></thead>
            <tbody>
              {ranking.map((d) => {
                const faixa = faixaRanking(d.stats.eficienciaGeral);
                return (
                  <tr key={d.empresa.id}>
                    <td style={{ fontWeight: 700, color: "var(--ink-soft)" }}>{getRank(d.stats.eficienciaGeral)}</td>
                    <td style={{ fontWeight: 600 }}>{d.empresa.nome}</td>
                    <td style={{ fontWeight: 700, color: faixa.color }}>{Math.round(d.stats.eficienciaGeral)}%</td>
                    <td><span className="insp-badge" style={{ background: faixa.color + "22", color: faixa.color }}>{faixa.label}</span></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------
   App principal
---------------------------------------------------------------- */

export default function Home() {
  const [authUser, setAuthUser] = useState(undefined);
  const [session, setSession] = useState(null);
  const [screen, setScreen] = useState(null);
  const [empresas, setEmpresas] = useState([]);
  const [inspetores, setInspetores] = useState([]);
  const [inspecoes, setInspecoes] = useState([]);
  const [dataLoading, setDataLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setAuthUser(data.session?.user || null));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, sess) => setAuthUser(sess?.user || null));
    return () => sub.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (authUser === undefined) return;
    if (!authUser) { setSession(null); return; }
    (async () => {
      const { data: adminRow } = await supabase.from("admins").select("user_id").eq("user_id", authUser.id).maybeSingle();
      if (adminRow) { setSession({ tipo: "adm" }); setScreen("validacao"); return; }
      const { data: empresaRow } = await supabase.from("empresas").select("id").eq("user_id", authUser.id).maybeSingle();
      if (empresaRow) { setSession({ tipo: "empresa", empresaId: empresaRow.id }); setScreen("nova-inspecao"); return; }
      setSession("sem-acesso");
    })();
  }, [authUser]);

  async function reload() {
    setDataLoading(true);
    const [e, i, ins] = await Promise.all([
      supabase.from("empresas").select("*").order("nome"),
      supabase.from("inspetores").select("*"),
      supabase.from("inspecoes").select("*"),
    ]);
    setEmpresas((e.data || []).map(mapEmpresa));
    setInspetores((i.data || []).map(mapInspetor));
    setInspecoes((ins.data || []).map(mapInspecao));
    setDataLoading(false);
  }

  useEffect(() => { if (session && session !== "sem-acesso") reload(); }, [session]);

  const codigoMap = useMemo(() => {
    const ordenadas = [...inspecoes].sort((a, b) => a.criadoEm - b.criadoEm);
    const map = {};
    ordenadas.forEach((ins, idx) => { map[ins.id] = `INSP-${String(idx + 1).padStart(4, "0")}`; });
    return map;
  }, [inspecoes]);
  const getCodigo = (id) => codigoMap[id] || id;
  const mesesDisponiveis = useMemo(() => distinctMeses(inspecoes), [inspecoes]);

  if (authUser === undefined) return <div className="insp-loading-wrap"><GlobalStyle />Carregando…</div>;
  if (!authUser) return <LoginScreen />;
  if (session === "sem-acesso") return <SemAcessoScreen onSair={() => supabase.auth.signOut()} />;
  if (!session) return <div className="insp-loading-wrap"><GlobalStyle />Carregando…</div>;

  const isAdm = session.tipo === "adm";

  return (
    <div className="insp-root">
      <GlobalStyle />
      <Sidebar session={session} empresas={empresas} screen={screen} setScreen={setScreen} onLogout={() => supabase.auth.signOut()} />
      <main className="insp-main">
        {dataLoading && empresas.length === 0 ? (
          <div className="insp-empty">Carregando dados…</div>
        ) : (
          <>
            {screen === "empresas" && isAdm && <EmpresasScreen empresas={empresas} inspetores={inspetores} reload={reload} />}
            {screen === "inspetores" && <InspetoresScreen inspetores={inspetores} empresas={empresas} restrictedEmpresaId={isAdm ? null : session.empresaId} reload={reload} />}
            {screen === "validacao" && isAdm && <ValidacaoScreen inspecoes={inspecoes} inspetores={inspetores} empresas={empresas} reload={reload} />}
            {screen === "base" && isAdm && <BaseScreen inspecoes={inspecoes} inspetores={inspetores} empresas={empresas} getCodigo={getCodigo} mesesDisponiveis={mesesDisponiveis} />}
            {screen === "pesquisa" && <PesquisaScreen inspecoes={inspecoes} inspetores={inspetores} empresas={empresas} restrictedEmpresaId={isAdm ? null : session.empresaId} getCodigo={getCodigo} />}
            {screen === "painel-empresa" && <PainelEmpresaScreen empresas={empresas} inspetores={inspetores} inspecoes={inspecoes} session={session} mesesDisponiveis={mesesDisponiveis} />}
            {screen === "painel-geral" && <PainelGeralScreen empresas={empresas} inspetores={inspetores} inspecoes={inspecoes} mesesDisponiveis={mesesDisponiveis} />}
            {screen === "nova-inspecao" && !isAdm && <NovaInspecaoScreen empresaId={session.empresaId} empresas={empresas} inspetores={inspetores} inspecoes={inspecoes} reload={reload} />}
            {screen === "minhas-inspecoes" && !isAdm && <MinhasInspecoesScreen empresaId={session.empresaId} empresas={empresas} inspetores={inspetores} inspecoes={inspecoes} getCodigo={getCodigo} />}
          </>
        )}
        <div className="insp-footer">feito por matheus manske</div>
      </main>
    </div>
  );
}
