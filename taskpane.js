// ─── HELPERS — getElementById avec null-guard ────────────────────────────────
const safeEl = id => document.getElementById(id);
const setDisplay = (id, val) => { const el = safeEl(id); if (el) el.style.display = val; };

/**
 * FleetInsight AI — Taskpane Controller v2.2
 * ============================================
 * - Gpe électrogène intégrée (N1/N2/N3 + onglet dédié)
 * - Dates corrigées (affichage JJ/MM/AAAA depuis le backend)
 * - Décisions 100% basées sur données réelles
 * - Gemini AI : résumé exécutif, économies, VT juridique, générateur
 */
"use strict";

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const CFG = {
  API: "https://parcautoia.onrender.com",  // ← URL Render
  TIMEOUT: 90_000,
};

const S = {
  sheets: null,
  result: null,
  gemini: null,
  charts: {},
};

const PAL = [
  "#f59e0b","#3b82f6","#10b981","#8b5cf6","#ef4444",
  "#f97316","#06b6d4","#ec4899","#14b8a6","#a855f7",
];

const C_OPT = {
  responsive:true, maintainAspectRatio:false,
  plugins:{
    legend:{labels:{color:"#8fa3c0",font:{size:9}}},
    tooltip:{backgroundColor:"#131e35",borderColor:"rgba(245,158,11,.3)",borderWidth:1,
             titleColor:"#e8edf5",bodyColor:"#8fa3c0"},
  },
  scales:{
    x:{ticks:{color:"#4a6080",font:{size:9}},grid:{color:"rgba(255,255,255,.04)"}},
    y:{ticks:{color:"#4a6080",font:{size:9}},grid:{color:"rgba(255,255,255,.04)"}},
  },
};

// ─── UTILS ───────────────────────────────────────────────────────────────────

const fmtF = n => {
  if (n == null || isNaN(n)) return "—";
  n = +n;
  if (n >= 1_000_000) return `${(n/1e6).toFixed(1)}M FCFA`;
  if (n >= 1_000)     return `${(n/1e3).toFixed(0)}K FCFA`;
  return `${Math.round(n)} FCFA`;
};
const fmtKm = n => {
  if (n == null || isNaN(n)) return "—";
  return n >= 1000 ? `${(n/1000).toFixed(1)}K km` : `${Math.round(n)} km`;
};
const scoreColor = s =>
  s>=80?"#10b981":s>=60?"#3b82f6":s>=40?"#f59e0b":s>=20?"#f97316":"#ef4444";

const killChart = id => { if (S.charts[id]) { S.charts[id].destroy(); delete S.charts[id]; } };

// ─── STATUS / LOADER / TOAST ─────────────────────────────────────────────────

const setStatus = (lbl, state="ok") => {
  document.getElementById("connLabel").textContent = lbl;
  const d = document.getElementById("connDot");
  d.className = "conn-dot";
  if (state==="busy") d.classList.add("busy");
  if (state==="off")  d.classList.add("off");
};
const showLoader = (txt,sub) => {
  document.getElementById("loaderText").textContent = txt||"Traitement…";
  document.getElementById("loaderSub").textContent  = sub||"";
  document.getElementById("loader").style.display   = "flex";
  setStatus(txt,"busy");
};
const hideLoader = () => {
  document.getElementById("loader").style.display = "none";
  setStatus("Connecté");
};
const toast = (title,msg,type="info",ms=4500) => {
  const icons={success:"✅",error:"❌",warning:"⚠️",info:"ℹ️"};
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.innerHTML = `<span class="toast-icon">${icons[type]}</span>
    <div><div class="toast-title">${title}</div>${msg?`<div class="toast-msg">${msg}</div>`:""}</div>`;
  document.getElementById("toasts").appendChild(el);
  setTimeout(()=>el.remove(), ms);
};

// ─── API ─────────────────────────────────────────────────────────────────────

async function callAPI(endpoint, body) {
  const ctrl = new AbortController();
  const t    = setTimeout(()=>ctrl.abort(), CFG.TIMEOUT);
  try {
    const res = await fetch(`${CFG.API}${endpoint}`, {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify(body),
      signal:ctrl.signal,
    });
    clearTimeout(t);
    if (!res.ok) {
      const e = await res.json().catch(()=>({detail:res.statusText}));
      throw new Error(e.detail||`HTTP ${res.status}`);
    }
    return await res.json();
  } catch(e) {
    clearTimeout(t);
    if (e.name==="AbortError")
      throw new Error("Délai dépassé — backend Render peut être en veille (30s). Réessayez.");
    throw e;
  }
}

// ─── NAV ─────────────────────────────────────────────────────────────────────

function initNav() {
  document.querySelectorAll(".nav-btn").forEach(btn => {
    btn.addEventListener("click", ()=>{
      document.querySelectorAll(".nav-btn").forEach(b=>b.classList.remove("active"));
      document.querySelectorAll(".tab").forEach(t=>t.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
    });
  });
  document.querySelectorAll(".niveau-btn").forEach(btn => {
    btn.addEventListener("click", ()=>{
      document.querySelectorAll(".niveau-btn").forEach(b=>b.classList.remove("active"));
      document.querySelectorAll(".niveau-panel").forEach(p=>p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`panel-${btn.dataset.niveau}`).classList.add("active");
    });
  });
}

// ─── CHARGEMENT CLASSEUR ─────────────────────────────────────────────────────

async function loadWorkbook() {
  showLoader("Lecture du classeur…","Parcours de toutes les feuilles");
  try {
    await Excel.run(async ctx => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      const payload = {};
      for (const sh of sheets.items) {
        try {
          const rng = sh.getUsedRange();
          rng.load(["values"]);
          await ctx.sync();
          if (!rng.values || rng.values.length < 2) continue;
          // Convertit en dicts indexés (Office.js → backend)
          payload[sh.name] = rng.values.map(row =>
            Object.fromEntries(row.map((v,i)=>[String(i), v]))
          );
        } catch {}
      }
      if (!Object.keys(payload).length) throw new Error("Aucune feuille avec données.");
      S.sheets = payload;
      renderSheets(payload);
      document.getElementById("btnAnalyze").disabled = false;
      buildImmatList();  // Activer l'autocomplete véhicule
      toast("Classeur chargé",`${Object.keys(payload).length} feuille(s) prêtes`,"success");
    });
  } catch(err) {
    toast("Erreur de lecture",err.message,"error");
  } finally { hideLoader(); }
}

function renderSheets(payload) {
  const ICONS = {
    "LISTE DE VEH":"🚗","SUIVI_CARBURANT":"⛽","ENTRETIEN ET REPARATIONS":"🔧",
    "SORTIES VEH":"📍","VT":"📋","Gpe électrogène":"⚡",
  };
  document.getElementById("sheetGrid").innerHTML =
    Object.entries(payload).map(([name,rows])=>`
      <div class="sheet-chip">
        <span class="sc-icon">${ICONS[name]||"📄"}</span>
        <div><div class="sc-name">${name}</div><div class="sc-rows">${rows.length} lignes</div></div>
      </div>`).join("");
  document.getElementById("sheetCard").style.display="";
}

// ─── ANALYSE ─────────────────────────────────────────────────────────────────

async function runAnalysis() {
  if (!S.sheets) { toast("Chargez d'abord le classeur","","warning"); return; }
  const useGemini = document.getElementById("useGemini").checked;
  showLoader(
    useGemini ? "Analyse IA + Gemini…" : "Analyse IA…",
    "N1 · N2 · N3" + (useGemini ? " · Gemini AI" : "")
  );
  try {
    const data = await callAPI("/analyze", { sheets:S.sheets, use_gemini:useGemini });
    S.result = data;
    S.gemini = data.gemini || null;

    // Badge Gemini dans le header
    const badge = document.getElementById("geminiBadge");
    if (S.gemini && !S.gemini.error) {
      badge.style.display = "flex";
    }

    renderKPIs(data.kpis);
    renderAlertes(data.niveau_3_decisions?.alertes||[]);
    renderCharts(data);
    renderVehicules(data);
    renderBudgets(data);
    renderGenerateur(data);
    renderDecisions(data);
    renderGemini(S.gemini);
    renderReportPreview(data, S.gemini);

    document.getElementById("btnPdf").disabled  = false;
    document.getElementById("btnWord").disabled = false;

    const msg = S.gemini && !S.gemini.error
      ? "3 niveaux + Gemini AI générés ✓"
      : "3 niveaux générés (Gemini désactivé)";
    toast("Analyse terminée", msg, "success");
    document.querySelector('[data-tab="tableau"]').click();
  } catch(err) {
    toast("Erreur d'analyse", err.message, "error");
  } finally { hideLoader(); }
}

// ─── KPIs ────────────────────────────────────────────────────────────────────

function renderKPIs(kpis=[]) {
  const g = document.getElementById("kpiGrid");
  if (!kpis.length) return;
  g.innerHTML = kpis.map(k=>`
    <div class="kpi-card" style="--kpi-color:${k.color}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-value">${k.value}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-niv">Niveau ${k.niveau}</div>
    </div>`).join("");
}

// ─── ALERTES ─────────────────────────────────────────────────────────────────

function renderAlertes(alertes) {
  const card = document.getElementById("alertesCard");
  const list = document.getElementById("alertesList");
  if (!alertes.length) { card.style.display="none"; return; }
  card.style.display = "";

  const groups = {};
  alertes.forEach(a => {
    const d = a.domaine||"AUTRE";
    groups[d] = groups[d]||[];
    groups[d].push(a);
  });

  const DOMAIN_COLORS = {
    ENTRETIEN:"#8b5cf6", VT:"#f59e0b", GENERATEUR:"#06b6d4", AUTRE:"#8fa3c0"
  };
  const DOMAIN_ICONS = { ENTRETIEN:"🔧", VT:"📋", GENERATEUR:"⚡", AUTRE:"📌" };

  list.innerHTML = Object.entries(groups).map(([dom, items])=>`
    <div class="alerte-group">
      <div class="alerte-group-title" style="color:${DOMAIN_COLORS[dom]||'#8fa3c0'}">
        ${DOMAIN_ICONS[dom]||"📌"} ${dom}
      </div>
      ${items.map(a=>`
        <div class="alerte-item alerte-${a.niveau}">
          <span class="ai-icon">${a.icon}</span>
          <div style="flex:1">
            <div style="display:flex;align-items:center;gap:8px;margin-bottom:3px">
              <div class="ai-titre">${a.titre}</div>
              <span class="alerte-badge badge-${a.niveau}">${a.niveau}</span>
            </div>
            <div class="ai-detail">${a.detail}</div>
            <div class="ai-action">→ ${a.action}</div>
          </div>
        </div>`).join("")}
    </div>`).join("");
}

// ─── CHARTS TABLEAU DE BORD ───────────────────────────────────────────────────

function renderCharts(result) {
  const ops = result.niveau_1_operationnel;
  const ind = result.niveau_2_indicateurs;

  // 1. Carburant
  const topCarb = (ops.carburant?.top_consommateurs||[]).slice(0,8);
  if (topCarb.length) {
    killChart("chartCarb");
    S.charts.chartCarb = new Chart(document.getElementById("chartCarb"), {
      type:"bar",
      data:{
        labels: topCarb.map(v=>v.vehicule),
        datasets:[{label:"FCFA",data:topCarb.map(v=>v.total),
          backgroundColor:"#f59e0bcc",borderColor:"#f59e0b",borderWidth:1}],
      },
      options:{...C_OPT,indexAxis:"y",
        plugins:{...C_OPT.plugins,legend:{display:false}},
        scales:{
          x:{ticks:{color:"#4a6080",font:{size:9},callback:v=>fmtF(v)},grid:{color:"rgba(255,255,255,.04)"}},
          y:{ticks:{color:"#8fa3c0",font:{size:9}},grid:{display:false}},
        }},
    });
  }

  // 2. Interventions entretien
  const types = (ops.entretien?.types||[]).slice(0,6);
  if (types.length) {
    killChart("chartEntr");
    S.charts.chartEntr = new Chart(document.getElementById("chartEntr"), {
      type:"doughnut",
      data:{labels:types.map(t=>t.type),
            datasets:[{data:types.map(t=>t.count),backgroundColor:PAL,borderWidth:0}]},
      options:{...C_OPT,cutout:"60%",
        plugins:{legend:{position:"right",labels:{color:"#8fa3c0",font:{size:9},boxWidth:10}}},
        scales:{}},
    });
  }

  // 3. VT Statuts (Oui/Non/Pas encore)
  const vt = ops.vt||{};
  if (vt.total) {
    killChart("chartVT");
    S.charts.chartVT = new Chart(document.getElementById("chartVT"), {
      type:"doughnut",
      data:{
        labels:["✅ Oui (faite)","❌ Non","⏳ Pas encore"],
        datasets:[{data:[vt.oui||0,vt.non||0,vt.pas_encore||0],
          backgroundColor:["#10b981","#ef4444","#f59e0b"],borderWidth:0}],
      },
      options:{...C_OPT,cutout:"60%",
        plugins:{legend:{position:"right",labels:{color:"#8fa3c0",font:{size:10},boxWidth:12}}},
        scales:{}},
    });
  }

  // 4. Groupe électrogène par agence
  const genAg = (ops.generateur?.cout_par_agence||[]).slice(0,6);
  if (genAg.length) {
    killChart("chartGen");
    S.charts.chartGen = new Chart(document.getElementById("chartGen"), {
      type:"bar",
      data:{
        labels: genAg.map(a=>a.agence),
        datasets:[{label:"FCFA",data:genAg.map(a=>a.total),
          backgroundColor:"#06b6d4aa",borderColor:"#06b6d4",borderWidth:1,borderRadius:4}],
      },
      options:{...C_OPT,indexAxis:"y",
        plugins:{...C_OPT.plugins,legend:{display:false}},
        scales:{
          x:{ticks:{color:"#4a6080",font:{size:9},callback:v=>fmtF(v)},grid:{color:"rgba(255,255,255,.04)"}},
          y:{ticks:{color:"#8fa3c0",font:{size:9}},grid:{display:false}},
        }},
    });
  }

  // 5. Disponibilité parc
  const taux = ind.taux_immobilisation||{};
  if (taux.total) {
    killChart("chartDispo");
    S.charts.chartDispo = new Chart(document.getElementById("chartDispo"), {
      type:"bar",
      data:{
        labels:["Disponibles","En atelier"],
        datasets:[{data:[taux.disponibles||0,taux.en_atelier||0],
          backgroundColor:["#10b981","#f97316"],borderWidth:0,borderRadius:6}],
      },
      options:{...C_OPT,plugins:{...C_OPT.plugins,legend:{display:false}}},
    });
  }
}

// ─── VÉHICULES ────────────────────────────────────────────────────────────────

function renderVehicules(result) {
  const sante  = result.niveau_2_indicateurs?.sante_vehicules||[];
  const entr   = result.niveau_1_operationnel?.entretien||{};
  const vt_op  = result.niveau_1_operationnel?.vt||{};
  const dec_vt = result.niveau_3_decisions?.plan_action_vt||{};

  // Tableau santé
  const tbody = document.getElementById("vehTbody");
  tbody.innerHTML = sante.length
    ? sante.map(v=>{
        const sc = scoreColor(v.score);
        return `<tr class="veh-row" data-statut="${v.statut}">
          <td><span class="veh-immat">${v.vehicule}</span></td>
          <td><div class="score-bar-wrap">
            <div class="score-bar"><div class="score-fill" style="width:${v.score}%;background:${sc}"></div></div>
            <span style="font-family:var(--font-display);font-size:11px;color:${sc}">${v.score}</span>
          </div></td>
          <td><span class="badge-statut s-${v.statut}">${v.statut}</span></td>
          <td style="font-family:var(--font-display);font-size:11px;color:#f59e0b">${fmtF(v.total_cout)}</td>
          <td style="color:${v.nb_pannes>=3?'#ef4444':v.nb_pannes>=1?'#f97316':'#10b981'};font-weight:600">${v.nb_pannes}</td>
          <td>${v.immobilise?'<span style="color:#ef4444">🔴 Oui</span>':'<span style="color:#10b981">✅ Non</span>'}</td>
        </tr>`;
      }).join("")
    : `<tr><td colspan="6" class="table-empty">Aucun véhicule analysé.</td></tr>`;

  document.querySelectorAll(".filter-btn").forEach(btn=>{
    btn.addEventListener("click",()=>{
      document.querySelectorAll(".filter-btn").forEach(b=>b.classList.remove("active"));
      btn.classList.add("active");
      const f = btn.dataset.filter;
      document.querySelectorAll(".veh-row").forEach(row=>{
        row.style.display=(f==="all"||row.dataset.statut===f)?"":"none";
      });
    });
  });

  // VT par statuts
  const vtCard = document.getElementById("vtCard");
  vtCard.style.display = "";
  const RISK_C = {"CRITIQUE":"#ef4444","ÉLEVÉ":"#f97316","MODÉRÉ":"#f59e0b","FAIBLE":"#10b981"};
  const conf = result.niveau_2_indicateurs?.conformite_vt||{};
  const riskC = RISK_C[conf.niveau_risque]||"#8fa3c0";

  document.getElementById("vtGrid").innerHTML = `
    <div class="vt-conformite-banner" style="border-color:${riskC}">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
        <div>
          <div style="font-family:var(--font-display);font-size:17px;font-weight:700;color:${riskC}">${conf.taux_conformite||0}% conformité VT</div>
          <div style="font-size:10px;color:var(--text-3)">Risque : <strong style="color:${riskC}">${conf.niveau_risque||"—"}</strong></div>
        </div>
        <div style="font-size:22px">${(conf.taux_conformite||0)>=80?"✅":"⚠️"}</div>
      </div>
      <div style="height:10px;border-radius:99px;overflow:hidden;display:flex;margin-bottom:8px">
        <div style="flex:${vt_op.oui||0};background:#10b981" title="Oui:${vt_op.oui}"></div>
        <div style="flex:${vt_op.non||0};background:#ef4444" title="Non:${vt_op.non}"></div>
        <div style="flex:${vt_op.pas_encore||0};background:#f59e0b" title="Pas encore:${vt_op.pas_encore}"></div>
      </div>
      <div style="display:flex;gap:16px;font-size:11px">
        <span style="color:#10b981">✅ Oui : <b>${vt_op.oui||0}</b></span>
        <span style="color:#ef4444">❌ Non : <b>${vt_op.non||0}</b></span>
        <span style="color:#f59e0b">⏳ Pas encore : <b>${vt_op.pas_encore||0}</b></span>
      </div>
    </div>

    ${(vt_op.liste_non||[]).length?`
    <div class="vt-section">
      <div class="vt-section-title" style="color:#ef4444">❌ P1 — NON : action immédiate (${vt_op.non} véhicule(s))</div>
      <div class="vt-section-desc">Ces véhicules ne doivent pas circuler sans VT valide.</div>
      ${(vt_op.liste_non||[]).map(v=>`
        <div class="vt-item">
          <div><div class="vt-immat">${v.vehicule}</div><div class="vt-info">${v.affectation}</div></div>
          <div style="text-align:right">
            <div class="vt-exp expired">NON FAITE ⛔</div>
            <div style="font-size:9px;color:var(--text-3)">Exp : ${v.expiration}</div>
          </div>
        </div>`).join("")}
    </div>`:""}

    ${(dec_vt.p2_planifier?.vehicules||[]).length?`
    <div class="vt-section">
      <div class="vt-section-title" style="color:#f59e0b">⏳ P2 — PAS ENCORE · &lt;60j (${dec_vt.p2_planifier.count})</div>
      <div class="vt-section-desc">Planifier les RDV sous 2 semaines.</div>
      ${dec_vt.p2_planifier.vehicules.map(v=>`
        <div class="vt-item urgent">
          <div><div class="vt-immat">${v.vehicule}</div><div class="vt-info">${v.affectation}</div></div>
          <div style="text-align:right">
            <div class="vt-exp urgent">⏳ ${v.jours_restants!=null?v.jours_restants+"j":"—"}</div>
            <div style="font-size:9px;color:var(--text-3)">${v.expiration}</div>
          </div>
        </div>`).join("")}
    </div>`:""}

    ${(dec_vt.p3_programmer?.count||0)?`
    <div class="vt-section">
      <div class="vt-section-title" style="color:var(--text-2)">📅 P3 — PAS ENCORE · >60j (${dec_vt.p3_programmer.count})</div>
      <div class="vt-section-desc">À programmer ce trimestre.</div>
      ${(dec_vt.p3_programmer.vehicules||[]).slice(0,6).map(v=>`
        <div class="vt-item" style="border-left-color:var(--text-3);opacity:.8">
          <div><div class="vt-immat">${v.vehicule}</div><div class="vt-info">${v.affectation}</div></div>
          <div style="text-align:right">
            <div style="font-size:10px;color:var(--text-2)">${v.jours_restants!=null?v.jours_restants+"j":"—"}</div>
            <div style="font-size:9px;color:var(--text-3)">${v.expiration}</div>
          </div>
        </div>`).join("")}
    </div>`:""}

    ${(vt_op.liste_oui||[]).length?`
    <div class="vt-section">
      <div class="vt-section-title" style="color:#10b981">✅ Conformes — OUI (${vt_op.oui})</div>
      <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:6px">
        ${(vt_op.liste_oui||[]).map(v=>`
          <div style="background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.2);border-radius:6px;padding:4px 8px;font-size:10px">
            <span style="color:#6ee7b7;font-family:var(--font-display)">${v.vehicule}</span>
            <span style="color:var(--text-3);font-size:9px;margin-left:4px">${v.expiration}</span>
          </div>`).join("")}
      </div>
    </div>`:""}`;

  // En atelier
  const atelier = entr.liste_atelier||[];
  if (atelier.length) {
    document.getElementById("atelierList").innerHTML = atelier.map(v=>`
      <div class="atelier-item">
        <div>
          <div class="atelier-veh">${v.vehicule}</div>
          <div class="atelier-type">${v.type}</div>
        </div>
        <div style="text-align:right">
          <div class="atelier-date">${v.date_depot}</div>
          ${v.duree_jours!=null?`<div style="font-size:10px;color:${v.duree_jours>14?'#ef4444':'#f59e0b'}">${v.duree_jours} j d'immo.</div>`:""}
        </div>
      </div>`).join("");
    document.getElementById("atelierCard").style.display = "";
  }
}

// ─── BUDGETS ─────────────────────────────────────────────────────────────────

function renderBudgets(result) {
  const ops   = result.niveau_1_operationnel;
  const dec   = result.niveau_3_decisions;
  const optim = dec.optimisation||{};
  const carb  = ops.carburant||{};
  const entr  = ops.entretien||{};
  const gen   = ops.generateur||{};
  const tc    = carb.total_depense||0;
  const te    = entr.total_depense||0;
  const tg    = gen.total_depense||0;
  const tt    = tc+te+tg;

  document.getElementById("budgetSummary").innerHTML = `
    <div class="bs-card">
      <div class="bs-val" style="color:#f59e0b">${fmtF(tc)}</div>
      <div class="bs-label">⛽ Carburant</div>
      <div class="bs-bar"><div class="bs-bar-fill" style="width:${tt?tc/tt*100:0}%;background:#f59e0b"></div></div>
      <div style="font-size:9px;color:var(--text-3)">${optim.part_carburant_pct||0}%</div>
    </div>
    <div class="bs-card">
      <div class="bs-val" style="color:#8b5cf6">${fmtF(te)}</div>
      <div class="bs-label">🔧 Entretien</div>
      <div class="bs-bar"><div class="bs-bar-fill" style="width:${tt?te/tt*100:0}%;background:#8b5cf6"></div></div>
      <div style="font-size:9px;color:var(--text-3)">${optim.part_entretien_pct||0}%</div>
    </div>
    <div class="bs-card">
      <div class="bs-val" style="color:#06b6d4">${fmtF(tg)||"—"}</div>
      <div class="bs-label">⚡ Groupe Élect.</div>
      <div class="bs-bar"><div class="bs-bar-fill" style="width:${tt?tg/tt*100:0}%;background:#06b6d4"></div></div>
      <div style="font-size:9px;color:var(--text-3)">${optim.part_generateur_pct||0}%</div>
    </div>
    <div class="bs-card">
      <div class="bs-val" style="color:#3b82f6">${optim.total_fmt||"—"}</div>
      <div class="bs-label">💰 Total parc</div>
      <div style="margin-top:8px;font-size:10px;color:#10b981">Éco. Gemini : ${optim.economie_fmt||"calculé…"}</div>
    </div>`;

  // Gemini savings
  const savings = result.gemini?.budget_et_economies;
  const scCard  = document.getElementById("savingsCard");
  if (savings && !savings.error && !savings.parse_error && savings.economie_fmt) {
    scCard.style.display = "";
    document.getElementById("savingsScore").textContent =
      `Score ${savings.score_gestion||"—"}/100 · ${savings.niveau_maturite||"—"}`;
    document.getElementById("savingsAmount").textContent = savings.economie_fmt||"—";
    document.getElementById("savingsSynthese").textContent = savings.synthese||"";
    const leviers = savings.leviers||[];
    document.getElementById("savingsLeviers").innerHTML = leviers.length
      ? leviers.map(l=>`
          <div class="levier-item">
            <span class="levier-prio prio-${l.priorite||'FAIBLE'}">${l.priorite||"—"}</span>
            <div class="levier-body">
              <div class="levier-title">${l.levier||"—"}</div>
              <div class="levier-eco">${fmtF(l.economie_potentielle)}</div>
              <div class="levier-detail">${l.detail||""}</div>
            </div>
          </div>`).join("")
      : "";
  } else { scCard.style.display="none"; }

  // Donut budget
  killChart("chartBudget");
  if (tt>0) {
    S.charts.chartBudget = new Chart(document.getElementById("chartBudget"), {
      type:"doughnut",
      data:{
        labels:["⛽ Carburant","🔧 Entretien","⚡ Groupe Élect."],
        datasets:[{data:[tc,te,tg],backgroundColor:["#f59e0b","#8b5cf6","#06b6d4"],borderWidth:0}],
      },
      options:{...C_OPT,cutout:"65%",
        plugins:{legend:{position:"right",labels:{color:"#8fa3c0",font:{size:10},boxWidth:12}}},
        scales:{}},
    });
  }

  // Bar coût/véhicule
  const couts = (result.niveau_2_indicateurs?.cout_par_vehicule||[]).slice(0,10);
  if (couts.length) {
    killChart("chartCoutVeh");
    S.charts.chartCoutVeh = new Chart(document.getElementById("chartCoutVeh"), {
      type:"bar",
      data:{
        labels:couts.map(v=>v.vehicule),
        datasets:[
          {label:"Carburant",data:couts.map(v=>v.carburant||0),backgroundColor:"#f59e0bcc",borderRadius:4},
          {label:"Entretien",data:couts.map(v=>v.entretien||0),backgroundColor:"#8b5cf6cc",borderRadius:4},
        ],
      },
      options:{...C_OPT,indexAxis:"y",
        scales:{
          x:{stacked:true,ticks:{color:"#4a6080",font:{size:9},callback:v=>fmtF(v)},grid:{color:"rgba(255,255,255,.04)"}},
          y:{stacked:true,ticks:{color:"#8fa3c0",font:{size:9}},grid:{display:false}},
        }},
    });
  }

  const top = carb.top_consommateurs||[];
  if (top.length) {
    document.getElementById("carbDetail").innerHTML = `
      <table class="carb-table">
        <thead><tr><th>#</th><th>Véhicule</th><th>Dépense</th></tr></thead>
        <tbody>${top.map((v,i)=>`
          <tr>
            <td style="color:var(--text-3)">${i+1}</td>
            <td style="font-family:var(--font-display);color:var(--amber)">${v.vehicule}</td>
            <td class="carb-val">${v.total_fmt}</td>
          </tr>`).join("")}
        </tbody>
      </table>`;
  }
}

// ─── GROUPE ÉLECTROGÈNE ───────────────────────────────────────────────────────

function renderGenerateur(result) {
  const gen    = result.niveau_1_operationnel?.generateur||{};
  const gen_ind= result.niveau_2_indicateurs?.indicateurs_generateur||{};
  const gen_dec= result.niveau_3_decisions?.decisions_generateur||{};
  const gen_ai = result.gemini?.groupe_electrogene||null;

  if (!gen.disponible) {
    document.getElementById("genN1Content").innerHTML =
      `<p class="card-desc" style="color:var(--orange)">⚠️ Feuille "Gpe électrogène" non détectée dans le classeur.</p>`;
    return;
  }

  // N1 — Opérationnel
  document.getElementById("genN1Content").innerHTML = `
    <div class="gen-stat-row">
      <div class="gen-stat"><span class="gen-val stat-cyan">${fmtF(gen.total_depense)}</span><span class="gen-lbl">Dépense totale</span></div>
      <div class="gen-stat"><span class="gen-val">${gen.nb_entrees||0}</span><span class="gen-lbl">Transactions</span></div>
      <div class="gen-stat"><span class="gen-val">${fmtF(gen.cout_moyen)}</span><span class="gen-lbl">Coût moyen</span></div>
    </div>
    <div style="font-size:11px;color:var(--text-2);margin-bottom:8px">
      📅 Période : <strong>${gen.date_min||"—"}</strong> → <strong>${gen.date_max||"—"}</strong>
      ${gen.nb_jours_couverts?` (${gen.nb_jours_couverts} jours)`:""}
    </div>
    ${(gen.repartition_produit||[]).length?`
      <div style="font-size:10px;color:var(--text-3);text-transform:uppercase;margin-bottom:4px">Répartition produit</div>
      ${gen.repartition_produit.map(p=>`
        <div class="mbar-row">
          <span class="mbar-label">${p.produit}</span>
          <div class="mbar-track"><div class="mbar-fill" style="width:${p.pct}%;background:var(--cyan)"></div></div>
          <span class="mbar-count">${p.pct}%</span>
        </div>`).join("")}`:""}
    ${(gen.par_mois||[]).length?`
      <div style="font-size:10px;color:var(--text-3);text-transform:uppercase;margin:8px 0 4px">Par mois</div>
      <table class="mini-tbl">
        <thead><tr><th>Mois</th><th>Total</th></tr></thead>
        <tbody>${gen.par_mois.map(m=>`
          <tr>
            <td style="color:var(--text-2)">${m.mois}</td>
            <td class="carb-val">${fmtF(m.total)}</td>
          </tr>`).join("")}
        </tbody>
      </table>`:""}`;

  // N2 — Indicateurs
  document.getElementById("genN2Card").style.display = "";
  document.getElementById("genN2Content").innerHTML = `
    <div class="gen-stat-row">
      <div class="gen-stat">
        <span class="gen-val" style="color:var(--cyan)">${gen_ind.nb_agences||0}</span>
        <span class="gen-lbl">Agences</span>
      </div>
      <div class="gen-stat">
        <span class="gen-val">${fmtF(gen_ind.cout_moyen_agence)}</span>
        <span class="gen-lbl">Coût moy./agence</span>
      </div>
      <div class="gen-stat">
        <span class="gen-val" style="color:${(gen_ind.frequence_par_semaine||0)>3?'#ef4444':'#10b981'}">
          ${gen_ind.frequence_par_semaine||0}
        </span>
        <span class="gen-lbl">Ravi./semaine</span>
      </div>
    </div>
    ${gen_ind.top_agence?`
      <div style="font-size:11px;color:var(--text-2)">
        🏆 Top agence : <strong style="color:var(--cyan)">${gen_ind.top_agence.agence}</strong>
        — ${gen_ind.top_agence.fmt}
      </div>`:""}`

  // Graphique par agence
  const ag_data = (gen.cout_par_agence||[]).slice(0,8);
  if (ag_data.length) {
    document.getElementById("genChartCard").style.display="";
    killChart("chartGenDetail");
    S.charts.chartGenDetail = new Chart(document.getElementById("chartGenDetail"), {
      type:"bar",
      data:{
        labels:ag_data.map(a=>a.agence),
        datasets:[{label:"FCFA",data:ag_data.map(a=>a.total),
          backgroundColor:"#06b6d4aa",borderColor:"#06b6d4",borderWidth:1,borderRadius:4}],
      },
      options:{...C_OPT,indexAxis:"y",
        plugins:{...C_OPT.plugins,legend:{display:false}},
        scales:{
          x:{ticks:{color:"#4a6080",font:{size:9},callback:v=>fmtF(v)},grid:{color:"rgba(255,255,255,.04)"}},
          y:{ticks:{color:"#8fa3c0",font:{size:9}},grid:{display:false}},
        }},
    });
  }

  // N3 — Décisions (Gemini AI)
  const geminiGenCard = document.getElementById("genGeminiCard");
  const geminiGenContent = document.getElementById("genGeminiContent");

  if (gen_ai && !gen_ai.error && !gen_ai.parse_error) {
    geminiGenCard.style.display="";
    const nivChip = gen_ai.niveau_consommation
      ? `<span class="gen-niveau-chip gen-${gen_ai.niveau_consommation}">${gen_ai.niveau_consommation}</span>`
      : "";
    geminiGenContent.innerHTML = `
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
        ${nivChip}
        <span style="font-size:11px;color:var(--text-2)">${gen_ai.diagnostic||""}</span>
      </div>
      ${gen_ai.alerte?`<div style="background:rgba(239,68,68,.1);border-left:3px solid #ef4444;border-radius:4px;padding:8px 12px;font-size:11px;color:#fca5a5;margin-bottom:8px">⚠️ ${gen_ai.alerte}</div>`:""}
      ${(gen_ai.recommandations||[]).map(r=>`
        <div class="reco-item"><span class="reco-arrow">→</span><span class="reco-text">${r}</span></div>`).join("")}
      ${gen_ai.optimisation_possible?`
        <div style="margin-top:10px;padding:8px;background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.2);border-radius:6px;font-size:11px;color:#6ee7b7">
          💡 ${gen_ai.optimisation_possible}
        </div>`:""}`;
  } else if (gen_dec.recommandations) {
    // Fallback sans Gemini : décisions locales
    geminiGenCard.style.display="";
    geminiGenContent.innerHTML = `
      ${gen_dec.recommandations.map(r=>`
        <div class="reco-item"><span class="reco-arrow">→</span><span class="reco-text">${r}</span></div>`).join("")}`;
  }
}

// ─── DÉCISIONS (3 NIVEAUX) ────────────────────────────────────────────────────

function renderDecisions(result) {
  const ops  = result.niveau_1_operationnel;
  const ind  = result.niveau_2_indicateurs;
  const dec  = result.niveau_3_decisions;
  const carb = ops.carburant||{};
  const entr = ops.entretien||{};
  const gen  = ops.generateur||{};
  const inv  = ops.inventaire||{};
  const vt   = ops.vt||{};
  const sort = ops.sorties||{};
  const conf = ind.conformite_vt||{};
  const taux = ind.taux_immobilisation||{};
  const RISK_C={"CRITIQUE":"#ef4444","ÉLEVÉ":"#f97316","MODÉRÉ":"#f59e0b","FAIBLE":"#10b981"};

  // ── N1 ────────────────────────────────────────────────────────────────────
  document.getElementById("n1Content").innerHTML = `
    <div class="niv-block">
      <div class="niv-block-title">🚗 Inventaire — ${inv.total||0} véhicules</div>
      ${(inv.par_type||[]).map(t=>`
        <div class="mbar-row">
          <span class="mbar-label">${t.label}</span>
          <div class="mbar-track"><div class="mbar-fill" style="width:${t.value/(inv.total||1)*100}%;background:var(--amber)"></div></div>
          <span class="mbar-count">${t.value}</span>
        </div>`).join("")}
    </div>

    <div class="niv-block" style="border-left:3px solid var(--amber)">
      <div class="niv-block-title">⛽ Carburant — faits bruts</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val stat-amber">${fmtF(carb.total_depense)}</span><span class="stat-lbl">Total</span></div>
        <div class="stat-chip"><span class="stat-val">${carb.nb_transactions||0}</span><span class="stat-lbl">Transactions</span></div>
        <div class="stat-chip"><span class="stat-val">${fmtF(carb.moyenne_par_transaction)}</span><span class="stat-lbl">Moy./transaction</span></div>
      </div>
    </div>

    <div class="niv-block" style="border-left:3px solid var(--purple)">
      <div class="niv-block-title">🔧 Entretien & Réparations — faits bruts</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val">${entr.total_interventions||0}</span><span class="stat-lbl">Interventions</span></div>
        <div class="stat-chip"><span class="stat-val stat-red">${entr.en_atelier||0}</span><span class="stat-lbl">En atelier ⚠️</span></div>
        <div class="stat-chip"><span class="stat-val stat-amber">${fmtF(entr.total_depense)}</span><span class="stat-lbl">Coût total</span></div>
      </div>
      ${(entr.categories||[]).map(c=>`
        <div class="mbar-row">
          <span class="mbar-label">${c.cat}</span>
          <div class="mbar-track"><div class="mbar-fill" style="width:${c.count/(entr.total_interventions||1)*100}%;background:var(--purple)"></div></div>
          <span class="mbar-count">${c.count}</span>
        </div>`).join("")}
    </div>

    <div class="niv-block" style="border-left:3px solid var(--amber)">
      <div class="niv-block-title">📋 VT — faits bruts (colonne Statuts)</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val">${vt.total||0}</span><span class="stat-lbl">Total VT</span></div>
        <div class="stat-chip"><span class="stat-val stat-green">${vt.oui||0}</span><span class="stat-lbl">✅ Oui</span></div>
        <div class="stat-chip"><span class="stat-val stat-red">${vt.non||0}</span><span class="stat-lbl">❌ Non</span></div>
        <div class="stat-chip"><span class="stat-val stat-orange">${vt.pas_encore||0}</span><span class="stat-lbl">⏳ Pas encore</span></div>
      </div>
    </div>

    ${gen.disponible?`
    <div class="niv-block" style="border-left:3px solid var(--cyan)">
      <div class="niv-block-title">⚡ Gpe électrogène — faits bruts</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="gen-val stat-cyan">${fmtF(gen.total_depense)}</span><span class="stat-lbl">Total dépenses</span></div>
        <div class="stat-chip"><span class="gen-val">${gen.nb_entrees||0}</span><span class="stat-lbl">Transactions</span></div>
        <div class="stat-chip"><span class="gen-val">${gen.date_min||"—"} → ${gen.date_max||"—"}</span><span class="stat-lbl">Période</span></div>
      </div>
    </div>`:""}

    <div class="niv-block">
      <div class="niv-block-title">📍 Sorties — faits bruts</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val">${sort.total_sorties||0}</span><span class="stat-lbl">Sorties</span></div>
        <div class="stat-chip"><span class="stat-val">${sort.vehicules_actifs||0}</span><span class="stat-lbl">Véhicules actifs</span></div>
        <div class="stat-chip"><span class="stat-val">${fmtKm(sort.km_total)}</span><span class="stat-lbl">KM total</span></div>
      </div>
    </div>`;

  // ── N2 ────────────────────────────────────────────────────────────────────
  const pannes = ind.frequence_pannes||[];
  const couts  = (ind.cout_par_vehicule||[]).slice(0,8);
  const gen_ind= ind.indicateurs_generateur||{};

  document.getElementById("n2Content").innerHTML = `
    <div class="niv-block" style="border-left:3px solid var(--purple)">
      <div class="niv-block-title">🔧 Disponibilité du parc (entretien mécanique)</div>
      <div class="gauge-wrap">
        <div class="gauge-bar"><div class="gauge-fill" style="width:${taux.taux_dispo||0}%;background:${(taux.taux_dispo||0)>=85?'#10b981':'#ef4444'}"></div></div>
        <div class="gauge-label">${taux.taux_dispo||0}% disponible — ${taux.disponibles||0}/${taux.total||0} véhicules · Taux immo. atelier : <strong>${taux.taux_immo||0}%</strong></div>
      </div>
    </div>

    <div class="niv-block" style="border-left:3px solid var(--amber)">
      <div class="niv-block-title">📋 Conformité VT (basée sur Statuts)</div>
      <div class="gauge-wrap">
        <div class="gauge-bar"><div class="gauge-fill" style="width:${conf.taux_conformite||0}%;background:${(conf.taux_conformite||0)>=80?'#10b981':'#ef4444'}"></div></div>
        <div class="gauge-label">${conf.taux_conformite||0}% conforme · Risque : <strong style="color:${RISK_C[conf.niveau_risque]||'#8fa3c0'}">${conf.niveau_risque||"—"}</strong></div>
      </div>
      <div style="display:flex;gap:12px;margin-top:6px;flex-wrap:wrap;font-size:10px">
        <span style="color:#10b981">✅ ${conf.nb_conformes||0} conformes</span>
        <span style="color:#ef4444">❌ ${conf.nb_non_conformes||0} non conformes</span>
        <span style="color:#f59e0b">⏳ ${conf.nb_en_attente||0} en attente</span>
      </div>
    </div>

    <div class="niv-block">
      <div class="niv-block-title">💰 Coût par véhicule (percentiles réels)</div>
      ${couts.map(v=>`
        <div class="mbar-row">
          <span class="mbar-label" title="${v.vehicule}">${v.vehicule}</span>
          <div class="mbar-track"><div class="mbar-fill" style="width:${v.total/(couts[0]?.total||1)*100}%;background:var(--amber)"></div></div>
          <span class="mbar-count">${fmtF(v.total)}</span>
        </div>`).join("")}
    </div>

    <div class="niv-block">
      <div class="niv-block-title">⚡ Pannes récurrentes (curatif uniquement)</div>
      ${pannes.length?pannes.map(p=>`
        <div class="mbar-row">
          <span class="mbar-label">${p.vehicule}</span>
          <div class="mbar-track"><div class="mbar-fill" style="width:${Math.min(100,p.nb_pannes*25)}%;background:${p.nb_pannes>=3?'#ef4444':'#f59e0b'}"></div></div>
          <span class="mbar-count" style="color:${p.nb_pannes>=3?'#ef4444':'#f59e0b'}">${p.nb_pannes} panne(s)</span>
        </div>`).join("")
      :"<p style='color:var(--text-3);font-size:11px'>Aucune panne récurrente détectée.</p>"}
    </div>

    ${gen.disponible?`
    <div class="niv-block" style="border-left:3px solid var(--cyan)">
      <div class="niv-block-title">⚡ Indicateurs Gpe électrogène</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val stat-cyan">${gen_ind.nb_agences||0}</span><span class="stat-lbl">Agences</span></div>
        <div class="stat-chip"><span class="stat-val">${fmtF(gen_ind.cout_moyen_agence)}</span><span class="stat-lbl">Coût moy./agence</span></div>
        <div class="stat-chip">
          <span class="stat-val" style="color:${(gen_ind.frequence_par_semaine||0)>3?'#ef4444':'#10b981'}">${gen_ind.frequence_par_semaine||0}</span>
          <span class="stat-lbl">Ravit./semaine</span>
        </div>
      </div>
    </div>`:""}`;

  // ── N3 ────────────────────────────────────────────────────────────────────
  const plan_vt  = dec.plan_action_vt||{};
  const plan_veh = dec.plan_renouvellement||{};
  const gen_dec  = dec.decisions_generateur||{};
  const optim    = dec.optimisation||{};
  const BADGES   = {
    "REMPLACER":["s-REMPLACER","⛔"],"CRITIQUE":["s-CRITIQUE","🔴"],
    "SURVEILLER":["s-SURVEILLER","🟡"],"BON":["s-BON","🟢"],"OPTIMAL":["s-OPTIMAL","✅"]
  };

  document.getElementById("n3Content").innerHTML = `
    <!-- Plan VT -->
    <div class="niv-block" style="border-left:3px solid var(--amber)">
      <div class="niv-block-title">📋 Plan d'action VT — 3 priorités</div>
      <div style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.2);border-radius:6px;padding:8px;font-size:11px;margin-bottom:10px">
        Conformité : <strong>${plan_vt.synthese?.taux_conformite||0}%</strong> ·
        Non conformes : <strong style="color:#ef4444">${plan_vt.synthese?.non_conformes||0}</strong> ·
        En attente : <strong style="color:#f59e0b">${plan_vt.synthese?.en_attente||0}</strong> ·
        Risque : <strong style="color:${RISK_C[plan_vt.synthese?.niveau_risque]||'#8fa3c0'}">${plan_vt.synthese?.niveau_risque||"—"}</strong>
      </div>

      ${(plan_vt.p1_immediat?.vehicules||[]).length?`
        <div style="margin-bottom:10px">
          <div style="font-size:11px;font-weight:700;color:#ef4444;margin-bottom:4px">🔴 P1 — Immédiat (${plan_vt.p1_immediat.count})</div>
          ${plan_vt.p1_immediat.vehicules.map(v=>`
            <div style="display:flex;justify-content:space-between;background:rgba(239,68,68,.08);border-left:3px solid #ef4444;border-radius:4px;padding:6px 10px;margin-bottom:4px">
              <div><span style="font-family:var(--font-display);color:#fca5a5">${v.vehicule}</span>
                   <span style="font-size:9px;color:var(--text-3);margin-left:8px">${v.affectation}</span></div>
              <span style="font-size:10px;color:var(--text-2)">Exp: ${v.expiration}</span>
            </div>`).join("")}
        </div>`:""}

      ${(plan_vt.p2_planifier?.vehicules||[]).length?`
        <div style="margin-bottom:10px">
          <div style="font-size:11px;font-weight:700;color:#f59e0b;margin-bottom:4px">🟡 P2 — Sous 2 semaines (${plan_vt.p2_planifier.count})</div>
          ${plan_vt.p2_planifier.vehicules.map(v=>`
            <div style="display:flex;justify-content:space-between;background:rgba(245,158,11,.08);border-left:3px solid #f59e0b;border-radius:4px;padding:6px 10px;margin-bottom:4px">
              <div><span style="font-family:var(--font-display);color:#fcd34d">${v.vehicule}</span>
                   <span style="font-size:9px;color:var(--text-3);margin-left:8px">${v.affectation}</span></div>
              <span style="font-size:10px;color:#f59e0b">${v.jours_restants!=null?v.jours_restants+"j":"—"} · ${v.expiration}</span>
            </div>`).join("")}
        </div>`:""}

      ${(plan_vt.p3_programmer?.count||0)?`
        <div>
          <div style="font-size:11px;font-weight:700;color:var(--text-2);margin-bottom:4px">🔵 P3 — Ce trimestre (${plan_vt.p3_programmer.count})</div>
          <div style="font-size:10px;color:var(--text-3)">Programmer dans le planning trimestriel.</div>
        </div>`:""}
    </div>

    <!-- Plan renouvellement -->
    <div class="niv-block" style="border-left:3px solid var(--purple)">
      <div class="niv-block-title">🔧 Plan de renouvellement mécanique</div>
      <div class="stat-row">
        <div class="stat-chip"><span class="stat-val stat-red">${plan_veh.nb_remplacement||0}</span><span class="stat-lbl">À remplacer</span></div>
        <div class="stat-chip"><span class="stat-val stat-orange">${plan_veh.nb_audit||0}</span><span class="stat-lbl">Audit requis</span></div>
        <div class="stat-chip"><span class="stat-val">${plan_veh.budget_fmt||"—"}</span><span class="stat-lbl">Budget estimé</span></div>
      </div>
      <div style="font-size:10px;color:var(--amber);margin:4px 0">⏱️ ${plan_veh.horizon||"—"}</div>
      ${(plan_veh.a_remplacer||[]).map(v=>`
        <div style="display:flex;justify-content:space-between;background:var(--navy-600);border-radius:6px;padding:7px 10px;margin-bottom:5px">
          <div>
            <div style="font-family:var(--font-display);font-size:12px;color:var(--amber)">${v.vehicule}</div>
            <div style="font-size:9px;color:var(--text-3)">Coût : ${fmtF(v.total_cout)} · Pannes : ${v.nb_pannes}</div>
          </div>
          <span class="badge-statut ${BADGES[v.statut]?.[0]||""}">${BADGES[v.statut]?.[1]||""} ${v.statut}</span>
        </div>`).join("")}
    </div>

    <!-- Gpe électrogène N3 -->
    ${gen.disponible?`
    <div class="niv-block" style="border-left:3px solid var(--cyan)">
      <div class="niv-block-title">⚡ Décisions Gpe électrogène</div>
      <div style="margin-bottom:6px">
        <span style="font-size:11px;color:var(--cyan)">Budget : <strong>${gen_dec.total_fmt||"—"}</strong></span>
        ${gen_dec.priorite?`<span class="levier-prio prio-${gen_dec.priorite==='HAUTE'?'HAUTE':'FAIBLE'}" style="margin-left:8px">${gen_dec.priorite}</span>`:""}
      </div>
      ${(gen_dec.recommandations||[]).map(r=>`
        <div class="reco-item"><span class="reco-arrow">→</span><span class="reco-text">${r}</span></div>`).join("")}
    </div>`:""}

    <!-- Budget global -->
    <div class="niv-block">
      <div class="niv-block-title">💡 Recommandations budgétaires</div>
      <div style="margin-bottom:8px;padding:8px;background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.2);border-radius:6px">
        <span style="font-size:11px;color:#6ee7b7">
          Budget total : <strong>${optim.total_fmt||"—"}</strong> ·
          Économie estimée : <strong>${optim.economie_fmt||"Calculé par Gemini AI"}</strong>
        </span>
      </div>
      ${(optim.anomalies_detectees||[]).map(a=>`
        <div style="background:rgba(249,115,22,.08);border-left:3px solid #f97316;border-radius:4px;padding:6px 10px;margin-bottom:4px;font-size:11px;color:var(--text-2)">
          ⚠️ ${a.detail||a.type}
        </div>`).join("")}
    </div>`;
}

// ─── GEMINI AI TAB ────────────────────────────────────────────────────────────

function renderGemini(gemini) {
  const hero = document.getElementById("geminiHero");

  if (!gemini || gemini.error) {
    hero.querySelector(".gh-sub").textContent =
      gemini?.error
        ? `Gemini indisponible : ${gemini.error}`
        : "Activez Gemini AI et relancez l'analyse.";
    return;
  }

  // Masquer le hero placeholder
  hero.style.display = "none";

  const exec = gemini.resume_executif||{};
  const sav  = gemini.budget_et_economies||{};
  const vt_g = gemini.conformite_vt||{};
  const gen_g= gemini.groupe_electrogene||{};
  const STAT_C = {EXCELLENT:"#10b981",BON:"#3b82f6",ACCEPTABLE:"#f59e0b",PRÉOCCUPANT:"#f97316",CRITIQUE:"#ef4444"};

  // ── Résumé exécutif (tableau de bord aussi) ──────────────────────────────
  if (exec && !exec.parse_error) {
    const execCard = document.getElementById("execCard");
    execCard.style.display = "";
    const statut = exec.statut_global||"ACCEPTABLE";
    const statEl = document.getElementById("execStatut");
    statEl.textContent = statut;
    statEl.className = `gec-statut ${statut}`;
    document.getElementById("execMessage").textContent   = exec.message_cle||"";
    document.getElementById("execNarrative").textContent = exec.resume_narratif||"";
    document.getElementById("execCols").innerHTML = `
      ${exec.points_forts?.length?`
        <div class="gec-col-block">
          <div class="gec-col-title" style="color:#10b981">✅ Points forts</div>
          <ul class="gec-col-list">${(exec.points_forts).map(p=>`<li>${p}</li>`).join("")}</ul>
        </div>`:""}
      ${exec.points_attention?.length?`
        <div class="gec-col-block">
          <div class="gec-col-title" style="color:#f59e0b">⚠️ Points d'attention</div>
          <ul class="gec-col-list">${exec.points_attention.map(p=>`<li>${p}</li>`).join("")}</ul>
        </div>`:""}`;
  }

  // ── Score & maturité ────────────────────────────────────────────────────
  if (sav && !sav.parse_error) {
    document.getElementById("geminiScoreCard").style.display = "";
    const scoreC = STAT_C[sav.niveau_maturite]||"#8fa3c0";
    document.getElementById("geminiScoreContent").innerHTML = `
      <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap">
        <div style="text-align:center">
          <div style="font-family:var(--font-display);font-size:40px;font-weight:700;color:${scoreC}">${sav.score_gestion||"—"}</div>
          <div style="font-size:10px;color:var(--text-3)">/ 100</div>
        </div>
        <div>
          <div style="font-family:var(--font-display);font-size:14px;color:${scoreC};font-weight:700">${sav.niveau_maturite||"—"}</div>
          <div style="font-size:11px;color:var(--text-2);margin-top:4px;max-width:200px">${sav.synthese||""}</div>
        </div>
      </div>`;
  }

  // ── Analyse juridique VT ────────────────────────────────────────────────
  if (vt_g && !vt_g.parse_error && !vt_g.error) {
    document.getElementById("geminiVtCard").style.display = "";
    const fields = [
      {k:"risque_juridique",l:"Risque juridique"},
      {k:"sanctions_possibles",l:"Sanctions possibles"},
      {k:"recommandation_urgente",l:"Recommandation urgente"},
      {k:"plan_mise_en_conformite",l:"Plan de mise en conformité"},
      {k:"impact_assurance",l:"Impact assurance"},
    ];
    document.getElementById("geminiVtContent").innerHTML = `
      ${vt_g.message_direction?`
        <div style="background:rgba(245,158,11,.1);border:1px solid rgba(245,158,11,.3);border-radius:8px;padding:10px 14px;font-family:var(--font-display);font-size:13px;font-weight:700;color:var(--amber);margin-bottom:10px">
          📢 ${vt_g.message_direction}
        </div>`:""}
      ${fields.filter(f=>vt_g[f.k]).map(f=>`
        <div class="juridique-item">
          <div class="jur-label">${f.l}</div>
          <div class="jur-value">${vt_g[f.k]}</div>
        </div>`).join("")}`;
  }

  // ── Leviers optimisation ────────────────────────────────────────────────
  if (sav?.leviers?.length) {
    document.getElementById("geminiLevCard").style.display = "";
    document.getElementById("geminiLevContent").innerHTML = sav.leviers.map(l=>`
      <div class="levier-item">
        <span class="levier-prio prio-${l.priorite||'FAIBLE'}">${l.priorite||"—"}</span>
        <div class="levier-body">
          <div class="levier-title">${l.levier||"—"}</div>
          <div class="levier-eco">${fmtF(l.economie_potentielle)}</div>
          <div class="levier-detail">${l.detail||""}</div>
        </div>
      </div>`).join("");
  }

  // ── Diagnostic groupe électrogène ───────────────────────────────────────
  if (gen_g && !gen_g.error && !gen_g.parse_error && gen_g.diagnostic) {
    document.getElementById("geminiGenCard2").style.display = "";
    const niv = gen_g.niveau_consommation||"NORMAL";
    document.getElementById("geminiGenContent2").innerHTML = `
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">
        <span class="gen-niveau-chip gen-${niv}">${niv}</span>
        <span style="font-size:11px;color:var(--text-2)">${gen_g.diagnostic}</span>
      </div>
      ${gen_g.alerte?`<div style="background:rgba(239,68,68,.1);border-left:3px solid #ef4444;border-radius:4px;padding:8px 12px;font-size:11px;color:#fca5a5;margin-bottom:8px">⚠️ ${gen_g.alerte}</div>`:""}
      ${(gen_g.recommandations||[]).map(r=>`<div class="reco-item"><span class="reco-arrow">→</span><span class="reco-text">${r}</span></div>`).join("")}
      ${gen_g.optimisation_possible?`<div style="margin-top:10px;padding:8px;background:rgba(16,185,129,.08);border:1px solid rgba(16,185,129,.2);border-radius:6px;font-size:11px;color:#6ee7b7">💡 ${gen_g.optimisation_possible}</div>`:""}`;
  }
}

// ─── APERÇU RAPPORT ──────────────────────────────────────────────────────────

function renderReportPreview(result, gemini) {
  const ops   = result.niveau_1_operationnel;
  const dec   = result.niveau_3_decisions;
  const optim = dec.optimisation||{};
  const carb  = ops.carburant||{};
  const entr  = ops.entretien||{};
  const gen   = ops.generateur||{};
  const vt    = ops.vt||{};
  const conf  = result.niveau_2_indicateurs?.conformite_vt||{};
  const sav   = gemini?.budget_et_economies||{};

  document.getElementById("reportPreviewContent").innerHTML = `
    <div class="rp-section">
      <div class="rp-heading">🚗 Parc</div>
      <div class="rp-row"><span class="rp-key">Total véhicules</span><span class="rp-value">${ops.inventaire?.total||"—"}</span></div>
    </div>
    <div class="rp-section">
      <div class="rp-heading">📋 Conformité VT</div>
      <div class="rp-row"><span class="rp-key">✅ Oui (faite)</span><span class="rp-value" style="color:#10b981">${vt.oui||0}</span></div>
      <div class="rp-row"><span class="rp-key">❌ Non (non faite)</span><span class="rp-value" style="color:#ef4444">${vt.non||0}</span></div>
      <div class="rp-row"><span class="rp-key">⏳ Pas encore</span><span class="rp-value" style="color:#f59e0b">${vt.pas_encore||0}</span></div>
      <div class="rp-row"><span class="rp-key">Taux conformité</span><span class="rp-value">${conf.taux_conformite||0}%</span></div>
    </div>
    <div class="rp-section">
      <div class="rp-heading">💰 Budget complet</div>
      <div class="rp-row"><span class="rp-key">Carburant</span><span class="rp-value">${fmtF(carb.total_depense)}</span></div>
      <div class="rp-row"><span class="rp-key">Entretien mécanique</span><span class="rp-value">${fmtF(entr.total_depense)}</span></div>
      <div class="rp-row"><span class="rp-key">Groupe électrogène</span><span class="rp-value">${fmtF(gen.total_depense)||"—"}</span></div>
      <div class="rp-row"><span class="rp-key">Total parc</span><span class="rp-value" style="color:#3b82f6">${optim.total_fmt||"—"}</span></div>
    </div>
    ${sav.economie_fmt?`
    <div class="rp-section">
      <div class="rp-heading">🤖 Gemini AI</div>
      <div class="rp-row"><span class="rp-key">Économie estimée</span><span class="rp-value" style="color:#10b981">${sav.economie_fmt}</span></div>
      <div class="rp-row"><span class="rp-key">Score gestion</span><span class="rp-value">${sav.score_gestion||"—"}/100</span></div>
      <div class="rp-row"><span class="rp-key">Maturité</span><span class="rp-value">${sav.niveau_maturite||"—"}</span></div>
    </div>`:""}`;
  document.getElementById("reportPreview").style.display = "";
}

// ─── RAPPORT ─────────────────────────────────────────────────────────────────

async function downloadReport(format) {
  if (!S.result) { toast("Lancez l'analyse d'abord","","warning"); return; }
  showLoader("Génération du rapport…", format.toUpperCase());
  try {
    const res = await fetch(`${CFG.API}/report`, {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({
        result: S.result,
        gemini: S.gemini,
        format,
        title: document.getElementById("reportTitle").value,
      }),
    });
    if (!res.ok) throw new Error(await res.text());
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `fleetinsight_rapport.${format==="pdf"?"pdf":"docx"}`;
    a.click();
    toast("Rapport téléchargé", format.toUpperCase(), "success");
  } catch(err) {
    toast("Erreur rapport", err.message, "error");
  } finally { hideLoader(); }
}

// ─── HEALTH + INIT ────────────────────────────────────────────────────────────

async function checkHealth() {
  setStatus("Connexion…","busy");
  try {
    const res = await fetch(`${CFG.API}/health`, {signal:AbortSignal.timeout(8000)});
    if (res.ok) {
      const data = await res.json();
      setStatus(data.gemini ? "Connecté + Gemini" : "Connecté (sans Gemini)");
      if (!data.gemini) {
        document.getElementById("useGemini").checked = false;
        toast("Gemini non configuré","Ajoutez GEMINI_API_KEY sur Render","warning",6000);
      }
    } else { setStatus("Erreur API","off"); }
  } catch { setStatus("Hors ligne","off"); }
}

function bindEvents() {
  document.getElementById("btnLoad").addEventListener("click", loadWorkbook);
  document.getElementById("btnAnalyze").addEventListener("click", runAnalysis);
  document.getElementById("btnPdf").addEventListener("click",  ()=>downloadReport("pdf"));
  document.getElementById("btnWord").addEventListener("click", ()=>downloadReport("word"));

}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    initNav();
    bindEvents();
    bindVehiculeEvents();
    checkHealth();
    toast("FleetInsight AI","Chargez votre classeur .xlsm pour commencer.","info",3000);
  }
});


// ═══════════════════════════════════════════════════════════════
// ONGLET VÉHICULE — Recherche par immatriculation
// ═══════════════════════════════════════════════════════════════

// Liste des immatriculations (remplie après chargement)
let immatList = [];

// ─── Navigation sections ─────────────────────────────────────────────────────

function initVehiculeNav() {
  document.querySelectorAll(".vsn-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".vsn-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".veh-section").forEach(s => s.classList.remove("active"));
      btn.classList.add("active");
      const sec = safeEl("vsec-" + btn.dataset.vsection);
      if (sec) sec.classList.add("active");
    });
  });
}

// ─── Autocomplete immatriculations ────────────────────────────────────────────

function buildImmatList() {
  if (!S.sheets) return;
  // Extraire depuis la feuille LISTE DE VEH
  const sheetName = Object.keys(S.sheets).find(n =>
    ["LISTE DE VEH","LISTE VEH","VEHICULES","PARC"].includes(n.toUpperCase().trim())
  );
  if (!sheetName) return;
  const rows = S.sheets[sheetName];
  if (!rows || !rows.length) return;

  // Détecter la colonne IMMATRICULATION (index 5 dans la structure connue)
  const header = rows[0];
  const immatIdx = Object.entries(header).find(([k, v]) =>
    v && String(v).toUpperCase().includes("IMMAT")
  )?.[0];

  immatList = [];
  rows.slice(1).forEach(row => {
    const val = row[immatIdx || "5"];
    if (val && String(val).trim() && !String(val).includes("IMMAT")) {
      immatList.push(String(val).trim().toUpperCase());
    }
  });

  // Activer la recherche
  const btn = safeEl("btnSearchVeh");
  const hint = safeEl("vehHint");
  if (btn) btn.disabled = false;
  if (hint) hint.textContent = `${immatList.length} véhicule(s) disponibles — tapez pour rechercher.`;
}

function setupImmatAutocomplete() {
  const input = safeEl("immatInput");
  const sugg  = safeEl("immatSuggestions");
  if (!input || !sugg) return;

  input.addEventListener("input", () => {
    const q = input.value.trim().toUpperCase();
    if (!q || !immatList.length) { sugg.style.display = "none"; return; }

    const matches = immatList.filter(im =>
      im.includes(q) || q.split("").every(c => im.includes(c))
    ).slice(0, 8);

    if (!matches.length) { sugg.style.display = "none"; return; }

    sugg.innerHTML = matches.map(im =>
      `<div class="suggestion-item" data-immat="${im}">${im}</div>`
    ).join("");
    sugg.style.display = "";

    sugg.querySelectorAll(".suggestion-item").forEach(el => {
      el.addEventListener("click", () => {
        input.value = el.dataset.immat;
        sugg.style.display = "none";
        input.focus();
      });
    });
  });

  // Fermer les suggestions si clic dehors
  document.addEventListener("click", e => {
    if (!input.contains(e.target) && !sugg.contains(e.target)) {
      sugg.style.display = "none";
    }
  });

  // Recherche sur Entrée
  input.addEventListener("keyup", e => {
    if (e.key === "Enter") {
      sugg.style.display = "none";
      searchVehicle();
    }
  });
}

// ─── Recherche principale ─────────────────────────────────────────────────────

async function searchVehicle() {
  const input = safeEl("immatInput");
  const immat = input ? input.value.trim() : "";
  if (!immat) { toast("Saisissez une immatriculation", "", "warning"); return; }
  if (!S.sheets) { toast("Chargez le classeur d'abord", "", "warning"); return; }

  showLoader("Recherche en cours…", `Analyse de ${immat.toUpperCase()}`);

  try {
    const res = await fetch(`${CFG.API}/vehicle`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheets: S.sheets, immat }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: res.statusText }));
      throw new Error(err.detail || `HTTP ${res.status}`);
    }

    const data = await res.json();

    if (!data.found) {
      renderNotFound(data);
    } else {
      renderVehicleProfile(data);
    }

    setDisplay("vehProfile",  data.found ? "" : "none");
    setDisplay("vehNotFound", data.found ? "none" : "");

  } catch (err) {
    toast("Erreur de recherche", err.message, "error");
  } finally {
    hideLoader();
  }
}

// ─── Rendu profil véhicule ────────────────────────────────────────────────────

/**
 * FleetInsight IA — renderVehicleProfile (CORRIGÉ)
 * =================================================
 * Corrections :
 *   - Missions : nb chip utilise sorties.nb_missions (pas ms.length)
 *   - km_parcouru : affiche "0 km" quand la mission a d'autres données
 *   - Carburant : nb_transactions reflète le vrai count filtré du backend
 */

function renderVehicleProfile(data) {
  const { immat, fiche, carburant, entretien, vt, sorties, score } = data;

  // ── En-tête ──────────────────────────────────────────────────────────────
  const SCORE_COLORS = {
    OPTIMAL:    "#10b981",
    BON:        "#3b82f6",
    SURVEILLER: "#f59e0b",
    CRITIQUE:   "#f97316",
    REMPLACER:  "#ef4444",
  };
  const sc = score || {};
  const scoreColor = SCORE_COLORS[sc.statut] || "#8fa3c0";

  const hdr = safeEl("vehHeader");
  if (hdr) {
    hdr.innerHTML = `
      <div class="vph-left">
        <div class="vph-icon">🚗</div>
        <div>
          <div class="vph-immat">${immat}</div>
          <div class="vph-info">
            ${fiche ? `<b>${fiche.marque}</b> ${fiche.type}` : "Véhicule"}
          </div>
          ${fiche?.numero && fiche.numero !== "—"
            ? `<div style="font-size:9px;color:var(--text-3)">N° ${fiche.numero}</div>`
            : ""}
        </div>
      </div>
      <div class="vph-right">
        <span class="vph-badge"
              style="background:${scoreColor}22;color:${scoreColor};border:1px solid ${scoreColor}44">
          ${sc.statut || "—"}
        </span>
        <div class="vph-stats">
          <div class="vph-stat">
            <span class="vph-stat-val" style="color:#f59e0b">${carburant?.total_fmt || "—"}</span>
            <span class="vph-stat-lbl">Carburant</span>
          </div>
          <div class="vph-stat">
            <span class="vph-stat-val" style="color:#8b5cf6">${entretien?.total_fmt || "—"}</span>
            <span class="vph-stat-lbl">Entretien</span>
          </div>
          <div class="vph-stat">
            <span class="vph-stat-val">${sorties?.nb_missions || 0}</span>
            <span class="vph-stat-lbl">Missions</span>
          </div>
          ${entretien?.en_atelier
            ? `<div class="vph-stat" style="color:#ef4444">
                 <span class="vph-stat-val">⚠</span>
                 <span class="vph-stat-lbl">Atelier</span>
               </div>`
            : ""}
        </div>
      </div>`;
  }

  // ── Score santé ──────────────────────────────────────────────────────────
  const scoreEl = safeEl("vehScoreContent");
  if (scoreEl && sc) {
    const r = 32, cx = 40, cy = 40;
    const circ = 2 * Math.PI * r;
    const offset = circ - (sc.score / 100) * circ;
    scoreEl.innerHTML = `
      <div class="score-gauge-wrap">
        <div class="score-circle">
          <svg width="80" height="80" viewBox="0 0 40 40">
            <circle class="score-circle-bg" cx="${cx}" cy="${cy}" r="${r}"/>
            <circle class="score-circle-fill"
              cx="${cx}" cy="${cy}" r="${r}"
              stroke="${scoreColor}"
              stroke-dasharray="${circ}"
              stroke-dashoffset="${offset}"/>
          </svg>
          <div class="score-circle-text" style="color:${scoreColor}">${sc.score}</div>
        </div>
        <div class="score-details">
          <div class="score-statut" style="color:${scoreColor}">${sc.statut || "—"}</div>
          ${(sc.details || []).length ?
            sc.details.map(d => `<div class="score-issue">• ${d}</div>`).join("") :
            `<div class="score-issue" style="color:#10b981">✓ Aucune pénalité détectée</div>`}
        </div>
      </div>`;
  }

  // ── VT ────────────────────────────────────────────────────────────────────
  const vtEl = safeEl("vehVtContent");
  if (vtEl) {
    if (!vt) {
      vtEl.innerHTML = `<p style="color:var(--text-3);font-size:12px">Aucune donnée VT trouvée.</p>`;
    } else {
      const ALERTE_CFG = {
        EXPIRÉE:      { label: "EXPIRÉE",       color: "#ef4444" },
        NON_CONFORME: { label: "NON CONFORME",  color: "#ef4444" },
        URGENTE:      { label: "URGENT",        color: "#f97316" },
        VIGILANCE:    { label: "VIGILANCE",     color: "#f59e0b" },
        EN_ATTENTE:   { label: "EN ATTENTE",    color: "#8fa3c0" },
        OK:           { label: "CONFORME",      color: "#10b981" },
      };
      const cfg = ALERTE_CFG[vt.alerte] || ALERTE_CFG.EN_ATTENTE;
      const jr  = vt.jours_restants;
      vtEl.innerHTML = `
        <div class="vt-detail">
          <div class="vt-row">
            <span class="vt-lbl">Statut</span>
            <span class="vt-val">${vt.statut}</span>
          </div>
          <div class="vt-row">
            <span class="vt-lbl">Expiration</span>
            <span class="vt-val">${vt.expiration}</span>
          </div>
          <div class="vt-row">
            <span class="vt-lbl">Affectation</span>
            <span class="vt-val">${vt.affectation}</span>
          </div>
          <div class="vt-row">
            <span class="vt-lbl">Alerte</span>
            <span class="vt-badge" style="background:${cfg.color}22;color:${cfg.color}">
              ${cfg.label}
            </span>
          </div>
          ${jr !== null && jr !== undefined ? `
          <div class="vt-jours-badge"
               style="background:${cfg.color}22;border:1px solid ${cfg.color}44;color:${cfg.color}">
            ${jr < 0
              ? `Expirée il y a ${Math.abs(jr)} jour(s)`
              : jr === 0
                ? "Expire aujourd'hui"
                : `${jr} jour(s) restants`}
          </div>` : ""}
        </div>`;
    }
  }

  // ── Carburant ─────────────────────────────────────────────────────────────
  const carbEl = safeEl("vehCarbContent");
  if (carbEl) {
    const tx = carburant?.transactions || [];
    carbEl.innerHTML = `
      <div class="veh-stat-row">
        <div class="veh-stat-chip">
          <span class="val" style="color:#f59e0b">${carburant?.total_fmt || "—"}</span>
          <span class="lbl">Total dépenses</span>
        </div>
        <div class="veh-stat-chip">
          <span class="val">${carburant?.nb_transactions ?? 0}</span>
          <span class="lbl">Transactions</span>
        </div>
      </div>
      ${tx.length ? `
      <div class="hist-table-wrap">
        <table class="hist-table">
          <thead><tr>
            <th>Date</th><th>Produit</th><th>Carte</th><th>Montant</th>
          </tr></thead>
          <tbody>
            ${tx.map(t => `
              <tr>
                <td>${t.date}</td>
                <td>${t.produit}</td>
                <td style="font-size:10px;color:var(--text-3)">${t.carte}</td>
                <td class="montant-val">${t.montant_fmt}</td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>` :
      `<p style="color:var(--text-3);font-size:12px">Aucune transaction carburant trouvée.</p>`}`;
  }

  // ── Entretien ─────────────────────────────────────────────────────────────
  const entrEl = safeEl("vehEntrContent");
  if (entrEl) {
    const iv = entretien?.interventions || [];
    entrEl.innerHTML = `
      <div class="veh-stat-row">
        <div class="veh-stat-chip">
          <span class="val" style="color:#8b5cf6">${entretien?.total_fmt || "—"}</span>
          <span class="lbl">Total coûts</span>
        </div>
        <div class="veh-stat-chip">
          <span class="val">${entretien?.nb_interventions ?? 0}</span>
          <span class="lbl">Interventions</span>
        </div>
        ${entretien?.en_atelier ? `
        <div class="veh-stat-chip" style="border-color:#ef4444">
          <span class="val" style="color:#ef4444">
            ${entretien.duree_atelier != null ? entretien.duree_atelier + "j" : "En cours"}
          </span>
          <span class="lbl">En atelier</span>
        </div>` : ""}
      </div>
      ${iv.length ? `
      <div class="hist-table-wrap">
        <table class="hist-table">
          <thead><tr>
            <th>Dépôt</th><th>Retour</th><th>Type</th><th>Coût</th>
          </tr></thead>
          <tbody>
            ${iv.map(i => `
              <tr>
                <td>${i.date_depot}</td>
                <td>${i.en_cours
                  ? '<span class="badge-encours">En cours</span>'
                  : i.date_retour}</td>
                <td style="max-width:120px;overflow:hidden;text-overflow:ellipsis"
                    title="${i.type}">${i.type}</td>
                <td class="montant-val">${i.montant_fmt}</td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>` :
      `<p style="color:var(--text-3);font-size:12px">Aucune intervention trouvée.</p>`}`;
  }

  // ── Sorties / Missions ────────────────────────────────────────────────────
  const sortEl = safeEl("vehSortContent");
  if (sortEl) {
    const ms = sorties?.missions || [];

    // ── BUG FIX : on utilise sorties.nb_missions (total réel du backend)
    // et non ms.length (limité à 30 par le backend)
    const nbTotal = sorties?.nb_missions ?? ms.length;

    const fmtKm = n => {
      if (n == null || n === 0) return "—";
      return n >= 1000 ? `${(n / 1000).toFixed(1)}K km` : `${n} km`;
    };

    sortEl.innerHTML = `
      <div class="veh-stat-row">
        <div class="veh-stat-chip">
          <span class="val">${nbTotal}</span>
          <span class="lbl">Sorties</span>
        </div>
        <div class="veh-stat-chip">
          <span class="val">${sorties?.km_total ? fmtKm(sorties.km_total) : "—"}</span>
          <span class="lbl">KM parcourus</span>
        </div>
        ${sorties?.km_dernier ? `
        <div class="veh-stat-chip">
          <span class="val">
            ${Number(sorties.km_dernier).toLocaleString("fr-FR")}
          </span>
          <span class="lbl">Dernier KM</span>
        </div>` : ""}
      </div>
      ${ms.length ? `
      <div class="hist-table-wrap">
        <table class="hist-table">
          <thead><tr>
            <th>Date</th><th>KM parcouru</th><th>Destination</th><th>Conducteur</th>
          </tr></thead>
          <tbody>
            ${ms.map(m => `
              <tr>
                <td>${m.date}</td>
                <td style="font-family:var(--font-display);color:var(--cyan)">
                  ${m.km_parcouru > 0
                    ? fmtKm(m.km_parcouru)
                    : (m.date !== "—" || m.destination !== "—"
                        ? "<span style='opacity:.5'>—</span>"
                        : "—")}
                </td>
                <td>${m.destination}</td>
                <td style="font-size:10px;color:var(--text-3)">${m.conducteur}</td>
              </tr>`).join("")}
          </tbody>
        </table>
        ${nbTotal > 30
          ? `<p style="font-size:10px;color:var(--text-3);text-align:center;padding:8px 0">
               Affichage des 30 dernières sur ${nbTotal} sorties
             </p>`
          : ""}
      </div>` :
      `<p style="color:var(--text-3);font-size:12px">Aucune mission trouvée.</p>`}`;
  }
}

// ─── Bind vehicle events ──────────────────────────────────────────────────────

function bindVehiculeEvents() {
  const btn = safeEl("btnSearchVeh");
  if (btn) btn.addEventListener("click", () => searchVehicle());
  initVehiculeNav();
  setupImmatAutocomplete();
}
