#!/usr/bin/env node
// ══════════════════════════════════════════════════════════════════════
// REPORT AUTO — Quadro de Aulas — Resiliência Sistêmica
// Busca dados direto do Azure DevOps e gera PPTX + JPG
//
// CONFIGURAÇÃO:
//   1. npm install pptxgenjs node-fetch
//   2. Preencha as constantes abaixo (PAT, datas)
//   3. node report_auto.js
// ══════════════════════════════════════════════════════════════════════

const pptxgen  = require("pptxgenjs");
const fetch    = (...a) => import("node-fetch").then(({default:f}) => f(...a));
const path     = require("path");
const fs       = require("fs");
const { createCanvas } = require("canvas");

// ── CONFIGURAÇÃO ──────────────────────────────────────────────────────
const PAT         = process.env.ADO_PAT || "";
const ORG         = "alm-animaeducacao";
const PROJECT     = "TMD Financas Docente e Planejamento Academico";
const QUERY_ID    = "3e461cff-34fa-4988-a28e-da5146c3afaa";

// Datas fixas (atualizar conforme necessário)
const DATA_BASE_E1   = "15/05/2026";
const DATA_PESS_E1   = "29/05/2026";
const DATA_BASE_OBS  = "29/05/2026";
const DATA_PESS_OBS  = "12/06/2026";

// Camadas com story cancelada (setar cancelado:true quando necessário)
const CAMADAS_CANCELADAS = ["COMPLEMENTARES"]; // vazio = nenhuma cancelada

// Output
const OUTPUT_PPTX = "./Quadro_Aulas_Report_Auto.pptx";
const OUTPUT_JPG  = "C:/Users/cintia/OneDrive - Laureate Education - LATAMBR/Reports_Quadro_Aulas/Quadro_Aulas_Report_Auto.jpg";

// ── CONSTANTES ────────────────────────────────────────────────────────
const SUB_TYPES   = ["Sub Imp","Sub Test","Sub Requirement","Sub Value Activation","Sub Bug"];
const EXCLUIDOS   = ["ricardofernandescardoso","ricardo.cardoso","tiagosa","tiago@labsit"];
const FERIADOS    = ["2026-04-21","2026-05-01","2026-09-07","2026-10-12","2026-11-02","2026-11-15","2026-12-25"];

const TEAM = [
  { key:"gabrielpigatto",    nome:"Gabriel Pigatto",   ini:"GP", role:"DEV" },
  { key:"artursilva",        nome:"Artur Silva",       ini:"AS", role:"DEV" },
  { key:"carlos.h.almeida",  nome:"Carlos H. Almeida", ini:"CA", role:"DEV" },
  { key:"maria.e.muniz",     nome:"Maria E. Muniz",    ini:"MM", role:"DEV" },
  { key:"fabiosantos",       nome:"Fabio Santos",      ini:"FS", role:"DEV" },
  { key:"albert.foureaux",   nome:"Albert V. Foureaux",ini:"AV", role:"DEV" },
  { key:"james.pereira",     nome:"James G. Pereira",  ini:"JG", role:"QA"  },
];

// ── HELPERS ADO ────────────────────────────────────────────────────────
function adoHeaders() {
  const token = Buffer.from(`:${PAT}`).toString("base64");
  return { "Authorization": `Basic ${token}`, "Content-Type": "application/json" };
}

async function fetchQueryIds() {
  const url = `https://dev.azure.com/${ORG}/${encodeURIComponent(PROJECT)}/_apis/wit/wiql/${QUERY_ID}?api-version=7.1`;
  const res = await fetch(url, { headers: adoHeaders() });
  if (!res.ok) throw new Error(`Query failed: ${res.status} ${await res.text()}`);
  const data = await res.json();
  const rows = data.workItemRelations || data.workItems || [];
  const ids = [...new Set(rows.map(r => r.target?.id || r.id).filter(Boolean))];
  return ids;
}

async function fetchItemsBatch(ids) {
  const FIELDS = [
    "System.Id","System.WorkItemType","System.Title",
    "System.AssignedTo","System.State","System.Tags",
    "System.Parent","System.CreatedDate","System.ChangedDate",
    "Microsoft.VSTS.Scheduling.OriginalEstimate",
    "Microsoft.VSTS.Scheduling.CompletedWork",
    "Microsoft.VSTS.Scheduling.RemainingWork"
  ];
  const url = `https://dev.azure.com/${ORG}/_apis/wit/workitemsbatch?api-version=7.1`;
  const BATCH = 200;
  let all = [];
  for (let i = 0; i < ids.length; i += BATCH) {
    const chunk = ids.slice(i, i + BATCH);
    const res = await fetch(url, {
      method: "POST",
      headers: adoHeaders(),
      body: JSON.stringify({ ids: chunk, fields: FIELDS })
    });
    if (!res.ok) throw new Error(`Batch failed: ${res.status} ${await res.text()}`);
    const data = await res.json();
    all = all.concat(data.value || []);
  }
  return all;
}

// ── HELPERS ────────────────────────────────────────────────────────────
function addBdays(startDate, days) {
  let current = new Date(startDate);
  let added = 0;
  while (added < Math.max(1, Math.floor(days))) {
    current.setDate(current.getDate() + 1);
    const dow = current.getDay();
    const iso = current.toISOString().split("T")[0];
    if (dow > 0 && dow < 6 && !FERIADOS.includes(iso)) added++;
  }
  return `${String(current.getDate()).padStart(2,"0")}/${String(current.getMonth()+1).padStart(2,"0")}/${current.getFullYear()}`;
}

function getTag(tags) {
  if (!tags) return null;
  if (tags.includes("Camada Principal"))       return "PRINCIPAL";
  if (tags.includes("Camada Redundante"))      return "REDUNDANTE";
  if (tags.includes("Camadas Complementares")) return "COMPLEMENTARES";
  if (tags.includes("Observabilidade"))        return "OBSERVABILIDADE";
  if (tags.includes("SUS"))                    return "SUS";
  return null;
}

function isExcluido(assignedTo) {
  if (!assignedTo) return false;
  const email = (assignedTo.uniqueName || assignedTo || "").toLowerCase();
  return EXCLUIDOS.some(e => email.includes(e));
}

// ── CALCULAR MÉTRICAS ──────────────────────────────────────────────────
function calcular(items) {
  // Filtrar Epic e Cancelado
  const work = items.filter(i => {
    const wt = i.fields["System.WorkItemType"];
    const st = i.fields["System.State"];
    return wt !== "Epic" && st !== "Cancelado";
  });

  // Geral
  const total = work.length;
  const conc  = work.filter(i => i.fields["System.State"] === "Finalizado").length;
  const prog  = work.filter(i => ["Em Desenvolvimento","Pronto para Desenvolvimento"].includes(i.fields["System.State"])).length;
  const afbl  = work.filter(i => ["A Fazer","Em Refinamento","Product Backlog"].includes(i.fields["System.State"])).length;
  const pct   = Math.round(conc / total * 100);

  // Horas
  const plan  = work.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0), 0);
  const real  = work.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0), 0);
  const falta = work.filter(i => i.fields["System.State"] !== "Finalizado")
                    .reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.RemainingWork"] || 0), 0);
  const proj  = real + falta;
  const deltaH = proj - plan;
  const deltaP = plan > 0 ? Math.round(deltaH / plan * 100) : 0;

  // Sem tag
  const semTag = work.filter(i => {
    const t = i.fields["System.Tags"];
    return !t || t.trim() === "";
  });

  // Camadas
  const CAMADAS_DEF = [
    { nome:"PRINCIPAL",       cor:"1D9E75", tagFilter: t => t && t.includes("Camada Principal") },
    { nome:"REDUNDANTE",      cor:"534AB7", tagFilter: t => t && t.includes("Camada Redundante") },
    { nome:"COMPLEMENTARES",  cor:"D97706", tagFilter: t => t && t.includes("Camadas Complementares") },
    { nome:"OBSERVABILIDADE", cor:"1E6FA8", tagFilter: t => t && t.includes("Observabilidade") },
  ];

  const camadas = CAMADAS_DEF.map(def => {
    const c = work.filter(i => def.tagFilter(i.fields["System.Tags"]));
    const subs = c.filter(i => SUB_TYPES.includes(i.fields["System.WorkItemType"]));
    const pais = c.filter(i => ["Story","Non Functional Task"].includes(i.fields["System.WorkItemType"]));
    const finSubs = subs.filter(i => i.fields["System.State"] === "Finalizado").length;
    const finPais = pais.filter(i => i.fields["System.State"] === "Finalizado").length;
    const stories = pais.filter(i => i.fields["System.WorkItemType"] === "Story").length;
    const nft     = pais.filter(i => i.fields["System.WorkItemType"] === "Non Functional Task").length;
    const estC  = c.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0), 0);
    const realC = c.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0), 0);
    const remC  = c.filter(i => i.fields["System.State"] !== "Finalizado")
                   .reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.RemainingWork"] || 0), 0);
    const projC   = realC + remC;
    const deltaC  = projC - estC;
    const deltaPc = estC > 0 ? Math.round(deltaC / estC * 100) : 0;
    const pctSubs = subs.length > 0 ? Math.round(finSubs / subs.length * 100) : 0;
    const pctPais = pais.length > 0 ? Math.round(finPais / pais.length * 100) : 0;

    // Estouros ativos
    const ativos = subs.filter(i => ["Em Desenvolvimento","Pronto para Desenvolvimento","A Fazer"].includes(i.fields["System.State"]));
    const estouros = ativos.filter(i => {
      const cw = i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0;
      const oe = i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0;
      return oe > 0 && cw > oe;
    }).map(i => {
      const cw = i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0;
      const oe = i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0;
      const at = i.fields["System.AssignedTo"];
      const nome = (at?.displayName || at || "?").split("<")[0].trim();
      return { id: i.id, titulo: i.fields["System.Title"], dH: +(cw-oe).toFixed(1), dP: Math.round((cw/oe-1)*100), dev: nome };
    });

    return {
      nome: def.nome, cor: def.cor,
      cancelado: CAMADAS_CANCELADAS.includes(def.nome),
      stories, nft,
      itens: subs.length, conc: finSubs, prog: subs.filter(i => ["Em Desenvolvimento","Pronto para Desenvolvimento"].includes(i.fields["System.State"])).length,
      af: subs.filter(i => ["A Fazer","Em Refinamento","Product Backlog"].includes(i.fields["System.State"])).length,
      pais_total: pais.length, pais_conc: finPais, pais_pct: pctPais,
      est: +estC.toFixed(1), real: +realC.toFixed(1), rem: +remC.toFixed(1), proj: +projC.toFixed(1),
      pct: pctSubs,
      deltaH: `${deltaC >= 0 ? "+" : ""}${deltaC.toFixed(1)}h`,
      deltaP: `${deltaPc >= 0 ? "+" : ""}${deltaPc}%`,
      estouros_ativos: estouros,
    };
  });

  // Time
  const allItems = items.filter(i => i.fields["System.WorkItemType"] !== "Epic");
  const teamData = TEAM.map(t => {
    const s = allItems.filter(i => {
      const at = i.fields["System.AssignedTo"];
      const email = (at?.uniqueName || at || "").toLowerCase();
      return email.includes(t.key);
    });
    const r2  = +s.reduce((x,i) => x + (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0), 0).toFixed(1);
    const e2  = +s.reduce((x,i) => x + (i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0), 0).toFixed(1);
    const fin = s.filter(i => SUB_TYPES.includes(i.fields["System.WorkItemType"]) && i.fields["System.State"] === "Finalizado").length;
    const ratio = e2 > 0 ? r2 / e2 : 0;
    const cor = ratio > 1.3 ? "D85A30" : ratio > 1.0 ? "D97706" : "1D9E75";
    return { ...t, est: e2, real: r2, subs: fin, cor };
  });

  // Bugs / SUS (R3)
  const allWork = items;
  const bugs = allWork.filter(i => i.fields["System.WorkItemType"] === "Bug" && i.fields["System.State"] !== "Cancelado");
  const srs  = allWork.filter(i => i.fields["System.WorkItemType"] === "Service Request" && i.fields["System.State"] !== "Cancelado");
  const paiIds = new Set([...bugs, ...srs].map(i => i.id));
  const subsBugsR3 = allWork.filter(i => SUB_TYPES.includes(i.fields["System.WorkItemType"]) && paiIds.has(i.fields["System.Parent"]));
  const horasR3 = +subsBugsR3.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0), 0).toFixed(1);

  // Expansão escopo (R4)
  const cutoff = new Date("2026-05-01");
  const novos = work.filter(i => new Date(i.fields["System.CreatedDate"]) >= cutoff);
  const novosPorCamada = {};
  ["PRINCIPAL","REDUNDANTE","COMPLEMENTARES","OBSERVABILIDADE"].forEach(nome => {
    const def = CAMADAS_DEF.find(d => d.nome === nome);
    const nc = novos.filter(i => def.tagFilter(i.fields["System.Tags"]));
    const hrs = +nc.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0), 0).toFixed(0);
    if (hrs > 0) novosPorCamada[nome] = { count: nc.length, hrs };
  });
  const novosTxt = Object.entries(novosPorCamada).map(([n,v]) => {
    const abrev = {PRINCIPAL:"Princ",REDUNDANTE:"Redund",COMPLEMENTARES:"Comp",OBSERVABILIDADE:"Obs"}[n];
    return `${abrev} +${v.hrs}h`;
  }).join(" · ");
  const novosTotal = novos.length;

  // Refinamento pendente (stories sem sub)
  const storiesNFT = work.filter(i => ["Story","Non Functional Task"].includes(i.fields["System.WorkItemType"]));
  const paiComSub  = new Set(allWork.filter(i => SUB_TYPES.includes(i.fields["System.WorkItemType"]))
                              .map(i => i.fields["System.Parent"]).filter(Boolean));
  const semSub = storiesNFT.filter(i => !paiComSub.has(i.id));
  const semSubCamadas = [...new Set(semSub.map(i => {
    const t = i.fields["System.Tags"] || "";
    if (t.includes("Camada Principal")) return "Principal";
    if (t.includes("Camada Redundante")) return "Redundante";
    if (t.includes("Camadas Complementares")) return "Complementares";
    if (t.includes("Observabilidade")) return "Observabilidade";
    return "Sem tag";
  }))].join(" · ");

  // Eficiência
  const withEst = work.filter(i => (i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"] || 0) > 0 &&
                                   (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0) > 0);
  const estourados = withEst.filter(i => i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] > i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"]).length;
  const dentro     = withEst.filter(i => i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] <= i.fields["Microsoft.VSTS.Scheduling.OriginalEstimate"]).length;
  const totEf      = estourados + dentro;
  const subBugsAll = work.filter(i => i.fields["System.WorkItemType"] === "Sub Bug");
  const horasSubBugs = +subBugsAll.reduce((s,i) => s + (i.fields["Microsoft.VSTS.Scheduling.CompletedWork"] || 0), 0).toFixed(1);

  // Prazo
  const today = new Date();
  const obsIdx = camadas.findIndex(c => c.nome === "OBSERVABILIDADE");
  const remObs = camadas[obsIdx].rem;
  const remE1  = falta - remObs;
  const dataProjE1  = remE1 > 0 ? addBdays(today, remE1 / 18) : "Concluido";
  const dataProjObs = addBdays(today, remObs / 18);

  return {
    // Geral
    total, conc, prog, afbl, pct,
    plan: +plan.toFixed(1), real: +real.toFixed(1), falta: +falta.toFixed(1),
    proj: +proj.toFixed(1), deltaH: +deltaH.toFixed(1), deltaP,
    // Camadas
    camadas,
    // Time
    devs: teamData.filter(t => t.role === "DEV"),
    qa:   teamData.filter(t => t.role === "QA"),
    // Riscos
    semTag, semSub, semSubCamadas,
    bugs: bugs.length, srs: srs.length, horasR3,
    novosTotal, novosTxt,
    // Eficiência
    estourados, dentro, totEf, subBugsAll: subBugsAll.length, horasSubBugs,
    // Prazo
    dataProjE1, dataProjObs,
    remE1: +remE1.toFixed(1), remObs: +remObs.toFixed(1),
  };
}

// ── EXIBIR RESUMO ──────────────────────────────────────────────────────
function exibirResumo(M) {
  console.log("\n" + "═".repeat(60));
  console.log("  RESUMO — QUADRO DE AULAS — RESILIÊNCIA SISTÊMICA");
  console.log("═".repeat(60));

  console.log(`\n📊 GERAL: ${M.total} itens · ${M.conc} concluídos (${M.pct}%) · ${M.prog} prog · ${M.afbl} a fazer`);
  console.log(`💰 HORAS: Plan ${M.plan}h · Gasto ${M.real}h · Falta ${M.falta}h · Proj ${M.proj}h · +${M.deltaH}h (+${M.deltaP}%)`);

  console.log("\n🏗️  CAMADAS:");
  M.camadas.forEach(c => {
    const status = c.cancelado ? "CANCELADO" : (c.pct === 100 && c.pais_pct === 100 ? "✓ ENTREGUE" : `${c.pct}%`);
    console.log(`   [${c.nome}] subs ${c.conc}/${c.itens} (${c.pct}%) · pais ${c.pais_conc}/${c.pais_total} (${c.pais_pct}%) ${status}`);
    console.log(`     est=${c.est}h real=${c.real}h rem=${c.rem}h delta=${c.deltaH}(${c.deltaP})`);
    if (c.estouros_ativos.length > 0) {
      c.estouros_ativos.forEach(e => console.log(`     ⚠️  ESTOURO #${e.id} +${e.dH}h(+${e.dP}%) [${e.dev}]`));
    }
  });

  console.log("\n⏰ PRAZO (18h/dia):");
  console.log(`   E1 Resiliência (${M.remE1}h) → ${M.dataProjE1}  [Base: ${DATA_BASE_E1} · Pess: ${DATA_PESS_E1}]`);
  console.log(`   E2 Observabilidade (${M.remObs}h) → ${M.dataProjObs}  [Base: ${DATA_BASE_OBS} · Pess: ${DATA_PESS_OBS}]`);

  console.log("\n👥 TIME:");
  M.devs.forEach(d => {
    const alerta = d.real > d.est * 1.3 ? "🔴" : d.real > d.est ? "🟠" : "🟢";
    console.log(`   [DEV] ${alerta} ${d.nome}: ${d.real}h / ${d.est}h est · ${d.subs} subs`);
  });
  M.qa.forEach(d => {
    const alerta = d.real > d.est * 1.3 ? "🔴" : d.real > d.est ? "🟠" : "🟢";
    console.log(`   [QA]  ${alerta} ${d.nome}: ${d.real}h / ${d.est}h est · ${d.subs} subs`);
  });

  console.log("\n⚠️  RISCOS:");
  if (M.semTag.length > 0) {
    console.log(`   🏷️  SEM TAG: ${M.semTag.length} itens`);
    M.semTag.forEach(i => console.log(`      #${i.id} ${i.fields["System.WorkItemType"]} "${i.fields["System.Title"]?.substring(0,50)}"`));
  } else {
    console.log("   🏷️  Sem tag: 0 ✅");
  }
  if (M.semSub.length > 0) {
    console.log(`   📋 REFINAMENTO PENDENTE: ${M.semSub.length} stories (${M.semSubCamadas})`);
    M.semSub.forEach(i => console.log(`      #${i.id} "${i.fields["System.Title"]?.substring(0,50)}"`));
  }
  console.log(`   🐛 BUG/SUS: ${M.bugs} Bugs · ${M.srs} SRs · ${M.horasR3}h`);
  console.log(`   📈 EXPANSÃO DE ESCOPO: ${M.novosTotal} itens pós 01/05 (${M.novosTxt})`);

  console.log("\n📉 EFICIÊNCIA:");
  const pEstourados = M.totEf > 0 ? (M.estourados/M.totEf*100).toFixed(1) : "0.0";
  const pDentro     = M.totEf > 0 ? (M.dentro/M.totEf*100).toFixed(1) : "0.0";
  const pRetrab     = M.totEf > 0 ? (M.subBugsAll/M.totEf*100).toFixed(1) : "0.0";
  console.log(`   Estourados: ${pEstourados}% · Retrabalho: ${pRetrab}% (${M.subBugsAll} Sub Bugs · ${M.horasSubBugs}h) · Dentro: ${pDentro}%`);

  console.log("\n" + "═".repeat(60));
}

// ── GERAR PPTX ─────────────────────────────────────────────────────────
async function gerarPPTX(M) {
  // Montar objeto D compatível com o layout do report_v75
  const e1Cams   = M.camadas.filter(c => c.nome !== "OBSERVABILIDADE");
  const e1SubsFin = e1Cams.reduce((s,c) => s + Math.round(c.pct/100 * c.itens), 0);
  const e1SubsTot = e1Cams.reduce((s,c) => s + c.itens, 0);
  const e1Pct    = e1SubsTot > 0 ? Math.round(e1SubsFin / e1SubsTot * 100) : 0;
  const e2Cam    = M.camadas.find(c => c.nome === "OBSERVABILIDADE");
  const e2Pct    = e2Cam ? e2Cam.pct : 0;

  const pEstourados = M.totEf > 0 ? (M.estourados/M.totEf*100).toFixed(1) : "0.0";
  const pDentro     = M.totEf > 0 ? (M.dentro/M.totEf*100).toFixed(1) : "0.0";
  const pRetrab     = M.totEf > 0 ? (M.subBugsAll/M.totEf*100).toFixed(1) : "0.0";

  // Riscos dinâmicos
  const riscos = [];

  // R1 — Refinamento ou Estouro
  const totalEstouros = M.camadas.reduce((s,c) => s + c.estouros_ativos.length, 0);
  if (totalEstouros > 0) {
    const linhas = M.camadas.flatMap(c => c.estouros_ativos.map(e => `#${e.id} +${e.dH}h(+${e.dP}%) [${e.dev}]`));
    riscos.push({
      cor:"DC2626", bordaCor:"DC2626", bgCor:"FFF1F2",
      tag:"RISCO ATIVO",
      badge:`${totalEstouros}\nestouros`, badgeCor:"DC2626",
      linha1: linhas[0] || "",
      linha2: linhas[1] || "",
    });
  } else if (M.semSub.length > 0) {
    riscos.push({
      cor:"6B3FA0", bordaCor:"6B3FA0", bgCor:"F5F3FF",
      tag:"REFINAMENTO PENDENTE",
      badge:`${M.semSub.length}\n${M.semSub.length===1?"story":"stories"}`, badgeCor:"6B3FA0",
      linha1:`${M.semSub.length} ${M.semSub.length===1?"story":"stories"} sem subtarefas`,
      linha2: M.semSubCamadas,
    });
  } else {
    riscos.push({
      cor:"1D9E75", bordaCor:"1D9E75", bgCor:"F0FDF4",
      tag:"SEM RISCO",
      badge:"0\nriscos", badgeCor:"1D9E75",
      linha1:"Nenhum estouro ou refinamento pendente",
      linha2:"",
    });
  }

  // R2 — Janela Observabilidade
  riscos.push({
    cor:"6B3FA0", bordaCor:"6B3FA0", bgCor:"F5F3FF",
    tag:"ATENCAO · Janela de Entrega",
    badge:"~25 dias\nuteis", badgeCor:"6B3FA0",
    linha1:`Observabilidade: Projetado ${M.dataProjObs} · Base ${DATA_BASE_OBS}`,
    linha2:"",
  });

  // R3 — Bug/SUS
  riscos.push({
    cor:"DC2626", bordaCor:"DC2626", bgCor:"FFF1F2",
    tag:"BUG / SUSTENTACAO",
    badge:`${M.horasR3}h\nrealizadas`, badgeCor:"DC2626",
    linha1:`${M.bugs} Bugs · ${M.srs} SRs no periodo`,
    linha2:`Horas nao planejadas: ${M.horasR3}h`,
  });

  // R4 — Expansão
  const novosPorTipo = {};
  // contar stories e subs nos novos
  riscos.push({
    cor:"D97706", bordaCor:"D97706", bgCor:"FFFBEB",
    tag:"EXPANSAO DE ESCOPO",
    badge:`${M.novosTotal}\nitens`, badgeCor:"D97706",
    linha1:`${M.novosTotal} itens criados pos 01/05`,
    linha2: M.novosTxt,
  });

  const D = {
    data: (() => {
      const now = new Date(new Date().toLocaleString("en-US", {timeZone:"America/Sao_Paulo"}));
      const meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
      return `${now.getDate()} de ${meses[now.getMonth()]} de ${now.getFullYear()} as ${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;
    })(),
    total: M.total, conc: M.conc, prog: M.prog, afbl: M.afbl, pct: M.pct, vel:"18h/dia",
    camadas: M.camadas,
    planejado: M.plan, jaGasto: +M.real.toFixed(1), falta: M.falta,
    aumentoH: `${M.deltaH >= 0 ? "+" : ""}${M.deltaH}h`,
    aumentoP: `${M.deltaP >= 0 ? "+" : ""}${M.deltaP}%`,
    proj_total: +M.proj.toFixed(1),
    data_proj: M.dataProjObs, data_proj_camadas: M.dataProjE1,
    data_base: DATA_BASE_E1, data_pess: DATA_PESS_E1,
    data_base_obs: DATA_BASE_OBS, data_pess_obs: DATA_PESS_OBS,
    ef_estourados: `${pEstourados}%`, ef_retrabalho_pct: `${pRetrab}%`,
    ef_retrabalho_bugs: M.subBugsAll, ef_retrabalho_h: `${M.horasSubBugs}h`,
    ef_dentro: `${pDentro}%`,
    riscos,
    devs: M.devs,
    qa:   M.qa,
    _e1Pct: e1Pct, _e2Pct: e2Pct,
  };

  // Cor do delta por camada
  D.camadas.forEach(cam => {
    const p = parseFloat(cam.deltaP);
    if (p < 0)        cam.projCor = "1D9E75";
    else if (p <= 10) cam.projCor = "1D9E75";
    else if (p <= 25) cam.projCor = "D97706";
    else              cam.projCor = "DC2626";
  });

  // ── LAYOUT PPTX ────────────────────────────────────────────────────
  const C = {
    NAVY:"1C2B4A", WHITE:"FFFFFF", OFFWHITE:"F8FAFC",
    LGRAY:"E8EDF2", GRAY:"888888", DARK:"222222",
    BORDER:"DDDDDD", SIDEBAR_DIM:"8FA3C0", SIDEBAR_SEP:"2D4A73",
  };
  const mkShadow = () => ({ type:"outer", blur:3, offset:2, angle:90, color:"CCCCCC", opacity:0.25 });

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.title  = "Status Report — Quadro de Aulas";

  const slide = pres.addSlide();
  slide.background = { color: C.OFFWHITE };

  const W=13.3, H=7.5, SBW=2.55;
  const MX=SBW+0.12, MW=W-MX-0.08;

  // SIDEBAR
  slide.addShape(pres.shapes.RECTANGLE, {x:0,y:0,w:SBW,h:H,fill:{color:C.NAVY},line:{color:C.NAVY}});
  slide.addText("STATUS REPORT",{x:0.15,y:0.18,w:SBW-0.20,h:0.18,fontSize:6.5,color:C.SIDEBAR_DIM,bold:true,charSpacing:4,align:"left",margin:0});
  slide.addText("Quadro de Aulas",{x:0.15,y:0.36,w:SBW-0.20,h:0.34,fontSize:13.5,color:C.WHITE,bold:true,align:"left",margin:0});
  slide.addText("Resiliencia Sistemica",{x:0.15,y:0.70,w:SBW-0.20,h:0.18,fontSize:8,color:C.SIDEBAR_DIM,italic:true,align:"left",margin:0});
  slide.addText(D.data,{x:0.15,y:0.90,w:SBW-0.20,h:0.16,fontSize:7,color:C.SIDEBAR_DIM,align:"left",margin:0});
  slide.addShape(pres.shapes.RECTANGLE,{x:0.15,y:1.10,w:SBW-0.30,h:0.015,fill:{color:C.SIDEBAR_SEP},line:{color:C.SIDEBAR_SEP}});

  const sbKpis=[{label:"TOTAL DE ITENS",val:String(D.total),vc:C.WHITE},{label:"CONCLUIDOS",val:String(D.conc),vc:"4ADE80"},{label:"EM PROGRESSO",val:String(D.prog),vc:"FCD34D"},{label:"A FAZER / BACKLOG",val:String(D.afbl),vc:"F87171"}];
  let sy=1.18;
  sbKpis.forEach(k=>{slide.addText(k.label,{x:0.15,y:sy,w:SBW-0.20,h:0.16,fontSize:6.5,color:C.SIDEBAR_DIM,bold:true,charSpacing:2,align:"left",margin:0});slide.addText(k.val,{x:0.15,y:sy+0.16,w:SBW-0.20,h:0.38,fontSize:24,color:k.vc,bold:true,align:"left",margin:0});sy+=0.62;});

  const pbY=sy+0.04;
  slide.addText("PROGRESSO GERAL",{x:0.15,y:pbY,w:SBW-0.20,h:0.15,fontSize:6.5,color:C.SIDEBAR_DIM,bold:true,charSpacing:2,align:"left",margin:0});
  const barY=pbY+0.17,barW=SBW-0.30;
  slide.addShape(pres.shapes.RECTANGLE,{x:0.15,y:barY,w:barW,h:0.10,fill:{color:"1E3A5F"},line:{color:"1E3A5F"}});
  slide.addShape(pres.shapes.RECTANGLE,{x:0.15,y:barY,w:barW*(D.pct/100),h:0.10,fill:{color:"4ADE80"},line:{color:"4ADE80"}});
  slide.addText(`${D.pct}% concluido  .  ${D.conc} de ${D.total} itens`,{x:0.15,y:barY+0.12,w:SBW-0.20,h:0.16,fontSize:7,color:C.SIDEBAR_DIM,align:"left",margin:0});
  slide.addText("Cancelados excluidos do escopo",{x:0.15,y:barY+0.28,w:SBW-0.20,h:0.14,fontSize:6.5,color:C.SIDEBAR_DIM,align:"left",margin:0});
  slide.addText("VELOCIDADE RECENTE",{x:0.15,y:barY+0.48,w:SBW-0.20,h:0.15,fontSize:6.5,color:C.SIDEBAR_DIM,bold:true,charSpacing:2,align:"left",margin:0});
  slide.addText(D.vel,{x:0.15,y:barY+0.63,w:SBW-0.20,h:0.34,fontSize:22,color:"FCD34D",bold:true,align:"left",margin:0});
  slide.addText("velocidade fixada pelo time",{x:0.15,y:barY+0.97,w:SBW-0.20,h:0.14,fontSize:6.5,color:C.SIDEBAR_DIM,align:"left",margin:0});
  slide.addText("Squad Lecionar",{x:0.15,y:7.10,w:SBW-0.20,h:0.16,fontSize:7,color:C.SIDEBAR_DIM,align:"left",margin:0});
  slide.addText("Cintia Baricatti",{x:0.15,y:7.27,w:SBW-0.20,h:0.16,fontSize:7,color:C.SIDEBAR_DIM,align:"left",margin:0});

  // SEÇÃO 1 — CAMADAS
  const S1Y=0.08,S1H=2.20,CW=(MW-0.09*3)/4;
  D.camadas.forEach((cam,ci)=>{
    const cx=MX+ci*(CW+0.09);
    const isCanceled=cam.cancelado===true;
    const isComplete=!isCanceled&&cam.pct===100&&cam.pais_pct===100;
    const bordaCor=isComplete?"1D9E75":isCanceled?"DC2626":C.BORDER;
    if(isCanceled){
      slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:S1Y,w:CW,h:0.28,fill:{color:cam.cor},line:{color:cam.cor}});
      slide.addText(cam.nome,{x:cx,y:S1Y,w:CW*0.62,h:0.28,fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle",margin:0});
      slide.addText("·",{x:cx+CW*0.59,y:S1Y,w:CW*0.08,h:0.28,fontSize:9,color:"FFAAAA",bold:false,align:"center",valign:"middle",margin:0});
      slide.addText("CANCELADO",{x:cx+CW*0.63,y:S1Y,w:CW*0.37,h:0.28,fontSize:8,color:"FFD5D5",bold:true,align:"center",valign:"middle",margin:0});
    }else if(isComplete){
      const darkCor=cam.cor==="1D9E75"?"085041":cam.cor==="534AB7"?"26215C":cam.cor==="D97706"?"633806":"042C53";
      slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:S1Y,w:CW,h:0.28,fill:{color:darkCor},line:{color:darkCor}});
      slide.addText(cam.nome,{x:cx,y:S1Y,w:CW*0.55,h:0.28,fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle",margin:0});
      slide.addText("·",{x:cx+CW*0.52,y:S1Y,w:CW*0.08,h:0.28,fontSize:9,color:"9FE1CB",bold:false,align:"center",valign:"middle",margin:0});
      slide.addText("v ENTREGUE",{x:cx+CW*0.56,y:S1Y,w:CW*0.44,h:0.28,fontSize:8,color:"9FE1CB",bold:true,align:"center",valign:"middle",margin:0});
    }else{
      slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:S1Y,w:CW,h:0.28,fill:{color:cam.cor},line:{color:cam.cor}});
      slide.addText(cam.nome,{x:cx,y:S1Y,w:CW,h:0.28,fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle",margin:0});
    }
    const bodyY=S1Y+0.28,bodyH=S1H-0.28;
    slide.addShape(pres.shapes.RECTANGLE,{x:cx,y:bodyY,w:CW,h:bodyH,fill:{color:C.WHITE},line:{color:bordaCor,pt:isComplete?1.5:1},shadow:mkShadow()});
    const storiesNFT=[cam.stories>0?`${cam.stories} ${cam.stories===1?"Story":"Stories"}`:null,cam.nft>0?`${cam.nft} NFT`:null].filter(Boolean).join(" · ");
    const l1=storiesNFT?`${storiesNFT} · ${cam.itens} itens`:`${cam.itens} itens`;
    slide.addText(l1,{x:cx+0.08,y:bodyY+0.08,w:CW-0.16,h:0.18,fontSize:7.5,color:C.DARK,align:"left",margin:0});
    slide.addText(`A Fazer: ${cam.af} . Em progresso: ${cam.prog} . Finalizados: ${cam.conc}`,{x:cx+0.08,y:bodyY+0.26,w:CW-0.16,h:0.16,fontSize:7.5,color:C.DARK,align:"left",margin:0});
    slide.addText(`Est: ${cam.est}h . Real: ${cam.real}h . Rem: ${cam.rem}h`,{x:cx+0.08,y:bodyY+0.42,w:CW-0.16,h:0.16,fontSize:7.5,color:C.DARK,align:"left",margin:0});
    const bY=bodyY+0.58,bW=CW-0.16;
    slide.addShape(pres.shapes.RECTANGLE,{x:cx+0.08,y:bY,w:bW,h:0.055,fill:{color:C.LGRAY},line:{color:C.LGRAY}});
    slide.addShape(pres.shapes.RECTANGLE,{x:cx+0.08,y:bY,w:bW*(cam.pais_pct/100),h:0.055,fill:{color:cam.cor},line:{color:cam.cor},transparency:40});
    slide.addText(`Stories/NFTs: ${cam.pais_pct}% · ${cam.pais_conc} de ${cam.pais_total}`,{x:cx+0.08,y:bY+0.06,w:CW-0.16,h:0.12,fontSize:6.5,color:C.GRAY,align:"left",margin:0});
    const bY2=bY+0.19;
    slide.addShape(pres.shapes.RECTANGLE,{x:cx+0.08,y:bY2,w:bW,h:0.055,fill:{color:C.LGRAY},line:{color:C.LGRAY}});
    slide.addShape(pres.shapes.RECTANGLE,{x:cx+0.08,y:bY2,w:bW*(cam.pct/100),h:0.055,fill:{color:cam.cor},line:{color:cam.cor}});
    slide.addText(`Subs: ${cam.pct}% · ${cam.conc} de ${cam.itens}`,{x:cx+0.08,y:bY2+0.06,w:CW-0.16,h:0.12,fontSize:6.5,color:C.GRAY,align:"left",margin:0});
    slide.addText(`Projetado: ${cam.proj}h (${cam.deltaH} / ${cam.deltaP})`,{x:cx+0.08,y:bY2+0.22,w:CW-0.16,h:0.18,fontSize:8,color:cam.projCor,bold:true,align:"left",margin:0});
    const eY=bY2+0.42;
    if(cam.estouros_ativos.length>0){
      slide.addText(`${cam.estouros_ativos.length} ativo(s) com horas estouradas`,{x:cx+0.08,y:eY,w:CW-0.16,h:0.16,fontSize:7.5,color:cam.cor,bold:true,align:"left",margin:0});
      const ecY=eY+0.18,ecH=bodyH-(ecY-bodyY)-0.06;
      cam.estouros_ativos.slice(0,1).forEach(ov=>{
        slide.addShape(pres.shapes.RECTANGLE,{x:cx+0.08,y:ecY,w:CW-0.16,h:ecH,fill:{color:"FFF8F8"},line:{color:"FCA5A5"}});
        slide.addText(ov.titulo?.substring(0,40)||"",{x:cx+0.10,y:ecY+0.02,w:CW-0.20,h:0.14,fontSize:6.5,color:C.DARK,align:"left",margin:0});
        slide.addText(`+${ov.dH}h (+${ov.dP}%)`,{x:cx+0.10,y:ecY+0.16,w:CW-0.20,h:0.14,fontSize:7,color:"DC2626",bold:true,align:"left",margin:0});
      });
    }
  });

  // SEÇÃO 2 — HORAS
  const S2Y=S1Y+S1H+0.04,S2H=0.94;
  slide.addShape(pres.shapes.RECTANGLE,{x:MX,y:S2Y,w:MW,h:S2H,fill:{color:C.WHITE},line:{color:C.BORDER,pt:1},shadow:mkShadow()});
  slide.addText("HORAS - ESTIMADO vs. PROJETADO",{x:MX+0.10,y:S2Y+0.06,w:MW-0.20,h:0.16,fontSize:7,color:C.GRAY,charSpacing:2,align:"left",margin:0});
  const hCols=[{label:"PLANEJADO",val:`${D.planejado}h`,sub:"original estimate",cor:C.DARK},{label:"JA GASTO",val:`${D.jaGasto}h`,sub:"completed work",cor:"D97706"},{label:"FALTA",val:`${D.falta}h`,sub:"remaining (devs)",cor:"F87171"},{label:"AUMENTO DE ESFORCO",val:D.aumentoH,sub:`projetado: ${D.proj_total}h (${D.aumentoP})`,cor:"D97706"}];
  const hW=MW/4;
  hCols.forEach((h,hi)=>{
    const hx=MX+hi*hW;
    if(hi>0)slide.addShape(pres.shapes.RECTANGLE,{x:hx,y:S2Y+0.22,w:0.01,h:S2H-0.28,fill:{color:C.LGRAY},line:{color:C.LGRAY}});
    slide.addText(h.label,{x:hx+0.10,y:S2Y+0.22,w:hW-0.20,h:0.16,fontSize:7,color:C.GRAY,charSpacing:1,bold:true,align:"left",margin:0});
    slide.addText(h.val,{x:hx+0.10,y:S2Y+0.36,w:hW-0.20,h:0.34,fontSize:24,color:h.cor,bold:true,align:"left",margin:0});
    slide.addText(h.sub,{x:hx+0.10,y:S2Y+0.72,w:hW-0.20,h:0.14,fontSize:7,color:C.GRAY,align:"left",margin:0});
  });
  const pbGastoW=MW*(D.jaGasto/D.proj_total),pbFaltaW=MW*(D.falta/D.proj_total);
  const pbgY=S2Y+S2H-0.07;
  slide.addShape(pres.shapes.RECTANGLE,{x:MX,y:pbgY,w:pbGastoW,h:0.06,fill:{color:"D97706"},line:{color:"D97706"}});
  slide.addShape(pres.shapes.RECTANGLE,{x:MX+pbGastoW,y:pbgY,w:pbFaltaW,h:0.06,fill:{color:"7C3AED"},line:{color:"7C3AED"}});

  // SEÇÃO 2B — PREVISÃO
  const S2BY=S2Y+S2H+0.03,S2BH=0.96;
  const halfW=(MW-0.10)/2;

  // E1
  const pv1X=MX,pv1W=halfW;
  slide.addShape(pres.shapes.RECTANGLE,{x:pv1X,y:S2BY,w:pv1W,h:S2BH,fill:{color:C.WHITE},line:{color:C.BORDER,pt:1},shadow:mkShadow()});
  slide.addText("Entrega 1 · Resiliencia Sistemica",{x:pv1X+0.12,y:S2BY+0.08,w:pv1W-0.84,h:0.18,fontSize:8,bold:true,color:"1F2937",align:"left",margin:0});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv1X+pv1W-0.72,y:S2BY+0.07,w:0.62,h:0.20,fill:{color:"E1F5EE"},line:{color:"E1F5EE"},rectRadius:0.08});
  slide.addText(`${D._e1Pct}% concluido`,{x:pv1X+pv1W-0.72,y:S2BY+0.07,w:0.62,h:0.20,fontSize:6.5,bold:true,color:"0F6E56",align:"center",valign:"middle",margin:0});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv1X+0.12,y:S2BY+0.31,w:pv1W-0.24,h:0.07,fill:{color:"E1F5EE"},line:{color:"E1F5EE"}});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv1X+0.12,y:S2BY+0.31,w:(pv1W-0.24)*(D._e1Pct/100),h:0.07,fill:{color:"1D9E75"},line:{color:"1D9E75"}});
  const pv1ColW=(pv1W-0.24)/3;
  [{label:"PROJETADO",cor:"1D9E75",val:D.data_proj_camadas},{label:"BASE",cor:"D97706",val:D.data_base},{label:"PESSIMISTA",cor:"DC2626",val:D.data_pess}].forEach((pv,pi)=>{
    const bx=pv1X+0.12+pi*pv1ColW;
    slide.addShape(pres.shapes.RECTANGLE,{x:bx,y:S2BY+0.44,w:pv1ColW-0.04,h:0.22,fill:{color:pv.cor},line:{color:pv.cor}});
    slide.addText(pv.label,{x:bx,y:S2BY+0.44,w:pv1ColW-0.04,h:0.22,fontSize:6.5,color:C.WHITE,bold:true,align:"center",valign:"middle",margin:0});
    slide.addText(pv.val,{x:bx,y:S2BY+0.68,w:pv1ColW-0.04,h:0.20,fontSize:10.5,color:pv.cor,bold:true,align:"center",margin:0});
  });

  // E2
  const pv2X=MX+halfW+0.10,pv2W=halfW;
  slide.addShape(pres.shapes.RECTANGLE,{x:pv2X,y:S2BY,w:pv2W,h:S2BH,fill:{color:C.WHITE},line:{color:C.BORDER,pt:1},shadow:mkShadow()});
  slide.addText("Entrega 2 · Observabilidade",{x:pv2X+0.12,y:S2BY+0.08,w:pv2W-0.84,h:0.18,fontSize:8,bold:true,color:"1F2937",align:"left",margin:0});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv2X+pv2W-0.72,y:S2BY+0.07,w:0.62,h:0.20,fill:{color:"FEE2E2"},line:{color:"FEE2E2"},rectRadius:0.08});
  slide.addText(`${D._e2Pct}% concluido`,{x:pv2X+pv2W-0.72,y:S2BY+0.07,w:0.62,h:0.20,fontSize:6.5,bold:true,color:"991B1B",align:"center",valign:"middle",margin:0});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv2X+0.12,y:S2BY+0.31,w:pv2W-0.24,h:0.07,fill:{color:"E6F1FB"},line:{color:"E6F1FB"}});
  slide.addShape(pres.shapes.RECTANGLE,{x:pv2X+0.12,y:S2BY+0.31,w:(pv2W-0.24)*(D._e2Pct/100),h:0.07,fill:{color:"1E6FA8"},line:{color:"1E6FA8"}});
  const pv2ColW=(pv2W-0.24)/3;
  [{label:"PROJETADO",cor:"1D9E75",val:D.data_proj},{label:"BASE",cor:"D97706",val:D.data_base_obs},{label:"PESSIMISTA",cor:"DC2626",val:D.data_pess_obs}].forEach((pv,pi)=>{
    const bx=pv2X+0.12+pi*pv2ColW;
    slide.addShape(pres.shapes.RECTANGLE,{x:bx,y:S2BY+0.44,w:pv2ColW-0.04,h:0.22,fill:{color:pv.cor},line:{color:pv.cor}});
    slide.addText(pv.label,{x:bx,y:S2BY+0.44,w:pv2ColW-0.04,h:0.22,fontSize:6.5,color:C.WHITE,bold:true,align:"center",valign:"middle",margin:0});
    slide.addText(pv.val,{x:bx,y:S2BY+0.68,w:pv2ColW-0.04,h:0.20,fontSize:10.5,color:pv.cor,bold:true,align:"center",margin:0});
  });

  // SEÇÃO 3 — EFICIÊNCIA
  const S3Y=S2BY+S2BH+0.03,S3H=0.90;
  slide.addShape(pres.shapes.RECTANGLE,{x:MX,y:S3Y,w:MW,h:S3H,fill:{color:C.WHITE},line:{color:C.BORDER,pt:1},shadow:mkShadow()});
  slide.addText("EFICIENCIA DA ENTREGA",{x:MX+0.10,y:S3Y+0.06,w:MW-0.20,h:0.16,fontSize:7,color:C.GRAY,charSpacing:2,align:"left",margin:0});
  const efCols=[{val:D.ef_estourados,label:"Itens Estourados",sub:"das subs com estimativa",cor:"D85A30"},{val:D.ef_retrabalho_pct,label:"Retrabalho",sub:`${D.ef_retrabalho_bugs} Sub Bug . ${D.ef_retrabalho_h} realizadas`,cor:"534AB7"},{val:D.ef_dentro,label:"Dentro da Estimativa",sub:"entregues dentro do planejado",cor:"1D9E75"}];
  const efColW=MW/3;
  efCols.forEach((ef,ei)=>{
    const efx=MX+ei*efColW;
    if(ei>0)slide.addShape(pres.shapes.RECTANGLE,{x:efx,y:S3Y+0.16,w:0.01,h:S3H-0.22,fill:{color:C.LGRAY},line:{color:C.LGRAY}});
    slide.addText(ef.val,{x:efx+0.10,y:S3Y+0.16,w:efColW-0.20,h:0.36,fontSize:22,color:ef.cor,bold:true,align:"center",margin:0});
    slide.addText(ef.label,{x:efx+0.10,y:S3Y+0.52,w:efColW-0.20,h:0.18,fontSize:8,color:C.DARK,bold:true,align:"center",margin:0});
    slide.addText(ef.sub,{x:efx+0.10,y:S3Y+0.70,w:efColW-0.20,h:0.18,fontSize:7,color:C.GRAY,align:"center",margin:0});
  });

  // SEÇÃO 4 — RISCOS
  const S4Y=S3Y+S3H+0.15,S4H=0.96,RW=(MW-0.09*3)/4;
  slide.addText("R I S C O S",{x:MX,y:S4Y-0.13,w:MW,h:0.12,fontSize:6.5,color:C.GRAY,charSpacing:3,align:"left",margin:0});
  D.riscos.forEach((r,ri)=>{
    const rx=MX+ri*(RW+0.09);
    slide.addShape(pres.shapes.RECTANGLE,{x:rx,y:S4Y,w:RW,h:S4H,fill:{color:r.bgCor},line:{color:r.bordaCor,pt:1.5},shadow:mkShadow()});
    slide.addText(r.tag,{x:rx+0.08,y:S4Y+0.07,w:RW-0.65,h:0.20,fontSize:7,color:r.cor,bold:true,align:"left",margin:0});
    slide.addShape(pres.shapes.RECTANGLE,{x:rx+RW-0.58,y:S4Y+0.06,w:0.52,h:0.32,fill:{color:r.badgeCor},line:{color:r.badgeCor}});
    slide.addText(r.badge,{x:rx+RW-0.58,y:S4Y+0.06,w:0.52,h:0.32,fontSize:6.5,color:C.WHITE,bold:true,align:"center",valign:"middle",margin:0});
    slide.addText(r.linha1,{x:rx+0.08,y:S4Y+0.34,w:RW-0.16,h:0.20,fontSize:7.5,color:C.DARK,align:"left",margin:0});
    slide.addText(r.linha2,{x:rx+0.08,y:S4Y+0.54,w:RW-0.16,h:0.18,fontSize:7.5,color:C.DARK,align:"left",margin:0});
  });

  // SEÇÃO 5 — TIME
  const S5Y=S4Y+S4H+0.05,S5H=H-S5Y-0.04;
  const nDevs=D.devs.length,nQA=D.qa.length,totalCards=nDevs+nQA;
  const TW=(MW-0.09*(totalCards-1))/totalCards,cardH=S5H-0.09;
  const devsFaixaW=nDevs*TW+(nDevs-1)*0.09;
  slide.addShape(pres.shapes.RECTANGLE,{x:MX,y:S5Y,w:devsFaixaW,h:0.08,fill:{color:"374151"},line:{color:"374151"}});
  slide.addText("D E V S",{x:MX,y:S5Y,w:devsFaixaW,h:0.08,fontSize:7,color:C.WHITE,bold:true,align:"center",valign:"middle",margin:0});
  const qaStartX=MX+nDevs*(TW+0.09),qaFaixaW=nQA*TW+(nQA-1)*0.09;
  slide.addShape(pres.shapes.RECTANGLE,{x:qaStartX,y:S5Y,w:qaFaixaW,h:0.08,fill:{color:"1E6FA8"},line:{color:"1E6FA8"}});
  slide.addText("Q A",{x:qaStartX,y:S5Y,w:qaFaixaW,h:0.08,fontSize:7,color:C.WHITE,bold:true,align:"center",valign:"middle",margin:0});

  [...D.devs,...D.qa].forEach((dev,di)=>{
    const isQA=di>=nDevs;
    const tx=isQA?qaStartX+(di-nDevs)*(TW+0.09):MX+di*(TW+0.09);
    const cardY=S5Y+0.09;
    slide.addShape(pres.shapes.RECTANGLE,{x:tx,y:cardY,w:TW,h:cardH,fill:{color:C.WHITE},line:{color:C.BORDER,pt:1},shadow:mkShadow()});
    slide.addText(dev.nome,{x:tx+0.06,y:cardY+0.08,w:TW-0.12,h:0.20,fontSize:8.5,color:C.DARK,bold:true,align:"center",margin:0});
    const bpY=cardY+0.30,bpW=TW-0.12;
    slide.addShape(pres.shapes.RECTANGLE,{x:tx+0.06,y:bpY,w:bpW,h:0.08,fill:{color:C.LGRAY},line:{color:C.LGRAY}});
    slide.addShape(pres.shapes.RECTANGLE,{x:tx+0.06,y:bpY,w:bpW*Math.min(dev.real/(dev.est||1),1),h:0.08,fill:{color:dev.cor},line:{color:dev.cor}});
    const subsLabel=dev.subs>0?`${dev.real}h · ${dev.subs} subs finalizadas`:`${dev.real}h realizadas`;
    slide.addText(subsLabel,{x:tx+0.06,y:bpY+0.10,w:bpW,h:0.24,fontSize:9.5,color:dev.cor,bold:true,align:"center",margin:0});
    if(dev.est>0)slide.addText(`${dev.est}h estimadas`,{x:tx+0.06,y:bpY+0.34,w:bpW,h:0.18,fontSize:8.5,color:C.GRAY,align:"center",margin:0});
  });

  await pres.writeFile({ fileName: OUTPUT_PPTX });
  console.log(`\n✅ PPTX gerado: ${OUTPUT_PPTX}`);
}

// ── GERAR JPG ──────────────────────────────────────────────────────────
// ── GERAR JPG (espelho exato do PPTX, fator 144px/polegada) ────────────
async function gerarJPG(M) {
  const S = 144; // scale: 1 polegada = 144px
  const W = Math.round(13.3*S), H = Math.round(7.5*S);
  const canvas = createCanvas(W, H);
  const ctx = canvas.getContext("2d");

  // Preparar D idêntico ao gerarPPTX
  const e1Cams    = M.camadas.filter(c => c.nome !== "OBSERVABILIDADE");
  const e1SubsFin = e1Cams.reduce((s,c) => s + Math.round(c.pct/100 * c.itens), 0);
  const e1SubsTot = e1Cams.reduce((s,c) => s + c.itens, 0);
  const e1Pct     = e1SubsTot > 0 ? Math.round(e1SubsFin / e1SubsTot * 100) : 0;
  const e2Cam     = M.camadas.find(c => c.nome === "OBSERVABILIDADE");
  const e2Pct     = e2Cam ? e2Cam.pct : 0;
  const pEstourados = M.totEf > 0 ? (M.estourados/M.totEf*100).toFixed(1) : "0.0";
  const pDentro     = M.totEf > 0 ? (M.dentro/M.totEf*100).toFixed(1) : "0.0";
  const pRetrab     = M.totEf > 0 ? (M.subBugsAll/M.totEf*100).toFixed(1) : "0.0";

  const totalEstouros = M.camadas.reduce((s,c) => s + c.estouros_ativos.length, 0);
  const riscos = [];
  if (totalEstouros > 0) {
    const linhas = M.camadas.flatMap(c => c.estouros_ativos.map(e => `#${e.id} +${e.dH}h(+${e.dP}%) [${e.dev}]`));
    riscos.push({ cor:"DC2626", bordaCor:"DC2626", bgCor:"FFF1F2", tag:"RISCO ATIVO",
      badge:`${totalEstouros}\nestouros`, badgeCor:"DC2626", linha1:linhas[0]||"", linha2:linhas[1]||"" });
  } else if (M.semSub.length > 0) {
    riscos.push({ cor:"6B3FA0", bordaCor:"6B3FA0", bgCor:"F5F3FF", tag:"REFINAMENTO PENDENTE",
      badge:`${M.semSub.length}\n${M.semSub.length===1?"story":"stories"}`, badgeCor:"6B3FA0",
      linha1:`${M.semSub.length} ${M.semSub.length===1?"story":"stories"} sem subtarefas`, linha2:M.semSubCamadas });
  } else {
    riscos.push({ cor:"1D9E75", bordaCor:"1D9E75", bgCor:"F0FDF4", tag:"SEM RISCO",
      badge:"0\nriscos", badgeCor:"1D9E75", linha1:"Nenhum estouro ou refinamento pendente", linha2:"" });
  }
  riscos.push({ cor:"6B3FA0", bordaCor:"6B3FA0", bgCor:"F5F3FF", tag:"ATENCAO · Janela de Entrega",
    badge:"~25 dias\nuteis", badgeCor:"6B3FA0",
    linha1:`Observabilidade: Projetado ${M.dataProjObs} · Base ${DATA_BASE_OBS}`, linha2:"" });
  riscos.push({ cor:"DC2626", bordaCor:"DC2626", bgCor:"FFF1F2", tag:"BUG / SUSTENTACAO",
    badge:`${M.horasR3}h\nrealizadas`, badgeCor:"DC2626",
    linha1:`${M.bugs} Bugs · ${M.srs} SRs no periodo`, linha2:`Horas nao planejadas: ${M.horasR3}h` });
  riscos.push({ cor:"D97706", bordaCor:"D97706", bgCor:"FFFBEB", tag:"EXPANSAO DE ESCOPO",
    badge:`${M.novosTotal}\nitens`, badgeCor:"D97706",
    linha1:`${M.novosTotal} itens criados pos 01/05`, linha2:M.novosTxt });

  const now = new Date(new Date().toLocaleString("en-US", {timeZone:"America/Sao_Paulo"}));
  const meses = ["Janeiro","Fevereiro","Marco","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
  const dataStr = `${now.getDate()} de ${meses[now.getMonth()]} de ${now.getFullYear()} as ${String(now.getHours()).padStart(2,"0")}:${String(now.getMinutes()).padStart(2,"0")}`;

  M.camadas.forEach(cam => {
    const p = parseFloat(cam.deltaP);
    cam.projCor = p <= 10 ? "1D9E75" : p <= 25 ? "D97706" : "DC2626";
  });

  const D = {
    data: dataStr, total:M.total, conc:M.conc, prog:M.prog, afbl:M.afbl, pct:M.pct, vel:"18h/dia",
    camadas:M.camadas, planejado:M.plan, jaGasto:+M.real.toFixed(1), falta:M.falta,
    aumentoH:`${M.deltaH>=0?"+":""}${M.deltaH}h`, aumentoP:`${M.deltaP>=0?"+":""}${M.deltaP}%`,
    proj_total:+M.proj.toFixed(1), data_proj:M.dataProjObs, data_proj_camadas:M.dataProjE1,
    data_base:DATA_BASE_E1, data_pess:DATA_PESS_E1, data_base_obs:DATA_BASE_OBS, data_pess_obs:DATA_PESS_OBS,
    ef_estourados:`${pEstourados}%`, ef_retrabalho_pct:`${pRetrab}%`,
    ef_retrabalho_bugs:M.subBugsAll, ef_retrabalho_h:`${M.horasSubBugs}h`, ef_dentro:`${pDentro}%`,
    riscos, devs:M.devs, qa:M.qa, _e1Pct:e1Pct, _e2Pct:e2Pct,
  };

  // Helpers canvas
  const p = v => Math.round(v * S); // polegadas → pixels
  const hex = c => c.startsWith("#") ? c : "#"+c;

  function fillRect(x,y,w,h,color) {
    ctx.fillStyle = hex(color); ctx.fillRect(p(x),p(y),p(w),p(h));
  }
  function strokeRect(x,y,w,h,color,lw=1) {
    ctx.strokeStyle = hex(color); ctx.lineWidth = lw; ctx.strokeRect(p(x)+0.5,p(y)+0.5,p(w)-1,p(h)-1);
  }
  function fillAndStroke(x,y,w,h,fill,stroke,lw=1) {
    fillRect(x,y,w,h,fill); if(stroke) strokeRect(x,y,w,h,stroke,lw);
  }
  function addText(text, x, y, w, h, opts={}) {
    const {fontSize=7, color="222222", bold=false, align="left", valign="top", italic=false} = opts;
    const px_size = Math.round(fontSize * S / 72 * 1.5);
    ctx.font = `${italic?"italic ":""}${bold?"bold ":""}${px_size}px Arial`;
    ctx.fillStyle = hex(color);
    ctx.textAlign = align === "center" ? "center" : align === "right" ? "right" : "left";
    const textX = align === "center" ? p(x)+p(w)/2 : align === "right" ? p(x)+p(w) : p(x);
    const lineH = px_size * 1.25;
    const str = String(text);
    // wrap text
    const words = str.split(" ");
    const maxW = p(w);
    let lines = [], line = "";
    for (const word of words) {
      const test = line ? line+" "+word : word;
      if (ctx.measureText(test).width > maxW && line) { lines.push(line); line = word; }
      else line = test;
    }
    lines.push(line);
    const totalH = lines.length * lineH;
    let startY = p(y);
    if (valign === "middle") startY = p(y) + (p(h) - totalH) / 2 + px_size;
    else startY = p(y) + px_size;
    lines.forEach((l, i) => {
      if (p(y) + p(h) > startY + i*lineH - px_size)
        ctx.fillText(l, textX, startY + i*lineH);
    });
    ctx.textAlign = "left";
  }
  function addBar(x,y,w,h,bgColor,fillColor,pct) {
    fillRect(x,y,w,h,bgColor);
    if (pct > 0) fillRect(x,y,w*(Math.min(pct,100)/100),h,fillColor);
  }

  // Fundo
  fillRect(0,0,13.3,7.5,"F8FAFC");

  const SBW=2.55, MX=SBW+0.12, MW=13.3-MX-0.08;

  // ── SIDEBAR ──
  fillRect(0,0,SBW,7.5,"1C2B4A");
  addText("STATUS REPORT",       0.15,0.18,SBW-0.20,0.18, {fontSize:6.5,color:"8FA3C0",bold:true});
  addText("Quadro de Aulas",     0.15,0.36,SBW-0.20,0.34, {fontSize:13.5,color:"FFFFFF",bold:true});
  addText("Resiliencia Sistemica",0.15,0.70,SBW-0.20,0.18,{fontSize:8,color:"8FA3C0",italic:true});
  addText(D.data,                 0.15,0.90,SBW-0.20,0.16, {fontSize:7,color:"8FA3C0"});
  fillRect(0.15,1.10,SBW-0.30,0.015,"2D4A73");

  const sbKpis=[{label:"TOTAL DE ITENS",val:String(D.total),vc:"FFFFFF"},{label:"CONCLUIDOS",val:String(D.conc),vc:"4ADE80"},{label:"EM PROGRESSO",val:String(D.prog),vc:"FCD34D"},{label:"A FAZER / BACKLOG",val:String(D.afbl),vc:"F87171"}];
  let sy=1.18;
  sbKpis.forEach(k=>{
    addText(k.label, 0.15,sy,SBW-0.20,0.16, {fontSize:6.5,color:"8FA3C0",bold:true});
    addText(k.val,   0.15,sy+0.16,SBW-0.20,0.38, {fontSize:24,color:k.vc,bold:true});
    sy+=0.62;
  });

  const pbY=sy+0.04, barW=SBW-0.30;
  addText("PROGRESSO GERAL", 0.15,pbY,SBW-0.20,0.15, {fontSize:6.5,color:"8FA3C0",bold:true});
  const barY=pbY+0.17;
  addBar(0.15,barY,barW,0.10,"1E3A5F","4ADE80",D.pct);
  addText(`${D.pct}% concluido  .  ${D.conc} de ${D.total} itens`, 0.15,barY+0.12,SBW-0.20,0.16, {fontSize:7,color:"8FA3C0"});
  addText("Cancelados excluidos do escopo", 0.15,barY+0.28,SBW-0.20,0.14, {fontSize:6.5,color:"8FA3C0"});
  addText("VELOCIDADE RECENTE", 0.15,barY+0.48,SBW-0.20,0.15, {fontSize:6.5,color:"8FA3C0",bold:true});
  addText(D.vel, 0.15,barY+0.63,SBW-0.20,0.34, {fontSize:22,color:"FCD34D",bold:true});
  addText("velocidade fixada pelo time", 0.15,barY+0.97,SBW-0.20,0.14, {fontSize:6.5,color:"8FA3C0"});
  addText("Squad Lecionar",  0.15,7.10,SBW-0.20,0.16, {fontSize:7,color:"8FA3C0"});
  addText("Cintia Baricatti",0.15,7.27,SBW-0.20,0.16, {fontSize:7,color:"8FA3C0"});

  // ── SEÇÃO 1 — CAMADAS ──
  const S1Y=0.08, S1H=2.20, CW=(MW-0.09*3)/4;
  D.camadas.forEach((cam,ci) => {
    const cx = MX+ci*(CW+0.09);
    const isCanceled = cam.cancelado===true;
    const isComplete = !isCanceled && cam.pct===100 && cam.pais_pct===100;
    const bordaCor = isComplete?"1D9E75": isCanceled?"DC2626":"DDDDDD";

    if (isCanceled) {
      fillRect(cx,S1Y,CW,0.28,cam.cor);
      addText(cam.nome,    cx,S1Y,CW*0.62,0.28, {fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
      addText("·",         cx+CW*0.59,S1Y,CW*0.08,0.28, {fontSize:9,color:"FFAAAA",align:"center",valign:"middle"});
      addText("CANCELADO", cx+CW*0.63,S1Y,CW*0.37,0.28, {fontSize:8,color:"FFD5D5",bold:true,align:"center",valign:"middle"});
    } else if (isComplete) {
      const darkCor = cam.cor==="1D9E75"?"085041":cam.cor==="534AB7"?"26215C":cam.cor==="D97706"?"633806":"042C53";
      fillRect(cx,S1Y,CW,0.28,darkCor);
      addText(cam.nome,     cx,S1Y,CW*0.55,0.28, {fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
      addText("·",          cx+CW*0.52,S1Y,CW*0.08,0.28, {fontSize:9,color:"9FE1CB",align:"center",valign:"middle"});
      addText("v ENTREGUE", cx+CW*0.56,S1Y,CW*0.44,0.28, {fontSize:8,color:"9FE1CB",bold:true,align:"center",valign:"middle"});
    } else {
      fillRect(cx,S1Y,CW,0.28,cam.cor);
      addText(cam.nome, cx,S1Y,CW,0.28, {fontSize:8.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
    }

    const bodyY=S1Y+0.28, bodyH=S1H-0.28;
    fillAndStroke(cx,bodyY,CW,bodyH,"FFFFFF",bordaCor, isComplete?1.5:1);

    const storiesNFT=[cam.stories>0?`${cam.stories} ${cam.stories===1?"Story":"Stories"}`:null,cam.nft>0?`${cam.nft} NFT`:null].filter(Boolean).join(" · ");
    const l1 = storiesNFT ? `${storiesNFT} · ${cam.itens} itens` : `${cam.itens} itens`;
    addText(l1, cx+0.08,bodyY+0.08,CW-0.16,0.18, {fontSize:7.5,color:"222222"});
    addText(`A Fazer: ${cam.af} . Em progresso: ${cam.prog} . Finalizados: ${cam.conc}`, cx+0.08,bodyY+0.26,CW-0.16,0.16, {fontSize:7.5,color:"222222"});
    addText(`Est: ${cam.est}h . Real: ${cam.real}h . Rem: ${cam.rem}h`, cx+0.08,bodyY+0.42,CW-0.16,0.16, {fontSize:7.5,color:"222222"});

    const bY=bodyY+0.58, bW=CW-0.16;
    addBar(cx+0.08,bY,bW,0.055,"E8EDF2",cam.cor,cam.pais_pct);
    addText(`Stories/NFTs: ${cam.pais_pct}% · ${cam.pais_conc} de ${cam.pais_total}`, cx+0.08,bY+0.06,CW-0.16,0.12, {fontSize:6.5,color:"888888"});

    const bY2=bY+0.19;
    addBar(cx+0.08,bY2,bW,0.055,"E8EDF2",cam.cor,cam.pct);
    addText(`Subs: ${cam.pct}% · ${cam.conc} de ${cam.itens}`, cx+0.08,bY2+0.06,CW-0.16,0.12, {fontSize:6.5,color:"888888"});
    addText(`Projetado: ${cam.proj}h (${cam.deltaH} / ${cam.deltaP})`, cx+0.08,bY2+0.22,CW-0.16,0.18, {fontSize:8,color:cam.projCor,bold:true});

    const eY=bY2+0.42;
    if (cam.estouros_ativos.length > 0) {
      addText(`${cam.estouros_ativos.length} ativo(s) com horas estouradas`, cx+0.08,eY,CW-0.16,0.16, {fontSize:7.5,color:cam.cor,bold:true});
      const ecY=eY+0.18, ecH=bodyH-(ecY-bodyY)-0.06;
      fillAndStroke(cx+0.08,ecY,CW-0.16,ecH,"FFF8F8","FCA5A5");
      cam.estouros_ativos.slice(0,1).forEach(ov => {
        addText(ov.titulo?.substring(0,40)||"", cx+0.10,ecY+0.02,CW-0.20,0.14, {fontSize:6.5,color:"222222"});
        addText(`+${ov.dH}h (+${ov.dP}%)`, cx+0.10,ecY+0.16,CW-0.20,0.14, {fontSize:7,color:"DC2626",bold:true});
      });
    }
  });

  // ── SEÇÃO 2 — HORAS ──
  const S2Y=S1Y+S1H+0.04, S2H=0.94;
  fillAndStroke(MX,S2Y,MW,S2H,"FFFFFF","DDDDDD");
  addText("HORAS - ESTIMADO vs. PROJETADO", MX+0.10,S2Y+0.06,MW-0.20,0.16, {fontSize:7,color:"888888"});
  const hCols=[{label:"PLANEJADO",val:`${D.planejado}h`,sub:"original estimate",cor:"222222"},{label:"JA GASTO",val:`${D.jaGasto}h`,sub:"completed work",cor:"D97706"},{label:"FALTA",val:`${D.falta}h`,sub:"remaining (devs)",cor:"F87171"},{label:"AUMENTO DE ESFORCO",val:D.aumentoH,sub:`projetado: ${D.proj_total}h (${D.aumentoP})`,cor:"D97706"}];
  const hW=MW/4;
  hCols.forEach((h,hi) => {
    const hx=MX+hi*hW;
    if (hi>0) fillRect(hx,S2Y+0.22,0.01,S2H-0.28,"E8EDF2");
    addText(h.label, hx+0.10,S2Y+0.22,hW-0.20,0.16, {fontSize:7,color:"888888",bold:true});
    addText(h.val,   hx+0.10,S2Y+0.36,hW-0.20,0.34, {fontSize:24,color:h.cor,bold:true});
    addText(h.sub,   hx+0.10,S2Y+0.72,hW-0.20,0.14, {fontSize:7,color:"888888"});
  });
  const pbGastoW=MW*(D.jaGasto/D.proj_total), pbFaltaW=MW*(D.falta/D.proj_total);
  const pbgY=S2Y+S2H-0.07;
  fillRect(MX,pbgY,pbGastoW,0.06,"D97706");
  fillRect(MX+pbGastoW,pbgY,pbFaltaW,0.06,"7C3AED");

  // ── SEÇÃO 2B — PREVISÃO ──
  const S2BY=S2Y+S2H+0.03, S2BH=0.96, halfW=(MW-0.10)/2;

  // E1
  const pv1X=MX, pv1W=halfW;
  fillAndStroke(pv1X,S2BY,pv1W,S2BH,"FFFFFF","DDDDDD");
  addText("Entrega 1 · Resiliencia Sistemica", pv1X+0.12,S2BY+0.08,pv1W-0.84,0.18, {fontSize:8,color:"1F2937",bold:true});
  fillRect(pv1X+pv1W-0.72,S2BY+0.07,0.62,0.20,"E1F5EE");
  addText(`${D._e1Pct}% concluido`, pv1X+pv1W-0.72,S2BY+0.07,0.62,0.20, {fontSize:6.5,color:"0F6E56",bold:true,align:"center",valign:"middle"});
  addBar(pv1X+0.12,S2BY+0.31,pv1W-0.24,0.07,"E1F5EE","1D9E75",D._e1Pct);
  const pv1ColW=(pv1W-0.24)/3;
  [{label:"PROJETADO",cor:"1D9E75",val:D.data_proj_camadas},{label:"BASE",cor:"D97706",val:D.data_base},{label:"PESSIMISTA",cor:"DC2626",val:D.data_pess}].forEach((pv,pi) => {
    const bx=pv1X+0.12+pi*pv1ColW;
    fillRect(bx,S2BY+0.44,pv1ColW-0.04,0.22,pv.cor);
    addText(pv.label, bx,S2BY+0.44,pv1ColW-0.04,0.22, {fontSize:6.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
    addText(pv.val,   bx,S2BY+0.68,pv1ColW-0.04,0.20, {fontSize:10.5,color:pv.cor,bold:true,align:"center"});
  });

  // E2
  const pv2X=MX+halfW+0.10, pv2W=halfW;
  fillAndStroke(pv2X,S2BY,pv2W,S2BH,"FFFFFF","DDDDDD");
  addText("Entrega 2 · Observabilidade", pv2X+0.12,S2BY+0.08,pv2W-0.84,0.18, {fontSize:8,color:"1F2937",bold:true});
  fillRect(pv2X+pv2W-0.72,S2BY+0.07,0.62,0.20,"FEE2E2");
  addText(`${D._e2Pct}% concluido`, pv2X+pv2W-0.72,S2BY+0.07,0.62,0.20, {fontSize:6.5,color:"991B1B",bold:true,align:"center",valign:"middle"});
  addBar(pv2X+0.12,S2BY+0.31,pv2W-0.24,0.07,"E6F1FB","1E6FA8",D._e2Pct);
  const pv2ColW=(pv2W-0.24)/3;
  [{label:"PROJETADO",cor:"1D9E75",val:D.data_proj},{label:"BASE",cor:"D97706",val:D.data_base_obs},{label:"PESSIMISTA",cor:"DC2626",val:D.data_pess_obs}].forEach((pv,pi) => {
    const bx=pv2X+0.12+pi*pv2ColW;
    fillRect(bx,S2BY+0.44,pv2ColW-0.04,0.22,pv.cor);
    addText(pv.label, bx,S2BY+0.44,pv2ColW-0.04,0.22, {fontSize:6.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
    addText(pv.val,   bx,S2BY+0.68,pv2ColW-0.04,0.20, {fontSize:10.5,color:pv.cor,bold:true,align:"center"});
  });

  // ── SEÇÃO 3 — EFICIÊNCIA ──
  const S3Y=S2BY+S2BH+0.03, S3H=0.90;
  fillAndStroke(MX,S3Y,MW,S3H,"FFFFFF","DDDDDD");
  addText("EFICIENCIA DA ENTREGA", MX+0.10,S3Y+0.06,MW-0.20,0.16, {fontSize:7,color:"888888"});
  const efCols=[{val:D.ef_estourados,label:"Itens Estourados",sub:"das subs com estimativa",cor:"D85A30"},{val:D.ef_retrabalho_pct,label:"Retrabalho",sub:`${D.ef_retrabalho_bugs} Sub Bug . ${D.ef_retrabalho_h} realizadas`,cor:"534AB7"},{val:D.ef_dentro,label:"Dentro da Estimativa",sub:"entregues dentro do planejado",cor:"1D9E75"}];
  const efColW=MW/3;
  efCols.forEach((ef,ei) => {
    const efx=MX+ei*efColW;
    if (ei>0) fillRect(efx,S3Y+0.16,0.01,S3H-0.22,"E8EDF2");
    addText(ef.val,   efx+0.10,S3Y+0.16,efColW-0.20,0.36, {fontSize:22,color:ef.cor,bold:true,align:"center"});
    addText(ef.label, efx+0.10,S3Y+0.52,efColW-0.20,0.18, {fontSize:8,color:"222222",bold:true,align:"center"});
    addText(ef.sub,   efx+0.10,S3Y+0.70,efColW-0.20,0.18, {fontSize:7,color:"888888",align:"center"});
  });

  // ── SEÇÃO 4 — RISCOS ──
  const S4Y=S3Y+S3H+0.15, S4H=0.96, RW=(MW-0.09*3)/4;
  addText("R I S C O S", MX,S4Y-0.13,MW,0.12, {fontSize:6.5,color:"888888"});
  D.riscos.forEach((r,ri) => {
    const rx=MX+ri*(RW+0.09);
    fillAndStroke(rx,S4Y,RW,S4H,r.bgCor,r.bordaCor,1.5);
    addText(r.tag,    rx+0.08,S4Y+0.07,RW-0.65,0.20, {fontSize:7,color:r.cor,bold:true});
    fillRect(rx+RW-0.58,S4Y+0.06,0.52,0.32,r.badgeCor);
    addText(r.badge,  rx+RW-0.58,S4Y+0.06,0.52,0.32, {fontSize:6.5,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
    addText(r.linha1, rx+0.08,S4Y+0.34,RW-0.16,0.20, {fontSize:7.5,color:"222222"});
    addText(r.linha2, rx+0.08,S4Y+0.54,RW-0.16,0.18, {fontSize:7.5,color:"222222"});
  });

  // ── SEÇÃO 5 — TIME ──
  const S5Y=S4Y+S4H+0.05, S5H=7.5-S5Y-0.04;
  const nDevs=D.devs.length, nQA=D.qa.length, totalCards=nDevs+nQA;
  const TW=(MW-0.09*(totalCards-1))/totalCards, cardH=S5H-0.09;
  const devsFaixaW=nDevs*TW+(nDevs-1)*0.09;
  fillRect(MX,S5Y,devsFaixaW,0.08,"374151");
  addText("D E V S", MX,S5Y,devsFaixaW,0.08, {fontSize:7,color:"FFFFFF",bold:true,align:"center",valign:"middle"});
  const qaStartX=MX+nDevs*(TW+0.09), qaFaixaW=nQA*TW+(nQA-1)*0.09;
  fillRect(qaStartX,S5Y,qaFaixaW,0.08,"1E6FA8");
  addText("Q A", qaStartX,S5Y,qaFaixaW,0.08, {fontSize:7,color:"FFFFFF",bold:true,align:"center",valign:"middle"});

  [...D.devs,...D.qa].forEach((dev,di) => {
    const isQA = di >= nDevs;
    const tx = isQA ? qaStartX+(di-nDevs)*(TW+0.09) : MX+di*(TW+0.09);
    const cardY = S5Y+0.09;
    fillAndStroke(tx,cardY,TW,cardH,"FFFFFF","DDDDDD");
    addText(dev.nome, tx+0.06,cardY+0.08,TW-0.12,0.20, {fontSize:8.5,color:"222222",bold:true,align:"center"});
    const bpY=cardY+0.30, bpW=TW-0.12;
    addBar(tx+0.06,bpY,bpW,0.08,"E8EDF2",dev.cor,dev.est>0?dev.real/dev.est*100:0);
    const subsLabel = dev.subs>0 ? `${dev.real}h · ${dev.subs} subs finalizadas` : `${dev.real}h realizadas`;
    addText(subsLabel, tx+0.06,bpY+0.10,bpW,0.24, {fontSize:9.5,color:dev.cor,bold:true,align:"center"});
    if (dev.est>0) addText(`${dev.est}h estimadas`, tx+0.06,bpY+0.34,bpW,0.18, {fontSize:8.5,color:"888888",align:"center"});
  });

  // Salvar
  const buf = canvas.toBuffer("image/jpeg", { quality: 0.95 });
  fs.writeFileSync(OUTPUT_JPG, buf);
  console.log(`✅ JPG gerado: ${OUTPUT_JPG}`);
}


// ── MAIN ───────────────────────────────────────────────────────────────
async function main() {
  if (PAT === "SEU_PAT_AQUI") {
    console.error("❌ Configure o PAT no topo do script antes de executar!");
    process.exit(1);
  }

  console.log("🔄 Buscando IDs da query no Azure DevOps...");
  const ids = await fetchQueryIds();
  console.log(`   → ${ids.length} IDs únicos encontrados`);

  console.log("🔄 Buscando campos dos work items (pode demorar ~5s)...");
  const items = await fetchItemsBatch(ids);
  console.log(`   → ${items.length} itens carregados`);

  console.log("🔄 Calculando métricas...");
  const M = calcular(items);

  // Exibir resumo no terminal
  exibirResumo(M);

  // Perguntar se deve gerar
  const readline = require("readline");
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  rl.question("\n👉 Deseja gerar o report PPTX? (s/n): ", async (resp) => {
    rl.close();
    if (resp.trim().toLowerCase() === "s") {
      console.log("🔄 Gerando PPTX...");
      await gerarPPTX(M);
      console.log("🔄 Gerando JPG...");
      await gerarJPG(M);
    } else {
      console.log("⏭️  Geração cancelada.");
    }
  });
}

main().catch(e => { console.error("❌ Erro:", e.message); process.exit(1); });