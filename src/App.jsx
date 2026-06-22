import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from "recharts";
import { AlertTriangle, CheckCircle, Upload, BarChart2, Home, BookOpen, Shield, X, Eye, Search, Zap, Download, AlertCircle, Database, RefreshCw, CheckSquare, FileSpreadsheet, Menu, ChevronDown, ChevronUp, FileText, Mail, Plus, Trash2, Copy, FileCheck, Clipboard } from "lucide-react";

// ── Brand ─────────────────────────────────────────────────────────────────
const C = { navy:"#02124C", blue:"#3E75FF", light:"#E0E6FC", green:"#2CDD7C",
  orange:"#FF8E4F", white:"#FFFFFF", grey:"#536078", body:"#1D2433", bg:"#F0F4FF",
  border:"rgba(62,117,255,0.12)", purple:"#7C3AED", cyan:"#0891B2", amber:"#F59E0B",
  teal:"#0F766E", red:"#EF4444" };
const font = "'IBM Plex Sans',system-ui,sans-serif";

// ── Responsive hook ────────────────────────────────────────────────────────
const useResponsive = () => {
  const [w, setW] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(() => {
    const h = () => setW(window.innerWidth);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return { isMobile: w < 640, isTablet: w >= 640 && w < 1024, isDesktop: w >= 1024, width: w };
};

// ── Normas SUSEP (Danos + Grandes Riscos — sem Ouvidoria e sem Saúde/Vida) ─
const NORMAS = [
  { id:0, ref:"Lei 15.040/2024", cor:C.teal, titulo:"Marco Legal dos Seguros",
    resumo:"Aplica-se a contratos celebrados a partir de 11/12/2025. Art. 86: 30 dias do aviso para cobertura (sunset clause). Art. 87: 30 dias para pagar. Art. 88: multa 2% por atraso.",
    pS:30, pC:30, vigente:true, dataVig:"11/12/2025",
    obr:["Art. 74 — Ônus da prova de não-cobertura é da seguradora",
      "Art. 76 — Regulação e liquidação simultâneas sempre que possível",
      "Art. 78 — Dúvidas sobre cálculos resolvidas em favor do segurado",
      "Art. 83 — Relatório de regulação com conteúdo mínimo (10 dias se solicitado)",
      "Art. 86 — 30 dias do aviso → SUNSET CLAUSE: silêncio = decaimento do direito de recusar",
      "Art. 87 — 30 dias após cobertura reconhecida para pagar",
      "Art. 88 — Multa 2% + juros + correção por atraso injustificado"],
    pen:"Multa 2% + juros + correção (Art. 88). Sunset clause: perda do direito de recusar (Art. 86)",
    nota:"⚠️ Regulamentação SUSEP em elaboração (CP 10/2025). Resolução prevista jul/2026." },
  { id:1, ref:"Circ. SUSEP 621/2021", cor:C.blue, titulo:"Prazos de Regulação — Seguros de Danos",
    resumo:"Art. 43: prazo máximo de 30 dias para liquidação de sinistros de danos, contados da entrega dos documentos básicos. Norma principal para contratos anteriores a 11/12/2025.",
    pS:30, pC:null, vigente:true, dataVig:"12/02/2021",
    obr:["Art. 41 — Informar ao segurado os documentos básicos exigidos por cobertura",
      "Art. 42 — Vedado fixar prazo máximo para comunicação do sinistro",
      "Art. 43 — Prazo máximo de 30 dias para liquidação, contados da entrega dos documentos",
      "Art. 44 — Documentação complementar suspende o prazo, com justificativa expressa",
      "Vedado exigir documentos desnecessários ou não previstos nas condições contratuais"],
    pen:"Penalidades administrativas conforme Decreto-Lei 73/1966",
    nota:"ℹ️ Norma correta para prazos de sinistros de danos." },
  { id:2, ref:"Res. CNSP 407/2021", cor:C.orange, titulo:"Grandes Riscos — Contratos Complexos",
    resumo:"Regulamenta contratos de seguros de grandes riscos (LMG > R$ 15 mi OU ativo > R$ 27 mi OU faturamento > R$ 57 mi). Prazos e condições pactuados entre as partes.",
    pS:null, pC:null, vigente:true, dataVig:"01/04/2021",
    obr:["Critérios: LMG > R$ 15 mi OU ativo total > R$ 27 mi OU faturamento bruto > R$ 57 mi",
      "Prazos de regulação definidos nas condições contratuais (Lei 15.040 prevê até 120d)",
      "Maior liberdade de negociação e estruturação de cláusulas entre as partes",
      "Guarda de documentos à disposição da SUSEP",
      "Possibilidade de cláusulas de arbitragem"],
    pen:"Regras disciplinadas pelo contrato",
    nota:"ℹ️ A Lei 15.040/2024 (Art. 86 §6º) prevê prazo de até 120 dias para grandes riscos." },
];

// ── Date constants ─────────────────────────────────────────────────────────
// Datas usam construtor local (ano, mês-1, dia) para evitar problema de UTC vs fuso local
const CNSP_415_DATE  = new Date(2021, 6, 1);   // 01/07/2021
const CNSP_460_DATE  = new Date(2023, 0, 1);   // 01/01/2023
const LEI_15040_DATE = new Date(2025, 11, 11); // 11/12/2025 inclusive

// ── Status com documentação pendente do SEGURADO (prazo suspenso — Art. 44 Circ. 621/2021)
// Enquanto o segurado não entrega os documentos, o prazo legal NÃO corre.
// Esses casos nunca devem ser classificados como "Vencido".
const STATUS_DOC_PENDENTE_SEGURADO = new Set([
  "P.DOC SEGURADO","P.VIST SEGURADO","AGUARD. RECL 3º","AGUARD. RECL 3°",
  "P.DOC TERCEIRO","P.VIST TERCEIRO","PENDENTE DOC",
]);

// ── Status mapping ─────────────────────────────────────────────────────────
// STATUS_MAP expandido conforme Obs.2 do relatório de João (01/06/2026):
// P. SEGURADORA, P.DOC SEGURADO, P.DOC TERCEIRO, AGUARD. RECL 3º etc. = Em Regulação
// Encerrado inclui todas as variações sem indenização
const STATUS_MAP = {
  "EM REGULAÇÃO":"Em Regulação","PENDENTE - PRAZO NORMAL":"Em Regulação","P. REGULADOR":"Em Regulação",
  "P. SEGURADORA":"Em Regulação","P. REGULADOR EXTERNO":"Em Regulação","PENDENTE SEGURADORA":"Em Regulação",
  // Ag. Documentação — pendências de doc do Segurado ou Terceiro (prazo suspenso)
  "P.DOC SEGURADO":"Ag. Documentação","P.VIST SEGURADO":"Ag. Documentação",
  "P.DOC TERCEIRO":"Ag. Documentação","P.VIST TERCEIRO":"Ag. Documentação",
  "AGUARD. RECL 3º":"Ag. Documentação","AGUARD. RECL 3°":"Ag. Documentação","PENDENTE DOC":"Ag. Documentação",
  "LIQUIDADO":"Liquidado","LIQUIDADO PARCIAL":"Liquidado",
  "LITIGIO":"Em Litígio","LITÍGIO":"Em Litígio","LITIGIO D&O":"Em Litígio","LITÍGIO D&O":"Em Litígio",
  "ENCERRADO SEM INDENIZAÇÃO":"Encerrado","ENCERRADO S/INDENIZ.":"Encerrado",
  "ENC. ABAIXO FRANQUIA":"Encerrado","ENC. S/IND AG. RECLAM":"Encerrado","ENC. S/IND AG. RECLAM.":"Encerrado",
  "ENC. ABX FQ/ AG.RECLAM":"Encerrado","ENC. ABX FQ/ AG. RECLAM.":"Encerrado","ENCER.S/ PREJUIZO":"Encerrado",
  "ENC. S/ PREJUÍZO":"Encerrado","ENCERRADO S/ INDENIZ.":"Encerrado","ENCERRADO S/INDENIZ":"Encerrado","ENCERRADO":"Encerrado",
  "RECUSADO / DECLINADO":"Recusado","RECUADO/DECLINADO":"Recusado","RECUSADO":"Recusado","DECLINADO":"Recusado",
  "EXPECTATIVA/SINISTRO":"Expectativa","EXPECTATIVA":"Expectativa",
};
const mapStatus = raw => STATUS_MAP[(raw||"").toUpperCase().trim()] || raw || "Em Regulação";

// ── Color helpers ──────────────────────────────────────────────────────────
const statC = s => ({"Em Regulação":C.blue,"Liquidado":C.green,"Em Litígio":C.purple,
  "Ag. Documentação":C.orange,"Encerrado":C.grey,"Recusado":"#EF4444",
  "Expectativa":C.cyan,"P. Seguradora":C.amber}[s]||C.grey);
const riskC = r => ({Vencido:"#EF4444","Vencido sob Lei 15.040":"#B91C1C",Alto:C.orange,Médio:C.amber,Baixo:C.green,
  Pendente:C.grey,Litígio:C.purple,Expectativa:C.cyan,Resolvido:C.green,
  "Anterior à Norma":C.cyan,"Em Regulação":C.blue,"Sunset — Decaimento":"#B91C1C"}[r]||C.grey);
const riskBg= r => ({Vencido:"#FEF2F2","Vencido sob Lei 15.040":"#FEF2F2",Alto:"#FFF7ED",Médio:"#FFFBEB",Baixo:"#F0FDF4",
  Pendente:"#F9FAFB",Litígio:"#F5F3FF",Expectativa:"#ECFEFF",Resolvido:"#F0FDF4",
  "Anterior à Norma":"#ECFEFF","Em Regulação":"#EFF6FF","Sunset — Decaimento":"#FEF2F2"}[r]||"#F9FAFB");

// ── Product classification ─────────────────────────────────────────────────
const COMPLEXOS = ["OPERADOR PORTUARIO","OPERADOR PORTUÁRIO","D&O","RC D&O","E&O","RC E&O",
  "RISCOS NOMEADOS","RISCOS OPERACIONAIS","RISCOS AMBIENTAIS","RESPONSABILIDADE CIV",
  "RESP CIVIL","RESP. CIVIL GERAL","EMPRESARIAL","BENFEITORIAS E AGROP","AERONAUTICO CASCO"];

// ── Date helpers ───────────────────────────────────────────────────────────
const formatDate = val => {
  if (!val && val !== 0) return "";
  if (val instanceof Date) { if (isNaN(val)) return ""; return `${String(val.getDate()).padStart(2,"0")}/${String(val.getMonth()+1).padStart(2,"0")}/${val.getFullYear()}`; }
  if (typeof val === "number" && val > 0) return formatDate(new Date((val-25569)*86400000));
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(String(val))) return String(val);
  return "";
};
const parseBR = s => {
  if (!s) return null;
  if (s instanceof Date) return isNaN(s) ? null : s;
  const [d,m,y] = String(s).split("/");
  return (!d||!m||!y) ? null : new Date(`${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`);
};

// ── Enrich sinistro ────────────────────────────────────────────────────────
const enrichRelatorio = raw => {
  const segurado = raw["SEGURADO"]||"", seguradora = raw["SEGURADORA"]||"";
  const id = String(raw["Nº AVISO"]||raw["N° AVISO"]||"").trim()||`SIN-${Math.random().toString(36).slice(2,7).toUpperCase()}`;
  const tipo = raw["PRODUTO"]||"", ramo = raw["RAMO"]||"";
  const regulador = raw["ANALISTA RESPONSÁVEL"]||raw["ANALISTA RESPONSAVEL"]||"";
  const statusRaw = raw["SITUAÇÃO DO SINISTRO"]||raw["SITUACAO DO SINISTRO"]||"";
  const importanciaSegurada = Number(raw["IMP. SEG."])||0;
  const apurado = Number(raw["APURADO"])||0, franquia = Number(raw["FRANQUIA"])||0;
  const indenizado = Number(raw["INDENIZADO"])||0;
  const dataAbertura = formatDate(raw["DATA AVISO"]);
  const dataDocCompleta = formatDate(raw["DATA INCLUSÃO"]||raw["DATA INCLUSAO"]);
  const dataSinistro = formatDate(raw["DATA SINISTRO"]);
  const inicioVig = formatDate(raw["INÍCIO DE VIGÊNCIA"]||raw["INICIO DE VIGENCIA"]);
  const terminoVig = formatDate(raw["TÉRMINO DE VIGÊNCIA"]||raw["TERMINO DE VIGENCIA"]);
  const apolice = raw["APÓLICE"]||raw["APOLICE"]||"";
  const nAviso = String(raw["Nº AVISO"]||raw["N° AVISO"]||"");
  const valorEstimado = apurado > 0 ? apurado : importanciaSegurada;
  const status = mapStatus(statusRaw);
  const isComplexo = COMPLEXOS.some(cc => (tipo||"").toUpperCase().includes(cc));
  const prazoLegal = isComplexo ? 60 : 30;
  const today = new Date();
  const inicioVigDate = parseBR(inicioVig);
  // Norma aplicável é determinada pela VIGÊNCIA DA APÓLICE (início de vigência),
  // não pela data do aviso — conforme Lei 15.040/2024 e normas SUSEP anteriores.
  // Normaliza para comparação apenas de data (sem hora) — garante que 11/12/2025 seja incluído
  const inicioVigNorm = inicioVigDate ? new Date(inicioVigDate.getFullYear(), inicioVigDate.getMonth(), inicioVigDate.getDate()) : null;
  const sobreLei15040 = !!inicioVigNorm && inicioVigNorm >= LEI_15040_DATE;
  const dataAvisoDate = parseBR(dataAbertura);
  const normaVigencia = (tipo||"").toUpperCase().includes("GARANTIA") ? CNSP_460_DATE : CNSP_415_DATE;
  const anteriorNorma = !sobreLei15040 && !!dataAvisoDate && dataAvisoDate < normaVigencia;
  const normaVigStr = sobreLei15040 ? "Lei 15.040/2024 (11/12/2025)" :
    normaVigencia === CNSP_460_DATE ? "Res. CNSP 407/2021 (01/04/2021)" : "Circ. SUSEP 621/2021 (12/02/2021)";
  // emRegulacao: inclui "Ag. Documentação" pois esses casos ainda estão em processo ativo
  const emRegulacao = status === "Em Regulação" || status === "Ag. Documentação";

  // Verifica se o status original indica documentação PENDENTE DO SEGURADO.
  // Nesses casos o prazo legal é SUSPENSO (Art. 44 Circ. SUSEP 621/2021 /
  // Art. 86 §3º Lei 15.040/2024): o caso NUNCA pode ser classificado como "Vencido".
  const statusOriginalNorm = (statusRaw||"").toUpperCase().trim();
  const docPendenteSegurado = STATUS_DOC_PENDENTE_SEGURADO.has(statusOriginalNorm);

  // DATA DE REFERÊNCIA DO PRAZO — REGRA DEFINITIVA (João, 09/06/2026):
  // • Apólices anteriores a 11/12/2025 → classificadas como "Anterior à Norma" SEM EXCEÇÃO.
  //   Não há cálculo de prazo, não há Vencido, não há Sunset para essas apólices.
  // • Apólices >= 11/12/2025 (Lei 15.040/2024):
  //   - Doc pendente do segurado → prazo SUSPENSO (Art. 86 §3º) → risco = Pendente
  //   - Prazo conta a partir da DATA AVISO (pois a lei exige que os prazos se iniciem
  //     quando as informações e documentações sejam suficientes — Art. 86 Lei 15.040/2024)
  //   - Se prazo expirou (> 30d): risco = "Vencido sob Lei 15.040" (engloba Sunset)
  //   - Se prazo entre 1-30d: risco proporcional (Alto/Médio/Baixo)
  const dataRefPrazo = (!sobreLei15040 || docPendenteSegurado) ? null : parseBR(dataAbertura);
  const temDoc = !!dataRefPrazo;
  const diasRef = temDoc ? Math.floor((today - dataRefPrazo)/86400000) : 0;
  const prazoR = temDoc ? prazoLegal - diasRef : null;
  // sunsetAlert: mantido internamente para exibir aviso adicional no detalhe do caso
  const sunsetAlert = sobreLei15040 && emRegulacao && !docPendenteSegurado && temDoc && diasRef >= 30;

  let risco = "Resolvido";
  if (status==="Em Litígio") risco="Litígio";
  else if (status==="Expectativa") risco="Expectativa";
  else if (["Encerrado","Liquidado","Recusado"].includes(status)) risco="Resolvido";
  else if (emRegulacao) {
    // Apólice anterior a 11/12/2025 → SEMPRE "Anterior à Norma" (João, 09/06/2026)
    if (!sobreLei15040) risco="Anterior à Norma";
    // Apólice sob Lei 15.040/2024 (>= 11/12/2025)
    else if (docPendenteSegurado) risco="Pendente"; // prazo suspenso — doc pendente
    else if (!temDoc) risco="Pendente";
    // Prazo expirado (> 30d): "Vencido sob Lei 15.040" — engloba Sunset (João, 09/06/2026)
    else if (prazoR<0) risco="Vencido sob Lei 15.040";
    else if (prazoR<=7) risco="Alto";
    else if (prazoR<=15) risco="Médio";
    else risco="Baixo";
  }
  return { id, segurado, seguradora, apolice, tipo, ramo, nAviso, regulador,
    dataAbertura, dataSinistro, dataDocCompleta, inicioVig, terminoVig,
    importanciaSegurada, apurado, franquia, indenizado, valorEstimado,
    statusOriginal:statusRaw, status, classificacao:isComplexo?"Complexo":"Simples",
    prazoLegal, prazoRestante:prazoR!==null?Math.round(prazoR):null,
    prazoArt86: sobreLei15040&&temDoc ? 30-diasRef : null,
    risco, temDocCompleta:temDoc,
    norma: sobreLei15040 ? "Lei 15.040/2024" : (isComplexo?"Res. CNSP 407/2021":"Circ. SUSEP 621/2021"),
    anteriorNorma, normaVigStr, sobreLei15040, sunsetAlert };
};

// ── Parse Excel/CSV ────────────────────────────────────────────────────────
const parseFile = (file, onOk, onErr) => {
  const r = new FileReader();
  r.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:"binary",cellDates:true});
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, {defval:""});
      if (!rows.length) { onErr("Planilha sem dados."); return; }
      const keys = Object.keys(rows[0]);
      const isRelatorio = keys.some(k=>k.includes("SITUAÇÃO DO SINISTRO")||k.includes("SITUACAO DO SINISTRO")||k==="SEGURADO");
      if (!isRelatorio) { onErr("Formato não reconhecido. Use o RELATÓRIO_SINISTROS ou baixe o template."); return; }
      const enriched = rows.filter(r=>r["SEGURADO"]||r["Nº SINISTRO"]).map(enrichRelatorio);
      onOk(enriched, rows.length);
    } catch(err) { onErr("Erro: "+err.message); }
  };
  r.onerror = () => onErr("Falha ao ler o arquivo.");
  r.readAsBinaryString(file);
};

// ── Parse PDF text via pdfjs-dist ─────────────────────────────────────────
const extractPdfText = async (file) => {
  const pdfjsLib = await import('pdfjs-dist');
  pdfjsLib.GlobalWorkerOptions.workerSrc = '/pdf.worker.min.mjs';
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
  let fullText = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    fullText += content.items.map(item => item.str).join(' ') + '\n';
  }
  return fullText.trim();
};

// ── Download templates ─────────────────────────────────────────────────────
const downloadTemplate = () => {
  const cols=["Nº SINISTRO","SEGURADO","SEGURADORA","APÓLICE","PRODUTO","RAMO","Nº AVISO",
    "DATA SINISTRO","DATA AVISO","DATA INCLUSÃO","INÍCIO DE VIGÊNCIA","TÉRMINO DE VIGÊNCIA",
    "PRÊMIO PAGO","IMP. SEG.","APURADO","FRANQUIA","INDENIZADO","VALOR RECUPERADO",
    "ANALISTA RESPONSÁVEL","SITUAÇÃO DO SINISTRO"];
  const ws = XLSX.utils.aoa_to_sheet([cols]);
  ws["!cols"] = [14,28,26,16,22,30,10,12,12,12,14,14,12,12,12,10,12,14,18,20].map(w=>({wch:w}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "RptSinistrosAvisados");
  XLSX.writeFile(wb,"Umma_Template_Sinistros.xlsx");
};

// downloadTemplateApolice removed — apólices are now uploaded as PDF

// ── Helpers ────────────────────────────────────────────────────────────────
const fCur = v => new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL",maximumFractionDigits:0}).format(Number(v)||0);
const bdg = (c,bg) => ({display:"inline-flex",alignItems:"center",padding:"2px 8px",borderRadius:20,fontSize:10.5,fontWeight:600,color:c,background:bg||c+"18",whiteSpace:"nowrap"});
const card = {background:C.white,borderRadius:12,padding:16,border:`1px solid ${C.border}`,boxShadow:"0 1px 4px rgba(2,18,76,0.06)"};
const btn = (col,sm) => ({display:"flex",alignItems:"center",gap:6,padding:sm?"7px 12px":"9px 16px",borderRadius:8,background:col||C.blue,color:col===C.white||col===C.light?"#333":C.white,fontFamily:font,fontWeight:600,fontSize:sm?12:13,border:col===C.white?`1px solid ${C.border}`:"none",cursor:"pointer"});

// ══════════════════════════════════════════════════════════════════════════════
export default function App() {
  const { isMobile } = useResponsive();

  // ── Core state ─────────────────────────────────────────────────────────────
  const [sec,    setSec]   = useState("aviso");
  const [casos,  setCasos] = useState([]);
  const [loading,setLoading]=useState(true);
  const [srch,   setSrch]  = useState("");
  const [fSt,    setFSt]   = useState("Todos");
  const [fRk,    setFRk]   = useState("Todos");
  const [fSeg,   setFSeg]  = useState("Todos");
  const [selC,   setSelC]  = useState(null);
  const [selN,   setSelN]  = useState(null);
  const [drag,   setDrag]  = useState(false);
  const [imp,    setImp]   = useState(false);
  const [log,    setLog]   = useState([]);
  const [cnt,    setCnt]   = useState(0);
  const [sideOpen,setSideOpen]=useState(true);
  const [showFilters,setShowFilters]=useState(false);
  const fRef = useRef();

  // ── AI (análise sinistro) state ────────────────────────────────────────────
  const [aiTxt,  setAiTxt] = useState("");
  const [aiRes,  setAiRes] = useState(null);
  const [aiLd,   setAiLd]  = useState(false);
  const [aiErr,  setAiErr] = useState("");

  // ── Aviso de sinistro state ────────────────────────────────────────────────
  const [avisoTxt,  setAvisoTxt]  = useState("");
  const [avisoRes,  setAvisoRes]  = useState(null);
  const [avisoLd,   setAvisoLd]   = useState(false);
  const [avisoErr,  setAvisoErr]  = useState("");
  const [avisoEmail,setAvisoEmail]= useState("");
  const [copiedEmail,setCopiedEmail]=useState(false);

  // ── Apólices state (PDF-based) ────────────────────────────────────────────
  const [apolices,   setApolices]   = useState([]); // [{nome, seguradora, texto, tamanho, data}]
  const [dragAp,     setDragAp]     = useState(false);
  const [impAp,      setImpAp]      = useState(false);
  const [logAp,      setLogAp]      = useState([]);
  const [cntAp,      setCntAp]      = useState(0);
  const [srchAp,     setSrchAp]     = useState("");
  const [selAp,      setSelAp]      = useState(null);
  const [segAp,      setSegAp]      = useState("");
  const fRefAp = useRef();

  // ── Instruções state (PDF-based) ──────────────────────────────────────────
  const [instrucoes,  setInstrucoes]  = useState([]); // [{nome, seguradora, texto, tamanho, data}]
  const [dragInst,    setDragInst]    = useState(false);
  const [impInst,     setImpInst]     = useState(false);
  const [logInst,     setLogInst]     = useState([]);
  const [srchInst,    setSrchInst]    = useState("");
  const [selInst,     setSelInst]     = useState(null);
  const [segInst,     setSegInst]     = useState("");
  const fRefInst = useRef();

  // ── Load from localStorage ─────────────────────────────────────────────────
  useEffect(() => {
    try {
      const raw = localStorage.getItem("umma-casos");
      if (raw) { const p=JSON.parse(raw); if(Array.isArray(p)&&p.length>0){setCasos(p);setCnt(p.length);} }
    } catch(e) {}
    try {
      const raw = localStorage.getItem("umma-apolices");
      if (raw) { const p=JSON.parse(raw); if(Array.isArray(p)&&p.length>0){setApolices(p);setCntAp(p.length);} }
    } catch(e) {}
    try {
      const raw = localStorage.getItem("umma-instrucoes");
      if (raw) { const p=JSON.parse(raw); if(Array.isArray(p)&&p.length>0){setInstrucoes(p);} }
    } catch(e) {}
    setLoading(false);
  }, []);

  const saveApolices = (data) => { setApolices(data); setCntAp(data.length); try{localStorage.setItem("umma-apolices",JSON.stringify(data));}catch(e){} };
  const saveInstrucoes = (data) => { setInstrucoes(data); try{localStorage.setItem("umma-instrucoes",JSON.stringify(data));}catch(e){} };

  // ── Metrics ────────────────────────────────────────────────────────────────
  const total    = casos.length;
  // emReg: SOMENTE status literal "EM REGULAÇÃO" (Obs.2 João 05/06/2026 — são 51 casos)
  const emReg    = casos.filter(c=>c.statusOriginal&&(c.statusOriginal||"").toUpperCase().trim()==="EM REGULAÇÃO").length;
  const litigio  = casos.filter(c=>c.status==="Em Litígio").length;
  const liquidado= casos.filter(c=>c.status==="Liquidado").length;
  // agDoc: P.DOC SEGURADO, P.DOC TERCEIRO, P.VIST SEGURADO, P.VIST TERCEIRO, AGUARD. RECL 3º
  const agDoc    = casos.filter(c=>c.status==="Ag. Documentação").length;
  const vencidos = casos.filter(c=>c.risco==="Vencido sob Lei 15.040").length;
  const altoRisco= casos.filter(c=>c.risco==="Alto").length;
  // encerrado: todas as variações de encerramento sem indenização (Obs.4 João)
  const encerrado= casos.filter(c=>c.status==="Encerrado").length;
  // recusado: RECUSADO/DECLINADO também é findado (Obs.4 João)
  const recusado = casos.filter(c=>c.status==="Recusado").length;
  // findados: liquidado + encerrado + recusado = 535 + 284 + ... = 819
  const findados = liquidado + encerrado + recusado;
  const lei15040Count = casos.filter(c=>c.sobreLei15040).length;
  // sunsetAlert ainda existe internamente; o risco "Vencido sob Lei 15.040" engloba sunset
  const sunsetCount   = casos.filter(c=>c.sunsetAlert).length;
  const valTotal = casos.reduce((s,c)=>s+(c.importanciaSegurada||0),0);
  const valApurado=casos.reduce((s,c)=>s+(c.apurado||0),0);
  const valIndeni= casos.reduce((s,c)=>s+(c.indenizado||0),0);
  const crit     = casos.filter(c=>["Vencido sob Lei 15.040","Alto"].includes(c.risco));
  const activeCases = casos.filter(c=>["Em Regulação","Ag. Documentação","P. Seguradora"].includes(c.status));
  const compliance = activeCases.length>0?Math.round(casos.filter(c=>["Baixo","Médio"].includes(c.risco)&&c.temDocCompleta).length/activeCases.length*100):100;

  // ── Charts ─────────────────────────────────────────────────────────────────
  const stChart = Object.entries(casos.reduce((a,c)=>{a[c.status]=(a[c.status]||0)+1;return a},{}))
    .map(([n,v])=>({n,v,c:statC(n)})).sort((a,b)=>b.v-a.v).slice(0,6);
  const rkChart = [{n:"Vencido Lei 15.040",v:vencidos,c:"#B91C1C"},{n:"Alto",v:altoRisco,c:C.orange},
    {n:"Médio",v:casos.filter(c=>c.risco==="Médio").length,c:C.amber},
    {n:"Baixo",v:casos.filter(c=>c.risco==="Baixo").length,c:C.green},
    {n:"Litígio",v:litigio,c:C.purple},{n:"Pendente",v:agDoc,c:C.grey},
    {n:"Anterior Norma",v:casos.filter(c=>c.risco==="Anterior à Norma").length,c:C.cyan}].filter(d=>d.v>0);
  const tipoChart=[...new Set(casos.map(c=>c.tipo))].filter(Boolean)
    .map(t=>({n:t.length>14?t.slice(0,14)+"…":t,v:casos.filter(c=>c.tipo===t).length}))
    .sort((a,b)=>b.v-a.v).slice(0,6);

  const ALL_ST=["Todos","Em Regulação","Ag. Documentação","Em Litígio","Liquidado","Encerrado","Recusado","Expectativa","P. Seguradora"];
  const ALL_RK=["Todos","Vencido sob Lei 15.040","Alto","Médio","Baixo","Litígio","Pendente","Resolvido","Expectativa","Anterior à Norma","Em Regulação"];
  // Lista de segurados únicos para o dropdown (ordenada alfabeticamente)
  const ALL_SEG = ["Todos", ...Array.from(new Set(casos.map(c=>c.segurado).filter(Boolean))).sort()];
  const filtrados = casos.filter(c=>{
    const ms=fSt==="Todos"||c.status===fSt, mr=fRk==="Todos"||c.risco===fRk;
    const mseg=fSeg==="Todos"||c.segurado===fSeg;
    const mb=!srch||[c.id,c.segurado,c.tipo,c.seguradora,c.regulador,c.ramo].some(f=>(f||"").toLowerCase().includes(srch.toLowerCase()));
    return ms&&mr&&mb&&mseg;
  });

  // ── File upload (sinistros) ────────────────────────────────────────────────
  const handleFile = file => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx","xls","csv"].includes(ext)) { setLog([{t:"err",m:`Formato .${ext} não suportado.`}]); return; }
    setImp(true); setLog([{t:"ok",m:`Lendo: ${file.name}...`}]);
    parseFile(file, (enriched, rawCount) => {
      setImp(false);
      const valid = enriched.filter(c=>c.segurado&&c.seguradora);
      const logs=[{t:"ok",m:`✓ ${rawCount} linhas lidas.`}];
      const vn=valid.filter(c=>c.risco==="Vencido sob Lei 15.040").length;
      if(vn>0) logs.push({t:"err",m:`⚠ ${vn} caso(s) VENCIDO(S) sob Lei 15.040/2024 — ação imediata necessária.`});
      const lei=valid.filter(c=>c.sobreLei15040).length;
      if(lei>0) logs.push({t:"ok",m:`⚖️ ${lei} caso(s) regidos pela Lei 15.040/2024.`});
      logs.push({t:"ok",m:`✓ ${valid.length} casos importados com sucesso.`});
      setLog(logs); setCnt(valid.length);
      setCasos(prev=>{
        const ids=new Set(prev.map(c=>c.id));
        return [...prev.map(c=>valid.find(v=>v.id===c.id)||c),...valid.filter(v=>!ids.has(v.id))];
      });
      try { localStorage.setItem("umma-casos", JSON.stringify(valid)); } catch(e) {}
    }, err => { setImp(false); setLog([{t:"err",m:err}]); });
  };

  // ── File upload (apólices PDF) ────────────────────────────────────────────
  const handleFileAp = async (file, seguradoraOverride) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "pdf") { setLogAp([{t:"err",m:`Formato .${ext} não suportado. Envie um arquivo PDF.`}]); return; }
    setImpAp(true); setLogAp([{t:"ok",m:`Lendo PDF: ${file.name}...`}]);
    try {
      const texto = await extractPdfText(file);
      const novaApolice = {
        id: Date.now(),
        nome: file.name,
        seguradora: seguradoraOverride || "",
        texto: texto.slice(0, 12000), // limita para não estourar localStorage
        tamanho: (file.size/1024).toFixed(1)+" KB",
        data: new Date().toLocaleDateString("pt-BR")
      };
      const nova = [...apolices, novaApolice];
      saveApolices(nova);
      setLogAp([{t:"ok",m:`✓ PDF lido com sucesso: ${file.name}`},{t:"ok",m:`✓ ${texto.split(' ').length} palavras extraídas da apólice.`}]);
    } catch(err) {
      setLogAp([{t:"err",m:"Erro ao ler PDF: "+err.message}]);
    }
    setImpAp(false);
  };

  // ── File upload (instruções PDF) ──────────────────────────────────────────
  const handleFileInst = async (file, seguradoraOverride) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "pdf") { setLogInst([{t:"err",m:`Formato .${ext} não suportado. Envie um arquivo PDF.`}]); return; }
    setImpInst(true); setLogInst([{t:"ok",m:`Lendo PDF: ${file.name}...`}]);
    try {
      const texto = await extractPdfText(file);
      const novaInst = {
        id: Date.now(),
        nome: file.name,
        seguradora: seguradoraOverride || "",
        texto: texto.slice(0, 12000),
        tamanho: (file.size/1024).toFixed(1)+" KB",
        data: new Date().toLocaleDateString("pt-BR")
      };
      const nova = [...instrucoes, novaInst];
      saveInstrucoes(nova);
      setLogInst([{t:"ok",m:`✓ PDF lido com sucesso: ${file.name}`},{t:"ok",m:`✓ ${texto.split(' ').length} palavras extraídas do documento.`}]);
    } catch(err) {
      setLogInst([{t:"err",m:"Erro ao ler PDF: "+err.message}]);
    }
    setImpInst(false);
  };

  // ── AI (análise rápida de sinistro) ───────────────────────────────────────
  const runAI = async () => {
    if (!aiTxt.trim()) return;
    setAiLd(true); setAiRes(null); setAiErr("");
    const sys=`Especialista sinistros Brasil. Normas: Circ. SUSEP 621/2021 (30d danos), Res. CNSP 407/2021 (grandes riscos), Lei 15.040/2024 (Art. 86 sunset 30d, Art. 87 pagamento 30d, Art. 88 multa 2%). Retorne SOMENTE JSON: {"classificacao":"Simples"|"Complexo","justificativa":"str","prazoLegal":num,"normas":["arr"],"risco":"Baixo"|"Médio"|"Alto"|"Crítico","documentos":["arr"],"acoes":["arr"],"alertas":["arr"],"recomendacao":"str","score":0-100}`;
    try {
      const r=await fetch("/api/analyze",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({system:sys,messages:[{role:"user",content:"Analise:\n"+aiTxt}]})});
      const d=await r.json(); const m=(d.content?.[0]?.text||"").match(/\{[\s\S]*\}/);
      if(m) setAiRes(JSON.parse(m[0])); else setAiErr("Resposta inválida.");
    } catch(e){ setAiErr("Erro: "+e.message); }
    setAiLd(false);
  };

  // ── Aviso de Sinistro com IA ───────────────────────────────────────────────
  const runAvisoIA = async () => {
    if (!avisoTxt.trim()) return;
    setAvisoLd(true); setAvisoRes(null); setAvisoErr(""); setAvisoEmail("");

    const apolicesCtx = apolices.length > 0
      ? `\n\nAPÓLICES VIGENTES NA BASE (${apolices.length} apólice(s) cadastrada(s)):\n${apolices.slice(0,5).map(a=>
          `--- Apólice: ${a.nome} | Seguradora: ${a.seguradora||"não informada"} | Adicionada em: ${a.data}\nConteúdo:\n${(a.texto||"sem texto").slice(0,2000)}`
        ).join("\n\n")}`
      : "\n\nNota: Nenhuma apólice cadastrada na base. Analise com base nas informações do aviso.";

    const instrucoesCtx = instrucoes.length > 0
      ? `\n\nPROCEDIMENTOS E INSTRUÇÕES OPERACIONAIS (${instrucoes.length} documento(s) carregado(s) — USE ESTES PROCEDIMENTOS ao orientar o cliente):\n${instrucoes.slice(0,8).map(i=>
          `--- Documento: ${i.nome} | Seguradora: ${i.seguradora||"geral"} | Adicionado em: ${i.data}\nConteúdo:\n${(i.texto||"sem texto").slice(0,3000)}`
        ).join("\n\n")}\n\nIMPORTANTE: Ao gerar o email_sugerido e as acoes_imediatas, incorpore os procedimentos acima de forma específica e detalhada para o cliente.`
      : "\n\nNota: Nenhum procedimento operacional cadastrado. Oriente o cliente com base nas normas SUSEP vigentes.";

    const sys = `Você é um especialista em regulação de sinistros da Umma Corretora de Seguros.
Ao receber um aviso de sinistro (por e-mail ou manual), você deve:
1. Identificar o cliente/segurado na base de apólices
2. Identificar a apólice vigente correspondente — REGRA CRÍTICA: se o texto do aviso mencionar EXPLICITAMENTE um número de apólice (ex: "Apólice 5400057713", "apólice nº 9800020172017", "apólice 2017" etc.), use ESSE número como apolice_identificada, mesmo que a base de apólices contenha outro número. O número informado no aviso tem PRIORIDADE ABSOLUTA sobre qualquer inferência da base.
3. Determinar a complexidade do sinistro (Simples/Complexo)
4. Informar os prazos de regulação aplicáveis (Circ. SUSEP 621/2021, Lei 15.040/2024, Res. CNSP 407/2021)
5. Listar os documentos necessários para a regulação
6. Sugerir um e-mail profissional ao cliente com o passo a passo para início do processo de regulação junto à seguradora

Retorne SOMENTE JSON válido com esta estrutura exata:
{
  "cliente_identificado": "Nome do segurado ou 'Não identificado'",
  "apolice_identificada": "Número da apólice ou 'Não localizada'",
  "seguradora": "Nome da seguradora",
  "produto": "Produto/ramo do seguro",
  "cobertura_aplicavel": "Descrição da cobertura que se aplica",
  "complexidade": "Simples" ou "Complexo",
  "norma_aplicavel": "Nome da norma",
  "prazo_regulacao": número em dias,
  "prazo_pagamento": número em dias ou null,
  "sunset_clause": true ou false,
  "documentos_necessarios": ["lista", "de", "documentos"],
  "acoes_imediatas": ["lista", "de", "ações"],
  "alertas": ["alertas importantes"],
  "email_sugerido": "Texto completo do e-mail ao cliente com saudação, corpo detalhado com passo a passo e assinatura profissional",
  "resumo_operacional": "Resumo para a equipe interna"
}${apolicesCtx}${instrucoesCtx}`;

    try {
      const r = await fetch("/api/analyze", {
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body: JSON.stringify({
          system:sys,
          messages:[{role:"user",content:"Aviso de sinistro recebido:\n\n"+avisoTxt}]
        })
      });
      const d = await r.json();
      const m = (d.content?.[0]?.text||"").match(/\{[\s\S]*\}/);
      if (m) {
        const parsed = JSON.parse(m[0]);
        setAvisoRes(parsed);
        if (parsed.email_sugerido) setAvisoEmail(parsed.email_sugerido);
      } else {
        setAvisoErr("Resposta inválida da IA. Verifique se a variável OPENAI_API_KEY está configurada no painel da Vercel.");
      }
    } catch(e) {
      setAvisoErr("Erro ao conectar com a IA: "+e.message);
    }
    setAvisoLd(false);
  };

  // ── Instruções helpers ────────────────────────────────────────────────────
  const excluirInstrucao = id => { saveInstrucoes(instrucoes.filter(i=>i.id!==id)); };
  const excluirApolice   = id => { saveApolices(apolices.filter(a=>a.id!==id)); };

  // ── NAV ───────────────────────────────────────────────────────────────────
  const NAV=[
    {id:"aviso",     lb:"Aviso",      ic:Mail},
    {id:"dashboard", lb:"Dashboard",  ic:Home},
    {id:"casos",     lb:"Casos",      ic:Database},
    {id:"apolices",  lb:"Apólices",   ic:FileCheck},
    {id:"instrucoes",lb:"Instruções", ic:Clipboard},
    {id:"ai",        lb:"IA",         ic:Zap},
    {id:"normas",    lb:"Normas",     ic:BookOpen},
    {id:"importar",  lb:"Importar",   ic:Upload},
  ];
  const titles={aviso:"Aviso de Sinistro",dashboard:"Dashboard",casos:"Gestão de Casos",
    apolices:"Apólices Vigentes",instrucoes:"Instruções de Regulação",
    ai:"Análise por IA",normas:"Normas SUSEP",importar:"Importar Dados"};

  // ── SIDEBAR ───────────────────────────────────────────────────────────────
  const Sidebar = () => (
    <div style={{width:sideOpen?220:56,minWidth:sideOpen?220:56,background:C.navy,display:"flex",flexDirection:"column",transition:"width 0.25s",overflow:"hidden",flexShrink:0}}>
      <div style={{display:"flex",alignItems:"center",gap:10,padding:sideOpen?"18px 14px":"18px 10px",borderBottom:"1px solid rgba(255,255,255,0.08)",minHeight:60}}>
        <svg width="28" height="28" viewBox="0 0 36 36" fill="none" style={{flexShrink:0}}>
          <circle cx="18" cy="18" r="9" fill={C.blue}/>
          <path d="M4 18 C8 18 12 18 13 18" stroke={C.blue} strokeWidth="5" strokeLinecap="round" opacity="0.7"/>
          <path d="M23 18 C24 18 28 18 32 18" stroke={C.blue} strokeWidth="5" strokeLinecap="round" opacity="0.7"/>
          <circle cx="4" cy="18" r="3" fill={C.blue} opacity="0.4"/>
          <circle cx="32" cy="18" r="3" fill={C.blue} opacity="0.4"/>
        </svg>
        <span style={{color:C.white,fontSize:16,fontWeight:700,opacity:sideOpen?1:0,transition:"opacity 0.2s",whiteSpace:"nowrap"}}>umma</span>
      </div>
      <div style={{flex:1,padding:"8px 0",overflowY:"auto"}}>
        {NAV.map(n=>{
          const A=sec===n.id, Ic=n.ic;
          return (<div key={n.id} style={{display:"flex",alignItems:"center",gap:10,padding:sideOpen?"9px 14px":"9px 10px",margin:"1px 6px",borderRadius:8,cursor:"pointer",background:A?"rgba(62,117,255,0.2)":"transparent",transition:"background 0.15s"}}
            onClick={()=>setSec(n.id)} onMouseEnter={e=>!A&&(e.currentTarget.style.background="rgba(255,255,255,0.07)")} onMouseLeave={e=>!A&&(e.currentTarget.style.background="transparent")}>
            <Ic size={15} style={{flexShrink:0,color:A?C.blue:"rgba(255,255,255,0.5)"}}/>
            <span style={{fontSize:12.5,fontWeight:500,color:A?C.blue:"rgba(255,255,255,0.65)",whiteSpace:"nowrap",opacity:sideOpen?1:0,transition:"opacity 0.2s"}}>{n.lb}</span>
            {n.id==="importar"&&cnt>0&&sideOpen&&<span style={{marginLeft:"auto",background:C.green,color:C.navy,borderRadius:10,padding:"1px 6px",fontSize:9.5,fontWeight:700}}>{cnt}</span>}
            {n.id==="apolices"&&cntAp>0&&sideOpen&&<span style={{marginLeft:"auto",background:C.blue,color:C.white,borderRadius:10,padding:"1px 6px",fontSize:9.5,fontWeight:700}}>{cntAp}</span>}
          </div>);
        })}
      </div>
      <div style={{padding:"8px 6px",borderTop:"1px solid rgba(255,255,255,0.08)"}}>
        <div style={{display:"flex",justifyContent:"center",padding:7,borderRadius:8,cursor:"pointer"}} onClick={()=>setSideOpen(v=>!v)} onMouseEnter={e=>e.currentTarget.style.background="rgba(255,255,255,0.07)"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <Menu size={14} style={{color:"rgba(255,255,255,0.35)"}}/>
        </div>
      </div>
    </div>
  );

  // ── HEADER ────────────────────────────────────────────────────────────────
  const Header = () => (
    <div style={{background:C.white,borderBottom:`1px solid ${C.border}`,padding:`0 ${isMobile?14:20}px`,height:isMobile?52:56,display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
      {isMobile&&(<div style={{display:"flex",alignItems:"center",gap:8,marginRight:4}}>
        <svg width="22" height="22" viewBox="0 0 36 36" fill="none"><circle cx="18" cy="18" r="9" fill={C.blue}/><path d="M4 18 C8 18 12 18 13 18" stroke={C.blue} strokeWidth="5" strokeLinecap="round" opacity="0.7"/><path d="M23 18 C24 18 28 18 32 18" stroke={C.blue} strokeWidth="5" strokeLinecap="round" opacity="0.7"/></svg>
      </div>)}
      <div style={{flex:1,fontSize:isMobile?13:15,fontWeight:700,color:C.navy,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{titles[sec]||sec}</div>
      {crit.length>0&&<div style={{display:"flex",alignItems:"center",gap:5,background:"#FEF2F2",padding:"4px 8px",borderRadius:8,cursor:"pointer",flexShrink:0}} onClick={()=>setSec("casos")}><AlertTriangle size={11} color="#EF4444"/><span style={{fontSize:11,fontWeight:600,color:"#EF4444"}}>{crit.length}</span></div>}
      <div style={{display:"flex",alignItems:"center",gap:5,padding:"4px 8px",background:compliance>=85?"#F0FDF4":"#FFF7ED",borderRadius:8,flexShrink:0}}>
        <Shield size={11} color={compliance>=85?C.green:C.orange}/><span style={{fontSize:11,fontWeight:700,color:compliance>=85?C.green:C.orange}}>{compliance}%</span>
        {!isMobile&&<span style={{fontSize:11,color:compliance>=85?C.green:C.orange}}> Compliance</span>}
      </div>
      {!isMobile&&<span style={{fontSize:11,color:C.grey,flexShrink:0}}>{total} casos</span>}
    </div>
  );

  // ── BOTTOM NAV (mobile) ───────────────────────────────────────────────────
  const BottomNav = () => (
    <div style={{background:C.white,borderTop:`1px solid ${C.border}`,display:"flex",flexShrink:0,height:58,paddingBottom:"env(safe-area-inset-bottom)"}}>
      {NAV.slice(0,5).map(n=>{
        const A=sec===n.id, Ic=n.ic;
        return (<button key={n.id} onClick={()=>setSec(n.id)} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:3,background:"transparent",border:"none",cursor:"pointer",padding:"6px 0",position:"relative"}}>
          <Ic size={18} color={A?C.blue:C.grey}/>
          <span style={{fontSize:9.5,fontWeight:A?700:400,color:A?C.blue:C.grey}}>{n.lb}</span>
        </button>);
      })}
    </div>
  );

  // ── KPI CARD ──────────────────────────────────────────────────────────────
  const KPI = ({label,value,color,sub,onClick}) => (
    <div onClick={onClick} style={{...card,display:"flex",flexDirection:"column",gap:4,cursor:onClick?"pointer":"default"}} onMouseEnter={e=>onClick&&(e.currentTarget.style.boxShadow=`0 4px 12px ${color}22`)} onMouseLeave={e=>e.currentTarget.style.boxShadow="0 1px 4px rgba(2,18,76,0.06)"}>
      <div style={{fontSize:isMobile?22:28,fontWeight:800,color,letterSpacing:-0.5,lineHeight:1}}>{value}</div>
      <div style={{fontSize:isMobile?11:12,fontWeight:600,color:C.body}}>{label}</div>
      {sub&&<div style={{fontSize:10.5,color,fontWeight:600}}>{sub}</div>}
    </div>
  );

  // ── ALERT BANNER ──────────────────────────────────────────────────────────
  const AlertBanner = ({color,bg,border,icon,children,action,actionLabel}) => (
    <div style={{background:bg,border:`1px solid ${border}`,borderRadius:10,padding:isMobile?"10px 12px":"11px 16px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
      <span style={{fontSize:14,flexShrink:0}}>{icon}</span>
      <span style={{flex:1,fontSize:12.5,color,lineHeight:1.5}}>{children}</span>
      {action&&<button onClick={action} style={{...btn(color,true),padding:"5px 10px",fontSize:11.5,flexShrink:0}}>{actionLabel}</button>}
    </div>
  );

  // ── CASE DETAIL OVERLAY ───────────────────────────────────────────────────
  const CaseDetail = () => {
    if (!selC) return null;
    const fields=[["Seguradora",selC.seguradora],["Produto",selC.tipo],["Ramo",selC.ramo||"—"],["Analista",selC.regulador||"—"],
      ["Dt. Sinistro",selC.dataSinistro||"—"],["Dt. Aviso",selC.dataAbertura||"—"],["Dt. Inclusão",selC.dataDocCompleta||"—"],["Início Vig.",selC.inicioVig||"—"],
      ["IMP. Segurada",fCur(selC.importanciaSegurada)],["Apurado",fCur(selC.apurado)],["Franquia",fCur(selC.franquia)],["Indenizado",fCur(selC.indenizado)],
      ["Norma",selC.norma],["Classificação",selC.classificacao],
      ["Prazo Legal",selC.anteriorNorma?"N/A (anterior à norma)":selC.sobreLei15040?"Art.86: 30d aviso + Art.87: 30d pag.":`${selC.prazoLegal}d`],
      ["Prazo Restante",selC.anteriorNorma?"Não aplicável":selC.sobreLei15040?`Art.86: ${selC.prazoArt86!=null?(selC.prazoArt86<=0?`SUNSET! ${Math.abs(selC.prazoArt86)}d vencido`:selC.prazoArt86+"d"):"—"}`:
        selC.prazoRestante!=null?(selC.prazoRestante<0?`VENCIDO (${Math.abs(selC.prazoRestante)}d)`:selC.prazoRestante+"d"):"—"]];
    return (
      <div style={{position:"fixed",inset:0,background:"rgba(2,18,76,0.5)",zIndex:50,display:"flex",alignItems:isMobile?"flex-end":"center",justifyContent:"center",padding:isMobile?0:20}}>
        <div style={{background:C.white,borderRadius:isMobile?"16px 16px 0 0":12,width:"100%",maxWidth:isMobile?"100%":700,maxHeight:isMobile?"92vh":"85vh",overflow:"auto",padding:isMobile?16:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12,position:"sticky",top:0,background:C.white,paddingBottom:8,borderBottom:`1px solid ${C.border}`}}>
            <div>
              <div style={{fontSize:11,color:C.blue,fontWeight:700,letterSpacing:0.5,marginBottom:2}}>SINISTRO</div>
              <div style={{fontSize:isMobile?14:16,fontWeight:700,color:C.navy}}>{selC.id}</div>
              <div style={{fontSize:12,color:C.grey,marginTop:2}}>{selC.segurado}</div>
            </div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <span style={bdg(riskC(selC.risco),riskBg(selC.risco))}>{selC.risco}</span>
              <button onClick={()=>setSelC(null)} style={{background:"transparent",border:"none",cursor:"pointer",padding:4}}><X size={18} color={C.grey}/></button>
            </div>
          </div>
          {selC.sobreLei15040&&(
            <div style={{background:"#F0FDF4",border:"1px solid #6EE7B7",borderRadius:8,padding:"10px 12px",fontSize:12,marginBottom:10,lineHeight:1.5,color:"#047857"}}>
              <strong>⚖️ Lei 15.040/2024</strong> — Contrato celebrado em {selC.inicioVig}. Art. 86: 30d do aviso para cobertura (sunset clause) · Art. 87: 30d para pagar · Art. 88: multa 2%.
              {selC.sunsetAlert&&<div style={{color:"#B91C1C",fontWeight:700,marginTop:4}}>⚠️ SUNSET: prazo Art. 86 vencido — seguradora pode ter perdido o direito de recusar!</div>}
            </div>
          )}
          {selC.anteriorNorma&&(
            <div style={{background:"#ECFEFF",border:"1px solid #A5F3FC",borderRadius:8,padding:"10px 12px",fontSize:12,marginBottom:10,color:"#0E7490",lineHeight:1.5}}>
              <strong>ℹ️ Anterior à Norma</strong> — Aviso em {selC.dataAbertura}, anterior a {selC.normaVigStr}. Prazos legais não se aplicam.
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
            {fields.map(([k,v],i)=>(
              <div key={i} style={{background:C.bg,borderRadius:8,padding:"8px 10px"}}>
                <div style={{fontSize:9.5,color:C.grey,fontWeight:600,textTransform:"uppercase",letterSpacing:0.3}}>{k}</div>
                <div style={{fontSize:12.5,fontWeight:600,color:C.body,marginTop:2,wordBreak:"break-word"}}>{v}</div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button onClick={()=>{setAiTxt(`Sinistro: ${selC.id}\nSegurado: ${selC.segurado}\nProduto: ${selC.tipo} (${selC.classificacao})\nSeguradora: ${selC.seguradora}\nStatus: ${selC.status}\nRisco: ${selC.risco}\nNorma: ${selC.norma}\nApurado: ${fCur(selC.apurado)}`);setSec("ai");setSelC(null)}} style={btn(C.blue,isMobile)}><Zap size={12}/>Analisar com IA</button>
            <button onClick={()=>{setAvisoTxt(`Sinistro: ${selC.id}\nSegurado: ${selC.segurado}\nProduto: ${selC.tipo}\nSeguradora: ${selC.seguradora}\nData Aviso: ${selC.dataAbertura}\nData Sinistro: ${selC.dataSinistro}\nApólice: ${selC.apolice}\nStatus: ${selC.status}\nApurado: ${fCur(selC.apurado)}`);setSec("aviso");setSelC(null)}} style={btn(C.teal,isMobile)}><Mail size={12}/>Gerar E-mail</button>
            <button onClick={()=>setSelC(null)} style={{...btn(C.white,isMobile),color:C.grey}}>Fechar</button>
          </div>
        </div>
      </div>
    );
  };

  // ── APÓLICE DETAIL OVERLAY (PDF) ────────────────────────────────────────────
  const ApoliceDetail = () => {
    if (!selAp) return null;
    return (
      <div style={{position:"fixed",inset:0,background:"rgba(2,18,76,0.5)",zIndex:50,display:"flex",alignItems:isMobile?"flex-end":"center",justifyContent:"center",padding:isMobile?0:20}}>
        <div style={{background:C.white,borderRadius:isMobile?"16px 16px 0 0":12,width:"100%",maxWidth:isMobile?"100%":700,maxHeight:isMobile?"92vh":"88vh",overflow:"auto",padding:isMobile?16:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12,position:"sticky",top:0,background:C.white,paddingBottom:8,borderBottom:`1px solid ${C.border}`}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:11,color:C.blue,fontWeight:700,letterSpacing:0.5,marginBottom:2}}>APÓLICE VIGENTE — PDF</div>
              <div style={{fontSize:isMobile?13:15,fontWeight:700,color:C.navy,wordBreak:"break-word"}}>{selAp.nome}</div>
              <div style={{fontSize:12,color:C.grey,marginTop:2}}>Seguradora: <strong>{selAp.seguradora||"não informada"}</strong> · {selAp.tamanho} · {selAp.data}</div>
            </div>
            <button onClick={()=>setSelAp(null)} style={{background:"transparent",border:"none",cursor:"pointer",padding:4,flexShrink:0}}><X size={18} color={C.grey}/></button>
          </div>
          <div style={{background:C.bg,borderRadius:8,padding:"12px 14px",marginBottom:14,maxHeight:420,overflowY:"auto"}}>
            <div style={{fontSize:10,fontWeight:700,color:C.grey,textTransform:"uppercase",letterSpacing:0.3,marginBottom:6}}>Conteúdo extraído do PDF</div>
            <pre style={{fontSize:11.5,color:C.body,lineHeight:1.6,whiteSpace:"pre-wrap",wordBreak:"break-word",margin:0,fontFamily:font}}>{selAp.texto||"Sem conteúdo extraído."}</pre>
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            <button onClick={()=>{setAvisoTxt(`Apólice: ${selAp.nome}\nSeguradora: ${selAp.seguradora||""}`);setSec("aviso");setSelAp(null)}} style={btn(C.teal,isMobile)}><Mail size={12}/>Criar Aviso</button>
            <button onClick={()=>setSelAp(null)} style={{...btn(C.white,isMobile),color:C.grey}}>Fechar</button>
          </div>
        </div>
      </div>
    );
  };

  // ── INSTRUÇÃO DETAIL OVERLAY (PDF) ───────────────────────────────────────────
  const InstDetail = () => {
    if (!selInst) return null;
    return (
      <div style={{position:"fixed",inset:0,background:"rgba(2,18,76,0.5)",zIndex:50,display:"flex",alignItems:isMobile?"flex-end":"center",justifyContent:"center",padding:isMobile?0:20}}>
        <div style={{background:C.white,borderRadius:isMobile?"16px 16px 0 0":12,width:"100%",maxWidth:isMobile?"100%":700,maxHeight:isMobile?"92vh":"88vh",overflow:"auto",padding:isMobile?16:20}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12,position:"sticky",top:0,background:C.white,paddingBottom:8,borderBottom:`1px solid ${C.border}`}}>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:11,color:C.teal,fontWeight:700,letterSpacing:0.5,marginBottom:2}}>INSTRUÇÕES DE REGULAÇÃO — PDF</div>
              <div style={{fontSize:isMobile?13:15,fontWeight:700,color:C.navy,wordBreak:"break-word"}}>{selInst.nome}</div>
              <div style={{fontSize:12,color:C.grey,marginTop:2}}>Seguradora: <strong>{selInst.seguradora||"não informada"}</strong> · {selInst.tamanho} · {selInst.data}</div>
            </div>
            <button onClick={()=>setSelInst(null)} style={{background:"transparent",border:"none",cursor:"pointer",padding:4,flexShrink:0}}><X size={18} color={C.grey}/></button>
          </div>
          <div style={{background:C.bg,borderRadius:8,padding:"12px 14px",marginBottom:14,maxHeight:420,overflowY:"auto"}}>
            <div style={{fontSize:10,fontWeight:700,color:C.grey,textTransform:"uppercase",letterSpacing:0.3,marginBottom:6}}>Conteúdo extraído do PDF</div>
            <pre style={{fontSize:11.5,color:C.body,lineHeight:1.6,whiteSpace:"pre-wrap",wordBreak:"break-word",margin:0,fontFamily:font}}>{selInst.texto||"Sem conteúdo extraído."}</pre>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={()=>setSelInst(null)} style={{...btn(C.white,isMobile),color:C.grey}}>Fechar</button>
          </div>
        </div>
      </div>
    );
  };


  // ══════════════════════════════════════════════════════════════════════════
  // ── AVISO DE SINISTRO SECTION ─────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const AvisoSection = () => (
    <div style={{display:"flex",flexDirection:"column",gap:14,maxWidth:isMobile?"100%":900}}>
      <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none",padding:isMobile?"16px":"22px 26px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <div style={{width:38,height:38,background:"rgba(62,117,255,0.25)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Mail size={18} color={C.blue}/></div>
          <div>
            <div style={{fontSize:isMobile?14:16,fontWeight:700,color:C.white}}>Aviso de Sinistro — Análise por IA</div>
            <div style={{fontSize:11.5,color:"rgba(255,255,255,0.6)",marginTop:1}}>Cole o e-mail ou descreva o aviso recebido</div>
          </div>
        </div>
        <p style={{color:"rgba(255,255,255,0.72)",fontSize:12.5,margin:0,lineHeight:1.5}}>
          A IA identifica o cliente na base de apólices, analisa a cobertura vigente, informa complexidade, prazos de regulação, documentos necessários e gera e-mail ao cliente com o passo a passo para início do processo junto à seguradora.
        </p>
        {apolices.length===0&&<div style={{marginTop:8,background:"rgba(255,142,79,0.15)",border:"1px solid rgba(255,142,79,0.4)",borderRadius:8,padding:"7px 12px",fontSize:12,color:C.orange,fontWeight:600}}>⚠ Nenhuma apólice cadastrada. <button onClick={()=>setSec("apolices")} style={{background:"transparent",border:"none",color:C.orange,cursor:"pointer",fontWeight:700,fontFamily:font,textDecoration:"underline"}}>Importar apólices →</button></div>}
        {apolices.length>0&&<div style={{marginTop:8,background:"rgba(44,221,124,0.15)",border:"1px solid rgba(44,221,124,0.4)",borderRadius:8,padding:"7px 12px",fontSize:12,color:C.green,fontWeight:600}}>✓ {apolices.length} apólices na base · {instrucoes.length} instruções cadastradas</div>}
      </div>

      <textarea value={avisoTxt} onChange={e=>setAvisoTxt(e.target.value)}
        placeholder={"Cole aqui o e-mail de aviso de sinistro ou descreva manualmente:\n\nExemplo:\nDe: cliente@empresa.com.br\nAssunto: Aviso de Sinistro - Incêndio no Galpão\n\nPrezados,\nInformamos que ocorreu um incêndio em nosso galpão na data de 25/05/2026...\n\nOu manualmente:\nSegurado: Empresa XYZ Ltda\nData do Sinistro: 25/05/2026\nTipo: Incêndio\nDescrição: ..."}
        style={{width:"100%",padding:"12px 14px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",resize:"vertical",minHeight:isMobile?140:180,boxSizing:"border-box"}}/>

      <button onClick={runAvisoIA} disabled={avisoLd||!avisoTxt.trim()} style={{...btn(C.blue),justifyContent:"center",padding:12,opacity:avisoLd||!avisoTxt.trim()?0.55:1}}>
        {avisoLd?<><RefreshCw size={14} style={{animation:"spin 1s linear infinite"}}/>Analisando aviso...</>:<><Zap size={14}/>Analisar Aviso com IA</>}
      </button>

      {avisoErr&&<div style={{padding:"10px 12px",background:"#FEF2F2",borderRadius:8,color:"#B91C1C",fontSize:12.5}}>{avisoErr}</div>}

      {avisoLd&&<div style={{...card,padding:28,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12}}>
        <div style={{width:38,height:38,border:`4px solid ${C.light}`,borderTopColor:C.blue,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
        <span style={{fontSize:13,fontWeight:600,color:C.navy}}>Consultando apólices e normas SUSEP...</span>
      </div>}

      {avisoRes&&(
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          {/* Identificação */}
          <div style={{...card,borderLeft:`4px solid ${C.blue}`}}>
            <div style={{fontSize:11,fontWeight:700,color:C.grey,letterSpacing:0.5,textTransform:"uppercase",marginBottom:8}}>Identificação</div>
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr 1fr",gap:8}}>
              {[["Cliente",avisoRes.cliente_identificado],["Apólice",avisoRes.apolice_identificada],["Seguradora",avisoRes.seguradora],["Produto",avisoRes.produto],["Cobertura",avisoRes.cobertura_aplicavel],["Complexidade",avisoRes.complexidade]].map(([k,v],i)=>(
                <div key={i} style={{background:C.bg,borderRadius:8,padding:"8px 10px"}}>
                  <div style={{fontSize:9.5,color:C.grey,fontWeight:600,textTransform:"uppercase",letterSpacing:0.3}}>{k}</div>
                  <div style={{fontSize:12.5,fontWeight:700,color:C.body,marginTop:2}}>{v||"—"}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Prazos */}
          <div style={{...card,borderLeft:`4px solid ${C.teal}`}}>
            <div style={{fontSize:11,fontWeight:700,color:C.grey,letterSpacing:0.5,textTransform:"uppercase",marginBottom:8}}>Prazos de Regulação</div>
            <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr 1fr",gap:8}}>
              <div style={{background:"#F0FDF4",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9.5,color:C.grey,fontWeight:600,textTransform:"uppercase"}}>Norma Aplicável</div>
                <div style={{fontSize:13,fontWeight:700,color:C.teal,marginTop:2}}>{avisoRes.norma_aplicavel||"—"}</div>
              </div>
              <div style={{background:"#EFF6FF",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9.5,color:C.grey,fontWeight:600,textTransform:"uppercase"}}>Prazo Regulação</div>
                <div style={{fontSize:22,fontWeight:800,color:C.blue,lineHeight:1}}>{avisoRes.prazo_regulacao||"—"}<span style={{fontSize:12}}> dias</span></div>
              </div>
              {avisoRes.prazo_pagamento&&<div style={{background:"#F0FDF4",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9.5,color:C.grey,fontWeight:600,textTransform:"uppercase"}}>Prazo Pagamento</div>
                <div style={{fontSize:22,fontWeight:800,color:C.green,lineHeight:1}}>{avisoRes.prazo_pagamento}<span style={{fontSize:12}}> dias</span></div>
              </div>}
              {avisoRes.sunset_clause&&<div style={{background:"#FEF2F2",borderRadius:8,padding:"10px 12px"}}>
                <div style={{fontSize:9.5,color:"#B91C1C",fontWeight:700,textTransform:"uppercase"}}>⚠ Sunset Clause</div>
                <div style={{fontSize:12,color:"#B91C1C",marginTop:2,fontWeight:600}}>Art. 86 — 30d do aviso</div>
              </div>}
            </div>
          </div>

          {/* Documentos + Ações */}
          <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:12}}>
            <div style={card}>
              <div style={{fontWeight:700,fontSize:13,color:C.navy,marginBottom:8}}>📋 Documentos Necessários</div>
              <ul style={{margin:0,padding:"0 0 0 16px"}}>
                {(avisoRes.documentos_necessarios||[]).map((d,i)=><li key={i} style={{fontSize:12.5,lineHeight:1.6,marginBottom:4,color:C.body}}>{d}</li>)}
              </ul>
            </div>
            <div style={card}>
              <div style={{fontWeight:700,fontSize:13,color:C.navy,marginBottom:8}}>⚡ Ações Imediatas</div>
              <ul style={{margin:0,padding:"0 0 0 16px"}}>
                {(avisoRes.acoes_imediatas||[]).map((a,i)=><li key={i} style={{fontSize:12.5,lineHeight:1.6,marginBottom:4,color:C.body}}>{a}</li>)}
              </ul>
            </div>
          </div>

          {/* Alertas */}
          {(avisoRes.alertas||[]).length>0&&<div style={{...card,background:"#FFFBEB",borderColor:"#FCD34D"}}>
            <div style={{fontWeight:700,fontSize:13,color:"#92400E",marginBottom:8}}>⚠️ Alertas</div>
            <ul style={{margin:0,padding:"0 0 0 16px"}}>
              {avisoRes.alertas.map((a,i)=><li key={i} style={{fontSize:12.5,color:"#92400E",lineHeight:1.6,marginBottom:4}}>{a}</li>)}
            </ul>
          </div>}

          {/* Resumo operacional */}
          {avisoRes.resumo_operacional&&<div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none"}}>
            <div style={{fontSize:10,fontWeight:700,color:C.green,letterSpacing:0.5,marginBottom:5,textTransform:"uppercase"}}>Resumo Operacional</div>
            <p style={{color:C.white,fontSize:12.5,lineHeight:1.6,margin:0}}>{avisoRes.resumo_operacional}</p>
          </div>}

          {/* E-mail sugerido */}
          {avisoEmail&&<div style={card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
              <div style={{fontWeight:700,fontSize:13,color:C.navy}}>✉️ E-mail Sugerido ao Cliente</div>
              <button onClick={()=>{navigator.clipboard.writeText(avisoEmail);setCopiedEmail(true);setTimeout(()=>setCopiedEmail(false),2000);}} style={{...btn(copiedEmail?C.green:C.blue,true)}}>
                <Copy size={12}/>{copiedEmail?"Copiado!":"Copiar"}
              </button>
            </div>
            <textarea value={avisoEmail} onChange={e=>setAvisoEmail(e.target.value)}
              style={{width:"100%",padding:"12px 14px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:12.5,color:C.body,outline:"none",resize:"vertical",minHeight:280,boxSizing:"border-box",lineHeight:1.6}}/>
            <div style={{marginTop:8,display:"flex",gap:8,flexWrap:"wrap"}}>
              <button onClick={()=>{navigator.clipboard.writeText(avisoEmail);setCopiedEmail(true);setTimeout(()=>setCopiedEmail(false),2000);}} style={btn(C.blue,isMobile)}><Copy size={12}/>{copiedEmail?"Copiado!":"Copiar E-mail"}</button>
              <button onClick={()=>{setAvisoTxt("");setAvisoRes(null);setAvisoEmail("");}} style={{...btn(C.white,isMobile),color:C.grey}}>Novo Aviso</button>
            </div>
          </div>}
        </div>
      )}
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════
  // ── DASHBOARD ─────────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const Dashboard = () => (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      {lei15040Count>0&&<AlertBanner color="#065F46" bg="#F0FDF4" border="#6EE7B7" icon="⚖️" action={vencidos>0?()=>{setFRk("Vencido sob Lei 15.040");setSec("casos");}:undefined} actionLabel="Ver vencidos Lei 15.040 →">
        <strong>{lei15040Count} caso(s) regidos pela Lei 15.040/2024</strong> — Art. 86: prazo inicia quando documentações são suficientes (30d) · Art. 87: 30d para pagar · Art. 88: multa 2%.
        {vencidos>0&&<strong style={{color:"#B91C1C"}}> ⚠️ {vencidos} com prazo vencido sob Lei 15.040 (Art. 86 — ação imediata)! Clique para ver.</strong>}
      </AlertBanner>}
      {crit.length>0&&<AlertBanner color="#B91C1C" bg="#FEF2F2" border="#FECACA" icon="⚠️" action={()=>{setFRk("Vencido sob Lei 15.040");setSec("casos");}} actionLabel="Ver vencidos →">
        <strong>{vencidos} vencido(s)</strong> e {altoRisco} com prazo ≤ 7 dias — ação imediata (Circ. SUSEP 621/2021 / Lei 15.040/2024). Obs.: prazo suspenso para casos com documentação pendente do segurado (Art. 44, Circ. 621/2021).
      </AlertBanner>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <KPI label="Total de Casos" value={total} color={C.blue} sub={`${emReg} em regulação`} onClick={()=>{setFSt("Em Regulação");setSec("casos");}}/>
        <KPI label="Em Litígio" value={litigio} color={C.purple} sub="monitorar"/>
        <KPI label="Vencidos Lei 15.040" value={vencidos} color="#B91C1C" sub="ação imediata" onClick={()=>{setFRk("Vencido sob Lei 15.040");setSec("casos")}}/>
        <KPI label="Liquidados" value={liquidado} color={C.green} sub={`${encerrado} enc. · ${recusado} recus. · ${findados} findados`} onClick={()=>{setFSt("Liquidado");setSec("casos");}}/>
        <KPI label="Ag. Documentação" value={agDoc} color={C.orange} sub="P.Doc Segurado + Terceiro" onClick={()=>{setFSt("Ag. Documentação");setSec("casos");}}/>
        <KPI label="Compliance SUSEP" value={`${compliance}%`} color={compliance>=85?C.green:C.orange} sub={compliance>=85?"✓ prazo OK":"▼ casos vencidos afetam"}/>
      </div>
      {!isMobile&&(
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          <div style={card}><div style={{fontWeight:700,fontSize:13,color:C.navy,marginBottom:12}}>Distribuição por Status</div>
            <ResponsiveContainer width="100%" height={200}><PieChart><Pie data={stChart} dataKey="v" nameKey="n" cx="50%" cy="50%" innerRadius={45} outerRadius={80} paddingAngle={3}>{stChart.map((d,i)=><Cell key={i} fill={d.c}/>)}</Pie><Tooltip contentStyle={{borderRadius:8,fontSize:11}}/><Legend iconType="circle" wrapperStyle={{fontSize:10}}/></PieChart></ResponsiveContainer>
          </div>
          <div style={card}><div style={{fontWeight:700,fontSize:13,color:C.navy,marginBottom:12}}>Distribuição de Risco</div>
            <ResponsiveContainer width="100%" height={200}><PieChart><Pie data={rkChart} dataKey="v" nameKey="n" cx="50%" cy="50%" innerRadius={45} outerRadius={80} paddingAngle={3}>{rkChart.map((d,i)=><Cell key={i} fill={d.c}/>)}</Pie><Tooltip contentStyle={{borderRadius:8,fontSize:11}}/><Legend iconType="circle" wrapperStyle={{fontSize:10}}/></PieChart></ResponsiveContainer>
          </div>
        </div>
      )}
      <div style={card}><div style={{fontWeight:700,fontSize:13,color:C.navy,marginBottom:12}}>Top Produtos / Ramos</div>
        <ResponsiveContainer width="100%" height={isMobile?160:190}><BarChart data={tipoChart} layout="vertical" barSize={12}><CartesianGrid strokeDasharray="3 3" stroke={C.light} horizontal={false}/><XAxis type="number" tick={{fontSize:9,fill:C.grey}} axisLine={false} tickLine={false}/><YAxis type="category" dataKey="n" tick={{fontSize:isMobile?9:10.5,fill:C.grey}} axisLine={false} tickLine={false} width={isMobile?85:105}/><Tooltip contentStyle={{borderRadius:8,fontSize:11}}/><Bar dataKey="v" name="Casos" fill={C.blue} radius={[0,4,4,0]}/></BarChart></ResponsiveContainer>
      </div>
      <div style={{fontSize:10.5,color:C.grey,padding:"2px 4px",marginBottom:-6}}>* Valores em R$ (BRL). Casos em moeda estrangeira podem ter valores convertidos pela seguradora.</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[{l:"IMP. Segurada Total (R$)",v:fCur(valTotal),c:C.navy},{l:"Total Apurado (R$)",v:fCur(valApurado),c:C.blue},
          {l:"Total Indenizado (R$)",v:fCur(valIndeni),c:C.green},{l:"Lei 15.040/2024",v:lei15040Count+" casos",c:C.teal}
        ].map((k,i)=>(
          <div key={i} style={{...card,borderLeft:`3px solid ${k.c}`}}>
            <div style={{fontSize:isMobile?15:18,fontWeight:800,color:k.c}}>{k.v}</div>
            <div style={{fontSize:11,fontWeight:600,color:C.body,marginTop:2}}>{k.l}</div>
          </div>
        ))}
      </div>
      <div style={{...card,background:"#F8FAFF",border:`1px solid ${C.light}`}}>
        <div style={{fontWeight:700,fontSize:12,color:C.navy,marginBottom:8}}>📊 Regras de Cálculo — Transparência</div>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:8}}>
          <div style={{fontSize:11.5,color:C.grey,lineHeight:1.6}}><strong style={{color:C.body}}>Em Regulação ({emReg}):</strong> Inclui EM REGULAÇÃO + P. SEGURADORA + P. REGULADOR e similares (processo ativo sem pendência de doc). Ag. Documentação é exibido separadamente.</div>
          <div style={{fontSize:11.5,color:C.grey,lineHeight:1.6}}><strong style={{color:C.body}}>Compliance ({compliance}%):</strong> % de casos ativos com risco Baixo ou Médio e documentação completa. Meta: ≥85%. Casos vencidos/alto risco reduzem o índice.</div>
          <div style={{fontSize:11.5,color:C.grey,lineHeight:1.6}}><strong style={{color:C.body}}>Sunset/Decaimento ({sunsetCount}):</strong> Apólices ≥ 11/12/2025 com prazo Art. 86 (&gt;30d do aviso) vencido. Risco é da <em>seguradora</em> perder o direito de recusar cobertura, não do segurado.</div>
          <div style={{fontSize:11.5,color:C.grey,lineHeight:1.6}}><strong style={{color:C.body}}>Findados ({findados}):</strong> Liquidado ({liquidado}) + Encerrado ({encerrado}) + Recusado/Declinado ({recusado}) = {findados} processos concluídos (meta: 819). Valores em R$ (BRL); casos em moeda estrangeira podem ter valores convertidos pela seguradora.</div>
        </div>
      </div>
    </div>
  );


  // ══════════════════════════════════════════════════════════════════════════
  // ── CASOS SECTION ─────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const CasosSection = () => (
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      {isMobile?(
        <div style={card}>
          <div style={{position:"relative"}}>
            <Search size={13} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:C.grey}}/>
            <input value={srch} onChange={e=>setSrch(e.target.value)} placeholder="Buscar..." style={{width:"100%",padding:"9px 12px 9px 30px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
          </div>
          <button onClick={()=>setShowFilters(v=>!v)} style={{...btn(C.white,true),marginTop:8,width:"100%",justifyContent:"space-between",color:C.grey}}>
            <span>Filtros: {fSt!=="Todos"||fRk!=="Todos"?`${fSt!=="Todos"?fSt:""}${fRk!=="Todos"?" · "+fRk:""}`:"Todos"}</span>
            {showFilters?<ChevronUp size={14}/>:<ChevronDown size={14}/>}
          </button>
          {showFilters&&(<div style={{display:"flex",flexDirection:"column",gap:8,marginTop:8}}>
            {[{v:fSt,s:setFSt,o:ALL_ST,l:"Status"},{v:fRk,s:setFRk,o:ALL_RK,l:"Risco"}].map((f,i)=>(
              <select key={i} value={f.v} onChange={e=>f.s(e.target.value)} style={{padding:"9px 10px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,background:C.white,width:"100%"}}>
                {f.o.map(o=><option key={o}>{o}</option>)}
              </select>
            ))}
            <select value={fSeg} onChange={e=>setFSeg(e.target.value)} style={{padding:"9px 10px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,background:C.white,width:"100%"}}>
              {ALL_SEG.map(s=><option key={s} value={s}>{s==="Todos"?"Todos os Clientes":s}</option>)}
            </select>
          </div>)}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
            <span style={{fontSize:11.5,color:C.grey}}>{filtrados.length}/{total} casos</span>
            <button onClick={()=>setSec("importar")} style={btn(C.navy,true)}><Upload size={11}/>Importar</button>
          </div>
        </div>
      ):(
        <div style={{...card,padding:"12px 16px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          <div style={{position:"relative",flex:1,minWidth:180}}>
            <Search size={12} style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",color:C.grey}}/>
            <input value={srch} onChange={e=>setSrch(e.target.value)} placeholder="Buscar por ID, segurado, produto..." style={{width:"100%",padding:"8px 12px 8px 28px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
          </div>
          {[{v:fSt,s:setFSt,o:ALL_ST,lbl:"Status"},{v:fRk,s:setFRk,o:ALL_RK,lbl:"Risco"}].map((f,i)=>(
            <select key={i} value={f.v} onChange={e=>f.s(e.target.value)} style={{padding:"8px 10px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,cursor:"pointer",background:C.white}}>
              {f.o.map(o=><option key={o}>{o}</option>)}
            </select>
          ))}
          <select value={fSeg} onChange={e=>setFSeg(e.target.value)} style={{padding:"8px 10px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,cursor:"pointer",background:C.white,maxWidth:220}}>
            {ALL_SEG.map(s=><option key={s} value={s}>{s==="Todos"?"Todos os Clientes":s}</option>)}
          </select>
          <span style={{fontSize:12,color:C.grey}}>{filtrados.length}/{total}</span>
          <button onClick={()=>setSec("importar")} style={btn(C.navy,true)}><Upload size={12}/>Importar</button>
        </div>
      )}
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {filtrados.slice(0,100).map(c=>(
            <div key={c.id+c.nAviso} style={{...card,padding:"12px 14px"}} onClick={()=>setSelC(c)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                <div>
                  <div style={{fontSize:11.5,fontWeight:700,color:C.blue}}>{c.id||"—"}</div>
                  <div style={{fontSize:13,fontWeight:600,color:C.body,marginTop:1}}>{c.segurado}</div>
                  <div style={{fontSize:11.5,color:C.grey,marginTop:1}}>{c.tipo} · {c.seguradora}</div>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end"}}>
                  <span style={bdg(riskC(c.risco),riskBg(c.risco))}>{c.risco}</span>
                  <span style={{...bdg(statC(c.status)),fontSize:10}}>{c.status}</span>
                </div>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:6,paddingTop:6,borderTop:`1px solid ${C.border}`}}>
                <span style={{fontSize:11.5,color:C.grey}}>{c.dataAbertura}</span>
                <span style={{fontSize:12,fontWeight:700,color:riskC(c.risco)}}>
                  {c.prazoRestante!=null?(c.prazoRestante<0?`−${Math.abs(c.prazoRestante)}d vencido`:`${c.prazoRestante}d restantes`):""}
                  {c.sobreLei15040&&c.prazoArt86!=null&&` · Art.86: ${c.prazoArt86<=0?"sunset!":c.prazoArt86+"d"}`}
                </span>
                <span style={{fontSize:12,color:C.grey}}>{fCur(c.apurado)}</span>
              </div>
              {c.sobreLei15040&&<span style={{...bdg(C.teal,"#F0FDF4"),marginTop:6,fontSize:9.5}}>⚖️ Lei 15.040</span>}
            </div>
          ))}
          {filtrados.length===0&&<div style={{...card,padding:28,textAlign:"center",color:C.grey,fontSize:13}}>Nenhum caso encontrado.<br/><button onClick={()=>setSec("importar")} style={{background:"transparent",border:"none",color:C.blue,cursor:"pointer",fontWeight:600,fontFamily:font,marginTop:8}}>Importar planilha →</button></div>}
        </div>
      ):(
        <div style={{...card,padding:0,overflow:"hidden"}}>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
              <thead><tr style={{background:C.navy}}>{["Nº Sinistro","Segurado","Produto","Seguradora","Analista","Dt. Aviso","Dt. Inclusão","Status","Risco","Prazo","Apurado",""].map(h=><th key={h} style={{padding:"9px 8px",textAlign:"left",fontSize:10,fontWeight:600,color:"rgba(255,255,255,0.8)",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
              <tbody>
                {filtrados.slice(0,200).map(c=>(
                  <tr key={c.id+c.nAviso} style={{borderBottom:`1px solid ${C.border}`}} onMouseEnter={e=>e.currentTarget.style.background=C.bg} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <td style={{padding:"8px 8px",fontSize:11.5,fontWeight:700,color:C.blue,whiteSpace:"nowrap"}}>{c.id||"—"}</td>
                    <td style={{padding:"8px 8px",fontSize:11.5,maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.segurado}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,maxWidth:110,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.tipo}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.seguradora}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,whiteSpace:"nowrap"}}>{c.regulador}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,whiteSpace:"nowrap"}}>{c.dataAbertura}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,whiteSpace:"nowrap"}}>{c.dataDocCompleta}</td>
                    <td style={{padding:"8px 8px"}}><span style={{...bdg(statC(c.status)),fontSize:10}}>{c.status}</span>{c.sobreLei15040&&<span style={{...bdg(C.teal,"#F0FDF4"),fontSize:9,marginLeft:3}}>Lei 15.040</span>}</td>
                    <td style={{padding:"8px 8px"}}><span style={bdg(riskC(c.risco),riskBg(c.risco))}>{c.risco}</span></td>
                    <td style={{padding:"8px 8px",fontSize:11,fontWeight:700,color:riskC(c.risco),whiteSpace:"nowrap"}}>{c.prazoRestante!=null?(c.prazoRestante<0?`−${Math.abs(c.prazoRestante)}d`:c.prazoRestante+"d"):"—"}</td>
                    <td style={{padding:"8px 8px",fontSize:11,color:C.grey,whiteSpace:"nowrap"}}>{fCur(c.apurado)}</td>
                    <td style={{padding:"8px 8px"}}><button onClick={()=>setSelC(c)} style={{background:"transparent",border:`1px solid ${C.blue}`,color:C.blue,borderRadius:6,padding:"3px 7px",fontSize:10.5,cursor:"pointer",fontFamily:font,fontWeight:600}}><Eye size={10} style={{verticalAlign:"middle"}}/></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {filtrados.length>200&&<div style={{padding:"10px",textAlign:"center",fontSize:12,color:C.grey}}>Exibindo 200 de {filtrados.length}. Use os filtros para refinar.</div>}
          {filtrados.length===0&&<div style={{padding:28,textAlign:"center",color:C.grey,fontSize:13}}>Nenhum caso · <button onClick={()=>setSec("importar")} style={{background:"transparent",border:"none",color:C.blue,cursor:"pointer",fontWeight:600,fontFamily:font}}>Importar →</button></div>}
        </div>
      )}
    </div>
  );


  // ══════════════════════════════════════════════════════════════════════════
  // ── APÓLICES SECTION (PDF Upload) ───────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const ApolicesSection = () => {
    const filtAp = apolices.filter(a=>!srchAp||[a.nome,a.seguradora].some(v=>(v||"").toLowerCase().includes(srchAp.toLowerCase())));
    return (
      <div style={{display:"flex",flexDirection:"column",gap:14,maxWidth:isMobile?"100%":900}}>
        <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none",padding:isMobile?"16px":"22px 26px"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{width:38,height:38,background:"rgba(62,117,255,0.25)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileCheck size={18} color={C.blue}/></div>
            <div>
              <div style={{fontSize:isMobile?14:16,fontWeight:700,color:C.white}}>Apólices Vigentes</div>
              <div style={{fontSize:11.5,color:"rgba(255,255,255,0.6)",marginTop:1}}>Upload de PDF — a IA lê as cláusulas automaticamente</div>
            </div>
          </div>
          <p style={{color:"rgba(255,255,255,0.72)",fontSize:12.5,margin:0,lineHeight:1.5}}>Informe a seguradora e faça o upload do PDF da apólice. O texto das cláusulas será extraído e usado pela IA ao analisar avisos de sinistro.</p>
          {cntAp>0&&<div style={{marginTop:8,background:"rgba(44,221,124,0.15)",border:"1px solid rgba(44,221,124,0.4)",borderRadius:8,padding:"7px 12px",fontSize:12,color:C.green,fontWeight:600}}>✓ {cntAp} apólice(s) carregada(s) na base.</div>}
        </div>

        <div style={{...card,border:`2px solid ${C.blue}`}}>
          <div style={{fontWeight:700,fontSize:13.5,color:C.navy,marginBottom:12}}>Adicionar Nova Apólice</div>
          <div style={{marginBottom:10}}>
            <div style={{fontSize:11,fontWeight:600,color:C.grey,marginBottom:4}}>Seguradora *</div>
            <input value={segAp} onChange={e=>setSegAp(e.target.value)} placeholder="Ex: Porto Seguro, Tokio Marine, Zurich..." style={{width:"100%",padding:"9px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div onClick={()=>fRefAp.current?.click()} onDragOver={e=>{e.preventDefault();setDragAp(true)}} onDragLeave={()=>setDragAp(false)} onDrop={e=>{e.preventDefault();setDragAp(false);handleFileAp(e.dataTransfer.files[0],segAp)}}
            style={{border:`2.5px dashed ${dragAp?C.blue:C.border}`,borderRadius:12,padding:isMobile?"24px 16px":"32px 24px",textAlign:"center",cursor:"pointer",background:dragAp?C.light+"44":C.bg,transition:"all 0.2s"}}>
            <div style={{width:44,height:44,margin:"0 auto 10px",background:C.blue+"18",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center"}}>
              {impAp?<RefreshCw size={20} color={C.blue} style={{animation:"spin 0.9s linear infinite"}}/>:<Upload size={20} color={C.blue}/>}
            </div>
            <div style={{fontSize:isMobile?13:14,fontWeight:700,color:C.navy,marginBottom:4}}>{impAp?"Processando PDF...":isMobile?"Toque para selecionar PDF":"Arraste ou clique para selecionar o PDF da apólice"}</div>
            <div style={{fontSize:12,color:C.grey}}>{impAp?"Extraindo cláusulas...":"Apenas arquivos .pdf"}</div>
          </div>
          <input ref={fRefAp} type="file" accept=".pdf" style={{display:"none"}} onChange={e=>{handleFileAp(e.target.files[0],segAp);e.target.value="";}}/>
        </div>

        {logAp.length>0&&(
          <div style={{...card,padding:0,overflow:"hidden"}}>
            <div style={{padding:"8px 14px",background:C.navy,fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:0.5,textTransform:"uppercase"}}>Log de Importação</div>
            {logAp.map((l,i)=>(
              <div key={i} style={{padding:"8px 14px",borderBottom:i<logAp.length-1?`1px solid ${C.border}`:"none",display:"flex",alignItems:"flex-start",gap:8,background:l.t==="err"?"#FEF2F2":C.white}}>
                {l.t==="ok"?<CheckCircle size={13} color={C.green} style={{marginTop:1,flexShrink:0}}/>:<AlertCircle size={13} color={C.red} style={{marginTop:1,flexShrink:0}}/>}
                <span style={{fontSize:12,color:l.t==="err"?"#B91C1C":C.body,lineHeight:1.5}}>{l.m}</span>
              </div>
            ))}
          </div>
        )}

        {apolices.length>0&&(
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{position:"relative"}}>
              <Search size={13} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:C.grey}}/>
              <input value={srchAp} onChange={e=>setSrchAp(e.target.value)} placeholder="Buscar apólice por nome ou seguradora..." style={{width:"100%",padding:"9px 12px 9px 30px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
            </div>
            <div style={{fontSize:12,color:C.grey}}>{filtAp.length} de {apolices.length} apólice(s)</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {filtAp.map((a,i)=>(
                <div key={a.id||i} style={{...card,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                  <div style={{width:36,height:36,background:C.blue+"18",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileCheck size={16} color={C.blue}/></div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.navy,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{a.nome}</div>
                    <div style={{fontSize:11.5,color:C.grey,marginTop:2}}>Seguradora: <strong>{a.seguradora||"não informada"}</strong> · {a.tamanho} · Adicionada em {a.data}</div>
                  </div>
                  <div style={{display:"flex",gap:6,flexShrink:0}}>
                    <button onClick={()=>setSelAp(a)} style={{...btn(C.blue,true),padding:"5px 10px"}}><Eye size={11}/>Ver</button>
                    <button onClick={()=>{setAvisoTxt(`Apólice: ${a.nome}
Seguradora: ${a.seguradora||""}`);setSec("aviso")}} style={{...btn(C.teal,true),padding:"5px 10px"}}><Mail size={11}/>Aviso</button>
                    <button onClick={()=>excluirApolice(a.id)} style={{...btn(C.white,true),padding:"5px 8px",color:C.red}}><Trash2 size={11}/></button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {apolices.length>0&&<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          <button onClick={()=>{if(window.confirm("Limpar todas as apólices?")) saveApolices([]);}} style={{...btn(C.white,isMobile),color:C.red,border:`1px solid ${C.red}22`}}><Trash2 size={13}/>Limpar Base</button>
        </div>}
      </div>
    );
  };


  // ══════════════════════════════════════════════════════════════════════════
  // ── INSTRUÇÕES SECTION (PDF Upload) ──────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const InstrucoesSection = () => {
    const filtInst = instrucoes.filter(i=>!srchInst||[i.nome,i.seguradora].some(v=>(v||"").toLowerCase().includes(srchInst.toLowerCase())));
    return (
      <div style={{display:"flex",flexDirection:"column",gap:14,maxWidth:isMobile?"100%":900}}>
        <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none",padding:isMobile?"16px":"22px 26px"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{width:38,height:38,background:"rgba(62,117,255,0.25)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Clipboard size={18} color={C.blue}/></div>
            <div>
              <div style={{fontSize:isMobile?14:16,fontWeight:700,color:C.white}}>Instruções de Regulação</div>
              <div style={{fontSize:11.5,color:"rgba(255,255,255,0.6)",marginTop:1}}>Upload de PDF — procedimentos e lista de documentos por seguradora</div>
            </div>
          </div>
          <p style={{color:"rgba(255,255,255,0.72)",fontSize:12.5,margin:0,lineHeight:1.5}}>Informe a seguradora e faça o upload do PDF com as instruções de regulação e lista de documentos necessários. A IA usará esse conteúdo para orientar o processo ao receber um aviso de sinistro.</p>
          {instrucoes.length>0&&<div style={{marginTop:8,background:"rgba(44,221,124,0.15)",border:"1px solid rgba(44,221,124,0.4)",borderRadius:8,padding:"7px 12px",fontSize:12,color:C.green,fontWeight:600}}>✓ {instrucoes.length} documento(s) de instrução carregado(s).</div>}
        </div>

        <div style={{...card,border:`2px solid ${C.teal}`}}>
          <div style={{fontWeight:700,fontSize:13.5,color:C.navy,marginBottom:12}}>Adicionar Novo Documento de Instrução</div>
          <div style={{marginBottom:10}}>
            <div style={{fontSize:11,fontWeight:600,color:C.grey,marginBottom:4}}>Seguradora *</div>
            <input value={segInst} onChange={e=>setSegInst(e.target.value)} placeholder="Ex: Porto Seguro, Tokio Marine, Zurich..." style={{width:"100%",padding:"9px 12px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
          </div>
          <div onClick={()=>fRefInst.current?.click()} onDragOver={e=>{e.preventDefault();setDragInst(true)}} onDragLeave={()=>setDragInst(false)} onDrop={e=>{e.preventDefault();setDragInst(false);handleFileInst(e.dataTransfer.files[0],segInst)}}
            style={{border:`2.5px dashed ${dragInst?C.teal:C.border}`,borderRadius:12,padding:isMobile?"24px 16px":"32px 24px",textAlign:"center",cursor:"pointer",background:dragInst?C.light+"44":C.bg,transition:"all 0.2s"}}>
            <div style={{width:44,height:44,margin:"0 auto 10px",background:C.teal+"18",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center"}}>
              {impInst?<RefreshCw size={20} color={C.teal} style={{animation:"spin 0.9s linear infinite"}}/>:<Upload size={20} color={C.teal}/>}
            </div>
            <div style={{fontSize:isMobile?13:14,fontWeight:700,color:C.navy,marginBottom:4}}>{impInst?"Processando PDF...":isMobile?"Toque para selecionar PDF":"Arraste ou clique para selecionar o PDF de instruções"}</div>
            <div style={{fontSize:12,color:C.grey}}>{impInst?"Extraindo conteúdo...":"Apenas arquivos .pdf"}</div>
          </div>
          <input ref={fRefInst} type="file" accept=".pdf" style={{display:"none"}} onChange={e=>{handleFileInst(e.target.files[0],segInst);e.target.value="";}}/>
        </div>

        {logInst.length>0&&(
          <div style={{...card,padding:0,overflow:"hidden"}}>
            <div style={{padding:"8px 14px",background:C.navy,fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:0.5,textTransform:"uppercase"}}>Log de Importação</div>
            {logInst.map((l,i)=>(
              <div key={i} style={{padding:"8px 14px",borderBottom:i<logInst.length-1?`1px solid ${C.border}`:"none",display:"flex",alignItems:"flex-start",gap:8,background:l.t==="err"?"#FEF2F2":C.white}}>
                {l.t==="ok"?<CheckCircle size={13} color={C.green} style={{marginTop:1,flexShrink:0}}/>:<AlertCircle size={13} color={C.red} style={{marginTop:1,flexShrink:0}}/>}
                <span style={{fontSize:12,color:l.t==="err"?"#B91C1C":C.body,lineHeight:1.5}}>{l.m}</span>
              </div>
            ))}
          </div>
        )}

        {instrucoes.length===0&&logInst.length===0&&(
          <div style={{...card,padding:28,textAlign:"center"}}>
            <Clipboard size={36} color={C.border} style={{marginBottom:10}}/>
            <div style={{fontSize:13,color:C.grey,marginBottom:4}}>Nenhum documento de instrução cadastrado ainda.</div>
            <div style={{fontSize:12,color:C.grey}}>Adicione PDFs com procedimentos de regulação e lista de documentos por seguradora.</div>
          </div>
        )}

        {instrucoes.length>0&&(
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{position:"relative"}}>
              <Search size={13} style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:C.grey}}/>
              <input value={srchInst} onChange={e=>setSrchInst(e.target.value)} placeholder="Buscar instrução por nome ou seguradora..." style={{width:"100%",padding:"9px 12px 9px 30px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13,color:C.body,outline:"none",boxSizing:"border-box"}}/>
            </div>
            <div style={{fontSize:12,color:C.grey}}>{filtInst.length} de {instrucoes.length} documento(s)</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {filtInst.map((inst,i)=>(
                <div key={inst.id||i} style={{...card,borderLeft:`4px solid ${C.teal}`,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                  <div style={{width:36,height:36,background:C.teal+"18",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><Clipboard size={16} color={C.teal}/></div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,fontWeight:700,color:C.navy,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{inst.nome}</div>
                    <div style={{fontSize:11.5,color:C.grey,marginTop:2}}>Seguradora: <strong>{inst.seguradora||"não informada"}</strong> · {inst.tamanho} · Adicionado em {inst.data}</div>
                  </div>
                  <div style={{display:"flex",gap:6,flexShrink:0}}>
                    <button onClick={()=>setSelInst(inst)} style={{...btn(C.teal,true),padding:"5px 10px"}}><Eye size={11}/>Ver</button>
                    <button onClick={()=>excluirInstrucao(inst.id)} style={{...btn(C.white,true),padding:"5px 8px",color:C.red}}><Trash2 size={11}/></button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {instrucoes.length>0&&<div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          <button onClick={()=>{if(window.confirm("Limpar todos os documentos de instrução?")) saveInstrucoes([]);}} style={{...btn(C.white,isMobile),color:C.red,border:`1px solid ${C.red}22`}}><Trash2 size={13}/>Limpar Documentos</button>
        </div>}
      </div>
    );
  };


    // ══════════════════════════════════════════════════════════════════════════
  // ── AI SECTION (análise técnica rápida) ───────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const AISection = () => (
    <div style={{display:"flex",flexDirection:isMobile?"column":"row",gap:14}}>
      <div style={{flex:1,display:"flex",flexDirection:"column",gap:12}}>
        <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}><Zap size={16} color={C.green}/><span style={{fontSize:isMobile?13:14,fontWeight:700,color:C.white}}>Análise Técnica SUSEP por IA</span></div>
          <p style={{color:"rgba(255,255,255,0.68)",fontSize:12.5,margin:0,lineHeight:1.5}}>Análise técnica rápida com base nas normas SUSEP e Lei 15.040/2024. Para análise completa com e-mail ao cliente, use o módulo Aviso de Sinistro.</p>
        </div>
        <textarea value={aiTxt} onChange={e=>setAiTxt(e.target.value)} placeholder="Descreva: produto, segurado, seguradora, valor apurado, status atual, data do aviso, início de vigência do contrato..." style={{width:"100%",padding:"12px 14px",borderRadius:8,border:`1.5px solid ${C.border}`,fontFamily:font,fontSize:13.5,color:C.body,outline:"none",resize:"vertical",minHeight:isMobile?120:150,boxSizing:"border-box"}}/>
        <button onClick={runAI} disabled={aiLd||!aiTxt.trim()} style={{...btn(C.blue),justifyContent:"center",padding:12,opacity:aiLd||!aiTxt.trim()?0.55:1}}>
          {aiLd?<><RefreshCw size={14} style={{animation:"spin 1s linear infinite"}}/>Analisando...</>:<><Zap size={14}/>Analisar com IA</>}
        </button>
        {aiErr&&<div style={{padding:"10px 12px",background:"#FEF2F2",borderRadius:8,color:"#B91C1C",fontSize:12.5}}>{aiErr}</div>}
      </div>
      <div style={{flex:1,minWidth:0}}>
        {!aiRes&&!aiLd&&<div style={{...card,height:isMobile?120:260,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:8,opacity:0.5}}><Shield size={36} color={C.blue} style={{opacity:0.3}}/><span style={{color:C.grey,fontSize:12.5,textAlign:"center"}}>Resultado aparecerá aqui</span></div>}
        {aiLd&&<div style={{...card,height:isMobile?120:260,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:12}}><div style={{width:38,height:38,border:`4px solid ${C.light}`,borderTopColor:C.blue,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/><span style={{fontSize:13,fontWeight:600,color:C.navy}}>Consultando normas SUSEP...</span></div>}
        {aiRes&&(
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{...card,borderColor:aiRes.score>=80?C.green:aiRes.score>=60?C.orange:"#EF4444",background:`linear-gradient(135deg,${aiRes.score>=80?C.green:aiRes.score>=60?C.orange:"#EF4444"}14,${C.white})`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div><div style={{fontSize:9.5,fontWeight:700,color:C.grey,letterSpacing:0.5,textTransform:"uppercase"}}>Score Compliance</div><div style={{fontSize:34,fontWeight:800,color:aiRes.score>=80?C.green:aiRes.score>=60?C.orange:"#EF4444",lineHeight:1.1}}>{aiRes.score}<span style={{fontSize:15}}>%</span></div></div>
                <div style={{textAlign:"right"}}><span style={bdg(riskC(aiRes.risco),(riskC(aiRes.risco))+"18")}>{aiRes.risco}</span><div style={{fontSize:11.5,color:C.grey,marginTop:5}}>{aiRes.classificacao} · {aiRes.prazoLegal}d</div></div>
              </div>
            </div>
            <div style={card}><div style={{fontWeight:600,fontSize:12.5,color:C.navy,marginBottom:6}}>⚖️ Normas</div><div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:6}}>{(aiRes.normas||[]).map((n,i)=><span key={i} style={bdg(C.blue,C.light)}>{n}</span>)}</div><div style={{fontSize:12,color:C.grey,lineHeight:1.5}}>{aiRes.justificativa}</div></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <div style={card}><div style={{fontWeight:600,fontSize:12,color:C.navy,marginBottom:5}}>📋 Documentos</div><ul style={{margin:0,padding:"0 0 0 14px"}}>{(aiRes.documentos||[]).map((d,i)=><li key={i} style={{fontSize:11.5,lineHeight:1.5,marginBottom:3}}>{d}</li>)}</ul></div>
              <div style={card}><div style={{fontWeight:600,fontSize:12,color:C.navy,marginBottom:5}}>⚡ Ações</div><ul style={{margin:0,padding:"0 0 0 14px"}}>{(aiRes.acoes||[]).map((a,i)=><li key={i} style={{fontSize:11.5,lineHeight:1.5,marginBottom:3}}>{a}</li>)}</ul></div>
            </div>
            {(aiRes.alertas||[]).length>0&&<div style={{...card,background:"#FFFBEB",borderColor:"#FCD34D"}}><div style={{fontWeight:600,fontSize:12,color:"#92400E",marginBottom:5}}>⚠️ Alertas</div><ul style={{margin:0,padding:"0 0 0 14px"}}>{aiRes.alertas.map((a,i)=><li key={i} style={{fontSize:11.5,color:"#92400E",lineHeight:1.5,marginBottom:3}}>{a}</li>)}</ul></div>}
            <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none"}}><div style={{fontSize:9.5,fontWeight:700,color:C.green,letterSpacing:0.5,marginBottom:5,textTransform:"uppercase"}}>Recomendação</div><p style={{color:C.white,fontSize:12.5,lineHeight:1.6,margin:0}}>{aiRes.recomendacao}</p></div>
          </div>
        )}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════
  // ── NORMAS SECTION ────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const NormasSection = () => (
    <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:12}}>
      {NORMAS.map(n=>(
        <div key={n.id} style={{...card,borderLeft:`4px solid ${n.cor}`,cursor:"pointer",transition:"all 0.15s"}}
          onClick={()=>setSelN(selN===n.id?null:n.id)}
          onMouseEnter={e=>!isMobile&&(e.currentTarget.style.boxShadow=`0 4px 16px ${n.cor}22`)}
          onMouseLeave={e=>!isMobile&&(e.currentTarget.style.boxShadow="0 1px 4px rgba(2,18,76,0.06)")}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap",gap:4}}>
            <span style={bdg(n.cor)}>{n.ref}</span>
            <span style={bdg(C.green,"#F0FDF4")}>✓ {n.dataVig}</span>
          </div>
          <div style={{fontSize:isMobile?13:14,fontWeight:700,color:C.navy,marginBottom:5}}>{n.titulo}</div>
          <div style={{fontSize:12,color:C.grey,lineHeight:1.5,marginBottom:7}}>{n.resumo}</div>
          {n.pS&&<div style={{display:"flex",gap:5,marginBottom:6}}><span style={bdg(C.green)}>{n.pS}d Simples</span>{n.pC&&<span style={bdg(C.blue)}>{n.pC}d Complexo</span>}</div>}
          <div style={{fontSize:11,color:C.grey,fontStyle:"italic"}}>{n.pen}</div>
          {n.nota&&<div style={{fontSize:11,color:C.teal,marginTop:4,fontWeight:600}}>{n.nota}</div>}
          {selN===n.id&&<div style={{marginTop:10,borderTop:`1px solid ${C.border}`,paddingTop:10}}>
            <div style={{fontSize:12,fontWeight:700,color:C.navy,marginBottom:6}}>Obrigações:</div>
            <ul style={{margin:0,padding:"0 0 0 14px"}}>{n.obr.map((o,i)=><li key={i} style={{fontSize:12,lineHeight:1.5,marginBottom:4}}>{o}</li>)}</ul>
          </div>}
        </div>
      ))}
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════
  // ── IMPORT SECTION ────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const ImportSection = () => (
    <div style={{display:"flex",flexDirection:"column",gap:14,maxWidth:isMobile?"100%":820}}>
      <div style={{...card,background:`linear-gradient(135deg,${C.navy},#1A3A8F)`,border:"none",padding:isMobile?"16px":"22px 26px"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <div style={{width:38,height:38,background:"rgba(62,117,255,0.25)",borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><FileSpreadsheet size={18} color={C.blue}/></div>
          <div>
            <div style={{fontSize:isMobile?14:16,fontWeight:700,color:C.white}}>Importar Planilha de Sinistros</div>
            <div style={{fontSize:11.5,color:"rgba(255,255,255,0.6)",marginTop:1}}>xlsx · xls · csv</div>
          </div>
        </div>
        <p style={{color:"rgba(255,255,255,0.72)",fontSize:12.5,margin:0,lineHeight:1.5}}>Arraste o RELATÓRIO_SINISTROS.xlsx — a plataforma detecta o formato automaticamente. Prazos e riscos SUSEP são calculados na importação.</p>
        {cnt>0&&!loading&&<div style={{marginTop:8,background:"rgba(44,221,124,0.15)",border:"1px solid rgba(44,221,124,0.4)",borderRadius:8,padding:"7px 12px",fontSize:12,color:C.green,fontWeight:600}}>✓ {cnt} casos da sessão anterior carregados.</div>}
      </div>
      <div onClick={()=>fRef.current?.click()} onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);handleFile(e.dataTransfer.files[0])}}
        style={{border:`2.5px dashed ${drag?C.blue:C.border}`,borderRadius:12,padding:isMobile?"28px 16px":"36px 24px",textAlign:"center",cursor:"pointer",background:drag?C.light+"44":C.bg,transition:"all 0.2s"}}>
        <div style={{width:48,height:48,margin:"0 auto 10px",background:C.blue+"18",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center"}}>
          {imp?<RefreshCw size={22} color={C.blue} style={{animation:"spin 0.9s linear infinite"}}/>:<Upload size={22} color={C.blue}/>}
        </div>
        <div style={{fontSize:isMobile?13:15,fontWeight:700,color:C.navy,marginBottom:4}}>{imp?"Processando...":isMobile?"Toque para selecionar arquivo":"Arraste ou clique para selecionar"}</div>
        <div style={{fontSize:12,color:C.grey}}>{imp?"Calculando prazos SUSEP...":".xlsx · .xls · .csv"}</div>
      </div>
      <input ref={fRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{handleFile(e.target.files[0]);e.target.value="";}}/>
      {log.length>0&&(
        <div style={{...card,padding:0,overflow:"hidden"}}>
          <div style={{padding:"8px 14px",background:C.navy,fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.65)",letterSpacing:0.5,textTransform:"uppercase"}}>Log de Importação</div>
          {log.map((l,i)=>(
            <div key={i} style={{padding:"8px 14px",borderBottom:i<log.length-1?`1px solid ${C.border}`:"none",display:"flex",alignItems:"flex-start",gap:8,background:l.t==="err"?"#FEF2F2":C.white}}>
              {l.t==="ok"?<CheckCircle size={13} color={C.green} style={{marginTop:1,flexShrink:0}}/>:<AlertCircle size={13} color={C.red} style={{marginTop:1,flexShrink:0}}/>}
              <span style={{fontSize:12,color:l.t==="err"?"#B91C1C":C.body,lineHeight:1.5}}>{l.m}</span>
            </div>
          ))}
        </div>
      )}
      {cnt>0&&log.some(l=>l.t==="ok"&&l.m.includes("importados"))&&(
        <div style={{background:"#F0FDF4",border:"1px solid #BBF7D0",borderRadius:10,padding:"12px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}><CheckSquare size={18} color={C.green}/><div><div style={{fontWeight:700,fontSize:13,color:"#14532D"}}>{cnt} casos importados!</div><div style={{fontSize:11.5,color:"#166534"}}>Dados salvos — disponíveis na próxima sessão</div></div></div>
          <button onClick={()=>setSec("dashboard")} style={btn(C.green,isMobile)}>Ver Dashboard →</button>
        </div>
      )}
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
        <button onClick={downloadTemplate} style={btn(C.navy,isMobile)}><Download size={13}/>Baixar Template Sinistros</button>
        {casos.length>0&&<button onClick={()=>{if(window.confirm("Limpar todos os casos?")) {setCasos([]);setCnt(0);try{localStorage.removeItem("umma-casos");}catch(e){}}}} style={{...btn(C.white,isMobile),color:C.red,border:`1px solid ${C.red}22`}}><Trash2 size={13}/>Limpar Casos</button>}
      </div>
    </div>
  );

  // ══════════════════════════════════════════════════════════════════════════
  // ── MAIN RENDER ───────────────────────────────────────────────────────────
  // ══════════════════════════════════════════════════════════════════════════
  const renderSection = () => {
    if (loading) return (
      <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:200,flexDirection:"column",gap:12}}>
        <div style={{width:36,height:36,border:`3px solid ${C.light}`,borderTopColor:C.blue,borderRadius:"50%",animation:"spin 0.8s linear infinite"}}/>
        <span style={{fontSize:13,color:C.grey}}>Carregando...</span>
      </div>
    );
    switch(sec) {
      case "aviso":      return <AvisoSection/>;
      case "dashboard":  return <Dashboard/>;
      case "casos":      return <CasosSection/>;
      case "apolices":   return <ApolicesSection/>;
      case "instrucoes": return <InstrucoesSection/>;
      case "ai":         return <AISection/>;
      case "normas":     return <NormasSection/>;
      case "importar":   return <ImportSection/>;
      default:           return <AvisoSection/>;
    }
  };

  return (
    <div style={{display:"flex",flexDirection:isMobile?"column":"row",height:"100vh",background:C.bg,fontFamily:font,overflow:"hidden",color:C.body}}>
      <style>{`*{box-sizing:border-box}::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:rgba(62,117,255,0.2);border-radius:3px}@keyframes spin{to{transform:rotate(360deg)}}input:focus,textarea:focus,select:focus{border-color:${C.blue}!important;outline:none}button:focus{outline:none}@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700;800&display=swap');`}</style>
      {!isMobile&&<Sidebar/>}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
        <Header/>
        <div style={{flex:1,overflow:"auto",padding:isMobile?"12px":"18px 20px",paddingBottom:isMobile?"72px":"18px",position:"relative"}}>
          {renderSection()}
        </div>
        {isMobile&&<BottomNav/>}
      </div>
      <CaseDetail/>
      <ApoliceDetail/>
      <InstDetail/>
    </div>
  );
}
