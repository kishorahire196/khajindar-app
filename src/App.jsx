import React, { useState, useEffect, useRef, useCallback } from "react";
import { storage } from "./storage";
import { IndianRupee, Plus, Share2, Receipt, Wallet, TrendingDown, X, Camera, Search, Settings, Check, ImageOff, ChevronRight, FileSpreadsheet, FileDown, Users, Crown, Eye, EyeOff, PenLine, LogOut, Lock, ShieldAlert, Copy } from "lucide-react";
import * as XLSX from "xlsx";

const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Yatra+One&family=Mukta:wght@400;500;600;700;800&family=Space+Mono:wght@400;700&display=swap');
.font-display { font-family: 'Yatra One', cursive; }
.font-body { font-family: 'Mukta', sans-serif; }
.font-num { font-family: 'Space Mono', monospace; font-variant-numeric: tabular-nums; }
.paper-texture {
  background-color: #F6EFDD;
  background-image:
    radial-gradient(circle at 1px 1px, rgba(122,32,72,0.07) 1px, transparent 0);
  background-size: 20px 20px;
}
.ledger-rules {
  background-image: repeating-linear-gradient(to bottom, transparent 0, transparent 27px, rgba(122,32,72,0.05) 27px, rgba(122,32,72,0.05) 28px);
}
.ledger-bind {
  position: relative;
  padding-left: 14px;
}
.ledger-bind::before {
  content: "";
  position: absolute;
  left: 3px; top: 4px; bottom: 4px;
  width: 0;
  border-left: 2px dashed rgba(122,32,72,0.35);
}
.receipt-notch {
  clip-path: polygon(
    0 0, 100% 0, 100% 100%,
    92% 100%, 88% 92%, 84% 100%, 80% 92%, 76% 100%, 72% 92%, 68% 100%, 64% 92%, 60% 100%,
    56% 92%, 52% 100%, 48% 92%, 44% 100%, 40% 92%, 36% 100%, 32% 92%, 28% 100%, 24% 92%,
    20% 100%, 16% 92%, 12% 100%, 8% 92%, 4% 100%, 0 92%
  );
  padding-bottom: 14px;
}
@keyframes stampIn {
  0% { transform: rotate(-14deg) scale(0.4); opacity: 0; }
  60% { transform: rotate(-11deg) scale(1.08); opacity: 1; }
  100% { transform: rotate(-12deg) scale(1); opacity: 1; }
}
.stamp-anim { animation: stampIn 0.5s ease-out; }
@media (prefers-reduced-motion: reduce) {
  .stamp-anim { animation: none; }
  * { transition-duration: 0.01ms !important; }
}
.focus-ring:focus-visible {
  outline: 2px solid #C98A2B;
  outline-offset: 2px;
}
#printArea { display: none; }
@media print {
  body * { visibility: hidden; }
  #printArea, #printArea * { visibility: visible; }
  #printArea { display: block; position: absolute; left: 0; top: 0; width: 100%; padding: 24px; }
}
`;

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const fmtDate = (d) => new Date(d).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
const fmtAmt = (n) => Number(n || 0).toLocaleString("en-IN");

const ONES99 = [
  "शून्य","एक","दोन","तीन","चार","पाच","सहा","सात","आठ","नऊ",
  "दहा","अकरा","बारा","तेरा","चौदा","पंधरा","सोळा","सतरा","अठरा","एकोणीस",
  "वीस","एकवीस","बावीस","तेवीस","चोवीस","पंचवीस","सव्वीस","सत्तावीस","अठ्ठावीस","एकोणतीस",
  "तीस","एकतीस","बत्तीस","तेहेतीस","चौतीस","पस्तीस","छत्तीस","सदतीस","अडतीस","एकोणचाळीस",
  "चाळीस","एक्केचाळीस","बेचाळीस","त्रेचाळीस","चव्वेचाळीस","पंचेचाळीस","शेहेचाळीस","सत्तेचाळीस","अठ्ठेचाळीस","एकोणपन्नास",
  "पन्नास","एक्कावन्न","बावन्न","त्रेपन्न","चोपन्न","पंचावन्न","छप्पन्न","सत्तावन्न","अठ्ठावन्न","एकोणसाठ",
  "साठ","एकसष्ठ","बासष्ठ","त्रेसष्ठ","चौसष्ठ","पासष्ठ","सहासष्ठ","सदुसष्ठ","अडुसष्ठ","एकोणसत्तर",
  "सत्तर","एक्काहत्तर","बहात्तर","त्र्याहत्तर","चौर्‍याहत्तर","पंच्याहत्तर","शहात्तर","सत्याहत्तर","अठ्ठ्याहत्तर","एकोणऐंशी",
  "ऐंशी","एक्क्याऐंशी","ब्याऐंशी","त्र्याऐंशी","चौऱ्याऐंशी","पंच्याऐंशी","शहाऐंशी","सत्त्याऐंशी","अठ्ठ्याऐंशी","एकोणनव्वद",
  "नव्वद","एक्क्याण्णव","ब्याण्णव","त्र्याण्णव","चौऱ्याण्णव","पंच्याण्णव","शहाण्णव","सत्त्याण्णव","अठ्ठ्याण्णव","नव्याण्णव",
];

function numberToMarathiWords(num) {
  let n = Math.round(Number(num) || 0);
  if (n === 0) return "शून्य";
  if (n < 0) return "उणे " + numberToMarathiWords(-n);
  const crore = Math.floor(n / 10000000); n %= 10000000;
  const lakh = Math.floor(n / 100000); n %= 100000;
  const thousand = Math.floor(n / 1000); n %= 1000;
  const hundred = Math.floor(n / 100); n %= 100;
  const rest = n;
  const parts = [];
  if (crore) parts.push(ONES99[crore] + " कोटी");
  if (lakh) parts.push(ONES99[lakh] + " लाख");
  if (thousand) parts.push(ONES99[thousand] + " हजार");
  if (hundred) parts.push(hundred === 1 ? "शंभर" : ONES99[hundred] + "शे");
  if (rest) parts.push(ONES99[rest]);
  return parts.join(" ");
}

async function loadCanvasFonts() {
  try {
    await Promise.all([
      document.fonts.load('700 42px "Yatra One"'),
      document.fonts.load('italic 800 34px "Yatra One"'),
      document.fonts.load('600 20px "Mukta"'),
      document.fonts.load('700 20px "Mukta"'),
      document.fonts.load('700 30px "Space Mono"'),
    ]);
    await document.fonts.ready;
  } catch (e) {}
}

function drawMedallion(ctx, cx, cy, r, color) {
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = 2;
  ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.stroke();
  ctx.beginPath(); ctx.arc(cx, cy, r - 6, 0, Math.PI * 2); ctx.stroke();
  const rays = 16;
  for (let i = 0; i < rays; i++) {
    const a = (i / rays) * Math.PI * 2;
    const x1 = cx + Math.cos(a) * (r - 10);
    const y1 = cy + Math.sin(a) * (r - 10);
    const x2 = cx + Math.cos(a) * (r + 4);
    const y2 = cy + Math.sin(a) * (r + 4);
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
  }
  ctx.beginPath(); ctx.arc(cx, cy, r - 16, 0, Math.PI * 2);
  ctx.fillStyle = color; ctx.globalAlpha = 0.15; ctx.fill(); ctx.globalAlpha = 1;
  ctx.restore();
}

async function drawReceiptCanvas(canvas, entry, settings) {
  await loadCanvasFonts();
  const W = 1500, H = 560;
  canvas.width = W; canvas.height = H;
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#FBF6EC";
  ctx.fillRect(0, 0, W, H);

  ctx.save();
  ctx.globalAlpha = 0.06;
  ctx.fillStyle = "#7A2048";
  ctx.strokeStyle = "#7A2048";
  // simple diya (lamp) silhouette watermark — generic across festivals
  ctx.save();
  ctx.translate(W / 2, H / 2 + 40);
  ctx.scale(3.2, 3.2);
  ctx.beginPath();
  ctx.moveTo(-70, 10);
  ctx.quadraticCurveTo(0, 55, 70, 10);
  ctx.quadraticCurveTo(40, 30, 0, 32);
  ctx.quadraticCurveTo(-40, 30, -70, 10);
  ctx.fill();
  ctx.beginPath();
  ctx.moveTo(0, -8);
  ctx.quadraticCurveTo(14, -30, 0, -48);
  ctx.quadraticCurveTo(-14, -30, 0, -8);
  ctx.fill();
  ctx.restore();
  ctx.restore();

  ctx.strokeStyle = "#7A2048";
  ctx.lineWidth = 6;
  ctx.strokeRect(16, 16, W - 32, H - 32);
  ctx.strokeStyle = "#C98A2B";
  ctx.lineWidth = 2;
  ctx.strokeRect(28, 28, W - 56, H - 56);

  ctx.fillStyle = "#7A2048";
  ctx.font = "600 20px 'Mukta'";
  ctx.textAlign = "left";
  ctx.textBaseline = "alphabetic";
  ctx.fillText("|| शुभम् भवतु ||", 50, 55);

  const bannerY = 68, bannerH = 78;
  ctx.fillStyle = "#7A2048";
  ctx.fillRect(50, bannerY, W - 100, bannerH);
  ctx.strokeStyle = "#C98A2B";
  ctx.lineWidth = 2;
  ctx.strokeRect(50, bannerY, W - 100, bannerH);

  drawMedallion(ctx, 110, bannerY + bannerH / 2, 30, "#F6EFDD");
  drawMedallion(ctx, W - 110, bannerY + bannerH / 2, 30, "#F6EFDD");

  ctx.fillStyle = "#F6EFDD";
  ctx.font = "700 42px 'Yatra One'";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(settings.mandalName, W / 2, bannerY + bannerH / 2 + 4);

  let y = bannerY + bannerH + 34;
  if (settings.place) {
    ctx.fillStyle = "#7A2048";
    ctx.font = "600 18px 'Mukta'";
    ctx.textAlign = "center";
    ctx.textBaseline = "alphabetic";
    ctx.fillText(settings.place, W / 2, y);
    y += 22;
  }

  ctx.fillStyle = "#7A2048";
  ctx.font = "700 26px 'Yatra One'";
  ctx.textAlign = "center";
  ctx.fillText("{ वर्गणी पावती }", W / 2, y + 16);
  y += 58;

  ctx.font = "600 20px 'Mukta'";
  ctx.fillStyle = "#3A2E22";
  ctx.textAlign = "left";
  ctx.fillText(`पावती क्र : ${entry.receiptNo}`, 60, y);
  ctx.textAlign = "right";
  ctx.fillText(`दिनांक : ${fmtDate(entry.date)}`, W - 60, y);
  y += 45;

  ctx.textAlign = "left";
  ctx.font = "600 20px 'Mukta'";
  const label1 = "श्री./श्रीमती/मेसर्स  ";
  ctx.fillText(label1, 60, y);
  const label1W = ctx.measureText(label1).width;
  ctx.setLineDash([4, 4]);
  ctx.strokeStyle = "#7A2048";
  ctx.beginPath(); ctx.moveTo(60 + label1W, y + 4); ctx.lineTo(W - 250, y + 4); ctx.stroke();
  ctx.setLineDash([]);
  ctx.font = "700 20px 'Mukta'";
  ctx.fillText(entry.name, 60 + label1W + 10, y);
  ctx.font = "600 20px 'Mukta'";
  ctx.textAlign = "right";
  ctx.fillText("यांजकडून", W - 60, y);
  y += 45;

  ctx.textAlign = "left";
  ctx.font = "600 20px 'Mukta'";
  ctx.fillText(`सार्वजनिक उत्सव वर्गणी${entry.note ? " (" + entry.note + ")" : ""} करिता`, 60, y);
  y += 45;

  ctx.font = "600 20px 'Mukta'";
  const label2 = "अक्षरी रुपये : ";
  ctx.fillText(label2, 60, y);
  const label2W = ctx.measureText(label2).width;
  ctx.font = "700 20px 'Mukta'";
  ctx.fillText(`${numberToMarathiWords(entry.amount)} रुपये फक्त`, 60 + label2W + 6, y);
  y += 55;

  ctx.strokeStyle = "#7A2048";
  ctx.lineWidth = 2;
  ctx.strokeRect(60, y, 220, 60);
  ctx.font = "700 30px 'Space Mono'";
  ctx.fillStyle = "#7A2048";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(`₹ ${fmtAmt(entry.amount)}`, 60 + 110, y + 32);

  ctx.textAlign = "left";
  ctx.textBaseline = "alphabetic";
  ctx.font = "600 18px 'Mukta'";
  ctx.fillStyle = "#3A2E22";
  ctx.fillText("रोख / ऑनलाईन स्वरूपात मिळाले.", 310, y + 25);
  ctx.fillText("वसूल करणार : खजिनदार", 310, y + 52);

  ctx.save();
  ctx.translate(W - 170, y + 38);
  ctx.rotate(-0.08);
  ctx.fillStyle = "#B23B3B";
  ctx.font = "italic 800 34px 'Yatra One'";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("धन्यवाद!", 0, 0);
  ctx.restore();
}

function compressImage(file, maxDim = 900, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDim) { height = (height * maxDim) / width; width = maxDim; }
        else if (height > maxDim) { width = (width * maxDim) / height; height = maxDim; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const DEFAULT_SETTINGS = { mandalName: "आपले उत्सव मंडळ", place: "", receiptStart: 101, setupDone: false };

const ACCOUNTS = [
  { id: "president", label: "अध्यक्ष", group: "officer", icon: "Crown" },
  { id: "treasurer", label: "खजिनदार", group: "officer", icon: "Wallet" },
  { id: "secretary", label: "सेक्रेटरी", group: "officer", icon: "PenLine" },
  { id: "viewer1", label: "कार्यकर्ता 1", group: "viewer" },
  { id: "viewer2", label: "कार्यकर्ता 2", group: "viewer" },
  { id: "viewer3", label: "कार्यकर्ता 3", group: "viewer" },
  { id: "viewer4", label: "कार्यकर्ता 4", group: "viewer" },
  { id: "viewer5", label: "कार्यकर्ता 5", group: "viewer" },
];
const ACCOUNT_ICONS = { Crown, Wallet, PenLine };
const SESSION_STALE_MS = 90 * 1000;
const HEARTBEAT_MS = 25 * 1000;
const newSid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

export default function KhajindarApp() {
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [chanda, setChanda] = useState([]);
  const [kharch, setKharch] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [account, setAccount] = useState(null); // account id from ACCOUNTS, or null
  const [kickedOut, setKickedOut] = useState(false);
  const sidRef = useRef(null);
  const [tab, setTab] = useState("home"); // home | chanda | kharch | list
  const [showChandaForm, setShowChandaForm] = useState(false);
  const [showKharchForm, setShowKharchForm] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [receiptToShow, setReceiptToShow] = useState(null);
  const [search, setSearch] = useState("");
  const [proofView, setProofView] = useState(null);
  const [saveErr, setSaveErr] = useState(false);
  const [excelBusy, setExcelBusy] = useState(false);
  const [excelMsg, setExcelMsg] = useState("");

  // load
  useEffect(() => {
    (async () => {
      try {
        const s = await storage.get("settings", true);
        if (s?.value) setSettings(JSON.parse(s.value));
      } catch (e) {}
      try {
        const t = await storage.get("transactions", true);
        if (t?.value) {
          const parsed = JSON.parse(t.value);
          setChanda(parsed.chanda || []);
          setKharch(parsed.kharch || []);
        }
      } catch (e) {}
      try {
        const local = await storage.get("session", false);
        if (local?.value) {
          const saved = JSON.parse(local.value);
          const sessions = await getSessions();
          const entry = sessions[saved.account];
          if (entry && entry.sid === saved.sid) {
            sidRef.current = saved.sid;
            setAccount(saved.account);
          } else {
            await storage.set("session", "", false).catch(() => {});
          }
        }
      } catch (e) {}
      setLoaded(true);
    })();
  }, []);

  const currentAccountInfo = ACCOUNTS.find((a) => a.id === account);
  const isOwner = currentAccountInfo?.group === "officer";

  // heartbeat: refresh our session, detect being kicked out by another login
  useEffect(() => {
    if (!account || !sidRef.current) return;
    let stopped = false;
    async function beat() {
      try {
        const sessions = await getSessions();
        const entry = sessions[account];
        if (entry && entry.sid !== sidRef.current) {
          if (!stopped) { setAccount(null); setKickedOut(true); }
          try { await storage.set("session", "", false); } catch (e) {}
          return;
        }
        sessions[account] = { sid: sidRef.current, ts: Date.now() };
        await storage.set("sessions", JSON.stringify(sessions), true);
      } catch (e) {}
    }
    beat();
    const id = setInterval(beat, HEARTBEAT_MS);
    return () => { stopped = true; clearInterval(id); };
  }, [account]);

  async function loginAccount(accId, sid) {
    sidRef.current = sid;
    setAccount(accId);
    setKickedOut(false);
    try { await storage.set("session", JSON.stringify({ account: accId, sid }), false); } catch (e) {}
  }

  async function logoutAccount() {
    try {
      const sessions = await getSessions();
      if (sessions[account]?.sid === sidRef.current) {
        delete sessions[account];
        await storage.set("sessions", JSON.stringify(sessions), true);
      }
    } catch (e) {}
    try { await storage.set("session", "", false); } catch (e) {}
    sidRef.current = null;
    setAccount(null);
  }

  async function getSessions() {
    try {
      const remote = await storage.get("sessions", true);
      return remote?.value ? JSON.parse(remote.value) : {};
    } catch (e) {
      // key doesn't exist yet (first-ever login) — treat as no active sessions
      return {};
    }
  }

  async function claimSession(accId, force = false) {
    let sessions = {};
    try {
      sessions = await getSessions();
    } catch (e) {}
    const entry = sessions[accId];
    const isFresh = entry && Date.now() - entry.ts < SESSION_STALE_MS;
    if (isFresh && !force) return { ok: false, reason: "locked" };
    const sid = newSid();
    sessions[accId] = { sid, ts: Date.now() };
    try {
      await storage.set("sessions", JSON.stringify(sessions), true);
    } catch (e) {
      // best-effort: a storage hiccup here shouldn't block a correctly-authenticated login
    }
    await loginAccount(accId, sid);
    return { ok: true };
  }

  async function attemptLogin(accId, pwd, force = false) {
    const passwords = settings.accountPasswords || {};
    if (passwords[accId] !== pwd) return { ok: false, reason: "wrong" };
    return claimSession(accId, force);
  }

  async function createPasswordAndLogin(accId, pwd) {
    const passwords = { ...(settings.accountPasswords || {}), [accId]: pwd };
    const next = { ...settings, accountPasswords: passwords };
    setSettings(next);
    await persistSettings(next);
    // password was just created — no need to re-verify it against (possibly stale) state
    return claimSession(accId, false);
  }

  async function changeOwnPassword(newPwd) {
    if (!account) return;
    const passwords = { ...(settings.accountPasswords || {}), [account]: newPwd };
    const next = { ...settings, accountPasswords: passwords };
    setSettings(next);
    await persistSettings(next);
  }

  const persistTransactions = useCallback(async (nextChanda, nextKharch) => {
    try {
      const res = await storage.set("transactions", JSON.stringify({ chanda: nextChanda, kharch: nextKharch }), true);
      if (!res) setSaveErr(true); else setSaveErr(false);
    } catch (e) { setSaveErr(true); }
  }, []);

  const persistSettings = useCallback(async (next) => {
    try { await storage.set("settings", JSON.stringify(next), true); } catch (e) {}
  }, []);

  const totalChanda = chanda.reduce((s, c) => s + Number(c.amount || 0), 0);
  const totalKharch = kharch.reduce((s, k) => s + Number(k.amount || 0), 0);
  const balance = totalChanda - totalKharch;

  function addChanda(entry) {
    const receiptNo = (settings.receiptStart || 101) + chanda.length;
    const newEntry = { id: uid(), ...entry, receiptNo, createdAt: Date.now() };
    const next = [newEntry, ...chanda];
    setChanda(next);
    persistTransactions(next, kharch);
    setShowChandaForm(false);
    setReceiptToShow(newEntry);
  }

  function addKharch(entry) {
    const next = [{ id: uid(), ...entry, createdAt: Date.now() }, ...kharch];
    setKharch(next);
    persistTransactions(chanda, next);
    setShowKharchForm(false);
  }

  function shareChandaWhatsapp(entry) {
    const msg =
      `🪔 *${settings.mandalName}*\n` +
      (settings.place ? `${settings.place}\n` : "") +
      `\n*पावती क्र:* ${entry.receiptNo}\n` +
      `*नाव:* ${entry.name}\n` +
      `*रक्कम:* ₹${fmtAmt(entry.amount)}\n` +
      `*दिनांक:* ${fmtDate(entry.date)}\n` +
      (entry.note ? `*तपशील:* ${entry.note}\n` : "") +
      `\nधन्यवाद! आपल्या योगदानाबद्दल मंडळ आभारी आहे. 🙏`;
    const phone = (entry.phone || "").replace(/\D/g, "");
    const url = `https://wa.me/${phone ? (phone.length === 10 ? "91" + phone : phone) : ""}?text=${encodeURIComponent(msg)}`;
    window.open(url, "_blank");
  }

  async function exportExcel() {
    setExcelBusy(true);
    setExcelMsg("");
    const chandaRows = chanda.map(c => ({
      "प्रकार": "वर्गणी", "पावती क्र": c.receiptNo, "नाव": c.name, "फोन": c.phone || "",
      "रक्कम": c.amount, "दिनांक": fmtDate(c.date), "तपशील": c.note || "",
    }));
    const kharchRows = kharch.map(k => ({
      "प्रकार": "खर्च", "पावती क्र": "", "नाव": k.desc, "फोन": "",
      "रक्कम": -k.amount, "दिनांक": fmtDate(k.date), "तपशील": k.proof ? "प्रूफ जोडलेला" : "",
    }));
    const all = [...chandaRows, ...kharchRows].sort((a, b) => new Date(b["दिनांक"]) - new Date(a["दिनांक"]));
    const summary = [
      { "तपशील": "एकूण जमा", "रक्कम": totalChanda },
      { "तपशील": "एकूण खर्च", "रक्कम": totalKharch },
      { "तपशील": "शिल्लक", "रक्कम": balance },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(all), "सर्व व्यवहार");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "सारांश");
    const filename = `${settings.mandalName}-hisab.xlsx`;

    // try the share sheet first — works even when the sandboxed preview blocks downloads
    try {
      const arrayBuf = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([arrayBuf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const file = new File([blob], filename, { type: blob.type });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: filename });
        setExcelBusy(false);
        return;
      }
    } catch (e) {
      if (e?.name === "AbortError") { setExcelBusy(false); return; }
    }

    // fall back to a direct download trigger
    try {
      XLSX.writeFile(wb, filename);
    } catch (e) {}
    setExcelBusy(false);
    setExcelMsg("Share/download पर्याय दिसला नाही तर लॅपटॉप/computer वर हाच artifact उघडून पुन्हा प्रयत्न करा.");
  }

  function exportPDF() {
    window.print();
  }

  function shareHisabWhatsapp() {
    const recentK = kharch.slice(0, 5).map(k => `  • ${k.desc} — ₹${fmtAmt(k.amount)} (${fmtDate(k.date)})`).join("\n");
    const msg =
      `📒 *${settings.mandalName} — हिशोब*\n\n` +
      `एकूण जमा: ₹${fmtAmt(totalChanda)}\n` +
      `एकूण खर्च: ₹${fmtAmt(totalKharch)}\n` +
      `शिल्लक: ₹${fmtAmt(balance)}\n` +
      (kharch.length ? `\n*अलीकडील खर्च:*\n${recentK}` : "") +
      `\n\n— पारदर्शक हिशोब, खजिनदार`;
    window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, "_blank");
  }

  if (!loaded) {
    return <div className="min-h-screen paper-texture flex items-center justify-center font-body text-[#7A2048]">लोड होत आहे…</div>;
  }

  if (!settings.setupDone) {
    return (
      <MandalSetup
        settings={settings}
        onSave={(s) => { const next = { ...s, setupDone: true }; setSettings(next); persistSettings(next); }}
      />
    );
  }

  if (!account) {
    return (
      <RoleGate
        mandalName={settings.mandalName}
        settings={settings}
        onCreatePassword={createPasswordAndLogin}
        onAttemptLogin={attemptLogin}
        kickedOut={kickedOut}
      />
    );
  }

  return (
    <div className="min-h-screen paper-texture font-body pb-24 relative" style={{ color: "#3A2E22" }}>
      <style>{FONT_STYLE}</style>
      <FestiveWatermark />

      {/* Header */}
      <div style={{ background: "linear-gradient(160deg,#FDEFDC 0%,#F8E3C8 100%)", borderColor: "rgba(122,32,72,0.1)" }} className="ledger-rules px-5 pt-6 pb-5 receipt-notch shadow-sm relative overflow-hidden border-b z-10">
        <div className="absolute -right-8 -top-8 w-36 h-36 rounded-full" style={{ background: "rgba(201,138,43,0.12)" }} />
        <div className="absolute right-12 bottom-4 w-16 h-16 rounded-full" style={{ background: "rgba(122,32,72,0.06)" }} />
        <div className="flex items-start justify-between relative z-10">
          <div className="flex items-start gap-3">
            <MonogramBadge letter={(settings.mandalName || "उ").charAt(0)} size={52} />
            <div>
              <p className="text-[11px] tracking-[0.2em] uppercase opacity-60 font-body font-semibold" style={{ color: "#7A2048" }}>खजिनदार वही</p>
              <h1 className="font-display leading-tight mt-1" style={{ fontSize: "clamp(1.35rem, 5vw, 1.75rem)", color: "#7A2048" }}>{settings.mandalName}</h1>
              {settings.place ? <p className="text-xs opacity-60 mt-0.5">{settings.place}</p> : null}
              <span
                className="inline-flex items-center gap-1 mt-2 px-2.5 py-1 rounded-full text-[11px] font-medium"
                style={isOwner ? { background: "#FBE9C6", color: "#8A5A0E" } : { background: "#E4EEF1", color: "#3D6B75" }}
              >
                {(() => { const Icon = ACCOUNT_ICONS[currentAccountInfo?.icon] || Eye; return <Icon size={12} />; })()}
                {currentAccountInfo?.label || "पाहणारा"}
              </span>
            </div>
          </div>
          <button onClick={() => setShowSettings(true)} className="focus-ring p-2.5 rounded-full bg-white/90 active:bg-white active:scale-90 transition-transform" style={{ color: "#7A2048" }} aria-label="सेटिंग्ज">
            <Settings size={18} />
          </button>
        </div>

        <div className="mt-6 relative z-10">
          <p className="text-xs opacity-60 tracking-wide" style={{ color: "#7A2048" }}>शिल्लक रक्कम</p>
          <p className="font-num text-[2.6rem] leading-none font-bold mt-1.5 flex items-center gap-0.5" style={{ color: "#7A2048" }}>
            <IndianRupee size={28} strokeWidth={2.75} />{fmtAmt(balance)}
          </p>
          <div className="flex gap-3 mt-5">
            <div className="ledger-bind flex-1 bg-white/90 rounded-xl px-3 py-2.5">
              <p className="text-[10px] opacity-60 tracking-wide" style={{ color: "#7A2048" }}>एकूण वर्गणी</p>
              <p className="font-num text-lg font-semibold" style={{ color: "#4C8058" }}>₹{fmtAmt(totalChanda)}</p>
            </div>
            <div className="ledger-bind flex-1 bg-white/90 rounded-xl px-3 py-2.5">
              <p className="text-[10px] opacity-60 tracking-wide" style={{ color: "#7A2048" }}>एकूण खर्च</p>
              <p className="font-num text-lg font-semibold" style={{ color: "#C15757" }}>₹{fmtAmt(totalKharch)}</p>
            </div>
          </div>
        </div>
      </div>

      {saveErr && (
        <div className="mx-5 mt-3 text-xs bg-red-50 text-red-700 border border-red-200 rounded-lg px-3 py-2">
          जतन करताना अडचण आली — पुन्हा प्रयत्न करा.
        </div>
      )}

      {/* Action buttons */}
      <div className="relative z-10">
      {isOwner && (
        <div className="px-5 mt-6 grid grid-cols-2 gap-3">
          <button
            onClick={() => setShowChandaForm(true)}
            className="focus-ring flex items-center justify-center gap-2 py-3.5 rounded-2xl font-semibold shadow-sm border active:scale-[0.97] transition-transform"
            style={{ background: "#E4F1E6", color: "#3C6B45", borderColor: "#C7E2CB" }}
          >
            <Plus size={18} /> वर्गणी जमा
          </button>
          <button
            onClick={() => setShowKharchForm(true)}
            className="focus-ring flex items-center justify-center gap-2 py-3.5 rounded-2xl font-semibold shadow-sm border active:scale-[0.97] transition-transform"
            style={{ background: "#FBE4E4", color: "#A83E3E", borderColor: "#F3C7C7" }}
          >
            <TrendingDown size={18} /> खर्च नोंद
          </button>
        </div>
      )}

      <button onClick={shareHisabWhatsapp} className="focus-ring mx-5 mt-3 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-medium border active:scale-[0.98] transition-transform" style={{ borderColor: "#D9A24B", color: "#7A2048", background: "#FDF8ED" }}>
        <Share2 size={15} /> संपूर्ण हिशोब WhatsApp वर पाठवा
      </button>

      <div className="mx-5 mt-2 grid grid-cols-2 gap-2">
        <button onClick={exportExcel} disabled={excelBusy} className="focus-ring flex items-center justify-center gap-1.5 py-2.5 rounded-xl text-sm font-medium border active:scale-[0.98] transition-transform disabled:opacity-60" style={{ borderColor: "#C7E2CB", color: "#3C6B45", background: "#FDF8ED" }}>
          <FileSpreadsheet size={15} /> {excelBusy ? "तयार होत आहे…" : "Excel"}
        </button>
        <button onClick={exportPDF} className="focus-ring flex items-center justify-center gap-1.5 py-2.5 rounded-xl text-sm font-medium border active:scale-[0.98] transition-transform" style={{ borderColor: "#F3C7C7", color: "#A83E3E", background: "#FDF8ED" }}>
          <FileDown size={15} /> PDF
        </button>
      </div>
      {excelMsg && (
        <p className="mx-5 mt-2 text-xs leading-relaxed" style={{ color: "#7A2048" }}>{excelMsg}</p>
      )}

      <div className="mx-5 mt-2.5 flex items-center gap-1.5 text-[11px] opacity-55">
        <Users size={12} /> हा हिशोब सर्वांना दिसतो — फक्त पदाधिकारी (अध्यक्ष/खजिनदार/सेक्रेटरी) बदलू शकतात
      </div>

      {/* List tabs */}
      <div className="px-5 mt-6">
        <div className="flex gap-2 mb-3">
          {[["home", "सर्व"], ["chanda", "वर्गणी"], ["kharch", "खर्च"]].map(([key, label]) => (
            <button
              key={key}
              onClick={() => setTab(key)}
              className="focus-ring px-4 py-1.5 rounded-full text-sm font-medium tracking-wide transition-colors"
              style={tab === key ? { background: "#F3DFC0", color: "#7A2048" } : { background: "rgba(122,32,72,0.06)", color: "#7A2048" }}
            >
              {label}
            </button>
          ))}
        </div>

        <div className="relative mb-3">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 opacity-50" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="नाव किंवा तपशील शोधा…"
            className="focus-ring w-full pl-9 pr-3 py-2.5 rounded-xl text-sm bg-white border outline-none"
            style={{ borderColor: "rgba(122,32,72,0.15)" }}
          />
        </div>

        <RecordsList
          tab={tab}
          chanda={chanda}
          kharch={kharch}
          search={search}
          onOpenReceipt={setReceiptToShow}
          onOpenProof={setProofView}
        />
      </div>

      </div>

      <div id="printArea">
        <h1 style={{ fontSize: 20, fontWeight: 700 }}>{settings.mandalName}</h1>
        {settings.place ? <p style={{ fontSize: 12, opacity: 0.7 }}>{settings.place}</p> : null}
        <p style={{ fontSize: 12, margin: "8px 0" }}>हिशोब अहवाल — {fmtDate(new Date())}</p>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11, marginBottom: 12 }}>
          <tbody>
            <tr><td style={{ padding: 4 }}>एकूण जमा</td><td style={{ padding: 4, textAlign: "right" }}>₹{fmtAmt(totalChanda)}</td></tr>
            <tr><td style={{ padding: 4 }}>एकूण खर्च</td><td style={{ padding: 4, textAlign: "right" }}>₹{fmtAmt(totalKharch)}</td></tr>
            <tr><td style={{ padding: 4, fontWeight: 700 }}>शिल्लक</td><td style={{ padding: 4, textAlign: "right", fontWeight: 700 }}>₹{fmtAmt(balance)}</td></tr>
          </tbody>
        </table>
        <h3 style={{ fontSize: 13, marginTop: 12 }}>वर्गणी जमा</h3>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
          <thead><tr style={{ borderBottom: "1px solid #999" }}>
            <th style={{ textAlign: "left", padding: 3 }}>पावती #</th><th style={{ textAlign: "left", padding: 3 }}>नाव</th>
            <th style={{ textAlign: "left", padding: 3 }}>दिनांक</th><th style={{ textAlign: "right", padding: 3 }}>रक्कम</th>
          </tr></thead>
          <tbody>
            {chanda.map(c => (
              <tr key={c.id} style={{ borderBottom: "1px solid #eee" }}>
                <td style={{ padding: 3 }}>{c.receiptNo}</td><td style={{ padding: 3 }}>{c.name}</td>
                <td style={{ padding: 3 }}>{fmtDate(c.date)}</td><td style={{ padding: 3, textAlign: "right" }}>₹{fmtAmt(c.amount)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        <h3 style={{ fontSize: 13, marginTop: 12 }}>खर्च</h3>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
          <thead><tr style={{ borderBottom: "1px solid #999" }}>
            <th style={{ textAlign: "left", padding: 3 }}>तपशील</th><th style={{ textAlign: "left", padding: 3 }}>दिनांक</th>
            <th style={{ textAlign: "right", padding: 3 }}>रक्कम</th>
          </tr></thead>
          <tbody>
            {kharch.map(k => (
              <tr key={k.id} style={{ borderBottom: "1px solid #eee" }}>
                <td style={{ padding: 3 }}>{k.desc}</td><td style={{ padding: 3 }}>{fmtDate(k.date)}</td>
                <td style={{ padding: 3, textAlign: "right" }}>₹{fmtAmt(k.amount)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {showChandaForm && isOwner && <ChandaForm onClose={() => setShowChandaForm(false)} onSave={addChanda} nextReceiptNo={(settings.receiptStart || 101) + chanda.length} />}
      {showKharchForm && isOwner && <KharchForm onClose={() => setShowKharchForm(false)} onSave={addKharch} />}
      {showSettings && (
        <SettingsSheet
          settings={settings}
          isOwner={isOwner}
          accountLabel={currentAccountInfo?.label}
          onClose={() => setShowSettings(false)}
          onSave={(s) => { setSettings(s); persistSettings(s); setShowSettings(false); }}
          onChangePassword={changeOwnPassword}
          onLogout={() => { setShowSettings(false); logoutAccount(); }}
        />
      )}
      {receiptToShow && (
        <ReceiptModal entry={receiptToShow} settings={settings} onClose={() => setReceiptToShow(null)} onShare={() => shareChandaWhatsapp(receiptToShow)} />
      )}
      {proofView && <ProofModal src={proofView} onClose={() => setProofView(null)} />}
    </div>
  );
}

function MonogramBadge({ letter, size = 64 }) {
  return (
    <div
      className="shrink-0 flex items-center justify-center rounded-full font-body font-bold"
      style={{
        width: size,
        height: size,
        background: "linear-gradient(135deg, #8C3A5E, #5C1836)",
        border: "1.5px solid #C98A2B",
        color: "#F3DFC0",
        fontSize: size * 0.44,
        lineHeight: 1,
        boxShadow: "0 2px 6px rgba(122,32,72,0.3)",
      }}
    >
      {letter}
    </div>
  );
}

function FestiveWatermark() {
  return (
    <svg
      viewBox="0 0 400 400"
      className="fixed pointer-events-none select-none"
      style={{ right: "-90px", bottom: "-90px", width: "min(64vw, 320px)", height: "auto", opacity: 0.05, zIndex: 0 }}
      fill="none"
      stroke="#7A2048"
    >
      <circle cx="200" cy="200" r="188" strokeWidth="1.5" />
      <circle cx="200" cy="200" r="158" strokeWidth="1" />
      <path d="M200 118 C217 145 217 167 200 184 C183 167 183 145 200 118 Z" strokeWidth="2.5" />
      <path d="M148 218 Q200 244 252 218" strokeWidth="2.5" strokeLinecap="round" />
    </svg>
  );
}

function PasswordInput({ value, onChange, placeholder, autoFocus, onKeyDown, borderColor, center }) {
  const [show, setShow] = useState(false);
  return (
    <div className="relative">
      <Lock size={14} className="absolute left-3.5 top-1/2 -translate-y-1/2 opacity-40" />
      <input
        type={show ? "text" : "password"}
        value={value}
        autoFocus={autoFocus}
        onChange={onChange}
        onKeyDown={onKeyDown}
        placeholder={placeholder}
        className={`focus-ring w-full pl-9 pr-10 py-2.5 rounded-xl bg-white border outline-none text-sm ${center ? "text-center" : ""}`}
        style={{ borderColor: borderColor || "rgba(122,32,72,0.2)" }}
      />
      <button
        type="button"
        onClick={() => setShow((s) => !s)}
        className="focus-ring absolute right-2.5 top-1/2 -translate-y-1/2 p-1 rounded-full active:scale-90 transition-transform"
        style={{ color: "#7A2048" }}
        aria-label={show ? "पासवर्ड लपवा" : "पासवर्ड दाखवा"}
      >
        {show ? <EyeOff size={15} /> : <Eye size={15} />}
      </button>
    </div>
  );
}

function MandalSetup({ settings, onSave }) {
  const [mandalName, setMandalName] = useState(settings.mandalName === DEFAULT_SETTINGS.mandalName ? "" : settings.mandalName);
  const [place, setPlace] = useState(settings.place || "");
  const [err, setErr] = useState("");

  function submit() {
    if (!mandalName.trim()) { setErr("मंडळाचं नाव टाका"); return; }
    onSave({ ...settings, mandalName: mandalName.trim(), place: place.trim() });
  }

  return (
    <div className="min-h-screen paper-texture font-body flex flex-col items-center justify-center px-6 py-10 text-center relative">
      <style>{FONT_STYLE}</style>
      <FestiveWatermark />
      <MonogramBadge letter={(mandalName || "उ").charAt(0)} size={92} />
      <h1 className="font-display text-2xl mt-3 relative z-10" style={{ color: "#7A2048" }}>स्वागत आहे</h1>

      <div className="relative z-10 w-full max-w-sm mt-6 text-left">
        <Field label="मंडळाचे नाव *">
          <input
            className={inputClass} style={inputStyle}
            value={mandalName}
            onChange={(e) => { setMandalName(e.target.value); setErr(""); }}
            placeholder="उदा. श्री गणेश मित्र मंडळ"
            autoFocus
          />
        </Field>
        <Field label="गाव / ठिकाण (ऐच्छिक)">
          <input className={inputClass} style={inputStyle} value={place} onChange={(e) => setPlace(e.target.value)} placeholder="उदा. ठाणे" />
        </Field>
        {err && <p className="text-xs mt-1 mb-2" style={{ color: "#A83E3E" }}>{err}</p>}
        <button
          onClick={submit}
          className="focus-ring w-full mt-3 py-3 rounded-xl font-semibold border active:scale-[0.98] transition-transform"
          style={{ background: "#FBE9C6", color: "#7A2048", borderColor: "#EFD8A0" }}
        >
          पुढे जा
        </button>
        <p className="text-[11px] opacity-45 mt-4 text-center">ही माहिती नंतर सेटिंग्जमधून कधीही बदलता येते</p>
      </div>
    </div>
  );
}

function RoleGate({ mandalName, settings, onCreatePassword, onAttemptLogin, kickedOut }) {
  const [step, setStep] = useState("pick"); // pick | setPassword | enterPassword
  const [accId, setAccId] = useState(null);
  const [pwd, setPwd] = useState("");
  const [pwd2, setPwd2] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  const [showKicked, setShowKicked] = useState(kickedOut);

  const passwords = settings.accountPasswords || {};
  const officers = ACCOUNTS.filter((a) => a.group === "officer");
  const viewers = ACCOUNTS.filter((a) => a.group === "viewer");
  const acc = ACCOUNTS.find((a) => a.id === accId);

  function pick(id) {
    setErr(""); setPwd(""); setPwd2(""); setShowKicked(false);
    setAccId(id);
    setStep(passwords[id] ? "enterPassword" : "setPassword");
  }

  async function submitSetPassword() {
    if (pwd.length < 4) { setErr("किमान 4 अंक/अक्षरांचा पासवर्ड ठेवा"); return; }
    if (pwd !== pwd2) { setErr("दोन्ही पासवर्ड जुळत नाहीत"); return; }
    setBusy(true);
    const res = await onCreatePassword(accId, pwd);
    setBusy(false);
    if (!res.ok) {
      if (res.reason === "locked") setErr("हे खातं सध्या दुसऱ्या डिव्हाइसवर सुरू आहे");
      else setErr("काहीतरी चुकलं, पुन्हा प्रयत्न करा");
    }
  }

  async function submitEnterPassword(force = false) {
    setBusy(true);
    const res = await onAttemptLogin(accId, pwd, force);
    setBusy(false);
    if (!res.ok) {
      if (res.reason === "wrong") { setErr("पासवर्ड चुकीचा आहे"); setPwd(""); }
      else if (res.reason === "locked") setErr("locked");
      else setErr("काहीतरी चुकलं, पुन्हा प्रयत्न करा");
    }
  }

  return (
    <div className="min-h-screen paper-texture font-body flex flex-col items-center justify-center px-6 py-10 text-center relative">
      <style>{FONT_STYLE}</style>
      <FestiveWatermark />
      <MonogramBadge letter={(mandalName || "उ").charAt(0)} size={116} />
      <p className="text-[11px] tracking-[0.25em] uppercase opacity-50 font-semibold relative z-10 mt-3" style={{ color: "#7A2048" }}>खजिनदार वही</p>
      <h1 className="font-display text-2xl mt-1.5 relative z-10" style={{ color: "#7A2048" }}>{mandalName}</h1>

      {showKicked && (
        <div className="relative z-10 mt-4 w-full max-w-sm flex items-start gap-2 text-left px-3.5 py-2.5 rounded-xl text-xs" style={{ background: "#FBE4E4", color: "#A83E3E" }}>
          <ShieldAlert size={15} className="mt-0.5 shrink-0" />
          तुमचं खातं दुसऱ्या डिव्हाइसवरून लॉगिन झाल्यामुळे इथे आपोआप लॉगआऊट झालं.
        </div>
      )}

      <div className="relative z-10 flex flex-col items-center w-full mt-2">
        {step === "pick" && (
          <>
            <p className="text-sm opacity-60 mt-4 mb-6 max-w-xs">सुरू करण्यापूर्वी तुमचं खातं निवडा</p>

            <div className="w-full max-w-sm">
              <p className="text-[11px] font-semibold tracking-wide uppercase opacity-50 text-left mb-2" style={{ color: "#7A2048" }}>पदाधिकारी</p>
              <div className="space-y-2.5 mb-6">
                {officers.map((o) => {
                  const Icon = ACCOUNT_ICONS[o.icon] || Crown;
                  return (
                    <button
                      key={o.id}
                      onClick={() => pick(o.id)}
                      className="focus-ring w-full flex items-center gap-4 p-3.5 rounded-2xl border text-left active:scale-[0.98] transition-transform shadow-sm"
                      style={{ background: "#FBE9C6", borderColor: "#EFD8A0" }}
                    >
                      <div className="p-2.5 rounded-full bg-white"><Icon size={19} style={{ color: "#8A5A0E" }} /></div>
                      <div>
                        <p className="font-semibold text-sm" style={{ color: "#7A2048" }}>{o.label}</p>
                        <p className="text-[11px] opacity-60 mt-0.5">वर्गणी/खर्च जोडू शकतात, सेटिंग्ज बदलू शकतात</p>
                      </div>
                    </button>
                  );
                })}
              </div>

              <p className="text-[11px] font-semibold tracking-wide uppercase opacity-50 text-left mb-2" style={{ color: "#7A2048" }}>कार्यकर्ता (फक्त पाहणारे)</p>
              <div className="grid grid-cols-5 gap-2">
                {viewers.map((v, i) => (
                  <button
                    key={v.id}
                    onClick={() => pick(v.id)}
                    className="focus-ring flex flex-col items-center gap-1 py-3 rounded-xl border active:scale-[0.96] transition-transform shadow-sm"
                    style={{ background: "#E4EEF1", borderColor: "#C9DEE2" }}
                  >
                    <Eye size={16} style={{ color: "#3D6B75" }} />
                    <span className="text-xs font-semibold" style={{ color: "#7A2048" }}>{i + 1}</span>
                  </button>
                ))}
              </div>
            </div>
            <p className="text-[11px] opacity-45 mt-8 max-w-xs">प्रत्येक खात्याचा स्वतःचा पासवर्ड असतो · एका वेळी एकाच डिव्हाइसवर लॉगिन शक्य</p>
          </>
        )}

        {step === "setPassword" && acc && (
          <div className="w-full max-w-sm mt-2">
            <div className="flex items-center justify-center gap-2 mb-4">
              {(() => { const Icon = ACCOUNT_ICONS[acc.icon] || Eye; return <Icon size={16} style={{ color: "#7A2048" }} />; })()}
              <p className="font-semibold text-sm" style={{ color: "#7A2048" }}>{acc.label}</p>
            </div>
            <p className="text-sm opacity-70 mb-4">पहिल्यांदाच लॉगिन करताय — या खात्यासाठी एक पासवर्ड तयार करा.</p>
            <div className="mb-2.5">
              <PasswordInput value={pwd} onChange={(e) => { setPwd(e.target.value); setErr(""); }} placeholder="नवीन पासवर्ड" />
            </div>
            <PasswordInput value={pwd2} onChange={(e) => { setPwd2(e.target.value); setErr(""); }} placeholder="पासवर्ड पुन्हा टाका" />
            {err && <p className="text-xs mt-2 text-left" style={{ color: "#A83E3E" }}>{err}</p>}
            <button disabled={busy} onClick={submitSetPassword} className="focus-ring w-full mt-4 py-3 rounded-xl font-semibold border active:scale-[0.98] transition-transform disabled:opacity-50" style={{ background: "#FBE9C6", color: "#7A2048", borderColor: "#EFD8A0" }}>
              {busy ? "जतन होत आहे…" : "पासवर्ड ठेवा व सुरू करा"}
            </button>
            <button onClick={() => setStep("pick")} className="focus-ring w-full mt-2 py-2 text-sm opacity-60">← मागे</button>
          </div>
        )}

        {step === "enterPassword" && acc && (
          <div className="w-full max-w-sm mt-2">
            <div className="flex items-center justify-center gap-2 mb-4">
              {(() => { const Icon = ACCOUNT_ICONS[acc.icon] || Eye; return <Icon size={16} style={{ color: "#7A2048" }} />; })()}
              <p className="font-semibold text-sm" style={{ color: "#7A2048" }}>{acc.label}</p>
            </div>
            <PasswordInput
              value={pwd}
              autoFocus
              onChange={(e) => { setPwd(e.target.value); setErr(""); }}
              onKeyDown={(e) => e.key === "Enter" && submitEnterPassword()}
              placeholder="पासवर्ड"
              borderColor={err ? "#C15757" : undefined}
            />
            {err === "locked" ? (
              <div className="mt-3 text-left px-3.5 py-2.5 rounded-xl text-xs" style={{ background: "#FBE4E4", color: "#A83E3E" }}>
                हे खातं सध्या दुसऱ्या डिव्हाइसवर लॉगिन आहे.
                <button disabled={busy} onClick={() => submitEnterPassword(true)} className="focus-ring block w-full mt-2 py-2 rounded-lg font-semibold text-xs border active:scale-[0.98] transition-transform" style={{ background: "white", borderColor: "#F3C7C7" }}>
                  जुनं सेशन बंद करून इथे लॉगिन करा
                </button>
              </div>
            ) : err ? (
              <p className="text-xs mt-2" style={{ color: "#A83E3E" }}>{err}</p>
            ) : null}
            <button disabled={busy} onClick={() => submitEnterPassword(false)} className="focus-ring w-full mt-4 py-3 rounded-xl font-semibold border active:scale-[0.98] transition-transform disabled:opacity-50" style={{ background: "#FBE9C6", color: "#7A2048", borderColor: "#EFD8A0" }}>
              {busy ? "तपासत आहे…" : "आत जा"}
            </button>
            <button onClick={() => setStep("pick")} className="focus-ring w-full mt-2 py-2 text-sm opacity-60">← मागे</button>
          </div>
        )}
      </div>
    </div>
  );
}

function RecordsList({ tab, chanda, kharch, search, onOpenReceipt, onOpenProof }) {
  const q = search.trim().toLowerCase();
  let items = [];
  if (tab === "home" || tab === "chanda") {
    items = items.concat(chanda.filter(c => !q || c.name.toLowerCase().includes(q) || (c.note || "").toLowerCase().includes(q)).map(c => ({ ...c, _type: "chanda" })));
  }
  if (tab === "home" || tab === "kharch") {
    items = items.concat(kharch.filter(k => !q || k.desc.toLowerCase().includes(q)).map(k => ({ ...k, _type: "kharch" })));
  }
  items.sort((a, b) => b.createdAt - a.createdAt);

  if (items.length === 0) {
    const emptyCopy = tab === "chanda" ? "अजून वर्गणी नोंद नाही — वरचं 'वर्गणी जमा' बटण दाबून सुरुवात करा."
      : tab === "kharch" ? "अजून खर्च नोंद नाही — वरचं 'खर्च नोंद' बटण दाबून सुरुवात करा."
      : "अजून काही नोंद नाही — वरून वर्गणी किंवा खर्च जोडा.";
    return (
      <div className="text-center py-14 opacity-70">
        <Receipt size={30} className="mx-auto mb-2.5" style={{ color: "#7A2048" }} strokeWidth={1.5} />
        <p className="text-sm px-6 leading-relaxed">{emptyCopy}</p>
      </div>
    );
  }

  return (
    <div className="space-y-2">
      {items.map((item) => (
        <div
          key={item.id}
          className="bg-white rounded-xl pl-3.5 pr-3 py-3 flex items-center justify-between shadow-sm border-l-4"
          style={{ borderLeftColor: item._type === "chanda" ? "#4C8058" : "#C15757" }}
        >
          <div className="min-w-0 flex-1">
            <p className="font-medium text-sm truncate">{item._type === "chanda" ? item.name : item.desc}</p>
            <p className="text-xs opacity-60 mt-0.5">
              {fmtDate(item.date)} {item._type === "chanda" ? `· पावती #${item.receiptNo}` : ""}
            </p>
          </div>
          <div className="flex items-center gap-2 ml-2 shrink-0">
            <span className="font-num font-semibold text-sm" style={{ color: item._type === "chanda" ? "#4C8058" : "#C15757" }}>
              {item._type === "chanda" ? "+" : "−"}₹{fmtAmt(item.amount)}
            </span>
            {item._type === "chanda" ? (
              <button onClick={() => onOpenReceipt(item)} className="focus-ring p-1.5 rounded-full active:scale-90 transition-transform" style={{ background: "rgba(122,32,72,0.08)" }} aria-label="पावती बघा">
                <ChevronRight size={15} style={{ color: "#7A2048" }} />
              </button>
            ) : item.proof ? (
              <button onClick={() => onOpenProof(item.proof)} className="focus-ring p-1.5 rounded-full active:scale-90 transition-transform" style={{ background: "rgba(122,32,72,0.08)" }} aria-label="प्रूफ बघा">
                <Camera size={15} style={{ color: "#7A2048" }} />
              </button>
            ) : (
              <span className="p-1.5" title="प्रूफ नाही"><ImageOff size={15} className="opacity-30" /></span>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

function Sheet({ title, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center" style={{ background: "rgba(36,28,21,0.72)" }} onClick={onClose}>
      <div
        className="w-full max-w-md paper-texture rounded-t-3xl px-5 pt-5 pb-8 max-h-[88vh] overflow-y-auto shadow-2xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="w-10 h-1 rounded-full bg-black/15 mx-auto mb-4" />
        <div className="flex items-center justify-between mb-4">
          <h2 className="font-display text-xl" style={{ color: "#7A2048" }}>{title}</h2>
          <button onClick={onClose} className="focus-ring p-1.5 rounded-full active:scale-90 transition-transform" style={{ background: "rgba(122,32,72,0.08)" }}>
            <X size={18} style={{ color: "#7A2048" }} />
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <div className="mb-3">
      <label className="text-xs font-medium opacity-70 mb-1 block">{label}</label>
      {children}
    </div>
  );
}

const inputClass = "focus-ring w-full px-3.5 py-2.5 rounded-xl bg-white border outline-none text-sm transition-shadow";
const inputStyle = { borderColor: "rgba(122,32,72,0.2)" };

function ChandaForm({ onClose, onSave, nextReceiptNo }) {
  const [name, setName] = useState("");
  const [phone, setPhone] = useState("");
  const [amount, setAmount] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [note, setNote] = useState("");

  const canSave = name.trim() && Number(amount) > 0;

  return (
    <Sheet title={`वर्गणी जमा — पावती #${nextReceiptNo}`} onClose={onClose}>
      <Field label="नाव *">
        <input className={inputClass} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="वर्गणीदाराचे नाव" autoFocus />
      </Field>
      <Field label="WhatsApp नंबर (ऐच्छिक)">
        <input className={inputClass} style={inputStyle} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="10 अंकी मोबाईल नंबर" inputMode="numeric" />
      </Field>
      <Field label="रक्कम (₹) *">
        <input className={inputClass} style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" inputMode="numeric" type="number" />
      </Field>
      <Field label="दिनांक">
        <input className={inputClass} style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} type="date" />
      </Field>
      <Field label="तपशील (ऐच्छिक)">
        <input className={inputClass} style={inputStyle} value={note} onChange={(e) => setNote(e.target.value)} placeholder="उदा. रोख / ऑनलाईन" />
      </Field>
      <button
        disabled={!canSave}
        onClick={() => onSave({ name: name.trim(), phone: phone.trim(), amount: Number(amount), date, note: note.trim() })}
        className="focus-ring w-full mt-2 py-3 rounded-xl font-semibold disabled:opacity-40 active:scale-[0.98] transition-transform border"
        style={{ background: "#E4F1E6", color: "#3C6B45", borderColor: "#C7E2CB" }}
      >
        पावती तयार करा
      </button>
    </Sheet>
  );
}

function KharchForm({ onClose, onSave }) {
  const [desc, setDesc] = useState("");
  const [amount, setAmount] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [proof, setProof] = useState(null);
  const [compressing, setCompressing] = useState(false);
  const [photoErr, setPhotoErr] = useState("");
  const fileRef = useRef(null);

  const canSave = desc.trim() && Number(amount) > 0;

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setCompressing(true);
    setPhotoErr("");
    try {
      const dataUrl = await compressImage(file);
      setProof(dataUrl);
    } catch (err) {
      setPhotoErr("फोटो जोडता आला नाही, पुन्हा प्रयत्न करा");
    }
    setCompressing(false);
    e.target.value = "";
  }

  return (
    <Sheet title="खर्च नोंद" onClose={onClose}>
      <Field label="तपशील *">
        <input className={inputClass} style={inputStyle} value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="उदा. डेकोरेशन, मंडप, प्रसाद" autoFocus />
      </Field>
      <Field label="रक्कम (₹) *">
        <input className={inputClass} style={inputStyle} value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="0" inputMode="numeric" type="number" />
      </Field>
      <Field label="दिनांक">
        <input className={inputClass} style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} type="date" />
      </Field>
      <Field label="प्रूफ / बिल फोटो (ऐच्छिक)">
        {proof ? (
          <div className="relative w-24 h-24">
            <img src={proof} alt="proof" className="w-24 h-24 object-cover rounded-xl border" style={{ borderColor: "rgba(122,32,72,0.2)" }} />
            <button onClick={() => setProof(null)} className="focus-ring absolute -top-2 -right-2 bg-white rounded-full p-1 shadow active:scale-90 transition-transform" style={{ color: "#9C3D3D" }}>
              <X size={13} />
            </button>
          </div>
        ) : (
          <button
            onClick={() => fileRef.current?.click()}
            className="focus-ring flex items-center gap-2 px-3.5 py-2.5 rounded-xl border text-sm w-full justify-center active:scale-[0.98] transition-transform"
            style={{ borderColor: "rgba(122,32,72,0.2)", background: "white" }}
          >
            <Camera size={16} style={{ color: "#7A2048" }} />
            {compressing ? "प्रोसेस होत आहे…" : "फोटो जोडा"}
          </button>
        )}
        <input ref={fileRef} type="file" accept="image/*" onChange={handleFile} className="hidden" />
        {photoErr && <p className="text-xs mt-1.5" style={{ color: "#A83E3E" }}>{photoErr}</p>}
      </Field>
      <button
        disabled={!canSave}
        onClick={() => onSave({ desc: desc.trim(), amount: Number(amount), date, proof })}
        className="focus-ring w-full mt-2 py-3 rounded-xl font-semibold disabled:opacity-40 active:scale-[0.98] transition-transform border"
        style={{ background: "#FBE4E4", color: "#A83E3E", borderColor: "#F3C7C7" }}
      >
        खर्च जतन करा
      </button>
    </Sheet>
  );
}

function SettingsSheet({ settings, isOwner, accountLabel, onClose, onSave, onChangePassword, onLogout }) {
  const [mandalName, setMandalName] = useState(settings.mandalName);
  const [place, setPlace] = useState(settings.place || "");
  const [receiptStart, setReceiptStart] = useState(settings.receiptStart || 101);
  const [newPwd, setNewPwd] = useState("");
  const [pwdSaved, setPwdSaved] = useState(false);

  return (
    <Sheet title="सेटिंग्ज" onClose={onClose}>
      <div className="flex items-center gap-2 mb-4 px-3 py-2 rounded-xl" style={{ background: "rgba(122,32,72,0.06)" }}>
        <Lock size={14} style={{ color: "#7A2048" }} />
        <p className="text-xs font-medium" style={{ color: "#7A2048" }}>लॉगिन: {accountLabel}</p>
      </div>

      <p className="text-[11px] font-semibold tracking-wide uppercase opacity-50 mb-2">मंडळ माहिती</p>
      <Field label="मंडळाचे नाव">
        <input className={inputClass} style={inputStyle} value={mandalName} onChange={(e) => setMandalName(e.target.value)} disabled={!isOwner} />
      </Field>
      <Field label="गाव / ठिकाण">
        <input className={inputClass} style={inputStyle} value={place} onChange={(e) => setPlace(e.target.value)} placeholder="उदा. ठाणे" disabled={!isOwner} />
      </Field>
      <Field label="पावती क्रमांक सुरुवात">
        <input className={inputClass} style={inputStyle} value={receiptStart} onChange={(e) => setReceiptStart(Number(e.target.value) || 1)} type="number" disabled={!isOwner} />
      </Field>
      {isOwner ? (
        <button
          onClick={() => onSave({ ...settings, mandalName: mandalName.trim() || DEFAULT_SETTINGS.mandalName, place: place.trim(), receiptStart })}
          className="focus-ring w-full mt-2 py-3 rounded-xl font-semibold active:scale-[0.98] transition-transform border"
          style={{ background: "#FBE9C6", color: "#7A2048", borderColor: "#EFD8A0" }}
        >
          जतन करा
        </button>
      ) : (
        <p className="text-xs opacity-60 text-center mt-1">पाहणारा खात्याने मंडळ माहिती बदलता येत नाही</p>
      )}

      <div className="h-px my-5" style={{ background: "rgba(122,32,72,0.12)" }} />

      <p className="text-[11px] font-semibold tracking-wide uppercase opacity-50 mb-2">तुमचा पासवर्ड</p>
      <Field label="नवीन पासवर्ड">
        <PasswordInput value={newPwd} onChange={(e) => { setNewPwd(e.target.value); setPwdSaved(false); }} placeholder="किमान 4 अंक/अक्षरे" />
      </Field>
      <button
        onClick={async () => { if (newPwd.trim().length < 4) return; await onChangePassword(newPwd.trim()); setNewPwd(""); setPwdSaved(true); }}
        disabled={newPwd.trim().length < 4}
        className="focus-ring w-full py-2.5 rounded-xl text-sm font-semibold border active:scale-[0.98] transition-transform disabled:opacity-40"
        style={{ background: "#E4F1E6", color: "#3C6B45", borderColor: "#C7E2CB" }}
      >
        {pwdSaved ? "पासवर्ड बदलला ✓" : "पासवर्ड अपडेट करा"}
      </button>

      <button onClick={onLogout} className="focus-ring w-full mt-5 py-2.5 rounded-xl text-sm font-medium border flex items-center justify-center gap-2 active:scale-[0.98] transition-transform" style={{ borderColor: "#F3C7C7", color: "#A83E3E", background: "#FDF8ED" }}>
        <LogOut size={15} /> लॉगआऊट करा
      </button>
    </Sheet>
  );
}

function ReceiptModal({ entry, settings, onClose, onShare }) {
  const canvasRef = useRef(null);
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");
  const [imgSrc, setImgSrc] = useState(null);

  useEffect(() => {
    let cancelled = false;
    setReady(false);
    setImgSrc(null);
    (async () => {
      if (canvasRef.current) {
        await drawReceiptCanvas(canvasRef.current, entry, settings);
        if (!cancelled) {
          setImgSrc(canvasRef.current.toDataURL("image/jpeg", 0.92));
          setReady(true);
        }
      }
    })();
    return () => { cancelled = true; };
  }, [entry, settings]);

  function getBlob() {
    return new Promise((resolve) => canvasRef.current.toBlob(resolve, "image/jpeg", 0.92));
  }

  async function handleDownload() {
    setBusy(true);
    const blob = await getBlob();
    setBusy(false);
    if (!blob) return;
    const filename = `वर्गणी-पावती-${entry.receiptNo}.jpg`;

    // 1. try the native share sheet — most reliable "save image" path on mobile,
    //    since it isn't blocked by the sandboxed preview's download restrictions
    try {
      const file = new File([blob], filename, { type: "image/jpeg" });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: "वर्गणी पावती" });
        return;
      }
    } catch (e) {}

    // 2. try a direct download link
    try {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 3000);
    } catch (e) {}

    // 3. this sandboxed preview often blocks both of the above — the reliable
    //    fallback here is a real <img> the person can long-press and save,
    //    shown just below the buttons.
    setStatusMsg("वरचा फोटो डाउनलोड झाला नसेल, तर खालच्या फोटोवर long-press करून \"इमेज सेव्ह करा / Save image\" निवडा.");
  }

  async function handleShareImage() {
    setBusy(true);
    setStatusMsg("");
    const blob = await getBlob();
    setBusy(false);
    if (!blob) return;
    const filename = `वर्गणी-पावती-${entry.receiptNo}.jpg`;
    const file = new File([blob], filename, { type: "image/jpeg" });

    // Only path that can hand the actual image (not just text) to WhatsApp:
    // the native share sheet, where the person picks WhatsApp themselves.
    try {
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: "वर्गणी पावती" });
        return;
      }
    } catch (e) {
      if (e?.name === "AbortError") return; // person cancelled the share sheet — not an error
    }

    // Native share isn't available here — don't fall back to a text-only
    // WhatsApp message (the image would be missing from it).
    setStatusMsg("इथून थेट फोटोसकट पाठवणं शक्य नाही. खालच्या फोटोवर long-press करून सेव्ह करा, मग WhatsApp मध्ये attach करा.");
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-5" style={{ background: "rgba(36,28,21,0.78)" }} onClick={onClose}>
      <div className="w-full max-w-md" onClick={(e) => e.stopPropagation()}>
        <div className="rounded-2xl overflow-hidden shadow-2xl relative border-2 bg-white" style={{ borderColor: "#C98A2B" }}>
          {!ready && <div className="aspect-[1500/560] flex items-center justify-center text-sm opacity-50" style={{ color: "#7A2048" }}>पावती तयार होत आहे…</div>}
          <canvas ref={canvasRef} className="hidden" />
          {ready && imgSrc && <img src={imgSrc} alt="वर्गणी पावती" className="w-full h-auto block" />}
        </div>

        <div className="flex gap-3 mt-4">
          <button onClick={onClose} className="focus-ring flex-1 py-3 rounded-xl font-medium bg-white active:scale-[0.98] transition-transform border" style={{ color: "#7A2048", borderColor: "rgba(122,32,72,0.15)" }}>बंद करा</button>
          <button onClick={handleDownload} disabled={busy || !ready} className="focus-ring flex-1 py-3 rounded-xl font-semibold flex items-center justify-center gap-2 active:scale-[0.98] transition-transform border disabled:opacity-50" style={{ background: "#FBE9C6", color: "#7A2048", borderColor: "#EFD8A0" }}>
            <FileDown size={16} /> Image
          </button>
          <button onClick={handleShareImage} disabled={busy || !ready} className="focus-ring flex-1 py-3 rounded-xl font-semibold flex items-center justify-center gap-2 active:scale-[0.98] transition-transform border disabled:opacity-50" style={{ background: "#E5F7EB", color: "#1F9C4C", borderColor: "#BEE8CB" }}>
            <Share2 size={16} /> WhatsApp
          </button>
        </div>
        {statusMsg && (
          <p className="text-xs mt-2.5 text-center leading-relaxed" style={{ color: "#7A2048" }}>{statusMsg}</p>
        )}
        {ready && (
          <p className="text-[11px] mt-2 text-center opacity-50">टीप: वरच्या फोटोवर बोट दाबून ठेवलं (long-press) की थेट "Save image" पर्याय मिळतो</p>
        )}
      </div>
    </div>
  );
}

function Row({ label, value, bold }) {
  return (
    <div className="flex justify-between">
      <span className="opacity-60">{label}</span>
      <span className={bold ? "font-num font-bold" : "font-medium"} style={bold ? { color: "#4C8058" } : {}}>{value}</span>
    </div>
  );
}

function ProofModal({ src, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-5" style={{ background: "rgba(36,28,21,0.75)" }} onClick={onClose}>
      <img src={src} alt="proof" className="max-w-full max-h-[80vh] rounded-xl shadow-2xl" onClick={(e) => e.stopPropagation()} />
    </div>
  );
}
