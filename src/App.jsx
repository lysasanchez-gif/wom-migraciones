import { useState, useCallback, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ============================================================
// WOM MIGRACIONES v5.0 — Con datos bancarios, apelaciones, filtro mes
// ============================================================

const MOTIV = [
  "Cada llamada es una nueva oportunidad. Los campeones no se rinden. 🔥",
  "La disciplina vence al talento cuando el talento no tiene disciplina. 💪",
  "No vendas un producto, resuelve un problema. 🎯",
  "El 80% de las ventas se cierran después del quinto intento. 📞",
  "Tu voz es tu herramienta más poderosa. 🎙️",
  "Los rechazos son escalones. Cada 'no' te acerca al 'sí'. 🪜",
  "Hoy alguien necesita lo que tú ofreces. Encuéntralo. 🔍",
  "¡Que el cliente sienta tu energía! ⚡",
  "Escucha el doble de lo que hablas. Escucha, conecta, cierra. 👂",
  "Compite con la versión de ti que fuiste ayer. 🏆",
  "El éxito es constancia. Cada día te acercas más. 📈",
  "Detrás de cada número hay alguien que confió en ti. 🤝",
  "Los grandes vendedores crean oportunidades. 🚀",
  "Arranca con mentalidad de campeón. 👑",
  "Hoy es una nueva chance de romper tu récord. 💥",
  "La constancia supera la perfección. Sigue cerrando. 🎯",
  "Cada cliente satisfecho es tu mejor publicidad. 🌟",
  "Enfrenta el miedo al rechazo y crece. 🦁",
  "Enfócate en el proceso. Los resultados llegarán. 🌱",
  "Eres más fuerte de lo que crees. 💎",
  "Sonríe al hablar. El cliente lo percibe. 😊",
  "Cada objeción es una pregunta. Respóndela y cierra. 🗝️",
  "Tu meta solo requiere más llamadas. ¡A darle! 📱",
  "No hay mal día con la actitud correcta. ✅",
  "El cliente compra confianza. Dásela. 🛡️",
  "Cuando no puedas más, recuerda por qué empezaste. 💭",
  "Hoy es perfecto para ser el #1. 🥇",
  "Las excusas no generan comisiones. La acción sí. 📲",
  "Sé la razón por la que tu equipo sube de nivel. 🌟",
  "Los lunes son para los valientes. ¡Arrasa! 🗓️",
];
const REJ_R=["Audio incorrecto (audio de otro cliente)","Audio incorrecto (contenido)","Audio no subido","Crédito cero","Documentos incorrectos","Pendiente subida VDI"];
const BANKS=["Banco de Venezuela","Banesco","Mercantil","Provincial","BNC","Bicentenario","Banco del Tesoro","Bancaribe","Exterior","Venezuela","BFC","Bancamiga","Banplus","Caroní","100% Banco","Bancrecer","Mi Banco","Banco Activo","Banco Plaza","Otro"];
const TIERS=[{min:1,max:4,rate:1.7},{min:5,max:7,rate:2.0},{min:8,max:9,rate:2.5},{min:10,max:Infinity,rate:3.0}];
function getRate(n){for(const t of TIERS)if(n>=t.min&&n<=t.max)return t.rate;return 0;}
function getNext(n){if(n<5)return{need:5-n,r:2.0};if(n<8)return{need:8-n,r:2.5};if(n<10)return{need:10-n,r:3.0};return null;}
function calcDay(ok){if(!ok)return{rate:0,total:0};const r=getRate(ok);return{rate:r,total:ok*r};}
function $(a){return`$${a.toFixed(2)}`;}
function tdy(){return new Date().toISOString().split("T")[0];}
function fd(d){return new Date(d+"T12:00:00").toLocaleDateString("es-CL",{day:"2-digit",month:"2-digit"});}
function fdf(d){return new Date(d+"T12:00:00").toLocaleDateString("es-CL");}
function wdL(){const t=new Date(),l=new Date(t.getFullYear(),t.getMonth()+1,0);let c=0;for(let d=new Date(t);d<=l;d.setDate(d.getDate()+1)){if(d.getDay()>=1&&d.getDay()<=6)c++;}return c;}
function fts(ts){return new Date(ts).toLocaleString("es-CL",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});}
function gMN(m){return["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][m];}

const U0=[
  {id:1,fullName:"Coordinadora WOM",username:"admin",password:"admin123",role:"super_admin",active:true,bank:null},
  {id:2,fullName:"Supervisor Operaciones",username:"supervisor",password:"super123",role:"admin",active:true,bank:null},
  {id:3,fullName:"José Ángel Lugo Vázquez",username:"jalugov",password:"venta123",role:"executive",active:true,bank:{account:"01020100000000123456",holder:"José Ángel Lugo",cedula:"12345678",bankName:"Mercantil",type:"ahorro"}},
  {id:4,fullName:"María Carolina Rodríguez P.",username:"mrodriguezp",password:"venta123",role:"executive",active:true,bank:null},
  {id:5,fullName:"Carlos Eduardo Martínez L.",username:"cemartinezl",password:"venta123",role:"executive",active:true,bank:null},
  {id:6,fullName:"Ana Gabriela Fernández D.",username:"agfernandezd",password:"venta123",role:"executive",active:true,bank:null},
  {id:7,fullName:"Pedro José Sánchez M.",username:"pjsanchezm",password:"venta123",role:"executive",active:true,bank:null},
  {id:8,fullName:"Laura Isabel Torres G.",username:"litorresg",password:"venta123",role:"executive",active:true,bank:null},
];

function genS(users){const s=[],ex=users.filter(u=>u.role==="executive"&&u.active);const t=new Date(),yr=t.getFullYear(),mo=t.getMonth();let id=1;
ex.forEach(e=>{for(let d=1;d<=Math.min(t.getDate(),22);d++){const dt=new Date(yr,mo,d);if(dt.getDay()===0)continue;const n=Math.floor(Math.random()*8)+2;
for(let i=0;i<n;i++){const st=d===t.getDate()?(Math.random()>.15?"ok":"pending"):["ok","ok","ok","ok","ok","ok","rejected","pending"][Math.floor(Math.random()*8)];
s.push({id:id++,date:dt.toISOString().split("T")[0],executiveId:e.id,username:e.username,phone:`569${String(Math.floor(Math.random()*90000000)+10000000)}`,
rut:String(Math.floor(Math.random()*900000000)+100000000),status:st,type:Math.random()>.92?"exception":"normal",
isVDI:Math.random()<.08,vdiUser:null,rejectionReason:st==="rejected"?REJ_R[Math.floor(Math.random()*REJ_R.length)]:null,
rejectionDetail:st==="rejected"?"Requiere corrección":null,recoveryAction:null,recoveryComment:null,managedByExec:false,appealed:false});}}});return s;}

const FN="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=DM+Sans:wght@300;400;500;600;700&display=swap";
const CSS=`@import url('${FN}');
*{margin:0;padding:0;box-sizing:border-box}
:root{--dp:#1a0a2e;--pu:#6b21a8;--vi:#8b5cf6;--ma:#d946ef;--pk:#f0abfc;--sf:#f8f7fc;--cd:#fff;--bd:#e9e5f0;--tx:#1e1b2e;--mt:#6b7280;--ok:#059669;--er:#dc2626;--wr:#d97706;--in:#2563eb;--sw:260px}
body{font-family:'DM Sans',sans-serif;background:var(--sf);color:var(--tx)}
.lbg{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#1a0a2e,#3b0764 40%,#6b21a8 70%,#d946ef)}
.lc{background:rgba(255,255,255,.07);backdrop-filter:blur(24px);border:1px solid rgba(255,255,255,.12);border-radius:24px;padding:48px 40px;width:420px;max-width:92vw;z-index:1;box-shadow:0 32px 64px rgba(0,0,0,.3)}
.ll{font-family:'Outfit';font-weight:800;font-size:32px;color:#fff;text-align:center;margin-bottom:4px}
.ls{font-family:'Outfit';font-size:13px;color:rgba(255,255,255,.5);text-align:center;margin-bottom:36px;letter-spacing:3px;text-transform:uppercase}
.li{width:100%;padding:14px 16px;border-radius:12px;border:1px solid rgba(255,255,255,.15);background:rgba(255,255,255,.06);color:#fff;font-size:15px;font-family:'DM Sans';outline:none;transition:.2s;margin-bottom:14px}
.li::placeholder{color:rgba(255,255,255,.35)}.li:focus{border-color:var(--ma);background:rgba(255,255,255,.1)}
.lb{width:100%;padding:14px;border-radius:12px;border:none;background:linear-gradient(135deg,var(--vi),var(--ma));color:#fff;font-size:16px;font-weight:600;font-family:'Outfit';cursor:pointer;transition:.25s;margin-top:8px}
.lb:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(217,70,239,.35)}
.le{background:rgba(220,38,38,.15);border:1px solid rgba(220,38,38,.3);color:#fca5a5;padding:10px 14px;border-radius:10px;font-size:13px;margin-bottom:14px;text-align:center}
.ld{margin-top:24px;padding-top:20px;border-top:1px solid rgba(255,255,255,.1);color:rgba(255,255,255,.4);font-size:12px;text-align:center;line-height:1.8}.ld b{color:rgba(255,255,255,.7)}
.app{display:flex;min-height:100vh}
.sb{width:var(--sw);min-height:100vh;background:var(--dp);color:#fff;display:flex;flex-direction:column;position:fixed;top:0;left:0;bottom:0;z-index:50;overflow-y:auto}
.sb-h{padding:28px 22px 20px;border-bottom:1px solid rgba(255,255,255,.06)}.sb-b{font-family:'Outfit';font-weight:700;font-size:20px}.sb-c{font-size:11px;color:rgba(255,255,255,.4);letter-spacing:2px;text-transform:uppercase;margin-top:4px}
.sb-n{flex:1;padding:16px 12px}.sb-l{font-size:10px;text-transform:uppercase;letter-spacing:2px;color:rgba(255,255,255,.25);padding:16px 12px 8px;font-weight:600}
.si{display:flex;align-items:center;gap:12px;padding:11px 14px;border-radius:10px;cursor:pointer;font-size:14px;color:rgba(255,255,255,.55);transition:.15s;margin-bottom:2px}.si:hover{color:#fff;background:rgba(255,255,255,.06)}.si.a{color:#fff;background:linear-gradient(135deg,rgba(139,92,246,.3),rgba(217,70,239,.15));font-weight:500}
.sb-f{padding:16px;border-top:1px solid rgba(255,255,255,.06)}.sb-u{display:flex;align-items:center;gap:10px;padding:8px}
.sb-av{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,var(--vi),var(--ma));display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;font-family:'Outfit';flex-shrink:0}
.sb-ui{flex:1;min-width:0}.sb-un{font-size:13px;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.sb-ur{font-size:11px;color:rgba(255,255,255,.35)}
.sb-o{background:none;border:none;color:rgba(255,255,255,.3);cursor:pointer;padding:6px;border-radius:6px}.sb-o:hover{color:#fff;background:rgba(255,255,255,.08)}
.mn{margin-left:var(--sw);flex:1;padding:32px;min-height:100vh}
.ph{margin-bottom:28px}.pg{font-family:'Outfit';font-size:28px;font-weight:700}.pg span{color:var(--vi)}
.mo{background:linear-gradient(135deg,#1a0a2e,#3b0764);border-radius:16px;padding:20px 24px;margin-bottom:28px;color:#fff;display:flex;align-items:flex-start;gap:14px}
.mo-l{font-size:10px;letter-spacing:2px;text-transform:uppercase;color:var(--pk);margin-bottom:6px;font-weight:600}.mo-t{font-size:15px;line-height:1.6;font-style:italic;color:rgba(255,255,255,.85)}
.mg{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px;margin-bottom:28px}
.mc{background:var(--cd);border:1px solid var(--bd);border-radius:14px;padding:20px;transition:.2s}.mc:hover{box-shadow:0 4px 16px rgba(0,0,0,.06);transform:translateY(-1px)}
.ml{font-size:12px;color:var(--mt);font-weight:500;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
.mv{font-family:'Outfit';font-size:28px;font-weight:700}.ms2{font-size:12px;color:var(--mt);margin-top:4px}
.mv.g{color:var(--ok)}.mv.p{color:var(--pu)}.mv.b{color:var(--in)}.mv.o{color:var(--wr)}.mv.r{color:var(--er)}
.nt{background:linear-gradient(135deg,#fef3c7,#fde68a);border:1px solid #f59e0b;border-radius:14px;padding:16px 20px;margin-bottom:28px;display:flex;align-items:center;gap:12px}
.nt-t{font-size:14px;color:#92400e;font-weight:500}.nt-t b{color:#78350f}
.cd{background:var(--cd);border:1px solid var(--bd);border-radius:16px;overflow:hidden;margin-bottom:24px}
.ch{padding:18px 22px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px}
.ct{font-family:'Outfit';font-weight:600;font-size:16px}
table{width:100%;border-collapse:collapse}
th{text-align:left;padding:12px 16px;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--mt);font-weight:600;background:#faf9fd;border-bottom:1px solid var(--bd);white-space:nowrap}
td{padding:11px 16px;font-size:13px;border-bottom:1px solid #f3f1f8}tr:last-child td{border-bottom:none}tr:hover td{background:#faf9fd}
.bg{display:inline-flex;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600}
.bg-ok{background:#d1fae5;color:#065f46}.bg-rejected{background:#fee2e2;color:#991b1b}.bg-pending{background:#fef3c7;color:#92400e}
.bg-active{background:#d1fae5;color:#065f46}.bg-inactive{background:#f3f4f6;color:#6b7280}
.bg-super_admin{background:#ede9fe;color:#5b21b6}.bg-admin{background:#dbeafe;color:#1e40af}.bg-executive{background:#e0e7ff;color:#3730a3}
.bg-recovered{background:#dbeafe;color:#1e40af}.bg-not_recoverable{background:#f3f4f6;color:#6b7280}.bg-appealed{background:#fef3c7;color:#92400e}
.rc{background:var(--cd);border:1px solid var(--bd);border-radius:16px;overflow:hidden}
.ri{display:flex;align-items:center;gap:12px;padding:14px 20px;border-bottom:1px solid #f3f1f8}.ri:last-child{border-bottom:none}
.rp{font-family:'Outfit';font-size:18px;font-weight:700;width:32px;text-align:center}.rn{flex:1;font-size:14px;font-weight:500}.rn.me{color:var(--vi)}
.ra{font-family:'Outfit';font-weight:700;font-size:15px;color:var(--ok)}
.tc{display:grid;grid-template-columns:1fr 340px;gap:24px;align-items:start}
.fg{margin-bottom:18px}.fl{display:block;font-size:13px;font-weight:600;margin-bottom:6px}
.fi,.fs{width:100%;padding:11px 14px;border-radius:10px;border:1px solid var(--bd);font-size:14px;font-family:'DM Sans';outline:none;background:#fff;transition:.15s}
.fi:focus,.fs:focus{border-color:var(--vi)}.fi.er{border-color:var(--er)}
.fh{font-size:12px;color:var(--mt);margin-top:4px}.fe{font-size:12px;color:var(--er);margin-top:4px}
.bt{display:inline-flex;align-items:center;gap:8px;padding:10px 20px;border-radius:10px;font-size:14px;font-weight:600;font-family:'DM Sans';cursor:pointer;border:none;transition:.15s;white-space:nowrap}
.bp{background:linear-gradient(135deg,var(--vi),var(--ma));color:#fff}.bp:hover{box-shadow:0 4px 16px rgba(139,92,246,.3);transform:translateY(-1px)}
.bs{background:#f3f4f6;color:var(--tx)}.bs:hover{background:#e5e7eb}
.bd{background:#fee2e2;color:#991b1b}.bd:hover{background:#fecaca}
.bo{background:#d1fae5;color:#065f46}.bo:hover{background:#a7f3d0}
.bw{background:#fef3c7;color:#92400e}.bw:hover{background:#fde68a}
.bsm{padding:6px 14px;font-size:13px}
.ov{position:fixed;inset:0;background:rgba(0,0,0,.5);backdrop-filter:blur(4px);display:flex;align-items:center;justify-content:center;z-index:100}
.md{background:#fff;border-radius:20px;padding:32px;width:520px;max-width:92vw;max-height:90vh;overflow-y:auto;box-shadow:0 24px 48px rgba(0,0,0,.2)}
.mdt{font-family:'Outfit';font-size:20px;font-weight:700;margin-bottom:20px}.ma{display:flex;gap:10px;justify-content:flex-end;margin-top:24px}
.nf{border-radius:12px;padding:14px 20px;margin-bottom:20px;display:flex;align-items:center;gap:10px;font-size:14px;font-weight:500}
.nf-r{background:#fef2f2;border:1px solid #fecaca;color:#991b1b}.nf-b{background:#eff6ff;border:1px solid #93c5fd;color:#1e40af}.nf-g{background:#f0fdf4;border:1px solid #86efac;color:#065f46}
.tb{display:flex;gap:4px;margin-bottom:20px;background:#f3f1f8;border-radius:12px;padding:4px;flex-wrap:wrap}
.ti{padding:9px 18px;border-radius:9px;font-size:13px;font-weight:500;color:var(--mt);cursor:pointer;border:none;background:none;font-family:'DM Sans'}.ti.a{background:#fff;color:var(--tx);box-shadow:0 1px 4px rgba(0,0,0,.08)}
.gl{background:linear-gradient(135deg,#ede9fe,#fce7f3);border:1px solid #c4b5fd;border-radius:16px;padding:24px;margin-bottom:24px}
.gt{font-family:'Outfit';font-weight:700;font-size:16px;margin-bottom:14px;color:#5b21b6}
.pb{width:100%;height:12px;background:rgba(255,255,255,.7);border-radius:6px;overflow:hidden;margin:10px 0}
.pf{height:100%;border-radius:6px;background:linear-gradient(90deg,var(--vi),var(--ma));transition:width .5s}
.gs{font-size:13px;color:#6d28d9;margin-top:8px}
.em{text-align:center;padding:60px 20px;color:var(--mt)}.em p{font-size:15px}
.ft{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px;align-items:end}.ft .fg{margin-bottom:0;min-width:150px}
.uz{border:2px dashed var(--bd);border-radius:16px;padding:48px;text-align:center;cursor:pointer;transition:.2s;margin-bottom:24px}.uz:hover{border-color:var(--vi);background:#faf5ff}
.pvok{background:#f0fdf4;border:1px solid #86efac;border-radius:10px;padding:12px 16px;margin-bottom:12px;font-size:14px;color:#065f46}
.pvw{background:#fefce8;border:1px solid #fde047;border-radius:10px;padding:12px 16px;margin-bottom:12px;font-size:14px;color:#854d0e}
.lgi{display:flex;gap:12px;padding:14px 20px;border-bottom:1px solid #f3f1f8;align-items:flex-start}
.lgi-i{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.lgi-c{flex:1}.lgi-t{font-size:14px;line-height:1.5}.lgi-tm{font-size:12px;color:var(--mt);margin-top:2px}
.msel{display:inline-flex;align-items:center;gap:8px;background:#fff;border:1px solid var(--bd);border-radius:10px;padding:6px 14px}
.msel select{border:none;outline:none;font-family:'DM Sans';font-size:14px;font-weight:600;color:var(--vi);background:transparent;cursor:pointer}
.ag-item{background:var(--cd);border:1px solid var(--bd);border-radius:12px;padding:16px;margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;gap:12px}
.bank-card{background:linear-gradient(135deg,#1a0a2e,#3b0764);border-radius:16px;padding:24px;color:#fff;margin-bottom:24px}
.bank-label{font-size:10px;letter-spacing:2px;text-transform:uppercase;color:var(--pk);margin-bottom:4px}.bank-val{font-size:16px;font-weight:600;letter-spacing:1px}
@keyframes fadeIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}.fade{animation:fadeIn .3s ease}
@media(max-width:900px){.sb{width:220px}:root{--sw:220px}.tc{grid-template-columns:1fr}.mn{padding:20px}}
@media(max-width:640px){.sb{display:none}:root{--sw:0px}.mn{margin-left:0}.mg{grid-template-columns:1fr 1fr}}`;

const Ic=({n,s=20})=>{const d={
dashboard:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>,
sale:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 2v20M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>,
list:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>,
alert:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>,
users:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>,
target:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>,
upload:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
download:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>,
log:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>,
receipt:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1z"/><path d="M16 8h-6a2 2 0 100 4h4a2 2 0 110 4H8"/><path d="M12 17.5v-11"/></svg>,
logout:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>,
plus:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>,
check:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"/></svg>,
key:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 11-7.778 7.778 5.5 5.5 0 017.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4"/></svg>,
edit:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>,
calendar:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>,
bank:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>,
trash:<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>,
};return d[n]||null;};

export default function App(){
  const[users,setUsers]=useState(U0);
  const[sales,setSales]=useState(()=>genS(U0));
  const[logs,setLogs]=useState([{id:1,ts:Date.now()-3600000,user:"admin",action:"login",detail:"Inicio de sesión"},{id:2,ts:Date.now()-2800000,user:"jalugov",action:"sale",detail:"Registró venta - Tel: 56912345678"}]);
  const[agendas,setAgendas]=useState([{id:1,executiveId:3,clientName:"Carlos Mendoza",phone:"56998765432",callTime:"14:30",date:tdy(),note:"Pedir documentos"}]);
  const[cur,setCur]=useState(null);
  const[pg,setPg]=useState("dashboard");
  const[showCU,setShowCU]=useState(false);
  const[showCP,setShowCP]=useState(null);
  const[lErr,setLErr]=useState("");
  const[lFrm,setLFrm]=useState({username:"",password:""});
  const[sFrm,setSFrm]=useState({date:tdy(),phone:"",rut:"",isVDI:false,vdiUser:"",type:"normal"});
  const[sErr,setSErr]=useState({});
  const[sOk,setSOk]=useState(false);
  const[goal,setGoal]=useState(200);
  const[rTab,setRTab]=useState("monthly");
  const[nU,setNU]=useState({fullName:"",username:"",password:"",role:"executive"});
  const[nP,setNP]=useState("");
  const[aFlt,setAFlt]=useState({exec:"",status:""});
  const[eS,setES]=useState(null);
  const[eFrm,setEFrm]=useState({status:"",rr:"",rd:""});
  const[uData,setUData]=useState(null);
  const[uRes,setURes]=useState(null);
  const[pRng,setPRng]=useState({from:"",to:""});
  const[lgFlt,setLgFlt]=useState({user:"",action:""});
  const[rjCmt,setRjCmt]=useState("");
  const[rjSel,setRjSel]=useState(null);
  const[rjAct,setRjAct]=useState(null);
  const[rejTab,setRejTab]=useState("unmanaged");
  const[dashMonth,setDashMonth]=useState(new Date().getMonth());
  const[dashYear,setDashYear]=useState(new Date().getFullYear());
  const[salesMonth,setSalesMonth]=useState(new Date().getMonth());
  const[salesYear,setSalesYear]=useState(new Date().getFullYear());
  const[agFrm,setAgFrm]=useState({clientName:"",phone:"",callTime:"",date:tdy(),note:""});
  const[showAgFrm,setShowAgFrm]=useState(false);
  const[bankFrm,setBankFrm]=useState({account:"",holder:"",cedula:"",bankName:"",type:"ahorro"});
  const[bankErr,setBankErr]=useState({});
  const[bankOk,setBankOk]=useState(false);
  const fRef=useRef();

  const tMsg=useMemo(()=>{const d=Math.floor((new Date()-new Date(new Date().getFullYear(),0,0))/86400000);return MOTIV[d%MOTIV.length];},[]);
  const addLog=useCallback((a,d)=>{setLogs(p=>[{id:Date.now(),ts:Date.now(),user:cur?.username||"sys",action:a,detail:d},...p]);},[cur]);

  const login=()=>{const u=users.find(x=>x.username===lFrm.username&&x.password===lFrm.password);if(!u){setLErr("Incorrecto");return;}if(!u.active){setLErr("Desactivada");return;}
    setCur(u);setPg("dashboard");setLErr("");if(u.bank)setBankFrm({account:u.bank.account,holder:u.bank.holder,cedula:u.bank.cedula,bankName:u.bank.bankName,type:u.bank.type});};
  const logout=()=>{setCur(null);setLFrm({username:"",password:""});setPg("dashboard");};
  const isAdm=cur&&(cur.role==="super_admin"||cur.role==="admin");const isSA=cur&&cur.role==="super_admin";const isEx=cur&&cur.role==="executive";

  const getStats=useCallback((eid,mo,yr)=>{
    const m=mo??new Date().getMonth(),y=yr??new Date().getFullYear();
    const ms=sales.filter(s=>{const d=new Date(s.date);return s.executiveId===eid&&d.getMonth()===m&&d.getFullYear()===y;});
    const dm={};ms.forEach(s=>{if(!dm[s.date])dm[s.date]={ok:0,rejected:0,pending:0};dm[s.date][s.status==="ok"?"ok":s.status==="rejected"?"rejected":"pending"]++;});
    let tc=0;const dc=[];Object.entries(dm).sort(([a],[b])=>b.localeCompare(a)).forEach(([date,d])=>{const c=calcDay(d.ok);tc+=c.total;dc.push({date,...d,rate:c.rate,commission:c.total});});
    const ts=tdy(),td2=dm[ts]||{ok:0,rejected:0,pending:0},todC=calcDay(td2.ok);
    return{monthlySales:ms,totalOk:ms.filter(s=>s.status==="ok").length,totalRejected:ms.filter(s=>s.status==="rejected").length,totalPending:ms.filter(s=>s.status==="pending").length,totalCommission:tc,dailyCommissions:dc,todayOk:td2.ok,todayRate:todC.rate,todayCommission:todC.total};
  },[sales]);

  const getRanking=useCallback((period)=>{const exs=users.filter(u=>u.role==="executive"&&u.active);const now=new Date(),ws=new Date(now);ws.setDate(now.getDate()-now.getDay()+1);
    return exs.map(e=>{const st=getStats(e.id);let wc=0;if(period==="weekly")st.dailyCommissions.forEach(d=>{if(new Date(d.date)>=ws)wc+=d.commission;});
      return{...e,commission:period==="monthly"?st.totalCommission:wc};}).sort((a,b)=>b.commission-a.commission);},[users,getStats]);

  const valSale=()=>{const e={};if(!sFrm.date)e.date="Requerido";if(!sFrm.phone)e.phone="Requerido";else if(!/^56\d{9}$/.test(sFrm.phone))e.phone="56 + 9 dígitos";
    if(!sFrm.rut)e.rut="Requerido";else if(!/^\d{7,9}$/.test(sFrm.rut))e.rut="Solo números";if(sFrm.isVDI&&!sFrm.vdiUser)e.vdiUser="Requerido";setSErr(e);return!Object.keys(e).length;};
  const regSale=()=>{if(!valSale())return;setSales(p=>[{id:p.length+Date.now(),date:sFrm.date,executiveId:cur.id,username:cur.username,phone:sFrm.phone,rut:sFrm.rut,status:"pending",type:sFrm.type,isVDI:sFrm.isVDI,vdiUser:sFrm.isVDI?sFrm.vdiUser:null,rejectionReason:null,rejectionDetail:null,recoveryAction:null,recoveryComment:null,managedByExec:false,appealed:false},...p]);
    addLog("sale",`Venta Tel: ${sFrm.phone}`);setSFrm({date:tdy(),phone:"",rut:"",isVDI:false,vdiUser:"",type:"normal"});setSErr({});setSOk(true);setTimeout(()=>setSOk(false),3000);};

  const createUser=()=>{if(!nU.fullName||!nU.username||!nU.password)return;if(users.find(u=>u.username===nU.username))return;
    setUsers(p=>[...p,{id:p.length+Date.now(),fullName:nU.fullName,username:nU.username,password:nU.password,role:nU.role,active:true,bank:null}]);addLog("user_create",`Creó ${nU.username}`);setNU({fullName:"",username:"",password:"",role:"executive"});setShowCU(false);};
  const toggleUsr=(uid)=>{const u=users.find(x=>x.id===uid);setUsers(users.map(x=>x.id===uid?{...x,active:!x.active}:x));addLog("user_toggle",`${u.active?"Desactivó":"Reactivó"} ${u.username}`);};
  const chgPw=()=>{if(!nP||!showCP)return;setUsers(users.map(x=>x.id===showCP?{...x,password:nP}:x));addLog("pw_change",`Cambió pw ${users.find(x=>x.id===showCP)?.username}`);setShowCP(null);setNP("");};

  // Bank
  const saveBank=()=>{const e={};if(!bankFrm.account)e.account="Requerido";else if(!/^\d{20}$/.test(bankFrm.account))e.account="Debe tener exactamente 20 dígitos";
    if(!bankFrm.holder)e.holder="Requerido";if(!bankFrm.cedula)e.cedula="Requerido";if(!bankFrm.bankName)e.bankName="Requerido";setBankErr(e);if(Object.keys(e).length)return;
    setUsers(users.map(u=>u.id===cur.id?{...u,bank:{...bankFrm}}:u));setCur({...cur,bank:{...bankFrm}});setBankOk(true);setTimeout(()=>setBankOk(false),3000);addLog("bank","Actualizó datos bancarios");};

  // Recovery
  const handleRecov=(sid,act)=>{setSales(sales.map(s=>s.id===sid?{...s,managedByExec:true,recoveryAction:act,recoveryComment:rjCmt}:s));
    addLog("recovery",`${act==="recovered"?"Recuperó":"No recup."} ${sales.find(s=>s.id===sid)?.phone}`);setRjSel(null);setRjCmt("");setRjAct(null);};
  // Appeal
  const appealSale=(sid)=>{setSales(sales.map(s=>s.id===sid?{...s,appealed:true}:s));addLog("appeal",`Apeló venta ${sales.find(s=>s.id===sid)?.phone}`);};

  const saveEdit=()=>{if(!eS)return;setSales(sales.map(s=>s.id===eS?{...s,status:eFrm.status,rejectionReason:eFrm.status==="rejected"?eFrm.rr:null,rejectionDetail:eFrm.status==="rejected"?eFrm.rd:null}:s));
    addLog("status_change",`${sales.find(s=>s.id===eS)?.phone} → ${eFrm.status}`);setES(null);};

  // Excel upload - now handles appealed sales
  const handleUpload=(e)=>{const file=e.target.files[0];if(!file)return;const r=new FileReader();r.onload=(ev)=>{try{
    const wb=XLSX.read(ev.target.result,{type:"array"});const ws=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(ws,{header:["fecha","usuario","telefono","rut","estado","detalle","observacion"],range:1});
    const matched=[],unmatched=[],updated=[],appealedOk=[];
    rows.forEach(row=>{if(!row.telefono&&!row.rut)return;const ph=String(row.telefono||"").replace(/\D/g,"");const rut=String(row.rut||"").replace(/\D/g,"");
      const est=String(row.estado||"").toLowerCase().trim();const st=est.includes("ok")?"ok":est.includes("rech")?"rejected":"pending";
      let found=sales.find(s=>(ph&&s.phone===ph)||(rut&&s.rut===rut));
      if(found){matched.push({saleId:found.id,newStatus:st,detalle:row.detalle||"",obs:row.observacion||""});
        if(found.status!==st)updated.push({saleId:found.id});
        if(found.appealed&&st==="ok")appealedOk.push({saleId:found.id});}
      else{unmatched.push({fecha:row.fecha,usuario:row.usuario,phone:ph,rut,estado:st});}});
    setUData({matched,unmatched,updated,appealedOk,totalRows:rows.length});}catch{setUData({error:true});}};r.readAsArrayBuffer(file);};
  const confirmUpload=()=>{if(!uData||uData.error)return;const ns=[...sales];
    uData.matched.forEach(m=>{const i=ns.findIndex(s=>s.id===m.saleId);if(i>=0){
      ns[i]={...ns[i],status:m.newStatus,rejectionReason:m.newStatus==="rejected"?m.detalle:null,rejectionDetail:m.newStatus==="rejected"?m.obs:null};
      if(m.newStatus==="ok"){ns[i].managedByExec=false;ns[i].appealed=false;ns[i].recoveryAction=null;ns[i].recoveryComment=null;ns[i].rejectionReason=null;ns[i].rejectionDetail=null;}}});
    setSales(ns);addLog("upload",`Excel: ${uData.matched.length} coinc, ${uData.updated.length} actual, ${uData.appealedOk.length} apeladas→OK, ${uData.unmatched.length} no enc`);
    setURes({matched:uData.matched.length,updated:uData.updated.length,unmatched:uData.unmatched.length,appealedOk:uData.appealedOk.length});setUData(null);if(fRef.current)fRef.current.value="";};

  // Exports
  const exportPayroll=()=>{if(!pRng.from||!pRng.to)return;const f=sales.filter(s=>s.date>=pRng.from&&s.date<=pRng.to);const execs=users.filter(u=>u.role==="executive");
    const detail=f.map(s=>{const ex=execs.find(e=>e.id===s.executiveId);return{Fecha:s.date,Ejecutivo:ex?.fullName||"",Usuario:s.username,Teléfono:s.phone,RUT:s.rut,Estado:s.status==="ok"?"OK":s.status==="rejected"?"Rechazada":"Pendiente",Detalle:s.rejectionReason||""};});
    const dates=[...new Set(f.map(s=>s.date))].sort();const commBD=[];
    execs.forEach(exec=>{dates.forEach(date=>{const dayOk=f.filter(s=>s.executiveId===exec.id&&s.date===date&&s.status==="ok").length;const dayRej=f.filter(s=>s.executiveId===exec.id&&s.date===date&&s.status==="rejected").length;
      if(dayOk>0||dayRej>0){const c=calcDay(dayOk);commBD.push({Fecha:date,Ejecutivo:exec.fullName,Usuario:exec.username,"Ventas OK":dayOk,Rechazadas:dayRej,"Escala USD":c.rate.toFixed(2),"Comisión Día":c.total.toFixed(2)});}});});
    const totals=[];execs.forEach(exec=>{const es=f.filter(s=>s.executiveId===exec.id);if(!es.length)return;const dm={};es.forEach(s=>{if(!dm[s.date])dm[s.date]={ok:0};if(s.status==="ok")dm[s.date].ok++;});
      let tc=0;Object.values(dm).forEach(d=>{tc+=calcDay(d.ok).total;});const bk=exec.bank;
      totals.push({Ejecutivo:exec.fullName,Usuario:exec.username,"Ventas OK":es.filter(s=>s.status==="ok").length,Rechazadas:es.filter(s=>s.status==="rejected").length,"Total USD":tc.toFixed(2),
        "Cuenta Banco":bk?.account||"SIN DATOS","Titular":bk?.holder||"","Cédula":bk?.cedula||"","Banco":bk?.bankName||"","Tipo Cuenta":bk?.type||""});});
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(detail),"Detalle");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(commBD),"Comisiones Día");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(totals),"Nómina Pago");
    XLSX.writeFile(wb,`Nomina_WOM_${pRng.from}_${pRng.to}.xlsx`);addLog("payroll",`Nómina ${pRng.from} a ${pRng.to}`);};
  const exportRecov=()=>{const recs=sales.filter(s=>s.managedByExec).map(s=>{const ex=users.find(u=>u.id===s.executiveId);
    return{Ejecutivo:ex?.fullName||"",Usuario:s.username,Teléfono:s.phone,RUT:s.rut,Motivo:s.rejectionReason||"",Recuperada:s.recoveryAction==="recovered"?"Sí":"No",Acción:s.recoveryComment||"",Apelada:s.appealed?"Sí":"No"};});
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(recs),"Recuperaciones");XLSX.writeFile(wb,`Recuperaciones_${tdy()}.xlsx`);};
  const exportAdm=()=>{const f=getFS();const d=f.map(s=>{const ex=users.find(u=>u.id===s.executiveId);return{Fecha:s.date,Ejecutivo:ex?.fullName||"",Usuario:s.username,Teléfono:s.phone,RUT:s.rut,Estado:s.status==="ok"?"OK":s.status==="rejected"?"Rechazada":"Pendiente"};});
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(d),"Ventas");XLSX.writeFile(wb,`Ventas_${tdy()}.xlsx`);};
  const dlTpl=()=>{const t=[{Fecha:"2026-05-24",Usuario:"jalugov",Teléfono:"56912345678",RUT:"183274765",Estado:"OK",Detalle:"",Observación:""}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t),"Plantilla");XLSX.writeFile(wb,"Plantilla_WOM.xlsx");};

  const addAg=()=>{if(!agFrm.clientName||!agFrm.phone||!agFrm.callTime)return;setAgendas(p=>[...p,{id:Date.now(),executiveId:cur.id,clientName:agFrm.clientName,phone:agFrm.phone,callTime:agFrm.callTime,date:agFrm.date||tdy(),note:agFrm.note}]);setAgFrm({clientName:"",phone:"",callTime:"",date:tdy(),note:""});setShowAgFrm(false);};
  const delAg=(id)=>{setAgendas(agendas.filter(a=>a.id!==id));};

  const getFS=()=>{const now=new Date(),cm=now.getMonth(),cy=now.getFullYear();return sales.filter(s=>{const d=new Date(s.date);if(d.getMonth()!==cm||d.getFullYear()!==cy)return false;if(aFlt.exec&&s.username!==aFlt.exec)return false;if(aFlt.status&&s.status!==aFlt.status)return false;return true;}).sort((a,b)=>b.id-a.id);};

  if(!cur)return (<><style>{CSS}</style><div className="lbg"><div className="lc fade"><div className="ll">WOM</div><div className="ls">Migraciones</div>
    {lErr&&<div className="le">{lErr}</div>}
    <input className="li" placeholder="Usuario" value={lFrm.username} onChange={e=>setLFrm({...lFrm,username:e.target.value})} onKeyDown={e=>e.key==="Enter"&&login()}/>
    <input className="li" type="password" placeholder="Contraseña" value={lFrm.password} onChange={e=>setLFrm({...lFrm,password:e.target.value})} onKeyDown={e=>e.key==="Enter"&&login()}/>
    <button className="lb" onClick={login}>Ingresar</button>
    <div className="ld"><b>Demo:</b><br/>admin / admin123 · jalugov / venta123</div></div></div></>);

  const fn=cur.fullName.split(" ")[0];const ini=cur.fullName.split(" ").map(n=>n[0]).join("").slice(0,2).toUpperCase();
  const rl={super_admin:"Coordinadora",admin:"Administrador",executive:"Ejecutivo"}[cur.role];
  const nav=isEx?[{id:"dashboard",l:"Dashboard",i:"dashboard"},{id:"register",l:"Registrar Venta",i:"sale"},{id:"my-sales",l:"Mis Ventas",i:"list"},
    {id:"my-rejections",l:"Mis Rechazos",i:"alert"},{id:"agenda",l:"Agendados",i:"calendar"},{id:"my-bank",l:"Mi Cuenta",i:"bank"},{id:"my-goal",l:"Mi Meta",i:"target"}]
    :[{id:"dashboard",l:"Dashboard",i:"dashboard"},{id:"all-sales",l:"Ventas",i:"list"},{id:"rejections",l:"Rechazos",i:"alert"},
    {id:"upload",l:"Subir Excel",i:"upload"},{id:"payroll",l:"Nómina",i:"receipt"},{id:"users",l:"Usuarios",i:"users"},{id:"log",l:"Log",i:"log"}];

  // ===== EXEC PAGES =====
  const pgDash=()=>{const st=getStats(cur.id);const rk=getRanking(rTab);const myR=rk.findIndex(r=>r.id===cur.id)+1;const nt=getNext(st.todayOk);
    const rejC=st.monthlySales.filter(s=>s.status==="rejected"&&!s.managedByExec).length;const prog=goal>0?Math.min((st.totalCommission/goal)*100,100):0;
    const myAg=agendas.filter(a=>a.executiveId===cur.id&&a.date===tdy());
    return (<div className="fade"><div className="ph"><div className="pg">¡Hola <span>{fn}</span>! 🔥</div></div>
      <div className="mo"><div style={{fontSize:28,flexShrink:0}}>💡</div><div><div className="mo-l">Mensaje del día</div><div className="mo-t">{tMsg}</div></div></div>
      {rejC>0&&<div className="nf nf-r">⚠️ <b style={{margin:"0 4px"}}>{rejC}</b> rechazada{rejC>1?"s":""} pendiente{rejC>1?"s":""}</div>}
      {myAg.length>0&&<div className="nf nf-b">📅 <b style={{margin:"0 4px"}}>{myAg.length}</b> agendado{myAg.length>1?"s":""} hoy</div>}
      {!cur.bank&&<div className="nf nf-r">🏦 No has registrado tus datos bancarios — ve a <b style={{margin:"0 4px",cursor:"pointer",textDecoration:"underline"}} onClick={()=>setPg("my-bank")}>Mi Cuenta</b></div>}
      {goal>0&&<div className="gl"><div className="gt">🎯 Meta: {$(goal)}</div><div className="pb"><div className="pf" style={{width:`${prog}%`}}/></div><div className="gs">Llevas <b>{$(st.totalCommission)}</b> — Faltan <b>{$(Math.max(goal-st.totalCommission,0))}</b></div></div>}
      <div className="mg">
        <div className="mc"><div className="ml">OK hoy</div><div className="mv p">{st.todayOk}</div><div className="ms2">Escala: {st.todayRate>0?$(st.todayRate):"—"}</div></div>
        <div className="mc"><div className="ml">Comisión hoy</div><div className="mv g">{$(st.todayCommission)}</div></div>
        <div className="mc"><div className="ml">Acumulado mes</div><div className="mv b">{$(st.totalCommission)}</div></div>
        <div className="mc"><div className="ml">Rechazadas</div><div className={`mv ${rejC>0?"r":"g"}`}>{rejC}</div></div>
      </div>
      {nt&&st.todayOk>0&&<div className="nt"><span style={{fontSize:24}}>🚀</span><div className="nt-t">¡Faltan <b>{nt.need}</b> para <b>{$(nt.r)}/venta</b>!</div></div>}
      <div className="tc">
        <div className="cd"><div className="ch"><div className="ct">📊 Comisiones por día</div></div><table><thead><tr><th>Fecha</th><th>OK</th><th>Rech.</th><th>Escala</th><th>Comisión</th></tr></thead><tbody>
          {st.dailyCommissions.slice(0,15).map(d=> <tr key={d.date}><td>{fd(d.date)}</td><td style={{fontWeight:600,color:"var(--ok)"}}>{d.ok}</td><td style={{color:d.rejected>0?"var(--er)":"var(--mt)"}}>{d.rejected}</td><td>{d.rate>0?$(d.rate):"—"}</td><td style={{fontFamily:"Outfit",fontWeight:700}}>{$(d.commission)}</td></tr>)}
          <tr style={{background:"#faf9fd"}}><td style={{fontWeight:700}}>TOTAL</td><td style={{fontWeight:700,color:"var(--ok)"}}>{st.totalOk}</td><td style={{fontWeight:700,color:"var(--er)"}}>{st.totalRejected}</td><td></td><td style={{fontFamily:"Outfit",fontWeight:800,color:"var(--pu)"}}>{$(st.totalCommission)}</td></tr></tbody></table></div>
        <div className="rc"><div className="ch"><div className="ct">🏆 Top 5</div></div>
          <div style={{padding:"4px 8px 0"}}><div className="tb"><button className={`ti ${rTab==="monthly"?"a":""}`} onClick={()=>setRTab("monthly")}>Mes</button><button className={`ti ${rTab==="weekly"?"a":""}`} onClick={()=>setRTab("weekly")}>Semana</button></div></div>
          {rk.slice(0,5).map((r,i)=> <div className="ri" key={r.id}><div className="rp" style={{color:i===0?"#f59e0b":i===1?"#9ca3af":i===2?"#b45309":"var(--mt)"}}>{i===0?"🥇":i===1?"🥈":i===2?"🥉":`${i+1}°`}</div><div className={`rn ${r.id===cur.id?"me":""}`}>{r.fullName.split(" ").slice(0,2).join(" ")}{r.id===cur.id&&" (Tú)"}</div><div className="ra">{$(r.commission)}</div></div>)}
          {myR>5&&<div className="ri" style={{background:"#faf5ff"}}><div className="rp">{myR}°</div><div className="rn me">{fn} (Tú)</div><div className="ra">{$(rk[myR-1]?.commission||0)}</div></div>}
        </div></div></div>);};

  const pgReg=()=>{const ae=users.filter(u=>u.role==="executive"&&u.active&&u.id!==cur.id);
    return (<div className="fade"><div className="ph"><div className="pg">Registrar <span>Venta</span></div></div><div style={{maxWidth:560}}><div className="cd" style={{padding:28}}>
      {sOk&&<div style={{background:"#d1fae5",border:"1px solid #059669",borderRadius:12,padding:"12px 18px",marginBottom:20,color:"#065f46",fontWeight:600}}>✅ ¡Registrada!</div>}
      <div style={{background:"#f8f7fc",borderRadius:10,padding:"12px 16px",marginBottom:20,fontSize:14,display:"flex",alignItems:"center",gap:10}}><div className="sb-av" style={{width:32,height:32,fontSize:12}}>{ini}</div><div><div style={{fontWeight:600}}>{cur.fullName}</div><div style={{fontSize:12,color:"var(--mt)"}}>@{cur.username}</div></div></div>
      <div className="fg"><label className="fl">Fecha</label><input type="date" className={`fi ${sErr.date?"er":""}`} value={sFrm.date} onChange={e=>setSFrm({...sFrm,date:e.target.value})}/></div>
      <div className="fg"><label className="fl">Teléfono</label><input className={`fi ${sErr.phone?"er":""}`} placeholder="56912345678" value={sFrm.phone} onChange={e=>setSFrm({...sFrm,phone:e.target.value.replace(/\D/g,"")})} maxLength={11}/><div className="fh">56 + 9 dígitos</div>{sErr.phone&&<div className="fe">{sErr.phone}</div>}</div>
      <div className="fg"><label className="fl">RUT</label><input className={`fi ${sErr.rut?"er":""}`} placeholder="183274765" value={sFrm.rut} onChange={e=>setSFrm({...sFrm,rut:e.target.value.replace(/\D/g,"")})} maxLength={9}/><div className="fh">Sin puntos ni guion</div>{sErr.rut&&<div className="fe">{sErr.rut}</div>}</div>
      <div className="fg"><label className="fl">Tipo</label><select className="fs" value={sFrm.type} onChange={e=>setSFrm({...sFrm,type:e.target.value})}><option value="normal">Normal</option><option value="exception">Excepción</option></select></div>
      <div className="fg"><label className="fl" style={{display:"flex",alignItems:"center",gap:10}}>¿VDI?<div onClick={()=>setSFrm({...sFrm,isVDI:!sFrm.isVDI,vdiUser:""})} style={{width:44,height:24,borderRadius:12,cursor:"pointer",background:sFrm.isVDI?"var(--vi)":"#d1d5db",position:"relative",transition:".2s"}}><div style={{width:18,height:18,borderRadius:9,background:"#fff",position:"absolute",top:3,left:sFrm.isVDI?23:3,transition:".2s",boxShadow:"0 1px 3px rgba(0,0,0,.2)"}}/></div></label></div>
      {sFrm.isVDI&&<div className="fg"><label className="fl">Usuario VDI</label><select className={`fs ${sErr.vdiUser?"er":""}`} value={sFrm.vdiUser} onChange={e=>setSFrm({...sFrm,vdiUser:e.target.value})}><option value="">Seleccionar...</option>{ae.map(u=> <option key={u.id} value={u.username}>{u.username} — {u.fullName}</option>)}</select></div>}
      <button className="bt bp" style={{width:"100%",justifyContent:"center",marginTop:8}} onClick={regSale}><Ic n="check" s={18}/> Registrar</button>
    </div></div></div>);};

  const pgMySales=()=>{const ms=sales.filter(s=>{const d=new Date(s.date);return s.executiveId===cur.id&&d.getMonth()===salesMonth&&d.getFullYear()===salesYear;}).sort((a,b)=>b.id-a.id);
    return (<div className="fade"><div className="ph" style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
      <div className="pg">Mis <span>Ventas</span></div>
      <div className="msel">📅 <select value={`${salesYear}-${salesMonth}`} onChange={e=>{const[y,m]=e.target.value.split("-");setSalesYear(Number(y));setSalesMonth(Number(m));}}>
        {Array.from({length:6},(_,i)=>{const d=new Date(new Date().getFullYear(),new Date().getMonth()-i,1);return <option key={i} value={`${d.getFullYear()}-${d.getMonth()}`}>{gMN(d.getMonth())} {d.getFullYear()}</option>;})}
      </select></div></div>
    <div className="cd"><div className="ch"><div className="ct">{gMN(salesMonth)} — {ms.length} ventas</div></div><div style={{overflowX:"auto"}}><table><thead><tr><th>Fecha</th><th>Tel.</th><th>RUT</th><th>Tipo</th><th>VDI</th><th>Estado</th></tr></thead><tbody>
      {ms.slice(0,60).map(s=> <tr key={s.id}><td>{fd(s.date)}</td><td style={{fontFamily:"monospace"}}>{s.phone}</td><td style={{fontFamily:"monospace"}}>{s.rut}</td><td>{s.type==="normal"?"N":"E"}</td><td>{s.isVDI?"Sí":"—"}</td><td><span className={`bg bg-${s.status}`}>{s.status==="ok"?"OK":s.status==="rejected"?"Rech.":"Pend."}</span></td></tr>)}
    </tbody></table></div></div></div>);};

  const pgMyRej=()=>{const rej=sales.filter(s=>s.executiveId===cur.id&&s.status==="rejected").sort((a,b)=>b.id-a.id);
    return (<div className="fade"><div className="ph"><div className="pg">Mis <span>Rechazos</span></div></div>
    {rej.length===0?<div className="cd"><div className="em"><p>🎉 ¡Sin rechazos!</p></div></div>:
      rej.map(s=> <div className="cd" key={s.id} style={{padding:20,marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"start",flexWrap:"wrap",gap:12}}>
          <div><div style={{fontSize:13,color:"var(--mt)",marginBottom:4}}>{fdf(s.date)} · {s.phone}</div><div style={{fontWeight:600,color:"var(--er)",marginBottom:4}}>❌ {s.rejectionReason}</div>{s.rejectionDetail&&s.rejectionDetail!=="null"&&<div style={{fontSize:13,color:"var(--mt)"}}>📝 {s.rejectionDetail}</div>}</div>
          {s.managedByExec?<span className={`bg ${s.appealed?"bg-appealed":"bg-"+s.recoveryAction}`}>{s.appealed?"📤 Apelada":s.recoveryAction==="recovered"?"✅ Gestionada":"🚫 No recup."}</span>:
            <div style={{display:"flex",gap:8}}><button className="bt bo bsm" onClick={()=>{setRjSel(s.id);setRjAct("recovered")}}>Recuperé</button><button className="bt bd bsm" onClick={()=>{setRjSel(s.id);setRjAct("not_recoverable")}}>No se puede</button></div>}
        </div>
        {rjSel===s.id&&<div style={{marginTop:14,paddingTop:14,borderTop:"1px solid var(--bd)"}}>
          <textarea className="fi" rows={2} placeholder={rjAct==="recovered"?"¿Qué hiciste?":"¿Por qué no?"} value={rjCmt} onChange={e=>setRjCmt(e.target.value)} style={{resize:"vertical"}}/>
          <div style={{display:"flex",gap:8,marginTop:10}}><button className="bt bp bsm" onClick={()=>handleRecov(s.id,rjAct)}>Confirmar</button><button className="bt bs bsm" onClick={()=>{setRjSel(null);setRjCmt("")}}>Cancelar</button></div>
        </div>}</div>)}</div>);};

  const pgBank=()=> (<div className="fade"><div className="ph"><div className="pg">Mi <span>Cuenta Bancaria</span> 🏦</div></div>
    {cur.bank&&<div className="bank-card"><div className="bank-label">Número de cuenta</div><div className="bank-val" style={{letterSpacing:2,fontSize:18}}>{cur.bank.account.replace(/(\d{4})/g,"$1 ").trim()}</div>
      <div style={{display:"flex",gap:24,marginTop:16,flexWrap:"wrap"}}><div><div className="bank-label">Titular</div><div style={{fontSize:14}}>{cur.bank.holder}</div></div><div><div className="bank-label">Cédula</div><div style={{fontSize:14}}>{cur.bank.cedula}</div></div><div><div className="bank-label">Banco</div><div style={{fontSize:14}}>{cur.bank.bankName}</div></div><div><div className="bank-label">Tipo</div><div style={{fontSize:14,textTransform:"capitalize"}}>{cur.bank.type}</div></div></div></div>}
    <div style={{maxWidth:560}}><div className="cd" style={{padding:28}}>
      {bankOk&&<div style={{background:"#d1fae5",border:"1px solid #059669",borderRadius:12,padding:"12px 18px",marginBottom:20,color:"#065f46",fontWeight:600}}>✅ Datos guardados</div>}
      <div className="fg"><label className="fl">Número de cuenta (20 dígitos)</label><input className={`fi ${bankErr.account?"er":""}`} placeholder="01020100000000123456" value={bankFrm.account} onChange={e=>setBankFrm({...bankFrm,account:e.target.value.replace(/\D/g,"")})} maxLength={20}/><div className="fh">Exactamente 20 dígitos numéricos</div>{bankErr.account&&<div className="fe">{bankErr.account}</div>}</div>
      <div className="fg"><label className="fl">Nombre del titular</label><input className={`fi ${bankErr.holder?"er":""}`} placeholder="José Ángel Lugo" value={bankFrm.holder} onChange={e=>setBankFrm({...bankFrm,holder:e.target.value})}/>{bankErr.holder&&<div className="fe">{bankErr.holder}</div>}</div>
      <div className="fg"><label className="fl">Cédula del titular</label><input className={`fi ${bankErr.cedula?"er":""}`} placeholder="12345678" value={bankFrm.cedula} onChange={e=>setBankFrm({...bankFrm,cedula:e.target.value.replace(/\D/g,"")})}/>{bankErr.cedula&&<div className="fe">{bankErr.cedula}</div>}</div>
      <div className="fg"><label className="fl">Banco</label><select className={`fs ${bankErr.bankName?"er":""}`} value={bankFrm.bankName} onChange={e=>setBankFrm({...bankFrm,bankName:e.target.value})}><option value="">Seleccionar banco...</option>{BANKS.map(b=> <option key={b} value={b}>{b}</option>)}</select>{bankErr.bankName&&<div className="fe">{bankErr.bankName}</div>}</div>
      <div className="fg"><label className="fl">Tipo de cuenta</label><select className="fs" value={bankFrm.type} onChange={e=>setBankFrm({...bankFrm,type:e.target.value})}><option value="ahorro">Ahorro</option><option value="corriente">Corriente</option></select></div>
      <button className="bt bp" style={{width:"100%",justifyContent:"center"}} onClick={saveBank}><Ic n="check" s={18}/> Guardar Datos Bancarios</button>
    </div></div></div>);

  const pgGoal=()=>{const st=getStats(cur.id);const prog=goal>0?Math.min((st.totalCommission/goal)*100,100):0;const rem=Math.max(goal-st.totalCommission,0);const wl=wdL();
    return (<div className="fade"><div className="ph"><div className="pg">Mi <span>Meta</span> 🎯</div></div><div style={{maxWidth:560}}><div className="cd" style={{padding:28}}>
      <div className="fg"><label className="fl">Meta mensual (USD)</label><input type="number" className="fi" value={goal} onChange={e=>setGoal(Number(e.target.value))} style={{fontSize:24,fontFamily:"Outfit",fontWeight:700,textAlign:"center"}}/></div>
      <div className="pb" style={{height:16,background:"#e5e7eb"}}><div className="pf" style={{width:`${prog}%`}}/></div>
      <div style={{textAlign:"center",marginTop:20}}><div style={{fontSize:36,fontFamily:"Outfit",fontWeight:800,color:"var(--pu)"}}>{$(st.totalCommission)}</div><div style={{fontSize:14,color:"var(--mt)"}}>de {$(goal)}</div></div>
      <div style={{background:"#f8f7fc",borderRadius:12,padding:20,marginTop:20,textAlign:"center"}}><div style={{fontSize:20,fontWeight:700,fontFamily:"Outfit"}}>{rem>0?`Faltan ${$(rem)}`:"🎉 ¡Meta alcanzada!"}</div>
        {wl>0&&rem>0&&<div style={{fontSize:14,color:"var(--mt)",marginTop:8}}>~<b>{Math.ceil(rem/(wl*2*1.7))}</b> ventas/día en <b>{wl}</b> días</div>}</div></div></div></div>);};

  const pgAgenda=()=>{const myAg=agendas.filter(a=>a.executiveId===cur.id).sort((a,b)=>a.date.localeCompare(b.date)||a.callTime.localeCompare(b.callTime));
    const todAg=myAg.filter(a=>a.date===tdy()),futAg=myAg.filter(a=>a.date>tdy()),pastAg=myAg.filter(a=>a.date<tdy());
    return (<div className="fade"><div className="ph" style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}><div className="pg">Mis <span>Agendados</span></div><button className="bt bp" onClick={()=>setShowAgFrm(true)}><Ic n="plus" s={18}/> Nuevo</button></div>
      {todAg.length>0&&<><div style={{fontSize:14,fontWeight:600,marginBottom:12,color:"var(--vi)"}}>📌 Hoy ({todAg.length})</div>
        {todAg.map(a=> <div className="ag-item" key={a.id} style={{borderLeft:"3px solid var(--vi)"}}><div style={{flex:1}}><div style={{fontWeight:600}}>{a.clientName}</div><div style={{fontSize:13,color:"var(--mt)",fontFamily:"monospace"}}>{a.phone}</div><div style={{fontSize:13,color:"var(--vi)",fontWeight:600}}>🕐 {a.callTime}</div>{a.note&&<div style={{fontSize:12,color:"var(--mt)",marginTop:4}}>📝 {a.note}</div>}</div><button className="bt bd bsm" onClick={()=>delAg(a.id)}><Ic n="trash" s={14}/></button></div>)}</>}
      {futAg.length>0&&<><div style={{fontSize:14,fontWeight:600,marginBottom:12,marginTop:20,color:"var(--wr)"}}>📆 Próximos ({futAg.length})</div>
        {futAg.map(a=> <div className="ag-item" key={a.id}><div style={{flex:1}}><div style={{fontWeight:600}}>{a.clientName}</div><div style={{fontSize:13,color:"var(--mt)",fontFamily:"monospace"}}>{a.phone}</div><div style={{fontSize:13,color:"var(--vi)"}}>📅 {fdf(a.date)} · 🕐 {a.callTime}</div></div><button className="bt bd bsm" onClick={()=>delAg(a.id)}><Ic n="trash" s={14}/></button></div>)}</>}
      {pastAg.length>0&&<><div style={{fontSize:14,fontWeight:600,marginBottom:12,marginTop:20,color:"var(--mt)"}}>⏰ Pasados ({pastAg.length})</div>{pastAg.map(a=> <div className="ag-item" key={a.id} style={{opacity:.5}}><div style={{flex:1}}><div style={{fontWeight:600}}>{a.clientName}</div><div style={{fontSize:13,color:"var(--mt)"}}>{fdf(a.date)}</div></div><button className="bt bd bsm" onClick={()=>delAg(a.id)}><Ic n="trash" s={14}/></button></div>)}</>}
      {myAg.length===0&&<div className="cd"><div className="em"><p>Sin agendados</p></div></div>}
      {showAgFrm&&<div className="ov" onClick={()=>setShowAgFrm(false)}><div className="md" onClick={e=>e.stopPropagation()}>
        <div className="mdt">Nuevo Agendado</div>
        <div className="fg"><label className="fl">Cliente</label><input className="fi" value={agFrm.clientName} onChange={e=>setAgFrm({...agFrm,clientName:e.target.value})}/></div>
        <div className="fg"><label className="fl">Teléfono</label><input className="fi" placeholder="56912345678" value={agFrm.phone} onChange={e=>setAgFrm({...agFrm,phone:e.target.value.replace(/\D/g,"")})} maxLength={11}/></div>
        <div className="fg"><label className="fl">Fecha</label><input type="date" className="fi" value={agFrm.date} onChange={e=>setAgFrm({...agFrm,date:e.target.value})}/></div>
        <div className="fg"><label className="fl">Hora</label><input type="time" className="fi" value={agFrm.callTime} onChange={e=>setAgFrm({...agFrm,callTime:e.target.value})}/></div>
        <div className="fg"><label className="fl">Nota</label><input className="fi" value={agFrm.note} onChange={e=>setAgFrm({...agFrm,note:e.target.value})}/></div>
        <div className="ma"><button className="bt bs" onClick={()=>setShowAgFrm(false)}>Cancelar</button><button className="bt bp" onClick={addAg}><Ic n="check" s={18}/> Guardar</button></div></div></div>}
    </div>);};

  // ===== ADMIN PAGES =====
  const pgADash=()=>{const execs=users.filter(u=>u.role==="executive"&&u.active);
    const ms=sales.filter(s=>{const d=new Date(s.date);return d.getMonth()===dashMonth&&d.getFullYear()===dashYear;});
    const tOk=ms.filter(s=>s.status==="ok").length,tRej=ms.filter(s=>s.status==="rejected").length,tPend=ms.filter(s=>s.status==="pending").length;
    const eS2=execs.map(e=>({...e,...getStats(e.id,dashMonth,dashYear)})).sort((a,b)=>b.totalCommission-a.totalCommission);let tc=0;eS2.forEach(e=>{tc+=e.totalCommission;});
    return (<div className="fade"><div className="ph" style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}><div className="pg">Dashboard <span>Admin</span></div>
      <div className="msel">📅 <select value={`${dashYear}-${dashMonth}`} onChange={e=>{const[y,m]=e.target.value.split("-");setDashYear(Number(y));setDashMonth(Number(m));}}>
        {Array.from({length:6},(_,i)=>{const d=new Date(new Date().getFullYear(),new Date().getMonth()-i,1);return <option key={i} value={`${d.getFullYear()}-${d.getMonth()}`}>{gMN(d.getMonth())} {d.getFullYear()}</option>;})}
      </select></div></div>
    <div className="mg">
      <div className="mc"><div className="ml">Total</div><div className="mv p">{ms.length}</div></div>
      <div className="mc"><div className="ml">OK</div><div className="mv g">{tOk}</div><div className="ms2">{ms.length?(tOk/ms.length*100).toFixed(1):0}%</div></div>
      <div className="mc"><div className="ml">Rechazadas</div><div className="mv r">{tRej}</div></div>
      <div className="mc"><div className="ml">Pendientes</div><div className="mv o">{tPend}</div></div>
      <div className="mc"><div className="ml">Comisiones</div><div className="mv b">{$(tc)}</div></div>
      <div className="mc"><div className="ml">Ejecutivos</div><div className="mv p">{execs.length}</div></div>
    </div>
    <div className="cd"><div className="ch"><div className="ct">📊 {gMN(dashMonth)} {dashYear}</div><button className="bt bs bsm" onClick={exportAdm}><Ic n="download" s={16}/> Excel</button></div><div style={{overflowX:"auto"}}><table><thead><tr><th>Ejecutivo</th><th>Usuario</th><th>Total</th><th>OK</th><th>Rech.</th><th>Comisión</th></tr></thead><tbody>
      {eS2.map(e=> <tr key={e.id}><td style={{fontWeight:500}}>{e.fullName}</td><td style={{fontFamily:"monospace",color:"var(--mt)"}}>{e.username}</td><td style={{fontWeight:600}}>{e.totalOk+e.totalRejected+e.totalPending}</td><td style={{color:"var(--ok)",fontWeight:600}}>{e.totalOk}</td><td style={{color:e.totalRejected>0?"var(--er)":"var(--mt)"}}>{e.totalRejected}</td><td style={{fontFamily:"Outfit",fontWeight:700,color:"var(--pu)"}}>{$(e.totalCommission)}</td></tr>)}
    </tbody></table></div></div></div>);};

  const pgASales=()=>{const f=getFS();const execs=users.filter(u=>u.role==="executive");
    return (<div className="fade"><div className="ph"><div className="pg">Gestión <span>Ventas</span></div></div>
    <div className="ft"><div className="fg"><label className="fl">Ejecutivo</label><select className="fs" value={aFlt.exec} onChange={e=>setAFlt({...aFlt,exec:e.target.value})}><option value="">Todos</option>{execs.map(u=> <option key={u.id} value={u.username}>{u.username}</option>)}</select></div>
      <div className="fg"><label className="fl">Estado</label><select className="fs" value={aFlt.status} onChange={e=>setAFlt({...aFlt,status:e.target.value})}><option value="">Todos</option><option value="ok">OK</option><option value="rejected">Rechazada</option><option value="pending">Pendiente</option></select></div>
      <button className="bt bs bsm" onClick={()=>setAFlt({exec:"",status:""})}>Limpiar</button><button className="bt bs bsm" onClick={exportAdm}><Ic n="download" s={16}/> Excel</button></div>
    <div className="cd"><div className="ch"><div className="ct">{f.length} ventas</div></div><div style={{overflowX:"auto"}}><table><thead><tr><th>Fecha</th><th>Ejecutivo</th><th>Tel.</th><th>RUT</th><th>Estado</th><th></th></tr></thead><tbody>
      {f.slice(0,80).map(s=>{const ex=users.find(u=>u.id===s.executiveId);return <tr key={s.id}><td>{fd(s.date)}</td><td style={{fontSize:12}}>{ex?.fullName?.split(" ").slice(0,2).join(" ")}<br/><span style={{fontFamily:"monospace",color:"var(--mt)"}}>{s.username}</span></td><td style={{fontFamily:"monospace",fontSize:12}}>{s.phone}</td><td style={{fontFamily:"monospace",fontSize:12}}>{s.rut}</td><td><span className={`bg bg-${s.status}`}>{s.status==="ok"?"OK":s.status==="rejected"?"Rech.":"Pend."}</span></td><td><button className="bt bs bsm" onClick={()=>{setES(s.id);setEFrm({status:s.status,rr:s.rejectionReason||"",rd:s.rejectionDetail||""})}}><Ic n="edit" s={14}/></button></td></tr>;})}
    </tbody></table></div></div>
    {eS&&<div className="ov" onClick={()=>setES(null)}><div className="md" onClick={e=>e.stopPropagation()}>
      <div className="mdt">Editar Estado</div>
      <div className="fg"><label className="fl">Estado</label><select className="fs" value={eFrm.status} onChange={e=>setEFrm({...eFrm,status:e.target.value})}><option value="pending">Pendiente</option><option value="ok">OK</option><option value="rejected">Rechazada</option></select></div>
      {eFrm.status==="rejected"&&<><div className="fg"><label className="fl">Motivo</label><select className="fs" value={eFrm.rr} onChange={e=>setEFrm({...eFrm,rr:e.target.value})}><option value="">Seleccionar...</option>{REJ_R.map(r=> <option key={r} value={r}>{r}</option>)}</select></div>
      <div className="fg"><label className="fl">Observación</label><textarea className="fi" rows={2} value={eFrm.rd} onChange={e=>setEFrm({...eFrm,rd:e.target.value})}/></div></>}
      <div className="ma"><button className="bt bs" onClick={()=>setES(null)}>Cancelar</button><button className="bt bp" onClick={saveEdit}><Ic n="check" s={18}/> Guardar</button></div></div></div>}</div>);};

  // ===== ADMIN REJECTIONS WITH APPEALED =====
  const pgARej=()=>{const rej=sales.filter(s=>s.status==="rejected");
    const unmanaged=rej.filter(s=>!s.managedByExec);const managed=rej.filter(s=>s.managedByExec&&!s.appealed);const appealed=rej.filter(s=>s.appealed);
    const shown=rejTab==="unmanaged"?unmanaged:rejTab==="managed"?managed:appealed;
    return (<div className="fade"><div className="ph" style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:12}}>
      <div className="pg">Seguimiento <span>Rechazos</span></div><button className="bt bs bsm" onClick={exportRecov}><Ic n="download" s={16}/> Excel</button></div>
    <div className="tb">
      <button className={`ti ${rejTab==="unmanaged"?"a":""}`} onClick={()=>setRejTab("unmanaged")}>Sin gestionar ({unmanaged.length})</button>
      <button className={`ti ${rejTab==="managed"?"a":""}`} onClick={()=>setRejTab("managed")}>Gestionadas ({managed.length})</button>
      <button className={`ti ${rejTab==="appealed"?"a":""}`} onClick={()=>setRejTab("appealed")}>Apeladas ({appealed.length})</button>
    </div>
    <div className="cd"><div style={{overflowX:"auto"}}><table><thead><tr><th>Fecha</th><th>Ejecutivo</th><th>Tel.</th><th>Motivo</th>
      {(rejTab==="managed"||rejTab==="appealed")&&<th>Acción ejec.</th>}
      {rejTab==="managed"&&<th>Apelar</th>}
      {rejTab==="appealed"&&<th>Estado</th>}
    </tr></thead><tbody>
      {shown.slice(0,60).map(s=>{const ex=users.find(u=>u.id===s.executiveId);return <tr key={s.id}><td>{fd(s.date)}</td><td style={{fontSize:12}}>{ex?.fullName?.split(" ").slice(0,2).join(" ")}</td><td style={{fontFamily:"monospace",fontSize:12}}>{s.phone}</td><td style={{fontSize:12}}>{s.rejectionReason}</td>
        {(rejTab==="managed"||rejTab==="appealed")&&<td style={{fontSize:12}}>{s.recoveryComment}</td>}
        {rejTab==="managed"&&<td><button className="bt bw bsm" onClick={()=>appealSale(s.id)}>📤 Apelar</button></td>}
        {rejTab==="appealed"&&<td><span className="bg bg-appealed">📤 En revisión</span></td>}
      </tr>;})}
      {shown.length===0&&<tr><td colSpan={6} style={{textAlign:"center",padding:40,color:"var(--mt)"}}>Sin ventas en esta categoría</td></tr>}
    </tbody></table></div></div></div>);};

  const pgUpload=()=> (<div className="fade"><div className="ph"><div className="pg">Subir <span>Excel</span></div></div><div style={{maxWidth:700}}>
    <div className="uz" onClick={()=>fRef.current?.click()}><input ref={fRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleUpload}/><div style={{fontSize:48,marginBottom:12}}>📤</div><div style={{fontSize:16,fontWeight:600,marginBottom:8}}>Subir Excel del back office</div></div>
    <button className="bt bs bsm" onClick={dlTpl} style={{marginBottom:24}}><Ic n="download" s={16}/> Plantilla</button>
    {uData?.error&&<div className="pvw">❌ Error de formato</div>}
    {uData&&!uData.error&&<>
      <div className="pvok">✅ {uData.matched.length} coinciden · {uData.updated.length} cambiarán</div>
      {uData.appealedOk.length>0&&<div className="pvok">🎉 <b>{uData.appealedOk.length}</b> apeladas pasarán a OK</div>}
      {uData.unmatched.length>0&&<div className="pvw">⚠️ {uData.unmatched.length} NO están en el sistema</div>}
      <div style={{display:"flex",gap:10}}><button className="bt bp" onClick={confirmUpload}><Ic n="check" s={18}/> Confirmar</button><button className="bt bs" onClick={()=>{setUData(null);if(fRef.current)fRef.current.value="";}}>Cancelar</button></div></>}
    {uRes&&<div className="nf nf-g" style={{marginTop:20}}>✅ {uRes.matched} coinc · {uRes.updated} actual{uRes.appealedOk>0?` · ${uRes.appealedOk} apeladas→OK`:""}</div>}
  </div></div>);

  const pgPayroll=()=>{const ok=pRng.from&&pRng.to;const f=ok?sales.filter(s=>s.date>=pRng.from&&s.date<=pRng.to):[];const execs=users.filter(u=>u.role==="executive");
    const preview=ok?execs.map(e=>{const es=f.filter(s=>s.executiveId===e.id);if(!es.length)return null;
      const dm={};es.forEach(s=>{if(!dm[s.date])dm[s.date]={ok:0};if(s.status==="ok")dm[s.date].ok++;});let tc=0;Object.values(dm).forEach(d=>{tc+=calcDay(d.ok).total;});
      return{...e,tot:es.length,ok:es.filter(s=>s.status==="ok").length,rej:es.filter(s=>s.status==="rejected").length,comm:tc};}).filter(Boolean).sort((a,b)=>b.comm-a.comm):[];
    const noBankExecs=preview.filter(e=>!e.bank);
    return (<div className="fade"><div className="ph"><div className="pg">Generar <span>Nómina</span></div></div><div style={{maxWidth:700}}><div className="cd" style={{padding:24}}>
      <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:20}}>
        <div className="fg" style={{flex:1,minWidth:180}}><label className="fl">Desde</label><input type="date" className="fi" value={pRng.from} onChange={e=>setPRng({...pRng,from:e.target.value})}/></div>
        <div className="fg" style={{flex:1,minWidth:180}}><label className="fl">Hasta</label><input type="date" className="fi" value={pRng.to} onChange={e=>setPRng({...pRng,to:e.target.value})}/></div></div>
      {ok&&preview.length>0&&<>
        {noBankExecs.length>0&&<div className="nf nf-r">⚠️ {noBankExecs.length} ejecutivo{noBankExecs.length>1?"s":""} sin datos bancarios: {noBankExecs.map(e=>e.username).join(", ")}</div>}
        <div style={{overflowX:"auto"}}><table><thead><tr><th>Ejecutivo</th><th>Total</th><th>OK</th><th>Rech.</th><th>Comisión</th><th>Banco</th></tr></thead><tbody>
          {preview.map(e=> <tr key={e.id}><td style={{fontWeight:500}}>{e.fullName}</td><td>{e.tot}</td><td style={{color:"var(--ok)",fontWeight:600}}>{e.ok}</td><td style={{color:"var(--er)"}}>{e.rej}</td><td style={{fontFamily:"Outfit",fontWeight:700,color:"var(--pu)"}}>{$(e.comm)}</td><td style={{fontSize:12}}>{e.bank?`${e.bank.bankName} ···${e.bank.account.slice(-4)}`:<span style={{color:"var(--er)"}}>Sin datos</span>}</td></tr>)}
          <tr style={{background:"#faf9fd"}}><td style={{fontWeight:700}}>TOTAL</td><td style={{fontWeight:700}}>{preview.reduce((a,e)=>a+e.tot,0)}</td><td style={{fontWeight:700,color:"var(--ok)"}}>{preview.reduce((a,e)=>a+e.ok,0)}</td><td style={{fontWeight:700,color:"var(--er)"}}>{preview.reduce((a,e)=>a+e.rej,0)}</td><td style={{fontFamily:"Outfit",fontWeight:800,color:"var(--pu)"}}>{$(preview.reduce((a,e)=>a+e.comm,0))}</td><td></td></tr>
        </tbody></table></div>
        <button className="bt bp" style={{marginTop:20}} onClick={exportPayroll}><Ic n="download" s={18}/> Descargar Nómina</button></>}
    </div></div></div>);};

  const pgUsers=()=>{const mgbl=isSA?users:users.filter(u=>u.role==="executive");
    return (<div className="fade"><div className="ph" style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div className="pg">Gestión <span>Usuarios</span></div><button className="bt bp" onClick={()=>setShowCU(true)}><Ic n="plus" s={18}/> Nuevo</button></div>
    <div className="cd"><div style={{overflowX:"auto"}}><table><thead><tr><th>Nombre</th><th>Usuario</th><th>Rol</th><th>Estado</th><th>Banco</th><th></th></tr></thead><tbody>
      {mgbl.map(u=> <tr key={u.id}><td style={{fontWeight:500}}>{u.fullName}</td><td style={{fontFamily:"monospace"}}>{u.username}</td><td><span className={`bg bg-${u.role}`}>{u.role==="super_admin"?"Coord.":u.role==="admin"?"Admin":"Ejec."}</span></td><td><span className={`bg bg-${u.active?"active":"inactive"}`}>{u.active?"Activo":"Inactivo"}</span></td><td style={{fontSize:12}}>{u.bank?`✅ ${u.bank.bankName}`:<span style={{color:"var(--er)"}}>❌</span>}</td><td><div style={{display:"flex",gap:6}}>
        {u.id!==cur.id&&(u.role==="executive"||isSA)&&<button className={`bt bsm ${u.active?"bd":"bo"}`} onClick={()=>toggleUsr(u.id)}>{u.active?"Desact.":"React."}</button>}
        {isSA&&u.id!==cur.id&&<button className="bt bs bsm" onClick={()=>{setShowCP(u.id);setNP("")}}><Ic n="key" s={14}/></button>}
      </div></td></tr>)}</tbody></table></div></div>
    {showCU&&<div className="ov" onClick={()=>setShowCU(false)}><div className="md" onClick={e=>e.stopPropagation()}>
      <div className="mdt">Crear Usuario</div>
      <div className="fg"><label className="fl">Nombre completo</label><input className="fi" value={nU.fullName} onChange={e=>setNU({...nU,fullName:e.target.value})}/></div>
      <div className="fg"><label className="fl">Usuario</label><input className="fi" value={nU.username} onChange={e=>setNU({...nU,username:e.target.value.toLowerCase().replace(/\s/g,"")})}/></div>
      <div className="fg"><label className="fl">Contraseña</label><input className="fi" type="password" value={nU.password} onChange={e=>setNU({...nU,password:e.target.value})}/></div>
      <div className="fg"><label className="fl">Rol</label><select className="fs" value={nU.role} onChange={e=>setNU({...nU,role:e.target.value})}><option value="executive">Ejecutivo</option>{isSA&&<option value="admin">Admin</option>}</select></div>
      <div className="ma"><button className="bt bs" onClick={()=>setShowCU(false)}>Cancelar</button><button className="bt bp" onClick={createUser}><Ic n="check" s={18}/> Crear</button></div></div></div>}
    {showCP&&<div className="ov" onClick={()=>setShowCP(null)}><div className="md" onClick={e=>e.stopPropagation()}>
      <div className="mdt">Cambiar Contraseña</div><p style={{fontSize:14,color:"var(--mt)",marginBottom:16}}>Usuario: <b>{users.find(u=>u.id===showCP)?.username}</b></p>
      <div className="fg"><label className="fl">Nueva contraseña</label><input className="fi" type="password" value={nP} onChange={e=>setNP(e.target.value)}/></div>
      <div className="ma"><button className="bt bs" onClick={()=>setShowCP(null)}>Cancelar</button><button className="bt bp" onClick={chgPw}><Ic n="check" s={18}/> Cambiar</button></div></div></div>}</div>);};

  const pgLog=()=>{const icons={login:"🔐",sale:"📝",recovery:"✅",upload:"📤",payroll:"📋",user_create:"👤",user_toggle:"🔄",pw_change:"🔑",status_change:"📊",bank:"🏦",appeal:"📤"};
    const colors={login:"#ede9fe",sale:"#d1fae5",recovery:"#dbeafe",upload:"#fef3c7",payroll:"#e0e7ff",user_create:"#d1fae5",user_toggle:"#fef3c7",pw_change:"#ede9fe",status_change:"#fce7f3",bank:"#dbeafe",appeal:"#fef3c7"};
    const f=logs.filter(l=>{if(lgFlt.user&&l.user!==lgFlt.user)return false;if(lgFlt.action&&l.action!==lgFlt.action)return false;return true;});
    return (<div className="fade"><div className="ph"><div className="pg">Log <span>Actividad</span></div></div>
    <div className="ft"><div className="fg"><label className="fl">Usuario</label><select className="fs" value={lgFlt.user} onChange={e=>setLgFlt({...lgFlt,user:e.target.value})}><option value="">Todos</option>{users.map(u=> <option key={u.id} value={u.username}>{u.username}</option>)}</select></div>
      <button className="bt bs bsm" onClick={()=>setLgFlt({user:"",action:""})}>Limpiar</button></div>
    <div className="cd">{f.length===0?<div className="em"><p>Sin registros</p></div>:
      f.slice(0,50).map(l=> <div className="lgi" key={l.id}><div className="lgi-i" style={{background:colors[l.action]||"#f3f4f6"}}>{icons[l.action]||"📌"}</div>
        <div className="lgi-c"><div className="lgi-t"><b>{l.user}</b> · {l.detail}</div><div className="lgi-tm">{fts(l.ts)}</div></div></div>)}
    </div></div>);};

  const renderPage=()=>{
    if(isEx){switch(pg){case "dashboard":return pgDash();case "register":return pgReg();case "my-sales":return pgMySales();case "my-rejections":return pgMyRej();
      case "agenda":return pgAgenda();case "my-bank":return pgBank();case "my-goal":return pgGoal();default:return pgDash();}}
    switch(pg){case "dashboard":return pgADash();case "all-sales":return pgASales();case "rejections":return pgARej();case "upload":return pgUpload();
      case "payroll":return pgPayroll();case "users":return pgUsers();case "log":return pgLog();default:return pgADash();}
  };

  return (<><style>{CSS}</style><div className="app">
    <div className="sb"><div className="sb-h"><div className="sb-b">WOM</div><div className="sb-c">Migraciones</div></div>
      <div className="sb-n"><div className="sb-l">{isEx?"Mi espacio":"Administración"}</div>
        {nav.map(item=> <div key={item.id} className={`si ${pg===item.id?"a":""}`} onClick={()=>setPg(item.id)}><Ic n={item.i} s={18}/>{item.l}</div>)}</div>
      <div className="sb-f"><div className="sb-u"><div className="sb-av">{ini}</div><div className="sb-ui"><div className="sb-un">{fn}</div><div className="sb-ur">{rl}</div></div><button className="sb-o" onClick={logout}><Ic n="logout" s={18}/></button></div></div></div>
    <div className="mn">{renderPage()}</div></div></>);
}
