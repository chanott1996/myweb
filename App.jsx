import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";


/* ---------- color tokens ---------- */
const C = {
  ink:"#0F172A", sub:"#475569", line:"#E2E8F0", line2:"#CBD5E1",
  bg:"#F6F8FA", surface:"#FFFFFF",
  teal:"#0E7C7B", tealDark:"#0B5E5D", tealSoft:"#D7ECEB",
  amber:"#D97706", amberSoft:"#FDE9CE",
  rose:"#E11D48", roseSoft:"#FCE2E8", emerald:"#059669",
  wt:"#0E7C7B", nwt:"#D97706",
};
const mono = { fontFamily:"'IBM Plex Mono', ui-monospace, monospace", fontVariantNumeric:"tabular-nums" };
const display = { fontFamily:"'Space Grotesk', system-ui, sans-serif" };

/* ---------- dilution helpers ---------- */
const LOW_LABELS = {"-1":"0.5","-2":"0.25","-3":"0.125","-4":"0.06","-5":"0.03","-6":"0.016","-7":"0.008","-8":"0.004","-9":"0.002","-10":"0.001"};
function log2Label(k){ if(k>=0) return String(Math.pow(2,k)); return LOW_LABELS[String(k)] || Number(Math.pow(2,k)).toPrecision(2); }

function parseMic(raw){
  if(raw===null||raw===undefined) return null;
  let s=String(raw).trim().replace(/≤/g,"<=").replace(/≥/g,">=").replace(/,/g,".").replace(/\s+/g,"");
  if(!s) return null;
  const m=s.match(/^(<=|>=|<|>|=)?(.+)$/); if(!m) return null;
  const censor=m[1]||""; const num=parseFloat(m[2]);
  if(!isFinite(num)||num<=0) return null;
  let log2=Math.round(Math.log2(num));
  if(censor===">") log2=log2+1;
  return { value:num, censor, log2 };
}

/* ---------- stats ---------- */
function erf(x){ const t=1/(1+0.3275911*Math.abs(x));
  const y=1-(((((1.061405429*t-1.453152027)*t)+1.421413741)*t-0.284496736)*t+0.254829592)*t*Math.exp(-x*x);
  return x>=0?y:-y; }
const normCdf=(x,mean,sd)=>0.5*(1+erf((x-mean)/(sd*Math.SQRT2)));

function ecoffFinder(distMap){
  const dist=[...distMap.entries()].map(([log2,count])=>({log2:Number(log2),count})).sort((a,b)=>a.log2-b.log2);
  if(!dist.length) return null;
  const N=dist.reduce((s,d)=>s+d.count,0); if(N<1) return null;
  let modeI=0; for(let i=1;i<dist.length;i++) if(dist[i].count>dist[modeI].count) modeI=i;
  let best=null;
  for(let end=modeI;end<dist.length;end++){
    const obs=[]; let c=0;
    for(let i=0;i<=end;i++){ c+=dist[i].count; obs.push({x:dist[i].log2,cum:c}); }
    const wtN=c; if(wtN<8||obs.length<3) continue;
    let local=null;
    for(let mean=dist[0].log2-1;mean<=dist[end].log2+1;mean+=0.1){
      for(let sd=0.3;sd<=2.5;sd+=0.05){
        const cdfs=obs.map(p=>normCdf(p.x+0.5,mean,sd));
        let sCC=0,sCC2=0;
        for(let i=0;i<obs.length;i++){ sCC+=cdfs[i]*obs[i].cum; sCC2+=cdfs[i]*cdfs[i]; }
        if(sCC2===0) continue;
        const scale=sCC/sCC2;
        if(scale<wtN*0.9||scale>N*1.6) continue;
        let sse=0; for(let i=0;i<obs.length;i++){ const e=scale*cdfs[i]-obs[i].cum; sse+=e*e; }
        if(!local||sse<local.sse) local={mean,sd,scale,sse};
      }
    }
    if(!local) continue;
    const meanCum=obs.reduce((s,p)=>s+p.cum,0)/obs.length;
    let sstot=0; obs.forEach(p=>sstot+=(p.cum-meanCum)**2);
    const r2=sstot>0?1-local.sse/sstot:1;
    const cand={end,...local,r2,wtN,fittedWT:local.scale};
    if(!best||cand.r2>best.r2+0.0015||(Math.abs(cand.r2-best.r2)<=0.0015&&cand.wtN<best.wtN)) best=cand;
  }
  if(!best) return null;
  const Z={0.95:1.6449,0.975:1.96,0.99:2.3263,0.995:2.5758};
  const dilFor=p=>Math.ceil(best.mean+Z[p]*best.sd-0.5);
  return { mean:best.mean, sd:best.sd, r2:best.r2, fittedWT:best.fittedWT, N,
    e95:dilFor(0.95), e975:dilFor(0.975), e99:dilFor(0.99), e995:dilFor(0.995) };
}

function micPercentiles(dist){
  const sorted=[...dist.entries()].map(([l,c])=>({l:Number(l),c})).sort((a,b)=>a.l-b.l);
  const N=sorted.reduce((s,d)=>s+d.c,0); if(!N) return null;
  const at=p=>{ const target=N*p; let cum=0; for(const d of sorted){ cum+=d.c; if(cum>=target) return d.l; } return sorted[sorted.length-1].l; };
  return { mic50:at(0.5), mic90:at(0.9), min:sorted[0].l, max:sorted[sorted.length-1].l, N };
}

function qcChecks(dist, ecoff, parseStats){
  const flags=[];
  const sorted=[...dist.entries()].map(([l,c])=>({l:Number(l),c})).sort((a,b)=>a.l-b.l);
  const N=sorted.reduce((s,d)=>s+d.c,0);
  if(N<100) flags.push({level:N<30?"error":"warn",code:"low_n",msg:`จำนวนเชื้อ n=${N} ต่ำ — EUCAST แนะนำรวมข้อมูล ≥100 (และจากหลายแหล่ง) เพื่อกำหนด ECOFF ที่เชื่อถือได้`});
  let modeI=0; for(let i=1;i<sorted.length;i++) if(sorted[i].c>sorted[modeI].c) modeI=i;
  if(modeI===0&&sorted.length>1) flags.push({level:"warn",code:"truncated_low",msg:"พีคของการกระจายอยู่ที่ความเข้มข้นต่ำสุดที่ทดสอบ — distribution อาจถูกตัด (truncated) ทำให้ประเมิน wild-type ไม่ครบ ควรทดสอบเจือจางต่ำลงอีก"});
  for(let i=1;i<sorted.length;i++){ if(sorted[i].l-sorted[i-1].l>1){ flags.push({level:"warn",code:"gap",msg:`มีช่องเจือจางที่หายไประหว่าง ${log2Label(sorted[i-1].l)} และ ${log2Label(sorted[i].l)} — ตรวจสอบความครบถ้วนของข้อมูล`}); break; } }
  if(parseStats){
    if(parseStats.lowOff/N>0.1) flags.push({level:"warn",code:"off_low",msg:`ค่า "≤" (off-scale ต่ำ) คิดเป็น ${(parseStats.lowOff/N*100).toFixed(0)}% — distribution อาจถูกตัดด้านล่าง`});
    if(parseStats.highOff/N>0.2) flags.push({level:"info",code:"off_high",msg:`ค่า ">" (off-scale สูง) คิดเป็น ${(parseStats.highOff/N*100).toFixed(0)}% — เชื้อกลุ่มดื้อจำนวนมาก เป็นเรื่องปกติได้แต่ควรตรวจสอบ`});
  }
  if(ecoff&&ecoff.r2<0.97) flags.push({level:"warn",code:"poor_fit",msg:`การ fit log-normal มี R²=${ecoff.r2.toFixed(3)} (ต่ำ) — distribution อาจ multimodal หรือไม่ใช่ log-normal ควรตรวจสอบ ECOFF ด้วยตา`});
  if(ecoff&&ecoff.sd>1.6) flags.push({level:"info",code:"wide_sd",msg:`ส่วนเบี่ยงเบนของ wild-type กว้างผิดปกติ (SD=${ecoff.sd.toFixed(2)} log2) — อาจมีการปนของหลายประชากร`});
  if(!flags.length) flags.push({level:"ok",code:"ok",msg:"ผ่านการตรวจสอบเบื้องต้น — ไม่พบสัญญาณผิดปกติชัดเจน"});
  return flags;
}

function buildAnalyses(records){
  const groups=new Map();
  for(const r of records){
    const key=`${r.organism}||${r.antibiotic}`;
    if(!groups.has(key)) groups.set(key,{organism:r.organism,antibiotic:r.antibiotic,dist:new Map(),lowOff:0,highOff:0});
    const g=groups.get(key);
    g.dist.set(r.log2,(g.dist.get(r.log2)||0)+r.count);
    if(r.censorLow) g.lowOff+=r.count;
    if(r.censorHigh) g.highOff+=r.count;
  }
  const out=[];
  for(const g of groups.values()){
    const ecoff=ecoffFinder(g.dist);
    const pct=micPercentiles(g.dist);
    const flags=qcChecks(g.dist,ecoff,{lowOff:g.lowOff,highOff:g.highOff});
    out.push({...g,ecoff,pct,flags});
  }
  out.sort((a,b)=>(a.organism+a.antibiotic).localeCompare(b.organism+b.antibiotic));
  return out;
}

/* ---------- format detection ---------- */
const ORG_HINTS=["organism","species","bug","เชื้อ","bacteria","strain"];
const AB_HINTS=["antibiotic","antimicrobial","drug","agent","ยา","atb"];
const MIC_HINTS=["mic","value","result","ค่า"];
const COUNT_HINTS=["count","n","number","freq","จำนวน","isolates"];
function looksLikeDilution(h){ const s=String(h).trim().replace(/≤|≥|<=|>=|<|>/g,""); const n=parseFloat(s);
  if(!isFinite(n)||n<=0) return false; return Math.abs(Math.log2(n)-Math.round(Math.log2(n)))<0.06; }
function findCol(headers,hints){ const lower=headers.map(h=>String(h).toLowerCase());
  for(const hint of hints){ const i=lower.findIndex(h=>h.includes(hint)); if(i>=0) return headers[i]; } return null; }
function detectFormat(headers){
  const dilCols=headers.filter(looksLikeDilution);
  const org=findCol(headers,ORG_HINTS), ab=findCol(headers,AB_HINTS), mic=findCol(headers,MIC_HINTS), count=findCol(headers,COUNT_HINTS);
  if(dilCols.length>=3) return {format:"wide",org,ab,dilCols};
  if(mic&&count) return {format:"aggregated",org,ab,mic,count};
  return {format:"isolate",org,ab,mic:mic||headers.find(h=>!ORG_HINTS.concat(AB_HINTS).some(x=>String(h).toLowerCase().includes(x)))};
}
function normalizeRows(rows,mapping){
  const recs=[]; const {format,org,ab,mic,count,dilCols}=mapping;
  for(const row of rows){
    const organism=(org?row[org]:"ไม่ระบุเชื้อ")??"ไม่ระบุเชื้อ";
    const antibiotic=(ab?row[ab]:"ไม่ระบุยา")??"ไม่ระบุยา";
    if(format==="wide"){
      for(const dc of dilCols){ const cnt=parseFloat(row[dc]); if(!isFinite(cnt)||cnt<=0) continue;
        const p=parseMic(dc); if(!p) continue;
        recs.push({organism,antibiotic,log2:p.log2,count:cnt,censorLow:false,censorHigh:false}); }
    } else if(format==="aggregated"){
      const p=parseMic(row[mic]); const cnt=parseFloat(row[count]);
      if(!p||!isFinite(cnt)||cnt<=0) continue;
      recs.push({organism,antibiotic,log2:p.log2,count:cnt,censorLow:p.censor.startsWith("<"),censorHigh:p.censor.startsWith(">")});
    } else {
      const p=parseMic(row[mic]); if(!p) continue;
      recs.push({organism,antibiotic,log2:p.log2,count:1,censorLow:p.censor.startsWith("<"),censorHigh:p.censor.startsWith(">")});
    }
  }
  return recs;
}

/* ---------- persistence via localStorage ---------- */
const LS = window.localStorage;
const hasStore = (()=>{ try{ LS.setItem("__t","1"); LS.removeItem("__t"); return true; }catch{ return false; } })();
async function listIndex(){ try{ const r=LS.getItem("mic:index"); return r?JSON.parse(r):[]; }catch{ return []; } }
async function saveDataset(ds){ if(!hasStore) return; try{ LS.setItem("mic:ds:"+ds.id,JSON.stringify(ds));
  const idx=await listIndex(); const meta={id:ds.id,name:ds.name,at:ds.at,n:ds.records.length};
  LS.setItem("mic:index",JSON.stringify([meta,...idx.filter(m=>m.id!==ds.id)])); }catch(e){ console.error(e); } }
async function loadDataset(id){ if(!hasStore) return null; try{ const r=LS.getItem("mic:ds:"+id); return r?JSON.parse(r):null; }catch{ return null; } }
async function deleteDataset(id){ if(!hasStore) return; try{ LS.removeItem("mic:ds:"+id);
  const idx=await listIndex(); LS.setItem("mic:index",JSON.stringify(idx.filter(m=>m.id!==id))); }catch(e){ console.error(e); } }

/* ---------- UI atoms ---------- */
function Chip({children,tone="teal"}){
  const m={teal:{bg:C.tealSoft,fg:C.tealDark},amber:{bg:C.amberSoft,fg:"#92400E"},rose:{bg:C.roseSoft,fg:"#9F1239"},slate:{bg:"#EEF2F6",fg:C.sub}}[tone];
  return <span className="inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium" style={{background:m.bg,color:m.fg}}>{children}</span>;
}
function Stat({label,value,sub}){
  return <div className="flex flex-col">
    <span className="text-xs uppercase" style={{color:C.sub,letterSpacing:"0.06em"}}>{label}</span>
    <span className="text-lg font-semibold" style={{...mono,color:C.ink}}>{value}</span>
    {sub&&<span className="text-xs" style={{color:C.sub}}>{sub}</span>}
  </div>;
}
function Section({title,children}){
  return <div className="rounded-lg border p-3" style={{borderColor:C.line,background:C.surface}}>
    <h3 className="text-xs font-semibold uppercase mb-2" style={{color:C.sub,letterSpacing:"0.06em"}}>{title}</h3>
    {children}
  </div>;
}

/* ---------- hand-rolled SVG bar chart ---------- */
function MicChart({chartData, eff}){
  const W=720,H=300,padL=38,padR=14,padT=26,padB=64;
  const innerW=W-padL-padR, innerH=H-padT-padB;
  const maxC=Math.max(1,...chartData.map(d=>d.count));
  const n=Math.max(1,chartData.length);
  const slot=innerW/n, bw=Math.min(46,slot*0.74);
  const yFor=c=>padT+innerH*(1-c/maxC);
  const yticks=[]; const steps=4;
  for(let i=0;i<=steps;i++){ const val=Math.round(maxC*i/steps); yticks.push(val); }
  const effIdx=eff!=null?chartData.findIndex(d=>d.log2===eff):-1;
  const effX=effIdx>=0?padL+(effIdx+1)*slot:null;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{width:"100%",height:"auto"}}>
      {yticks.map((v,i)=>{ const y=yFor(v); return (
        <g key={i}>
          <line x1={padL} y1={y} x2={W-padR} y2={y} stroke={C.line} strokeWidth="1" />
          <text x={padL-6} y={y+3} textAnchor="end" fontSize="10" fill={C.sub} style={mono}>{v}</text>
        </g>); })}
      {chartData.map((d,i)=>{
        const x=padL+i*slot+(slot-bw)/2; const y=yFor(d.count); const h=padT+innerH-y;
        const fill=eff!=null?(d.log2>eff?C.nwt:C.wt):C.teal;
        return (
          <g key={i}>
            {d.count>0&&<rect x={x} y={y} width={bw} height={h} rx="3" fill={fill}><title>{`MIC ${d.label} = ${d.count}`}</title></rect>}
            {d.count>0&&<text x={x+bw/2} y={y-4} textAnchor="middle" fontSize="9" fill={C.sub} style={mono}>{d.count}</text>}
            <text x={padL+i*slot+slot/2} y={H-padB+16} textAnchor="end" fontSize="10" fill={C.sub} style={mono}
              transform={`rotate(-40 ${padL+i*slot+slot/2} ${H-padB+16})`}>{d.label}</text>
          </g>);
      })}
      {effX!=null&&(
        <g>
          <line x1={effX} y1={padT-8} x2={effX} y2={padT+innerH} stroke={C.rose} strokeWidth="2" strokeDasharray="5 3" />
          <text x={effX+4} y={padT-12} fontSize="11" fontWeight="600" fill={C.rose} style={display}>{`ECOFF ≤ ${log2Label(eff)}`}</text>
        </g>)}
      <line x1={padL} y1={padT+innerH} x2={W-padR} y2={padT+innerH} stroke={C.line2} strokeWidth="1" />
    </svg>
  );
}

/* ---------- empty state ---------- */
function EmptyState({onUpload,onTemplate,onHelp}){
  return <div className="rounded-lg border p-10 text-center" style={{borderColor:C.line,background:C.surface}}>
    <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-xl" style={{background:C.tealSoft}}>
      <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke={C.teal} strokeWidth="2"><path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/></svg>
    </div>
    <h2 className="text-xl font-bold mb-1" style={display}>เริ่มต้นด้วยการอัปโหลดข้อมูล MIC</h2>
    <p className="text-sm mb-5 mx-auto max-w-lg" style={{color:C.sub}}>
      รองรับไฟล์ CSV และ Excel ทั้งแบบ <b>รายเชื้อ</b> (1 แถวต่อ 1 isolate), แบบ <b>นับรวม</b> (เชื้อ–ยา–MIC–จำนวน) และแบบ <b>ตารางกว้าง</b> (คอลัมน์เป็นค่าเจือจาง) ระบบจะตรวจจับรูปแบบให้อัตโนมัติ
    </p>
    <div className="flex items-center justify-center gap-2">
      <button onClick={onUpload} className="rounded-md px-4 py-2 text-sm font-medium text-white" style={{background:C.teal}}>เลือกไฟล์</button>
      <button onClick={onTemplate} className="rounded-md border px-4 py-2 text-sm font-medium" style={{borderColor:C.line2,color:C.sub}}>ดาวน์โหลดเทมเพลตตัวอย่าง</button>
      {onHelp&&<button onClick={onHelp} className="rounded-md border px-4 py-2 text-sm font-medium" style={{borderColor:C.line2,color:C.teal}}>อ่านคู่มือ & อภิธานศัพท์</button>}
    </div>
  </div>;
}

/* ---------- import modal ---------- */
function ImportModal({pending,setPending,onConfirm}){
  const {headers,mapping,rows}=pending;
  const [m,setM]=useState(mapping);
  const set=(k,v)=>setM(p=>({...p,[k]:v||null}));
  const Sel=({label,value,onChange,opts})=>(
    <label className="block text-sm">
      <span className="text-xs" style={{color:C.sub}}>{label}</span>
      <select value={value||""} onChange={e=>onChange(e.target.value)} className="mt-1 w-full rounded-md border px-2 py-1.5 text-sm" style={{borderColor:C.line2}}>
        <option value="">— ไม่ใช้ —</option>
        {opts.map(h=><option key={h} value={h}>{h}</option>)}
      </select>
    </label>);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{background:"rgba(15,23,42,0.45)"}}>
      <div className="w-full max-w-2xl rounded-lg p-5" style={{background:C.surface}}>
        <h3 className="text-lg font-bold mb-1" style={display}>ยืนยันการจับคู่คอลัมน์</h3>
        <p className="text-sm mb-4" style={{color:C.sub}}>
          ตรวจพบรูปแบบ: <Chip tone="teal">{m.format==="wide"?"ตารางกว้าง":m.format==="aggregated"?"นับรวม":"รายเชื้อ"}</Chip> — ปรับการจับคู่ได้หากไม่ถูกต้อง
        </p>
        <div className="grid grid-cols-2 gap-3 mb-4">
          <Sel label="คอลัมน์เชื้อ (Organism)" value={m.org} onChange={v=>set("org",v)} opts={headers} />
          <Sel label="คอลัมน์ยา (Antibiotic)" value={m.ab} onChange={v=>set("ab",v)} opts={headers} />
          {m.format!=="wide"&&<Sel label="คอลัมน์ค่า MIC" value={m.mic} onChange={v=>set("mic",v)} opts={headers} />}
          {m.format==="aggregated"&&<Sel label="คอลัมน์จำนวน (Count)" value={m.count} onChange={v=>set("count",v)} opts={headers} />}
        </div>
        {m.format==="wide"&&<p className="text-xs mb-4 rounded-md p-2" style={{background:C.bg,color:C.sub}}>คอลัมน์ค่าเจือจางที่ตรวจพบ: {m.dilCols.join(", ")}</p>}
        <p className="text-xs mb-4" style={{color:C.sub}}>ทั้งหมด {rows.length} แถว</p>
        <div className="flex justify-end gap-2">
          <button onClick={()=>setPending(null)} className="rounded-md border px-3 py-2 text-sm" style={{borderColor:C.line2,color:C.sub}}>ยกเลิก</button>
          <button onClick={()=>{pending.mapping=m;onConfirm();}} className="rounded-md px-3 py-2 text-sm font-medium text-white" style={{background:C.teal}}>นำเข้าข้อมูล</button>
        </div>
      </div>
    </div>);
}

/* ---------- combo dashboard ---------- */
function ComboDashboard({a,onOverride}){
  const eff=a.manualEcoff??a.ecoff?.e99??null;
  const chartData=useMemo(()=>{
    const entries=[...a.dist.entries()].map(([l,c])=>({l:Number(l),c})).sort((x,y)=>x.l-y.l);
    if(entries.length){
      const filled=[]; const min=entries[0].l,max=entries[entries.length-1].l;
      for(let k=min;k<=max;k++){ const f=entries.find(e=>e.l===k); filled.push({l:k,c:f?f.c:0}); }
      return filled.map(e=>({label:log2Label(e.l),log2:e.l,count:e.c}));
    }
    return [];
  },[a,eff]);
  const nwtCount=useMemo(()=>eff==null?0:[...a.dist.entries()].filter(([l])=>Number(l)>eff).reduce((s,[,c])=>s+c,0),[a,eff]);
  const N=a.pct?.N??0; const wtPct=N?((N-nwtCount)/N*100):0;
  const ecoffOpts=[]; for(let k=(a.pct?.min??-9);k<=(a.pct?.max??8)+1;k++) ecoffOpts.push(k);
  return (
    <div className="space-y-4">
      <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div><h2 className="text-xl font-bold italic" style={display}>{a.organism}</h2><p className="text-sm" style={{color:C.sub}}>{a.antibiotic}</p></div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-5">
            <Stat label="n" value={N} />
            <Stat label="MIC50" value={a.pct?log2Label(a.pct.mic50):"—"} />
            <Stat label="MIC90" value={a.pct?log2Label(a.pct.mic90):"—"} />
            <Stat label="ช่วง" value={a.pct?`${log2Label(a.pct.min)}–${log2Label(a.pct.max)}`:"—"} />
          </div>
        </div>
      </div>
      <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-sm font-semibold" style={display}>การกระจายค่า MIC (mg/L)</h3>
          <div className="flex items-center gap-3 text-xs" style={{color:C.sub}}>
            <span className="flex items-center gap-1"><span className="h-2.5 w-2.5 rounded-sm" style={{background:C.wt}}/> Wild-type</span>
            <span className="flex items-center gap-1"><span className="h-2.5 w-2.5 rounded-sm" style={{background:C.nwt}}/> Non-wild-type</span>
          </div>
        </div>
        <MicChart chartData={chartData} eff={eff} />
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
          <h3 className="text-sm font-semibold mb-3" style={display}>ค่า ECOFF & การแปลผล</h3>
          {a.ecoff?(<>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <Stat label="ECOFF (95%)" value={log2Label(a.ecoff.e95)} sub="mg/L" />
              <Stat label="ECOFF (99%)" value={log2Label(a.ecoff.e99)} sub="แนะนำ · mg/L" />
              <Stat label="Mean (log2)" value={a.ecoff.mean.toFixed(2)} />
              <Stat label="Goodness R²" value={a.ecoff.r2.toFixed(3)} />
            </div>
            <div className="rounded-md p-3 mb-3" style={{background:C.bg}}>
              <div className="flex items-center justify-between text-sm mb-1">
                <span style={{color:C.sub}}>Wild-type (≤ ECOFF)</span>
                <span className="font-semibold" style={{...mono,color:C.teal}}>{wtPct.toFixed(1)}%</span>
              </div>
              <div className="flex items-center justify-between text-sm">
                <span style={{color:C.sub}}>Non-wild-type (&gt; ECOFF)</span>
                <span className="font-semibold" style={{...mono,color:C.amber}}>{(100-wtPct).toFixed(1)}% ({nwtCount}/{N})</span>
              </div>
              <div className="mt-2 h-2 w-full rounded-full overflow-hidden" style={{background:C.amberSoft}}>
                <div className="h-full" style={{width:`${wtPct}%`,background:C.teal}} />
              </div>
            </div>
            <label className="block text-sm">
              <span className="text-xs" style={{color:C.sub}}>กำหนด ECOFF เอง (override) — เช่น ใช้ค่าทางการของ EUCAST</span>
              <select value={a.manualEcoff??""} onChange={e=>onOverride(a.organism,a.antibiotic,e.target.value===""?null:Number(e.target.value))}
                className="mt-1 w-full rounded-md border px-2 py-1.5 text-sm" style={{...mono,borderColor:C.line2}}>
                <option value="">ใช้ค่าที่คำนวณ (99%) = {log2Label(a.ecoff.e99)}</option>
                {ecoffOpts.map(k=><option key={k} value={k}>{log2Label(k)} mg/L</option>)}
              </select>
            </label>
          </>):(<p className="text-sm" style={{color:C.sub}}>ข้อมูลไม่พอสำหรับการ fit — ต้องการการกระจายที่มีจำนวนเพียงพอและมีพีคชัดเจน</p>)}
        </div>
        <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
          <h3 className="text-sm font-semibold mb-3" style={display}>การตรวจสอบคุณภาพ (QC) & ข้อสังเกต</h3>
          <ul className="space-y-2">
            {a.flags.map((f,i)=>{
              const map={error:{c:C.rose,t:"ข้อผิดพลาด",bg:C.roseSoft},warn:{c:C.amber,t:"ควรตรวจสอบ",bg:C.amberSoft},info:{c:C.teal,t:"ข้อมูล",bg:C.tealSoft},ok:{c:C.emerald,t:"ผ่าน",bg:"#D1FAE5"}}[f.level];
              return (
                <li key={i} className="flex gap-2 rounded-md p-2 text-sm" style={{background:map.bg}}>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={map.c} strokeWidth="2" className="mt-0.5 shrink-0">
                    {f.level==="ok"?<path d="M20 6 9 17l-5-5"/>:<><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></>}
                  </svg>
                  <span style={{color:C.ink}}><b style={{color:map.c}}>{map.t}:</b> {f.msg}</span>
                </li>);
            })}
          </ul>
        </div>
      </div>
      <p className="text-xs leading-relaxed rounded-md p-3" style={{background:C.bg,color:C.sub}}>
        <b>หมายเหตุวิธีคำนวณ:</b> ECOFF ประเมินด้วยการ fit การกระจายแบบ log-normal กับประชากร wild-type (วิธี ECOFFinder, Turnidge et al. 2006) ค่าที่ได้เป็นการประมาณเชิงสถิติ ควรตรวจสอบด้วยตาและเทียบกับค่าทางการของ EUCAST ก่อนใช้งานจริง การกำหนด ECOFF ที่เชื่อถือได้ควรใช้ข้อมูลรวม ≥100 isolates จากหลายแหล่ง
      </p>
    </div>);
}

/* ---------- raw data ---------- */
function RawData({analyses,onEdit}){
  const [q,setQ]=useState("");
  const filtered=analyses.filter(a=>(a.organism+a.antibiotic).toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
      <div className="flex items-center justify-between mb-3 gap-3">
        <h3 className="text-sm font-semibold" style={display}>ข้อมูลดิบ — แก้ไขจำนวนได้ทันที (อัปเดตการวิเคราะห์แบบ real-time)</h3>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="ค้นหาเชื้อ/ยา…" className="rounded-md border px-2 py-1.5 text-sm w-48" style={{borderColor:C.line2}} />
      </div>
      <div className="overflow-auto mic-scroll" style={{maxHeight:520}}>
        {filtered.map((a,idx)=>{
          const entries=[...a.dist.entries()].map(([l,c])=>({l:Number(l),c})).sort((x,y)=>x.l-y.l);
          return (
            <div key={idx} className="mb-4">
              <div className="flex items-center gap-2 mb-1 sticky top-0 py-1" style={{background:C.surface}}>
                <span className="text-sm font-semibold italic" style={display}>{a.organism}</span>
                <span className="text-xs" style={{color:C.sub}}>· {a.antibiotic}</span>
              </div>
              <table className="w-full text-sm" style={mono}>
                <thead><tr style={{color:C.sub}}>
                  <th className="text-left font-medium py-1 pr-3">MIC (mg/L)</th>
                  {entries.map(e=><th key={e.l} className="px-1 text-center font-medium">{log2Label(e.l)}</th>)}
                </tr></thead>
                <tbody><tr>
                  <td className="py-1 pr-3" style={{color:C.sub}}>จำนวน</td>
                  {entries.map(e=>(
                    <td key={e.l} className="px-1">
                      <input type="number" min="0" value={e.c} onChange={ev=>onEdit(a.organism,a.antibiotic,e.l,ev.target.value)}
                        className="w-14 rounded border px-1 py-0.5 text-center" style={{borderColor:C.line2}} />
                    </td>))}
                </tr></tbody>
              </table>
            </div>);
        })}
        {filtered.length===0&&<p className="text-sm" style={{color:C.sub}}>ไม่พบรายการที่ตรงกับการค้นหา</p>}
      </div>
      <p className="text-xs mt-3" style={{color:C.sub}}>เปลี่ยนค่าจำนวนในตารางแล้วกราฟ ECOFF และการแปลผลจะคำนวณใหม่ทันที พร้อมบันทึกอัตโนมัติ</p>
    </div>);
}

/* ---------- grouped overview (by drug or by organism) ---------- */
function GroupedView({rows, by, pct, onPick}){
  const groups=useMemo(()=>{
    const map=new Map();
    for(const r of rows){
      const gname=by==="ab"?r.a.antibiotic:r.a.organism;
      const mname=by==="ab"?r.a.organism:r.a.antibiotic;
      if(!map.has(gname)) map.set(gname,{name:gname,items:[]});
      map.get(gname).items.push({...r, mname, hasE:r.eff!=null});
    }
    const arr=[...map.values()];
    arr.forEach(g=>{ g.total=g.items.length; g.withE=g.items.filter(i=>i.hasE).length; g.iso=g.items.reduce((s,i)=>s+i.n,0); });
    arr.sort((a,b)=>a.name.localeCompare(b.name));
    return arr;
  },[rows,by]);

  if(!groups.length) return <div className="rounded-lg border p-6 text-center text-sm" style={{borderColor:C.line,background:C.surface,color:C.sub}}>ไม่มีข้อมูลตามตัวกรอง</div>;

  return (
    <div className="space-y-3">
      <p className="text-xs" style={{color:C.sub}}>
        {by==="ab"?"แต่ละยา แสดงเชื้อทั้งหมดที่ทดสอบ พร้อมสัดส่วนที่คำนวณ ECOFF ได้":"แต่ละเชื้อ แสดงยาทั้งหมดที่ทดสอบ พร้อมสัดส่วนที่คำนวณ ECOFF ได้"} · {groups.length} กลุ่ม
      </p>
      {groups.map(g=>{
        const pctWith=g.total?Math.round(g.withE/g.total*100):0;
        return (
          <div key={g.name} className="rounded-lg border" style={{borderColor:C.line,background:C.surface}}>
            <div className="p-3 border-b" style={{borderColor:C.line}}>
              <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
                <div className="flex items-baseline gap-2">
                  <span className={by==="ab"?"text-base font-bold":"text-base font-bold italic"} style={display}>{g.name}</span>
                  <span className="text-xs" style={{color:C.sub}}>{by==="ab"?`${g.total} เชื้อ`:`${g.total} ยา`} · n รวม {g.iso.toLocaleString()}</span>
                </div>
                <span className="text-sm font-semibold" style={{...mono}}>
                  <span style={{color:C.teal}}>มี ECOFF {g.withE}</span>
                  <span style={{color:C.sub}}> / ไม่มี {g.total-g.withE}</span>
                  <span style={{color:C.ink}}> ({pctWith}%)</span>
                </span>
              </div>
              <div className="flex h-2.5 w-full rounded-full overflow-hidden" style={{background:"#EEF2F6"}}>
                <div style={{width:`${pctWith}%`,background:C.teal}} title={`มี ECOFF ${g.withE}/${g.total}`} />
                <div style={{width:`${100-pctWith}%`,background:C.line2}} title={`ไม่มี ECOFF ${g.total-g.withE}/${g.total}`} />
              </div>
            </div>
            <div className="divide-y" style={{borderColor:C.line}}>
              {g.items.map(it=>(
                <div key={it.key} onClick={()=>onPick(it.key)} className="flex items-center gap-3 px-3 py-1.5 cursor-pointer text-sm" style={{borderTop:`1px solid ${C.line}`}}>
                  <span className="h-1.5 w-1.5 rounded-full shrink-0" style={{background:it.hasE?C.teal:C.line2}} />
                  <span className={by==="ab"?"flex-1 truncate italic":"flex-1 truncate"} style={by==="ab"?display:{fontFamily:"Inter"}}>{it.mname}</span>
                  <span className="text-xs shrink-0" style={{...mono,color:C.sub}}>n={it.n}</span>
                  <span className="shrink-0 text-right" style={{...mono,width:92}}>
                    {it.hasE?<><span style={{color:C.sub}}>ECOFF </span><b>{log2Label(it.eff)}</b></>:<span style={{color:C.amber}}>no ECOFF</span>}
                  </span>
                  <span className="shrink-0 text-right" style={{...mono,width:56,color:it.nwt==null?C.sub:it.nwt>20?C.amber:C.ink}}>{it.nwt!=null?it.nwt.toFixed(1)+"%":"—"}</span>
                </div>))}
            </div>
          </div>);
      })}
    </div>
  );
}

/* ---------- overview (summary) page ---------- */
function OverviewPage({analyses, anyFilter, onPick}){
  const [pct,setPct]=useState("99");
  const [sortBy,setSortBy]=useState("nwt");
  const [dir,setDir]=useState("desc");
  const [mode,setMode]=useState("table"); // table | chart
  const [groupBy,setGroupBy]=useState("none"); // none | ab | org
  const [cols,setCols]=useState({mic50:true,mic90:true,range:true,r2:true,qc:true});
  const pctKey={"95":"e95","97.5":"e975","99":"e99","99.5":"e995"}[pct];
  const effFor=a=>a.manualEcoff??(a.ecoff?a.ecoff[pctKey]:null);
  const nwtOf=a=>{ const eff=effFor(a); if(eff==null) return null; return [...a.dist.entries()].filter(([l])=>Number(l)>eff).reduce((s,[,c])=>s+c,0); };
  const nwtPct=a=>{ const n=a.pct?.N??0; const x=nwtOf(a); return (n&&x!=null)?(x/n*100):null; };

  const rows=useMemo(()=>{
    const arr=analyses.map(a=>({ a, key:`${a.organism}||${a.antibiotic}`, n:a.pct?.N??0,
      eff:effFor(a), nwt:nwtPct(a), r2:a.ecoff?a.ecoff.r2:null,
      worst:a.flags.some(f=>f.level==="error")?2:a.flags.some(f=>f.level==="warn")?1:0 }));
    const cmp={ org:(x,y)=>x.a.organism.localeCompare(y.a.organism)||x.a.antibiotic.localeCompare(y.a.antibiotic),
      ab:(x,y)=>x.a.antibiotic.localeCompare(y.a.antibiotic), n:(x,y)=>x.n-y.n,
      nwt:(x,y)=>(x.nwt??-1)-(y.nwt??-1), ecoff:(x,y)=>(x.eff??-99)-(y.eff??-99),
      r2:(x,y)=>(x.r2??-1)-(y.r2??-1), qc:(x,y)=>x.worst-y.worst }[sortBy];
    arr.sort((x,y)=>{ const v=cmp(x,y); return dir==="asc"?v:-v; });
    return arr;
  },[analyses,pct,sortBy,dir]);

  // KPIs
  const totalCombos=analyses.length;
  const totalIso=analyses.reduce((s,a)=>s+(a.pct?.N??0),0);
  const withE=analyses.filter(a=>(a.manualEcoff!=null)||(a.ecoff!=null)).length;
  const flagged=analyses.filter(a=>a.flags.some(f=>f.level==="warn"||f.level==="error")).length;
  const avgNwt=(()=>{ const v=rows.map(r=>r.nwt).filter(x=>x!=null); return v.length?(v.reduce((s,x)=>s+x,0)/v.length):null; })();

  const Th=({id,label,right})=>(
    <th className={`py-2 px-2 font-medium cursor-pointer select-none ${right?"text-right":"text-left"}`} style={{color:C.sub}} onClick={()=>{ if(sortBy===id) setDir(d=>d==="asc"?"desc":"asc"); else { setSortBy(id); setDir("desc"); } }}>
      {label}{sortBy===id?(dir==="asc"?" ▲":" ▼"):""}
    </th>);

  const kpi=(label,value,sub,tone)=>(
    <div className="rounded-lg border p-3" style={{borderColor:C.line,background:C.surface}}>
      <div className="text-xs uppercase mb-1" style={{color:C.sub,letterSpacing:"0.06em"}}>{label}</div>
      <div className="text-2xl font-bold" style={{...mono,color:tone||C.ink}}>{value}</div>
      {sub&&<div className="text-xs" style={{color:C.sub}}>{sub}</div>}
    </div>);

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 lg:grid-cols-5 gap-3">
        {kpi("คู่เชื้อ–ยา", totalCombos, anyFilter?"ตามตัวกรอง":"ทั้งหมด")}
        {kpi("จำนวนเชื้อรวม", totalIso.toLocaleString())}
        {kpi("มี ECOFF", `${withE}/${totalCombos}`, `${totalCombos?Math.round(withE/totalCombos*100):0}%`, C.teal)}
        {kpi("ต้องตรวจสอบ (QC)", flagged, "มี warning/error", flagged?C.amber:C.emerald)}
        {kpi("ค่าเฉลี่ย %NWT", avgNwt!=null?avgNwt.toFixed(1)+"%":"—", `ที่ระดับ ${pct}%`, C.amber)}
      </div>

      {/* display controls */}
      <div className="rounded-lg border p-3 flex flex-wrap items-center gap-x-5 gap-y-2" style={{borderColor:C.line,background:C.surface}}>
        <label className="flex items-center gap-2 text-sm"><span style={{color:C.sub}}>จัดกลุ่มตาม</span>
          <select value={groupBy} onChange={e=>setGroupBy(e.target.value)} className="rounded-md border px-2 py-1 text-sm" style={{borderColor:C.line2}}>
            <option value="none">ไม่จัดกลุ่ม (รายคู่)</option>
            <option value="ab">ตามยา → เชื้อทั้งหมด</option>
            <option value="org">ตามเชื้อ → ยาทั้งหมด</option>
          </select></label>
        <label className="flex items-center gap-2 text-sm"><span style={{color:C.sub}}>ระดับครอบคลุม WT</span>
          <select value={pct} onChange={e=>setPct(e.target.value)} className="rounded-md border px-2 py-1 text-sm" style={{...mono,borderColor:C.line2}}>
            {["95","97.5","99","99.5"].map(p=><option key={p} value={p}>{p}%</option>)}
          </select></label>
        <label className="flex items-center gap-2 text-sm"><span style={{color:C.sub}}>เรียงตาม</span>
          <select value={sortBy} onChange={e=>setSortBy(e.target.value)} className="rounded-md border px-2 py-1 text-sm" style={{borderColor:C.line2}}>
            {[["nwt","%NWT"],["org","เชื้อ"],["ab","ยา"],["n","จำนวน (n)"],["ecoff","ECOFF"],["r2","R²"],["qc","สถานะ QC"]].map(([v,t])=><option key={v} value={v}>{t}</option>)}
          </select>
          <button onClick={()=>setDir(d=>d==="asc"?"desc":"asc")} className="rounded-md border px-2 py-1 text-sm" style={{borderColor:C.line2,color:C.sub}}>{dir==="asc"?"▲ น้อย→มาก":"▼ มาก→น้อย"}</button>
        </label>
        {groupBy==="none"&&(
        <div className="inline-flex rounded-md border" style={{borderColor:C.line2}}>
          {[["table","ตาราง"],["chart","กราฟ %NWT"]].map(([k,t])=>(
            <button key={k} onClick={()=>setMode(k)} className="px-2.5 py-1 text-sm font-medium" style={{background:mode===k?C.teal:C.surface,color:mode===k?"#fff":C.sub}}>{t}</button>))}
        </div>)}
        {groupBy==="none"&&mode==="table"&&(
          <div className="flex items-center gap-1.5 text-xs">
            <span style={{color:C.sub}}>คอลัมน์:</span>
            {[["mic50","MIC50"],["mic90","MIC90"],["range","ช่วง"],["r2","R²"],["qc","QC"]].map(([k,t])=>(
              <button key={k} onClick={()=>setCols(c=>({...c,[k]:!c[k]}))} className="rounded-full px-2 py-0.5 border"
                style={{borderColor:cols[k]?C.teal:C.line2,background:cols[k]?C.tealSoft:"transparent",color:cols[k]?C.tealDark:C.sub}}>{t}</button>))}
          </div>)}
      </div>

      {groupBy!=="none"&&(
        <GroupedView rows={rows} by={groupBy} pct={pct} onPick={onPick} />
      )}
      {groupBy==="none"&&(mode==="table"?(
        <div className="rounded-lg border overflow-auto mic-scroll" style={{borderColor:C.line,background:C.surface,maxHeight:560}}>
          <table className="w-full text-sm" style={{borderCollapse:"collapse"}}>
            <thead className="sticky top-0" style={{background:C.surface,boxShadow:`inset 0 -1px 0 ${C.line}`}}>
              <tr><Th id="org" label="เชื้อ" /><Th id="ab" label="ยา" /><Th id="n" label="n" right />
                {cols.mic50&&<Th id="m50" label="MIC50" right />}{cols.mic90&&<Th id="m90" label="MIC90" right />}
                {cols.range&&<th className="py-2 px-2 text-right font-medium" style={{color:C.sub}}>ช่วง</th>}
                <Th id="ecoff" label={`ECOFF (${pct}%)`} right /><Th id="nwt" label="%NWT" right />
                {cols.r2&&<Th id="r2" label="R²" right />}{cols.qc&&<th className="py-2 px-2 text-left font-medium" style={{color:C.sub}}>QC</th>}
              </tr>
            </thead>
            <tbody style={mono}>
              {rows.map(r=>{
                const a=r.a; const wt=r.nwt!=null?100-r.nwt:null;
                const flagTone=r.worst===2?C.rose:r.worst===1?C.amber:C.emerald;
                return (
                  <tr key={r.key} className="cursor-pointer" style={{borderTop:`1px solid ${C.line}`}} onClick={()=>onPick(r.key)}>
                    <td className="py-1.5 px-2 italic" style={{...display,color:C.ink}}>{a.organism}</td>
                    <td className="py-1.5 px-2" style={{fontFamily:"Inter"}}>{a.antibiotic}</td>
                    <td className="py-1.5 px-2 text-right">{r.n}</td>
                    {cols.mic50&&<td className="py-1.5 px-2 text-right">{a.pct?log2Label(a.pct.mic50):"—"}</td>}
                    {cols.mic90&&<td className="py-1.5 px-2 text-right">{a.pct?log2Label(a.pct.mic90):"—"}</td>}
                    {cols.range&&<td className="py-1.5 px-2 text-right">{a.pct?`${log2Label(a.pct.min)}–${log2Label(a.pct.max)}`:"—"}</td>}
                    <td className="py-1.5 px-2 text-right">{r.eff!=null?log2Label(r.eff):<span style={{color:C.amber}}>—</span>}{a.manualEcoff!=null&&<span className="ml-1 text-xs" style={{color:C.teal}}>✎</span>}</td>
                    <td className="py-1.5 px-2 text-right">
                      {r.nwt!=null?(
                        <span className="inline-flex items-center gap-1.5 justify-end">
                          <span style={{width:42,height:6,background:C.amberSoft,borderRadius:4,overflow:"hidden",display:"inline-block"}}>
                            <span style={{display:"block",height:"100%",width:`${Math.min(100,r.nwt)}%`,background:r.nwt>20?C.amber:C.teal}} />
                          </span>
                          <b style={{color:r.nwt>20?C.amber:C.ink}}>{r.nwt.toFixed(1)}%</b>
                        </span>):"—"}
                    </td>
                    {cols.r2&&<td className="py-1.5 px-2 text-right">{r.r2!=null?r.r2.toFixed(3):"—"}</td>}
                    {cols.qc&&<td className="py-1.5 px-2"><span className="inline-block h-2 w-2 rounded-full" style={{background:flagTone}} title={r.worst===2?"error":r.worst===1?"warning":"ok"} /></td>}
                  </tr>);
              })}
              {rows.length===0&&<tr><td colSpan="10" className="py-6 text-center" style={{color:C.sub}}>ไม่มีข้อมูลตามตัวกรอง</td></tr>}
            </tbody>
          </table>
        </div>
      ):(
        <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
          <div className="text-sm font-semibold mb-3" style={display}>สัดส่วน Non-wild-type ต่อคู่เชื้อ–ยา (ระดับ {pct}%)</div>
          <div className="space-y-1.5 overflow-auto mic-scroll" style={{maxHeight:520}}>
            {rows.map(r=>(
              <div key={r.key} className="flex items-center gap-3 cursor-pointer" onClick={()=>onPick(r.key)}>
                <div className="text-xs truncate" style={{width:230,color:C.sub}}><span className="italic" style={display}>{r.a.organism}</span> · {r.a.antibiotic}</div>
                <div className="flex-1 h-5 rounded" style={{background:C.bg,position:"relative",overflow:"hidden"}}>
                  {r.nwt!=null&&<div style={{height:"100%",width:`${Math.max(1,Math.min(100,r.nwt))}%`,background:r.nwt>20?C.amber:C.teal,borderRadius:4}} />}
                </div>
                <div className="text-xs text-right" style={{...mono,width:64,color:r.nwt==null?C.sub:r.nwt>20?C.amber:C.ink}}>{r.nwt!=null?r.nwt.toFixed(1)+"%":"no ECOFF"}</div>
              </div>))}
            {rows.length===0&&<p className="text-sm" style={{color:C.sub}}>ไม่มีข้อมูลตามตัวกรอง</p>}
          </div>
        </div>
      ))}
      <p className="text-xs" style={{color:C.sub}}>คลิกที่แถว/แท่งเพื่อเปิดดูรายละเอียดของคู่เชื้อ–ยานั้น · ✎ = ใช้ค่า ECOFF ที่กำหนดเอง · ตาราง/กราฟ/กลุ่มนี้แสดงเฉพาะรายการที่ผ่านตัวกรองด้านซ้าย</p>
    </div>
  );
}

/* ---------- help / glossary page ---------- */
function HelpPage({onBack}){
  const Term=({th,en,children})=>(
    <div className="rounded-lg border p-3" style={{borderColor:C.line,background:C.surface}}>
      <div className="flex items-baseline gap-2 mb-1">
        <span className="font-semibold" style={{...display,color:C.tealDark}}>{th}</span>
        {en&&<span className="text-xs" style={{color:C.sub}}>{en}</span>}
      </div>
      <p className="text-sm leading-relaxed" style={{color:C.ink}}>{children}</p>
    </div>);
  const Step=({n,title,children})=>(
    <div className="flex gap-3">
      <div className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full text-xs font-bold text-white" style={{background:C.teal,...mono}}>{n}</div>
      <div><div className="text-sm font-semibold" style={display}>{title}</div><p className="text-sm" style={{color:C.sub}}>{children}</p></div>
    </div>);
  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <h2 className="text-xl font-bold" style={display}>คู่มือการใช้งาน & อภิธานศัพท์</h2>
        <button onClick={onBack} className="rounded-md border px-3 py-1.5 text-sm" style={{borderColor:C.line2,color:C.teal}}>← กลับ</button>
      </div>

      <div className="rounded-lg border p-4" style={{borderColor:C.line,background:C.surface}}>
        <h3 className="text-sm font-semibold uppercase mb-3" style={{color:C.sub,letterSpacing:"0.06em"}}>ขั้นตอนการใช้งาน</h3>
        <div className="space-y-3">
          <Step n="1" title="เตรียมและอัปโหลดไฟล์">รองรับ CSV/Excel 3 รูปแบบ — รายเชื้อ (1 แถว/isolate มีคอลัมน์ MIC), นับรวม (เชื้อ–ยา–MIC–จำนวน) และตารางกว้าง (คอลัมน์หัวเป็นค่าเจือจาง เซลล์เป็นจำนวน) ค่าตัวเลขรองรับเครื่องหมาย ≤ และ &gt; เช่น "≤0.25", "&gt;32" ระบบตรวจจับรูปแบบให้อัตโนมัติ</Step>
          <Step n="2" title="ยืนยันการจับคู่คอลัมน์">หน้าต่างจะให้ยืนยันว่าให้คอลัมน์ไหนเป็นเชื้อ/ยา/MIC/จำนวน ปรับได้หากตรวจจับผิด</Step>
          <Step n="3" title="ดูภาพรวม">หน้า "ภาพรวม" สรุปทุกคู่เชื้อ–ยาในตารางเดียว ปรับระดับครอบคลุม WT, เรียงลำดับ, เลือกคอลัมน์ และสลับเป็นกราฟ %NWT ได้</Step>
          <Step n="4" title="เจาะรายคู่">คลิกที่แถวหรือเลือกจากแถบซ้าย เพื่อดูกราฟการกระจาย MIC, ค่า ECOFF, การแปลผล WT/NWT และผลตรวจสอบคุณภาพ (QC)</Step>
          <Step n="5" title="ปรับค่า ECOFF เอง">ในหน้ารายคู่ สามารถกำหนด ECOFF เอง (เช่น ใช้ค่าทางการของ EUCAST) แทนค่าที่เครื่องคำนวณ ระบบจะแปลผลใหม่ทันที</Step>
          <Step n="6" title="แก้ไขข้อมูลดิบ">หน้า "ข้อมูลดิบ" แก้จำนวนในแต่ละช่องเจือจางได้ กราฟและการวิเคราะห์อัปเดตทันที พร้อมบันทึกอัตโนมัติในเบราว์เซอร์</Step>
          <Step n="7" title="กรองและส่งออก">ใช้ตัวกรอง (เชื้อ/ยา/สถานะ ECOFF) ที่แถบซ้าย แล้วส่งออกเป็น Excel — ไฟล์จะยึดตามรายการที่กรองไว้</Step>
        </div>
      </div>

      <div>
        <h3 className="text-sm font-semibold uppercase mb-3" style={{color:C.sub,letterSpacing:"0.06em"}}>อภิธานศัพท์ทางวิชาการ</h3>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <Term th="MIC" en="Minimum Inhibitory Concentration">ความเข้มข้นต่ำสุดของยาต้านจุลชีพที่ยับยั้งการเจริญที่มองเห็นได้ของเชื้อ วัดเป็น mg/L (หรือ µg/mL) โดยทั่วไปทดสอบเป็นชุดเจือจางสองเท่า (doubling/two-fold dilution) เช่น 0.25, 0.5, 1, 2, 4 …</Term>
          <Term th="การกระจาย MIC" en="MIC distribution">จำนวนเชื้อในแต่ละค่าความเข้มข้น เมื่อพล็อตจะเห็นเป็นฮิสโตแกรม ประชากรที่ไม่มีกลไกดื้อยา (wild-type) มักรวมกันเป็นพีคเดียวรูประฆังบนสเกล log2 ส่วนเชื้อที่มีกลไกดื้อจะแยกออกไปทางค่าสูง</Term>
          <Term th="ECOFF" en="Epidemiological Cut-Off">ค่าจุดตัดทางระบาดวิทยา ใช้แบ่งประชากรเชื้อออกเป็น wild-type (ไม่มีกลไกดื้อที่ได้มา) กับ non-wild-type โดยอิงจากการกระจาย MIC ของประชากร ไม่ได้อิงผลการรักษาทางคลินิก เป็นเครื่องมือหลักในการเฝ้าระวังการดื้อยาและตรวจจับกลไกดื้อที่เกิดใหม่</Term>
          <Term th="Wild-type (WT)">เชื้อที่ไม่มีกลไกการดื้อยาที่ได้มา (acquired/mutational) ต่อยานั้น มีค่า MIC ≤ ECOFF</Term>
          <Term th="Non-wild-type (NWT)">เชื้อที่มี MIC สูงกว่า ECOFF บ่งชี้ว่าน่าจะมีกลไกดื้อยา (เช่น เอนไซม์ทำลายยา, ปั๊มขับยา, การกลายพันธุ์ของเป้า) — ไม่จำเป็นต้องเท่ากับ "ดื้อทางคลินิก" เสมอไป</Term>
          <Term th="Breakpoint ทางคลินิก" en="Clinical breakpoint">ต่างจาก ECOFF — breakpoint (S/I/R) กำหนดโดยพิจารณาผลการรักษา เภสัชจลนศาสตร์/เภสัชพลศาสตร์ และขนาดยา ใช้ทำนายผลการรักษาในผู้ป่วย ขณะที่ ECOFF บอกเพียงว่าเชื้อมีกลไกดื้อหรือไม่</Term>
          <Term th="MIC50 / MIC90">ค่า MIC ที่ครอบคลุม 50% และ 90% ของเชื้อทั้งหมดในชุดข้อมูล (เปอร์เซ็นไทล์ที่ 50 และ 90) ใช้สรุปแนวโน้มความไวของประชากรเชื้อโดยรวม</Term>
          <Term th="ชุดเจือจางสองเท่า" en="Doubling / two-fold dilution">มาตรฐานการทดสอบ MIC ที่แต่ละขั้นต่างกัน 2 เท่า ทำให้สเกลเป็น log ฐาน 2 — แอปนี้คำนวณบนสเกล log2 เพื่อให้การ fit ทางสถิติถูกต้อง</Term>
          <Term th="วิธี ECOFFinder" en="Turnidge et al., 2006">วิธีประเมิน ECOFF เชิงสถิติ โดย fit เส้นโค้งสะสมแบบ log-normal เข้ากับส่วน wild-type ของการกระจาย แล้วหาความเข้มข้นที่ครอบคลุมประชากร WT ตามสัดส่วนที่กำหนด (เช่น 95/97.5/99/99.5%) อ้างอิง: Turnidge J, Kahlmeter G, Kronvall G. Clin Microbiol Infect 2006;12:418–425</Term>
          <Term th="ระดับครอบคลุม WT" en="% coverage">สัดส่วนของประชากร wild-type (ตามโมเดล) ที่ ECOFF ครอบคลุม ค่าที่สูงขึ้น (99.5%) ให้ ECOFF ที่สูงขึ้น = อนุรักษ์นิยมกว่า (ตัดเชื้อปกติเป็น NWT น้อยลง) EUCAST มักใช้ช่วง 97.5–99%</Term>
          <Term th="R² (ความพอดีของโมเดล)" en="Goodness of fit">บอกว่าเส้นโค้ง log-normal อธิบายข้อมูลได้ดีเพียงใด (เข้าใกล้ 1 = ดี) ถ้าต่ำ อาจแปลว่าการกระจายมีหลายพีค (multimodal) หรือไม่เป็น log-normal ควรตรวจ ECOFF ด้วยตา</Term>
          <Term th="ค่า off-scale / censored" en="≤ และ > values">ค่าที่อยู่นอกช่วงที่ทดสอบ เช่น "≤0.25" (ต่ำกว่าหรือเท่าขอบล่าง) หรือ "&gt;32" (สูงกว่าขอบบน) ถ้ามีสัดส่วนสูงอาจทำให้ประเมิน ECOFF คลาดเคลื่อน</Term>
          <Term th="Truncated distribution">การกระจายที่ถูกตัด — พีคของ wild-type อยู่ชิดขอบความเข้มข้นต่ำสุดที่ทดสอบ ทำให้มองไม่เห็นส่วนล่างของประชากร ควรทดสอบเจือจางให้ต่ำลงอีกเพื่อให้ประเมิน WT ได้ครบ</Term>
          <Term th="จำนวนเชื้อ (n) ที่เหมาะสม">การกำหนด ECOFF ที่เชื่อถือได้ควรใช้ข้อมูลรวม ≥100 isolates และควรมาจากหลายห้องปฏิบัติการ/แหล่ง เพื่อให้การกระจายเป็นตัวแทนของประชากร wild-type จริง</Term>
        </div>
      </div>

      <p className="text-xs leading-relaxed rounded-md p-3" style={{background:C.bg,color:C.sub}}>
        <b>ข้อจำกัดสำคัญ:</b> ค่า ECOFF ที่แอปคำนวณเป็นการประมาณเชิงสถิติจากข้อมูลที่ป้อนเข้า ไม่ใช่ค่าทางการ ควรตรวจสอบด้วยตา (เทียบรูปการกระจาย) และเทียบกับค่าทางการของ EUCAST (eucast.org) ก่อนนำไปใช้อ้างอิงหรือทำรายงานเฝ้าระวัง การแปลผลทางคลินิก (S/I/R) ต้องใช้ clinical breakpoint ไม่ใช่ ECOFF
      </p>
    </div>
  );
}

/* ---------- main app ---------- */
function App(){
  const [datasets,setDatasets]=useState([]);
  const [active,setActive]=useState(null);
  const [analyses,setAnalyses]=useState([]);
  const [selectedKey,setSelectedKey]=useState(null);
  const [fOrg,setFOrg]=useState("");      // filter: organism
  const [fAb,setFAb]=useState("");        // filter: antibiotic
  const [fEcoff,setFEcoff]=useState("");  // filter: "" all | "has" | "none"
  const [pendingFile,setPendingFile]=useState(null);
  const [view,setView]=useState("dashboard");
  const [toast,setToast]=useState(null);
  const fileRef=useRef();
  const keyOf=a=>`${a.organism}||${a.antibiotic}`;
  const hasEcoff=a=>(a.manualEcoff!=null)||(a.ecoff!=null);
  const notify=(msg,tone="teal")=>{ setToast({msg,tone}); setTimeout(()=>setToast(null),2600); };

  useEffect(()=>{ listIndex().then(setDatasets); },[]);
  useEffect(()=>{
    if(!active){ setAnalyses([]); return; }
    const a=buildAnalyses(active.records);
    const ov=active.overrides||{};
    a.forEach(x=>{ const k=`${x.organism}||${x.antibiotic}`; if(ov[k]!=null) x.manualEcoff=ov[k]; });
    setAnalyses(a);
    setSelectedKey(prev=>{ const keys=a.map(x=>`${x.organism}||${x.antibiotic}`); return prev&&keys.includes(prev)?prev:(keys[0]??null); });
  },[active]);

  function onFile(e){
    const file=e.target.files?.[0]; if(!file) return;
    const ext=file.name.split(".").pop().toLowerCase();
    if(["csv","tsv","txt"].includes(ext)){
      Papa.parse(file,{header:true,skipEmptyLines:true,
        complete:res=>ingest(res.data,res.meta.fields||Object.keys(res.data[0]||{}),file.name),
        error:()=>notify("อ่านไฟล์ CSV ไม่สำเร็จ ตรวจสอบรูปแบบไฟล์","rose")});
    } else if(["xlsx","xls"].includes(ext)){
      const reader=new FileReader();
      reader.onload=ev=>{ try{ const wb=XLSX.read(ev.target.result,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]]; const json=XLSX.utils.sheet_to_json(ws,{defval:""});
        ingest(json,Object.keys(json[0]||{}),file.name); }catch{ notify("อ่านไฟล์ Excel ไม่สำเร็จ","rose"); } };
      reader.readAsArrayBuffer(file);
    } else notify("รองรับเฉพาะไฟล์ .csv และ .xlsx","rose");
    e.target.value="";
  }
  function ingest(rows,headers,name){
    rows=rows.filter(r=>Object.values(r).some(v=>v!==""&&v!=null));
    if(!rows.length){ notify("ไม่พบข้อมูลในไฟล์","rose"); return; }
    setPendingFile({rows,headers,mapping:detectFormat(headers),name});
  }
  async function confirmImport(){
    const {rows,mapping,name}=pendingFile;
    const records=normalizeRows(rows,mapping);
    if(!records.length){ notify("แปลงข้อมูลไม่ได้ — ตรวจสอบการจับคู่คอลัมน์","rose"); return; }
    const ds={id:"ds_"+Date.now(),name,at:new Date().toISOString(),records,overrides:{}};
    setActive(ds); setPendingFile(null); setSelectedKey(null); setFOrg(""); setFAb(""); setFEcoff(""); setView("overview");
    await saveDataset(ds); setDatasets(await listIndex());
    notify(`นำเข้า ${records.reduce((s,r)=>s+r.count,0)} เชื้อ สำเร็จ`);
  }
  async function openDataset(id){ const ds=await loadDataset(id); if(ds){ setActive(ds); setSelectedKey(null); setFOrg(""); setFAb(""); setFEcoff(""); setView("overview"); } else notify("เปิดชุดข้อมูลไม่สำเร็จ","rose"); }
  async function removeDataset(id){ await deleteDataset(id); setDatasets(await listIndex()); if(active?.id===id){ setActive(null); setAnalyses([]); } notify("ลบชุดข้อมูลแล้ว","slate"); }

  const updateCount=useCallback((organism,antibiotic,log2,newCount)=>{
    setActive(prev=>{
      if(!prev) return prev;
      const val=Math.max(0,parseFloat(newCount)||0);
      const others=prev.records.filter(r=>!(r.organism===organism&&r.antibiotic===antibiotic&&r.log2===log2));
      if(val>0) others.push({organism,antibiotic,log2,count:val,censorLow:false,censorHigh:false});
      const next={...prev,records:others}; saveDataset(next); return next;
    });
  },[]);
  const setOverride=useCallback((organism,antibiotic,ecoffLog2)=>{
    setActive(prev=>{
      if(!prev) return prev;
      const overrides={...(prev.overrides||{})}; const k=`${organism}||${antibiotic}`;
      if(ecoffLog2===null) delete overrides[k]; else overrides[k]=ecoffLog2;
      const next={...prev,overrides}; saveDataset(next); return next;
    });
  },[]);

  function exportAll(){
    const rowsSrc=filteredAnalyses;
    if(!rowsSrc.length) return;
    const wb=XLSX.utils.book_new();
    const summary=rowsSrc.map(a=>{
      const eff=a.manualEcoff??a.ecoff?.e99;
      const nwt=a.pct?[...a.dist.entries()].filter(([l])=>Number(l)>eff).reduce((s,[,c])=>s+c,0):0;
      return { Organism:a.organism, Antibiotic:a.antibiotic, N:a.pct?.N??0,
        MIC50:a.pct?log2Label(a.pct.mic50):"", MIC90:a.pct?log2Label(a.pct.mic90):"",
        Range:a.pct?`${log2Label(a.pct.min)}-${log2Label(a.pct.max)}`:"",
        "ECOFF_computed_99%":a.ecoff?log2Label(a.ecoff.e99):"",
        "ECOFF_used":eff!=null?log2Label(eff):"",
        "ECOFF_source":a.manualEcoff!=null?"manual":"computed",
        "Fit_R2":a.ecoff?a.ecoff.r2.toFixed(3):"",
        "WT_%":a.pct?(((a.pct.N-nwt)/a.pct.N)*100).toFixed(1):"",
        "NWT_%":a.pct?((nwt/a.pct.N)*100).toFixed(1):"",
        QC:a.flags.map(f=>f.code).join("; ") };
    });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(summary),"Summary");
    const raw=[];
    rowsSrc.forEach(a=>[...a.dist.entries()].sort((x,y)=>x[0]-y[0]).forEach(([l,c])=>raw.push({Organism:a.organism,Antibiotic:a.antibiotic,MIC:log2Label(Number(l)),Count:c})));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(raw),"Distributions");
    XLSX.writeFile(wb,`MIC_ECOFF_${active?.name?.replace(/\.[^.]+$/,"")||"export"}.xlsx`);
    notify("ส่งออกไฟล์ Excel แล้ว","emerald");
  }
  function downloadTemplate(){
    const rows=[
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:"0.008",Count:5},
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:"0.016",Count:60},
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:"0.03",Count:180},
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:"0.06",Count:90},
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:"0.125",Count:25},
      {Organism:"Escherichia coli",Antibiotic:"Ciprofloxacin",MIC:">2",Count:47},
    ];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Template");
    XLSX.writeFile(wb,"MIC_template.xlsx");
  }

  // ---- filter derivations ----
  const orgOptions=useMemo(()=>[...new Set(analyses.map(a=>a.organism))].sort(),[analyses]);
  const abOptions=useMemo(()=>[...new Set(analyses.filter(a=>!fOrg||a.organism===fOrg).map(a=>a.antibiotic))].sort(),[analyses,fOrg]);
  const filteredAnalyses=useMemo(()=>analyses.filter(a=>{
    if(fOrg&&a.organism!==fOrg) return false;
    if(fAb&&a.antibiotic!==fAb) return false;
    if(fEcoff==="has"&&!hasEcoff(a)) return false;
    if(fEcoff==="none"&&hasEcoff(a)) return false;
    return true;
  }),[analyses,fOrg,fAb,fEcoff]);
  const anyFilter=fOrg||fAb||fEcoff;
  const clearFilters=()=>{ setFOrg(""); setFAb(""); setFEcoff(""); };
  const current=filteredAnalyses.find(a=>keyOf(a)===selectedKey)||filteredAnalyses[0]||null;
  return (
    <div style={{color:C.ink,minHeight:"100vh"}}>
      <header className="border-b" style={{borderColor:C.line,background:C.surface}}>
        <div className="mx-auto max-w-7xl px-5 py-3 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-md" style={{background:C.teal}}>
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2"><path d="M3 3v18h18"/><path d="M7 14l3-5 3 4 4-7"/></svg>
            </div>
            <div>
              <div className="text-base font-bold leading-tight" style={display}>MIC · ECOFF Analyzer</div>
              <div className="text-xs" style={{color:C.sub}}>วิเคราะห์การกระจาย MIC และกำหนดค่า ECOFF</div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={()=>fileRef.current?.click()} className="rounded-md px-3 py-2 text-sm font-medium text-white" style={{background:C.teal}}>อัปโหลดไฟล์</button>
            <input ref={fileRef} type="file" accept=".csv,.tsv,.txt,.xlsx,.xls" onChange={onFile} className="hidden" />
            <button onClick={downloadTemplate} className="rounded-md border px-3 py-2 text-sm font-medium" style={{borderColor:C.line2,color:C.sub}}>เทมเพลต</button>
            <button onClick={()=>setView(v=>v==="help"?(active?"overview":"home"):"help")} className="rounded-md border px-3 py-2 text-sm font-medium" style={{borderColor:view==="help"?C.teal:C.line2,color:view==="help"?C.teal:C.sub}}>คู่มือ</button>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-7xl px-5 py-5 grid grid-cols-12 gap-5">
        <aside className="col-span-12 lg:col-span-3 space-y-4">
          <Section title="ชุดข้อมูล">
            {!hasStore&&<p className="text-xs mb-2" style={{color:C.amber}}>โหมดชั่วคราว: เบราว์เซอร์นี้ไม่บันทึกข้อมูลถาวร</p>}
            {datasets.length===0&&<p className="text-sm" style={{color:C.sub}}>ยังไม่มีข้อมูล — อัปโหลดไฟล์เพื่อเริ่ม</p>}
            <ul className="space-y-1">
              {datasets.map(d=>(
                <li key={d.id} className="flex items-center justify-between rounded-md px-2 py-1.5 text-sm cursor-pointer"
                  style={{background:active?.id===d.id?C.tealSoft:"transparent"}} onClick={()=>openDataset(d.id)}>
                  <span className="truncate">{d.name}</span>
                  <button onClick={e=>{e.stopPropagation();removeDataset(d.id);}} className="text-xs ml-2 shrink-0" style={{color:C.sub}}>ลบ</button>
                </li>))}
            </ul>
          </Section>
          {analyses.length>0&&(
            <Section title="ตัวกรอง (Filter)">
              <div className="space-y-2">
                <label className="block">
                  <span className="text-xs" style={{color:C.sub}}>เชื้อ (Organism)</span>
                  <select value={fOrg} onChange={e=>{setFOrg(e.target.value); setFAb("");}} className="mt-1 w-full rounded-md border px-2 py-1.5 text-sm" style={{borderColor:C.line2}}>
                    <option value="">ทั้งหมด ({orgOptions.length})</option>
                    {orgOptions.map(o=><option key={o} value={o}>{o}</option>)}
                  </select>
                </label>
                <label className="block">
                  <span className="text-xs" style={{color:C.sub}}>ยา (Antibiotic)</span>
                  <select value={fAb} onChange={e=>setFAb(e.target.value)} className="mt-1 w-full rounded-md border px-2 py-1.5 text-sm" style={{borderColor:C.line2}}>
                    <option value="">ทั้งหมด ({abOptions.length})</option>
                    {abOptions.map(o=><option key={o} value={o}>{o}</option>)}
                  </select>
                </label>
                <label className="block">
                  <span className="text-xs" style={{color:C.sub}}>สถานะ ECOFF</span>
                  <select value={fEcoff} onChange={e=>setFEcoff(e.target.value)} className="mt-1 w-full rounded-md border px-2 py-1.5 text-sm" style={{borderColor:C.line2}}>
                    <option value="">ทั้งหมด</option>
                    <option value="has">มี ECOFF</option>
                    <option value="none">ไม่มี ECOFF</option>
                  </select>
                </label>
                {anyFilter&&(
                  <button onClick={clearFilters} className="w-full rounded-md border px-2 py-1.5 text-xs font-medium" style={{borderColor:C.line2,color:C.teal}}>
                    ล้างตัวกรอง · แสดง {filteredAnalyses.length}/{analyses.length}
                  </button>)}
              </div>
            </Section>)}
          {analyses.length>0&&(
            <Section title={`คู่เชื้อ–ยา (${filteredAnalyses.length}${anyFilter?`/${analyses.length}`:""})`}>
              <ul className="space-y-1 max-h-96 overflow-auto mic-scroll">
                {filteredAnalyses.map((a)=>{
                  const k=keyOf(a);
                  const worst=a.flags.some(f=>f.level==="error")?"rose":a.flags.some(f=>f.level==="warn")?"amber":"emerald";
                  return (
                    <li key={k} onClick={()=>{setSelectedKey(k);setView("dashboard");}} className="rounded-md px-2 py-1.5 cursor-pointer"
                      style={{background:selectedKey===k?C.tealSoft:"transparent"}}>
                      <div className="flex items-center gap-2">
                        <span className="h-2 w-2 rounded-full shrink-0" style={{background:worst==="rose"?C.rose:worst==="amber"?C.amber:C.emerald}} title={hasEcoff(a)?"มี ECOFF":"ไม่มี ECOFF"} />
                        <span className="text-sm italic truncate" style={display}>{a.organism}</span>
                        {!hasEcoff(a)&&<span className="text-xs shrink-0" style={{color:C.sub}}>· no ECOFF</span>}
                      </div>
                      <span className="text-xs ml-4" style={{color:C.sub}}>{a.antibiotic} · n={a.pct?.N??0}</span>
                    </li>);
                })}
                {filteredAnalyses.length===0&&<li className="text-sm px-2 py-1" style={{color:C.sub}}>ไม่มีรายการตรงกับตัวกรอง</li>}
              </ul>
            </Section>)}
        </aside>

        <main className="col-span-12 lg:col-span-9 space-y-4">
          {view==="help"
            ? <HelpPage onBack={()=>setView(active?"overview":"home")} />
            : !active
              ? <EmptyState onUpload={()=>fileRef.current?.click()} onTemplate={downloadTemplate} onHelp={()=>setView("help")} />
              : (<>
            <div className="flex items-center justify-between gap-3">
              <div className="inline-flex rounded-md border" style={{borderColor:C.line2}}>
                {[["overview","ภาพรวม"],["dashboard","รายคู่เชื้อ–ยา"],["raw","ข้อมูลดิบ"]].map(([k,t])=>(
                  <button key={k} onClick={()=>setView(k)} className="px-3 py-1.5 text-sm font-medium"
                    style={{background:view===k?C.teal:C.surface,color:view===k?"#fff":C.sub}}>{t}</button>))}
              </div>
              <button onClick={exportAll} disabled={!filteredAnalyses.length} className="rounded-md border px-3 py-2 text-sm font-medium" style={{borderColor:C.teal,color:filteredAnalyses.length?C.teal:C.line2}}>ส่งออก Excel{anyFilter?` (${filteredAnalyses.length})`:""}</button>
            </div>
            {view==="overview"&&<OverviewPage analyses={filteredAnalyses} anyFilter={anyFilter} onPick={(k)=>{setSelectedKey(k);setView("dashboard");}} />}
            {view==="dashboard"&&current&&<ComboDashboard a={current} onOverride={setOverride} />}
            {view==="dashboard"&&!current&&(
              <div className="rounded-lg border p-8 text-center text-sm" style={{borderColor:C.line,background:C.surface,color:C.sub}}>
                ไม่มีคู่เชื้อ–ยาที่ตรงกับตัวกรอง — ลองปรับหรือ <button onClick={clearFilters} className="underline" style={{color:C.teal}}>ล้างตัวกรอง</button>
              </div>)}
            {view==="raw"&&<RawData analyses={filteredAnalyses} onEdit={updateCount} />}
          </>)}
        </main>
      </div>

      {pendingFile&&<ImportModal pending={pendingFile} setPending={setPendingFile} onConfirm={confirmImport} />}
      {toast&&(
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 rounded-md px-4 py-2 text-sm text-white shadow-lg z-50"
          style={{background:toast.tone==="rose"?C.rose:toast.tone==="emerald"?C.emerald:toast.tone==="slate"?C.sub:C.teal}}>{toast.msg}</div>)}
    </div>);
}

export default App;
