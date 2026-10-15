import React from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
window.jspdf = { jsPDF };
window.autoTable = autoTable;

// ---- calculation logic (mirrors the sheet formulas) ----
const norm=s=>String(s==null?"":s).replace(/\u00a0/g," ").trim().toUpperCase();
const num=v=>{const n=typeof v==="number"?v:parseFloat(String(v==null?"":v).replace(/,/g,""));return isFinite(n)?n:0};
function sumBy(rows,kc,qc,flt){const m=new Map();for(const r of rows){if(flt&&!flt(r))continue;const k=norm(r[kc]);if(!k)continue;m.set(k,(m.get(k)||0)+num(r[qc]))}return m}
// Consumption = Usage / dataMonths * months ; Shortfall = Available stock - Consumption (negative = short)
function compute(T,months,dataMonths,safety,basis){
 const g=(m,k)=>m.get(k)||0;
 const imp=sumBy(T.DWH,0,2),inh=sumBy(T.DWH,0,3),dwh=sumBy(T.DWH,0,4),bwh=sumBy(T.BWH,1,5),lwh=sumBy(T.LWH,0,2),mwh=sumBy(T.MWH,0,2);
 const po=sumBy(T.PO,2,5),tr=sumBy(T.PO,2,5,r=>norm(r[6])==="IN TRANSIT");
 const del=sumBy(T.DEL,6,9),out=sumBy(T.OUT,6,9),up=sumBy(T.UP,8,11),blr=sumBy(T.BLR,8,11);
 const dm=dataMonths>0?dataMonths:1;
 return T.TD.map(m=>{const k=norm(m[3]);
  const o={v:m[0],t:m[1],cat:m[2],code:m[3],desc:m[4],imp:g(imp,k),inh:g(inh,k),dwh:g(dwh,k),bwh:g(bwh,k),lwh:g(lwh,k),mwh:g(mwh,k),tr:g(tr,k),po:g(po,k),del:g(del,k),out:g(out,k),up:g(up,k),blr:g(blr,k)};
  const use=(o.del+o.out+o.up+o.blr)/dm*months;
  const avail=basis==="ALL"?o.dwh+o.bwh+o.lwh+o.mwh+o.po:o.mwh;
  const sf=avail-use;
  const ord=Math.max(0,Math.ceil(Math.round((-sf+safety)*1e6)/1e6));
  o.use=use;o.avail=avail;o.sf=sf;o.ord=ord;o.st=use===0?"No Usage":ord>0?"Shortfall":"Sufficient";return o})}
// Child table: every transaction that makes up a Cat. No.'s consumption
const isoDate=v=>{if(v==null||v==="")return"";const s=String(v);let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);if(m)return m[1]+"-"+m[2]+"-"+m[3];m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);if(m)return m[3]+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0");return s};
const SRC=[["DEL","Del",6,9,2,1,3,5],["OUT","Outstation",6,9,1,0,2,4],["UP","UP",8,11,1,0,2,5],["BLR","Bangalore",8,11,1,0,3,4]];
function consumptionOf(T,code){const k=norm(code),rows=[];
 for(const[id,label,kc,qc,dc,doc,party,person]of SRC)for(const r of T[id]){if(norm(r[kc])!==k)continue;rows.push({src:label,date:isoDate(r[dc]),doc:r[doc]==null?"":String(r[doc]),party:r[party]==null?"":String(r[party]),person:r[person]==null?"":String(r[person]),qty:num(r[qc])})}
 rows.sort((a,b)=>b.date.localeCompare(a.date));return rows}

// ---- UI ----
const h=React.createElement,{useState,useMemo,useEffect}=React;
const TD_KEYS=["v","t","cat","code","desc","imp","inh","dwh","bwh","lwh","mwh","tr","po","del","out","up","blr","use","sf","ord","st"];
const tdHead=(mo,dm,basis)=>["vendor","Implant/ Instrument","category","Cat. No.","Description","DWH - Imprest","DWH- IN HAND","DWH- Tot. stock","BWH","LWH","MWH","IN TRANSIT","pending as on - 22.09.26","Del","outstation supply","up","Bangalore","Total Usage / Consumption (÷"+dm+" ×"+mo+")","Shortfall ("+(basis==="ALL"?"Stock":"MWH")+" − Consumption)","Qty to be ordered","Status"];
const TABS=[
 {id:"TD",name:"Transaction Data"},
 {id:"PO",name:"PENDING ORDER",heads:["PO NO","PO Dt","CAT. NO.","DESCRIPTION","PO QTY","PENDING QTY.","STATUS"],chk:"Cat. No.",key:2},
 {id:"MWH",name:"MWH",heads:["Item Code","Item","Quantity","Brand Name"],chk:"CHECK",key:0},
 {id:"DWH",name:"DWH",heads:["Item Code","Particulars","IMPREST QTY.","IN HAND QTY.","TOTAL QTY.","Division Name"],chk:"CHECK",key:0},
 {id:"BWH",name:"BWH",heads:["MAKE","Cat No","Product Description","Consignment Pending","Stock at Office","Total Stock"],chk:"CHECK",key:1},
 {id:"LWH",name:"LWH",heads:["Item Code","Particulars","Net Quantity","Brand Name","Warehouse Name","Division Name"],chk:"CHECK",key:0},
 {id:"UP",name:"up usage",heads:["Doc No","Date","Customer Account","Location.Name","Warehouse","Doctor","Sales Executive","Narration","Item.Code","Item.Name","Item.Brand","Quantity"],chk:"CHECK",key:8},
 {id:"DEL",name:"del usage",heads:["Slip No.","Slip No.","SLIP DATE","HOSPITAL NAME","Region","SURGEON NAME","CAT No.","PRODUCT DESCRIPTION","Item. Brand","QTY"],chk:"CHECK",key:6},
 {id:"OUT",name:"outstation",heads:["Focus Slip No.","SLIP DATE","HOSPITAL NAME","Region","SURGEON NAME","PRINCIPAL COMPANY","CAT No.","PRODUCT DESCRIPTION","Item.Brand","QTY","CONSIGNMENT"],chk:"CHECK",key:6},
 {id:"BLR",name:"blr usage",heads:["Surgery No.","Surgery Date","Attended Person","Hospital's Name","Surgeon's Name","Patient's Name","Supplied Through","Area/ City","Cat. No.","Product Used","Item.Brand","Qty. Used"],chk:"CHECK",key:8}
];
const STAT=["All","Shortfall","Sufficient","No Usage"];
const fmt=n=>Math.round(n).toLocaleString("en-IN");
const f1=n=>(Math.round(n*10)/10).toLocaleString("en-IN");
const nk=s=>String(s==null?"":s).toLowerCase().replace(/[^a-z0-9]/g,"");
const cellStr=v=>v==null?"":v instanceof Date?v.toISOString().slice(0,10):String(v);
let DLP=null;const getDl=()=>DLP||(DLP=(window.claude&&window.claude.use?window.claude.use("downloads"):Promise.resolve(null)).catch(()=>null));
async function saveFile(name,data){const dl=await getDl();
 if(dl){try{await dl.save({filename:name,data});return}catch(e){if(e&&e.code==="declined")return;throw e}}
 const blob=new Blob([data]);const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),2000)}
function makeXlsx(sheets){const wb=XLSX.utils.book_new();sheets.forEach(([n,a])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(a),n.slice(0,31)));return XLSX.write(wb,{bookType:"xlsx",type:"array"})}
function makePdf(title,head,body){const doc=new window.jspdf.jsPDF({orientation:"landscape",unit:"pt",format:"a3"});doc.setFontSize(12);doc.text(title,24,26);
 const opts={head:[head],body,startY:36,styles:{fontSize:6,cellPadding:2},headStyles:{fillColor:[31,56,100]},margin:{left:24,right:24},
  didParseCell:d=>{if(d.section==="body"&&d.column.index===head.length-1&&d.cell.raw==="Shortfall"){d.cell.styles.textColor=[192,57,43];d.cell.styles.fontStyle="bold"}}};
 if(typeof doc.autoTable==="function")doc.autoTable(opts);else window.autoTable(doc,opts);return doc.output("arraybuffer")}

function ImportModal({tab,onClose,onApply}){
 const[file,setFile]=useState(null),[sheets,setSheets]=useState([]),[wbk,setWbk]=useState(null),[sn,setSn]=useState(""),[aoa,setAoa]=useState([]),[hr,setHr]=useState(1),[map,setMap]=useState([]),[mode,setMode]=useState("replace"),[err,setErr]=useState(""),[busy,setBusy]=useState(false);
 const targets=tab.heads;
 const suggest=(raw,hi)=>{const hd=(raw[hi-1]||[]);const used=new Set();setMap(targets.map(t=>{const a=nk(t);let j=hd.findIndex((x,i)=>!used.has(i)&&nk(x)===a);if(j<0)j=hd.findIndex((x,i)=>!used.has(i)&&x!=null&&a.length>2&&(nk(x).includes(a)||a.includes(nk(x))));if(j>=0)used.add(j);return j}))};
 const load=(wb,name)=>{const raw=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:null,blankrows:false});setAoa(raw);let hi=raw.findIndex(r=>r.filter(x=>x!=null&&x!=="").length>=3)+1;if(hi<1)hi=1;setHr(hi);suggest(raw,hi)};
 const onFile=async e=>{const f=e.target.files[0];if(!f)return;setErr("");try{const buf=await f.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:true});setFile(f.name);setWbk(wb);setSheets(wb.SheetNames);const pick=wb.SheetNames.find(s=>nk(s)===nk(tab.name))||wb.SheetNames[0];setSn(pick);load(wb,pick)}catch(x){setErr("Could not read file: "+x.message)}};
 const hd=aoa[hr-1]||[];const colLabel=i=>XLSX.utils.encode_col(i)+": "+(cellStr(hd[i])||"(blank)");
 const rows=useMemo(()=>aoa.slice(hr).map(r=>map.map(j=>j<0?null:(r[j]instanceof Date?r[j].toISOString().slice(0,10):r[j]))).filter(r=>r.some(v=>v!=null&&v!=="")),[aoa,hr,map]);
 const ncols=Math.max(0,...aoa.slice(0,50).map(r=>r.length));const miss=map[tab.key]<0;
 return h("div",{className:"ov"},h("div",{className:"md"},
  h("div",{className:"mh"},h("b",null,"Import into “"+tab.name+"”"),h("button",{onClick:onClose},"✕")),
  h("div",{className:"mb"},
   h("p",{className:"mut"},"1. Choose an Excel / CSV file · 2. Pick the header row · 3. Map every column manually · 4. Import."),
   h("input",{type:"file",accept:".xlsx,.xls,.csv",onChange:onFile}),err&&h("p",{className:"er"},err),
   wbk&&h("div",null,
    h("div",{className:"bar",style:{marginTop:10}},
     sheets.length>1&&h("label",null,"Sheet ",h("select",{value:sn,onChange:e=>{setSn(e.target.value);load(wbk,e.target.value)}},sheets.map(s=>h("option",{key:s},s)))),
     h("label",null,"Header row ",h("input",{type:"number",min:1,value:hr,style:{width:60},onChange:e=>{const v=Math.max(1,+e.target.value||1);setHr(v);suggest(aoa,v)}})),
     h("label",null,"Mode ",h("select",{value:mode,onChange:e=>setMode(e.target.value)},h("option",{value:"replace"},"Replace existing data"),h("option",{value:"append"},"Append to existing data")))),
    h("table",{className:"mp"},h("thead",null,h("tr",null,h("th",null,"Dashboard column"),h("th",null,"← Column in your file"))),
     h("tbody",null,targets.map((t,i)=>h("tr",{key:i},h("td",null,t,i===tab.key&&h("span",{className:"mut"}," (item key)")),
      h("td",null,h("select",{value:map[i]==null?-1:map[i],onChange:e=>{const m=[...map];m[i]=+e.target.value;setMap(m)}},h("option",{value:-1},"— not mapped —"),Array.from({length:ncols},(_,j)=>h("option",{key:j,value:j},colLabel(j))))))))),
    h("p",{className:"mut"},"Preview (first 4 of "+fmt(rows.length)+" rows):"),
    h("div",{className:"tw",style:{maxHeight:140}},h("table",null,h("thead",null,h("tr",null,targets.map((t,i)=>h("th",{key:i,style:{cursor:"default"}},t)))),h("tbody",null,rows.slice(0,4).map((r,i)=>h("tr",{key:i},r.map((v,j)=>h("td",{key:j},cellStr(v)))))))),
    miss&&h("p",{className:"er"},"Map the item key column ("+targets[tab.key]+") – without it nothing can be matched."))),
  h("div",{className:"mf"},h("button",{onClick:onClose},"Cancel"),h("button",{className:"pr",disabled:!wbk||miss||!rows.length||busy,onClick:async()=>{setBusy(true);try{await onApply(tab,rows,mode);onClose()}catch(e){setErr(String(e.message||e));setBusy(false)}}},busy?"Importing…":"Import "+(wbk?fmt(rows.length)+" rows":"")))));
}

function ChildModal({T,row,mo,dm,basis,onClose,flash}){
 const ent=useMemo(()=>consumptionOf(T,row.code),[T,row.code]);
 const months=useMemo(()=>Array.from(new Set(ent.map(e=>e.date.slice(0,7)).filter(Boolean))).sort(),[ent]);
 const SRCN=["Del","Outstation","UP","Bangalore"];
 const piv=months.map(m=>{const o={m};let t=0;SRCN.forEach(s=>{o[s]=ent.filter(e=>e.date.startsWith(m)&&e.src===s).reduce((a,e)=>a+e.qty,0);t+=o[s]});o.t=t;return o});
 const total=ent.reduce((a,e)=>a+e.qty,0);const[srcF,setSrcF]=useState("All");
 const list=ent.filter(e=>srcF==="All"||e.src===srcF);
 const kp=(l,v,c)=>h("div",{className:"k "+(c||"")},h("small",null,l),h("b",null,v));
 const exp=async()=>{try{await saveFile("Consumption_"+row.code.replace(/[^A-Za-z0-9-]/g,"_")+".xlsx",makeXlsx([["Monthly",[["Month",...SRCN,"Total"],...piv.map(p=>[p.m,...SRCN.map(s=>p[s]),p.t])]],["Transactions",[["Source","Date","Doc No","Hospital / Customer","Surgeon / Doctor","Qty"],...ent.map(e=>[e.src,e.date,e.doc,e.party,e.person,e.qty])]]]))}catch(e){flash("Export failed: "+(e.message||e))}};
 return h("div",{className:"ov",onClick:onClose},h("div",{className:"md wide",onClick:e=>e.stopPropagation()},
  h("div",{className:"mh"},h("div",null,h("b",null,row.code),h("div",{className:"mut"},row.desc+" · "+row.v)),h("div",{className:"bar",style:{margin:0}},h("button",{onClick:exp},"⬇ Export"),h("button",{onClick:onClose},"✕"))),
  h("div",{className:"mb"},
   h("div",{className:"kp"},kp("Actual usage",fmt(total)),kp("Consumption ("+mo+"M)",f1(row.use)),kp("MWH stock",fmt(row.mwh)),kp("Shortfall",f1(row.sf),row.sf<0?"o":""),kp("Qty to be ordered",fmt(row.ord),row.ord>0?"o":"")),
   h("p",{className:"mut"},"Other stock: DWH "+fmt(row.dwh)+" · BWH "+fmt(row.bwh)+" · LWH "+fmt(row.lwh)+" · Pending PO "+fmt(row.po)),
   h("b",null,"Consumption by month"),
   ent.length===0?h("p",{className:"mut"},"No consumption recorded for this Cat. No."):
   h("div",{className:"tw",style:{maxHeight:220,margin:"6px 0 12px"}},h("table",null,h("thead",null,h("tr",null,["Month",...SRCN,"Total"].map(x=>h("th",{key:x,style:{cursor:"default"}},x)))),
    h("tbody",null,piv.map(p=>h("tr",{key:p.m},h("td",{className:"l"},p.m),SRCN.map(s=>h("td",{key:s},p[s]?fmt(p[s]):"–")),h("td",null,h("b",null,fmt(p.t))))),
     h("tr",null,h("td",{className:"l"},h("b",null,"Total")),SRCN.map(s=>h("td",{key:s},h("b",null,fmt(ent.filter(e=>e.src===s).reduce((a,e)=>a+e.qty,0))))),h("td",null,h("b",null,fmt(total))))))),
   ent.length>0&&h("div",null,h("div",{className:"bar",style:{marginBottom:6}},h("b",null,"Transactions ("+list.length+")"),h("select",{value:srcF,onChange:e=>setSrcF(e.target.value)},["All",...SRCN].map(s=>h("option",{key:s},s)))),
    h("div",{className:"tw",style:{maxHeight:300}},h("table",null,h("thead",null,h("tr",null,["Source","Date","Doc No","Hospital / Customer","Surgeon / Doctor","Qty"].map(x=>h("th",{key:x,style:{cursor:"default"}},x)))),
     h("tbody",null,list.slice(0,500).map((e,i)=>h("tr",{key:i},h("td",{className:"l"},e.src),h("td",{className:"l"},e.date),h("td",{className:"l"},e.doc),h("td",{className:"l"},e.party),h("td",{className:"l"},e.person),h("td",null,fmt(e.qty)))))))))));
}

function Dashboard({initial,onImport,onSave,onReload,sheetUrl}){
 const[T,setT]=useState(initial);
 const[tab,setTab]=useState("TD");
 const[q,setQ]=useState(""),[v,setV]=useState("All"),[ty,setTy]=useState("All"),[c,setC]=useState("All"),[st,setSt]=useState("All"),[dm,setDm]=useState(9),[mo,setMo]=useState(6),[basis,setBasis]=useState("MWH"),[s,setS]=useState({k:null,d:1}),[pg,setPg]=useState(0),[imp,setImp]=useState(null),[child,setChild]=useState(null),[toast,setToast]=useState("");
 useEffect(()=>{setPg(0)},[tab,q,v,ty,c,st,mo,dm,basis,s]);
 useEffect(()=>{setQ("");setS({k:null,d:1})},[tab]);
 const flash=m=>{setToast(m);setTimeout(()=>setToast(""),4000)};
 const all=useMemo(()=>compute(T,mo,dm,0,basis),[T,mo,dm,basis]);
 const master=useMemo(()=>new Set(T.TD.map(r=>norm(r[3]))),[T]);
 const def=TABS.find(t=>t.id===tab);
 const uniq=k=>["All",...Array.from(new Set(all.map(r=>r[k]).filter(Boolean))).sort()];
 const HEAD=tdHead(mo,dm,basis);
 let heads,rows;
 if(tab==="TD"){heads=HEAD;
  let r=all.filter(x=>(v==="All"||x.v===v)&&(ty==="All"||x.t===ty)&&(c==="All"||x.cat===c)&&(st==="All"||x.st===st)&&(!q||(x.code+" "+x.desc).toLowerCase().includes(q.toLowerCase())));
  if(s.k!==null){const k=TD_KEYS[s.k];r=[...r].sort((a,b)=>(typeof a[k]==="number"?a[k]-b[k]:String(a[k]).localeCompare(String(b[k])))*s.d)}rows=r}
 else{heads=[...def.heads,def.chk];
  let r=T[tab].map(x=>[...x,master.has(norm(x[def.key]))?"OK":"Not in master"]);
  if(q)r=r.filter(x=>x.some(y=>String(y==null?"":y).toLowerCase().includes(q.toLowerCase())));
  if(s.k!==null)r=[...r].sort((a,b)=>{const A=a[s.k],B=b[s.k];return(typeof A==="number"&&typeof B==="number"?A-B:String(A==null?"":A).localeCompare(String(B==null?"":B)))*s.d});rows=r}
 const PS=200,pages=Math.max(1,Math.ceil(rows.length/PS)),shown=rows.slice(pg*PS,pg*PS+PS);
 const sum=k=>rows.reduce((a,r)=>a+r[k],0);
 const apply=async(t,newRows,mode)=>{if(onImport)await onImport(t.id,newRows,mode,setT);else setT(p=>({...p,[t.id]:mode==="replace"?newRows:[...p[t.id],...newRows]}));flash("Imported "+fmt(newRows.length)+" rows into "+t.name)};
 const tdAoa=()=>[HEAD,...rows.map(r=>TD_KEYS.map(k=>typeof r[k]==="number"&&(k==="use"||k==="sf")?Math.round(r[k]*10)/10:r[k]))];
 const run=async fn=>{try{await fn()}catch(e){flash("Export failed: "+((e&&e.message)||e))}};
 const expXlsx=()=>run(async()=>{const a=tab==="TD"?tdAoa():[heads,...rows];await saveFile("SPINE_"+def.name.replace(/\s+/g,"_")+".xlsx",makeXlsx([[def.name,a]]))});
 const expPdf=()=>run(async()=>{const a=tdAoa();await saveFile("SPINE_Transaction_Data.pdf",makePdf("SPINE – Transaction Data · Consumption = Usage ÷ "+dm+" × "+mo+" months · Shortfall vs "+(basis==="ALL"?"all stock + PO":"MWH stock")+" · "+rows.length+" items",a[0],a.slice(1).map(r=>r.map(x=>typeof x==="number"?String(x):x))))});
 const refresh=async()=>{try{await onReload(setT);flash("Refreshed from Google Sheet")}catch(e){flash("Refresh failed: "+(e.message||e))}};
 const save=async()=>{const items=rows.map(r=>({code:r.code,qty:r.ord,status:r.st}));
  if(!window.confirm("Write Qty to be ordered and Status for "+items.length+" item(s) into the Transaction Data sheet?\nExisting values in those two columns will be overwritten (previous values are kept in the SPINE_Log tab)."))return;
  try{await onSave(items,{months:mo,dataMonths:dm,basis},setT);flash("Saved "+items.length+" items to the Google Sheet")}catch(e){flash("Save failed: "+(e.message||e))}};
 const template=()=>run(async()=>saveFile("Template_"+def.name.replace(/\s+/g,"_")+".xlsx",makeXlsx([[def.name,[def.heads]]])));
 const sel=(val,set,opts,lab)=>h("select",{value:val,onChange:e=>set(e.target.value)},opts.map(o=>h("option",{key:o,value:o},o==="All"&&lab?lab:o)));
 const kp=(l,val,cls)=>h("div",{className:"k "+(cls||"")},h("small",null,l),h("b",null,val));
 const th=(l,i)=>h("th",{key:i,onClick:()=>setS({k:i,d:s.k===i?-s.d:1})},l+(s.k===i?(s.d>0?" ▲":" ▼"):""));
 const nCheck=tab==="TD"?0:rows.filter(r=>r[r.length-1]!=="OK").length;
 const numIn=(val,set)=>h("input",{type:"number",min:0,step:"any",value:val,style:{width:62},onChange:e=>set(e.target.value===""?0:+e.target.value)});
 return h("div",null,
  h("div",{className:"top"},h("h1",null,"SPINE – Transaction Data"),h("div",{className:"bar",style:{margin:0}},onReload&&h("button",{onClick:refresh},"↻ Refresh"),sheetUrl&&h("a",{className:"lk",href:sheetUrl,target:"_blank",rel:"noreferrer"},"Open Google Sheet ↗"))),
  h("div",{className:"tabs"},TABS.map(t=>h("button",{key:t.id,className:t.id===tab?"on":"",onClick:()=>setTab(t.id)},t.name,h("span",{className:"cnt"},fmt(t.id==="TD"?T.TD.length:T[t.id].length))))),
  h("div",{className:"bar"},
   h("input",{placeholder:tab==="TD"?"Search Cat. No. / Description":"Search this tab",value:q,onChange:e=>setQ(e.target.value)}),
   tab==="TD"&&[sel(v,setV,uniq("v"),"All vendors"),sel(ty,setTy,uniq("t"),"All Implant/Instrument"),sel(c,setC,uniq("cat"),"All categories"),
    h("select",{key:"st",value:st,onChange:e=>setSt(e.target.value)},STAT.map(o=>h("option",{key:o,value:o},o==="All"?"All status":o))),
    h("button",{key:"rs",onClick:()=>{setQ("");setV("All");setTy("All");setC("All");setSt("All")}},"Reset")],
   h("span",{className:"sp"}),
   tab!=="TD"&&h("button",{className:"pr",onClick:()=>setImp(def)},"⬆ Import"),
   tab==="TD"&&onSave&&h("button",{className:"pr",onClick:save,title:"Write Qty to be ordered + Status of the rows shown into the Google Sheet"},"⬆ Save to Sheet"),h("button",{onClick:expXlsx},"⬇ Excel"),tab==="TD"&&h("button",{onClick:expPdf},"⬇ PDF"),tab!=="TD"&&h("button",{onClick:template},"Template")),
  tab==="TD"&&h("div",{className:"bar var"},h("b",null,"Variables:"),h("span",null,"Consumption = Usage ÷"),numIn(dm,setDm),h("span",null,"×"),numIn(mo,setMo),h("span",null,"months"),
   h("span",{className:"mut"},"· Shortfall = "),h("select",{value:basis,onChange:e=>setBasis(e.target.value)},h("option",{value:"MWH"},"MWH stock − Consumption"),h("option",{value:"ALL"},"All warehouses + PO − Consumption"))),
  tab==="TD"?h("div",{className:"kp"},kp("Items shown",fmt(rows.length)),kp("Consumption ("+mo+"M)",fmt(sum("use"))),kp(basis==="ALL"?"Available stock":"MWH stock",fmt(sum("avail"))),kp("Shortfall items",fmt(rows.filter(r=>r.st==="Shortfall").length),"o"),kp("Qty to be ordered",fmt(sum("ord")),"o"))
   :h("div",{className:"kp"},kp("Rows",fmt(rows.length)),kp("Cat. nos not in master",fmt(nCheck),nCheck?"o":"")),
  h("div",{className:"tw"},h("table",null,
   h("thead",null,h("tr",null,heads.map((l,i)=>th(l,i)))),
   h("tbody",null,tab==="TD"?shown.map((r,i)=>h("tr",{key:i,className:r.st==="Shortfall"?"hot":""},TD_KEYS.map((k,j)=>{const x=r[k];
     if(k==="st")return h("td",{key:k},h("span",{className:"bd "+x.replace(" ","")},x));
     if(k==="code")return h("td",{key:k,className:"l"},h("a",{className:"lk",title:"Show consumption detail",onClick:()=>setChild(r)},x));
     return h("td",{key:k,className:j<5?"l":""},typeof x==="number"?(k==="use"?f1(x):k==="sf"?f1(x):fmt(x)):x)}))):
    shown.map((r,i)=>h("tr",{key:i},r.map((x,j)=>h("td",{key:j,className:typeof x==="number"?"":"l"},j===r.length-1?h("span",{className:"bd "+(x==="OK"?"Sufficient":"Shortfall")},x):typeof x==="number"?fmt(x):cellStr(x))))))) ),
  h("div",{className:"pgn"},h("button",{disabled:pg===0,onClick:()=>setPg(pg-1)},"‹ Prev"),h("span",null,"Page "+(pg+1)+" of "+pages+" · "+fmt(rows.length)+" rows"),h("button",{disabled:pg>=pages-1,onClick:()=>setPg(pg+1)},"Next ›")),
  tab==="TD"&&h("div",{className:"note"},"Click a Cat. No. to see its consumption transactions. Total Usage = Del + outstation + up + Bangalore. Consumption = Usage ÷ "+dm+" × "+mo+". Shortfall = "+(basis==="ALL"?"(DWH + BWH + LWH + MWH + pending PO)":"MWH")+" − Consumption; negative means short. Qty to be ordered = shortfall, rounded up."),
  imp&&h(ImportModal,{tab:imp,onClose:()=>setImp(null),onApply:apply}),
  child&&h(ChildModal,{T,row:child,mo,dm,basis,onClose:()=>setChild(null),flash}),
  toast&&h("div",{className:"toast"},toast));
}

export default Dashboard;
