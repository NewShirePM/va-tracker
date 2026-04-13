const {useState,useEffect,useCallback,useRef}=React;

// ── FAVICON ──
(()=>{const l=document.querySelector("link[rel='icon']")||document.createElement("link");l.rel="icon";l.href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%23CDA04B'/><text x='16' y='23' font-family='Georgia,serif' font-size='22' font-weight='bold' fill='%2328434C' text-anchor='middle'>V</text></svg>";document.head.appendChild(l)})();

// ── CONFIG ──
const CONFIG={
  clientId:"32e75ffa-747a-4cf0-8209-6a19150c4547",
  tenantId:"33575d04-ca7b-4396-8011-9eaea4030b46",
  siteId:"vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
  appName:"VA PRODUCTIVITY TRACKER",
};
const GRAPH="https://graph.microsoft.com/v1.0";
const SITE=`${GRAPH}/sites/${CONFIG.siteId}`;
const SCOPES=["Sites.ReadWrite.All","User.Read"];

// ── ROLE DETECTION ──
const ROLE_MAP={"Virtual Assistant":"va","Property Manager":"manager","Regional/Portfolio Manager":"regional","Owner/Operator":"admin"};
function detectRole(emp){
  const r=(emp.VATrackerRole||"").toLowerCase().trim();
  if(["va","manager","regional","admin"].includes(r))return r;
  return ROLE_MAP[emp.JobTitle]||null;
}

// ── DESIGN SYSTEM (matching V3 preview) ──
const C={
  teal:"#1C3740",t2:"#28434C",t3:"#3A6577",t4:"#4A7E91",t5:"#7AAFC0",
  tl:"#D6E7EC",tl0:"#EDF4F7",tl00:"#F4F9FB",
  gold:"#CDA04B",g2:"#B8922E",gl:"#F8F0DB",gl0:"#FFFBF2",
  bg:"#EDEEF0",white:"#FFFFFF",
  b1:"#DDE3E6",b2:"#C2CDD1",b3:"#9EAFB5",b4:"#6B8590",b6:"#3A5058",
  ok:"#1A7A46",okl:"rgba(26,122,70,0.1)",okb:"#EAF5EE",
  er:"#B83B2A",erl:"rgba(184,59,42,0.09)",erb:"#FDF0EE",
  wn:"#A86F08",wnl:"rgba(168,111,8,0.1)",wnb:"#FDF5E5",
  inf:"#2B5FA8",infl:"rgba(43,95,168,0.1)",infb:"#EBF1FB",
  pu:"#5B3FA8",pul:"rgba(91,63,168,0.09)",pub:"#F2EEFB",
};
const fnt="'DM Sans',system-ui,sans-serif";
const mono="'DM Mono',monospace";

// ── GRAPH API ──
async function gGet(token,url){const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});if(!r.ok)throw new Error(`GET ${r.status}`);return r.json();}
async function gAll(token,url){let a=[],n=url;while(n){const d=await gGet(token,n);a=a.concat(d.value||[]);n=d["@odata.nextLink"]||null;}return a;}
async function gPost(token,url,fields){const r=await fetch(url,{method:"POST",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify({fields})});if(!r.ok)throw new Error(`POST ${r.status}`);return r.json();}
async function gPatch(token,url,fields){const r=await fetch(url,{method:"PATCH",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify(fields)});if(!r.ok)throw new Error(`PATCH ${r.status}`);return r.json();}
function lUrl(n){return`${SITE}/lists/${n}/items`;}
function iUrl(n,id){return`${SITE}/lists/${n}/items/${id}/fields`;}
async function safeGet(token,name,url){try{const r=await gAll(token,url);return r;}catch(e){console.warn(`[VT] ${name} failed:`,e.message);return[];}}

async function loadAll(token){
  const[eR,cR,pR,ptR,aR]=await Promise.all([
    safeGet(token,"Employees",`${lUrl("Employees")}?expand=fields&$top=200`),
    safeGet(token,"Config",`${lUrl("VA_TrackerConfig")}?expand=fields&$top=10`),
    safeGet(token,"Properties",`${lUrl("VA_Properties")}?expand=fields&$top=200`),
    safeGet(token,"Portfolios",`${lUrl("VA_Portfolios")}?expand=fields&$top=200`),
    safeGet(token,"Activity",`${lUrl("VA_Activity")}?expand=fields&$top=1000`),
  ]);
  const employees=eR.map(e=>({id:e.id,...e.fields}));
  const config=cR.length>0?JSON.parse(cR[0].fields.ConfigJSON||"{}"):{};
  const properties=pR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const portfolios=ptR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const allActs=aR.map(a=>({id:a.id,...a.fields}));
  // Filter out activity before dataStartDate (if set in config)
  const cutoff=config.dataStartDate||null;
  const activities=cutoff?allActs.filter(a=>{const d=a.ActivityDate||a.StartTime||"";return d>=cutoff;}):allActs;
  const vas=employees.filter(e=>e.JobTitle==="Virtual Assistant"&&e.EmployeeActive!==false);
  const pms=employees.filter(e=>(e.JobTitle==="Property Manager"||e.JobTitle==="Regional/Portfolio Manager"||e.JobTitle==="Owner/Operator")&&e.EmployeeActive!==false);
  return{employees,config,properties,portfolios,activities,vas,pms};
}

// ── MSAL ──
function useMsal(){
  const[inst,setInst]=useState(null);const[acct,setAcct]=useState(null);const[token,setToken]=useState(null);const[err,setErr]=useState(null);
  useEffect(()=>{
    const s=document.createElement("script");s.src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js";
    s.onload=()=>{const i=new window.msal.PublicClientApplication({auth:{clientId:CONFIG.clientId,authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`,redirectUri:window.location.origin},cache:{cacheLocation:"sessionStorage"}});
      i.initialize().then(()=>{setInst(i);const a=i.getAllAccounts();if(a.length>0){setAcct(a[0]);i.acquireTokenSilent({scopes:SCOPES,account:a[0]}).then(r=>setToken(r.accessToken)).catch(()=>{});}});};
    document.head.appendChild(s);
  },[]);
  const login=useCallback(async()=>{if(!inst)return;try{const r=await inst.loginPopup({scopes:SCOPES});setAcct(r.account);const t=await inst.acquireTokenSilent({scopes:SCOPES,account:r.account});setToken(t.accessToken);setErr(null);}catch(e){if(e.errorCode!=="user_cancelled")setErr(e.message);}},[inst]);
  const refresh=useCallback(async()=>{if(!inst||!acct)return null;try{const r=await inst.acquireTokenSilent({scopes:SCOPES,account:acct});setToken(r.accessToken);return r.accessToken;}catch{const r=await inst.acquireTokenPopup({scopes:SCOPES});setToken(r.accessToken);return r.accessToken;}},[inst,acct]);
  return{acct,token,login,refresh,err};
}


// ── HELPERS ──
function fD(d){return d?new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}):"—";}
function fT(d){return d?new Date(d).toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit"}):"—";}
function fM(m){if(!m&&m!==0)return"—";const h=Math.floor(m/60),r=m%60;return h>0?`${h}h ${r}m`:`${r}m`;}
function fTm(s){const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),ss=s%60;const p=n=>String(n).padStart(2,"0");return h>0?`${h}:${p(m)}:${p(ss)}`:`${p(m)}:${p(ss)}`;}
function today(){return new Date().toISOString().slice(0,10);}
function dAgo(d){return Math.floor((Date.now()-new Date(d).getTime())/864e5);}
function inRange(d,from,to){if(!d)return false;const t=new Date(d).getTime();return t>=new Date(from).getTime()&&t<=new Date(to).getTime()+864e5;}
function initials(name){if(!name)return"?";const p=name.split(" ");return p.length>1?(p[0][0]+p[p.length-1][0]).toUpperCase():name.slice(0,2).toUpperCase();}
const catIcon={"Work Orders":"\u{1F527}",Marketing:"\u{1F4E2}","Tenant Comms":"\u{1F4AC}",Reporting:"\u{1F4CA}",Inspections:"\u{1F50D}",Renewals:"\u{1F4DD}",Accounts:"\u{1F4B0}","Admin/Other":"\u{1F4C1}",Leasing:"\u{1F511}"};

// ── SHARED STYLES ──
const base={fontFamily:fnt,color:C.t2};
const ss={
  app:{...base,display:"flex",flexDirection:"column",height:"100vh",background:C.bg,overflow:"hidden"},
  hdr:{background:C.teal,borderBottom:`2px solid ${C.gold}`,padding:"0 18px",display:"flex",alignItems:"center",justifyContent:"space-between",height:52,flexShrink:0,gap:10},
  logo:{width:30,height:30,background:C.gold,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:700,color:C.teal,flexShrink:0},
  tabs:{background:C.white,borderBottom:`1px solid ${C.b1}`,display:"flex",padding:"0 16px",overflowX:"auto",flexShrink:0,WebkitOverflowScrolling:"touch"},
  tab:a=>({padding:"0 15px",fontSize:12,fontWeight:a?700:500,color:a?C.t2:C.b4,borderBottom:`2.5px solid ${a?C.gold:"transparent"}`,cursor:"pointer",whiteSpace:"nowrap",background:"none",border:"none",borderTop:"none",borderLeft:"none",borderRight:"none",fontFamily:fnt,height:43,display:"inline-flex",alignItems:"center",gap:5}),
  content:{flex:1,overflowY:"auto",padding:"14px 14px 32px"},
  card:{background:C.white,border:`1px solid ${C.b1}`,borderRadius:8,boxShadow:"0 1px 3px rgba(28,55,64,0.07),0 1px 2px rgba(28,55,64,0.04)",padding:15,marginBottom:12},
  cardT:{fontSize:13,fontWeight:700,color:C.t2,lineHeight:1.3},
  cardS:{fontSize:10.5,color:C.b4,marginTop:2,lineHeight:1.4},
  sec:{fontSize:10,fontWeight:700,color:C.b4,textTransform:"uppercase",letterSpacing:"0.07em",marginBottom:8},
  label:{display:"block",fontSize:11,fontWeight:700,color:C.t2,marginBottom:4},
  input:{width:"100%",padding:"8px 10px",fontSize:12,fontFamily:fnt,color:C.t2,background:C.white,border:`1px solid ${C.b2}`,borderRadius:6,outline:"none",boxSizing:"border-box"},
  select:{width:"100%",padding:"8px 10px",fontSize:12,fontFamily:fnt,color:C.t2,background:C.white,border:`1px solid ${C.b2}`,borderRadius:6,cursor:"pointer",boxSizing:"border-box"},
  btn:(bg,fg)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:5,fontSize:12,fontWeight:600,fontFamily:fnt,borderRadius:6,cursor:"pointer",whiteSpace:"nowrap",border:"none",padding:"8px 14px",minHeight:36,color:fg||"#fff",background:bg||C.teal,transition:"all .12s"}),
  btnO:(fg,bdr)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:5,fontSize:12,fontWeight:600,fontFamily:fnt,borderRadius:6,cursor:"pointer",whiteSpace:"nowrap",padding:"8px 14px",minHeight:36,color:fg||C.t2,background:"transparent",border:`1px solid ${bdr||C.b2}`,transition:"all .12s"}),
  sm:{padding:"5px 10px",fontSize:11,minHeight:28},
  xs:{padding:"3px 8px",fontSize:10,minHeight:24},
  kpi:{background:C.white,border:`1px solid ${C.b1}`,borderRadius:8,padding:"11px 8px",textAlign:"center",flex:"1 1 75px",minWidth:68,boxShadow:"0 1px 3px rgba(28,55,64,0.07)"},
  th:{textAlign:"left",padding:"9px 11px",fontSize:11,fontWeight:700,color:C.t2,background:C.tl00,borderBottom:`1px solid ${C.b1}`,whiteSpace:"nowrap"},
  td:{padding:"9px 11px",fontSize:12,color:C.b6,borderBottom:`1px solid ${C.b1}`},
  av:(size,bgC,fgC)=>({width:size,height:size,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,flexShrink:0,fontFamily:fnt,fontSize:size*0.33,background:bgC||C.tl0,color:fgC||C.t2}),
  pill:{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",background:"rgba(255,255,255,0.1)",borderRadius:99,fontSize:10,color:"rgba(255,255,255,0.6)",whiteSpace:"nowrap",border:"1px solid rgba(255,255,255,0.08)"},
};

// ── COMPONENTS ──
const avColors=[{bg:C.tl0,fg:C.t2},{bg:C.gl,fg:C.g2},{bg:C.erb,fg:C.er},{bg:C.pub,fg:C.pu},{bg:C.okb,fg:C.ok},{bg:C.infb,fg:C.inf}];
function Avatar({name,size=30,colorIdx=0,isOut}){const c=isOut?{bg:C.erb,fg:C.er}:avColors[colorIdx%avColors.length];return<div style={ss.av(size,c.bg,c.fg)}>{initials(name)}</div>;}
function Badge({type="ne",children,dot=true}){const m={ok:{c:C.ok,b:C.okl},er:{c:C.er,b:C.erl},wn:{c:C.wn,b:C.wnl},ne:{c:C.b4,b:C.b1},pu:{c:C.pu,b:C.pul},"in":{c:C.inf,b:C.infl}}[type]||{c:C.b4,b:C.b1};return<span style={{display:"inline-flex",alignItems:"center",gap:3,padding:"2px 8px",fontSize:10,fontWeight:700,borderRadius:99,whiteSpace:"nowrap",color:m.c,background:m.b}}>{dot&&<span style={{width:5,height:5,borderRadius:"50%",background:"currentColor",flexShrink:0}}/>}{children}</span>;}
function KPI({label,value,color,sub}){return<div style={ss.kpi}><div style={{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:3}}>{label}</div><div style={{fontSize:22,fontWeight:700,fontFamily:mono,color:color||C.t2,lineHeight:1.1}}>{value}</div>{sub&&<div style={{fontSize:9,color:C.b4,marginTop:2}}>{sub}</div>}</div>;}
function CountBadge({n,bg,fg}){return n>0?<span style={{display:"inline-flex",alignItems:"center",justifyContent:"center",borderRadius:99,fontSize:9,fontWeight:700,minWidth:16,height:16,padding:"0 4px",background:bg,color:fg}}>{n}</span>:null;}
function Dot(){return<span style={{width:3,height:3,borderRadius:"50%",background:C.b3,flexShrink:0}}/>;}


// ══════════════════════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════════════════════
function App(){
  const{acct,token,login,refresh,err:authErr}=useMsal();
  const[tab,setTab]=useState(0);
  const[data,setData]=useState(null);
  const[loading,setLoading]=useState(false);
  const[error,setError]=useState(null);
  const[role,setRole]=useState(null);
  const[myEmail,setMyEmail]=useState(null);
  const[myEmp,setMyEmp]=useState(null);
  const[flash,setFlash]=useState("");
  const[timers,setTimers]=useState([]);
  const timersRef=useRef([]);
  useEffect(()=>{timersRef.current=timers;},[timers]);
  const[shift,setShift]=useState(null);
  const[covQ,setCovQ]=useState([]);
  const[queue,setQueue]=useState([]);
  const[tick,setTick]=useState(0);
  const[dfFrom,setDfFrom]=useState(()=>{const d=new Date();d.setDate(d.getDate()-7);return d.toISOString().slice(0,10);});
  const[dfTo,setDfTo]=useState(today());

  useEffect(()=>{const id=setInterval(()=>setTick(t=>t+1),1000);return()=>clearInterval(id);},[]);
  const fl=useCallback(msg=>{setFlash(msg);setTimeout(()=>setFlash(""),2500);},[]);
  async function gT(){return(await refresh())||token;}

  // ── Load Data ──
  useEffect(()=>{
    if(!token||!acct)return;
    const email=acct.username.toLowerCase();
    setMyEmail(email);setLoading(true);
    loadAll(token).then(d=>{
      setData(d);
      const me=d.employees.find(e=>(e.Email&&e.Email.toLowerCase()===email)||(e.M365UserId&&e.M365UserId.toLowerCase()===email)||(e.Email&&email.split("@")[0]===e.Email.toLowerCase().split("@")[0]));
      if(!me){setRole(null);setError("access_denied");}
      else{const r=detectRole(me);if(!r){setRole(null);setError("access_denied");}else{setRole(r);setError(null);setMyEmp(me);}}
      buildQueue(d,email,timersRef.current);
      setLoading(false);
    }).catch(e=>{setError("load_error: "+e.message);setLoading(false);});
  },[token,acct]);

  function buildQueue(d,email,currentTimers=[]){
    const timerSpIds=new Set(currentTimers.map(t=>t._spId).filter(Boolean));
    const persistedQueued=d.activities.filter(a=>a.ActivityType==="Task"&&a.Status==="Queued"&&!timerSpIds.has(a.id));
    const persistedInProgress=d.activities.filter(a=>a.ActivityType==="Task"&&a.Status==="In Progress"&&!timerSpIds.has(a.id));
    const q=[],cv=[];
    persistedQueued.forEach(a=>{
      const t={...a,_localId:a.id,_spId:a.id};
      if(a.Source==="Coverage"||a.CoverageForEmail){cv.push(t);}else{q.push(t);}
    });
    // Restore orphaned In Progress tasks as timers
    const restoredTimers=persistedInProgress.map(a=>({...a,_localId:a.id,_spId:a.id,_pMs:(a.PausedMin||0)*60000,_pS:null}));
    if(restoredTimers.length>0){
      setTimers(prev=>{const existingIds=new Set(prev.map(t=>t._spId));const fresh=restoredTimers.filter(t=>!existingIds.has(t._spId));if(!fresh.length)return prev;return[...prev,...fresh];});
    }
    // Generate recurring daily tasks
    if(d.config.recurringTasks){
      const td=today();
      const dow=new Date().getDay();
      if(dow===0||dow===6){setQueue(q);setCovQ(cv);return;} // Skip weekends
      const todayTasks=d.activities.filter(a=>a.ActivityType==="Task"&&a.ActivityDate&&a.ActivityDate.slice(0,10)===td);
      const existing=new Set(todayTasks.map(t=>`${t.VAEmail}|${t.Title}`));
      persistedQueued.forEach(a=>{existing.add(`${a.VAEmail}|${a.Title}`);});
      const toSave=[];
      d.config.recurringTasks.forEach(rt=>{
        if(!rt.active)return;
        const freq=rt.frequency||"daily";
        if(freq==="weekly"){const wd=rt.weekDay||1;if(dow!==wd)return;}
        const key=`${rt.vaEmail}|${rt.description}`;
        if(existing.has(key))return;
        const va=d.vas.find(v=>v.Email&&v.Email.toLowerCase()===rt.vaEmail.toLowerCase());
        if(!va)return;
        const prop=rt.propertyId?d.properties.find(p=>p.Title===rt.propertyId):null;
        const cat=d.config.categories?.find(c=>c.id===rt.category);
        const task={Title:rt.description,VAEmail:va.Email,VAName:va.Name,PropertyId:rt.propertyId||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Source:"Daily",Status:"Queued",Priority:"Normal",ActivityDate:td,ActivityType:"Task"};
        if(va.VATrackerStatus==="Out"){task.Source="Coverage";task.CoverageForEmail=va.Email;task.CoverageForName=va.Name;}
        toSave.push(task);
      });
      if(toSave.length>0){
        (async()=>{try{const tk=await gT();for(const task of toSave){const res=await gPost(tk,lUrl("VA_Activity"),{Title:task.Title,ActivityType:"Task",VAEmail:task.VAEmail,VAName:task.VAName,ActivityDate:task.ActivityDate,PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category,Source:task.Source,Status:task.Status,Priority:task.Priority||"Normal",CoverageForEmail:task.CoverageForEmail||"",CoverageForName:task.CoverageForName||""});
          const saved={...task,_localId:res.id,_spId:res.id,id:res.id};
          if(task.Source==="Coverage"){setCovQ(p=>[...p,saved]);}else{setQueue(p=>[...p,saved]);}
        }}catch(e){console.error("[VT] Daily task gen error:",e);}})();
      }
    }
    setQueue(q);setCovQ(cv);
  }

  async function reload(){const t=await gT();const d=await loadAll(t);setData(d);buildQueue(d,myEmail,timersRef.current);return d;}

  // ── Task CRUD ──
  async function saveTask(task){
    const t=await gT();
    return gPost(t,lUrl("VA_Activity"),{Title:task.Title,ActivityType:"Task",VAEmail:task.VAEmail,VAName:task.VAName,ActivityDate:task.ActivityDate||new Date().toISOString(),PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category,Source:task.Source,Status:task.Status,Priority:task.Priority||"Normal",StartTime:task.StartTime||null,EndTime:task.EndTime||null,DurationMin:task.DurationMin||0,PausedMin:task.PausedMin||0,Notes:task.Notes||"",CoverageForEmail:task.CoverageForEmail||"",CoverageForName:task.CoverageForName||"",AssignedByEmail:task.AssignedByEmail||"",AssignedByName:task.AssignedByName||"",GroupId:task.GroupId||""});
  }
  async function addTask(task){
    const t2={...task,Status:"Queued",ActivityDate:today(),ActivityType:"Task"};
    const targetVa=data?.vas.find(v=>v.Email?.toLowerCase()===task.VAEmail?.toLowerCase());
    if(targetVa&&targetVa.VATrackerStatus==="Out"){t2.Source="Coverage";t2.CoverageForEmail=targetVa.Email;t2.CoverageForName=targetVa.Name;}
    try{const tk=await gT();const res=await gPost(tk,lUrl("VA_Activity"),{Title:t2.Title,ActivityType:"Task",VAEmail:t2.VAEmail,VAName:t2.VAName,ActivityDate:t2.ActivityDate,PropertyId:t2.PropertyId||"",PropertyName:t2.PropertyName||"General",PMName:t2.PMName||"",Category:t2.Category,Source:t2.Source,Status:t2.Status,Priority:t2.Priority||"Normal",Notes:t2.Notes||"",CoverageForEmail:t2.CoverageForEmail||"",CoverageForName:t2.CoverageForName||"",AssignedByEmail:t2.AssignedByEmail||"",AssignedByName:t2.AssignedByName||"",GroupId:t2.GroupId||""});
      const saved={...t2,_localId:res.id,_spId:res.id,id:res.id};
      if(t2.Source==="Coverage"){setCovQ(p=>[saved,...p]);fl(`${targetVa.Name} is OUT — task sent to coverage`);}
      else{setQueue(p=>[saved,...p]);fl("Task added!");}
    }catch(e){fl("Error: "+e.message);}
  }
  async function deleteTask(task){
    if(task._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",task._spId),{Status:"Incomplete",Notes:"Removed by admin"});}catch(e){fl("Error: "+e.message);return;}}
    setQueue(p=>p.filter(t=>t._localId!==task._localId));setCovQ(p=>p.filter(t=>t._localId!==task._localId));fl("Task removed");
  }

  // ── Timer actions ──
  async function startTimer(task){
    const timer={...task,Status:"In Progress",StartTime:new Date().toISOString(),_pMs:0,_pS:null};
    if(task._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",task._spId),{Status:"In Progress",StartTime:new Date().toISOString()});}catch(e){fl("Could not start timer — try again.");return;}}
    setTimers(p=>[...p,timer]);setQueue(p=>p.filter(t=>t._localId!==task._localId));fl("Timer started!");
  }
  function pauseTimer(id){setTimers(p=>p.map(t=>t._localId===id?{...t,_pS:Date.now()}:t));}
  function resumeTimer(id){setTimers(p=>p.map(t=>{if(t._localId!==id||!t._pS)return t;return{...t,_pMs:(t._pMs||0)+(Date.now()-t._pS),_pS:null};}));}
  async function finishTimer(id,status,notes){
    const t=timers.find(x=>x._localId===id);if(!t)return;
    const now=Date.now();let pMs=t._pMs||0;if(t._pS)pMs+=(now-t._pS);
    const dur=Math.max(1,Math.round((now-new Date(t.StartTime).getTime()-pMs)/6e4));
    const fields={Status:status,EndTime:status==="Completed"?new Date(now).toISOString():null,DurationMin:dur,PausedMin:Math.round(pMs/6e4),Notes:notes||""};
    try{const tk=await gT();if(t._spId){await gPatch(tk,iUrl("VA_Activity",t._spId),fields);}else{await saveTask({...t,...fields});}
      setTimers(p=>p.filter(x=>x._localId!==id));fl(status==="Completed"?"Task completed!":"Task: "+status);
      await new Promise(r=>setTimeout(r,1500));await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function cancelTimer(id){
    const t=timers.find(x=>x._localId===id);if(!t)return;
    if(t._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",t._spId),{Status:"Queued",StartTime:null});}catch(e){}}
    setTimers(p=>p.filter(x=>x._localId!==id));setQueue(p=>[{...t,Status:"Queued",StartTime:null,_pMs:0,_pS:null},...p]);
  }

  // ── Shift clock ──
  function clockIn(){if(shift){fl("Already clocked in!");return;}setShift({ClockIn:new Date().toISOString(),Breaks:[],_ob:false,_bs:null});fl("Clocked in!");}
  function startBreak(){setShift(p=>p?{...p,_ob:true,_bs:new Date().toISOString()}:p);}
  function endBreak(){setShift(p=>{if(!p||!p._bs)return p;return{...p,_ob:false,Breaks:[...p.Breaks,{s:p._bs,e:new Date().toISOString()}],_bs:null};});}
  async function clockOut(){
    if(!shift)return;const now=new Date();const bks=[...shift.Breaks];
    if(shift._ob&&shift._bs)bks.push({s:shift._bs,e:now.toISOString()});
    const bMs=bks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);
    const bMin=Math.round(bMs/6e4);const wMin=Math.round((now-new Date(shift.ClockIn)-bMs)/6e4);
    try{const t=await gT();await gPost(t,lUrl("VA_Activity"),{Title:`${myEmp?.Name||myEmail}-${today()}`,ActivityType:"Shift",VAEmail:myEmail,VAName:myEmp?.Name||myEmail,ActivityDate:shift.ClockIn,StartTime:shift.ClockIn,EndTime:now.toISOString(),BreakMinutes:bMin,WorkMinutes:wMin,BreaksJSON:JSON.stringify(bks)});
      setShift(null);fl("Clocked out!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Absence ──
  async function toggleAbsence(va){
    const t=await gT();const ns=va.VATrackerStatus==="Out"?"Active":"Out";
    await gPatch(t,iUrl("Employees",va.id),{VATrackerStatus:ns});
    if(ns==="Out"){
      await gPost(t,lUrl("VA_Activity"),{Title:`${va.Name}-Out-${today()}`,ActivityType:"Absence",VAEmail:va.Email,VAName:va.Name,ActivityDate:new Date().toISOString(),StartTime:new Date().toISOString(),Status:"Out",MarkedByEmail:myEmail,MarkedByName:acct.name||myEmail});
      const mv=queue.filter(q=>q.VAEmail.toLowerCase()===va.Email.toLowerCase());
      for(const task of mv){if(task._spId){try{await gPatch(t,iUrl("VA_Activity",task._spId),{Source:"Coverage",CoverageForEmail:va.Email,CoverageForName:va.Name});}catch(e){}}}
      const rest=queue.filter(q=>q.VAEmail.toLowerCase()!==va.Email.toLowerCase());
      mv.forEach(q=>{q.Source="Coverage";q.CoverageForEmail=va.Email;q.CoverageForName=va.Name;});
      setQueue(rest);setCovQ(p=>[...p,...mv]);fl(`${va.Name} marked OUT — ${mv.length} tasks to coverage`);
    }else{
      const returning=covQ.filter(q=>q.CoverageForEmail?.toLowerCase()===va.Email.toLowerCase());
      for(const task of returning){if(task._spId){try{await gPatch(t,iUrl("VA_Activity",task._spId),{Source:"Daily",VAEmail:va.Email,VAName:va.Name,CoverageForEmail:"",CoverageForName:""});}catch(e){}}}
      returning.forEach(q=>{q.Source="Daily";q.CoverageForEmail="";q.CoverageForName="";q.VAEmail=va.Email;q.VAName=va.Name;});
      setCovQ(p=>p.filter(q=>q.CoverageForEmail?.toLowerCase()!==va.Email.toLowerCase()));
      setQueue(p=>[...p,...returning]);fl(`${va.Name} marked IN — ${returning.length} tasks returned`);
    }
    await reload();
  }

  // ── Coverage ──
  async function claimCov(id){
    const t2=covQ.find(x=>x._localId===id);if(!t2)return;
    if(t2._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",t2._spId),{VAEmail:myEmail,VAName:myEmp?.Name||myEmail});}catch(e){}}
    setCovQ(p=>p.filter(x=>x._localId!==id));setQueue(p=>[{...t2,VAEmail:myEmail,VAName:myEmp?.Name||myEmail},...p]);fl(`Claimed! Covering for ${t2.CoverageForName}`);
  }

  // ── Config ──
  async function updateConfig(nc){
    try{const t=await gT();const ci=await gAll(t,`${lUrl("VA_TrackerConfig")}?expand=fields&$top=10`);
      const item=ci.find(c=>c.fields.Title==="VATrackerSettings");if(!item){fl("Config not found");return;}
      await gPatch(t,iUrl("VA_TrackerConfig",item.id),{ConfigJSON:JSON.stringify(nc)});await reload();fl("Config saved!");
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Portfolio ──
  async function assignProp(vaEmail,vaName,propId){
    try{const t=await gT();const prop=data.properties.find(p=>p.Title===propId);if(!prop)return;
      const ex=data.portfolios.find(p=>p.VAEmail?.toLowerCase()===vaEmail.toLowerCase()&&p.PropertyId===propId);
      if(ex)await gPatch(t,iUrl("VA_Portfolios",ex.id),{IsActive:true});
      else await gPost(t,lUrl("VA_Portfolios"),{Title:`${vaName.split(" ")[0]}-${prop.PropertyName.replace(/\s+/g,"")}`.slice(0,50),VAEmail:vaEmail,VAName:vaName,PropertyId:propId,PropertyName:prop.PropertyName,AssignedDate:new Date().toISOString(),IsActive:true});
      await reload();fl(`${prop.PropertyName} → ${vaName}`);
    }catch(e){fl("Error: "+e.message);}
  }
  async function unassignProp(portId,propName,vaName){
    try{const t=await gT();await gPatch(t,iUrl("VA_Portfolios",portId),{IsActive:false});await reload();fl(`${propName} removed from ${vaName}`);}catch(e){fl("Error: "+e.message);}
  }
  async function addProperty(p){
    try{const t=await gT();await gPost(t,lUrl("VA_Properties"),{Title:`PROP-${String(data.properties.length+1).padStart(3,"0")}`,PropertyName:p.name,PropertyGroup:p.group,Units:p.units,PMEmail:p.pmEmail,PMName:p.pmName,AppFolioId:p.appFolioId||"",IsActive:true});
      await reload();fl(`${p.name} added!`);
    }catch(e){fl("Error: "+e.message);}
  }
  async function editProperty(id,fields){try{const t=await gT();await gPatch(t,iUrl("VA_Properties",id),fields);await reload();fl("Updated!");}catch(e){fl("Error: "+e.message);}}
  async function editEmployee(id,fields){try{const t=await gT();await gPatch(t,iUrl("Employees",id),fields);await reload();fl("Employee updated!");}catch(e){fl("Error: "+e.message);}}

  // ── Full VA reassignment cascade for a property ──
  async function reassignPropertyVA(propId,oldVaEmail,newVaEmail,newVaName){
    try{
      const t=await gT();
      const prop=data.properties.find(p=>p.Title===propId);
      const propName=prop?.PropertyName||propId;
      let moved=0;

      // 1. Portfolio: unassign old, assign new
      const oldPort=data.portfolios.find(p=>p.PropertyId===propId&&p.VAEmail?.toLowerCase()===oldVaEmail?.toLowerCase());
      if(oldPort)await gPatch(t,iUrl("VA_Portfolios",oldPort.id),{IsActive:false});
      if(newVaEmail){
        const exNew=data.portfolios.find(p=>p.PropertyId===propId&&p.VAEmail?.toLowerCase()===newVaEmail.toLowerCase());
        if(exNew)await gPatch(t,iUrl("VA_Portfolios",exNew.id),{IsActive:true,VAEmail:newVaEmail,VAName:newVaName});
        else await gPost(t,lUrl("VA_Portfolios"),{Title:`${newVaName.split(" ")[0]}-${propName.replace(/\s+/g,"")}`.slice(0,50),VAEmail:newVaEmail,VAName:newVaName,PropertyId:propId,PropertyName:propName,AssignedDate:new Date().toISOString(),IsActive:true});
      }

      // 2. Recurring tasks: update config JSON — change vaEmail on matching propertyId
      if(data.config.recurringTasks&&newVaEmail){
        const updated=data.config.recurringTasks.map(rt=>{
          if(rt.propertyId===propId&&rt.vaEmail?.toLowerCase()===oldVaEmail?.toLowerCase()){
            return{...rt,vaEmail:newVaEmail};
          }
          return rt;
        });
        const changed=updated.some((rt,i)=>rt.vaEmail!==data.config.recurringTasks[i].vaEmail);
        if(changed){
          const ci=await gAll(t,`${lUrl("VA_TrackerConfig")}?expand=fields&$top=10`);
          const item=ci.find(c=>c.fields.Title==="VATrackerSettings");
          if(item){await gPatch(t,iUrl("VA_TrackerConfig",item.id),{ConfigJSON:JSON.stringify({...data.config,recurringTasks:updated})});}
        }
      }

      // 3. Queued tasks in SharePoint: PATCH all Queued tasks for this property from old VA to new VA
      if(newVaEmail&&oldVaEmail){
        const pendingTasks=data.activities.filter(a=>a.ActivityType==="Task"&&(a.Status==="Queued"||a.Status==="In Progress")&&a.PropertyId===propId&&a.VAEmail?.toLowerCase()===oldVaEmail.toLowerCase());
        for(const task of pendingTasks){
          try{await gPatch(t,iUrl("VA_Activity",task.id),{VAEmail:newVaEmail,VAName:newVaName});moved++;}catch(e){console.warn("[VT] Task reassign failed:",task.id,e);}
        }
        // Update local queue state
        setQueue(prev=>prev.map(q=>{
          if(q.PropertyId===propId&&q.VAEmail?.toLowerCase()===oldVaEmail.toLowerCase()){return{...q,VAEmail:newVaEmail,VAName:newVaName};}
          return q;
        }));
        setCovQ(prev=>prev.map(q=>{
          if(q.PropertyId===propId&&q.VAEmail?.toLowerCase()===oldVaEmail.toLowerCase()){return{...q,VAEmail:newVaEmail,VAName:newVaName};}
          return q;
        }));
      }

      await reload();
      fl(`${propName} → ${newVaName||"unassigned"}${moved>0?` · ${moved} pending tasks moved`:""}`);
    }catch(e){fl("Error reassigning: "+e.message);}
  }

  // ── Helpers ──
  function getVAForProperty(propId){const port=data?.portfolios.find(p=>p.PropertyId===propId);if(!port)return null;return data?.vas.find(v=>v.Email?.toLowerCase()===port.VAEmail?.toLowerCase())||null;}
  function myManagedProps(){if(!data||!myEmail)return[];return data.properties.filter(p=>p.PMEmail?.toLowerCase()===myEmail);}

  // ── Computed ──
  const isAdmin=role==="admin";
  const isRegional=role==="regional";
  const isMgr=isAdmin||isRegional||role==="manager";
  const isVA=role==="va";
  const myVa=data?.vas.find(v=>v.Email&&v.Email.toLowerCase()===myEmail);
  const myPort=data?data.portfolios.filter(p=>p.VAEmail&&p.VAEmail.toLowerCase()===myEmail):[];
  const myProps=data?myPort.map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean):[];
  const mgrProps=myManagedProps();
  const myQ=queue.filter(t=>isVA?t.VAEmail?.toLowerCase()===myEmail:true);
  const myTm=timers.filter(t=>isVA?t.VAEmail?.toLowerCase()===myEmail:true);
  const outVAs=data?data.vas.filter(v=>v.VATrackerStatus==="Out"):[];
  // Overdue: tasks from previous days still queued
  const overdueTasks=data?data.activities.filter(a=>a.ActivityType==="Task"&&(a.Status==="Queued"||a.Status==="In Progress")&&a.ActivityDate&&a.ActivityDate.slice(0,10)<today()):[];
  const myOverdue=overdueTasks.filter(t=>isVA?t.VAEmail?.toLowerCase()===myEmail:true);

  // ── Auth Screens ──
  if(!acct)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700,letterSpacing:"0.04em"}}>{CONFIG.appName}</div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,letterSpacing:"0.08em",textTransform:"uppercase"}}>NewShire Property Management</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{...ss.card,textAlign:"center",maxWidth:360,padding:32}}><div style={{fontSize:36,marginBottom:12}}>⏱</div><div style={{fontSize:18,fontWeight:700,color:C.t2,marginBottom:6}}>VA Productivity Tracker</div><div style={{color:C.b4,marginBottom:20,fontSize:12}}>Sign in with your NewShire account.</div><button style={ss.btn(C.teal)} onClick={login}>Sign In with Microsoft</button>{authErr&&<div style={{color:C.er,marginTop:12,fontSize:11}}>{authErr}</div>}</div></div></div>);
  if(loading)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>{CONFIG.appName}</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{fontSize:16,color:C.b4}}>Loading...</div></div></div>);
  if(error||!role)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>{CONFIG.appName}</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{...ss.card,borderTop:`3px solid ${C.er}`,maxWidth:400,textAlign:"center",padding:32}}><div style={{fontSize:36,marginBottom:12}}>🚫</div><div style={{fontSize:16,fontWeight:700,color:C.er,marginBottom:6}}>{error==="access_denied"?"Access Denied":"Error"}</div><div style={{color:C.b4,fontSize:12}}>{error==="access_denied"?"Your account is not authorized.":error}</div><div style={{color:C.b3,fontSize:10,marginTop:12}}>Signed in as: {myEmail}</div></div></div></div>);

  // ── Tabs ──
  const TABS=[];
  if(isVA||isAdmin)TABS.push({n:"My Day",k:"myday"});
  if(isMgr)TABS.push({n:"Manager View",k:"mgr",badge:myOverdue.length>0?myOverdue.length:0,badgeBg:C.er});
  TABS.push({n:"Dashboard",k:"dash"});
  if(isAdmin||isRegional)TABS.push({n:"Coaching",k:"coach"});
  TABS.push({n:"History",k:"hist"});
  if(isAdmin)TABS.push({n:"Admin",k:"admin",badge:covQ.length+outVAs.length,badgeBg:C.t3});
  const ck=TABS[tab]?.k||TABS[0]?.k;

  // ── Header pills ──
  const pills=[];
  if(timers.length>0)pills.push({c:C.ok,n:timers.length,l:"timing"});
  if(covQ.length>0)pills.push({c:C.wn,n:covQ.length,l:"coverage"});
  if(myOverdue.length>0)pills.push({c:C.er,n:myOverdue.length,l:"overdue"});
  if(outVAs.length>0)pills.push({c:C.er,n:outVAs.length,l:"out"});

  return(
    <div style={ss.app}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@500;600&display=swap" rel="stylesheet"/>
      {/* Header */}
      <div style={ss.hdr}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={ss.logo}>V</div>
          <div><div style={{color:"#fff",fontSize:14,fontWeight:700,letterSpacing:"0.04em"}}>{CONFIG.appName}</div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,letterSpacing:"0.08em",textTransform:"uppercase"}}>NewShire Property Management</div></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,flex:1,justifyContent:"center",flexWrap:"wrap"}}>
          {pills.map((p,i)=><div key={i} style={ss.pill}><span style={{width:5,height:5,borderRadius:"50%",background:p.c}}/><strong style={{color:"#fff",fontWeight:700}}>{p.n}</strong>&nbsp;{p.l}</div>)}
        </div>
        <div style={{textAlign:"right",flexShrink:0}}><strong style={{display:"block",fontSize:12,fontWeight:600,color:"#fff"}}>{acct.name||myEmail}</strong><em style={{fontSize:9,fontStyle:"normal",color:C.gold,textTransform:"uppercase",letterSpacing:"0.07em"}}>{role}</em></div>
      </div>
      {/* Tabs */}
      <div style={ss.tabs}>
        {TABS.map((t,i)=><button key={t.k} style={ss.tab(tab===i)} onClick={()=>setTab(i)}>{t.n}{t.badge>0&&<CountBadge n={t.badge} bg={t.badgeBg||C.er} fg="#fff"/>}</button>)}
      </div>
      {flash&&<div style={{background:C.gl0,borderBottom:`1px solid ${C.gold}`,padding:"7px 18px",fontSize:12,fontWeight:600,color:C.g2,textAlign:"center"}}>{flash}</div>}
      {/* Content */}
      <div style={ss.content}>
        {ck==="myday"&&<MyDayView data={data} role={role} myEmail={myEmail} myVa={myVa} myProps={myProps} queue={queue} covQ={covQ} shift={shift} timers={timers} tick={tick} overdue={myOverdue} config={data?.config} onClockIn={clockIn} onBreakStart={startBreak} onBreakEnd={endBreak} onClockOut={clockOut} onStartTimer={startTimer} onPause={pauseTimer} onResume={resumeTimer} onFinish={finishTimer} onCancel={cancelTimer} onClaimCov={claimCov} onAddTask={addTask} onDeleteTask={deleteTask} isAdmin={isAdmin} fl={fl}/>}
        {ck==="mgr"&&<ManagerView data={data} myEmail={myEmail} myEmp={myEmp} mgrProps={isAdmin?data.properties:mgrProps} queue={queue} timers={timers} covQ={covQ} overdue={overdueTasks} onAddTask={addTask} getVA={getVAForProperty} isAdmin={isAdmin} isRegional={isRegional}/>}
        {ck==="dash"&&<DashboardView data={data} queue={queue} timers={timers} covQ={covQ} overdue={overdueTasks} dfFrom={dfFrom} dfTo={dfTo} setDfFrom={setDfFrom} setDfTo={setDfTo} isAdmin={isAdmin} role={role} mgrProps={mgrProps}/>}
        {ck==="coach"&&<CoachingView data={data}/>}
        {ck==="hist"&&<HistoryView data={data} role={role} myEmail={myEmail} isMgr={isMgr} mgrProps={mgrProps}/>}
        {ck==="admin"&&<AdminView data={data} myEmail={myEmail} acct={acct} config={data?.config} queue={queue} covQ={covQ} onToggleAbsence={toggleAbsence} onAssignTask={addTask} onUpdateConfig={updateConfig} onAssignProp={assignProp} onUnassignProp={unassignProp} onReassignVA={reassignPropertyVA} onAddProperty={addProperty} onEditProperty={editProperty} onEditEmployee={editEmployee} onDeleteTask={deleteTask}/>}
      </div>
    </div>
  );
}


// ══════════════════════════════════════════════════════
// MY DAY VIEW
// ══════════════════════════════════════════════════════
function MyDayView({data,role,myEmail,myVa,myProps,queue,covQ,shift,timers,tick,overdue,config,onClockIn,onBreakStart,onBreakEnd,onClockOut,onStartTimer,onPause,onResume,onFinish,onCancel,onClaimCov,onAddTask,onDeleteTask,isAdmin,fl}){
  const[showForm,setShowForm]=useState(false);const[fCat,setFCat]=useState("");const[fProp,setFProp]=useState("");const[fPri,setFPri]=useState("Normal");const[fDesc,setFDesc]=useState("");
  const cats=config?.categories||[];
  const portProps=data?data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===myEmail).map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean):[];
  const myTasks=isAdmin?queue:queue.filter(t=>t.VAEmail?.toLowerCase()===myEmail);
  const myTimers=isAdmin?timers:timers.filter(t=>t.VAEmail?.toLowerCase()===myEmail);
  const myOverdue=isAdmin?overdue:overdue.filter(t=>t.VAEmail?.toLowerCase()===myEmail);

  let shE=0,bkE=0;
  if(shift){const now=Date.now();let bMs=shift.Breaks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);if(shift._ob&&shift._bs)bMs+=(now-new Date(shift._bs).getTime());bkE=Math.floor(bMs/1000);shE=Math.floor((now-new Date(shift.ClockIn).getTime()-bMs)/1000);}

  function handleAdd(){if(!fDesc||!fCat)return;const cat=cats.find(c=>c.id===fCat);const prop=fProp?data.properties.find(p=>p.Title===fProp):null;
    onAddTask({Title:fDesc,VAEmail:myEmail,VAName:myVa?.Name||myEmail,PropertyId:fProp||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Priority:fPri,Source:"Ad-Hoc"});
    setFDesc("");setFCat("");setFProp("");setFPri("Normal");setShowForm(false);}

  return(<div style={{maxWidth:700,margin:"0 auto"}}>
    {/* Shift Clock */}
    <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:8,padding:14,marginBottom:12,boxShadow:"0 1px 3px rgba(28,55,64,0.07)"}}>
      {!shift?(<div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:11,fontWeight:700,color:C.b4,textTransform:"uppercase"}}>Shift Clock</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Not clocked in</div></div><button style={ss.btn(C.inf)} onClick={onClockIn}>☀ Clock In</button></div>
      ):(<div>
        <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:11}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:shift._ob?C.wn:C.ok,position:"relative"}}/>
          <div style={{flex:1}}><div style={{fontSize:11,fontWeight:700,color:shift._ob?C.wn:C.ok,textTransform:"uppercase",letterSpacing:"0.05em"}}>{shift._ob?"On Break":"Clocked In"}</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Since {fT(shift.ClockIn)} · {shift.Breaks.length+(shift._ob?1:0)} breaks</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase",marginBottom:1}}>Working Time</div><div style={{fontSize:26,fontWeight:700,fontFamily:mono,color:shift._ob?C.wn:C.t2}}>{fTm(shE)}</div></div>
        </div>
        <div style={{display:"flex",gap:7}}>
          {shift._ob?<button style={{...ss.btn(C.ok),flex:1}} onClick={onBreakEnd}>▶ End Break</button>:<button style={{...ss.btnO(C.wn,C.wn),flex:1}} onClick={onBreakStart}>☕ Break</button>}
          <button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),flex:1}} onClick={onClockOut}>🌙 Clock Out</button>
        </div></div>)}
    </div>

    {/* Active Timers — inline */}
    {myTimers.map(t=>{const now=Date.now(),st=new Date(t.StartTime).getTime();let pMs=t._pMs||0;const ip=!!t._pS;if(ip)pMs+=(now-t._pS);const el=ip?Math.floor((t._pS-st-(t._pMs||0))/1000):Math.floor((now-st-pMs)/1000);
      return(<div key={t._localId} style={{...ss.card,borderLeft:`4px solid ${ip?C.wn:C.ok}`,background:ip?C.wnb:C.okb,padding:0,overflow:"hidden"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px"}}>
          <div style={{width:7,height:7,borderRadius:"50%",background:ip?C.wn:C.ok,flexShrink:0}}/> 
          <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:C.t2}}>{t.Title}</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{t.PropertyName} · {t.Category}{isAdmin&&t.VAName?` · ${t.VAName}`:""}</div></div>
          <div style={{fontSize:20,fontWeight:700,fontFamily:mono,color:ip?C.wn:C.ok}}>{fTm(el)}</div>
        </div>
        <div style={{display:"flex",gap:5,padding:"0 14px 10px"}}>
          {ip?<button style={{...ss.btn(C.ok),...ss.sm,flex:1}} onClick={()=>onResume(t._localId)}>▶ Resume</button>:<button style={{...ss.btnO(C.wn,C.wn),...ss.sm,flex:1}} onClick={()=>onPause(t._localId)}>⏸ Pause</button>}
          <button style={{...ss.btn(C.ok),...ss.sm,flex:1}} onClick={()=>{const n=prompt("Notes (optional):");onFinish(t._localId,"Completed",n||"");}}>✓ Done</button>
          <button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),...ss.sm}} onClick={()=>{const n=prompt("What's blocking?");if(n)onFinish(t._localId,"Blocked",n);}}>⚠</button>
          <button style={{...ss.btnO(C.b4,C.b2),...ss.sm}} onClick={()=>onCancel(t._localId)}>↩</button>
        </div>
      </div>);})}

    {/* Overdue */}
    {myOverdue.length>0&&<div style={{background:C.erb,border:`1px solid rgba(184,59,42,0.18)`,borderLeft:`4px solid ${C.er}`,borderRadius:8,padding:13,marginBottom:12}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:9}}>
        <div style={{width:22,height:22,background:C.er,borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#fff",fontWeight:700}}>!</div>
        <div><div style={{fontSize:12,fontWeight:700,color:C.er}}>{myOverdue.length} Overdue Tasks from Previous Days</div><div style={{fontSize:10,color:"rgba(184,59,42,0.7)",marginTop:1}}>Start, mark incomplete, or flag for review</div></div>
      </div>
      {myOverdue.slice(0,5).map(t=><TaskRow key={t.id} task={t} onStart={()=>onStartTimer({...t,_localId:t.id,_spId:t.id})} showVA={isAdmin} isOverdue/>)}
    </div>}

    {/* Coverage */}
    {covQ.length>0&&<div style={{...ss.card,borderLeft:`4px solid ${C.wn}`,background:C.wnb}}>
      <div style={{fontSize:13,fontWeight:700,color:C.wn,marginBottom:8}}>🚨 Coverage Needed ({covQ.length})</div>
      {covQ.map(t=><div key={t._localId} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid rgba(168,111,8,0.1)`}}>
        <span>{catIcon[t.Category]||"📁"}</span><div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{t.Title}</div><div style={{fontSize:10,color:C.b4}}>{t.PropertyName} · covering {t.CoverageForName}</div></div>
        <button style={{...ss.btn(C.wn),...ss.xs}} onClick={()=>onClaimCov(t._localId)}>✋ Claim</button></div>)}
    </div>}

    {/* Daily Tasks */}
    <div style={ss.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div><div style={ss.cardT}>📋 {isAdmin?"All Tasks":"Daily Tasks"}</div><div style={ss.cardS}>Today — {new Date().toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric",year:"numeric"})}</div></div>
        <Badge type="ne" dot={false}>{myTasks.length} remaining</Badge>
      </div>
      {myTasks.length===0&&<div style={{textAlign:"center",padding:"20px 0",color:C.b4,fontSize:12}}>All done! 🎉</div>}
      {myTasks.map(t=><TaskRow key={t._localId} task={t} onStart={()=>onStartTimer(t)} onDelete={isAdmin?onDeleteTask:null} showVA={isAdmin}/>)}
    </div>

    {/* Add Task */}
    <div style={ss.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={ss.cardT}>➕ Add Task</div><button style={ss.btnO(C.t2,C.b2)} onClick={()=>setShowForm(!showForm)}>{showForm?"Cancel":"New"}</button></div>
      {showForm&&<div style={{marginTop:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
          <div style={{flex:1,minWidth:130}}><label style={ss.label}>Category *</label><select style={ss.select} value={fCat} onChange={e=>setFCat(e.target.value)}><option value="">...</option>{(config?.categories||[]).sort((a,b)=>a.name==="Admin/Other"?1:b.name==="Admin/Other"?-1:a.name.localeCompare(b.name)).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div style={{flex:1,minWidth:130}}><label style={ss.label}>Property</label><select style={ss.select} value={fProp} onChange={e=>setFProp(e.target.value)}><option value="">General</option>{portProps.map(p=><option key={p.Title} value={p.Title}>{p.PropertyName}</option>)}</select></div>
          <div style={{minWidth:80}}><label style={ss.label}>Priority</label><select style={ss.select} value={fPri} onChange={e=>setFPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div></div>
        <label style={ss.label}>Task *</label><input style={{...ss.input,marginBottom:10}} value={fDesc} onChange={e=>setFDesc(e.target.value)} placeholder="Describe..."/>
        <button style={{...ss.btn(C.teal),width:"100%"}} onClick={handleAdd}>+ Queue</button>
      </div>}
    </div>
  </div>);
}

// ── Task Row ──
function TaskRow({task,onStart,onDelete,showVA,isOverdue}){
  return(<div style={{display:"flex",alignItems:"flex-start",gap:9,padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
    <span style={{fontSize:15,width:22,textAlign:"center",flexShrink:0,paddingTop:1}}>{catIcon[task.Category]||"📁"}</span>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontSize:12,fontWeight:600,color:C.t2,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap",lineHeight:1.4}}>{task.Title}
        {isOverdue&&<Badge type="er" dot={false}>OVERDUE — {fD(task.ActivityDate)}</Badge>}
        {task.CoverageForName&&<Badge type="wn" dot={false}>Coverage</Badge>}
        {task.Priority==="Urgent"&&<Badge type="er" dot={false}>Urgent</Badge>}
        {task.Priority==="High"&&<Badge type="wn" dot={false}>High</Badge>}
      </div>
      <div style={{fontSize:10,color:C.b4,marginTop:3,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
        {showVA&&task.VAName&&<>{task.VAName}<Dot/></>}
        {task.PropertyName}<Dot/>{task.Source}
        {task.Notes&&<><Dot/>💬 {task.Notes}</>}
      </div>
    </div>
    <div style={{display:"flex",gap:4,flexShrink:0,flexWrap:"wrap",alignItems:"flex-start"}}>
      {onStart&&<button style={{...ss.btn(C.ok),...ss.xs}} onClick={onStart}>▶ Start</button>}
      {onDelete&&<button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),...ss.xs}} onClick={()=>{if(window.confirm("Remove?"))onDelete(task);}}>✕</button>}
    </div>
  </div>);
}

// ══════════════════════════════════════════════════════
// MANAGER VIEW
// ══════════════════════════════════════════════════════
function ManagerView({data,myEmail,myEmp,mgrProps,queue,timers,covQ,overdue,onAddTask,getVA,isAdmin,isRegional}){
  const[selProp,setSelProp]=useState("");const[tCat,setTCat]=useState("");const[tDesc,setTDesc]=useState("");const[tPri,setTPri]=useState("Normal");const[tNotes,setTNotes]=useState("");
  const cats=data?.config?.categories||[];
  function handleSubmit(){if(!selProp||!tCat||!tDesc)return;const prop=data.properties.find(p=>p.Title===selProp);const va=getVA(selProp);if(!va){alert("No VA assigned.");return;}const cat=cats.find(c=>c.id===tCat);
    onAddTask({Title:tDesc,VAEmail:va.Email,VAName:va.Name,PropertyId:selProp,PropertyName:prop?.PropertyName||"",PMName:myEmp?.Name||myEmail,Category:cat?.name||"Admin/Other",Priority:tPri,Source:"Assigned",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||myEmail,Notes:tNotes});setTDesc("");setTCat("");setTPri("Normal");setTNotes("");}

  // Active timers on my properties
  const propIds=new Set(mgrProps.map(p=>p.Title));
  const activeOnMine=timers.filter(t=>propIds.has(t.PropertyId));
  const overdueOnMine=overdue.filter(t=>propIds.has(t.PropertyId));

  return(<div>
    {/* Live Active Tasks */}
    {activeOnMine.length>0&&<div style={{...ss.card,borderTop:`3px solid ${C.ok}`}}>
      <div style={ss.cardT}>🟢 Currently Active on Your Properties</div><div style={{...ss.cardS,marginBottom:10}}>Live · {activeOnMine.length} tasks in progress</div>
      {activeOnMine.map((t,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:C.okb,border:`1px solid rgba(26,122,70,0.18)`,borderRadius:6,marginBottom:7}}>
        <div style={{width:7,height:7,borderRadius:"50%",background:C.ok,flexShrink:0}}/>
        <Avatar name={t.VAName} size={24}/>
        <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:C.t2}}>{t.VAName} — {t.Title}</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{t.PropertyName} · {t.Category} · Started {fT(t.StartTime)}</div></div>
      </div>)}
    </div>}

    {/* Portfolio Overview (Regional only) */}
    {isRegional&&<div style={{...ss.card,borderTop:`3px solid ${C.inf}`,marginBottom:12}}>
      <div style={{...ss.cardT,color:C.inf}}>📊 Portfolio Overview — Regional View</div>
      <div style={{...ss.cardS,marginBottom:12}}>PM-by-PM rollup across all managed properties</div>
      {data.pms.filter(pm=>pm.JobTitle==="Property Manager").map((pm,pi)=>{
        const pmProps=data.properties.filter(p=>p.PMName===pm.Name||p.PMEmail?.toLowerCase()===pm.Email?.toLowerCase());
        if(!pmProps.length)return null;
        const pmPropIds=new Set(pmProps.map(p=>p.Title));
        const pmOD=overdue.filter(t=>pmPropIds.has(t.PropertyId));
        const pmTasks7=data.activities.filter(a=>a.ActivityType==="Task"&&pmPropIds.has(a.PropertyId)&&dAgo(a.ActivityDate)<=7);
        const pmDone=pmTasks7.filter(t=>t.Status==="Completed");
        const rate=pmTasks7.length?Math.round(pmDone.length/pmTasks7.length*100):0;
        return<div key={pi} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"11px 13px",marginBottom:8}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}><Avatar name={pm.Name} size={24} colorIdx={pi+1}/><div><div style={{fontSize:12,fontWeight:700,color:C.t2}}>{pm.Name}</div><div style={{fontSize:10,color:C.b4}}>Property Manager · {pmProps.length} properties</div></div></div>
            <div style={{display:"flex",gap:6}}><Badge type={rate>=80?"ok":"wn"} dot={false}>{rate}% rate</Badge>{pmOD.length>0&&<Badge type="er" dot={false}>{pmOD.length} overdue</Badge>}</div>
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{pmProps.map(p=>{const va=getVA(p.Title);const pOD=overdue.filter(t=>t.PropertyId===p.Title);
            return<div key={p.Title} style={{flex:"1 1 140px",minWidth:140,background:C.white,border:`1px solid ${C.b1}`,borderRadius:6,padding:"8px 10px"}}>
              <div style={{fontSize:11,fontWeight:700,color:C.t2}}>{p.PropertyName}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:1}}>{p.Units}u · {va?va.Name:"No VA"}</div>
              {pOD.length>0&&<div style={{display:"flex",gap:4,marginTop:4}}><Badge type="er" dot={false}>{pOD.length} overdue</Badge></div>}
            </div>;})}</div>
        </div>;})}
    </div>}

    {/* Property Cards */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10,marginBottom:12}}>
      {mgrProps.map(prop=>{const va=getVA(prop.Title);const pendQ=queue.filter(t=>t.PropertyId===prop.Title);const activeT=timers.filter(t=>t.PropertyId===prop.Title);const r7=data.activities.filter(a=>a.ActivityType==="Task"&&a.PropertyId===prop.Title&&dAgo(a.ActivityDate)<=7);const done=r7.filter(a=>a.Status==="Completed");const blocked=r7.filter(a=>a.Status==="Blocked");const od=overdue.filter(t=>t.PropertyId===prop.Title);
        return<div key={prop.Title} style={{background:C.white,border:`1px solid ${C.b1}`,borderRadius:8,padding:14,boxShadow:"0 1px 3px rgba(28,55,64,0.07)"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
            <div><div style={{fontSize:13,fontWeight:700}}>{prop.PropertyName}</div><div style={{fontSize:10,color:C.b4}}>{prop.Units} units · {prop.PropertyGroup}</div></div>
            {va?<div style={{display:"flex",alignItems:"center",gap:5}}><Avatar name={va.Name} size={20} colorIdx={0}/><Badge type={va.VATrackerStatus==="Out"?"er":"ok"} dot={false}>{va.Name.split(" ")[0]}</Badge></div>:<Badge type="er" dot={false}>No VA</Badge>}
          </div>
          <div style={{display:"flex",gap:7,marginBottom:10}}>
            {[{v:done.length,l:"Done 7d",c:C.ok},{v:pendQ.length,l:"Pending",c:pendQ.length?C.wn:C.t2},{v:od.length,l:"Overdue",c:od.length?C.er:C.ok},{v:blocked.length,l:"Blocked",c:blocked.length?C.er:C.ok}].map((s,i)=>
              <div key={i} style={{flex:1,textAlign:"center",padding:"8px 4px",background:C.tl00,borderRadius:6}}><div style={{fontSize:18,fontWeight:700,fontFamily:mono,color:s.c,lineHeight:1}}>{s.v}</div><div style={{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase",marginTop:2}}>{s.l}</div></div>)}
          </div>
          {va&&<button style={{...ss.btn(C.teal),...ss.xs,width:"100%"}} onClick={()=>{setSelProp(prop.Title);setTCat("");setTDesc("");setTPri("Normal");setTNotes("");}}>+ Assign Task to {va.Name.split(" ")[0]}</button>}
        </div>;})}
    </div>

    {/* Assign Task Form */}
    <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
      <div style={ss.cardT}>📌 Assign Task to VA</div><div style={{...ss.cardS,marginBottom:12}}>Select a property — routes automatically to the assigned VA</div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
        <div style={{flex:2,minWidth:180}}><label style={ss.label}>Property *</label><select style={ss.select} value={selProp} onChange={e=>setSelProp(e.target.value)}><option value="">Select property...</option>{mgrProps.map(p=>{const va=getVA(p.Title);return<option key={p.Title} value={p.Title}>{p.PropertyName} ({p.Units}u){va?` → ${va.Name}`:""}</option>;})}</select></div>
        <div style={{flex:1,minWidth:150}}><label style={ss.label}>Category *</label><select style={ss.select} value={tCat} onChange={e=>setTCat(e.target.value)}><option value="">Select...</option>{cats.sort((a,b)=>a.name.localeCompare(b.name)).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
        <div style={{flex:0,minWidth:88}}><label style={ss.label}>Priority</label><select style={ss.select} value={tPri} onChange={e=>setTPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
      </div>
      <input style={{...ss.input,marginBottom:9}} value={tDesc} onChange={e=>setTDesc(e.target.value)} placeholder="Task description *"/>
      <input style={{...ss.input,marginBottom:9}} value={tNotes} onChange={e=>setTNotes(e.target.value)} placeholder="Notes for VA (optional)"/>
      {selProp&&(()=>{const va=getVA(selProp);return va?<div style={{fontSize:11,color:C.ok,marginBottom:9}}>→ Routes to: {va.Name}{va.VATrackerStatus==="Out"?` (⚠ OUT — goes to coverage)`:""}</div>:<div style={{fontSize:11,color:C.er,marginBottom:9}}>⚠ No VA assigned</div>;})()}
      <button style={{...ss.btn(C.gold,C.teal),width:"100%"}} onClick={handleSubmit}>📌 Add to VA's Queue</button>
    </div>
  </div>);
}

// ══════════════════════════════════════════════════════
// DASHBOARD VIEW
// ══════════════════════════════════════════════════════
function DashboardView({data,queue,timers,covQ,overdue,dfFrom,dfTo,setDfFrom,setDfTo,isAdmin,role,mgrProps}){
  if(!data)return null;
  const propIds=role==="manager"?new Set(mgrProps.map(p=>p.Title)):null;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&inRange(a.ActivityDate,dfFrom,dfTo)&&(!propIds||propIds.has(a.PropertyId)));
  const done=tasks.filter(a=>a.Status==="Completed");const blocked=tasks.filter(a=>a.Status==="Blocked");const inc=tasks.filter(a=>a.Status==="Incomplete");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&inRange(a.ActivityDate,dfFrom,dfTo));
  const shiftMin=shifts.reduce((s,a)=>s+(a.WorkMinutes||0),0);const taskMin=done.reduce((s,a)=>s+(a.DurationMin||0),0);
  const rate=tasks.length?Math.round(done.length/tasks.length*100):0;const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;
  const filteredVAs=propIds?data.vas.filter(v=>data.portfolios.some(p=>propIds.has(p.PropertyId)&&p.VAEmail?.toLowerCase()===v.Email.toLowerCase())):data.vas;

  return(<div>
    {/* Date filter */}
    <div style={{...ss.card,display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",padding:12,marginBottom:12}}>
      <div><label style={ss.label}>From</label><input type="date" style={{...ss.input,width:140}} value={dfFrom} onChange={e=>setDfFrom(e.target.value)}/></div>
      <div><label style={ss.label}>To</label><input type="date" style={{...ss.input,width:140}} value={dfTo} onChange={e=>setDfTo(e.target.value)}/></div>
      <div style={{display:"flex",gap:5}}>{[{l:"7d",d:7},{l:"14d",d:14},{l:"30d",d:30}].map(p=><button key={p.l} style={{...ss.btnO(C.t2,C.b2),...ss.xs}} onClick={()=>{const f=new Date();f.setDate(f.getDate()-p.d);setDfFrom(f.toISOString().slice(0,10));setDfTo(today());}}>{p.l}</button>)}</div>
    </div>
    {/* KPIs */}
    <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
      <KPI label="Tasks" value={tasks.length}/><KPI label="Done" value={done.length} color={C.ok} sub={`${rate}% rate`}/><KPI label="Utilization" value={`${util}%`} color={util>=65?C.ok:util>=50?C.wn:shiftMin>0?C.er:C.t2}/><KPI label="Overdue" value={overdue.length} color={overdue.length?C.er:C.ok}/><KPI label="Blocked" value={blocked.length} color={blocked.length?C.er:C.ok}/><KPI label="Coverage" value={covQ.length} color={covQ.length?C.wn:C.ok}/>
    </div>
    {/* Needs Attention */}
    {(blocked.length>0||overdue.length>0)&&<div style={{...ss.card,borderTop:`3px solid ${C.er}`,background:C.erb}}>
      <div style={{fontSize:13,fontWeight:700,color:C.er,marginBottom:8}}>🚨 Needs Attention</div>
      {blocked.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:10,fontWeight:700,color:C.er,marginBottom:4}}>Blocked ({blocked.length})</div>{blocked.slice(0,5).map((t,i)=><div key={i} style={{fontSize:11,color:C.b6,marginBottom:2}}>{t.VAName}: {t.Title} — {t.PropertyName}{t.Notes?` · ${t.Notes}`:""}</div>)}</div>}
      {overdue.length>0&&<div><div style={{fontSize:10,fontWeight:700,color:C.er,marginBottom:4}}>Overdue ({overdue.length})</div>{overdue.slice(0,5).map((t,i)=><div key={i} style={{fontSize:11,color:C.b6,marginBottom:2}}>{t.VAName}: {t.Title} — {t.PropertyName} · from {fD(t.ActivityDate)}</div>)}</div>}
    </div>}
    {/* Kanban */}
    <div style={ss.card}>
      <div style={{...ss.cardT,marginBottom:11}}>📋 Task Board — Today</div>
      <div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:6}}>
        {filteredVAs.map((va,vi)=>{const isOut=va.VATrackerStatus==="Out";const vQ=queue.filter(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());const vT=timers.filter(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());const vOD=overdue.filter(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());const port=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase());
          return<div key={va.Email} style={{flex:"0 0 185px",background:isOut?C.erb:C.bg,borderRadius:8,padding:11,border:`1px solid ${isOut?"rgba(184,59,42,0.2)":C.b1}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:9,paddingBottom:8,borderBottom:`2px solid ${C.tl}`}}>
              <div style={{display:"flex",alignItems:"center",gap:6}}><Avatar name={va.Name} size={24} colorIdx={vi} isOut={isOut}/><div><div style={{fontSize:11,fontWeight:700,color:isOut?C.er:C.t2}}>{va.Name.split(" ")[0]}</div><div style={{fontSize:9,color:C.b4}}>{port.length} props</div></div></div>
              {isOut?<Badge type="er">OUT</Badge>:vT.length>0?<Badge type="ok" dot={false}>Active</Badge>:null}
            </div>
            {vT.map(t=><div key={t._localId} style={{background:C.white,border:`1px solid ${C.b1}`,borderLeft:`3px solid ${C.ok}`,borderRadius:6,padding:8,marginBottom:5,boxShadow:"0 1px 2px rgba(28,55,64,0.04)"}}><div style={{fontSize:8,fontWeight:700,color:C.ok,textTransform:"uppercase",marginBottom:2}}>● Active</div><div style={{fontSize:11,fontWeight:600}}>{t.Title}</div><div style={{fontSize:9,color:C.b4}}>{t.PropertyName}</div></div>)}
            {vOD.map(t=><div key={t.id} style={{background:C.erb,border:`1px solid ${C.b1}`,borderLeft:`3px solid ${C.er}`,borderRadius:6,padding:8,marginBottom:5}}><div style={{fontSize:9,color:C.er,marginBottom:2}}>⚠ Overdue {fD(t.ActivityDate)}</div><div style={{fontSize:11,fontWeight:600}}>{t.Title}</div><div style={{fontSize:9,color:C.b4}}>{t.PropertyName}</div></div>)}
            {vQ.filter((_,i)=>i<3).map(t=><div key={t._localId} style={{background:C.white,border:`1px solid ${C.b1}`,borderLeft:`3px solid ${t.Priority==="Urgent"?C.er:t.Priority==="High"?C.wn:C.b2}`,borderRadius:6,padding:8,marginBottom:5}}><div style={{fontSize:11,fontWeight:600}}>{t.Title}</div><div style={{fontSize:9,color:C.b4}}>{t.PropertyName}</div></div>)}
            {!vT.length&&!vQ.length&&!vOD.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:10,fontStyle:"italic"}}>{isOut?"Tasks in coverage":"No pending"}</div>}
          </div>;})}
      </div>
    </div>
    {/* Performance Table */}
    <div style={ss.card}>
      <div style={{...ss.cardT,marginBottom:11}}>VA Performance</div>
      <div style={{overflowX:"auto",borderRadius:8,border:`1px solid ${C.b1}`}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["VA","Done","Rate","Utilization","Overdue","Blocked","Incomplete"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead><tbody>
          {filteredVAs.map((va,vi)=>{const vT=tasks.filter(a=>a.VAEmail===va.Email);const vD=vT.filter(a=>a.Status==="Completed");const vB=vT.filter(a=>a.Status==="Blocked");const vI=vT.filter(a=>a.Status==="Incomplete");
            const vS=shifts.filter(a=>a.VAEmail===va.Email);const vSm=vS.reduce((s,a)=>s+(a.WorkMinutes||0),0);const vTm=vD.reduce((s,a)=>s+(a.DurationMin||0),0);
            const vR=vT.length?Math.round(vD.length/vT.length*100):0;const vU=vSm>0?Math.round(vTm/vSm*100):0;const vOD=overdue.filter(t=>t.VAEmail===va.Email);
            return<tr key={va.Email} style={{opacity:va.VATrackerStatus==="Out"?0.5:1}}>
              <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:7}}><Avatar name={va.Name} size={24} colorIdx={vi} isOut={va.VATrackerStatus==="Out"}/><strong style={{color:va.VATrackerStatus==="Out"?C.er:C.t2}}>{va.Name}</strong></div></td>
              <td style={ss.td}>{vD.length}/{vT.length}</td>
              <td style={ss.td}><Badge type={vR>=85?"ok":vR>=60?"wn":"er"}>{vR}%</Badge></td>
              <td style={ss.td}><div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:40,height:5,background:C.b1,borderRadius:3,overflow:"hidden"}}><div style={{width:`${Math.min(vU,100)}%`,height:"100%",borderRadius:3,background:vU>=65?C.ok:vU>=50?C.wn:C.er}}/></div><span style={{fontSize:10,fontWeight:700,color:vU>=65?C.ok:vU>=50?C.wn:vSm>0?C.er:C.t2,fontFamily:mono}}>{vU}%</span></div></td>
              <td style={ss.td}>{vOD.length>0?<Badge type="er" dot={false}>{vOD.length}</Badge>:"0"}</td>
              <td style={ss.td}>{vB.length>0?<Badge type="er" dot={false}>{vB.length}</Badge>:"0"}</td>
              <td style={ss.td}>{vI.length>0?<Badge type="er" dot={false}>{vI.length}</Badge>:"0"}</td>
            </tr>;})}
        </tbody></table></div>
    </div>
  </div>);
}

// ══════════════════════════════════════════════════════
// COACHING VIEW
// ══════════════════════════════════════════════════════
function CoachingView({data}){
  const[selVa,setSelVa]=useState("");const[period,setPeriod]=useState(7);
  if(!data)return null;
  const va=data.vas.find(v=>v.Email?.toLowerCase()===(selVa||data.vas[0]?.Email||"").toLowerCase());
  const vaEmail=va?.Email?.toLowerCase()||"";
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=period);
  const done=tasks.filter(t=>t.Status==="Completed");const blocked=tasks.filter(t=>t.Status==="Blocked");const inc=tasks.filter(t=>t.Status==="Incomplete");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=period);
  const shiftMin=shifts.reduce((s,a)=>s+(a.WorkMinutes||0),0);const taskMin=done.reduce((s,a)=>s+(a.DurationMin||0),0);
  const rate=tasks.length?Math.round(done.length/tasks.length*100):0;const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;
  // Category breakdown
  const catMap={};done.forEach(t=>{if(!catMap[t.Category])catMap[t.Category]={c:0,m:0};catMap[t.Category].c++;catMap[t.Category].m+=(t.DurationMin||0);});
  const catList=Object.entries(catMap).sort((a,b)=>b[1].m-a[1].m);
  // Pattern detection: repeated incompletes
  const incMap={};inc.forEach(t=>{const k=t.Title;if(!incMap[k])incMap[k]={c:0,prop:t.PropertyName};incMap[k].c++;});
  const patterns=Object.entries(incMap).filter(([_,v])=>v.c>=2).sort((a,b)=>b[1].c-a[1].c);

  return(<div>
    <div style={{...ss.card,display:"flex",gap:9,alignItems:"flex-end",flexWrap:"wrap",padding:12,marginBottom:12}}>
      <div style={{flex:1,minWidth:150}}><label style={ss.label}>VA</label><select style={ss.select} value={selVa||data.vas[0]?.Email||""} onChange={e=>setSelVa(e.target.value)}>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}{v.VATrackerStatus==="Out"?" [OUT]":""}</option>)}</select></div>
      <div style={{flex:1,minWidth:150}}><label style={ss.label}>Date Range</label><select style={ss.select} value={period} onChange={e=>setPeriod(Number(e.target.value))}><option value={7}>Last 7 days</option><option value={14}>Last 14 days</option><option value={30}>Last 30 days</option></select></div>
    </div>
    <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:12}}>
      <KPI label="Tasks" value={tasks.length}/><KPI label="Rate" value={`${rate}%`} color={rate>=85?C.ok:rate>=60?C.wn:tasks.length?C.er:C.t2}/><KPI label="Util" value={`${util}%`} color={util>=65?C.ok:util>=50?C.wn:shiftMin>0?C.er:C.t2}/><KPI label="Blocked" value={blocked.length} color={blocked.length?C.er:C.ok}/><KPI label="Incomplete" value={inc.length} color={inc.length>3?C.er:inc.length>0?C.wn:C.ok}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
      <div style={{...ss.card,margin:0}}><div style={ss.cardT}>Time by Category</div>
        {catList.map(([cat,d])=><div key={cat} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 0",borderBottom:`1px solid ${C.b1}`}}>
          <span>{catIcon[cat]||"📁"}</span><div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{cat}</div><div style={{fontSize:10,color:C.b4}}>{d.c} tasks</div></div><div style={{fontFamily:mono,fontSize:12,fontWeight:700,color:C.t2}}>{fM(d.m)}</div>
        </div>)}{!catList.length&&<div style={{color:C.b4,fontSize:12,padding:"12px 0"}}>No data</div>}
      </div>
      <div style={{...ss.card,margin:0}}><div style={ss.cardT}>Property Coverage</div>
        {(data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===vaEmail).map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean)).map(p=>{
          const ok=tasks.some(t=>t.PropertyId===p.Title);
          return<div key={p.Title} style={{display:"flex",alignItems:"center",gap:5,padding:"3px 0",borderBottom:`1px solid ${C.b1}`}}>
            <span style={{color:ok?C.ok:C.er,fontSize:12}}>{ok?"✓":"✗"}</span><span style={{flex:1,fontSize:12,color:ok?C.t2:C.b4}}>{p.PropertyName}</span></div>;})}
      </div>
    </div>
    {/* Pattern Detection */}
    {patterns.length>0&&<div style={ss.card}>
      <div style={{...ss.cardT,marginBottom:10}}>Pattern Detection</div>
      {patterns.map(([title,d],i)=><div key={i} style={{display:"flex",alignItems:"flex-start",gap:9,padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
        <span style={{fontSize:15}}>🔧</span><div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>Recurring incomplete: "{title}"</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{d.c}× this period — {d.prop}</div></div><Badge type="er" dot={false}>{d.c}×</Badge>
      </div>)}
    </div>}
    {blocked.length>0&&<div style={{...ss.card,borderTop:`3px solid ${C.er}`}}><div style={{...ss.cardT,color:C.er,marginBottom:10}}>⚠ Blocked Tasks</div>{blocked.map((t,i)=><div key={i} style={{padding:"5px 0",borderBottom:`1px solid ${C.b1}`}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{t.Title}</div><div style={{fontSize:10,color:C.b4}}>{t.PropertyName}{t.Notes?` · ${t.Notes}`:""}</div></div>)}</div>}
  </div>);
}

// ══════════════════════════════════════════════════════
// HISTORY VIEW
// ══════════════════════════════════════════════════════
function HistoryView({data,role,myEmail,isMgr,mgrProps}){
  const[view,setView]=useState("tasks");
  if(!data)return null;
  const propFilter=role==="manager"?new Set(mgrProps.map(p=>p.Title)):null;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.Status!=="Queued"&&a.Status!=="In Progress"&&(isMgr?(propFilter?propFilter.has(a.PropertyId):true):a.VAEmail?.toLowerCase()===myEmail)).slice(0,50);
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&(role==="admin"||a.VAEmail?.toLowerCase()===myEmail)).slice(0,30);

  return(<div>
    <div style={{display:"flex",gap:8,marginBottom:12}}>
      <button style={view==="tasks"?ss.btn(C.teal):ss.btnO(C.t2,C.b2)} onClick={()=>setView("tasks")}>Tasks</button>
      {role!=="manager"&&<button style={view==="shifts"?ss.btn(C.teal):ss.btnO(C.t2,C.b2)} onClick={()=>setView("shifts")}>Shifts</button>}
    </div>
    {view==="tasks"&&<div style={{...ss.card,overflowX:"auto"}}><div style={{borderRadius:8,border:`1px solid ${C.b1}`,overflow:"hidden"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Date","VA","Task","Property","Source","Duration","Status"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead><tbody>
      {tasks.map((t,i)=><tr key={i}><td style={{...ss.td,whiteSpace:"nowrap"}}>{fD(t.ActivityDate)}</td><td style={{...ss.td,fontWeight:600,fontSize:11}}>{t.VAName}</td><td style={{...ss.td,maxWidth:200}}>{t.Title}{t.Notes&&<div style={{fontSize:10,color:C.b4}}>💬 {t.Notes}</div>}</td><td style={{...ss.td,fontSize:11}}>{t.PropertyName}</td><td style={ss.td}><Badge type={{Daily:"in",Assigned:"wn","Ad-Hoc":"ne",Coverage:"wn"}[t.Source]||"ne"} dot={false}>{t.Source}</Badge></td><td style={{...ss.td,fontFamily:mono,fontSize:11}}>{t.DurationMin?fM(t.DurationMin):"—"}</td><td style={ss.td}><Badge type={{Completed:"ok",Blocked:"er",Incomplete:"er"}[t.Status]||"ne"}>{t.Status}</Badge></td></tr>)}
    </tbody></table></div>{!tasks.length&&<div style={{textAlign:"center",padding:40,color:C.b4}}>No history.</div>}</div>}
    {view==="shifts"&&<div style={{...ss.card,overflowX:"auto"}}><div style={{borderRadius:8,border:`1px solid ${C.b1}`,overflow:"hidden"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Date","VA","In","Out","Break","Working"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead><tbody>
      {shifts.map((s,i)=><tr key={i}><td style={{...ss.td,whiteSpace:"nowrap"}}>{fD(s.ActivityDate)}</td><td style={{...ss.td,fontWeight:600}}>{s.VAName}</td><td style={{...ss.td,fontFamily:mono,fontSize:11}}>{fT(s.StartTime)}</td><td style={{...ss.td,fontFamily:mono,fontSize:11}}>{fT(s.EndTime)}</td><td style={{...ss.td,fontFamily:mono}}>{fM(s.BreakMinutes)}</td><td style={{...ss.td,fontFamily:mono,fontWeight:600}}>{fM(s.WorkMinutes)}</td></tr>)}
    </tbody></table></div>{!shifts.length&&<div style={{textAlign:"center",padding:40,color:C.b4}}>No shifts.</div>}</div>}
  </div>);
}

// ══════════════════════════════════════════════════════
// ADMIN VIEW (with sub-navigation)
// ══════════════════════════════════════════════════════
function AdminView({data,myEmail,acct,config,queue,covQ,onToggleAbsence,onAssignTask,onUpdateConfig,onAssignProp,onUnassignProp,onReassignVA,onAddProperty,onEditProperty,onEditEmployee,onDeleteTask}){
  const[sub,setSub]=useState("team");
  const[showAssign,setShowAssign]=useState(false);const[aVa,setAVa]=useState("");const[aCat,setACat]=useState("");const[aPri,setAPri]=useState("Normal");const[aDesc,setADesc]=useState("");const[aNotes,setANotes]=useState("");const[aProps,setAProps]=useState([]);
  const[showRec,setShowRec]=useState(false);const[rVa,setRVa]=useState("");const[rCat,setRCat]=useState("");const[rDesc,setRDesc]=useState("");const[rProps,setRProps]=useState([]);
  const[editIdx,setEditIdx]=useState(null);const[editDesc,setEditDesc]=useState("");
  const[portVa,setPortVa]=useState("");const[portProp,setPortProp]=useState("");
  const[showAddProp,setShowAddProp]=useState(false);const[npName,setNpName]=useState("");const[npGroup,setNpGroup]=useState("Multifamily");const[npUnits,setNpUnits]=useState("");const[npPm,setNpPm]=useState("");
  const[roleFilter,setRoleFilter]=useState("all");
  // Employee edit state
  const[editEmpId,setEditEmpId]=useState(null);const[eName,setEName]=useState("");const[eEmail,setEEmail]=useState("");const[eTitle,setETitle]=useState("");const[eRole,setERole]=useState("");
  // Property edit state
  const[editPropId,setEditPropId]=useState(null);const[epName,setEpName]=useState("");const[epUnits,setEpUnits]=useState("");const[epGroup,setEpGroup]=useState("");const[epPm,setEpPm]=useState("");const[epVa,setEpVa]=useState("");
  if(!data)return null;
  const cats=config?.categories||[];const rTasks=config?.recurringTasks||[];
  function vaProps(email){return data.properties.filter(p=>data.portfolios.some(pt=>pt.VAEmail?.toLowerCase()===email.toLowerCase()&&pt.PropertyId===p.Title));}
  function toggleAProp(id){setAProps(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);}
  function toggleRProp(id){setRProps(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);}

  function handleAssign(){if(!aVa||!aCat||!aDesc)return;const va=data.vas.find(v=>v.Email===aVa);const cat=cats.find(c=>c.id===aCat);
    const propList=aProps.length>0?aProps:[""];
    propList.forEach(pid=>{const prop=pid?data.properties.find(p=>p.Title===pid):null;
      onAssignTask({Title:aDesc,VAEmail:aVa,VAName:va?.Name||aVa,PropertyId:pid||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Priority:aPri,Source:"Assigned",AssignedByEmail:myEmail,AssignedByName:acct?.name||myEmail,Notes:aNotes});
    });setADesc("");setAVa("");setACat("");setAProps([]);setAPri("Normal");setANotes("");setShowAssign(false);}

  function saveRec(newList){onUpdateConfig({...config,recurringTasks:newList});}
  function addRec(){if(!rVa||!rCat||!rDesc)return;const propList=rProps.length>0?rProps:[""];const newTasks=propList.map(pid=>({vaEmail:rVa,category:rCat,description:rDesc,propertyId:pid,active:true}));saveRec([...rTasks,...newTasks]);setRDesc("");setRVa("");setRCat("");setRProps([]);setShowRec(false);}
  function toggleRec(i){const u=[...rTasks];u[i]={...u[i],active:!u[i].active};saveRec(u);}
  function deleteRec(i){if(!window.confirm("Delete this recurring task?"))return;saveRec(rTasks.filter((_,j)=>j!==i));}
  function saveEdit(i){const u=[...rTasks];u[i]={...u[i],description:editDesc};saveRec(u);setEditIdx(null);}

  // Sub-nav tabs
  const subTabs=[{k:"team",l:"👥 Team"},{k:"sched",l:"🔄 Recurring"},{k:"port",l:"🏠 Portfolio"},{k:"settings",l:"⚙ Settings"}];
  // Group employees by role
  const emps=data.employees.filter(e=>e.EmployeeActive!==false);
  const roleGroups=[{key:"va",label:"Virtual Assistants",bg:C.tl0,fg:C.t2,emps:emps.filter(e=>detectRole(e)==="va")},{key:"mgr",label:"Property Managers",bg:C.gl,fg:C.g2,emps:emps.filter(e=>detectRole(e)==="manager")},{key:"regional",label:"Regional / Portfolio",bg:C.infb,fg:C.inf,emps:emps.filter(e=>detectRole(e)==="regional")},{key:"admin",label:"Admins",bg:C.erb,fg:C.er,emps:emps.filter(e=>detectRole(e)==="admin")}];
  const inactiveEmps=data.employees.filter(e=>e.EmployeeActive===false);

  return(<div>
    {/* Sub-navigation */}
    <div style={{display:"flex",gap:4,marginBottom:14,padding:4,background:C.b1,borderRadius:8,flexWrap:"wrap"}}>
      {subTabs.map(t=><button key={t.k} style={{flex:1,padding:"7px 8px",fontSize:11,fontWeight:600,color:sub===t.k?C.t2:C.b4,background:sub===t.k?C.white:"transparent",border:"none",cursor:"pointer",borderRadius:6,fontFamily:fnt,textAlign:"center",whiteSpace:"nowrap",boxShadow:sub===t.k?"0 1px 3px rgba(28,55,64,0.07)":"none"}} onClick={()=>setSub(t.k)}>{t.l}</button>)}
    </div>

    {/* ── TEAM ── */}
    {sub==="team"&&<div>
      {/* Role filter */}
      <div style={{display:"flex",gap:5,marginBottom:12,flexWrap:"wrap"}}>
        {[{k:"all",l:`All (${emps.length})`},...roleGroups.map(g=>({k:g.key,l:`${g.label.split(" ")[0]} (${g.emps.length})`}))].map(f=>
          <button key={f.k} style={{padding:"5px 12px",fontSize:11,fontWeight:600,border:`1px solid ${roleFilter===f.k?C.teal:C.b2}`,borderRadius:99,background:roleFilter===f.k?C.teal:"transparent",color:roleFilter===f.k?"#fff":C.b4,cursor:"pointer",fontFamily:fnt}} onClick={()=>setRoleFilter(f.k)}>{f.l}</button>)}
      </div>
      <div style={{...ss.card,padding:0,overflow:"hidden"}}>
        {roleGroups.filter(g=>roleFilter==="all"||roleFilter===g.key).map(g=><div key={g.key}>
          <div style={{padding:"8px 13px",borderBottom:`1px solid ${C.b1}`,borderTop:`1px solid ${C.b1}`,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.06em",background:g.bg,color:g.fg}}>{g.label}</div>
          {g.emps.map(emp=>{const port=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===emp.Email?.toLowerCase());const isEditing=editEmpId===emp.id;const isSelf=emp.Email?.toLowerCase()===myEmail;
            return<div key={emp.id}>
            <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 13px",borderBottom:isEditing?"none":`1px solid ${C.b1}`}}>
              <Avatar name={emp.Name} size={30} isOut={emp.VATrackerStatus==="Out"}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:13,fontWeight:700,color:emp.VATrackerStatus==="Out"?C.er:C.t2}}>{emp.Name}</div>
                <div style={{fontSize:10,color:C.b4,marginTop:1}}>{emp.Email}</div>
                <div style={{display:"flex",gap:5,marginTop:4,flexWrap:"wrap"}}>
                  <span style={{display:"inline-flex",padding:"2px 8px",fontSize:10,fontWeight:700,borderRadius:99,background:g.bg,color:g.fg}}>{g.key==="va"?"VA":g.key==="mgr"?"Manager":g.key==="regional"?"Regional":"Admin"}</span>
                  <Badge type={emp.VATrackerStatus==="Out"?"er":"ok"} dot={false}>{emp.VATrackerStatus||"Active"}</Badge>
                  {port.length>0&&<Badge type="ne" dot={false}>{port.length} props</Badge>}
                  {emp.VATrackerRole&&<span style={{fontSize:9,fontWeight:700,padding:"2px 7px",background:C.gold,color:C.teal,borderRadius:99}}>Role Override: {emp.VATrackerRole}</span>}
                  {isSelf&&<em style={{fontSize:10,color:C.b4}}>You</em>}
                </div>
                {emp.VATrackerRole&&<div style={{fontSize:10,color:C.b4,marginTop:3}}>Job Title: {emp.JobTitle} · Tracker Role Override: <strong style={{color:C.t2}}>{emp.VATrackerRole}</strong></div>}
              </div>
              <div style={{display:"flex",gap:5,flexShrink:0}}>
                <button style={{...ss.btnO(C.t2,C.b2),...ss.xs}} onClick={()=>{if(isEditing){setEditEmpId(null);}else{setEditEmpId(emp.id);setEName(emp.Name||"");setEEmail(emp.Email||"");setETitle(emp.JobTitle||"");setERole(emp.VATrackerRole||"");}}}>{isEditing?"Cancel":"Edit"}</button>
                {detectRole(emp)==="va"&&<button style={{...ss.btnO(emp.VATrackerStatus==="Out"?C.ok:C.er,emp.VATrackerStatus==="Out"?`rgba(26,122,70,0.3)`:`rgba(184,59,42,0.3)`),...ss.xs}} onClick={()=>onToggleAbsence(emp)}>{emp.VATrackerStatus==="Out"?"Mark In":"Mark Out"}</button>}
              </div>
            </div>
            {isEditing&&<div style={{margin:0,padding:"11px 13px",background:C.tl00,border:`1px solid ${C.tl}`,borderBottom:`1px solid ${C.b1}`}}>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
                <div style={{flex:1,minWidth:140}}><label style={ss.label}>Name</label><input style={ss.input} value={eName} onChange={e=>setEName(e.target.value)}/></div>
                <div style={{flex:1,minWidth:140}}><label style={ss.label}>Email</label><input style={ss.input} value={eEmail} onChange={e=>setEEmail(e.target.value)}/></div>
              </div>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
                <div style={{flex:1,minWidth:140}}><label style={ss.label}>Job Title</label><select style={ss.select} value={eTitle} onChange={e=>setETitle(e.target.value)}><option>Virtual Assistant</option><option>Property Manager</option><option>Regional/Portfolio Manager</option><option>Owner/Operator</option></select></div>
                <div style={{flex:1,minWidth:140}}><label style={{...ss.label,display:"flex",alignItems:"center",gap:5}}>Tracker Role Override <span style={{fontSize:9,fontWeight:700,padding:"1px 6px",background:C.gold,color:C.teal,borderRadius:99}}>KEY FIELD</span></label><select style={ss.select} value={eRole} onChange={e=>setERole(e.target.value)}><option value="">Auto — match job title</option><option value="va">va — task logging only</option><option value="manager">manager — see their properties</option><option value="regional">regional — portfolio overview + coaching</option><option value="admin">admin — full access to everything</option></select></div>
              </div>
              <div style={{background:C.wnb,border:"1px solid rgba(168,111,8,0.25)",borderRadius:6,padding:"9px 11px",fontSize:11,color:C.b6,marginBottom:10,lineHeight:1.6}}>
                <strong style={{color:C.wn}}>⚠ Tracker Role overrides job title.</strong> Example: a "Property Manager" with role set to <strong>admin</strong> gets full admin access. A "Regional/Portfolio Manager" with role set to <strong>regional</strong> gets the coaching tab and portfolio overview. Leave blank to auto-assign from job title.
              </div>
              <div style={{display:"flex",gap:6}}>
                <button style={{...ss.btn(C.ok),...ss.xs}} onClick={()=>{const fields={};if(eName!==emp.Name)fields.Name=eName;if(eEmail!==emp.Email)fields.Email=eEmail;if(eTitle!==emp.JobTitle)fields.JobTitle=eTitle;if(eRole!==(emp.VATrackerRole||""))fields.VATrackerRole=eRole;if(Object.keys(fields).length===0){setEditEmpId(null);return;}onEditEmployee(emp.id,fields);setEditEmpId(null);}}>✓ Save</button>
                <button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setEditEmpId(null)}>Cancel</button>
              </div>
            </div>}
            </div>;})}
        </div>)}
      </div>
    </div>}

    {/* ── RECURRING ── */}
    {sub==="sched"&&<div>
      <div style={{...ss.card,borderTop:`3px solid ${C.t3}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}><div><div style={ss.cardT}>🔄 Recurring Daily Tasks</div><div style={ss.cardS}>Generated each morning — skips weekends</div></div><button style={{...ss.btn(C.teal),...ss.xs}} onClick={()=>setShowRec(!showRec)}>{showRec?"Cancel":"+ Add"}</button></div>
        {showRec&&<div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:12,marginBottom:12}}>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
            <div style={{flex:1,minWidth:130}}><label style={ss.label}>VA *</label><select style={ss.select} value={rVa} onChange={e=>{setRVa(e.target.value);setRProps([]);}}><option value="">Select...</option>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
            <div style={{flex:1,minWidth:130}}><label style={ss.label}>Category *</label><select style={ss.select} value={rCat} onChange={e=>setRCat(e.target.value)}><option value="">Select...</option>{cats.sort((a,b)=>a.name.localeCompare(b.name)).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          </div>
          <div style={{marginBottom:10}}><label style={ss.label}>Properties {rProps.length>0&&`(${rProps.length})`}</label>
            <div style={{background:C.white,border:`1px solid ${C.b2}`,borderRadius:6,padding:8,maxHeight:140,overflowY:"auto"}}>
              {rVa&&<div style={{display:"flex",gap:6,marginBottom:6}}><button style={{...ss.btn(C.t3),...ss.xs}} onClick={()=>setRProps(vaProps(rVa).map(p=>p.Title))}>ALL</button><button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setRProps([])}>None</button></div>}
              <label style={{display:"flex",gap:6,fontSize:12,marginBottom:4,cursor:"pointer",color:C.b4}}><input type="checkbox" checked={rProps.length===0} onChange={()=>setRProps([])}/>General</label>
              {(rVa?vaProps(rVa):data.properties).map(p=><label key={p.Title} style={{display:"flex",gap:6,fontSize:12,marginBottom:4,cursor:"pointer",color:C.t2}}><input type="checkbox" checked={rProps.includes(p.Title)} onChange={()=>toggleRProp(p.Title)}/>{p.PropertyName}</label>)}
            </div>
          </div>
          <input style={{...ss.input,marginBottom:9}} value={rDesc} onChange={e=>setRDesc(e.target.value)} placeholder="Task description *"/>
          <button style={{...ss.btn(C.ok),width:"100%"}} onClick={addRec}>✓ {rProps.length>1?`Add ${rProps.length} Tasks`:"Add"}</button>
        </div>}
        {data.vas.map(va=>{const vr=rTasks.map((r,i)=>({...r,_i:i})).filter(r=>r.vaEmail?.toLowerCase()===va.Email.toLowerCase());if(!vr.length)return null;
          return<div key={va.Email} style={{marginBottom:12}}>
            <div style={ss.sec}>{va.Name} — {vr.filter(r=>r.active).length} active</div>
            {vr.map(r=>{const cat=cats.find(c=>c.id===r.category);const prop=r.propertyId?data.properties.find(p=>p.Title===r.propertyId):null;const isEd=editIdx===r._i;
              return<div key={r._i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 0",borderBottom:`1px solid ${C.b1}`,opacity:r.active?1:0.5}}>
                <span style={{fontSize:15}}>{catIcon[cat?.name]||"📁"}</span>
                <div style={{flex:1,minWidth:0}}>{isEd?<div style={{display:"flex",gap:4}}><input style={{...ss.input,padding:"4px 8px",fontSize:12}} value={editDesc} onChange={e=>setEditDesc(e.target.value)} autoFocus onKeyDown={e=>{if(e.key==="Enter")saveEdit(r._i);if(e.key==="Escape")setEditIdx(null);}}/><button style={{...ss.btn(C.ok),...ss.xs}} onClick={()=>saveEdit(r._i)}>✓</button></div>
                  :<div onClick={()=>{setEditIdx(r._i);setEditDesc(r.description);}} style={{cursor:"pointer"}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{r.description}</div>{prop&&<div style={{fontSize:10,color:C.b4}}>{prop.PropertyName}</div>}</div>}</div>
                {!isEd&&<div style={{display:"flex",gap:3}}>
                  <button style={{...ss.btnO(r.active?C.wn:C.ok,r.active?C.wn:C.ok),...ss.xs}} onClick={()=>toggleRec(r._i)}>{r.active?"Pause":"On"}</button>
                  <button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),...ss.xs}} onClick={()=>deleteRec(r._i)}>✕</button></div>}
              </div>;})}
          </div>;})}
      </div>
    </div>}

    {/* ── PORTFOLIO ── */}
    {sub==="port"&&<div>
      <div style={{...ss.card,borderTop:`3px solid ${C.t3}`}}>
        <div style={ss.cardT}>🏠 Portfolio & Property Management</div><div style={{...ss.cardS,marginBottom:12}}>Add properties, assign to VAs — tasks auto-route by assignment</div>
        {/* Add Property */}
        <div style={{background:C.gl,border:"1px solid rgba(205,160,75,0.3)",borderRadius:6,padding:12,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:10}}>+ Add Property</div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
            <div style={{flex:2,minWidth:150}}><label style={ss.label}>Name *</label><input style={ss.input} value={npName} onChange={e=>setNpName(e.target.value)} placeholder="e.g. Sunset Ridge"/></div>
            <div style={{flex:1,minWidth:100}}><label style={ss.label}>Group</label><select style={ss.select} value={npGroup} onChange={e=>setNpGroup(e.target.value)}><option>Multifamily</option><option>Single Family</option><option>Lease-Up</option></select></div>
            <div style={{flex:0,minWidth:70}}><label style={ss.label}>Units</label><input style={ss.input} type="number" value={npUnits} onChange={e=>setNpUnits(e.target.value)} placeholder="0"/></div>
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:9}}>
            <div style={{flex:1,minWidth:180}}><label style={ss.label}>Property Manager *</label><select style={ss.select} value={npPm} onChange={e=>setNpPm(e.target.value)}><option value="">Select PM...</option>
              <optgroup label="Admins / Regional">{data.employees.filter(e=>["admin","regional"].includes(detectRole(e))&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name} ({detectRole(e)})</option>)}</optgroup>
              <optgroup label="Property Managers">{data.pms.filter(e=>e.JobTitle==="Property Manager").map(pm=><option key={pm.Email} value={pm.Email}>{pm.Name}</option>)}</optgroup>
            </select></div>
          </div>
          <button style={{...ss.btn(C.gold,C.teal),...ss.xs}} onClick={()=>{if(!npName||!npUnits||!npPm)return;const pm=data.pms.find(p=>p.Email===npPm)||data.employees.find(e=>e.Email===npPm);onAddProperty({name:npName,group:npGroup,units:parseInt(npUnits)||0,pmEmail:npPm,pmName:pm?.Name||npPm});setNpName("");setNpUnits("");setNpPm("");}}>✓ Add Property</button>
        </div>
        {/* VA-to-Property assignment */}
        <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:11,marginBottom:12,display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div style={{flex:1,minWidth:140}}><label style={ss.label}>Assign VA</label><select style={ss.select} value={portVa} onChange={e=>setPortVa(e.target.value)}><option value="">Select VA...</option>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
          <div style={{flex:2,minWidth:180}}><label style={ss.label}>to Property</label><select style={ss.select} value={portProp} onChange={e=>setPortProp(e.target.value)}><option value="">Select...</option>{data.properties.filter(p=>!portVa||!data.portfolios.some(pt=>pt.VAEmail?.toLowerCase()===portVa.toLowerCase()&&pt.PropertyId===p.Title)).map(p=>{const a=data.portfolios.find(pt=>pt.PropertyId===p.Title);return<option key={p.Title} value={p.Title}>{p.PropertyName} ({p.Units}u){a?` — ${a.VAName}`:""}</option>;})}</select></div>
          <button style={{...ss.btn(C.ok),...ss.xs,opacity:(!portVa||!portProp)?0.5:1}} onClick={()=>{if(!portVa||!portProp)return;const va=data.vas.find(v=>v.Email===portVa);onAssignProp(portVa,va?.Name||portVa,portProp);setPortProp("");}}>Assign →</button>
        </div>
        {/* Property list with VA assignments — editable */}
        <div style={{borderRadius:8,border:`1px solid ${C.b1}`,overflow:"hidden"}}>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 50px 1.5fr 1.5fr 80px",gap:0,background:C.tl00,borderBottom:`1px solid ${C.b1}`}}>
            {["Property","Group","Units","PM","VA",""].map(h=><div key={h} style={{padding:"8px 11px",fontSize:11,fontWeight:700,color:C.t2}}>{h}</div>)}
          </div>
          {data.properties.map(p=>{const port=data.portfolios.find(pt=>pt.PropertyId===p.Title);const va=port?data.vas.find(v=>v.Email?.toLowerCase()===port.VAEmail?.toLowerCase()):null;const isEd=editPropId===p.id;
            return<div key={p.id} style={{borderBottom:`1px solid ${C.b1}`}}>
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 50px 1.5fr 1.5fr 80px",gap:0,alignItems:"center",background:!va?"rgba(184,59,42,0.03)":"transparent"}}>
                <div style={{padding:"9px 11px",fontSize:12,fontWeight:700,color:!va?C.er:C.t2}}>{p.PropertyName}</div>
                <div style={{padding:"9px 11px",fontSize:12,color:C.b6}}>{p.PropertyGroup}</div>
                <div style={{padding:"9px 11px",fontSize:12,color:C.b6}}>{p.Units}</div>
                <div style={{padding:"9px 11px",fontSize:12,color:C.b6}}>{p.PMName}</div>
                <div style={{padding:"9px 11px"}}>{va?<div style={{display:"flex",alignItems:"center",gap:5}}><Avatar name={va.Name} size={20}/><span style={{fontSize:11,color:C.b6}}>{va.Name}</span></div>:<Badge type="er" dot={false}>No VA assigned</Badge>}</div>
                <div style={{padding:"9px 11px"}}><button style={{...ss.btnO(C.t2,C.b2),...ss.xs}} onClick={()=>{if(isEd){setEditPropId(null);}else{setEditPropId(p.id);setEpName(p.PropertyName||"");setEpUnits(String(p.Units||0));setEpGroup(p.PropertyGroup||"Multifamily");setEpPm(p.PMEmail||"");setEpVa(port?.VAEmail||"");}}}>{isEd?"Cancel":"✎ Edit"}</button></div>
              </div>
              {isEd&&<div style={{padding:"10px 12px",background:C.tl00,borderTop:`1px solid ${C.b1}`}}>
                <div style={{fontSize:11,fontWeight:700,color:C.t3,marginBottom:8}}>Edit {p.PropertyName}</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                  <div style={{flex:2,minWidth:150}}><label style={ss.label}>Property Name</label><input style={ss.input} value={epName} onChange={e=>setEpName(e.target.value)}/></div>
                  <div style={{flex:1,minWidth:100}}><label style={ss.label}>Group</label><select style={ss.select} value={epGroup} onChange={e=>setEpGroup(e.target.value)}><option>Multifamily</option><option>Single Family</option><option>Lease-Up</option></select></div>
                  <div style={{flex:0,minWidth:70}}><label style={ss.label}>Units</label><input style={ss.input} type="number" value={epUnits} onChange={e=>setEpUnits(e.target.value)}/></div>
                </div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
                  <div style={{flex:1,minWidth:180}}><label style={ss.label}>Property Manager</label><select style={ss.select} value={epPm} onChange={e=>setEpPm(e.target.value)}>
                    <optgroup label="Admins / Regional">{data.employees.filter(e=>["admin","regional"].includes(detectRole(e))&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name} ({detectRole(e)})</option>)}</optgroup>
                    <optgroup label="Property Managers">{data.pms.filter(e=>e.JobTitle==="Property Manager").map(pm=><option key={pm.Email} value={pm.Email}>{pm.Name}</option>)}</optgroup>
                  </select></div>
                  <div style={{flex:1,minWidth:180}}><label style={ss.label}>VA Assigned</label><select style={ss.select} value={epVa} onChange={e=>setEpVa(e.target.value)}>
                    <option value="">— Unassign VA —</option>
                    {data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}{v.VATrackerStatus==="Out"?" (OUT)":""}</option>)}
                  </select></div>
                </div>
                <div style={{display:"flex",gap:6}}>
                  <button style={{...ss.btn(C.ok),...ss.xs}} onClick={async()=>{
                    // 1. Update property fields if changed
                    const propFields={};
                    if(epName!==p.PropertyName)propFields.PropertyName=epName;
                    if(parseInt(epUnits)!==p.Units)propFields.Units=parseInt(epUnits)||p.Units;
                    if(epGroup!==p.PropertyGroup)propFields.PropertyGroup=epGroup;
                    if(epPm!==(p.PMEmail||"")){propFields.PMEmail=epPm;const pm2=data.employees.find(e=>e.Email===epPm);propFields.PMName=pm2?.Name||epPm;}
                    if(Object.keys(propFields).length>0)onEditProperty(p.id,propFields);
                    // 2. Handle VA reassignment with full cascade
                    const oldVaEmail=(port?.VAEmail||"").toLowerCase();
                    const newVaEmail=epVa.toLowerCase();
                    if(oldVaEmail!==newVaEmail){
                      const nv=newVaEmail?data.vas.find(v=>v.Email?.toLowerCase()===newVaEmail):null;
                      await onReassignVA(p.Title,oldVaEmail,newVaEmail,nv?.Name||"");
                    }
                    setEditPropId(null);
                  }}>✓ Save Changes</button>
                  <button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setEditPropId(null)}>Cancel</button>
                  {port&&<button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),...ss.xs,marginLeft:"auto"}} onClick={()=>{if(window.confirm(`Remove ${va?.Name} from ${p.PropertyName}?`)){onUnassignProp(port.id,p.PropertyName,va?.Name||"VA");setEditPropId(null);}}}>Remove VA Assignment</button>}
                </div>
              </div>}
            </div>;})}
        </div>
        {/* Unassigned warning */}
        {(()=>{const aIds=new Set(data.portfolios.map(p=>p.PropertyId));const un=data.properties.filter(p=>!aIds.has(p.Title));if(!un.length)return null;
          return<div style={{padding:10,background:C.wnb,borderRadius:6,marginTop:8,border:`1px solid rgba(168,111,8,0.2)`}}><div style={{fontSize:10,fontWeight:700,color:C.wn,marginBottom:4}}>⚠ Unassigned ({un.length})</div>{un.map(p=><div key={p.Title} style={{fontSize:11,color:C.b6,padding:"2px 0"}}>{p.PropertyName} ({p.Units}u) — {p.PMName}</div>)}</div>;})()}
      </div>
    </div>}

    {/* ── SETTINGS ── */}
    {sub==="settings"&&<div>
      {/* Absence Management */}
      <div style={{...ss.card,borderTop:`3px solid ${C.er}`}}>
        <div style={ss.cardT}>🔒 Absence Management</div><div style={{...ss.cardS,marginBottom:12}}>Mark VAs out — their tasks move to coverage pool automatically</div>
        {data.vas.map(va=><div key={va.Email} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <Avatar name={va.Name} size={24} isOut={va.VATrackerStatus==="Out"}/>
            <div><div style={{fontSize:13,fontWeight:700,color:va.VATrackerStatus==="Out"?C.er:C.t2}}>{va.Name}{va.VATrackerStatus==="Out"?" 🛑":""}</div><div style={{fontSize:10,color:C.b4}}>{data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase()).length} properties</div></div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:7}}>
            <Badge type={va.VATrackerStatus==="Out"?"er":"ok"}>{va.VATrackerStatus||"Active"}</Badge>
            <button style={{...(va.VATrackerStatus==="Out"?ss.btn(C.ok):ss.btnO(C.er,`rgba(184,59,42,0.3)`)),...ss.xs}} onClick={()=>onToggleAbsence(va)}>{va.VATrackerStatus==="Out"?"Mark In":"Mark Out"}</button>
          </div>
        </div>)}
        {covQ.length>0&&<div style={{marginTop:10,padding:8,background:C.wnb,borderRadius:6,fontSize:11,color:C.wn}}>⚠ {covQ.length} unclaimed coverage tasks</div>}
      </div>
      {/* Assign Task */}
      <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={ss.cardT}>📌 Quick Assign Task</div><button style={ss.btn(C.gold,C.teal)} onClick={()=>setShowAssign(!showAssign)}>{showAssign?"Cancel":"Assign"}</button></div>
        {showAssign&&<div style={{marginTop:14}}>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
            <div style={{flex:1,minWidth:140}}><label style={ss.label}>VA *</label><select style={ss.select} value={aVa} onChange={e=>{setAVa(e.target.value);setAProps([]);}}><option value="">Select...</option>{data.vas.filter(v=>v.VATrackerStatus!=="Out").map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
            <div style={{flex:1,minWidth:140}}><label style={ss.label}>Category *</label><select style={ss.select} value={aCat} onChange={e=>setACat(e.target.value)}><option value="">Select...</option>{cats.sort((a,b)=>a.name.localeCompare(b.name)).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
            <div style={{minWidth:80}}><label style={ss.label}>Priority</label><select style={ss.select} value={aPri} onChange={e=>setAPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
          </div>
          {aVa&&<div style={{marginBottom:10}}><label style={ss.label}>Properties {aProps.length>0&&`(${aProps.length})`}</label>
            <div style={{background:C.white,border:`1px solid ${C.b2}`,borderRadius:6,padding:8,maxHeight:140,overflowY:"auto"}}>
              <div style={{display:"flex",gap:6,marginBottom:6}}><button style={{...ss.btn(C.t3),...ss.xs}} onClick={()=>setAProps(vaProps(aVa).map(p=>p.Title))}>ALL</button><button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setAProps([])}>None</button></div>
              <label style={{display:"flex",gap:6,fontSize:12,marginBottom:4,cursor:"pointer",color:C.b4}}><input type="checkbox" checked={aProps.length===0} onChange={()=>setAProps([])}/>General</label>
              {vaProps(aVa).map(p=><label key={p.Title} style={{display:"flex",gap:6,fontSize:12,marginBottom:4,cursor:"pointer"}}><input type="checkbox" checked={aProps.includes(p.Title)} onChange={()=>toggleAProp(p.Title)}/>{p.PropertyName} ({p.Units}u)</label>)}
            </div>
          </div>}
          <input style={{...ss.input,marginBottom:8}} value={aDesc} onChange={e=>setADesc(e.target.value)} placeholder="Task description *"/>
          <input style={{...ss.input,marginBottom:10}} value={aNotes} onChange={e=>setANotes(e.target.value)} placeholder="Notes (optional)"/>
          <button style={{...ss.btn(C.teal),width:"100%"}} onClick={handleAssign}>📌 {aProps.length>1?`Add ${aProps.length} Tasks`:"Add to Queue"}</button>
        </div>}
      </div>
      {/* Queued Task Management */}
      {queue.length>0&&<div style={ss.card}>
        <div style={ss.cardT}>🗑️ Queued Tasks ({queue.length})</div><div style={{...ss.cardS,marginBottom:10}}>Remove tasks that should no longer be in the queue</div>
        {data.vas.map(va=>{const vq=queue.filter(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());if(!vq.length)return null;
          return<div key={va.Email} style={{marginBottom:10}}><div style={ss.sec}>{va.Name} ({vq.length})</div>{vq.map(t=><TaskRow key={t._localId} task={t} onDelete={onDeleteTask}/>)}</div>;})}
      </div>}
      {/* Data Management */}
      <div style={{...ss.card,borderTop:`3px solid ${C.wn}`}}>
        <div style={ss.cardT}>🗂 Data Management</div>
        <div style={{...ss.cardS,marginBottom:12}}>Control which activity data is visible in the tracker. Old data stays in SharePoint but is hidden from all views.</div>
        <div style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,marginBottom:10}}>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:700,color:C.t2}}>Data Start Date</div>
            <div style={{fontSize:10,color:C.b4,marginTop:2}}>{config?.dataStartDate?`Showing activity from ${fD(config.dataStartDate)} onward. Everything before is hidden.`:"No cutoff set — showing all historical activity."}</div>
          </div>
          {config?.dataStartDate&&<Badge type="wn" dot={false}>Filtered from {fD(config.dataStartDate)}</Badge>}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end",marginBottom:10}}>
          <button style={ss.btn(C.wn)} onClick={()=>{if(!window.confirm("Reset data to start fresh from today?\n\nThis hides ALL activity before today from every tab — Dashboard, History, Scorecard, Coaching, everything.\n\nThe old data is NOT deleted from SharePoint. You can undo this by clearing the start date."))return;onUpdateConfig({...config,dataStartDate:today()});}}>🔄 Start Fresh from Today</button>
          <button style={ss.btnO(C.t2,C.b2)} onClick={()=>{const d=prompt("Enter custom start date (YYYY-MM-DD):",config?.dataStartDate||today());if(!d)return;if(!/^\d{4}-\d{2}-\d{2}$/.test(d)){alert("Use YYYY-MM-DD format");return;}onUpdateConfig({...config,dataStartDate:d});}}>📅 Set Custom Date</button>
          {config?.dataStartDate&&<button style={ss.btnO(C.er,`rgba(184,59,42,0.3)`)} onClick={()=>{if(!window.confirm("Remove the data cutoff? All historical activity will be visible again."))return;const nc={...config};delete nc.dataStartDate;onUpdateConfig(nc);}}>✕ Clear Cutoff</button>}
        </div>
        {config?.dataStartDate&&<div style={{background:C.wnb,border:"1px solid rgba(168,111,8,0.25)",borderRadius:6,padding:"8px 11px",fontSize:11,color:C.b6,lineHeight:1.6}}>
          <strong style={{color:C.wn}}>Note:</strong> This only hides data from the app's views. The records still exist in SharePoint's VA_Activity list. To permanently delete old records, do that directly in SharePoint. To undo this filter, click "Clear Cutoff" above.
        </div>}
      </div>
      {/* Data Sources */}
      <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:9,fontSize:11,color:C.b4}}><strong style={{color:C.t2}}>Data:</strong> Employees → VAs/PMs/Roles. VA_Properties → registry. VA_Portfolios → assignments. VA_TrackerConfig → categories/recurring. VA_Activity → tasks/shifts/absences. Auth: MSAL.</div>
    </div>}
  </div>);
}
