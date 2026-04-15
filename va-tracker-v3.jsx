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
const SCOPES=["Sites.ReadWrite.All","User.Read","Mail.Send"];

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
async function sendEmail(token,to,subject,body){
  try{await fetch("https://graph.microsoft.com/v1.0/me/sendMail",{method:"POST",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify({message:{subject,body:{contentType:"Text",content:body},toRecipients:[{emailAddress:{address:to}}]},saveToSentItems:false})});
  }catch(e){console.warn("[VT] Email failed:",e.message);}
}

async function safeGet(token,name,url){try{const r=await gAll(token,url);return r;}catch(e){console.warn(`[VT] ${name} failed:`,e.message);return[];}}

async function loadAll(token){
  const[eR,cR,pR,ptR,aR,gR,toR]=await Promise.all([
    safeGet(token,"Employees",`${lUrl("Employees")}?expand=fields&$top=200`),
    safeGet(token,"Config",`${lUrl("VA_TrackerConfig")}?expand=fields&$top=10`),
    safeGet(token,"Properties",`${lUrl("VA_Properties")}?expand=fields&$top=200`),
    safeGet(token,"Portfolios",`${lUrl("VA_Portfolios")}?expand=fields&$top=200`),
    safeGet(token,"Activity",`${lUrl("VA_Activity")}?expand=fields&$top=1000`),
    safeGet(token,"Guests",`${lUrl("VA_Guests")}?expand=fields&$top=200`),
    safeGet(token,"TimeOff",`${lUrl("VA_TimeOff")}?expand=fields&$top=500`),
  ]);
  const employees=eR.map(e=>({id:e.id,...e.fields}));
  const config=cR.length>0?JSON.parse(cR[0].fields.ConfigJSON||"{}"):{};
  const properties=pR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const portfolios=ptR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const allActs=aR.map(a=>({id:a.id,...a.fields}));
  // Filter out activity before dataStartDate (if set in config)
  const cutoff=config.dataStartDate||null;
  const activities=(cutoff?allActs.filter(a=>{const d=a.ActivityDate||a.StartTime||"";return d>=cutoff;}):allActs).filter(a=>a.Status!=="Deleted");
  const vas=employees.filter(e=>e.JobTitle==="Virtual Assistant"&&e.EmployeeActive!==false);
  const guests=gR.map(g=>({id:g.id,...g.fields})).filter(g=>g.IsActive!==false);
  const timeOff=toR.map(t=>({id:t.id,...t.fields}));
  const pms=employees.filter(e=>(e.JobTitle==="Property Manager"||e.JobTitle==="Regional/Portfolio Manager"||e.JobTitle==="Owner/Operator")&&e.EmployeeActive!==false);
  return{employees,config,properties,portfolios,activities,vas,pms,guests,timeOff};
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
  // ── Guest token detection ──
  const guestToken=useRef(new URLSearchParams(window.location.search).get("guest")).current;
  const[guestInfo,setGuestInfo]=useState(null);

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
  const[lastRefresh,setLastRefresh]=useState(Date.now());
  const[dfFrom,setDfFrom]=useState(()=>{const d=new Date();d.setDate(d.getDate()-7);return d.toISOString().slice(0,10);});
  const[dfTo,setDfTo]=useState(today());

  useEffect(()=>{const id=setInterval(()=>setTick(t=>t+1),1000);return()=>clearInterval(id);},[]);
  // Auto-refresh every 5 minutes
  useEffect(()=>{if(!token||!acct||guestToken)return;const id=setInterval(()=>{reload().catch(e=>console.warn("[VT] Auto-refresh failed:",e));},300000);return()=>clearInterval(id);},[token,acct]);
  const fl=useCallback(msg=>{setFlash(msg);setTimeout(()=>setFlash(""),2500);},[]);
  async function gT(){return(await refresh())||token;}

  // ── Guest Data Loading (via Cloudflare Worker) ──
  useEffect(()=>{
    if(!guestToken)return;
    setLoading(true);
    fetch(`https://va-tracker-guest.newshire-pm.workers.dev?token=${guestToken}`)
      .then(r=>r.json())
      .then(d=>{
        if(d.error){setError(d.error);setLoading(false);return;}
        setGuestInfo(d.guest);
        setRole("guest");
        setMyEmail(d.guest?.name||"Guest");
        // Transform Worker response to match loadAll format
        const employees=d.employees||[];
        const vas=employees.filter(e=>e.JobTitle==="Virtual Assistant"&&e.EmployeeActive!==false);
        const pms=employees.filter(e=>(e.JobTitle==="Property Manager"||e.JobTitle==="Regional/Portfolio Manager")&&e.EmployeeActive!==false);
        setData({employees,config:d.config||{},properties:d.properties||[],portfolios:d.portfolios||[],activities:d.activities||[],vas,pms,guests:[]});
        setLoading(false);
      })
      .catch(e=>{setError("Failed to load: "+e.message);setLoading(false);});
  },[guestToken]);

  // ── Load Data (authenticated users) ──
  useEffect(()=>{
    if(guestToken)return; // Skip if guest mode
    if(!token||!acct)return;
    const email=acct.username.toLowerCase();
    setMyEmail(email);setLoading(true);
    loadAll(token).then(d=>{
      setData(d);
      const me=d.employees.find(e=>(e.Email&&e.Email.toLowerCase()===email)||(e.M365UserId&&e.M365UserId.toLowerCase()===email)||(e.Email&&email.split("@")[0]===e.Email.toLowerCase().split("@")[0]));
      if(!me){setRole(null);setError("access_denied");}
      else{const r=detectRole(me);if(!r){setRole(null);setError("access_denied");}else{setRole(r);setError(null);setMyEmp(me);}}
      buildQueue(d,email,timersRef.current);
      // Restore active shift from SharePoint (clock-in persists across page refresh)
      const activeShift=d.activities.find(a=>a.ActivityType==="Shift"&&a.VAEmail?.toLowerCase()===email&&a.StartTime&&!a.EndTime&&a.Status!=="Deleted");
      if(activeShift&&!shift){
        const bks=activeShift.BreaksJSON?JSON.parse(activeShift.BreaksJSON):[];
        setShift({ClockIn:activeShift.StartTime,Breaks:bks,_ob:false,_bs:null,_spId:activeShift.id});
      }
      setLoading(false);
    }).catch(e=>{setError("load_error: "+e.message);setLoading(false);});
  },[token,acct]);

  function buildQueue(d,email,currentTimers=[]){
    const timerSpIds=new Set(currentTimers.map(t=>t._spId).filter(Boolean));
    const td=today();
    // Only load TODAY's queued tasks into the active queue — older ones show in overdue section only
    const persistedQueued=d.activities.filter(a=>a.ActivityType==="Task"&&a.Status==="Queued"&&!timerSpIds.has(a.id)&&(a.ActivityDate?.slice(0,10)===td||!a.ActivityDate));
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
    // Auto-out for scheduled time off
    if(d.timeOff){
      d.timeOff.filter(to=>to.Status==="Approved"&&to.StartDate?.slice(0,10)<=td&&(to.EndDate?.slice(0,10)>=td||!to.EndDate)).forEach(to=>{
        const va=d.vas.find(v=>v.Email?.toLowerCase()===to.VAEmail?.toLowerCase());
        if(va&&va.VATrackerStatus!=="Out"){console.log("[VT] Auto-out:",va.Name,"- scheduled time off");}
      });
    }
      // Build existing set from ALL today's tasks (any status) — case insensitive
      const todayTasks=d.activities.filter(a=>a.ActivityType==="Task"&&a.ActivityDate&&a.ActivityDate.slice(0,10)===td);
      const existing=new Set(todayTasks.map(t=>`${(t.VAEmail||"").toLowerCase()}|${t.Title}`));
      // Also include queued tasks we just loaded
      persistedQueued.forEach(a=>{existing.add(`${(a.VAEmail||"").toLowerCase()}|${a.Title}`);});
      const toSave=[];
      d.config.recurringTasks.forEach(rt=>{
        if(!rt.active)return;
        const freq=rt.frequency||"daily";
        if(freq==="weekly"){const wd=rt.weekDay||1;if(dow!==wd)return;}
        const key=`${(rt.vaEmail||"").toLowerCase()}|${rt.description}`;
        if(existing.has(key))return;
        existing.add(key); // Prevent intra-run duplicates
        const va=d.vas.find(v=>v.Email&&v.Email.toLowerCase()===rt.vaEmail.toLowerCase());
        if(!va)return;
        const prop=rt.propertyId?d.properties.find(p=>p.Title===rt.propertyId):null;
        const cat=d.config.categories?.find(c=>c.id===rt.category);
        const task={Title:rt.description,VAEmail:va.Email,VAName:va.Name,PropertyId:rt.propertyId||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Source:"Daily",Status:"Queued",Priority:"Normal",ActivityDate:td,ActivityType:"Task"};
        if(va.VATrackerStatus==="Out"){task.Source="Coverage";task.CoverageForEmail=va.Email;task.CoverageForName=va.Name;}
        toSave.push(task);
      });
      // Set queue FIRST with what we loaded from SP, then append newly generated tasks
      setQueue(q);setCovQ(cv);
      if(toSave.length>0){
        (async()=>{try{const tk=await gT();for(const task of toSave){const res=await gPost(tk,lUrl("VA_Activity"),{Title:task.Title,ActivityType:"Task",VAEmail:task.VAEmail,VAName:task.VAName,ActivityDate:task.ActivityDate,PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category,Source:task.Source,Status:task.Status,Priority:task.Priority||"Normal",CoverageForEmail:task.CoverageForEmail||"",CoverageForName:task.CoverageForName||""});
          const saved={...task,_localId:res.id,_spId:res.id,id:res.id};
          if(task.Source==="Coverage"){setCovQ(p=>[...p,saved]);}else{setQueue(p=>[...p,saved]);}
        }}catch(e){console.error("[VT] Daily task gen error:",e);}})();
      }
      return; // Already set queue/covQ above
    }
    setQueue(q);setCovQ(cv);
  }

  async function reload(){const t=await gT();const d=await loadAll(t);setData(d);buildQueue(d,myEmail,timersRef.current);setLastRefresh(Date.now());return d;}

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

  // ── Guest management ──
  async function addGuest({name,email,company,vaEmails}){
    try{const tk=await gT();
      const token=Date.now().toString(36)+Math.random().toString(36).slice(2,8);
      await gPost(tk,lUrl("VA_Guests"),{Title:token,GuestName:name,GuestEmail:email.toLowerCase(),Company:company||"",VAEmails:vaEmails.join(","),IsActive:true,InvitedDate:new Date().toISOString(),InvitedBy:acct?.name||myEmail});
      await reload();fl(`${name} added! Copy their guest URL from the Guests tab.`);
    }catch(e){fl("Error: "+e.message);}
  }
  async function deactivateGuest(id,name){
    if(!window.confirm(`Deactivate guest access for ${name}?`))return;
    try{const tk=await gT();await gPatch(tk,iUrl("VA_Guests",id),{IsActive:false});await reload();fl(`${name} deactivated`);}catch(e){fl("Error: "+e.message);}
  }

  // ── Time Off management ──
  async function requestTimeOff(req){
    try{const tk=await gT();
      await gPost(tk,lUrl("VA_TimeOff"),{Title:`TO-${Date.now().toString(36)}`,VAEmail:req.vaEmail,VAName:req.vaName,RequestType:req.type,StartDate:req.startDate,EndDate:req.endDate,HoursRequested:req.hours||0,Status:"Pending",PaidStatus:"TBD",VANotes:req.notes||"",RequestedDate:new Date().toISOString()});
      fl("Time off request submitted!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function updateTimeOff(id,fields){
    try{const tk=await gT();await gPatch(tk,iUrl("VA_TimeOff",id),fields);await reload();fl("Time off updated!");}catch(e){fl("Error: "+e.message);}
  }
  async function approveTimeOff(id,paidStatus,adminNotes,supervisorConfirmed){
    try{const tk=await gT();
      await gPatch(tk,iUrl("VA_TimeOff",id),{Status:"Approved",PaidStatus:paidStatus||"TBD",AdminNotes:adminNotes||"",SupervisorConfirmed:supervisorConfirmed||false,ApprovedBy:myEmp?.Name||acct?.name||myEmail,ApprovedDate:new Date().toISOString()});
      fl("Time off approved!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function denyTimeOff(id,adminNotes){
    try{const tk=await gT();
      await gPatch(tk,iUrl("VA_TimeOff",id),{Status:"Denied",AdminNotes:adminNotes||"",ApprovedBy:myEmp?.Name||acct?.name||myEmail,ApprovedDate:new Date().toISOString()});
      fl("Time off denied.");await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function logCallout(vaEmail,vaName,notes){
    try{const tk=await gT();
      await gPost(tk,lUrl("VA_TimeOff"),{Title:`CO-${Date.now().toString(36)}`,VAEmail:vaEmail,VAName:vaName,RequestType:"Callout",StartDate:new Date().toISOString(),EndDate:new Date().toISOString(),Status:"Approved",PaidStatus:"Unpaid",AdminNotes:notes||"Same-day callout",ApprovedBy:myEmp?.Name||myEmail,ApprovedDate:new Date().toISOString(),RequestedDate:new Date().toISOString()});
      fl(`${vaName} callout logged`);await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function logEarlyDeparture(vaEmail,vaName,departureTime,notes){
    try{const tk=await gT();
      await gPost(tk,lUrl("VA_TimeOff"),{Title:`ED-${Date.now().toString(36)}`,VAEmail:vaEmail,VAName:vaName,RequestType:"Early Departure",StartDate:new Date().toISOString(),DepartureTime:departureTime,Status:"Approved",PaidStatus:"Partial",AdminNotes:notes||"",ApprovedBy:myEmp?.Name||myEmail,ApprovedDate:new Date().toISOString(),RequestedDate:new Date().toISOString()});
      fl(`${vaName} early departure logged`);await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Shift management ──
  async function editShift(shiftId,fields){
    try{const tk=await gT();
      // Recalculate WorkMinutes if times changed
      if(fields.StartTime&&fields.EndTime){
        const bMin=fields.BreakMinutes||0;
        fields.WorkMinutes=Math.max(0,Math.round((new Date(fields.EndTime)-new Date(fields.StartTime))/6e4)-bMin);
      }
      await gPatch(tk,iUrl("VA_Activity",shiftId),fields);await reload();fl("Shift updated!");
    }catch(e){fl("Error: "+e.message);}
  }
  async function deleteShift(shiftId){
    if(!window.confirm("Delete this shift record? This cannot be undone."))return;
    try{const tk=await gT();
      // Mark as deleted rather than actual delete (preserves audit trail)
      await gPatch(tk,iUrl("VA_Activity",shiftId),{Status:"Deleted",Notes:"Deleted by admin"});
      await reload();fl("Shift deleted!");
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Review request (VA → Manager) ──
  async function sendReview(task,note){
    try{const tk=await gT();
      const groupId=task._spId||task._localId||task.Title;
      await gPost(tk,lUrl("VA_Activity"),{Title:`Review-${task.Title}`,ActivityType:"Review",VAEmail:task.VAEmail||myEmail,VAName:task.VAName||myEmp?.Name||myEmail,ActivityDate:new Date().toISOString(),PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category||"",Notes:note,Status:"Pending",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||myEmail,GroupId:groupId});
      // Email the PM
      const prop=data.properties.find(p=>p.Title===task.PropertyId);
      const pmEmail=prop?.PMEmail||task.PMEmail;
      if(pmEmail){
        await sendEmail(tk,pmEmail,`VA Review Request: ${task.Title} — ${task.PropertyName||""}`,
          `${myEmp?.Name||myEmail} needs your review on "${task.Title}" for ${task.PropertyName||"a property"}.\n\n"${note}"\n\nRespond in the VA Tracker: https://newshirepm.github.io/va-tracker/`);
      }
      fl("Review sent to PM"+(pmEmail?" — email notification sent":"")+"!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  // Reply to a review (works for both VA and PM — creates a response and toggles status)
  async function replyToReview(reviewId,note,newStatus){
    try{const tk=await gT();
      const orig=data.activities.find(a=>a.id===reviewId);if(!orig)return;
      // Create response record
      await gPost(tk,lUrl("VA_Activity"),{Title:`Response-${orig.Title||"Review"}`,ActivityType:"ReviewResponse",VAEmail:orig.VAEmail,VAName:orig.VAName,ActivityDate:new Date().toISOString(),PropertyId:orig.PropertyId||"",PropertyName:orig.PropertyName||"General",PMName:orig.PMName||"",Category:orig.Category||"",Notes:note,Status:"Completed",GroupId:orig.GroupId||"",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||acct?.name||myEmail});
      // Update review status
      await gPatch(tk,iUrl("VA_Activity",reviewId),{Status:newStatus||"Responded"});
      // Email notification to the other party
      const iAmVA=orig.VAEmail?.toLowerCase()===myEmail;
      if(iAmVA){
        // VA is replying — notify PM
        const prop=data.properties.find(p=>p.Title===orig.PropertyId);
        if(prop?.PMEmail)await sendEmail(tk,prop.PMEmail,`VA Reply: ${orig.Title?.replace("Review-","")} — ${orig.PropertyName||""}`,`${myEmp?.Name||myEmail} replied to a review request for "${orig.Title?.replace("Review-","")}" at ${orig.PropertyName||"property"}.\n\n"${note}"\n\nRespond in the VA Tracker: https://newshirepm.github.io/va-tracker/`);
      }else{
        // PM is replying — notify VA
        if(orig.VAEmail)await sendEmail(tk,orig.VAEmail,`PM Response: ${orig.Title?.replace("Review-","")} — ${orig.PropertyName||""}`,`${myEmp?.Name||acct?.name||"Your PM"} responded to your review request for "${orig.Title?.replace("Review-","")}" at ${orig.PropertyName||"property"}.\n\n"${note}"\n\nCheck your My Reviews in the VA Tracker: https://newshirepm.github.io/va-tracker/`);
      }
      fl(newStatus==="Resolved"?"Review resolved — notification sent!":"Reply sent — notification sent!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function resolveReview(reviewId,responseNote,responderName){
    // Backward compat wrapper — routes to replyToReview with Resolved status
    return replyToReview(reviewId,responseNote||"Resolved","Resolved");
  }
  async function assignReviewProperty(reviewId,propTitle,propName){
    try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",reviewId),{PropertyId:propTitle,PropertyName:propName});await reload();fl("Review assigned to "+propName);}catch(e){fl("Error: "+e.message);}
  }

  // ── Coaching notes ──
  async function saveCoachingNote(vaEmail,vaName,note){
    try{const tk=await gT();
      await gPost(tk,lUrl("VA_Activity"),{Title:`Coaching-${vaName}-${today()}`,ActivityType:"CoachingNote",VAEmail:vaEmail,VAName:vaName,ActivityDate:new Date().toISOString(),Notes:note,Status:"Completed",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||myEmail});
      fl("Coaching note saved!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Interruption logging ──
  async function logInterruption(int){
    try{const tk=await gT();
      await gPost(tk,lUrl("VA_Activity"),{Title:`Interruption-${int.type}`,ActivityType:"Interruption",VAEmail:int.vaEmail,VAName:int.vaName,ActivityDate:new Date().toISOString(),PropertyId:int.propertyId||"",PropertyName:int.propertyName||"General",Category:int.type,DurationMin:int.duration||0,Notes:int.notes||"",Status:"Completed"});
      // If convert-to-task, also create the task
      if(int.task){await addTask(int.task);}
      fl(`Interruption logged${int.task?" + task added":""}`);
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Daily metrics — live auto-save ──
  const metricsSpIdRef=useRef(null);
  async function saveDailyMetrics(metricsObj,notes){
    try{const tk=await gT();
      const combined=notes!==undefined?{...metricsObj,notes}:metricsObj;
      const notesStr=JSON.stringify(combined);
      if(metricsSpIdRef.current){
        await gPatch(tk,iUrl("VA_Activity",metricsSpIdRef.current),{Notes:notesStr,ActivityDate:new Date().toISOString()});
      }else{
        const existing=data?.activities.find(a=>a.ActivityType==="DailyMetrics"&&a.VAEmail?.toLowerCase()===myEmail&&a.ActivityDate?.slice(0,10)===today());
        if(existing){
          metricsSpIdRef.current=existing.id;
          await gPatch(tk,iUrl("VA_Activity",existing.id),{Notes:notesStr,ActivityDate:new Date().toISOString()});
        }else{
          const res=await gPost(tk,lUrl("VA_Activity"),{Title:`Metrics-${myEmp?.Name||myEmail}-${today()}`,ActivityType:"DailyMetrics",VAEmail:myEmail,VAName:myEmp?.Name||myEmail,ActivityDate:new Date().toISOString(),Notes:notesStr,Status:"Active"});
          metricsSpIdRef.current=res.id;
        }
      }
    }catch(e){console.warn("[VT] Metrics save error:",e.message);}
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
  async function clockIn(vaEmail,vaName){
    // Guard against React event objects being passed as first arg
    if(vaEmail&&typeof vaEmail==="object"){vaEmail=undefined;vaName=undefined;}
    const targetEmail=vaEmail||myEmail;const targetName=vaName||myEmp?.Name||myEmail;
    if(!vaEmail&&shift){fl("Already clocked in!");return;}
    try{const t=await gT();
      const now=new Date().toISOString();
      const res=await gPost(t,lUrl("VA_Activity"),{Title:`${targetName}-${today()}`,ActivityType:"Shift",VAEmail:targetEmail,VAName:targetName,ActivityDate:now,StartTime:now,Status:"Active"});
      if(!vaEmail){setShift({ClockIn:now,Breaks:[],_ob:false,_bs:null,_spId:res.id});fl("Clocked in!");}
      else{fl(`${targetName} clocked in!`);}
      await reload();
    }catch(e){fl("Error clocking in: "+e.message);}
  }
  function startBreak(){setShift(p=>p?{...p,_ob:true,_bs:new Date().toISOString()}:p);}
  function endBreak(){setShift(p=>{if(!p||!p._bs)return p;return{...p,_ob:false,Breaks:[...p.Breaks,{s:p._bs,e:new Date().toISOString()}],_bs:null};});}
  async function clockOut(){
    if(!shift)return;const now=new Date();const bks=[...shift.Breaks];
    if(shift._ob&&shift._bs)bks.push({s:shift._bs,e:now.toISOString()});
    const bMs=bks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);
    const bMin=Math.round(bMs/6e4);const wMin=Math.round((now-new Date(shift.ClockIn)-bMs)/6e4);
    try{const t=await gT();
      if(shift._spId){
        await gPatch(t,iUrl("VA_Activity",shift._spId),{EndTime:now.toISOString(),BreakMinutes:bMin,WorkMinutes:wMin,BreaksJSON:JSON.stringify(bks),Status:"Completed"});
      }else{
        await gPost(t,lUrl("VA_Activity"),{Title:`${myEmp?.Name||myEmail}-${today()}`,ActivityType:"Shift",VAEmail:myEmail,VAName:myEmp?.Name||myEmail,ActivityDate:shift.ClockIn,StartTime:shift.ClockIn,EndTime:now.toISOString(),BreakMinutes:bMin,WorkMinutes:wMin,BreaksJSON:JSON.stringify(bks),Status:"Completed"});
      }
      setShift(null);fl("Clocked out!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Absence ──
  async function toggleAbsence(va){
    const t=await gT();const ns=va.VATrackerStatus==="Out"?"Active":"Out";
    await gPatch(t,iUrl("Employees",va.id),{VATrackerStatus:ns});
    if(ns==="Out"){
      await gPost(t,lUrl("VA_Activity"),{Title:`${va.Name}-Out-${today()}`,ActivityType:"Absence",VAEmail:va.Email,VAName:va.Name,ActivityDate:new Date().toISOString(),StartTime:new Date().toISOString(),Status:"Out",MarkedByEmail:myEmail,MarkedByName:acct.name||myEmail});
      // Only move today's queued tasks to coverage
      const mv=queue.filter(q=>q.VAEmail.toLowerCase()===va.Email.toLowerCase());
      for(const task of mv){if(task._spId){try{await gPatch(t,iUrl("VA_Activity",task._spId),{Source:"Coverage",CoverageForEmail:va.Email,CoverageForName:va.Name});}catch(e){}}}
      const rest=queue.filter(q=>q.VAEmail.toLowerCase()!==va.Email.toLowerCase());
      mv.forEach(q=>{q.Source="Coverage";q.CoverageForEmail=va.Email;q.CoverageForName=va.Name;});
      setQueue(rest);setCovQ(p=>[...p,...mv]);fl(`${va.Name} marked OUT — ${mv.length} tasks to coverage`);    }else{
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
  if(!acct&&!guestToken)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700,letterSpacing:"0.04em"}}>{CONFIG.appName}</div><div style={{color:"rgba(255,255,255,0.4)",fontSize:9,letterSpacing:"0.08em",textTransform:"uppercase"}}>NewShire Property Management</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{...ss.card,textAlign:"center",maxWidth:360,padding:32}}><div style={{fontSize:36,marginBottom:12}}>⏱</div><div style={{fontSize:18,fontWeight:700,color:C.t2,marginBottom:6}}>VA Productivity Tracker</div><div style={{color:C.b4,marginBottom:20,fontSize:12}}>Sign in with your NewShire account.</div><button style={ss.btn(C.teal)} onClick={login}>Sign In with Microsoft</button>{authErr&&<div style={{color:C.er,marginTop:12,fontSize:11}}>{authErr}</div>}</div></div></div>);
  if(loading)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>{CONFIG.appName}</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{fontSize:16,color:C.b4}}>Loading...</div></div></div>);
  if(error||!role)return(<div style={ss.app}><div style={ss.hdr}><div style={{display:"flex",alignItems:"center",gap:10}}><div style={ss.logo}>V</div><div><div style={{color:"#fff",fontSize:14,fontWeight:700}}>{CONFIG.appName}</div></div></div></div><div style={{...ss.content,display:"flex",alignItems:"center",justifyContent:"center"}}><div style={{...ss.card,borderTop:`3px solid ${C.er}`,maxWidth:400,textAlign:"center",padding:32}}><div style={{fontSize:36,marginBottom:12}}>🚫</div><div style={{fontSize:16,fontWeight:700,color:C.er,marginBottom:6}}>{error==="access_denied"?"Access Denied":"Error"}</div><div style={{color:C.b4,fontSize:12}}>{error==="access_denied"?"Your account is not authorized.":error}</div><div style={{color:C.b3,fontSize:10,marginTop:12}}>Signed in as: {myEmail}</div></div></div></div>);

  // ── Tabs ──
  const isGuest=role==="guest";
  const TABS=[];
  if(!isGuest&&(isVA||isAdmin))TABS.push({n:"My Day",k:"myday"});
  if(!isGuest&&isMgr)TABS.push({n:"Manager View",k:"mgr",badge:myOverdue.length>0?myOverdue.length:0,badgeBg:C.er});
  TABS.push({n:"Dashboard",k:"dash"});
  if(!isGuest&&(isAdmin||isRegional))TABS.push({n:"Coaching",k:"coach"});
  TABS.push({n:"History",k:"hist"});
  if(!isGuest&&isAdmin)TABS.push({n:"Admin",k:"admin",badge:covQ.length+outVAs.length,badgeBg:C.t3});
  const ck=TABS[tab]?.k||TABS[0]?.k;

  // ── Header pills ──
  const refreshAgo=Math.floor((Date.now()-lastRefresh)/60000);
  const pills=[];
  if(!isGuest&&timers.length>0)pills.push({c:C.ok,n:timers.length,l:"timing"});
  if(!isGuest&&covQ.length>0)pills.push({c:C.wn,n:covQ.length,l:"coverage"});
  if(!isGuest&&myOverdue.length>0)pills.push({c:C.er,n:myOverdue.length,l:"overdue"});
  if(!isGuest&&outVAs.length>0)pills.push({c:C.er,n:outVAs.length,l:"out"});

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
        <div style={{textAlign:"right",flexShrink:0}}><div style={{fontSize:9,color:C.t4,marginBottom:2,cursor:"pointer"}} onClick={()=>{if(!guestToken)reload();}}>{refreshAgo<1?"Just refreshed":`Refreshed ${refreshAgo}m ago`} · ↻</div><strong style={{display:"block",fontSize:12,fontWeight:600,color:"#fff"}}>{isGuest?`${guestInfo?.name||"Guest"} · ${guestInfo?.company||""}`:acct?.name||myEmail}</strong><em style={{fontSize:9,fontStyle:"normal",color:isGuest?C.pu:C.gold,textTransform:"uppercase",letterSpacing:"0.07em"}}>{isGuest?"Guest (read-only)":role}</em></div>
      </div>
      {/* Tabs */}
      <div style={ss.tabs}>
        {TABS.map((t,i)=><button key={t.k} style={ss.tab(tab===i)} onClick={()=>setTab(i)}>{t.n}{t.badge>0&&<CountBadge n={t.badge} bg={t.badgeBg||C.er} fg="#fff"/>}</button>)}
      </div>
      {flash&&<div style={{background:C.gl0,borderBottom:`1px solid ${C.gold}`,padding:"7px 18px",fontSize:12,fontWeight:600,color:C.g2,textAlign:"center"}}>{flash}</div>}
      {/* Content */}
      <div style={ss.content}>
        {ck==="myday"&&<MyDayView data={data} role={role} myEmail={myEmail} myVa={myVa} myProps={myProps} queue={queue} covQ={covQ} shift={shift} timers={timers} tick={tick} overdue={myOverdue} config={data?.config} onClockIn={clockIn} onBreakStart={startBreak} onBreakEnd={endBreak} onClockOut={clockOut} onStartTimer={startTimer} onPause={pauseTimer} onResume={resumeTimer} onFinish={finishTimer} onCancel={cancelTimer} onClaimCov={claimCov} onAddTask={addTask} onDeleteTask={deleteTask} onLogInterruption={logInterruption} onSaveMetrics={saveDailyMetrics} onSendReview={sendReview} onResolveReview={resolveReview} onReplyToReview={replyToReview} onRequestTimeOff={requestTimeOff} reviews={data.activities.filter(a=>a.ActivityType==="Review"||a.ActivityType==="ReviewResponse")} isAdmin={isAdmin} fl={fl}/>}
        {ck==="mgr"&&<ManagerView data={data} onResolveReview={resolveReview} onReplyToReview={replyToReview} onAssignReviewProp={assignReviewProperty} myEmail={myEmail} acct={acct} myEmp={myEmp} mgrProps={isAdmin?data.properties:mgrProps} queue={queue} timers={timers} covQ={covQ} overdue={overdueTasks} onAddTask={addTask} getVA={getVAForProperty} isAdmin={isAdmin} isRegional={isRegional}/>}
        {ck==="dash"&&<DashboardView data={data} queue={queue} timers={timers} covQ={covQ} overdue={overdueTasks} dfFrom={dfFrom} dfTo={dfTo} setDfFrom={setDfFrom} setDfTo={setDfTo} isAdmin={isAdmin} role={role} mgrProps={mgrProps} myEmail={myEmail}/>}
        {ck==="coach"&&<CoachingView data={data} onSaveNote={saveCoachingNote}/>}
        {ck==="hist"&&<HistoryView data={data} role={role} myEmail={myEmail} isMgr={isMgr} mgrProps={mgrProps}/>}
        {ck==="admin"&&<AdminView data={data} myEmail={myEmail} myEmp={myEmp} acct={acct} config={data?.config} queue={queue} covQ={covQ} onToggleAbsence={toggleAbsence} onAssignTask={addTask} onUpdateConfig={updateConfig} onAssignProp={assignProp} onUnassignProp={unassignProp} onReassignVA={reassignPropertyVA} onApproveTimeOff={approveTimeOff} onDenyTimeOff={denyTimeOff} onLogCallout={logCallout} onLogEarlyDeparture={logEarlyDeparture} onEditShift={editShift} onDeleteShift={deleteShift} onClockInVA={clockIn} onAddGuest={addGuest} onDeactivateGuest={deactivateGuest} onAddProperty={addProperty} onEditProperty={editProperty} onEditEmployee={editEmployee} onDeleteTask={deleteTask}/>}
      </div>
    </div>
  );
}


// ══════════════════════════════════════════════════════
// MY DAY VIEW
// ══════════════════════════════════════════════════════
function MyDayView({data,role,myEmail,myVa,myProps,queue,covQ,shift,timers,tick,overdue,config,onClockIn,onBreakStart,onBreakEnd,onClockOut,onStartTimer,onPause,onResume,onFinish,onCancel,onClaimCov,onAddTask,onDeleteTask,onLogInterruption,onSaveMetrics,onSendReview,onResolveReview,onReplyToReview,onRequestTimeOff,reviews,isAdmin,fl}){
  const[showForm,setShowForm]=useState(false);const[fCat,setFCat]=useState("");const[fProp,setFProp]=useState("");const[fPri,setFPri]=useState("Normal");const[fDesc,setFDesc]=useState("");
  // Interruption state
  const[iType,setIType]=useState("Prospect Call");const[iProp,setIProp]=useState("");const[iDur,setIDur]=useState("");const[iNotes,setINotes]=useState("");const[iConvert,setIConvert]=useState(false);const[iCat,setICat]=useState("");const[iPri,setIPri]=useState("Normal");const[iTaskDesc,setITaskDesc]=useState("");
  // Daily metrics state
  // Daily metrics — restore from SP if exists, auto-save on change
  const[metrics,setMetrics]=useState(()=>{
    const existing=data?.activities.find(a=>a.ActivityType==="DailyMetrics"&&a.VAEmail?.toLowerCase()===myEmail&&a.ActivityDate?.slice(0,10)===today());
    if(existing?.Notes){try{const p=JSON.parse(existing.Notes);return{Leads:p.Leads||0,Apps:p.Apps||0,Showings:p.Showings||0,"Res. Comms":p["Res. Comms"]||0,"WOs In":p["WOs In"]||0,"WOs Upd.":p["WOs Upd."]||0,Renewals:p.Renewals||0,"Mgr Calls":p["Mgr Calls"]||0};}catch(e){}}
    return{Leads:0,Apps:0,Showings:0,"Res. Comms":0,"WOs In":0,"WOs Upd.":0,Renewals:0,"Mgr Calls":0};
  });const[mNotes,setMNotes]=useState(()=>{
    const existing=data?.activities.find(a=>a.ActivityType==="DailyMetrics"&&a.VAEmail?.toLowerCase()===myEmail&&a.ActivityDate?.slice(0,10)===today());
    if(existing?.Notes){try{return JSON.parse(existing.Notes).notes||"";}catch(e){}}return"";
  });const[mSaving,setMSaving]=useState(false);
  const metricsTimerRef=useRef(null);
  function updateMetric(key,delta){
    setMetrics(m=>{const updated={...m,[key]:Math.max(0,(m[key]||0)+delta)};
      // Debounce save — 800ms after last change
      if(metricsTimerRef.current)clearTimeout(metricsTimerRef.current);
      metricsTimerRef.current=setTimeout(()=>{setMSaving(true);onSaveMetrics(updated,mNotes).then(()=>setMSaving(false));},800);
      return updated;
    });
  }
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

  // Review request handler
  const[revTask,setRevTask]=useState(null);const[revNote,setRevNote]=useState("");

  // Active notices
  const activeNotices=(config?.notices||[]).filter(n=>{const td=today();return n.startDate<=td&&n.endDate>=td;});

  return(<div style={{maxWidth:1100,margin:"0 auto"}}>
    {/* Notice Banner */}
    {activeNotices.length>0&&<div style={{marginBottom:12}}>
      {activeNotices.map((n,i)=><div key={i} style={{background:C.gl,border:`1px solid ${C.gold}`,borderLeft:`4px solid ${C.gold}`,borderRadius:8,padding:"10px 14px",marginBottom:i<activeNotices.length-1?6:0,display:"flex",alignItems:"flex-start",gap:10}}>
        <span style={{fontSize:16,flexShrink:0}}>📢</span>
        <div style={{flex:1}}>
          <div style={{fontSize:12,fontWeight:700,color:C.g2}}>{n.title||"Notice"}</div>
          <div style={{fontSize:12,color:C.t2,marginTop:2,lineHeight:1.5}}>{n.text}</div>
          <div style={{fontSize:9,color:C.b4,marginTop:4}}>Posted by {n.createdBy||"Admin"} · {fD(n.startDate)} – {fD(n.endDate)}</div>
        </div>
      </div>)}
    </div>}

    {/* Shift Clock — full width */}
    <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:8,padding:14,marginBottom:12,boxShadow:"0 1px 3px rgba(28,55,64,0.07)"}}>
      {!shift?(<div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:11,fontWeight:700,color:C.b4,textTransform:"uppercase"}}>Shift Clock</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Not clocked in</div></div><button style={ss.btn(C.inf)} onClick={()=>onClockIn()}>☀ Clock In</button></div>
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

    {/* Review Panel — threaded, multiple per task */}
    {revTask&&(()=>{
      const taskKey=revTask._spId||revTask._localId||revTask.Title;
      const taskReviews=(reviews||[]).filter(r=>r.GroupId===taskKey).sort((a,b)=>(a.ActivityDate||"").localeCompare(b.ActivityDate||""));
      const pending=taskReviews.filter(r=>r.Status==="Pending").length;
      return<div style={{...ss.card,borderTop:`3px solid ${C.pu}`,marginBottom:12,padding:0,overflow:"hidden"}}>
        <div style={{padding:"12px 14px",borderBottom:`1px solid ${C.b1}`}}>
          <div style={{fontSize:13,fontWeight:700,color:C.pu}}>⚑ Manager Review — {revTask.Title}</div>
          <div style={{fontSize:10,color:C.b4,marginTop:2}}>{revTask.PropertyName} · {revTask.PMName||"PM"} · {taskReviews.length} request{taskReviews.length!==1?"s":""}{pending>0?` · ${pending} pending`:""}</div>
        </div>
        {/* Existing thread */}
        {taskReviews.length>0&&<div style={{padding:"10px 14px",background:C.pub,maxHeight:200,overflowY:"auto"}}>
          {taskReviews.map((r,i)=><div key={i} style={{display:"flex",gap:8,marginBottom:i<taskReviews.length-1?9:0}}>
            <div style={{display:"flex",flexDirection:"column",alignItems:"center"}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:r.Status==="Resolved"?C.ok:C.wn,marginTop:3,flexShrink:0}}/>
              {i<taskReviews.length-1&&<div style={{width:2,flex:1,background:"rgba(91,63,168,0.15)",marginTop:3,minHeight:12}}/>}
            </div>
            <div style={{flex:1}}>
              <div style={{fontSize:10,fontWeight:700,color:C.t2,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
                {r.VAName||"VA"} <span style={{fontSize:9,color:C.b4}}>{fD(r.ActivityDate)} {fT(r.ActivityDate)}</span>
                <Badge type={r.Status==="Resolved"?"ok":"wn"} dot={false}>{r.Status}</Badge>
              </div>
              <div style={{fontSize:11,color:C.b6,background:C.white,border:"1px solid rgba(91,63,168,0.15)",borderRadius:6,padding:"6px 8px",marginTop:4,lineHeight:1.5}}>{r.Notes}</div>
            </div>
          </div>)}
        </div>}
        {/* New review request */}
        <div style={{padding:"12px 14px"}}>
          <textarea style={{...ss.input,minHeight:50,marginBottom:8}} value={revNote} onChange={e=>setRevNote(e.target.value)} placeholder={taskReviews.length>0?"Add another review request...":"What do you need from the PM?"}/>
          <div style={{display:"flex",gap:6}}>
            <button style={{...ss.btn(C.pu),flex:1}} onClick={()=>{if(!revNote){fl("Enter a note");return;}onSendReview(revTask,revNote);setRevNote("");}}>⚑ {taskReviews.length>0?"Add Review Request":"Send Review Request"}</button>
            <button style={{...ss.btnO(C.b4,C.b2)}} onClick={()=>{setRevTask(null);setRevNote("");}}>Close</button>
          </div>
        </div>
      </div>;
    })()}

    {/* Two-column layout */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}} className="myday-grid">
    {/* LEFT COLUMN — Tasks */}
    <div style={{minWidth:0}}>

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
      {myTasks.map(t=>{const rc=(reviews||[]).filter(r=>r.GroupId===(t._spId||t._localId||t.Title)&&r.Status==="Pending").length;return<TaskRow key={t._localId} task={t} onStart={()=>onStartTimer(t)} onDelete={isAdmin?onDeleteTask:null} onReview={t2=>setRevTask(t2)} showVA={isAdmin} reviewCount={rc}/>})}
    </div>

    </div>{/* end left column */}

    {/* RIGHT COLUMN — Tools */}
    <div style={{minWidth:0}}>

    {/* My Reviews — always visible */}
    {(()=>{
      const myReviews=(reviews||[]).filter(a=>a.ActivityType==="Review"&&a.VAEmail?.toLowerCase()===myEmail);
      const allResponses=(reviews||[]).filter(a=>a.ActivityType==="ReviewResponse");
      const pending=myReviews.filter(r=>r.Status==="Pending");
      const responded=myReviews.filter(r=>r.Status==="Responded");
      const resolved=myReviews.filter(r=>r.Status==="Resolved").sort((a,b)=>(b.ActivityDate||"").localeCompare(a.ActivityDate||"")).slice(0,10);
      return<div style={{...ss.card,borderTop:`3px solid ${C.pu}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
          <div><div style={ss.cardT}>⚑ My Review Requests</div><div style={ss.cardS}>{pending.length} pending · {responded.length} responded · {resolved.length} resolved</div></div>
        </div>
        {!pending.length&&!responded.length&&!resolved.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:12}}>No pending reviews. Use the purple ⚑ button on any task to request PM input.</div>}
        {/* Responded — PM replied, VA can respond back or mark resolved */}
        {responded.length>0&&<div style={{marginBottom:10}}>
          <div style={{fontSize:10,fontWeight:700,color:C.inf,textTransform:"uppercase",marginBottom:6}}>PM Responded — Action Needed</div>
          {responded.map((r,i)=>{
            const thread=r.GroupId?allResponses.filter(a=>a.GroupId===r.GroupId).sort((a,b)=>(a.ActivityDate||"").localeCompare(b.ActivityDate||"")):[];
            const lastResponse=thread.length?thread[thread.length-1]:null;
            return<div key={i} style={{padding:"8px 10px",background:C.infb,border:`1px solid rgba(43,95,168,0.15)`,borderRadius:6,marginBottom:6}}>
              <div style={{fontSize:12,fontWeight:600,color:C.t2}}>{r.Title?.replace("Review-","")}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:2}}>{r.PropertyName} · {r.PMName||"PM"}</div>
              <div style={{fontSize:11,color:C.pu,background:C.pub,padding:"4px 8px",borderRadius:4,marginTop:4}}>You asked: "{r.Notes}"</div>
              {lastResponse&&<div style={{fontSize:11,color:C.ok,background:C.okb,border:`1px solid rgba(26,122,70,0.15)`,padding:"4px 8px",borderRadius:4,marginTop:4}}>
                <span style={{fontSize:9,fontWeight:700,color:C.ok}}>{lastResponse.AssignedByName||"PM"}:</span> "{lastResponse.Notes}"
              </div>}
              <ReviewReplyBox reviewId={r.id} onReply={onReplyToReview} label="VA"/>
            </div>;})}
        </div>}
        {/* Pending */}
        {pending.length>0&&<div style={{marginBottom:10}}>
          <div style={{fontSize:10,fontWeight:700,color:C.wn,textTransform:"uppercase",marginBottom:6}}>Awaiting PM Response</div>
          {pending.map((r,i)=>{
            const thread=r.GroupId?allResponses.filter(a=>a.GroupId===r.GroupId).sort((a,b)=>(a.ActivityDate||"").localeCompare(b.ActivityDate||"")):[];
            const lastResponse=thread.length?thread[thread.length-1]:null;
            return<div key={i} style={{padding:"8px 0",borderBottom:`1px solid ${C.b1}`}}>
              <div style={{fontSize:12,fontWeight:600,color:C.t2}}>{r.Title?.replace("Review-","")}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:2}}>{r.PropertyName} · {r.PMName||"PM"} · Sent {fD(r.ActivityDate)} {fT(r.ActivityDate)}</div>
              <div style={{fontSize:11,color:C.pu,background:C.pub,padding:"4px 8px",borderRadius:4,marginTop:4}}>"{r.Notes}"</div>
              {lastResponse&&<div style={{fontSize:11,color:C.ok,background:C.okb,border:`1px solid rgba(26,122,70,0.15)`,padding:"4px 8px",borderRadius:4,marginTop:4}}>
                <span style={{fontSize:9,fontWeight:700,color:C.ok}}>{lastResponse.AssignedByName||"PM"}:</span> "{lastResponse.Notes}"
              </div>}
              <ReviewReplyBox reviewId={r.id} onReply={onReplyToReview} label="VA"/>
            </div>;})}
        </div>}
        {/* Resolved */}
        {resolved.length>0&&<div>
          <div style={{fontSize:10,fontWeight:700,color:C.ok,textTransform:"uppercase",marginBottom:6}}>Resolved</div>
          {resolved.slice(0,5).map((r,i)=>{
            const thread=r.GroupId?allResponses.filter(a=>a.GroupId===r.GroupId).sort((a,b)=>(a.ActivityDate||"").localeCompare(b.ActivityDate||"")):[];
            const lastResponse=thread.length?thread[thread.length-1]:null;
            return<div key={i} style={{padding:"8px 0",borderBottom:`1px solid ${C.b1}`,opacity:0.7}}>
              <div style={{fontSize:12,fontWeight:600,color:C.t2}}>{r.Title?.replace("Review-","")}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:2}}>{r.PropertyName}</div>
              {lastResponse&&<div style={{fontSize:10,color:C.ok,marginTop:2}}>{lastResponse.AssignedByName}: "{lastResponse.Notes?.slice(0,80)}{lastResponse.Notes?.length>80?"...":""}"</div>}
            </div>;})}
        </div>}
      </div>;
    })()}

    {/* Interruption Logger */}
    <div style={{...ss.card,borderTop:`3px solid ${C.t3}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div><div style={ss.cardT}>📞 Log Interruption</div><div style={ss.cardS}>{data.activities.filter(a=>a.ActivityType==="Interruption"&&a.VAEmail?.toLowerCase()===myEmail&&a.ActivityDate?.slice(0,10)===today()).length} logged today</div></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6,marginBottom:10}}>
        {["Prospect Call","Resident Call","Vendor Call","Owner Call","Manager Call","Other"].map(t=>
          <button key={t} style={{padding:"9px 4px",fontSize:10,fontWeight:iType===t?700:600,border:`1.5px solid ${iType===t?C.t3:C.b2}`,borderRadius:6,background:iType===t?C.tl0:C.white,color:iType===t?C.t2:C.b4,cursor:"pointer",textAlign:"center",lineHeight:1.35,fontFamily:fnt}} onClick={()=>setIType(t)}>{t.replace(" ","\n")}</button>)}
      </div>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
        <div style={{flex:2,minWidth:120}}><label style={ss.label}>Property</label><select style={ss.select} value={iProp} onChange={e=>setIProp(e.target.value)}><option value="">General</option>{portProps.map(p=><option key={p.Title} value={p.Title}>{p.PropertyName}</option>)}</select></div>
        <div style={{flex:1,minWidth:75}}><label style={ss.label}>Duration (min)</label><input style={ss.input} type="number" value={iDur} onChange={e=>setIDur(e.target.value)} placeholder="0"/></div>
      </div>
      <input style={{...ss.input,marginBottom:10}} value={iNotes} onChange={e=>setINotes(e.target.value)} placeholder="Notes (optional)"/>
      {/* Convert to task toggle */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 12px",background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,marginBottom:iConvert?0:10,cursor:"pointer"}} onClick={()=>setIConvert(!iConvert)}>
        <div><div style={{fontSize:12,fontWeight:700,color:C.t2}}>➕ Convert to Task</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Add a follow-up task to your queue from this interruption</div></div>
        <div style={{width:40,height:22,background:iConvert?C.ok:C.b2,borderRadius:11,position:"relative",transition:"background 0.2s",flexShrink:0}}>
          <div style={{position:"absolute",width:16,height:16,background:"#fff",borderRadius:"50%",top:3,left:iConvert?21:3,transition:"left 0.2s",boxShadow:"0 1px 2px rgba(0,0,0,0.15)"}}/>
        </div>
      </div>
      {iConvert&&<div style={{background:C.tl0,border:`1px solid ${C.tl}`,borderRadius:"0 0 6px 6px",padding:12,marginBottom:10}}>
        <div style={{fontSize:11,fontWeight:700,color:C.t3,marginBottom:8,display:"flex",alignItems:"center",gap:5}}>📋 New Task from This Interruption</div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:8}}>
          <div style={{flex:1,minWidth:120}}><label style={ss.label}>Category</label><select style={ss.select} value={iCat} onChange={e=>setICat(e.target.value)}><option value="">Select...</option>{(config?.categories||[]).sort((a,b)=>a.name.localeCompare(b.name)).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div style={{flex:1,minWidth:100}}><label style={ss.label}>Priority</label><select style={ss.select} value={iPri} onChange={e=>setIPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
        </div>
        <input style={{...ss.input,marginBottom:8}} value={iTaskDesc} onChange={e=>setITaskDesc(e.target.value)} placeholder="Task description"/>
        <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"8px 11px",fontSize:11,color:C.b4}}><strong style={{color:C.t2}}>On submit:</strong> ✓ Interruption logged + ✓ Task added to your queue</div>
      </div>}
      <button style={{...ss.btn(C.teal),width:"100%"}} onClick={()=>{
        if(!iDur&&!iNotes){fl("Enter duration or notes");return;}
        const prop=iProp?data.properties.find(p=>p.Title===iProp):null;
        const cat=iCat?config?.categories?.find(c=>c.id===iCat):null;
        const int={type:iType,vaEmail:myEmail,vaName:myVa?.Name||myEmail,propertyId:iProp,propertyName:prop?prop.PropertyName:"General",duration:parseInt(iDur)||0,notes:iNotes};
        if(iConvert&&iTaskDesc&&iCat){
          int.task={Title:iTaskDesc,VAEmail:myEmail,VAName:myVa?.Name||myEmail,PropertyId:iProp||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Priority:iPri,Source:"Ad-Hoc",Notes:`From ${iType} interruption`};
        }
        onLogInterruption(int);
        setIDur("");setINotes("");setIConvert(false);setITaskDesc("");setICat("");setIPri("Normal");
      }}>{iConvert&&iTaskDesc?"Log Interruption & Add Task":"Log Interruption"}</button>
    </div>

    {/* Daily Activity Metrics — LIVE */}
    <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div><div style={ss.cardT}>📈 Daily Activity Metrics</div><div style={ss.cardS}>Tap + / − to track · Saves automatically</div></div>
        {mSaving?<Badge type="wn" dot={false}>Saving...</Badge>:<Badge type="ok" dot={false}>Live</Badge>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7,marginBottom:10}}>
        {Object.entries(metrics).map(([key,val])=>
          <div key={key} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:8,padding:"9px 5px",textAlign:"center"}}>
            <div style={{fontSize:8.5,fontWeight:700,color:C.t3,textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>{key}</div>
            <div style={{fontSize:24,fontWeight:700,fontFamily:mono,color:C.t2,lineHeight:1}}>{val}</div>
            <div style={{display:"flex",justifyContent:"center",gap:4,marginTop:6}}>
              <button style={{width:25,height:25,border:`1px solid ${C.b2}`,borderRadius:4,background:C.white,color:C.t2,fontSize:14,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:fnt}} onClick={()=>updateMetric(key,-1)}>−</button>
              <button style={{width:25,height:25,border:`1px solid ${C.b2}`,borderRadius:4,background:C.white,color:C.t2,fontSize:14,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:fnt}} onClick={()=>updateMetric(key,1)}>+</button>
            </div>
          </div>)}
      </div>
      <input style={{...ss.input,marginBottom:0}} value={mNotes} onChange={e=>{setMNotes(e.target.value);if(metricsTimerRef.current)clearTimeout(metricsTimerRef.current);metricsTimerRef.current=setTimeout(()=>{setMSaving(true);onSaveMetrics(metrics,e.target.value).then(()=>setMSaving(false));},1200);}} placeholder="Notes (optional — saves automatically)"/>
    </div>

    {/* Request Time Off */}
    {!isAdmin&&<TimeOffRequestCard myEmail={myEmail} myVa={myVa} onRequest={onRequestTimeOff}/>}

    {/* Upcoming Team Time Off — visible to ALL */}
    {(()=>{
      const upcoming=(data.timeOff||[]).filter(t=>t.Status==="Approved"&&t.StartDate?.slice(0,10)>=today()).sort((a,b)=>(a.StartDate||"").localeCompare(b.StartDate||""));
      return<div style={ss.card}>
        <div style={ss.cardT}>📅 Upcoming Team Time Off</div>
        {!upcoming.length?<div style={{textAlign:"center",padding:"12px 0",color:C.b4,fontSize:12}}>No upcoming time off scheduled.</div>
        :upcoming.slice(0,8).map((to,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${C.b1}`}}>
          <Avatar name={to.VAName} size={22}/>
          <div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{to.VAName}</div><div style={{fontSize:10,color:C.b4}}>{fD(to.StartDate)}{to.EndDate&&to.EndDate!==to.StartDate?` – ${fD(to.EndDate)}`:""} · {to.RequestType}</div></div>
          <Badge type={to.RequestType==="PTO"?"in":"wn"} dot={false}>{to.RequestType}</Badge>
        </div>)}
      </div>;
    })()}

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
    </div>{/* end right column */}
    </div>{/* end grid */}
  </div>);
}

// ── Time Off Request Card (VA view) ──
function TimeOffRequestCard({myEmail,myVa,onRequest}){
  const[open,setOpen]=useState(false);const[type,setType]=useState("PTO");const[start,setStart]=useState("");const[end,setEnd]=useState("");const[hours,setHours]=useState("");const[notes,setNotes]=useState("");
  return<div style={ss.card}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
      <div style={ss.cardT}>📅 Request Time Off</div>
      <button style={ss.btnO(C.t2,C.b2)} onClick={()=>setOpen(!open)}>{open?"Cancel":"Request"}</button>
    </div>
    {open&&<div style={{marginTop:12}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
        <div style={{flex:1,minWidth:100}}><label style={ss.label}>Type *</label><select style={ss.select} value={type} onChange={e=>setType(e.target.value)}><option>PTO</option><option>Personal</option><option>Sick</option></select></div>
        <div style={{flex:1,minWidth:120}}><label style={ss.label}>Start Date *</label><input style={ss.input} type="date" value={start} onChange={e=>setStart(e.target.value)}/></div>
        <div style={{flex:1,minWidth:120}}><label style={ss.label}>End Date</label><input style={ss.input} type="date" value={end} onChange={e=>setEnd(e.target.value)}/></div>
        <div style={{minWidth:70}}><label style={ss.label}>Hours</label><input style={ss.input} type="number" value={hours} onChange={e=>setHours(e.target.value)} placeholder="8"/></div>
      </div>
      <label style={ss.label}>Notes</label>
      <textarea style={{...ss.input,minHeight:40,marginBottom:10}} value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Reason or details (optional)"/>
      <button style={{...ss.btn(C.teal),width:"100%"}} onClick={()=>{if(!start)return;onRequest({vaEmail:myEmail,vaName:myVa?.Name||myEmail,type,startDate:start,endDate:end||start,hours:parseInt(hours)||0,notes});setOpen(false);setStart("");setEnd("");setHours("");setNotes("");}}>Submit Request</button>
    </div>}
  </div>;
}

// ── Review Reply Box (inline reply + resolve for both VA and PM) ──
function ReviewReplyBox({reviewId,onReply,label}){
  const[open,setOpen]=useState(false);const[note,setNote]=useState("");
  if(!open)return<div style={{display:"flex",gap:5,marginTop:6}}>
    <button style={{...ss.btn(C.teal),...ss.xs}} onClick={()=>setOpen(true)}>Reply</button>
    <button style={{...ss.btn(C.ok),...ss.xs}} onClick={()=>onReply(reviewId,"","Resolved")}>✓ Mark Resolved</button>
  </div>;
  return<div style={{marginTop:6}}>
    <textarea style={{...ss.input,minHeight:40,marginBottom:5,fontSize:11}} value={note} onChange={e=>setNote(e.target.value)} placeholder="Type your reply..."/>
    <div style={{display:"flex",gap:5}}>
      <button style={{...ss.btn(C.teal),...ss.xs,flex:1}} onClick={()=>{if(!note.trim())return;onReply(reviewId,note,"Pending");setNote("");setOpen(false);}}>Send Reply</button>
      <button style={{...ss.btn(C.ok),...ss.xs,flex:1}} onClick={()=>{onReply(reviewId,note||"","Resolved");setNote("");setOpen(false);}}>Reply & Resolve</button>
      <button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>{setNote("");setOpen(false);}}>Cancel</button>
    </div>
  </div>;
}

// ── Task Row ──
function TaskRow({task,onStart,onDelete,onReview,showVA,isOverdue,reviewCount=0}){
  return(<div style={{display:"flex",alignItems:"flex-start",gap:9,padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
    <span style={{fontSize:15,width:22,textAlign:"center",flexShrink:0,paddingTop:1}}>{catIcon[task.Category]||"📁"}</span>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontSize:12,fontWeight:600,color:C.t2,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap",lineHeight:1.4}}>{task.Title}
        {isOverdue&&<Badge type="er" dot={false}>OVERDUE — {fD(task.ActivityDate)}</Badge>}
        {task.CoverageForName&&<Badge type="wn" dot={false}>Coverage</Badge>}
        {task.Priority==="Urgent"&&<Badge type="er" dot={false}>Urgent</Badge>}
        {task.Priority==="High"&&<Badge type="wn" dot={false}>High</Badge>}
        {reviewCount>0&&<Badge type="pu" dot={false}>{reviewCount} review{reviewCount>1?"s":""}</Badge>}
      </div>
      <div style={{fontSize:10,color:C.b4,marginTop:3,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
        {showVA&&task.VAName&&<>{task.VAName}<Dot/></>}
        {task.PropertyName}<Dot/>{task.Source}
        {task.Notes&&<><Dot/>💬 {task.Notes}</>}
      </div>
    </div>
    <div style={{display:"flex",gap:4,flexShrink:0,flexWrap:"wrap",alignItems:"flex-start"}}>
      {onStart&&<button style={{...ss.btn(C.ok),...ss.xs}} onClick={onStart}>▶ Start</button>}
      {onReview&&<button style={{...ss.btn(C.pu),...ss.xs}} onClick={()=>onReview(task)}>⚑</button>}
      {onDelete&&<button style={{...ss.btnO(C.er,`rgba(184,59,42,0.3)`),...ss.xs}} onClick={()=>{if(window.confirm("Remove?"))onDelete(task);}}>✕</button>}
    </div>
  </div>);
}

// ══════════════════════════════════════════════════════
// MANAGER VIEW
// ══════════════════════════════════════════════════════
function ManagerView({data,myEmail,myEmp,mgrProps,queue,timers,covQ,overdue,onAddTask,getVA,isAdmin,isRegional,onResolveReview,onReplyToReview,onAssignReviewProp,acct}){
  const[selProp,setSelProp]=useState("");const[tCat,setTCat]=useState("");const[tDesc,setTDesc]=useState("");const[tPri,setTPri]=useState("Normal");const[tNotes,setTNotes]=useState("");
  const[expandedThreads,setExpandedThreads]=useState(new Set());
  const cats=data?.config?.categories||[];
  function handleSubmit(){if(!selProp||!tCat||!tDesc)return;const prop=data.properties.find(p=>p.Title===selProp);const va=getVA(selProp);if(!va){alert("No VA assigned.");return;}const cat=cats.find(c=>c.id===tCat);
    onAddTask({Title:tDesc,VAEmail:va.Email,VAName:va.Name,PropertyId:selProp,PropertyName:prop?.PropertyName||"",PMName:myEmp?.Name||myEmail,Category:cat?.name||"Admin/Other",Priority:tPri,Source:"Assigned",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||myEmail,Notes:tNotes});setTDesc("");setTCat("");setTPri("Normal");setTNotes("");}

  // Active timers on my properties
  const propIds=new Set(mgrProps.map(p=>p.Title));
  const activeOnMine=timers.filter(t=>propIds.has(t.PropertyId));
  const overdueOnMine=overdue.filter(t=>propIds.has(t.PropertyId));

  // Pending reviews — admin/regional see ALL, managers see their properties + unassigned
  const pendingReviews=data.activities.filter(a=>a.ActivityType==="Review"&&(a.Status==="Pending"||a.Status==="Responded")&&(isAdmin||isRegional||propIds.has(a.PropertyId)||!a.PropertyId));

  return(<div>
    {/* Review Inbox */}
    <div style={{...ss.card,borderTop:`3px solid ${C.pu}`,padding:0,overflow:"hidden"}}>
      <div style={{padding:"13px 14px 11px",borderBottom:`1px solid ${C.b1}`}}>
        <div style={{...ss.cardT,color:C.pu}}>⚑ Needs Your Review — {pendingReviews.length} item{pendingReviews.length!==1?"s":""}</div>
        <div style={ss.cardS}>VAs flagged these tasks and are waiting on your response</div>
      </div>
      {pendingReviews.map((r,i)=>{
        // Find all messages in this thread (same GroupId)
        const thread=r.GroupId?data.activities.filter(a=>(a.ActivityType==="Review"||a.ActivityType==="ReviewResponse")&&a.GroupId===r.GroupId).sort((a,b)=>(a.ActivityDate||"").localeCompare(b.ActivityDate||"")):[];
        const needsAssign=!r.PropertyId;
        const hasThread=thread.length>1;
        const isExpanded=expandedThreads.has(r.id);
        const lastResponse=thread.length>1?thread[thread.length-1]:null;
        return<div key={i} style={{padding:"12px 14px",borderBottom:`1px solid ${C.b1}`,background:needsAssign?C.wnb:"transparent"}}>
          <div style={{display:"flex",alignItems:"flex-start",gap:9,marginBottom:8}}>
            <Avatar name={r.VAName} size={30} colorIdx={4}/>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:700,color:C.t2}}>{r.VAName} · {r.PropertyName||"⚠ No property assigned"}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:1}}>{r.Category||"Review"} · {fD(r.ActivityDate)} {fT(r.ActivityDate)}{hasThread?<span style={{color:C.pu,fontWeight:600}}> · {thread.length} messages</span>:""}</div>
            </div>
          </div>
          {needsAssign&&<div style={{display:"flex",gap:6,alignItems:"center",marginBottom:8,padding:"6px 8px",background:C.white,border:`1px solid ${C.wn}`,borderRadius:4}}>
            <span style={{fontSize:10,fontWeight:700,color:C.wn,whiteSpace:"nowrap"}}>Assign to:</span>
            <select style={{...ss.select,flex:1,padding:"4px 8px",fontSize:11}} onChange={e=>{if(!e.target.value)return;const p=data.properties.find(pr=>pr.Title===e.target.value);if(p)onAssignReviewProp(r.id,p.Title,p.PropertyName);}} defaultValue="">
              <option value="">Select property...</option>
              {data.properties.map(p=><option key={p.Title} value={p.Title}>{p.PropertyName}</option>)}
            </select>
          </div>}
          {/* Thread display */}
          {!hasThread?<div style={{fontSize:11,color:C.pu,background:C.pub,padding:"6px 8px",borderRadius:4,marginBottom:7,lineHeight:1.5}}>"{r.Notes}"</div>
          :<div style={{marginBottom:7}}>
            {/* Always show the original request */}
            <div style={{fontSize:11,color:C.pu,background:C.pub,padding:"6px 8px",borderRadius:4,lineHeight:1.5,marginBottom:4}}>
              <span style={{fontSize:9,fontWeight:700,color:C.pu}}>{r.VAName} · {fD(r.ActivityDate)} {fT(r.ActivityDate)}</span><br/>"{r.Notes}"
            </div>
            {/* Expand/collapse for thread */}
            {!isExpanded&&<div>
              {lastResponse&&<div style={{fontSize:11,color:lastResponse.ActivityType==="ReviewResponse"?C.ok:C.pu,background:lastResponse.ActivityType==="ReviewResponse"?C.okb:C.pub,padding:"6px 8px",borderRadius:4,lineHeight:1.5,marginBottom:4}}>
                <span style={{fontSize:9,fontWeight:700}}>{lastResponse.AssignedByName||lastResponse.VAName} · {fD(lastResponse.ActivityDate)} {fT(lastResponse.ActivityDate)}</span><br/>"{lastResponse.Notes}"
              </div>}
              <button style={{fontSize:10,fontWeight:600,color:C.pu,background:"none",border:"none",cursor:"pointer",padding:"2px 0",textDecoration:"underline"}} onClick={()=>setExpandedThreads(p=>{const n=new Set(p);n.add(r.id);return n;})}>Show full thread ({thread.length} messages)</button>
            </div>}
            {isExpanded&&<div>
              {thread.slice(1).map((msg,mi)=>{
                const isResponse=msg.ActivityType==="ReviewResponse";
                const isVA=msg.VAEmail?.toLowerCase()===r.VAEmail?.toLowerCase()&&!isResponse;
                return<div key={mi} style={{fontSize:11,color:isResponse?C.ok:C.pu,background:isResponse?C.okb:C.pub,padding:"6px 8px",borderRadius:4,lineHeight:1.5,marginBottom:4,borderLeft:`3px solid ${isResponse?C.ok:C.pu}`}}>
                  <span style={{fontSize:9,fontWeight:700}}>{msg.AssignedByName||msg.VAName} · {fD(msg.ActivityDate)} {fT(msg.ActivityDate)} · <Badge type={isResponse?"ok":"pu"} dot={false}>{isResponse?"Response":"Request"}</Badge></span><br/>"{msg.Notes}"
                </div>;
              })}
              <button style={{fontSize:10,fontWeight:600,color:C.b4,background:"none",border:"none",cursor:"pointer",padding:"2px 0",textDecoration:"underline"}} onClick={()=>setExpandedThreads(p=>{const n=new Set(p);n.delete(r.id);return n;})}>Collapse thread</button>
            </div>}
          </div>}
          <ReviewReplyBox reviewId={r.id} onReply={onReplyToReview} label="PM"/>
        </div>;})}
      {!pendingReviews.length&&<div style={{padding:"20px 14px",textAlign:"center",color:C.b4,fontSize:12}}>No pending review requests from your VAs.</div>}
    </div>

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
// Helper: calculate live shift minutes including active (clocked-in) shifts
function liveShiftMin(shifts,vaEmail){
  let total=shifts.reduce((s,a)=>s+(a.WorkMinutes||0),0);
  // Add elapsed time for active shifts (no EndTime)
  shifts.filter(a=>a.StartTime&&!a.EndTime&&(!vaEmail||a.VAEmail?.toLowerCase()===vaEmail)).forEach(a=>{
    total+=Math.max(0,Math.round((Date.now()-new Date(a.StartTime).getTime())/6e4));
  });
  return total;
}

function DashboardView({data,queue,timers,covQ,overdue,dfFrom,dfTo,setDfFrom,setDfTo,isAdmin,role,mgrProps,myEmail}){
  if(!data)return null;
  const propIds=role==="manager"?new Set(mgrProps.map(p=>p.Title)):null;
  const vaEmail=role==="va"?myEmail?.toLowerCase():null;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&inRange(a.ActivityDate,dfFrom,dfTo)&&(!propIds||propIds.has(a.PropertyId))&&(!vaEmail||a.VAEmail?.toLowerCase()===vaEmail));
  const done=tasks.filter(a=>a.Status==="Completed");const blocked=tasks.filter(a=>a.Status==="Blocked");const inc=tasks.filter(a=>a.Status==="Incomplete");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&inRange(a.ActivityDate,dfFrom,dfTo)&&(!vaEmail||a.VAEmail?.toLowerCase()===vaEmail));
  const shiftMin=liveShiftMin(shifts);const taskMin=done.reduce((s,a)=>s+(a.DurationMin||0),0);
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
    {/* Team Status — who's clocked in */}
    <div style={{...ss.card,marginBottom:12}}>
      <div style={{...ss.cardT,marginBottom:10}}>👥 Team Status — Today</div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
        {filteredVAs.map((va,vi)=>{
          const isOut=va.VATrackerStatus==="Out";
          const todayShifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.VAEmail?.toLowerCase()===va.Email.toLowerCase()&&a.ActivityDate?.slice(0,10)===today()&&a.Status!=="Deleted");
          const activeShift=todayShifts.find(s=>s.StartTime&&!s.EndTime);
          const completedShift=todayShifts.find(s=>s.StartTime&&s.EndTime);
          const activeTimer=timers.find(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());
          let status,statusColor,statusBg,detail;
          if(isOut){status="OUT";statusColor=C.er;statusBg=C.erb;detail="Marked out";}
          else if(activeShift){status=activeTimer?"Working":"Clocked In";statusColor=C.ok;statusBg=C.okb;detail=`Since ${fT(activeShift.StartTime)}${activeTimer?` · ${activeTimer.Title}`:""}`;}
          else if(completedShift){status="Clocked Out";statusColor=C.b4;statusBg=C.b1;detail=`${fT(completedShift.StartTime)} – ${fT(completedShift.EndTime)} · ${fM(completedShift.WorkMinutes)}`;}
          else{status="Not Clocked In";statusColor=C.wn;statusBg=C.wnb;detail="No shift started today";}
          return<div key={va.Email} style={{flex:"1 1 160px",minWidth:150,background:statusBg,border:`1px solid ${statusColor}22`,borderRadius:8,padding:"10px 12px"}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
              <Avatar name={va.Name} size={28} colorIdx={vi} isOut={isOut}/>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,fontWeight:700,color:isOut?C.er:C.t2}}>{va.Name}</div>
                <div style={{fontSize:10,color:C.b4,marginTop:1}}>{detail}</div>
              </div>
            </div>
            <Badge type={status==="Working"||status==="Clocked In"?"ok":status==="OUT"?"er":status==="Not Clocked In"?"wn":"ne"} dot={status==="Working"||status==="Clocked In"}>{status}</Badge>
          </div>;})}
      </div>
    </div>
    {/* Daily Activity Metrics Rollup */}
    {(isAdmin||role==="regional")&&(()=>{
      const td=today();
      const metricRecords=data.activities.filter(a=>a.ActivityType==="DailyMetrics"&&a.ActivityDate?.slice(0,10)===td&&a.Status!=="Deleted");
      const metricKeys=["Leads","Apps","Showings","Res. Comms","WOs In","WOs Upd.","Renewals","Mgr Calls"];
      const byVA=filteredVAs.map(va=>{
        const rec=metricRecords.find(m=>m.VAEmail?.toLowerCase()===va.Email.toLowerCase());
        let parsed=null;
        if(rec?.Notes){try{parsed=JSON.parse(rec.Notes);}catch(e){}}
        return{va,submitted:!!rec,metrics:parsed};
      });
      const totals={};metricKeys.forEach(k=>{totals[k]=byVA.reduce((s,v)=>s+((v.metrics?.[k])||0),0);});
      const submittedCount=byVA.filter(v=>v.submitted).length;
      return<div style={{...ss.card,marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
          <div><div style={ss.cardT}>📊 Daily Activity Metrics — Today</div><div style={ss.cardS}>{submittedCount}/{filteredVAs.length} VAs tracking · Updates on page refresh</div></div>
          {submittedCount<filteredVAs.length&&<Badge type="wn" dot={false}>{filteredVAs.length-submittedCount} not started</Badge>}
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><th style={ss.th}>VA</th>{metricKeys.map(k=><th key={k} style={{...ss.th,textAlign:"center",fontSize:10}}>{k}</th>)}<th style={{...ss.th,textAlign:"center"}}>Status</th></tr></thead>
            <tbody>
              {byVA.map(({va,submitted,metrics:m})=><tr key={va.Email} style={{opacity:submitted?1:0.5}}>
                <td style={{...ss.td,fontWeight:600,fontSize:12}}>{va.Name}</td>
                {metricKeys.map(k=><td key={k} style={{...ss.td,textAlign:"center",fontFamily:mono,fontSize:12,fontWeight:600,color:(m?.[k]||0)>0?C.t2:C.b3}}>{m?.[k]||0}</td>)}
                <td style={{...ss.td,textAlign:"center"}}>{submitted?<Badge type="ok" dot={false}>Live</Badge>:<Badge type="wn" dot={false}>—</Badge>}</td>
              </tr>)}
              <tr style={{background:C.tl00}}>
                <td style={{...ss.td,fontWeight:700,fontSize:12,color:C.t2}}>TOTAL</td>
                {metricKeys.map(k=><td key={k} style={{...ss.td,textAlign:"center",fontFamily:mono,fontSize:13,fontWeight:700,color:totals[k]>0?C.teal:C.b3}}>{totals[k]}</td>)}
                <td style={{...ss.td,textAlign:"center"}}><span style={{fontSize:10,color:C.b4}}>{submittedCount}/{filteredVAs.length}</span></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>;
    })()}
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
            const vS=shifts.filter(a=>a.VAEmail===va.Email);const vSm=liveShiftMin(vS,va.Email.toLowerCase());const vTm=vD.reduce((s,a)=>s+(a.DurationMin||0),0);
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
function CoachingView({data,onSaveNote}){
  const[selVa,setSelVa]=useState("");const[period,setPeriod]=useState(7);const[cNote,setCNote]=useState("");
  if(!data)return null;
  const va=data.vas.find(v=>v.Email?.toLowerCase()===(selVa||data.vas[0]?.Email||"").toLowerCase());
  const vaEmail=va?.Email?.toLowerCase()||"";
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=period);
  const done=tasks.filter(t=>t.Status==="Completed");const blocked=tasks.filter(t=>t.Status==="Blocked");const inc=tasks.filter(t=>t.Status==="Incomplete");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=period);
  const shiftMin=liveShiftMin(shifts,vaEmail);const taskMin=done.reduce((s,a)=>s+(a.DurationMin||0),0);
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

    {/* Coaching Notes */}
    <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
      <div style={ss.cardT}>Coaching Notes <span style={{fontSize:10,fontWeight:400,color:C.b4}}>— admin only, not visible to VA</span></div>
      <div style={{marginTop:10}}>
        <textarea style={{...ss.input,minHeight:70,marginBottom:8}} value={cNote} onChange={e=>setCNote(e.target.value)} placeholder={`Add a coaching note for ${va?.Name||"this VA"}...`}/>
        <button style={{...ss.btn(C.teal),...ss.sm}} onClick={()=>{if(!cNote.trim())return;onSaveNote(vaEmail,va?.Name||"",cNote);setCNote("");}}>Save Note</button>
      </div>
      {(()=>{const notes=data.activities.filter(a=>a.ActivityType==="CoachingNote"&&a.VAEmail?.toLowerCase()===vaEmail).sort((a,b)=>(b.ActivityDate||"").localeCompare(a.ActivityDate||""));
        if(!notes.length)return<div style={{marginTop:10,fontSize:11,color:C.b4,fontStyle:"italic"}}>No coaching notes yet.</div>;
        return<div style={{marginTop:12}}><div style={{fontSize:10,fontWeight:700,color:C.b4,textTransform:"uppercase",marginBottom:6}}>History</div>
          {notes.slice(0,20).map((n,i)=><div key={i} style={{padding:"8px 0",borderBottom:`1px solid ${C.b1}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
              <span style={{fontSize:10,fontWeight:700,color:C.t2}}>{n.AssignedByName||"Admin"}</span>
              <span style={{fontSize:9,color:C.b4}}>{fD(n.ActivityDate)} {fT(n.ActivityDate)}</span>
            </div>
            <div style={{fontSize:12,color:C.b6,lineHeight:1.5,whiteSpace:"pre-wrap"}}>{n.Notes}</div>
          </div>)}</div>;
      })()}
    </div>
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
// TIME OFF ADMIN COMPONENT
// ══════════════════════════════════════════════════════
function TimeOffAdmin({data,myEmail,acct,myEmp,onApprove,onDeny,onLogCallout,onLogEarlyDeparture}){
  const[appId,setAppId]=useState(null);const[paid,setPaid]=useState("TBD");const[aNotes,setANotes]=useState("");const[supConf,setSupConf]=useState(false);
  const[coVa,setCoVa]=useState("");const[coNotes,setCoNotes]=useState("");
  const[edVa,setEdVa]=useState("");const[edTime,setEdTime]=useState("");const[edNotes,setEdNotes]=useState("");
  const[view,setView]=useState("pending");
  if(!data)return(<div style={ss.card}><div style={{color:C.b4,fontSize:12}}>Loading time off data...</div></div>);
  const timeOff=data.timeOff||[];
  const vas=data.vas||[];
  const pending=timeOff.filter(t=>t.Status==="Pending");
  const approved=timeOff.filter(t=>t.Status==="Approved").sort((a,b)=>(a.StartDate||"").localeCompare(b.StartDate||""));
  const upcoming=approved.filter(t=>t.StartDate?.slice(0,10)>=today());
  const past=timeOff.filter(t=>t.Status!=="Pending").sort((a,b)=>(b.StartDate||"").localeCompare(a.StartDate||""));
  const typeColors={PTO:C.inf,Personal:C.pu,Sick:C.wn,Callout:C.er,"Early Departure":C.er};

  return<div>
    {/* Pending Requests */}
    <div style={{...ss.card,borderTop:`3px solid ${C.wn}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div><div style={ss.cardT}>📋 Time Off Requests{pending.length>0?` — ${pending.length} pending`:""}</div><div style={ss.cardS}>Approve, deny, and track paid/unpaid status</div></div>
      </div>
      {!pending.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:12}}>No pending requests.</div>}
      {pending.map(to=>{const isApp=appId===to.id;
        return<div key={to.id} style={{padding:"12px 0",borderBottom:`1px solid ${C.b1}`}}>
          <div style={{display:"flex",alignItems:"flex-start",gap:10}}>
            <Avatar name={to.VAName} size={30}/>
            <div style={{flex:1}}>
              <div style={{fontSize:13,fontWeight:700,color:C.t2}}>{to.VAName}</div>
              <div style={{fontSize:11,color:C.b4,marginTop:2}}>
                <Badge type={to.RequestType==="PTO"?"in":to.RequestType==="Sick"?"wn":"pu"} dot={false}>{to.RequestType}</Badge>
                {" "}{fD(to.StartDate)}{to.EndDate&&to.EndDate!==to.StartDate?` – ${fD(to.EndDate)}`:""}{to.HoursRequested?` · ${to.HoursRequested}h`:""} · Requested {fD(to.RequestedDate)}
              </div>
              {to.VANotes&&<div style={{fontSize:11,color:C.b6,background:C.tl00,padding:"4px 8px",borderRadius:4,marginTop:4}}>"{to.VANotes}"</div>}
            </div>
            {!isApp&&<div style={{display:"flex",gap:4}}>
              <button style={{...ss.btn(C.ok),...ss.xs}} onClick={()=>{setAppId(to.id);setPaid("TBD");setANotes("");setSupConf(false);}}>Review</button>
              <button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>{const n=prompt("Reason for denial:");if(n!==null)onDeny(to.id,n);}}>Deny</button>
            </div>}
          </div>
          {isApp&&<div style={{marginTop:10,background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:12}}>
            <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:8}}>
              <div style={{flex:1,minWidth:120}}><label style={ss.label}>Paid Status *</label><select style={ss.select} value={paid} onChange={e=>setPaid(e.target.value)}><option value="TBD">TBD — Decide later</option><option value="Paid">Paid</option><option value="Unpaid">Unpaid</option><option value="Partial">Partial</option></select></div>
              <div style={{flex:0,display:"flex",alignItems:"flex-end",gap:6,paddingBottom:2}}>
                <label style={{fontSize:11,fontWeight:600,color:C.t2,display:"flex",alignItems:"center",gap:6,cursor:"pointer"}}><input type="checkbox" checked={supConf} onChange={e=>setSupConf(e.target.checked)}/> Supervisor confirmed paid/unpaid</label>
              </div>
            </div>
            <label style={ss.label}>Admin Notes</label>
            <textarea style={{...ss.input,minHeight:40,marginBottom:8}} value={aNotes} onChange={e=>setANotes(e.target.value)} placeholder="Notes on approval, supervisor confirmation, etc."/>
            <div style={{display:"flex",gap:6}}>
              <button style={{...ss.btn(C.ok),...ss.xs,flex:1}} onClick={()=>{onApprove(to.id,paid,aNotes,supConf);setAppId(null);}}>✓ Approve</button>
              <button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>{const n=prompt("Reason?");if(n!==null){onDeny(to.id,n);setAppId(null);}}}>Deny</button>
              <button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setAppId(null)}>Cancel</button>
            </div>
          </div>}
        </div>;})}
    </div>

    {/* Quick Log — Callout & Early Departure */}
    <div style={{...ss.card,borderTop:`3px solid ${C.er}`}}>
      <div style={ss.cardT}>⚡ Quick Log — Same Day</div><div style={{...ss.cardS,marginBottom:12}}>Log callouts and early departures for today</div>
      <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:200,background:C.erb,border:`1px solid rgba(184,59,42,0.15)`,borderRadius:6,padding:12}}>
          <div style={{fontSize:12,fontWeight:700,color:C.er,marginBottom:8}}>📞 Callout</div>
          <div style={{marginBottom:8}}><label style={ss.label}>VA</label><select style={ss.select} value={coVa} onChange={e=>setCoVa(e.target.value)}><option value="">Select...</option>{vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
          <input style={{...ss.input,marginBottom:8}} value={coNotes} onChange={e=>setCoNotes(e.target.value)} placeholder="Reason (optional)"/>
          <button style={{...ss.btn(C.er),...ss.xs,width:"100%"}} onClick={()=>{if(!coVa)return;const va=vas.find(v=>v.Email===coVa);onLogCallout(coVa,va?.Name||coVa,coNotes);setCoVa("");setCoNotes("");}}>Log Callout</button>
        </div>
        <div style={{flex:1,minWidth:200,background:C.wnb,border:`1px solid rgba(168,111,8,0.15)`,borderRadius:6,padding:12}}>
          <div style={{fontSize:12,fontWeight:700,color:C.wn,marginBottom:8}}>🕐 Early Departure</div>
          <div style={{marginBottom:8}}><label style={ss.label}>VA</label><select style={ss.select} value={edVa} onChange={e=>setEdVa(e.target.value)}><option value="">Select...</option>{vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
          <div style={{marginBottom:8}}><label style={ss.label}>Departure Time</label><input style={ss.input} type="time" value={edTime} onChange={e=>setEdTime(e.target.value)}/></div>
          <input style={{...ss.input,marginBottom:8}} value={edNotes} onChange={e=>setEdNotes(e.target.value)} placeholder="Reason (optional)"/>
          <button style={{...ss.btn(C.wn),...ss.xs,width:"100%"}} onClick={()=>{if(!edVa||!edTime)return;const va=vas.find(v=>v.Email===edVa);onLogEarlyDeparture(edVa,va?.Name||edVa,edTime,edNotes);setEdVa("");setEdTime("");setEdNotes("");}}>Log Early Departure</button>
        </div>
      </div>
    </div>

    {/* Upcoming Approved */}
    <div style={ss.card}>
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        {[{k:"pending",l:"Upcoming"},{k:"history",l:"History"}].map(v=><button key={v.k} style={view===v.k?ss.btn(C.teal):ss.btnO(C.t2,C.b2)} onClick={()=>setView(v.k)}>{v.l}</button>)}
      </div>
      {view==="pending"&&<div>
        {!upcoming.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:12}}>No upcoming approved time off.</div>}
        {upcoming.map(to=><div key={to.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
          <Avatar name={to.VAName} size={24}/>
          <div style={{flex:1}}>
            <div style={{fontSize:12,fontWeight:600,color:C.t2}}>{to.VAName}</div>
            <div style={{fontSize:10,color:C.b4}}>{fD(to.StartDate)}{to.EndDate&&to.EndDate!==to.StartDate?` – ${fD(to.EndDate)}`:""} · <Badge type={to.RequestType==="PTO"?"in":to.RequestType==="Sick"?"wn":"pu"} dot={false}>{to.RequestType}</Badge></div>
          </div>
          <Badge type={to.PaidStatus==="Paid"?"ok":to.PaidStatus==="Unpaid"?"er":"wn"} dot={false}>{to.PaidStatus}</Badge>
          {to.SupervisorConfirmed&&<span style={{fontSize:9,fontWeight:700,padding:"2px 6px",background:C.ok,color:"#fff",borderRadius:99}}>✓ Confirmed</span>}
        </div>)}
      </div>}
      {view==="history"&&<div style={{overflowX:"auto"}}>
        <div style={{borderRadius:8,border:`1px solid ${C.b1}`,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["VA","Type","Dates","Status","Paid","Supervisor","Notes"].map(h=><th key={h} style={ss.th}>{h}</th>)}</tr></thead><tbody>
          {past.slice(0,30).map((to,i)=><tr key={i}>
            <td style={{...ss.td,fontWeight:600}}>{to.VAName}</td>
            <td style={ss.td}><Badge type={to.RequestType==="Callout"||to.RequestType==="Early Departure"?"er":to.RequestType==="PTO"?"in":"wn"} dot={false}>{to.RequestType}</Badge></td>
            <td style={{...ss.td,fontSize:11}}>{fD(to.StartDate)}{to.EndDate&&to.EndDate!==to.StartDate?` – ${fD(to.EndDate)}`:""}{to.DepartureTime?` @ ${to.DepartureTime}`:""}</td>
            <td style={ss.td}><Badge type={to.Status==="Approved"?"ok":to.Status==="Denied"?"er":"wn"} dot={false}>{to.Status}</Badge></td>
            <td style={ss.td}><Badge type={to.PaidStatus==="Paid"?"ok":to.PaidStatus==="Unpaid"?"er":"wn"} dot={false}>{to.PaidStatus||"TBD"}</Badge></td>
            <td style={ss.td}>{to.SupervisorConfirmed?<span style={{color:C.ok,fontWeight:700}}>✓</span>:<span style={{color:C.b4}}>—</span>}</td>
            <td style={{...ss.td,fontSize:11,maxWidth:200}}>{to.AdminNotes||to.VANotes||"—"}</td>
          </tr>)}
        </tbody></table></div>
      </div>}
    </div>
  </div>;
}

// ══════════════════════════════════════════════════════
// SHIFT MANAGER COMPONENT
// ══════════════════════════════════════════════════════
function ShiftManager({data,onEditShift,onDeleteShift,onClockInVA}){
  const[editId,setEditId]=useState(null);const[eIn,setEIn]=useState("");const[eOut,setEOut]=useState("");const[eBrk,setEBrk]=useState("");const[filterVa,setFilterVa]=useState("");const[ciVa,setCiVa]=useState("");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.Status!=="Deleted").sort((a,b)=>(b.ActivityDate||b.StartTime||"").localeCompare(a.ActivityDate||a.StartTime||""));
  const filtered=filterVa?shifts.filter(s=>s.VAEmail?.toLowerCase()===filterVa.toLowerCase()):shifts;
  // Active shifts (clocked in but not out)
  const activeShifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.StartTime&&!a.EndTime&&a.Status!=="Deleted");
  const clockedInEmails=new Set(activeShifts.map(s=>s.VAEmail?.toLowerCase()));
  const notClockedIn=data.vas.filter(v=>!clockedInEmails.has(v.Email?.toLowerCase())&&v.VATrackerStatus!=="Out");

  return<div style={{...ss.card,borderTop:`3px solid ${C.inf}`}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
      <div><div style={ss.cardT}>⏱ Shift Management</div><div style={ss.cardS}>Edit or delete shift records, manually clock in VAs</div></div>
      <div style={{minWidth:150}}><select style={ss.select} value={filterVa} onChange={e=>setFilterVa(e.target.value)}><option value="">All VAs</option>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
    </div>
    {/* Manual Clock-In */}
    {notClockedIn.length>0&&<div style={{background:C.wnb,border:`1px solid rgba(168,111,8,0.15)`,borderRadius:6,padding:12,marginBottom:12}}>
      <div style={{fontSize:12,fontWeight:700,color:C.wn,marginBottom:8}}>⚠ Not Clocked In Today ({notClockedIn.length})</div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
        {notClockedIn.map(v=><button key={v.Email} style={{...ss.btnO(C.wn,C.wnb),...ss.xs,display:"flex",alignItems:"center",gap:4}} onClick={()=>{if(window.confirm(`Manually clock in ${v.Name}?`))onClockInVA(v.Email,v.Name);}}><Avatar name={v.Name} size={16}/> Clock In {v.Name}</button>)}
      </div>
    </div>}
    {filtered.length===0&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:12}}>No shift records found.</div>}
    <div style={{borderRadius:8,border:`1px solid ${C.b1}`,overflow:"hidden"}}>
      {filtered.slice(0,20).map((s,i)=>{const isEd=editId===s.id;const hasIssue=(()=>{const otherShifts=shifts.filter(x=>x.id!==s.id&&x.VAEmail===s.VAEmail&&x.ActivityDate?.slice(0,10)===s.ActivityDate?.slice(0,10));return otherShifts.length>0;})();
        return<div key={s.id} style={{borderBottom:`1px solid ${C.b1}`,background:hasIssue?C.erb:"transparent"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px"}}>
            <Avatar name={s.VAName} size={24}/>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontSize:12,fontWeight:700,color:hasIssue?C.er:C.t2}}>{s.VAName}{hasIssue?" ⚠ Multiple shifts this day":""}</div>
              <div style={{fontSize:10,color:C.b4}}>{fD(s.ActivityDate||s.StartTime)} · In: {fT(s.StartTime)} · Out: {fT(s.EndTime)} · Break: {fM(s.BreakMinutes)} · Work: {fM(s.WorkMinutes)}</div>
            </div>
            <div style={{display:"flex",gap:4,flexShrink:0}}>
              <button style={{...ss.btnO(C.t2,C.b2),...ss.xs}} onClick={()=>{if(isEd){setEditId(null);}else{setEditId(s.id);setEIn(s.StartTime?.slice(0,16)||"");setEOut(s.EndTime?.slice(0,16)||"");setEBrk(String(s.BreakMinutes||0));}}}>✎</button>
              <button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>onDeleteShift(s.id)}>✕</button>
            </div>
          </div>
          {isEd&&<div style={{padding:"10px 12px",background:C.tl00,borderTop:`1px solid ${C.b1}`}}>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
              <div style={{flex:1,minWidth:160}}><label style={ss.label}>Clock In</label><input style={ss.input} type="datetime-local" value={eIn} onChange={e=>setEIn(e.target.value)}/></div>
              <div style={{flex:1,minWidth:160}}><label style={ss.label}>Clock Out</label><input style={ss.input} type="datetime-local" value={eOut} onChange={e=>setEOut(e.target.value)}/></div>
              <div style={{minWidth:80}}><label style={ss.label}>Break (min)</label><input style={ss.input} type="number" value={eBrk} onChange={e=>setEBrk(e.target.value)}/></div>
            </div>
            {eIn&&eOut&&<div style={{fontSize:11,color:C.b4,marginBottom:8}}>Calculated work time: {fM(Math.max(0,Math.round((new Date(eOut)-new Date(eIn))/6e4)-parseInt(eBrk||0)))}</div>}
            {eIn&&!eOut&&<div style={{fontSize:11,color:C.wn,marginBottom:8}}>Active shift — leave Clock Out empty to keep it open</div>}
            <div style={{display:"flex",gap:6}}>
              <button style={{...ss.btn(C.ok),...ss.xs}} onClick={()=>{if(!eIn){alert("Clock In time is required");return;}const fields={StartTime:new Date(eIn).toISOString(),BreakMinutes:parseInt(eBrk)||0,ActivityDate:new Date(eIn).toISOString()};if(eOut){fields.EndTime=new Date(eOut).toISOString();fields.WorkMinutes=Math.max(0,Math.round((new Date(eOut)-new Date(eIn))/6e4)-(parseInt(eBrk)||0));fields.Status="Completed";}onEditShift(s.id,fields);setEditId(null);}}>✓ Save</button>
              <button style={{...ss.btnO(C.b4,C.b2),...ss.xs}} onClick={()=>setEditId(null)}>Cancel</button>
            </div>
          </div>}
        </div>;})}
    </div>
    {filtered.length>20&&<div style={{fontSize:10,color:C.b4,textAlign:"center",marginTop:6}}>Showing 20 of {filtered.length} shifts</div>}
  </div>;
}

// ══════════════════════════════════════════════════════
// NOTICE MANAGER COMPONENT
// ══════════════════════════════════════════════════════
function NoticeManager({config,onUpdateConfig}){
  const[showAdd,setShowAdd]=useState(false);const[nTitle,setNTitle]=useState("");const[nText,setNText]=useState("");
  const[nStart,setNStart]=useState(today());const[nEnd,setNEnd]=useState(()=>{const d=new Date();d.setDate(d.getDate()+7);return d.toISOString().slice(0,10);});
  const notices=config?.notices||[];
  const td=today();

  function addNotice(){
    if(!nText.trim())return;
    const n={id:Date.now().toString(36),title:nTitle||"Notice",text:nText,startDate:nStart,endDate:nEnd,createdBy:"Admin",createdAt:new Date().toISOString()};
    onUpdateConfig({...config,notices:[...notices,n]});
    setNTitle("");setNText("");setShowAdd(false);
  }
  function removeNotice(id){if(!window.confirm("Remove this notice?"))return;onUpdateConfig({...config,notices:notices.filter(n=>n.id!==id)});}

  return(<div>
    {/* Active notices */}
    {notices.filter(n=>n.startDate<=td&&n.endDate>=td).map(n=><div key={n.id} style={{background:C.gl,border:`1px solid ${C.gold}`,borderLeft:`4px solid ${C.gold}`,borderRadius:6,padding:"10px 12px",marginBottom:8,display:"flex",alignItems:"flex-start",gap:10}}>
      <span style={{fontSize:14}}>📢</span>
      <div style={{flex:1}}>
        <div style={{fontSize:12,fontWeight:700,color:C.g2}}>{n.title}</div>
        <div style={{fontSize:11,color:C.t2,marginTop:2,lineHeight:1.5}}>{n.text}</div>
        <div style={{fontSize:9,color:C.b4,marginTop:3}}>{fD(n.startDate)} – {fD(n.endDate)} · {n.createdBy}</div>
      </div>
      <div style={{display:"flex",gap:4,flexShrink:0}}><Badge type="ok" dot={false}>Live</Badge><button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>removeNotice(n.id)}>✕</button></div>
    </div>)}

    {/* Scheduled / expired notices */}
    {notices.filter(n=>n.endDate<td||n.startDate>td).map(n=><div key={n.id} style={{padding:"8px 12px",borderBottom:`1px solid ${C.b1}`,display:"flex",alignItems:"center",gap:10,opacity:n.endDate<td?0.5:0.8}}>
      <div style={{flex:1}}>
        <div style={{fontSize:11,fontWeight:600,color:C.t2}}>{n.title}: {n.text.slice(0,60)}{n.text.length>60?"...":""}</div>
        <div style={{fontSize:9,color:C.b4}}>{fD(n.startDate)} – {fD(n.endDate)} · {n.endDate<td?"Expired":"Scheduled"}</div>
      </div>
      <Badge type={n.endDate<td?"ne":"in"} dot={false}>{n.endDate<td?"Expired":"Scheduled"}</Badge>
      <button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>removeNotice(n.id)}>✕</button>
    </div>)}

    {/* Add notice form */}
    {!showAdd?<button style={{...ss.btn(C.gold,C.teal),...ss.sm,marginTop:8}} onClick={()=>setShowAdd(true)}>+ Add Notice</button>
    :<div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:12,marginTop:8}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:8}}>
        <div style={{flex:2,minWidth:150}}><label style={ss.label}>Title</label><input style={ss.input} value={nTitle} onChange={e=>setNTitle(e.target.value)} placeholder="e.g. Reminder, Update, Important"/></div>
        <div style={{flex:1,minWidth:120}}><label style={ss.label}>Start Date</label><input style={ss.input} type="date" value={nStart} onChange={e=>setNStart(e.target.value)}/></div>
        <div style={{flex:1,minWidth:120}}><label style={ss.label}>End Date</label><input style={ss.input} type="date" value={nEnd} onChange={e=>setNEnd(e.target.value)}/></div>
      </div>
      <label style={ss.label}>Message *</label>
      <textarea style={{...ss.input,minHeight:60,marginBottom:8}} value={nText} onChange={e=>setNText(e.target.value)} placeholder="What do you need the team to know?"/>
      <div style={{display:"flex",gap:6}}>
        <button style={{...ss.btn(C.ok),...ss.sm}} onClick={addNotice}>✓ Post Notice</button>
        <button style={{...ss.btnO(C.b4,C.b2),...ss.sm}} onClick={()=>{setShowAdd(false);setNTitle("");setNText("");}}>Cancel</button>
      </div>
    </div>}
  </div>);
}

// ══════════════════════════════════════════════════════
// ADMIN VIEW (with sub-navigation)
// ══════════════════════════════════════════════════════
function AdminView({data,myEmail,myEmp,acct,config,queue,covQ,onToggleAbsence,onAssignTask,onUpdateConfig,onAssignProp,onUnassignProp,onReassignVA,onApproveTimeOff,onDenyTimeOff,onLogCallout,onLogEarlyDeparture,onEditShift,onDeleteShift,onClockInVA,onAddGuest,onDeactivateGuest,onAddProperty,onEditProperty,onEditEmployee,onDeleteTask}){
  const[sub,setSub]=useState("team");
  const[showAssign,setShowAssign]=useState(false);const[aVa,setAVa]=useState("");const[aCat,setACat]=useState("");const[aPri,setAPri]=useState("Normal");const[aDesc,setADesc]=useState("");const[aNotes,setANotes]=useState("");const[aProps,setAProps]=useState([]);
  const[showRec,setShowRec]=useState(false);const[rVa,setRVa]=useState("");const[rCat,setRCat]=useState("");const[rDesc,setRDesc]=useState("");const[rProps,setRProps]=useState([]);
  const[editIdx,setEditIdx]=useState(null);const[editDesc,setEditDesc]=useState("");
  const[portVa,setPortVa]=useState("");const[portProp,setPortProp]=useState("");
  const[showAddProp,setShowAddProp]=useState(false);const[npName,setNpName]=useState("");const[npGroup,setNpGroup]=useState("Multifamily");const[npUnits,setNpUnits]=useState("");const[npPm,setNpPm]=useState("");
  const[roleFilter,setRoleFilter]=useState("all");
  // Employee edit state
  const[editEmpId,setEditEmpId]=useState(null);const[eName,setEName]=useState("");const[eEmail,setEEmail]=useState("");const[eTitle,setETitle]=useState("");const[eRole,setERole]=useState("");
  // Guest state
  const[gName,setGName]=useState("");const[gEmail,setGEmail]=useState("");const[gCompany,setGCompany]=useState("");const[gVAs,setGVAs]=useState([]);
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
  const subTabs=[{k:"team",l:"👥 Team"},{k:"sched",l:"🔄 Recurring"},{k:"port",l:"🏠 Portfolio"},{k:"timeoff",l:"📅 Time Off"},{k:"guests",l:"🔑 Guests"},{k:"settings",l:"⚙ Settings"}];
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

    {/* ── TIME OFF ── */}
    {sub==="timeoff"&&<TimeOffAdmin data={data} myEmail={myEmail} acct={acct} myEmp={myEmp} onApprove={onApproveTimeOff} onDeny={onDenyTimeOff} onLogCallout={onLogCallout} onLogEarlyDeparture={onLogEarlyDeparture}/>}

    {/* ── GUESTS ── */}
    {sub==="guests"&&<div>
      <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div><div style={ss.cardT}>🔑 Guest Access Management</div><div style={ss.cardS}>Give coordinators a read-only view of their VAs' activity — no Microsoft account needed</div></div>
        </div>
        {/* Add Guest Form */}
        <div style={{background:C.gl,border:"1px solid rgba(205,160,75,0.3)",borderRadius:6,padding:12,marginBottom:12}}>
          <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:10}}>+ Add Guest</div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:10}}>
            <div style={{flex:2,minWidth:140}}><label style={ss.label}>Full Name *</label><input style={ss.input} value={gName} onChange={e=>setGName(e.target.value)} placeholder="e.g. Jessica Park"/></div>
            <div style={{flex:2,minWidth:140}}><label style={ss.label}>Email *</label><input style={ss.input} value={gEmail} onChange={e=>setGEmail(e.target.value)} placeholder="jessica@staffpro.com"/></div>
            <div style={{flex:1,minWidth:120}}><label style={ss.label}>Company</label><input style={ss.input} value={gCompany} onChange={e=>setGCompany(e.target.value)} placeholder="StaffPro"/></div>
          </div>
          <div style={{marginBottom:10}}><label style={ss.label}>VAs this guest can see *</label>
            <div style={{background:C.white,border:`1px solid ${C.b2}`,borderRadius:6,padding:8}}>
              {data.vas.map(v=><label key={v.Email} style={{display:"flex",gap:6,fontSize:12,marginBottom:4,cursor:"pointer"}}><input type="checkbox" checked={gVAs.includes(v.Email)} onChange={()=>setGVAs(p=>p.includes(v.Email)?p.filter(x=>x!==v.Email):[...p,v.Email])}/>{v.Name}</label>)}
            </div>
          </div>
          <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"8px 11px",fontSize:11,color:C.b4,marginBottom:10,lineHeight:1.6}}>
            <strong style={{color:C.t2}}>How it works:</strong> Guests get a unique URL with a token. They open it in any browser — no login required. They see a read-only Dashboard, History, and Scorecard filtered to only the VAs you assign.
          </div>
          <button style={{...ss.btn(C.ok),...ss.sm}} onClick={()=>{if(!gName||!gEmail||!gVAs.length){return;}
            onAddGuest({name:gName,email:gEmail,company:gCompany,vaEmails:gVAs});
            setGName("");setGEmail("");setGCompany("");setGVAs([]);
          }}>✓ Add Guest & Generate URL</button>
        </div>
        {/* Existing Guests */}
        {(()=>{const guests=(data.guests||[]).filter(g=>g.IsActive!==false);
          if(!guests.length)return<div style={{textAlign:"center",padding:"20px 0",color:C.b4,fontSize:12}}>No guests configured yet. Add one above.</div>;
          return guests.map((g,i)=>{
            const guestUrl=`https://va-tracker-guest.newshire-pm.workers.dev?token=${g.Title}`;
            const vaNames=(g.VAEmails||"").split(",").map(e=>{const v=data.vas.find(va=>va.Email?.toLowerCase()===e.trim().toLowerCase());return v?v.Name:e;}).join(", ");
            return<div key={i} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"12px 0",borderBottom:`1px solid ${C.b1}`}}>
            <Avatar name={g.GuestName} size={30} colorIdx={4}/>
            <div style={{flex:1}}>
              <div style={{fontSize:12,fontWeight:700}}>{g.GuestName}{g.Company?` — ${g.Company}`:""}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:1}}>VAs: {vaNames||"none assigned"}</div>
              <div style={{fontSize:10,color:C.b4,marginTop:1}}>Email: {g.GuestEmail}</div>
              <div style={{fontFamily:mono,fontSize:9,color:C.inf,background:C.infb,padding:"3px 7px",borderRadius:3,marginTop:4,display:"inline-block",wordBreak:"break-all"}}>{guestUrl}</div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:4,alignItems:"flex-end"}}>
              <Badge type="ok" dot={false}>Active</Badge>
              <button style={{...ss.btnO(C.t2,C.b2),...ss.xs}} onClick={()=>{navigator.clipboard?.writeText(guestUrl).then(()=>alert("URL copied to clipboard!")).catch(()=>{const ta=document.createElement("textarea");ta.value=guestUrl;document.body.appendChild(ta);ta.select();document.execCommand("copy");document.body.removeChild(ta);alert("URL copied!");});}}>📋 Copy URL</button>
              <button style={{...ss.btnO(C.er,"rgba(184,59,42,0.3)"),...ss.xs}} onClick={()=>onDeactivateGuest(g.id,g.GuestName)}>Deactivate</button>
            </div>
          </div>;});
        })()}
      </div>
    </div>}

    {/* ── SETTINGS ── */}
    {sub==="settings"&&<div>
      {/* Notice Banner Management */}
      <div style={{...ss.card,borderTop:`3px solid ${C.gold}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div><div style={ss.cardT}>📢 Notice Banner</div><div style={ss.cardS}>Post notices that appear at the top of every VA's My Day page</div></div>
        </div>
        <NoticeManager config={config} onUpdateConfig={onUpdateConfig}/>
      </div>

      {/* Shift Management */}
      <ShiftManager data={data} onEditShift={onEditShift} onDeleteShift={onDeleteShift} onClockInVA={onClockInVA}/>

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
