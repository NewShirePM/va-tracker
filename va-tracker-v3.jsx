const {useState,useEffect,useCallback,useRef}=React;

// ============================================================
// FAVICON
// ============================================================
(()=>{const l=document.querySelector("link[rel='icon']")||document.createElement("link");l.rel="icon";l.href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%23CDA04B'/><text x='16' y='23' font-family='Georgia,serif' font-size='22' font-weight='bold' fill='%2328434C' text-anchor='middle'>V</text></svg>";document.head.appendChild(l)})();

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  clientId: "32e75ffa-747a-4cf0-8209-6a19150c4547",
  tenantId: "33575d04-ca7b-4396-8011-9eaea4030b46",
  siteId: "vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
  appName: "VA PRODUCTIVITY TRACKER",
};
const GRAPH = "https://graph.microsoft.com/v1.0";
const SITE = `${GRAPH}/sites/${CONFIG.siteId}`;
const SCOPES = ["Sites.ReadWrite.All", "User.Read"];

// ============================================================
// ROLE DETECTION
// ============================================================
const ROLE_MAP = {
  "Virtual Assistant": "va",
  "Property Manager": "manager",
  "Regional/Portfolio Manager": "admin",
  "Owner/Operator": "admin",
};
function detectRole(emp) {
  if (emp.VATrackerRole) return emp.VATrackerRole.toLowerCase();
  return ROLE_MAP[emp.JobTitle] || null;
}

// ============================================================
// PALETTE
// ============================================================
const C = {
  headerBg:"#1C3740",headerHover:"#213F4A",
  teal700:"#28434C",teal600:"#2F5260",teal500:"#3A6577",teal400:"#4A7E91",teal100:"#D6E7EC",teal50:"#EDF4F7",
  gold700:"#9E7B2F",gold600:"#B8922E",gold500:"#CDA04B",gold400:"#D4AF61",gold100:"#F8F0DB",gold50:"#FFFBF0",
  white:"#FFFFFF",pageBg:"#F7F8F7",gray100:"#E8EAEA",gray200:"#D0D8DC",gray300:"#A8B0B0",gray400:"#7A8585",gray600:"#3E4A4A",dark:"#1A2A30",
  success:"#2D8A5A",successBg:"rgba(45,138,90,0.08)",error:"#C44B3B",errorBg:"rgba(196,75,59,0.06)",
  warning:"#D4960A",warningBg:"rgba(212,150,10,0.08)",info:"#4A78B0",infoBg:"rgba(74,120,176,0.08)",
};
const font="'Source Sans 3','Segoe UI',-apple-system,sans-serif";
const mono="'Source Code Pro',Consolas,monospace";

// ============================================================
// STYLES
// ============================================================
const S={
  page:{fontFamily:font,background:C.pageBg,minHeight:"100vh",color:C.teal700},
  header:{background:C.headerBg,borderBottom:`2px solid ${C.gold500}`,padding:"0 20px",display:"flex",alignItems:"center",justifyContent:"space-between",minHeight:56,flexWrap:"wrap",gap:8},
  headerTitle:{color:"#FFF",fontSize:16,fontWeight:700,letterSpacing:"0.05em"},
  headerSub:{color:C.gold500,fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase"},
  headerUser:{color:C.teal100,fontSize:13,textAlign:"right"},
  tabBar:{background:C.white,borderBottom:`1px solid ${C.gray200}`,display:"flex",gap:0,padding:"0 16px",overflowX:"auto",WebkitOverflowScrolling:"touch"},
  tab:a=>({padding:"12px 16px",fontSize:13,fontWeight:a?600:400,color:a?C.teal700:C.gray400,borderBottom:`2px solid ${a?C.gold500:"transparent"}`,cursor:"pointer",whiteSpace:"nowrap",background:"none",border:"none",fontFamily:font,minHeight:44}),
  content:{maxWidth:1200,margin:"0 auto",padding:"20px 16px"},
  card:{background:C.white,border:`1px solid ${C.gray200}`,borderRadius:6,boxShadow:"0 1px 3px rgba(28,55,64,0.06)",padding:16,marginBottom:14},
  cardTitle:{fontSize:16,fontWeight:600,color:C.teal700,paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:12},
  label:{display:"block",fontSize:13,fontWeight:500,color:C.teal700,marginBottom:3},
  input:{width:"100%",padding:"9px 11px",fontSize:14,fontFamily:font,color:C.teal700,background:C.white,border:`1px solid ${C.gray200}`,borderRadius:4,outline:"none",boxSizing:"border-box"},
  select:{width:"100%",padding:"9px 11px",fontSize:14,fontFamily:font,color:C.teal700,background:C.white,border:`1px solid ${C.gray200}`,borderRadius:4,cursor:"pointer",boxSizing:"border-box"},
  btn:(bg,fg)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,padding:"9px 16px",fontSize:13,fontWeight:600,fontFamily:font,color:fg||"#FFF",background:bg||C.headerBg,border:"none",borderRadius:4,cursor:"pointer",minHeight:40,whiteSpace:"nowrap"}),
  btnO:(fg,bdr)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:6,padding:"9px 16px",fontSize:13,fontWeight:600,fontFamily:font,color:fg||C.teal700,background:C.white,border:`1px solid ${bdr||C.teal100}`,borderRadius:4,cursor:"pointer",minHeight:40,whiteSpace:"nowrap"}),
  btnSm:{padding:"5px 12px",fontSize:11,minHeight:32},
  th:{textAlign:"left",padding:"9px 10px",fontSize:12,fontWeight:500,color:C.teal700,background:C.teal50,borderBottom:`2px solid ${C.teal100}`,whiteSpace:"nowrap"},
  td:{padding:"9px 10px",fontSize:13,color:C.gray600,borderBottom:`1px solid ${C.gray100}`},
  kpi:{background:C.white,border:`1px solid ${C.gray200}`,borderRadius:6,padding:"12px 8px",textAlign:"center",flex:"1 1 100px",minWidth:80},
  kpiL:{fontSize:11,fontWeight:500,color:C.gray400,textTransform:"uppercase",letterSpacing:"0.04em"},
  kpiV:{fontSize:24,fontWeight:700,fontFamily:mono},
  row:{display:"flex",gap:12,flexWrap:"wrap"},
};

// ============================================================
// COMPONENTS
// ============================================================
const badgeC={success:{c:C.success,b:C.successBg},error:{c:C.error,b:C.errorBg},warning:{c:C.warning,b:C.warningBg},info:{c:C.info,b:C.infoBg},neutral:{c:C.gray400,b:C.gray100}};
function Badge({type="neutral",children}){const m=badgeC[type]||badgeC.neutral;return<span style={{display:"inline-flex",padding:"2px 8px",fontSize:11,fontWeight:600,borderRadius:99,textTransform:"uppercase",letterSpacing:"0.03em",color:m.c,background:m.b}}>{children}</span>;}
function KPI({label,value,color,sub}){return<div style={S.kpi}><div style={S.kpiL}>{label}</div><div style={{...S.kpiV,color:color||C.teal700}}>{value}</div>{sub&&<div style={{fontSize:10,color:C.gray400,marginTop:2}}>{sub}</div>}</div>;}

const stBadge={Completed:"success",Blocked:"error","In Progress":"warning",Queued:"neutral",Incomplete:"error"};
const srcBadge={Daily:"info",Assigned:"warning","Ad-Hoc":"neutral",Coverage:"warning"};
const catIcon={"Work Orders":"\u{1F527}",Marketing:"\u{1F4E2}","Tenant Comms":"\u{1F4AC}",Reporting:"\u{1F4CA}",Inspections:"\u{1F50D}",Renewals:"\u{1F4DD}",Accounts:"\u{1F4B0}","Admin/Other":"\u{1F4C1}"};

// ============================================================
// HELPERS
// ============================================================
function fD(d){return d?new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}):"\u2014";}
function fT(d){return d?new Date(d).toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit"}):"\u2014";}
function fM(m){if(!m&&m!==0)return"\u2014";const h=Math.floor(m/60),r=m%60;return h>0?`${h}h ${r}m`:`${r}m`;}
function fTm(s){const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),ss=s%60;const p=n=>String(n).padStart(2,"0");return h>0?`${h}:${p(m)}:${p(ss)}`:`${p(m)}:${p(ss)}`;}
function today(){return new Date().toISOString().slice(0,10);}
function dAgo(d){return Math.floor((Date.now()-new Date(d).getTime())/864e5);}
function uid(){return Date.now()+"-"+Math.random().toString(36).slice(2,6);}
function sortCats(cats){if(!cats)return[];const other=cats.filter(c=>c.name==="Admin/Other");const rest=cats.filter(c=>c.name!=="Admin/Other").sort((a,b)=>a.name.localeCompare(b.name));return[...rest,...other];}
function inRange(d,from,to){if(!d)return false;const t=new Date(d).getTime();return t>=new Date(from).getTime()&&t<=new Date(to).getTime()+864e5;}

// ============================================================
// GRAPH API
// ============================================================
async function gGet(token,url){const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});if(!r.ok)throw new Error(`GET ${r.status}`);return r.json();}
async function gAll(token,url){let a=[],n=url;while(n){const d=await gGet(token,n);a=a.concat(d.value||[]);n=d["@odata.nextLink"]||null;}return a;}
async function gPost(token,url,fields){const r=await fetch(url,{method:"POST",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify({fields})});if(!r.ok)throw new Error(`POST ${r.status}`);return r.json();}
async function gPatch(token,url,fields){const r=await fetch(url,{method:"PATCH",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify(fields)});if(!r.ok)throw new Error(`PATCH ${r.status}`);return r.json();}
function lUrl(n){return`${SITE}/lists/${n}/items`;}
function iUrl(n,id){return`${SITE}/lists/${n}/items/${id}/fields`;}

async function safeGet(token,name,url){try{console.log(`[VT] Loading ${name}...`);const r=await gAll(token,url);console.log(`[VT] ${name}: ${r.length}`);return r;}catch(e){console.warn(`[VT] ${name} failed: ${e.message}`);return[];}}

async function loadAll(token){
  const[eR,cR,pR,ptR,aR]=await Promise.all([
    safeGet(token,"Employees",`${lUrl("Employees")}?expand=fields&$top=200`),
    safeGet(token,"Config",`${lUrl("VA_TrackerConfig")}?expand=fields&$top=10`),
    safeGet(token,"Properties",`${lUrl("VA_Properties")}?expand=fields&$top=200`),
    safeGet(token,"Portfolios",`${lUrl("VA_Portfolios")}?expand=fields&$top=200`),
    safeGet(token,"Activity",`${lUrl("VA_Activity")}?expand=fields&$top=200`),
  ]);
  const employees=eR.map(e=>({id:e.id,...e.fields}));
  const config=cR.length>0?JSON.parse(cR[0].fields.ConfigJSON||"{}"):{}; 
  const properties=pR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const portfolios=ptR.map(p=>({id:p.id,...p.fields})).filter(p=>p.IsActive!==false);
  const activities=aR.map(a=>({id:a.id,...a.fields}));
  const vas=employees.filter(e=>e.JobTitle==="Virtual Assistant"&&e.EmployeeActive!==false);
  const pms=employees.filter(e=>e.JobTitle==="Property Manager"&&e.EmployeeActive!==false);
  console.log("[VT] Loaded:",{emp:employees.length,props:properties.length,port:portfolios.length,act:activities.length,vas:vas.length,pms:pms.length});
  return{employees,config,properties,portfolios,activities,vas,pms};
}

// ============================================================
// MSAL
// ============================================================
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

// ============================================================
// MAIN APP
// ============================================================
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
  const[shift,setShift]=useState(null);
  const[covQ,setCovQ]=useState([]);
  const[queue,setQueue]=useState([]);
  const[tick,setTick]=useState(0);
  // Dashboard date filter
  const[dfFrom,setDfFrom]=useState(()=>{const d=new Date();d.setDate(d.getDate()-7);return d.toISOString().slice(0,10);});
  const[dfTo,setDfTo]=useState(today());

  useEffect(()=>{const id=setInterval(()=>setTick(t=>t+1),1000);return()=>clearInterval(id);},[]);
  const fl=useCallback(msg=>{setFlash(msg);setTimeout(()=>setFlash(""),2500);},[]);
  async function gT(){return(await refresh())||token;}

  // Load data
  useEffect(()=>{
    if(!token||!acct)return;
    const email=acct.username.toLowerCase();
    console.log("[VT] MSAL user:",email);
    setMyEmail(email);setLoading(true);
    loadAll(token).then(d=>{
      setData(d);
      const me=d.employees.find(e=>(e.Email&&e.Email.toLowerCase()===email)||(e.M365UserId&&e.M365UserId.toLowerCase()===email)||(e.Email&&email.split("@")[0]===e.Email.toLowerCase().split("@")[0]));
      console.log("[VT] Match:",me?`${me.Name} ${me.JobTitle} role=${me.VATrackerRole||"auto"}`:"NOT FOUND");
      if(!me){setRole(null);setError("access_denied");}
      else{const r=detectRole(me);if(!r){setRole(null);setError("access_denied");}else{setRole(r);setError(null);setMyEmp(me);}}
      buildQueue(d,email);
      setLoading(false);
    }).catch(e=>{console.error("[VT] Load error:",e);setError("load_error: "+e.message);setLoading(false);});
  },[token,acct]);

  function buildQueue(d,email){
    // 1. Load all persisted Queued tasks from VA_Activity
    const persistedQueued=d.activities.filter(a=>a.ActivityType==="Task"&&(a.Status==="Queued"||a.Status==="In Progress"));
    const q=[],cv=[];
    
    persistedQueued.forEach(a=>{
      const t={...a,_localId:a.id,_spId:a.id}; // _spId = SharePoint item ID for PATCH
      if(a.Source==="Coverage"||a.CoverageForEmail){cv.push(t);}
      else{q.push(t);}
    });

    // 2. Generate recurring daily tasks that don't already exist
    if(d.config.recurringTasks){
      const td=today();
      // Check what's already in SharePoint for today (any status)
      const todayTasks=d.activities.filter(a=>a.ActivityType==="Task"&&a.ActivityDate&&a.ActivityDate.slice(0,10)===td);
      const existing=new Set(todayTasks.map(t=>`${t.VAEmail}|${t.Title}`));
      // Also check what we just loaded as queued (could be from previous days)
      persistedQueued.forEach(a=>{existing.add(`${a.VAEmail}|${a.Title}`);});

      const toSave=[];
      d.config.recurringTasks.forEach(rt=>{
        if(!rt.active)return;
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

      // Save new daily tasks to SharePoint in background
      if(toSave.length>0){
        (async()=>{
          try{
            const tk=await gT();
            for(const task of toSave){
              const res=await gPost(tk,lUrl("VA_Activity"),{Title:task.Title,ActivityType:"Task",VAEmail:task.VAEmail,VAName:task.VAName,ActivityDate:task.ActivityDate,PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category,Source:task.Source,Status:task.Status,Priority:task.Priority||"Normal",CoverageForEmail:task.CoverageForEmail||"",CoverageForName:task.CoverageForName||""});
              const saved={...task,_localId:res.id,_spId:res.id,id:res.id};
              if(task.Source==="Coverage"){setCovQ(p=>[...p,saved]);}
              else{setQueue(p=>[...p,saved]);}
            }
            console.log(`[VT] Generated ${toSave.length} daily tasks`);
          }catch(e){console.error("[VT] Daily task generation error:",e);}
        })();
      }
    }

    setQueue(q);setCovQ(cv);
  }

  async function reload(){const t=await gT();const d=await loadAll(t);setData(d);buildQueue(d,myEmail);return d;}

  // ── Save task to SharePoint ──
  async function saveTask(task){
    const t=await gT();
    return gPost(t,lUrl("VA_Activity"),{Title:task.Title,ActivityType:"Task",VAEmail:task.VAEmail,VAName:task.VAName,ActivityDate:task.ActivityDate||new Date().toISOString(),PropertyId:task.PropertyId||"",PropertyName:task.PropertyName||"General",PMName:task.PMName||"",Category:task.Category,Source:task.Source,Status:task.Status,Priority:task.Priority||"Normal",StartTime:task.StartTime||null,EndTime:task.EndTime||null,DurationMin:task.DurationMin||0,PausedMin:task.PausedMin||0,Notes:task.Notes||"",CoverageForEmail:task.CoverageForEmail||"",CoverageForName:task.CoverageForName||"",AssignedByEmail:task.AssignedByEmail||"",AssignedByName:task.AssignedByName||""});
  }

  // ── Save shift ──
  async function saveShift(s){
    const t=await gT();
    return gPost(t,lUrl("VA_Activity"),{Title:`${s.VAName}-${today()}`,ActivityType:"Shift",VAEmail:s.VAEmail,VAName:s.VAName,ActivityDate:s.ClockIn,StartTime:s.ClockIn,EndTime:s.ClockOut,BreakMinutes:s.BreakMinutes,WorkMinutes:s.WorkMinutes,BreaksJSON:JSON.stringify(s.Breaks)});
  }

  // ── Absence toggle ──
  async function toggleAbsence(va){
    const t=await gT();const ns=va.VATrackerStatus==="Out"?"Active":"Out";
    await gPatch(t,iUrl("Employees",va.id),{VATrackerStatus:ns});
    if(ns==="Out"){
      await gPost(t,lUrl("VA_Activity"),{Title:`${va.Name}-Out-${today()}`,ActivityType:"Absence",VAEmail:va.Email,VAName:va.Name,ActivityDate:new Date().toISOString(),StartTime:new Date().toISOString(),Status:"Out",MarkedByEmail:myEmail,MarkedByName:acct.name||myEmail});
      // PATCH all their Queued tasks in SharePoint to Coverage
      const mv=queue.filter(q=>q.VAEmail.toLowerCase()===va.Email.toLowerCase());
      for(const task of mv){
        if(task._spId){try{await gPatch(t,iUrl("VA_Activity",task._spId),{Source:"Coverage",CoverageForEmail:va.Email,CoverageForName:va.Name});}catch(e){console.warn("[VT] Coverage patch failed:",e);}}
      }
      const rest=queue.filter(q=>q.VAEmail.toLowerCase()!==va.Email.toLowerCase());
      mv.forEach(q=>{q.Source="Coverage";q.CoverageForEmail=va.Email;q.CoverageForName=va.Name;});
      setQueue(rest);setCovQ(p=>[...p,...mv]);fl(`${va.Name} marked OUT — ${mv.length} tasks to coverage`);
    }else{
      // When marking back IN, move their unclaimed coverage tasks back
      const returning=covQ.filter(q=>q.CoverageForEmail?.toLowerCase()===va.Email.toLowerCase());
      for(const task of returning){
        if(task._spId){try{await gPatch(t,iUrl("VA_Activity",task._spId),{Source:"Daily",VAEmail:va.Email,VAName:va.Name,CoverageForEmail:"",CoverageForName:""});}catch(e){console.warn("[VT] Return patch failed:",e);}}
      }
      returning.forEach(q=>{q.Source="Daily";q.CoverageForEmail="";q.CoverageForName="";q.VAEmail=va.Email;q.VAName=va.Name;});
      setCovQ(p=>p.filter(q=>q.CoverageForEmail?.toLowerCase()!==va.Email.toLowerCase()));
      setQueue(p=>[...p,...returning]);
      fl(`${va.Name} marked IN — ${returning.length} tasks returned`);
    }
    await reload();
  }

  // ── Timer actions ──
  async function startTimer(task){
    const timer={...task,Status:"In Progress",StartTime:new Date().toISOString(),_pMs:0,_pS:null};
    // Update SharePoint if persisted
    if(task._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",task._spId),{Status:"In Progress",StartTime:new Date().toISOString()});}catch(e){console.warn("[VT] Timer start patch failed:",e);}}
    setTimers(p=>[...p,timer]);setQueue(p=>p.filter(t=>t._localId!==task._localId));fl("Timer started!");
  }
  function pauseTimer(id){setTimers(p=>p.map(t=>t._localId===id?{...t,_pS:Date.now()}:t));}
  function resumeTimer(id){setTimers(p=>p.map(t=>{if(t._localId!==id||!t._pS)return t;return{...t,_pMs:(t._pMs||0)+(Date.now()-t._pS),_pS:null};}));}
  async function finishTimer(id,status,notes){
    const t=timers.find(x=>x._localId===id);if(!t)return;
    const now=Date.now();let pMs=t._pMs||0;if(t._pS)pMs+=(now-t._pS);
    const dur=Math.max(1,Math.round((now-new Date(t.StartTime).getTime()-pMs)/6e4));
    const fields={Status:status,EndTime:status==="Completed"?new Date(now).toISOString():null,DurationMin:dur,PausedMin:Math.round(pMs/6e4),Notes:notes||""};
    try{
      const tk=await gT();
      if(t._spId){await gPatch(tk,iUrl("VA_Activity",t._spId),fields);}
      else{await saveTask({...t,...fields});}
      setTimers(p=>p.filter(x=>x._localId!==id));fl(status==="Completed"?"Task completed!":"Task: "+status);await reload();
    }catch(e){fl("Error: "+e.message);}
  }
  async function cancelTimer(id){
    const t=timers.find(x=>x._localId===id);if(!t)return;
    if(t._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",t._spId),{Status:"Queued",StartTime:null});}catch(e){console.warn("[VT] Cancel patch failed:",e);}}
    setTimers(p=>p.filter(x=>x._localId!==id));setQueue(p=>[{...t,Status:"Queued",StartTime:null,_pMs:0,_pS:null},...p]);
  }

  // ── Shift clock ──
  function clockIn(){
    // Validate: prevent double clock-in
    if(shift){fl("Already clocked in!");return;}
    setShift({ClockIn:new Date().toISOString(),Breaks:[],_ob:false,_bs:null});fl("Clocked in!");
  }
  function startBreak(){setShift(p=>p?{...p,_ob:true,_bs:new Date().toISOString()}:p);}
  function endBreak(){setShift(p=>{if(!p||!p._bs)return p;return{...p,_ob:false,Breaks:[...p.Breaks,{s:p._bs,e:new Date().toISOString()}],_bs:null};});}
  async function clockOut(){
    if(!shift)return;const now=new Date();const bks=[...shift.Breaks];
    if(shift._ob&&shift._bs)bks.push({s:shift._bs,e:now.toISOString()});
    const bMs=bks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);
    const bMin=Math.round(bMs/6e4);const wMin=Math.round((now-new Date(shift.ClockIn)-bMs)/6e4);
    // Warn if 10+ hours
    if(wMin>600&&!window.confirm(`You've been clocked in for ${fM(wMin)}. Clock out?`))return;
    try{await saveShift({VAEmail:myEmail,VAName:myEmp?.Name||myEmail,ClockIn:shift.ClockIn,ClockOut:now.toISOString(),BreakMinutes:bMin,WorkMinutes:wMin,Breaks:bks});
      setShift(null);fl("Clocked out!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Coverage ──
  async function claimCov(id){
    const t=covQ.find(x=>x._localId===id);if(!t)return;
    const meVa=myEmp?.Name||myEmail;
    if(t._spId){try{const tk=await gT();await gPatch(tk,iUrl("VA_Activity",t._spId),{VAEmail:myEmail,VAName:meVa});}catch(e){console.warn("[VT] Claim patch failed:",e);}}
    setCovQ(p=>p.filter(x=>x._localId!==id));setQueue(p=>[{...t,VAEmail:myEmail,VAName:meVa},...p]);fl(`Claimed! Covering for ${t.CoverageForName}`);
  }

  // ── Close Day ──
  async function closeDay(){
    if(!window.confirm("Close the day? Marks remaining daily tasks Incomplete."))return;
    const tk=await gT();
    for(const t of queue.filter(t=>t.Source==="Daily")){
      if(t._spId){try{await gPatch(tk,iUrl("VA_Activity",t._spId),{Status:"Incomplete",Notes:"Not started"});}catch(e){console.warn("[VT] Close patch failed:",e);}}
      else{await saveTask({...t,Status:"Incomplete",Notes:"Not started"});}
    }
    for(const t of covQ){
      if(t._spId){try{await gPatch(tk,iUrl("VA_Activity",t._spId),{Status:"Incomplete",Notes:"Uncovered"});}catch(e){console.warn("[VT] Close cov patch failed:",e);}}
      else{await saveTask({...t,Status:"Incomplete",Notes:"Uncovered"});}
    }
    setQueue(p=>p.filter(t=>t.Source!=="Daily"));setCovQ([]);
    if(shift)await clockOut();fl("Day closed");await reload();
  }

  // ── Ad-hoc / Assign ──
  async function addTask(task){
    const t={...task,Status:"Queued",ActivityDate:today(),ActivityType:"Task"};
    // Check if target VA is Out — route to coverage
    const targetVa=data?.vas.find(v=>v.Email?.toLowerCase()===task.VAEmail?.toLowerCase());
    if(targetVa&&targetVa.VATrackerStatus==="Out"){
      t.Source="Coverage";t.CoverageForEmail=targetVa.Email;t.CoverageForName=targetVa.Name;
    }
    try{
      const tk=await gT();
      const res=await gPost(tk,lUrl("VA_Activity"),{Title:t.Title,ActivityType:"Task",VAEmail:t.VAEmail,VAName:t.VAName,ActivityDate:t.ActivityDate,PropertyId:t.PropertyId||"",PropertyName:t.PropertyName||"General",PMName:t.PMName||"",Category:t.Category,Source:t.Source,Status:t.Status,Priority:t.Priority||"Normal",Notes:t.Notes||"",CoverageForEmail:t.CoverageForEmail||"",CoverageForName:t.CoverageForName||"",AssignedByEmail:t.AssignedByEmail||"",AssignedByName:t.AssignedByName||""});
      const saved={...t,_localId:res.id,_spId:res.id,id:res.id};
      if(t.Source==="Coverage"){setCovQ(p=>[saved,...p]);fl(`${targetVa.Name} is OUT — task sent to coverage pool`);}
      else{setQueue(p=>[saved,...p]);fl("Task added!");}
    }catch(e){fl("Error saving task: "+e.message);}
  }

  // ── Delete queued task (admin/manager) ──
  async function deleteTask(task){
    if(!task._spId){
      // Local only — just remove from state
      setQueue(p=>p.filter(t=>t._localId!==task._localId));
      setCovQ(p=>p.filter(t=>t._localId!==task._localId));
      fl("Task removed");return;
    }
    try{
      const tk=await gT();
      // Mark as "Cancelled" instead of deleting — preserves audit trail
      await gPatch(tk,iUrl("VA_Activity",task._spId),{Status:"Incomplete",Notes:"Removed by admin"});
      setQueue(p=>p.filter(t=>t._localId!==task._localId));
      setCovQ(p=>p.filter(t=>t._localId!==task._localId));
      fl("Task removed");
    }catch(e){fl("Error: "+e.message);}
  }

  // ── Weekly Report ──
  function generateWeeklyReport(){
    if(!data)return null;
    const now=new Date();const mon=new Date(now);mon.setDate(mon.getDate()-(mon.getDay()+6)%7-7);mon.setHours(0,0,0,0);
    const sun=new Date(mon);sun.setDate(sun.getDate()+6);sun.setHours(23,59,59,999);
    const from=mon.toISOString().slice(0,10),to=sun.toISOString().slice(0,10);
    const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&inRange(a.ActivityDate,from,to));
    const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&inRange(a.ActivityDate,from,to));
    const done=tasks.filter(t=>t.Status==="Completed");
    const blocked=tasks.filter(t=>t.Status==="Blocked");
    const inc=tasks.filter(t=>t.Status==="Incomplete");

    let report=`VA PRODUCTIVITY TRACKER — WEEKLY REPORT\nNewShire Property Management\nWeek: ${fD(from)} – ${fD(to)}\nGenerated: ${new Date().toLocaleString()}\nPrepared by: Brandy Turner\n${"=".repeat(60)}\n\n`;
    report+=`SUMMARY\n${"-".repeat(40)}\nTotal Tasks: ${tasks.length}\nCompleted: ${done.length} (${tasks.length?Math.round(done.length/tasks.length*100):0}%)\nBlocked: ${blocked.length}\nIncomplete: ${inc.length}\n\n`;

    data.vas.forEach(va=>{
      const vT=tasks.filter(t=>t.VAEmail===va.Email);const vD=vT.filter(t=>t.Status==="Completed");const vB=vT.filter(t=>t.Status==="Blocked");const vI=vT.filter(t=>t.Status==="Incomplete");
      const vS=shifts.filter(s=>s.VAEmail===va.Email);const vSm=vS.reduce((s,a)=>s+(a.WorkMinutes||0),0);const vTm=vD.reduce((s,t)=>s+(t.DurationMin||0),0);
      const vU=vSm>0?Math.round(vTm/vSm*100):0;const vR=vT.length?Math.round(vD.length/vT.length*100):0;
      report+=`\n${"=".repeat(60)}\n${va.Name.toUpperCase()}${va.VATrackerStatus==="Out"?" [OUT]":""}\n${"=".repeat(60)}\n`;
      report+=`Tasks: ${vD.length}/${vT.length} completed (${vR}%)\nShift Time: ${fM(vSm)} | Task Time: ${fM(vTm)} | Utilization: ${vU}%\n`;
      if(vB.length)report+=`\nBLOCKED (${vB.length}):\n${vB.map(t=>`  • ${t.Title} — ${t.PropertyName}${t.Notes?" — "+t.Notes:""}`).join("\n")}\n`;
      if(vI.length)report+=`\nINCOMPLETE (${vI.length}):\n${vI.map(t=>`  • ${t.Title} — ${t.PropertyName}${t.Notes?" — "+t.Notes:""}`).join("\n")}\n`;

      // Category breakdown
      const catM={};vD.forEach(t=>{if(!catM[t.Category])catM[t.Category]={c:0,m:0};catM[t.Category].c++;catM[t.Category].m+=(t.DurationMin||0);});
      const catL=Object.entries(catM).sort((a,b)=>b[1].m-a[1].m);
      if(catL.length){report+=`\nTIME BY CATEGORY:\n${catL.map(([cat,d])=>`  ${cat}: ${d.c} tasks, ${fM(d.m)}`).join("\n")}\n`;}

      // Property breakdown
      const propM={};vT.forEach(t=>{const pn=t.PropertyName||"General";if(!propM[pn])propM[pn]={done:0,total:0};propM[pn].total++;if(t.Status==="Completed")propM[pn].done++;});
      report+=`\nPROPERTY BREAKDOWN:\n${Object.entries(propM).map(([pn,d])=>`  ${pn}: ${d.done}/${d.total}`).join("\n")}\n`;
    });

    // Coverage summary
    const covTasks=tasks.filter(t=>t.CoverageForEmail);
    if(covTasks.length)report+=`\n${"=".repeat(60)}\nCOVERAGE TASKS: ${covTasks.length}\n${covTasks.map(t=>`  ${t.VAName} covered for ${t.CoverageForName}: ${t.Title}`).join("\n")}\n`;

    return{report,from,to};
  }

  // ── Config update ──
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

  // ── Property CRUD ──
  async function addProperty(p){
    try{const t=await gT();await gPost(t,lUrl("VA_Properties"),{Title:`PROP-${String(data.properties.length+1).padStart(3,"0")}`,PropertyName:p.name,PropertyGroup:p.group,Units:p.units,PMEmail:p.pmEmail,PMName:p.pmName,AppFolioId:p.appFolioId||"",IsActive:true});
      await reload();fl(`${p.name} added!`);
    }catch(e){fl("Error: "+e.message);}
  }
  async function editProperty(id,fields){
    try{const t=await gT();await gPatch(t,iUrl("VA_Properties",id),fields);await reload();fl("Property updated!");}catch(e){fl("Error: "+e.message);}
  }
  async function deactivateProperty(id,name){
    if(!window.confirm(`Deactivate ${name}? It will be hidden but not deleted.`))return;
    try{const t=await gT();await gPatch(t,iUrl("VA_Properties",id),{IsActive:false});await reload();fl(`${name} deactivated`);}catch(e){fl("Error: "+e.message);}
  }

  // ── Helpers for manager view ──
  function getVAForProperty(propId){
    const port=data?.portfolios.find(p=>p.PropertyId===propId);
    if(!port)return null;
    return data?.vas.find(v=>v.Email?.toLowerCase()===port.VAEmail?.toLowerCase())||null;
  }
  function myManagedProps(){
    if(!data||!myEmail)return[];
    return data.properties.filter(p=>p.PMEmail?.toLowerCase()===myEmail);
  }

  // ============================================================
  // COMPUTED
  // ============================================================
  const isAdmin=role==="admin";
  const isMgr=role==="admin"||role==="manager";
  const isVA=role==="va";
  const myVa=data?.vas.find(v=>v.Email&&v.Email.toLowerCase()===myEmail);
  const myPort=data?data.portfolios.filter(p=>p.VAEmail&&p.VAEmail.toLowerCase()===myEmail):[];
  const myProps=data?myPort.map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean):[];
  const mgrProps=myManagedProps();
  const myQ=queue.filter(t=>isVA?t.VAEmail.toLowerCase()===myEmail:true);
  const myTm=timers.filter(t=>isVA?t.VAEmail.toLowerCase()===myEmail:true);
  const outVAs=data?data.vas.filter(v=>v.VATrackerStatus==="Out"):[];

  // ============================================================
  // RENDER: Auth screens
  // ============================================================
  if(!acct)return(<div style={S.page}><div style={S.header}><div><div style={S.headerTitle}>{CONFIG.appName}</div><div style={S.headerSub}>NewShire Property Management</div></div></div><div style={{...S.content,textAlign:"center",paddingTop:80}}><div style={S.card}><div style={{fontSize:40,marginBottom:16}}>{"\u23F1"}</div><div style={{fontSize:20,fontWeight:700,color:C.teal700,marginBottom:8}}>VA Productivity Tracker</div><div style={{color:C.gray400,marginBottom:24}}>Sign in with your NewShire account.</div><button style={S.btn(C.headerBg)} onClick={login}>Sign In with Microsoft</button>{authErr&&<div style={{color:C.error,marginTop:12,fontSize:13}}>{authErr}</div>}</div></div></div>);
  if(loading)return(<div style={S.page}><div style={S.header}><div><div style={S.headerTitle}>{CONFIG.appName}</div><div style={S.headerSub}>NewShire Property Management</div></div></div><div style={{...S.content,textAlign:"center",paddingTop:80}}><div style={{fontSize:18,color:C.gray400}}>Loading...</div></div></div>);
  if(error||!role)return(<div style={S.page}><div style={S.header}><div><div style={S.headerTitle}>{CONFIG.appName}</div><div style={S.headerSub}>NewShire Property Management</div></div></div><div style={{...S.content,textAlign:"center",paddingTop:80}}><div style={{...S.card,borderLeft:`3px solid ${C.error}`}}><div style={{fontSize:40,marginBottom:16}}>{"\u{1F6AB}"}</div><div style={{fontSize:18,fontWeight:600,color:C.error,marginBottom:8}}>{error==="access_denied"?"Access Denied":"Error"}</div><div style={{color:C.gray400}}>{error==="access_denied"?"Your account is not authorized.":error}</div><div style={{color:C.gray300,fontSize:12,marginTop:12}}>Signed in as: {myEmail}</div></div></div></div>);

  // ============================================================
  // TABS
  // ============================================================
  const TABS=[];
  if(isMgr)TABS.push({n:"Dashboard",k:isAdmin?"dash":"mdash"});
  if(isVA||isAdmin)TABS.push({n:`My Tasks${myQ.length?` (${myQ.length})`:""}`,k:"tasks"});
  if(isVA||isAdmin)TABS.push({n:`Active${(isAdmin?timers.length:myTm.length)?` (${isAdmin?timers.length:myTm.length})`:""}`,k:"active"});
  if(role==="manager"||isAdmin)TABS.push({n:`My Properties${mgrProps.length?` (${mgrProps.length})`:isAdmin?" (All)":""}`,k:"mprops"});
  TABS.push({n:"History",k:"history"});
  TABS.push({n:"Scorecard",k:"score"});
  if(isAdmin)TABS.push({n:"Admin",k:"admin"});
  const ck=TABS[tab]?.k||TABS[0]?.k;

  // ============================================================
  // RENDER
  // ============================================================
  return(
    <div style={S.page}>
      <link href="https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600;700&family=Source+Code+Pro:wght@500;700&display=swap" rel="stylesheet"/>
      <div style={S.header}>
        <div><div style={S.headerTitle}>{CONFIG.appName}</div><div style={S.headerSub}>NewShire Property Management</div></div>
        <div style={S.headerUser}>
          <div style={{fontWeight:600}}>{acct.name||myEmail}</div>
          <div style={{fontSize:11,color:C.gold500,textTransform:"uppercase"}}>{role}{role==="manager"?` · ${mgrProps.length} properties`:""}</div>
          <div style={{display:"flex",gap:4,justifyContent:"flex-end",marginTop:2}}>
            {timers.length>0&&<Badge type="success">{timers.length} timing</Badge>}
            {covQ.length>0&&<Badge type="warning">{covQ.length} coverage</Badge>}
            {outVAs.length>0&&<Badge type="error">{outVAs.length} out</Badge>}
          </div>
        </div>
      </div>
      <div style={S.tabBar}>{TABS.map((t,i)=><button key={t.k} style={S.tab(tab===i)} onClick={()=>setTab(i)}>{t.n}</button>)}</div>
      {flash&&<div style={{background:C.gold50,borderBottom:`1px solid ${C.gold500}`,padding:"8px 20px",fontSize:13,fontWeight:600,color:C.gold700,textAlign:"center"}}>{flash}</div>}
      <div style={S.content}>
        {ck==="dash"&&<DashboardView data={data} queue={queue} timers={timers} covQ={covQ} dfFrom={dfFrom} dfTo={dfTo} setDfFrom={setDfFrom} setDfTo={setDfTo} filterProps={null}/>}
        {ck==="mdash"&&<DashboardView data={data} queue={queue} timers={timers} covQ={covQ} dfFrom={dfFrom} dfTo={dfTo} setDfFrom={setDfFrom} setDfTo={setDfTo} filterProps={mgrProps}/>}
        {ck==="tasks"&&<MyTasksView data={data} role={role} myEmail={myEmail} myVa={myVa} myProps={myProps} queue={queue} covQ={covQ} shift={shift} timers={timers} tick={tick} onClockIn={clockIn} onBreakStart={startBreak} onBreakEnd={endBreak} onClockOut={clockOut} onStartTimer={startTimer} onClaimCov={claimCov} onAddTask={addTask} onDeleteTask={deleteTask} config={data?.config}/>}
        {ck==="active"&&<ActiveView timers={isVA?timers.filter(t=>t.VAEmail.toLowerCase()===myEmail):timers} tick={tick} isMgr={isMgr} onPause={pauseTimer} onResume={resumeTimer} onFinish={finishTimer} onCancel={cancelTimer}/>}
        {ck==="mprops"&&<ManagerPropsView data={data} myEmail={myEmail} myEmp={myEmp} mgrProps={isAdmin?data.properties:mgrProps} queue={queue} timers={timers} covQ={covQ} onAddTask={addTask} getVA={getVAForProperty} isAdmin={isAdmin}/>}
        {ck==="history"&&<HistoryView data={data} role={role} myEmail={myEmail} isMgr={isMgr} mgrProps={mgrProps}/>}
        {ck==="score"&&<ScorecardView data={data} role={role} myEmail={myEmail} isMgr={isMgr} myVa={myVa} myProps={myProps} mgrProps={mgrProps}/>}
        {ck==="admin"&&<AdminView data={data} myEmail={myEmail} acct={acct} config={data?.config} queue={queue} covQ={covQ} onToggleAbsence={toggleAbsence} onAssignTask={addTask} onCloseDay={closeDay} onUpdateConfig={updateConfig} onAssignProp={assignProp} onUnassignProp={unassignProp} onAddProperty={addProperty} onEditProperty={editProperty} onDeactivateProperty={deactivateProperty} onDeleteTask={deleteTask} generateReport={generateWeeklyReport}/>}
      </div>
    </div>
  );
}

// ============================================================
// DASHBOARD — shared by Admin (all) and Manager (filtered)
// ============================================================
function DashboardView({data,queue,timers,covQ,dfFrom,dfTo,setDfFrom,setDfTo,filterProps}){
  if(!data)return null;
  // Filter activities by date range and optionally by properties
  const propIds=filterProps?new Set(filterProps.map(p=>p.Title)):null;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&inRange(a.ActivityDate,dfFrom,dfTo)&&(!propIds||propIds.has(a.PropertyId)));
  const done=tasks.filter(a=>a.Status==="Completed");
  const blocked=tasks.filter(a=>a.Status==="Blocked");
  const inc=tasks.filter(a=>a.Status==="Incomplete");
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&inRange(a.ActivityDate,dfFrom,dfTo)&&(!propIds||true));
  const shiftMin=shifts.reduce((s,a)=>s+(a.WorkMinutes||0),0);
  const taskMin=done.reduce((s,a)=>s+(a.DurationMin||0),0);
  const rate=tasks.length?Math.round(done.length/tasks.length*100):0;
  const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;
  const filteredVAs=propIds?data.vas.filter(v=>data.portfolios.some(p=>propIds.has(p.PropertyId)&&p.VAEmail?.toLowerCase()===v.Email.toLowerCase())):data.vas;

  return(
    <div>
      {/* Date filter */}
      <div style={{...S.card,display:"flex",gap:12,alignItems:"flex-end",flexWrap:"wrap",padding:12}}>
        <div><label style={S.label}>From</label><input type="date" style={{...S.input,width:150}} value={dfFrom} onChange={e=>setDfFrom(e.target.value)}/></div>
        <div><label style={S.label}>To</label><input type="date" style={{...S.input,width:150}} value={dfTo} onChange={e=>setDfTo(e.target.value)}/></div>
        <div style={{display:"flex",gap:6}}>
          {[{l:"7d",d:7},{l:"14d",d:14},{l:"30d",d:30},{l:"90d",d:90}].map(p=><button key={p.l} style={{...S.btnO(C.teal600,C.teal100),...S.btnSm}} onClick={()=>{const f=new Date();f.setDate(f.getDate()-p.d);setDfFrom(f.toISOString().slice(0,10));setDfTo(today());}}>{p.l}</button>)}
        </div>
        {filterProps&&<div style={{fontSize:12,color:C.gold700,fontWeight:600,marginLeft:"auto"}}>Showing {filterProps.length} managed properties</div>}
      </div>

      {/* KPIs */}
      <div style={{...S.row,marginBottom:16}}>
        <KPI label="Tasks" value={tasks.length}/>
        <KPI label="Done" value={done.length} color={C.success}/>
        <KPI label="Rate" value={`${rate}%`} color={rate>=85?C.success:rate>=60?C.warning:tasks.length?C.error:C.teal700}/>
        {!filterProps&&<KPI label="Shift Hrs" value={fM(shiftMin)}/>}
        {!filterProps&&<KPI label="Utilization" value={`${util}%`} color={util>=75?C.success:util>=50?C.warning:shiftMin>0?C.error:C.teal700}/>}
        <KPI label="Blocked" value={blocked.length} color={blocked.length>0?C.error:C.success}/>
        <KPI label="Coverage" value={covQ.length} color={covQ.length>0?C.warning:C.success}/>
        {!filterProps&&<KPI label="VAs Out" value={data.vas.filter(v=>v.VATrackerStatus==="Out").length} color={data.vas.filter(v=>v.VATrackerStatus==="Out").length>0?C.warning:C.teal700}/>}
      </div>

      {/* Needs Attention */}
      {(blocked.length>0||inc.length>0||covQ.length>0)&&(
        <div style={{...S.card,borderLeft:`3px solid ${C.error}`,background:C.errorBg}}>
          <div style={{fontSize:16,fontWeight:600,color:C.error,marginBottom:10}}>{"\u{1F6A8}"} Needs Attention</div>
          {blocked.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:12,fontWeight:600,color:C.error,marginBottom:4}}>Blocked Tasks ({blocked.length})</div>{blocked.slice(0,5).map((t,i)=><div key={i} style={{fontSize:13,color:C.gray600,padding:"3px 0"}}>{t.VAName}: {t.Title} — {t.PropertyName}{t.Notes?` · ${t.Notes}`:""}</div>)}{blocked.length>5&&<div style={{fontSize:11,color:C.gray400}}>+{blocked.length-5} more</div>}</div>}
          {covQ.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:12,fontWeight:600,color:C.warning,marginBottom:4}}>Unclaimed Coverage ({covQ.length})</div>{covQ.slice(0,5).map((t,i)=><div key={i} style={{fontSize:13,color:C.gray600,padding:"3px 0"}}>{t.Title} — covering {t.CoverageForName}</div>)}</div>}
          {inc.length>3&&<div><div style={{fontSize:12,fontWeight:600,color:C.gray400,marginBottom:4}}>High Incomplete Rate</div><div style={{fontSize:13,color:C.gray600}}>{inc.length} tasks marked incomplete in this period</div></div>}
        </div>
      )}

      {/* Kanban */}
      <div style={S.card}>
        <div style={S.cardTitle}>{"\u{1F4CB}"} Task Board</div>
        <div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:8}}>
          {filteredVAs.map(va=>{
            const isOut=va.VATrackerStatus==="Out";
            const vQ=queue.filter(t=>t.VAEmail.toLowerCase()===va.Email.toLowerCase());
            const vT=timers.filter(t=>t.VAEmail.toLowerCase()===va.Email.toLowerCase());
            const port=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase());
            return(
              <div key={va.Email} style={{flex:"1 0 200px",minWidth:200,background:isOut?C.errorBg:C.pageBg,borderRadius:6,padding:10,border:`1px solid ${isOut?"rgba(196,75,59,0.2)":C.gray100}`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8,paddingBottom:6,borderBottom:`2px solid ${C.teal100}`}}>
                  <div><div style={{fontSize:13,fontWeight:700,color:isOut?C.error:C.teal700}}>{va.Name}{isOut?" \u{1F6D1}":""}</div><div style={{fontSize:10,color:C.gray400}}>{port.length} props</div></div>
                  {isOut?<Badge type="error">OUT</Badge>:vT.length>0?<Badge type="success">Active</Badge>:<span style={{fontSize:10,color:C.gray300}}>Offline</span>}
                </div>
                {vT.map(t=><div key={t._localId} style={{background:C.white,border:`1px solid ${C.success}`,borderLeft:`3px solid ${C.success}`,borderRadius:4,padding:8,marginBottom:5}}><div style={{fontSize:10,fontWeight:600,color:C.success,textTransform:"uppercase"}}>{"\u25CF"} In Progress</div><div style={{fontSize:12,fontWeight:500,color:C.teal700}}>{t.Title}</div><div style={{fontSize:10,color:C.gray400}}>{t.PropertyName}</div></div>)}
                {vQ.map(t=><div key={t._localId} style={{background:C.white,border:`1px solid ${C.gray200}`,borderLeft:`3px solid ${t.Priority==="Urgent"?C.error:t.Priority==="High"?C.warning:C.gray200}`,borderRadius:4,padding:8,marginBottom:5}}><div style={{fontSize:12,fontWeight:500,color:C.teal700}}>{t.Title}</div><div style={{fontSize:10,color:C.gray400}}>{t.PropertyName}</div></div>)}
                {!vT.length&&!vQ.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.gray300,fontSize:11,fontStyle:"italic"}}>{isOut?"Tasks in coverage":"No pending"}</div>}
                <div style={{marginTop:6,paddingTop:6,borderTop:`1px solid ${C.gray100}`,fontSize:10,color:C.gray400,display:"flex",justifyContent:"space-between"}}><span>{vQ.length+vT.length} pending</span><span>{vT.length} active</span></div>
              </div>
            );
          })}
        </div>
      </div>

      {/* VA Performance Table */}
      <div style={S.card}>
        <div style={S.cardTitle}>VA Performance</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["VA","Portfolio","Done","Rate","Shift","Tasks","Util","Blocked","Incomplete"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {filteredVAs.map(va=>{
                const vT=tasks.filter(a=>a.VAEmail===va.Email);const vD=vT.filter(a=>a.Status==="Completed");const vI=vT.filter(a=>a.Status==="Incomplete");const vB=vT.filter(a=>a.Status==="Blocked");
                const vS=shifts.filter(a=>a.VAEmail===va.Email);const vSm=vS.reduce((s,a)=>s+(a.WorkMinutes||0),0);const vTm=vD.reduce((s,a)=>s+(a.DurationMin||0),0);
                const vR=vT.length?Math.round(vD.length/vT.length*100):0;const vU=vSm>0?Math.round(vTm/vSm*100):0;
                const port=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase());
                return(<tr key={va.Email} style={{opacity:va.VATrackerStatus==="Out"?0.6:1}}>
                  <td style={{...S.td,fontWeight:600,color:va.VATrackerStatus==="Out"?C.error:C.teal700}}>{va.Name}</td>
                  <td style={{...S.td,fontSize:12}}>{port.length} props</td>
                  <td style={S.td}>{vD.length}/{vT.length}</td>
                  <td style={S.td}><Badge type={vR>=85?"success":vR>=60?"warning":"error"}>{vR}%</Badge></td>
                  <td style={{...S.td,fontFamily:mono,fontSize:12}}>{fM(vSm)}</td>
                  <td style={{...S.td,fontFamily:mono,fontSize:12}}>{fM(vTm)}</td>
                  <td style={S.td}><div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:40,height:6,background:C.gray100,borderRadius:4,overflow:"hidden"}}><div style={{width:`${Math.min(vU,100)}%`,height:"100%",background:vU>=75?C.success:vU>=50?C.warning:C.error,borderRadius:4}}/></div><span style={{fontFamily:mono,fontSize:10,fontWeight:600,color:vU>=75?C.success:vU>=50?C.warning:C.error}}>{vU}%</span></div></td>
                  <td style={S.td}>{vB.length>0?<Badge type="error">{vB.length}</Badge>:"0"}</td>
                  <td style={S.td}>{vI.length>0?<Badge type="error">{vI.length}</Badge>:"0"}</td>
                </tr>);
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// TASK ROW
// ============================================================
function TaskRow({task,onStart,onDelete,showVA}){
  return(<div style={{display:"flex",alignItems:"center",gap:8,padding:"9px 0",borderBottom:`1px solid ${C.gray100}`}}>
    <span style={{fontSize:16}}>{catIcon[task.Category]||"\u{1F4C1}"}</span>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontSize:13,fontWeight:500,color:C.teal700}}>{task.Title}{task.CoverageForName&&<>{" "}<Badge type="warning">Coverage</Badge></>}</div>
      <div style={{fontSize:11,color:C.gray400}}>{showVA?`${task.VAName} · `:""}{task.PropertyName}{task.Priority!=="Normal"&&<>{" · "}<Badge type={task.Priority==="Urgent"?"error":"warning"}>{task.Priority}</Badge></>}{task.Notes&&` · \u{1F4AC} ${task.Notes}`}</div>
    </div>
    <div style={{display:"flex",gap:4,flexShrink:0}}>
      {onStart&&<button style={{...S.btn(C.success),...S.btnSm}} onClick={onStart}>{"\u25B6"} Start</button>}
      {onDelete&&<button style={{...S.btnO(C.error,C.error),...S.btnSm,padding:"5px 8px"}} onClick={()=>{if(window.confirm("Remove this task?"))onDelete(task);}} title="Remove task">{"\u2715"}</button>}
    </div>
  </div>);
}

// ============================================================
// MY TASKS (VA view)
// ============================================================
function MyTasksView({data,role,myEmail,myVa,myProps,queue,covQ,shift,timers,tick,onClockIn,onBreakStart,onBreakEnd,onClockOut,onStartTimer,onClaimCov,onAddTask,onDeleteTask,config}){
  const[showForm,setShowForm]=useState(false);const[fCat,setFCat]=useState("");const[fProp,setFProp]=useState("");const[fPri,setFPri]=useState("Normal");const[fDesc,setFDesc]=useState("");
  const isAdm=role==="admin";
  const myTasks=isAdm?queue:queue.filter(t=>t.VAEmail.toLowerCase()===myEmail);
  const daily=myTasks.filter(t=>t.Source==="Daily");
  const assigned=myTasks.filter(t=>t.Source==="Assigned"||t.Source==="Coverage");
  const adhoc=myTasks.filter(t=>t.Source==="Ad-Hoc");
  const myActive=isAdm?timers.length:timers.filter(t=>t.VAEmail.toLowerCase()===myEmail).length;
  const cats=config?.categories||[];
  const portProps=data?data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===myEmail).map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean):[];
  let shE=0,bkE=0;
  if(shift){const now=Date.now();let bMs=shift.Breaks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);if(shift._ob&&shift._bs)bMs+=(now-new Date(shift._bs).getTime());bkE=Math.floor(bMs/1000);shE=Math.floor((now-new Date(shift.ClockIn).getTime()-bMs)/1000);}

  function handleAdd(){if(!fDesc||!fCat)return;const cat=cats.find(c=>c.id===fCat);const prop=fProp?data.properties.find(p=>p.Title===fProp):null;
    onAddTask({Title:fDesc,VAEmail:myEmail,VAName:myVa?.Name||myEmail,PropertyId:fProp||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Priority:fPri,Source:"Ad-Hoc"});
    setFDesc("");setFCat("");setFProp("");setFPri("Normal");setShowForm(false);}

  return(<div style={{maxWidth:700,margin:"0 auto"}}>
    {/* Identity */}
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
      <div><div style={{fontSize:18,fontWeight:700,color:C.teal700}}>{myVa?.Name||myEmail}</div><div style={{fontSize:12,color:C.gray400}}>{portProps.length} properties · {portProps.reduce((s,p)=>s+(p.Units||0),0)} units</div></div>
      {myActive>0&&<Badge type="success">{myActive} timing</Badge>}
    </div>
    {/* Shift Clock */}
    <div style={{...S.card,borderLeft:`3px solid ${shift?(shift._ob?C.warning:C.success):C.info}`}}>
      {!shift?(<div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:15,fontWeight:600,color:C.teal700}}>{"\u23F0"} Shift Clock</div><div style={{fontSize:12,color:C.gray400}}>Not clocked in</div></div><button style={S.btn(C.info)} onClick={onClockIn}>{"\u2600"} Clock In</button></div>
      ):(<div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div><div style={{fontSize:13,fontWeight:600,color:shift._ob?C.warning:C.success,textTransform:"uppercase"}}>{shift._ob?"\u2615 On Break":"\u23F0 Clocked In"}</div><div style={{fontSize:11,color:C.gray400}}>Since {fT(shift.ClockIn)} · {shift.Breaks.length+(shift._ob?1:0)} breaks</div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:10,color:C.gray400,textTransform:"uppercase"}}>Working</div><div style={{fontSize:22,fontWeight:700,fontFamily:mono,color:shift._ob?C.warning:C.teal700}}>{fTm(shE)}</div></div></div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {shift._ob?<button style={{...S.btn(C.success),flex:1}} onClick={onBreakEnd}>{"\u25B6"} End Break</button>:<button style={{...S.btnO(C.warning,C.warning),flex:1}} onClick={onBreakStart}>{"\u2615"} Break</button>}
          <button style={{...S.btn(C.error),flex:1}} onClick={onClockOut}>{"\u{1F319}"} Clock Out</button>
        </div></div>)}
    </div>
    {/* Coverage */}
    {covQ.length>0&&(<div style={{...S.card,borderLeft:`3px solid ${C.warning}`,background:C.warningBg}}>
      <div style={{fontSize:15,fontWeight:600,color:C.warning,marginBottom:8}}>{"\u{1F6A8}"} Coverage Needed ({covQ.length})</div>
      {covQ.map(t=><div key={t._localId} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:"1px solid rgba(212,150,10,0.1)"}}>
        <span>{catIcon[t.Category]||"\u{1F4C1}"}</span><div style={{flex:1}}><div style={{fontSize:13,fontWeight:500,color:C.teal700}}>{t.Title}</div><div style={{fontSize:11,color:C.gray400}}>{t.PropertyName} · covering {t.CoverageForName}</div></div>
        <button style={{...S.btn(C.warning,C.dark),...S.btnSm}} onClick={()=>onClaimCov(t._localId)}>{"\u270B"} Claim</button></div>)}
    </div>)}
    {/* Daily */}
    <div style={S.card}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",...(daily.length?{paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:6}:{})}}><div style={{fontSize:15,fontWeight:600,color:C.teal700}}>{"\u{1F4CB}"} Daily Tasks</div><span style={{fontSize:12,color:C.gray400}}>{daily.length} left</span></div>
      {!daily.length&&<p style={{color:C.gray400,fontSize:13,marginTop:10}}>All done!</p>}
      {daily.map(t=><TaskRow key={t._localId} task={t} onStart={()=>onStartTimer(t)} onDelete={isAdm?onDeleteTask:null} showVA={isAdm}/>)}</div>
    {/* Assigned */}
    {assigned.length>0&&<div style={{...S.card,borderLeft:`3px solid ${C.gold500}`}}><div style={{fontSize:15,fontWeight:600,color:C.teal700,marginBottom:6}}>{"\u{1F4CC}"} Assigned & Coverage ({assigned.length})</div>{assigned.map(t=><TaskRow key={t._localId} task={t} onStart={()=>onStartTimer(t)} onDelete={isAdm?onDeleteTask:null} showVA={isAdm}/>)}</div>}
    {/* Ad-Hoc */}
    {adhoc.length>0&&<div style={S.card}><div style={{fontSize:15,fontWeight:600,color:C.teal700,marginBottom:6}}>{"\u{1F4DD}"} Extra ({adhoc.length})</div>{adhoc.map(t=><TaskRow key={t._localId} task={t} onStart={()=>onStartTimer(t)} onDelete={isAdm?onDeleteTask:null} showVA={isAdm}/>)}</div>}
    {/* Add Task */}
    <div style={S.card}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{fontSize:14,fontWeight:600,color:C.teal700}}>➕ Add Task</div><button style={S.btnO(C.teal700,C.teal100)} onClick={()=>setShowForm(!showForm)}>{showForm?"Cancel":"New"}</button></div>
      {showForm&&<div style={{marginTop:14}}>
        <div style={S.row}><div style={{flex:1,minWidth:130}}><label style={S.label}>Category *</label><select style={S.select} value={fCat} onChange={e=>setFCat(e.target.value)}><option value="">...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div style={{flex:1,minWidth:130}}><label style={S.label}>Property</label><select style={S.select} value={fProp} onChange={e=>setFProp(e.target.value)}><option value="">General</option>{portProps.map(p=><option key={p.Title} value={p.Title}>{p.PropertyName}</option>)}</select></div>
          <div style={{minWidth:80}}><label style={S.label}>Priority</label><select style={S.select} value={fPri} onChange={e=>setFPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div></div>
        <div style={{marginTop:10,marginBottom:10}}><label style={S.label}>Task *</label><input style={S.input} value={fDesc} onChange={e=>setFDesc(e.target.value)} placeholder="Describe..."/></div>
        <button style={{...S.btn(C.headerBg),width:"100%"}} onClick={handleAdd}>+ Queue</button>
      </div>}</div>
  </div>);
}

// ============================================================
// ACTIVE VIEW
// ============================================================
function ActiveView({timers:tms,tick,isMgr,onPause,onResume,onFinish,onCancel}){
  if(!tms.length)return(<div style={{...S.card,textAlign:"center",padding:50}}><div style={{fontSize:36,marginBottom:10}}>{"\u23F1"}</div><div style={{fontSize:17,fontWeight:600,color:C.teal700}}>No Active Timers</div><div style={{fontSize:13,color:C.gray400}}>Start a task to begin tracking.</div></div>);
  return(<div>{tms.map(t=>{const now=Date.now(),st=new Date(t.StartTime).getTime();let pMs=t._pMs||0;const ip=!!t._pS;if(ip)pMs+=(now-t._pS);const el=ip?Math.floor((t._pS-st-(t._pMs||0))/1000):Math.floor((now-st-pMs)/1000);
    return(<div key={t._localId} style={{...S.card,borderLeft:`4px solid ${ip?C.warning:C.success}`,padding:0,overflow:"hidden"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",background:ip?C.warningBg:C.successBg}}>
        <div><div style={{fontSize:12,fontWeight:600,color:ip?C.warning:C.success,textTransform:"uppercase"}}>{ip?"\u23F8 Paused":"\u25CF Recording"}</div><div style={{fontSize:11,color:C.gray400,marginTop:2}}>{fT(t.StartTime)} · <Badge type={srcBadge[t.Source]||"neutral"}>{t.Source}</Badge>{isMgr&&` · ${t.VAName}`}</div></div>
        <div style={{fontSize:26,fontWeight:700,fontFamily:mono,color:ip?C.warning:C.teal700}}>{fTm(el)}</div></div>
      <div style={{padding:"12px 16px"}}><div style={{fontSize:14,fontWeight:600,color:C.teal700,marginBottom:3}}>{t.Title}</div><div style={{fontSize:12,color:C.gray400}}>{"\u{1F3E0}"} {t.PropertyName} · {catIcon[t.Category]||""} {t.Category}</div>{t.CoverageForName&&<div style={{fontSize:11,color:C.warning,marginTop:2}}>{"\u{1F504}"} Covering for {t.CoverageForName}</div>}</div>
      <div style={{display:"flex",gap:6,padding:"0 16px 12px",flexWrap:"wrap"}}>
        {ip?<button style={{...S.btn(C.success),flex:1}} onClick={()=>onResume(t._localId)}>{"\u25B6"} Resume</button>:<button style={{...S.btnO(C.warning,C.warning),flex:1}} onClick={()=>onPause(t._localId)}>{"\u23F8"} Pause</button>}
        <button style={{...S.btn(C.success),flex:1}} onClick={()=>{const n=prompt("Notes (optional):");onFinish(t._localId,"Completed",n||"");}}>{"\u2713"} Done</button>
        <button style={{...S.btnO(C.error,C.error),padding:"9px 12px"}} onClick={()=>{const n=prompt("What's blocking?");if(n)onFinish(t._localId,"Blocked",n);}}>{"\u26A0"}</button>
        <button style={{...S.btnO(C.gray400,C.gray200),padding:"9px 12px"}} onClick={()=>onCancel(t._localId)}>{"\u21A9"}</button>
      </div></div>);})}</div>);
}

// ============================================================
// MANAGER PROPERTIES VIEW
// ============================================================
function ManagerPropsView({data,myEmail,myEmp,mgrProps,queue,timers,covQ,onAddTask,getVA}){
  const[selProp,setSelProp]=useState("");
  const[tCat,setTCat]=useState("");const[tDesc,setTDesc]=useState("");const[tPri,setTPri]=useState("Normal");const[tNotes,setTNotes]=useState("");
  const cats=data?.config?.categories||[];

  function handleSubmit(){
    if(!selProp||!tCat||!tDesc)return;
    const prop=data.properties.find(p=>p.Title===selProp);
    const va=getVA(selProp);
    if(!va){alert("No VA assigned to this property. Ask admin to assign one.");return;}
    const cat=cats.find(c=>c.id===tCat);
    onAddTask({Title:tDesc,VAEmail:va.Email,VAName:va.Name,PropertyId:selProp,PropertyName:prop?.PropertyName||"",PMName:myEmp?.Name||myEmail,Category:cat?.name||"Admin/Other",Priority:tPri,Source:"Assigned",AssignedByEmail:myEmail,AssignedByName:myEmp?.Name||myEmail,Notes:tNotes});
    setTDesc("");setTCat("");setTPri("Normal");setTNotes("");
  }

  return(<div>
    {/* Property overview cards */}
    <div style={S.row}>
      {mgrProps.map(prop=>{
        const va=getVA(prop.Title);
        const pendingQ=queue.filter(t=>t.PropertyId===prop.Title);
        const activeT=timers.filter(t=>t.PropertyId===prop.Title);
        const recentTasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.PropertyId===prop.Title&&dAgo(a.ActivityDate)<=7);
        const done=recentTasks.filter(a=>a.Status==="Completed");
        const blocked=recentTasks.filter(a=>a.Status==="Blocked");

        return(<div key={prop.Title} style={{...S.card,flex:"1 1 280px",minWidth:260,borderLeft:`3px solid ${blocked.length>0?C.error:activeT.length>0?C.success:C.teal100}`}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
            <div><div style={{fontSize:15,fontWeight:600,color:C.teal700}}>{prop.PropertyName}</div><div style={{fontSize:11,color:C.gray400}}>{prop.Units} units · {prop.PropertyGroup}</div></div>
            {va?<Badge type={va.VATrackerStatus==="Out"?"error":"success"}>{va.Name.split(" ")[0]}</Badge>:<Badge type="error">No VA</Badge>}
          </div>
          <div style={{display:"flex",gap:8,marginBottom:8}}>
            <div style={{flex:1,textAlign:"center",padding:6,background:C.teal50,borderRadius:4}}><div style={{fontSize:18,fontWeight:700,fontFamily:mono,color:C.teal700}}>{done.length}</div><div style={{fontSize:10,color:C.gray400}}>Done (7d)</div></div>
            <div style={{flex:1,textAlign:"center",padding:6,background:pendingQ.length>0?C.warningBg:C.teal50,borderRadius:4}}><div style={{fontSize:18,fontWeight:700,fontFamily:mono,color:pendingQ.length>0?C.warning:C.teal700}}>{pendingQ.length}</div><div style={{fontSize:10,color:C.gray400}}>Pending</div></div>
            <div style={{flex:1,textAlign:"center",padding:6,background:blocked.length>0?C.errorBg:C.teal50,borderRadius:4}}><div style={{fontSize:18,fontWeight:700,fontFamily:mono,color:blocked.length>0?C.error:C.teal700}}>{blocked.length}</div><div style={{fontSize:10,color:C.gray400}}>Blocked</div></div>
          </div>
          {blocked.length>0&&<div style={{fontSize:12,color:C.error,marginBottom:4}}>{blocked.map((b,i)=><div key={i}>{"\u26A0"} {b.Title}{b.Notes?` — ${b.Notes}`:""}</div>)}</div>}
          {activeT.length>0&&<div style={{fontSize:12,color:C.success}}>{activeT.map((a,i)=><div key={i}>{"\u25CF"} {a.Title} (in progress)</div>)}</div>}
        </div>);
      })}
    </div>

    {/* Submit task form */}
    <div style={{...S.card,borderLeft:`3px solid ${C.gold500}`}}>
      <div style={S.cardTitle}>{"\u{1F4CC}"} Submit Task for VA</div>
      <div style={{fontSize:12,color:C.gray400,marginBottom:12}}>Select a property — the task auto-routes to the assigned VA.</div>
      <div style={S.row}>
        <div style={{flex:1,minWidth:180}}><label style={S.label}>Property *</label>
          <select style={S.select} value={selProp} onChange={e=>setSelProp(e.target.value)}>
            <option value="">Select property...</option>
            {mgrProps.map(p=>{const va=getVA(p.Title);return<option key={p.Title} value={p.Title}>{p.PropertyName} ({p.Units}u){va?` → ${va.Name}`:""}</option>;})}
          </select>
        </div>
        <div style={{flex:1,minWidth:150}}><label style={S.label}>Category *</label>
          <select style={S.select} value={tCat} onChange={e=>setTCat(e.target.value)}><option value="">Select...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select>
        </div>
        <div style={{minWidth:80}}><label style={S.label}>Priority</label><select style={S.select} value={tPri} onChange={e=>setTPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
      </div>
      <div style={{marginTop:10}}><label style={S.label}>Task *</label><input style={S.input} value={tDesc} onChange={e=>setTDesc(e.target.value)} placeholder="What needs to be done..."/></div>
      <div style={{marginTop:8}}><label style={S.label}>Notes for VA</label><input style={S.input} value={tNotes} onChange={e=>setTNotes(e.target.value)} placeholder="Additional context (optional)"/></div>
      <button style={{...S.btn(C.headerBg),width:"100%",marginTop:12}} onClick={handleSubmit}>{"\u{1F4CC}"} Submit Task</button>
      {selProp&&(()=>{const va=getVA(selProp);return va?<div style={{fontSize:12,color:C.success,marginTop:8}}>{"\u2192"} Routes to: {va.Name}{va.VATrackerStatus==="Out"?` (\u26A0 currently OUT — will go to coverage)`:""}</div>:<div style={{fontSize:12,color:C.error,marginTop:8}}>{"\u26A0"} No VA assigned to this property</div>;})()}
    </div>

    {/* Recent activity on my properties */}
    <div style={S.card}>
      <div style={S.cardTitle}>Recent Activity — My Properties (7 days)</div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr>{["Date","Property","VA","Task","Duration","Status"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>{data.activities.filter(a=>a.ActivityType==="Task"&&mgrProps.some(p=>p.Title===a.PropertyId)&&dAgo(a.ActivityDate)<=7).slice(0,30).map((t,i)=>(
            <tr key={i}><td style={{...S.td,whiteSpace:"nowrap"}}>{fD(t.ActivityDate)}</td><td style={{...S.td,fontSize:12}}>{t.PropertyName}</td><td style={{...S.td,fontWeight:500}}>{t.VAName}</td><td style={{...S.td,maxWidth:200}}>{t.Title}{t.Notes&&<div style={{fontSize:11,color:C.gray400}}>{"\u{1F4AC}"} {t.Notes}</div>}</td><td style={{...S.td,fontFamily:mono,fontSize:12}}>{t.DurationMin?fM(t.DurationMin):"\u2014"}</td><td style={S.td}><Badge type={stBadge[t.Status]||"neutral"}>{t.Status}</Badge></td></tr>
          ))}</tbody>
        </table>
        {data.activities.filter(a=>a.ActivityType==="Task"&&mgrProps.some(p=>p.Title===a.PropertyId)&&dAgo(a.ActivityDate)<=7).length===0&&<div style={{textAlign:"center",padding:30,color:C.gray400}}>No activity yet.</div>}
      </div>
    </div>
  </div>);
}

// ============================================================
// HISTORY VIEW
// ============================================================
function HistoryView({data,role,myEmail,isMgr,mgrProps}){
  const[view,setView]=useState("tasks");
  if(!data)return null;
  const propFilter=role==="manager"?new Set(mgrProps.map(p=>p.Title)):null;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.Status!=="Queued"&&a.Status!=="In Progress"&&(isMgr?(propFilter?propFilter.has(a.PropertyId):true):a.VAEmail?.toLowerCase()===myEmail)).slice(0,50);
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&(role==="admin"||a.VAEmail?.toLowerCase()===myEmail)).slice(0,30);

  return(<div>
    <div style={{display:"flex",gap:8,marginBottom:12}}><button style={view==="tasks"?S.btn(C.headerBg):S.btnO(C.teal700,C.teal100)} onClick={()=>setView("tasks")}>Tasks</button>{role!=="manager"&&<button style={view==="shifts"?S.btn(C.headerBg):S.btnO(C.teal700,C.teal100)} onClick={()=>setView("shifts")}>Shifts</button>}</div>
    {view==="tasks"&&<div style={{...S.card,overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Date","VA","Task","Property","Source","Duration","Status"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead><tbody>
      {tasks.map((t,i)=><tr key={i}><td style={{...S.td,whiteSpace:"nowrap"}}>{fD(t.ActivityDate)}</td><td style={{...S.td,fontWeight:500,fontSize:12}}>{t.VAName}</td><td style={{...S.td,maxWidth:200}}>{t.Title}{t.CoverageForName&&<div style={{fontSize:10,color:C.warning}}>{"\u{1F504}"} coverage for {t.CoverageForName}</div>}{t.Notes&&<div style={{fontSize:11,color:C.gray400}}>{"\u{1F4AC}"} {t.Notes}</div>}</td><td style={{...S.td,fontSize:12}}>{t.PropertyName}</td><td style={S.td}><Badge type={srcBadge[t.Source]||"neutral"}>{t.Source}</Badge></td><td style={{...S.td,fontFamily:mono,fontWeight:600,fontSize:12}}>{t.DurationMin?fM(t.DurationMin):"\u2014"}</td><td style={S.td}><Badge type={stBadge[t.Status]||"neutral"}>{t.Status}</Badge></td></tr>)}
    </tbody></table>{!tasks.length&&<div style={{textAlign:"center",padding:40,color:C.gray400}}>No history.</div>}</div>}
    {view==="shifts"&&<div style={{...S.card,overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Date","VA","In","Out","Break","Working"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead><tbody>
      {shifts.map((s,i)=><tr key={i}><td style={{...S.td,whiteSpace:"nowrap"}}>{fD(s.ActivityDate)}</td><td style={{...S.td,fontWeight:500}}>{s.VAName}</td><td style={{...S.td,fontFamily:mono,fontSize:12}}>{fT(s.StartTime)}</td><td style={{...S.td,fontFamily:mono,fontSize:12}}>{fT(s.EndTime)}</td><td style={{...S.td,fontFamily:mono}}>{fM(s.BreakMinutes)}</td><td style={{...S.td,fontFamily:mono,fontWeight:600}}>{fM(s.WorkMinutes)}</td></tr>)}
    </tbody></table>{!shifts.length&&<div style={{textAlign:"center",padding:40,color:C.gray400}}>No shifts.</div>}</div>}
  </div>);
}

// ============================================================
// SCORECARD VIEW
// ============================================================
function ScorecardView({data,role,myEmail,isMgr,myVa,myProps,mgrProps}){
  const[selVa,setSelVa]=useState(myEmail);
  if(!data)return null;
  const viewableVAs=role==="manager"?data.vas.filter(v=>data.portfolios.some(p=>mgrProps.some(mp=>mp.Title===p.PropertyId)&&p.VAEmail?.toLowerCase()===v.Email.toLowerCase())):data.vas;
  const va=data.vas.find(v=>v.Email?.toLowerCase()===(isMgr?selVa:myEmail)?.toLowerCase())||myVa;
  const vaEmail=va?.Email?.toLowerCase()||myEmail;
  const per=7;
  const tasks=data.activities.filter(a=>a.ActivityType==="Task"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=per);
  const done=tasks.filter(t=>t.Status==="Completed");const blocked=tasks.filter(t=>t.Status==="Blocked");const cov=tasks.filter(t=>t.CoverageForEmail);
  const taskMin=done.reduce((s,t)=>s+(t.DurationMin||0),0);const rate=tasks.length?Math.round(done.length/tasks.length*100):0;
  const shifts=data.activities.filter(a=>a.ActivityType==="Shift"&&a.VAEmail?.toLowerCase()===vaEmail&&dAgo(a.ActivityDate)<=per);
  const shiftMin=shifts.reduce((s,a)=>s+(a.WorkMinutes||0),0);const breakMin=shifts.reduce((s,a)=>s+(a.BreakMinutes||0),0);
  const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;const gap=Math.max(0,shiftMin-taskMin);
  const port=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===vaEmail);
  const vaPropsL=port.map(p=>data.properties.find(pr=>pr.Title===p.PropertyId)).filter(Boolean);
  const catMap={};tasks.forEach(t=>{if(!catMap[t.Category])catMap[t.Category]={c:0,m:0};catMap[t.Category].c++;catMap[t.Category].m+=(t.DurationMin||0);});
  const catList=Object.entries(catMap).sort((a,b)=>b[1].m-a[1].m);

  return(<div>
    {isMgr&&<div style={{...S.card,display:"flex",gap:12,flexWrap:"wrap",alignItems:"flex-end"}}><div style={{flex:1,minWidth:200}}><label style={S.label}>VA</label><select style={S.select} value={selVa} onChange={e=>setSelVa(e.target.value)}>{viewableVAs.map(v=><option key={v.Email} value={v.Email.toLowerCase()}>{v.Name}{v.VATrackerStatus==="Out"?" [OUT]":""}</option>)}</select></div></div>}
    {!isMgr&&va&&<div style={{marginBottom:12}}><div style={{fontSize:18,fontWeight:700,color:C.teal700}}>{va.Name}</div><div style={{fontSize:12,color:C.gray400}}>{vaPropsL.length} properties · {vaPropsL.reduce((s,p)=>s+(p.Units||0),0)} units</div></div>}
    <div style={{...S.row,marginBottom:16}}>{[{l:"Tasks",v:tasks.length},{l:"Done",v:done.length,c:C.success},{l:"Rate",v:`${rate}%`,c:rate>=85?C.success:undefined},{l:"Shift",v:fM(shiftMin)},{l:"Util",v:`${util}%`,c:util>=75?C.success:util<50&&shiftMin>0?C.error:undefined},{l:"Gap",v:fM(gap),c:gap>120?C.error:undefined},{l:"Coverage",v:cov.length,c:cov.length>0?C.warning:undefined}].map((k,i)=><KPI key={i} label={k.l} value={k.v} color={k.c}/>)}</div>
    {shiftMin>0&&<div style={{...S.card,padding:14}}><div style={{fontSize:13,fontWeight:600,color:C.teal700,marginBottom:8}}>Time Breakdown</div><div style={{display:"flex",height:22,borderRadius:4,overflow:"hidden",background:C.gray100}}><div style={{width:`${Math.min(util,100)}%`,background:C.success,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:600,color:"#FFF"}}>{util>14?`${util}%`:""}</div>{breakMin>0&&<div style={{width:`${Math.min(Math.round(breakMin/(shiftMin+breakMin)*100),40)}%`,background:C.warning}}/>}<div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,color:C.gray400}}>{gap>0?`${fM(gap)} gap`:""}</div></div></div>}
    <div style={S.row}>
      <div style={{...S.card,flex:1,minWidth:240}}><div style={S.cardTitle}>Time by Category</div>{catList.map(([cat,d])=><div key={cat} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 0",borderBottom:`1px solid ${C.gray100}`}}><span>{catIcon[cat]||"\u{1F4C1}"}</span><div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.teal700}}>{cat}</div><div style={{fontSize:10,color:C.gray400}}>{d.c} tasks</div></div><div style={{fontFamily:mono,fontSize:12,fontWeight:700,color:C.teal700}}>{fM(d.m)}</div></div>)}{!catList.length&&<p style={{color:C.gray400,fontSize:13}}>No data</p>}</div>
      <div style={{...S.card,flex:1,minWidth:200}}><div style={S.cardTitle}>Property Coverage</div>{vaPropsL.map(p=>{const ok=tasks.some(t=>t.PropertyId===p.Title);return<div key={p.Title} style={{display:"flex",alignItems:"center",gap:5,padding:"3px 0",borderBottom:`1px solid ${C.gray100}`}}><span style={{color:ok?C.success:C.error,fontSize:12}}>{ok?"\u2713":"\u2717"}</span><span style={{flex:1,fontSize:12,color:ok?C.teal700:C.gray400}}>{p.PropertyName}</span></div>;})}</div>
    </div>
    {blocked.length>0&&<div style={{...S.card,borderLeft:`3px solid ${C.error}`}}><div style={{...S.cardTitle,color:C.error}}>{"\u26A0"} Blocked</div>{blocked.map((t,i)=><div key={i} style={{padding:"5px 0",borderBottom:`1px solid ${C.gray100}`}}><div style={{fontSize:13,fontWeight:500,color:C.teal700}}>{t.Title}</div><div style={{fontSize:11,color:C.gray400}}>{t.PropertyName}{t.Notes?` · ${t.Notes}`:""}</div></div>)}</div>}
  </div>);
}

// ============================================================
// ADMIN VIEW
// ============================================================
function AdminView({data,myEmail,acct,config,queue,covQ,onToggleAbsence,onAssignTask,onCloseDay,onUpdateConfig,onAssignProp,onUnassignProp,onAddProperty,onEditProperty,onDeactivateProperty,onDeleteTask,generateReport}){
  const[showAssign,setShowAssign]=useState(false);const[aVa,setAVa]=useState("");const[aCat,setACat]=useState("");const[aPri,setAPri]=useState("Normal");const[aDesc,setADesc]=useState("");const[aNotes,setANotes]=useState("");
  const[showRec,setShowRec]=useState(false);const[rVa,setRVa]=useState("");const[rCat,setRCat]=useState("");const[rDesc,setRDesc]=useState("");
  const[editIdx,setEditIdx]=useState(null);const[editDesc,setEditDesc]=useState("");
  const[portVa,setPortVa]=useState("");const[portProp,setPortProp]=useState("");
  const[showAddProp,setShowAddProp]=useState(false);const[npName,setNpName]=useState("");const[npGroup,setNpGroup]=useState("Multifamily");const[npUnits,setNpUnits]=useState("");const[npPm,setNpPm]=useState("");
  const[editPropId,setEditPropId]=useState(null);const[epName,setEpName]=useState("");const[epUnits,setEpUnits]=useState("");
  // Category management
  const[showAddCat,setShowAddCat]=useState(false);const[ncName,setNcName]=useState("");const[ncIcon,setNcIcon]=useState("folder");
  const[editCatId,setEditCatId]=useState(null);const[ecName,setEcName]=useState("");

  if(!data)return null;
  const cats=config?.categories||[];const rTasks=config?.recurringTasks||[];

  // Get properties for a specific VA
  function vaProps(email){if(!email)return data.properties;return data.properties.filter(p=>data.portfolios.some(pt=>pt.VAEmail?.toLowerCase()===email.toLowerCase()&&pt.PropertyId===p.Title));}
  const sCats=sortCats(cats);
  // Multi-property state for assign
  const[aProps,setAProps]=useState([]);
  function toggleAProp(id){setAProps(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);}

  function handleAssign(){if(!aVa||!aCat||!aDesc)return;const va=data.vas.find(v=>v.Email===aVa);const cat=cats.find(c=>c.id===aCat);
    const propList=aProps.length>0?aProps:[""];// empty string = General
    propList.forEach(pid=>{const prop=pid?data.properties.find(p=>p.Title===pid):null;
      onAssignTask({Title:aDesc,VAEmail:aVa,VAName:va?.Name||aVa,PropertyId:pid||"",PropertyName:prop?prop.PropertyName:"General",PMName:prop?prop.PMName:"",Category:cat?.name||"Admin/Other",Priority:aPri,Source:"Assigned",AssignedByEmail:myEmail,AssignedByName:acct?.name||myEmail,Notes:aNotes});
    });
    setADesc("");setAVa("");setACat("");setAProps([]);setAPri("Normal");setANotes("");setShowAssign(false);}

  // Multi-property for recurring
  const[rProps,setRProps]=useState([]);
  function toggleRProp(id){setRProps(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);}

  function addRec(){if(!rVa||!rCat||!rDesc)return;
    const propList=rProps.length>0?rProps:[""];
    const newTasks=propList.map(pid=>({vaEmail:rVa,category:rCat,description:rDesc,propertyId:pid,active:true}));
    saveRec([...rTasks,...newTasks]);setRDesc("");setRVa("");setRCat("");setRProps([]);setShowRec(false);}

  // Duplicate recurring task for different VA
  function duplicateRec(i){const r=rTasks[i];setRVa("");setRCat(r.category);setRDesc(r.description);setRProps(r.propertyId?[r.propertyId]:[]);setShowRec(true);}

  return(<div>
    {/* Absence */}
    <div style={{...S.card,borderLeft:`3px solid ${C.error}`}}>
      <div style={S.cardTitle}>{"\u{1F6D1}"} VA Absence Management</div>
      {data.vas.map(va=><div key={va.Email} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${C.gray100}`}}>
        <div><div style={{fontSize:14,fontWeight:600,color:va.VATrackerStatus==="Out"?C.error:C.teal700}}>{va.Name}</div><div style={{fontSize:11,color:C.gray400}}>{data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase()).length} properties</div></div>
        <div style={{display:"flex",alignItems:"center",gap:8}}><Badge type={va.VATrackerStatus==="Out"?"error":"success"}>{va.VATrackerStatus||"Active"}</Badge><button style={{...(va.VATrackerStatus==="Out"?S.btn(C.success):S.btn(C.error)),...S.btnSm}} onClick={()=>onToggleAbsence(va)}>{va.VATrackerStatus==="Out"?"Mark In":"Mark Out"}</button></div></div>)}
      {covQ.length>0&&<div style={{marginTop:10,padding:8,background:C.warningBg,borderRadius:4,fontSize:12,color:C.warning}}>{"\u26A0"} {covQ.length} unclaimed coverage tasks</div>}
    </div>

    {/* Assign */}
    <div style={{...S.card,borderLeft:`3px solid ${C.gold500}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:16,fontWeight:600,color:C.teal700}}>{"\u{1F4CC}"} Assign Task</div></div><button style={S.btn(C.gold500,C.dark)} onClick={()=>setShowAssign(!showAssign)}>{showAssign?"Cancel":"Assign"}</button></div>
      {showAssign&&<div style={{marginTop:14}}>
        <div style={S.row}><div style={{flex:1,minWidth:140}}><label style={S.label}>VA *</label><select style={S.select} value={aVa} onChange={e=>{setAVa(e.target.value);setAProps([]);}}><option value="">Select...</option>{data.vas.filter(v=>v.VATrackerStatus!=="Out").map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
          <div style={{flex:1,minWidth:140}}><label style={S.label}>Category *</label><select style={S.select} value={aCat} onChange={e=>setACat(e.target.value)}><option value="">Select...</option>{sCats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div></div>
        <div style={{marginTop:10}}><label style={S.label}>Properties {aProps.length>0&&`(${aProps.length} selected)`}</label>
          <div style={{border:`1px solid ${C.gray200}`,borderRadius:4,padding:8,maxHeight:160,overflowY:"auto",background:C.white}}>
            <label style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0",fontSize:13,color:C.gray400,cursor:"pointer",borderBottom:`1px solid ${C.gray100}`}}><input type="checkbox" checked={aProps.length===0} onChange={()=>setAProps([])}/>General (no specific property)</label>
            {(aVa?vaProps(aVa):data.properties).map(p=><label key={p.Title} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 0",fontSize:13,color:C.teal700,cursor:"pointer"}}><input type="checkbox" checked={aProps.includes(p.Title)} onChange={()=>toggleAProp(p.Title)}/>{p.PropertyName} ({p.Units}u)</label>)}
          </div>
        </div>
        <div style={{...S.row,marginTop:10}}><div style={{flex:1}}><label style={S.label}>Priority</label><select style={S.select} value={aPri} onChange={e=>setAPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div></div>
        <div style={{marginTop:10}}><label style={S.label}>Task *</label><input style={S.input} value={aDesc} onChange={e=>setADesc(e.target.value)} placeholder="What needs to be done..."/></div>
        <div style={{marginTop:8}}><label style={S.label}>Notes</label><input style={S.input} value={aNotes} onChange={e=>setANotes(e.target.value)} placeholder="Context for the VA (optional)"/></div>
        <button style={{...S.btn(C.headerBg),width:"100%",marginTop:10}} onClick={handleAssign}>{"\u{1F4CC}"} {aProps.length>1?`Add ${aProps.length} Tasks`:"Add to Queue"}</button>
      </div>}
    </div>

    {/* Recurring */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:10}}><div style={{fontSize:16,fontWeight:600,color:C.teal700}}>{"\u{1F4CB}"} Recurring Daily Tasks</div><button style={{...S.btn(C.headerBg),...S.btnSm}} onClick={()=>setShowRec(!showRec)}>{showRec?"Cancel":"+ Add"}</button></div>
      {showRec&&<div style={{background:C.teal50,borderRadius:6,padding:14,marginBottom:14,border:`1px solid ${C.teal100}`}}>
        <div style={S.row}><div style={{flex:1,minWidth:130}}><label style={S.label}>VA *</label><select style={S.select} value={rVa} onChange={e=>{setRVa(e.target.value);setRProps([]);}}><option value="">Select...</option>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
          <div style={{flex:1,minWidth:130}}><label style={S.label}>Category *</label><select style={S.select} value={rCat} onChange={e=>setRCat(e.target.value)}><option value="">Select...</option>{sCats.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div></div>
        <div style={{marginTop:10}}><label style={S.label}>Properties {rProps.length>0&&`(${rProps.length} selected)`}</label>
          <div style={{border:`1px solid ${C.gray200}`,borderRadius:4,padding:8,maxHeight:140,overflowY:"auto",background:C.white}}>
            <label style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",fontSize:12,color:C.gray400,cursor:"pointer",borderBottom:`1px solid ${C.gray100}`}}><input type="checkbox" checked={rProps.length===0} onChange={()=>setRProps([])}/>General</label>
            {(rVa?vaProps(rVa):data.properties).map(p=><label key={p.Title} style={{display:"flex",alignItems:"center",gap:6,padding:"3px 0",fontSize:12,color:C.teal700,cursor:"pointer"}}><input type="checkbox" checked={rProps.includes(p.Title)} onChange={()=>toggleRProp(p.Title)}/>{p.PropertyName}</label>)}
          </div>
        </div>
        <div style={{marginTop:10,marginBottom:10}}><label style={S.label}>Description *</label><input style={S.input} value={rDesc} onChange={e=>setRDesc(e.target.value)} placeholder="e.g., Check AppFolio for new work orders"/></div>
        <button style={{...S.btn(C.success),width:"100%"}} onClick={addRec}>{"\u2713"} {rProps.length>1?`Add ${rProps.length} Tasks`:"Add"}</button></div>}
      {data.vas.map(va=>{const vr=rTasks.map((r,i)=>({...r,_i:i})).filter(r=>r.vaEmail?.toLowerCase()===va.Email.toLowerCase());if(!vr.length)return null;
        return<div key={va.Email} style={{marginBottom:12}}><div style={{fontSize:12,fontWeight:600,color:C.teal600,textTransform:"uppercase",marginBottom:4,paddingBottom:3,borderBottom:`1px solid ${C.teal100}`}}>{va.Name} ({vr.filter(r=>r.active).length} active)</div>
          {vr.map(r=>{const cat=cats.find(c=>c.id===r.category);const prop=r.propertyId?data.properties.find(p=>p.Title===r.propertyId):null;const isEd=editIdx===r._i;
            return<div key={r._i} style={{display:"flex",alignItems:"center",gap:6,padding:"6px 0",borderBottom:`1px solid ${C.gray100}`,opacity:r.active?1:0.5}}>
              <span>{catIcon[cat?.name]||"\u{1F4C1}"}</span>
              <div style={{flex:1,minWidth:0}}>{isEd?<div style={{display:"flex",gap:4}}><input style={{...S.input,padding:"4px 8px",fontSize:12}} value={editDesc} onChange={e=>setEditDesc(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")saveEdit(r._i);if(e.key==="Escape")setEditIdx(null);}} autoFocus/><button style={{...S.btn(C.success),...S.btnSm}} onClick={()=>saveEdit(r._i)}>{"\u2713"}</button></div>
                :<div><div style={{fontSize:12,fontWeight:500,color:C.teal700,cursor:"pointer"}} onClick={()=>{setEditIdx(r._i);setEditDesc(r.description);}}>{r.description}</div>{prop&&<div style={{fontSize:10,color:C.gray400}}>{prop.PropertyName}</div>}</div>}</div>
              {!isEd&&<div style={{display:"flex",gap:3}}>
                <button style={{...S.btnO(C.info,C.info),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>duplicateRec(r._i)} title="Duplicate for another VA">{"\u{1F4CB}"}</button>
                <button style={{...S.btnO(r.active?C.warning:C.success,r.active?C.warning:C.success),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>toggleRec(r._i)}>{r.active?"Pause":"On"}</button>
                <button style={{...S.btnO(C.error,C.error),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>deleteRec(r._i)}>{"\u2715"}</button></div>}
            </div>;})}</div>;})}
    </div>

    {/* Portfolio */}
    <div style={S.card}>
      <div style={S.cardTitle}>{"\u{1F3E0}"} Portfolio Assignments</div>
      <div style={{background:C.teal50,borderRadius:6,padding:12,marginBottom:14,border:`1px solid ${C.teal100}`,display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end"}}>
        <div style={{flex:1,minWidth:140}}><label style={{...S.label,fontSize:11}}>VA</label><select style={S.select} value={portVa} onChange={e=>setPortVa(e.target.value)}><option value="">Select...</option>{data.vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
        <div style={{flex:1,minWidth:180}}><label style={{...S.label,fontSize:11}}>Property</label><select style={S.select} value={portProp} onChange={e=>setPortProp(e.target.value)}><option value="">Select...</option>{data.properties.filter(p=>!portVa||!data.portfolios.some(pt=>pt.VAEmail?.toLowerCase()===portVa.toLowerCase()&&pt.PropertyId===p.Title)).map(p=>{const a=data.portfolios.find(pt=>pt.PropertyId===p.Title);return<option key={p.Title} value={p.Title}>{p.PropertyName} ({p.Units}u){a?` — ${a.VAName}`:""}</option>;})}</select></div>
        <button style={{...S.btn(C.success),...S.btnSm,opacity:(!portVa||!portProp)?0.5:1}} onClick={()=>{if(!portVa||!portProp)return;const va=data.vas.find(v=>v.Email===portVa);onAssignProp(portVa,va?.Name||portVa,portProp);setPortProp("");}}>Assign</button>
      </div>
      {data.vas.map(va=>{const vp=data.portfolios.filter(p=>p.VAEmail?.toLowerCase()===va.Email.toLowerCase());const tu=vp.reduce((s,p)=>{const pr=data.properties.find(x=>x.Title===p.PropertyId);return s+(pr?.Units||0);},0);
        return<div key={va.Email} style={{marginBottom:14}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:4,paddingBottom:3,borderBottom:`2px solid ${C.teal100}`}}><div style={{fontSize:13,fontWeight:700,color:C.teal700}}>{va.Name}</div><div style={{fontSize:11,color:C.gray400}}>{vp.length} props · {tu}u</div></div>
          {!vp.length&&<div style={{fontSize:12,color:C.gray400,fontStyle:"italic",padding:"4px 0"}}>None assigned</div>}
          {vp.map(p=>{const pr=data.properties.find(x=>x.Title===p.PropertyId);return<div key={p.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:`1px solid ${C.gray100}`}}><span>{"\u{1F3E0}"}</span><div style={{flex:1}}><div style={{fontSize:12,fontWeight:500,color:C.teal700}}>{pr?.PropertyName||p.PropertyName}</div><div style={{fontSize:10,color:C.gray400}}>{pr?.Units||"?"}u · {pr?.PMName||"?"}</div></div><button style={{...S.btnO(C.error,C.error),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>{if(window.confirm(`Remove ${pr?.PropertyName} from ${va.Name}?`))onUnassignProp(p.id,pr?.PropertyName||p.PropertyName,va.Name);}}>Remove</button></div>;})}</div>;})}
      {(()=>{const aIds=new Set(data.portfolios.map(p=>p.PropertyId));const un=data.properties.filter(p=>!aIds.has(p.Title));if(!un.length)return null;return<div style={{padding:10,background:C.warningBg,borderRadius:6,marginTop:8}}><div style={{fontSize:12,fontWeight:600,color:C.warning,marginBottom:4}}>{"\u26A0"} Unassigned ({un.length})</div>{un.map(p=><div key={p.Title} style={{fontSize:12,color:C.gray600,padding:"2px 0"}}>{p.PropertyName} ({p.Units}u) — {p.PMName}</div>)}</div>;})()}
    </div>

    {/* Property Management */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:10}}><div style={{fontSize:16,fontWeight:600,color:C.teal700}}>{"\u{1F3D7}\u{FE0F}"} Property Management</div><button style={{...S.btn(C.headerBg),...S.btnSm}} onClick={()=>setShowAddProp(!showAddProp)}>{showAddProp?"Cancel":"+ Add Property"}</button></div>
      {showAddProp&&<div style={{background:C.teal50,borderRadius:6,padding:14,marginBottom:14,border:`1px solid ${C.teal100}`}}>
        <div style={S.row}><div style={{flex:2,minWidth:150}}><label style={S.label}>Name *</label><input style={S.input} value={npName} onChange={e=>setNpName(e.target.value)} placeholder="Property name"/></div><div style={{flex:1,minWidth:100}}><label style={S.label}>Group *</label><select style={S.select} value={npGroup} onChange={e=>setNpGroup(e.target.value)}><option>Multifamily</option><option>Single Family</option><option>Lease-Up</option></select></div><div style={{minWidth:70}}><label style={S.label}>Units *</label><input style={S.input} type="number" value={npUnits} onChange={e=>setNpUnits(e.target.value)} placeholder="#"/></div></div>
        <div style={{marginTop:10}}><label style={S.label}>Property Manager *</label><select style={S.select} value={npPm} onChange={e=>setNpPm(e.target.value)}><option value="">Select PM...</option>{data.pms.map(pm=><option key={pm.Email} value={pm.Email}>{pm.Name}</option>)}</select></div>
        <button style={{...S.btn(C.success),width:"100%",marginTop:10}} onClick={()=>{if(!npName||!npUnits||!npPm)return;const pm=data.pms.find(p=>p.Email===npPm);onAddProperty({name:npName,group:npGroup,units:parseInt(npUnits)||0,pmEmail:npPm,pmName:pm?.Name||npPm});setNpName("");setNpUnits("");setNpPm("");setShowAddProp(false);}}>{"\u2713"} Add Property</button>
      </div>}
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr>{["Property","Group","Units","PM","VA","Actions"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead><tbody>
        {data.properties.map(p=>{const port=data.portfolios.find(pt=>pt.PropertyId===p.Title);const va=port?data.vas.find(v=>v.Email?.toLowerCase()===port.VAEmail?.toLowerCase()):null;const isEd=editPropId===p.id;
          return<tr key={p.id}><td style={S.td}>{isEd?<input style={{...S.input,padding:"4px 8px",fontSize:12,width:140}} value={epName} onChange={e=>setEpName(e.target.value)}/>:<span style={{fontWeight:500,color:C.teal700}}>{p.PropertyName}</span>}</td>
            <td style={{...S.td,fontSize:12}}>{p.PropertyGroup}</td>
            <td style={S.td}>{isEd?<input type="number" style={{...S.input,padding:"4px 8px",fontSize:12,width:60}} value={epUnits} onChange={e=>setEpUnits(e.target.value)}/>:p.Units}</td>
            <td style={{...S.td,fontSize:12}}>{p.PMName}</td>
            <td style={{...S.td,fontSize:12}}>{va?va.Name:<span style={{color:C.gray400,fontStyle:"italic"}}>Unassigned</span>}</td>
            <td style={S.td}>{isEd?<div style={{display:"flex",gap:3}}><button style={{...S.btn(C.success),...S.btnSm,padding:"3px 8px"}} onClick={()=>{onEditProperty(p.id,{PropertyName:epName,Units:parseInt(epUnits)||p.Units});setEditPropId(null);}}>{"\u2713"}</button><button style={{...S.btnO(C.gray400,C.gray200),...S.btnSm,padding:"3px 8px"}} onClick={()=>setEditPropId(null)}>{"\u2715"}</button></div>:<div style={{display:"flex",gap:3}}><button style={{...S.btnO(C.teal600,C.teal100),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>{setEditPropId(p.id);setEpName(p.PropertyName);setEpUnits(String(p.Units||0));}}>Edit</button><button style={{...S.btnO(C.error,C.error),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>onDeactivateProperty(p.id,p.PropertyName)}>Deactivate</button></div>}</td></tr>;
        })}</tbody></table></div>
    </div>

    {/* Category Management */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:10}}>
        <div style={{fontSize:16,fontWeight:600,color:C.teal700}}>{"\u{1F3F7}\u{FE0F}"} Task Categories</div>
        <button style={{...S.btn(C.headerBg),...S.btnSm}} onClick={()=>setShowAddCat(!showAddCat)}>{showAddCat?"Cancel":"+ Add"}</button>
      </div>
      {showAddCat&&<div style={{background:C.teal50,borderRadius:6,padding:14,marginBottom:14,border:`1px solid ${C.teal100}`}}>
        <div style={S.row}>
          <div style={{flex:2,minWidth:150}}><label style={S.label}>Category Name *</label><input style={S.input} value={ncName} onChange={e=>setNcName(e.target.value)} placeholder="e.g., Leasing"/></div>
          <div style={{flex:1,minWidth:120}}><label style={S.label}>Icon</label>
            <select style={S.select} value={ncIcon} onChange={e=>setNcIcon(e.target.value)}>
              <option value="wrench">{"\u{1F527}"} Wrench</option>
              <option value="megaphone">{"\u{1F4E2}"} Megaphone</option>
              <option value="chat">{"\u{1F4AC}"} Chat</option>
              <option value="chart">{"\u{1F4CA}"} Chart</option>
              <option value="search">{"\u{1F50D}"} Search</option>
              <option value="document">{"\u{1F4DD}"} Document</option>
              <option value="money">{"\u{1F4B0}"} Money</option>
              <option value="folder">{"\u{1F4C1}"} Folder</option>
              <option value="key">{"\u{1F511}"} Key</option>
              <option value="house">{"\u{1F3E0}"} House</option>
              <option value="handshake">{"\u{1F91D}"} Handshake</option>
              <option value="clipboard">{"\u{1F4CB}"} Clipboard</option>
              <option value="phone">{"\u{1F4DE}"} Phone</option>
              <option value="email">{"\u{1F4E7}"} Email</option>
            </select>
          </div>
        </div>
        <button style={{...S.btn(C.success),width:"100%",marginTop:10}} onClick={()=>{
          if(!ncName)return;
          const newId="c"+Date.now().toString(36);
          const newCats=[...cats,{id:newId,name:ncName,icon:ncIcon}];
          onUpdateConfig({...config,categories:newCats});
          // Also add to VA_Activity Category choice column (note: SharePoint choice columns auto-expand when new values are written)
          setNcName("");setNcIcon("folder");setShowAddCat(false);
        }}>{"\u2713"} Add Category</button>
      </div>}
      {sortCats(cats).map((c,i)=>{const origIdx=cats.findIndex(x=>x.id===c.id);
        const iconMap={wrench:"\u{1F527}",megaphone:"\u{1F4E2}",chat:"\u{1F4AC}",chart:"\u{1F4CA}",search:"\u{1F50D}",document:"\u{1F4DD}",money:"\u{1F4B0}",folder:"\u{1F4C1}",key:"\u{1F511}",house:"\u{1F3E0}",handshake:"\u{1F91D}",clipboard:"\u{1F4CB}",phone:"\u{1F4DE}",email:"\u{1F4E7}"};
        const isEd=editCatId===c.id;
        return<div key={c.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${C.gray100}`}}>
          <span style={{fontSize:16}}>{iconMap[c.icon]||catIcon[c.name]||"\u{1F4C1}"}</span>
          <div style={{flex:1}}>
            {isEd?<div style={{display:"flex",gap:4}}><input style={{...S.input,padding:"4px 8px",fontSize:12}} value={ecName} onChange={e=>setEcName(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){const u=[...cats];u[origIdx]={...u[origIdx],name:ecName};onUpdateConfig({...config,categories:u});setEditCatId(null);}if(e.key==="Escape")setEditCatId(null);}} autoFocus/><button style={{...S.btn(C.success),...S.btnSm}} onClick={()=>{const u=[...cats];u[origIdx]={...u[origIdx],name:ecName};onUpdateConfig({...config,categories:u});setEditCatId(null);}}>{"\u2713"}</button></div>
            :<div style={{fontSize:13,fontWeight:500,color:C.teal700,cursor:"pointer"}} onClick={()=>{setEditCatId(c.id);setEcName(c.name);}}>{c.name}</div>}
          </div>
          {!isEd&&<div style={{display:"flex",gap:3}}>
            <button style={{...S.btnO(C.error,C.error),...S.btnSm,padding:"3px 8px",fontSize:10}} onClick={()=>{
              if(!window.confirm(`Delete "${c.name}" category? Existing tasks with this category won't be affected.`))return;
              onUpdateConfig({...config,categories:cats.filter((_,j)=>j!==origIdx)});
            }}>{"\u2715"}</button>
          </div>}
        </div>;
      })}
    </div>

    {/* Data Sources */}
    <div style={{...S.card,padding:12,background:C.teal50,border:`1px solid ${C.teal100}`}}><div style={{fontSize:12,color:C.teal600}}><strong>Data:</strong> Employees → VAs/PMs/Roles. VA_Properties → property registry. VA_Portfolios → assignments. VA_TrackerConfig → categories/recurring/settings. VA_Activity → tasks/shifts/absences. Auth: MSAL.</div></div>

    {/* Weekly Report */}
    <div style={S.card}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10,borderBottom:`1px solid ${C.gray100}`,marginBottom:10}}>
        <div><div style={{fontSize:16,fontWeight:600,color:C.teal700}}>{"\u{1F4CA}"} Weekly Report</div><div style={{fontSize:12,color:C.gray400}}>Previous week's performance summary</div></div>
        <button style={S.btn(C.headerBg)} onClick={()=>{
          const r=generateReport();if(!r)return;
          const blob=new Blob([r.report],{type:"text/plain"});
          const url=URL.createObjectURL(blob);const a=document.createElement("a");
          a.href=url;a.download=`VA-Weekly-Report-${r.from}-to-${r.to}.txt`;a.click();URL.revokeObjectURL(url);
        }}>{"\u{1F4E5}"} Download Report</button>
      </div>
      <button style={{...S.btnO(C.teal600,C.teal100),width:"100%"}} onClick={()=>{
        const r=generateReport();if(!r)return;
        const w=window.open("","_blank","width=800,height=600");
        w.document.write(`<pre style="font-family:Consolas,monospace;font-size:13px;padding:20px;white-space:pre-wrap;">${r.report.replace(/</g,"&lt;")}</pre>`);
        w.document.title=`VA Report ${r.from} to ${r.to}`;
      }}>Preview Report</button>
    </div>

    {/* Queued Task Management */}
    {queue.length>0&&<div style={S.card}>
      <div style={S.cardTitle}>{"\u{1F5D1}\u{FE0F}"} Queued Tasks ({queue.length})</div>
      <div style={{fontSize:12,color:C.gray400,marginBottom:10}}>Remove tasks that should no longer be in the queue.</div>
      {data.vas.map(va=>{const vq=queue.filter(t=>t.VAEmail?.toLowerCase()===va.Email.toLowerCase());if(!vq.length)return null;
        return<div key={va.Email} style={{marginBottom:10}}><div style={{fontSize:12,fontWeight:600,color:C.teal600,textTransform:"uppercase",marginBottom:4}}>{va.Name} ({vq.length})</div>
          {vq.map(t=><TaskRow key={t._localId} task={t} onDelete={onDeleteTask}/>)}</div>;})}
    </div>}

    {/* Close Day */}
    <div style={{...S.card,borderLeft:`3px solid ${C.warning}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:15,fontWeight:600,color:C.teal700}}>{"\u{1F504}"} Close Day</div><div style={{fontSize:12,color:C.gray400}}>Marks unfinished daily tasks Incomplete, clears coverage.</div></div><button style={S.btn(C.warning,C.dark)} onClick={onCloseDay}>Close Day</button></div></div>
  </div>);
}
