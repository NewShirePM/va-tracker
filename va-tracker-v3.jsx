import{useState,useEffect,useCallback,useRef}from"react";

// FAVICON
(()=>{const l=document.querySelector("link[rel='icon']")||document.createElement("link");l.rel="icon";l.href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='6' fill='%23CDA04B'/><text x='16' y='23' font-family='Georgia,serif' font-size='22' font-weight='bold' fill='%2328434C' text-anchor='middle'>V</text></svg>";document.head.appendChild(l);})();

// CONFIG
const CONFIG={
  clientId:"32e75ffa-747a-4cf0-8209-6a19150c4547",
  tenantId:"33575d04-ca7b-4396-8011-9eaea4030b46",
  siteId:"vanrockre.sharepoint.com,a02c1cd8-9f1f-4827-8286-7b6b7ce74232,01202419-6625-4499-b0d5-8ceb1cffdba3",
  // Cloudflare Worker URL for guest access (no MSAL required for guests)
  // Deploy va-tracker-worker.js to Cloudflare, paste the worker URL here:
  workerUrl:"YOUR_CLOUDFLARE_WORKER_URL",
  appName:"VA PRODUCTIVITY TRACKER",
  teamsEnabled:true,
};

const GRAPH="https://graph.microsoft.com/v1.0";
const SITE=`${GRAPH}/sites/${CONFIG.siteId}`;
const SCOPES=["Sites.ReadWrite.All","User.Read","Chat.Create","ChatMessage.Send"];

// ROLE DETECTION — va | manager | regional | admin
const TITLE_ROLE={"Virtual Assistant":"va","Property Manager":"manager","Regional/Portfolio Manager":"regional"};
function detectRole(emp){
  const ov=(emp.VATrackerRole||"").toLowerCase().trim();
  if(["va","manager","regional","admin"].includes(ov))return ov;
  return TITLE_ROLE[emp.JobTitle]||null;
}

// PALETTE
const C={
  hdr:"#1C3740",t2:"#28434C",t3:"#3A6577",t4:"#4A7E91",
  tl:"#D6E7EC",tl0:"#EDF4F7",tl00:"#F4F9FB",
  gold:"#CDA04B",g2:"#B8922E",gl:"#F8F0DB",
  bg:"#EDEEF0",w:"#FFFFFF",
  b1:"#DDE3E6",b2:"#C2CDD1",b4:"#6B8590",b6:"#3A5058",
  ok:"#1A7A46",okl:"rgba(26,122,70,0.1)",okb:"#EAF5EE",
  er:"#B83B2A",erl:"rgba(184,59,42,0.09)",erb:"#FDF0EE",
  wn:"#A86F08",wnl:"rgba(168,111,8,0.1)",
  inf:"#2B5FA8",infl:"rgba(43,95,168,0.1)",infb:"#EBF1FB",
  pu:"#5B3FA8",pul:"rgba(91,63,168,0.09)",pub:"#F2EEFB",
  tm:"#4A4E8A",
};
const F="'DM Sans','Segoe UI',system-ui,sans-serif";
const M="'DM Mono','Cascadia Code',monospace";

// STYLES
const S={
  page:{fontFamily:F,background:C.bg,minHeight:"100vh",color:C.t2,fontSize:13},
  hdr:{background:C.hdr,borderBottom:`2px solid ${C.gold}`,padding:"0 18px",display:"flex",alignItems:"center",justifyContent:"space-between",height:52,flexShrink:0,gap:10},
  hdrT:{color:"#fff",fontSize:14,fontWeight:700,letterSpacing:".04em"},
  hdrS:{color:"rgba(255,255,255,.4)",fontSize:9,letterSpacing:".08em",textTransform:"uppercase"},
  tabs:{background:C.w,borderBottom:`1px solid ${C.b1}`,display:"flex",padding:"0 16px",overflowX:"auto",WebkitOverflowScrolling:"touch"},
  tab:on=>({padding:"0 14px",fontSize:12,fontWeight:on?700:500,color:on?C.t2:C.b4,borderBottom:`2.5px solid ${on?C.gold:"transparent"}`,cursor:"pointer",whiteSpace:"nowrap",background:"none",borderTop:"none",borderLeft:"none",borderRight:"none",fontFamily:F,height:43,display:"inline-flex",alignItems:"center",gap:5}),
  con:{maxWidth:1200,margin:"0 auto",padding:"14px 14px 32px"},
  card:{background:C.w,border:`1px solid ${C.b1}`,borderRadius:8,boxShadow:"0 1px 3px rgba(28,55,64,.07)",padding:15,marginBottom:12},
  at:{borderTop:`3px solid ${C.gold}`},ac:{borderTop:`3px solid ${C.t3}`},
  ao:{borderTop:`3px solid ${C.ok}`},ae:{borderTop:`3px solid ${C.er}`},
  ap:{borderTop:`3px solid ${C.pu}`},ain:{borderTop:`3px solid ${C.inf}`},
  lbl:{display:"block",fontSize:11,fontWeight:700,color:C.t2,marginBottom:4},
  inp:{width:"100%",padding:"8px 10px",fontSize:12,fontFamily:F,color:C.t2,background:C.w,border:`1px solid ${C.b2}`,borderRadius:6,outline:"none",boxSizing:"border-box"},
  sel:{width:"100%",padding:"8px 10px",fontSize:12,fontFamily:F,color:C.t2,background:C.w,border:`1px solid ${C.b2}`,borderRadius:6,cursor:"pointer",boxSizing:"border-box"},
  btn:(bg,fg)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:5,padding:"8px 14px",fontSize:12,fontWeight:600,fontFamily:F,color:fg||"#fff",background:bg||C.hdr,border:"none",borderRadius:6,cursor:"pointer",minHeight:36,whiteSpace:"nowrap"}),
  btnO:(fg,bdr)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",gap:5,padding:"8px 14px",fontSize:12,fontWeight:600,fontFamily:F,color:fg||C.t2,background:C.w,border:`1px solid ${bdr||C.b2}`,borderRadius:6,cursor:"pointer",minHeight:36,whiteSpace:"nowrap"}),
  sm:{padding:"5px 10px",fontSize:11,minHeight:28},
  xs:{padding:"3px 8px",fontSize:10,minHeight:24},
  th:{textAlign:"left",padding:"9px 11px",fontSize:11,fontWeight:700,color:C.t2,background:C.tl00,borderBottom:`1px solid ${C.b1}`,whiteSpace:"nowrap"},
  td:{padding:"9px 11px",fontSize:12,color:C.b6,borderBottom:`1px solid ${C.b1}`},
  kpi:{background:C.w,border:`1px solid ${C.b1}`,borderRadius:8,padding:"11px 8px",textAlign:"center",flex:"1 1 75px",minWidth:68},
  kl:{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase",letterSpacing:".06em",marginBottom:3},
  kv:{fontSize:22,fontWeight:700,fontFamily:M,color:C.t2,lineHeight:1.1},
  row:{display:"flex",gap:10,flexWrap:"wrap"},
};

// COMPONENTS
const BD={ok:{c:C.ok,b:C.okl},er:{c:C.er,b:C.erl},wn:{c:C.wn,b:C.wnl},in:{c:C.inf,b:C.infl},pu:{c:C.pu,b:C.pul},ne:{c:C.b4,b:C.b1}};
function Badge({type="ne",nd,children}){const m=BD[type]||BD.ne;return<span style={{display:"inline-flex",alignItems:"center",gap:nd?0:3,padding:"2px 8px",fontSize:10,fontWeight:700,borderRadius:99,whiteSpace:"nowrap",color:m.c,background:m.b}}>{!nd&&<span style={{width:5,height:5,borderRadius:"50%",background:"currentColor",flexShrink:0}}/>}{children}</span>;}
function KPI({label,value,color,sub}){return<div style={S.kpi}><div style={S.kl}>{label}</div><div style={{...S.kv,color:color||C.t2}}>{value}</div>{sub&&<div style={{fontSize:9,color:C.b4,marginTop:2}}>{sub}</div>}</div>;}
function Av({initials,color="tl",size=30}){const bgs={tl:{bg:C.tl0,c:C.t2},gd:{bg:C.gl,c:C.g2},er:{bg:C.erb,c:C.er},ok:{bg:C.okb,c:C.ok},in:{bg:C.infb,c:C.inf},pu:{bg:C.pub,c:C.pu},ne:{bg:C.b1,c:C.b4}};const m=bgs[color]||bgs.tl;return<div style={{width:size,height:size,borderRadius:"50%",background:m.bg,color:m.c,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:size*.36,flexShrink:0,fontFamily:F}}>{initials}</div>;}

// HELPERS
const catI={"Work Orders":"🔧",Marketing:"📢","Tenant Comms":"💬",Reporting:"📊",Inspections:"🔍",Renewals:"📝",Accounts:"💰","Admin/Other":"📁"};
const stB={Completed:"ok",Blocked:"er","In Progress":"wn",Queued:"ne",Incomplete:"er"};
const srcB={Daily:"in",Assigned:"wn","Ad-Hoc":"ne",Coverage:"wn"};
function fD(d){return d?new Date(d).toLocaleDateString("en-US",{month:"short",day:"numeric"}):"—";}
function fT(d){return d?new Date(d).toLocaleTimeString("en-US",{hour:"numeric",minute:"2-digit"}):"—";}
function fM(m){if(!m&&m!==0)return"—";const h=Math.floor(m/60),r=m%60;return h>0?`${h}h ${r}m`:`${r}m`;}
function fTm(s){const h=Math.floor(s/3600),m=Math.floor((s%3600)/60),ss=s%60,p=n=>String(n).padStart(2,"0");return h>0?`${h}:${p(m)}:${p(ss)}`:`${p(m)}:${p(ss)}`;}
function today(){return new Date().toISOString().slice(0,10);}
function daysAgoStr(n){const d=new Date();d.setDate(d.getDate()-n);return d.toISOString().slice(0,10);}
function dAgo(d){return Math.floor((Date.now()-new Date(d).getTime())/864e5);}
function isOD(t){return t.task_date&&t.task_date<today()&&(t.status==="Queued"||t.status==="In Progress");}
function sortCats(c){if(!c)return[];return[...c.filter(x=>x.name!=="Admin/Other").sort((a,b)=>a.name.localeCompare(b.name)),...c.filter(x=>x.name==="Admin/Other")];}
function inRange(d,from,to){if(!d)return false;const t=new Date(d).getTime();return t>=new Date(from).getTime()&&t<=new Date(to).getTime()+864e5;}

// SHAREPOINT DATA LAYER
// Field name normalizer — maps SharePoint PascalCase to snake_case used by views
function spNorm(item,type){
  const f=item.fields||item;const id=item.id||f.id;
  if(type==="activity")return{id,activity_type:f.ActivityType,va_email:(f.VAEmail||"").toLowerCase(),va_name:f.VAName,title:f.Title,task_date:f.TaskDate||(f.ActivityDate?f.ActivityDate.slice(0,10):null),property_code:f.PropertyId||"",property_name:f.PropertyName||"General",pm_name:f.PMName||"",category:f.Category||"Admin/Other",source:f.Source||"Daily",status:f.Status||"Queued",priority:f.Priority||"Normal",recurring_id:f.RecurringId||"",start_time:f.StartTime||null,end_time:f.EndTime||null,duration_min:f.DurationMin||0,paused_min:f.PausedMin||0,notes:f.Notes||"",coverage_for_email:f.CoverageForEmail||"",coverage_for_name:f.CoverageForName||"",assigned_by_email:f.AssignedByEmail||"",assigned_by_name:f.AssignedByName||"",clock_in:f.ClockIn||null,clock_out:f.ClockOut||null,break_minutes:f.BreakMinutes||0,work_minutes:f.WorkMinutes||0,breaks_json:f.BreaksJSON?JSON.parse(f.BreaksJSON):[],marked_by_email:f.MarkedByEmail||""};
  if(type==="property")return{id,property_code:f.Title,property_name:f.PropertyName,property_group:f.PropertyGroup||"Multifamily",units:f.Units||0,pm_email:(f.PMEmail||"").toLowerCase(),pm_name:f.PMName||"",appfolio_id:f.AppFolioId||"",is_active:f.IsActive!==false};
  if(type==="portfolio")return{id,va_email:(f.VAEmail||"").toLowerCase(),va_name:f.VAName||"",property_code:f.PropertyId||"",property_name:f.PropertyName||"",is_active:f.IsActive!==false};
  if(type==="interruption")return{id,va_email:(f.VAEmail||"").toLowerCase(),va_name:f.VAName||"",interruption_date:f.InterruptionDate||(f.LoggedAt?f.LoggedAt.slice(0,10):today()),logged_at:f.LoggedAt||null,interruption_type:f.InterruptionType||"Other",property_code:f.PropertyId||"",property_name:f.PropertyName||"General",duration_min:f.DurationMin||0,notes:f.Notes||"",spawned_task_id:f.SpawnedTaskId||null};
  if(type==="metric")return{id,va_email:(f.VAEmail||"").toLowerCase(),va_name:f.VAName||"",metric_date:f.MetricDate||today(),leads_followed_up:f.LeadsFollowedUp||0,applications_processed:f.ApplicationsProcessed||0,showings_scheduled:f.ShowingsScheduled||0,resident_comms:f.ResidentComms||0,work_orders_entered:f.WorkOrdersEntered||0,work_orders_updated:f.WorkOrdersUpdated||0,renewals_touched:f.RenewalsTouched||0,manager_calls:f.ManagerCalls||0,notes:f.MetricNotes||"",submitted:f.Submitted||false,submitted_at:f.SubmittedAt||null};
  if(type==="review")return{id,activity_id:f.ActivityId||null,va_email:(f.VAEmail||"").toLowerCase(),va_name:f.VAName||"",pm_email:(f.PMEmail||"").toLowerCase(),pm_name:f.PMName||"",review_type:f.ReviewType||"Other",property_code:f.PropertyId||"",property_name:f.PropertyName||"",notes:f.ReviewNotes||"",status:f.Status||"Pending",pm_response:f.PMResponse||"",teams_sent:f.TeamsSent||false,created_at:f.Created||null,resolved_at:f.ResolvedAt||null,resolved_by:f.ResolvedBy||""};
  if(type==="guest")return{id,token:f.GuestToken||"",guest_name:f.GuestName||"",guest_email:f.GuestEmail||"",company:f.Company||"",va_emails:(f.VAEmails||"").split(",").map(e=>e.trim().toLowerCase()).filter(Boolean),is_active:f.IsActive!==false,invited_by:f.InvitedBy||"",last_accessed:f.LastAccessed||null};
  return{id,...f};
}

async function spAll(token,listName,filter=""){
  const q=filter?`&$filter=${encodeURIComponent(filter)}`:"";
  return gAll(token,`${SITE}/lists/${listName}/items?expand=fields&$top=2000${q}`);
}

async function spLoad(token){
  const f30=daysAgoStr(30),f7=daysAgoStr(7);
  const[eActs,eProps,ePorts,eCfg,eInts,eMets,eRevs]=await Promise.all([
    spAll(token,"VA_Activity",`fields/TaskDate ge '${f30}' or (fields/Status eq 'Queued') or (fields/Status eq 'In Progress')`),
    spAll(token,"VA_Properties","fields/IsActive eq 1"),
    spAll(token,"VA_Portfolios","fields/IsActive eq 1"),
    gAll(token,`${SITE}/lists/VA_TrackerConfig/items?expand=fields&$top=5`),
    spAll(token,"VA_Interruptions",`fields/InterruptionDate ge '${f7}'`),
    spAll(token,"VA_DailyMetrics",`fields/MetricDate ge '${f7}'`),
    spAll(token,"VA_TaskReviews",""),
  ]);
  const cfgFields=eCfg[0]?.fields||{};
  const config={id:"main",categories:cfgFields.CategoriesJSON?JSON.parse(cfgFields.CategoriesJSON):[],recurring_tasks:cfgFields.RecurringJSON?JSON.parse(cfgFields.RecurringJSON):[],settings:cfgFields.SettingsJSON?JSON.parse(cfgFields.SettingsJSON):{},_spId:eCfg[0]?.id};
  return{
    properties:eProps.map(x=>spNorm(x,"property")),
    portfolios:ePorts.map(x=>spNorm(x,"portfolio")),
    config,
    activities:eActs.map(x=>spNorm(x,"activity")),
    interruptions:eInts.map(x=>spNorm(x,"interruption")),
    dailyMetrics:eMets.map(x=>spNorm(x,"metric")),
    taskReviews:eRevs.map(x=>spNorm(x,"review")),
  };
}

// CRUD helpers — always return normalized objects
async function spInsActivity(token,fields){
  const body={Title:fields.title||"Task",ActivityType:fields.activity_type,VAEmail:fields.va_email,VAName:fields.va_name,TaskDate:fields.task_date||today(),PropertyId:fields.property_code||"",PropertyName:fields.property_name||"General",PMName:fields.pm_name||"",Category:fields.category,Source:fields.source,Status:fields.status,Priority:fields.priority||"Normal",RecurringId:fields.recurring_id||"",Notes:fields.notes||"",CoverageForEmail:fields.coverage_for_email||"",CoverageForName:fields.coverage_for_name||"",AssignedByEmail:fields.assigned_by_email||"",AssignedByName:fields.assigned_by_name||""};
  const r=await gPost(token,`${SITE}/lists/VA_Activity/items`,{fields:body});
  return spNorm(r,"activity");
}
async function spUpdActivity(token,id,fields){
  const body={};
  if(fields.status!==undefined)body.Status=fields.status;
  if(fields.start_time!==undefined)body.StartTime=fields.start_time;
  if(fields.end_time!==undefined)body.EndTime=fields.end_time;
  if(fields.duration_min!==undefined)body.DurationMin=fields.duration_min;
  if(fields.paused_min!==undefined)body.PausedMin=fields.paused_min;
  if(fields.notes!==undefined)body.Notes=fields.notes;
  if(fields.source!==undefined)body.Source=fields.source;
  if(fields.va_email!==undefined)body.VAEmail=fields.va_email;
  if(fields.va_name!==undefined)body.VAName=fields.va_name;
  if(fields.coverage_for_email!==undefined)body.CoverageForEmail=fields.coverage_for_email;
  if(fields.coverage_for_name!==undefined)body.CoverageForName=fields.coverage_for_name;
  await gPatch(token,`${SITE}/lists/VA_Activity/items/${id}/fields`,body);
}
async function spInsMetric(token,fields){
  const body={Title:`${fields.va_email}-${fields.metric_date}`,VAEmail:fields.va_email,VAName:fields.va_name,MetricDate:fields.metric_date,LeadsFollowedUp:fields.leads_followed_up||0,ApplicationsProcessed:fields.applications_processed||0,ShowingsScheduled:fields.showings_scheduled||0,ResidentComms:fields.resident_comms||0,WorkOrdersEntered:fields.work_orders_entered||0,WorkOrdersUpdated:fields.work_orders_updated||0,RenewalsTouched:fields.renewals_touched||0,ManagerCalls:fields.manager_calls||0,MetricNotes:fields.notes||"",Submitted:fields.submitted||false};
  const r=await gPost(token,`${SITE}/lists/VA_DailyMetrics/items`,{fields:body});
  return spNorm(r,"metric");
}
async function spUpdMetric(token,id,fields){
  const body={};
  const map={leads_followed_up:"LeadsFollowedUp",applications_processed:"ApplicationsProcessed",showings_scheduled:"ShowingsScheduled",resident_comms:"ResidentComms",work_orders_entered:"WorkOrdersEntered",work_orders_updated:"WorkOrdersUpdated",renewals_touched:"RenewalsTouched",manager_calls:"ManagerCalls",notes:"MetricNotes",submitted:"Submitted",submitted_at:"SubmittedAt"};
  Object.entries(fields).forEach(([k,v])=>{if(map[k])body[map[k]]=v;});
  await gPatch(token,`${SITE}/lists/VA_DailyMetrics/items/${id}/fields`,body);
}
async function spUpsMetric(token,vaEmail,vaName,field,delta,existing){
  const map={leads_followed_up:"LeadsFollowedUp",applications_processed:"ApplicationsProcessed",showings_scheduled:"ShowingsScheduled",resident_comms:"ResidentComms",work_orders_entered:"WorkOrdersEntered",work_orders_updated:"WorkOrdersUpdated",renewals_touched:"RenewalsTouched",manager_calls:"ManagerCalls"};
  const spField=map[field];if(!spField)return null;
  const cur=existing?existing[field]||0:0;const nv=Math.max(0,cur+delta);
  if(existing&&existing.id){await spUpdMetric(token,existing.id,{[field]:nv});return{...existing,[field]:nv};}
  return spInsMetric(token,{va_email:vaEmail,va_name:vaName,metric_date:today(),[field]:nv});
}
async function spUpdConfig(token,spId,config){
  await gPatch(token,`${SITE}/lists/VA_TrackerConfig/items/${spId}/fields`,{CategoriesJSON:JSON.stringify(config.categories),RecurringJSON:JSON.stringify(config.recurring_tasks),SettingsJSON:JSON.stringify(config.settings)});
}

// GRAPH/MSAL HELPERS
async function gGet(token,url){const r=await fetch(url,{headers:{Authorization:`Bearer ${token}`}});if(!r.ok)throw new Error(`GET ${r.status} ${url}`);return r.json();}
async function gAll(token,url){let a=[],n=url;while(n){const d=await gGet(token,n);a=a.concat(d.value||[]);n=d["@odata.nextLink"]||null;}return a;}
async function gPost(token,url,body){const r=await fetch(url,{method:"POST",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify(body)});if(!r.ok)throw new Error(`POST ${r.status}`);return r.json();}
async function gPatch(token,url,body){const r=await fetch(url,{method:"PATCH",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/json"},body:JSON.stringify(body)});if(!r.ok)throw new Error(`PATCH ${r.status}`);return r.json();}
function empUrl(id){return`${SITE}/lists/Employees/items/${id}/fields`;}

async function loadEmployees(token){
  try{const r=await gAll(token,`${SITE}/lists/Employees/items?expand=fields&$top=200`);return r.map(e=>({id:e.id,...e.fields}));}
  catch(e){console.warn("[VT] Employees:",e.message);return[];}
}

async function sendTeamsMsg(token,pmEmail,html){
  if(!CONFIG.teamsEnabled)return;
  try{
    const[me,pm]=await Promise.all([gGet(token,`${GRAPH}/me`),gGet(token,`${GRAPH}/users/${encodeURIComponent(pmEmail)}`)]);
    const chat=await gPost(token,`${GRAPH}/chats`,{chatType:"oneOnOne",members:[
      {"@odata.type":"#microsoft.graph.aadUserConversationMember",roles:["owner"],"user@odata.bind":`${GRAPH}/users/${me.id}`},
      {"@odata.type":"#microsoft.graph.aadUserConversationMember",roles:["owner"],"user@odata.bind":`${GRAPH}/users/${pm.id}`},
    ]});
    await gPost(token,`${GRAPH}/chats/${chat.id}/messages`,{body:{contentType:"html",content:html}});
    console.log("[VT] Teams message sent to",pmEmail);
  }catch(e){console.warn("[VT] Teams msg failed:",e.message);}
}

// MSAL HOOK
function useMsal(){
  const[inst,setInst]=useState(null);const[acct,setAcct]=useState(null);const[token,setToken]=useState(null);const[err,setErr]=useState(null);
  useEffect(()=>{
    const s=document.createElement("script");s.src="https://alcdn.msauth.net/browser/2.38.0/js/msal-browser.min.js";
    s.onload=()=>{
      const i=new window.msal.PublicClientApplication({auth:{clientId:CONFIG.clientId,authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`,redirectUri:window.location.origin},cache:{cacheLocation:"sessionStorage"}});
      i.initialize().then(()=>{setInst(i);const a=i.getAllAccounts();if(a.length>0){setAcct(a[0]);i.acquireTokenSilent({scopes:SCOPES,account:a[0]}).then(r=>setToken(r.accessToken)).catch(()=>{});}});
    };
    document.head.appendChild(s);
  },[]);
  const login=useCallback(async()=>{if(!inst)return;try{const r=await inst.loginPopup({scopes:SCOPES});setAcct(r.account);const t=await inst.acquireTokenSilent({scopes:SCOPES,account:r.account});setToken(t.accessToken);setErr(null);}catch(e){if(e.errorCode!=="user_cancelled")setErr(e.message);}},[inst]);
  const refresh=useCallback(async()=>{if(!inst||!acct)return null;try{const r=await inst.acquireTokenSilent({scopes:SCOPES,account:acct});setToken(r.accessToken);return r.accessToken;}catch{const r=await inst.acquireTokenPopup({scopes:SCOPES});setToken(r.accessToken);return r.accessToken;}},[inst,acct]);
  return{acct,token,login,refresh,err};
}

// ═══════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════
export default function App(){
  const{acct,token,login,refresh,err:authErr}=useMsal();
  const[tab,setTab]=useState(0);
  const[employees,setEmployees]=useState([]);
  const[sbData,setSbData]=useState(null);
  const[loading,setLoading]=useState(false);
  const[error,setError]=useState(null);
  const[role,setRole]=useState(null);
  const[myEmail,setMyEmail]=useState(null);
  const[myEmp,setMyEmp]=useState(null);
  const[flash,setFlash]=useState("");
  const[timers,setTimers]=useState([]);
  const[shift,setShift]=useState(null);
  const[queue,setQueue]=useState([]);
  const[covQ,setCovQ]=useState([]);
  const[tick,setTick]=useState(0);
  const[dfFrom,setDfFrom]=useState(daysAgoStr(7));
  const[dfTo,setDfTo]=useState(today());
  const genLock=useRef(false);
  const[guestToken]=useState(()=>new URLSearchParams(window.location.search).get("guest"));
  const[guestVAs,setGuestVAs]=useState(null);

  useEffect(()=>{const id=setInterval(()=>setTick(t=>t+1),1000);return()=>clearInterval(id);},[]);
  const fl=useCallback(msg=>{setFlash(msg);setTimeout(()=>setFlash(""),3000);},[]);
  async function gT(){return(await refresh())||token;}

  // Load Supabase CDN — NOT NEEDED (SharePoint backend)
  // Guest token path — uses Cloudflare Worker (no MSAL for guests)
  useEffect(()=>{
    if(!guestToken)return;
    setLoading(true);
    fetch(`${CONFIG.workerUrl}?token=${encodeURIComponent(guestToken)}`)
      .then(r=>{if(!r.ok)throw new Error(r.status);return r.json();})
      .then(data=>{
        setGuestVAs(data.vaEmails||[]);
        setRole("guest");setMyEmail(data.guest_email||"guest");
        setSbData({properties:data.properties||[],portfolios:data.portfolios||[],activities:data.activities||[],config:{categories:[],recurring_tasks:[],settings:{}},interruptions:[],dailyMetrics:[],taskReviews:[]});
      })
      .catch(()=>setError("invalid_guest"))
      .finally(()=>setLoading(false));
  },[guestToken]);

  // Internal user load
  useEffect(()=>{
    if(!token||!acct||guestToken)return;
    const email=acct.username.toLowerCase();
    setMyEmail(email);setLoading(true);
    Promise.all([loadEmployees(token),spLoad(token)]).then(([emps,sbD])=>{
      setEmployees(emps);setSbData(sbD);
      const me=emps.find(e=>(e.Email&&e.Email.toLowerCase()===email)||(e.M365UserId&&e.M365UserId.toLowerCase()===email));
      if(!me){setRole(null);setError("access_denied");}
      else{const r=detectRole(me);if(!r){setRole(null);setError("access_denied");}else{setRole(r);setMyEmp(me);buildQueue(sbD,emps);}}
      setLoading(false);
    }).catch(e=>{setError("load_error: "+e.message);setLoading(false);});
  },[token,acct]);

  function buildQueue(sbD,emps){
    const q=[],cv=[];
    sbD.activities.filter(a=>a.activity_type==="Task"&&(a.status==="Queued"||a.status==="In Progress")).forEach(a=>{
      if(a.source==="Coverage"||a.coverage_for_email)cv.push(a);else q.push(a);
    });
    setQueue(q);setCovQ(cv);
    if(!genLock.current&&sbD.config.recurring_tasks?.length>0){
      genLock.current=true;
      generateRecurring(sbD,emps).finally(()=>{genLock.current=false;});
    }
  }

  async function generateRecurring(sbD,emps){
    const td=today();const tk=await gT();
    for(const rt of sbD.config.recurring_tasks){
      if(!rt.active||!rt.recurringId)continue;
      const va=emps.find(v=>v.Email?.toLowerCase()===rt.vaEmail?.toLowerCase());
      if(!va)continue;
      // Dedup check via SharePoint $filter (RecurringId column added by provisioning script)
      try{
        const ex=await spAll(tk,"VA_Activity",`fields/RecurringId eq '${rt.recurringId}' and fields/TaskDate eq '${td}' and fields/VAEmail eq '${va.Email}'`);
        if(ex.length>0)continue;
      }catch{continue;}
      const prop=rt.propertyCode?sbD.properties.find(p=>p.property_code===rt.propertyCode):null;
      const cat=sbD.config.categories?.find(c=>c.id===rt.category);
      const isOut=va.VATrackerStatus==="Out";
      const task={activity_type:"Task",va_email:va.Email,va_name:va.Name,title:rt.description,task_date:td,property_code:rt.propertyCode||"",property_name:prop?.property_name||"General",pm_name:prop?.pm_name||"",category:cat?.name||"Admin/Other",source:isOut?"Coverage":"Daily",status:"Queued",priority:"Normal",recurring_id:rt.recurringId,coverage_for_email:isOut?va.Email:"",coverage_for_name:isOut?va.Name:""};
      try{
        const saved=await spInsActivity(tk,task);
        if(isOut)setCovQ(p=>[...p,saved]);else setQueue(p=>[...p,saved]);
      }catch(e){console.warn("[VT] Recurring insert:",e.message);}
    }
  }

  async function reload(){
    const tk=await gT();
    const[emps,sbD]=await Promise.all([loadEmployees(tk),spLoad(tk)]);
    setEmployees(emps);setSbData(sbD);buildQueue(sbD,emps);return{emps,sbD};
  }

  // ── SHIFT ──
  function clockIn(){if(shift){fl("Already clocked in!");return;}setShift({ClockIn:new Date().toISOString(),Breaks:[],_ob:false,_bs:null});fl("Clocked in!");}
  function startBreak(){setShift(p=>p?{...p,_ob:true,_bs:new Date().toISOString()}:p);}
  function endBreak(){setShift(p=>{if(!p||!p._bs)return p;return{...p,_ob:false,Breaks:[...p.Breaks,{s:p._bs,e:new Date().toISOString()}],_bs:null};});}
  async function clockOut(){
    if(!shift)return;const now=new Date();const bks=[...shift.Breaks];
    if(shift._ob&&shift._bs)bks.push({s:shift._bs,e:now.toISOString()});
    const bMs=bks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);
    const bMin=Math.round(bMs/60000),wMin=Math.round((now-new Date(shift.ClockIn)-bMs)/60000);
    if(wMin>600&&!window.confirm(`${fM(wMin)} elapsed. Clock out?`))return;
    try{
      const tk=await gT();
      await gPost(tk,`${SITE}/lists/VA_Activity/items`,{fields:{Title:`${myEmp?.Name||myEmail}-${today()}`,ActivityType:"Shift",VAEmail:myEmail,VAName:myEmp?.Name||myEmail,TaskDate:today(),ClockIn:shift.ClockIn,ClockOut:now.toISOString(),BreakMinutes:bMin,WorkMinutes:wMin,BreaksJSON:JSON.stringify(bks)}});
      setShift(null);fl("Clocked out!");await reload();}
    catch(e){fl("Error: "+e.message);}
  }

  // ── TIMERS ──
  async function startTimer(task){
    const now=new Date().toISOString();
    try{const tk=await gT();await spUpdActivity(tk,task.id,{status:"In Progress",start_time:now});}catch{}
    setTimers(p=>[...p,{...task,status:"In Progress",start_time:now,_pMs:0,_pS:null}]);
    setQueue(p=>p.filter(t=>t.id!==task.id));fl("Timer started!");
  }
  function pauseTimer(id){setTimers(p=>p.map(t=>t.id===id?{...t,_pS:Date.now()}:t));}
  function resumeTimer(id){setTimers(p=>p.map(t=>{if(t.id!==id||!t._pS)return t;return{...t,_pMs:(t._pMs||0)+(Date.now()-t._pS),_pS:null};}));}
  async function finishTimer(id,status,notes){
    const t=timers.find(x=>x.id===id);if(!t)return;
    const now=Date.now();let pMs=t._pMs||0;if(t._pS)pMs+=(now-t._pS);
    const dur=Math.max(1,Math.round((now-new Date(t.start_time).getTime()-pMs)/60000));
    try{const tk=await gT();await spUpdActivity(tk,t.id,{status,end_time:status==="Completed"?new Date(now).toISOString():null,duration_min:dur,paused_min:Math.round(pMs/60000),notes:notes||""});setTimers(p=>p.filter(x=>x.id!==id));fl(status==="Completed"?"✓ Completed!":"Saved");await reload();}
    catch(e){fl("Error: "+e.message);}
  }
  async function cancelTimer(id){
    const t=timers.find(x=>x.id===id);if(!t)return;
    try{const tk=await gT();await spUpdActivity(tk,t.id,{status:"Queued",start_time:null});}catch{}
    setTimers(p=>p.filter(x=>x.id!==id));setQueue(p=>[{...t,status:"Queued",start_time:null,_pMs:0,_pS:null},...p]);
  }

  // ── COVERAGE ──
  async function claimCov(id){
    const t=covQ.find(x=>x.id===id);if(!t)return;
    try{const tk=await gT();await spUpdActivity(tk,t.id,{va_email:myEmail,va_name:myEmp?.Name||myEmail});}catch(e){console.warn(e);}
    setCovQ(p=>p.filter(x=>x.id!==id));setQueue(p=>[{...t,va_email:myEmail,va_name:myEmp?.Name||myEmail},...p]);
    fl(`Claimed! Covering for ${t.coverage_for_name}`);
  }

  // ── ABSENCE ──
  async function toggleAbsence(va){
    const tk=await gT();const ns=va.VATrackerStatus==="Out"?"Active":"Out";
    try{await gPatch(tk,empUrl(va.id),{VATrackerStatus:ns});}catch(e){fl("Error: "+e.message);return;}
    if(ns==="Out"){
      try{const tk2=await gT();await gPost(tk2,`${SITE}/lists/VA_Activity/items`,{fields:{Title:`${va.Name}-Out-${today()}`,ActivityType:"Absence",VAEmail:va.Email,VAName:va.Name,TaskDate:today(),Status:"Out",MarkedByEmail:myEmail}});}catch{}
      const mv=queue.filter(q=>q.va_email?.toLowerCase()===va.Email.toLowerCase());
      for(const t of mv){try{const tk2=await gT();await spUpdActivity(tk2,t.id,{source:"Coverage",coverage_for_email:va.Email,coverage_for_name:va.Name});}catch{}}
      setQueue(r=>r.filter(q=>q.va_email?.toLowerCase()!==va.Email.toLowerCase()));
      setCovQ(p=>[...p,...mv.map(q=>({...q,source:"Coverage",coverage_for_email:va.Email,coverage_for_name:va.Name}))]);
      fl(`${va.Name} marked OUT — ${mv.length} tasks to coverage`);
    }else{
      const ret=covQ.filter(q=>q.coverage_for_email?.toLowerCase()===va.Email.toLowerCase());
      for(const t of ret){try{const tk2=await gT();await spUpdActivity(tk2,t.id,{source:"Daily",va_email:va.Email,va_name:va.Name,coverage_for_email:"",coverage_for_name:""});}catch{}}
      setCovQ(p=>p.filter(q=>q.coverage_for_email?.toLowerCase()!==va.Email.toLowerCase()));
      setQueue(p=>[...p,...ret.map(q=>({...q,source:"Daily",va_email:va.Email,va_name:va.Name}))]);
      fl(`${va.Name} marked IN`);
    }
    await reload();
  }

  // ── TASKS ──
  async function addTask(task){
    const tVa=employees.find(v=>v.Email?.toLowerCase()===task.va_email?.toLowerCase());
    const isOut=tVa?.VATrackerStatus==="Out";
    const t={...task,activity_type:"Task",task_date:today(),status:"Queued",source:isOut?"Coverage":(task.source||"Ad-Hoc"),coverage_for_email:isOut?tVa.Email:"",coverage_for_name:isOut?tVa.Name:""};
    try{const tk=await gT();const saved=await spInsActivity(tk,t);if(isOut){setCovQ(p=>[saved,...p]);fl(`${tVa.Name} is OUT — task sent to coverage`);}else{setQueue(p=>[saved,...p]);fl("Task added!");}}
    catch(e){fl("Error: "+e.message);}
  }

  async function removeTask(task){
    try{const tk=await gT();await spUpdActivity(tk,task.id,{status:"Incomplete",notes:"Removed by admin"});setQueue(p=>p.filter(t=>t.id!==task.id));setCovQ(p=>p.filter(t=>t.id!==task.id));fl("Task removed");}
    catch(e){fl("Error: "+e.message);}
  }

  // ── INTERRUPTIONS ──
  async function logInterruption(fields,spawnTask){
    try{
      const tk=await gT();
      const intFields=fields;
      const spInt=await gPost(tk,`${SITE}/lists/VA_Interruptions/items`,{fields:{Title:`INT-${Date.now()}`,VAEmail:intFields.va_email,VAName:intFields.va_name,InterruptionDate:intFields.interruption_date,LoggedAt:new Date().toISOString(),InterruptionType:intFields.interruption_type,PropertyId:intFields.property_code||"",PropertyName:intFields.property_name||"General",DurationMin:intFields.duration_min||0,Notes:intFields.notes||""}});
      if(spawnTask){
        const saved=await spInsActivity(tk,{...spawnTask,activity_type:"Task",task_date:today(),status:"Queued",source:"Ad-Hoc"});
        await gPatch(tk,`${SITE}/lists/VA_Interruptions/items/${spInt.id}/fields`,{SpawnedTaskId:saved.id});
        setQueue(p=>[saved,...p]);
      }
      await reload();fl("Interruption logged!");
    }catch(e){fl("Error: "+e.message);}
  }

  // ── DAILY METRICS ──
  async function nudgeMetric(field,delta){
    const td=today();const existing=sbData?.dailyMetrics.find(m=>m.va_email===myEmail&&m.metric_date===td);
    try{
      const tk=await gT();
      const updated=await spUpsMetric(tk,myEmail,myEmp?.Name||myEmail,field,delta,existing);
      if(updated)setSbData(prev=>({...prev,dailyMetrics:[...prev.dailyMetrics.filter(m=>!(m.va_email===myEmail&&m.metric_date===td)),updated]}));
    }catch(e){console.warn("[VT] Metric:",e.message);}
  }

  async function submitMetrics(notes){
    const td=today();const existing=sbData?.dailyMetrics.find(m=>m.va_email===myEmail&&m.metric_date===td);
    try{
      const tk=await gT();
      if(existing?.id){await spUpdMetric(tk,existing.id,{notes,submitted:true,submitted_at:new Date().toISOString()});}
      else{await spInsMetric(tk,{va_email:myEmail,va_name:myEmp?.Name||myEmail,metric_date:td,notes,submitted:true,submitted_at:new Date().toISOString()});}
      fl("Metrics submitted!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // ── REVIEWS ──
  async function submitReview(activityId,fields){
    const tk=await gT();
    try{
      const revRes=await gPost(tk,`${SITE}/lists/VA_TaskReviews/items`,{fields:{Title:`REV-${Date.now()}`,ActivityId:activityId,VAEmail:fields.va_email,VAName:fields.va_name,PMEmail:fields.pm_email,PMName:fields.pm_name,ReviewType:fields.review_type,PropertyId:fields.property_code||"",PropertyName:fields.property_name||"",ReviewNotes:fields.notes,Status:"Pending",TeamsSent:false}});
      const html=`<b>🔔 Manager Review Requested</b><br/><b>From:</b> ${fields.va_name} | <b>Type:</b> ${fields.review_type}<br/><b>Property:</b> ${fields.property_name||"General"}<br/><br/><i>${fields.notes}</i><br/><br/><a href="${window.location.href}">→ View in VA Tracker</a>`;
      await sendTeamsMsg(tk,fields.pm_email,html);
      await gPatch(tk,`${SITE}/lists/VA_TaskReviews/items/${revRes.id}/fields`,{TeamsSent:true});
      fl("Review submitted — Teams message sent!");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  async function resolveReview(id,response){
    try{const tk=await gT();await gPatch(tk,`${SITE}/lists/VA_TaskReviews/items/${id}/fields`,{Status:"Resolved",PMResponse:response,ResolvedAt:new Date().toISOString(),ResolvedBy:myEmail});fl("Resolved!");await reload();}
    catch(e){fl("Error: "+e.message);}
  }

  // ── PORTFOLIO ──
  async function assignProp(vaEmail,vaName,propCode){
    const prop=sbData?.properties.find(p=>p.property_code===propCode);if(!prop)return;
    try{
      const tk=await gT();
      const ex=sbData?.portfolios.find(p=>p.va_email===vaEmail.toLowerCase()&&p.property_code===propCode);
      if(ex){await gPatch(tk,`${SITE}/lists/VA_Portfolios/items/${ex.id}/fields`,{IsActive:true});}
      else{await gPost(tk,`${SITE}/lists/VA_Portfolios/items`,{fields:{Title:`${vaName.split(" ")[0]}-${prop.property_name.replace(/\s+/g,"")}`.slice(0,50),VAEmail:vaEmail,VAName:vaName,PropertyId:propCode,PropertyName:prop.property_name,IsActive:true}});}
      await reload();fl(`${prop.property_name} → ${vaName}`);
    }catch(e){fl("Error: "+e.message);}
  }
  async function unassignProp(portId,pn,vn){
    try{const tk=await gT();await gPatch(tk,`${SITE}/lists/VA_Portfolios/items/${portId}/fields`,{IsActive:false});await reload();fl(`${pn} removed from ${vn}`);}catch(e){fl("Error: "+e.message);}
  }
  async function addProperty(p){
    const code=`PROP-${String((sbData?.properties.length||0)+1).padStart(3,"0")}`;
    try{const tk=await gT();await gPost(tk,`${SITE}/lists/VA_Properties/items`,{fields:{Title:code,PropertyName:p.name,PropertyGroup:p.group,Units:parseInt(p.units)||0,PMEmail:p.pmEmail,PMName:p.pmName,AppFolioId:p.afId||"",IsActive:true}});await reload();fl(`${p.name} added!`);}
    catch(e){fl("Error: "+e.message);}
  }
  async function saveProperty(id,fields){
    try{
      const tk=await gT();
      const spFields={};
      if(fields.property_name)spFields.PropertyName=fields.property_name;
      if(fields.units!==undefined)spFields.Units=fields.units;
      if(fields.pm_email)spFields.PMEmail=fields.pm_email;
      if(fields.pm_name)spFields.PMName=fields.pm_name;
      await gPatch(tk,`${SITE}/lists/VA_Properties/items/${id}/fields`,spFields);
      await reload();fl("Property saved!");
    }catch(e){fl("Error: "+e.message);}
  }

  // ── CONFIG ──
  async function updateConfig(nc){
    try{const tk=await gT();await spUpdConfig(tk,sbData.config._spId,nc);await reload();fl("Saved!");}
    catch(e){fl("Error: "+e.message);}
  }

  // ── CLOSE DAY ──
  async function closeDay(){
    if(!window.confirm("Close the day? Marks remaining daily tasks Incomplete and locks all metrics."))return;
    const td=today();const tk=await gT();
    try{
      // Mark daily queued tasks as Incomplete
      const todayQ=queue.filter(t=>t.source==="Daily"&&t.task_date===td);
      await Promise.all(todayQ.map(t=>spUpdActivity(tk,t.id,{status:"Incomplete",notes:"Auto-closed end of day"}).catch(()=>{})));
      // Submit any unsubmitted metrics for today
      const todayMets=sbData?.dailyMetrics.filter(m=>m.metric_date===td&&!m.submitted)||[];
      await Promise.all(todayMets.map(m=>spUpdMetric(tk,m.id,{submitted:true,submitted_at:new Date().toISOString()}).catch(()=>{})));
      if(shift)await clockOut();
      setQueue(p=>p.filter(t=>t.source!=="Daily"||t.task_date!==td));
      fl("Day closed");await reload();
    }catch(e){fl("Error: "+e.message);}
  }

  // COMPUTED
  const isAdmin=role==="admin";
  const isMgr=role==="admin"||role==="manager"||role==="regional";
  const isRegional=role==="regional";
  const isVA=role==="va";
  const isGuest=role==="guest";
  const vas=employees.filter(e=>detectRole(e)==="va"&&e.EmployeeActive!==false);
  const pms=employees.filter(e=>["manager","regional","admin"].includes(detectRole(e))&&e.EmployeeActive!==false);
  const myVa=isVA?vas.find(v=>v.Email?.toLowerCase()===myEmail):null;
  const myPort=sbData?sbData.portfolios.filter(p=>p.va_email?.toLowerCase()===myEmail):[];
  const myProps=sbData?myPort.map(p=>sbData.properties.find(pr=>pr.property_code===p.property_code)).filter(Boolean):[];
  const mgrProps=isMgr&&sbData?sbData.properties.filter(p=>p.pm_email?.toLowerCase()===myEmail):[];
  const allMgrProps=isAdmin&&sbData?sbData.properties:mgrProps;
  const outVAs=vas.filter(v=>v.VATrackerStatus==="Out");
  const myQ=isVA?queue.filter(t=>t.va_email?.toLowerCase()===myEmail):queue;
  const myTm=isVA?timers.filter(t=>t.va_email?.toLowerCase()===myEmail):timers;
  const pendingRevs=sbData?.taskReviews.filter(r=>r.status==="Pending"&&(isAdmin||r.pm_email?.toLowerCase()===myEmail))||[];
  const todayMetrics=sbData?.dailyMetrics.find(m=>m.va_email===myEmail&&m.metric_date===today());
  const todayInts=sbData?.interruptions.filter(i=>i.va_email===myEmail&&i.interruption_date===today())||[];

  // AUTH SCREENS
  if(!acct&&!guestToken)return(
    <div style={S.page}>
      <div style={S.hdr}><div><div style={S.hdrT}>{CONFIG.appName}</div><div style={S.hdrS}>NewShire Property Management</div></div></div>
      <div style={{...S.con,textAlign:"center",paddingTop:80}}>
        <div style={S.card}>
          <div style={{fontSize:40,marginBottom:16}}>⏱</div>
          <div style={{fontSize:20,fontWeight:700,color:C.t2,marginBottom:8}}>VA Productivity Tracker V3</div>
          <div style={{color:C.b4,marginBottom:24}}>Sign in with your NewShire Microsoft account.</div>
          <button style={S.btn(C.hdr)} onClick={login}>Sign In with Microsoft</button>
          {authErr&&<div style={{color:C.er,marginTop:12,fontSize:13}}>{authErr}</div>}
        </div>
      </div>
    </div>
  );
  if(loading)return(<div style={S.page}><div style={S.hdr}><div><div style={S.hdrT}>{CONFIG.appName}</div><div style={S.hdrS}>NewShire Property Management</div></div></div><div style={{...S.con,textAlign:"center",paddingTop:80}}><div style={{fontSize:18,color:C.b4}}>Loading...</div></div></div>);
  if(error||!role)return(
    <div style={S.page}>
      <div style={S.hdr}><div><div style={S.hdrT}>{CONFIG.appName}</div><div style={S.hdrS}>NewShire Property Management</div></div></div>
      <div style={{...S.con,textAlign:"center",paddingTop:80}}>
        <div style={{...S.card,borderLeft:`3px solid ${C.er}`}}>
          <div style={{fontSize:40,marginBottom:16}}>🚫</div>
          <div style={{fontSize:18,fontWeight:600,color:C.er,marginBottom:8}}>{error==="access_denied"?"Access Denied":error==="invalid_guest"?"Invalid or expired link":"Error"}</div>
          <div style={{color:C.b4}}>{error==="access_denied"?"Your account is not authorized for this app.":error==="invalid_guest"?"This guest link is no longer active. Contact your NewShire coordinator.":error}</div>
          {myEmail&&<div style={{color:C.b4,fontSize:12,marginTop:12}}>Signed in as: {myEmail}</div>}
        </div>
      </div>
    </div>
  );

  // TABS
  const TABS=[];
  if(isMgr||isGuest)TABS.push({n:"Dashboard",k:"dash"});
  if(isVA||isAdmin)TABS.push({n:`My Day${myQ.length?` (${myQ.length})`:""}`,k:"myday"});
  if(!isGuest)TABS.push({n:`Active${myTm.length?` (${myTm.length})`:""}`,k:"active"});
  if(isMgr||isGuest)TABS.push({n:"Manager View"+(pendingRevs.length?` (${pendingRevs.length})`:isAdmin?" — All":""),k:"manager"});
  TABS.push({n:"History",k:"history"});
  TABS.push({n:"Scorecard",k:"score"});
  if(isAdmin)TABS.push({n:"Admin",k:"admin"});
  const ck=TABS[tab]?.k||TABS[0]?.k;

  return(
    <div style={S.page}>
      <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@500;600&display=swap" rel="stylesheet"/>
      <div style={S.hdr}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:28,height:28,background:C.gold,borderRadius:5,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:C.hdr}}>V</div>
          <div><div style={S.hdrT}>{CONFIG.appName}</div><div style={S.hdrS}>NewShire Property Management</div></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,flex:1,justifyContent:"center",flexWrap:"wrap"}}>
          {myTm.length>0&&<span style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",background:"rgba(255,255,255,.1)",borderRadius:99,fontSize:10,color:"rgba(255,255,255,.7)"}}><span style={{width:5,height:5,borderRadius:"50%",background:C.ok}}/>  <strong style={{color:"#fff"}}>{myTm.length}</strong> timing</span>}
          {covQ.length>0&&<span style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",background:"rgba(255,255,255,.1)",borderRadius:99,fontSize:10,color:"rgba(255,255,255,.7)"}}><span style={{width:5,height:5,borderRadius:"50%",background:C.wn}}/><strong style={{color:"#fff"}}>{covQ.length}</strong> coverage</span>}
          {pendingRevs.length>0&&<span style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",background:"rgba(255,255,255,.1)",borderRadius:99,fontSize:10,color:"rgba(255,255,255,.7)"}}><span style={{width:5,height:5,borderRadius:"50%",background:C.pu}}/><strong style={{color:"#fff"}}>{pendingRevs.length}</strong> reviews</span>}
          {outVAs.length>0&&<span style={{display:"flex",alignItems:"center",gap:4,padding:"3px 9px",background:"rgba(255,255,255,.1)",borderRadius:99,fontSize:10,color:"rgba(255,255,255,.7)"}}><span style={{width:5,height:5,borderRadius:"50%",background:C.er}}/><strong style={{color:"#fff"}}>{outVAs.length}</strong> out</span>}
        </div>
        <div style={{textAlign:"right",flexShrink:0}}>
          <strong style={{display:"block",fontSize:12,fontWeight:600,color:"#fff"}}>{isGuest?guestData?.guest_name||"Guest":(acct?.name||myEmail)}</strong>
          <em style={{fontSize:9,fontStyle:"normal",color:C.gold,textTransform:"uppercase",letterSpacing:".07em"}}>{role}</em>
        </div>
      </div>
      <div style={S.tabs}>{TABS.map((t,i)=><button key={t.k} style={S.tab(tab===i)} onClick={()=>setTab(i)}>{t.n}</button>)}</div>
      {flash&&<div style={{background:C.gl,borderBottom:`1px solid ${C.gold}`,padding:"8px 20px",fontSize:13,fontWeight:600,color:C.g2,textAlign:"center"}}>{flash}</div>}
      <div style={S.con}>
        {ck==="dash"&&<DashboardView sbData={sbData} queue={queue} timers={timers} covQ={covQ} dfFrom={dfFrom} dfTo={dfTo} setDfFrom={setDfFrom} setDfTo={setDfTo} role={role} myEmail={myEmail} guestVAs={guestVAs} employees={employees}/>}
        {ck==="myday"&&<MyDayView sbData={sbData} myEmp={myEmp} myEmail={myEmail} myProps={myProps} queue={myQ} covQ={covQ} timers={myTm} tick={tick} shift={shift} todayMetrics={todayMetrics} todayInts={todayInts} role={role} employees={employees} onClockIn={clockIn} onBreakStart={startBreak} onBreakEnd={endBreak} onClockOut={clockOut} onStart={startTimer} onClaim={claimCov} onAdd={addTask} onRemove={isAdmin?removeTask:null} onLogInt={logInterruption} onNudge={nudgeMetric} onSubmitMetrics={submitMetrics} onReview={submitReview} config={sbData?.config}/>}
        {ck==="active"&&<ActiveView timers={myTm} tick={tick} isMgr={isMgr} isGuest={isGuest} onPause={pauseTimer} onResume={resumeTimer} onFinish={finishTimer} onCancel={cancelTimer}/>}
        {ck==="manager"&&<ManagerView sbData={sbData} myEmail={myEmail} myEmp={myEmp} role={role} allProps={isAdmin?sbData?.properties:allMgrProps} pendingRevs={pendingRevs} timers={timers} queue={queue} employees={employees} isAdmin={isAdmin} isRegional={isRegional} onAdd={addTask} onResolve={resolveReview} pms={pms}/>}
        {ck==="history"&&<HistoryView sbData={sbData} role={role} myEmail={myEmail} mgrProps={allMgrProps} guestVAs={guestVAs}/>}
        {ck==="score"&&<ScorecardView sbData={sbData} role={role} myEmail={myEmail} myVa={myVa} myProps={myProps} isMgr={isMgr} employees={employees} vas={vas}/>}
        {ck==="admin"&&<AdminView sbData={sbData} employees={employees} vas={vas} pms={pms} myEmail={myEmail} acct={acct} queue={queue} covQ={covQ} config={sbData?.config} onToggleAbsence={toggleAbsence} onAdd={addTask} onRemove={removeTask} onCloseDay={closeDay} onUpdateConfig={updateConfig} onAssignProp={assignProp} onUnassignProp={unassignProp} onAddProp={addProperty} onSaveProp={saveProperty} onGT={gT} reload={reload} fl={fl}/>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// DASHBOARD VIEW
// ═══════════════════════════════════════════
function DashboardView({sbData,queue,timers,covQ,dfFrom,dfTo,setDfFrom,setDfTo,role,myEmail,guestVAs,employees}){
  if(!sbData)return null;
  const isGuest=role==="guest";
  const filtVAEmails=guestVAs?new Set(guestVAs):null;
  const tasks=sbData.activities.filter(a=>a.activity_type==="Task"&&inRange(a.task_date,dfFrom,dfTo)&&(!filtVAEmails||filtVAEmails.has(a.va_email)));
  const done=tasks.filter(a=>a.status==="Completed");
  const blocked=tasks.filter(a=>a.status==="Blocked");
  const inc=tasks.filter(a=>a.status==="Incomplete");
  const overdue=tasks.filter(a=>isOD(a));
  const shifts=sbData.activities.filter(a=>a.activity_type==="Shift"&&inRange(a.task_date,dfFrom,dfTo));
  const shiftMin=shifts.reduce((s,a)=>s+(a.work_minutes||0),0);
  const taskMin=done.reduce((s,a)=>s+(a.duration_min||0),0);
  const rate=tasks.length?Math.round(done.length/tasks.length*100):0;
  const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;
  const vas=employees.filter(e=>detectRole(e)==="va"&&e.EmployeeActive!==false).filter(v=>!filtVAEmails||filtVAEmails.has(v.Email?.toLowerCase()));

  return(
    <div>
      {/* Date filter */}
      <div style={{...S.card,display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",padding:12}}>
        <div><label style={S.lbl}>From</label><input type="date" style={{...S.inp,width:150}} value={dfFrom} onChange={e=>setDfFrom(e.target.value)}/></div>
        <div><label style={S.lbl}>To</label><input type="date" style={{...S.inp,width:150}} value={dfTo} onChange={e=>setDfTo(e.target.value)}/></div>
        <div style={{display:"flex",gap:5}}>
          {[{l:"7d",d:7},{l:"14d",d:14},{l:"30d",d:30}].map(p=><button key={p.l} style={{...S.btnO(C.t3,C.tl),...S.sm}} onClick={()=>{setDfFrom(daysAgoStr(p.d));setDfTo(today());}}>{p.l}</button>)}
        </div>
      </div>
      {/* KPIs */}
      <div style={{...S.row,marginBottom:14}}>
        <KPI label="Tasks" value={tasks.length}/>
        <KPI label="Done" value={done.length} color={C.ok}/>
        <KPI label="Rate" value={`${rate}%`} color={rate>=85?C.ok:rate>=60?C.wn:tasks.length?C.er:C.t2}/>
        {!isGuest&&<KPI label="Shift Hrs" value={fM(shiftMin)}/>}
        {!isGuest&&<KPI label="Util" value={`${util}%`} color={util>=75?C.ok:util>=50?C.wn:shiftMin>0?C.er:C.t2}/>}
        <KPI label="Blocked" value={blocked.length} color={blocked.length>0?C.er:C.ok}/>
        <KPI label="Overdue" value={overdue.length} color={overdue.length>0?C.er:C.ok}/>
        {!isGuest&&<KPI label="Coverage" value={covQ.length} color={covQ.length>0?C.wn:C.ok}/>}
      </div>
      {/* Needs Attention */}
      {(blocked.length>0||overdue.length>0||covQ.length>0)&&(
        <div style={{...S.card,...S.ae,background:C.erb,marginBottom:14}}>
          <div style={{fontSize:13,fontWeight:700,color:C.er,marginBottom:10}}>🚨 Needs Attention</div>
          {blocked.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:11,fontWeight:700,color:C.er,marginBottom:4}}>Blocked ({blocked.length})</div>{blocked.slice(0,5).map((t,i)=><div key={i} style={{fontSize:12,color:C.b6,padding:"2px 0"}}>{t.va_name}: {t.title} — {t.property_name}{t.notes?` · ${t.notes}`:""}</div>)}</div>}
          {overdue.length>0&&<div style={{marginBottom:8}}><div style={{fontSize:11,fontWeight:700,color:C.er,marginBottom:4}}>Overdue ({overdue.length})</div>{overdue.slice(0,5).map((t,i)=><div key={i} style={{fontSize:12,color:C.b6,padding:"2px 0"}}>{t.va_name}: {t.title} — <Badge type="er" nd>{`OD ${fD(t.task_date)}`}</Badge></div>)}</div>}
          {covQ.length>0&&<div><div style={{fontSize:11,fontWeight:700,color:C.wn,marginBottom:4}}>Unclaimed Coverage ({covQ.length})</div>{covQ.slice(0,3).map((t,i)=><div key={i} style={{fontSize:12,color:C.b6,padding:"2px 0"}}>{t.title} — covering {t.coverage_for_name}</div>)}</div>}
        </div>
      )}
      {/* Kanban */}
      <div style={S.card}>
        <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:12}}>📋 Task Board</div>
        <div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:8}}>
          {vas.map(va=>{
            const isOut=va.VATrackerStatus==="Out";
            const vQ=queue.filter(t=>t.va_email?.toLowerCase()===va.Email?.toLowerCase());
            const vT=timers.filter(t=>t.va_email?.toLowerCase()===va.Email?.toLowerCase());
            const vOD=vQ.filter(t=>isOD(t));
            const port=sbData.portfolios.filter(p=>p.va_email?.toLowerCase()===va.Email?.toLowerCase());
            return(
              <div key={va.Email} style={{flex:"0 0 190px",background:isOut?C.erb:C.bg,borderRadius:8,padding:11,border:`1px solid ${isOut?"rgba(184,59,42,0.2)":C.b1}`}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:9,paddingBottom:8,borderBottom:`2px solid ${C.tl}`}}>
                  <div><div style={{fontSize:11,fontWeight:700,color:isOut?C.er:C.t2}}>{va.Name}{isOut?" 🛑":""}</div><div style={{fontSize:9,color:C.b4}}>{port.length} props</div></div>
                  {isOut?<Badge type="er">OUT</Badge>:vT.length>0?<Badge type="ok">Active</Badge>:vOD.length>0?<Badge type="er" nd>{vOD.length} OD</Badge>:<span style={{fontSize:9,color:C.b4}}>Offline</span>}
                </div>
                {vT.map(t=><div key={t.id} style={{background:C.w,border:`1px solid ${C.ok}`,borderLeft:`3px solid ${C.ok}`,borderRadius:6,padding:8,marginBottom:5,background:C.okb}}><div style={{fontSize:8,fontWeight:700,color:C.ok,textTransform:"uppercase",marginBottom:2}}>● Active</div><div style={{fontSize:11,fontWeight:600,color:C.t2}}>{t.title}</div><div style={{fontSize:9,color:C.b4}}>{t.property_name}</div></div>)}
                {vQ.filter(t=>isOD(t)).map(t=><div key={t.id} style={{background:C.erb,border:`1px solid rgba(184,59,42,.2)`,borderLeft:`3px solid ${C.er}`,borderRadius:6,padding:8,marginBottom:5}}><div style={{fontSize:8,fontWeight:700,color:C.er,marginBottom:2}}>⚠ OD {fD(t.task_date)}</div><div style={{fontSize:11,fontWeight:600,color:C.t2}}>{t.title}</div><div style={{fontSize:9,color:C.b4}}>{t.property_name}</div></div>)}
                {vQ.filter(t=>!isOD(t)).slice(0,3).map(t=><div key={t.id} style={{background:C.w,border:`1px solid ${C.b1}`,borderLeft:`3px solid ${t.priority==="Urgent"?C.er:t.priority==="High"?C.wn:C.b2}`,borderRadius:6,padding:8,marginBottom:5}}><div style={{fontSize:11,fontWeight:600,color:C.t2}}>{t.title}</div><div style={{fontSize:9,color:C.b4}}>{t.property_name}</div></div>)}
                {!vT.length&&!vQ.length&&<div style={{textAlign:"center",padding:"16px 0",color:C.b4,fontSize:10,fontStyle:"italic"}}>{isOut?"Tasks in coverage":"No pending"}</div>}
              </div>
            );
          })}
        </div>
      </div>
      {/* Performance Table */}
      <div style={S.card}>
        <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:12}}>VA Performance</div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["VA","Props","Done","Rate","Shift","Util","Blocked","Overdue","Interruptions"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {vas.map(va=>{
                const vT=tasks.filter(a=>a.va_email===va.Email);const vD=vT.filter(a=>a.status==="Completed");const vB=vT.filter(a=>a.status==="Blocked");const vOD=vT.filter(a=>isOD(a));
                const vS=shifts.filter(a=>a.va_email===va.Email);const vSm=vS.reduce((s,a)=>s+(a.work_minutes||0),0);const vTm=vD.reduce((s,a)=>s+(a.duration_min||0),0);
                const vR=vT.length?Math.round(vD.length/vT.length*100):0;const vU=vSm>0?Math.round(vTm/vSm*100):0;
                const vInts=sbData.interruptions.filter(i=>i.va_email===va.Email&&inRange(i.interruption_date,dfFrom,dfTo)).length;
                const port=sbData.portfolios.filter(p=>p.va_email?.toLowerCase()===va.Email?.toLowerCase());
                return(<tr key={va.Email} style={{opacity:va.VATrackerStatus==="Out"?.6:1}}>
                  <td style={{...S.td,fontWeight:700,color:va.VATrackerStatus==="Out"?C.er:C.t2}}>{va.Name}</td>
                  <td style={{...S.td,fontSize:11}}>{port.length}</td>
                  <td style={S.td}>{vD.length}/{vT.length}</td>
                  <td style={S.td}><Badge type={vR>=85?"ok":vR>=60?"wn":"er"} nd>{vR}%</Badge></td>
                  <td style={{...S.td,fontFamily:M,fontSize:11}}>{fM(vSm)}</td>
                  <td style={S.td}><div style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:36,height:5,background:C.b1,borderRadius:3,overflow:"hidden"}}><div style={{width:`${Math.min(vU,100)}%`,height:"100%",background:vU>=75?C.ok:vU>=50?C.wn:C.er,borderRadius:3}}/></div><span style={{fontFamily:M,fontSize:10,fontWeight:700,color:vU>=75?C.ok:vU>=50?C.wn:C.er}}>{vU}%</span></div></td>
                  <td style={S.td}>{vB.length>0?<Badge type="er" nd>{vB.length}</Badge>:"0"}</td>
                  <td style={S.td}>{vOD.length>0?<Badge type="er" nd>{vOD.length}</Badge>:"0"}</td>
                  <td style={S.td}>{vInts}</td>
                </tr>);
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// OVERDUE SECTION
// ═══════════════════════════════════════════
function OverdueSection({tasks,onStart,onRemove,onReview,myEmail,myEmp,sbData,pms}){
  const od=tasks.filter(isOD);
  if(!od.length)return null;
  return(
    <div style={{background:C.erb,border:`1px solid rgba(184,59,42,.18)`,borderLeft:`4px solid ${C.er}`,borderRadius:8,padding:13,marginBottom:12}}>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
        <div style={{width:22,height:22,background:C.er,borderRadius:4,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#fff",fontWeight:700}}>!</div>
        <div><div style={{fontSize:12,fontWeight:700,color:C.er}}>{od.length} Overdue Task{od.length>1?"s":""} from Previous Days</div><div style={{fontSize:10,color:"rgba(184,59,42,.7)",marginTop:1}}>Start, mark incomplete, or flag for manager review</div></div>
      </div>
      {od.map(t=><TaskRowFull key={t.id} task={t} onStart={onStart} onRemove={onRemove} onReview={onReview} myEmail={myEmail} myEmp={myEmp} sbData={sbData} pms={pms} showOD/>)}
    </div>
  );
}

// ═══════════════════════════════════════════
// TASK ROW (full — with review button)
// ═══════════════════════════════════════════
function TaskRowFull({task,onStart,onRemove,onReview,myEmail,myEmp,sbData,pms,showOD}){
  const[showRevForm,setShowRevForm]=useState(false);
  const[revType,setRevType]=useState("Application Review");
  const[revNote,setRevNote]=useState("");
  const prop=sbData?.properties.find(p=>p.property_code===task.property_code);
  const pmEmail=prop?.pm_email||"";
  const pm=pms?.find(p=>p.Email?.toLowerCase()===pmEmail?.toLowerCase());
  const revCount=sbData?.taskReviews.filter(r=>r.activity_id===task.id).length||0;

  async function doReview(){
    if(!revNote||!pmEmail){alert("Add a note and make sure this property has a PM assigned.");return;}
    await onReview(task.id,{va_email:myEmail,va_name:myEmp?.Name||myEmail,pm_email:pmEmail,pm_name:pm?.Name||"PM",review_type:revType,property_code:task.property_code,property_name:task.property_name,notes:revNote});
    setShowRevForm(false);setRevNote("");
  }

  return(
    <div style={{borderBottom:`1px solid ${C.b1}`,paddingBottom:8,marginBottom:8}}>
      <div style={{display:"flex",alignItems:"flex-start",gap:9}}>
        <span style={{fontSize:15,paddingTop:1,flexShrink:0}}>{catI[task.category]||"📁"}</span>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:12,fontWeight:600,color:C.t2,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
            {task.title}
            {showOD&&<Badge type="er" nd>{`OD — ${fD(task.task_date)}`}</Badge>}
            {task.coverage_for_name&&<Badge type="wn" nd>Coverage</Badge>}
            {task.priority!=="Normal"&&<Badge type={task.priority==="Urgent"?"er":"wn"} nd>{task.priority}</Badge>}
          </div>
          <div style={{fontSize:10,color:C.b4,marginTop:3,display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
            {task.property_name}<span style={{width:3,height:3,borderRadius:"50%",background:C.b4}}/>
            <Badge type={srcB[task.source]||"ne"} nd>{task.source}</Badge>
            {revCount>0&&<span style={{color:C.pu,fontWeight:700}}>⚑ {revCount} review{revCount>1?"s":""}</span>}
          </div>
        </div>
        <div style={{display:"flex",gap:4,flexShrink:0,flexWrap:"wrap"}}>
          {onStart&&<button style={{...S.btn(C.ok),...S.xs}} onClick={()=>onStart(task)}>▶ Start</button>}
          {onReview&&<button style={{...S.btn(C.pu),...S.xs}} onClick={()=>setShowRevForm(v=>!v)}>⚑ Review</button>}
          {onRemove&&<button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={()=>{if(window.confirm("Remove this task?"))onRemove(task);}}>✕</button>}
        </div>
      </div>
      {showRevForm&&(
        <div style={{background:C.pub,border:`1px solid rgba(91,63,168,.2)`,borderRadius:6,padding:12,marginTop:8}}>
          <div style={{fontSize:11,fontWeight:700,color:C.pu,marginBottom:8}}>⚑ Request Manager Review — {task.title}</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:6,marginBottom:9}}>
            {["Application Review","Prospect Question","Resident Question","Other"].map(t=>(
              <button key={t} style={{padding:"9px 8px",fontSize:11,fontWeight:600,border:`1.5px solid ${revType===t?C.pu:C.b2}`,borderRadius:6,background:revType===t?C.pub:C.w,color:revType===t?C.pu:C.b4,cursor:"pointer",textAlign:"left",display:"flex",alignItems:"center",gap:7}} onClick={()=>setRevType(t)}>
                {t==="Application Review"?"📋":t==="Prospect Question"?"🏠":t==="Resident Question"?"💬":"📁"} {t}
              </button>
            ))}
          </div>
          <div style={{marginBottom:9}}><label style={S.lbl}>Notify PM</label><div style={{fontSize:12,color:C.b6,padding:"8px 10px",background:C.w,border:`1px solid ${C.b2}`,borderRadius:6}}>{pm?.Name||"PM"} {pmEmail?`(${pmEmail})`:""}{!pmEmail&&<span style={{color:C.er}}> — No PM assigned to this property</span>}</div></div>
          <div style={{marginBottom:9}}><label style={S.lbl}>What needs their attention? *</label><textarea style={S.inp} rows={3} value={revNote} onChange={e=>setRevNote(e.target.value)} placeholder="Describe what you need the manager to review or decide..."/></div>
          {revNote&&pmEmail&&<div style={{background:C.tmb,border:`1px solid rgba(74,78,138,.2)`,borderRadius:8,padding:10,marginBottom:9}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
              <div style={{width:18,height:18,background:C.tm,borderRadius:3,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:700,color:"#fff"}}>T</div>
              <span style={{fontSize:10,fontWeight:700,color:C.tm}}>Teams message preview — sent immediately</span>
            </div>
            <div style={{background:C.w,border:`1px solid rgba(74,78,138,.15)`,borderRadius:6,padding:"8px 10px",fontSize:11,color:C.b6,lineHeight:1.5}}>
              <strong>🔔 Manager Review Requested</strong><br/>
              <strong>From:</strong> {myEmp?.Name||myEmail} | <strong>Type:</strong> {revType}<br/>
              <strong>Property:</strong> {task.property_name}<br/><br/>
              <em style={{color:C.b4}}>{revNote}</em>
            </div>
          </div>}
          <div style={{display:"flex",gap:7}}>
            <button style={{...S.btn(C.tm),flex:1}} onClick={doReview}>Send Teams Message &amp; Create Review</button>
            <button style={{...S.btnO(C.t2,C.b2),...S.sm}} onClick={()=>setShowRevForm(false)}>Cancel</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════
// MY DAY VIEW
// ═══════════════════════════════════════════
function MyDayView({sbData,myEmp,myEmail,myProps,queue,covQ,timers,tick,shift,todayMetrics,todayInts,role,employees,onClockIn,onBreakStart,onBreakEnd,onClockOut,onStart,onClaim,onAdd,onRemove,onLogInt,onNudge,onSubmitMetrics,onReview,config}){
  const[showForm,setShowForm]=useState(false);const[fCat,setFCat]=useState("");const[fProp,setFProp]=useState("");const[fPri,setFPri]=useState("Normal");const[fDesc,setFDesc]=useState("");
  const[intType,setIntType]=useState("Prospect Call");const[intProp,setIntProp]=useState("");const[intDur,setIntDur]=useState("5");const[intNote,setIntNote]=useState("");const[cvt,setCvt]=useState(false);const[cvtDesc,setCvtDesc]=useState("");const[cvtPri,setCvtPri]=useState("Normal");
  const[metNotes,setMetNotes]=useState(todayMetrics?.notes||"");
  const isAdm=role==="admin";
  const cats=config?.categories||[];
  const portProps=sbData?sbData.portfolios.filter(p=>p.va_email?.toLowerCase()===myEmail).map(p=>sbData.properties.find(pr=>pr.property_code===p.property_code)).filter(Boolean):[];
  const daily=queue.filter(t=>t.source==="Daily"&&!isOD(t));
  const assigned=queue.filter(t=>(t.source==="Assigned"||t.source==="Coverage")&&!isOD(t));
  const adhoc=queue.filter(t=>t.source==="Ad-Hoc"&&!isOD(t));
  const pms=employees.filter(e=>["manager","regional","admin"].includes(detectRole(e))&&e.EmployeeActive!==false);

  let shElapsed=0;
  if(shift){const now=Date.now();let bMs=shift.Breaks.reduce((s,b)=>s+(new Date(b.e)-new Date(b.s)),0);if(shift._ob&&shift._bs)bMs+=(now-new Date(shift._bs).getTime());shElapsed=Math.floor((now-new Date(shift.ClockIn).getTime()-bMs)/1000);}

  function handleAdd(){
    if(!fDesc||!fCat)return;
    const cat=cats.find(c=>c.id===fCat);const prop=fProp?portProps.find(p=>p.property_code===fProp):null;
    onAdd({title:fDesc,va_email:myEmail,va_name:myEmp?.Name||myEmail,property_code:fProp||"",property_name:prop?.property_name||"General",pm_name:prop?.pm_name||"",category:cat?.name||"Admin/Other",priority:fPri,source:"Ad-Hoc"});
    setFDesc("");setFCat("");setFProp("");setFPri("Normal");setShowForm(false);
  }

  function handleLogInt(){
    if(!intType)return;
    const prop=intProp?portProps.find(p=>p.property_code===intProp):null;
    const cat=cats.find(c=>c.id===fCat);
    const intFields={va_email:myEmail,va_name:myEmp?.Name||myEmail,interruption_date:today(),interruption_type:intType,property_code:intProp||"",property_name:prop?.property_name||"General",duration_min:parseInt(intDur)||0,notes:intNote};
    const taskFields=cvt&&cvtDesc?{title:cvtDesc,va_email:myEmail,va_name:myEmp?.Name||myEmail,property_code:intProp||"",property_name:prop?.property_name||"General",pm_name:prop?.pm_name||"",category:"Tenant Comms",priority:cvtPri}:null;
    onLogInt(intFields,taskFields);
    setIntNote("");setCvt(false);setCvtDesc("");setIntDur("5");
  }

  const metFields=[
    {k:"leads_followed_up",l:"Leads"},
    {k:"applications_processed",l:"Apps"},
    {k:"showings_scheduled",l:"Showings"},
    {k:"resident_comms",l:"Res. Comms"},
    {k:"work_orders_entered",l:"WOs In"},
    {k:"work_orders_updated",l:"WOs Upd."},
    {k:"renewals_touched",l:"Renewals"},
    {k:"manager_calls",l:"Mgr Calls"},
  ];

  return(
    <div style={{maxWidth:720,margin:"0 auto"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
        <div><div style={{fontSize:18,fontWeight:700,color:C.t2}}>{myEmp?.Name||myEmail}</div><div style={{fontSize:11,color:C.b4}}>{portProps.length} properties &nbsp;·&nbsp; {portProps.reduce((s,p)=>s+(p.units||0),0)} units</div></div>
        {timers.length>0&&<Badge type="ok">{timers.length} timing</Badge>}
      </div>

      {/* SHIFT CLOCK */}
      <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:8,padding:14,marginBottom:12}}>
        {!shift?(
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div><div style={{fontSize:11,fontWeight:700,color:C.inf,textTransform:"uppercase",letterSpacing:".05em"}}>Not Clocked In</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Tap to start your shift</div></div>
            <button style={S.btn(C.inf)} onClick={onClockIn}>☀ Clock In</button>
          </div>
        ):(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:11}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:8,height:8,borderRadius:"50%",background:shift._ob?C.wn:C.ok}}/><div style={{fontSize:11,fontWeight:700,color:shift._ob?C.wn:C.ok,textTransform:"uppercase",letterSpacing:".05em"}}>{shift._ob?"On Break":"Clocked In"}</div></div>
                <div style={{fontSize:10,color:C.b4,marginTop:1}}>Since {fT(shift.ClockIn)} · {shift.Breaks.length+(shift._ob?1:0)} breaks</div>
              </div>
              <div style={{textAlign:"right"}}><div style={{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase"}}>Working Time</div><div style={{fontSize:24,fontWeight:700,fontFamily:M,color:shift._ob?C.wn:C.t2}}>{fTm(shElapsed)}</div></div>
            </div>
            <div style={{display:"flex",gap:8}}>
              {shift._ob?<button style={{...S.btn(C.ok),flex:1}} onClick={onBreakEnd}>▶ End Break</button>:<button style={{...S.btnO(C.wn,C.wn),flex:1}} onClick={onBreakStart}>☕ Break</button>}
              <button style={{...S.btn(C.er),flex:1}} onClick={onClockOut}>🌙 Clock Out</button>
            </div>
          </div>
        )}
      </div>

      {/* COVERAGE */}
      {covQ.length>0&&(
        <div style={{...S.card,borderLeft:`4px solid ${C.wn}`,background:C.wnb}}>
          <div style={{fontSize:13,fontWeight:700,color:C.wn,marginBottom:8}}>🚨 Coverage Needed ({covQ.length})</div>
          {covQ.map(t=>(
            <div key={t.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid rgba(168,111,8,.1)`}}>
              <span style={{fontSize:14}}>{catI[t.category]||"📁"}</span>
              <div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{t.title}</div><div style={{fontSize:10,color:C.b4}}>{t.property_name} · covering {t.coverage_for_name}</div></div>
              <button style={{...S.btn(C.wn,"#1a2a30"),...S.sm}} onClick={()=>onClaim(t.id)}>✋ Claim</button>
            </div>
          ))}
        </div>
      )}

      {/* OVERDUE */}
      <OverdueSection tasks={queue} onStart={onStart} onRemove={isAdm?onRemove:null} onReview={onReview} myEmail={myEmail} myEmp={myEmp} sbData={sbData} pms={pms}/>

      {/* DAILY */}
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:9,borderBottom:`1px solid ${C.b1}`,marginBottom:9}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2}}>📋 Daily Tasks</div>
          <span style={{fontSize:11,color:C.b4}}>{daily.length} remaining</span>
        </div>
        {!daily.length&&<p style={{color:C.b4,fontSize:12}}>All done! ✓</p>}
        {daily.map(t=><TaskRowFull key={t.id} task={t} onStart={onStart} onRemove={isAdm?onRemove:null} onReview={onReview} myEmail={myEmail} myEmp={myEmp} sbData={sbData} pms={pms}/>)}
      </div>

      {/* ASSIGNED */}
      {assigned.length>0&&(
        <div style={{...S.card,borderLeft:`4px solid ${C.gold}`}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:9}}>📌 Assigned &amp; Coverage ({assigned.length})</div>
          {assigned.map(t=><TaskRowFull key={t.id} task={t} onStart={onStart} onRemove={isAdm?onRemove:null} onReview={onReview} myEmail={myEmail} myEmp={myEmp} sbData={sbData} pms={pms}/>)}
        </div>
      )}

      {/* AD-HOC */}
      {adhoc.length>0&&(
        <div style={S.card}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:9}}>📝 Extra ({adhoc.length})</div>
          {adhoc.map(t=><TaskRowFull key={t.id} task={t} onStart={onStart} onRemove={isAdm?onRemove:null} onReview={onReview} myEmail={myEmail} myEmp={myEmp} sbData={sbData} pms={pms}/>)}
        </div>
      )}

      {/* INTERRUPTION LOGGER */}
      <div style={{...S.card,...S.ac}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:11}}>
          <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>📞 Log Interruption</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{todayInts.length} logged today · {todayInts.reduce((s,i)=>s+(i.duration_min||0),0)} min total</div></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:6,marginBottom:10}}>
          {["Prospect Call","Resident Call","Vendor Call","Owner Call","Manager Call","Other"].map(t=>(
            <button key={t} style={{padding:"9px 4px",fontSize:10,fontWeight:600,border:`1.5px solid ${intType===t?C.t3:C.b2}`,borderRadius:6,background:intType===t?C.tl0:C.w,color:intType===t?C.t2:C.b4,cursor:"pointer",textAlign:"center",lineHeight:1.35}} onClick={()=>setIntType(t)}>{t}</button>
          ))}
        </div>
        <div style={{...S.row,marginBottom:9}}>
          <div style={{flex:2,minWidth:120}}><label style={S.lbl}>Property</label><select style={S.sel} value={intProp} onChange={e=>setIntProp(e.target.value)}><option value="">General</option>{portProps.map(p=><option key={p.property_code} value={p.property_code}>{p.property_name}</option>)}</select></div>
          <div style={{flex:1,minWidth:75}}><label style={S.lbl}>Duration (min)</label><input style={S.inp} type="number" value={intDur} onChange={e=>setIntDur(e.target.value)} min="1"/></div>
        </div>
        <input style={{...S.inp,marginBottom:10}} type="text" value={intNote} onChange={e=>setIntNote(e.target.value)} placeholder="Notes (optional)"/>
        {/* Convert to task toggle */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 12px",background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,marginBottom:cvt?0:10,cursor:"pointer"}} onClick={()=>setCvt(v=>!v)}>
          <div><div style={{fontSize:12,fontWeight:700,color:C.t2}}>➕ Convert to Task</div><div style={{fontSize:10,color:C.b4,marginTop:1}}>Add a follow-up task to your queue</div></div>
          <button style={{width:40,height:22,background:cvt?C.ok:C.b2,borderRadius:11,position:"relative",cursor:"pointer",border:"none",transition:"background .2s"}} onClick={e=>{e.stopPropagation();setCvt(v=>!v);}}>
            <span style={{position:"absolute",width:16,height:16,background:"#fff",borderRadius:"50%",top:3,left:cvt?21:3,transition:"left .2s",boxShadow:"0 1px 2px rgba(0,0,0,.15)"}}/>
          </button>
        </div>
        {cvt&&(
          <div style={{background:C.tl0,border:`1px solid ${C.tl}`,borderRadius:6,padding:12,marginBottom:10}}>
            <div style={{fontSize:11,fontWeight:700,color:C.t3,marginBottom:8}}>📋 Task from this interruption</div>
            <div style={{...S.row,marginBottom:8}}>
              <div style={{flex:2}}><label style={S.lbl}>Category</label><select style={S.sel} value={fCat} onChange={e=>setFCat(e.target.value)}><option value="">Select...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
              <div style={{flex:1}}><label style={S.lbl}>Priority</label><select style={S.sel} value={cvtPri} onChange={e=>setCvtPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
            </div>
            <input style={S.inp} type="text" value={cvtDesc} onChange={e=>setCvtDesc(e.target.value)} placeholder="Task description *"/>
          </div>
        )}
        <button style={{...S.btn(C.t2),width:"100%"}} onClick={handleLogInt}>{cvt&&cvtDesc?"Log Interruption & Add Task to Queue":"Log Interruption"}</button>
      </div>

      {/* DAILY METRICS */}
      <div style={{...S.card,...S.at}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:11}}>
          <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>📈 Daily Activity Metrics</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>Tap +/− to track as you go</div></div>
          {todayMetrics?.submitted?<Badge type="ok" nd>Submitted</Badge>:<Badge type="wn" nd>Not submitted</Badge>}
        </div>
        {!todayMetrics?.submitted&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7,marginBottom:10}}>
              {metFields.map(({k,l})=>{
                const val=todayMetrics?todayMetrics[k]||0:0;
                return(
                  <div key={k} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:8,padding:"9px 5px",textAlign:"center"}}>
                    <div style={{fontSize:8.5,fontWeight:700,color:C.t3,textTransform:"uppercase",letterSpacing:".05em",marginBottom:3}}>{l}</div>
                    <div style={{fontSize:24,fontWeight:700,fontFamily:M,color:C.t2,lineHeight:1}}>{val}</div>
                    <div style={{display:"flex",justifyContent:"center",gap:4,marginTop:6}}>
                      <button style={{width:25,height:25,border:`1px solid ${C.b2}`,borderRadius:4,background:C.w,color:C.t2,fontSize:14,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>onNudge(k,-1)}>−</button>
                      <button style={{width:25,height:25,border:`1px solid ${C.b2}`,borderRadius:4,background:C.w,color:C.t2,fontSize:14,fontWeight:700,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}} onClick={()=>onNudge(k,1)}>+</button>
                    </div>
                  </div>
                );
              })}
            </div>
            <input style={{...S.inp,marginBottom:9}} type="text" value={metNotes} onChange={e=>setMetNotes(e.target.value)} placeholder="End-of-day notes (optional)"/>
            <button style={{...S.btn(C.gold,C.t2),width:"100%"}} onClick={()=>onSubmitMetrics(metNotes)}>✓ Submit Day</button>
          </>
        )}
        {todayMetrics?.submitted&&(
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7}}>
            {metFields.map(({k,l})=>(
              <div key={k} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"8px 5px",textAlign:"center",opacity:.8}}>
                <div style={{fontSize:8,fontWeight:700,color:C.t3,textTransform:"uppercase",letterSpacing:".05em",marginBottom:2}}>{l}</div>
                <div style={{fontSize:20,fontWeight:700,fontFamily:M,color:C.t2}}>{todayMetrics[k]||0}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ADD TASK */}
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2}}>➕ Add Task</div>
          <button style={S.btnO(C.t2,C.b2)} onClick={()=>setShowForm(v=>!v)}>{showForm?"Cancel":"New"}</button>
        </div>
        {showForm&&(
          <div style={{marginTop:12}}>
            <div style={{...S.row,marginBottom:9}}>
              <div style={{flex:1,minWidth:120}}><label style={S.lbl}>Category *</label><select style={S.sel} value={fCat} onChange={e=>setFCat(e.target.value)}><option value="">...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
              <div style={{flex:1,minWidth:120}}><label style={S.lbl}>Property</label><select style={S.sel} value={fProp} onChange={e=>setFProp(e.target.value)}><option value="">General</option>{portProps.map(p=><option key={p.property_code} value={p.property_code}>{p.property_name}</option>)}</select></div>
              <div style={{minWidth:80}}><label style={S.lbl}>Priority</label><select style={S.sel} value={fPri} onChange={e=>setFPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
            </div>
            <input style={{...S.inp,marginBottom:9}} type="text" value={fDesc} onChange={e=>setFDesc(e.target.value)} placeholder="Task description *"/>
            <button style={{...S.btn(C.hdr),width:"100%"}} onClick={handleAdd}>+ Add to Queue</button>
          </div>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// ACTIVE VIEW
// ═══════════════════════════════════════════
function ActiveView({timers,tick,isMgr,isGuest,onPause,onResume,onFinish,onCancel}){
  if(!timers.length)return(<div style={{...S.card,textAlign:"center",padding:50}}><div style={{fontSize:36,marginBottom:10}}>⏱</div><div style={{fontSize:17,fontWeight:700,color:C.t2}}>No Active Timers</div><div style={{fontSize:13,color:C.b4}}>Start a task to begin tracking.</div></div>);
  return(
    <div>
      {timers.map(t=>{
        const now=Date.now(),st=new Date(t.start_time).getTime();
        let pMs=t._pMs||0;const ip=!!t._pS;if(ip)pMs+=(now-t._pS);
        const el=ip?Math.floor((t._pS-st-(t._pMs||0))/1000):Math.floor((now-st-pMs)/1000);
        return(
          <div key={t.id} style={{...S.card,padding:0,overflow:"hidden",borderLeft:`4px solid ${ip?C.wn:C.ok}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 15px",background:ip?"rgba(168,111,8,.07)":C.okb}}>
              <div><div style={{fontSize:10,fontWeight:700,color:ip?C.wn:C.ok,textTransform:"uppercase",letterSpacing:".05em"}}>{ip?"⏸ Paused":"● Recording"}</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{fT(t.start_time)} · <Badge type={srcB[t.source]||"ne"} nd>{t.source}</Badge>{isMgr&&` · ${t.va_name}`}</div></div>
              <div style={{fontSize:26,fontWeight:700,fontFamily:M,color:ip?C.wn:C.t2}}>{fTm(el)}</div>
            </div>
            <div style={{padding:"12px 15px"}}><div style={{fontSize:14,fontWeight:700,color:C.t2,marginBottom:3}}>{t.title}</div><div style={{fontSize:11,color:C.b4}}>{catI[t.category]||""} {t.category} · {t.property_name}{t.coverage_for_name?` · 🔄 covering ${t.coverage_for_name}`:""}</div></div>
            {!isGuest&&(
              <div style={{display:"flex",gap:6,padding:"0 15px 13px",flexWrap:"wrap"}}>
                {ip?<button style={{...S.btn(C.ok),flex:1}} onClick={()=>onResume(t.id)}>▶ Resume</button>:<button style={{...S.btnO(C.wn,C.wn),flex:1}} onClick={()=>onPause(t.id)}>⏸ Pause</button>}
                <button style={{...S.btn(C.ok),flex:1}} onClick={()=>{const n=prompt("Notes (optional):");onFinish(t.id,"Completed",n||"");}}>✓ Done</button>
                <button style={{...S.btnO(C.er,C.er),padding:"8px 12px"}} onClick={()=>{const n=prompt("What's blocking?");if(n)onFinish(t.id,"Blocked",n);}}>⚠</button>
                <button style={{...S.btnO(C.b4,C.b2),padding:"8px 12px"}} onClick={()=>onCancel(t.id)}>↩</button>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════
// MANAGER VIEW
// ═══════════════════════════════════════════
function ManagerView({sbData,myEmail,myEmp,role,allProps,pendingRevs,timers,queue,employees,isAdmin,isRegional,onAdd,onResolve,pms}){
  const[selProp,setSelProp]=useState("");const[tCat,setTCat]=useState("");const[tDesc,setTDesc]=useState("");const[tPri,setTPri]=useState("Normal");const[tNotes,setTNotes]=useState("");
  const[revReply,setRevReply]=useState({});
  const cats=sbData?.config?.categories||[];
  const vas=employees.filter(e=>detectRole(e)==="va"&&e.EmployeeActive!==false);

  function getVAForProp(propCode){
    const port=sbData?.portfolios.find(p=>p.property_code===propCode&&p.is_active);
    if(!port)return null;
    return employees.find(e=>e.Email?.toLowerCase()===port.va_email?.toLowerCase())||null;
  }

  function handleSubmit(){
    if(!selProp||!tCat||!tDesc)return;
    const prop=sbData?.properties.find(p=>p.property_code===selProp);
    const va=getVAForProp(selProp);
    if(!va){alert("No VA assigned to this property.");return;}
    const cat=cats.find(c=>c.id===tCat);
    onAdd({title:tDesc,va_email:va.Email,va_name:va.Name,property_code:selProp,property_name:prop?.property_name||"",pm_name:myEmp?.Name||myEmail,category:cat?.name||"Admin/Other",priority:tPri,source:"Assigned",assigned_by_email:myEmail,assigned_by_name:myEmp?.Name||myEmail,notes:tNotes});
    setTDesc("");setTCat("");setTPri("Normal");setTNotes("");
  }

  // Regional Portfolio Overview — PM breakdown
  const pmGroups=isRegional||isAdmin?pms.map(pm=>{
    const pmProps=sbData?.properties.filter(p=>p.pm_email?.toLowerCase()===pm.Email?.toLowerCase())||[];
    const pmVAs=[...new Set(pmProps.map(p=>{const port=sbData?.portfolios.find(pt=>pt.property_code===p.property_code&&pt.is_active);return port?.va_email;}).filter(Boolean))].map(e=>employees.find(v=>v.Email?.toLowerCase()===e.toLowerCase())).filter(Boolean);
    const pmTasks=sbData?.activities.filter(a=>a.activity_type==="Task"&&pmProps.some(p=>p.property_code===a.property_code)&&dAgo(a.task_date)<=7)||[];
    const pmDone=pmTasks.filter(t=>t.status==="Completed").length;
    const pmBlocked=pmTasks.filter(t=>t.status==="Blocked").length;
    const pmOD=pmTasks.filter(t=>isOD(t)).length;
    const pmRevs=pendingRevs.filter(r=>pmProps.some(p=>p.property_code===r.property_code)).length;
    return{pm,props:pmProps,vas:pmVAs,done:pmDone,blocked:pmBlocked,od:pmOD,revs:pmRevs,total:pmTasks.length};
  }).filter(g=>g.props.length>0):[];

  return(
    <div>
      {/* Pending Reviews inbox */}
      {pendingRevs.length>0&&(
        <div style={{...S.card,padding:0,overflow:"hidden",borderTop:`3px solid ${C.er}`}}>
          <div style={{padding:"13px 14px 11px",borderBottom:`1px solid ${C.b1}`}}>
            <div style={{fontSize:13,fontWeight:700,color:C.pu}}>⚑ Needs Your Review — {pendingRevs.length} item{pendingRevs.length>1?"s":""}</div>
            <div style={{fontSize:10,color:C.b4,marginTop:2}}>VAs flagged these · Teams message already sent</div>
          </div>
          {pendingRevs.map(rev=>(
            <div key={rev.id} style={{padding:"12px 14px",borderBottom:`1px solid ${C.b1}`}}>
              <div style={{display:"flex",alignItems:"flex-start",gap:9,marginBottom:8}}>
                <div style={{flex:1}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                    <span style={{fontSize:12,fontWeight:700,color:C.t2}}>{rev.va_name}</span>
                    <span style={{fontSize:10,color:C.b4}}>·</span>
                    <span style={{fontSize:12,color:C.b6}}>{rev.property_name||"General"}</span>
                    <span style={{display:"inline-flex",padding:"1px 7px",fontSize:9,fontWeight:700,borderRadius:99,background:C.pu,color:"#fff"}}>{rev.review_type}</span>
                    <Badge type="ok" nd>{fD(rev.created_at)}</Badge>
                  </div>
                  <div style={{fontSize:11,color:C.pu,background:"rgba(91,63,168,.07)",padding:"6px 8px",borderRadius:4,marginTop:6,lineHeight:1.5}}>{rev.notes}</div>
                </div>
              </div>
              {/* Thread history */}
              {sbData?.taskReviews.filter(r=>r.activity_id===rev.activity_id&&r.id!==rev.id).map(pr=>(
                <div key={pr.id} style={{background:pr.status==="Resolved"?C.okb:C.tl00,border:`1px solid ${pr.status==="Resolved"?"rgba(26,122,70,.15)":C.b1}`,borderRadius:5,padding:"7px 9px",marginBottom:6,fontSize:11}}>
                  <div style={{fontWeight:700,color:pr.status==="Resolved"?C.ok:C.t2,marginBottom:2}}>{pr.va_name} — {pr.review_type} <span style={{fontWeight:400,color:C.b4}}>{fD(pr.created_at)}</span> <Badge type={pr.status==="Resolved"?"ok":"ne"} nd>{pr.status}</Badge></div>
                  <div style={{color:C.b6}}>{pr.notes}</div>
                  {pr.pm_response&&<div style={{color:C.ok,marginTop:4,fontStyle:"italic"}}>Reply: {pr.pm_response}</div>}
                </div>
              ))}
              <input style={{...S.inp,marginBottom:7}} type="text" value={revReply[rev.id]||""} onChange={e=>setRevReply(r=>({...r,[rev.id]:e.target.value}))} placeholder="Reply to VA (optional — they'll see this in the thread)..."/>
              <div style={{display:"flex",gap:6}}>
                <button style={{...S.btn(C.hdr),...S.sm,flex:1}} onClick={()=>onResolve(rev.id,revReply[rev.id]||"")}>✓ Resolve</button>
                <button style={{...S.btnO(C.b4,C.b2),...S.sm}} onClick={()=>onResolve(rev.id,"Acknowledged")}>Acknowledge</button>
                <button style={{...S.btnO(C.er,C.er),...S.sm}} onClick={()=>onResolve(rev.id,"Declined — "+( revReply[rev.id]||""))}>Decline</button>
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Live Active Feed */}
      <div style={{...S.card,...S.ao}}>
        <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:10}}>🟢 Currently Active on Your Properties</div>
        {timers.filter(t=>allProps?.some(p=>p.property_code===t.property_code)).map(t=>(
          <div key={t.id} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px",background:C.okb,border:`1px solid rgba(26,122,70,.18)`,borderRadius:6,marginBottom:7}}>
            <div style={{width:7,height:7,borderRadius:"50%",background:C.ok,flexShrink:0}}/>
            <div style={{flex:1}}><div style={{fontSize:12,fontWeight:700,color:C.t2}}>{t.va_name} — {t.title}</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>{t.property_name} · Started {fT(t.start_time)}</div></div>
          </div>
        ))}
        {!timers.filter(t=>allProps?.some(p=>p.property_code===t.property_code)).length&&<div style={{fontSize:12,color:C.b4,fontStyle:"italic"}}>No tasks currently in progress on your properties.</div>}
      </div>

      {/* Regional Portfolio Overview */}
      {(isRegional||isAdmin)&&pmGroups.length>0&&(
        <div style={{...S.card,...S.ain}}>
          <div style={{fontSize:13,fontWeight:700,color:C.inf,marginBottom:12}}>📊 Portfolio Overview — {isRegional?"Regional":"Admin"} View</div>
          {pmGroups.map(({pm,props,vas,done,blocked,od,revs,total})=>(
            <div key={pm.Email} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"11px 13px",marginBottom:8}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{fontSize:12,fontWeight:700,color:C.t2}}>{pm.Name}</div>
                  <span style={{fontSize:10,color:C.b4}}>{props.length} props · {vas.length} VA{vas.length!==1?"s":""}</span>
                </div>
                <div style={{display:"flex",gap:5}}>
                  {total>0&&<Badge type={Math.round(done/total*100)>=80?"ok":"wn"} nd>{Math.round(done/total*100)}% rate</Badge>}
                  {blocked>0&&<Badge type="er" nd>{blocked} blocked</Badge>}
                  {od>0&&<Badge type="er" nd>{od} OD</Badge>}
                  {revs>0&&<Badge type="pu" nd>{revs} reviews</Badge>}
                </div>
              </div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {props.map(prop=>{
                  const va=employees.find(e=>e.Email?.toLowerCase()===sbData?.portfolios.find(p=>p.property_code===prop.property_code&&p.is_active)?.va_email?.toLowerCase());
                  const propOD=queue.filter(t=>t.property_code===prop.property_code&&isOD(t)).length;
                  const propRevs=pendingRevs.filter(r=>r.property_code===prop.property_code).length;
                  return(
                    <div key={prop.property_code} style={{flex:"1 1 140px",minWidth:130,background:propOD>0||propRevs>0?C.erb:C.w,border:`1px solid ${propOD>0?`rgba(184,59,42,.2)`:C.b1}`,borderRadius:6,padding:"8px 10px"}}>
                      <div style={{fontSize:11,fontWeight:700,color:propOD>0?C.er:C.t2}}>{prop.property_name}</div>
                      <div style={{fontSize:9,color:C.b4,marginTop:1}}>{prop.units}u · {va?va.Name:"No VA"}</div>
                      {(propOD>0||propRevs>0)&&<div style={{display:"flex",gap:3,marginTop:4}}>{propOD>0&&<Badge type="er" nd>{propOD} OD</Badge>}{propRevs>0&&<Badge type="pu" nd>{propRevs} ⚑</Badge>}</div>}
                    </div>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Property Cards */}
      <div style={{fontSize:10,fontWeight:700,color:C.b4,textTransform:"uppercase",letterSpacing:".07em",marginBottom:8,display:"flex",alignItems:"center",gap:6}}>My Directly Managed Properties<span style={{flex:1,height:1,background:C.b1}}/></div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(230px,1fr))",gap:10,marginBottom:14}}>
        {(allProps||[]).filter(p=>p.pm_email?.toLowerCase()===myEmail).map(prop=>{
          const va=getVAForProp(prop.property_code);
          const pQ=queue.filter(t=>t.property_code===prop.property_code);
          const pOD=pQ.filter(t=>isOD(t)).length;
          const pBlocked=sbData?.activities.filter(a=>a.property_code===prop.property_code&&a.status==="Blocked"&&dAgo(a.task_date)<=7).length||0;
          const pDone=sbData?.activities.filter(a=>a.property_code===prop.property_code&&a.status==="Completed"&&dAgo(a.task_date)<=7).length||0;
          const pRevs=pendingRevs.filter(r=>r.property_code===prop.property_code).length;
          return(
            <div key={prop.property_code} style={{background:C.w,border:`1px solid ${C.b1}`,borderRadius:8,padding:14}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>{prop.property_name}</div><div style={{fontSize:10,color:C.b4}}>{prop.units}u · {prop.property_group}</div></div>
                {va?<Badge type={va.VATrackerStatus==="Out"?"er":"ok"} nd>{va.Name.split(" ")[0]}</Badge>:<Badge type="er" nd>No VA</Badge>}
              </div>
              <div style={{display:"flex",gap:7,marginBottom:10}}>
                {[{v:pDone,l:"Done 7d",c:C.ok},{v:pQ.length,l:"Pending",c:pQ.length?C.wn:C.t2},{v:pBlocked,l:"Blocked",c:pBlocked?C.er:C.t2},{v:pOD,l:"Overdue",c:pOD?C.er:C.t2}].map(({v,l,c})=>(
                  <div key={l} style={{flex:1,textAlign:"center",padding:"7px 3px",background:C.tl00,borderRadius:5}}>
                    <div style={{fontSize:17,fontWeight:700,fontFamily:M,color:c}}>{v}</div>
                    <div style={{fontSize:9,fontWeight:700,color:C.b4,textTransform:"uppercase"}}>{l}</div>
                  </div>
                ))}
              </div>
              {pRevs>0&&<div style={{background:C.pub,borderRadius:4,padding:"5px 8px",fontSize:11,color:C.pu,marginBottom:8}}>⚑ {pRevs} pending review{pRevs>1?"s":""}</div>}
              <button style={{...S.btn(C.hdr),width:"100%",fontSize:11}} onClick={()=>setSelProp(prop.property_code)}>+ Assign Task to {va?va.Name:"VA"}</button>
            </div>
          );
        })}
      </div>

      {/* Assign Task Form */}
      <div style={{...S.card,...S.at}}>
        <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:4}}>📌 Assign Task to VA</div>
        <div style={{fontSize:11,color:C.b4,marginBottom:12}}>Select a property — routes automatically to the assigned VA</div>
        <div style={{...S.row,marginBottom:9}}>
          <div style={{flex:2,minWidth:160}}><label style={S.lbl}>Property *</label>
            <select style={S.sel} value={selProp} onChange={e=>setSelProp(e.target.value)}>
              <option value="">Select property...</option>
              {(allProps||[]).map(p=>{const va=getVAForProp(p.property_code);return<option key={p.property_code} value={p.property_code}>{p.property_name} ({p.units}u){va?` → ${va.Name}`:""}</option>;})}
            </select>
          </div>
          <div style={{flex:1,minWidth:130}}><label style={S.lbl}>Category *</label><select style={S.sel} value={tCat} onChange={e=>setTCat(e.target.value)}><option value="">Select...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
          <div style={{minWidth:80}}><label style={S.lbl}>Priority</label><select style={S.sel} value={tPri} onChange={e=>setTPri(e.target.value)}><option>Normal</option><option>High</option><option>Urgent</option></select></div>
        </div>
        <input style={{...S.inp,marginBottom:9}} type="text" value={tDesc} onChange={e=>setTDesc(e.target.value)} placeholder="Task description *"/>
        <input style={{...S.inp,marginBottom:9}} type="text" value={tNotes} onChange={e=>setTNotes(e.target.value)} placeholder="Notes for VA (optional)"/>
        {selProp&&(()=>{const va=getVAForProp(selProp);return va?<div style={{fontSize:11,color:C.ok,marginBottom:9}}>→ Routes to: {va.Name}{va.VATrackerStatus==="Out"?" (⚠ currently OUT — will go to coverage)":""}</div>:<div style={{fontSize:11,color:C.er,marginBottom:9}}>⚠ No VA assigned to this property</div>;})()}
        <button style={{...S.btn(C.gold,C.t2),width:"100%"}} onClick={handleSubmit}>📌 Add to VA's Queue</button>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// HISTORY VIEW
// ═══════════════════════════════════════════
function HistoryView({sbData,role,myEmail,mgrProps,guestVAs}){
  const[view,setView]=useState("tasks");
  if(!sbData)return null;
  const propFilter=role==="manager"?new Set(mgrProps.map(p=>p.property_code)):null;
  const vaFilter=guestVAs?new Set(guestVAs):null;
  const tasks=sbData.activities.filter(a=>a.activity_type==="Task"&&a.status!=="Queued"&&a.status!=="In Progress"&&(!propFilter||propFilter.has(a.property_code))&&(!vaFilter||vaFilter.has(a.va_email))&&(role==="va"?a.va_email?.toLowerCase()===myEmail:true)).slice(0,100);
  const shifts=sbData.activities.filter(a=>a.activity_type==="Shift"&&(role==="admin"||role==="regional"||a.va_email?.toLowerCase()===myEmail)).slice(0,50);
  return(
    <div>
      <div style={{display:"flex",gap:7,marginBottom:12}}>
        <button style={view==="tasks"?S.btn(C.hdr):S.btnO(C.t2,C.b2)} onClick={()=>setView("tasks")}>Tasks</button>
        {role!=="manager"&&role!=="guest"&&<button style={view==="shifts"?S.btn(C.hdr):S.btnO(C.t2,C.b2)} onClick={()=>setView("shifts")}>Shifts</button>}
      </div>
      {view==="tasks"&&(
        <div style={{...S.card,overflowX:"auto",padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["Date","VA","Task","Property","Source","Duration","Status"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {tasks.map((t,i)=>(
                <tr key={i}>
                  <td style={{...S.td,whiteSpace:"nowrap"}}>{fD(t.task_date)}</td>
                  <td style={{...S.td,fontWeight:500,fontSize:11}}>{t.va_name}</td>
                  <td style={{...S.td,maxWidth:200}}>{t.title}{t.coverage_for_name&&<div style={{fontSize:10,color:C.wn}}>🔄 covering {t.coverage_for_name}</div>}{t.notes&&<div style={{fontSize:10,color:C.b4}}>💬 {t.notes}</div>}</td>
                  <td style={{...S.td,fontSize:11}}>{t.property_name}</td>
                  <td style={S.td}><Badge type={srcB[t.source]||"ne"} nd>{t.source}</Badge></td>
                  <td style={{...S.td,fontFamily:M,fontSize:11}}>{t.duration_min?fM(t.duration_min):"—"}</td>
                  <td style={S.td}><Badge type={stB[t.status]||"ne"} nd>{t.status}</Badge></td>
                </tr>
              ))}
            </tbody>
          </table>
          {!tasks.length&&<div style={{textAlign:"center",padding:40,color:C.b4}}>No history yet.</div>}
        </div>
      )}
      {view==="shifts"&&(
        <div style={{...S.card,overflowX:"auto",padding:0}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["Date","VA","Clock In","Clock Out","Break","Working"].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {shifts.map((s,i)=>(
                <tr key={i}>
                  <td style={{...S.td,whiteSpace:"nowrap"}}>{fD(s.task_date)}</td>
                  <td style={{...S.td,fontWeight:500}}>{s.va_name}</td>
                  <td style={{...S.td,fontFamily:M,fontSize:11}}>{fT(s.clock_in)}</td>
                  <td style={{...S.td,fontFamily:M,fontSize:11}}>{fT(s.clock_out)}</td>
                  <td style={{...S.td,fontFamily:M}}>{fM(s.break_minutes)}</td>
                  <td style={{...S.td,fontFamily:M,fontWeight:700}}>{fM(s.work_minutes)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {!shifts.length&&<div style={{textAlign:"center",padding:40,color:C.b4}}>No shifts yet.</div>}
        </div>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════
// SCORECARD VIEW
// ═══════════════════════════════════════════
function ScorecardView({sbData,role,myEmail,myVa,myProps,isMgr,employees,vas}){
  const[selVa,setSelVa]=useState(myEmail);
  if(!sbData)return null;
  const va=isMgr?vas.find(v=>v.Email?.toLowerCase()===selVa?.toLowerCase())||myVa:myVa;
  const vaEmail=va?.Email?.toLowerCase()||myEmail;
  const tasks=sbData.activities.filter(a=>a.activity_type==="Task"&&a.va_email?.toLowerCase()===vaEmail&&dAgo(a.task_date)<=7);
  const done=tasks.filter(t=>t.status==="Completed");const blocked=tasks.filter(t=>t.status==="Blocked");const cov=tasks.filter(t=>t.coverage_for_email);
  const taskMin=done.reduce((s,t)=>s+(t.duration_min||0),0);const rate=tasks.length?Math.round(done.length/tasks.length*100):0;
  const shifts=sbData.activities.filter(a=>a.activity_type==="Shift"&&a.va_email?.toLowerCase()===vaEmail&&dAgo(a.task_date)<=7);
  const shiftMin=shifts.reduce((s,a)=>s+(a.work_minutes||0),0);const util=shiftMin>0?Math.round(taskMin/shiftMin*100):0;
  const gap=Math.max(0,shiftMin-taskMin);
  const port=sbData.portfolios.filter(p=>p.va_email?.toLowerCase()===vaEmail);
  const vaPropsL=port.map(p=>sbData.properties.find(pr=>pr.property_code===p.property_code)).filter(Boolean);
  const catMap={};tasks.forEach(t=>{if(!catMap[t.category])catMap[t.category]={c:0,m:0};catMap[t.category].c++;catMap[t.category].m+=(t.duration_min||0);});
  const ints=sbData.interruptions.filter(i=>i.va_email?.toLowerCase()===vaEmail&&dAgo(i.interruption_date)<=7);
  const mets=sbData.dailyMetrics.find(m=>m.va_email?.toLowerCase()===vaEmail&&m.metric_date===today());
  return(
    <div>
      {isMgr&&(
        <div style={{...S.card,display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap",marginBottom:14}}>
          <div style={{flex:1,minWidth:200}}><label style={S.lbl}>VA</label><select style={S.sel} value={selVa} onChange={e=>setSelVa(e.target.value)}>{vas.map(v=><option key={v.Email} value={v.Email?.toLowerCase()}>{v.Name}{v.VATrackerStatus==="Out"?" [OUT]":""}</option>)}</select></div>
        </div>
      )}
      {!isMgr&&va&&<div style={{marginBottom:12}}><div style={{fontSize:18,fontWeight:700,color:C.t2}}>{va.Name}</div><div style={{fontSize:11,color:C.b4}}>{vaPropsL.length} properties · {vaPropsL.reduce((s,p)=>s+(p.units||0),0)} units</div></div>}
      <div style={{...S.row,marginBottom:14}}>
        {[{l:"Tasks",v:tasks.length},{l:"Done",v:done.length,c:C.ok},{l:"Rate",v:`${rate}%`,c:rate>=85?C.ok:rate>=60?C.wn:tasks.length?C.er:undefined},{l:"Shift",v:fM(shiftMin)},{l:"Util",v:`${util}%`,c:util>=75?C.ok:util<50&&shiftMin>0?C.er:undefined},{l:"Gap",v:fM(gap),c:gap>120?C.er:undefined},{l:"Coverage",v:cov.length,c:cov.length>0?C.wn:undefined},{l:"Interruptions",v:ints.length,c:ints.length>=5?C.wn:undefined}].map((k,i)=><KPI key={i} label={k.l} value={k.v} color={k.c}/>)}
      </div>
      <div style={{...S.row,marginBottom:12}}>
        <div style={{...S.card,flex:1,minWidth:240,margin:0}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:10}}>Time by Category</div>
          {Object.entries(catMap).sort((a,b)=>b[1].m-a[1].m).map(([cat,d])=>(
            <div key={cat} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 0",borderBottom:`1px solid ${C.b1}`}}>
              <span style={{fontSize:13}}>{catI[cat]||"📁"}</span>
              <div style={{flex:1}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{cat}</div><div style={{fontSize:10,color:C.b4}}>{d.c} tasks</div></div>
              <div style={{fontFamily:M,fontSize:12,fontWeight:700,color:C.t2}}>{fM(d.m)}</div>
            </div>
          ))}
          {!Object.keys(catMap).length&&<p style={{color:C.b4,fontSize:12}}>No completed tasks.</p>}
        </div>
        <div style={{...S.card,flex:1,minWidth:200,margin:0}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:10}}>Interruptions (7d)</div>
          {["Prospect Call","Resident Call","Vendor Call","Owner Call","Manager Call","Other"].map(type=>{const c=ints.filter(i=>i.interruption_type===type).length,min=ints.filter(i=>i.interruption_type===type).reduce((s,i)=>s+(i.duration_min||0),0);if(!c)return null;return(<div key={type} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:`1px solid ${C.b1}`,fontSize:12}}><span style={{color:C.b6}}>{type}</span><span style={{fontFamily:M,fontWeight:700,color:C.t2}}>{c} · {min}m</span></div>);})}
          {!ints.length&&<p style={{color:C.b4,fontSize:12}}>No interruptions logged.</p>}
        </div>
      </div>
      {mets&&(
        <div style={{...S.card,marginBottom:12}}>
          <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:10}}>Today's Activity Metrics</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7}}>
            {[["leads_followed_up","Leads"],["applications_processed","Apps"],["showings_scheduled","Showings"],["resident_comms","Res. Comms"],["work_orders_entered","WOs In"],["work_orders_updated","WOs Upd."],["renewals_touched","Renewals"],["manager_calls","Mgr Calls"]].map(([k,l])=>(
              <div key={k} style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:"8px 5px",textAlign:"center"}}>
                <div style={{fontSize:8.5,fontWeight:700,color:C.t3,textTransform:"uppercase",marginBottom:2}}>{l}</div>
                <div style={{fontSize:20,fontWeight:700,fontFamily:M,color:C.t2}}>{mets[k]||0}</div>
              </div>
            ))}
          </div>
          {mets.submitted&&<div style={{marginTop:8,fontSize:10,color:C.ok,fontWeight:700}}>✓ Submitted {fD(mets.submitted_at)}</div>}
        </div>
      )}
      {blocked.length>0&&(
        <div style={{...S.card,...S.ae,marginBottom:12}}>
          <div style={{fontSize:13,fontWeight:700,color:C.er,marginBottom:10}}>⚠ Blocked Tasks</div>
          {blocked.map((t,i)=><div key={i} style={{padding:"5px 0",borderBottom:`1px solid ${C.b1}`}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{t.title}</div><div style={{fontSize:10,color:C.b4}}>{t.property_name}{t.notes?` · ${t.notes}`:""}</div></div>)}
        </div>
      )}
      {/* Phase 2: Full Coaching view will be added here */}
      {isMgr&&<div style={{...S.card,background:C.tl00,border:`1px dashed ${C.tl}`,textAlign:"center",padding:20}}><div style={{fontSize:12,color:C.b4,fontWeight:600}}>📈 Full Coaching &amp; Trend Analysis — Phase 2</div><div style={{fontSize:11,color:C.b4,marginTop:4}}>Deep-dive performance coaching view coming in Phase 2.</div></div>}
    </div>
  );
}

// ═══════════════════════════════════════════
// ADMIN VIEW
// ═══════════════════════════════════════════
function AdminView({sbData,employees,vas,pms,myEmail,acct,queue,covQ,config,onToggleAbsence,onAdd,onRemove,onCloseDay,onUpdateConfig,onAssignProp,onUnassignProp,onAddProp,onSaveProp,onGT,reload,fl}){
  const[adminTab,setAdminTab]=useState("team");
  const cats=config?.categories||[];const rTasks=config?.recurring_tasks||[];

  // Team state
  const[empFilter,setEmpFilter]=useState("all");
  const[editEmpId,setEditEmpId]=useState(null);
  const[editEmpData,setEditEmpData]=useState({});
  const[newEmp,setNewEmp]=useState({name:"",email:"",jobTitle:"Virtual Assistant",roleOverride:""});

  // Recurring state
  const[showRec,setShowRec]=useState(false);
  const[rVa,setRVa]=useState("");const[rCat,setRCat]=useState("");const[rDesc,setRDesc]=useState("");const[rProp,setRProp]=useState([]);

  // Portfolio state
  const[portVa,setPortVa]=useState("");const[portProp,setPortProp]=useState("");
  const[editPropId,setEditPropId]=useState(null);const[editPropData,setEditPropData]=useState({});
  const[showAddProp,setShowAddProp]=useState(false);const[newProp,setNewProp]=useState({name:"",group:"Multifamily",units:"",pmEmail:"",afId:""});

  // Guest token state
  const[showAddGuest,setShowAddGuest]=useState(false);const[guestName,setGuestName]=useState("");const[guestCompany,setGuestCompany]=useState("");const[guestVAEmails,setGuestVAEmails]=useState([]);
  const[guests,setGuests]=useState([]);

  // Load guests
  useEffect(()=>{if(adminTab==="guests"&&sbData){(async()=>{try{const tk=await onGT();const rows=await spAll(tk,"VA_GuestTokens","");setGuests(rows.map(r=>spNorm(r,"guest")).sort((a,b)=>b.id.localeCompare(a.id)));}catch(e){console.warn("[VT] Guests:",e.message);}})();};},[adminTab,sbData]);

  // PM-eligible employees (managers, regional, admin — NOT VAs)
  const pmEligible=[...pms.filter(e=>["regional","admin"].includes(detectRole(e))).map(e=>({...e,_group:"Admin / Regional"})),...pms.filter(e=>detectRole(e)==="manager").map(e=>({...e,_group:"Property Managers"}))];

  function vaProps(email){return sbData?.portfolios.filter(p=>p.va_email?.toLowerCase()===email?.toLowerCase()).map(p=>sbData.properties.find(pr=>pr.property_code===p.property_code)).filter(Boolean)||[];}

  function addRec(){
    if(!rVa||!rCat||!rDesc)return;
    const props=rProp.length>0?rProp:[""];
    const newTasks=props.map(pc=>({vaEmail:rVa,category:rCat,description:rDesc,propertyCode:pc,active:true,recurringId:`rt-${Date.now()}-${Math.random().toString(36).slice(2,6)}`}));
    onUpdateConfig({...config,recurring_tasks:[...rTasks,...newTasks]});
    setRDesc("");setRVa("");setRCat("");setRProp([]);setShowRec(false);
  }

  function toggleRec(i){const u=[...rTasks];u[i]={...u[i],active:!u[i].active};onUpdateConfig({...config,recurring_tasks:u});}
  function deleteRec(i){if(!window.confirm("Delete this recurring task?"))return;const u=rTasks.filter((_,j)=>j!==i);onUpdateConfig({...config,recurring_tasks:u});}

  async function saveEditProp(){
    if(!editPropId)return;
    await onSaveProp(editPropId,editPropData);
    setEditPropId(null);setEditPropData({});
  }

  async function generateGuestToken(){
    if(!guestName||!guestVAEmails.length)return;
    const tk=await onGT();
    const token=`gt-${Date.now().toString(36)}-${Math.random().toString(36).slice(2,8)}`;
    try{
      const res=await gPost(tk,`${SITE}/lists/VA_GuestTokens/items`,{fields:{Title:guestName,GuestToken:token,GuestName:guestName,GuestEmail:"",Company:guestCompany,VAEmails:guestVAEmails.join(","),IsActive:true,InvitedBy:myEmail}});
      const newGuest={id:res.id,token,guest_name:guestName,company:guestCompany,va_emails:guestVAEmails,is_active:true,invited_by:myEmail,last_accessed:null};
      const url=`${window.location.origin}${window.location.pathname}?guest=${token}`;
      try{await navigator.clipboard.writeText(url);}catch{}
      setGuests(g=>[newGuest,...g]);setGuestName("");setGuestCompany("");setGuestVAEmails([]);setShowAddGuest(false);
      fl("Token created — URL copied to clipboard!");
    }catch(e){fl("Error: "+e.message);}
  }

  async function deactivateGuest(id){
    try{
      const tk=await onGT();
      await gPatch(tk,`${SITE}/lists/VA_GuestTokens/items/${id}/fields`,{IsActive:false});
      setGuests(g=>g.map(x=>x.id===id?{...x,is_active:false}:x));
    }catch(e){fl("Error: "+e.message);}
  }

  async function copyGuestUrl(token){
    const url=`${window.location.origin}${window.location.pathname}?guest=${token}`;
    try{await navigator.clipboard.writeText(url);fl("URL copied!");}catch{fl("Copy: "+url);}
  }

  const snStyle=on=>({flex:1,padding:"7px 8px",fontSize:11,fontWeight:600,color:on?C.t2:C.b4,background:on?C.w:"transparent",border:"none",cursor:"pointer",borderRadius:6,fontFamily:F,textAlign:"center",transition:"all .15s",boxShadow:on?"0 1px 3px rgba(28,55,64,.07)":"none"});

  return(
    <div>
      {/* Sub-nav */}
      <div style={{display:"flex",gap:4,marginBottom:14,padding:4,background:C.b1,borderRadius:8,flexWrap:"wrap"}}>
        {[["team","👥 Team"],["sched","🔄 Recurring"],["port","🏠 Portfolio"],["guests","🔑 Guests"],["settings","⚙ Settings"]].map(([k,n])=><button key={k} style={snStyle(adminTab===k)} onClick={()=>setAdminTab(k)}>{n}</button>)}
      </div>

      {/* ── TEAM ── */}
      {adminTab==="team"&&(
        <div>
          <div style={{...S.card,...S.ac,marginBottom:14}}>
            <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:12}}>+ Add Team Member</div>
            <div style={{...S.row,marginBottom:9}}>
              <div style={{flex:2,minWidth:140}}><label style={S.lbl}>Full Name *</label><input style={S.inp} value={newEmp.name} onChange={e=>setNewEmp(p=>({...p,name:e.target.value}))} placeholder="e.g. Jessica Park"/></div>
              <div style={{flex:2,minWidth:160}}><label style={S.lbl}>Work Email *</label><input style={S.inp} value={newEmp.email} onChange={e=>setNewEmp(p=>({...p,email:e.target.value}))} placeholder="jessica@staffpro.com"/></div>
            </div>
            <div style={{...S.row,marginBottom:9}}>
              <div style={{flex:1}}><label style={S.lbl}>Job Title</label><select style={S.sel} value={newEmp.jobTitle} onChange={e=>setNewEmp(p=>({...p,jobTitle:e.target.value}))}><option>Virtual Assistant</option><option>Property Manager</option><option>Regional/Portfolio Manager</option></select></div>
              <div style={{flex:1}}>
                <label style={{...S.lbl,display:"flex",alignItems:"center",gap:5}}>Tracker Role Override <span style={{fontSize:9,fontWeight:700,padding:"1px 5px",background:C.gold,color:C.hdr,borderRadius:99}}>KEY</span></label>
                <select style={S.sel} value={newEmp.roleOverride} onChange={e=>setNewEmp(p=>({...p,roleOverride:e.target.value}))}>
                  <option value="">Auto — match job title</option>
                  <option value="va">va — task logging only</option>
                  <option value="manager">manager — see their properties</option>
                  <option value="regional">regional — portfolio overview + coaching</option>
                  <option value="admin">admin — full access to everything</option>
                </select>
              </div>
            </div>
            <div style={{background:C.wnb,border:`1px solid rgba(168,111,8,.25)`,borderRadius:6,padding:"9px 11px",fontSize:11,color:C.b6,marginBottom:10,lineHeight:1.6}}>
              <strong style={{color:C.wn}}>⚠ Role Override takes priority over job title.</strong> A "Property Manager" with override set to <strong>admin</strong> gets full access. Leave blank to auto-detect from title. <strong>VAs are never PM-eligible regardless of override.</strong>
            </div>
            <div style={{fontSize:11,color:C.b4,marginBottom:10,lineHeight:1.5}}><strong style={{color:C.t2}}>Sign-in:</strong> Employee uses their Microsoft account at the app URL. Email must match exactly. No IT setup required.</div>
            <div style={{display:"flex",gap:7}}>
              <button style={{...S.btn(C.ok),flex:1,...S.sm}} onClick={async()=>{
                if(!newEmp.name||!newEmp.email)return;
                const tk=await onGT();
                try{
                  await gPost(tk,`${SITE}/lists/Employees/items`,{fields:{Title:newEmp.name,Name:newEmp.name,Email:newEmp.email,JobTitle:newEmp.jobTitle,VATrackerRole:newEmp.roleOverride,EmployeeActive:true}});
                  fl(`${newEmp.name} added!`);setNewEmp({name:"",email:"",jobTitle:"Virtual Assistant",roleOverride:""});await reload();
                }catch(e){fl("Error: "+e.message);}
              }}>✓ Add to Team</button>
              <button style={{...S.btnO(C.t2,C.b2),...S.sm}} onClick={()=>setNewEmp({name:"",email:"",jobTitle:"Virtual Assistant",roleOverride:""})}>Clear</button>
            </div>
          </div>

          {/* Filter */}
          <div style={{display:"flex",gap:5,marginBottom:12,flexWrap:"wrap"}}>
            {[["all",`All (${employees.filter(e=>e.EmployeeActive!==false).length})`],["va","VAs"],["manager","Managers"],["regional","Regional"],["admin","Admins"],["dis","Inactive"]].map(([k,n])=>(
              <button key={k} style={{padding:"5px 12px",fontSize:11,fontWeight:600,border:`1px solid ${empFilter===k?C.hdr:C.b2}`,borderRadius:99,background:empFilter===k?C.hdr:"transparent",color:empFilter===k?"#fff":C.b4,cursor:"pointer"}} onClick={()=>setEmpFilter(k)}>{n}</button>
            ))}
          </div>

          {/* Employee list */}
          <div style={{...S.card,padding:0,overflow:"hidden"}}>
            {[{label:"Virtual Assistants",role:"va",bg:C.tl00,tc:C.b4},{label:"Property Managers",role:"manager",bg:C.gl,tc:C.g2},{label:"Regional / Portfolio Managers",role:"regional",bg:C.infb,tc:C.inf},{label:"Admins",role:"admin",bg:C.erb,tc:C.er},{label:"Inactive",role:"dis",bg:C.b1,tc:C.b4}].map(grp=>{
              const grpEmps=employees.filter(e=>{const r=detectRole(e);if(grp.role==="dis")return e.EmployeeActive===false;return r===grp.role&&e.EmployeeActive!==false;});
              if(!grpEmps.length)return null;
              if(empFilter!=="all"&&empFilter!==grp.role)return null;
              return(
                <div key={grp.role}>
                  <div style={{padding:"8px 13px",background:grp.bg,borderBottom:`1px solid ${C.b1}`,borderTop:`1px solid ${C.b1}`,fontSize:10,fontWeight:700,textTransform:"uppercase",letterSpacing:".06em",color:grp.tc}}>{grp.label}</div>
                  {grpEmps.map(emp=>{
                    const r=detectRole(emp);const hasOverride=emp.VATrackerRole&&emp.VATrackerRole!==r;const isMe=emp.Email?.toLowerCase()===myEmail;
                    return(
                      <div key={emp.id}>
                        <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 13px",borderBottom:`1px solid ${C.b1}`,cursor:"default"}}>
                          <div style={{width:30,height:30,borderRadius:"50%",background:grp.bg,color:grp.tc,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:10,flexShrink:0}}>{emp.Name?.split(" ").map(w=>w[0]).join("").slice(0,2)}</div>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:13,fontWeight:700,color:emp.EmployeeActive===false?C.b4:C.t2}}>{emp.Name}</div>
                            <div style={{fontSize:10,color:C.b4,marginTop:1}}>{emp.Email}</div>
                            <div style={{display:"flex",alignItems:"center",gap:5,marginTop:4,flexWrap:"wrap"}}>
                              <span style={{display:"inline-flex",padding:"2px 8px",fontSize:10,fontWeight:700,borderRadius:99,background:grp.bg,color:grp.tc}}>{r||"inactive"}</span>
                              <Badge type={emp.VATrackerStatus==="Out"?"er":"ok"} nd>{emp.VATrackerStatus||"Active"}</Badge>
                              {hasOverride&&<span style={{fontSize:9,fontWeight:700,padding:"1px 6px",background:C.gold,color:C.hdr,borderRadius:99}}>Role Override</span>}
                              {isMe&&<em style={{fontSize:10,color:C.b4}}>You</em>}
                            </div>
                            {hasOverride&&<div style={{fontSize:10,color:C.b4,marginTop:2}}>Job Title: {emp.JobTitle} · Override: {emp.VATrackerRole}</div>}
                          </div>
                          <div style={{display:"flex",gap:5,flexShrink:0}}>
                            {r!=="dis"&&r==="va"&&<button style={{...S.btn(emp.VATrackerStatus==="Out"?C.ok:C.er),...S.xs}} onClick={()=>onToggleAbsence(emp)}>{emp.VATrackerStatus==="Out"?"Mark In":"Mark Out"}</button>}
                            <button style={{...S.btnO(C.t2,C.b2),...S.xs}} onClick={()=>{setEditEmpId(emp.id);setEditEmpData({name:emp.Name,email:emp.Email,jobTitle:emp.JobTitle,roleOverride:emp.VATrackerRole||""});}}>Edit</button>
                            {!isMe&&emp.EmployeeActive!==false&&<button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={async()=>{if(!window.confirm(`Deactivate ${emp.Name}?`))return;const tk=await onGT();try{await gPatch(tk,empUrl(emp.id),{EmployeeActive:false});fl(`${emp.Name} deactivated`);await reload();}catch(e){fl("Error: "+e.message);}}}>Deactivate</button>}
                            {emp.EmployeeActive===false&&<button style={{...S.btn(C.ok),...S.xs}} onClick={async()=>{const tk=await onGT();try{await gPatch(tk,empUrl(emp.id),{EmployeeActive:true});fl(`${emp.Name} reactivated`);await reload();}catch(e){fl("Error: "+e.message);}}}>Reactivate</button>}
                          </div>
                        </div>
                        {editEmpId===emp.id&&(
                          <div style={{background:C.tl00,border:`1px solid ${C.tl}`,padding:"11px 13px",borderBottom:`1px solid ${C.b1}`}}>
                            <div style={{...S.row,marginBottom:8}}>
                              <div style={{flex:1}}><label style={S.lbl}>Name</label><input style={S.inp} value={editEmpData.name} onChange={e=>setEditEmpData(p=>({...p,name:e.target.value}))}/></div>
                              <div style={{flex:1}}><label style={S.lbl}>Email</label><input style={S.inp} value={editEmpData.email} onChange={e=>setEditEmpData(p=>({...p,email:e.target.value}))}/></div>
                            </div>
                            <div style={{...S.row,marginBottom:8}}>
                              <div style={{flex:1}}><label style={S.lbl}>Job Title</label><select style={S.sel} value={editEmpData.jobTitle} onChange={e=>setEditEmpData(p=>({...p,jobTitle:e.target.value}))}><option>Virtual Assistant</option><option>Property Manager</option><option>Regional/Portfolio Manager</option></select></div>
                              <div style={{flex:1}}>
                                <label style={{...S.lbl,display:"flex",alignItems:"center",gap:5}}>Role Override <span style={{fontSize:9,fontWeight:700,padding:"1px 5px",background:C.gold,color:C.hdr,borderRadius:99}}>KEY</span></label>
                                <select style={S.sel} value={editEmpData.roleOverride} onChange={e=>setEditEmpData(p=>({...p,roleOverride:e.target.value}))}>
                                  <option value="">Auto (from job title)</option>
                                  <option value="va">va</option><option value="manager">manager</option><option value="regional">regional</option><option value="admin">admin — full access</option>
                                </select>
                              </div>
                            </div>
                            <div style={{display:"flex",gap:6}}>
                              <button style={{...S.btn(C.ok),...S.xs}} onClick={async()=>{const tk=await onGT();try{await gPatch(tk,empUrl(emp.id),{Title:editEmpData.name,Name:editEmpData.name,Email:editEmpData.email,JobTitle:editEmpData.jobTitle,VATrackerRole:editEmpData.roleOverride});fl("Saved!");setEditEmpId(null);await reload();}catch(e){fl("Error: "+e.message);}}}>✓ Save</button>
                              <button style={{...S.btnO(C.t2,C.b2),...S.xs}} onClick={()=>setEditEmpId(null)}>Cancel</button>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── RECURRING ── */}
      {adminTab==="sched"&&(
        <div style={{...S.card,...S.ac}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>🔄 Recurring Daily Tasks</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>Generated each morning · UNIQUE constraint prevents duplicates at DB level</div></div>
            <button style={{...S.btn(C.hdr),...S.sm}} onClick={()=>setShowRec(v=>!v)}>{showRec?"Cancel":"+ Add"}</button>
          </div>
          {showRec&&(
            <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:13,marginBottom:13}}>
              <div style={{...S.row,marginBottom:9}}>
                <div style={{flex:1}}><label style={S.lbl}>VA *</label><select style={S.sel} value={rVa} onChange={e=>{setRVa(e.target.value);setRProp([]);}}><option value="">Select...</option>{vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
                <div style={{flex:1}}><label style={S.lbl}>Category *</label><select style={S.sel} value={rCat} onChange={e=>setRCat(e.target.value)}><option value="">Select...</option>{sortCats(cats).map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select></div>
              </div>
              <div style={{marginBottom:9}}>
                <label style={S.lbl}>Properties (check all — one task per property)</label>
                <div style={{background:C.w,border:`1px solid ${C.b2}`,borderRadius:6,padding:8,maxHeight:140,overflowY:"auto"}}>
                  <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,marginBottom:4,cursor:"pointer"}}><input type="checkbox" checked={rProp.length===0} onChange={()=>setRProp([])}/>General (no specific property)</label>
                  {(rVa?vaProps(rVa):sbData?.properties||[]).map(p=>(
                    <label key={p.property_code} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,marginBottom:3,cursor:"pointer"}}>
                      <input type="checkbox" checked={rProp.includes(p.property_code)} onChange={()=>setRProp(prev=>prev.includes(p.property_code)?prev.filter(x=>x!==p.property_code):[...prev,p.property_code])}/>
                      {p.property_name}
                    </label>
                  ))}
                </div>
              </div>
              <input style={{...S.inp,marginBottom:9}} type="text" value={rDesc} onChange={e=>setRDesc(e.target.value)} placeholder="Task description *"/>
              <button style={{...S.btn(C.ok),width:"100%"}} onClick={addRec}>✓ {rProp.length>1?`Add ${rProp.length} Tasks`:"Add Recurring Task"}</button>
            </div>
          )}
          {vas.map(va=>{
            const vr=rTasks.map((r,i)=>({...r,_i:i})).filter(r=>r.vaEmail?.toLowerCase()===va.Email?.toLowerCase());
            if(!vr.length)return null;
            return(
              <div key={va.Email} style={{marginBottom:12}}>
                <div style={{fontSize:11,fontWeight:700,color:C.t3,textTransform:"uppercase",marginBottom:5,paddingBottom:3,borderBottom:`1px solid ${C.tl}`}}>{va.Name} ({vr.filter(r=>r.active).length} active)</div>
                {vr.map(r=>{const cat=cats.find(c=>c.id===r.category);const prop=r.propertyCode?sbData?.properties.find(p=>p.property_code===r.propertyCode):null;
                  return(
                    <div key={r._i} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${C.b1}`,opacity:r.active?1:.55}}>
                      <span style={{fontSize:14}}>{catI[cat?.name]||"📁"}</span>
                      <div style={{flex:1,minWidth:0}}><div style={{fontSize:12,fontWeight:600,color:C.t2}}>{r.description}</div>{prop&&<div style={{fontSize:10,color:C.b4}}>{prop.property_name}</div>}</div>
                      <div style={{display:"flex",gap:3}}>
                        <button style={{...S.btnO(r.active?C.wn:C.ok,r.active?C.wn:C.ok),...S.xs}} onClick={()=>toggleRec(r._i)}>{r.active?"Pause":"Resume"}</button>
                        <button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={()=>deleteRec(r._i)}>✕</button>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      )}

      {/* ── PORTFOLIO ── */}
      {adminTab==="port"&&(
        <div style={{...S.card,...S.ac}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:13,fontWeight:700,color:C.t2}}>🏠 Portfolio &amp; Property Management</div>
            <button style={{...S.btn(C.gold,C.hdr),...S.sm}} onClick={()=>setShowAddProp(v=>!v)}>{showAddProp?"Cancel":"+ Add Property"}</button>
          </div>
          {showAddProp&&(
            <div style={{background:C.gl,border:`1px solid rgba(205,160,75,.3)`,borderRadius:6,padding:13,marginBottom:13}}>
              <div style={{fontSize:12,fontWeight:700,color:C.t2,marginBottom:9}}>+ Add Property</div>
              <div style={{...S.row,marginBottom:9}}>
                <div style={{flex:2}}><label style={S.lbl}>Property Name *</label><input style={S.inp} value={newProp.name} onChange={e=>setNewProp(p=>({...p,name:e.target.value}))} placeholder="e.g. Sunset Ridge"/></div>
                <div style={{flex:1}}><label style={S.lbl}>Group</label><select style={S.sel} value={newProp.group} onChange={e=>setNewProp(p=>({...p,group:e.target.value}))}><option>Multifamily</option><option>Single Family</option><option>Lease-Up</option></select></div>
                <div style={{minWidth:75}}><label style={S.lbl}>Units</label><input style={S.inp} type="number" value={newProp.units} onChange={e=>setNewProp(p=>({...p,units:e.target.value}))} placeholder="0"/></div>
              </div>
              <div style={{...S.row,marginBottom:9}}>
                <div style={{flex:1}}>
                  <label style={S.lbl}>Property Manager *</label>
                  <select style={S.sel} value={newProp.pmEmail} onChange={e=>{const emp=employees.find(x=>x.Email===e.target.value);setNewProp(p=>({...p,pmEmail:e.target.value,pmName:emp?.Name||""}));}}>
                    <option value="">Select PM...</option>
                    <optgroup label="Admins / Regional">{employees.filter(e=>["regional","admin"].includes(detectRole(e))&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name}</option>)}</optgroup>
                    <optgroup label="Property Managers">{employees.filter(e=>detectRole(e)==="manager"&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name}</option>)}</optgroup>
                  </select>
                </div>
                <div style={{flex:1}}><label style={S.lbl}>AppFolio ID (optional)</label><input style={S.inp} value={newProp.afId} onChange={e=>setNewProp(p=>({...p,afId:e.target.value}))} placeholder="AF-0000"/></div>
              </div>
              <button style={{...S.btn(C.gold,C.hdr),width:"100%"}} onClick={()=>{if(!newProp.name||!newProp.units||!newProp.pmEmail)return;onAddProp({name:newProp.name,group:newProp.group,units:newProp.units,pmEmail:newProp.pmEmail,pmName:newProp.pmName,afId:newProp.afId});setNewProp({name:"",group:"Multifamily",units:"",pmEmail:"",afId:""});setShowAddProp(false);}}>✓ Add Property</button>
            </div>
          )}
          {/* VA Assignment */}
          <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:11,marginBottom:13,display:"flex",gap:8,flexWrap:"wrap",alignItems:"flex-end"}}>
            <div style={{flex:1,minWidth:140}}><label style={S.lbl}>Assign VA</label><select style={S.sel} value={portVa} onChange={e=>setPortVa(e.target.value)}><option value="">Select VA...</option>{vas.map(v=><option key={v.Email} value={v.Email}>{v.Name}</option>)}</select></div>
            <div style={{flex:2,minWidth:180}}><label style={S.lbl}>To Property</label><select style={S.sel} value={portProp} onChange={e=>setPortProp(e.target.value)}><option value="">Select...</option>{(sbData?.properties||[]).map(p=>{const cur=sbData?.portfolios.find(pt=>pt.property_code===p.property_code&&pt.is_active);return<option key={p.property_code} value={p.property_code}>{p.property_name} ({p.units}u){cur?` — ${cur.va_name}`:""}</option>;})</select></div>
            <button style={{...S.btn(C.ok),...S.sm,opacity:(!portVa||!portProp)?.5:1}} onClick={()=>{if(!portVa||!portProp)return;const va=vas.find(v=>v.Email===portVa);onAssignProp(portVa,va?.Name||portVa,portProp);setPortProp("");}}>Assign</button>
          </div>
          {/* Property table */}
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><th style={S.th}>Property</th><th style={S.th}>Group</th><th style={S.th}>Units</th><th style={S.th}>Property Manager</th><th style={S.th}>VA Assigned</th><th style={S.th}></th></tr></thead>
              <tbody>
                {(sbData?.properties||[]).map(p=>{
                  const port=sbData?.portfolios.find(pt=>pt.property_code===p.property_code&&pt.is_active);
                  const va=port?employees.find(e=>e.Email?.toLowerCase()===port.va_email?.toLowerCase()):null;
                  const isEditP=editPropId===p.id;
                  return(
                    <tr key={p.id}>
                      <td style={S.td}>{isEditP?<input style={{...S.inp,width:150,padding:"4px 8px",fontSize:11}} value={editPropData.property_name??p.property_name} onChange={e=>setEditPropData(d=>({...d,property_name:e.target.value}))}/>:<strong style={{color:p.pm_email?C.t2:C.er}}>{p.property_name}</strong>}</td>
                      <td style={{...S.td,fontSize:11}}>{p.property_group}</td>
                      <td style={S.td}>{isEditP?<input style={{...S.inp,width:60,padding:"4px 8px",fontSize:11}} type="number" value={editPropData.units??p.units} onChange={e=>setEditPropData(d=>({...d,units:parseInt(e.target.value)||0}))}/>:p.units}</td>
                      <td style={{...S.td,fontSize:11}}>
                        {isEditP?(
                          <select style={{...S.sel,fontSize:11,padding:"4px 8px"}} value={editPropData.pm_email??p.pm_email} onChange={e=>{const emp=employees.find(x=>x.Email===e.target.value);setEditPropData(d=>({...d,pm_email:e.target.value,pm_name:emp?.Name||""}));}}>
                            <optgroup label="Admins / Regional">{employees.filter(e=>["regional","admin"].includes(detectRole(e))&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name}</option>)}</optgroup>
                            <optgroup label="Property Managers">{employees.filter(e=>detectRole(e)==="manager"&&e.EmployeeActive!==false).map(e=><option key={e.Email} value={e.Email}>{e.Name}</option>)}</optgroup>
                          </select>
                        ):<span style={{color:p.pm_name?C.b6:C.er}}>{p.pm_name||"⚠ Unassigned"}</span>}
                      </td>
                      <td style={{...S.td,fontSize:11}}>{va?va.Name:<span style={{color:C.b4,fontStyle:"italic"}}>Unassigned</span>}</td>
                      <td style={S.td}>
                        {isEditP?(
                          <div style={{display:"flex",gap:3}}>
                            <button style={{...S.btn(C.ok),...S.xs}} onClick={saveEditProp}>✓</button>
                            <button style={{...S.btnO(C.b4,C.b2),...S.xs}} onClick={()=>setEditPropId(null)}>✕</button>
                          </div>
                        ):(
                          <div style={{display:"flex",gap:3}}>
                            <button style={{...S.btnO(C.t3,C.tl),...S.xs}} onClick={()=>{setEditPropId(p.id);setEditPropData({});}}>Edit</button>
                            {port&&<button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={()=>{if(window.confirm(`Remove ${va?.Name} from ${p.property_name}?`))onUnassignProp(port.id,p.property_name,va?.Name||"");}} >Remove VA</button>}
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          {/* Unassigned warning */}
          {(sbData?.properties||[]).filter(p=>!sbData?.portfolios.some(pt=>pt.property_code===p.property_code&&pt.is_active)).length>0&&(
            <div style={{marginTop:10,padding:9,background:C.wnb,borderRadius:6,fontSize:11,color:C.wn}}>
              ⚠ {(sbData?.properties||[]).filter(p=>!sbData?.portfolios.some(pt=>pt.property_code===p.property_code&&pt.is_active)).length} propert(ies) have no VA assigned.
            </div>
          )}
        </div>
      )}

      {/* ── GUESTS ── */}
      {adminTab==="guests"&&(
        <div style={{...S.card,...S.at}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>🔑 Guest Token Management</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>External coordinators get a URL — no Microsoft account required</div></div>
            <button style={{...S.btn(C.gold,C.hdr),...S.sm}} onClick={()=>setShowAddGuest(v=>!v)}>{showAddGuest?"Cancel":"+ Add Guest"}</button>
          </div>
          {showAddGuest&&(
            <div style={{background:C.tl00,border:`1px solid ${C.tl}`,borderRadius:6,padding:12,marginBottom:12}}>
              <div style={{...S.row,marginBottom:9}}>
                <div style={{flex:1}}><label style={S.lbl}>Name *</label><input style={S.inp} value={guestName} onChange={e=>setGuestName(e.target.value)} placeholder="Coordinator name"/></div>
                <div style={{flex:1}}><label style={S.lbl}>Company</label><input style={S.inp} value={guestCompany} onChange={e=>setGuestCompany(e.target.value)} placeholder="StaffPro etc."/></div>
              </div>
              <div style={{marginBottom:9}}><label style={S.lbl}>Assign VAs (can see these VAs' data)</label>
                <div style={{background:C.w,border:`1px solid ${C.b2}`,borderRadius:6,padding:8}}>
                  {vas.map(v=><label key={v.Email} style={{display:"flex",alignItems:"center",gap:6,fontSize:12,marginBottom:4,cursor:"pointer"}}><input type="checkbox" checked={guestVAEmails.includes(v.Email?.toLowerCase())} onChange={()=>setGuestVAEmails(prev=>prev.includes(v.Email?.toLowerCase())?prev.filter(x=>x!==v.Email?.toLowerCase()):[...prev,v.Email?.toLowerCase()])}/>{v.Name}</label>)}
                </div>
              </div>
              <button style={{...S.btn(C.gold,C.hdr),width:"100%"}} onClick={generateGuestToken}>🔑 Generate Token &amp; Copy URL</button>
            </div>
          )}
          {guests.map(g=>(
            <div key={g.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"11px 0",borderBottom:`1px solid ${C.b1}`}}>
              <div style={{flex:1}}>
                <div style={{fontSize:12,fontWeight:700,color:g.is_active?C.t2:C.b4}}>{g.guest_name}{g.company?` — ${g.company}`:""}</div>
                <div style={{fontSize:10,color:C.b4,marginTop:1}}>VAs: {g.va_emails?.join(", ")||"none"} · Last accessed: {g.last_accessed?fD(g.last_accessed):"never"}</div>
                <div style={{fontFamily:M,fontSize:9,color:C.inf,background:C.infb,padding:"2px 6px",borderRadius:3,marginTop:4,display:"inline-block",wordBreak:"break-all"}}>{window.location.origin}{window.location.pathname}?guest={g.token}</div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:4,flexShrink:0}}>
                <Badge type={g.is_active?"ok":"ne"} nd>{g.is_active?"Active":"Inactive"}</Badge>
                {g.is_active&&<button style={{...S.btnO(C.t2,C.b2),...S.xs}} onClick={()=>copyGuestUrl(g.token)}>📋 Copy</button>}
                {g.is_active&&<button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={()=>deactivateGuest(g.id)}>Deactivate</button>}
              </div>
            </div>
          ))}
          {!guests.length&&<div style={{textAlign:"center",padding:30,color:C.b4,fontSize:12}}>No guest tokens yet.</div>}
        </div>
      )}

      {/* ── SETTINGS ── */}
      {adminTab==="settings"&&(
        <div>
          <div style={{...S.card,...S.ao}}>
            <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:12}}>⚙ Automation Settings</div>
            <div style={{fontSize:11,color:C.b4,marginBottom:12,lineHeight:1.6}}>These automations run via Supabase pg_cron. Enable in Supabase Dashboard → Database → Extensions → pg_cron, then run the SQL from DEPLOY.md. The toggles below store your preferences in config.settings.</div>
            {[{k:"autoClose",t:"Auto-close incomplete daily tasks at 11:59 PM",s:"Queued Daily tasks auto-mark Incomplete at end of day. Close Day button no longer needed."},{k:"autoSubmitMetrics",t:"Auto-submit daily metrics at 11:59 PM",s:"Locks the metrics tally for VAs who didn't manually submit."},{k:"teamsOnReview",t:"Teams message on manager review request",s:"Sends a direct Teams message to the property PM when a VA submits a review. Requires Chat.Create + ChatMessage.Send Azure permissions."},{k:"teamsOnBlock",t:"Teams message when VA marks task Blocked",s:"Notifies the property PM when their VA marks a task Blocked with the reason."},{k:"dailySummary",t:"Daily summary email to admin at 6:00 PM",s:"Sends Brandy a digest of completions, incompletes, blocked tasks, and reviews for the day via Resend."}].map(s=>{
              const on=config?.settings?.[s.k]!==false;
              return(
                <div key={s.k} style={{display:"flex",alignItems:"flex-start",gap:12,padding:"12px 0",borderBottom:`1px solid ${C.b1}`}}>
                  <div style={{flex:1}}><div style={{fontSize:13,fontWeight:600,color:C.t2}}>{s.t}</div><div style={{fontSize:11,color:C.b4,marginTop:2,lineHeight:1.5}}>{s.s}</div></div>
                  <button style={{width:40,height:22,background:on?C.ok:C.b2,borderRadius:11,position:"relative",cursor:"pointer",border:"none",transition:"background .2s",marginTop:2,flexShrink:0}} onClick={()=>onUpdateConfig({...config,settings:{...(config?.settings||{}),[s.k]:!on}})}>
                    <span style={{position:"absolute",width:16,height:16,background:"#fff",borderRadius:"50%",top:3,left:on?21:3,transition:"left .2s",boxShadow:"0 1px 2px rgba(0,0,0,.15)"}}/>
                  </button>
                </div>
              );
            })}
          </div>
          <div style={{...S.card,...S.ae}}>
            <div style={{fontSize:13,fontWeight:700,color:C.t2,marginBottom:12}}>🛑 Absence Management</div>
            {vas.map(va=>(
              <div key={va.Email} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 0",borderBottom:`1px solid ${C.b1}`}}>
                <div><div style={{fontSize:13,fontWeight:700,color:va.VATrackerStatus==="Out"?C.er:C.t2}}>{va.Name}</div><div style={{fontSize:10,color:C.b4}}>{sbData?.portfolios.filter(p=>p.va_email?.toLowerCase()===va.Email?.toLowerCase()).length} properties</div></div>
                <div style={{display:"flex",alignItems:"center",gap:7}}>
                  <Badge type={va.VATrackerStatus==="Out"?"er":"ok"}>{va.VATrackerStatus||"Active"}</Badge>
                  <button style={{...S.btn(va.VATrackerStatus==="Out"?C.ok:C.er),...S.xs}} onClick={()=>onToggleAbsence(va)}>{va.VATrackerStatus==="Out"?"Mark In":"Mark Out"}</button>
                </div>
              </div>
            ))}
          </div>
          <div style={{...S.card,borderLeft:`4px solid ${C.wn}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontSize:13,fontWeight:700,color:C.t2}}>🔄 Close Day</div><div style={{fontSize:10,color:C.b4,marginTop:2}}>Marks remaining daily tasks Incomplete, locks all metrics. Auto-close above makes this optional.</div></div>
              <button style={S.btn(C.wn,"#1a2a30")} onClick={onCloseDay}>Close Day</button>
            </div>
          </div>
          {/* Task Category Management */}
          <div style={S.card}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div style={{fontSize:13,fontWeight:700,color:C.t2}}>🏷 Task Categories</div>
              <button style={{...S.btn(C.hdr),...S.sm}} onClick={()=>{const n=prompt("New category name:");if(n){const id="c"+Date.now().toString(36);onUpdateConfig({...config,categories:[...(cats||[]),{id,name:n,icon:"folder"}]});}}}>+ Add</button>
            </div>
            {sortCats(cats).map((c,i)=>(
              <div key={c.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:`1px solid ${C.b1}`}}>
                <span style={{fontSize:15}}>{catI[c.name]||"📁"}</span>
                <div style={{flex:1,fontSize:12,fontWeight:600,color:C.t2}}>{c.name}</div>
                <button style={{...S.btnO(C.er,C.er),...S.xs}} onClick={()=>{if(!window.confirm(`Delete "${c.name}"? Existing tasks won't be affected.`))return;onUpdateConfig({...config,categories:cats.filter(x=>x.id!==c.id)});}}>✕</button>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
