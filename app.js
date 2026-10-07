'use strict';
const $=s=>document.querySelector(s),CR='Developed by Dr. Vikas Panthi, PC CSE Core, SCOPE';
const ALIAS={REG_NO:['REGNO','REGISTRATION_NUMBER','REG_NUMBER'],STUDENT_NAME:['NAME','STUDENT'],COURSE_CODE:['COURSE','COURSECODE'],CLASS_ID:['CLASS','CLASSID'],FACULTY_NAME:['FACULTY','FACULTYNAME'],STUDENT_STATUS:['STATUS'],MARK_MODE:['MARKMODE','MODE'],MARK_CONSIDER:['MARKS','SCORE','MARK'],MAX_MARK:['MAXIMUM_MARK','MAXMARK','MAX_MARKS'],ERP_ID:['ERPID'],TITLE:['COURSE_TITLE'],PROGRAMME_CODE:['PROGRAMME','PROGRAM_CODE']};
const REQ=['REG_NO','COURSE_CODE','CLASS_ID','STUDENT_STATUS','MARK_MODE','MARK_CONSIDER','MAX_MARK'];
const AL={};Object.entries(ALIAS).forEach(([k,v])=>v.forEach(a=>AL[a]=k));
const TABS=['Dashboard','Visual Analytics','Analysis','Students','Grade Levels','Validation','Data Description','Summary & Conclusion'];
var T={},S={raw:[],recs:[],cnt:{},ch:{},cuts:[1.5,1,.5,0,-.5,-1],gb:'mean',gs:'p',gr:false,tab:0,sub:'course',ver:0,memo:{}};
const esc=x=>String(x??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const f=(x,d=2)=>Number.isFinite(x)?x.toFixed(d):'–';
const nk=h=>{const k=String(h).trim().toUpperCase().replace(/[^A-Z0-9]+/g,'_').replace(/^_|_$/g,'');return AL[k]||k};
const num=v=>(v===''||v==null)?NaN:Number(String(v).trim());
const msg=(t,c='err')=>{$('#msg').innerHTML=t?`<div class="${c}">${esc(t)}</div>`:''};
/* ---------- statistics ---------- */
function q(s,p){const i=(s.length-1)*p,l=Math.floor(i);return s[l]+(s[Math.min(l+1,s.length-1)]-s[l])*(i-l)}
function st(a){const n=a.length;if(!n)return null;const s=[...a].sort((x,y)=>x-y),m=a.reduce((x,y)=>x+y,0)/n;let s2=0,s3=0,s4=0;a.forEach(x=>{const d=x-m;s2+=d*d;s3+=d**3;s4+=d**4});
const v=s2/n,sd=Math.sqrt(v),q1=q(s,.25),q3=q(s,.75),iq=q3-q1,fr={};a.forEach(x=>fr[x]=(fr[x]||0)+1);let mo=null,mc=1;for(const k in fr)if(fr[k]>mc){mc=fr[k];mo=+k}
return{n,mean:m,med:q(s,.5),mode:mo,sd,ssd:n>1?Math.sqrt(s2/(n-1)):NaN,v,cv:m?sd/m*100:NaN,min:s[0],max:s[n-1],range:s[n-1]-s[0],p10:q(s,.1),q1,q3,p90:q(s,.9),iqr:iq,sk:sd?s3/n/sd**3:NaN,ku:sd?s4/n/sd**4-3:NaN,out:a.filter(x=>x<q1-1.5*iq||x>q3+1.5*iq).length}}
function pearson(x,y){const n=x.length;if(n<3)return NaN;const mx=x.reduce((a,b)=>a+b)/n,my=y.reduce((a,b)=>a+b)/n;let a=0,b=0,c=0;for(let i=0;i<n;i++){a+=(x[i]-mx)*(y[i]-my);b+=(x[i]-mx)**2;c+=(y[i]-my)**2}return b&&c?a/Math.sqrt(b*c):NaN}
/* ---------- load / validate ---------- */
async function load(files){try{msg('Reading files…','warn');let rows=[];
for(const fl of files){const wb=XLSX.read(await fl.arrayBuffer(),{type:'array'});wb.SheetNames.forEach(n=>{XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:''}).forEach(r=>{const o={};for(const k in r)o[nk(k)]=typeof r[k]==='string'?r[k].trim():r[k];rows.push(o)})})}
if(!rows.length)throw new Error('No data rows found.');const cols=new Set(Object.keys(rows[0]));const miss=REQ.filter(c=>!cols.has(c));
if(miss.length)throw new Error('Unable to analyze: missing column(s) '+miss.join(', ')+'. Rename the headers or use a recognised alias.');
S.raw=rows;process();msg('Loaded '+rows.length+' rows.','ok')}catch(e){msg(e.message)}}
function process(){const raw=S.raw,seen=new Set(),c={up:raw.length,dup:0,dupKey:0,missMarks:0,invMarks:0,badMode:0,badStatus:0,zero:0,full:0,maxMis:0,missFac:0,missId:0};const recs=[],kc={},mx={};
raw.forEach(r=>{const js=JSON.stringify(r);if(seen.has(js)){c.dup++;return}seen.add(js);
const mm=String(r.MARK_MODE).toUpperCase().replace(/[^A-Z0-9]/g,''),mode=['CAT1','CATI'].includes(mm)?'CAT1':['CAT2','CATII'].includes(mm)?'CAT2':'',
sr=String(r.STUDENT_STATUS).trim(),status=sr?sr[0].toUpperCase()+sr.slice(1).toLowerCase():'',marks=num(r.MARK_CONSIDER),max=num(r.MAX_MARK),notes=[];let lv='Valid';
const bad=(t,l)=>{notes.push(t);if(['Invalid','Review'].indexOf(l)>=0&&(lv!=='Invalid'))lv=l;else if(l==='Warning'&&lv==='Valid')lv='Warning'};
if(!r.REG_NO){bad('Missing REG_NO','Invalid');c.missId++}if(!r.COURSE_CODE)bad('Missing COURSE_CODE','Invalid');if(!r.CLASS_ID)bad('Missing CLASS_ID','Invalid');
if(!r.FACULTY_NAME){bad('Missing faculty','Warning');c.missFac++}
if(!mode){bad('Unsupported/missing MARK_MODE','Invalid');c.badMode++}
if(!['Present','Absent','Debarred'].includes(status)){bad('Invalid status','Review');c.badStatus++}
if(status==='Present'){if(Number.isNaN(marks)){bad('Missing/non-numeric marks','Invalid');c.missMarks++}else if(marks<0||!(max>0)||marks>max){bad('Invalid marks (negative / above MAX_MARK / MAX_MARK<=0)','Invalid');c.invMarks++}else{if(marks===0){bad('Present with zero marks','Warning');c.zero++}if(marks===max){bad('Full marks','Warning');c.full++}}}
const rec={reg:String(r.REG_NO),name:r.STUDENT_NAME||'',prog:r.PROGRAMME_CODE||'',course:String(r.COURSE_CODE),title:r.TITLE||'',cls:String(r.CLASS_ID),erp:String(r.ERP_ID||''),fn:r.FACULTY_NAME||'Unknown',status,mode,marks,max,notes,lv,valid:false};
rec.fac=(rec.erp?rec.erp+' | ':'')+rec.fn;
if(mode&&r.REG_NO){const k=[rec.reg,rec.course,rec.cls,mode].join('|');if(kc[k]){bad('Duplicate analytical key (first record kept)','Review');rec.lv='Review';c.dupKey++}else kc[k]=1}
rec.valid=status==='Present'&&!!mode&&rec.lv!=='Invalid'&&rec.lv!=='Review';if(rec.valid)rec.pct=marks/max*100;
if(mode&&max>0){(mx[mode+'|'+rec.course]??=new Set()).add(max)}recs.push(rec)});
c.maxMis=Object.values(mx).filter(s=>s.size>1).length;c.rows=recs.length;S.recs=recs;S.cnt=c;S.ver++;
const opt=(id,vals,all)=>{$(id).innerHTML=`<option value="">${all}</option>`+[...new Set(vals)].sort().map(v=>`<option>${esc(v)}</option>`).join('')};
opt('#fc',recs.map(r=>r.course),'All courses');opt('#ff',recs.map(r=>r.fac),'All faculty');render()}
/* ---------- analytical dataset (cached) ---------- */
function view(){const mode=$('#mode').value,w1=+$('#w1').value||0,w2=+$('#w2').value||0,fc=$('#fc').value,ff=$('#ff').value,sig=[S.ver,mode,w1,w2,fc,ff].join('~');
if(S.memo.sig===sig)return S.memo.v;const rs=S.recs.filter(r=>(!fc||r.course===fc)&&(!ff||r.fac===ff));let it=[],sr,u1=0,u2=0;
if(mode==='COMB'){const m={};rs.forEach(r=>{if(r.valid)(m[r.reg+'|'+r.course+'|'+r.cls]??={})[r.mode]=r});const w=w1+w2||1;
for(const k in m){const a=m[k].CAT1,b=m[k].CAT2;if(a&&b)it.push({...a,m2:b.max,marks:NaN,c1:a.marks,c2:b.marks,p1:a.pct,p2:b.pct,pct:(a.pct*w1+b.pct*w2)/w,d:b.pct-a.pct});else a?u1++:u2++}
sr=rs.filter(r=>r.mode==='CAT1')}else{sr=rs.filter(r=>r.mode===mode);it=sr.filter(r=>r.valid).map(r=>({...r}))}
const s=st(it.map(i=>i.pct)),m=new Map(),sorted=it.map(i=>i.pct).sort((a,b)=>a-b);
sorted.forEach((x,i)=>{const e=m.get(x)||{l:i,c:0};e.c++;m.set(x,e)});
it.forEach(i=>{const e=m.get(i.pct);i.pr=(e.l+e.c/2)/it.length*100;i.z=s.sd?(i.pct-s.mean)/s.sd:0});
const ms=new Set();it.forEach(i=>{ms.add(i.max);if(i.m2)ms.add(i.m2)});const K=ms.size===1?[...ms][0]:100;it.forEach(i=>{i.sc=(mode!=='COMB'&&i.max===K)?i.marks:+(i.pct*K/100).toFixed(4)});const ss=st(it.map(i=>i.sc));
const v={mode,w1,w2,K,ss,rs,sr,it,s,u1,u2,raw:mode==='COMB'?null:st(it.map(i=>i.marks)),fc,ff};S.memo={sig,v};return v}
function grades(v){const c=S.cuts.map(Number),ok=c.every((x,i)=>i===0||x<c[i-1]),s=v.ss,b=S.gb==='med'?s.med:s.mean,sd=S.gs==='s'?s.ssd:s.sd,rd=x=>S.gr?Math.ceil(x):+x.toFixed(2);
const g='SABCDEF'.split('').map((x,i)=>({g:x,k:i<6?c[i]:null,lo:i<6?rd(b+c[i]*sd):-Infinity,n:0}));
v.it.forEach(i=>{const k=g.findIndex(x=>i.sc>=x.lo);i.grade=g[k].g;g[k].n++});return{g,ok,b,sd}}
function grp(v,kf){const k=v.K/100;const M=new Map();v.sr.forEach(r=>{let g=M.get(kf(r));if(!g)M.set(kf(r),g={r,Present:0,Absent:0,Debarred:0,c:new Set(),cl:new Set(),a:[]});if(r.status in g)g[r.status]++;g.c.add(r.course);g.cl.add(r.course+r.cls)});
v.it.forEach(i=>{const g=M.get(kf(i));if(g)g.a.push(i.pct)});
return[...M.values()].map(g=>{const s=st(g.a)||{};return{course:g.r.course,title:g.r.title,cls:g.r.cls,erp:g.r.erp,fn:g.r.fn,fac:g.r.fac,nc:g.c.size,ncl:g.cl.size,n:g.a.length,mean:s.mean*k,mp:s.mean,med:s.med*k,sd:s.sd*k,cv:s.cv,q1:s.q1*k,q3:s.q3*k,min:s.min*k,max:s.max*k,P:g.Present,A:g.Absent,D:g.Debarred}})}
/* ---------- table helper ---------- */
function tbl(id,cols,data,ps=25,search=false){T[id]={cols,data,s:null,d:1,p:0,q:'',ps};return`<div class="tw">${search?`<input class="q" placeholder="Search…" oninput="T['${id}'].q=this.value.toLowerCase();T['${id}'].p=0;dt('${id}')">`:''}<div id="${id}"></div></div>`}
function dt(id){const t=T[id];let d=t.data;if(t.q)d=d.filter(r=>t.cols.some(c=>String(r[c[1]]??'').toLowerCase().includes(t.q)));
if(t.s!=null){const k=t.cols[t.s][1];d=[...d].sort((a,b)=>{const x=a[k],y=b[k];return(typeof x==='number'&&typeof y==='number'?(Number.isNaN(x)?-1e9:x)-(Number.isNaN(y)?-1e9:y):String(x).localeCompare(String(y)))*t.d})}
const pg=Math.ceil(d.length/t.ps)||1;t.p=Math.min(t.p,pg-1);const rows=d.slice(t.p*t.ps,(t.p+1)*t.ps);
$('#'+id).innerHTML=`<table><thead><tr>${t.cols.map((c,i)=>`<th onclick="T['${id}'].d=T['${id}'].s===${i}?-T['${id}'].d:1;T['${id}'].s=${i};dt('${id}')">${c[0]}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${t.cols.map(c=>`<td>${c[2]!==undefined?f(r[c[1]],c[2]):esc(r[c[1]])}</td>`).join('')}</tr>`).join('')||'<tr><td>No rows</td></tr>'}</tbody></table>
<div class="pg"><button onclick="T['${id}'].p--;dt('${id}')" ${t.p<1?'disabled':''}>Prev</button> Page ${t.p+1}/${pg} (${d.length} rows) <button onclick="T['${id}'].p++;dt('${id}')" ${t.p>=pg-1?'disabled':''}>Next</button></div>`}
/* ---------- chart helpers ---------- */
const pan=(id,t,h)=>`<div class="panel"><h3>${t}<button class="tg">Hide Chart</button></h3><div class="cw" ${h?`style="height:${h}px"`:''}><canvas id="${id}"></canvas></div></div>`;
function mk(id,cfg){const c=document.getElementById(id);if(!c)return;try{if(S.ch[id])S.ch[id].destroy();cfg.options=Object.assign({responsive:true,maintainAspectRatio:false},cfg.options||{});S.ch[id]=new Chart(c,cfg)}catch(e){c.parentElement.innerHTML='<p class="err">Chart failed: '+esc(e.message)+'</p>'}}
const ax=(x,y)=>({scales:{x:{title:{display:!!x,text:x}},y:{title:{display:!!y,text:y},beginAtZero:true}}});
const bar=(id,labels,data,label,x,y,col='#2f6fd0')=>mk(id,{type:'bar',data:{labels,datasets:[{label,data,backgroundColor:col}]},options:ax(x,y)});
document.addEventListener('click',e=>{if(e.target.classList.contains('tg')){const w=e.target.closest('.panel').querySelector('.cw'),h=w.style.display==='none';w.style.display=h?'':'none';e.target.textContent=h?'Hide Chart':'Show Chart';if(h)Object.values(S.ch).forEach(c=>c.resize())}});
function allCharts(show){document.querySelectorAll('.panel').forEach(p=>{const w=p.querySelector('.cw'),b=p.querySelector('.tg');if(w){w.style.display=show?'':'none';b.textContent=show?'Hide Chart':'Show Chart'}});Object.values(S.ch).forEach(c=>c.resize())}
/* ---------- pages ---------- */
const kp=(l,v,h)=>`<div class="kpi" title="${esc(h||'')}"><b>${v}</b><span>${l}</span></div>`;
function noData(v){if(!S.recs.length)return'<div class="warn">Upload one or more Excel/CSV files to begin.</div>';
if(!v.it.length)return`<div class="err">${v.mode==='COMB'?'Unable to create Combined analysis: no valid CAT-I / CAT-II matched pairs.':'No valid Present analytical records for the current selection.'}</div>`;return''}
const SC=[['Total records (after exact-dup removal)','rows'],['Present','P'],['Absent','A'],['Debarred','D']];
function pDash(v){const n=noData(v);if(n)return n;const s=v.s,ss=v.ss,K=v.K,c=S.cnt,cnt=x=>v.sr.filter(r=>r.status===x).length;let h='<div class="kpis">'+
kp('Total records',c.rows)+kp('Valid analytical records',v.it.length)+kp('Unique students',new Set(v.it.map(i=>i.reg)).size)+kp('Courses',new Set(v.sr.map(r=>r.course)).size)+kp('Faculty',new Set(v.sr.map(r=>r.fac)).size)+kp('Classes',new Set(v.sr.map(r=>r.course+r.cls)).size)+
kp('Present',cnt('Present'))+kp('Absent',cnt('Absent'))+kp('Debarred',cnt('Debarred'))+kp('Mean (out of '+K+')',f(ss.mean),'Arithmetic average on the marks scale')+kp('Mean %',f(s.mean))+kp('Median (out of '+K+')',f(ss.med))+kp('Median %',f(s.med))+kp('Population SD (marks)',f(ss.sd))+kp('Q1 (marks)',f(ss.q1))+kp('Q3 (marks)',f(ss.q3))+kp('Minimum (marks)',f(ss.min))+kp('Maximum (marks)',f(ss.max));
if(v.mode==='COMB'){const r=pearson(v.it.map(i=>i.p1),v.it.map(i=>i.p2));h+=kp('Matched pairs',v.it.length)+kp('Unmatched CAT-I',v.u1)+kp('Unmatched CAT-II',v.u2)+kp('Improved',v.it.filter(i=>i.d>0).length)+kp('Declined',v.it.filter(i=>i.d<0).length)+kp('Unchanged',v.it.filter(i=>i.d===0).length)+kp('Pearson r',f(r,3),'Linear correlation between CAT-I % and CAT-II %')}
return h+'</div>'}
function pVis(v){const n=noData(v);if(n)return n;return`<button onclick="allCharts(false)">Hide All Charts</button> <button onclick="allCharts(true)">Show All Charts</button><div class="grid2">${pan('c1','Marks Distribution (normalized %)')}${pan('c2','Student Status')}${pan('c3','Percentile / Quartile Profile')}${pan('c4','Performance Bands')}${pan('c5','Data Quality')}${pan('c6','Statistical Diagnostics')}${pan('c7','Grade Distribution (scenario)')}${v.mode==='COMB'?pan('c8','CAT-I vs CAT-II (%)')+pan('c9','Improvement / Decline'):''}</div>`}
function vis(v){if(!v.it.length)return;const s=v.s,p=v.it.map(i=>i.pct),bins=Array(10).fill(0);p.forEach(x=>bins[Math.min(9,Math.floor(x/10))]++);
bar('c1',bins.map((_,i)=>`${i*10}–${i*10+10}`),bins,'Students','Normalized %','Student count');
const sc=x=>v.sr.filter(r=>r.status===x).length;bar('c2',['Present','Absent','Debarred'],[sc('Present'),sc('Absent'),sc('Debarred')],'Records','Status','Count','#3b8c6e');
bar('c3',['P10','Q1','Median','Q3','P90'],[s.p10,s.q1,s.med,s.q3,s.p90].map(x=>+x.toFixed(2)),'Score %','Percentile','Score %');
const E=[0,20,40,60,70,80,90,100],bd=E.slice(0,-1).map((lo,i)=>p.filter(x=>x>=lo&&(i===6?x<=100:x<E[i+1])).length);bar('c4',E.slice(0,-1).map((l,i)=>`${l}–${E[i+1]}%`),bd,'Students','Band','Students','#8a6fd0');
const c=S.cnt,pi=v.sr.filter(r=>r.status==='Present'),vp=pi.filter(r=>r.valid).length;bar('c5',['Valid Present','Invalid/Missing Present','Absent','Debarred','Other'],[vp,pi.length-vp,sc('Absent'),sc('Debarred'),v.sr.length-pi.length-sc('Absent')-sc('Debarred')],'Records','Category','Records','#d08a2f');
bar('c6',['Outliers (Tukey)','Zero marks','Full marks','Invalid marks','Missing Present marks','Duplicate keys'],[s.out,p.filter(x=>x===0).length,p.filter(x=>x>=100).length,c.invMarks,c.missMarks,c.dupKey],'Count','Diagnostic','Count','#c0504d');
const g=grades(v);bar('c7',g.g.map(x=>x.g),g.g.map(x=>x.n),'Students','Grade','Students','#2f6fd0');
if(v.mode==='COMB'){mk('c8',{type:'scatter',data:{datasets:[{label:'Students',data:v.it.map(i=>({x:i.p1,y:i.p2})),backgroundColor:'rgba(47,111,208,.5)'},{label:'y = x',type:'line',data:[{x:0,y:0},{x:100,y:100}],borderColor:'#c0504d',pointRadius:0}]},options:{scales:{x:{title:{display:true,text:'CAT-I %'},min:0,max:100},y:{title:{display:true,text:'CAT-II %'},min:0,max:100}}}});
bar('c9',['Improved','Declined','Unchanged'],[v.it.filter(i=>i.d>0).length,v.it.filter(i=>i.d<0).length,v.it.filter(i=>i.d===0).length],'Students','Change','Students','#3b8c6e')}}
const GC={course:[['Course','course'],['Title','title'],['N','n',0],['Mean (marks)','mean'],['Mean %','mp'],['Median (marks)','med'],['SD (marks)','sd'],['CV %','cv'],['Q1','q1'],['Q3','q3'],['Min','min'],['Max','max'],['Present','P'],['Absent','A'],['Debarred','D']],
faculty:[['ERP ID','erp'],['Faculty','fn'],['Courses','nc'],['Classes','ncl'],['N','n'],['Mean (marks)','mean'],['Mean %','mp'],['Median (marks)','med'],['SD (marks)','sd'],['CV %','cv'],['Q1','q1'],['Q3','q3'],['Min','min'],['Max','max'],['Present','P'],['Absent','A'],['Debarred','D']],
class:[['Course','course'],['Class ID','cls'],['Faculty','fac'],['N','n'],['Mean (marks)','mean'],['Mean %','mp'],['Median (marks)','med'],['SD (marks)','sd'],['CV %','cv'],['Present','P'],['Absent','A'],['Debarred','D']]};
GC.course.forEach(c=>{if(['mean','mp','med','sd','cv','q1','q3','min','max'].includes(c[1]))c[2]=2});GC.faculty.forEach(c=>{if(['mean','mp','med','sd','cv','q1','q3','min','max'].includes(c[1]))c[2]=2});GC.class.forEach(c=>{if(['mean','mp','med','sd','cv'].includes(c[1]))c[2]=2});
const KF={course:r=>r.course,faculty:r=>r.fac,class:r=>r.course+'|'+r.cls};
function pAn(v){const n=noData(v);if(n)return n;const rows=grp(v,KF[S.sub]).sort((a,b)=>(b.mean||-1)-(a.mean||-1));
return`<p>${['course','faculty','class'].map(k=>`<button class="${S.sub===k?'on':''}" onclick="S.sub='${k}';render()">${k[0].toUpperCase()+k.slice(1)} Analysis</button>`).join(' ')}</p>
${S.sub==='faculty'?'<div class="warn">Faculty statistics are descriptive and can be influenced by course difficulty, cohort composition, assessment design, class size, and other confounding factors.</div>':''}
${pan('a1','Mean marks (top 15, horizontal)',Math.max(260,Math.min(15,rows.length)*26+60))}${pan('a2','Q1–Q3 range in marks (central 50%, top 15 by mean)',Math.max(260,Math.min(15,rows.length)*26+60))}${tbl('ta',GC[S.sub],rows,20,true)}`}
function an(v){if(!v.it.length)return;const rows=grp(v,KF[S.sub]).filter(r=>r.n).sort((a,b)=>b.mean-a.mean).slice(0,15),lb=rows.map(r=>S.sub==='faculty'?r.fac:S.sub==='class'?r.course+' / '+r.cls:r.course);
mk('a1',{type:'bar',data:{labels:lb,datasets:[{label:'Mean (marks)',data:rows.map(r=>+r.mean.toFixed(2)),backgroundColor:'#2f6fd0'}]},options:{indexAxis:'y',scales:{x:{title:{display:true,text:'Mean (marks out of '+v.K+')'},beginAtZero:true},y:{title:{display:true,text:S.sub}}}}});
mk('a2',{type:'bar',data:{labels:lb,datasets:[{label:'Q1–Q3',data:rows.map(r=>[+r.q1.toFixed(2),+r.q3.toFixed(2)]),backgroundColor:'#8a6fd0'}]},options:{indexAxis:'y',scales:{x:{title:{display:true,text:'Score (marks)'},min:0,max:v.K}}}})}
function pStu(v){const n=noData(v);if(n)return n;grades(v);const c=v.mode==='COMB'?[['REG_NO','reg'],['Name','name'],['Course','course'],['Class','cls'],['CAT-I','c1',2],['CAT-I %','p1',2],['CAT-II','c2',2],['CAT-II %','p2',2],['Combined %','pct',2],['Combined score','sc',2],['Diff % pts','d',2],['Change','chg'],['Grade (scenario)','grade']]:
[['REG_NO','reg'],['Name','name'],['Programme','prog'],['Course','course'],['Class','cls'],['Faculty','fac'],['Status','status'],['Marks','marks',2],['Max','max',2],['%','pct',2],['Percentile','pr',1],['Z','z',2],['Grade (scenario)','grade']];
v.it.forEach(i=>i.chg=i.d>0?'Improved':i.d<0?'Declined':'No Change');return tbl('ts',c,v.it,25,true)}
function pGr(v){const n=noData(v);if(n)return n;const g=grades(v),bn=S.gb==='med'?'Median':'Mean',sel=(k,o)=>o.map(x=>`<option value="${x[0]}" ${S[k]===x[0]?'selected':''}>${x[1]}</option>`).join('');
const rows=g.g.map((x,i)=>({g:x.g,fm:i<6?`${bn} ${x.k<0?'−':'+'} ${Math.abs(x.k)} × SD`:'Below E boundary',lo:x.lo===-Infinity?NaN:x.lo,n:x.n,sh:x.n/v.it.length*100}));
return`<div class="warn"><b>STATISTICAL SCENARIO — NOT AUTOMATIC OFFICIAL GRADING</b></div>
<p>Scores are on the original marks scale (out of <b>${v.K}</b>). Base (${bn}) = <b>${f(g.b)}</b>, SD (${S.gs==='s'?'sample':'population'}) = <b>${f(g.sd)}</b>. Boundary = Base + k × SD.</p>
<p>Base <select onchange="S.gb=this.value;render()">${sel('gb',[['mean','Mean'],['med','Median']])}</select>
SD <select onchange="S.gs=this.value;render()">${sel('gs',[['p','Population'],['s','Sample']])}</select>
<label><input type="checkbox" ${S.gr?'checked':''} onchange="S.gr=this.checked;render()"> Round boundaries up to whole marks</label></p>
<p>k for S, A, B, C, D, E: ${S.cuts.map((c,i)=>`<input type="number" step="0.1" style="width:60px" value="${c}" onchange="S.cuts[${i}]=+this.value;render()">`).join(' ')}
<span class="tag ${g.ok?'Valid':'Invalid'}">${g.ok?'VERIFIED':'CHECK BOUNDARIES'}</span> <button onclick="S.cuts=[1.5,1,.5,0,-.5,-1];S.gb='mean';S.gs='p';S.gr=false;render()">Reset defaults</button></p>
${tbl('tg',[['Grade','g'],['Formula','fm'],['Lower boundary (marks)','lo',2],['Students','n',0],['Share %','sh',2]],rows,10)}${pan('g1','Grade Distribution')}`}
function gr(v){if(!v.it.length)return;const g=grades(v);bar('g1',g.g.map(x=>x.g),g.g.map(x=>x.n),'Students','Grade','Students')}
function pVal(v){const c=S.cnt;if(!S.recs.length)return noData(v);const L=(n,x)=>({s:x?'Warning':'Valid'}),rows=[['Rows uploaded',c.up,0],['Rows accepted (after exact-duplicate removal)',c.rows,0],['Exact duplicate rows removed',c.dup,1],['Duplicate analytical keys (kept first, excluded rest)',c.dupKey,2],['Missing Present marks',c.missMarks,2],['Invalid marks',c.invMarks,2],['Unsupported / missing MARK_MODE',c.badMode,2],['Invalid student status',c.badStatus,2],['Missing REG_NO',c.missId,2],['Missing faculty',c.missFac,1],['Course/mode groups with MAX_MARK mismatch',c.maxMis,1],['Present zero marks',c.zero,1],['Present full marks',c.full,1],['Unmatched CAT-I (current view)',v.u1||0,1],['Unmatched CAT-II (current view)',v.u2||0,1]].map(r=>({a:r[0],n:r[1],s:r[1]?(r[2]===2?'Review Required':r[2]===1?'Warning':'Valid'):'Valid'}));
const bad=S.recs.filter(r=>r.lv!=='Valid').map(r=>({reg:r.reg,course:r.course,cls:r.cls,mode:r.mode,status:r.status,lv:r.lv,notes:r.notes.join('; ')}));
return tbl('tv',[['Check','a'],['Count','n',0],['Result','s']],rows,20)+'<h3>Record-level issues (nothing is discarded silently)</h3>'+tbl('tvi',[['REG_NO','reg'],['Course','course'],['Class','cls'],['Mode','mode'],['Status','status'],['Level','lv'],['Notes','notes']],bad,20,true)}
const DEF=[['Mean','Average of scores.','Σx / N','Typical level; sensitive to extremes.'],['Median','Middle value.','50th percentile','Typical value robust to extremes.'],['Population SD','Spread around mean.','√(Σ(x−μ)²/N)','Larger = more dispersion.'],['Sample SD','SD with Bessel correction.','√(Σ(x−x̄)²/(N−1))','Estimate for a wider population.'],['Variance','Squared SD.','SD²','Spread in squared units.'],['CV','Relative variability.','Pop. SD / Mean × 100','<15% low, 15–30% moderate, >30% high (rule of thumb).'],['Quartiles / Percentiles','Linear interpolation (position = (N−1)p).','Q1=P25, Q3=P75','Score below which p% of records fall.'],['IQR','Central 50% width.','Q3 − Q1','Robust spread.'],['Skewness','Asymmetry.','Σ(x−μ)³/(N·SD³)','Negative: tail to low scores; positive: tail to high scores.'],['Excess kurtosis','Tail weight.','Σ(x−μ)⁴/(N·SD⁴) − 3','>0 heavier tails than normal.'],['Outliers (Tukey)','Statistically unusual observations.','< Q1−1.5·IQR or > Q3+1.5·IQR','Flag for review, not errors.'],['Z-score','Distance from mean in SDs.','(x − mean)/SD','|z|>2 is unusual.'],['Percentile rank','Position among peers (ties: average rank).','(less + ½·equal)/N × 100','Higher = stronger relative standing.'],['Pearson r','Linear association CAT-I vs CAT-II.','Σ(x−x̄)(y−ȳ)/√(Σ(x−x̄)²Σ(y−ȳ)²)','−1…1; needs ≥3 matched pairs.']];
const COLD=[['REG_NO / STUDENT_NAME','Student identity'],['COURSE_CODE / TITLE','Course'],['CLASS_ID','Class/section'],['FACULTY_NAME / ERP_ID','Faculty (ERP_ID+name used when present)'],['STUDENT_STATUS','Present/Absent/Debarred'],['MARK_MODE','CAT1 or CAT2 — the only assessment identifier'],['MARK_CONSIDER / MAX_MARK','Marks and maximum; normalized % = marks/max×100']];
function pDesc(v){const s=S.recs.length?`<p>Mode: <b>${v.mode}</b> | Course filter: ${esc(v.fc||'All')} | Faculty filter: ${esc(v.ff||'All')} | Records: ${S.cnt.rows||0} | Valid analytical: ${v.it?v.it.length:0}</p>`:noData(v);
return s+'<h3>Columns</h3>'+tbl('td1',[['Column','a'],['Meaning','b']],COLD.map(r=>({a:r[0],b:r[1]})),20)+'<h3>Statistics</h3>'+tbl('td2',[['Statistic','a'],['Definition','b'],['Formula','c'],['Interpretation','d']],DEF.map(r=>({a:r[0],b:r[1],c:r[2],d:r[3]})),20)}
/* ---------- summary text ---------- */
function summ(v){if(!v.it.length)return[['Summary','No valid analytical records for the current selection.']];const s=v.s,M={CAT1:'CAT-I',CAT2:'CAT-II',COMB:'Combined CAT-I + CAT-II'}[v.mode],c=S.cnt,out=[];
const sk=Math.abs(s.sk)<.5?'approximately symmetric':(Math.abs(s.sk)<1?'moderately ':'strongly ')+(s.sk<0?'negatively skewed, indicating greater concentration toward the upper score range':'positively skewed, indicating greater concentration toward the lower score range');
const cv=s.cv<15?'low':s.cv<30?'moderate':'high';
out.push(['Executive Summary',`The current ${M} dataset contains ${s.n} valid analytical records. On the marks scale (out of ${v.K}) the mean is ${f(v.ss.mean)} and the median is ${f(v.ss.med)}; the normalized mean is ${f(s.mean)}% with a median of ${f(s.med)}%.`]);
out.push(['Performance Distribution',`The distribution is ${sk} (skewness ${f(s.sk)}). The central 50% of scores lie between ${f(s.q1)}% and ${f(s.q3)}%; P10 is ${f(s.p10)}% and P90 is ${f(s.p90)}%.`]);
out.push(['Variability',`Population SD is ${f(s.sd)} percentage points (sample SD ${f(s.ssd)}). The coefficient of variation is ${f(s.cv)}%, indicating ${cv} score dispersion. ${s.out} statistically unusual observation(s) fall outside the Tukey fences.`]);
out.push(['Data Quality',`${c.up} rows were uploaded; ${c.dup} exact duplicate row(s) removed; ${c.dupKey} duplicate analytical key(s), ${c.invMarks} invalid mark(s), ${c.missMarks} missing Present mark(s) and ${c.badMode} unsupported MARK_MODE value(s) were flagged and excluded from statistics but kept in the Validation page. Absent/Debarred records are never treated as zero.`]);
const cr=grp(v,KF.course).filter(r=>r.n),fr=grp(v,KF.faculty).filter(r=>r.n),sr=grp(v,KF.class).filter(r=>r.n),rg=(a,k)=>a.length?`highest mean ${a.reduce((x,y)=>y.mean>x.mean?y:x)[k]} (${f(a.reduce((x,y)=>y.mean>x.mean?y:x).mean)} marks), lowest mean ${a.reduce((x,y)=>y.mean<x.mean?y:x)[k]} (${f(a.reduce((x,y)=>y.mean<x.mean?y:x).mean)} marks)`:'none';
out.push(['Course Summary',`${cr.length} course(s) analysed; ${rg(cr,'course')}.`]);
out.push(['Faculty Summary',`${fr.length} faculty record(s) analysed; ${rg(fr,'fac')}. These are descriptive values that may be influenced by course difficulty, cohort composition and assessment design; a section mean below the overall mean may merit contextual review and does not imply causation.`]);
out.push(['Class Summary',`${sr.length} class(es) analysed; ${rg(sr,'cls')}.`]);
const g=grades(v);out.push(['Grade Scenario',`Statistical scenario only (not official grading), boundaries ${g.ok?'verified':'NOT verified — check boundaries'}: ${g.g.map(x=>`${x.g}=${x.n}`).join(', ')}.`]);
if(v.mode==='COMB'){const a=v.it,r=pearson(a.map(i=>i.p1),a.map(i=>i.p2)),dm=st(a.map(i=>i.d));out.push(['CAT-I vs CAT-II',`${a.length} matched pairs (${v.u1} unmatched CAT-I, ${v.u2} unmatched CAT-II). Mean CAT-I ${f(a.reduce((x,i)=>x+i.p1,0)/a.length)}%, mean CAT-II ${f(a.reduce((x,i)=>x+i.p2,0)/a.length)}%; mean change ${f(dm.mean)} points, median change ${f(dm.med)}. Improved ${a.filter(i=>i.d>0).length}, declined ${a.filter(i=>i.d<0).length}, unchanged ${a.filter(i=>i.d===0).length}. Pearson r = ${f(r,3)}.`])}
out.push(['Review Notes','Outliers are statistically unusual observations, not incorrect marks. Descriptive statistics do not establish causation. Final academic decisions rest with the competent authority.']);
out.push(['Conclusion',`Based on the calculated values above, ${M} performance shows a mean of ${f(s.mean)}% with ${cv} dispersion across ${s.n} valid records.`]);return out}
function pSum(v){return summ(v).map(([h,t])=>`<div class="panel"><h3>${h}</h3><p>${esc(t)}</p></div>`).join('')+`<p><i>${CR}</i></p>`}
/* ---------- router ---------- */
const PAGES=[pDash,pVis,pAn,pStu,pGr,pVal,pDesc,pSum],POST=[null,vis,an,null,gr];
function render(){try{T={};const v=view();$('#nav').innerHTML=TABS.map((t,i)=>`<button class="${i===S.tab?'on':''}" onclick="S.tab=${i};S.ch={};render()">${t}</button>`).join('');
Object.values(S.ch).forEach(c=>c.destroy());S.ch={};$('#main').innerHTML=PAGES[S.tab](v);Object.keys(T).forEach(k=>$('#'+k)&&dt(k));
if(POST[S.tab])requestAnimationFrame(()=>{try{POST[S.tab](v)}catch(e){msg('Chart error: '+e.message)}})}catch(e){$('#main').innerHTML='<div class="err">Unable to render: '+esc(e.message)+'</div>'}}
/* ---------- exports ---------- */
const dl=(n,t,m)=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([t],{type:m}));a.download=n;a.click()};
$('#csv').onclick=()=>{try{const v=view();if(!v.it.length)throw new Error('Nothing to export for the current selection.');const c=['reg','name','course','cls','fac','status','marks','max','sc','pct','pr','z'],g=grades(v);
dl('filtered_analysis.csv',[c.concat('grade').join(',')].concat(v.it.map(i=>c.map(k=>'"'+String(i[k]??'').replace(/"/g,'""')+'"').concat(i.grade).join(','))).join('\n'),'text/csv')}catch(e){msg(e.message)}};
$('#shtml').onclick=()=>{try{const v=view();dl('summary.html',`<html><head><meta charset="utf-8"><title>Summary</title></head><body style="font-family:Arial;max-width:900px;margin:auto"><h1>Relative Grade Analytics — Summary</h1>${summ(v).map(([h,t])=>`<h3>${h}</h3><p>${esc(t)}</p>`).join('')}<hr><p>Student data processed locally. ${CR}</p></body></html>`,'text/html')}catch(e){msg(e.message)}};
$('#pdf').onclick=()=>{try{const v=view();if(!v.it.length)throw new Error('Nothing to report for the current selection.');const D=new jspdf.jsPDF({orientation:'landscape',format:'a4'}),s=v.s,c=S.cnt;
D.setFontSize(26);D.text('Relative Grade Analytics',148,70,{align:'center'});D.setFontSize(14);D.text('Student Assessment & Result Analysis System',148,82,{align:'center'});D.text('Mode: '+{CAT1:'CAT-I Only',CAT2:'CAT-II Only',COMB:'Combined (weights '+v.w1+'/'+v.w2+')'}[v.mode],148,100,{align:'center'});D.text('Filters — Course: '+(v.fc||'All')+' | Faculty: '+(v.ff||'All'),148,110,{align:'center'});D.text('Generated '+new Date().toLocaleString(),148,120,{align:'center'});
const T1=(h,b)=>{D.addPage();D.setFontSize(14);D.text(h,14,16);D.autoTable({startY:22,head:[b[0]],body:b.slice(1),styles:{fontSize:8},margin:{left:14,right:14,bottom:16}})};
T1('Dataset, Validation & Statistical Summary',[['Metric','Value'],['Rows uploaded',c.up],['Exact duplicates removed',c.dup],['Duplicate keys',c.dupKey],['Invalid marks',c.invMarks],['Missing Present marks',c.missMarks],['Valid analytical records',s.n],['Scale (out of)',v.K],['Mean (marks)',f(v.ss.mean)],['Median (marks)',f(v.ss.med)],['Pop SD (marks)',f(v.ss.sd)],['Mean %',f(s.mean)],['Median %',f(s.med)],['Pop SD %',f(s.sd)],['Sample SD',f(s.ssd)],['CV %',f(s.cv)],['Q1 / Q3',f(s.q1)+' / '+f(s.q3)],['P10 / P90',f(s.p10)+' / '+f(s.p90)],['Min / Max',f(s.min)+' / '+f(s.max)],['Skewness',f(s.sk)],['Excess kurtosis',f(s.ku)],['Tukey outliers',s.out]]);
const gt=(k,cols)=>{const r=grp(v,KF[k]).sort((a,b)=>(b.mean||-1)-(a.mean||-1));return[cols.map(x=>x[0])].concat(r.map(x=>cols.map(y=>y[2]!==undefined?f(x[y[1]],y[2]):x[y[1]]??'')))};
Object.keys(S.ch).forEach(id=>{try{D.addPage();D.setFontSize(14);D.text('Chart',14,16);D.addImage(S.ch[id].toBase64Image(),'PNG',14,22,180,90)}catch(e){}});
T1('Course Analysis',gt('course',GC.course));T1('Faculty Analysis (descriptive; confounders apply)',gt('faculty',GC.faculty));T1('Class Analysis',gt('class',GC.class));
const g=grades(v);T1('Grade Levels (statistical scenario, not official grading)',[['Grade','Lower boundary (marks out of '+v.K+')','Students']].concat(g.g.map(x=>[x.g,f(x.lo),x.n])));
if(v.mode==='COMB')T1('CAT-I vs CAT-II',[['Metric','Value'],['Matched pairs',v.it.length],['Pearson r',f(pearson(v.it.map(i=>i.p1),v.it.map(i=>i.p2)),3)],['Improved',v.it.filter(i=>i.d>0).length],['Declined',v.it.filter(i=>i.d<0).length],['Unchanged',v.it.filter(i=>i.d===0).length]]);
D.addPage();D.setFontSize(14);D.text('Summary & Conclusion',14,16);D.setFontSize(9);let y=24;summ(v).forEach(([h,t])=>{const L=D.splitTextToSize(t,265);if(y+L.length*4+8>195){D.addPage();y=16}D.setFont(undefined,'bold');D.text(h,14,y);D.setFont(undefined,'normal');D.text(L,14,y+5);y+=L.length*4+9});
const n=D.getNumberOfPages();for(let i=1;i<=n;i++){D.setPage(i);D.setFontSize(8);D.text('Relative Grade Analytics | '+CR,14,205);D.text('Page '+i+' / '+n,283,205,{align:'right'})}D.save('Relative_Grade_Analytics_Report.pdf')}catch(e){msg('PDF failed: '+e.message)}};
$('#file').onchange=e=>load([...e.target.files]);['mode','w1','w2','fc','ff'].forEach(id=>$('#'+id).onchange=()=>{S.ch={};render()});
render();
