const VERSION='28',DEFAULT_USER='Rabbit',DEFAULT_PASS='October28',MAX_TRIES_1=10,MAX_TRIES_2=20,AUTO_LOGOUT_H=8;
const S={orbitWB:null,orbitSheet:null,orbitHeaders:[],destWB:null,destSheet:null,destHeaders:[],orbitFile:null,destFile:null,mappings:[],newCols:[],staticRows:[],updatedWB:null,updatedName:'',dndPairs:{},manualPairs:[],changeLog:[],auditLog:[]};
const $=id=>document.getElementById(id);
const show=id=>$(id)&&$(id).classList.remove('hidden');
const hide=id=>$(id)&&$(id).classList.add('hidden');
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function getStoredUser(){return localStorage.getItem('ef_user')||DEFAULT_USER}
function getStoredPass(){return localStorage.getItem('ef_pass')||DEFAULT_PASS}
function getAttempts(){return parseInt(sessionStorage.getItem('ef_attempts')||'0')}
function setAttempts(n){sessionStorage.setItem('ef_attempts',String(n))}
function getLockUntil(){return parseInt(localStorage.getItem('ef_lock_until')||'0')}
function setLockUntil(ts){localStorage.setItem('ef_lock_until',String(ts))}

let autoLogoutTimer=null,autoLogoutWarningTimer=null,activityTime=Date.now();

function resetActivity(){
  activityTime=Date.now();
  hide('autoLogoutWarning');
  clearTimeout(autoLogoutWarningTimer);
  startAutoLogoutTimer();
}

function startAutoLogoutTimer(){
  clearTimeout(autoLogoutTimer);
  clearTimeout(autoLogoutWarningTimer);
  const logoutMs=AUTO_LOGOUT_H*60*60*1000;
  const warnMs=logoutMs-60000;
  autoLogoutWarningTimer=setTimeout(()=>{
    show('autoLogoutWarning');
    let countdown=60;
    $('alCountdown').textContent=countdown;
    const interval=setInterval(()=>{
      countdown--;
      if($('alCountdown'))$('alCountdown').textContent=countdown;
      if(countdown<=0){clearInterval(interval);doLogout();}
    },1000);
  },warnMs>0?warnMs:logoutMs);
}

['click','keydown','mousemove','touchstart'].forEach(ev=>{
  document.addEventListener(ev,resetActivity,{passive:true});
});

$('btnStayLoggedIn').addEventListener('click',()=>{hide('autoLogoutWarning');resetActivity();});
$('btnLogout').addEventListener('click',doLogout);

function doLogout(){
  hide('mainApp');
  show('loginPage');
  $('loginUser').value='';
  $('loginPass').value='';
  clearTimeout(autoLogoutTimer);
  clearTimeout(autoLogoutWarningTimer);
}

window.addEventListener('DOMContentLoaded',()=>{
  applyStoredPrefs();
  const splash=$('splash');
  setTimeout(()=>{
    splash.classList.add('fade');
    setTimeout(()=>{
      splash.style.display='none';
      const rem=localStorage.getItem('ef_remember');
      if(rem==='1'){$('loginUser').value=getStoredUser();$('rememberMe').checked=true;}
      const lastLogin=localStorage.getItem('ef_last_login');
      if(lastLogin)$('lastLoginInfo').textContent=`Last signed in: ${lastLogin}`;
      show('loginPage');
    },500);
  },2200);
});

$('togglePass').addEventListener('click',()=>{
  const inp=$('loginPass');
  inp.type=inp.type==='password'?'text':'password';
});

$('loginPass').addEventListener('keydown',e=>{if(e.key==='Enter')$('btnLogin').click();});
$('loginUser').addEventListener('keydown',e=>{if(e.key==='Enter')$('loginPass').focus();});

$('btnLogin').addEventListener('click',()=>{
  const now=Date.now();
  const lockUntil=getLockUntil();
  if(now<lockUntil){
    const secs=Math.ceil((lockUntil-now)/1000);
    show('loginLockout');
    $('loginLockout').textContent=`Too many failed attempts. Try again in ${secs} second${secs!==1?'s':''}.`;
    return;
  }
  hide('loginError');hide('loginLockout');
  const user=$('loginUser').value.trim();
  const pass=$('loginPass').value;
  if(user===getStoredUser()&&pass===getStoredPass()){
    setAttempts(0);
    const now_str=new Date().toLocaleString('en-CA',{dateStyle:'medium',timeStyle:'short'});
    localStorage.setItem('ef_last_login',now_str);
    if($('rememberMe').checked)localStorage.setItem('ef_remember','1');
    else localStorage.removeItem('ef_remember');
    hide('loginPage');show('mainApp');
    $('ribbonUser').textContent=getStoredUser();
    resetActivity();addAudit('Signed in');
  } else {
    const tries=getAttempts()+1;setAttempts(tries);
    show('loginError');
    $('loginError').textContent=`Incorrect username or password. Attempt ${tries}.`;
    if(tries>=MAX_TRIES_2){setLockUntil(Date.now()+10*60*1000);setAttempts(0);hide('loginError');show('loginLockout');$('loginLockout').textContent='Too many failed attempts. Locked for 10 minutes.';}
    else if(tries>=MAX_TRIES_1){setLockUntil(Date.now()+60*1000);hide('loginError');show('loginLockout');$('loginLockout').textContent='Too many failed attempts. Locked for 1 minute.';}
  }
});

function applyStoredPrefs(){
  const theme=localStorage.getItem('ef_theme')||'theme-saas';
  const mode=localStorage.getItem('ef_mode')||'light';
  const accent=localStorage.getItem('ef_accent')||null;
  const font=localStorage.getItem('ef_font')||'Plus Jakarta Sans';
  const fontSize=localStorage.getItem('ef_fontsize')||'14px';
  applyTheme(theme);applyMode(mode);
  if(accent)document.documentElement.style.setProperty('--accent',accent);
  applyFont(font,fontSize);
}

function applyTheme(theme){
  document.body.className=document.body.className.replace(/theme-\S+/g,'').trim();
  document.body.classList.add(theme);
  localStorage.setItem('ef_theme',theme);
  document.querySelectorAll('.theme-btn').forEach(b=>b.classList.toggle('active',b.dataset.theme===theme));
}

function applyMode(mode){
  localStorage.setItem('ef_mode',mode);
  document.querySelectorAll('.mode-btn').forEach(b=>b.classList.toggle('active',b.dataset.mode===mode));
  if(mode==='dark')applyTheme('theme-dark');
  else if(mode==='system'){const d=window.matchMedia('(prefers-color-scheme: dark)').matches;applyTheme(d?'theme-dark':'theme-saas');}
  $('quickThemeToggle').textContent=mode==='dark'?'☀️':'🌙';
}

function applyFont(font,size){
  document.documentElement.style.setProperty('--font',`'${font}',sans-serif`);
  document.documentElement.style.setProperty('--font-size',size);
  localStorage.setItem('ef_font',font);localStorage.setItem('ef_fontsize',size);
}

$('quickThemeToggle').addEventListener('click',()=>{
  const cur=localStorage.getItem('ef_mode')||'light';
  applyMode(cur==='dark'?'light':'dark');
});

document.querySelectorAll('.theme-btn').forEach(btn=>btn.addEventListener('click',()=>applyTheme(btn.dataset.theme)));
document.querySelectorAll('.mode-btn').forEach(btn=>btn.addEventListener('click',()=>applyMode(btn.dataset.mode)));

$('accentColor').addEventListener('input',e=>{
  document.documentElement.style.setProperty('--accent',e.target.value);
  localStorage.setItem('ef_accent',e.target.value);
});

$('fontSelect').addEventListener('change',e=>applyFont(e.target.value,$('fontSizeSelect').value));
$('fontSizeSelect').addEventListener('change',e=>applyFont($('fontSelect').value,e.target.value));

let openPanel=null;
document.querySelectorAll('.rnav-btn').forEach(btn=>{
  btn.addEventListener('click',()=>{
    const panelId='panel-'+btn.dataset.panel;
    if(openPanel===panelId){hide(panelId);btn.classList.remove('active');openPanel=null;}
    else{
      if(openPanel){hide(openPanel);document.querySelector(`.rnav-btn[data-panel="${openPanel.replace('panel-','')}"]`)?.classList.remove('active');}
      show(panelId);btn.classList.add('active');openPanel=panelId;
      if(panelId==='panel-shortcuts')renderShortcutsList();
      if(panelId==='panel-audit')renderAuditPanel();
      if(panelId==='panel-settings')loadSettingsUI();
    }
  });
});

document.addEventListener('click',e=>{
  if(openPanel&&!e.target.closest('.ribbon-panel')&&!e.target.closest('.rnav-btn')){
    hide(openPanel);
    document.querySelector(`.rnav-btn[data-panel="${openPanel.replace('panel-','')}"]`)?.classList.remove('active');
    openPanel=null;
  }
});

$('rpNewTransfer').addEventListener('click',()=>location.reload());
$('rpDownloadAudit').addEventListener('click',downloadAuditLog);
$('rpClearAudit').addEventListener('click',()=>{S.auditLog=[];saveAudit();renderAuditPanel();});

function loadAudit(){try{S.auditLog=JSON.parse(localStorage.getItem('ef_audit')||'[]');}catch{S.auditLog=[];}}
function saveAudit(){localStorage.setItem('ef_audit',JSON.stringify(S.auditLog.slice(-200)));}
function addAudit(msg){loadAudit();S.auditLog.push({time:new Date().toLocaleString('en-CA',{dateStyle:'medium',timeStyle:'short'}),msg});saveAudit();}
function renderAuditPanel(){
  loadAudit();
  const list=$('auditList');
  if(!S.auditLog.length){list.innerHTML='<div class="audit-empty">No audit entries yet.</div>';return;}
  list.innerHTML=[...S.auditLog].reverse().map(a=>`<div class="audit-item"><span class="audit-time">${esc(a.time)}</span><span>${esc(a.msg)}</span></div>`).join('');
}
function downloadAuditLog(){
  loadAudit();
  const txt=S.auditLog.map(a=>`[${a.time}] ${a.msg}`).join('\n');
  const blob=new Blob([txt],{type:'text/plain'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download='excelflow_audit.txt';a.click();URL.revokeObjectURL(url);
}

const DEFAULT_SHORTCUTS=[{key:'Enter',action:'Next Step',custom:false},{key:'Escape',action:'Go Back',custom:false},{key:'Ctrl+S',action:'Run Transfer',custom:false},{key:'Ctrl+D',action:'Dry Run',custom:false},{key:'Ctrl+/',action:'Open Shortcuts',custom:false}];
function loadCustomShortcuts(){try{return JSON.parse(localStorage.getItem('ef_shortcuts')||'[]');}catch{return[];}}
function saveCustomShortcuts(arr){localStorage.setItem('ef_shortcuts',JSON.stringify(arr));}
function getAllShortcuts(){const custom=loadCustomShortcuts();const customKeys=custom.map(s=>s.key);return[...DEFAULT_SHORTCUTS.filter(s=>!customKeys.includes(s.key)),...custom];}

function renderShortcutsList(){
  $('shortcutsList').innerHTML=getAllShortcuts().map(s=>`<div class="sc-item ${s.custom?'sc-custom':''}"><span>${esc(s.action)}</span><span class="sc-key">${esc(s.key)}</span></div>`).join('');
}

function renderShortcutsEditor(){
  const custom=loadCustomShortcuts();
  $('shortcutsEditor').innerHTML=custom.map((s,i)=>`<div class="sc-edit-row"><input type="text" value="${esc(s.action)}" data-idx="${i}" data-field="action" placeholder="Action name"/><input type="text" value="${esc(s.key)}" data-idx="${i}" data-field="key" placeholder="e.g. Ctrl+K"/><button class="sc-del" onclick="deleteShortcut(${i})">✕</button></div>`).join('');
}

window.deleteShortcut=i=>{const arr=loadCustomShortcuts();arr.splice(i,1);saveCustomShortcuts(arr);renderShortcutsEditor();};

$('btnAddShortcut').addEventListener('click',()=>{const arr=loadCustomShortcuts();arr.push({key:'',action:'',custom:true});saveCustomShortcuts(arr);renderShortcutsEditor();});
$('btnSaveShortcuts').addEventListener('click',()=>{
  const rows=$('shortcutsEditor').querySelectorAll('.sc-edit-row');
  const arr=[];
  rows.forEach(row=>{const inputs=row.querySelectorAll('input');arr.push({action:inputs[0].value.trim(),key:inputs[1].value.trim(),custom:true});});
  saveCustomShortcuts(arr.filter(s=>s.key&&s.action));
  renderShortcutsEditor();showSettingsMsg('Shortcuts saved!','ok');
});

document.addEventListener('keydown',e=>{
  const custom=loadCustomShortcuts();
  const keyStr=(e.ctrlKey?'Ctrl+':'')+e.key;
  const hit=custom.find(s=>s.key===keyStr||s.key===e.key);
  if(hit){handleShortcutAction(hit.action,e);return;}
  if(e.key==='Escape')handleShortcutAction('Go Back',e);
  if((e.ctrlKey||e.metaKey)&&e.key==='s')handleShortcutAction('Run Transfer',e);
  if((e.ctrlKey||e.metaKey)&&e.key==='d')handleShortcutAction('Dry Run',e);
  if((e.ctrlKey||e.metaKey)&&e.key==='/')handleShortcutAction('Open Shortcuts',e);
});

function handleShortcutAction(action,e){
  switch(action){
    case 'Next Step':
      if(!$('step-upload').classList.contains('hidden')&&!$('btnScan').disabled)$('btnScan').click();
      else if(!$('step-mapping').classList.contains('hidden'))$('btnPreview').click();
      else if(!$('step-preview').classList.contains('hidden'))$('btnTransfer').click();
      break;
    case 'Go Back':
      if(!$('step-mapping').classList.contains('hidden'))$('btnBackUpload').click();
      else if(!$('step-preview').classList.contains('hidden'))$('btnBackMapping').click();
      break;
    case 'Run Transfer':e?.preventDefault();if(!$('step-preview').classList.contains('hidden'))$('btnTransfer').click();break;
    case 'Dry Run':e?.preventDefault();if(!$('step-preview').classList.contains('hidden'))$('btnDryRun').click();break;
    case 'Open Shortcuts':e?.preventDefault();document.querySelector('.rnav-btn[data-panel="shortcuts"]')?.click();break;
  }
}

function loadSettingsUI(){
  $('setNewUser').value=getStoredUser();
  renderShortcutsEditor();
  const savedAccent=localStorage.getItem('ef_accent');
  if(savedAccent)$('accentColor').value=savedAccent;
  $('fontSelect').value=localStorage.getItem('ef_font')||'Plus Jakarta Sans';
  $('fontSizeSelect').value=localStorage.getItem('ef_fontsize')||'14px';
}

function showSettingsMsg(msg,type){
  const el=$('accountMsg');el.textContent=msg;el.className=`sf-msg ${type}`;show('accountMsg');
  setTimeout(()=>hide('accountMsg'),3000);
}

$('btnSaveAccount').addEventListener('click',()=>{
  const newUser=$('setNewUser').value.trim();
  const curPass=$('setCurrentPass').value;
  const newPass=$('setNewPass').value;
  const confirmPass=$('setConfirmPass').value;
  if(!curPass){showSettingsMsg('Enter current password to confirm.','err');return;}
  if(curPass!==getStoredPass()){showSettingsMsg('Current password is incorrect.','err');return;}
  if(newUser)localStorage.setItem('ef_user',newUser);
  if(newPass){
    if(newPass!==confirmPass){showSettingsMsg('New passwords do not match.','err');return;}
    localStorage.setItem('ef_pass',newPass);
  }
  $('ribbonUser').textContent=getStoredUser();
  $('setCurrentPass').value='';$('setNewPass').value='';$('setConfirmPass').value='';
  showSettingsMsg('Account updated!','ok');addAudit('Account credentials updated');
});

function parseFile(file){
  return new Promise((resolve,reject)=>{
    const r=new FileReader();
    r.onload=e=>{try{resolve(XLSX.read(e.target.result,{type:'array',cellText:false,cellDates:true}));}catch(err){reject(err);}};
    r.onerror=reject;r.readAsArrayBuffer(file);
  });
}

function colNumToLetter(n){let s='';n=n+1;while(n>0){const rem=(n-1)%26;s=String.fromCharCode(65+rem)+s;n=Math.floor((n-1)/26);}return s;}
function colLetterToNum(s){let n=0;for(let i=0;i<s.length;i++){n=n*26+s.charCodeAt(i)-64;}return n-1;}

function getHeadersFromSheet(wb,sheetName,headerRow,fromCol,toCol){
  const ws=wb.Sheets[sheetName];
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const hRowIdx=Math.max(0,(parseInt(headerRow)||1)-1);
  const row=data[hRowIdx]||[];
  let headers=row.map((c,i)=>({col:c,idx:i,letter:colNumToLetter(i)}));
  if(fromCol){const fi=colLetterToNum(fromCol.toUpperCase());headers=headers.filter(h=>h.idx>=fi);}
  if(toCol){const ti=colLetterToNum(toCol.toUpperCase());headers=headers.filter(h=>h.idx<=ti);}
  return headers.map(h=>({name:String(h.col).trim(),letter:h.letter,idx:h.idx})).filter(h=>h.name);
}

function getRowsFromSheet(wb,sheetName,headerRow,fromCol,toCol){
  const ws=wb.Sheets[sheetName];
  const hRowIdx=Math.max(0,(parseInt(headerRow)||1)-1);
  const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const headerRowData=data[hRowIdx]||[];
  let colFilter=null;
  if(fromCol||toCol){const fi=fromCol?colLetterToNum(fromCol.toUpperCase()):0;const ti=toCol?colLetterToNum(toCol.toUpperCase()):9999;colFilter={from:fi,to:ti};}
  const rows=[];
  for(let i=hRowIdx+1;i<data.length;i++){
    const rowArr=data[i];const obj={};
    headerRowData.forEach((h,idx)=>{
      if(!h)return;
      if(colFilter&&(idx<colFilter.from||idx>colFilter.to))return;
      obj[String(h).trim()]=rowArr[idx]??'';
    });
    rows.push(obj);
  }
  return rows;
}

async function handleFileUpload(file,side){
  const statusEl=$(side==='orbit'?'orbitStatus':'destStatus');
  const cardEl=$(side==='orbit'?'card-orbit':'card-dest');
  const wrapEl=$(side==='orbit'?'orbitSheetWrap':'destSheetWrap');
  const selEl=$(side==='orbit'?'orbitSheetSelect':'destSheetSelect');
  statusEl.textContent='Reading…';statusEl.className='ub-status';
  try{
    const wb=await parseFile(file);
    const names=wb.SheetNames;
    selEl.innerHTML=names.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
    wrapEl.style.display=names.length>1?'flex':'none';
    if(side==='orbit'){S.orbitWB=wb;S.orbitSheet=names[0];S.orbitFile=file;}
    else{S.destWB=wb;S.destSheet=names[0];S.destFile=file;}
    statusEl.textContent=`✓ ${file.name}`;statusEl.className='ub-status ok';
    cardEl.classList.add('loaded');checkBothLoaded();
  }catch(err){statusEl.textContent='✗ Could not read file';console.error(err);}
}

$('orbitFile').addEventListener('change',e=>{if(e.target.files[0])handleFileUpload(e.target.files[0],'orbit');});
$('destFile').addEventListener('change',e=>{if(e.target.files[0])handleFileUpload(e.target.files[0],'dest');});
$('orbitSheetSelect').addEventListener('change',e=>{S.orbitSheet=e.target.value;});
$('destSheetSelect').addEventListener('change',e=>{S.destSheet=e.target.value;});
function checkBothLoaded(){$('btnScan').disabled=!(S.orbitWB&&S.destWB);}

$('btnScan').addEventListener('click',()=>{
  S.orbitHeaders=getHeadersFromSheet(S.orbitWB,S.orbitSheet,$('orbitHeaderRow').value,$('orbitFromCol').value,$('orbitToCol').value);
  S.destHeaders=getHeadersFromSheet(S.destWB,S.destSheet,$('destHeaderRow').value,$('destFromCol').value,$('destToCol').value);
  buildMappingTable();buildDnD();buildManualGrid();populateFilterCols();buildCustomRowFields();
  setStep(2);hide('step-upload');show('step-mapping');
  addAudit(`Scanned — Orbit: ${S.orbitFile.name} | Dest: ${S.destFile.name}`);
});

document.querySelectorAll('.ef-tab').forEach(tab=>{
  tab.addEventListener('click',()=>{
    tab.closest('.ef-card').querySelectorAll('.ef-tab').forEach(t=>t.classList.remove('active'));
    tab.classList.add('active');
    const panelId='panel-'+tab.dataset.tab;
    tab.closest('.ef-card').querySelectorAll('.ef-panel').forEach(p=>p.classList.add('hidden'));
    show(panelId);
  });
});

function autoMatch(name,destHeaders){
  const norm=s=>s.toLowerCase().replace(/[\s_\-\.]/g,'');
  const n1=norm(name);
  let match=destHeaders.find(d=>norm(d.name)===n1);
  if(match)return{header:match,type:'exact',confidence:95};
  match=destHeaders.find(d=>{const n2=norm(d.name);return n1.includes(n2)||n2.includes(n1);});
  if(match)return{header:match,type:'fuzzy',confidence:60};
  return null;
}

function buildMappingTable(){
  const tbody=$('mappingBody');tbody.innerHTML='';S.mappings=[];
  const search=($('colSearch').value||'').toLowerCase();
  S.orbitHeaders.forEach(oh=>{
    if(search&&!oh.name.toLowerCase().includes(search))return;
    const matched=autoMatch(oh.name,S.destHeaders);
    const mapping={orbitCol:oh.name,orbitRef:oh.letter,destCol:matched?matched.header.name:'__SKIP__',skip:false,createNew:false,confidence:matched?matched.confidence:0,matchType:matched?matched.type:'none'};
    S.mappings.push(mapping);
    const tr=document.createElement('tr');
    tr.className=matched?`match-${matched.type}`:'match-none';
    const tdSrc=document.createElement('td');
    tdSrc.innerHTML=`<span class="col-tag">${esc(oh.name)}</span><span class="col-ref">${oh.letter}</span>`;
    const tdArr=document.createElement('td');tdArr.className='arrow-cell';tdArr.textContent='→';
    const tdDst=document.createElement('td');
    const sel=document.createElement('select');sel.className='map-select';
    sel.innerHTML=`<option value="__SKIP__">— skip —</option>`+S.destHeaders.map(h=>`<option value="${esc(h.name)}" ${matched&&matched.header.name===h.name?'selected':''}>${esc(h.name)} (${h.letter})</option>`).join('');
    tdDst.appendChild(sel);
    sel.addEventListener('change',()=>{mapping.destCol=sel.value;chkCreate.disabled=sel.value!=='__SKIP__';});
    const tdConf=document.createElement('td');
    if(matched){const pct=matched.confidence;const cls=pct>=90?'high':pct>=50?'mid':'low';tdConf.innerHTML=`<div class="confidence-bar"><div class="conf-track"><div class="conf-fill ${cls}" style="width:${pct}%"></div></div><span class="conf-pct">${pct}%</span></div>`;}
    const tdCreate=document.createElement('td');const chkCreate=document.createElement('input');chkCreate.type='checkbox';chkCreate.disabled=!!matched;
    chkCreate.addEventListener('change',()=>{mapping.createNew=chkCreate.checked;if(chkCreate.checked){mapping.destCol=oh.name;sel.disabled=true;}else{sel.disabled=false;mapping.destCol=sel.value;}});
    tdCreate.appendChild(chkCreate);
    const tdSkip=document.createElement('td');const chkSkip=document.createElement('input');chkSkip.type='checkbox';
    chkSkip.addEventListener('change',()=>{mapping.skip=chkSkip.checked;sel.disabled=chkSkip.checked;chkCreate.disabled=chkSkip.checked;});
    tdSkip.appendChild(chkSkip);
    tr.append(tdSrc,tdArr,tdDst,tdConf,tdCreate,tdSkip);tbody.appendChild(tr);
  });
}

$('colSearch').addEventListener('input',buildMappingTable);

let mapView='table';
$('btnSwitchMapView').addEventListener('click',()=>{
  mapView=mapView==='table'?'dnd':'table';
  $('mapTableView').classList.toggle('hidden',mapView!=='table');
  $('mapDndView').classList.toggle('hidden',mapView!=='dnd');
  $('btnSwitchMapView').textContent=mapView==='table'?'Switch to Drag & Drop':'Switch to Table View';
});

function buildDnD(){
  const orbitRow=$('dndOrbit');const destRow=$('dndDest');
  orbitRow.innerHTML='';destRow.innerHTML='';S.dndPairs={};
  S.orbitHeaders.forEach(oh=>{
    const card=document.createElement('div');card.className='dnd-card source-card';card.textContent=oh.name;card.draggable=true;card.dataset.col=oh.name;
    card.addEventListener('dragstart',e=>{e.dataTransfer.setData('text/plain',oh.name);card.classList.add('dragging');});
    card.addEventListener('dragend',()=>card.classList.remove('dragging'));
    orbitRow.appendChild(card);
  });
  S.destHeaders.forEach(dh=>{
    const card=document.createElement('div');card.className='dnd-card dest-card';card.textContent=dh.name;card.dataset.col=dh.name;
    card.addEventListener('dragover',e=>{e.preventDefault();card.classList.add('drag-over');});
    card.addEventListener('dragleave',()=>card.classList.remove('drag-over'));
    card.addEventListener('drop',e=>{
      e.preventDefault();card.classList.remove('drag-over');
      const src=e.dataTransfer.getData('text/plain');
      S.dndPairs[src]=dh.name;card.classList.add('matched');
      const srcCard=orbitRow.querySelector(`[data-col="${src}"]`);
      if(srcCard)srcCard.classList.add('matched');
      $('dndConnections').textContent=`${Object.keys(S.dndPairs).length} connection(s): ${Object.entries(S.dndPairs).map(([s,d])=>`${s}→${d}`).join(', ')}`;
      const m=S.mappings.find(mp=>mp.orbitCol===src);if(m)m.destCol=dh.name;
    });
    destRow.appendChild(card);
  });
}

$('btnToggleManual').addEventListener('click',()=>$('manualView').classList.toggle('hidden'));

function buildManualGrid(){$('manualGrid').innerHTML='';S.manualPairs=[];addManualRow();}
$('btnAddManualRow').addEventListener('click',addManualRow);

function addManualRow(){
  const pair={from:'',to:''};S.manualPairs.push(pair);const idx=S.manualPairs.length-1;
  const row=document.createElement('div');row.className='manual-row';
  row.innerHTML=`<input type="text" placeholder="Source e.g. C121" data-idx="${idx}" data-field="from"/><span class="manual-arrow">→</span><input type="text" placeholder="Destination e.g. D5" data-idx="${idx}" data-field="to"/><button class="manual-del" onclick="removeManualRow(this)">✕</button>`;
  row.querySelectorAll('input').forEach(inp=>inp.addEventListener('input',e=>{S.manualPairs[parseInt(e.target.dataset.idx)][e.target.dataset.field]=e.target.value.trim();}));
  $('manualGrid').appendChild(row);
}
window.removeManualRow=btn=>btn.closest('.manual-row').remove();

document.querySelectorAll('input[name="rowMode"]').forEach(r=>r.addEventListener('change',()=>{$('rowRangePanel').classList.toggle('hidden',r.value!=='range');$('rowFilterPanel').classList.toggle('hidden',r.value!=='filter');}));
document.querySelectorAll('input[name="addRowMode"]').forEach(r=>r.addEventListener('change',()=>{$('customRowPanel').classList.toggle('hidden',r.value!=='custom');}));
function populateFilterCols(){$('filterColSelect').innerHTML='<option value="">— choose —</option>'+S.orbitHeaders.map(h=>`<option value="${esc(h.name)}">${esc(h.name)}</option>`).join('');}
function buildCustomRowFields(){const cols=[...new Set([...S.destHeaders.map(h=>h.name),...S.newCols])];$('customRowFields').innerHTML=cols.map(c=>`<div class="custom-row-field"><label>${esc(c)}</label><input type="text" class="ef-input" data-col="${esc(c)}" placeholder="value…"/></div>`).join('');}

$('btnAddCol').addEventListener('click',()=>{const name=$('newColName').value.trim();if(!name||S.newCols.includes(name))return;S.newCols.push(name);$('newColName').value='';renderPills('newColsList',S.newCols);buildCustomRowFields();});
$('btnAddStaticRow').addEventListener('click',()=>{const label=$('newRowLabel').value.trim();if(!label||S.staticRows.includes(label))return;S.staticRows.push(label);$('newRowLabel').value='';renderPills('newRowsList',S.staticRows);});
function renderPills(cid,arr){$(cid).innerHTML=arr.map((c,i)=>`<div class="pill">${esc(c)}<button onclick="removePill('${cid}',${i})">×</button></div>`).join('');}
window.removePill=(cid,i)=>{if(cid==='newColsList'){S.newCols.splice(i,1);renderPills(cid,S.newCols);buildCustomRowFields();}else{S.staticRows.splice(i,1);renderPills(cid,S.staticRows);}};

$('btnBackUpload').addEventListener('click',()=>{setStep(1);hide('step-mapping');show('step-upload');});
$('btnPreview').addEventListener('click',()=>{buildPreview();setStep(3);hide('step-mapping');show('step-preview');});
$('btnBackMapping').addEventListener('click',()=>{setStep(2);hide('step-preview');show('step-mapping');});

function applyRowFilter(rows){
  const mode=document.querySelector('input[name="rowMode"]:checked')?.value||'all';
  if(mode==='range'){const from=parseInt($('rowFrom').value||'1')-1;const to=$('rowTo').value.trim()?parseInt($('rowTo').value):rows.length;return rows.slice(Math.max(0,from),to);}
  if(mode==='filter'){const col=$('filterColSelect').value;const val=$('filterValue').value.trim();if(col&&val)return rows.filter(r=>String(r[col]??'').trim()===val);}
  return rows;
}

function getActiveMappings(){
  return S.mappings.map(m=>({...m,destCol:S.dndPairs[m.orbitCol]||m.destCol})).filter(m=>!m.skip&&m.destCol!=='__SKIP__');
}

function buildPreview(){
  let orbitRows=getRowsFromSheet(S.orbitWB,S.orbitSheet,$('orbitHeaderRow').value,$('orbitFromCol').value,$('orbitToCol').value);
  orbitRows=applyRowFilter(orbitRows);
  const activeMaps=getActiveMappings();
  const destRowCount=getRowsFromSheet(S.destWB,S.destSheet,$('destHeaderRow').value,$('destFromCol').value,$('destToCol').value).length;
  $('previewSummary').innerHTML=`<div class="preview-chip"><strong>${orbitRows.length}</strong>Orbit Rows</div><div class="preview-chip"><strong>${activeMaps.length}</strong>Columns Mapped</div><div class="preview-chip"><strong>${destRowCount}</strong>Existing Rows</div><div class="preview-chip"><strong>${S.newCols.length}</strong>New Columns</div>`;
  const preview=orbitRows.slice(0,8);
  const headers=activeMaps.map(m=>m.destCol);
  let html='<table><thead><tr>'+headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr></thead><tbody>';
  preview.forEach(row=>{html+='<tr>'+activeMaps.map(m=>`<td>${esc(String(row[m.orbitCol]??''))}</td>`).join('')+'</tr>';});
  html+='</tbody></table>';
  $('previewScroll').innerHTML=html;
  if(orbitRows.length>8)$('previewScroll').insertAdjacentHTML('beforeend',`<p style="padding:6px 14px;font-size:11px;color:var(--text-muted)">Showing 8 of ${orbitRows.length} rows</p>`);
  detectConflicts(orbitRows,activeMaps);
}

function detectConflicts(orbitRows,activeMaps){
  const destRows=getRowsFromSheet(S.destWB,S.destSheet,$('destHeaderRow').value,$('destFromCol').value,$('destToCol').value);
  const conflicts=[];
  orbitRows.forEach((row,ri)=>{activeMaps.forEach(m=>{const destRow=destRows[ri];if(destRow&&destRow[m.destCol]!==''&&destRow[m.destCol]!==undefined)conflicts.push(`Row ${ri+1}: "${m.destCol}" already has "${destRow[m.destCol]}"`);});});
  if(conflicts.length>0){show('conflictBox');$('conflictList').innerHTML=conflicts.slice(0,5).map(c=>`<div>• ${esc(c)}</div>`).join('')+(conflicts.length>5?`<div>…and ${conflicts.length-5} more</div>`:'');}
  else hide('conflictBox');
}

document.querySelectorAll('input[name="insertMode"]').forEach(r=>r.addEventListener('change',()=>{$('newSheetNameRow').classList.toggle('hidden',r.value!=='new_sheet');}));

$('btnDryRun').addEventListener('click',()=>{
  let orbitRows=getRowsFromSheet(S.orbitWB,S.orbitSheet,$('orbitHeaderRow').value,$('orbitFromCol').value,$('orbitToCol').value);
  orbitRows=applyRowFilter(orbitRows);
  const activeMaps=getActiveMappings();
  $('drContent').innerHTML=`<p style="font-size:13px;margin-bottom:8px;color:var(--text-mid)">Would transfer <strong>${orbitRows.length}</strong> rows across <strong>${activeMaps.length}</strong> columns.</p><p style="font-size:12px;color:var(--text-muted)">Columns: ${activeMaps.map(m=>`${esc(m.orbitCol)} → ${esc(m.destCol)}`).join(', ')}</p><p style="font-size:12px;color:var(--green);margin-top:10px;font-weight:600">✓ No changes made to any file.</p>`;
  show('dryRunResult');
  addAudit(`Dry run — ${orbitRows.length} rows, ${activeMaps.length} columns`);
});

$('btnTransfer').addEventListener('click',()=>doTransfer());

async function doTransfer(){
  let orbitRows=getRowsFromSheet(S.orbitWB,S.orbitSheet,$('orbitHeaderRow').value,$('orbitFromCol').value,$('orbitToCol').value);
  orbitRows=applyRowFilter(orbitRows);
  const activeMaps=getActiveMappings();
  const insertMode=document.querySelector('input[name="insertMode"]:checked')?.value||'append';
  const addRowMode=document.querySelector('input[name="addRowMode"]:checked')?.value||'none';
  const conflictMode=document.querySelector('input[name="conflictMode"]:checked')?.value||'skip';
  const recalcTotals=$('optRecalcTotals').checked;
  S.changeLog=[];
  showProgress('Preparing…',10);await sleep(200);
  const wb=S.destWB;
  if(insertMode==='new_sheet'){
    showProgress('Writing new sheet…',50);await sleep(100);
    const sheetName=$('newSheetName').value.trim()||'Orbit Import';
    const headers=activeMaps.map(m=>m.destCol).concat(S.newCols);
    const data=[headers];
    orbitRows.forEach(row=>{const r=activeMaps.map(m=>row[m.orbitCol]??'');S.newCols.forEach(()=>r.push(''));data.push(r);});
    appendStaticRowsToData(data,headers,addRowMode);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(data),sheetName);
  } else {
    const ws=wb.Sheets[S.destSheet];
    const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
    const hRowIdx=Math.max(0,(parseInt($('destHeaderRow').value)||1)-1);
    const headerRow=raw[hRowIdx]?[...raw[hRowIdx]]:[];
    showProgress('Preparing columns…',30);await sleep(100);
    activeMaps.forEach(m=>{if(m.createNew&&!headerRow.includes(m.destCol)){headerRow.push(m.destCol);S.changeLog.push({cell:`New col: ${m.destCol}`,type:'col',oldVal:'',newVal:m.destCol});}});
    S.newCols.forEach(c=>{if(!headerRow.includes(c)){headerRow.push(c);S.changeLog.push({cell:`New col: ${c}`,type:'col',oldVal:'',newVal:c});}});
    const colIdx={};headerRow.forEach((h,i)=>{colIdx[h]=i;});
    const bodyRows=[...raw.slice(hRowIdx+1)];
    showProgress('Writing rows…',60);await sleep(100);
    const newRows=orbitRows.map((row,ri)=>{
      const r=new Array(headerRow.length).fill('');
      activeMaps.forEach(m=>{
        const ci=colIdx[m.destCol];if(ci===undefined)return;
        const existing=bodyRows[ri]?bodyRows[ri][ci]:'';
        if(existing!==''&&existing!==undefined&&conflictMode==='skip')return;
        r[ci]=row[m.orbitCol]??'';
        S.changeLog.push({cell:`${colNumToLetter(ci)}${hRowIdx+2+ri}`,type:'new',oldVal:existing,newVal:r[ci]});
      });
      return r;
    });
    let finalData=[headerRow,...bodyRows,...newRows];
    appendStaticRowsToData(finalData,headerRow,addRowMode);
    if(recalcTotals){showProgress('Updating totals…',80);await sleep(100);finalData=recalcTotalRows(finalData,S.changeLog);}
    wb.Sheets[S.destSheet]=XLSX.utils.aoa_to_sheet(finalData);
  }
  showProgress('Finalizing…',90);await sleep(200);
  S.updatedWB=wb;
  const ts=new Date().toISOString().slice(0,16).replace('T','_').replace(/:/g,'-');
  S.updatedName=`${S.destFile.name.replace(/(\.[^.]+)$/,'')}_ExcelFlow_${ts}.xlsx`;
  showProgress('Done!',100);await sleep(300);hide('transferProgressBar');
  buildChangeSummary();show('changeSummary');show('dlPanel');
  $('btnTransfer').disabled=true;
  addAudit(`Transfer complete — ${orbitRows.length} rows | ${activeMaps.length} cols | ${insertMode}`);
}

function showProgress(label,pct){show('transferProgressBar');$('tpbLabel').textContent=label;$('tpbFill').style.width=pct+'%';$('tpbPct').textContent=pct+'%';}
function sleep(ms){return new Promise(r=>setTimeout(r,ms));}

function appendStaticRowsToData(data,headerRow,addRowMode){
  if(addRowMode==='blank')data.push(new Array(headerRow.length).fill(''));
  if(addRowMode==='custom'){const r=new Array(headerRow.length).fill('');$('customRowFields').querySelectorAll('input[data-col]').forEach(inp=>{const idx=Array.isArray(headerRow)?headerRow.indexOf(inp.dataset.col):-1;if(idx!==-1)r[idx]=inp.value;});data.push(r);}
  S.staticRows.forEach(label=>{const r=new Array(headerRow.length).fill('');r[0]=label;data.push(r);});
}

function recalcTotalRows(data,changeLog){
  data.forEach((row,ri)=>{
    if(ri===0)return;
    const firstCell=String(row[0]||'').toLowerCase();
    if(firstCell.includes('total')){
      row.forEach((cell,ci)=>{
        if(ci===0)return;
        let sum=0;let hasNum=false;
        for(let r2=1;r2<ri;r2++){const v=parseFloat(data[r2][ci]);if(!isNaN(v)){sum+=v;hasNum=true;}}
        if(hasNum){const old=row[ci];row[ci]=Math.round(sum*100)/100;changeLog.push({cell:`${colNumToLetter(ci)}${ri+1}`,type:'total',oldVal:old,newVal:row[ci]});}
      });
    }
  });
  return data;
}

function buildChangeSummary(){
  if(!S.changeLog.length){$('csGrid').innerHTML='<p style="font-size:13px;color:var(--text-muted)">No changes recorded.</p>';return;}
  let html='<table><thead><tr><th>Cell</th><th>Type</th><th>Old Value</th><th>New Value</th></tr></thead><tbody>';
  S.changeLog.slice(0,50).forEach(c=>{const cls=c.type==='new'?'cell-new':c.type==='total'?'cell-total':'cell-col';html+=`<tr><td class="${cls}">${esc(c.cell)}</td><td>${esc(c.type)}</td><td>${esc(String(c.oldVal??''))}</td><td>${esc(String(c.newVal??''))}</td></tr>`;});
  if(S.changeLog.length>50)html+=`<tr><td colspan="4" style="color:var(--text-muted);font-size:11px">…and ${S.changeLog.length-50} more</td></tr>`;
  html+='</tbody></table>';$('csGrid').innerHTML=html;
}

$('btnDlUpdated').addEventListener('click',()=>{XLSX.writeFile(S.updatedWB,S.updatedName);addAudit(`Downloaded updated: ${S.updatedName}`);});
$('btnDlOriginal').addEventListener('click',()=>{const url=URL.createObjectURL(S.destFile);const a=document.createElement('a');a.href=url;a.download=S.destFile.name.replace(/(\.[^.]+)$/,'_original$1');a.click();URL.revokeObjectURL(url);addAudit('Downloaded original backup');});
$('btnDlReport').addEventListener('click',()=>{
  const lines=[`EXCEL FLOW v${VERSION} — Transfer Report`,`Generated: ${new Date().toLocaleString()}`,`User: ${getStoredUser()}`,`Source: ${S.orbitFile?.name||'—'}`,`Destination: ${S.destFile?.name||'—'}`,`Output: ${S.updatedName}`,``,`CHANGES (${S.changeLog.length} total):`, ...S.changeLog.map(c=>`  [${c.type.toUpperCase()}] ${c.cell}: "${c.oldVal}" → "${c.newVal}"`)];
  const blob=new Blob([lines.join('\n')],{type:'text/plain'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download='ExcelFlow_Report.txt';a.click();URL.revokeObjectURL(url);addAudit('Downloaded transfer report');
});
$('btnStartOver').addEventListener('click',()=>location.reload());

function setStep(n){[1,2,3].forEach(i=>{const el=$(`sp${i}`);el.classList.remove('active','done');if(i<n)el.classList.add('done');if(i===n)el.classList.add('active');});}

loadAudit();
