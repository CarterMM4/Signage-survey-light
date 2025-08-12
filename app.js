// ====== Step 1: details + bulk edit ======

/** Sign types + short labels + colors **/
const SIGN_TYPES = {
  'FOH':'#0ea5e9','BOH':'#64748b','Elevator':'#8b5cf6','UNIT ID':'#22c55e',
  'Stair - Ingress':'#22c55e','Stair - Egress':'#ef4444','Ingress':'#22c55e',
  'Egress':'#ef4444','Hall Direct':'#f59e0b','Callbox':'#06b6d4',
  'Evac':'#84cc16','Exit':'#ef4444','Restroom':'#10b981'
};
const SHORT = {
  'Elevator':'ELV','Hall Direct':'HD','Callbox':'CB','Evac':'EV','Ingress':'ING',
  'Egress':'EGR','Exit':'EXIT','Restroom':'WC','UNIT ID':'UNIT','FOH':'FOH',
  'BOH':'BOH','Stair - Ingress':'ING','Stair - Egress':'EGR'
};

/** Tiny helpers **/
const $ = (id) => document.getElementById(id);
const on = (el, evt, fn) => el && el.addEventListener(evt, fn);
const id = () => Math.random().toString(36).slice(2,10);

/** State **/
let projects = loadProjects();
let currentProjectId = localStorage.getItem('survey:lastOpenProjectId') || null;
let project = currentProjectId ? projects.find(p=>p.id===currentProjectId) : null;
if(!project){ project = newProject('Untitled Project'); saveProject(project); }
selectProject(project.id);

// DOM refs
const thumbsEl = $('thumbs');
const stage = $('stage');
const stageImage = $('stageImage');
const pinLayer = $('pinLayer');
const projectLabel = $('projectLabel');
const inputUpload = $('inputUpload');
const inputBuilding = $('inputBuilding');
const inputLevel = $('inputLevel');
const filterType = $('filterType');
const inputSearch = $('inputSearch');
const toggleField = $('toggleField');
const selId = $('selId');
const fieldType = $('fieldType');
const fieldRoomNum = $('fieldRoomNum');
const fieldRoomName = $('fieldRoomName');
const fieldBuilding = $('fieldBuilding');
const fieldLevel = $('fieldLevel');
const fieldNotes = $('fieldNotes');
const pinList = $('pinList');

// Populate type dropdowns
(function initTypeSelects() {
  Object.keys(SIGN_TYPES).forEach(t=>{
    const o = document.createElement('option');
    o.value = t; o.textContent = t;
    filterType.appendChild(o.cloneNode(true));
    fieldType.appendChild(o);
  });
})();

// Project context (pre-fill)
const projectContext = { building:'', level:'' };
on(inputBuilding,'input',()=> projectContext.building = inputBuilding.value);
on(inputLevel,'input',()=> projectContext.level = inputLevel.value);

// Toolbar
on($('btnNew'),'click',()=>{
  const name = prompt('New project name?','New Project'); if(!name) return;
  const p = newProject(name); saveProject(p); selectProject(p.id);
});
on($('btnOpen'),'click',()=>{
  const items = projects.map(p=>`• ${p.name} (${new Date(p.updatedAt).toLocaleString()}) [${p.id}]`).join('\n');
  const id = prompt('Projects:\n'+items+'\n\nEnter project id to open:'); if(!id) return;
  const found = projects.find(p=>p.id===id); if(found) selectProject(found.id); else alert('Not found.');
});
on($('btnSaveAs'),'click',()=>{
  const name = prompt('Duplicate as name:', project.name+' (copy)'); if(!name) return;
  const copy = JSON.parse(JSON.stringify(project)); copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now();
  saveProject(copy); selectProject(copy.id);
});
on($('btnRename'),'click',()=>{
  const name = prompt('Rename project:', project.name); if(!name) return;
  project.name = name; saveProject(project); renderProjectLabel();
});

on($('btnUpload'),'click',()=> inputUpload.click());
on(inputUpload,'change', async (e)=>{
  const files=[...e.target.files]; if(!files.length) return;
  for(const f of files){
    if(f.type==='application/pdf'){
      await addPdfPages(f);
    } else if(f.type.startsWith('image/')){
      const url = URL.createObjectURL(f);
      addImagePage(url, f.name.replace(/\.[^.]+$/,''));
    }
  }
  renderAll();
});

on($('btnAddPin'),'click',()=> startAddPin());
on($('btnClearPins'),'click',()=>{
  if(!confirm('Clear ALL pins on this page?')) return;
  const pg = currentPage(); if(!pg) return;
  pg.pins = []; saveProject(project); renderAll();
});

on(filterType,'change',()=>{ renderPins(); renderPinsList(); });
on(inputSearch,'input',()=> renderPinsList());
on(toggleField,'change',()=>{ project.settings.fieldMode = !!toggleField.checked; saveProject(project); renderPins(); });

// Right-panel live edits
[fieldType, fieldRoomNum, fieldRoomName, fieldBuilding, fieldLevel, fieldNotes].forEach(el=>{
  on(el,'input',()=>{
    const p = selectedPin(); if(!p) return;
    if(el===fieldType) p.sign_type = el.value || '';
    if(el===fieldRoomNum) p.room_number = el.value || '';
    if(el===fieldRoomName) p.room_name = el.value || '';
    if(el===fieldBuilding) p.building = el.value || '';
    if(el===fieldLevel) p.level = el.value || '';
    if(el===fieldNotes) p.notes = el.value || '';
    p.lastEdited = Date.now();
    saveProject(project); renderPins(); renderPinsList(); updateDetails();
  });
});

// Bulk buttons
on($('btnBulkType'),'click',()=>{
  const type = prompt('Enter type for selected rows (exact name, case-sensitive):'); 
  if(type==null) return;
  const ids = checkedPinIds();
  ids.forEach(idv=>{ const p=findPin(idv); if(p){ p.sign_type=type; p.lastEdited=Date.now(); }});
  saveProject(project); renderPins(); renderPinsList();
});
on($('btnBulkBL'),'click',()=>{
  const b = prompt('Building value (blank to keep):', inputBuilding.value||''); if(b==null) return;
  const l = prompt('Level value (blank to keep):', inputLevel.value||''); if(l==null) return;
  const ids = checkedPinIds();
  ids.forEach(idv=>{
    const p=findPin(idv);
    if(p){ if(b!=='') p.building=b; if(l!=='') p.level=l; p.lastEdited=Date.now(); }
  });
  saveProject(project); renderPins(); renderPinsList();
});

// Stage interactions
let addingPin=false;
let dragging=null;

on($('stage'),'pointerdown',(e)=>{
  // only add pin when in add mode and clicking empty area
  if(e.target.classList.contains('pin')) return;
  if(!addingPin) return;
  const {x_pct,y_pct} = toPctCoords(e);
  const p = makePin(x_pct,y_pct);
  currentPage().pins.push(p);
  saveProject(project);
  renderPins(); renderPinsList(); selectPin(p.id);
  addingPin=false; $('btnAddPin').classList.remove('ok');
});

on(pinLayer,'pointerdown',(e)=>{
  const el = e.target.closest('.pin'); if(!el) return;
  selectPin(el.dataset.id);
  dragging = el;
  el.setPointerCapture?.(e.pointerId);
});
on(pinLayer,'pointermove',(e)=>{
  if(!dragging) return; e.preventDefault();
  const pin = findPin(dragging.dataset.id); if(!pin) return;
  const {x_pct,y_pct} = toPctCoords(e);
  dragging.style.left = x_pct+'%';
  dragging.style.top = y_pct+'%';
});
on(pinLayer,'pointerup',(e)=>{
  if(!dragging) return;
  const pin = findPin(dragging.dataset.id);
  if(pin){
    const {x_pct,y_pct} = toPctCoords(e);
    pin.x_pct=x_pct; pin.y_pct=y_pct; pin.lastEdited=Date.now();
    saveProject(project); renderPinsList();
  }
  dragging.releasePointerCapture?.(e.pointerId);
  dragging=null;
});

// ========== Renderers ==========
function renderAll(){
  renderThumbs();
  renderStage();
  renderPins();
  renderPinsList();
  renderProjectLabel();
  toggleField.checked = !!project.settings.fieldMode;
}
function renderProjectLabel(){
  const totalPins = project.pages.reduce((a,p)=>a+(p.pins?.length||0),0);
  projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${totalPins}`;
}
function renderThumbs(){
  thumbsEl.innerHTML='';
  project.pages.forEach(pg=>{
    const d=document.createElement('div');
    d.className='thumb'+(project._pageId===pg.id?' active':'');
    const im=document.createElement('img'); im.src=pg.blobUrl; d.appendChild(im);
    const inp=document.createElement('input'); inp.value=pg.name;
    inp.oninput=()=>{ pg.name=inp.value; pg.updatedAt=Date.now(); saveProject(project); renderProjectLabel(); };
    d.appendChild(inp);
    d.onclick=()=>{ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); renderPinsList(); };
    thumbsEl.appendChild(d);
  });
}
function renderStage(){
  const pg=currentPage();
  if(!pg){ stageImage.removeAttribute('src'); return; }
  stageImage.src = pg.blobUrl;
}
function renderPins(){
  const pg=currentPage(); if(!pg){ pinLayer.innerHTML=''; return; }
  pinLayer.innerHTML='';
  const q = (inputSearch.value||'').toLowerCase();
  const typeFilter = filterType.value;
  const fieldMode = !!project.settings.fieldMode;

  (pg.pins||[]).forEach(p=>{
    if(typeFilter && p.sign_type!==typeFilter) return;
    const line = [p.sign_type,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
    if(q && !line.includes(q)) return;

    const el=document.createElement('div');
    el.className='pin'+(p.id===project._sel?' selected':'');
    el.dataset.id=p.id;
    el.textContent = SHORT[p.sign_type] || (p.sign_type||'').slice(0,3).toUpperCase() || 'PIN';
    el.style.left = p.x_pct+'%';
    el.style.top = p.y_pct+'%';
    el.style.background = SIGN_TYPES[p.sign_type] || '#a3e635';
    el.style.padding = fieldMode? '.28rem .5rem' : '.18rem .35rem';
    el.style.fontSize = fieldMode? '0.9rem' : '0.75rem';
    pinLayer.appendChild(el);
  });

  updateDetails();
}
function renderPinsList(){
  const pg=currentPage(); if(!pg){ pinList.innerHTML=''; return; }
  pinList.innerHTML='';
  const q = (inputSearch.value||'').toLowerCase();
  const typeFilter = filterType.value;

  (pg.pins||[]).forEach(p=>{
    const line = [p.sign_type,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
    if(q && !line.includes(q)) return;
    if(typeFilter && p.sign_type!==typeFilter) return;

    const row=document.createElement('div'); row.className='item';
    const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; row.appendChild(cb);
    const txt=document.createElement('div');
    txt.innerHTML=`<strong>${p.sign_type||'-'}</strong> • ${p.room_number||'-'} ${p.room_name||''} <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
    row.appendChild(txt);
    const go=document.createElement('button'); go.textContent='Go'; go.onclick=()=> selectPin(p.id); row.appendChild(go);
    pinList.appendChild(row);
  });
}
function updateDetails(){
  const p = selectedPin();
  selId.textContent = p ? p.id : 'None';
  fieldType.value = p?.sign_type || '';
  fieldRoomNum.value = p?.room_number || '';
  fieldRoomName.value = p?.room_name || '';
  fieldBuilding.value = p?.building || '';
  fieldLevel.value = p?.level || '';
  fieldNotes.value = p?.notes || '';
}

// ========== Actions / helpers ==========
function startAddPin(){
  addingPin = !addingPin;
  $('btnAddPin').classList.toggle('ok', addingPin);
}
function toPctCoords(e){
  const rect=stageImage.getBoundingClientRect();
  const x=Math.min(Math.max(0,e.clientX-rect.left),rect.width);
  const y=Math.min(Math.max(0,e.clientY-rect.top),rect.height);
  return { x_pct:+(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
}
function currentPage(){
  if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
  return project.pages.find(p=>p.id===project._pageId);
}
function makePin(x_pct,y_pct){
  return {
    id:id(), sign_type:'', room_number:'', room_name:'',
    building: inputBuilding.value||projectContext.building||'',
    level: inputLevel.value||projectContext.level||'',
    x_pct, y_pct, notes:'', lastEdited: Date.now(), photos:[]
  };
}
function findPin(idv){
  for(const pg of project.pages){ const f=(pg.pins||[]).find(p=>p.id===idv); if(f) return f; }
  return null;
}
function checkedPinIds(){
  return [...pinList.querySelectorAll('input[type="checkbox"]:checked')].map(cb=>cb.dataset.id);
}
function selectedPin(){
  return project._sel ? findPin(project._sel) : null;
}
function selectPin(idv){
  project._sel=idv; saveProject(project); renderPins();
  const p = findPin(idv);
  // ensure page sync if needed
  if(p){
    const pg = project.pages.find(pg=> (pg.pins||[]).some(x=>x.id===p.id));
    if(pg && project._pageId!==pg.id){ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); renderPinsList(); }
  }
}

// ========== Storage / pages ==========
function newProject(name){
  return { id:id(), name, createdAt:Date.now(), updatedAt:Date.now(), pages:[], settings:{ fieldMode:false } };
}
function selectProject(pid){
  project = projects.find(p=>p.id===pid); if(!project) return;
  localStorage.setItem('survey:lastOpenProjectId', pid);
  renderAll();
}
function saveProject(p){
  p.updatedAt = Date.now();
  const i=projects.findIndex(x=>x.id===p.id);
  if(i>=0) projects[i]=p; else projects.push(p);
  localStorage.setItem('survey:projects', JSON.stringify(projects));
  // rehydrate from storage to keep array refs fresh
  projects = loadProjects();
}
function loadProjects(){
  try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); }catch{ return []; }
}

function addImagePage(url, name){
  const pg={ id:id(), name:name||'Image', kind:'image', blobUrl:url, pins:[], measurements:[], updatedAt:Date.now() };
  project.pages.push(pg);
  if(!project._pageId) project._pageId = pg.id;
  saveProject(project);
}
async function addPdfPages(file){
  const url=URL.createObjectURL(file);
  const pdf=await pdfjsLib.getDocument(url).promise;
  for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i);
    const viewport=page.getViewport({scale:1.5});
    const canvas=document.createElement('canvas');
    canvas.width=viewport.width; canvas.height=viewport.height;
    const ctx=canvas.getContext('2d');
    await page.render({canvasContext:ctx, viewport}).promise;
    const data=canvas.toDataURL('image/png');
    const pg={ id:id(), name:`${file.name.replace(/\.[^.]+$/,'')} · p${i}`, kind:'pdf', pdfPage:i, blobUrl:data, pins:[], measurements:[], updatedAt:Date.now() };
    project.pages.push(pg);
  }
  if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
  saveProject(project);
}
