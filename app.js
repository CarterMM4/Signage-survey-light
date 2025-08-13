// app.js
document.addEventListener('DOMContentLoaded', () => {
  /***********************
   * Helpers & constants *
   ***********************/
  const $ = (id) => document.getElementById(id);
  const on = (el, evt, fn) => el && el.addEventListener(evt, fn);

  // Colors for suggested types
  const SIGN_TYPES = {
    'FOH':'#0ea5e9','BOH':'#64748b','Elevator':'#8b5cf6','UNIT ID':'#22c55e',
    'Stair - Ingress':'#22c55e','Stair - Egress':'#ef4444','Ingress':'#22c55e',
    'Egress':'#ef4444','Hall Direct':'#f59e0b','Callbox':'#06b6d4','Evac':'#84cc16',
    'Exit':'#ef4444','Restroom':'#10b981'
  };
  const GREEN_DEFAULT = '#22c55e';

  /***********************
   * Grab DOM references *
   ***********************/
  // Toolbar + lists
  const projectPicker = $('projectPicker');
  const projectsList = $('projectsList');
  const projectLabel = $('projectLabel');

  // Left pages
  const thumbsEl = $('thumbs');

  // Stage
  const stage = $('stage');
  const stageImage = $('stageImage');
  const pinLayer = $('pinLayer');
  const measureSvg = $('measureSvg');

  // Inputs/buttons
  const inputUpload = $('inputUpload');
  const inputImportCSV = $('inputImportCSV');
  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');
  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleField = $('toggleField');

  // Pin details
  const fieldType = $('fieldType');     // input with datalist
  const fieldLabel = $('fieldLabel');
  const fieldRoomNum = $('fieldRoomNum');
  const fieldRoomName = $('fieldRoomName');
  const fieldBuilding = $('fieldBuilding');
  const fieldLevel = $('fieldLevel');
  const fieldNotes = $('fieldNotes');
  const selId = $('selId');
  const posLabel = $('posLabel');
  const warnsEl = $('warns');
  const pinList = $('pinList');

  // Photo modal
  const photoModal = $('photoModal');
  const photoImg = $('photoImg');
  const photoMeasureSvg = $('photoMeasureSvg');
  const photoPinId = $('photoPinId');
  const photoName = $('photoName');
  const photoMeaCount = $('photoMeaCount');

  /************************
   * App state / Undo/redo *
   ************************/
  let projects = loadProjects();
  let currentProjectId = localStorage.getItem('survey:lastOpenProjectId') || null;
  let project = projects.find(p => p.id === currentProjectId) || null;

  const UNDO = [];
  const REDO = [];
  const MAX_UNDO = 60;

  // Ephemeral defaults when adding
  const projectContext = { building:'', level:'' };

  // Stage interaction state
  let addingPin = false;
  let draggingEl = null;
  let measureMode = false;
  let calibFirst = null;
  let measureFirst = null;
  let calibAwait = null; // 'main' | 'photo'
  let fieldMode = false;

  // Photo modal state
  const photoState = { pin:null, idx:0, measure:false, calib:null };
  let photoMeasureFirst = null;

  /*******************
   * Small utilities *
   *******************/
  function id(){ return Math.random().toString(36).slice(2,10); }
  function fix(n){ return typeof n==='number'? +Number(n||0).toFixed(3):''; }
  function pctLabel(x,y){ return `${(x||0).toFixed(2)}%, ${(y||0).toFixed(2)}%`; }
  function dist(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return Math.sqrt(dx*dx+dy*dy); }
  function toPctCoords(e){
    const rect = stageImage.getBoundingClientRect();
    const x = Math.min(Math.max(0, e.clientX - rect.left), rect.width);
    const y = Math.min(Math.max(0, e.clientY - rect.top), rect.height);
    return { x_pct: +(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
  }
  function toLocal(e){ const r=stageImage.getBoundingClientRect(); return { x:e.clientX-r.left, y:e.clientY-r.top }; }
  function toScreen(pt){ const r=stageImage.getBoundingClientRect(); return { x:r.left+pt.x, y:r.top+pt.y }; }
  function csvEscape(v){ const s=(v==null? '': String(v)); return /["\n,]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s; }

  function parseCsvLine(line){
    const out=[]; let cur=''; let i=0; let inQ=false;
    while(i<line.length){
      const ch=line[i++];
      if(inQ){
        if(ch==='"'){ if(line[i]==='"'){ cur+='"'; i++; } else { inQ=false; } }
        else cur+=ch;
      }else{
        if(ch===','){ out.push(cur); cur=''; }
        else if(ch==='"'){ inQ=true; }
        else cur+=ch;
      }
    }
    out.push(cur); return out;
  }
  function fileToDataURL(file){
    return new Promise((res,rej)=>{ const fr=new FileReader(); fr.onload=()=>res(fr.result); fr.onerror=rej; fr.readAsDataURL(file); });
  }
  function dataURLtoArrayBuffer(dataURL){
    const bstr=atob(dataURL.split(',')[1]); const len=bstr.length; const buf=new Uint8Array(len);
    for(let i=0;i<len;i++) buf[i]=bstr.charCodeAt(i); return buf;
  }
  function dataURLtoBlob(dataURL){
    const arr=dataURL.split(','); const match=(arr[0].match(/:(.*?);/)||[]); const mime=match[1]||'image/png';
    return new Blob([dataURLtoArrayBuffer(dataURL)],{type:mime});
  }
  function downloadText(name, text){ downloadFile(name, new Blob([text],{type:'text/plain'})); }
  function downloadFile(name, blob){
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 900);
  }

  /********************
   * Data model utils *
   ********************/
  function newProject(name){
    return { id:id(), name, createdAt:Date.now(), updatedAt:Date.now(), pages:[], settings:{ fieldMode:false } };
  }
  function currentPage(){
    if(!project) return null;
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId) || null;
  }
  function makePin(x_pct,y_pct){
    return {
      id:id(),
      sign_type: fieldType.value || '',
      label: fieldLabel.value || '',
      room_number:'', room_name:'',
      building: inputBuilding.value || projectContext.building || '',
      level: inputLevel.value || projectContext.level || '',
      x_pct, y_pct, notes:'', photos:[], lastEdited:Date.now()
    };
  }
  function findPin(idv){
    for(const pg of project.pages){ const f=(pg.pins||[]).find(p=>p.id===idv); if(f) return f; }
    return null;
  }
  function checkedPinIds(){ return [...pinList.querySelectorAll('input[type="checkbox"]:checked')].map(cb=>cb.dataset.id); }

  function snapshot(){ return JSON.stringify(project); }
  function loadSnapshot(s){
    project=JSON.parse(s);
    const i=projects.findIndex(p=>p.id===project.id);
    if(i>=0) projects[i]=project; else projects.push(project);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    localStorage.setItem('survey:lastOpenProjectId', project.id);
  }
  function commit(){ if(!project) return; UNDO.push(snapshot()); if(UNDO.length>MAX_UNDO) UNDO.shift(); REDO.length=0; }
  function undo(){ if(!UNDO.length) return; REDO.push(snapshot()); const s=UNDO.pop(); loadSnapshot(s); renderAll(); }
  function redo(){ if(!REDO.length) return; UNDO.push(snapshot()); const s=REDO.pop(); loadSnapshot(s); renderAll(); }

  function saveProject(p){
    p.updatedAt = Date.now();
    const i=projects.findIndex(x=>x.id===p.id);
    if(i>=0) projects[i]=p; else projects.push(p);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    projects = loadProjects(); // refresh canonical
    fillProjectPicker();
    renderProjectsList();
  }
  function loadProjects(){
    try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); }catch{ return []; }
  }
  function selectProject(pid){
    project = projects.find(p=>p.id===pid) || null;
    if(!project){
      alert('Project not found.');
      return;
    }
    localStorage.setItem('survey:lastOpenProjectId', pid);
    renderAll();
  }

  function addImagePage(url, name){
    const pg={ id:id(), name:name||'Image', kind:'image', blobUrl:url, pins:[], measurements:[], updatedAt:Date.now() };
    project.pages.push(pg);
    if(!project._pageId) project._pageId=pg.id;
    saveProject(project);
  }
  async function addPdfPages(file){
    const url=URL.createObjectURL(file);
    const pdf=await pdfjsLib.getDocument(url).promise;
    for(let i=1;i<=pdf.numPages;i++){
      const page=await pdf.getPage(i);
      const viewport=page.getViewport({scale:2});
      const canvas=document.createElement('canvas');
      canvas.width=viewport.width; canvas.height=viewport.height;
      const ctx=canvas.getContext('2d');
      await page.render({canvasContext:ctx, viewport}).promise;
      const data=canvas.toDataURL('image/png');
      const pg={ id:id(), name:`${file.name.replace(/\.[^.]+$/,'')} · p${i}`, kind:'pdf', pdfPage:i, blobUrl=data, pins:[], measurements:[], updatedAt:Date.now() };
      project.pages.push(pg);
    }
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    saveProject(project);
  }

  /****************
   * Render funcs *
   ****************/
  function renderAll(){
    fillTypeDatalist();
    fillProjectPicker();
    renderProjectsList();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    renderProjectLabel();
    drawMeasurements();
    toggleField.checked = !!project.settings.fieldMode;
  }

  function fillTypeDatalist(){
    const dl = $('typeList');
    if(!dl.dataset.filled){
      Object.keys(SIGN_TYPES).forEach(t=>{
        const o=document.createElement('option'); o.value=t; dl.appendChild(o);
        const o2=document.createElement('option'); o2.value=t; filterType.appendChild(o2);
      });
      dl.dataset.filled = '1';
    }
  }

  function fillProjectPicker(){
    projectPicker.innerHTML='';
    projects.forEach(p=>{
      const o=document.createElement('option');
      o.value=p.id; o.textContent=p.name;
      if(project && p.id===project.id) o.selected=true;
      projectPicker.appendChild(o);
    });
  }

  function renderProjectsList(){
    projectsList.innerHTML='';
    projects.forEach(p=>{
      const row=document.createElement('div'); row.className='project-row'+(project && p.id===project.id?' active':'');
      const left=document.createElement('div');
      left.innerHTML=`<div>${p.name}</div><small>${new Date(p.updatedAt).toLocaleString()}</small>`;
      const btn=document.createElement('button'); btn.textContent='Open';
      btn.onclick=()=> selectProject(p.id);
      row.appendChild(left); row.appendChild(btn);
      projectsList.appendChild(row);
    });
  }

  function renderProjectLabel(){
    if(!project){ projectLabel.textContent='—'; return; }
    const pinCount = project.pages.reduce((a,pg)=> a + ((pg.pins && pg.pins.length) || 0), 0);
    projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${pinCount}`;
  }

  function renderThumbs(){
    thumbsEl.innerHTML='';
    if(!project) return;
    project.pages.forEach(pg=>{
      const d=document.createElement('div'); d.className='thumb'+(project._pageId===pg.id?' active':'');
      const im=document.createElement('img'); im.src=pg.blobUrl; d.appendChild(im);
      const inp=document.createElement('input'); inp.value=pg.name;
      inp.oninput=()=>{ pg.name=inp.value; pg.updatedAt=Date.now(); saveProject(project); renderProjectLabel(); };
      d.appendChild(inp);
      d.onclick=()=>{ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); drawMeasurements(); };
      thumbsEl.appendChild(d);
    });
  }

  function renderStage(){
    const pg=currentPage();
    if(!pg){ stageImage.removeAttribute('src'); return; }
    stageImage.src=pg.blobUrl;
  }

  function pinText(p){
    if(p.label && p.label.trim()) return p.label.trim().slice(0,12);
    if(p.sign_type && SIGN_TYPES[p.sign_type]) return p.sign_type;
    if(p.sign_type) return p.sign_type.slice(0,10);
    return 'PIN';
  }

  function renderPins(){
    pinLayer.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    fieldMode=!!project.settings.fieldMode;

    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type!==typeFilter) return;

      const line=[p.sign_type,p.label,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el=document.createElement('div');
      el.className='pin'+(p.id===project._sel?' selected':'');
      el.dataset.id=p.id;
      el.textContent=pinText(p);
      el.style.left=p.x_pct+'%';
      el.style.top=p.y_pct+'%';
      el.style.background=SIGN_TYPES[p.sign_type] || GREEN_DEFAULT;
      el.style.padding = fieldMode? '.32rem .6rem' : '.22rem .45rem';
      el.style.fontSize= fieldMode? '0.95rem' : '0.8rem';
      pinLayer.appendChild(el);

      // start drag
      el.addEventListener('pointerdown', (e)=>{
        e.preventDefault();
        selectPin(p.id);
        draggingEl = el;
        el.setPointerCapture?.(e.pointerId);
      });
    });

    updatePinDetails();
  }

  function renderPinsList(){
    pinList.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      const line=[p.sign_type,p.label,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;
      if(typeFilter && p.sign_type!==typeFilter) return;

      const row=document.createElement('div'); row.className='item';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; row.appendChild(cb);
      const txt=document.createElement('div');
      txt.innerHTML=`<strong>${p.label||p.sign_type||'-'}</strong> • ${p.room_number||'-'} ${p.room_name||''} <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
      row.appendChild(txt);
      const go=document.createElement('button'); go.textContent='Go'; go.onclick=()=>{ selectPin(p.id); }; row.appendChild(go);
      pinList.appendChild(row);
    });
  }

  function selectedPin(){
    const idv = project? project._sel : null;
    if(!idv) return null;
    return findPin(idv);
  }

  function updatePinDetails(){
    const p=selectedPin();
    selId.textContent=p? p.id : 'None';
    fieldType.value=p?.sign_type||'';
    fieldLabel.value=p?.label||'';
    fieldRoomNum.value=p?.room_number||'';
    fieldRoomName.value=p?.room_name||'';
    fieldBuilding.value=p?.building||'';
    fieldLevel.value=p?.level||'';
    fieldNotes.value=p?.notes||'';
    posLabel.textContent=p? pctLabel(p.x_pct,p.y_pct) : '—';
    renderWarnings();
  }

  function renderWarnings(){
    warnsEl.innerHTML='';
    const p=selectedPin(); if(!p) return;
    const list=[];
    const rn=(p.room_name||'').toUpperCase();
    if(['Callbox','Evac','Hall Direct'].includes(p.sign_type) && rn!=='ELEV. LOBBY'){ list.push('Tip: Room name usually “ELEV. LOBBY”.'); }
    if((/ELECTRICAL|DATA/).test(rn) && p.sign_type!=='BOH'){ list.push('Recommended BOH for ELECTRICAL/DATA rooms.'); }
    list.forEach(s=>{ const t=document.createElement('span'); t.className='tag warn'; t.textContent=s; warnsEl.appendChild(t); });
  }

  function selectPin(idv){
    if(!project) return;
    project._sel=idv; saveProject(project); renderPins();
    const el=[...pinLayer.children].find(x=>x.dataset.id===idv);
    if(el){ el.scrollIntoView({block:'center', inline:'center', behavior:'smooth'}); }
    const p=findPin(idv); if(p){
      const pgId=project.pages.find(pg=>pg.pins.includes(p))?.id;
      if(pgId && project._pageId!==pgId){ project._pageId=pgId; saveProject(project); renderStage(); renderPins(); drawMeasurements(); }
    }
  }

  /********************
   * Measuring (main) *
   ********************/
  function startCalibration(scope){
    calibFirst=null; measureFirst=null;
    alert('Calibration: click two points on the image, then enter real distance.');
    measureMode=false; $('btnMeasureToggle').textContent='Measure: OFF';
    calibAwait=scope; // 'main' expected here
  }

  function toggleMeasuring(scope){
    measureMode=!measureMode;
    if(scope==='main') $('btnMeasureToggle').textContent = 'Measure: '+(measureMode?'ON':'OFF');
  }

  function resetMeasurements(scope){
    if(scope==='main'){
      const pg=currentPage(); if(!pg) return;
      pg.measurements=[]; drawMeasurements(); saveProject(project);
    }else{
      const ph = photoState.pin?.photos[photoState.idx];
      if(ph){ ph.measurements=[]; drawPhotoMeasurements(); saveProject(project); }
    }
  }

  on(stage,'click',(e)=>{
    const page=currentPage(); if(!page || !stageImage.src) return;
    const pt=toLocal(e);

    if(calibAwait==='main'){
      if(!calibFirst){ calibFirst=pt; return; }
      const px = dist(calibFirst, pt);
      const val = prompt('Enter real distance (e.g. feet):','10');
      const ft = val? parseFloat(val) : NaN;
      if(!isNaN(ft) && ft>0){ page.scalePxPerFt = px/ft; alert('Calibrated: '+(px/ft).toFixed(2)+' px/ft'); }
      calibFirst=null; calibAwait=null;
      return;
    }

    if(measureMode){
      if(!measureFirst){ measureFirst=pt; return; }
      const m={id:id(), kind:'main', points:[measureFirst, pt]};
      // Optional per-line manual entry
      const manual = prompt('Enter real length for THIS line (optional, e.g. 6.5):','');
      if(manual && !isNaN(parseFloat(manual))) m.userFeet = parseFloat(manual);
      if(page.scalePxPerFt) m.autoFeet = dist(measureFirst, pt)/page.scalePxPerFt;

      page.measurements = page.measurements||[]; page.measurements.push(m);
      measureFirst=null; drawMeasurements(); saveProject(project);
    }
  });

  function drawMeasurements(){
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    const page=currentPage(); if(!page||!page.measurements) return;

    page.measurements.forEach(m=>{
      const a=toScreen(m.points[0]); const b=toScreen(m.points[1]);
      const dx=Math.abs(a.x-b.x), dy=Math.abs(a.y-b.y);
      const color=(dx>dy)? '#ef4444' : (dy>dx)? '#3b82f6' : '#f59e0b';

      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y);
      line.setAttribute('x2',b.x); line.setAttribute('y2',b.y);
      line.setAttribute('stroke',color); line.setAttribute('stroke-width','3');
      measureSvg.appendChild(line);

      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      const label = (typeof m.userFeet==='number') ? m.userFeet.toFixed(2)+' ft'
                   : (page.scalePxPerFt ? (dist(m.points[0], m.points[1])/page.scalePxPerFt).toFixed(2)+' ft'
                   : dist(m.points[0], m.points[1]).toFixed(0)+' px');
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=label;
      measureSvg.appendChild(text);
    });
  }

  /*************************
   * Photo modal & measure *
   *************************/
  function openPhotoModal(){
    const p=selectedPin(); if(!p) return alert('Select a pin first.');
    if(!p.photos.length) return alert('No photos attached.');
    photoState.pin=p; photoState.idx=0; showPhoto(); photoModal.style.display='flex';
  }
  function showPhoto(){
    const ph=photoState.pin.photos[photoState.idx];
    photoImg.src=ph.dataUrl; photoPinId.textContent=photoState.pin.id; photoName.textContent=ph.name;
    photoMeaCount.textContent=(ph.measurements?.length||0); drawPhotoMeasurements();
  }
  function closePhotoModal(){ photoModal.style.display='none'; }

  function drawPhotoMeasurements(){
    while(photoMeasureSvg.firstChild) photoMeasureSvg.removeChild(photoMeasureSvg.firstChild);
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph||!ph.measurements) return;

    ph.measurements.forEach(m=>{
      const a=m.points[0], b=m.points[1];
      const dx=Math.abs(a.x-b.x), dy=Math.abs(a.y-b.y);
      const color=(dx>dy)? '#ef4444' : (dy>dx)? '#3b82f6' : '#f59e0b';

      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y);
      line.setAttribute('x2',b.x); line.setAttribute('y2',b.y);
      line.setAttribute('stroke',color); line.setAttribute('stroke-width','3');
      photoMeasureSvg.appendChild(line);

      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      const label = (typeof m.userFeet==='number') ? m.userFeet.toFixed(2)+' ft'
                   : (ph.scalePxPerFt ? (dist(a,b)/ph.scalePxPerFt).toFixed(2)+' ft'
                   : dist(a,b).toFixed(0)+' px');
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=label;
      photoMeasureSvg.appendChild(text);
    });
  }

  on($('btnPhotoClose'),'click',closePhotoModal);
  on($('btnPhotoMeasure'),'click',()=>{
    photoState.measure=!photoState.measure;
    $('btnPhotoMeasure').textContent='Measure: '+(photoState.measure?'ON':'OFF');
  });
  on($('btnPhotoCalib'),'click',()=>{
    photoState.calib=null;
    alert('Photo calibration: click two points on the image, then enter real distance.');
  });
  on($('btnPhotoPrev'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on($('btnPhotoNext'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on($('btnPhotoDelete'),'click',()=>{ if(!photoState.pin) return; if(!confirm('Delete this photo?')) return; commit(); photoState.pin.photos.splice(photoState.idx,1); photoState.idx=Math.max(0,photoState.idx-1); saveProject(project); if(photoState.pin.photos.length===0){ closePhotoModal(); } else { showPhoto(); } });
  on($('btnPhotoDownload'),'click',()=>{ const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return; downloadFile(ph.name||'photo.png', dataURLtoBlob(ph.dataUrl)); });

  on(photoMeasureSvg,'click',(e)=>{
    const ph = photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };

    // Calibration clicks on photo
    if(photoState.calib===null && !photoState.measure){ photoState.calib = pt; return; }
    if(photoState.calib && !photoState.measure){
      const px = dist(photoState.calib, pt);
      const val = prompt('Enter real distance (e.g. feet):','10');
      const ft = val? parseFloat(val) : NaN;
      if(!isNaN(ft) && ft>0){ ph.scalePxPerFt = px/ft; alert('Photo calibrated: '+(px/ft).toFixed(2)+' px/ft'); }
      photoState.calib=null; saveProject(project); drawPhotoMeasurements(); return;
    }
  });

  on(photoImg,'click',(e)=>{
    if(!photoState.measure) return; const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect(); const pt={x:e.clientX-rect.left,y:e.clientY-rect.top};
    if(!photoMeasureFirst){ photoMeasureFirst=pt; return; }
    const m={id:id(), kind:'photo', points:[photoMeasureFirst, pt]};
    const manual = prompt('Enter real length for THIS line (optional, e.g. 6.5):','');
    if(manual && !isNaN(parseFloat(manual))) m.userFeet = parseFloat(manual);
    if(ph.scalePxPerFt){ m.autoFeet=dist(photoMeasureFirst,pt)/ph.scalePxPerFt; }
    ph.measurements=ph.measurements||[]; ph.measurements.push(m);
    photoMeasureFirst=null; drawPhotoMeasurements(); photoMeaCount.textContent=ph.measurements.length; saveProject(project);
  });

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    project.pages.forEach(pg=> (pg.pins||[]).forEach(p=> rows.push({
      id:p.id, sign_type:p.sign_type||'', label:p.label||'', room_number:p.room_number||'', room_name:p.room_name||'',
      building:p.building||'', level:p.level||'', x_pct:fix(p.x_pct), y_pct:fix(p.y_pct),
      notes:p.notes||'', page_name:pg.name, last_edited: new Date(p.lastEdited||project.updatedAt).toISOString()
    })));
    rows.sort((a,b)=> (a.building||'').localeCompare(b.building||'')
      || (a.level||'').localeCompare(b.level||'')
      || a.room_number.localeCompare(b.room_number, undefined, {numeric:true,sensitivity:'base'}) );
    return rows;
  }

  function exportCSV(){
    const rows = toRows();
    const csv=[Object.keys(rows[0]||{}).join(','), ...rows.map(r=>Object.values(r).map(v=>csvEscape(v)).join(','))].join('\n');
    downloadText((project.name||'project').replace(/\W+/g,'_')+"_signage.csv", csv);
  }
  function exportXLSX(){
    const rows = toRows(); const wb = XLSX.utils.book_new();
    const info=[[ 'Project', project.name ], [ 'Exported', new Date().toLocaleString() ], [ 'Total Pins', rows.length ], [ 'Total Pages', project.pages.length ]];
    const counts = {}; rows.forEach(r=> counts[r.sign_type]=(counts[r.sign_type]||0)+1 ); info.push([]); info.push(['Breakdown']); Object.entries(counts).forEach(([k,v])=> info.push([k,v]));
    const wsInfo = XLSX.utils.aoa_to_sheet(info);
    const ws = XLSX.utils.json_to_sheet(rows, {header: Object.keys(rows[0]||{})});
    XLSX.utils.book_append_sheet(wb, wsInfo, 'Project Info'); XLSX.utils.book_append_sheet(wb, ws, 'Pins');
    const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    downloadFile((project.name||'project').replace(/\W+/g,'_')+'_signage.xlsx', new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  }
  async function exportZIP(){
    const rows=toRows(); const zip=new JSZip();
    zip.file('signage.csv', [Object.keys(rows[0]||{}).join(','), ...rows.map(r=>Object.values(r).map(v=>csvEscape(v)).join(','))].join('\n'));
    project.pages.forEach(pg=> (pg.pins||[]).forEach(pin=> (pin.photos||[]).forEach((ph, idx)=>{ const folder=zip.folder(`photos/${pin.id}`); folder.file(ph.name||`photo_${idx+1}.png`, dataURLtoArrayBuffer(ph.dataUrl)); }) ));
    const blob = await zip.generateAsync({type:'blob'});
    downloadFile((project.name||'project').replace(/\W+/g,'_')+'_export.zip', blob);
  }
  function importCSV(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2) return;
    const hdr=lines[0].split(',').map(h=>h.trim()); commit();
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]); const row={}; hdr.forEach((h,idx)=> row[h]=cells[idx]||'');
      const x = parseFloat(row.x_pct), y = parseFloat(row.y_pct);
      const p = makePin(isNaN(x)?50:x, isNaN(y)?50:y);
      p.sign_type=row.sign_type||''; p.label=row.label||''; p.room_number=row.room_number||''; p.room_name=row.room_name||'';
      p.building=row.building||''; p.level=row.level||''; p.notes=row.notes||'';
      currentPage().pins.push(p);
    }
    saveProject(project); renderPins(); renderPinsList(); alert('Imported rows into current page.');
  }

  /***********************
   * OCR (optional use)  *
   ***********************/
  async function ocrCurrentView(){
    const img=stageImage; if(!img || !img.src){ alert('No page image.'); return; }
    const canvas = document.createElement('canvas');
    const w=img.naturalWidth||img.width, h=img.naturalHeight||img.height; canvas.width=w; canvas.height=h;
    const ctx=canvas.getContext('2d'); ctx.drawImage(img,0,0);
    const { data: { text } } = await Tesseract.recognize(canvas, 'eng');
    const out = (text||'').trim();
    if(!out){ alert('No text recognized.'); return; }
    const head = out.slice(0,200);
    try { await navigator.clipboard.writeText(out); } catch {}
    inputSearch.value=head; renderPinsList(); alert('OCR done. First 200 chars placed in search. Full text copied.');
  }

  /**********************
   * Event wiring (UI)  *
   **********************/
  // Project picker and list
  on(projectPicker,'change', (e)=> selectProject(e.target.value));
  on($('btnNew'),'click',()=>{
    const name = prompt('New project name?','New Project');
    if(!name) return;
    const p=newProject(name); projects.push(p); saveProject(p); selectProject(p.id);
  });
  on($('btnSaveAs'),'click',()=>{
    if(!project) return;
    const name=prompt('Duplicate as name:', project.name+' (copy)');
    if(!name) return;
    commit();
    const copy=JSON.parse(JSON.stringify(project));
    copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now();
    projects.push(copy); localStorage.setItem('survey:projects', JSON.stringify(projects));
    selectProject(copy.id);
  });
  on($('btnRename'),'click',()=>{
    if(!project) return;
    const name=prompt('Rename project:', project.name);
    if(!name) return;
    commit(); project.name=name; saveProject(project); renderProjectLabel(); fillProjectPicker(); renderProjectsList();
  });

  // Upload pages (images/pdfs)
  on($('btnUpload'),'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){ const url=URL.createObjectURL(f); const name=f.name.replace(/\.[^.]+$/,''); addImagePage(url,name); }
    }
    renderAll();
    inputUpload.value='';
  });

  // Exports / imports
  on($('btnExportCSV'),'click',exportCSV);
  on($('btnExportXLSX'),'click',exportXLSX);
  on($('btnExportZIP'),'click',exportZIP);
  on($('btnImportCSV'),'click',()=> inputImportCSV.click());
  on(inputImportCSV,'change', async (e)=>{
    const f=e.target.files?.[0]; if(!f) return;
    const text=await f.text(); importCSV(text); inputImportCSV.value='';
  });

  // OCR & measure controls
  on($('btnOCR'),'click',ocrCurrentView);
  on($('btnCalibrate'),'click',()=> startCalibration('main'));
  on($('btnMeasureToggle'),'click',()=> toggleMeasuring('main'));
  on($('btnMeasureReset'),'click',()=> resetMeasurements('main'));

  // Undo/Redo
  on($('btnUndo'),'click',undo);
  on($('btnRedo'),'click',redo);

  // Pin bulk actions & search/filter
  on($('btnClearPins'),'click',()=>{
    if(!currentPage()) return;
    if(!confirm('Clear ALL pins on this page?')) return;
    commit(); currentPage().pins=[]; saveProject(project); renderAll();
  });
  on($('btnAddPin'),'click',()=>{
    addingPin = !addingPin;
    $('btnAddPin').classList.toggle('ok', addingPin);
  });
  on(inputSearch,'input',renderPinsList);
  on(filterType,'change',()=>{ renderPins(); renderPinsList(); });
  on(toggleField,'change',()=>{
    project.settings.fieldMode = !!toggleField.checked; saveProject(project); renderPins();
  });
  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; });

  // Stage: place new pin
  on(stage,'pointerdown',(e)=>{
    if(e.target.classList.contains('pin')) return; // handled by pin itself
    if(!addingPin) return;
    const {x_pct,y_pct} = toPctCoords(e);
    commit();
    const p = makePin(x_pct,y_pct);
    currentPage().pins.push(p);
    saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
    addingPin=false; $('btnAddPin').classList.remove('ok');
  });

  // Stage: drag pins
  on(pinLayer,'pointermove',(e)=>{
    if(!draggingEl) return;
    e.preventDefault();
    const pin = findPin(draggingEl.dataset.id); if(!pin) return;
    const {x_pct,y_pct} = toPctCoords(e);
    draggingEl.style.left = x_pct+'%';
    draggingEl.style.top = y_pct+'%';
    posLabel.textContent = pctLabel(x_pct,y_pct);
  });
  on(pinLayer,'pointerup',(e)=>{
    if(!draggingEl) return;
    const pin = findPin(draggingEl.dataset.id);
    if(pin){
      commit();
      const {x_pct,y_pct} = toPctCoords(e);
      pin.x_pct=x_pct; pin.y_pct=y_pct; pin.lastEdited=Date.now();
      saveProject(project); renderPinsList();
    }
    draggingEl.releasePointerCapture?.(e.pointerId);
    draggingEl=null;
  });

  // Pin detail inputs (free text)
  [fieldType, fieldLabel, fieldRoomNum, fieldRoomName, fieldBuilding, fieldLevel, fieldNotes].forEach(el=> on(el,'input',()=>{
    const pin = selectedPin(); if(!pin) return; commit();
    if(el===fieldType) pin.sign_type = el.value || '';
    if(el===fieldLabel) pin.label = el.value || '';
    if(el===fieldRoomNum) pin.room_number = el.value || '';
    if(el===fieldRoomName) pin.room_name = el.value || '';
    if(el===fieldBuilding) pin.building = el.value || '';
    if(el===fieldLevel) pin.level = el.value || '';
    if(el===fieldNotes) pin.notes = el.value || '';
    pin.lastEdited=Date.now(); saveProject(project); renderPins(); renderWarnings(); renderPinsList();
  }));

  // Photos on pin
  on($('btnAddPhoto'),'click',()=> $('inputPhoto').click());
  on($('inputPhoto'),'change', async (e)=>{
    const pin = selectedPin(); if(!pin) return alert('Select a pin first.');
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){ const url=await fileToDataURL(f); pin.photos.push({name:f.name,dataUrl:url,measurements:[]}); }
    pin.lastEdited=Date.now(); saveProject(project); renderPinsList(); alert('Photo(s) added.');
    $('inputPhoto').value='';
  });
  on($('btnOpenPhoto'),'click',openPhotoModal);
  on($('btnDuplicate'),'click',()=>{
    const pin=selectedPin(); if(!pin) return;
    commit(); const p=JSON.parse(JSON.stringify(pin));
    p.id=id(); p.x_pct=Math.min(100,p.x_pct+2); p.y_pct=Math.min(100,p.y_pct+2);
    p.lastEdited=Date.now(); currentPage().pins.push(p);
    saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
  });
  on($('btnDelete'),'click',()=>{
    const pin=selectedPin(); if(!pin) return;
    if(!confirm('Delete selected pin?')) return;
    commit();
    currentPage().pins=currentPage().pins.filter(x=>x.id!==pin.id);
    saveProject(project); renderPins(); renderPinsList(); project._sel=null; updatePinDetails();
  });

  // Keyboard helpers
  window.addEventListener('keydown',(e)=>{
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); redo(); }
    if(e.key==='Delete'){ const p=selectedPin(); if(p){ commit(); currentPage().pins=currentPage().pins.filter(x=>x.id!==p.id); saveProject(project); renderPins(); renderPinsList(); project._sel=null; updatePinDetails(); } }
    if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)){
      const p=selectedPin(); if(!p) return; e.preventDefault();
      commit();
      const delta = project.settings.fieldMode? 0.5 : 0.2;
      if(e.key==='ArrowUp') p.y_pct=Math.max(0,p.y_pct-delta);
      if(e.key==='ArrowDown') p.y_pct=Math.min(100,p.y_pct+delta);
      if(e.key==='ArrowLeft') p.x_pct=Math.max(0,p.x_pct-delta);
      if(e.key==='ArrowRight') p.x_pct=Math.min(100,p.x_pct+delta);
      p.lastEdited=Date.now(); saveProject(project); renderPins(); updatePinDetails();
    }
  });

  // Resize => keep SVG overlay sized
  const ro=new ResizeObserver(()=>{
    measureSvg.setAttribute('width', stage.clientWidth);
    measureSvg.setAttribute('height', stage.clientHeight);
  });
  ro.observe(stage);

  /****************
   * First boot   *
   ****************/
  if(!project){
    project = newProject('Untitled Project');
    projects.push(project);
    saveProject(project);
  }
  selectProject(project.id); // triggers renderAll()
});
