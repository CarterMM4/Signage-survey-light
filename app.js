// app.js
document.addEventListener('DOMContentLoaded', () => {
  /***********************
   * Helpers & constants *
   ***********************/
  const $ = (id) => document.getElementById(id);
  const on = (el, evt, fn) => el && el.addEventListener(evt, fn);
  const toast = $('toast');
  let toastTimer = null;
  function savedToast(){
    if(!toast) return;
    toast.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(()=>toast.classList.remove('show'), 1200);
  }

  const SIGN_TYPES = {
    'FOH':'#0ea5e9','BOH':'#64748b','Elevator':'#8b5cf6','UNIT ID':'#22c55e',
    'Stair - Ingress':'#22c55e','Stair - Egress':'#ef4444','Ingress':'#22c55e',
    'Egress':'#ef4444','Hall Direct':'#f59e0b','Callbox':'#06b6d4','Evac':'#84cc16',
    'Exit':'#ef4444','Restroom':'#10b981'
  };
  const SHORT = {
    'Elevator':'ELV','Hall Direct':'HD','Callbox':'CB','Evac':'EV','Ingress':'ING',
    'Egress':'EGR','Exit':'EXIT','Restroom':'WC','UNIT ID':'UNIT','FOH':'FOH','BOH':'BOH',
    'Stair - Ingress':'ING','Stair - Egress':'EGR'
  };

  /***********************
   * Grab DOM references *
   ***********************/
  // Views
  const homeView = $('homeView');
  const projectsView = $('projectsView');
  const editorView = $('editorView');

  // Home / Projects nav
  on($('homeNew'), 'click', ()=> createAndOpenProject());
  on($('homeOpen'), 'click', ()=> { renderProjectsGrid(); showView('projectsView'); });
  on($('projBack'), 'click', ()=> showView('homeView'));
  on($('projNew'), 'click', ()=> createAndOpenProject());

  on($('btnHome'), 'click', ()=> showView('homeView'));
  on($('btnProjects'), 'click', ()=> { renderProjectsGrid(); showView('projectsView'); });

  // Editor: Left column
  const thumbsEl = $('thumbs');

  // Editor: Stage
  const stage = $('stage');
  const stageImage = $('stageImage');
  const pinLayer = $('pinLayer');
  const measureSvg = $('measureSvg');

  // Editor: Toolbar bits
  const projectLabel = $('projectLabel');
  const inputUpload = $('inputUpload');
  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');
  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleField = $('toggleField');

  // Editor: Right panel fields
  const fieldType = $('fieldType');
  const fieldRoomNum = $('fieldRoomNum');
  const fieldRoomName = $('fieldRoomName'); // custom name
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
  let project = currentProjectId ? projects.find(p=>p.id===currentProjectId) : null;

  const UNDO = [];
  const REDO = [];
  const MAX_UNDO = 50;

  // Ephemeral defaults (for adding)
  const projectContext = { building:'', level:'' };

  // Stage interaction state
  let addingPin = false;
  let dragging = null;
  let measureMode = false;
  let calibFirst = null;
  let measureFirst = null;
  let fieldMode = false;
  let calibAwait = null; // 'main'

  // Photo modal state
  const photoState = { pin:null, idx:0, measuring:false, calib:null, tempLine:null }; // tempLine for live rubberband
  let photoMeasureFirst = null;

  /*******************
   * Utilities
   *******************/
  function showView(id){
    document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
    $(id).classList.add('active');
  }

  function id(){ return Math.random().toString(36).slice(2,10); }
  function fix(n){ return typeof n==='number'? +Number(n||0).toFixed(3):'' }
  function pctLabel(x,y){ return `${(x||0).toFixed(2)}%, ${(y||0).toFixed(2)}%`; }
  function dist(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return Math.sqrt(dx*dx+dy*dy); }

  function toPctCoords(e){
    const rect=stageImage.getBoundingClientRect();
    const x=Math.min(Math.max(0,e.clientX-rect.left),rect.width);
    const y=Math.min(Math.max(0,e.clientY-rect.top),rect.height);
    return { x_pct: +(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
  }
  function toLocal(e){ const r=stageImage.getBoundingClientRect(); return { x:e.clientX-r.left, y:e.clientY-r.top }; }
  function toScreen(pt){ const r=stageImage.getBoundingClientRect(); return { x:r.left+pt.x, y:r.top+pt.y }; }

  function csvEscape(v){ const s=(v==null? '': String(v)); if(/["\n,]/.test(s)) return '"'+s.replace(/"/g,'""')+'"'; return s; }
  function parseCsvLine(line){
    const out=[]; let cur=''; let i=0; let inQ=false;
    while(i<line.length){
      const ch=line[i++];
      if(inQ){
        if(ch==='"'){ if(line[i]==='"'){ cur+='"'; i++; } else { inQ=false; } }
        else { cur+=ch; }
      }else{
        if(ch===','){ out.push(cur); cur=''; }
        else if(ch==='"'){ inQ=true; }
        else { cur+=ch; }
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
    const arr=dataURL.split(',');
    const match=(arr[0].match(/:(.*?);/)||[]);
    const mime=match[1]||'image/png';
    const b=dataURLtoArrayBuffer(dataURL);
    return new Blob([b],{type:mime});
  }
  function downloadText(name, text){ downloadFile(name, new Blob([text],{type:'text/plain'})); }
  function downloadFile(name, blob){
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 800);
  }

  /*********************
   * Data model helpers *
   *********************/
  function newProject(name){
    return { id:id(), name, createdAt:Date.now(), updatedAt:Date.now(), pages:[], settings:{ colorsByType:SIGN_TYPES, fieldMode:false } };
  }
  function currentPage(){
    if(!project) return null;
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId) || null;
  }
  function makePin(x_pct,y_pct){
    return { id:id(), sign_type:'', room_number:'', room_name:'', building: inputBuilding.value||projectContext.building||'', level: inputLevel.value||projectContext.level||'', x_pct, y_pct, notes:'', photos:[], lastEdited:Date.now() };
  }
  function findPin(idv){
    if(!project) return null;
    for(const pg of project.pages){ const f=(pg.pins||[]).find(p=>p.id===idv); if(f) return f; }
    return null;
  }
  function checkedPinIds(){ return [...pinList.querySelectorAll('input[type="checkbox"]:checked')].map(cb=>cb.dataset.id); }

  function selectProject(pid){
    project = projects.find(p=>p.id===pid) || null;
    if(!project){ alert('Project not found'); return; }
    localStorage.setItem('survey:lastOpenProjectId', pid);
    renderAll();
    showView('editorView');
  }
  function saveProject(p){
    p.updatedAt=Date.now();
    const i=projects.findIndex(x=>x.id===p.id);
    if(i>=0) projects[i]=p; else projects.push(p);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    projects = loadProjects();
    savedToast();
  }
  function loadProjects(){
    try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); }catch{ return []; }
  }

  // Projects grid (Projects view)
  function renderProjectsGrid(){
    const grid=$('projectsGrid'); if(!grid) return;
    grid.innerHTML='';
    const list = loadProjects().sort((a,b)=> (b.updatedAt||0)-(a.updatedAt||0));
    if(!list.length){
      grid.innerHTML = `<div class="muted">No projects yet. Create one with “New Project”.</div>`;
      return;
    }
    list.forEach(p=>{
      const card=document.createElement('div'); card.className='project-card';
      const pages = p.pages?.length||0;
      const pins = p.pages?.reduce((a,pg)=>a+(pg.pins?.length||0),0) || 0;
      card.innerHTML = `
        <h4>${p.name}</h4>
        <div class="meta">${pages} page${pages!==1?'s':''} • ${pins} pin${pins!==1?'s':''}</div>
        <div class="meta">Updated ${new Date(p.updatedAt||p.createdAt).toLocaleString()}</div>
      `;
      card.onclick=()=> selectProject(p.id);
      grid.appendChild(card);
    });
  }
  function createAndOpenProject(){
    const name=prompt('New project name?','New Project'); if(!name) return;
    const p=newProject(name); projects.push(p); saveProject(p); selectProject(p.id);
  }

  /****************
   * Rendering
   ****************/
  function renderAll(){
    if(!project){
      // first run init
      const any = loadProjects();
      if(any.length){ selectProject(any[0].id); return; }
      const p = newProject('Untitled Project'); projects.push(p); saveProject(p); selectProject(p.id); return;
    }
    renderTypeOptionsOnce();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    renderProjectLabel();
    drawMeasurements();
    toggleField.checked=!!project.settings.fieldMode;
  }

  let _typesFilled = false;
  function renderTypeOptionsOnce(){
    if(_typesFilled) return;
    Object.keys(SIGN_TYPES).forEach(t=>{
      const o=document.createElement('option'); o.value=t; o.textContent=t;
      filterType.appendChild(o.cloneNode(true));
      fieldType.appendChild(o);
    });
    _typesFilled = true;
  }

  function renderProjectLabel(){
    projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${project.pages.reduce((a,p)=>a+(p.pins?.length||0),0)}`;
  }

  function renderThumbs(){
    thumbsEl.innerHTML='';
    (project.pages||[]).forEach((pg)=>{
      const d=document.createElement('div'); d.className='thumb'+(project._pageId===pg.id?' active':'');
      const im=document.createElement('img'); im.src=pg.blobUrl; d.appendChild(im);
      const inp=document.createElement('input'); inp.value=pg.name; inp.oninput=()=>{ pg.name=inp.value; pg.updatedAt=Date.now(); saveProject(project); renderProjectLabel(); }; d.appendChild(inp);
      d.onclick=()=>{ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); drawMeasurements(); };
      thumbsEl.appendChild(d);
    });
  }

  function renderStage(){
    const pg=currentPage();
    if(!pg){ stageImage.removeAttribute('src'); return; }
    stageImage.src=pg.blobUrl;
  }

  function renderPins(){
    pinLayer.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    fieldMode=!!project.settings.fieldMode;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type!==typeFilter) return;
      const line=[p.sign_type,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el=document.createElement('div');
      el.className='pin'+(p.id===project._sel?' selected':'');
      el.dataset.id=p.id;
      el.textContent= p.room_name?.trim() ? p.room_name : (SHORT[p.sign_type]|| (p.sign_type||'PIN').slice(0,4).toUpperCase());
      el.style.left=p.x_pct+'%';
      el.style.top=p.y_pct+'%';
      el.style.background=SIGN_TYPES[p.sign_type]||'#22c55e';
      el.style.padding = fieldMode? '.32rem .55rem' : '.22rem .45rem';
      el.style.fontSize= fieldMode? '0.95rem' : '0.78rem';
      pinLayer.appendChild(el);

      // drag handlers
      el.addEventListener('pointerdown',(e)=>{
        selectPin(p.id);
        dragging = el;
        el.setPointerCapture?.(e.pointerId);
      });
    });

    updatePinDetails();
  }

  // drag move
  on(pinLayer,'pointermove',(e)=>{
    if(!dragging) return;
    e.preventDefault();
    const idv = dragging.dataset.id;
    const pin = findPin(idv); if(!pin) return;
    const {x_pct,y_pct} = toPctCoords(e);
    dragging.style.left = x_pct+'%';
    dragging.style.top = y_pct+'%';
    posLabel.textContent = pctLabel(x_pct,y_pct);
  });
  on(pinLayer,'pointerup',(e)=>{
    if(!dragging) return;
    const idv = dragging.dataset.id;
    const pin = findPin(idv);
    if(pin){
      const {x_pct,y_pct}=toPctCoords(e);
      commit(); pin.x_pct=x_pct; pin.y_pct=y_pct; pin.lastEdited=Date.now(); saveProject(project); renderPinsList();
    }
    dragging.releasePointerCapture?.(e.pointerId); dragging=null;
  });

  function renderPinsList(){
    pinList.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      const line=[p.sign_type,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;
      if(typeFilter && p.sign_type!==typeFilter) return;

      const row=document.createElement('div'); row.className='item';
      row.style.display='grid'; row.style.gridTemplateColumns='20px 1fr auto'; row.style.alignItems='center'; row.style.gap='.5rem';
      row.style.padding='.35rem .45rem'; row.style.border='1px solid #27335f'; row.style.borderRadius='10px'; row.style.background='#0e1634';

      const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; row.appendChild(cb);
      const txt=document.createElement('div');
      txt.innerHTML=`<strong>${p.sign_type||'-'}</strong> • ${p.room_number||'-'} ${p.room_name||''} <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
      row.appendChild(txt);
      const go=document.createElement('button'); go.textContent='Go'; go.onclick=()=>{ selectPin(p.id); }; row.appendChild(go);
      pinList.appendChild(row);
    });
  }

  function updatePinDetails(){
    const p=selectedPin();
    selId.textContent=p? p.id : 'None';
    fieldType.value=p?.sign_type||'';
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
    if(['Callbox','Evac','Hall Direct'].includes(p.sign_type) && (p.room_name||'').toUpperCase()!=='ELEV. LOBBY'){
      list.push('Tip: Room name usually “ELEV. LOBBY”.');
    }
    const rn=(p.room_name||'').toUpperCase();
    if((/ELECTRICAL|DATA/).test(rn) && p.sign_type!=='BOH'){
      list.push('Recommended BOH for ELECTRICAL/DATA rooms.');
    }
    list.forEach(s=>{ const t=document.createElement('span'); t.className='tag'; t.textContent=s; warnsEl.appendChild(t); });
  }

  function selectedPin(){
    const idv=project? project._sel : null;
    if(!idv) return null;
    return findPin(idv);
  }

  function selectPin(idv){
    if(!project) return;
    project._sel=idv; saveProject(project); renderPins();
    const el=[...pinLayer.children].find(x=>x.dataset.id===idv);
    if(el){ el.scrollIntoView({block:'center', inline:'center', behavior:'smooth'}); }
    const p=findPin(idv);
    if(p){
      const pgId=project.pages.find(pg=>pg.pins.includes(p))?.id;
      if(pgId && project._pageId!==pgId){ project._pageId=pgId; saveProject(project); renderStage(); renderPins(); drawMeasurements(); }
    }
  }

  /***********************
   * Measuring (main view)
   ***********************/
  function startCalibration(scope){
    calibFirst=null; measureFirst=null;
    alert('Calibration: click two points on the page to set real feet.');
    measureMode=false; $('btnMeasureToggle').textContent='Measure: OFF';
    calibAwait=scope;
  }
  function toggleMeasuring(scope){
    measureMode=!measureMode;
    if(scope==='main') $('btnMeasureToggle').textContent = 'Measure: '+(measureMode?'ON':'OFF');
  }
  function resetMeasurements(scope){
    if(scope==='main'){
      const page=currentPage(); if(!page) return;
      page.measurements=[]; drawMeasurements(); saveProject(project);
    }else{
      const ph = currentPhoto(); if(ph){ ph.measurements=[]; drawPhotoMeasurements(); saveProject(project); }
    }
  }

  on(stage,'click',(e)=>{
    const page=currentPage(); if(!page) return;
    const pt=toLocal(e);
    if(calibAwait==='main'){
      if(!calibFirst) { calibFirst=pt; return; }
      const px = dist(calibFirst, pt);
      const ft = parseFloat(prompt('Enter real distance (feet):','10')) || 10;
      page.scalePxPerFt = px/ft; calibFirst=null; calibAwait=null;
      alert('Calibrated: '+(px/ft).toFixed(2)+' px/ft');
      return;
    }
    if(measureMode){
      if(!measureFirst){ measureFirst=pt; return; }
      const m={id:id(), kind:'main', points:[measureFirst, pt]};
      if(page.scalePxPerFt){
        const calc = dist(measureFirst, pt)/page.scalePxPerFt;
        const input = prompt(`Measured ≈ ${calc.toFixed(2)} ft. Enter value to save (or leave to accept):`, calc.toFixed(2));
        m.feet = parseFloat(input)||calc;
      }else{
        const input = prompt('Enter measured feet:','10');
        m.feet = parseFloat(input)||0;
      }
      page.measurements = page.measurements||[]; page.measurements.push(m); measureFirst=null; drawMeasurements(); saveProject(project);
    }
  });

  function drawMeasurements(){
    const page=currentPage();
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    if(!page||!page.measurements) return;
    page.measurements.forEach(m=>{
      const a=toScreen(m.points[0]); const b=toScreen(m.points[1]);
      const dx=Math.abs(a.x-b.x), dy=Math.abs(a.y-b.y);
      const color = (dx>dy)? '#ef4444' : (dy>dx)? '#3b82f6' : '#f59e0b';
      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y); line.setAttribute('x2',b.x); line.setAttribute('y2',b.y); line.setAttribute('stroke',color); measureSvg.appendChild(line);
      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2; const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      text.setAttribute('x',midx); text.setAttribute('y',midy-6);
      const label = (typeof m.feet==='number' && !isNaN(m.feet)) ? m.feet.toFixed(2)+' ft' :
        (page.scalePxPerFt? (dist(m.points[0],m.points[1])/page.scalePxPerFt).toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px'));
      text.textContent=label; measureSvg.appendChild(text);
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
  function closePhotoModal(){ photoModal.style.display='none'; }

  function currentPhoto(){
    const ph=photoState.pin?.photos?.[photoState.idx];
    return ph || null;
  }

  function showPhoto(){
    const ph=currentPhoto(); if(!ph) return;
    // set onload to sync overlay size to natural image size
    photoImg.onload = () => {
      photoMeasureSvg.setAttribute('width', photoImg.naturalWidth);
      photoMeasureSvg.setAttribute('height', photoImg.naturalHeight);
      drawPhotoMeasurements();
    };
    photoImg.src=ph.dataUrl;
    photoPinId.textContent=photoState.pin.id;
    photoName.textContent=ph.name || ('Photo '+(photoState.idx+1));
    photoMeaCount.textContent=(ph.measurements?.length||0);
  }

  function drawPhotoMeasurements(){
    while(photoMeasureSvg.firstChild) photoMeasureSvg.removeChild(photoMeasureSvg.firstChild);
    const ph=currentPhoto(); if(!ph) return;

    // draw temp rubberband if any
    if(photoState.tempLine){
      photoMeasureSvg.appendChild(photoState.tempLine);
    }

    (ph.measurements||[]).forEach(m=>{
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
      const label = (typeof m.feet==='number' && !isNaN(m.feet)) ? m.feet.toFixed(2)+' ft' :
        (ph.scalePxPerFt? (dist(a,b)/ph.scalePxPerFt).toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px'));
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=label;
      text.setAttribute('fill','#fff'); text.setAttribute('stroke','#000'); text.setAttribute('stroke-width','3'); text.setAttribute('paint-order','stroke');
      photoMeasureSvg.appendChild(text);
    });
  }

  // Rubberband helpers
  function makeTempLine(ax,ay,bx,by){
    const line=document.createElementNS('http://www.w3.org/2000/svg','line');
    line.setAttribute('x1',ax); line.setAttribute('y1',ay);
    line.setAttribute('x2',bx); line.setAttribute('y2',by);
    line.setAttribute('stroke','#22c55e'); line.setAttribute('stroke-width','3'); line.setAttribute('stroke-dasharray','6 6');
    return line;
  }

  on($('btnPhotoClose'),'click',()=> closePhotoModal());
  on($('btnPhotoMeasure'),'click',()=>{
    photoState.measuring=!photoState.measuring;
    $('btnPhotoMeasure').textContent='Measure: '+(photoState.measuring?'ON':'OFF');
    photoMeasureFirst=null; photoState.tempLine=null; drawPhotoMeasurements();
  });
  on($('btnPhotoCalib'),'click',()=>{
    photoState.calib=null; alert('Photo calibration: click two points on the image representing a known length. After the second click, enter the feet.');
  });
  on($('btnPhotoPrev'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on($('btnPhotoNext'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on($('btnPhotoDelete'),'click',()=>{ const phs=photoState.pin?.photos; if(!phs) return; if(!confirm('Delete this photo?')) return; commit(); phs.splice(photoState.idx,1); photoState.idx=Math.max(0,photoState.idx-1); saveProject(project); if(phs.length===0){ closePhotoModal(); } else { showPhoto(); } });
  on($('btnPhotoDownload'),'click',()=>{ const ph=currentPhoto(); if(!ph) return; downloadFile(ph.name||'photo.png', dataURLtoBlob(ph.dataUrl)); });

  // Clicks on overlay to measure + live rubberband
  on(photoMeasureSvg,'pointerdown',(e)=>{
    if(!photoState.measuring && photoState.calib===null) return; // measuring off and not in calib prompt state
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };

    if(photoState.calib===null && !photoMeasureFirst){
      photoMeasureFirst=pt;
      photoState.tempLine = makeTempLine(pt.x,pt.y,pt.x,pt.y);
      drawPhotoMeasurements();
      photoMeasureSvg.setPointerCapture?.(e.pointerId);
    }else if(photoState.calib===null && photoMeasureFirst){
      // finish measure line
      const start = photoMeasureFirst;
      const end = pt;
      const ph=currentPhoto(); if(!ph) return;
      const pxLen = dist(start,end);
      let defaultFeet = 0;
      if(ph.scalePxPerFt) defaultFeet = pxLen / ph.scalePxPerFt;
      const input = prompt(defaultFeet? `Measured ≈ ${defaultFeet.toFixed(2)} ft. Enter value to save (or leave to accept):`
                                   : 'Enter measured feet:','10');
      const feet = (input!==null && input.trim()!=='') ? parseFloat(input) : (defaultFeet||0);
      const m={ id:id(), kind:'photo', points:[start, end], feet: (isNaN(feet)?0:feet) };
      ph.measurements=ph.measurements||[]; ph.measurements.push(m);
      photoMeasureFirst=null; photoState.tempLine=null; drawPhotoMeasurements(); $('photoMeaCount').textContent=ph.measurements.length; saveProject(project);
    }
  });

  on(photoMeasureSvg,'pointermove',(e)=>{
    if(photoMeasureFirst && photoState.tempLine){
      const rect = photoImg.getBoundingClientRect();
      const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };
      photoState.tempLine.setAttribute('x2', pt.x);
      photoState.tempLine.setAttribute('y2', pt.y);
      // live redraw keeps temp line on top
      drawPhotoMeasurements();
    }
  });

  // Calibration by two clicks anywhere in viewer (uses same overlay)
  on(photoMeasureSvg,'dblclick',(e)=>{
    // Optional shortcut: double-click cancels measurement
    photoMeasureFirst=null; photoState.tempLine=null; drawPhotoMeasurements();
  });

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    (project.pages||[]).forEach(pg=> (pg.pins||[]).forEach(p=> rows.push({
      id:p.id, sign_type:p.sign_type||'', room_number:p.room_number||'', room_name:p.room_name||'',
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
    downloadText(project.name.replace(/\W+/g,'_')+"_signage.csv", csv);
  }

  function exportXLSX(){
    const rows = toRows(); const wb = XLSX.utils.book_new();
    const info=[[ 'Project', project.name ], [ 'Exported', new Date().toLocaleString() ], [ 'Total Signs', rows.length ], [ 'Total Pages', project.pages.length ]];
    const counts = {}; rows.forEach(r=> counts[r.sign_type]=(counts[r.sign_type]||0)+1 );
    info.push([]); info.push(['Breakdown']); Object.entries(counts).forEach(([k,v])=> info.push([k,v]));
    const wsInfo = XLSX.utils.aoa_to_sheet(info);
    const ws = XLSX.utils.json_to_sheet(rows, {header: Object.keys(rows[0]||{})});
    XLSX.utils.book_append_sheet(wb, wsInfo, 'Project Info'); XLSX.utils.book_append_sheet(wb, ws, 'Signage');
    const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    downloadFile(project.name.replace(/\W+/g,'_')+'_signage.xlsx', new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  }

  async function exportZIP(){
    const rows=toRows(); const zip=new JSZip();
    zip.file('signage.csv', [Object.keys(rows[0]||{}).join(','), ...rows.map(r=>Object.values(r).map(v=>csvEscape(v)).join(','))].join('\n'));
    // photos
    (project.pages||[]).forEach(pg=> (pg.pins||[]).forEach(pin=> (pin.photos||[]).forEach((ph, idx)=>{
      const folder=zip.folder(`photos/${pin.id}`); folder.file(ph.name||`photo_${idx+1}.png`, dataURLtoArrayBuffer(ph.dataUrl));
    })));
    const blob = await zip.generateAsync({type:'blob'}); downloadFile(project.name.replace(/\W+/g,'_')+'_export.zip', blob);
  }

  function importCSVText(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2) return;
    const hdr=lines[0].split(',').map(h=>h.trim()); commit();
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]); const row={}; hdr.forEach((h,idx)=> row[h]=cells[idx]||'');
      const p = makePin(parseFloat(row.x_pct)||50, parseFloat(row.y_pct)||50);
      p.sign_type=row.sign_type||''; p.room_number=row.room_number||''; p.room_name=row.room_name||''; p.building=row.building||''; p.level=row.level||''; p.notes=row.notes||'';
      currentPage().pins.push(p);
    }
    saveProject(project); renderPins(); renderPinsList(); alert('Imported rows into current page.');
  }

  /***************
   * Persistence / Undo
   ***************/
  function snapshot(){ return JSON.stringify(project); }
  function loadSnapshot(s){ project=JSON.parse(s); const i=projects.findIndex(p=>p.id===project.id); if(i>=0) projects[i]=project; else projects.push(project); localStorage.setItem('survey:projects', JSON.stringify(projects)); localStorage.setItem('survey:lastOpenProjectId', project.id); }
  function commit(){ UNDO.push(snapshot()); if(UNDO.length>MAX_UNDO) UNDO.shift(); REDO.length=0; }
  function undo(){ if(!UNDO.length) return; REDO.push(snapshot()); const s=UNDO.pop(); loadSnapshot(s); renderAll(); }
  function redo(){ if(!REDO.length) return; UNDO.push(snapshot()); const s=REDO.pop(); loadSnapshot(s); renderAll(); }

  /****************
   * Toolbar wires
   ****************/
  on($('btnNew'),'click',()=> createAndOpenProject());
  on($('btnSaveAs'),'click',()=>{ const name=prompt('Duplicate as name:', project.name+' (copy)'); if(!name) return; commit(); const copy=JSON.parse(JSON.stringify(project)); copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now(); saveProject(copy); selectProject(copy.id); });
  on($('btnRename'),'click',()=>{ const name=prompt('Rename project:', project.name); if(!name) return; commit(); project.name=name; saveProject(project); renderProjectLabel(); });

  on($('btnUpload'),'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){ const url=URL.createObjectURL(f); const name=f.name.replace(/\.[^.]+$/,''); addImagePage(url,name); }
    }
    renderAll();
    e.target.value='';
  });

  on($('btnExportCSV'),'click',()=> exportCSV());
  on($('btnExportXLSX'),'click',()=> exportXLSX());
  on($('btnExportZIP'),'click',()=> exportZIP());

  on($('btnImportCSV'),'click',()=> $('inputImportCSV').click());
  on($('inputImportCSV'),'change', async (e)=>{ const f=e.target.files?.[0]; if(!f) return; const text=await f.text(); importCSVText(text); e.target.value=''; });

  on($('btnOCR'),'click',()=> ocrCurrentView());

  const btnCalibrate = $('btnCalibrate');
  const btnMeasureToggle = $('btnMeasureToggle');
  const btnMeasureReset = $('btnMeasureReset');
  on(btnCalibrate,'click',()=> startCalibration('main'));
  on(btnMeasureToggle,'click',()=> toggleMeasuring('main'));
  on(btnMeasureReset,'click',()=> resetMeasurements('main'));

  on($('btnUndo'),'click',()=> undo());
  on($('btnRedo'),'click',()=> redo());

  on($('btnClearPins'),'click',()=>{ if(!confirm('Clear ALL pins on this page?')) return; commit(); const pg=currentPage(); if(pg){ pg.pins=[]; } renderAll(); saveProject(project); });
  on($('btnAddPin'),'click',()=> startAddPin());

  on(toggleField,'change',()=>{ project.settings.fieldMode = !!toggleField.checked; saveProject(project); renderPins(); });

  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; });

  on(inputSearch,'input',()=> renderPinsList());
  on(filterType,'change',()=> renderPins());

  // Right panel field changes
  ;[fieldType, fieldRoomNum, fieldRoomName, fieldBuilding, fieldLevel, fieldNotes].forEach(el=> on(el,'input',()=>{
    const pin = selectedPin(); if(!pin) return; commit();
    if(el===fieldType) pin.sign_type = el.value || '';
    if(el===fieldRoomNum) pin.room_number = el.value || '';
    if(el===fieldRoomName) pin.room_name = el.value || '';
    if(el===fieldBuilding) pin.building = el.value || '';
    if(el===fieldLevel) pin.level = el.value || '';
    if(el===fieldNotes) pin.notes = el.value || '';
    pin.lastEdited=Date.now(); saveProject(project); renderPins(); renderWarnings(); renderPinsList();
  }));

  // Photo add/open/duplicate/delete
  on($('btnAddPhoto'),'click',()=> $('inputPhoto').click());
  on($('inputPhoto'),'change', async (e)=>{
    const pin = selectedPin(); if(!pin) return alert('Select a pin first.');
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){ const url=await fileToDataURL(f); pin.photos.push({name:f.name||'photo.png',dataUrl:url,measurements:[]}); }
    pin.lastEdited=Date.now(); saveProject(project); renderPinsList(); alert('Photo(s) added.');
    e.target.value='';
  });
  on($('btnOpenPhoto'),'click',()=> openPhotoModal());
  on($('btnDuplicate'),'click',()=>{ const pin=selectedPin(); if(!pin) return; commit(); const p=JSON.parse(JSON.stringify(pin)); p.id=id(); p.x_pct=Math.min(100,p.x_pct+2); p.y_pct=Math.min(100,p.y_pct+2); p.lastEdited=Date.now(); currentPage().pins.push(p); saveProject(project); renderPins(); renderPinsList(); selectPin(p.id); });
  on($('btnDelete'),'click',()=>{ const pin=selectedPin(); if(!pin) return; if(!confirm('Delete selected pin?')) return; commit(); const pg=currentPage(); pg.pins=pg.pins.filter(x=>x.id!==pin.id); saveProject(project); renderPins(); renderPinsList(); clearSelection(); });

  on($('btnBulkType'),'click',()=>{ const type = prompt('Enter type for selected checkboxes (or empty to cancel):'); if(type==null) return; commit(); const ids = checkedPinIds(); ids.forEach(idv=>{ const p=findPin(idv); if(p){ p.sign_type=type; p.lastEdited=Date.now(); }}); saveProject(project); renderPins(); renderPinsList(); });
  on($('btnBulkBL'),'click',()=>{ const b = prompt('Building value (blank to keep):', inputBuilding.value||''); if(b==null) return; const l = prompt('Level value (blank to keep):', inputLevel.value||''); if(l==null) return; commit(); const ids = checkedPinIds(); ids.forEach(idv=>{ const p=findPin(idv); if(p){ if(b!=='') p.building=b; if(l!=='') p.level=l; p.lastEdited=Date.now(); }}); saveProject(project); renderPins(); renderPinsList(); });

  // Stage interactions
  let startAdd=false;
  function startAddPin(){ addingPin=!addingPin; $('btnAddPin').classList.toggle('primary', addingPin); }

  on($('stage'),'pointerdown',(e)=>{
    if(e.target.classList.contains('pin')) return; // handled by pin
    if(!addingPin) return;
    const {x_pct,y_pct} = toPctCoords(e);
    commit();
    const p = makePin(x_pct,y_pct);
    currentPage().pins.push(p); saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
    addingPin=false; $('btnAddPin').classList.remove('primary');
  });

  function clearSelection(){ if(!project) return; project._sel=null; saveProject(project); updatePinDetails(); }

  /***********
   * OCR
   ***********/
  async function ocrCurrentView(){
    const canvas = document.createElement('canvas');
    const img=stageImage; if(!img || !img.src){ alert('No page image.'); return; }
    const w=img.naturalWidth||img.width, h=img.naturalHeight||img.height; canvas.width=w; canvas.height=h;
    const ctx=canvas.getContext('2d'); ctx.drawImage(img,0,0);
    const { data: { text } } = await Tesseract.recognize(canvas, 'eng');
    const out = (text||'').trim();
    if(!out){ alert('No text recognized.'); return; }
    try { await navigator.clipboard.writeText(out); } catch{}
    inputSearch.value=out.slice(0,200);
    renderPinsList();
    alert('OCR done. First 200 chars placed in search. Full text copied to clipboard.');
  }

  /*********************
   * Page adders
   *********************/
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
      const canvas=document.createElement('canvas'); canvas.width=viewport.width; canvas.height=viewport.height;
      const ctx=canvas.getContext('2d'); await page.render({canvasContext:ctx, viewport}).promise;
      const data=canvas.toDataURL('image/png');
      const pg={ id:id(), name:`${file.name.replace(/\.[^.]+$/,'')} · p${i}`, kind:'pdf', pdfPage:i, blobUrl=data, pins:[], measurements:[], updatedAt:Date.now() };
      project.pages.push(pg);
    }
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    saveProject(project);
  }

  /********************
   * Resize observer
   ********************/
  renderAll();
  const ro=new ResizeObserver(()=>{ measureSvg.setAttribute('width', stage.clientWidth); measureSvg.setAttribute('height', stage.clientHeight); });
  ro.observe(stage);

  // Keyboard shortcuts
  window.addEventListener('keydown',(e)=>{
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); redo(); }
    if(e.key==='Delete'){ const p=selectedPin(); if(p){ commit(); const pg=currentPage(); pg.pins=pg.pins.filter(x=>x.id!==p.id); saveProject(project); renderPins(); renderPinsList(); clearSelection(); } }
    if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)){ const p=selectedPin(); if(!p) return; e.preventDefault(); commit(); const delta = project.settings.fieldMode? 0.5 : 0.2; if(e.key==='ArrowUp') p.y_pct=Math.max(0,p.y_pct-delta); if(e.key==='ArrowDown') p.y_pct=Math.min(100,p.y_pct+delta); if(e.key==='ArrowLeft') p.x_pct=Math.max(0,p.x_pct-delta); if(e.key==='ArrowRight') p.x_pct=Math.min(100,p.x_pct+delta); p.lastEdited=Date.now(); saveProject(project); renderPins(); updatePinDetails(); }
  });

  // Initial view route
  window.addEventListener('load', ()=>{
    const last = localStorage.getItem('survey:lastOpenProjectId');
    const hasAny = (loadProjects().length>0);
    if(last && projects.find(p=>p.id===last)){ selectProject(last); showView('editorView'); }
    else if(hasAny){ renderProjectsGrid(); showView('projectsView'); }
    else { showView('homeView'); }
  });
});
