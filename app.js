// app.js
document.addEventListener('DOMContentLoaded', () => {
  /***********************
   * Helpers & constants *
   ***********************/
  const $ = (id) => document.getElementById(id);
  const on = (el, evt, fn) => el && el.addEventListener(evt, fn);

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
  // Left column
  const thumbsEl = $('thumbs');

  // Stage
  const stage = $('stage');
  const stageImage = $('stageImage');
  const pinLayer = $('pinLayer');
  const measureSvg = $('measureSvg');

  // Toolbar bits
  const projectLabel = $('projectLabel');
  const inputUpload = $('inputUpload');
  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');
  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleStrict = $('toggleStrict');
  const toggleField = $('toggleField');

  // Right panel — fields
  const fieldType = $('fieldType');
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

  // Top toolbar buttons (by id)
  const btnNew = $('btnNew');
  const btnOpen = $('btnOpen');
  const btnSaveAs = $('btnSaveAs');
  const btnRename = $('btnRename');
  const btnUpload = $('btnUpload');
  const btnExportCSV = $('btnExportCSV');
  const btnExportXLSX = $('btnExportXLSX');
  const btnExportZIP = $('btnExportZIP');
  const btnImportCSV = $('btnImportCSV');
  const inputImportCSV = $('inputImportCSV');

  const btnCalibrate = $('btnCalibrate');
  const btnMeasureToggle = $('btnMeasureToggle');
  const btnMeasureReset = $('btnMeasureReset');

  const btnUndo = $('btnUndo');
  const btnRedo = $('btnRedo');
  const btnClearPins = $('btnClearPins');
  const btnAddPin = $('btnAddPin');

  const btnAddPhoto = $('btnAddPhoto');
  const inputPhoto = $('inputPhoto');
  const btnOpenPhoto = $('btnOpenPhoto');
  const btnDuplicate = $('btnDuplicate');
  const btnDelete = $('btnDelete');

  // Photo modal buttons
  const btnPhotoClose = $('btnPhotoClose');
  const btnPhotoMeasure = $('btnPhotoMeasure');
  const btnPhotoCalib = $('btnPhotoCalib');
  const btnPhotoPrev = $('btnPhotoPrev');
  const btnPhotoNext = $('btnPhotoNext');
  const btnPhotoDelete = $('btnPhotoDelete');
  const btnPhotoDownload = $('btnPhotoDownload');

  /************************
   * App state / Undo/redo
   ************************/
  let projects = [];
  let currentProjectId = null;
  let project = null;

  const UNDO = [];
  const REDO = [];
  const MAX_UNDO = 50;

  // Ephemeral context (for default building/level when adding)
  const projectContext = { building:'', level:'' };

  // Stage interaction state
  let addingPin = false;
  let dragging = null;        // HTMLElement of the pin being dragged
  let dragData = null;        // {startX, startY, startPctX, startPctY}
  let measureMode = false;
  let calibFirst = null;
  let measureFirst = null;    // main view measurement temp start
  let fieldMode = false;
  let calibAwait = null;      // 'main' | 'photo'

  // Photo modal state
  const photoState = { pin:null, idx:0, measure:false, calib:null };
  let photoMeasureFirst = null;          // photo measurement start
  let photoRubberLine = null;            // SVG line element for rubber band
  let mainRubberLine = null;             // SVG line element in main view

  /*******************
   * Small utilities *
   *******************/
  function id(){ return Math.random().toString(36).slice(2,10); }
  function fix(n){ return typeof n==='number'? +Number(n||0).toFixed(3):''; }
  function pctLabel(x,y){ return `${(x||0).toFixed(2)}%, ${(y||0).toFixed(2)}%`; }
  function dist(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return Math.sqrt(dx*dx+dy*dy); }

  function toPctCoordsFromPoint(clientX, clientY){
    const rect=stageImage.getBoundingClientRect();
    const x=Math.min(Math.max(0,clientX-rect.left),rect.width);
    const y=Math.min(Math.max(0,clientY-rect.top),rect.height);
    return { x_pct: +(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
  }
  function toPctCoords(e){ return toPctCoordsFromPoint(e.clientX, e.clientY); }
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
    const ab=dataURLtoArrayBuffer(dataURL);
    return new Blob([ab],{type:mime});
  }
  function downloadText(name, text){ downloadFile(name, new Blob([text],{type:'text/plain'})); }
  function downloadFile(name, blob){
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  /*********************
   * Data model helpers *
   *********************/
  function newProject(name){
    return {
      id: id(),
      name: name,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      pages: [],
      settings: { colorsByType: SIGN_TYPES, fieldMode: false, strictRules: false }
    };
  }
  function currentPage(){
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId);
  }
  function makePin(x_pct,y_pct){
    return {
      id: id(),
      sign_type: '',
      room_number: '',
      room_name: '',
      building: inputBuilding.value || projectContext.building || '',
      level: inputLevel.value || projectContext.level || '',
      x_pct: x_pct,
      y_pct: y_pct,
      custom_name: '',       // allow user-defined name
      notes: '',
      photos: [],
      lastEdited: Date.now()
    };
  }
  function findPin(idv){
    for(const pg of project.pages){
      const f=(pg.pins||[]).find(p=>p.id===idv);
      if(f) return f;
    }
    return null;
  }
  function pinPage(pinId){
    for(const pg of project.pages){
      if((pg.pins||[]).some(p=>p.id===pinId)) return pg;
    }
    return null;
  }
  function checkedPinIds(){ return [...pinList.querySelectorAll('input[type="checkbox"]:checked')].map(cb=>cb.dataset.id); }

  function selectProject(pid){
    project = projects.find(p=>p.id===pid);
    if(!project) return;
    localStorage.setItem('survey:lastOpenProjectId', pid);
    renderAll();
  }
  function saveProject(p){
    p.updatedAt=Date.now();
    const i=projects.findIndex(x=>x.id===p.id);
    if(i>=0) projects[i]=p; else projects.push(p);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    projects = loadProjects();
  }
  function loadProjects(){
    try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); }catch{ return []; }
  }

  function addImagePage(url, name){
    const pg={
      id: id(),
      name: name || 'Image',
      kind: 'image',
      blobUrl: url,
      pins: [],
      measurements: [],
      updatedAt: Date.now()
    };
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
      await page.render({canvasContext:ctx, viewport:viewport}).promise;
      const data=canvas.toDataURL('image/png');
      const pg={
        id: id(),
        name: `${file.name.replace(/\.[^.]+$/,'')} · p${i}`,
        kind: 'pdf',
        pdfPage: i,
        blobUrl: data,
        pins: [],
        measurements: [],
        updatedAt: Date.now()
      };
      project.pages.push(pg);
    }
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    saveProject(project);
  }

  /****************
   * Render funcs *
   ****************/
  function renderAll(){
    renderTypeOptionsOnce();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    renderProjectLabel();
    drawMeasurements();
    toggleStrict.checked=!!project.settings.strictRules;
    toggleField.checked=!!project.settings.fieldMode;
  }

  let _typesFilled = false;
  function renderTypeOptionsOnce(){
    if(_typesFilled) return;
    // Add "Custom..." option at top for free text entry via prompt
    const customOpt = document.createElement('option');
    customOpt.value = '__CUSTOM__';
    customOpt.textContent = 'Custom (type your own)…';
    fieldType.appendChild(customOpt);

    Object.keys(SIGN_TYPES).forEach(t=>{
      const o=document.createElement('option'); o.value=t; o.textContent=t;
      filterType.appendChild(o.cloneNode(true));
      fieldType.appendChild(o);
    });
    _typesFilled = true;
  }

  function renderProjectLabel(){
    const totalPins = project.pages.reduce((a,p)=>a+(p.pins?.length||0),0);
    projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${totalPins}`;
  }

  function renderThumbs(){
    thumbsEl.innerHTML='';
    project.pages.forEach((pg)=>{
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

  function renderPins(){
    pinLayer.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    fieldMode=!!project.settings.fieldMode;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type!==typeFilter) return;
      const line=[p.custom_name||'', p.sign_type, p.room_number, p.room_name, p.building, p.level, p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el=document.createElement('div');
      el.className='pin'+(p.id===project._sel?' selected':'');
      el.dataset.id=p.id;
      const label = p.custom_name || SHORT[p.sign_type] || (p.sign_type||'---').slice(0,3).toUpperCase();
      el.textContent=label;
      el.style.left=p.x_pct+'%';
      el.style.top=p.y_pct+'%';
      el.style.background=SIGN_TYPES[p.sign_type]||'#22c55e'; // default Southwood-green
      el.style.padding = fieldMode? '.28rem .5rem' : '.18rem .35rem';
      el.style.fontSize= fieldMode? '0.9rem' : '0.75rem';
      pinLayer.appendChild(el);

      // Dragging
      el.addEventListener('pointerdown', (e)=>{
        e.preventDefault();
        selectPin(p.id);
        dragging = el;
        const rect=stageImage.getBoundingClientRect();
        dragData = {
          startX: e.clientX,
          startY: e.clientY,
          rectLeft: rect.left,
          rectTop: rect.top,
          rectW: rect.width,
          rectH: rect.height,
          startPctX: p.x_pct,
          startPctY: p.y_pct
        };
        el.setPointerCapture?.(e.pointerId);
      });

      el.addEventListener('dblclick', ()=>{ selectPin(p.id); openPhotoModal(); });
    });

    updatePinDetails();
  }

  function renderPinsList(){
    pinList.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      const line=[p.custom_name||'', p.sign_type, p.room_number, p.room_name, p.building, p.level, p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;
      if(typeFilter && p.sign_type!==typeFilter) return;

      const row=document.createElement('div'); row.className='item';
      const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; row.appendChild(cb);
      const nameTxt = p.custom_name ? ` (${p.custom_name})` : '';
      const txt=document.createElement('div');
      txt.innerHTML=`<strong>${p.sign_type||'-'}${nameTxt}</strong> • ${p.room_number||'-'} ${p.room_name||''} <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
      row.appendChild(txt);
      const go=document.createElement('button'); go.textContent='Go'; go.onclick=()=>{ selectPin(p.id); }; row.appendChild(go);
      pinList.appendChild(row);
    });
  }

  function updatePinDetails(){
    const p=selectedPin();
    selId.textContent=p? p.id : 'None';
    fieldType.value = p ? (p.sign_type || '') : '';
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
    list.forEach(s=>{
      const t=document.createElement('span'); t.className='tag warn'; t.textContent=s; warnsEl.appendChild(t);
    });
  }

  function selectedPin(){
    const idv=project._sel;
    if(!idv) return null;
    return findPin(idv);
  }

  function selectPin(idv){
    project._sel=idv; saveProject(project); renderPins();
    const el=[...pinLayer.children].find(x=>x.dataset.id===idv);
    if(el){ el.scrollIntoView({block:'center', inline:'center', behavior:'smooth'}); }
    const p=findPin(idv);
    if(p){
      const pg = pinPage(p.id);
      if(pg && project._pageId!==pg.id){ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); drawMeasurements(); }
    }
  }

  /***********************
   * Measuring (main view)
   ***********************/
  function startCalibration(scope){
    calibFirst=null; measureFirst=null;
    alert('Calibration: click two points on the page to set real feet.');
    measureMode=false; btnMeasureToggle.textContent='Measuring: OFF';
    calibAwait=scope;
  }
  function toggleMeasuring(scope){
    measureMode=!measureMode;
    if(scope==='main') btnMeasureToggle.textContent = 'Measuring: '+(measureMode?'ON':'OFF');
    // clear rubber line if turning off
    if(!measureMode && mainRubberLine){ measureSvg.removeChild(mainRubberLine); mainRubberLine=null; }
  }
  function resetMeasurements(scope){
    if(scope==='main'){
      const pg=currentPage(); if(pg){ pg.measurements=[]; }
      drawMeasurements();
    }else{
      const ph = photoState.pin?.photos[photoState.idx];
      if(ph){ ph.measurements=[]; drawPhotoMeasurements(); }
    }
  }

  function ensureRubberLine(svg){
    const line=document.createElementNS('http://www.w3.org/2000/svg','line');
    line.setAttribute('stroke','#22c55e');
    line.setAttribute('stroke-dasharray','4 3');
    line.setAttribute('stroke-width','3');
    return line;
  }

  on(stage,'pointermove',(e)=>{
    if(measureMode && measureFirst){
      // update rubber band in main view
      const a = toLocalPoint(measureFirst);
      const b = toLocal(e);
      if(!mainRubberLine){
        mainRubberLine = ensureRubberLine(measureSvg);
        measureSvg.appendChild(mainRubberLine);
      }
      mainRubberLine.setAttribute('x1', a.x); mainRubberLine.setAttribute('y1', a.y);
      mainRubberLine.setAttribute('x2', b.x); mainRubberLine.setAttribute('y2', b.y);
    }
  });

  on(stage,'click',(e)=>{
    const page=currentPage(); if(!page) return;
    const pt=toLocal(e);

    if(calibAwait==='main'){
      if(!calibFirst) { calibFirst=pt; return; }
      const px = dist(calibFirst, pt);
      const ftInput = prompt('Enter real distance (feet) for calibration:', '10');
      const ft = ftInput ? parseFloat(ftInput) : NaN;
      if(ft && ft>0){ page.scalePxPerFt = px/ft; alert(`Calibrated: ${(px/ft).toFixed(2)} px/ft`); }
      calibFirst=null; calibAwait=null;
      return;
    }

    if(measureMode){
      if(!measureFirst){ measureFirst = { x: pt.x, y: pt.y }; return; }
      // finalize measurement
      const start = measureFirst;
      const end = pt;
      if(mainRubberLine){ measureSvg.removeChild(mainRubberLine); mainRubberLine=null; }
      const pxLen = dist(start, end);
      let ftVal = null;
      if(page.scalePxPerFt){
        const autoFt = pxLen / page.scalePxPerFt;
        const entered = prompt(`Distance (feet). Calibrated value = ${autoFt.toFixed(2)} ft. Press OK to accept or type your own:`, autoFt.toFixed(2));
        ftVal = entered ? parseFloat(entered) : autoFt;
      } else {
        const entered = prompt('Enter distance (feet) for this line:', '10');
        ftVal = entered ? parseFloat(entered) : NaN;
      }
      const m = {
        id: id(),
        kind: 'main',
        points: [ {x:start.x, y:start.y}, {x:end.x, y:end.y} ],
        feet: (ftVal && !isNaN(ftVal)) ? ftVal : undefined,
        px: pxLen
      };
      page.measurements = page.measurements || [];
      page.measurements.push(m);
      measureFirst=null;
      drawMeasurements();
    }
  });

  function toLocalPoint(pt){ return {x: pt.x, y: pt.y}; }

  function drawMeasurements(){
    const page=currentPage();
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    if(!page||!page.measurements) return;
    page.measurements.forEach(m=>{
      const a=toScreen(m.points[0]); const b=toScreen(m.points[1]);
      const dx=Math.abs(a.x-b.x), dy=Math.abs(a.y-b.y);
      const color = (dx>dy)? '#ef4444' : (dy>dx)? '#3b82f6' : '#f59e0b';
      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y);
      line.setAttribute('x2',b.x); line.setAttribute('y2',b.y);
      line.setAttribute('stroke',color);
      line.setAttribute('stroke-width','3');
      measureSvg.appendChild(line);

      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      let ft;
      if(typeof m.feet==='number'){
        ft = m.feet.toFixed(2)+' ft';
      } else if(page.scalePxPerFt){
        ft = (dist(a,b)/page.scalePxPerFt).toFixed(2)+' ft';
      } else {
        ft = (dist(a,b).toFixed(0)+' px');
      }
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=ft;
      measureSvg.appendChild(text);
    });
  }

  /*************************
   * Photo modal & measure *
   *************************/
  function openPhotoModal(){
    const p=selectedPin(); if(!p) { alert('Select a pin first.'); return; }
    if(!p.photos.length) { alert('No photos attached.'); return; }
    photoState.pin=p; photoState.idx=0; showPhoto(); photoModal.style.display='flex';
    // allow scroll inside the viewer if large image
    const viewer = photoModal.querySelector('.viewer');
    if(viewer) viewer.style.overflow = 'auto';
  }
  function showPhoto(){
    const ph=photoState.pin.photos[photoState.idx];
    photoImg.src=ph.dataUrl;
    photoPinId.textContent=photoState.pin.id;
    photoName.textContent=ph.name;
    photoMeaCount.textContent=(ph.measurements?.length||0);
    drawPhotoMeasurements();
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
      line.setAttribute('stroke',color);
      line.setAttribute('stroke-width','3');
      photoMeasureSvg.appendChild(line);

      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      let ft;
      if(typeof m.feet==='number'){
        ft = m.feet.toFixed(2)+' ft';
      } else if(ph.scalePxPerFt){
        const pxLen = Math.hypot(a.x-b.x, a.y-b.y);
        ft = (pxLen/ph.scalePxPerFt).toFixed(2)+' ft';
      } else {
        ft = (Math.hypot(a.x-b.x, a.y-b.y).toFixed(0)+' px');
      }
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=ft;
      photoMeasureSvg.appendChild(text);
    });
  }

  // Rubber band in photo view
  on(photoImg,'pointermove',(e)=>{
    if(!photoState.measure || !photoMeasureFirst) return;
    const rect = photoImg.getBoundingClientRect();
    const cur = { x: e.clientX-rect.left, y: e.clientY-rect.top };
    if(!photoRubberLine){
      photoRubberLine = ensureRubberLine(photoMeasureSvg);
      photoMeasureSvg.appendChild(photoRubberLine);
    }
    photoRubberLine.setAttribute('x1', photoMeasureFirst.x);
    photoRubberLine.setAttribute('y1', photoMeasureFirst.y);
    photoRubberLine.setAttribute('x2', cur.x);
    photoRubberLine.setAttribute('y2', cur.y);
  });

  on(btnPhotoClose,'click',()=> closePhotoModal());
  on(btnPhotoMeasure,'click',()=>{
    photoState.measure=!photoState.measure;
    btnPhotoMeasure.textContent='Measure: '+(photoState.measure?'ON':'OFF');
    if(!photoState.measure && photoRubberLine){ photoMeasureSvg.removeChild(photoRubberLine); photoRubberLine=null; photoMeasureFirst=null; }
  });
  on(btnPhotoCalib,'click',()=>{ photoState.calib=null; alert('Photo calibration: click two points on the image, then enter real feet.'); });

  on(btnPhotoPrev,'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on(btnPhotoNext,'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on(btnPhotoDelete,'click',()=>{
    if(!photoState.pin) return;
    if(!confirm('Delete this photo?')) return;
    commit();
    photoState.pin.photos.splice(photoState.idx,1);
    photoState.idx=Math.max(0,photoState.idx-1);
    saveProject(project);
    if(photoState.pin.photos.length===0){ closePhotoModal(); } else { showPhoto(); }
  });
  on(btnPhotoDownload,'click',()=>{
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    downloadFile(ph.name || 'photo.png', dataURLtoBlob(ph.dataUrl));
  });

  // Calibrate or measure click handling on photo overlay SVG
  on(photoMeasureSvg,'click',(e)=>{
    const ph = photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };

    // Calibration clicks
    if(photoState.calib===null && btnPhotoCalib){ /* waiting to start */ }
    if(photoState.calib===null && e.isTrusted){
      // first calibration point
      photoState.calib = { x: pt.x, y: pt.y };
      return;
    }
    if(photoState.calib && !photoState.measureTmp){
      // second calibration point
      const px = Math.hypot(photoState.calib.x-pt.x, photoState.calib.y-pt.y);
      const ft = parseFloat(prompt('Enter real distance (feet) for calibration:', '10')) || 10;
      ph.measurements = ph.measurements || [];
      ph.scalePxPerFt = px/ft;
      photoState.calib=null;
      drawPhotoMeasurements();
      return;
    }
  });

  on(photoImg,'click',(e)=>{
    if(!photoState.measure) return;
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect();
    const pt={x:e.clientX-rect.left, y:e.clientY-rect.top};

    if(!photoMeasureFirst){
      photoMeasureFirst = { x: pt.x, y: pt.y };
      return;
    }

    // finalize a measurement on photo
    const start = photoMeasureFirst;
    const end = pt;
    if(photoRubberLine){ photoMeasureSvg.removeChild(photoRubberLine); photoRubberLine=null; }
    const pxLen = Math.hypot(start.x-end.x, start.y-end.y);

    let ftVal = null;
    if(ph.scalePxPerFt){
      const autoFt = pxLen / ph.scalePxPerFt;
      const entered = prompt(`Distance (feet). Calibrated value = ${autoFt.toFixed(2)} ft. Press OK to accept or type your own:`, autoFt.toFixed(2));
      ftVal = entered ? parseFloat(entered) : autoFt;
    } else {
      const entered = prompt('Enter distance (feet) for this line:', '10');
      ftVal = entered ? parseFloat(entered) : NaN;
    }

    const m={
      id: id(),
      kind: 'photo',
      points: [ {x:start.x, y:start.y}, {x:end.x, y:end.y} ],
      feet: (ftVal && !isNaN(ftVal)) ? ftVal : undefined,
      px: pxLen
    };
    ph.measurements = ph.measurements || [];
    ph.measurements.push(m);
    photoMeasureFirst=null;
    drawPhotoMeasurements();
    photoMeaCount.textContent = ph.measurements.length.toString();
  });

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    project.pages.forEach(pg=> (pg.pins||[]).forEach(p=> rows.push({
      id: p.id,
      sign_type: p.sign_type || '',
      custom_name: p.custom_name || '',
      room_number: p.room_number || '',
      room_name: p.room_name || '',
      building: p.building || '',
      level: p.level || '',
      x_pct: fix(p.x_pct),
      y_pct: fix(p.y_pct),
      notes: p.notes || '',
      page_name: pg.name,
      last_edited: new Date(p.lastEdited || project.updatedAt).toISOString()
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
    project.pages.forEach(pg=>{
      (pg.pins||[]).forEach(pin=>{
        (pin.photos||[]).forEach((ph, idx)=>{
          const folder=zip.folder(`photos/${pin.id}`);
          folder.file(ph.name || `photo_${idx+1}.png`, dataURLtoArrayBuffer(ph.dataUrl));
        });
      });
    });
    const blob = await zip.generateAsync({type:'blob'});
    downloadFile(project.name.replace(/\W+/g,'_')+'_export.zip', blob);
  }

  function importCSV(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2) return;
    const hdr=lines[0].split(',').map(h=>h.trim()); commit();
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]); const row={};
      hdr.forEach((h,idx)=> row[h]=cells[idx]||'');
      const p = makePin(parseFloat(row.x_pct)||50, parseFloat(row.y_pct)||50);
      p.sign_type=row.sign_type||'';
      p.custom_name=row.custom_name||'';
      p.room_number=row.room_number||'';
      p.room_name=row.room_name||'';
      p.building=row.building||'';
      p.level=row.level||'';
      p.notes=row.notes||'';
      currentPage().pins.push(p);
    }
    saveProject(project); renderPins(); renderPinsList(); alert('Imported rows into current page.');
  }

  /***********************
   * Toolbar wiring, etc *
   ***********************/
  // Project storage init
  projects = loadProjects();
  currentProjectId = localStorage.getItem('survey:lastOpenProjectId') || null;
  project = currentProjectId ? projects.find(p=>p.id===currentProjectId) : null;
  if(!project){ project = newProject('Untitled Project'); saveProject(project); }
  selectProject(project.id);

  // Simple Project Switcher overlay (no index change)
  function openProjectSwitcher(){
    const overlay=document.createElement('div');
    overlay.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;z-index:9999;';
    const box=document.createElement('div');
    box.style.cssText='background:#0f1733;border:1px solid #2b3566;border-radius:12px;min-width:360px;max-width:90vw;max-height:80vh;overflow:auto;padding:12px;';
    const title=document.createElement('div');
    title.textContent='Open Project'; title.style.cssText='font-weight:700;margin-bottom:8px;';
    const list=document.createElement('div');
    projects.forEach(p=>{
      const item=document.createElement('div');
      item.textContent=`${p.name}  (${new Date(p.updatedAt).toLocaleString()})`;
      item.style.cssText='padding:8px;border:1px solid #2b3566;border-radius:8px;margin:6px 0;cursor:pointer;background:#121a3b;';
      item.onclick=()=>{ selectProject(p.id); document.body.removeChild(overlay); };
      list.appendChild(item);
    });
    const newBtn=document.createElement('button');
    newBtn.textContent='New Project';
    newBtn.style.cssText='margin-top:8px;';
    newBtn.onclick=()=>{
      const name=prompt('New project name?','New Project'); if(!name) return;
      commit(); const np=newProject(name); saveProject(np); selectProject(np.id); document.body.removeChild(overlay);
    };
    box.appendChild(title); box.appendChild(list); box.appendChild(newBtn);
    overlay.appendChild(box);
    overlay.addEventListener('click',(e)=>{ if(e.target===overlay) document.body.removeChild(overlay); });
    document.body.appendChild(overlay);
  }

  // Toolbar handlers
  on(btnNew,'click',()=>{ const name=prompt('New project name?','New Project'); if(!name) return; commit(); const p=newProject(name); saveProject(p); selectProject(p.id); });
  on(btnOpen,'click',openProjectSwitcher);
  on(btnSaveAs,'click',()=>{ const name=prompt('Duplicate as name:', project.name+' (copy)'); if(!name) return; commit(); const copy=JSON.parse(JSON.stringify(project)); copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now(); saveProject(copy); selectProject(copy.id); });
  on(btnRename,'click',()=>{ const name=prompt('Rename project:', project.name); if(!name) return; commit(); project.name=name; saveProject(project); renderProjectLabel(); });

  on(btnUpload,'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){ const url=URL.createObjectURL(f); const name=f.name.replace(/\.[^.]+$/,''); addImagePage(url,name); }
    }
    renderAll();
  });

  on(btnExportCSV,'click',exportCSV);
  on(btnExportXLSX,'click',exportXLSX);
  on(btnExportZIP,'click',exportZIP);

  on(btnImportCSV,'click',()=> inputImportCSV.click());
  on(inputImportCSV,'change', async (e)=>{
    const f=e.target.files && e.target.files[0]; if(!f) return; const text=await f.text(); importCSV(text);
  });

  // Pin CRUD & fields
  on(btnClearPins,'click',()=>{ if(!confirm('Clear ALL pins on this page?')) return; commit(); const pg=currentPage(); if(pg){ pg.pins=[]; } renderAll(); });
  on(btnAddPin,'click',()=>{ addingPin=!addingPin; btnAddPin.classList.toggle('ok', addingPin); });

  // Custom type via fieldType
  on(fieldType,'change',()=>{
    const pin = selectedPin(); if(!pin) return;
    if(fieldType.value==='__CUSTOM__'){
      const val = prompt('Enter custom pin name (label):','');
      if(val && val.trim()){
        pin.custom_name = val.trim();
        if(!pin.sign_type) pin.sign_type = ''; // keep sign_type empty if purely custom
      }
      // reset dropdown to blank (or keep previous)
      fieldType.value = pin.sign_type || '';
    } else {
      pin.sign_type = fieldType.value || '';
      // clear custom name if switching back to a known type (optional)
      // pin.custom_name = '';
    }
    pin.lastEdited = Date.now(); saveProject(project); renderPins(); renderPinsList(); updatePinDetails();
  });

  [fieldRoomNum, fieldRoomName, fieldBuilding, fieldLevel, fieldNotes].forEach(el=> on(el,'input',()=>{
    const pin = selectedPin(); if(!pin) return; commit();
    if(el===fieldRoomNum) pin.room_number = el.value || '';
    if(el===fieldRoomName) pin.room_name = el.value || '';
    if(el===fieldBuilding) pin.building = el.value || '';
    if(el===fieldLevel) pin.level = el.value || '';
    if(el===fieldNotes) pin.notes = el.value || '';
    pin.lastEdited=Date.now(); saveProject(project); renderPins(); renderWarnings(); renderPinsList();
  }));

  on(btnAddPhoto,'click',()=> inputPhoto.click());
  on(inputPhoto,'change', async (e)=>{
    const pin = selectedPin(); if(!pin) { alert('Select a pin first.'); return; }
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      const url=await fileToDataURL(f);
      pin.photos.push({ name: f.name, dataUrl: url, measurements: [] });
    }
    pin.lastEdited=Date.now(); saveProject(project); renderPinsList(); alert('Photo(s) added.');
  });

  on(btnOpenPhoto,'click',openPhotoModal);
  on(btnDuplicate,'click',()=>{
    const pin=selectedPin(); if(!pin) return; commit();
    const p=JSON.parse(JSON.stringify(pin));
    p.id=id(); p.x_pct=Math.min(100,p.x_pct+2); p.y_pct=Math.min(100,p.y_pct+2);
    p.lastEdited=Date.now();
    const pg=currentPage(); if(pg){ pg.pins.push(p); }
    saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
  });
  on(btnDelete,'click',()=>{
    const pin=selectedPin(); if(!pin) return;
    if(!confirm('Delete selected pin?')) return; commit();
    const pg=currentPage(); if(!pg) return;
    pg.pins=pg.pins.filter(x=>x.id!==pin.id);
    saveProject(project); renderPins(); renderPinsList(); clearSelection();
  });

  on(toggleStrict,'change',()=>{ project.settings.strictRules = !!toggleStrict.checked; saveProject(project); renderWarnings(); });
  on(toggleField,'change',()=>{ project.settings.fieldMode = !!toggleField.checked; saveProject(project); renderPins(); });
  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; });
  on(inputSearch,'input',()=> renderPinsList());
  on(filterType,'change',()=> renderPins());

  // Stage interactions: add pin & dragging
  on(stage,'pointerdown',(e)=>{
    if(!addingPin) return;
    if(e.target && e.target.classList && e.target.classList.contains('pin')) return;
    const pct = toPctCoords(e);
    commit();
    const p = makePin(pct.x_pct, pct.y_pct);
    const pg=currentPage(); if(pg){ pg.pins.push(p); }
    saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
    addingPin=false; btnAddPin.classList.remove('ok');
  });

  // Global dragging listeners (document-level)
  document.addEventListener('pointermove', (e)=>{
    if(!dragging || !dragData) return;
    e.preventDefault();
    const pin = findPin(dragging.dataset.id); if(!pin) return;
    const x = Math.min(Math.max(dragData.rectLeft, e.clientX), dragData.rectLeft+dragData.rectW);
    const y = Math.min(Math.max(dragData.rectTop, e.clientY), dragData.rectTop+dragData.rectH);
    const pct = toPctCoordsFromPoint(x, y);
    dragging.style.left = pct.x_pct+'%';
    dragging.style.top = pct.y_pct+'%';
    posLabel.textContent = pctLabel(pct.x_pct, pct.y_pct);
  });

  document.addEventListener('pointerup', (e)=>{
    if(!dragging || !dragData) return;
    const pin = findPin(dragging.dataset.id);
    if(pin){
      const pct = toPctCoordsFromPoint(e.clientX, e.clientY);
      commit();
      pin.x_pct = pct.x_pct;
      pin.y_pct = pct.y_pct;
      pin.lastEdited = Date.now();
      saveProject(project); renderPinsList();
    }
    dragging.releasePointerCapture?.(e.pointerId);
    dragging=null; dragData=null;
  });

  // Measuring toolbar
  on(btnCalibrate,'click',()=> startCalibration('main'));
  on(btnMeasureToggle,'click',()=> toggleMeasuring('main'));
  on(btnMeasureReset,'click',()=> resetMeasurements('main'));

  function clearSelection(){ project._sel=null; saveProject(project); updatePinDetails(); }

  // Undo/Redo
  function snapshot(){ return JSON.stringify(project); }
  function loadSnapshot(s){
    project=JSON.parse(s);
    const i=projects.findIndex(p=>p.id===project.id);
    if(i>=0) projects[i]=project; else projects.push(project);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    localStorage.setItem('survey:lastOpenProjectId', project.id);
  }
  function commit(){ UNDO.push(snapshot()); if(UNDO.length>MAX_UNDO) UNDO.shift(); REDO.length=0; }
  function undo(){ if(!UNDO.length) return; REDO.push(snapshot()); const s=UNDO.pop(); loadSnapshot(s); renderAll(); }
  function redo(){ if(!REDO.length) return; UNDO.push(snapshot()); const s=REDO.pop(); loadSnapshot(s); renderAll(); }

  on(btnUndo,'click',undo);
  on(btnRedo,'click',redo);

  // Keyboard shortcuts
  window.addEventListener('keydown',(e)=>{
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); redo(); }
    if(e.key==='Delete'){ const p=selectedPin(); if(p){ commit(); const pg=currentPage(); if(pg){ pg.pins=pg.pins.filter(x=>x.id!==p.id); } saveProject(project); renderPins(); renderPinsList(); clearSelection(); } }
    if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)){
      const p=selectedPin(); if(!p) return; e.preventDefault(); commit();
      const delta = project.settings.fieldMode? 0.5 : 0.2;
      if(e.key==='ArrowUp') p.y_pct=Math.max(0,p.y_pct-delta);
      if(e.key==='ArrowDown') p.y_pct=Math.min(100,p.y_pct+delta);
      if(e.key==='ArrowLeft') p.x_pct=Math.max(0,p.x_pct-delta);
      if(e.key==='ArrowRight') p.x_pct=Math.min(100,p.x_pct+delta);
      p.lastEdited=Date.now(); saveProject(project); renderPins(); updatePinDetails();
    }
  });

  // Initial render & observers
  renderAll();
  const ro=new ResizeObserver(()=>{
    measureSvg.setAttribute('width', stage.clientWidth);
    measureSvg.setAttribute('height', stage.clientHeight);
  }); 
  if(stage) ro.observe(stage);
});
