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
  const gpsLabel = $('gpsLabel');
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
  let dragging = null;
  let measureMode = false;
  let calibFirst = null;
  let measureFirst = null;
  let fieldMode = false;
  let calibAwait = null; // 'main' | 'photo'

  // Photo modal state
  const photoState = { pin:null, idx:0, measure:false, calib:null };
  let photoMeasureFirst = null;

  /*******************
   * Small utilities *
   *******************/
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
    return { id:id(), name, createdAt:Date.now(), updatedAt:Date.now(), pages:[], settings:{ colorsByType:SIGN_TYPES, fieldMode:false, strictRules:false } };
  }
  function currentPage(){
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId);
  }
  function makePin(x_pct,y_pct){
    return { id:id(), sign_type:'', room_number:'', room_name:'', building: inputBuilding.value||projectContext.building||'', level: inputLevel.value||projectContext.level||'', x_pct, y_pct, notes:'', photos:[], lastEdited:Date.now() };
  }
  function findPin(idv){
    for(const pg of project.pages){ const f=(pg.pins||[]).find(p=>p.id===idv); if(f) return f; }
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
    project.pages.forEach((pg)=>{
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
      el.textContent=SHORT[p.sign_type]||p.sign_type.slice(0,3).toUpperCase();
      el.style.left=p.x_pct+'%';
      el.style.top=p.y_pct+'%';
      el.style.background=SIGN_TYPES[p.sign_type]||'#a3e635';
      el.style.padding = fieldMode? '.28rem .5rem' : '.18rem .35rem';
      el.style.fontSize= fieldMode? '0.9rem' : '0.75rem';
      pinLayer.appendChild(el);

      el.addEventListener('dblclick',()=>{ selectPin(p.id); openPhotoModal(); });
    });

    updatePinDetails();
  }

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
    fieldType.value=p?.sign_type||''; fieldRoomNum.value=p?.room_number||''; fieldRoomName.value=p?.room_name||''; fieldBuilding.value=p?.building||''; fieldLevel.value=p?.level||''; fieldNotes.value=p?.notes||'';
    posLabel.textContent=p? pctLabel(p.x_pct,p.y_pct) : '—';
    gpsLabel.textContent=(p&&p.gps)? `${p.gps.lat}, ${p.gps.lon}` : '—';
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
    list.forEach(s=>{ const t=document.createElement('span'); t.className='tag warn'; t.textContent=s; warnsEl.appendChild(t); });
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
    measureMode=false; $('btnMeasureToggle').textContent='Measuring: OFF';
    calibAwait=scope;
  }
  function toggleMeasuring(scope){
    measureMode=!measureMode;
    if(scope==='main') $('btnMeasureToggle').textContent = 'Measuring: '+(measureMode?'ON':'OFF');
  }
  function resetMeasurements(scope){
    if(scope==='main'){
      currentPage().measurements=[];
      drawMeasurements();
    }else{
      if(photoState.photo){ photoState.photo.measurements=[]; drawPhotoMeasurements(); }
    }
  }

  function drawMeasurements(){
    const page=currentPage();
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    if(!page||!page.measurements) return;
    page.measurements.forEach(m=>{
      const a=toScreen(m.points[0]); const b=toScreen(m.points[1]);
      const dx=Math.abs(a.x-b.x), dy=Math.abs(a.y-b.y);
      let color = (dx>dy)? '#ef4444' : (dy>dx)? '#3b82f6' : '#f59e0b';
      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y); line.setAttribute('x2',b.x); line.setAttribute('y2',b.y); line.setAttribute('stroke',color); measureSvg.appendChild(line);
      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2; const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      const ft = m.feet ? m.feet.toFixed(2)+' ft' : ((currentPage().scalePxPerFt)? (dist(a,b)/currentPage().scalePxPerFt).toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px'));
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=ft; measureSvg.appendChild(text);
    });
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
      if(page.scalePxPerFt){ m.feet = (dist(measureFirst, pt)/page.scalePxPerFt); }
      page.measurements = page.measurements||[]; page.measurements.push(m); measureFirst=null; drawMeasurements();
    }
  });

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
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y); line.setAttribute('x2',b.x); line.setAttribute('y2',b.y); line.setAttribute('stroke',color); photoMeasureSvg.appendChild(line);
      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2; const text=document.createElementNS('http://www.w3.org/2000/svg','text');
      const ft = m.feet ? m.feet.toFixed(2)+' ft' : (ph.scalePxPerFt? (dist(a,b)/ph.scalePxPerFt).toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px'));
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=ft; photoMeasureSvg.appendChild(text);
    });
  }

  on($('btnPhotoClose'),'click',()=> closePhotoModal());
  on($('btnPhotoMeasure'),'click',()=>{ photoState.measure=!photoState.measure; $('btnPhotoMeasure').textContent='Measure: '+(photoState.measure?'ON':'OFF'); });
  on($('btnPhotoCalib'),'click',()=>{ photoState.calib=null; alert('Photo calibration: click two points on the image.'); });
  on($('btnPhotoPrev'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on($('btnPhotoNext'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on($('btnPhotoDelete'),'click',()=>{ if(!photoState.pin) return; if(!confirm('Delete this photo?')) return; commit(); photoState.pin.photos.splice(photoState.idx,1); photoState.idx=Math.max(0,photoState.idx-1); saveProject(project); if(photoState.pin.photos.length===0){ closePhotoModal(); } else { showPhoto(); } });
  on($('btnPhotoDownload'),'click',()=>{ const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return; downloadFile(ph.name, dataURLtoBlob(ph.dataUrl)); });

  on($('photoMeasureSvg'),'click',(e)=>{
    const ph = photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };
    if(photoState.calib===null){ photoState.calib = pt; return; }
    if(photoState.calib && !photoState.measureTmp){
      const px = dist(photoState.calib, pt); const ft=parseFloat(prompt('Enter feet for calibration:','10'))||10;
      ph.measurements=ph.measurements||[]; ph.scalePxPerFt = px/ft; photoState.calib=null; drawPhotoMeasurements(); return;
    }
  });

  on(photoImg,'click',(e)=>{
    if(!photoState.measure) return; const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect(); const pt={x:e.clientX-rect.left,y:e.clientY-rect.top};
    if(!photoMeasureFirst){ photoMeasureFirst=pt; return; }
    const m={id:id(), kind:'photo', points:[photoMeasureFirst, pt]};
    if(ph.scalePxPerFt){ m.feet=dist(photoMeasureFirst,pt)/ph.scalePxPerFt; }
    ph.measurements=ph.measurements||[]; ph.measurements.push(m); photoMeasureFirst=null; drawPhotoMeasurements(); $('photoMeaCount').textContent=ph.measurements.length;
  });

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    project.pages.forEach(pg=> (pg.pins||[]).forEach(p=> rows.push({
      id:p.id, sign_type:p.sign_type||'', room_number:p.room_number||'', room_name:p.room_name||'',
      building:p.building||'', level:p.level||'', x_pct:fix(p.x_pct), y_pct:fix(p.y_pct),
      notes:p.notes||'', lat:p.gps?.lat||'', lon:p.gps?.lon||'', page_name:pg.name,
      last_edited: new Date(p.lastEdited||project.updatedAt).toISOString()
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
    XLSX.utils.book_append_sheet(wb, wsInfo, 'Project Info'); XLSX.utils.book_append_sheet(w
