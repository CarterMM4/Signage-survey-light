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
    'Exit':'#ef4444','Restroom':'#10b981','Custom':'#22c55e'
  };
  const SHORT = {
    'Elevator':'ELV','Hall Direct':'HD','Callbox':'CB','Evac':'EV','Ingress':'ING',
    'Egress':'EGR','Exit':'EXIT','Restroom':'WC','UNIT ID':'UNIT','FOH':'FOH','BOH':'BOH',
    'Stair - Ingress':'ING','Stair - Egress':'EGR','Custom':'CSTM'
  };

  /* Views */
  const viewHome = $('viewHome');
  const viewEditor = $('viewEditor');

  /* HOME refs */
  const homeGrid = $('homeGrid');
  const homeSearch = $('homeSearch');
  const btnNewProject = $('btnNewProject');

  /* Editor – left */
  const thumbsEl = $('thumbs');

  /* Editor – stage */
  const stage = $('stage');
  const stageImage = $('stageImage');
  const pinLayer = $('pinLayer');
  const measureSvg = $('measureSvg');

  /* Editor – toolbar */
  const projectLabel = $('projectLabel');
  const inputUpload = $('inputUpload');
  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');
  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleField = $('toggleField');

  /* Editor – right panel */
  const fieldType = $('fieldType');
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

  /* Photo modal */
  const photoModal = $('photoModal');
  const photoImg = $('photoImg');
  const photoMeasureSvg = $('photoMeasureSvg');
  const photoOverlay = $('photoOverlay');
  const photoPinId = $('photoPinId');
  const photoName = $('photoName');
  const photoMeaCount = $('photoMeaCount');
  const photoThumbRow = $('photoThumbRow');

  /************************
   * App state & Undo/redo
   ************************/
  let projects = [];
  let project = null;

  const UNDO = [];
  const REDO = [];
  const MAX_UNDO = 50;

  // Ephemeral context (defaults while adding)
  const projectContext = { building:'', level:'' };

  // Stage interaction state
  let addingPin = false;
  let draggingEl = null;
  let draggingPin = null;
  let measureMode = false;
  let calibFirst = null;
  let measureFirst = null;
  let fieldMode = false;
  let calibAwait = null; // 'main'

  // Photo modal state
  const photoState = { pin:null, idx:0, measuring:false, start:null }; // measuring for rubber-band

  /*******************
   * Small utilities *
   *******************/
  function id(){ return Math.random().toString(36).slice(2,10); }
  function fix(n){ return (typeof n==='number') ? +Number(n||0).toFixed(3) : ''; }
  function pctLabel(x,y){ return `${(x||0).toFixed(2)}%, ${(y||0).toFixed(2)}%`; }
  function dist(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return Math.sqrt(dx*dx+dy*dy); }

  function toPctCoords(clientX, clientY){
    const rect=stageImage.getBoundingClientRect();
    const x=Math.min(Math.max(0,clientX-rect.left),rect.width);
    const y=Math.min(Math.max(0,clientY-rect.top),rect.height);
    return { x_pct: +(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
  }
  function toLocal(clientX, clientY){
    const r=stageImage.getBoundingClientRect();
    return { x: clientX - r.left, y: clientY - r.top };
  }
  function toScreen(pt){
    const r=stageImage.getBoundingClientRect();
    return { x: r.left + pt.x, y: r.top + pt.y };
  }

  function csvEscape(v){ const s=(v==null? '': String(v)); return (/["\n,]/.test(s)) ? '"'+s.replace(/"/g,'""')+'"' : s; }
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
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 600);
  }

  /*********************
   * Data model helpers *
   *********************/
  function newProject(name){
    return {
      id:id(), name, createdAt:Date.now(), updatedAt:Date.now(),
      pages:[], settings:{ colorsByType:SIGN_TYPES, fieldMode:false }
    };
  }
  function currentPage(){
    if(!project) return null;
    if(!project._pageId && project.pages[0]) project._pageId=project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId) || null;
  }
  function makePin(x_pct, y_pct){
    return {
      id:id(),
      sign_type:'',
      label:'',            // custom display label
      room_number:'',
      room_name:'',
      building: inputBuilding.value || projectContext.building || '',
      level: inputLevel.value || projectContext.level || '',
      x_pct, y_pct,
      notes:'',
      photos:[],           // [{name,dataUrl,measurements:[], scalePxPerFt?}]
      lastEdited:Date.now()
    };
  }
  function findPin(idv){
    for(const pg of project.pages){
      const f=(pg.pins||[]).find(p=>p.id===idv);
      if(f) return f;
    }
    return null;
  }
  function checkedPinIds(){
    return [...pinList.querySelectorAll('input[type="checkbox"]:checked')].map(cb=>cb.dataset.id);
  }

  function selectProjectById(pid){
    const p = projects.find(x=>x.id===pid);
    if(!p) return;
    project = p;
    localStorage.setItem('survey:lastOpenProjectId', pid);
    switchView('editor');
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
    saveProject(project); renderAll();
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
    saveProject(project); renderAll();
  }

  /***************
   * View switch *
   ***************/
  function switchView(which){
    if(which==='home'){
      viewHome.classList.add('active'); viewEditor.classList.remove('active');
      renderHome();
    }else{
      viewEditor.classList.add('active'); viewHome.classList.remove('active');
    }
  }

  /****************
   * Home (hub)   *
   ****************/
  function renderHome(){
    const q=(homeSearch.value||'').toLowerCase();
    homeGrid.innerHTML='';
    const list = projects
      .slice()
      .sort((a,b)=> (b.updatedAt||0)-(a.updatedAt||0))
      .filter(p=> !q || p.name.toLowerCase().includes(q));
    if(!list.length){
      homeGrid.innerHTML = `<div class="muted">No projects yet. Click <strong>New Project</strong> to start.</div>`;
      return;
    }
    list.forEach(p=>{
      const totalPins = p.pages.reduce((a,pg)=>a+(pg.pins?.length||0),0);
      const card = document.createElement('div');
      card.className='card';
      card.innerHTML = `
        <header>
          <strong>${csvEscape(p.name)}</strong>
          <button class="btn acc open">Open</button>
        </header>
        <main>
          <div class="muted">Pages: ${p.pages.length} • Pins: ${totalPins}</div>
          <div class="muted">Updated: ${new Date(p.updatedAt).toLocaleString()}</div>
          <div class="muted">ID: ${p.id}</div>
        </main>
        <footer>
          <button class="btn rename">Rename</button>
          <button class="btn">Duplicate</button>
          <button class="btn ghost delete">Delete</button>
        </footer>
      `;
      card.querySelector('.open').onclick = ()=> selectProjectById(p.id);
      card.querySelector('.rename').onclick = ()=>{
        const name = prompt('New name:', p.name);
        if(!name) return;
        p.name=name; saveProject(p); renderHome();
      };
      card.querySelector('.delete').onclick = ()=>{
        if(!confirm('Delete this project?')) return;
        projects = projects.filter(x=>x.id!==p.id);
        localStorage.setItem('survey:projects', JSON.stringify(projects));
        renderHome();
      };
      card.querySelector('footer .btn:nth-child(2)').onclick = ()=>{
        const copy = JSON.parse(JSON.stringify(p));
        copy.id = id(); copy.name = p.name + ' (copy)';
        copy.createdAt = Date.now(); copy.updatedAt=Date.now();
        projects.push(copy);
        localStorage.setItem('survey:projects', JSON.stringify(projects));
        renderHome();
      };
      homeGrid.appendChild(card);
    });
  }

  on(btnNewProject,'click',()=>{
    const name = prompt('Project name','New Project');
    if(!name) return;
    const p = newProject(name);
    projects.push(p);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    selectProjectById(p.id);
  });

  on(homeSearch,'input',()=> renderHome());

  /****************
   * Rendering    *
   ****************/
  function renderAll(){
    renderTypeOptionsOnce();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    renderProjectLabel();
    drawMeasurements();
    toggleField.checked = !!project.settings.fieldMode;
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
    if(!project){ projectLabel.textContent='—'; return; }
    const pins = project.pages.reduce((a,p)=>a+(p.pins?.length||0),0);
    projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${pins}`;
  }

  function renderThumbs(){
    thumbsEl.innerHTML='';
    if(!project) return;
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
    fieldMode = !!project.settings.fieldMode;
    const q=(inputSearch.value||'').toLowerCase();
    const typeFilter=filterType.value;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type!==typeFilter) return;
      const line=[p.sign_type,p.label,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el=document.createElement('div');
      el.className='pin'+(p.id===project._sel?' selected':'');
      el.dataset.id=p.id;
      const tag = (p.label && p.label.trim()) ? p.label.trim() : (SHORT[p.sign_type] || (p.sign_type||'-').slice(0,3).toUpperCase());
      el.textContent = tag;
      el.style.left=p.x_pct+'%';
      el.style.top=p.y_pct+'%';
      el.style.background=SIGN_TYPES[p.sign_type]||'#22c55e';
      el.style.padding = fieldMode? '.32rem .55rem' : '.22rem .45rem';
      el.style.fontSize= fieldMode? '0.95rem' : '0.8rem';
      pinLayer.appendChild(el);

      // Start drag
      el.addEventListener('pointerdown',(e)=>{
        draggingEl = el;
        draggingPin = p;
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
      row.style.display='grid';
      row.style.gridTemplateColumns='20px 1fr auto';
      row.style.alignItems='center';
      row.style.gap='.5rem';
      row.style.padding='.35rem .45rem';
      row.style.border='1px solid #1e2a2d';
      row.style.borderRadius='10px';
      row.style.background='#0f1b1e';

      const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; row.appendChild(cb);
      const txt=document.createElement('div');
      const tag = (p.label && p.label.trim()) ? p.label.trim() : (p.sign_type||'-');
      txt.innerHTML=`<strong>${tag}</strong> • ${p.room_number||'-'} ${p.room_name||''} <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
      row.appendChild(txt);
      const go=document.createElement('button'); go.className='btn'; go.textContent='Go'; go.onclick=()=>{ selectPin(p.id); }; row.appendChild(go);
      pinList.appendChild(row);
    });
  }

  function updatePinDetails(){
    const p = selectedPin();
    selId.textContent = p ? p.id : 'None';
    fieldType.value = p?.sign_type || '';
    fieldLabel.value = p?.label || '';
    fieldRoomNum.value = p?.room_number || '';
    fieldRoomName.value = p?.room_name || '';
    fieldBuilding.value = p?.building || '';
    fieldLevel.value = p?.level || '';
    fieldNotes.value = p?.notes || '';
    posLabel.textContent = p ? pctLabel(p.x_pct,p.y_pct) : '—';
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
      const t=document.createElement('div');
      t.className='muted'; t.textContent='• '+s;
      warnsEl.appendChild(t);
    });
  }

  function selectedPin(){
    const idv=project? project._sel : null;
    if(!idv) return null;
    return findPin(idv);
  }

  function selectPin(idv){
    project._sel=idv; saveProject(project); renderPins();
    const el=[...pinLayer.children].find(x=>x.dataset.id===idv);
    if(el){ el.scrollIntoView({block:'center', inline:'center', behavior:'smooth'}); }
    const p=findPin(idv);
    if(p){
      const pgId=project.pages.find(pg=> (pg.pins||[]).includes(p))?.id;
      if(pgId && project._pageId!==pgId){ project._pageId=pgId; saveProject(project); renderStage(); renderPins(); drawMeasurements(); }
    }
  }

  /***********************
   * Measuring (main view)
   ***********************/
  function startCalibration(scope){
    calibFirst=null; measureFirst=null;
    alert('Calibration: click two points on the stage to define real feet.');
    measureMode=false; $('btnMeasureToggle').textContent='Measuring: OFF';
    calibAwait=scope;
  }
  function toggleMeasuring(scope){
    measureMode=!measureMode;
    if(scope==='main') $('btnMeasureToggle').textContent = 'Measuring: '+(measureMode?'ON':'OFF');
  }
  function resetMeasurements(scope){
    if(scope==='main'){
      const pg=currentPage(); if(!pg) return;
      pg.measurements=[]; drawMeasurements(); saveProject(project);
    }
  }

  function drawMeasurements(){
    const page=currentPage();
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    if(!page||!page.measurements) return;
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
      const ft = m.feet ? m.feet.toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px');
      text.setAttribute('x',midx); text.setAttribute('y',midy-6); text.textContent=ft; measureSvg.appendChild(text);
    });
  }

  /* Stage clicks for measure & add pin */
  on(stage,'click',(e)=>{
    const pg=currentPage(); if(!pg) return;
    const pt=toLocal(e.clientX, e.clientY);
    if(calibAwait==='main'){
      if(!calibFirst){ calibFirst=pt; return; }
      const px = dist(calibFirst, pt);
      const ft = parseFloat(prompt('Enter real distance (feet):','10')) || 10;
      pg.scalePxPerFt = px/ft; calibFirst=null; calibAwait=null;
      alert('Calibrated: '+(px/ft).toFixed(2)+' px/ft');
      return;
    }
    if(measureMode){
      if(!measureFirst){ measureFirst=pt; return; }
      const m={id:id(), points:[measureFirst, pt], feet:null};
      if(pg.scalePxPerFt){ m.feet = dist(measureFirst,pt)/pg.scalePxPerFt; }
      pg.measurements = pg.measurements||[]; pg.measurements.push(m); measureFirst=null; drawMeasurements(); saveProject(project);
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
  function closePhotoModal(){ photoModal.style.display='none'; photoState.measuring=false; photoState.start=null; clearPhotoOverlay(); }

  function clearPhotoOverlay(){
    while(photoOverlay.firstChild) photoOverlay.removeChild(photoOverlay.firstChild);
  }
  function drawPhotoOverlayLine(a,b){
    clearPhotoOverlay();
    const line=document.createElementNS('http://www.w3.org/2000/svg','line');
    line.setAttribute('x1',a.x); line.setAttribute('y1',a.y);
    line.setAttribute('x2',b.x); line.setAttribute('y2',b.y);
    line.setAttribute('stroke','#22c55e'); line.setAttribute('stroke-width','3');
    photoOverlay.appendChild(line);
  }

  function showPhoto(){
    const ph=photoState.pin.photos[photoState.idx];
    photoImg.src=ph.dataUrl;
    photoPinId.textContent=photoState.pin.id;
    photoName.textContent=ph.name;
    photoMeaCount.textContent=(ph.measurements?.length||0);
    drawPhotoMeasurements();
    renderPhotoThumbs();
  }

  function renderPhotoThumbs(){
    photoThumbRow.innerHTML='';
    const arr = photoState.pin.photos || [];
    arr.forEach((ph, i)=>{
      const im = document.createElement('img');
      im.className = 'phThumb'+(i===photoState.idx?' active':'');
      im.src = ph.dataUrl;
      im.onclick = ()=>{ photoState.idx = i; showPhoto(); };
      photoThumbRow.appendChild(im);
    });
  }

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
      const ft = m.feet!=null ? m.feet.toFixed(2)+' ft'
               : (ph.scalePxPerFt ? (dist(a,b)/ph.scalePxPerFt).toFixed(2)+' ft' : (dist(a,b).toFixed(0)+' px'));
      text.setAttribute('x',midx); text.setAttribute('y',midy-6);
      text.textContent=ft; photoMeasureSvg.appendChild(text);
    });
  }

  /* Rubber-band measurement on photo */
  function photoLocalFromEvent(ev){
    const r=photoImg.getBoundingClientRect();
    return { x: ev.clientX - r.left + photoImg.scrollLeft, y: ev.clientY - r.top + photoImg.scrollTop };
  }

  on($('btnPhotoMeasure'),'click',()=>{
    photoState.measuring=!photoState.measuring;
    $('btnPhotoMeasure').textContent = 'Measure: ' + (photoState.measuring?'ON':'OFF');
    photoState.start=null; clearPhotoOverlay();
  });
  on($('btnPhotoCalib'),'click',()=>{
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    ph.scalePxPerFt = null; alert('Calibration cleared. Draw a line and enter a known length to re-calibrate.');
  });
  on($('btnPhotoClose'),'click',()=> closePhotoModal());
  on($('btnPhotoPrev'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on($('btnPhotoNext'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on($('btnPhotoDelete'),'click',()=>{ const pin=photoState.pin; if(!pin) return; if(!confirm('Delete this photo?')) return; const arr=pin.photos; arr.splice(photoState.idx,1); photoState.idx=Math.max(0,photoState.idx-1); saveProject(project); if(!arr.length){ closePhotoModal(); } else { showPhoto(); } });
  on($('btnPhotoDownload'),'click',()=>{ const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return; downloadFile(ph.name||'photo.png', dataURLtoBlob(ph.dataUrl)); });

  // Mouse interactions on viewer
  photoImg.addEventListener('mousedown',(e)=>{
    if(!photoState.measuring) return;
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    photoState.start = photoLocalFromEvent(e);
    drawPhotoOverlayLine(photoState.start, photoState.start);
  });
  photoImg.addEventListener('mousemove',(e)=>{
    if(!photoState.measuring || !photoState.start) return;
    const cur = photoLocalFromEvent(e);
    drawPhotoOverlayLine(photoState.start, cur);
  });
  photoImg.addEventListener('mouseup',(e)=>{
    if(!photoState.measuring || !photoState.start) return;
    const ph=photoState.pin?.photos[photoState.idx]; if(!ph) return;
    const end = photoLocalFromEvent(e);
    const px = dist(photoState.start, end);

    // Ask for real length (ft) — this both stores the measure and sets/updates calibration
    const entered = prompt('Enter measured length (feet) for this line (e.g., 6.5):','');
    let feet = null;
    if(entered!=null && entered.trim()!==''){
      const val = parseFloat(entered);
      if(!isNaN(val) && val>0){
        feet = val;
        ph.scalePxPerFt = px / feet; // (re)calibrate from this entry
      }
    }
    // Store measurement
    ph.measurements = ph.measurements || [];
    ph.measurements.push({ id:id(), points:[photoState.start, end], feet: feet });
    saveProject(project);
    clearPhotoOverlay();
    photoState.start=null;
    drawPhotoMeasurements();
    photoMeaCount.textContent = ph.measurements.length;
  });

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    project.pages.forEach(pg=> (pg.pins||[]).forEach(p=>{
      rows.push({
        id:p.id,
        sign_type:p.sign_type||'',
        label:p.label||'',
        room_number:p.room_number||'',
        room_name:p.room_name||'',
        building:p.building||'',
        level:p.level||'',
        x_pct:fix(p.x_pct),
        y_pct:fix(p.y_pct),
        notes:p.notes||'',
        page_name:pg.name,
        last_edited:new Date(p.lastEdited||project.updatedAt).toISOString()
      });
    }));
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
    project.pages.forEach(pg=> (pg.pins||[]).forEach(pin=>{
      (pin.photos||[]).forEach((ph, idx)=>{
        const folder=zip.folder(`photos/${pin.id}`);
        folder.file(ph.name||`photo_${idx+1}.png`, dataURLtoArrayBuffer(ph.dataUrl));
      });
    }));
    const blob = await zip.generateAsync({type:'blob'});
    downloadFile(project.name.replace(/\W+/g,'_')+'_export.zip', blob);
  }

  function importCSV(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2) return;
    const hdr=lines[0].split(',').map(h=>h.trim()); commit();
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]); const row={}; hdr.forEach((h,idx)=> row[h]=cells[idx]||'');
      const p = makePin(parseFloat(row.x_pct)||50, parseFloat(row.y_pct)||50);
      p.sign_type=row.sign_type||''; p.label=row.label||''; p.room_number=row.room_number||''; p.room_name=row.room_name||'';
      p.building=row.building||''; p.level=row.level||''; p.notes=row.notes||'';
      currentPage().pins.push(p);
    }
    saveProject(project); renderPins(); renderPinsList(); alert('Imported rows into current page.');
  }

  /****************
   * Toolbar hooks
   ****************/
  on($('btnBackHome'),'click',()=> switchView('home'));
  on($('btnRename'),'click',()=>{
    const name=prompt('Rename project:', project.name);
    if(!name) return;
    commit(); project.name=name; saveProject(project); renderProjectLabel(); renderHome();
  });
  on($('btnSaveAs'),'click',()=>{
    const name=prompt('Duplicate name:', project.name+' (copy)');
    if(!name) return;
    commit();
    const copy=JSON.parse(JSON.stringify(project));
    copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now();
    saveProject(copy);
    alert('Duplicated as "'+name+'". Find it in Projects.');
  });

  on($('btnUpload'),'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){ const url=URL.createObjectURL(f); const name=f.name.replace(/\.[^.]+$/,''); addImagePage(url,name); }
    }
  });

  on($('btnExportCSV'),'click',()=> exportCSV());
  on($('btnExportXLSX'),'click',()=> exportXLSX());
  on($('btnExportZIP'),'click',()=> exportZIP());
  on($('btnImportCSV'),'click',()=> $('inputImportCSV').click());
  on($('inputImportCSV'),'change', async (e)=>{ const f=e.target.files?.[0]; if(!f) return; const text=await f.text(); importCSV(text); });
  on($('btnOCR'),'click',()=> ocrCurrentView());

  const btnCalibrate = $('btnCalibrate');
  const btnMeasureToggle = $('btnMeasureToggle');
  const btnMeasureReset = $('btnMeasureReset');
  on(btnCalibrate,'click',()=> startCalibration('main'));
  on(btnMeasureToggle,'click',()=> toggleMeasuring('main'));
  on(btnMeasureReset,'click',()=> resetMeasurements('main'));

  on($('btnUndo'),'click',()=> undo());
  on($('btnRedo'),'click',()=> redo());

  on($('btnClearPins'),'click',()=>{ if(!confirm('Clear ALL pins on this page?')) return; commit(); currentPage().pins=[]; saveProject(project); renderAll(); });
  on($('btnAddPin'),'click',()=> startAddPin());

  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; });

  on(inputSearch,'input',()=> renderPinsList());
  on(filterType,'change',()=> renderPins());
  on(toggleField,'change',()=>{ project.settings.fieldMode = !!toggleField.checked; saveProject(project); renderPins(); });

  // Right fields update
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

  on($('btnAddPhoto'),'click',()=> $('inputPhoto').click());
  on($('inputPhoto'),'change', async (e)=>{
    const pin = selectedPin(); if(!pin) return alert('Select a pin first.');
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){ const url=await fileToDataURL(f); pin.photos.push({name:f.name,dataUrl:url,measurements:[]}); }
    pin.lastEdited=Date.now(); saveProject(project); renderPinsList(); alert('Photo(s) added.');
  });

  on($('btnOpenPhoto'),'click',()=> openPhotoModal());
  on($('btnDuplicate'),'click',()=>{ const pin=selectedPin(); if(!pin) return; commit(); const p=JSON.parse(JSON.stringify(pin)); p.id=id(); p.x_pct=Math.min(100, pin.x_pct+2); p.y_pct=Math.min(100, pin.y_pct+2); p.lastEdited=Date.now(); currentPage().pins.push(p); saveProject(project); renderPins(); renderPinsList(); selectPin(p.id); });
  on($('btnDelete'),'click',()=>{ const pin=selectedPin(); if(!pin) return; if(!confirm('Delete selected pin?')) return; commit(); const pg=currentPage(); pg.pins=(pg.pins||[]).filter(x=>x.id!==pin.id); saveProject(project); renderPins(); renderPinsList(); project._sel=null; updatePinDetails(); });

  function startAddPin(){
    addingPin=!addingPin;
    $('btnAddPin').classList.toggle('acc', addingPin);
    if(addingPin){ alert('Click on the page to place a new pin.'); }
  }

  /********************
   * Stage interactions
   ********************/
  stage.addEventListener('pointerdown',(e)=>{
    if(e.target.classList.contains('pin')) return;
    if(!addingPin) return;
    const local = toPctCoords(e.clientX, e.clientY);
    commit();
    const p = makePin(local.x_pct, local.y_pct);
    const pg=currentPage();
    pg.pins = pg.pins || [];
    pg.pins.push(p);
    saveProject(project); renderPins(); renderPinsList(); selectPin(p.id);
    addingPin=false; $('btnAddPin').classList.remove('acc');
  });

  // Dragging pins
  pinLayer.addEventListener('pointermove',(e)=>{
    if(!draggingEl || !draggingPin) return;
    e.preventDefault();
    const pc = toPctCoords(e.clientX, e.clientY);
    draggingEl.style.left = pc.x_pct+'%';
    draggingEl.style.top = pc.y_pct+'%';
    posLabel.textContent = pctLabel(pc.x_pct, pc.y_pct);
  });
  pinLayer.addEventListener('pointerup',(e)=>{
    if(!draggingEl || !draggingPin) return;
    const pc = toPctCoords(e.clientX, e.clientY);
    commit();
    draggingPin.x_pct = pc.x_pct;
    draggingPin.y_pct = pc.y_pct;
    draggingPin.lastEdited = Date.now();
    saveProject(project); renderPinsList();
    draggingEl.releasePointerCapture?.(e.pointerId);
    draggingEl=null; draggingPin=null;
    updatePinDetails();
  });

  /****************
   * OCR (stage)  *
   ****************/
  async function ocrCurrentView(){
    const canvas = document.createElement('canvas');
    const img=stageImage; if(!img || !img.src){ alert('No page image.'); return; }
    const w=img.naturalWidth||img.width, h=img.naturalHeight||img.height;
    canvas.width=w; canvas.height=h;
    const ctx=canvas.getContext('2d'); ctx.drawImage(img,0,0);
    const { data: { text } } = await Tesseract.recognize(canvas, 'eng');
    let out = (text||'').trim(); if(!out){ alert('No text recognized.'); return; }
    const head = out.slice(0,200);
    try { await navigator.clipboard.writeText(out); } catch{}
    inputSearch.value=head; renderPinsList(); alert('OCR done. First 200 chars placed in search. Full text copied to clipboard.');
  }

  /*************
   * Undo/Redo *
   *************/
  function snapshot(){ return JSON.stringify(project); }
  function loadSnapshot(s){
    const restored=JSON.parse(s);
    const i=projects.findIndex(p=>p.id===restored.id);
    if(i>=0) projects[i]=restored; else projects.push(restored);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    localStorage.setItem('survey:lastOpenProjectId', restored.id);
    project = restored;
  }
  function commit(){ UNDO.push(snapshot()); if(UNDO.length>MAX_UNDO) UNDO.shift(); REDO.length=0; }
  function undo(){ if(!UNDO.length) return; REDO.push(snapshot()); const s=UNDO.pop(); loadSnapshot(s); renderAll(); }
  function redo(){ if(!REDO.length) return; UNDO.push(snapshot()); const s=REDO.pop(); loadSnapshot(s); renderAll(); }

  // Keyboard shortcuts
  window.addEventListener('keydown',(e)=>{
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); redo(); }
    if(e.key==='Delete'){ const p=selectedPin(); if(p){ commit(); const pg=currentPage(); pg.pins=(pg.pins||[]).filter(x=>x.id!==p.id); saveProject(project); renderPins(); renderPinsList(); project._sel=null; updatePinDetails(); } }
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

  /********************
   * Initial bootstrap *
   ********************/
  projects = loadProjects();
  const last = localStorage.getItem('survey:lastOpenProjectId');
  if(last && projects.find(p=>p.id===last)){
    selectProjectById(last);
  }else{
    switchView('home');
  }

  // Buttons in editor toolbar that rely on DOM after setup
  on($('btnOpenPhoto'),'click',()=> openPhotoModal());
});
