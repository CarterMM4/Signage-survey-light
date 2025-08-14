// app.js
document.addEventListener('DOMContentLoaded', () => {
  /***********************
   * Helpers & constants *
   ***********************/
  function $(id){ return document.getElementById(id); }
  function on(el, evt, fn){ if(el) el.addEventListener(evt, fn); }

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
  const BRAND_FALLBACK = '#43d68a';

  /***********************
   * State & references  *
   ***********************/
  // Screens
  const screenHome = $('screen-home');
  const screenProjects = $('screen-projects');
  const screenEditor = $('screen-editor');

  // Home
  const homeNew = $('homeNew');
  const homeOpenList = $('homeOpenList');
  const homeImportCsv = $('homeImportCsv');
  const inputImportCSV = $('inputImportCSV');

  // Projects screen
  const projectList = $('projectList');
  const plNew = $('plNew');
  const plRename = $('plRename');
  const plDelete = $('plDelete');
  const plExport = $('plExport');
  const plImport = $('plImport');
  const plImportBtn = $('plImportBtn');

  // Editor refs
  const thumbsEl = $('thumbs');
  const stage = $('stage');
  const stageImage = $('stageImage');
  const pinLayer = $('pinLayer');
  const measureSvg = $('measureSvg');

  const inputUpload = $('inputUpload');
  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');
  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleField = $('toggleField');
  const projectLabel = $('projectLabel');

  const fieldType = $('fieldType'); // input with datalist
  const typeList = $('typeList');
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

  // Top nav
  on($('btnHome'), 'click', () => showScreen('home'));
  on($('btnProjects'), 'click', () => { renderProjectList(); showScreen('projects'); });

  // Screen nav buttons
  on(homeNew, 'click', () => { doNewProject(); showScreen('editor'); });
  on(homeOpenList, 'click', () => { renderProjectList(); showScreen('projects'); });
  on(homeImportCsv, 'click', () => inputImportCSV.click());
  on(inputImportCSV, 'change', async (e) => {
    const f = e.target.files && e.target.files[0]; if(!f) return;
    const text = await f.text(); importCSV(text); e.target.value='';
  });

  on(plNew, 'click', () => { doNewProject(); showScreen('editor'); });
  on(plRename, 'click', () => promptRenameProject());
  on(plDelete, 'click', () => deleteProject());
  on(plExport, 'click', () => exportAllProjectsJSON());
  on(plImportBtn, 'click', () => plImport.click());
  on(plImport, 'change', async (e) => {
    const f = e.target.files && e.target.files[0]; if(!f) return;
    try {
      const txt = await f.text();
      const imported = JSON.parse(txt);
      if(Array.isArray(imported)){
        localStorage.setItem('survey:projects', JSON.stringify(imported));
        projects = loadProjects();
        renderProjectList();
        alert('Imported projects JSON.');
      } else alert('Invalid JSON format.');
    } catch(err){ alert('Failed to import: '+err); }
    e.target.value='';
  });

  // Toolbar (editor)
  on($('btnNew'),'click',()=> doNewProject());
  on($('btnOpenList'),'click',()=> { renderProjectList(); showScreen('projects'); });
  on($('btnSaveAs'),'click',()=> duplicateProject());
  on($('btnRename'),'click',()=> promptRenameProject());

  on($('btnUpload'),'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    const files=[...e.target.files]; if(!files.length) return; commit();
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){ const url=URL.createObjectURL(f); const name=f.name.replace(/\.[^.]+$/,''); addImagePage(url, name); }
    }
    renderAll();
    inputUpload.value='';
  });

  on($('btnExportCSV'),'click',()=> exportCSV());
  on($('btnExportXLSX'),'click',()=> exportXLSX());
  on($('btnExportZIP'),'click',()=> exportZIP());

  on($('btnOCR'),'click',()=> ocrCurrentView());
  on($('btnCalibrate'),'click',()=> startCalibration('main'));
  on($('btnMeasureToggle'),'click',()=> toggleMeasuring('main'));
  on($('btnMeasureReset'),'click',()=> resetMeasurements('main'));

  on($('btnUndo'),'click',()=> undo());
  on($('btnRedo'),'click',()=> redo());

  on($('btnClearPins'),'click',()=>{ if(!currentPage()) return; if(!confirm('Clear ALL pins on this page?')) return; commit(); currentPage().pins=[]; renderAll(); });
  on($('btnAddPin'),'click',()=> startAddPin());

  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; saveProject(project); });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; saveProject(project); });

  on(inputSearch,'input',()=> renderPinsList());
  on(filterType,'change',()=> renderPins());

  // Right panel fields
  ;[fieldType, fieldRoomNum, fieldRoomName, fieldBuilding, fieldLevel, fieldNotes].forEach(el => on(el,'input',()=>{
    const pin = selectedPin(); if(!pin) return; commit();
    if(el===fieldType) pin.sign_type = el.value || '';
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
    e.target.value = '';
  });

  on($('btnOpenPhoto'),'click',()=> openPhotoModal());
  on($('btnDuplicate'),'click',()=>{ const pin=selectedPin(); if(!pin) return; commit(); const p=JSON.parse(JSON.stringify(pin)); p.id=id(); p.x_pct = Math.min(100, (pin.x_pct||0)+2); p.y_pct=Math.min(100,(pin.y_pct||0)+2); p.lastEdited=Date.now(); currentPage().pins.push(p); saveProject(project); renderPins(); renderPinsList(); selectPin(p.id); });
  on($('btnDelete'),'click',()=>{ const pin=selectedPin(); if(!pin) return; if(!confirm('Delete selected pin?')) return; commit(); currentPage().pins=currentPage().pins.filter(x=>x.id!==pin.id); saveProject(project); renderPins(); renderPinsList(); clearSelection(); });

  on($('btnBulkType'),'click',()=>{ const type = prompt('Enter type for selected pins:'); if(type==null) return; commit(); checkedPinIds().forEach(idv=>{ const p=findPin(idv); if(p){ p.sign_type=type; p.lastEdited=Date.now(); }}); saveProject(project); renderPins(); renderPinsList(); });
  on($('btnBulkBL'),'click',()=>{ const b = prompt('Building value (blank keep):', inputBuilding.value||''); if(b==null) return; const l = prompt('Level value (blank keep):', inputLevel.value||''); if(l==null) return; commit(); checkedPinIds().forEach(idv=>{ const p=findPin(idv); if(p){ if(b!=='') p.building=b; if(l!=='') p.level=l; p.lastEdited=Date.now(); }}); saveProject(project); renderPins(); renderPinsList(); });

  // Stage interactions
  let addingPin=false;
  let dragging=null;
  let measureMode=false;
  let calibFirst=null;
  let measureFirst=null;
  let fieldMode=false;
  let calibAwait=null; // 'main'|'photo'

  on(stage,'pointerdown',(e)=>{
    // add-pin only: click anywhere that isn't a pin
    if(e.target.classList.contains('pin')) return;
    if(!addingPin) return;
    const pt = toPctCoords(e);
    commit();
    const newp = makePin(pt.x_pct, pt.y_pct);
    currentPage().pins.push(newp); saveProject(project);
    renderPins(); renderPinsList(); selectPin(newp.id);
    addingPin=false; $('btnAddPin').classList.remove('btn-brand');
  });

  on(pinLayer,'pointerdown',(e)=>{
    const el = e.target.closest('.pin'); if(!el) return;
    selectPin(el.dataset.id);
    dragging = el;
    if(dragging.setPointerCapture) dragging.setPointerCapture(e.pointerId);
  });
  on(pinLayer,'pointermove',(e)=>{
    if(!dragging) return; e.preventDefault();
    const pin = findPin(dragging.dataset.id); if(!pin) return;
    const pt = toPctCoords(e);
    dragging.style.left = pt.x_pct+'%';
    dragging.style.top  = pt.y_pct+'%';
    posLabel.textContent = pctLabel(pt.x_pct, pt.y_pct);
  });
  on(pinLayer,'pointerup',(e)=>{
    if(!dragging) return;
    const pin = findPin(dragging.dataset.id);
    if(pin){
      commit();
      const pt = toPctCoords(e);
      pin.x_pct = pt.x_pct;
      pin.y_pct = pt.y_pct;
      pin.lastEdited = Date.now();
      saveProject(project);
      renderPinsList();
    }
    if(dragging.releasePointerCapture) dragging.releasePointerCapture(e.pointerId);
    dragging = null;
  });

  function startAddPin(){
    addingPin = !addingPin;
    $('btnAddPin').classList.toggle('btn-brand', addingPin);
  }

  /***********************
   * Measurement (main)  *
   ***********************/
  function startCalibration(scope){
    calibFirst=null; measureFirst=null; measureMode=false;
    $('btnMeasureToggle').textContent='Measuring: OFF';
    calibAwait=scope;
    alert('Calibration: click two points on the page to define a known distance, then enter real feet.');
  }
  function toggleMeasuring(scope){
    measureMode = !measureMode;
    if(scope==='main') $('btnMeasureToggle').textContent = 'Measuring: ' + (measureMode?'ON':'OFF');
  }
  function resetMeasurements(scope){
    if(scope==='main'){
      const page=currentPage(); if(!page) return; page.measurements=[]; drawMeasurements();
    } else {
      const ph = photoState.pin && photoState.pin.photos[photoState.idx];
      if(ph){ ph.measurements=[]; drawPhotoMeasurements(); }
    }
  }

  on(stage,'click',(e)=>{
    const page=currentPage(); if(!page) return;
    const pt=toLocal(e);
    if(calibAwait==='main'){
      if(!calibFirst){ calibFirst=pt; return; }
      // Second click; prompt real feet
      const px = dist(calibFirst, pt);
      const ft = parseFloat(prompt('Enter real distance (feet):','10')) || 10;
      page.scalePxPerFt = px/ft;
      calibFirst=null; calibAwait=null;
      alert('Calibrated at '+(px/ft).toFixed(2)+' px/ft');
      return;
    }
    if(measureMode){
      // rubberband measurement on main page
      if(!measureFirst){ measureFirst=pt; drawMeasurements(pt); return; }
      const m={ id:id(), kind:'main', points:[measureFirst, pt] };
      // Ask user for their measured feet (site note)
      const userFt = prompt('Enter measured distance (feet). Leave blank to auto from calibration:','');
      if(userFt && !isNaN(parseFloat(userFt))){
        m.feet = parseFloat(userFt);
      } else if(page.scalePxPerFt){
        m.feet = dist(measureFirst, pt)/page.scalePxPerFt;
      }
      page.measurements = page.measurements || [];
      page.measurements.push(m);
      measureFirst=null;
      drawMeasurements();
    }
  });

  function drawMeasurements(livePoint){
    // Clear
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);

    const page=currentPage();
    if(!page) return;

    // Persisted
    if(page.measurements){
      page.measurements.forEach(m=>{
        const a = toScreen(m.points[0]);
        const b = toScreen(m.points[1]);
        drawSvgLine(measureSvg, a, b);
        const mid = { x:(a.x+b.x)/2, y:(a.y+b.y)/2 };
        const label = (typeof m.feet==='number')
          ? m.feet.toFixed(2)+' ft'
          : (page.scalePxPerFt ? (dist(m.points[0], m.points[1])/page.scalePxPerFt).toFixed(2)+' ft' : Math.round(dist(a,b))+' px');
        drawSvgText(measureSvg, mid.x, mid.y-6, label);
      });
    }

    // Live rubberband
    if(measureFirst && livePoint){
      const a = toScreen(measureFirst);
      const b = toScreen(livePoint);
      drawSvgLine(measureSvg, a, b, true);
    }
  }

  function drawSvgLine(svg, a, b, dashed){
    const line=document.createElementNS('http://www.w3.org/2000/svg','line');
    line.setAttribute('x1', a.x); line.setAttribute('y1', a.y);
    line.setAttribute('x2', b.x); line.setAttribute('y2', b.y);
    line.setAttribute('stroke', dashed ? '#43d68a' : (Math.abs(a.x-b.x)>Math.abs(a.y-b.y)? '#ef4444' : '#4cc9f0'));
    line.setAttribute('stroke-width', '3');
    if(dashed) line.setAttribute('stroke-dasharray','6 6');
    svg.appendChild(line);
  }
  function drawSvgText(svg, x, y, textVal){
    const t=document.createElementNS('http://www.w3.org/2000/svg','text');
    t.setAttribute('x', x); t.setAttribute('y', y);
    t.textContent = textVal;
    svg.appendChild(t);
  }

  /***********************
   * Photo modal measure *
   ***********************/
  const photoState = { pin:null, idx:0, measure:false, calib:null };
  let photoMeasureFirst = null;

  on($('btnPhotoClose'),'click',()=> closePhotoModal());
  on($('btnPhotoMeasure'),'click',()=>{ photoState.measure=!photoState.measure; $('btnPhotoMeasure').textContent='Measure: '+(photoState.measure?'ON':'OFF'); });
  on($('btnPhotoCalib'),'click',()=>{ photoState.calib=null; alert('Photo calibration: click two points, then enter real feet.'); });
  on($('btnPhotoPrev'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.max(0,photoState.idx-1); showPhoto(); });
  on($('btnPhotoNext'),'click',()=>{ if(!photoState.pin) return; photoState.idx=Math.min(photoState.pin.photos.length-1,photoState.idx+1); showPhoto(); });
  on($('btnPhotoDelete'),'click',()=>{ if(!photoState.pin) return; if(!confirm('Delete this photo?')) return; commit(); photoState.pin.photos.splice(photoState.idx,1); photoState.idx=Math.max(0,photoState.idx-1); saveProject(project); if(photoState.pin.photos.length===0){ closePhotoModal(); } else { showPhoto(); } });
  on($('btnPhotoDownload'),'click',()=>{ const ph=photoState.pin && photoState.pin.photos[photoState.idx]; if(!ph) return; downloadFile(ph.name, dataURLtoBlob(ph.dataUrl)); });

  on($('photoMeasureSvg'),'pointermove',(e)=>{
    if(!photoState.measure || !photoMeasureFirst) return;
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };
    drawPhotoMeasurements(pt); // live rubberband
  });
  on($('photoMeasureSvg'),'click',(e)=>{
    const ph = photoState.pin && photoState.pin.photos[photoState.idx]; if(!ph) return;
    const rect = photoImg.getBoundingClientRect();
    const pt = { x: e.clientX-rect.left, y:e.clientY-rect.top };

    // calibration flow
    if(photoState.calib===null && !photoState.measure){ photoState.calib = pt; return; }
    if(photoState.calib && !photoState.measure){ // second click for calib
      const px = dist(photoState.calib, pt);
      const ft=parseFloat(prompt('Enter real feet for calibration:','10'))||10;
      ph.scalePxPerFt = px/ft; photoState.calib=null; drawPhotoMeasurements(); return;
    }

    // measure flow (with rubberband)
    if(photoState.measure){
      if(!photoMeasureFirst){ photoMeasureFirst = pt; drawPhotoMeasurements(pt); return; }
      const m = { id:id(), kind:'photo', points:[photoMeasureFirst, pt] };
      const userFt = prompt('Enter measured distance (feet). Leave blank to compute from calibration:','');
      if(userFt && !isNaN(parseFloat(userFt))){
        m.feet = parseFloat(userFt);
      } else if(ph.scalePxPerFt){
        m.feet = dist(photoMeasureFirst, pt)/ph.scalePxPerFt;
      }
      ph.measurements = ph.measurements || [];
      ph.measurements.push(m);
      photoMeasureFirst = null;
      drawPhotoMeasurements();
      $('photoMeaCount').textContent = String(ph.measurements.length);
    }
  });

  function drawPhotoMeasurements(livePoint){
    while(photoMeasureSvg.firstChild) photoMeasureSvg.removeChild(photoMeasureSvg.firstChild);
    const ph = photoState.pin && photoState.pin.photos[photoState.idx]; if(!ph) return;

    // Persisted
    (ph.measurements||[]).forEach(m=>{
      const a=m.points[0], b=m.points[1];
      drawSvgLine(photoMeasureSvg, a, b);
      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      const label = (typeof m.feet==='number')
        ? m.feet.toFixed(2)+' ft'
        : (ph.scalePxPerFt ? (dist(a,b)/ph.scalePxPerFt).toFixed(2)+' ft' : Math.round(dist(a,b))+' px');
      drawSvgText(photoMeasureSvg, midx, midy-6, label);
    });

    // Live rubberband
    if(photoMeasureFirst && livePoint){
      drawSvgLine(photoMeasureSvg, photoMeasureFirst, livePoint, true);
    }
  }

  function openPhotoModal(){
    const p=selectedPin(); if(!p) return alert('Select a pin first.');
    if(!p.photos || !p.photos.length) return alert('No photos attached.');
    photoState.pin=p; photoState.idx=0; showPhoto(); photoModal.style.display='flex';
  }
  function showPhoto(){
    const ph=photoState.pin.photos[photoState.idx];
    photoImg.src=ph.dataUrl; photoPinId.textContent=photoState.pin.id; photoName.textContent=ph.name;
    photoMeaCount.textContent=String((ph.measurements && ph.measurements.length) || 0);
    drawPhotoMeasurements();
  }
  function closePhotoModal(){ photoModal.style.display='none'; }

  /***********
   * OCR     *
   ***********/
  async function ocrCurrentView(){
    const img=stageImage; if(!img || !img.src){ alert('No page image.'); return; }
    const canvas=document.createElement('canvas');
    const w=img.naturalWidth||img.width, h=img.naturalHeight||img.height; canvas.width=w; canvas.height=h;
    const ctx=canvas.getContext('2d'); ctx.drawImage(img,0,0);
    const res = await Tesseract.recognize(canvas, 'eng');
    const out = (res && res.data && res.data.text || '').trim();
    if(!out){ alert('No text recognized.'); return; }
    const head = out.slice(0,200);
    try { await navigator.clipboard.writeText(out); } catch(_){}
    inputSearch.value=head; renderPinsList(); alert('OCR complete. First 200 chars in search. Full text copied to clipboard.');
  }

  /****************
   * Rendering    *
   ****************/
  function renderAll(){
    ensureTypeOptions();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    renderProjectLabel();
    drawMeasurements();
    toggleField.checked = !!project.settings.fieldMode;
  }
  function renderProjectLabel(){
    const totalPins = project.pages.reduce((a,p)=> a + ((p.pins && p.pins.length) || 0), 0);
    projectLabel.textContent = 'Project: ' + project.name + ' • Pages: ' + project.pages.length + ' • Pins: ' + totalPins;
  }

  function ensureTypeOptions(){
    // datalist + filter select
    if(!typeList.dataset.filled){
      Object.keys(SIGN_TYPES).forEach(t=>{
        const o=document.createElement('option'); o.value=t; typeList.appendChild(o);
        const selOpt=document.createElement('option'); selOpt.value=t; selOpt.textContent=t; filterType.appendChild(selOpt);
      });
      typeList.dataset.filled = '1';
    }
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
    const pg=currentPage(); if(!pg){ stageImage.removeAttribute('src'); return; }
    stageImage.src=pg.blobUrl;
  }

  function renderPins(){
    pinLayer.innerHTML='';
    const pg=currentPage(); if(!pg) return;
    fieldMode = !!project.settings.fieldMode;
    const q = (inputSearch.value || '').toLowerCase();
    const typeFilter = filterType.value;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type!==typeFilter) return;
      const line = [p.sign_type,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el=document.createElement('div');
      el.className='pin'+(p.id===project._sel?' selected':''); el.dataset.id=p.id;
      el.textContent = (SHORT[p.sign_type] || (p.sign_type ? p.sign_type.slice(0,3).toUpperCase() : 'PIN'));
      const color = SIGN_TYPES[p.sign_type] || BRAND_FALLBACK;
      el.style.left = String(p.x_pct) + '%';
      el.style.top  = String(p.y_pct) + '%';
      el.style.background = color;
      el.style.padding = fieldMode ? '.28rem .5rem' : '.22rem .45rem';
      el.style.fontSize = fieldMode ? '0.95rem' : '0.8rem';
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

      const row=document.createElement('div'); row.className='row';
      const meta=document.createElement('div'); meta.className='meta';
      meta.innerHTML = '<strong>'+ (p.sign_type || '-') +'</strong><span class="small">'
        + (p.room_number || '-') + ' ' + (p.room_name || '')
        + ' · [' + (p.building || '-') + '/' + (p.level || '-') + ']</span>';
      const right=document.createElement('div');
      const cb=document.createElement('input'); cb.type='checkbox'; cb.dataset.id=p.id; cb.title='Select for bulk'; right.appendChild(cb);
      const go=document.createElement('button'); go.className='pill-btn'; go.textContent='Go'; go.onclick=()=> selectPin(p.id); right.appendChild(go);
      row.appendChild(meta); row.appendChild(right);
      pinList.appendChild(row);
    });
  }

  function updatePinDetails(){
    const p=selectedPin();
    selId.textContent = p ? p.id : 'None';
    fieldType.value = (p && p.sign_type) || '';
    fieldRoomNum.value = (p && p.room_number) || '';
    fieldRoomName.value = (p && p.room_name) || '';
    fieldBuilding.value = (p && p.building) || '';
    fieldLevel.value = (p && p.level) || '';
    fieldNotes.value = (p && p.notes) || '';
    posLabel.textContent = p ? pctLabel(p.x_pct, p.y_pct) : '—';
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
    if((/ELECTRICAL|DATA/).test(rn) && p.sign_type!=='BOH'){ list.push('Recommended BOH for ELECTRICAL/DATA rooms.'); }
    list.forEach(s=>{ const t=document.createElement('span'); t.className='tag'; t.textContent=s; warnsEl.appendChild(t); });
  }

  function selectedPin(){
    const idv=project && project._sel; if(!idv) return null;
    return findPin(idv);
  }
  function selectPin(idv){
    project._sel=idv; saveProject(project); renderPins();
    const el = Array.from(pinLayer.children).find(x=>x.dataset.id===idv);
    if(el && el.scrollIntoView) el.scrollIntoView({block:'center', inline:'center', behavior:'smooth'});
    const p=findPin(idv); if(p){
      const pg = project.pages.find(pg=> (pg.pins||[]).includes(p));
      if(pg && project._pageId!==pg.id){ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); drawMeasurements(); }
    }
  }
  function clearSelection(){ if(!project) return; project._sel=null; saveProject(project); updatePinDetails(); }

  /****************
   * CSV / XLSX / ZIP
   ****************/
  function toRows(){
    const rows=[];
    project.pages.forEach(pg=> (pg.pins||[]).forEach(p=> rows.push({
      id: p.id,
      sign_type: p.sign_type || '',
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
    rows.sort((a,b)=>{
      const byB = (a.building||'').localeCompare(b.building||'');
      if(byB) return byB;
      const byL = (a.level||'').localeCompare(b.level||'');
      if(byL) return byL;
      return a.room_number.localeCompare(b.room_number, undefined, {numeric:true,sensitivity:'base'});
    });
    return rows;
  }
  function exportCSV(){
    const rows = toRows();
    const csvHead = Object.keys(rows[0]||{}).join(',');
    const csvBody = rows.map(r=> Object.values(r).map(v=>csvEscape(v)).join(',')).join('\n');
    downloadText(safeName(project.name)+'_signage.csv', csvHead+'\n'+csvBody);
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
    downloadFile(safeName(project.name)+'_signage.xlsx', new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  }
  async function exportZIP(){
    const rows=toRows(); const zip=new JSZip();
    const csvHead = Object.keys(rows[0]||{}).join(',');
    const csvBody = rows.map(r=> Object.values(r).map(v=>csvEscape(v)).join(',')).join('\n');
    zip.file('signage.csv', csvHead+'\n'+csvBody);
    // photos
    project.pages.forEach(pg=> (pg.pins||[]).forEach(pin=> (pin.photos||[]).forEach((ph, idx)=>{
      const folder=zip.folder('photos/'+pin.id);
      folder.file(ph.name||('photo_'+String(idx+1)+'.png'), dataURLtoArrayBuffer(ph.dataUrl));
    })));
    const blob = await zip.generateAsync({type:'blob'});
    downloadFile(safeName(project.name)+'_export.zip', blob);
  }

  function importCSV(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2){ alert('No rows found.'); return; }
    const hdr=lines[0].split(',').map(h=>h.trim()); commit();
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]);
      const row={}; hdr.forEach((h,idx)=> row[h]=cells[idx]||'');
      const xp = parseFloat(row.x_pct); const yp=parseFloat(row.y_pct);
      const p = makePin(isFinite(xp)?xp:50, isFinite(yp)?yp:50);
      p.sign_type=row.sign_type||''; p.room_number=row.room_number||''; p.room_name=row.room_name||'';
      p.building=row.building||''; p.level=row.level||''; p.notes=row.notes||'';
      currentPage().pins.push(p);
    }
    saveProject(project); renderPins(); renderPinsList(); alert('Imported rows into current page.');
  }

  /***************
   * Helpers     *
   ***************/
  const projectContext={building:'', level:''};
  const UNDO=[]; const REDO=[]; const MAX_UNDO=50;

  function id(){ return Math.random().toString(36).slice(2,10); }
  function fix(n){ return typeof n==='number'? +Number(n||0).toFixed(3):''; }
  function pctLabel(x,y){ return ( (x||0).toFixed(2)+'%, '+(y||0).toFixed(2)+'%' ); }
  function dist(a,b){ const dx=a.x-b.x, dy=a.y-b.y; return Math.sqrt(dx*dx+dy*dy); }
  function safeName(s){ return String(s||'project').replace(/\W+/g,'_'); }

  function toPctCoords(e){
    const rect=stageImage.getBoundingClientRect();
    const x=Math.min(Math.max(0,e.clientX-rect.left),rect.width);
    const y=Math.min(Math.max(0,e.clientY-rect.top),rect.height);
    return { x_pct: +(x/rect.width*100).toFixed(3), y_pct:+(y/rect.height*100).toFixed(3) };
  }
  function toLocal(e){
    const r=stageImage.getBoundingClientRect();
    return { x:e.clientX-r.left, y:e.clientY-r.top };
  }
  function toScreen(pt){
    const r=stageImage.getBoundingClientRect();
    return { x:r.left+pt.x, y:r.top+pt.y };
  }

  function csvEscape(v){ const s=(v==null? '': String(v)); if(/["\n,]/.test(s)) return '"'+s.replace(/"/g,'""')+'"'; return s; }
  function parseCsvLine(line){
    const out=[]; let cur=''; let i=0; let inQ=false;
    while(i<line.length){
      const ch=line[i++];
      if(inQ){
        if(ch==='"'){ if(line[i]==='"'){ cur+='"'; i++; } else { inQ=false; } }
        else { cur+=ch; }
      } else {
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

  function newProject(name){
    return {
      id: id(),
      name: name,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      pages: [],
      settings: { colorsByType: SIGN_TYPES, fieldMode: false }
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
      notes: '',
      photos: [],
      lastEdited: Date.now()
    };
  }
  function findPin(idv){
    for(const pg of project.pages){ const f = (pg.pins||[]).find(p=>p.id===idv); if(f) return f; }
    return null;
  }
  function checkedPinIds(){
    return Array.from(pinList.querySelectorAll('input[type="checkbox"]:checked')).map(cb=>cb.dataset.id);
  }

  let projects = loadProjects();
  let currentProjectId = localStorage.getItem('survey:lastOpenProjectId') || null;
  let project = currentProjectId ? projects.find(p=>p.id===currentProjectId) : null;
  if(!project){ project = newProject('Untitled Project'); saveProject(project); }
  selectProject(project.id);
  showScreen('editor');

  function selectProject(pid){
    project = projects.find(p=>p.id===pid);
    if(!project) return;
    localStorage.setItem('survey:lastOpenProjectId', pid);
    renderAll();
  }
  function saveProject(p){
    p.updatedAt = Date.now();
    const i = projects.findIndex(x=>x.id===p.id);
    if(i>=0) projects[i]=p; else projects.push(p);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    projects = loadProjects();
  }
  function loadProjects(){
    try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); }catch(_){ return []; }
  }

  function addImagePage(url, name){
    const pg = {
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
      const canvas=document.createElement('canvas'); canvas.width=viewport.width; canvas.height=viewport.height;
      const ctx=canvas.getContext('2d'); await page.render({canvasContext:ctx, viewport:viewport}).promise;
      const data=canvas.toDataURL('image/png');
      const pg = {
        id: id(),
        name: file.name.replace(/\.[^.]+$/,'') + ' · p' + String(i),
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

  // Undo/Redo
  function snapshot(){ return JSON.stringify(project); }
  function loadSnapshot(s){
    project = JSON.parse(s);
    const i = projects.findIndex(p=>p.id===project.id);
    if(i>=0) projects[i]=project; else projects.push(project);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    localStorage.setItem('survey:lastOpenProjectId', project.id);
  }
  function commit(){ UNDO.push(snapshot()); if(UNDO.length>MAX_UNDO) UNDO.shift(); REDO.length=0; }
  function undo(){ if(!UNDO.length) return; REDO.push(snapshot()); const s=UNDO.pop(); loadSnapshot(s); renderAll(); }
  function redo(){ if(!REDO.length) return; UNDO.push(snapshot()); const s=REDO.pop(); loadSnapshot(s); renderAll(); }

  // Keyboard
  window.addEventListener('keydown',(e)=>{
    if((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); undo(); }
    if((e.ctrlKey||e.metaKey) && e.shiftKey && e.key.toLowerCase()==='z'){ e.preventDefault(); redo(); }
    if(e.key==='Delete'){ const p=selectedPin(); if(p){ commit(); currentPage().pins=currentPage().pins.filter(x=>x.id!==p.id); saveProject(project); renderPins(); renderPinsList(); clearSelection(); } }
    if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)){
      const p=selectedPin(); if(!p) return; e.preventDefault(); commit();
      const delta = project.settings.fieldMode? 0.5 : 0.2;
      if(e.key==='ArrowUp') p.y_pct=Math.max(0,(p.y_pct||0)-delta);
      if(e.key==='ArrowDown') p.y_pct=Math.min(100,(p.y_pct||0)+delta);
      if(e.key==='ArrowLeft') p.x_pct=Math.max(0,(p.x_pct||0)-delta);
      if(e.key==='ArrowRight') p.x_pct=Math.min(100,(p.x_pct||0)+delta);
      p.lastEdited=Date.now(); saveProject(project); renderPins(); updatePinDetails();
    }
  });

  /***************
   * Project list
   ***************/
  function renderProjectList(){
    projectList.innerHTML='';
    projects = loadProjects();
    if(!projects.length){
      const empty=document.createElement('div'); empty.className='muted'; empty.textContent='No projects yet.';
      projectList.appendChild(empty); return;
    }
    projects.sort((a,b)=> (b.updatedAt||0)-(a.updatedAt||0));
    projects.forEach(p=>{
      const row=document.createElement('div'); row.className='row';
      const meta=document.createElement('div'); meta.className='meta';
      const date=new Date(p.updatedAt||p.createdAt||Date.now()).toLocaleString();
      meta.innerHTML='<strong>'+p.name+'</strong><span class="small">Updated '+date+' — '+(p.pages?.length||0)+' page(s)</span>';
      const act=document.createElement('div');
      const open=document.createElement('button'); open.className='pill-btn'; open.textContent='Open'; open.onclick=()=>{ selectProject(p.id); showScreen('editor'); };
      act.appendChild(open);
      row.appendChild(meta); row.appendChild(act);
      projectList.appendChild(row);
    });
  }
  function doNewProject(){
    commit();
    const name = prompt('Project name:','New Project') || 'New Project';
    const p=newProject(name); saveProject(p); selectProject(p.id);
  }
  function duplicateProject(){
    const name=prompt('Duplicate as:', project.name+' (copy)'); if(!name) return;
    commit();
    const copy=JSON.parse(JSON.stringify(project));
    copy.id=id(); copy.name=name; copy.createdAt=Date.now(); copy.updatedAt=Date.now();
    saveProject(copy); selectProject(copy.id);
  }
  function promptRenameProject(){
    const name=prompt('Rename project:', project.name); if(!name) return;
    commit(); project.name=name; saveProject(project); renderProjectLabel();
  }
  function deleteProject(){
    if(!project) return;
    if(!confirm('Delete the currently selected project? This cannot be undone.')) return;
    projects = projects.filter(p=>p.id!==project.id);
    localStorage.setItem('survey:projects', JSON.stringify(projects));
    if(projects[0]){ selectProject(projects[0].id); }
    else { const p=newProject('Untitled Project'); saveProject(p); selectProject(p.id); }
    renderProjectList();
  }
  function exportAllProjectsJSON(){
    const data = localStorage.getItem('survey:projects') || '[]';
    downloadText('southwood_projects.json', data);
  }

  /***************
   * Screen swap *
   ***************/
  function showScreen(name){
    screenHome.classList.remove('active');
    screenProjects.classList.remove('active');
    screenEditor.classList.remove('active');
    if(name==='home') screenHome.classList.add('active');
    else if(name==='projects') screenProjects.classList.add('active');
    else screenEditor.classList.add('active');
  }
});
