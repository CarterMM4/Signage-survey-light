// app.js
document.addEventListener('DOMContentLoaded', () => {
  /******************
   * Utilities
   ******************/
  const $ = (id) => document.getElementById(id);
  const on = (el, evt, fn) => el && el.addEventListener(evt, fn);

  function id(){ return Math.random().toString(36).slice(2,10); }
  function clamp(n,min,max){ return Math.max(min, Math.min(max, n)); }

  /******************
   * Screens
   ******************/
  const screenHome = $('screenHome');
  const screenProjects = $('screenProjects');
  const screenEditor = $('screenEditor');

  function showScreen(which){
    [screenHome, screenProjects, screenEditor].forEach(s=> s.classList.remove('show'));
    if(which==='home') screenHome.classList.add('show');
    if(which==='projects') screenProjects.classList.add('show');
    if(which==='editor') screenEditor.classList.add('show');
    if(which==='home') renderHomeRecent();
  }

  on($('navHome'),'click',()=> showScreen('home'));
  on($('navProjects'),'click',()=> { renderProjectsGrid(); showScreen('projects'); });
  on($('navEditor'),'click',()=> showScreen('editor'));
  on($('homeNew'),'click',()=> createProjectFlow());
  on($('homeOpen'),'click',()=> { renderProjectsGrid(); showScreen('projects'); });

  /******************
   * Data Model
   ******************/
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

  const NAME_COLOR_RULES = [
    { re:/exit/i, color:'#ef4444' },
    { re:/ingress|enter|entry/i, color:'#22c55e' },
    { re:/egress|leave/i, color:'#ef4444' },
    { re:/elev|lift/i, color:'#8b5cf6' },
    { re:/restroom|wc|toilet/i, color:'#10b981' },
    { re:/callbox|intercom/i, color:'#06b6d4' },
    { re:/evac|stair/i, color:'#84cc16' }
  ];

  let projects = loadProjects();
  let project = null; // current project object
  const projectContext = { building:'', level:'' };

  function loadProjects(){
    try{ return JSON.parse(localStorage.getItem('survey:projects')||'[]'); } catch { return []; }
  }
  function saveProjectsList(){
    localStorage.setItem('survey:projects', JSON.stringify(projects));
  }
  function saveProject(p){
    p.updatedAt = Date.now();
    const i = projects.findIndex(x=>x.id===p.id);
    if(i>=0) projects[i]=p; else projects.push(p);
    saveProjectsList();
  }
  function newProject(name){
    return {
      id: id(),
      name: name,
      createdAt: Date.now(),
      updatedAt: Date.now(),
      pages: [],
      settings: { fieldMode: false }
    };
  }
  function selectProject(pid){
    const found = projects.find(p=>p.id===pid);
    if(!found){ alert('Project not found.'); return; }
    project = found;
    localStorage.setItem('survey:lastOpenProjectId', pid);
    renderEditorAll();
    showScreen('editor');
  }

  /******************
   * HOME – recent projects
   ******************/
  const homeRecent = $('homeRecent');
  function renderHomeRecent(){
    projects = loadProjects();
    homeRecent.innerHTML = '';
    projects
      .slice()
      .sort((a,b)=> b.updatedAt - a.updatedAt)
      .slice(0,6)
      .forEach(p=>{
        const card = document.createElement('div'); card.className='card';
        const h = document.createElement('h4'); h.textContent=p.name; card.appendChild(h);
        const meta = document.createElement('div'); meta.className='muted';
        const pinCount = (p.pages||[]).reduce((acc,pg)=>acc+(pg.pins?pg.pins.length:0),0);
        meta.textContent = `Pages: ${p.pages.length} • Pins: ${pinCount}`;
        card.appendChild(meta);
        const row = document.createElement('div'); row.className='spaced';
        const open = document.createElement('button'); open.textContent='Open'; open.className='btn-primary';
        on(open,'click',()=> selectProject(p.id));
        row.appendChild(open);
        card.appendChild(row);
        homeRecent.appendChild(card);
      });
  }

  /******************
   * Projects screen
   ******************/
  const projectsGrid = $('projectsGrid');
  const projSearch = $('projSearch');
  on(projSearch,'input',()=> renderProjectsGrid());

  function duplicateProject(p){
    const copy = JSON.parse(JSON.stringify(p));
    copy.id = id();
    copy.name = p.name + ' (copy)';
    copy.createdAt = Date.now();
    copy.updatedAt = Date.now();
    saveProject(copy);
    return copy;
  }

  function renderProjectsGrid(){
    const q = (projSearch.value||'').toLowerCase();
    projects = loadProjects();
    projectsGrid.innerHTML = '';
    projects
      .filter(p=>!q || p.name.toLowerCase().includes(q))
      .sort((a,b)=> b.updatedAt - a.updatedAt)
      .forEach(p=>{
        const card = document.createElement('div'); card.className='card';
        const h = document.createElement('h4'); h.textContent = p.name; card.appendChild(h);
        const meta = document.createElement('div'); meta.className='muted';
        const pinCount = (p.pages||[]).reduce((acc,pg)=>acc+(pg.pins?pg.pins.length:0),0);
        meta.textContent = `Pages: ${p.pages.length} • Pins: ${pinCount}`;
        card.appendChild(meta);

        const row = document.createElement('div'); row.className='spaced';
        const openBtn = document.createElement('button'); openBtn.textContent='Open'; openBtn.className='btn-primary';
        on(openBtn,'click',()=> selectProject(p.id));

        const dupBtn = document.createElement('button'); dupBtn.textContent='Duplicate';
        on(dupBtn,'click',()=>{ const cp = duplicateProject(p); renderProjectsGrid(); renderHomeRecent(); alert('Duplicated as: ' + cp.name); });

        const renameBtn = document.createElement('button'); renameBtn.textContent='Rename';
        on(renameBtn,'click',()=>{
          const name = prompt('Rename project:', p.name);
          if(!name) return;
          p.name = name; saveProject(p); renderProjectsGrid(); renderHomeRecent();
        });

        const delBtn = document.createElement('button'); delBtn.textContent='Delete';
        on(delBtn,'click',()=>{
          if(!confirm('Delete this project?')) return;
          projects = projects.filter(x=>x.id!==p.id);
          saveProjectsList();
          renderProjectsGrid(); renderHomeRecent();
        });
        row.appendChild(openBtn); row.appendChild(dupBtn); row.appendChild(renameBtn); row.appendChild(delBtn);
        card.appendChild(row);
        projectsGrid.appendChild(card);
      });
  }

  on($('btnCreateProj'),'click',()=> createProjectFlow());
  function createProjectFlow(){
    const name = prompt('Project name?','New Project');
    if(!name) return;
    const p = newProject(name);
    saveProject(p);
    selectProject(p.id);
  }

  /******************
   * Editor DOM refs
   ******************/
  const thumbsEl = $('thumbs');
  const stage = $('stage');
  const stageInner = $('stageInner');
  const stageImage = $('stageImage');
  const measureSvg = $('measureSvg');
  const pinLayer = $('pinLayer');

  const projectLabel = $('projectLabel');
  const inputUpload = $('inputUpload');
  const btnUpload = $('btnUpload');

  const inputBuilding = $('inputBuilding');
  const inputLevel = $('inputLevel');

  const inputSearch = $('inputSearch');
  const filterType = $('filterType');
  const toggleField = $('toggleField');

  const fieldType = $('fieldType');
  const fieldCustom = $('fieldCustom');
  const fieldRoomNum = $('fieldRoomNum');
  const fieldRoomName = $('fieldRoomName');
  const fieldBuilding = $('fieldBuilding');
  const fieldLevel = $('fieldLevel');
  const fieldNotes = $('fieldNotes');
  const selId = $('selId');
  const posLabel = $('posLabel');
  const pinList = $('pinList');

  const btnAddPin = $('btnAddPin');
  const btnClearPins = $('btnClearPins');
  const btnDuplicate = $('btnDuplicate');
  const btnDelete = $('btnDelete');
  const btnOpenPhoto = $('btnOpenPhoto');
  const btnAddPhoto = $('btnAddPhoto');
  const inputPhoto = $('inputPhoto');

  // Measure controls
  const btnMeasureToggle = $('btnMeasureToggle');
  const btnMeasureClear = $('btnMeasureClear');

  // Zoom controls
  const btnZoomIn = $('btnZoomIn');
  const btnZoomOut = $('btnZoomOut');
  const btnZoomReset = $('btnZoomReset');

  // Export/Import
  const btnExportCSV = $('btnExportCSV');
  const btnExportXLSX = $('btnExportXLSX');
  const btnExportZIP = $('btnExportZIP');
  const btnImportCSV = $('btnImportCSV');
  const inputImportCSV = $('inputImportCSV');

  // Photo modal
  const photoModal = $('photoModal');
  const photoImg = $('photoImg');
  const photoMeasureSvg = $('photoMeasureSvg');
  const photoPinId = $('photoPinId');
  const photoName = $('photoName');
  const photoMeaCount = $('photoMeaCount');
  const btnPhotoClose = $('btnPhotoClose');
  const btnPhotoMeasure = $('btnPhotoMeasure');
  const btnPhotoPrev = $('btnPhotoPrev');
  const btnPhotoNext = $('btnPhotoNext');
  const btnPhotoDelete = $('btnPhotoDelete');
  const btnPhotoDownload = $('btnPhotoDownload');

  /******************
   * Zoom (stage)
   ******************/
  let stageZoom = 1;
  function applyZoom(){
    stageInner.style.transform = `scale(${stageZoom})`;
    // Resize overlay SVG to match image's rendered box at current zoom
    const rect = stageImage.getBoundingClientRect();
    measureSvg.setAttribute('width', rect.width);
    measureSvg.setAttribute('height', rect.height);
  }
  function zoomBy(f){ stageZoom = clamp(stageZoom * f, 0.25, 6); applyZoom(); }
  function zoomReset(){ stageZoom = 1; applyZoom(); }

  on(btnZoomIn,'click',()=> zoomBy(1.2));
  on(btnZoomOut,'click',()=> zoomBy(1/1.2));
  on(btnZoomReset,'click', zoomReset);

  // Ctrl + wheel zoom (pinch on trackpads triggers ctrlKey=true in many browsers)
  on(stage,'wheel',(e)=>{
    if(!e.ctrlKey) return; // avoid hijacking normal scroll
    e.preventDefault();
    const dir = e.deltaY > 0 ? (1/1.15) : 1.15;
    zoomBy(dir);
  }, {passive:false});

  /******************
   * Editor helpers
   ******************/
  function currentPage(){
    if(!project) return null;
    if(!project._pageId && project.pages[0]) project._pageId = project.pages[0].id;
    return project.pages.find(p=>p.id===project._pageId) || null;
  }
  function makePin(x_pct,y_pct){
    return {
      id: id(),
      sign_type: '',
      custom_name: '',
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
    for(const pg of (project?.pages||[])){
      const f = (pg.pins||[]).find(p=>p.id===idv);
      if(f) return f;
    }
    return null;
  }

  /******************
   * Upload pages
   ******************/
  on(btnUpload,'click',()=> inputUpload.click());
  on(inputUpload,'change', async (e)=>{
    if(!project) return;
    const files = [...e.target.files];
    if(!files.length) return;
    for(const f of files){
      if(f.type==='application/pdf'){ await addPdfPages(f); }
      else if(f.type.startsWith('image/')){
        const url = URL.createObjectURL(f);
        addImagePage(url, f.name.replace(/\.[^.]+$/,''));
      }
    }
    renderEditorAll();
    inputUpload.value = '';
  });

  function addImagePage(blobUrl, name){
    const pg = {
      id: id(),
      name: name || 'Image',
      kind: 'image',
      blobUrl: blobUrl,
      pins: [],
      measurements: [],
      updatedAt: Date.now()
    };
    project.pages.push(pg);
    if(!project._pageId) project._pageId = pg.id;
    saveProject(project);
  }

  async function addPdfPages(file){
    const url = URL.createObjectURL(file);
    const pdf = await pdfjsLib.getDocument(url).promise;
    for(let i=1;i<=pdf.numPages;i++){
      const page = await pdf.getPage(i);
      const viewport = page.getViewport({scale:2});
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');
      await page.render({canvasContext:ctx, viewport:viewport}).promise;
      const dataUrl = canvas.toDataURL('image/png');
      const pg = {
        id: id(),
        name: `${file.name.replace(/\.[^.]+$/,'')} · p${i}`,
        kind: 'pdf',
        pdfPage: i,
        blobUrl: dataUrl,
        pins: [],
        measurements: [],
        updatedAt: Date.now()
      };
      project.pages.push(pg);
    }
    if(!project._pageId && project.pages[0]) project._pageId = project.pages[0].id;
    saveProject(project);
  }

  /******************
   * Rendering
   ******************/
  let typesFilled = false;
  function ensureTypes(){
    if(typesFilled) return;
    Object.keys(SIGN_TYPES).forEach(t=>{
      const o=document.createElement('option'); o.value=t; o.textContent=t;
      filterType.appendChild(o.cloneNode(true));
      fieldType.appendChild(o);
    });
    typesFilled = true;
  }

  function renderEditorAll(){
    ensureTypes();
    renderProjectLabel();
    renderThumbs();
    renderStage();
    renderPins();
    renderPinsList();
    toggleField.checked = !!project.settings.fieldMode;
    drawMeasurements();
    applyZoom(); // sync overlay sizes
  }

  function renderProjectLabel(){
    const pinCount = (project.pages||[]).reduce((acc,pg)=>acc + (pg.pins?pg.pins.length:0),0);
    projectLabel.textContent = `Project: ${project.name} • Pages: ${project.pages.length} • Pins: ${pinCount}`;
  }

  function renderThumbs(){
    thumbsEl.innerHTML = '';
    (project.pages||[]).forEach(pg=>{
      const d = document.createElement('div'); d.className='thumb'+(project._pageId===pg.id?' active':'');
      const im = document.createElement('img'); im.src = pg.blobUrl; d.appendChild(im);
      const inp = document.createElement('input'); inp.value = pg.name;
      on(inp,'input',()=>{ pg.name = inp.value; pg.updatedAt=Date.now(); saveProject(project); renderProjectLabel(); });
      d.appendChild(inp);
      on(d,'click',()=>{ project._pageId=pg.id; saveProject(project); renderStage(); renderPins(); drawMeasurements(); applyZoom(); });
      thumbsEl.appendChild(d);
    });
  }

  function renderStage(){
    const pg = currentPage();
    if(!pg){ stageImage.removeAttribute('src'); return; }
    stageImage.src = pg.blobUrl;
    // Reset zoom when switching pages for clarity but keep user option
    zoomReset();
  }

  function pctFromEvent(e){
    const rect = stageImage.getBoundingClientRect();
    const x = clamp(e.clientX - rect.left, 0, rect.width);
    const y = clamp(e.clientY - rect.top , 0, rect.height);
    return {
      x_pct: +(x/rect.width*100).toFixed(3),
      y_pct: + (y/rect.height*100).toFixed(3)
    };
  }

  function colorForPin(p){
    if(p.sign_type) return SIGN_TYPES[p.sign_type] || '#2ecc71';
    const name = (p.custom_name || p.room_name || '').trim();
    for(const r of NAME_COLOR_RULES){
      if(r.re.test(name)) return r.color;
    }
    return '#2ecc71';
  }

  function labelForPin(p){
    if(p.custom_name) return p.custom_name;
    if(p.sign_type) return SHORT[p.sign_type] || p.sign_type.slice(0,4).toUpperCase();
    if(p.room_number) return p.room_number;
    return 'PIN';
  }

  function renderPins(){
    pinLayer.innerHTML = '';
    const pg = currentPage(); if(!pg) return;
    const q = (inputSearch.value||'').toLowerCase();
    const typeFilter = filterType.value;
    const fieldMode = !!project.settings.fieldMode;

    (pg.pins||[]).forEach(p=>{
      if(typeFilter && p.sign_type !== typeFilter) return;
      const line = [p.sign_type,p.custom_name,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;

      const el = document.createElement('div');
      el.className = 'pin'+(p.id===project._sel?' selected':'');
      el.dataset.id = p.id;
      el.textContent = labelForPin(p);
      el.style.left = p.x_pct + '%';
      el.style.top = p.y_pct + '%';
      el.style.background = colorForPin(p);
      el.style.padding = fieldMode? '.36rem .68rem' : '.28rem .6rem';
      el.style.fontSize = fieldMode? '0.98rem' : '0.82rem';

      on(el,'pointerdown',(ev)=>{
        ev.preventDefault();
        selectPin(p.id,false);
        el.setPointerCapture?.(ev.pointerId);
        draggingPin = { el:el, pin:p };
      });

      on(el,'dblclick',()=>{ selectPin(p.id,true); openPhotoModal(); });

      pinLayer.appendChild(el);
    });

    updatePinFields();
  }

  function renderPinsList(){
    pinList.innerHTML = '';
    const pg = currentPage(); if(!pg) return;
    const q = (inputSearch.value||'').toLowerCase();
    const typeFilter = filterType.value;

    (pg.pins||[]).forEach(p=>{
      const line=[p.sign_type,p.custom_name,p.room_number,p.room_name,p.building,p.level,p.notes].join(' ').toLowerCase();
      if(q && !line.includes(q)) return;
      if(typeFilter && p.sign_type!==typeFilter) return;

      const card = document.createElement('div'); card.className='card';
      const title = document.createElement('div');
      title.innerHTML = `<strong>${labelForPin(p)}</strong> <span class="muted">[${p.building||'-'}/${p.level||'-'}]</span>`;
      card.appendChild(title);

      const row = document.createElement('div'); row.className='spaced';
      const go = document.createElement('button'); go.textContent='Go';
      on(go,'click',()=> selectPin(p.id,true));
      const openPhotos = document.createElement('button'); openPhotos.textContent='Photos';
      on(openPhotos,'click',()=>{ selectPin(p.id,true); openPhotoModal(); });
      row.appendChild(go); row.appendChild(openPhotos);
      card.appendChild(row);

      pinList.appendChild(card);
    });
  }

  function updatePinFields(){
    const p = selectedPin();
    selId.textContent = p ? p.id : 'None';
    // Guard against missing inputs
    if(fieldType) fieldType.value = p?.sign_type || '';
    if(fieldCustom) fieldCustom.value = p?.custom_name || '';
    if(fieldRoomNum) fieldRoomNum.value = p?.room_number || '';
    if(fieldRoomName) fieldRoomName.value = p?.room_name || '';
    if(fieldBuilding) fieldBuilding.value = p?.building || '';
    if(fieldLevel) fieldLevel.value = p?.level || '';
    if(fieldNotes) fieldNotes.value = p?.notes || '';
    const x = Number(p?.x_pct||0).toFixed(2), y=Number(p?.y_pct||0).toFixed(2);
    posLabel.textContent = p ? `${x}%, ${y}%` : '—';
  }

  function selectedPin(){
    const idv = project?._sel;
    if(!idv) return null;
    return findPin(idv);
  }

  function selectPin(idv, center){
    project._sel = idv;
    saveProject(project);
    renderPins();
    if(center){
      const el = [...pinLayer.children].find(x=>x.dataset.id===idv);
      el?.scrollIntoView({block:'center',inline:'center',behavior:'smooth'});
    }
  }

  /******************
   * Editing & Pins
   ******************/
  let addingPin = false;
  on(btnAddPin,'click',()=>{
    addingPin = !addingPin;
    btnAddPin.textContent = addingPin ? 'Click on page…' : 'Add Pin';
  });

  on(stage,'pointerdown',(e)=>{
    if(!addingPin) return;
    if(e.target && e.target.classList && e.target.classList.contains('pin')) return;
    const pt = pctFromEvent(e);
    const p = makePin(pt.x_pct, pt.y_pct);
    const pg = currentPage();
    pg.pins.push(p);
    saveProject(project);
    addingPin = false; btnAddPin.textContent = 'Add Pin';
    renderPins(); renderPinsList();
    selectPin(p.id,true);
  });

  // Dragging pins (fixed: listen on window so pointer capture works everywhere)
  let draggingPin = null;
  on(window,'pointermove',(e)=>{
    if(!draggingPin) return;
    const pt = pctFromEvent(e);
    draggingPin.el.style.left = pt.x_pct + '%';
    draggingPin.el.style.top = pt.y_pct + '%';
    posLabel.textContent = `${pt.x_pct.toFixed(2)}%, ${pt.y_pct.toFixed(2)}%`;
  });
  on(window,'pointerup',()=>{
    if(!draggingPin) return;
    const rect = stageImage.getBoundingClientRect();
    // If mouse left the image entirely, still commit last known position from element style
    const leftPct = parseFloat(draggingPin.el.style.left) || 0;
    const topPct  = parseFloat(draggingPin.el.style.top)  || 0;
    draggingPin.pin.x_pct = clamp(leftPct, 0, 100);
    draggingPin.pin.y_pct = clamp(topPct , 0, 100);
    draggingPin.pin.lastEdited = Date.now();
    saveProject(project);
    draggingPin = null;
    renderPinsList();
  });

  // Field bindings
  [fieldType,fieldCustom,fieldRoomNum,fieldRoomName,fieldBuilding,fieldLevel,fieldNotes].forEach(el=>{
    on(el,'input',()=>{
      const p = selectedPin(); if(!p) return;
      if(el===fieldType) p.sign_type = el.value || '';
      if(el===fieldCustom) p.custom_name = el.value || '';
      if(el===fieldRoomNum) p.room_number = el.value || '';
      if(el===fieldRoomName) p.room_name = el.value || '';
      if(el===fieldBuilding) p.building = el.value || '';
      if(el===fieldLevel) p.level = el.value || '';
      if(el===fieldNotes) p.notes = el.value || '';
      p.lastEdited = Date.now();
      saveProject(project);
      renderPins();
      renderPinsList();
    });
  });

  on(btnDuplicate,'click',()=>{
    const p = selectedPin(); if(!p) return;
    const copy = JSON.parse(JSON.stringify(p));
    copy.id = id();
    copy.x_pct = clamp((p.x_pct||50)+2,0,100);
    copy.y_pct = clamp((p.y_pct||50)+2,0,100);
    copy.lastEdited = Date.now();
    currentPage().pins.push(copy);
    saveProject(project);
    renderPins(); renderPinsList();
    selectPin(copy.id,true);
  });

  on(btnDelete,'click',()=>{
    const p = selectedPin(); if(!p) return;
    if(!confirm('Delete selected pin?')) return;
    const pg = currentPage();
    pg.pins = (pg.pins||[]).filter(x=>x.id!==p.id);
    project._sel = null;
    saveProject(project);
    renderPins(); renderPinsList();
    updatePinFields();
  });

  on(btnClearPins,'click',()=>{
    if(!currentPage()) return;
    if(!confirm('Clear ALL pins on this page?')) return;
    currentPage().pins = [];
    saveProject(project);
    renderPins(); renderPinsList();
    updatePinFields();
  });

  on(toggleField,'change',()=>{
    if(!project) return;
    project.settings.fieldMode = !!toggleField.checked;
    saveProject(project);
    renderPins();
  });

  on(inputSearch,'input',()=>{ renderPins(); renderPinsList(); });
  on(filterType,'change',()=>{ renderPins(); renderPinsList(); });

  on(inputBuilding,'input',()=>{ projectContext.building = inputBuilding.value; });
  on(inputLevel,'input',()=>{ projectContext.level = inputLevel.value; });

  /******************
   * Measuring (main canvas)
   ******************/
  let measureActive = false;
  let measureDrag = null; // {start:{x,y}, end:{x,y}}
  const stagePreview = document.createElementNS('http://www.w3.org/2000/svg','line');
  stagePreview.setAttribute('stroke','#2ecc71');
  stagePreview.setAttribute('stroke-width','3');
  stagePreview.setAttribute('stroke-dasharray','6,6');
  stagePreview.style.display='none';
  measureSvg.appendChild(stagePreview);

  on(btnMeasureToggle,'click',()=>{
    measureActive = !measureActive;
    btnMeasureToggle.textContent = 'Measure: ' + (measureActive ? 'ON' : 'OFF');
    if(!measureActive){ stagePreview.style.display='none'; measureDrag=null; }
  });
  on(btnMeasureClear,'click',()=>{
    const pg = currentPage(); if(!pg) return;
    pg.measurements = [];
    saveProject(project);
    drawMeasurements();
  });

  function stageLocal(e){
    const r = stageImage.getBoundingClientRect();
    return { x: clamp(e.clientX-r.left,0,r.width), y: clamp(e.clientY-r.top,0,r.height) };
    // Coordinates auto-scale with zoom via getBoundingClientRect
  }

  on(stage,'pointerdown',(e)=>{
    if(!measureActive) return;
    const s = stageLocal(e);
    measureDrag = { start:s, end:s };
    stagePreview.setAttribute('x1', s.x); stagePreview.setAttribute('y1', s.y);
    stagePreview.setAttribute('x2', s.x); stagePreview.setAttribute('y2', s.y);
    stagePreview.style.display = 'block';
  });
  on(stage,'pointermove',(e)=>{
    if(!measureActive || !measureDrag) return;
    const p = stageLocal(e);
    measureDrag.end = p;
    stagePreview.setAttribute('x2', p.x);
    stagePreview.setAttribute('y2', p.y);
  });
  on(stage,'pointerup',()=>{
    if(!measureActive || !measureDrag) return;
    const length = prompt('Enter measured length (ft):','10');
    if(length){
      const pg = currentPage();
      const m = { id:id(), kind:'main', points:[measureDrag.start, measureDrag.end], feet: parseFloat(length)||0 };
      pg.measurements = pg.measurements || [];
      pg.measurements.push(m);
      saveProject(project);
      drawMeasurements();
    }
    stagePreview.style.display='none';
    measureDrag = null;
  });

  function drawMeasurements(){
    while(measureSvg.firstChild) measureSvg.removeChild(measureSvg.firstChild);
    measureSvg.appendChild(stagePreview); // keep preview element
    const pg = currentPage(); if(!pg || !pg.measurements) return;
    pg.measurements.forEach(m=>{
      const a = m.points[0], b = m.points[1];
      const line = document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1', a.x); line.setAttribute('y1', a.y);
      line.setAttribute('x2', b.x); line.setAttribute('y2', b.y);
      line.setAttribute('stroke', '#4cc9f0');
      line.setAttribute('stroke-width','3');
      measureSvg.appendChild(line);

      const label = document.createElementNS('http://www.w3.org/2000/svg','text');
      const midx = (a.x+b.x)/2, midy=(a.y+b.y)/2;
      label.setAttribute('x', midx); label.setAttribute('y', midy-6);
      label.setAttribute('fill', '#fff');
      label.setAttribute('stroke', '#000');
      label.setAttribute('stroke-width', '3');
      label.textContent = (m.feet||0).toFixed(2)+' ft';
      measureSvg.appendChild(label);
    });
  }

  /******************
   * Photo modal + measurement
   ******************/
  const photoState = { pin: null, idx: 0, measure: false };
  let photoDrag = null; // {start:{x,y}, end:{x,y}}
  const photoPreview = document.createElementNS('http://www.w3.org/2000/svg','line');
  photoPreview.setAttribute('stroke','#2ecc71');
  photoPreview.setAttribute('stroke-width','3');
  photoPreview.setAttribute('stroke-dasharray','6,6');
  photoPreview.style.display='none';
  photoMeasureSvg.appendChild(photoPreview);

  function openPhotoModal(){
    const p = selectedPin();
    if(!p) return alert('Select a pin first.');
    if(!p.photos || p.photos.length===0) return alert('No photos attached.');
    photoState.pin = p;
    photoState.idx = 0;
    showPhoto();
    photoModal.classList.add('show');
  }
  function closePhotoModal(){ photoModal.classList.remove('show'); photoDrag=null; photoPreview.style.display='none'; }

  on(btnOpenPhoto,'click',()=> openPhotoModal());
  on(btnPhotoClose,'click',()=> closePhotoModal());

  function showPhoto(){
    const ph = photoState.pin.photos[photoState.idx];
    photoImg.src = ph.dataUrl;
    photoPinId.textContent = photoState.pin.id;
    photoName.textContent = ph.name || `Photo ${photoState.idx+1}`;
    photoMeaCount.textContent = (ph.measurements?.length||0);
    drawPhotoMeasurements();
  }

  function drawPhotoMeasurements(){
    while(photoMeasureSvg.firstChild) photoMeasureSvg.removeChild(photoMeasureSvg.firstChild);
    photoMeasureSvg.appendChild(photoPreview);
    const ph = photoState.pin?.photos[photoState.idx]; if(!ph||!ph.measurements) return;
    ph.measurements.forEach(m=>{
      const a=m.points[0], b=m.points[1];
      const line=document.createElementNS('http://www.w3.org/2000/svg','line');
      line.setAttribute('x1',a.x); line.setAttribute('y1',a.y);
      line.setAttribute('x2',b.x); line.setAttribute('y2',b.y);
      line.setAttribute('stroke','#4cc9f0'); line.setAttribute('stroke-width','3');
      photoMeasureSvg.appendChild(line);

      const label=document.createElementNS('http://www.w3.org/2000/svg','text');
      const midx=(a.x+b.x)/2, midy=(a.y+b.y)/2;
      label.setAttribute('x',midx); label.setAttribute('y',midy-6);
      label.setAttribute('fill','#fff'); label.setAttribute('stroke','#000'); label.setAttribute('stroke-width','3');
      label.textContent=(m.feet||0).toFixed(2)+' ft';
      photoMeasureSvg.appendChild(label);
    });
  }

  on(btnPhotoMeasure,'click',()=>{
    photoState.measure = !photoState.measure;
    btnPhotoMeasure.textContent = 'Measure: ' + (photoState.measure?'ON':'OFF');
    if(!photoState.measure){ photoPreview.style.display='none'; photoDrag=null; }
  });

  on(photoMeasureSvg,'pointerdown',(e)=>{
    if(!photoState.measure) return;
    const r = photoImg.getBoundingClientRect();
    const s = { x: clamp(e.clientX-r.left,0,r.width), y: clamp(e.clientY-r.top,0,r.height) };
    photoDrag = { start:s, end:s };
    photoPreview.setAttribute('x1', s.x); photoPreview.setAttribute('y1', s.y);
    photoPreview.setAttribute('x2', s.x); photoPreview.setAttribute('y2', s.y);
    photoPreview.style.display='block';
  });
  on(photoMeasureSvg,'pointermove',(e)=>{
    if(!photoDrag) return;
    const r = photoImg.getBoundingClientRect();
    const p = { x: clamp(e.clientX-r.left,0,r.width), y: clamp(e.clientY-r.top,0,r.height) };
    photoDrag.end = p;
    photoPreview.setAttribute('x2', p.x);
    photoPreview.setAttribute('y2', p.y);
  });
  on(photoMeasureSvg,'pointerup',()=>{
    if(!photoDrag) return;
    const ph = photoState.pin.photos[photoState.idx];
    const length = prompt('Enter measured length (ft):','10');
    if(length){
      ph.measurements = ph.measurements || [];
      ph.measurements.push({
        id: id(),
        kind: 'photo',
        points: [photoDrag.start, photoDrag.end],
        feet: parseFloat(length)||0
      });
      saveProject(project);
      drawPhotoMeasurements();
      photoMeaCount.textContent = ph.measurements.length;
    }
    photoPreview.style.display='none';
    photoDrag = null;
  });

  on(btnPhotoPrev,'click',()=>{
    if(!photoState.pin) return;
    photoState.idx = Math.max(0, photoState.idx-1);
    showPhoto();
  });
  on(btnPhotoNext,'click',()=>{
    if(!photoState.pin) return;
    photoState.idx = Math.min(photoState.pin.photos.length-1, photoState.idx+1);
    showPhoto();
  });
  on(btnPhotoDelete,'click',()=>{
    if(!photoState.pin) return;
    if(!confirm('Delete this photo?')) return;
    const phs = photoState.pin.photos;
    phs.splice(photoState.idx,1);
    photoState.idx = Math.max(0, photoState.idx-1);
    saveProject(project);
    if(phs.length===0) closePhotoModal(); else showPhoto();
  });
  on(btnPhotoDownload,'click',()=>{
    const ph = photoState.pin?.photos[photoState.idx]; if(!ph) return;
    downloadFile(ph.name||'photo.png', dataURLtoBlob(ph.dataUrl));
  });

  // Add photos to pin
  on(btnAddPhoto,'click',()=> inputPhoto.click());
  on(inputPhoto,'change', async (e)=>{
    const p = selectedPin(); if(!p) return alert('Select a pin first.');
    const files = [...e.target.files]; if(!files.length) return;
    for(const f of files){
      const fr = await toDataURL(f);
      p.photos.push({ name:f.name, dataUrl:fr, measurements:[] });
    }
    p.lastEdited = Date.now();
    saveProject(project);
    renderPinsList();
    inputPhoto.value = '';
    alert('Photo(s) added.');
  });

  function toDataURL(file){
    return new Promise((res,rej)=>{
      const r = new FileReader();
      r.onload = () => res(r.result);
      r.onerror = rej;
      r.readAsDataURL(file);
    });
  }
  function dataURLtoBlob(dataURL){
    const parts = dataURL.split(',');
    const header = parts[0] || '';
    const match = header.match(/:(.*?);/);
    const mime = (match && match[1]) || 'image/png';
    const bstr = atob(parts[1]||'');
    const len = bstr.length;
    const u8 = new Uint8Array(len);
    for(let i=0;i<len;i++) u8[i] = bstr.charCodeAt(i);
    return new Blob([u8], {type:mime});
  }
  function downloadFile(name, blob){
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=name;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 700);
  }

  /******************
   * Export / Import
   ******************/
  function toRows(){
    const rows=[];
    (project.pages||[]).forEach(pg=>{
      (pg.pins||[]).forEach(p=>{
        rows.push({
          id: p.id,
          sign_type: p.sign_type || '',
          custom_name: p.custom_name || '',
          room_number: p.room_number || '',
          room_name: p.room_name || '',
          building: p.building || '',
          level: p.level || '',
          x_pct: typeof p.x_pct==='number' ? +p.x_pct.toFixed(3) : '',
          y_pct: typeof p.y_pct==='number' ? +p.y_pct.toFixed(3) : '',
          notes: p.notes || '',
          page_name: pg.name || '',
          last_edited: new Date(p.lastEdited || project.updatedAt || Date.now()).toISOString()
        });
      });
    });
    rows.sort((a,b)=>
      (a.building||'').localeCompare(b.building||'') ||
      (a.level||'').localeCompare(b.level||'') ||
      a.room_number.localeCompare(b.room_number, undefined, {numeric:true,sensitivity:'base'})
    );
    return rows;
  }

  function csvEscape(v){ const s=(v==null?'':String(v)); return /["\n,]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s; }

  function exportCSV(){
    const rows = toRows();
    const header = Object.keys(rows[0]||{placeholder:''});
    const csv = [header.join(','), ...rows.map(r=>header.map(h=>csvEscape(r[h])).join(','))].join('\n');
    downloadFile((project.name||'project').replace(/\W+/g,'_')+'_signage.csv', new Blob([csv], {type:'text/csv'}));
  }

  function exportXLSX(){
    const rows = toRows();
    const wb = XLSX.utils.book_new();
    const info = [
      ['Project', project.name],
      ['Exported', new Date().toLocaleString()],
      ['Total Signs', rows.length],
      ['Total Pages', (project.pages||[]).length],
      []
    ];
    const counts={};
    rows.forEach(r=> counts[r.sign_type]=(counts[r.sign_type]||0)+1 );
    info.push(['Breakdown']);
    Object.entries(counts).forEach(([k,v])=> info.push([k,v]));
    const wsInfo = XLSX.utils.aoa_to_sheet(info);
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, wsInfo, 'Project Info');
    XLSX.utils.book_append_sheet(wb, ws, 'Signage');
    const out = XLSX.write(wb, {bookType:'xlsx', type:'array'});
    downloadFile((project.name||'project').replace(/\W+/g,'_')+'_signage.xlsx',
      new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  }

  async function exportZIP(){
    const rows = toRows();
    const zip = new JSZip();
    // CSV
    const header = Object.keys(rows[0]||{placeholder:''});
    const csv = [header.join(','), ...rows.map(r=>header.map(h=>csvEscape(r[h])).join(','))].join('\n');
    zip.file('signage.csv', csv);
    // Photos
    (project.pages||[]).forEach(pg=>{
      (pg.pins||[]).forEach(pin=>{
        (pin.photos||[]).forEach((ph,idx)=>{
          const folder = zip.folder(`photos/${pin.id}`);
          const base = ph.name || `photo_${idx+1}.png`;
          const bin = dataURLtoArrayBuffer(ph.dataUrl);
          folder.file(base, bin);
        });
      });
    });
    const blob = await zip.generateAsync({type:'blob'});
    downloadFile((project.name||'project').replace(/\W+/g,'_')+'_export.zip', blob);
  }

  function dataURLtoArrayBuffer(dataURL){
    const bstr = atob((dataURL||'').split(',')[1]||'');
    const len=bstr.length; const buf=new Uint8Array(len);
    for(let i=0;i<len;i++) buf[i]=bstr.charCodeAt(i);
    return buf;
  }

  function importCSVText(text){
    const lines = text.split(/\r?\n/).filter(Boolean);
    if(lines.length<2) return;
    const hdr = lines[0].split(',').map(h=>h.trim());
    const idx = (k)=> hdr.indexOf(k);
    const pg = currentPage(); if(!pg) return;
    for(let i=1;i<lines.length;i++){
      const cells = parseCsvLine(lines[i]);
      const get = (k)=> cells[idx(k)] ?? '';
      const p = makePin(parseFloat(get('x_pct'))||50, parseFloat(get('y_pct'))||50);
      p.sign_type = get('sign_type')||'';
      p.custom_name = get('custom_name')||'';
      p.room_number = get('room_number')||'';
      p.room_name = get('room_name')||'';
      p.building = get('building')||'';
      p.level = get('level')||'';
      p.notes = get('notes')||'';
      pg.pins.push(p);
    }
    saveProject(project);
    renderPins(); renderPinsList();
    alert('CSV imported into current page.');
  }

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

  on(btnExportCSV,'click',exportCSV);
  on(btnExportXLSX,'click',exportXLSX);
  on(btnExportZIP,'click',exportZIP);
  on(btnImportCSV,'click',()=> inputImportCSV.click());
  on(inputImportCSV,'change', async (e)=>{
    const f=e.target.files?.[0]; if(!f) return;
    const text = await f.text();
    importCSVText(text);
    inputImportCSV.value='';
  });

  /******************
   * Keyboard: nudge pin
   ******************/
  window.addEventListener('keydown',(e)=>{
    if(['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key)){
      const p = selectedPin(); if(!p) return;
      e.preventDefault();
      const delta = project?.settings?.fieldMode ? 0.5 : 0.2;
      if(e.key==='ArrowUp') p.y_pct = clamp(p.y_pct - delta, 0, 100);
      if(e.key==='ArrowDown') p.y_pct = clamp(p.y_pct + delta, 0, 100);
      if(e.key==='ArrowLeft') p.x_pct = clamp(p.x_pct - delta, 0, 100);
      if(e.key==='ArrowRight') p.x_pct = clamp(p.x_pct + delta, 0, 100);
      p.lastEdited = Date.now();
      saveProject(project);
      renderPins();
      updatePinFields();
    }
  });

  /******************
   * First load – always start on Home
   ******************/
  showScreen('home');
  renderHomeRecent();
});
