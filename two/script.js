// -- Constants, Defaults & State
    const LS_KEY = 'transfer_logistics_data_v6';
    const defaults = {
      exceptions: {
        'Quelimane': { city: 'Quelimane' }, 'Quelimane District': { city: 'Quelimane' }, 'Nhamatanda': { city: 'Nhamatanda' }, 'Marromeu': { city: 'Marromeu' }, 'Caia': { city: 'Caia' }, 'ZONA INHAMIZUA': { city: 'Beira' }, 'ZONA MANGA': { city: 'Beira' }, 'ZONA MUNHAVA': { city: 'Beira' }
      },
      columnSettings: {
          city: ['type', 'lastName', 'originArea', 'destArea', 'transport', 'date', 'time', 'instructions', 'new', 'leader'],
          transport: ['type', 'lastName', 'originCity', 'destCity', 'transport', 'date', 'time', 'instructions', 'new', 'leader'],
          master: ['type', 'lastName', 'originCity', 'destCity', 'transport', 'date', 'time', 'instructions', 'new', 'leader']
      }
    };
    let state = loadFromLS() || { groups: {}, exceptions: defaults.exceptions, lang: 'pt', columnSettings: JSON.parse(JSON.stringify(defaults.columnSettings)) };
    let fileOldRows = null; let fileNewRows = null;
    let mapping = { name: null, type: null, area: null, district: null, zone: null, companion: null };
    let pendingMissionaries = [];

function setLanguage(lang) {
        state.lang = lang;
        document.querySelectorAll('[data-translate]').forEach(el => {
            const key = el.getAttribute('data-translate');
            if (translations[lang] && translations[lang][key]) {
                const isPlaceholder = el.hasAttribute('placeholder');
                if (isPlaceholder) { el.setAttribute('placeholder', translations[lang][key]); } 
                else { el.textContent = translations[lang][key]; }
            }
        });
        document.querySelectorAll('.lang-select').forEach(sel => { sel.value = lang; M.FormSelect.init(sel); });
        populateTransportDropdowns();
        renderAllViews();
    }
    
    function populateTransportDropdowns() {
        const lang = state.lang || 'pt';
        const transportOptions = `<option value="Bus">${translations[lang].transport_bus}</option><option value="Plane">${translations[lang].transport_plane}</option><option value="Chapa">${translations[lang].transport_chapa}</option><option value="Txopela/Taxi">${translations[lang].transport_taxi}</option><option value="Ride">${translations[lang].transport_ride}</option>`;
        document.getElementById('add-transport').innerHTML = transportOptions;
        document.getElementById('edit-transport').innerHTML = transportOptions;
    }

    // -- State Management
    function loadFromLS() { try { return JSON.parse(localStorage.getItem(LS_KEY)); } catch(e){return null} }
    function saveToLS() { localStorage.setItem(LS_KEY, JSON.stringify(state)); }
    function loadState(){
        renderExceptions();
        renderAllViews();
        setupColumnManager();
        setLanguage(state.lang || 'pt');

        if (Object.keys(state.groups || {}).length > 0) {
            document.getElementById('upload-section').style.display = 'none';
            document.getElementById('controls-section').style.display = 'block';
            document.getElementById('tabs-section').style.display = 'block';
            M.Tabs.init(document.querySelector('.tabs'));
            document.querySelectorAll('.tabs .tab a').forEach(tab => {
                tab.addEventListener('click', () => {
                    setTimeout(() => setupColumnManager(), 50);
                });
            });
        }
    }

    // -- Helpers
    function getActiveView() {
        const activeTab = document.querySelector('.tabs .tab a.active');
        if (!activeTab) return 'city';
        return activeTab.getAttribute('href').replace('#tab-', '');
    }
    function normalizeZoneName(name){ if(!name) return ''; name = name.replace(/ZONA/i,'').trim(); return name.charAt(0).toUpperCase()+name.slice(1).toLowerCase(); }
    function autoResizeTextarea(element) { element.style.height = 'auto'; element.style.height = (element.scrollHeight) + 'px'; }
    function determineCity(loc) {
        if (loc.area && state.exceptions[loc.area]) return state.exceptions[loc.area].city;
        if (loc.district && state.exceptions[loc.district]) return state.exceptions[loc.district].city;
        if (loc.zone && state.exceptions[loc.zone]) return state.exceptions[loc.zone].city;
        return normalizeZoneName(loc.zone || loc.district || loc.area);
    }
    function findMissionary(missionaryId) {
        for (const key in state.groups) {
            const missionary = state.groups[key].find(m => m.id === missionaryId);
            if (missionary) return { missionary, groupKey: key };
        }
        return { missionary: null, groupKey: null };
    }


    // -- File Processing
    document.getElementById('btn-process').addEventListener('click', async ()=>{
      const fOld = document.getElementById('file-old').files[0]; const fNew = document.getElementById('file-new').files[0];
      if(!fOld || !fNew){ M.toast({html: translations[state.lang].toast_select_files}); return; }
      fileOldRows = await readFileRows(fOld); fileNewRows = await readFileRows(fNew);
      showMappingModal(fileOldRows[0] || fileNewRows[0]);
    });

    async function readFileRows(file){ const data = await file.arrayBuffer(); const wb = XLSX.read(data); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval: ''}); }

    function showMappingModal(sampleRow){
      const container = document.getElementById('mapping-forms'); container.innerHTML = '';
      const keys = Object.keys(sampleRow || {});
      const fields = [ {id:'name', label:'Name (Last, First)'}, {id:'type', label:'Type (Elder/Sister)'}, {id:'area', label:'Area/Neighborhood'}, {id:'district', label:'District'}, {id:'zone', label:'Zone'}, {id:'companion', label:'Companion'} ];
      fields.forEach(f=>{
        const div = document.createElement('div'); div.style.marginBottom = '20px';
        const h6 = document.createElement('h6'); h6.textContent = f.label;
        const select = document.createElement('select'); select.id = 'map-'+f.id; select.className = 'browser-default';
        select.innerHTML = `<option value="">-- Select column --</option>` + keys.map(k=>`<option value="${k}">${k}</option>`).join('');
        div.appendChild(h6); div.appendChild(select); container.appendChild(div);
      });
      M.Modal.getInstance(document.getElementById('modal-map')).open();
    }

    document.getElementById('btn-map-apply').addEventListener('click', ()=>{
      ['name','type','area', 'district', 'zone', 'companion'].forEach(f=>mapping[f]=document.getElementById('map-'+f).value || null);
      prepareExceptionsAssignment();
    });

    function prepareExceptionsAssignment(){
      const allUnique = { zones: new Set(), districts: new Set(), areas: new Set() };
      [fileOldRows,fileNewRows].forEach(rows=>{ (rows||[]).forEach(r=>{
        if (mapping.zone && r[mapping.zone]) allUnique.zones.add(r[mapping.zone].toString().trim());
        if (mapping.district && r[mapping.district]) allUnique.districts.add(r[mapping.district].toString().trim());
        if (mapping.area && r[mapping.area]) allUnique.areas.add(r[mapping.area].toString().trim());
      }); });
      const container = document.getElementById('exceptions-assign'); container.innerHTML='';
      const createSection = (title, list) => {
        if (list.size === 0) return;
        const titleEl = document.createElement('h5'); titleEl.textContent = title; container.appendChild(titleEl);
        Array.from(list).sort((x,y)=>x.localeCompare(y)).forEach(val=>{
          const row = document.createElement('div'); row.className='row valign-wrapper'; row.style.marginBottom = '0px';
          const existingCity = state.exceptions[val]?.city || '';
          row.innerHTML = `<div class="input-field col s5"><input value="${val}" readonly></div><div class="input-field col s5"><input class="ex-city" data-key="${val}" value="${existingCity}" placeholder="e.g., Beira" /></div><div class="col s2"><label><input type="checkbox" class="ex-apply" data-key="${val}" ${existingCity ? 'checked' : ''} /><span>Apply</span></label></div>`;
          container.appendChild(row);
        });
      };
      createSection('Area Overrides', allUnique.areas); createSection('District Overrides', allUnique.districts); createSection('Zone Overrides', allUnique.zones);
      M.AutoInit(); M.Modal.getInstance(document.getElementById('modal-exceptions')).open();
    }

    document.getElementById('btn-exceptions-apply').addEventListener('click', ()=>{
      document.querySelectorAll('.ex-apply').forEach(chk => {
          const key = chk.dataset.key; const city = document.querySelector(`.ex-city[data-key="${key}"]`).value.trim();
          if (chk.checked && city) { state.exceptions[key] = { city: city }; } 
          else if (state.exceptions[key] && !defaults.exceptions[key]) { delete state.exceptions[key]; }
      });
      state.exceptions = Object.assign({}, defaults.exceptions, state.exceptions);
      saveToLS(); prepareAndShowTransportModal();
    });
    
    function prepareAndShowTransportModal() {
        const oldIdx = {};
        (fileOldRows||[]).forEach(r=>{ const nameVal = mapping.name ? (r[mapping.name]||'') : ''; const key = nameVal.toString().trim().toLowerCase(); if(key) oldIdx[key]=r; });
        pendingMissionaries = [];
        (fileNewRows||[]).forEach(r=>{
            const nameVal = (mapping.name ? r[mapping.name] : '').toString().trim();
            const id = nameVal.toLowerCase(); if(!id) return;
            const nameParts = nameVal.split(',').map(p => p.trim());
            const old = oldIdx[id];
            const newLoc = { zone: (mapping.zone ? r[mapping.zone] : '').toString(), district: (mapping.district ? r[mapping.district] : '').toString(), area: (mapping.area ? r[mapping.area] : '').toString() };
            const baseMissionary = { id, name: nameVal, lastName: nameParts[0] || '', firstName: nameParts[1] || '', type: /^(elder|sister)$/i.test((r[mapping.type]||'').toString().trim()) ? (r[mapping.type]||'').toString().trim() : '', date: '', time: '', instructions:'', leader:false, companion: (mapping.companion && r[mapping.companion]) ? r[mapping.companion].toString().trim() : '' };
            if (old) {
                const oldLoc = { zone: (mapping.zone ? old[mapping.zone] : '').toString(), district: (mapping.district ? old[mapping.district] : '').toString(), area: (mapping.area ? old[mapping.area] : '').toString() };
                if (oldLoc.area !== newLoc.area) { // Simplified transfer condition
                    pendingMissionaries.push({ ...baseMissionary, originCity: determineCity(oldLoc), originArea: oldLoc.area, originDistrict: oldLoc.district, originZone: oldLoc.zone, destinationCity: determineCity(newLoc), destinationArea: newLoc.area, destinationDistrict: newLoc.district, destinationZone: newLoc.zone, isNew: false });
                }
            } else {
                pendingMissionaries.push({ ...baseMissionary, originCity: 'Beira', originArea: '', originDistrict: '', originZone: '', destinationCity: determineCity(newLoc), destinationArea: newLoc.area, destinationDistrict: newLoc.district, destinationZone: newLoc.zone, isNew: true });
            }
        });
        const uniqueRoutes = new Set(pendingMissionaries.map(m => `${m.originCity} -> ${m.destinationCity}`));
        const container = document.getElementById('transport-defaults-list'); container.innerHTML = ''; const lang = state.lang || 'pt'; const T = translations[lang];
        const transportOptions = `<option value="Bus">${T.transport_bus}</option><option value="Plane">${T.transport_plane}</option><option value="Chapa">${T.transport_chapa}</option><option value="Txopela/Taxi">${T.transport_taxi}</option><option value="Ride">${T.transport_ride}</option>`;
        Array.from(uniqueRoutes).sort().forEach(route => {
            const [origin, destination] = route.split(' -> ');
            const row = document.createElement('div'); row.className = 'row valign-wrapper';
            row.innerHTML = `<div class="col s6"><h6>${route}</h6></div><div class="col s6"><select class="browser-default transport-default-select" data-route="${route}">${transportOptions}</select></div>`;
            container.appendChild(row); row.querySelector('select').value = defaultTransportForPair(origin, destination);
        });
        M.Modal.getInstance(document.getElementById('modal-transport-defaults')).open();
    }

    document.getElementById('btn-transport-apply').addEventListener('click', () => {
        const transportDefaults = {}; document.querySelectorAll('.transport-default-select').forEach(sel => { transportDefaults[sel.dataset.route] = sel.value; });
        pendingMissionaries.forEach(m => { m.transport = transportDefaults[`${m.originCity} -> ${m.destinationCity}`] || defaultTransportForPair(m.originCity, m.destinationCity); });
        state.groups = {};
        pendingMissionaries.forEach(m => { const key = `${m.originCity} -> ${m.destinationCity}`; if (!state.groups[key]) state.groups[key] = []; state.groups[key].push(m); });
        saveToLS(); M.toast({html: translations[state.lang].toast_processed.replace('{count}', pendingMissionaries.length)}); location.reload();
    });

    function defaultTransportForPair(from, to){
      const f = (from||'').toLowerCase(), t = (to||'').toLowerCase();
      if((f === 'nampula' && t === 'quelimane') || (t === 'nampula' && f === 'quelimane')) return 'Bus'; if (f === 'nampula' || t === 'nampula') return 'Plane'; const main = ['tete','chimoio','beira']; if(main.includes(f) && main.includes(t) && f !== t) return 'Bus'; if(f === t){ if(f==='beira') return 'Ride'; return 'Txopela/Taxi'; } return 'Bus';
    }

    // -- UI Rendering
    function renderAllViews(){ renderGroupsByCity(); renderGroupsByTransport(); renderMasterTable(); populateCityDropdowns(); M.AutoInit(); applyColumnOrderAndVisibility(); }
    function renderGroupsByCity(){
      const container = document.getElementById('groups-by-city'); container.innerHTML = '';
      Object.keys(state.groups||{}).sort().forEach(key => {
        const card = document.createElement('div'); card.className='card group-card';
        const content = document.createElement('div'); content.className='card-content'; content.innerHTML = `<span class="card-title">${key}</span>`;
        content.appendChild(createMissionaryTable(state.groups[key], 'city', key)); card.appendChild(content); container.appendChild(card);
      });
    }
    function renderGroupsByTransport() {
        const container = document.getElementById('groups-by-transport'); container.innerHTML = '';
        const byTransport = {};
        Object.values(state.groups || {}).flat().forEach(m => {
            const transport = m.transport || 'Unassigned';
            if (!byTransport[transport]) byTransport[transport] = []; byTransport[transport].push(m);
        });
        Object.keys(byTransport).sort().forEach(key => {
            const card = document.createElement('div'); card.className = 'card group-card';
            const content = document.createElement('div'); content.className = 'card-content';
            const lang = state.lang || 'pt'; const T = translations[lang]; const transportName = T[`transport_${key.toLowerCase().replace('/','_')}`] || key;
            content.innerHTML = `<span class="card-title">${transportName}</span>`;
            content.appendChild(createMissionaryTable(byTransport[key], 'transport', key)); card.appendChild(content); container.appendChild(card);
        });
    }
    function renderMasterTable() {
        const container = document.getElementById('groups-master'); container.innerHTML = '';
        const allMissionaries = Object.values(state.groups || {}).flat();
        container.appendChild(createMissionaryTable(allMissionaries, 'master', null));
    }

    function createMissionaryTable(missionaries, groupType, groupKey) {
      const table = document.createElement('table'); table.className='striped table-col-toggle'; const lang = state.lang || 'pt'; const T = translations[lang];
      const sortedMissionaries = [...missionaries].sort((a,b) => (a.isNew - b.isNew) || a.lastName.localeCompare(b.lastName) || a.firstName.localeCompare(b.firstName));
      const allowBulkEdit = groupType === 'city' || groupType === 'transport';
      
      const bulkEditRow = allowBulkEdit ? `<tr class="bulk-edit-row">
            <th class="col-type"></th><th class="col-lastName"></th><th class="col-firstName"></th>
            <th class="col-originCity"></th><th class="col-destCity"></th>
            <th class="col-originArea"></th><th class="col-destArea"></th><th class="col-companion"></th>
            <th class="col-transport"><select class="bulk-transport-select browser-default"><option value="" selected disabled>${T.bulk_edit_transport_placeholder}</option><option value="Bus">${T.transport_bus}</option><option value="Plane">${T.transport_plane}</option><option value="Chapa">${T.transport_chapa}</option><option value="Txopela/Taxi">${T.transport_taxi}</option><option value="Ride">${T.transport_ride}</option></select></th>
            <th class="col-date"><input type="date" class="bulk-date-input"></th>
            <th class="col-time"><input type="time" class="bulk-time-input"></th>
            <th class="col-instructions"><textarea class="materialize-textarea bulk-instr-input" placeholder="${T.bulk_edit_instructions_placeholder}"></textarea></th>
            <th class="col-new"></th><th class="col-leader"></th><th class="col-actions"></th>
        </tr>` : '';

      table.innerHTML = `<thead>
        <tr class="header-row">
            <th class="col-type">${T.table_header_type}</th><th class="col-lastName">${T.table_header_last_name}</th><th class="col-firstName">${T.table_header_first_name}</th>
            <th class="col-originCity">${T.origin_city}</th><th class="col-destCity">${T.dest_city}</th>
            <th class="col-originArea">${T.table_header_origin_area}</th><th class="col-destArea">${T.table_header_dest_area}</th><th class="col-companion">${T.table_header_companion}</th>
            <th class="col-transport">${T.table_header_transport}</th><th class="col-date">${T.table_header_date}</th><th class="col-time">${T.table_header_time}</th>
            <th class="col-instructions">${T.table_header_instructions}</th><th class="col-new">${T.table_header_new}</th><th class="col-leader">${T.table_header_leader}</th>
            <th class="col-actions">${T.table_header_actions}</th>
        </tr>
        ${bulkEditRow}
        </thead>`;
      const tbody = document.createElement('tbody');
      sortedMissionaries.forEach(m =>{
          const tr = document.createElement('tr'); tr.dataset.id = m.id;
          let typeCellContent = (m.type === 'Elder' || m.type === 'Sister') ? `<span>${m.type}</span>` : `<select class="type-select-editable browser-default"><option value="" selected disabled>${T.choose_option}</option><option value="Elder">Elder</option><option value="Sister">Sister</option></select>`;

          tr.innerHTML = `
            <td class="col-type">${typeCellContent}</td><td class="col-lastName">${m.lastName}</td><td class="col-firstName">${m.firstName}</td>
            <td class="col-originCity">${m.originCity}</td><td class="col-destCity">${m.destinationCity}</td>
            <td class="col-originArea">${m.originArea || ''}</td><td class="col-destArea">${m.destinationArea}</td>
            <td class="col-companion"><span>${m.companion || ''}</span></td>
            <td class="col-transport"><select class="transport-select browser-default"><option value="Bus">${T.transport_bus}</option><option value="Plane">${T.transport_plane}</option><option value="Chapa">${T.transport_chapa}</option><option value="Txopela/Taxi">${T.transport_taxi}</option><option value="Ride">${T.transport_ride}</option></select></td>
            <td class="col-date"><input type="date" class="date-input" value="${m.date}"></td>
            <td class="col-time"><input type="time" class="time-input" value="${m.time}"></td>
            <td class="col-instructions"><textarea class="instr-input materialize-textarea">${m.instructions || ''}</textarea></td>
            <td class="col-new" style="text-align:center;">${m.isNew ? '<i class="material-icons green-text">check_box</i>' : ''}</td>
            <td class="col-leader"><label><input type="checkbox" class="leader-checkbox" ${m.leader ? 'checked':''}><span></span></label></td>
            <td class="col-actions">
            <div class="action-buttons">
                <a class="btn-small yellow calendar-btn modal-trigger" href="#modal-calendar" data-group-key="${groupKey}" data-group-type="${groupType}" data-lang="${lang}">
                    <i class="fa fa-calendar-alt"></i>
                </a>
                <a class="btn-small blue edit-btn"><i class="fa fa-pencil"></i></a>
                <a class="btn-small red remove-btn"><i class="fa fa-trash"></i></a>
            </div>
            </td>`;
          tbody.appendChild(tr);

          const instrTextarea = tr.querySelector('.instr-input');
          instrTextarea.addEventListener('input', () => autoResizeTextarea(instrTextarea)); setTimeout(() => autoResizeTextarea(instrTextarea), 0);
          tr.querySelector('.transport-select').value = m.transport;
          
          tr.querySelectorAll('input, select, textarea').forEach(el => {
              el.addEventListener('change', (e) => {
                  const trElement = e.target.closest('tr'); const missionaryId = trElement.dataset.id;
                  const { missionary } = findMissionary(missionaryId);
                  if (missionary) {
                      missionary.transport = trElement.querySelector('.transport-select').value; missionary.date = trElement.querySelector('.date-input').value; missionary.time = trElement.querySelector('.time-input').value; missionary.instructions = trElement.querySelector('.instr-input').value; missionary.leader = trElement.querySelector('.leader-checkbox').checked; 
                      if (el.classList.contains('type-select-editable')) { missionary.type = el.value; }
                      saveToLS();
                      if (el.classList.contains('type-select-editable')) renderAllViews();
                  }
              });
          });
          tr.querySelector('.remove-btn').addEventListener('click', (e)=>{ 
            if(confirm(T.confirm_remove_missionary)){
                const missionaryId = e.currentTarget.closest('tr').dataset.id;
                const { groupKey } = findMissionary(missionaryId);
                if (groupKey) {
                    const indexToRemove = state.groups[groupKey].findIndex(miss => miss.id === missionaryId);
                    if(indexToRemove > -1) { state.groups[groupKey].splice(indexToRemove, 1); if(state.groups[groupKey].length===0) delete state.groups[groupKey]; }
                }
                saveToLS(); renderAllViews();
            }
          });
          tr.querySelector('.edit-btn').addEventListener('click', (e) => openEditModal(e.currentTarget.closest('tr').dataset.id));
        });
      table.appendChild(tbody);
      if(allowBulkEdit && groupKey) {
        const bulkUpdate = (field, value) => {
            if(!value && field !== 'instructions') return;
            let missionariesToUpdate = [];
            if (groupType === 'city') {
                missionariesToUpdate = state.groups[groupKey] || [];
            } else if (groupType === 'transport') {
                missionariesToUpdate = Object.values(state.groups).flat().filter(m => (m.transport || 'Unassigned') === groupKey);
            }
            missionariesToUpdate.forEach(m => { m[field] = value; });
            saveToLS();
            renderAllViews();
            M.toast({html: T.toast_updated.replace('{field}', field)});
        };
        const bulkInstr = table.querySelector('.bulk-instr-input');
        if(bulkInstr) { bulkInstr.addEventListener('input', () => autoResizeTextarea(bulkInstr)); bulkInstr.addEventListener('change', (e) => bulkUpdate('instructions', e.target.value)); }
        table.querySelector('.bulk-transport-select').addEventListener('change', (e) => bulkUpdate('transport', e.target.value));
        table.querySelector('.bulk-date-input').addEventListener('change', (e) => bulkUpdate('date', e.target.value));
        table.querySelector('.bulk-time-input').addEventListener('change', (e) => bulkUpdate('time', e.target.value));
      }
      return table;
    }
    
    function populateCityDropdowns() {
        const cities = new Set(Object.values(state.groups || {}).flat().flatMap(m => [m.originCity, m.destinationCity]));
        const sortedCities = Array.from(cities).sort(); const lang = state.lang || 'pt';
        const selects = document.querySelectorAll('#add-origin-select, #add-dest-select, #edit-origin-select, #edit-dest-select');
        selects.forEach(select => {
            select.innerHTML = `<option value="" disabled selected>${translations[lang].choose_option}</option>${sortedCities.map(c=>`<option value="${c}">${c}</option>`).join('')}<option value="other">${translations[lang].other_option}</option>`;
        });
    }

    function openEditModal(missionaryId) {
        const { missionary } = findMissionary(missionaryId); if (!missionary) return;
        document.getElementById('edit-id').value = missionary.id;
        document.getElementById('edit-name').value = missionary.name; document.getElementById('edit-type').value = missionary.type;
        const setupSelect = (type, city) => {
            const select = document.getElementById(`edit-${type}-select`), other = document.getElementById(`edit-${type}-other`);
            if (Array.from(select.options).some(o => o.value === city)) { select.value = city; other.style.display = 'none'; } 
            else { select.value = 'other'; other.value = city; other.style.display = 'block'; }
        };
        setupSelect('origin', missionary.originCity); setupSelect('dest', missionary.destinationCity);
        document.getElementById('edit-originarea').value = missionary.originArea || ''; document.getElementById('edit-destarea').value = missionary.destinationArea; document.getElementById('edit-transport').value = missionary.transport; document.getElementById('edit-companion').value = missionary.companion;
        document.getElementById('edit-date').value = missionary.date; document.getElementById('edit-time').value = missionary.time; document.getElementById('edit-instructions').value = missionary.instructions; document.getElementById('edit-leader').checked = missionary.leader;
        M.Modal.getInstance(document.getElementById('modal-edit')).open();
    }
    
    // -- Column Visibility & Reordering
    const ALL_COLUMNS = [
        { key: 'type', labelKey: 'table_header_type' }, { key: 'lastName', labelKey: 'table_header_last_name' }, { key: 'firstName', labelKey: 'table_header_first_name' },
        { key: 'originCity', labelKey: 'origin_city' }, { key: 'destCity', labelKey: 'dest_city' }, { key: 'originArea', labelKey: 'table_header_origin_area' },
        { key: 'destArea', labelKey: 'table_header_dest_area' }, { key: 'companion', labelKey: 'table_header_companion' }, { key: 'transport', labelKey: 'table_header_transport' },
        { key: 'date', labelKey: 'table_header_date' }, { key: 'time', labelKey: 'table_header_time' }, { key: 'instructions', labelKey: 'table_header_instructions' },
        { key: 'new', labelKey: 'table_header_new' }, { key: 'leader', labelKey: 'table_header_leader' }
    ];

    function setupColumnManager() {
        const view = getActiveView();
        if (!state.columnSettings[view]) {
            state.columnSettings[view] = [...defaults.columnSettings[view]];
        }
        const visibleContainer = document.getElementById('visible-cols-container');
        const hiddenContainer = document.getElementById('hidden-cols-container');
        visibleContainer.innerHTML = ''; hiddenContainer.innerHTML = '';
        const lang = state.lang || 'pt';
        const visibleKeys = new Set(state.columnSettings[view]);

        state.columnSettings[view].forEach(key => {
            const col = ALL_COLUMNS.find(c => c.key === key);
            if(col) visibleContainer.appendChild(createColumnChip(col, true, lang));
        });

        ALL_COLUMNS.forEach(col => {
            if (!visibleKeys.has(col.key)) {
                hiddenContainer.appendChild(createColumnChip(col, false, lang));
            }
        });
        
        const sortableOptions = { group: 'columns', animation: 150, onEnd: updateColumnState };
        new Sortable(visibleContainer, sortableOptions);
        new Sortable(hiddenContainer, sortableOptions);
        applyColumnOrderAndVisibility();
    }
    
    function createColumnChip(col, isVisible, lang) {
        const chip = document.createElement('div');
        chip.className = 'col-chip';
        chip.dataset.key = col.key;
        chip.innerHTML = `${translations[lang][col.labelKey]} <i class="material-icons visibility-toggle">${isVisible ? 'visibility' : 'visibility_off'}</i>`;
        return chip;
    }

    document.getElementById('column-manager-wrapper').addEventListener('click', e => {
        if (e.target.classList.contains('visibility-toggle')) {
            const chip = e.target.closest('.col-chip');
            const targetContainer = chip.parentElement.id === 'visible-cols-container' ? document.getElementById('hidden-cols-container') : document.getElementById('visible-cols-container');
            targetContainer.appendChild(chip);
            updateColumnState();
        }
    });

    function updateColumnState() {
        const view = getActiveView();
        const visibleContainer = document.getElementById('visible-cols-container');
        state.columnSettings[view] = [...visibleContainer.children].map(chip => chip.dataset.key);
        saveToLS();
        applyColumnOrderAndVisibility();
        setupColumnManager();
    }
    
    function applyColumnOrderAndVisibility() {
        const view = getActiveView();
        const visibleColumns = state.columnSettings[view] || [];
        ALL_COLUMNS.forEach(col => document.body.classList.toggle(`show-${col.key}`, visibleColumns.includes(col.key)));
        document.querySelectorAll('.table-col-toggle').forEach(table => {
            table.querySelectorAll('thead tr, tbody tr').forEach(row => {
                const cells = new Map([...row.children].map(cell => {
                    const key = [...cell.classList].find(c => c.startsWith('col-')).replace('col-', '');
                    return [key, cell];
                }));
                visibleColumns.forEach(key => {
                    if (cells.has(key)) row.appendChild(cells.get(key));
                });
                if(cells.has('actions')) row.appendChild(cells.get('actions'));
            });
        });
    }

    // -- Event Listeners
    ['add', 'edit'].forEach(prefix => {
        document.getElementById(`${prefix}-origin-select`).addEventListener('change', (e) => { document.getElementById(`${prefix}-origin-other`).style.display = e.target.value === 'other' ? 'block' : 'none'; });
        document.getElementById(`${prefix}-dest-select`).addEventListener('change', (e) => { document.getElementById(`${prefix}-dest-other`).style.display = e.target.value === 'other' ? 'block' : 'none'; });
    });

    document.getElementById('btn-add-save').addEventListener('click', ()=>{
      const originCity = document.getElementById('add-origin-select').value === 'other' ? document.getElementById('add-origin-other').value : document.getElementById('add-origin-select').value;
      const destCity = document.getElementById('add-dest-select').value === 'other' ? document.getElementById('add-dest-other').value : document.getElementById('add-dest-select').value;
      const nameVal = document.getElementById('add-name').value.trim();
      const nameParts = nameVal.split(',').map(p => p.trim());
      const m = {
        id: nameVal.toLowerCase(), name: nameVal, lastName: nameParts[0]||'', firstName: nameParts[1]||'', type: document.getElementById('add-type').value, 
        originCity: normalizeZoneName(originCity), destinationCity: normalizeZoneName(destCity),
        originArea: document.getElementById('add-originarea').value, destinationArea: document.getElementById('add-destarea').value, 
        transport: document.getElementById('add-transport').value, companion: document.getElementById('add-companion').value,
        date: document.getElementById('add-date').value, time: document.getElementById('add-time').value, 
        instructions: document.getElementById('add-instructions').value, leader: document.getElementById('add-leader').checked, isNew: document.getElementById('add-new').checked
      };
      const key = `${m.originCity} -> ${m.destinationCity}`; if(!state.groups[key]) state.groups[key]=[]; state.groups[key].push(m);
      saveToLS(); M.toast({html: translations[state.lang].toast_added}); location.reload();
    });
    
    document.getElementById('btn-edit-save').addEventListener('click', () => {
        const id = document.getElementById('edit-id').value;
        const { missionary, groupKey } = findMissionary(id);
        if (!missionary) return;
        const originCity = document.getElementById('edit-origin-select').value === 'other' ? document.getElementById('edit-origin-other').value : document.getElementById('edit-origin-select').value;
        const destCity = document.getElementById('edit-dest-select').value === 'other' ? document.getElementById('edit-dest-other').value : document.getElementById('edit-dest-select').value;
        const nameVal = document.getElementById('edit-name').value.trim();
        const nameParts = nameVal.split(',').map(p => p.trim());

        Object.assign(missionary, {
            id: nameVal.toLowerCase(), name: nameVal, lastName: nameParts[0]||'', firstName: nameParts[1]||'', type: document.getElementById('edit-type').value, 
            originCity: normalizeZoneName(originCity), destinationCity: normalizeZoneName(destCity),
            originArea: document.getElementById('edit-originarea').value, destinationArea: document.getElementById('edit-destarea').value, 
            transport: document.getElementById('edit-transport').value, companion: document.getElementById('edit-companion').value,
            date: document.getElementById('edit-date').value, time: document.getElementById('edit-time').value, 
            instructions: document.getElementById('edit-instructions').value, leader: document.getElementById('edit-leader').checked
        });
        
        const newGroupKey = `${missionary.originCity} -> ${missionary.destinationCity}`;
        if (newGroupKey !== groupKey) {
            const index = state.groups[groupKey].findIndex(m => m.id === id);
            state.groups[groupKey].splice(index, 1);
            if (state.groups[groupKey].length === 0) delete state.groups[groupKey];
            if (!state.groups[newGroupKey]) state.groups[newGroupKey] = [];
            state.groups[newGroupKey].push(missionary);
        }
        saveToLS(); location.reload();
    });
    
    document.getElementById('btn-save-exception').addEventListener('click', () => {
      const name = document.getElementById('ex-add-name').value.trim(), city = document.getElementById('ex-add-city').value.trim();
      if (name && city) { if (!state.exceptions) state.exceptions = {}; state.exceptions[name] = { city }; saveToLS(); renderExceptions(); document.getElementById('ex-add-name').value = ''; document.getElementById('ex-add-city').value = ''; M.updateTextFields(); }
    });

    document.getElementById('btn-reset-cols').addEventListener('click', () => {
        const view = getActiveView();
        state.columnSettings[view] = [...defaults.columnSettings[view]];
        saveToLS();
        setupColumnManager();
        applyColumnOrderAndVisibility();
        M.toast({html: translations[state.lang].toast_columns_reset});
    });

    function renderExceptions(){ const el = document.getElementById('exceptions-list'); el.innerHTML=''; Object.keys(state.exceptions||{}).forEach(k => { const row = document.createElement('div'); row.className='row'; row.innerHTML = `<div class="col s4">${k}</div><div class="col s4">${state.exceptions[k].city || ''}</div><div class="col s4"><a data-key="${k}" class="btn-small red btn-ex-remove">Remove</a></div>`; el.appendChild(row); row.querySelector('.btn-ex-remove').addEventListener('click', (e)=>{ delete state.exceptions[e.currentTarget.dataset.key]; saveToLS(); renderExceptions(); }); }); }
    
    document.getElementById('btn-clear-local').addEventListener('click', ()=>{ if(confirm(translations[state.lang].confirm_clear_local)){ localStorage.removeItem(LS_KEY); location.reload(); } });
    document.querySelectorAll('.lang-select').forEach(sel => { sel.addEventListener('change', (e)=>{ state.lang = e.target.value; saveToLS(); location.reload(); }); });\

    document.getElementById('btn-restore').addEventListener('click', async () => {
        const file = document.getElementById('file-restore').files[0]; if (!file) { M.toast({ html: translations[state.lang].toast_restore_file }); return; }
        const data = await file.text(); const rows = data.split('\n').slice(1); const restoredMissionaries = [];
        for (const row of rows) {
            if (!row.trim()) continue;
            const cols = row.split(';').map(col => col.startsWith('"') && col.endsWith('"') ? col.slice(1, -1).replace(/""/g, '"') : col);
            const transportVal = Object.keys(translations.en).find(k => k.startsWith('transport_') && translations.en[k] === cols[8]) || Object.keys(translations.pt).find(k => k.startsWith('transport_') && translations.pt[k] === cols[8]);
            const missionary = {
                lastName: cols[0], firstName: cols[1], type: cols[2], originCity: cols[3], destinationCity: cols[4], originArea: cols[5], destinationArea: cols[6], companion: cols[7],
                transport: (transportVal || '').replace('transport_', ''), date: cols[9], time: cols[10], instructions: cols[11],
                isNew: cols[12]?.toLowerCase() === 'yes', leader: cols[13]?.toLowerCase() === 'yes', 
                name: `${cols[0]}, ${cols[1]}`, id: `${cols[0]}, ${cols[1]}`.toLowerCase()
            };
            restoredMissionaries.push(missionary);
        }
        state.groups = {};
        restoredMissionaries.forEach(m => { const key = `${m.originCity} -> ${m.destinationCity}`; if (!state.groups[key]) state.groups[key] = []; state.groups[key].push(m); });
        saveToLS(); M.toast({ html: translations[state.lang].toast_restored }); location.reload();
    });
    
    // Global variable to hold the list of missionaries for the currently open group
    let currentMissionaryGroup = [];
    let currentGroupKey = null;

    function formatCalendarDate(date, time) {
        if (!date) return 'N/A';
        const dateObj = new Date(`${date}T${time || '00:00:00'}`);
        // Basic formatting, assuming local timezone is sufficient for display
        return dateObj.toLocaleString(state.lang, { 
            weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', 
            hour: '2-digit', minute: '2-digit', hour12: true 
        });
    }

    function createCalendarDescription() {
        const T = translations[state.lang];
        let description = ``;
        
        const selectedIds = new Set(Array.from(document.querySelectorAll('#missionaries-checkbox-list input[type="checkbox"]:checked')).map(chk => chk.value));
        const selectedMissionaries = currentMissionaryGroup.filter(m => selectedIds.has(m.id));
        
        // Use unique cities from selected missionaries
        const originCities = [...new Set(selectedMissionaries.map(m => m.originCity))].join(', ');
        const destCities = [...new Set(selectedMissionaries.map(m => m.destinationCity))].join(', ');

        description += `${T.origin_city_label}: ${originCities}\n`;
        description += `${T.destination_city_label}: ${destCities}\n\n`;
        description += `${T.missionaries_names}:\n`;

        selectedMissionaries.forEach(m => {
            const transportName = m.transport ? (T[`transport_${m.transport.toLowerCase().replace('/','_')}`] || m.transport) : '';
            description += `- ${m.name} (${m.type}, ${transportName}, ${m.companion || ''})\n`;
        });
        
        return description.trim();
    }
    
    function updateDescriptionPreview() {
        document.getElementById('calendar-description-preview').textContent = createCalendarDescription();
    }

    document.querySelector('body').addEventListener('click', (e) => {
        if (e.target.closest('.calendar-btn')) {
            const btn = e.target.closest('.calendar-btn');
            currentGroupKey = btn.dataset.groupKey;
            const groupType = btn.dataset.groupType;

            // Get missionaries based on group type and key
            if (groupType === 'city' && state.groups[currentGroupKey]) {
                currentMissionaryGroup = state.groups[currentGroupKey];
            } else if (groupType === 'transport') {
                currentMissionaryGroup = Object.values(state.groups).flat().filter(m => (m.transport || 'Unassigned') === currentGroupKey);
            } else { // Master List (use a common attribute like date to group)
                // For Master List, find the group by date/time if available, or just use all.
                const tr = btn.closest('tr');
                const missionaryId = tr ? tr.dataset.id : null;
                const m = missionaryId ? findMissionary(missionaryId).missionary : null;
                
                if (m && m.date) {
                    currentGroupKey = `${m.date} ${m.time || ''}`;
                    currentMissionaryGroup = Object.values(state.groups).flat().filter(
                        miss => miss.date === m.date && (miss.time || '') === (m.time || '')
                    );
                } else {
                    currentMissionaryGroup = Object.values(state.groups).flat(); // Fallback to all
                }
            }

            // Get the first missionary for time/date reference
            const refMissionary = currentMissionaryGroup.find(m => m.date) || currentMissionaryGroup[0];

            // Set modal values
            const formattedDateTime = refMissionary ? formatCalendarDate(refMissionary.date, refMissionary.time) : translations[state.lang].choose_option;
            document.getElementById('calendar-time-date').textContent = formattedDateTime;

            // Populate checkbox list
            const checkboxList = document.getElementById('missionaries-checkbox-list');
            checkboxList.innerHTML = '';
            
            currentMissionaryGroup.sort((a,b) => a.lastName.localeCompare(b.lastName)).forEach(m => {
                const p = document.createElement('p');
                p.innerHTML = `<label>
                    <input type="checkbox" class="missionary-checkbox" value="${m.id}" checked />
                    <span>${m.name} (${m.originCity} &rarr; ${m.destinationCity})</span>
                </label>`;
                checkboxList.appendChild(p);
            });
            
            // Initial description and event listeners for updates
            updateDescriptionPreview();
            checkboxList.querySelectorAll('.missionary-checkbox').forEach(chk => {
                chk.addEventListener('change', updateDescriptionPreview);
            });

            M.Modal.getInstance(document.getElementById('modal-calendar')).open();
        }
    });

    // Calendar Export Buttons
    document.getElementById('btn-export-google').addEventListener('click', () => {
        const title = document.getElementById('calendar-event-title').value;
        const description = createCalendarDescription();
        const firstMissionary = currentMissionaryGroup.find(m => m.date) || currentMissionaryGroup[0];
        
        if (!firstMissionary || !firstMissionary.date) {
            M.toast({html: 'Please set a date and time for at least one missionary in the group.'});
            return;
        }

        const date = firstMissionary.date;
        const time = firstMissionary.time || '00:00';
        
        // Helper to convert YYYY-MM-DD and HH:MM to Google Calendar format YYYYMMDDTHHMMSSZ (UTC or local with no Z)
        function formatDateTimeForGoogle(dateStr, timeStr) {
            const dt = new Date(`${dateStr}T${timeStr}`);
            const offsetHours = dt.getTimezoneOffset() / 60;
            // Simplest for local time: YYYYMMDDTHHMM00
            return `${dateStr.replace(/-/g, '')}T${timeStr.replace(/:/g, '')}00`;
        }
        
        const startTime = formatDateTimeForGoogle(date, time);
        // Assume 2 hours duration for simplicity, or just use start time for all-day like entry.
        const dtEnd = new Date(`${date}T${time}`);
        dtEnd.setHours(dtEnd.getHours() + 2); 
        const endTime = formatDateTimeForGoogle(dtEnd.toISOString().substring(0, 10), dtEnd.toTimeString().substring(0, 5));
        
        const cityInfo = [...new Set(currentMissionaryGroup.map(m => m.originCity))].join(', ');

        const googleUrl = new URL('https://calendar.google.com/calendar/render');
        googleUrl.searchParams.set('action', 'TEMPLATE');
        googleUrl.searchParams.set('text', title);
        googleUrl.searchParams.set('dates', `${startTime}/${endTime}`);
        googleUrl.searchParams.set('details', description);
        googleUrl.searchParams.set('location', cityInfo);
        
        window.open(googleUrl.toString(), '_blank');
    });

    document.getElementById('btn-export-outlook').addEventListener('click', () => {
        const title = document.getElementById('calendar-event-title').value;
        const description = createCalendarDescription();
        const firstMissionary = currentMissionaryGroup.find(m => m.date) || currentMissionaryGroup[0];
        
        if (!firstMissionary || !firstMissionary.date) {
            M.toast({html: 'Please set a date and time for at least one missionary in the group.'});
            return;
        }

        const date = firstMissionary.date;
        const time = firstMissionary.time || '00:00';
        
        // Helper to convert YYYY-MM-DD and HH:MM to iCalendar format (UTC or local with no Z)
        function formatDateTimeForOutlook(dateStr, timeStr) {
            const dt = new Date(`${dateStr}T${timeStr}`);
            // Use ISO format without milliseconds and Z for local time, then clean up
            return dt.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '').substring(0, 15);
        }
        
        const startTime = formatDateTimeForOutlook(date, time);
        // Assume 2 hours duration for simplicity
        const dtEnd = new Date(`${date}T${time}`);
        dtEnd.setHours(dtEnd.getHours() + 2); 
        const endTime = formatDateTimeForOutlook(dtEnd.toISOString().substring(0, 10), dtEnd.toTimeString().substring(0, 5));
        
        const cityInfo = [...new Set(currentMissionaryGroup.map(m => m.originCity))].join(', ');

        // Outlook/Office 365 URL uses slightly different parameters
        const outlookUrl = new URL('https://outlook.live.com/calendar/0/action/compose');
        outlookUrl.searchParams.set('subject', title);
        outlookUrl.searchParams.set('body', description);
        outlookUrl.searchParams.set('location', cityInfo);
        outlookUrl.searchParams.set('startdt', `${startTime}`); 
        outlookUrl.searchParams.set('enddt', `${endTime}`);
        
        window.open(outlookUrl.toString(), '_blank');
    });

    // -- Initial Load
    document.addEventListener('DOMContentLoaded', ()=>{
        const ls = loadFromLS(); if(ls){ state=ls; }
        state.exceptions = Object.assign({}, defaults.exceptions, state.exceptions||{});
        if (!state.columnSettings) { state.columnSettings = JSON.parse(JSON.stringify(defaults.columnSettings)); }
        Object.values(state.groups || {}).flat().forEach(m => delete m.tbd); // Clean up old data
        saveToLS();
        M.AutoInit();
        loadState();
    });