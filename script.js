document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    let appState = {
        missionaries: [],
        settings: { language: 'en' },
        cities: ['Tete', 'Chimoio', 'Beira', 'Quelimane', 'Nampula', 'Nhamatanda', 'Marromeu', 'Caia'],
        // Simplified exceptions as the modal flow was removed for reliability
        cityExceptions: [
            { city: 'Quelimane', type: 'district', name: 'Quelimane' },
            { city: 'Nhamatanda', type: 'area', name: 'Nhamatanda' },
            { city: 'Marromeu', type: 'area', name: 'Marromeu' },
            { city: 'Caia', type: 'area', name: 'Caia' },
        ]
    };

    // --- DOM ELEMENTS ---
    const processFilesBtn = document.getElementById('process-files');
    const oldBoardInput = document.getElementById('old-transfer-board');
    const newBoardInput = document.getElementById('new-transfer-board');
    const mainContent = document.getElementById('main-content');
    const uploadSection = document.getElementById('upload-section');
    const masterTableBody = document.querySelector('#master-table tbody');
    const separateTablesContainer = document.getElementById('separate-tables-container');
    const deleteAllBtn = document.getElementById('delete-all-missionaries');
    const printBtn = document.getElementById('print-btn');
    const downloadExcelBtn = document.getElementById('download-excel-btn');

    // --- INITIALIZATION ---
    loadState();
    if (appState.missionaries.length > 0) {
        uploadSection.style.display = 'none';
        mainContent.style.display = 'block';
        renderTables();
    }

    // --- FILE PROCESSING ---
    processFilesBtn.addEventListener('click', () => {
        const oldFile = oldBoardInput.files[0];
        const newFile = newBoardInput.files[0];
        if (!oldFile || !newFile) { alert('Please upload both Excel files.'); return; }

        Promise.all([readExcelFile(oldFile), readExcelFile(newFile)])
            .then(([oldData, newData]) => {
                processData(oldData, newData);
                uploadSection.style.display = 'none';
                mainContent.style.display = 'block';
                renderTables();
            }).catch(error => {
                console.error("Error processing files:", error);
                alert("Could not process files. Ensure they are correct and not corrupted.");
            });
    });

    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const workbook = XLSX.read(event.target.result, { type: 'binary' });
                    resolve(XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]));
                } catch (e) { reject(e); }
            };
            reader.onerror = reject;
            reader.readAsBinaryString(file);
        });
    }

    function processData(oldData, newData) {
        const transferred = [];
        const oldDataMap = new Map(oldData.map(m => [m['Missionary Name'], m]));
        newData.forEach(newM => {
            const oldM = oldDataMap.get(newM['Missionary Name']);
            if (oldM && (oldM.Area !== newM.Area || oldM.Zone !== newM.Zone)) {
                 transferred.push({
                    type: (newM.Position?.includes('Sister')) ? 'Sister' : 'Elder', name: newM['Missionary Name'],
                    originZone: oldM.Zone, originDistrict: oldM.District, originArea: oldM.Area,
                    destinationZone: newM.Zone, destinationDistrict: newM.District, destinationArea: newM.Area,
                });
            }
        });
        appState.missionaries = transferred.map((m, index) => {
            const originCity = getCity(m.originZone, m.originDistrict, m.originArea);
            const destinationCity = getCity(m.destinationZone, m.destinationDistrict, m.destinationArea);
            const transport = getDefaultTransport(originCity, destinationCity);
            return { ...m, originCity, destinationCity, transport, date: '', time: '', tbd: false, instructions: '', leader: false, id: Date.now() + index };
        });
        saveState();
    }

    // --- CITY AND TRANSPORT LOGIC ---
    function getCity(zone = '', district = '', area = '') {
        const exception = appState.cityExceptions.find(ex => (ex.type === 'area' && ex.name === area) || (ex.type === 'district' && ex.name === district));
        if (exception) return exception.city;
        const upperZone = zone.toUpperCase();
        if (['ZONA MUNHAVA', 'ZONA INHAMIZUA', 'ZONA MANGA'].includes(upperZone)) return 'Beira';
        let cityName = zone.replace(/ZONA /i, '').trim();
        return cityName.charAt(0).toUpperCase() + cityName.slice(1).toLowerCase();
    }

    function getDefaultTransport(from, to) {
        const majorCities = ['Tete', 'Chimoio', 'Beira'];
        if (from === to && from === 'Beira') return 'Ride';
        if (from === to) return 'Txopela/Taxi';
        if (majorCities.includes(from) && majorCities.includes(to)) return 'Bus';
        if ((from === 'Quelimane' && to === 'Nampula') || (from === 'Nampula' && to === 'Quelimane')) return 'Bus';
        if ((majorCities.includes(from) && to === 'Nampula') || (from === 'Nampula' && majorCities.includes(to))) return 'Airplane';
        return 'Bus';
    }

    // --- TABLE RENDERING ---
    function renderTables() {
        renderMasterTable();
        renderSeparateTables();
    }

    function renderMasterTable() {
        masterTableBody.innerHTML = '';
        const zones = [...new Set(appState.missionaries.map(m => m.originZone))];
        const colors = ['#f8d7da', '#d1ecf1', '#d4edda', '#fff3cd', '#d6d8db', '#cce5ff', '#f5c6cb'];
        const zoneColorMap = zones.reduce((acc, zone, i) => ({ ...acc, [zone]: colors[i % colors.length] }), {});

        appState.missionaries.forEach(m => {
            const row = document.createElement('tr');
            row.style.backgroundColor = zoneColorMap[m.originZone];
            row.dataset.id = m.id;
            row.innerHTML = `
                <td>${m.type}</td>
                <td>${m.name}</td>
                <td>${m.originCity}</td>
                <td>${m.destinationCity}</td>
                <td>${m.destinationArea}</td>
                <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride', 'Boleia'], m.transport, 'transport')}</td>
                <td><input type="date" class="form-control form-control-sm date-input" data-field="date" value="${m.date}" ${m.tbd ? 'disabled' : ''}> <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div></td>
                <td><input type="time" class="form-control form-control-sm" data-field="time" value="${m.time}"></td>
                <td><textarea class="form-control form-control-sm" data-field="instructions">${m.instructions}</textarea></td>
                <td class="text-center align-middle"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? 'checked' : ''}></td>
                <td><button class="btn btn-danger btn-sm trash-btn"><i class="fas fa-trash"></i></button></td>
            `;
            masterTableBody.appendChild(row);
        });
    }

    function renderSeparateTables() {
        separateTablesContainer.innerHTML = '';
        const groups = appState.missionaries.reduce((acc, m) => {
            const key = `${m.originCity} to ${m.destinationCity}`;
            if (!acc[key]) acc[key] = [];
            acc[key].push(m);
            return acc;
        }, {});
        for (const groupName in groups) {
            const container = document.createElement('div');
            container.innerHTML = `<h3 class="mt-4">${groupName}</h3>`;
            const table = document.createElement('table');
            table.className = 'table table-bordered';
            table.innerHTML = `<thead class="table-dark"><tr><th>Name</th><th>Transport</th><th>Date & Time</th><th>Instructions</th><th>Leader</th></tr></thead>`;
            const tbody = document.createElement('tbody');
            groups[groupName].forEach(m => {
                tbody.innerHTML += `
                    <tr>
                        <td>${m.name} (${m.type})</td>
                        <td>${m.transport}</td>
                        <td>${m.tbd ? 'TBD' : `${m.date} ${m.time}`}</td>
                        <td>${m.instructions}</td>
                        <td>${m.leader ? 'Yes' : 'No'}</td>
                    </tr>
                `;
            });
            table.appendChild(tbody);
            container.appendChild(table);
        }
    }
    
    // --- EVENT LISTENERS & UI ---
    if (deleteAllBtn) deleteAllBtn.addEventListener('click', () => { if (confirm("Are you sure you want to delete all data?")) { localStorage.removeItem('missionaryTransferState'); window.location.reload(); } });
    if (printBtn) printBtn.addEventListener('click', () => window.print());
    if (downloadExcelBtn) downloadExcelBtn.addEventListener('click', exportToExcel);
    
    mainContent.addEventListener('change', (e) => {
        const target = e.target; const row = target.closest('tr'); if (!row) return;
        const id = parseFloat(row.dataset.id); const missionary = appState.missionaries.find(m => m.id === id); if (!missionary) return;
        
        if (target.matches('.tbd-checkbox')) missionary.tbd = target.checked;
        else if (target.matches('.leader-checkbox')) missionary.leader = target.checked;
        else if (target.dataset.field) missionary[target.dataset.field] = target.value;
        
        saveState();
        renderTables(); // Re-render both tables to keep them in sync
    });

    mainContent.addEventListener('click', (e) => {
         if (e.target.closest('.trash-btn')) {
            const row = e.target.closest('tr'); const id = parseFloat(row.dataset.id);
            if (confirm("Remove this missionary?")) {
                appState.missionaries = appState.missionaries.filter(m => m.id !== id); saveState(); renderTables();
            }
        }
    });

    // --- DATA EXPORT ---
    function exportToExcel() {
        const missionariesToExport = appState.missionaries.map(m => ({
            Type: m.type,
            Name: m.name,
            Origin: m.originCity,
            Destination: m.destinationCity,
            Area: m.destinationArea,
            Transportation: m.transport,
            Date: m.tbd ? 'TBD' : m.date,
            Time: m.tbd ? '' : m.time,
            Leader: m.leader ? 'Yes' : 'No',
            Instructions: m.instructions
        }));
        
        const ws = XLSX.utils.json_to_sheet(missionariesToExport);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Transfers");
        XLSX.writeFile(wb, "Missionary_Transfers.xlsx");
    }
    
    // --- HELPERS & LOCALSTORAGE ---
    function createDropdown(options, selected, field) { return `<select class="form-select form-select-sm" data-field="${field}">${options.map(o => `<option value="${o}" ${o === selected ? 'selected' : ''}>${o}</option>`).join('')}</select>`; }
    function saveState() { localStorage.setItem('missionaryTransferState', JSON.stringify(appState)); }
    function loadState() {
        const saved = localStorage.getItem('missionaryTransferState');
        if (saved) appState = {...appState, ...JSON.parse(saved)};
    }
});