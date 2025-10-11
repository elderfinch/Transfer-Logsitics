document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    let appState = {
        missionaries: [],
        settings: { language: 'en' },
        cityExceptions: [
            { city: 'Quelimane', type: 'district', name: 'Quelimane' },
            { city: 'Nhamatanda', type: 'area', name: 'Nhamatanda' },
            { city: 'Marromeu', type: 'area', name: 'Marromeu' },
            { city: 'Caia', type: 'area', name: 'Caia' },
        ],
        cities: ['Tete', 'Chimoio', 'Beira', 'Quelimane', 'Nampula', 'Nhamatanda', 'Marromeu', 'Caia']
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

        if (!oldFile || !newFile) {
            alert('Please upload both transfer board Excel files.');
            return;
        }

        Promise.all([readExcelFile(oldFile), readExcelFile(newFile)])
            .then(([oldData, newData]) => {
                processData(oldData, newData);
                
                uploadSection.style.display = 'none';
                mainContent.style.display = 'block';
                renderTables();
            })
            .catch(error => {
                console.error("Error processing files:", error);
                alert("There was an error reading the Excel files. Please ensure they are in the correct format.");
            });
    });

    /**
     * Reads an Excel file and converts the first sheet to an array of objects.
     */
    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = event.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet);
                    resolve(json);
                } catch (e) {
                    reject(e);
                }
            };
            reader.onerror = (error) => reject(error);
            reader.readAsBinaryString(file);
        });
    }

    /**
     * Compares old and new data to find transferred missionaries.
     */
    function processData(oldData, newData) {
        const transferred = [];
        const oldDataMap = new Map(oldData.map(m => [m['Missionary Name'], m]));

        newData.forEach(newM => {
            const oldM = oldDataMap.get(newM['Missionary Name']);
            if (oldM && (oldM.Area !== newM.Area || oldM.Zone !== newM.Zone)) {
                 transferred.push({
                    type: (newM.Position && newM.Position.includes('Sister')) ? 'Sister' : 'Elder',
                    name: newM['Missionary Name'],
                    originZone: oldM.Zone,
                    originDistrict: oldM.District,
                    originArea: oldM.Area,
                    destinationZone: newM.Zone,
                    destinationDistrict: newM.District,
                    destinationArea: newM.Area,
                });
            }
        });

        appState.missionaries = transferred.map((m, index) => {
            const originCity = getCity(m.originZone, m.originDistrict, m.originArea);
            const destinationCity = getCity(m.destinationZone, m.destinationDistrict, m.destinationArea);
            return {
                ...m,
                originCity,
                destinationCity,
                transport: getDefaultTransport(originCity, destinationCity),
                date: '',
                time: '',
                tbd: true,
                instructions: '',
                leader: false,
                id: Date.now() + index 
            };
        });
        saveState();
    }

    // --- CITY AND TRANSPORT LOGIC ---
    function getCity(zone = '', district = '', area = '') {
        const exception = appState.cityExceptions.find(ex =>
            (ex.type === 'area' && ex.name === area) ||
            (ex.type === 'district' && ex.name === district)
        );
        if (exception) return exception.city;

        const upperZone = zone.toUpperCase();
        if (['ZONA MUNHAVA', 'ZONA INHAMIZUA', 'ZONA MANGA'].includes(upperZone)) {
            return 'Beira';
        }
        
        // Improved default to handle names like "ZONA TETE" -> "Tete"
        return zone.replace(/ZONA /i, '').charAt(0).toUpperCase() + zone.slice(zone.indexOf(' ') + 1).toLowerCase();
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
        const colors = ['#f8d7da', '#d1ecf1', '#d4edda', '#fff3cd', '#d6d8db', '#cce5ff', '#f5c6cb', '#e2e3e5'];
        const zoneColorMap = zones.reduce((acc, zone, index) => {
            acc[zone] = colors[index % colors.length];
            return acc;
        }, {});

        appState.missionaries.forEach(m => {
            const row = document.createElement('tr');
            row.style.backgroundColor = zoneColorMap[m.originZone];
            row.dataset.id = m.id;
            row.innerHTML = `
                <td>${createDropdown(['Elder', 'Sister'], m.type, 'type')}</td>
                <td><input type="text" class="form-control" data-field="name" value="${m.name}"></td>
                <td>${createDropdown(appState.cities, m.originCity, 'originCity')}</td>
                <td>${createDropdown(appState.cities, m.destinationCity, 'destinationCity')}</td>
                <td><input type="text" class="form-control" data-field="destinationArea" value="${m.destinationArea}"></td>
                <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride'], m.transport, 'transport')}</td>
                <td>
                    <input type="date" class="form-control date-input" data-field="date" value="${m.date}" ${m.tbd ? 'disabled' : ''}>
                    <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div>
                </td>
                <td><input type="time" class="form-control" data-field="time" value="${m.time}"></td>
                <td><textarea class="form-control" data-field="instructions">${m.instructions}</textarea></td>
                <td><div class="form-check d-flex justify-content-center align-items-center h-100"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? 'checked' : ''}></div></td>
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
            const missionariesInGroup = groups[groupName];
            const container = document.createElement('div');
            container.innerHTML = `<h3 class="mt-4">${groupName}</h3>`;
            const table = document.createElement('table');
            table.className = 'table table-bordered';
            table.innerHTML = `<thead class="table-dark"><tr><th>Type</th><th>Name</th><th>Transportation</th><th>Date</th><th>Time</th><th>Instructions</th><th>Leader</th><th></th></tr></thead>`;
            const tbody = document.createElement('tbody');
            missionariesInGroup.forEach(m => {
                 const row = document.createElement('tr');
                 row.dataset.id = m.id;
                 row.innerHTML = `
                    <td>${m.type}</td>
                    <td>${m.name}</td>
                    <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride'], m.transport, 'transport')}</td>
                    <td>
                        <input type="date" class="form-control date-input" data-field="date" value="${m.date}" ${m.tbd ? 'disabled' : ''}>
                        <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div>
                    </td>
                    <td><input type="time" class="form-control" data-field="time" value="${m.time}"></td>
                    <td><textarea class="form-control" data-field="instructions">${m.instructions}</textarea></td>
                    <td><div class="form-check d-flex justify-content-center align-items-center h-100"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? 'checked' : ''}></div></td>
                    <td><button class="btn btn-danger btn-sm trash-btn"><i class="fas fa-trash"></i></button></td>
                 `;
                 tbody.appendChild(row);
            });
            table.appendChild(tbody);
            container.appendChild(table);
            separateTablesContainer.appendChild(container);
        }
    }

    // --- HELPER to create dropdowns ---
    function createDropdown(options, selectedValue, fieldName) {
        let optionsHTML = options.map(option =>
            `<option value="${option}" ${option === selectedValue ? 'selected' : ''}>${option}</option>`
        ).join('');
        return `<select class="form-select" data-field="${fieldName}">${optionsHTML}</select>`;
    }

    // --- EVENT LISTENERS on Tables & UI ---
    function handleTableInput(e) {
        const target = e.target;
        const row = target.closest('tr');
        if (!row) return;

        const id = parseFloat(row.dataset.id);
        const missionary = appState.missionaries.find(m => m.id === id);
        if (!missionary) return;

        // Handle different input types
        if (target.matches('.tbd-checkbox')) {
            missionary.tbd = target.checked;
            const dateInput = row.querySelector('.date-input');
            if (dateInput) dateInput.disabled = target.checked;
        } else if (target.matches('.leader-checkbox')) {
            missionary.leader = target.checked;
        } else if (target.dataset.field) {
            missionary[target.dataset.field] = target.value;
        }

        saveState();
        // A full re-render can be slow, let's just sync the other table if needed
        // For simplicity, we'll still re-render here.
        renderTables();
    }

    function handleTableClick(e) {
        if (e.target.closest('.trash-btn')) {
            const row = e.target.closest('tr');
            const id = parseFloat(row.dataset.id);
            const missionary = appState.missionaries.find(m => m.id === id);
            if (confirm(`Are you sure you want to remove ${missionary.name} from the travel plans?`)) {
                appState.missionaries = appState.missionaries.filter(m => m.id !== id);
                saveState();
                renderTables();
            }
        }
    }
    
    // Attach listeners to a parent element
    mainContent.addEventListener('change', handleTableInput);
    mainContent.addEventListener('click', handleTableClick);

    if (deleteAllBtn) {
        deleteAllBtn.addEventListener('click', () => {
            if (confirm('Are you sure you want to delete all missionary travel data? This cannot be undone.')) {
                appState.missionaries = [];
                saveState();
                localStorage.removeItem('missionaryTransferState');
                window.location.reload();
            }
        });
    }

    // --- LOCALSTORAGE ---
    function saveState() {
        localStorage.setItem('missionaryTransferState', JSON.stringify(appState));
    }

    function loadState() {
        const savedState = localStorage.getItem('missionaryTransferState');
        if (savedState) {
            appState = JSON.parse(savedState);
        }
    }
});