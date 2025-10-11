document.addEventListener('DOMContentLoaded', () => {
    // --- STATE MANAGEMENT ---
    let appState = {
        missionaries: [],
        settings: { language: 'en' },
        cityExceptions: [
            // Default exceptions
            { city: 'Quelimane', type: 'district', name: 'Quelimane' },
            { city: 'Nhamatanda', type: 'area', name: 'Nhamatanda' },
            { city: 'Marromeu', type: 'area', name: 'Marromeu' },
            { city: 'Caia', type: 'area', name: 'Caia' },
        ],
        // A list of all possible cities for dropdowns
        cities: ['Tete', 'Chimoio', 'Beira', 'Quelimane', 'Nampula', 'Nhamatanda', 'Marromeu', 'Caia']
    };

    // --- DOM ELEMENTS ---
    const processFilesBtn = document.getElementById('process-files');
    const oldBoardInput = document.getElementById('old-transfer-board');
    const newBoardInput = document.getElementById('new-transfer-board');
    const mainContent = document.getElementById('main-content');
    const masterTableBody = document.querySelector('#master-table tbody');
    const separateTablesContainer = document.getElementById('separate-tables-container');
    
    // ... (Add other DOM elements you need for modals and buttons)

    // --- INITIALIZATION ---
    loadState();
    if (appState.missionaries.length > 0) {
        document.getElementById('upload-section').style.display = 'none';
        mainContent.style.display = 'block';
        renderTables();
    }
    // updateUIText(); // You can implement this for language switching

    // --- FILE PROCESSING ---
    processFilesBtn.addEventListener('click', () => {
        const oldFile = oldBoardInput.files[0];
        const newFile = newBoardInput.files[0];

        if (!oldFile || !newFile) {
            alert('Please upload both transfer board Excel files.');
            return;
        }

        // Use Promise.all to handle both file readings asynchronously
        Promise.all([readExcelFile(oldFile), readExcelFile(newFile)])
            .then(([oldData, newData]) => {
                processData(oldData, newData);
                
                // Once data is processed, show the main content and hide upload
                document.getElementById('upload-section').style.display = 'none';
                mainContent.style.display = 'block';
                renderTables();
                
                // You can open your city exceptions modal here if needed
                // openCityExceptionsModal();
            })
            .catch(error => {
                console.error("Error processing files:", error);
                alert("There was an error reading the Excel files. Please ensure they are in the correct format.");
            });
    });

    /**
     * Reads an Excel file and converts the first sheet to an array of objects.
     * @param {File} file - The Excel file to read.
     * @returns {Promise<Array<Object>>} A promise that resolves with the data.
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
        
        // Create a map of the old data for quick lookups
        const oldDataMap = new Map(oldData.map(m => [m['Missionary Name'], m]));

        newData.forEach(newM => {
            const oldM = oldDataMap.get(newM['Missionary Name']);
            // Check if the missionary existed before and their area/zone has changed
            if (oldM && (oldM.Area !== newM.Area || oldM.Zone !== newM.Zone)) {
                 transferred.push({
                    type: newM.Position.includes('Sister') ? 'Sister' : 'Elder',
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
                id: Date.now() + index // Simple unique ID
            };
        });
        saveState();
    }


    // --- CITY AND TRANSPORT LOGIC ---
    function getCity(zone = '', district = '', area = '') {
        const upperZone = zone.toUpperCase();
        // Check for specific exceptions first
        const exception = appState.cityExceptions.find(ex =>
            (ex.type === 'area' && ex.name === area) ||
            (ex.type === 'district' && ex.name === district)
        );
        if (exception) return exception.city;

        // Zone-based rules for Beira
        if (['ZONA MUNHAVA', 'ZONA INHAMIZUA', 'ZONA MANGA'].includes(upperZone)) {
            return 'Beira';
        }

        // Default: normalize zone name
        return zone.replace(/ZONA /i, '').charAt(0).toUpperCase() + zone.slice(zone.indexOf(' ') + 1).toLowerCase();
    }


    function getDefaultTransport(from, to) {
        const majorCities = ['Tete', 'Chimoio', 'Beira'];
        if (from === to && from === 'Beira') return 'Ride';
        if (from === to) return 'Txopela/Taxi';
        if (majorCities.includes(from) && majorCities.includes(to)) return 'Bus';
        if ((from === 'Quelimane' && to === 'Nampula') || (from === 'Nampula' && to === 'Quelimane')) return 'Bus';
        if ((majorCities.includes(from) && to === 'Nampula') || (from === 'Nampula' && majorCities.includes(to))) return 'Airplane';
        return 'Bus'; // A sensible default
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

            // ... The logic to build the cells for the row goes here ...
            // This needs to be fully implemented with inputs, dropdowns, etc.
            row.innerHTML = `
                <td>${createDropdown(['Elder', 'Sister'], m.type)}</td>
                <td><input type="text" class="form-control" value="${m.name}"></td>
                <td>${createDropdown(appState.cities, m.originCity)}</td>
                <td>${createDropdown(appState.cities, m.destinationCity)}</td>
                <td><input type="text" class="form-control" value="${m.destinationArea}"></td>
                <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride'], m.transport)}</td>
                <td>
                    <input type="date" class="form-control" value="${m.date}" ${m.tbd ? 'disabled' : ''}>
                    <div class="form-check"><input class="form-check-input" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div>
                </td>
                <td><input type="time" class="form-control" value="${m.time}"></td>
                <td><textarea class="form-control">${m.instructions}</textarea></td>
                <td><div class="form-check"><input class="form-check-input" type="checkbox" ${m.leader ? 'checked' : ''}></div></td>
                <td><button class="btn btn-danger btn-sm"><i class="fas fa-trash"></i></button></td>
            `;

            masterTableBody.appendChild(row);
        });
        addTableEventListeners();
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
            container.innerHTML = `<h3>${groupName}</h3>`;
            const table = document.createElement('table');
            table.className = 'table table-bordered';
            table.innerHTML = `
                <thead class="table-dark">
                    <tr>
                        <th>Type</th><th>Name</th><th>Transportation</th><th>Date</th><th>Time</th><th>Instructions</th><th>Leader</th><th></th>
                    </tr>
                </thead>
            `;
            const tbody = document.createElement('tbody');
            missionariesInGroup.forEach(m => {
                 const row = document.createElement('tr');
                 row.dataset.id = m.id;
                 // Add cells for the separate table view
                 row.innerHTML = `
                    <td>${createDropdown(['Elder', 'Sister'], m.type)}</td>
                    <td><input type="text" class="form-control" value="${m.name}" disabled></td>
                    <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride'], m.transport)}</td>
                    <td>
                        <input type="date" class="form-control" value="${m.date}" ${m.tbd ? 'disabled' : ''}>
                        <div class="form-check"><input class="form-check-input" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div>
                    </td>
                    <td><input type="time" class="form-control" value="${m.time}"></td>
                    <td><textarea class="form-control">${m.instructions}</textarea></td>
                    <td><div class="form-check"><input class="form-check-input" type="checkbox" ${m.leader ? 'checked' : ''}></div></td>
                    <td><button class="btn btn-danger btn-sm"><i class="fas fa-trash"></i></button></td>
                 `;
                 tbody.appendChild(row);
            });
            table.appendChild(tbody);
            container.appendChild(table);
            separateTablesContainer.appendChild(container);
        }
        addTableEventListeners();
    }

    // --- HELPER FUNCTIONS for rendering ---
    function createDropdown(options, selectedValue) {
        const select = document.createElement('select');
        select.className = 'form-select';
        options.forEach(option => {
            const opt = document.createElement('option');
            opt.value = option;
            opt.textContent = option;
            if (option === selectedValue) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
        return select.outerHTML;
    }


    // --- EVENT LISTENERS on Tables ---
    function addTableEventListeners() {
        const allTables = document.querySelectorAll('#master-table, #separate-tables-container table');
        allTables.forEach(table => {
            table.addEventListener('change', (e) => {
                const target = e.target;
                const row = target.closest('tr');
                const id = parseFloat(row.dataset.id);
                const missionary = appState.missionaries.find(m => m.id === id);

                if (!missionary) return;

                // Logic to update appState when an input/select/checkbox in the table changes
                // This is a simplified example. You'll need to expand this for each column.
                const cellIndex = target.closest('td').cellIndex;
                if (cellIndex === 0) missionary.type = target.value;
                if (cellIndex === 5) missionary.transport = target.value;
                // Add more cases for each editable column...

                saveState();
                renderTables(); // Re-render to ensure both tables are in sync
            });

            table.addEventListener('click', (e) => {
                if (e.target.closest('.btn-danger')) {
                    const row = e.target.closest('tr');
                    const id = parseFloat(row.dataset.id);
                    if (confirm('Are you sure you want to remove this missionary from the travel plans?')) {
                        appState.missionaries = appState.missionaries.filter(m => m.id !== id);
                        saveState();
                        renderTables();
                    }
                }
            });
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