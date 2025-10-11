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
        cities: ['Tete', 'Chimoio', 'Beira', 'Quelimane', 'Nampula', 'Nhamatanda', 'Marromeu', 'Caia']
    };

    // --- DOM ELEMENTS ---
    const processFilesBtn = document.getElementById('process-files');
    const oldBoardInput = document.getElementById('old-transfer-board');
    const newBoardInput = document.getElementById('new-transfer-board');
    const mainContent = document.getElementById('main-content');
    const masterTableBody = document.querySelector('#master-table tbody');
    const separateTablesContainer = document.getElementById('separate-tables-container');
    const addMissionaryBtn = document.querySelector('[data-bs-target="#add-missionary-modal"]');
    const addMissionaryForm = document.getElementById('add-missionary-form');
    const settingsForm = document.getElementById('settings-form');
    const deleteAllBtn = document.getElementById('delete-all-missionaries');
    const downloadMasterPdfBtn = document.getElementById('download-master-pdf');
    const downloadMasterExcelBtn = document.getElementById('download-master-excel');
    const downloadSeparatePdfBtn = document.getElementById('download-separate-pdf');
    const downloadSeparateExcelBtn = document.getElementById('download-separate-excel');
    const cityExceptionsForm = document.getElementById('city-exceptions-form');
    const addCityExceptionBtn = document.getElementById('add-city-exception');
    const saveExceptionsBtn = document.getElementById('save-exceptions');


    // --- INITIALIZATION ---
    loadState();
    if (appState.missionaries.length > 0) {
        mainContent.style.display = 'block';
        document.getElementById('upload-section').style.display = 'none';
        renderTables();
    }
    updateUIText();


    // --- FILE PROCESSING ---
    processFilesBtn.addEventListener('click', () => {
        const oldFile = oldBoardInput.files[0];
        const newFile = newBoardInput.files[0];

        if (!oldFile || !newFile) {
            alert('Please upload both transfer board files.');
            return;
        }

        Papa.parse(oldFile, {
            header: true,
            complete: (oldResults) => {
                Papa.parse(newFile, {
                    header: true,
                    complete: (newResults) => {
                        processData(oldResults.data, newResults.data);
                        openCityExceptionsModal();
                    }
                });
            }
        });
    });

    function processData(oldData, newData) {
        const transferred = [];
        newData.forEach(newM => {
            const oldM = oldData.find(o => o['Missionary Name'] === newM['Missionary Name']);
            if (oldM && (oldM.Area !== newM.Area || oldM.District !== newM.District || oldM.Zone !== newM.Zone)) {
                transferred.push({
                    type: newM.Type,
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

        appState.missionaries = transferred.map(m => {
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
                id: Date.now() + Math.random()
            }
        });
        saveState();
    }

    // --- CITY AND TRANSPORT LOGIC ---
    function getCity(zone, district, area) {
        // Check for specific exceptions first
        const exception = appState.cityExceptions.find(ex =>
            (ex.type === 'area' && ex.name === area) ||
            (ex.type === 'district' && ex.name === district)
        );
        if (exception) return exception.city;

        // Zone-based rules
        if (['ZONA MUNHAVA', 'ZONA INHAMIZUA', 'ZONA MANGA'].includes(zone.toUpperCase())) {
            return 'Beira';
        }

        // Default: normalize zone name
        return zone.replace(/ZONA /i, '').charAt(0).toUpperCase() + zone.slice(1).toLowerCase();
    }


    function getDefaultTransport(from, to) {
        const majorCities = ['Tete', 'Chimoio', 'Beira'];
        if (majorCities.includes(from) && majorCities.includes(to) && from !== to) return 'Bus';
        if ((from === 'Quelimane' && to === 'Nampula') || (from === 'Nampula' && to === 'Quelimane')) return 'Bus';
        if (majorCities.includes(from) && to === 'Nampula' || from === 'Nampula' && majorCities.includes(to)) return 'Airplane';
        if (from === to && from !== 'Beira') return 'Txopela/Taxi';
        if (from === to && from === 'Beira') return 'Ride';
        return ''; // Default empty
    }


    // --- TABLE RENDERING ---
    function renderTables() {
        renderMasterTable();
        renderSeparateTables();
    }

    function renderMasterTable() {
        masterTableBody.innerHTML = '';
        const zoneColors = {};
        let colorIndex = 0;
        const colors = ['#f8d7da', '#d1ecf1', '#d4edda', '#fff3cd', '#d6d8db', '#cce5ff', '#f5c6cb'];

        appState.missionaries.forEach(m => {
            if (!zoneColors[m.originZone]) {
                zoneColors[m.originZone] = colors[colorIndex % colors.length];
                colorIndex++;
            }

            const row = document.createElement('tr');
            row.style.backgroundColor = zoneColors[m.originZone];
            row.dataset.id = m.id;

            // ... (Build and append cells for master table)

            masterTableBody.appendChild(row);
        });
        addCellEventListeners();
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
            // ... (Build and append table header for separate tables)

            const tbody = document.createElement('tbody');
            missionariesInGroup.forEach(m => {
                 const row = document.createElement('tr');
                 row.dataset.id = m.id;
                 // ... (Build and append cells for separate tables)
                 tbody.appendChild(row);
            });
            table.appendChild(tbody);
            container.appendChild(table);
            separateTablesContainer.appendChild(container);
        }
         addCellEventListeners();
    }

    // --- EVENT LISTENERS & UI ---
    function addCellEventListeners() {
        // Logic for making cells editable
    }

    deleteAllBtn.addEventListener('click', () => {
        if (confirm('Are you sure you want to delete all missionary travel data? This cannot be undone.')) {
            appState.missionaries = [];
            saveState();
            window.location.reload();
        }
    });


    // --- LOCALIZATION ---
    function updateUIText() {
        // Logic for updating UI text based on selected language
    }

    // --- MODALS ---
    function openCityExceptionsModal() {
        // Logic to populate and show the city exceptions modal
    }

    // --- DATA EXPORT ---
    downloadMasterPdfBtn.addEventListener('click', () => exportTableToPdf('master-table'));
    downloadMasterExcelBtn.addEventListener('click', () => exportTableToExcel('master-table'));
    downloadSeparatePdfBtn.addEventListener('click', () => exportAllToPdf('separate-tables-container'));
    downloadSeparateExcelBtn.addEventListener('click', () => exportAllToExcel('separate-tables-container'));

    function exportTableToPdf(tableId) { /* ... PDF export logic ... */ }
    function exportTableToExcel(tableId) { /* ... Excel export logic ... */ }
    function exportAllToPdf(containerId) { /* ... PDF export logic for multiple tables ... */ }
    function exportAllToExcel(containerId) { /* ... Excel export logic for multiple tables ... */ }


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