document.addEventListener('DOMContentLoaded', () => {
    // --- TRANSLATION DICTIONARY ---
    const translations = {
        en: {
            "missionary_transfer_logistics": "Missionary Transfer Logistics", "upload_transfer_boards": "Upload Transfer Boards",
            "old_transfer_board": "Old Transfer Board (Excel)", "new_transfer_board": "New Transfer Board (Excel)",
            "process_files": "Process Files", "master_table": "Master Table", "separate_tables": "Separate Tables",
            "add_missionary": "Add Missionary", "download_pdf": "PDF", "download_excel": "Excel", "type": "Type",
            "name": "Name", "origin_city": "Origin City", "destination_city": "Destination City",
            "destination_area": "Destination Area", "transportation": "Transportation", "date_of_travel": "Date of Travel",
            "departure_time": "Departure Time", "instructions": "Instructions", "travel_leader": "Travel Leader",
            "settings": "Settings", "language": "Language", "danger_zone": "Danger Zone",
            "delete_all_travel_data": "Delete All Travel Data", "delete_confirmation": "Are you sure you want to delete all missionary travel data? This cannot be undone.",
            "remove_confirmation": "Are you sure you want to remove this missionary from the travel plans?",
            "assign_city_exceptions": "Assign City Exceptions", "assign_city_exceptions_desc": "Assign areas or districts to a specific city if they are not in the main zone city.",
            "add_new_exception": "Add New Exception", "process_with_cities": "Process with These Cities", "city": "City",
            "area": "Area", "district": "District",
        },
        pt: {
            "missionary_transfer_logistics": "Logística de Transferência de Missionários", "upload_transfer_boards": "Carregar Planilhas de Transferência",
            "old_transfer_board": "Planilha Antiga (Excel)", "new_transfer_board": "Planilha Nova (Excel)",
            "process_files": "Processar Arquivos", "master_table": "Tabela Principal", "separate_tables": "Tabelas Separadas",
            "add_missionary": "Adicionar Missionário", "download_pdf": "PDF", "download_excel": "Excel", "type": "Tipo",
            "name": "Nome", "origin_city": "Cidade de Origem", "destination_city": "Cidade de Destino",
            "destination_area": "Área de Destino", "transportation": "Transporte", "date_of_travel": "Data da Viagem",
            "departure_time": "Hora de Partida", "instructions": "Instruções", "travel_leader": "Líder de Viagem",
            "settings": "Configurações", "language": "Idioma", "danger_zone": "Zona de Perigo",
            "delete_all_travel_data": "Apagar Todos os Dados", "delete_confirmation": "Tem certeza que quer apagar todos os dados? Esta ação não pode ser desfeita.",
            "remove_confirmation": "Tem certeza que quer remover este missionário dos planos de viagem?",
            "assign_city_exceptions": "Atribuir Exceções de Cidade", "assign_city_exceptions_desc": "Atribua áreas ou distritos a uma cidade específica se não estiverem na cidade principal da zona.",
            "add_new_exception": "Adicionar Nova Exceção", "process_with_cities": "Processar com Cidades", "city": "Cidade",
            "area": "Área", "district": "Distrito",
        }
    };

    // --- STATE MANAGEMENT ---
    let appState = {
        missionaries: [],
        settings: { language: 'en' },
        cityExceptions: [
            { city: 'Quelimane', type: 'district', name: 'Quelimane' }, { city: 'Nhamatanda', type: 'area', name: 'Nhamatanda' },
            { city: 'Marromeu', type: 'area', name: 'Marromeu' }, { city: 'Caia', type: 'area', name: 'Caia' },
        ],
        cities: ['Tete', 'Chimoio', 'Beira', 'Quelimane', 'Nampula', 'Nhamatanda', 'Marromeu', 'Caia'],
        tempData: null
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
    const languageSelect = document.getElementById('language-select');
    const saveExceptionsBtn = document.getElementById('save-exceptions');
    const exceptionsModal = new bootstrap.Modal(document.getElementById('city-exceptions-modal'));
    const exceptionsForm = document.getElementById('add-exception-form');
    const exceptionsList = document.getElementById('city-exceptions-list');

    // --- INITIALIZATION ---
    loadState();
    if (appState.missionaries.length > 0) {
        uploadSection.style.display = 'none';
        mainContent.style.display = 'block';
        renderTables();
    }
    updateUIText();

    // --- FILE PROCESSING ---
    processFilesBtn.addEventListener('click', () => {
        const oldFile = oldBoardInput.files[0];
        const newFile = newBoardInput.files[0];
        if (!oldFile || !newFile) { alert('Please upload both transfer board Excel files.'); return; }
        Promise.all([readExcelFile(oldFile), readExcelFile(newFile)])
            .then(([oldData, newData]) => {
                appState.tempData = { oldData, newData };
                populateExceptionsModal(newData);
                exceptionsModal.show();
            })
            .catch(error => console.error("Error processing files:", error));
    });
    
    saveExceptionsBtn.addEventListener('click', () => {
        if (appState.tempData) {
            processData(appState.tempData.oldData, appState.tempData.newData);
            appState.tempData = null;
            uploadSection.style.display = 'none';
            mainContent.style.display = 'block';
            renderTables();
            exceptionsModal.hide();
        }
    });

    exceptionsForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const city = e.target.elements['new-city-name'].value;
        const type = e.target.elements['exception-type'].value;
        const name = e.target.elements['exception-name'].value;
        if(city && type && name){
            appState.cityExceptions.push({ city, type, name });
            if(!appState.cities.includes(city)) appState.cities.push(city);
            saveState();
            populateExceptionsModal(appState.tempData.newData);
            e.target.reset();
        }
    });

    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => {
                try {
                    const data = event.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const sheetName = workbook.SheetNames[0];
                    resolve(XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]));
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
                    type: (newM.Position?.includes('Sister')) ? 'Sister' : 'Elder',
                    name: newM['Missionary Name'],
                    originZone: oldM.Zone, originDistrict: oldM.District, originArea: oldM.Area,
                    destinationZone: newM.Zone, destinationDistrict: newM.District, destinationArea: newM.Area,
                });
            }
        });
        appState.missionaries = transferred.map((m, index) => {
            const originCity = getCity(m.originZone, m.originDistrict, m.originArea);
            const destinationCity = getCity(m.destinationZone, m.destinationDistrict, m.destinationArea);
            return { ...m, originCity, destinationCity, transport: getDefaultTransport(originCity, destinationCity), date: '', time: '', tbd: true, instructions: '', leader: false, id: Date.now() + index };
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
        const lang = appState.settings.language;
        if (from === to && from === 'Beira') return (lang === 'pt' ? 'Boleia' : 'Ride');
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
        updateUIText();
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
                <td>${m.type}</td> <td>${m.name}</td> <td>${m.originCity}</td>
                <td>${createDropdown(appState.cities, m.destinationCity, 'destinationCity')}</td>
                <td><input type="text" class="form-control form-control-sm" data-field="destinationArea" value="${m.destinationArea}"></td>
                <td>${createDropdown(['Bus', 'Airplane', 'Chapa', 'Txopela/Taxi', 'Ride', 'Boleia'], m.transport, 'transport')}</td>
                <td>
                    <input type="date" class="form-control form-control-sm date-input" data-field="date" value="${m.date}" ${m.tbd ? 'disabled' : ''}>
                    <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? 'checked' : ''}> TBD</div>
                </td>
                <td><input type="time" class="form-control form-control-sm" data-field="time" value="${m.time}"></td>
                <td><textarea class="form-control form-control-sm" data-field="instructions">${m.instructions}</textarea></td>
                <td class="text-center align-middle"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? 'checked' : ''}></td>
                <td>
                    <button class="btn btn-sm btn-light" disabled><i class="fas fa-pencil-alt"></i></button> 
                    <button class="btn btn-danger btn-sm trash-btn"><i class="fas fa-trash"></i></button>
                </td>
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
            table.className = 'table table-bordered table-striped';
            table.innerHTML = `<thead class="table-dark"><tr><th>Type</th><th>Name</th><th>Transportation</th><th>Date</th><th>Instructions</th><th>Leader</th></tr></thead>`;
            const tbody = document.createElement('tbody');
            groups[groupName].forEach(m => {
                 tbody.innerHTML += `<tr><td>${m.type}</td><td>${m.name}</td><td>${m.transport}</td><td>${m.tbd ? 'TBD' : m.date}</td><td>${m.instructions}</td><td>${m.leader ? 'Yes' : 'No'}</td></tr>`;
            });
            table.appendChild(tbody);
            container.appendChild(table);
            separateTablesContainer.appendChild(container);
        }
    }

    // --- MODAL & EXCEPTIONS ---
    function populateExceptionsModal(data) {
        const uniqueAreas = [...new Set(data.map(item => item.Area).filter(Boolean))];
        const uniqueDistricts = [...new Set(data.map(item => item.District).filter(Boolean))];

        exceptionsList.innerHTML = `<ul class="list-group mb-4">${appState.cityExceptions.map(ex => `<li class="list-group-item">${ex.type.charAt(0).toUpperCase() + ex.type.slice(1)} '${ex.name}' is in <strong>${ex.city}</strong></li>`).join('')}</ul>`;
        
        const nameSelect = exceptionsForm.querySelector('#exception-name');
        const typeSelect = exceptionsForm.querySelector('#exception-type');

        // ROBUSTNESS CHECK: Ensure elements exist before adding listeners
        if (nameSelect && typeSelect) {
            const updateNameOptions = () => {
                const options = typeSelect.value === 'area' ? uniqueAreas : uniqueDistricts;
                nameSelect.innerHTML = options.map(opt => `<option value="${opt}">${opt}</option>`).join('');
            };
            typeSelect.onchange = updateNameOptions; // Use onchange to avoid multiple listeners
            updateNameOptions();
        }
    }

    // --- EVENT LISTENERS & UI ---
    languageSelect.addEventListener('change', (e) => {
        appState.settings.language = e.target.value;
        saveState();
        renderTables();
    });

    if (deleteAllBtn) {
        deleteAllBtn.addEventListener('click', () => {
            if (confirm(getText("delete_confirmation"))) {
                localStorage.removeItem('missionaryTransferState');
                window.location.reload();
            }
        });
    }
    
    mainContent.addEventListener('change', (e) => {
        const target = e.target;
        const row = target.closest('tr');
        if (!row) return;
        const id = parseFloat(row.dataset.id);
        const missionary = appState.missionaries.find(m => m.id === id);
        if (!missionary) return;

        if (target.matches('.tbd-checkbox')) missionary.tbd = target.checked;
        else if (target.matches('.leader-checkbox')) missionary.leader = target.checked;
        else if (target.dataset.field) missionary[target.dataset.field] = target.value;
        
        saveState();
        renderTables();
    });

    mainContent.addEventListener('click', (e) => {
         if (e.target.closest('.trash-btn')) {
            const row = e.target.closest('tr');
            const id = parseFloat(row.dataset.id);
            if (confirm(getText("remove_confirmation"))) {
                appState.missionaries = appState.missionaries.filter(m => m.id !== id);
                saveState();
                renderTables();
            }
        }
    });

    // --- DATA EXPORT ---
    document.getElementById('download-master-pdf').addEventListener('click', () => exportToPdf('master-table'));
    document.getElementById('download-master-excel').addEventListener('click', () => exportToExcel('master-table', 'Master_Transfers.xlsx'));
    document.getElementById('download-separate-pdf').addEventListener('click', () => exportToPdf('separate-tables-container'));
    document.getElementById('download-separate-excel').addEventListener('click', () => exportToExcel('separate-tables-container', 'Separate_Transfers.xlsx'));

    function exportToPdf(elementId) {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'landscape' });
        const element = document.getElementById(elementId);
        const tables = element.getElementsByTagName('table');
        
        for (let i = 0; i < tables.length; i++) {
            if (i > 0) doc.addPage();
            const title = tables[i].previousElementSibling?.textContent || "Transfers";
            doc.text(title, 14, 15);
            doc.autoTable({
                html: tables[i],
                startY: 20,
                theme: 'grid',
                styles: { fontSize: 8 },
                headStyles: { fillColor: [41, 128, 185] }
            });
        }
        doc.save('transfers.pdf');
    }

    function exportToExcel(elementId, fileName) {
        const element = document.getElementById(elementId);
        const tables = element.getElementsByTagName('table');
        const wb = XLSX.utils.book_new();

        for (let i = 0; i < tables.length; i++) {
            let title = `Sheet${i+1}`;
            if(elementId === 'separate-tables-container' && tables[i].previousElementSibling) {
                title = tables[i].previousElementSibling.textContent;
            } else if (elementId === 'master-table') {
                title = 'Master List';
            }
            const ws = XLSX.utils.table_to_sheet(tables[i]);
            XLSX.utils.book_append_sheet(wb, ws, title.replace(/[^a-zA-Z0-9]/g, '').substring(0, 31));
        }
        XLSX.writeFile(wb, fileName);
    }
    
    // --- TRANSLATION & HELPERS ---
    function updateUIText() {
        const lang = appState.settings.language;
        languageSelect.value = lang;
        document.querySelectorAll("[data-translate-key]").forEach(el => {
            const key = el.getAttribute("data-translate-key");
            if (translations[lang] && translations[lang][key]) el.innerText = translations[lang][key];
        });
        document.querySelectorAll("#master-table thead th").forEach((th, i) => {
            const keys = ["type", "name", "origin_city", "destination_city", "destination_area", "transportation", "date_of_travel", "departure_time", "instructions", "travel_leader"];
            if (keys[i]) th.innerText = getText(keys[i]);
        });
    }

    function getText(key) { return translations[appState.settings.language]?.[key] || key; }
    function createDropdown(options, selected, field) { return `<select class="form-select form-select-sm" data-field="${field}">${options.map(o => `<option value="${o}" ${o === selected ? 'selected' : ''}>${o}</option>`).join('')}</select>`; }
    
    // --- LOCALSTORAGE ---
    function saveState() { localStorage.setItem('missionaryTransferState', JSON.stringify(appState)); }
    function loadState() {
        const saved = localStorage.getItem('missionaryTransferState');
        if (saved) appState = {...appState, ...JSON.parse(saved)};
    }
});