document.addEventListener("DOMContentLoaded", () => {
  // --- TRANSLATION DICTIONARY ---
  const translations = {
    en: {
      actions: "Actions",
      add_missionary: "Add Missionary",
      add_new_exception: "Add New Exception",
      area: "Area",
      assign_city_exceptions: "Assign City Exceptions",
      assign_city_exceptions_desc: "Assign areas or districts to a city.",
      city: "City",
      danger_zone: "Danger Zone",
      date_of_travel: "Date",
      delete_all_travel_data: "Delete All Travel Data",
      delete_confirmation:
        "Are you sure you want to delete all missionary travel data? This cannot be undone.",
      departure_time: "Time",
      destination_area: "Area",
      destination_city: "Destination",
      district: "District",
      download_excel: "Excel",
      download_pdf: "PDF",
      instructions: "Instructions",
      language: "Language",
      master_table: "Master Table",
      missionary_transfer_logistics: "Missionary Transfer Logistics",
      name: "Name",
      new_transfer_board: "New Transfer Board (Excel)",
      old_transfer_board: "Old Transfer Board (Excel)",
      origin_city: "Origin",
      process_files: "Process Files",
      process_with_cities: "Continue to Transportation",
      remove_confirmation: "Are you sure you want to remove this missionary?",
      separate_tables: "Separate Tables",
      settings: "Settings",
      transportation: "Transport",
      travel_leader: "Leader",
      type: "Type",
      upload_transfer_boards: "Upload Transfer Boards",
      Bus: "Bus",
      Airplane: "Airplane",
      Chapa: "Chapa",
      "Txopela/Taxi": "Txopela/Taxi",
      Ride: "Ride",
      Boleia: "Boleia",
    },
    pt: {
      actions: "Ações",
      add_missionary: "Adicionar Missionário",
      add_new_exception: "Adicionar Nova Exceção",
      area: "Área",
      assign_city_exceptions: "Atribuir Exceções de Cidade",
      assign_city_exceptions_desc: "Atribua áreas ou distritos a uma cidade.",
      city: "Cidade",
      danger_zone: "Zona de Perigo",
      date_of_travel: "Data",
      delete_all_travel_data: "Apagar Todos os Dados",
      delete_confirmation:
        "Tem certeza que quer apagar todos os dados? Esta ação não pode ser desfeita.",
      departure_time: "Hora",
      destination_area: "Área",
      destination_city: "Destino",
      district: "Distrito",
      download_excel: "Excel",
      download_pdf: "PDF",
      instructions: "Instruções",
      language: "Idioma",
      master_table: "Tabela Principal",
      missionary_transfer_logistics:
        "Logística de Transferência de Missionários",
      name: "Nome",
      new_transfer_board: "Planilha Nova (Excel)",
      old_transfer_board: "Planilha Antiga (Excel)",
      origin_city: "Origem",
      process_files: "Processar Arquivos",
      process_with_cities: "Continuar para Transporte",
      remove_confirmation: "Tem certeza que quer remover este missionário?",
      separate_tables: "Tabelas Separadas",
      settings: "Configurações",
      transportation: "Transporte",
      travel_leader: "Líder",
      type: "Tipo",
      upload_transfer_boards: "Carregar Planilhas",
      Bus: "Autocarro",
      Airplane: "Avião",
      Chapa: "Chapa",
      "Txopela/Taxi": "Txopela/Táxi",
      Ride: "Boleia",
      Boleia: "Boleia",
    },
  };

  // --- STATE MANAGEMENT ---
  let appState = {
    missionaries: [],
    settings: { language: "en" },
    cityExceptions: [
      { city: "Quelimane", type: "district", name: "Quelimane" },
      { city: "Nhamatanda", type: "area", name: "Nhamatanda" },
      { city: "Marromeu", type: "area", name: "Marromeu" },
      { city: "Caia", type: "area", name: "Caia" },
    ],
    cities: [
      "Tete",
      "Chimoio",
      "Beira",
      "Quelimane",
      "Nampula",
      "Nhamatanda",
      "Marromeu",
      "Caia",
    ],
    tempData: null,
    transportDefaults: {},
  };

  // --- DOM ELEMENTS ---
  const processFilesBtn = document.getElementById("process-files");
  const oldBoardInput = document.getElementById("old-transfer-board");
  const newBoardInput = document.getElementById("new-transfer-board");
  const mainContent = document.getElementById("main-content");
  const uploadSection = document.getElementById("upload-section");
  const masterTableBody = document.querySelector("#master-table tbody");
  const separateTablesContainer = document.getElementById(
    "separate-tables-container",
  );
  const deleteAllBtn = document.getElementById("delete-all-missionaries");
  const languageSelect = document.getElementById("language-select");
  const exceptionsModal = new bootstrap.Modal(
    document.getElementById("city-exceptions-modal"),
  );
  const transportModal = new bootstrap.Modal(
    document.getElementById("transportation-defaults-modal"),
  );
  const exceptionsForm = document.getElementById("add-exception-form");
  const exceptionsList = document.getElementById("city-exceptions-list");
  const continueToTransportBtn = document.getElementById(
    "continue-to-transport",
  );
  const finishProcessingBtn = document.getElementById("finish-processing");

  // --- INITIALIZATION ---
  loadState();
  if (appState.missionaries.length > 0) {
    uploadSection.style.display = "none";
    mainContent.style.display = "block";
    renderTables();
  }
  updateUIText();

  // --- FILE PROCESSING & MODAL FLOW ---
  if (processFilesBtn) {
    processFilesBtn.addEventListener("click", () => {
      const oldFile = oldBoardInput.files[0];
      const newFile = newBoardInput.files[0];
      if (!oldFile || !newFile) {
        alert("Please upload both Excel files.");
        return;
      }
      Promise.all([readExcelFile(oldFile), readExcelFile(newFile)])
        .then(([oldData, newData]) => {
          appState.tempData = { oldData, newData };
          populateExceptionsModal(newData);
          exceptionsModal.show();
        })
        .catch((error) => console.error("Error reading files:", error));
    });
  }

  if (continueToTransportBtn) {
    continueToTransportBtn.addEventListener("click", () => {
      populateTransportModal();
      exceptionsModal.hide();
      transportModal.show();
    });
  }

  if (finishProcessingBtn) {
    finishProcessingBtn.addEventListener("click", () => {
      if (appState.tempData) {
        processData(appState.tempData.oldData, appState.tempData.newData);
        appState.tempData = null;
        uploadSection.style.display = "none";
        mainContent.style.display = "block";
        renderTables();
        transportModal.hide();
      }
    });
  }

  if (exceptionsForm) {
    exceptionsForm.addEventListener("submit", (e) => {
      e.preventDefault();
      const city = e.target.elements["new-city-name"].value;
      const type = e.target.elements["exception-type"].value;
      const name = e.target.elements["exception-name"].value;
      if (city && type && name) {
        appState.cityExceptions.push({ city, type, name });
        if (!appState.cities.includes(city)) appState.cities.push(city);
        saveState();
        populateExceptionsModal(appState.tempData.newData);
        e.target.reset();
      }
    });
  }

  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const workbook = XLSX.read(event.target.result, { type: "binary" });
          resolve(
            XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]),
          );
        } catch (e) {
          reject(e);
        }
      };
      reader.onerror = reject;
      reader.readAsBinaryString(file);
    });
  }

  function processData(oldData, newData) {
    const transferred = [];
    const oldDataMap = new Map(oldData.map((m) => [m["Missionary Name"], m]));
    newData.forEach((newM) => {
      const oldM = oldDataMap.get(newM["Missionary Name"]);
      if (oldM && (oldM.Area !== newM.Area || oldM.Zone !== newM.Zone)) {
        transferred.push({
          type: newM.Position?.includes("Sister") ? "Sister" : "Elder",
          name: newM["Missionary Name"],
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
      const destinationCity = getCity(
        m.destinationZone,
        m.destinationDistrict,
        m.destinationArea,
      );
      const routeKey = `${originCity}->${destinationCity}`;
      const transport = appState.transportDefaults[routeKey] || "Bus";
      return {
        ...m,
        originCity,
        destinationCity,
        transport,
        date: "",
        time: "",
        tbd: false,
        instructions: "",
        leader: false,
        id: Date.now() + index,
      };
    });
    saveState();
  }

  // --- CITY AND TRANSPORT LOGIC ---
  function getCity(zone = "", district = "", area = "") {
    const exception = appState.cityExceptions.find(
      (ex) =>
        (ex.type === "area" && ex.name === area) ||
        (ex.type === "district" && ex.name === district),
    );
    if (exception) return exception.city;
    const upperZone = zone.toUpperCase().trim();
    if (["ZONA MUNHAVA", "ZONA INHAMIZUA", "ZONA MANGA"].includes(upperZone))
      return "Beira";
    let cityName = zone.replace(/ZONA /i, "").trim();
    return cityName.charAt(0).toUpperCase() + cityName.slice(1).toLowerCase();
  }

  function getDefaultTransport(from, to) {
    const majorCities = ["Tete", "Chimoio", "Beira"];
    if (from === to && from === "Beira") return "Ride";
    if (from === to) return "Txopela/Taxi";
    if (majorCities.includes(from) && majorCities.includes(to)) return "Bus";
    if (
      (from === "Quelimane" && to === "Nampula") ||
      (from === "Nampula" && to === "Quelimane")
    )
      return "Bus";
    if (
      (majorCities.includes(from) && to === "Nampula") ||
      (from === "Nampula" && majorCities.includes(to))
    )
      return "Airplane";
    return "Bus";
  }

  // --- TABLE RENDERING ---
  function renderTables() {
    renderMasterTable();
    renderSeparateTables();
    updateUIText();
  }

  function renderMasterTable() {
    masterTableBody.innerHTML = "";
    const zones = [
      ...new Set(
        appState.missionaries.map((m) => m.originZone.trim().toUpperCase()),
      ),
    ];
    const colors = [
      "#f8d7da",
      "#d1ecf1",
      "#d4edda",
      "#fff3cd",
      "#d6d8db",
      "#cce5ff",
      "#f5c6cb",
    ];
    const zoneColorMap = zones.reduce(
      (acc, zone, i) => ({ ...acc, [zone]: colors[i % colors.length] }),
      {},
    );

    appState.missionaries.forEach((m) => {
      const row = document.createElement("tr");
      const zoneKey = m.originZone.trim().toUpperCase();
      row.style.backgroundColor = zoneColorMap[zoneKey];
      row.dataset.id = m.id;
      row.innerHTML = `
                <td>${createDropdown(["Elder", "Sister"], m.type, "type")}</td>
                <td><input type="text" class="form-control form-control-sm" data-field="name" value="${m.name}"></td>
                <td><input type="text" class="form-control form-control-sm" data-field="originCity" value="${m.originCity}"></td>
                <td>${createDropdown(appState.cities, m.destinationCity, "destinationCity")}</td>
                <td><input type="text" class="form-control form-control-sm" data-field="destinationArea" value="${m.destinationArea}"></td>
                <td>${createDropdown(["Bus", "Airplane", "Chapa", "Txopela/Taxi", "Ride", "Boleia"], m.transport, "transport")}</td>
                <td><input type="date" class="form-control form-control-sm date-input" data-field="date" value="${m.date}" ${m.tbd ? "disabled" : ""}> <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? "checked" : ""}> TBD</div></td>
                <td><input type="time" class="form-control form-control-sm" data-field="time" value="${m.time}"></td>
                <td><textarea class="form-control form-control-sm" data-field="instructions">${m.instructions}</textarea></td>
                <td class="text-center align-middle"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? "checked" : ""}></td>
                <td><button class="btn btn-danger btn-sm trash-btn"><i class="fas fa-trash"></i></button></td>
            `;
      masterTableBody.appendChild(row);
    });
  }

  function renderSeparateTables() {
    separateTablesContainer.innerHTML = "";
    const groups = appState.missionaries.reduce((acc, m) => {
      const key = `${m.originCity} to ${m.destinationCity}`;
      if (!acc[key]) acc[key] = [];
      acc[key].push(m);
      return acc;
    }, {});
    for (const groupName in groups) {
      const container = document.createElement("div");
      container.innerHTML = `<h3 class="mt-4">${groupName}</h3>`;
      const table = document.createElement("table");
      table.className = "table table-bordered table-hover";
      table.innerHTML = `<thead class="table-dark"><tr><th>Type</th><th>Name</th><th>Origin</th><th>Destination</th><th>Area</th><th>Transport</th><th>Date</th><th>Time</th><th>Instructions</th><th>Leader</th><th>Actions</th></tr></thead>`;
      const tbody = document.createElement("tbody");
      groups[groupName].forEach((m) => {
        const row = document.createElement("tr");
        row.dataset.id = m.id;
        row.style.backgroundColor = "white";
        row.innerHTML = `
                    <td>${createDropdown(["Elder", "Sister"], m.type, "type")}</td>
                    <td><input type="text" class="form-control form-control-sm" data-field="name" value="${m.name}"></td>
                    <td><input type="text" class="form-control form-control-sm" data-field="originCity" value="${m.originCity}"></td>
                    <td>${createDropdown(appState.cities, m.destinationCity, "destinationCity")}</td>
                    <td><input type="text" class="form-control form-control-sm" data-field="destinationArea" value="${m.destinationArea}"></td>
                    <td>${createDropdown(["Bus", "Airplane", "Chapa", "Txopela/Taxi", "Ride", "Boleia"], m.transport, "transport")}</td>
                    <td><input type="date" class="form-control form-control-sm date-input" data-field="date" value="${m.date}" ${m.tbd ? "disabled" : ""}> <div class="form-check form-switch mt-1"><input class="form-check-input tbd-checkbox" type="checkbox" ${m.tbd ? "checked" : ""}> TBD</div></td>
                    <td><input type="time" class="form-control form-control-sm" data-field="time" value="${m.time}"></td>
                    <td><textarea class="form-control form-control-sm" data-field="instructions">${m.instructions}</textarea></td>
                    <td class="text-center align-middle"><input class="form-check-input leader-checkbox" type="checkbox" ${m.leader ? "checked" : ""}></td>
                    <td><button class="btn btn-danger btn-sm trash-btn"><i class="fas fa-trash"></i></button></td>
                `;
        tbody.appendChild(row);
      });
      table.appendChild(tbody);
      container.appendChild(table);
    }
  }

  // --- MODAL POPULATION ---
  function populateExceptionsModal(data) {
    const uniqueAreas = [
      ...new Set(data.map((item) => item.Area).filter(Boolean)),
    ];
    const uniqueDistricts = [
      ...new Set(data.map((item) => item.District).filter(Boolean)),
    ];
    exceptionsList.innerHTML = `<ul class="list-group mb-4">${appState.cityExceptions.map((ex) => `<li class="list-group-item">${getText(ex.type)} '${ex.name}' is in <strong>${ex.city}</strong></li>`).join("")}</ul>`;
    const nameSelect = exceptionsForm.querySelector("#exception-name");
    const typeSelect = exceptionsForm.querySelector("#exception-type");
    if (nameSelect && typeSelect) {
      const updateNameOptions = () => {
        const options =
          typeSelect.value === "area" ? uniqueAreas : uniqueDistricts;
        nameSelect.innerHTML = options
          .map((opt) => `<option value="${opt}">${opt}</option>`)
          .join("");
      };
      typeSelect.onchange = updateNameOptions;
      updateNameOptions();
    }
  }

  function populateTransportModal() {
    const routes = new Set();
    const oldDataMap = new Map(
      appState.tempData.oldData.map((m) => [m["Missionary Name"], m]),
    );
    appState.tempData.newData.forEach((newM) => {
      const oldM = oldDataMap.get(newM["Missionary Name"]);
      if (oldM && (oldM.Area !== newM.Area || oldM.Zone !== newM.Zone)) {
        const from = getCity(oldM.Zone, oldM.District, oldM.Area);
        const to = getCity(newM.Zone, newM.District, newM.Area);
        routes.add(`${from}->${to}`);
      }
    });

    let tableHTML =
      '<table class="table"><thead><tr><th>From</th><th>To</th><th>Default Transport</th></tr></thead><tbody>';
    appState.transportDefaults = {};
    routes.forEach((route) => {
      const [from, to] = route.split("->");
      const defaultTransport = getDefaultTransport(from, to);
      appState.transportDefaults[route] = defaultTransport;
      tableHTML += `<tr><td>${from}</td><td>${to}</td><td>${createDropdown(["Bus", "Airplane", "Chapa", "Txopela/Taxi", "Ride", "Boleia"], defaultTransport, route, "transport-default")}</td></tr>`;
    });
    tableHTML += "</tbody></table>";
    document.getElementById("transportation-defaults-list").innerHTML =
      tableHTML;
  }

  document
    .getElementById("transportation-defaults-list")
    .addEventListener("change", (e) => {
      if (e.target.classList.contains("transport-default")) {
        appState.transportDefaults[e.target.dataset.field] = e.target.value;
      }
    });

  // --- EVENT LISTENERS & UI ---
  if (languageSelect)
    languageSelect.addEventListener("change", (e) => {
      appState.settings.language = e.target.value;
      saveState();
      renderTables();
    });
  if (deleteAllBtn)
    deleteAllBtn.addEventListener("click", () => {
      if (confirm(getText("delete_confirmation"))) {
        localStorage.removeItem("missionaryTransferState");
        window.location.reload();
      }
    });

  if (mainContent) {
    mainContent.addEventListener("change", (e) => {
      const target = e.target;
      const row = target.closest("tr");
      if (!row) return;
      const id = parseFloat(row.dataset.id);
      const missionary = appState.missionaries.find((m) => m.id === id);
      if (!missionary) return;
      if (target.matches(".tbd-checkbox")) missionary.tbd = target.checked;
      else if (target.matches(".leader-checkbox"))
        missionary.leader = target.checked;
      else if (target.dataset.field)
        missionary[target.dataset.field] = target.value;
      saveState();
      renderTables();
    });
    mainContent.addEventListener("click", (e) => {
      if (e.target.closest(".trash-btn")) {
        const row = e.target.closest("tr");
        const id = parseFloat(row.dataset.id);
        if (confirm(getText("remove_confirmation"))) {
          appState.missionaries = appState.missionaries.filter(
            (m) => m.id !== id,
          );
          saveState();
          renderTables();
        }
      }
    });
  }

  // --- DATA EXPORT ---
  document
    .getElementById("print-btn")
    .addEventListener("click", () => window.print());
  document
    .getElementById("download-excel-btn")
    .addEventListener("click", () =>
      exportToExcel("master-table", "Master_Transfers.xlsx"),
    );

  function exportToExcel(elementId, fileName) {
    const element = document.getElementById(elementId);
    const tables = element.getElementsByTagName("table");
    const wb = XLSX.utils.book_new();
    for (let i = 0; i < tables.length; i++) {
      let title =
        elementId === "master-table"
          ? "Master List"
          : tables[i].previousElementSibling?.textContent || `Sheet${i + 1}`;
      const ws = XLSX.utils.table_to_sheet(tables[i]);
      XLSX.utils.book_append_sheet(
        wb,
        ws,
        title.replace(/[^a-zA-Z0-9]/g, "").substring(0, 31),
      );
    }
    XLSX.writeFile(wb, fileName);
  }

  // --- TRANSLATION & HELPERS ---
  function updateUIText() {
    const lang = appState.settings.language;
    languageSelect.value = lang;
    document.querySelectorAll("[data-translate-key]").forEach((el) => {
      if (translations[lang]?.[el.dataset.translateKey])
        el.innerText = translations[lang][el.dataset.translateKey];
    });
    document
      .querySelectorAll(
        "#master-table thead th, #separate-tables-container thead th",
      )
      .forEach((th, i) => {
        const keys = [
          "type",
          "name",
          "origin_city",
          "destination_city",
          "destination_area",
          "transportation",
          "date_of_travel",
          "departure_time",
          "instructions",
          "travel_leader",
          "actions",
        ];
        const key = th.parentElement.children[i % keys.length].innerText
          .toLowerCase()
          .replace(" ", "_"); // Heuristic
        if (keys.includes(key)) th.innerText = getText(key);
      });
  }

  function getText(key) {
    return translations[appState.settings.language]?.[key] || key;
  }
  function createDropdown(options, selected, field, className = "") {
    const lang = appState.settings.language;
    return `<select class="form-select form-select-sm ${className}" data-field="${field}">${options.map((o) => `<option value="${o}" ${o === selected ? "selected" : ""}>${getText(o)}</option>`).join("")}</select>`;
  }

  // --- LOCALSTORAGE ---
  function saveState() {
    localStorage.setItem("missionaryTransferState", JSON.stringify(appState));
  }
  function loadState() {
    const saved = localStorage.getItem("missionaryTransferState");
    if (saved) appState = { ...appState, ...JSON.parse(saved) };
  }
});
