// -- Export and Update Functions
document.getElementById("btn-export-pdf").addEventListener("click", () => {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "l", unit: "mm", format: "a4" });
  const lang = state.lang || "pt";
  const T = translations[lang];
  M.toast({ html: T.toast_pdf_generating });

  const view = getActiveView();
  const visibleCols = state.columnSettings[view];
  const allHeaders = {
    type: T.table_header_type,
    lastName: T.table_header_last_name,
    firstName: T.table_header_first_name,
    originCity: T.origin_city,
    destCity: T.dest_city,
    originArea: T.table_header_origin_area,
    destArea: T.table_header_dest_area,
    companion: T.table_header_companion,
    transport: T.table_header_transport,
    date: T.table_header_date,
    time: T.table_header_time,
    instructions: T.table_header_instructions,
    new: T.table_header_new,
    leader: T.table_header_leader,
  };
  const head = [visibleCols.map((key) => allHeaders[key])];

  const getRowData = (m) =>
    visibleCols.map((key) => {
      switch (key) {
        case "transport":
          return m.transport
            ? T[`transport_${m.transport.toLowerCase().replace("/", "_")}`] ||
                m.transport
            : "";
        case "new":
          return m.isNew ? "✅" : "";
        case "leader":
          return m.leader ? "✅" : "";
        default:
          return m[key] || "";
      }
    });

  let finalY = 15;
  const autoTableOptions = {
    head,
    startY: finalY,
    theme: "striped",
    headStyles: { fillColor: [0, 150, 136] },
  };

  if (view === "city") {
    Object.keys(state.groups)
      .sort()
      .forEach((groupKey) => {
        if (finalY > 180) {
          doc.addPage();
          finalY = 15;
        }
        doc.setFontSize(14);
        doc.text(groupKey, 14, finalY);
        finalY += 7;
        autoTableOptions.body = state.groups[groupKey]
          .sort((a, b) => a.isNew - b.isNew)
          .map(getRowData);
        autoTableOptions.startY = finalY;
        doc.autoTable(autoTableOptions);
        finalY = doc.previousAutoTable.finalY + 10;
      });
  } else {
    let groupsToPrint = {};
    if (view === "transport") {
      const byTransport = {};
      Object.values(state.groups)
        .flat()
        .forEach((m) => {
          const t = m.transport || "Unassigned";
          if (!byTransport[t]) byTransport[t] = [];
          byTransport[t].push(m);
        });
      groupsToPrint = byTransport;
    } else {
      // Master List
      const all = Object.values(state.groups)
        .flat()
        .sort(
          (a, b) =>
            a.isNew - b.isNew ||
            a.lastName.localeCompare(b.lastName) ||
            a.firstName.localeCompare(b.firstName),
        );
      groupsToPrint = { [T.tab_master]: all };
    }
    Object.keys(groupsToPrint)
      .sort()
      .forEach((groupKey) => {
        if (finalY > 180 && view !== "master") {
          doc.addPage();
          finalY = 15;
        }
        const groupTitle =
          T[`transport_${groupKey.toLowerCase().replace("/", "_")}`] ||
          groupKey;
        doc.setFontSize(14);
        doc.text(groupTitle, 14, finalY);
        finalY += 7;
        autoTableOptions.body = groupsToPrint[groupKey]
          .sort((a, b) => a.isNew - b.isNew)
          .map(getRowData);
        autoTableOptions.startY = finalY;
        doc.autoTable(autoTableOptions);
        finalY = doc.previousAutoTable.finalY + 10;
      });
  }

  doc.save("transfer_logistics.pdf");
  M.toast({ html: T.toast_pdf_generated });
});

document.getElementById("btn-export-csv").addEventListener("click", () => {
  const lang = state.lang || "pt";
  const T = translations[lang];
  const delimiter = ";";
  const headers = [
    T.table_header_last_name,
    T.table_header_first_name,
    T.table_header_type,
    T.origin_city,
    T.dest_city,
    T.table_header_origin_area,
    T.table_header_dest_area,
    T.table_header_companion,
    T.table_header_transport,
    T.table_header_date,
    T.table_header_time,
    T.table_header_instructions,
    T.table_header_new,
    T.table_header_leader,
  ];
  let csvContent =
    "data:text/csv;charset=utf-8," + headers.join(delimiter) + "\n";
  Object.values(state.groups || {})
    .flat()
    .forEach((m) => {
      const sanitize = (str) => `"${(str || "").replace(/"/g, '""')}"`;
      const row = [
        sanitize(m.lastName),
        sanitize(m.firstName),
        m.type,
        m.originCity,
        m.destinationCity,
        sanitize(m.originArea),
        sanitize(m.destinationArea),
        sanitize(m.companion),
        T[`transport_${(m.transport || "").toLowerCase().replace("/", "_")}`] ||
          m.transport,
        m.date,
        m.time,
        sanitize(m.instructions),
        m.isNew ? "Yes" : "No",
        m.leader ? "Yes" : "No",
      ];
      csvContent += row.join(delimiter) + "\n";
    });
  const link = document.createElement("a");
  link.setAttribute("href", encodeURI(csvContent));
  link.setAttribute("download", "travel_plans.csv");
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
});
