// -- Export and Update Functions
document.getElementById("btn-export-pdf").addEventListener("click", () => {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "l", unit: "mm", format: "a4" });
  const lang = state.lang || "pt";
  const T = translations[lang];
  M.toast({ html: T.toast_pdf_generating });

  const view = getActiveView();
  const visibleCols = state.columnSettings[view];
  // support multiple key naming conventions: destCity/destinationCity, destArea/destinationArea
  const allHeaders = {
    type: T.table_header_type,
    lastName: T.table_header_last_name,
    firstName: T.table_header_first_name,
    originCity: T.origin_city,
    destCity: T.dest_city,
    destinationCity: T.dest_city,
    originArea: T.table_header_origin_area,
    destArea: T.table_header_dest_area,
    destinationArea: T.table_header_dest_area,
    companion: T.table_header_companion,
    transport: T.table_header_transport,
    date: T.table_header_date,
    time: T.table_header_time,
    instructions: T.table_header_instructions,
    new: T.table_header_new,
    leader: T.table_header_leader,
  };

  const head = [visibleCols.map((key) => allHeaders[key] || key)];

  // determine which visible columns are checkboxes so we can center them and give fixed width
  const checkboxColIndices = visibleCols.reduce((acc, k, i) => {
    if (k === 'new' || k === 'leader') acc.push(i);
    return acc;
  }, []);

  const resolveFieldValue = (m, key) => {
    // normalize known aliases
    if (key === "destCity" || key === "destinationCity")
      return m.destinationCity || m.destCity || "";
    if (key === "destArea" || key === "destinationArea")
      return m.destinationArea || m.destArea || "";
    if (key === "originCity") return m.originCity || m.origin || "";
    if (key === "originArea") return m.originArea || m.origin_area || "";
    if (key === "new") return m.isNew;
    // leader may be boolean or truthy flag
    if (key === "leader") return m.leader;
    // default fallback
    return m[key] !== undefined ? m[key] : "";
  };

  const renderCell = (val, key) => {
    // transport has localized label
    if (key === "transport") {
      if (!val) return "";
      return (
        T[`transport_${String(val).toLowerCase().replace("/", "_")}`] || val
      );
    }
    // booleans / checkbox-like fields
    if (typeof val === "boolean") return val ? "✅" : "";
    if (val === "true" || val === "True" || val === 1) return "✅";
    if (val === "false" || val === "False" || val === 0) return "";
    return String(val || "");
  };

  // we'll build rows and a parallel checkbox matrix so we can draw pretty boxes with jsPDF
  const buildRowsAndCheckboxMatrix = (list) => {
    const rows = [];
    const checkboxMatrix = [];
    list.forEach((m) => {
      const row = [];
      const cbRow = [];
      visibleCols.forEach((key) => {
        const raw = resolveFieldValue(m, key);
        const isCheckbox = key === 'new' || key === 'leader';
        if (isCheckbox) {
          // leave text empty; will draw box in didDrawCell
          row.push('');
          cbRow.push(Boolean(raw));
        } else {
          row.push(renderCell(raw, key));
          cbRow.push(null);
        }
      });
      rows.push(row);
      checkboxMatrix.push(cbRow);
    });
    return { rows, checkboxMatrix };
  };

  let finalY = 15;
  // currentCheckboxMatrix will be set before each autoTable call so didDrawCell can read it
  let currentCheckboxMatrix = [];
  const autoTableOptions = {
    head,
    startY: finalY,
    theme: "striped",
    styles: { fontSize: 10 },
  headStyles: { fillColor: [0, 150, 136], textColor: 255, fontStyle: 'bold', fontSize: 10 },
    columnStyles: (function(){
      const cs = {};
      checkboxColIndices.forEach(i => { cs[i] = { halign: 'center', cellWidth: 14 }; });
      return cs;
    })(),
    didDrawCell: function (data) {
      // draw checkbox box + check for boolean checkboxMatrix entries
      if (data.section !== 'body') return;
      const r = data.row.index;
      const c = data.column.index;
      const val = currentCheckboxMatrix[r] && currentCheckboxMatrix[r][c];
      if (val === null || val === undefined) return;

      // compute a small square centered in the cell
  // make checkbox larger (proportional to cell size), but cap to a reasonable size
  let boxSize = Math.min(data.cell.width, data.cell.height) * 0.6; // relative
  boxSize = Math.min(boxSize, 18); // max 18mm
      const x = data.cell.x + (data.cell.width - boxSize) / 2;
      const y = data.cell.y + (data.cell.height - boxSize) / 2;

      // draw square
      data.doc.setDrawColor(0);
      data.doc.setLineWidth(0.5);
      data.doc.rect(x, y, boxSize, boxSize);

      if (val) {
        // draw a thicker checkmark scaled to the boxSize
        const pad = boxSize * 0.18;
        const x1 = x + pad;
        const y1 = y + boxSize * 0.55;
        const x2 = x + boxSize * 0.44;
        const y2 = y + boxSize - pad;
        const x3 = x + boxSize - pad;
        const y3 = y + pad;
        data.doc.setLineWidth(Math.max(0.8, boxSize * 0.09));
        data.doc.line(x1, y1, x2, y2);
        data.doc.line(x2, y2, x3, y3);
      }
    }
  };

  if (view === "city") {
    Object.keys(state.groups)
      .sort()
      .forEach((groupKey) => {
        if (finalY > 180) {
          doc.addPage();
          finalY = 15;
        }
  const pageWidth = doc.internal.pageSize.getWidth();
  // render group title using same font and size as table header for a matching look
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.text(groupKey, pageWidth / 2, finalY, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  finalY += 8;
        const list = state.groups[groupKey].sort((a, b) => a.isNew - b.isNew);
        const built = buildRowsAndCheckboxMatrix(list);
        autoTableOptions.body = built.rows;
        currentCheckboxMatrix = built.checkboxMatrix;
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
  const pageWidth = doc.internal.pageSize.getWidth();
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.text(groupTitle, pageWidth / 2, finalY, { align: 'center' });
  doc.setFont('helvetica', 'normal');
  finalY += 8;
        const list = groupsToPrint[groupKey].sort((a, b) => a.isNew - b.isNew);
        const built = buildRowsAndCheckboxMatrix(list);
        autoTableOptions.body = built.rows;
        currentCheckboxMatrix = built.checkboxMatrix;
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
