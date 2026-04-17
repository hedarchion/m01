const { PDFDocument, StandardFonts, rgb } = window.PDFLib;

const PAGE = { width: 595, height: 842 };
const PREVIEW_SCALE = 2;
const LOGO_PATH = "./Pasted%20image.png";
const TYPE = {
  body: 9.2,
  bodySmall: 8.4,
  bodyTight: 8.8,
  label: 9.2,
  section: 10.2,
  title: 13,
  subtitle: 12,
  lineHeight: 1.16,
  lineHeightTight: 1.1,
  insetX: 4,
  insetY: 3,
};

const FIELD_DEFS = [
  {
    id: "championship",
    label: "Kejohanan",
    type: "text",
    fallback: "KEJOHANAN CATUR PERINGKAT NEGERI PERAK",
    aliases: ["kejohanan", "championship", "event", "tournament"],
  },
  {
    id: "district",
    label: "MSS Daerah",
    type: "text",
    fallback: "MUALLIM",
    aliases: ["mss daerah", "daerah", "district", "ppd"],
  },
  {
    id: "category",
    label: "Kategori",
    type: "text",
    fallback: "SM P15",
    aliases: ["kategori", "category", "division", "class"],
  },
  {
    id: "fullName",
    label: "Nama Penuh",
    type: "text",
    fallback: "",
    aliases: ["nama penuh", "full name", "participant name", "student name", "name"],
  },
  {
    id: "gender",
    label: "Jantina",
    type: "text",
    fallback: "",
    aliases: ["jantina", "gender", "sex"],
  },
  {
    id: "icNumber",
    label: "No Kad Pengenalan",
    type: "text",
    fallback: "",
    aliases: ["no kad pengenalan", "ic number", "mykad", "nric", "ic"],
  },
  {
    id: "birthDate",
    label: "Tarikh Lahir",
    type: "text",
    fallback: "",
    aliases: ["tarikh lahir", "birth date", "dob", "date of birth"],
  },
  {
    id: "homeAddress",
    label: "Alamat Rumah",
    type: "text",
    fallback: "",
    aliases: ["alamat rumah", "address", "home address", "alamat"],
  },
  {
    id: "homePhone",
    label: "Telefon Rumah",
    type: "text",
    fallback: "",
    aliases: ["telefon rumah", "phone", "telephone", "contact number", "home phone"],
  },
  {
    id: "registrationDate",
    label: "Tarikh Daftar",
    type: "text",
    fallback: "",
    aliases: ["tarikh daftar", "registration date", "register date", "date registered"],
  },
  {
    id: "studentId",
    label: "No ID Murid",
    type: "text",
    fallback: "",
    aliases: ["no id murid", "student id", "murid id", "school id", "id murid"],
  },
  {
    id: "passportPhoto",
    label: "Passport Photo Path",
    type: "image",
    fallback: "",
    aliases: ["passport photo", "photo", "passport photo path", "gambar pasport", "image photo"],
  },
  {
    id: "parentName",
    label: "Nama Ibu/Bapa/Penjaga",
    type: "text",
    fallback: "",
    aliases: ["parent name", "guardian name", "nama penjaga", "nama ibu bapa"],
  },
  {
    id: "parentIc",
    label: "No KP Ibu/Bapa/Penjaga",
    type: "text",
    fallback: "",
    aliases: ["parent ic", "guardian ic", "no kp penjaga", "no kp ibu bapa"],
  },
  {
    id: "parentDate",
    label: "Tarikh Akuan Ibu/Bapa",
    type: "text",
    fallback: "",
    aliases: ["parent date", "guardian date", "tarikh penjaga", "tarikh ibu bapa"],
  },
  {
    id: "principalName",
    label: "Nama Guru Besar / Pengetua",
    type: "text",
    fallback: "",
    aliases: ["principal name", "pengetua", "guru besar", "headmaster"],
  },
  {
    id: "principalIc",
    label: "No KP Guru Besar / Pengetua",
    type: "text",
    fallback: "",
    aliases: ["principal ic", "no kp pengetua", "no kp guru besar"],
  },
  {
    id: "schoolName",
    label: "Nama Sekolah",
    type: "text",
    fallback: "SMK KATHOLIK",
    aliases: ["school name", "nama sekolah", "school"],
  },
  {
    id: "schoolPhone",
    label: "Telefon Sekolah",
    type: "text",
    fallback: "054596241",
    aliases: ["school phone", "telefon sekolah", "tel sekolah"],
  },
  {
    id: "principalDate",
    label: "Tarikh Akuan Pengetua",
    type: "text",
    fallback: "",
    aliases: ["principal date", "tarikh pengetua", "date principal"],
  },
  {
    id: "managerName",
    label: "Nama Pengurus Pasukan",
    type: "text",
    fallback: "MUHAMMAD ASHRAF BIN AZAMMUDIN",
    aliases: ["manager name", "pengurus pasukan", "team manager"],
  },
  {
    id: "managerIc",
    label: "No KP Pengurus Pasukan",
    type: "text",
    fallback: "981216086001",
    aliases: ["manager ic", "no kp pengurus", "team manager ic"],
  },
  {
    id: "managerDate",
    label: "Tarikh Pengurus Pasukan",
    type: "text",
    fallback: "",
    aliases: ["manager date", "tarikh pengurus"],
  },
  {
    id: "secretaryName",
    label: "Nama Setiausaha MSS / Unit Sukan",
    type: "text",
    fallback: "",
    aliases: ["secretary name", "setiausaha mss", "unit sukan", "mss secretary"],
  },
  {
    id: "secretaryStamp",
    label: "Cop Rasmi / Catatan Setiausaha",
    type: "text",
    fallback: "",
    aliases: ["secretary stamp", "cop rasmi", "official stamp"],
  },
  {
    id: "secretaryDate",
    label: "Tarikh Setiausaha MSS / Unit Sukan",
    type: "text",
    fallback: "",
    aliases: ["secretary date", "tarikh setiausaha"],
  },
  {
    id: "icScan",
    label: "IC / Birth Cert Scan Path",
    type: "image",
    fallback: "",
    aliases: ["ic scan", "scan ic", "surat beranak", "ic image", "dokumen ic"],
  },
];

const state = {
  workbook: null,
  rows: [],
  headers: [],
  sheetNames: [],
  currentSheet: "",
  currentRowIndex: 0,
  mappings: {},
  manualValues: Object.fromEntries(FIELD_DEFS.map((field) => [field.id, field.fallback])),
  assetIndex: new Map(),
  assetFiles: [],
  previewToken: 0,
};

const els = {
  dataFile: document.querySelector("#dataFile"),
  sheetField: document.querySelector("#sheetField"),
  sheetSelect: document.querySelector("#sheetSelect"),
  fileStatus: document.querySelector("#fileStatus"),
  assetFolder: document.querySelector("#assetFolder"),
  assetStatus: document.querySelector("#assetStatus"),
  prevRow: document.querySelector("#prevRow"),
  nextRow: document.querySelector("#nextRow"),
  rowIndex: document.querySelector("#rowIndex"),
  rowStatus: document.querySelector("#rowStatus"),
  autoMap: document.querySelector("#autoMap"),
  exportCurrent: document.querySelector("#exportCurrent"),
  exportAll: document.querySelector("#exportAll"),
  exportStatus: document.querySelector("#exportStatus"),
  previewMeta: document.querySelector("#previewMeta"),
  mappingGrid: document.querySelector("#mappingGrid"),
  page1Canvas: document.querySelector("#page1Canvas"),
  page2Canvas: document.querySelector("#page2Canvas"),
};

init();

function init() {
  buildMappingUi();
  bindEvents();
  drawEmptyPreview();
  autoMapHeaders();
  renderPreview();
}

function bindEvents() {
  els.dataFile.addEventListener("change", onDataFileChange);
  els.sheetSelect.addEventListener("change", () => {
    state.currentSheet = els.sheetSelect.value;
    hydrateRowsFromSheet();
  });
  els.assetFolder.addEventListener("change", onAssetFolderChange);
  els.prevRow.addEventListener("click", () => setCurrentRow(state.currentRowIndex - 1));
  els.nextRow.addEventListener("click", () => setCurrentRow(state.currentRowIndex + 1));
  els.rowIndex.addEventListener("change", () => setCurrentRow(Number(els.rowIndex.value) - 1));
  els.autoMap.addEventListener("click", () => {
    autoMapHeaders();
    syncMappingInputs();
    renderPreview();
  });
  els.exportCurrent.addEventListener("click", () => exportPdf(false));
  els.exportAll.addEventListener("click", () => exportPdf(true));
}

function buildMappingUi() {
  const fragment = document.createDocumentFragment();

  for (const field of FIELD_DEFS) {
    const wrap = document.createElement("section");
    wrap.className = "mapping-item";
    wrap.innerHTML = `
      <h3>${field.label}</h3>
      <label class="field">
        <span>Column</span>
        <select data-map-id="${field.id}">
          <option value="">Manual / fixed value</option>
        </select>
      </label>
      <label class="field">
        <span>${field.type === "image" ? "Fallback path or URL" : "Fallback value"}</span>
        <input data-manual-id="${field.id}" type="text" value="${escapeHtml(field.fallback)}">
      </label>
    `;
    fragment.appendChild(wrap);
  }

  els.mappingGrid.innerHTML = "";
  els.mappingGrid.appendChild(fragment);

  els.mappingGrid.querySelectorAll("select[data-map-id]").forEach((select) => {
    select.addEventListener("change", (event) => {
      state.mappings[event.target.dataset.mapId] = event.target.value;
      renderPreview();
    });
  });

  els.mappingGrid.querySelectorAll("input[data-manual-id]").forEach((input) => {
    input.addEventListener("input", (event) => {
      state.manualValues[event.target.dataset.manualId] = event.target.value;
      renderPreview();
    });
  });
}

function populateMappingOptions() {
  const options = state.headers
    .map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`)
    .join("");

  els.mappingGrid.querySelectorAll("select[data-map-id]").forEach((select) => {
    const selected = state.mappings[select.dataset.mapId] || "";
    select.innerHTML = `<option value="">Manual / fixed value</option>${options}`;
    select.value = state.headers.includes(selected) ? selected : "";
  });
}

function syncMappingInputs() {
  els.mappingGrid.querySelectorAll("select[data-map-id]").forEach((select) => {
    select.value = state.headers.includes(state.mappings[select.dataset.mapId]) ? state.mappings[select.dataset.mapId] : "";
  });

  els.mappingGrid.querySelectorAll("input[data-manual-id]").forEach((input) => {
    input.value = state.manualValues[input.dataset.manualId] ?? "";
  });
}

async function onDataFileChange(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const data = await file.arrayBuffer();
    const workbook = window.XLSX.read(data, {
      type: "array",
      cellDates: true,
      raw: false,
    });

    state.workbook = workbook;
    state.sheetNames = workbook.SheetNames || [];
    state.currentSheet = state.sheetNames[0] || "";

    els.sheetSelect.innerHTML = state.sheetNames
      .map((sheetName) => `<option value="${escapeHtml(sheetName)}">${escapeHtml(sheetName)}</option>`)
      .join("");
    els.sheetField.classList.toggle("hidden", state.sheetNames.length <= 1);
    els.fileStatus.textContent = `${file.name} loaded with ${state.sheetNames.length} sheet${state.sheetNames.length === 1 ? "" : "s"}.`;

    hydrateRowsFromSheet();
  } catch (error) {
    console.error(error);
    els.fileStatus.textContent = `Failed to read file: ${error.message}`;
  }
}

function hydrateRowsFromSheet() {
  if (!state.workbook || !state.currentSheet) {
    return;
  }

  const sheet = state.workbook.Sheets[state.currentSheet];
  const rows = window.XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
  });

  state.rows = rows;
  state.headers = collectHeaders(rows);
  state.currentRowIndex = 0;
  els.sheetSelect.value = state.currentSheet;
  populateMappingOptions();
  autoMapHeaders();
  syncMappingInputs();
  updateRowUi();
  renderPreview();
}

async function onAssetFolderChange(event) {
  const files = Array.from(event.target.files || []);
  state.assetFiles = files;
  state.assetIndex = buildAssetIndex(files);
  els.assetStatus.textContent = files.length
    ? `${files.length} asset file${files.length === 1 ? "" : "s"} indexed.`
    : "No asset folder selected.";
  renderPreview();
}

function collectHeaders(rows) {
  const set = new Set();
  for (const row of rows) {
    Object.keys(row || {}).forEach((key) => set.add(key));
  }
  return Array.from(set);
}

function autoMapHeaders() {
  const headerMap = new Map(state.headers.map((header) => [normalizeKey(header), header]));

  for (const field of FIELD_DEFS) {
    let winner = "";
    for (const alias of [field.label, ...field.aliases]) {
      const normalAlias = normalizeKey(alias);
      for (const [normalizedHeader, sourceHeader] of headerMap.entries()) {
        if (
          normalizedHeader === normalAlias ||
          normalizedHeader.includes(normalAlias) ||
          normalAlias.includes(normalizedHeader)
        ) {
          winner = sourceHeader;
          break;
        }
      }
      if (winner) {
        break;
      }
    }
    if (winner) {
      state.mappings[field.id] = winner;
    }
  }
}

function setCurrentRow(index) {
  if (!state.rows.length) {
    state.currentRowIndex = 0;
    updateRowUi();
    renderPreview();
    return;
  }
  state.currentRowIndex = Math.max(0, Math.min(index, state.rows.length - 1));
  updateRowUi();
  renderPreview();
}

function updateRowUi() {
  const total = state.rows.length;
  els.rowIndex.value = total ? String(state.currentRowIndex + 1) : "1";
  els.rowIndex.max = String(Math.max(1, total));
  els.rowStatus.textContent = total
    ? `Showing record ${state.currentRowIndex + 1} of ${total}.`
    : "No records available.";
}

function getCurrentRow() {
  return state.rows[state.currentRowIndex] || {};
}

function getFieldValue(fieldId, row) {
  const mappedColumn = state.mappings[fieldId];
  const mappedValue = mappedColumn ? row?.[mappedColumn] : "";
  const fallbackValue = state.manualValues[fieldId] ?? "";
  return `${mappedValue || fallbackValue || ""}`.trim();
}

function getRecord(row) {
  const values = {};
  for (const field of FIELD_DEFS) {
    values[field.id] = getFieldValue(field.id, row);
  }

  return {
    ...values,
    formCode: "M01",
  };
}

function drawEmptyPreview() {
  [els.page1Canvas, els.page2Canvas].forEach((canvas) => {
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  });
}

async function renderPreview() {
  const token = ++state.previewToken;
  const record = getRecord(getCurrentRow());

  els.previewMeta.textContent = state.rows.length
    ? `Previewing row ${state.currentRowIndex + 1} from ${state.currentSheet || "source"}.`
    : "Previewing fallback values only.";

  const [logoAsset, photoAsset, icAsset] = await Promise.all([
    resolveImageAsset(LOGO_PATH),
    resolveImageAsset(record.passportPhoto),
    resolveImageAsset(record.icScan),
  ]);

  if (token !== state.previewToken) {
    return;
  }

  await renderToCanvas(els.page1Canvas, 1, record, {
    logo: logoAsset,
    passportPhoto: photoAsset,
  });
  await renderToCanvas(els.page2Canvas, 2, record, {
    icScan: icAsset,
  });
}

async function renderToCanvas(canvas, pageNumber, record, assets) {
  const ctx = canvas.getContext("2d");
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.scale(PREVIEW_SCALE, PREVIEW_SCALE);
  const renderer = new CanvasRenderer(ctx);
  await drawFormPage(renderer, pageNumber, record, assets);
}

async function exportPdf(allRows) {
  try {
    const rows = allRows ? state.rows : [getCurrentRow()];
    if (!rows.length) {
      els.exportStatus.textContent = "No data rows to export.";
      return;
    }

    els.exportStatus.textContent = `Preparing ${allRows ? rows.length : 1} record${allRows && rows.length !== 1 ? "s" : ""}...`;

    const pdfDoc = await PDFDocument.create();
    const fonts = {
      regular: await pdfDoc.embedFont(StandardFonts.Helvetica),
      bold: await pdfDoc.embedFont(StandardFonts.HelveticaBold),
    };

    for (let index = 0; index < rows.length; index += 1) {
      const record = getRecord(rows[index]);
      const assets = {
        logo: await resolveImageAsset(LOGO_PATH),
        passportPhoto: await resolveImageAsset(record.passportPhoto),
        icScan: await resolveImageAsset(record.icScan),
      };

      const page1 = pdfDoc.addPage([PAGE.width, PAGE.height]);
      const page2 = pdfDoc.addPage([PAGE.width, PAGE.height]);
      const renderer1 = new PdfRenderer(page1, fonts);
      const renderer2 = new PdfRenderer(page2, fonts);
      await drawFormPage(renderer1, 1, record, assets);
      await drawFormPage(renderer2, 2, record, assets);
      els.exportStatus.textContent = `Rendered ${index + 1} of ${rows.length} record${rows.length === 1 ? "" : "s"}...`;
    }

    const bytes = await pdfDoc.save();
    const fileName = allRows ? "m01-bulk-export.pdf" : `${safeFileName(getRecord(getCurrentRow()).fullName || "m01-record")}.pdf`;
    downloadBlob(new Blob([bytes], { type: "application/pdf" }), fileName);
    els.exportStatus.textContent = `Exported ${fileName}.`;
  } catch (error) {
    console.error(error);
    els.exportStatus.textContent = `Export failed: ${error.message}`;
  }
}

async function drawFormPage(renderer, pageNumber, record, assets) {
  renderer.fillPage("#ffffff");
  if (pageNumber === 1) {
    await drawPageOne(renderer, record, assets);
  } else {
    await drawPageTwo(renderer, record, assets);
  }
}

async function drawPageOne(renderer, record, assets) {
  const left = 18;
  const width = PAGE.width - 36;

  renderer.rect(left, 18, width, 44, { stroke: "#000000", lineWidth: 1.2 });
  renderer.rect(left, 18, 76, 44, { stroke: "#000000", lineWidth: 1 });
  renderer.rect(left + 76, 18, width - 170, 22, { stroke: "#000000", fill: "#d9d9d9", lineWidth: 1 });
  renderer.rect(left + 76, 40, width - 170, 22, { stroke: "#000000", fill: "#d9d9d9", lineWidth: 1 });
  renderer.rect(left + width - 94, 18, 94, 44, { stroke: "#000000", lineWidth: 1 });

  if (assets.logo) {
    await renderer.image(assets.logo, left + 6, 22, 64, 36, { mode: "contain" });
  }

  renderer.textBox("Majlis Sukan Sekolah-Sekolah Perak", left + 84, 21, width - 186, 16, {
    fontSize: TYPE.title,
    fontWeight: "bold",
    align: "center",
    valign: "middle",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox("BORANG PENDAFTARAN INDIVIDU", left + 84, 43, width - 186, 14, {
    fontSize: TYPE.subtitle,
    fontWeight: "bold",
    align: "center",
    valign: "middle",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox("M01", left + width - 94, 28, 94, 14, {
    fontSize: 15,
    fontWeight: "bold",
    align: "center",
    valign: "middle",
  });

  drawLabeledRow(renderer, 62, [
    { label: "KEJOHANAN", value: record.championship, labelWidth: 112, valueWidth: width - 112 },
  ]);
  drawLabeledRow(renderer, 86, [
    { label: "MSS DAERAH", value: record.district, labelWidth: 112, valueWidth: 170 },
    { label: "KATEGORI", value: record.category, labelWidth: 110, valueWidth: width - 392 },
  ]);

  sectionHeader(renderer, 122, "MAKLUMAT PESERTA");

  const tableX = left;
  const tableY = 144;
  const leftTableWidth = 423;
  const photoWidth = width - leftTableWidth;
  const numberWidth = 28;
  const labelWidth = 126;
  const valueWidth = leftTableWidth - numberWidth - labelWidth;
  const rowHeights = [22, 22, 22, 22, 44, 22, 22, 22];
  const labels = [
    ["1", "Nama Penuh", record.fullName],
    ["2", "Jantina", record.gender],
    ["3", "No Kad Pengenalan", record.icNumber],
    ["4", "Tarikh Lahir", record.birthDate],
    ["5", "Alamat Rumah", record.homeAddress],
    ["6", "Telefon Rumah", record.homePhone],
    ["7", "Tarikh Daftar", record.registrationDate],
    ["8", "No ID Murid", record.studentId],
  ];

  let runningY = tableY;
  for (let i = 0; i < labels.length; i += 1) {
    const rowHeight = rowHeights[i];
    renderer.rect(tableX, runningY, numberWidth, rowHeight, { stroke: "#000000", lineWidth: 1 });
    renderer.rect(tableX + numberWidth, runningY, labelWidth, rowHeight, { stroke: "#000000", lineWidth: 1 });
    renderer.rect(tableX + numberWidth + labelWidth, runningY, valueWidth, rowHeight, {
      stroke: "#000000",
      lineWidth: 1,
    });
    renderer.textBox(labels[i][0], tableX + 2, runningY + TYPE.insetY, numberWidth - 4, rowHeight - TYPE.insetY * 2, {
      fontSize: TYPE.bodySmall,
      align: "center",
      valign: "middle",
      lineHeight: TYPE.lineHeightTight,
    });
    renderer.textBox(
      labels[i][1],
      tableX + numberWidth + TYPE.insetX,
      runningY + TYPE.insetY,
      labelWidth - TYPE.insetX * 2,
      rowHeight - TYPE.insetY * 2,
      {
      fontSize: TYPE.label,
      fontWeight: "bold",
      valign: "middle",
      lineHeight: TYPE.lineHeightTight,
      },
    );
    renderer.textBox(
      labels[i][2],
      tableX + numberWidth + labelWidth + TYPE.insetX,
      runningY + TYPE.insetY,
      valueWidth - TYPE.insetX * 2,
      rowHeight - TYPE.insetY * 2,
      {
      fontSize: i === 4 ? TYPE.bodyTight : TYPE.body,
      valign: i === 4 ? "top" : "middle",
      lineHeight: i === 4 ? TYPE.lineHeightTight : TYPE.lineHeight,
      paddingTop: i === 4 ? 1 : 0,
      },
    );
    runningY += rowHeight;
  }

  renderer.rect(tableX + leftTableWidth, tableY, photoWidth, rowHeights.reduce((sum, value) => sum + value, 0), {
    stroke: "#000000",
    lineWidth: 1,
  });
  renderer.textBox("<passport photo>", tableX + leftTableWidth + 6, tableY + 6, photoWidth - 12, 14, {
    fontSize: TYPE.bodySmall,
    align: "center",
    color: "#666666",
    lineHeight: TYPE.lineHeightTight,
  });
  if (assets.passportPhoto) {
    await renderer.image(
      assets.passportPhoto,
      tableX + leftTableWidth + 8,
      tableY + 24,
      photoWidth - 16,
      rowHeights.reduce((sum, value) => sum + value, 0) - 32,
      { mode: "contain" },
    );
  }

  drawConsentSection(renderer, 334, record);
  drawPrincipalSection(renderer, 458, record);
  drawManagerSection(renderer, 560, record);
}

async function drawPageTwo(renderer, record, assets) {
  const left = 18;
  const width = PAGE.width - 36;
  sectionHeader(renderer, 18, "SALINAN KAD PENGENALAN/SURAT BERANAK");
  renderer.rect(left, 38, width, 660, { stroke: "#000000", lineWidth: 1 });
  renderer.textBox("[ SCANNED IC IMAGE ]", left + 10, 50, 180, 14, {
    fontSize: TYPE.label,
    fontWeight: "bold",
    color: "#666666",
    lineHeight: TYPE.lineHeightTight,
  });

  if (assets.icScan) {
    await renderer.image(assets.icScan, left + 18, 68, width - 36, 612, { mode: "contain" });
  }

  sectionHeader(renderer, 718, "SURAT GABUNGAN");
}

function drawLabeledRow(renderer, y, entries) {
  let x = 18;
  const height = 24;
  for (const entry of entries) {
    renderer.rect(x, y, entry.labelWidth, height, { stroke: "#000000", lineWidth: 1 });
    renderer.rect(x + entry.labelWidth, y, entry.valueWidth, height, { stroke: "#000000", lineWidth: 1 });
    renderer.textBox(entry.label, x + TYPE.insetX, y + TYPE.insetY, entry.labelWidth - TYPE.insetX * 2, height - TYPE.insetY * 2, {
      fontSize: TYPE.label,
      fontWeight: "bold",
      valign: "middle",
      lineHeight: TYPE.lineHeightTight,
    });
    renderer.textBox(entry.value, x + entry.labelWidth + TYPE.insetX, y + TYPE.insetY, entry.valueWidth - TYPE.insetX * 2, height - TYPE.insetY * 2, {
      fontSize: TYPE.body,
      valign: "middle",
      lineHeight: TYPE.lineHeightTight,
    });
    x += entry.labelWidth + entry.valueWidth;
  }
}

function sectionHeader(renderer, y, title) {
  renderer.rect(18, y, PAGE.width - 36, 20, { stroke: "#000000", fill: "#d9d9d9", lineWidth: 1 });
  renderer.textBox(title, 24, y + 3, PAGE.width - 48, 14, {
    fontSize: TYPE.section,
    fontWeight: "bold",
    align: "center",
    valign: "middle",
    lineHeight: TYPE.lineHeightTight,
  });
}

function drawConsentSection(renderer, y, record) {
  sectionHeader(renderer, y, "AKUAN KEBENARAN IBUBAPA / PENJAGA");
  renderer.rect(18, y + 20, PAGE.width - 36, 104, { stroke: "#000000", lineWidth: 1 });
  renderer.textBox("Saya,", 24, y + 32, 28, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(55, y + 44, 308, y + 44, { lineWidth: 0.9 });
  renderer.textBox(`No KP`, 322, y + 32, 36, 12, {
    fontSize: TYPE.body,
    fontWeight: "bold",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.line(360, y + 44, 468, y + 44, { lineWidth: 0.9 });
  renderer.textBox(
    consentParagraph(record),
    24,
    y + 46,
    PAGE.width - 48,
    40,
    { fontSize: 7.9, lineHeight: 1.14, paddingTop: 0.5 },
  );
  renderer.textBox(record.parentName || "", 60, y + 32, 244, 10, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(record.parentIc || "", 364, y + 32, 98, 10, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox("Tandatangan :", 24, y + 88, 70, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(98, y + 100, 248, y + 100, { lineWidth: 0.9 });
  renderer.textBox("Tarikh :", 252, y + 88, 38, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(290, y + 100, 414, y + 100, { lineWidth: 0.9 });
  renderer.textBox(record.parentDate || "", 294, y + 88, 116, 12, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
}

function drawPrincipalSection(renderer, y, record) {
  sectionHeader(renderer, y, "AKUAN GURU BESAR / PENGETUA");
  renderer.rect(18, y + 20, PAGE.width - 36, 82, { stroke: "#000000", lineWidth: 1 });
  renderer.textBox("Saya,", 24, y + 30, 28, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(55, y + 42, 308, y + 42, { lineWidth: 0.9 });
  renderer.textBox("No KP", 322, y + 30, 36, 12, {
    fontSize: TYPE.body,
    fontWeight: "bold",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.line(360, y + 42, 478, y + 42, { lineWidth: 0.9 });
  renderer.textBox(record.principalName || "", 60, y + 30, 244, 10, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(record.principalIc || "", 364, y + 30, 108, 10, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(
    principalParagraph(),
    24,
    y + 44,
    PAGE.width - 48,
    20,
    { fontSize: 7.9, lineHeight: 1.14, paddingTop: 0.5 },
  );
  renderer.textBox("Nama Sekolah :", 24, y + 60, 74, 12, {
    fontSize: TYPE.body,
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox(record.schoolName || "", 98, y + 60, 162, 12, {
    fontSize: TYPE.body,
    fontWeight: "bold",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox("Tel Sekolah :", 266, y + 60, 64, 12, {
    fontSize: TYPE.body,
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox(record.schoolPhone || "", 330, y + 60, 84, 12, {
    fontSize: TYPE.body,
    fontWeight: "bold",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox("Tanda Tangan:", 24, y + 80, 80, 12, { fontSize: TYPE.body, lineHeight: TYPE.lineHeightTight });
  renderer.line(104, y + 92, 214, y + 92, { lineWidth: 0.9 });
  renderer.textBox("Tarikh :", 252, y + 80, 38, 12, { fontSize: TYPE.body, lineHeight: TYPE.lineHeightTight });
  renderer.line(290, y + 92, 394, y + 92, { lineWidth: 0.9 });
  renderer.textBox(record.principalDate || "", 294, y + 80, 96, 12, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox("Cop Rasmi", 182, y + 86, 72, 10, { fontSize: TYPE.bodySmall, color: "#222222", align: "center", lineHeight: TYPE.lineHeightTight });
}

function drawManagerSection(renderer, y, record) {
  sectionHeader(renderer, y, "AKUAN PENGURUS PASUKAN dan SETIAUSAHA MSS DAERAH/UNIT SUKAN PPD");
  renderer.rect(18, y + 20, PAGE.width - 36, 158, { stroke: "#000000", lineWidth: 1 });
  renderer.textBox(
    managerParagraph(record),
    24,
    y + 28,
    PAGE.width - 48,
    24,
    { fontSize: 7.9, lineHeight: 1.14, paddingTop: 0.5 },
  );
  renderer.textBox("PENGURUS PASUKAN", 50, y + 58, 180, 12, {
    fontSize: TYPE.body,
    fontWeight: "bold",
    align: "center",
    lineHeight: TYPE.lineHeightTight,
  });
  renderer.textBox("SETIAUSAHA MSS DAERAH / UNIT SUKAN PPD", 312, y + 58, 220, 12, {
    fontSize: TYPE.bodySmall,
    fontWeight: "bold",
    align: "center",
    lineHeight: TYPE.lineHeightTight,
  });

  drawSignatureColumn(renderer, 34, y + 78, {
    ttLabel: "T.Tangan",
    nameLabel: "Nama",
    idLabel: "No KP",
    dateLabel: "Tarikh",
    nameValue: record.managerName,
    idValue: record.managerIc,
    dateValue: record.managerDate,
  });

  drawSignatureColumn(renderer, 304, y + 78, {
    ttLabel: "T.Tangan",
    nameLabel: "Nama",
    idLabel: "Cop Rasmi",
    dateLabel: "Tarikh",
    nameValue: record.secretaryName,
    idValue: record.secretaryStamp,
    dateValue: record.secretaryDate,
  });
}

function drawSignatureColumn(renderer, x, y, values) {
  renderer.textBox(`${values.ttLabel} :`, x, y, 56, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(x + 62, y + 12, x + 228, y + 12, { lineWidth: 0.9 });
  renderer.textBox(`${values.nameLabel} :`, x, y + 30, 48, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(values.nameValue || "", x + 58, y + 30, 168, 12, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(`${values.idLabel} :`, x, y + 56, 54, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(values.idValue || "", x + 58, y + 56, 168, 12, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
  renderer.textBox(`${values.dateLabel} :`, x, y + 82, 54, 12, { fontSize: TYPE.body, fontWeight: "bold", lineHeight: TYPE.lineHeightTight });
  renderer.line(x + 62, y + 94, x + 228, y + 94, { lineWidth: 0.9 });
  renderer.textBox(values.dateValue || "", x + 66, y + 82, 156, 12, { fontSize: TYPE.body, valign: "middle", lineHeight: TYPE.lineHeightTight });
}

function consentParagraph(record) {
  return `adalah bapa/ibu/penjaga kepada pelajar di atas, mengesahkan segala keterangan di atas adalah benar dan gambar terbaru adalah betul. Saya juga membenarkan anak/jagaan saya didaftarkan sebagai pemain ${record.category ? `(${record.category}) ` : ""}untuk pasukan ${record.district || "__________________"} dan bersetuju mematuhi Perlembagaan, undang-undang, peraturan-peraturan, pekeliling-pekeliling dan syarat-syarat mengenai pertandingan ini. Saya juga faham bahawa pihak tuan sentiasa memberi segala penerangan dan akan mengambil langkah-langkah keselamatan dan pengawasan yang diperlukan sepanjang masa kejohanan tersebut.`;
}

function principalParagraph() {
  return "mengesahkan segala keterangan di atas adalah betul dan gambar di atas adalah terbaru dan benar.";
}

function managerParagraph(record) {
  return `Diakui penama di atas adalah ahli pasukan ini yang menyertai ${record.championship || "Kejohanan / Pertandingan Majlis Sukan Sekolah-sekolah Perak"} dan sepanjang pengetahuan saya semua maklumat yang diberikan adalah benar.`;
}

class CanvasRenderer {
  constructor(ctx) {
    this.ctx = ctx;
  }

  fillPage(color) {
    this.ctx.fillStyle = color;
    this.ctx.fillRect(0, 0, PAGE.width, PAGE.height);
  }

  rect(x, y, width, height, options = {}) {
    this.ctx.save();
    if (options.fill) {
      this.ctx.fillStyle = options.fill;
      this.ctx.fillRect(x, y, width, height);
    }
    this.ctx.lineWidth = options.lineWidth || 1;
    this.ctx.strokeStyle = options.stroke || "#000000";
    this.ctx.strokeRect(x, y, width, height);
    this.ctx.restore();
  }

  line(x1, y1, x2, y2, options = {}) {
    this.ctx.save();
    this.ctx.beginPath();
    this.ctx.moveTo(x1, y1);
    this.ctx.lineTo(x2, y2);
    this.ctx.lineWidth = options.lineWidth || 1;
    this.ctx.strokeStyle = options.stroke || "#000000";
    this.ctx.stroke();
    this.ctx.restore();
  }

  measureText(text, fontSize, fontWeight = "regular") {
    this.ctx.save();
    this.ctx.font = `${fontWeight === "bold" ? "700" : "400"} ${fontSize}px Helvetica, Arial, sans-serif`;
    const width = this.ctx.measureText(text).width;
    this.ctx.restore();
    return width;
  }

  textBox(text, x, y, width, height, options = {}) {
    drawTextBoxGeneric(this, text, x, y, width, height, options);
  }

  drawLines(lines, x, y, fontSize, options = {}) {
    this.ctx.save();
    this.ctx.font = `${options.fontWeight === "bold" ? "700" : "400"} ${fontSize}px Helvetica, Arial, sans-serif`;
    this.ctx.fillStyle = options.color || "#000000";
    this.ctx.textBaseline = "top";
    for (let index = 0; index < lines.length; index += 1) {
      const lineWidth = this.measureText(lines[index], fontSize, options.fontWeight);
      let drawX = x;
      if (options.align === "center") {
        drawX = x + (options.width - lineWidth) / 2;
      } else if (options.align === "right") {
        drawX = x + options.width - lineWidth;
      }
      this.ctx.fillText(lines[index], drawX, y + index * fontSize * (options.lineHeight || 1.18));
    }
    this.ctx.restore();
  }

  async image(asset, x, y, width, height, options = {}) {
    const image = await loadImageElement(asset.url);
    const rect = fitRect(image.naturalWidth, image.naturalHeight, x, y, width, height, options.mode || "contain");
    this.ctx.drawImage(image, rect.x, rect.y, rect.width, rect.height);
  }
}

class PdfRenderer {
  constructor(page, fonts) {
    this.page = page;
    this.fonts = fonts;
  }

  fillPage(color) {
    this.page.drawRectangle({
      x: 0,
      y: 0,
      width: PAGE.width,
      height: PAGE.height,
      color: toRgb(color),
    });
  }

  rect(x, y, width, height, options = {}) {
    this.page.drawRectangle({
      x,
      y: PAGE.height - y - height,
      width,
      height,
      borderColor: toRgb(options.stroke || "#000000"),
      borderWidth: options.lineWidth || 1,
      color: options.fill ? toRgb(options.fill) : undefined,
    });
  }

  line(x1, y1, x2, y2, options = {}) {
    this.page.drawLine({
      start: { x: x1, y: PAGE.height - y1 },
      end: { x: x2, y: PAGE.height - y2 },
      thickness: options.lineWidth || 1,
      color: toRgb(options.stroke || "#000000"),
    });
  }

  measureText(text, fontSize, fontWeight = "regular") {
    const font = fontWeight === "bold" ? this.fonts.bold : this.fonts.regular;
    return font.widthOfTextAtSize(text, fontSize);
  }

  textBox(text, x, y, width, height, options = {}) {
    drawTextBoxGeneric(this, text, x, y, width, height, options);
  }

  drawLines(lines, x, y, fontSize, options = {}) {
    const font = options.fontWeight === "bold" ? this.fonts.bold : this.fonts.regular;
    const lineHeight = fontSize * (options.lineHeight || 1.18);
    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index];
      const lineWidth = this.measureText(line, fontSize, options.fontWeight);
      let drawX = x;
      if (options.align === "center") {
        drawX = x + (options.width - lineWidth) / 2;
      } else if (options.align === "right") {
        drawX = x + options.width - lineWidth;
      }
      this.page.drawText(line, {
        x: drawX,
        y: PAGE.height - y - fontSize - index * lineHeight,
        size: fontSize,
        font,
        color: toRgb(options.color || "#000000"),
      });
    }
  }

  async image(asset, x, y, width, height, options = {}) {
    const bytes = await readAssetBytes(asset);
    const embedded = asset.mimeType === "image/png" ? await this.page.doc.embedPng(bytes) : await this.page.doc.embedJpg(bytes);
    const rect = fitRect(embedded.width, embedded.height, x, y, width, height, options.mode || "contain");
    this.page.drawImage(embedded, {
      x: rect.x,
      y: PAGE.height - rect.y - rect.height,
      width: rect.width,
      height: rect.height,
    });
  }
}

function drawTextBoxGeneric(renderer, text, x, y, width, height, options = {}) {
  const content = `${text || ""}`.trim();
  if (!content) {
    return;
  }

  let fontSize = options.fontSize || 10;
  const minFontSize = options.minFontSize || 7;
  const lineHeight = options.lineHeight || 1.18;
  let lines = [];

  while (fontSize >= minFontSize) {
    lines = wrapText(content, width, (value) => renderer.measureText(value, fontSize, options.fontWeight));
    const totalHeight = lines.length * fontSize * lineHeight;
    if (totalHeight <= height || fontSize === minFontSize) {
      break;
    }
    fontSize -= 0.4;
  }

  const totalHeight = lines.length * fontSize * lineHeight;
  let drawY = y;
  if (options.valign === "middle") {
    drawY = y + (height - totalHeight) / 2;
  } else if (options.valign === "bottom") {
    drawY = y + height - totalHeight;
  }
  drawY += options.paddingTop || 0;

  renderer.drawLines(lines, x, drawY, fontSize, {
    width,
    align: options.align || "left",
    color: options.color || "#000000",
    fontWeight: options.fontWeight || "regular",
    lineHeight,
  });
}

function wrapText(text, maxWidth, measure) {
  const paragraphs = text.split(/\r?\n/);
  const lines = [];

  for (const paragraph of paragraphs) {
    const words = paragraph.split(/\s+/).filter(Boolean);
    if (!words.length) {
      lines.push("");
      continue;
    }
    let line = words[0];
    for (let index = 1; index < words.length; index += 1) {
      const candidate = `${line} ${words[index]}`;
      if (measure(candidate) <= maxWidth) {
        line = candidate;
      } else if (measure(words[index]) <= maxWidth) {
        lines.push(line);
        line = words[index];
      } else {
        lines.push(line);
        const chunks = breakLongWord(words[index], maxWidth, measure);
        lines.push(...chunks.slice(0, -1));
        line = chunks[chunks.length - 1];
      }
    }
    lines.push(line);
  }

  return lines;
}

function breakLongWord(word, maxWidth, measure) {
  const chunks = [];
  let current = "";
  for (const char of word) {
    const candidate = current + char;
    if (measure(candidate) <= maxWidth || !current) {
      current = candidate;
    } else {
      chunks.push(current);
      current = char;
    }
  }
  if (current) {
    chunks.push(current);
  }
  return chunks;
}

function normalizeKey(value) {
  return `${value || ""}`
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildAssetIndex(files) {
  const map = new Map();
  for (const file of files) {
    const relativePath = normalizePath(file.webkitRelativePath || file.name);
    const baseName = normalizePath(file.name);
    map.set(relativePath, file);
    map.set(baseName, file);
    const pieces = relativePath.split("/");
    if (pieces.length > 1) {
      map.set(pieces.slice(1).join("/"), file);
    }
  }
  return map;
}

function normalizePath(value) {
  return `${value || ""}`
    .replace(/\\/g, "/")
    .replace(/^\.?\//, "")
    .toLowerCase()
    .trim();
}

async function resolveImageAsset(rawValue) {
  if (!rawValue) {
    return null;
  }

  if (rawValue === LOGO_PATH) {
    return {
      url: LOGO_PATH,
      mimeType: "image/png",
      bytes: null,
    };
  }

  const normalized = normalizePath(rawValue);
  const basename = normalized.split("/").pop();
  const localFile = state.assetIndex.get(normalized) || state.assetIndex.get(basename);
  if (localFile) {
    return {
      url: URL.createObjectURL(localFile),
      file: localFile,
      mimeType: localFile.type || inferMimeType(localFile.name),
    };
  }

  if (/^https?:\/\//i.test(rawValue) || rawValue.startsWith("blob:") || rawValue.startsWith("data:")) {
    return {
      url: rawValue,
      mimeType: rawValue.startsWith("data:image/png") ? "image/png" : "image/jpeg",
    };
  }

  return null;
}

async function readAssetBytes(asset) {
  if (asset.file) {
    return asset.file.arrayBuffer();
  }
  if (!asset.bytes) {
    const response = await fetch(asset.url);
    asset.bytes = await response.arrayBuffer();
    if (!asset.mimeType) {
      asset.mimeType = response.headers.get("content-type") || "image/png";
    }
  }
  return asset.bytes;
}

function inferMimeType(name) {
  const lower = name.toLowerCase();
  if (lower.endsWith(".png")) {
    return "image/png";
  }
  return "image/jpeg";
}

function fitRect(sourceWidth, sourceHeight, x, y, width, height, mode = "contain") {
  const ratio = sourceWidth / sourceHeight;
  const boxRatio = width / height;
  let drawWidth;
  let drawHeight;

  if ((mode === "contain" && ratio > boxRatio) || (mode === "cover" && ratio < boxRatio)) {
    drawWidth = width;
    drawHeight = width / ratio;
  } else {
    drawHeight = height;
    drawWidth = height * ratio;
  }

  return {
    x: x + (width - drawWidth) / 2,
    y: y + (height - drawHeight) / 2,
    width: drawWidth,
    height: drawHeight,
  };
}

function toRgb(hex) {
  const value = hex.replace("#", "");
  const expanded = value.length === 3 ? value.split("").map((char) => char + char).join("") : value;
  const red = parseInt(expanded.slice(0, 2), 16) / 255;
  const green = parseInt(expanded.slice(2, 4), 16) / 255;
  const blue = parseInt(expanded.slice(4, 6), 16) / 255;
  return rgb(red, green, blue);
}

function safeFileName(value) {
  return `${value || "export"}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function escapeHtml(value) {
  return `${value || ""}`
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function loadImageElement(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = url;
  });
}
