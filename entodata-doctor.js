(function () {
  "use strict";

  const root = document.getElementById("entodata-root");
  if (!root) return;

  const FIELD_DEFINITIONS = [
    {
      key: "specimenId",
      label: "Specimen ID",
      hint: "Unique identifier",
      aliases: [
        "specimen id",
        "specimenid",
        "catalog number",
        "catalognumber",
        "catalogue number",
        "collection id",
        "voucher id",
        "belegnummer",
        "objektnummer",
        "nummer",
        "id",
      ],
    },
    {
      key: "scientificName",
      label: "Scientific name",
      hint: "Taxon or determination",
      aliases: [
        "scientific name",
        "scientificname",
        "taxon",
        "species",
        "art",
        "bestimmung",
        "determination",
        "name",
      ],
    },
    {
      key: "eventDate",
      label: "Collection date",
      hint: "Date or event date",
      aliases: [
        "event date",
        "eventdate",
        "collection date",
        "collecting date",
        "date collected",
        "sammeldatum",
        "funddatum",
        "datum",
        "date",
      ],
    },
    {
      key: "eventTime",
      label: "Collection time",
      hint: "Single time or time range",
      aliases: [
        "event time",
        "eventtime",
        "collection time",
        "collecting time",
        "sampling time",
        "uhrzeit",
        "zeit",
        "time",
      ],
    },
    {
      key: "locality",
      label: "Locality",
      hint: "Verbatim place name",
      aliases: [
        "verbatim locality",
        "verbatimlocality",
        "locality",
        "location",
        "fundort",
        "fundstelle",
        "ort",
        "place",
      ],
    },
    {
      key: "collector",
      label: "Collector",
      hint: "Person who collected it",
      aliases: [
        "recorded by",
        "recordedby",
        "collector",
        "collectors",
        "sammler",
        "legit",
        "leg",
      ],
    },
    {
      key: "latitude",
      label: "Latitude",
      hint: "Decimal or DMS",
      aliases: [
        "decimal latitude",
        "decimallatitude",
        "latitude",
        "lat",
        "breitengrad",
        "gps lat",
      ],
    },
    {
      key: "longitude",
      label: "Longitude",
      hint: "Decimal or DMS",
      aliases: [
        "decimal longitude",
        "decimallongitude",
        "longitude",
        "long",
        "lon",
        "lng",
        "langengrad",
        "gps lon",
      ],
    },
    {
      key: "altitude",
      label: "Altitude",
      hint: "Elevation in metres",
      aliases: [
        "verbatim elevation",
        "elevation",
        "altitude",
        "hohe",
        "height",
        "meters above sea level",
        "masl",
      ],
    },
    {
      key: "sex",
      label: "Sex",
      hint: "Female, male or unknown",
      aliases: ["sex", "geschlecht", "gender"],
    },
    {
      key: "lifeStage",
      label: "Life stage",
      hint: "Adult, larva, pupa…",
      aliases: [
        "life stage",
        "lifestage",
        "stage",
        "stadium",
        "entwicklungsstadium",
      ],
    },
    {
      key: "method",
      label: "Collecting method",
      hint: "Trap or sampling method",
      aliases: [
        "sampling protocol",
        "samplingprotocol",
        "collecting method",
        "collection method",
        "fangmethode",
        "sammelmethode",
        "method",
      ],
    },
    {
      key: "habitat",
      label: "Habitat",
      hint: "Habitat description",
      aliases: ["habitat", "biotope", "biotop", "lebensraum"],
    },
    {
      key: "host",
      label: "Host plant / host",
      hint: "Associated organism",
      aliases: [
        "host plant",
        "hostplant",
        "host",
        "wirtspflanze",
        "wirt",
        "associated taxon",
      ],
    },
  ];

  const EMPTY_MARKERS = new Set([
    "-",
    "—",
    "n/a",
    "na",
    "none",
    "null",
    "not available",
  ]);

  const CONTROLLED_TERMS = {
    sex: {
      female: "female",
      f: "female",
      weibchen: "female",
      "♀": "female",
      male: "male",
      m: "male",
      mannchen: "male",
      "♂": "male",
      unknown: "unknown",
      unbekannt: "unknown",
    },
    lifeStage: {
      adult: "adult",
      imago: "adult",
      erwachsen: "adult",
      larva: "larva",
      larve: "larva",
      larvae: "larva",
      pupa: "pupa",
      puppe: "pupa",
      nymph: "nymph",
      nymphe: "nymph",
      egg: "egg",
      ei: "egg",
    },
    method: {
      lt: "light trap",
      "lighttrap": "light trap",
      "light trap": "light trap",
      "at light": "at light",
      hand: "hand collected",
      "hand collecting": "hand collected",
      "hand collected": "hand collected",
      sweep: "sweep net",
      "sweep net": "sweep net",
      pitfall: "pitfall trap",
      "pitfall trap": "pitfall trap",
      malaise: "Malaise trap",
      "malaise trap": "Malaise trap",
    },
  };

  const state = {
    screen: "upload",
    fileName: "",
    workbook: null,
    sheetName: "",
    matrix: [],
    headerRow: 0,
    rawHeaders: [],
    columns: [],
    rows: [],
    originalRows: [],
    excelRows: [],
    mapping: {},
    issues: [],
    changes: [],
    filter: "all",
    issueView: "grouped",
    issueStatus: "open",
    ignoredIssueKeys: new Set(),
    bulkFill: null,
    previewMode: "current",
    previewPage: 0,
    previewRowsPerPage: 20,
    previewScrollLeft: 0,
    previewScrollTop: 0,
    error: "",
    exported: false,
  };

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function displayValue(value) {
    if (value === "" || value == null) return "empty";
    return String(value);
  }

  function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function normalizedText(value) {
    return String(value == null ? "" : value)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/ß/g, "ss")
      .replace(/[^a-zA-Z0-9♀♂]+/g, " ")
      .trim()
      .toLowerCase();
  }

  function normalizedHeader(value) {
    return normalizedText(value).replace(/\s+/g, "");
  }

  function cleanWhitespace(value) {
    if (typeof value !== "string") return value;
    return value.replace(/[\t\r\n]+/g, " ").replace(/\s{2,}/g, " ").trim();
  }

  function headerMarkup(activeIndex) {
    const steps = ["Upload", "Diagnose", "Fix", "Export"];
    return `
      <header class="app-header">
        <div class="brand">
          <div class="brand-mark" aria-hidden="true">🪲<span class="brand-cross">+</span></div>
          <div class="brand-copy">
            <h1 class="brand-name">EntoData Doctor</h1>
            <p class="brand-tag">Free &amp; open source · data stays local</p>
          </div>
        </div>
        <nav class="stepper" aria-label="Workflow">
          ${steps
            .map((step, index) => {
              const status =
                index < activeIndex
                  ? "is-complete"
                  : index === activeIndex
                    ? "is-active"
                    : "";
              return `<div class="step ${status}">
                <span class="step-dot" aria-hidden="true"></span>
                <span>${step}</span>
              </div>`;
            })
            .join("")}
        </nav>
      </header>`;
  }

  function shell(content, activeIndex) {
    root.innerHTML = `<div class="app-shell">${headerMarkup(activeIndex)}${content}</div>`;
  }

  function renderUpload() {
    shell(
      `
      <main class="screen hero">
        <section class="hero-left">
          <p class="eyebrow">A careful check-up for collection data</p>
          <h2>Give messy specimen data a careful check-up.</h2>
          <p class="hero-copy">
            Clean, validate, and prepare entomological spreadsheets—without
            changing your original file or guessing scientific information.
          </p>
          ${
            state.error
              ? `<div class="error-banner" role="alert"><strong>Could not open the table.</strong><span>${escapeHtml(
                  state.error,
                )}</span></div>`
              : ""
          }
          <label class="upload-zone" id="upload-zone" tabindex="0">
            <input id="file-input" type="file" accept=".xlsx,.xls,.csv,.tsv" />
            <span class="file-symbol" aria-hidden="true"></span>
            <span class="upload-title">Drop your specimen table here</span>
            <span class="upload-subtitle">or choose a file</span>
            <span class="file-types">.xlsx&nbsp;&nbsp; .xls&nbsp;&nbsp; .csv&nbsp;&nbsp; .tsv</span>
            <span class="primary-button">Choose file <span class="button-arrow">→</span></span>
          </label>
          <div class="hero-actions">
            <button class="ghost-button" id="sample-button" type="button">Try the intentionally messy example</button>
          </div>
          <p class="privacy-note"><span class="shield" aria-hidden="true">✓</span>Your data stays in your browser</p>
        </section>

        <aside class="diagnosis-card" aria-label="Example diagnosis">
          <h3 class="card-heading"><span class="clipboard" aria-hidden="true">✓</span>Example diagnosis</h3>
          <div class="mini-table" aria-hidden="true">
            <div class="mini-row mini-head"><span>ID</span><span>Species</span><span>Date</span><span>Locality</span></div>
            <div class="mini-row"><span>ENT-001</span><span>Carabus violaceus</span><span>2026-07-12</span><span>Alpine pass</span></div>
            <div class="mini-row"><span>ENT-001</span><span>Pterostichus niger</span><span class="mini-bad">31.02.2026</span><span>River bend</span></div>
            <div class="mini-row"><span>ENT-003</span><span>Staphylinus olens</span><span>2026-07-14</span><span>Pine hollow</span></div>
          </div>
          <div class="diagnostic-list">
            <div class="diagnostic-item"><span class="diagnostic-icon icon-critical">!</span><span>2 duplicate specimen IDs</span><span class="diagnostic-count">critical</span></div>
            <div class="diagnostic-item"><span class="diagnostic-icon icon-warning">△</span><span>3 dates need review</span><span class="diagnostic-count">review</span></div>
            <div class="diagnostic-item"><span class="diagnostic-icon icon-safe">✓</span><span>14 safe fixes available</span><span class="diagnostic-count">one click</span></div>
          </div>
          <div class="diagnosis-footer"><span>53 records checked</span><div class="progress-bar"><div class="progress-fill"></div></div></div>
        </aside>
      </main>`,
      0,
    );

    const input = document.getElementById("file-input");
    const zone = document.getElementById("upload-zone");
    const sampleButton = document.getElementById("sample-button");

    input.addEventListener("change", (event) => {
      const file = event.target.files && event.target.files[0];
      if (file) loadFile(file);
    });

    zone.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        input.click();
      }
    });

    ["dragenter", "dragover"].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => {
        event.preventDefault();
        zone.classList.remove("is-dragging");
      });
    });

    zone.addEventListener("drop", (event) => {
      const file = event.dataTransfer.files && event.dataTransfer.files[0];
      if (file) loadFile(file);
    });

    sampleButton.addEventListener("click", loadSample);
  }

  async function loadFile(file) {
    if (!window.XLSX) {
      state.error = "The spreadsheet reader did not load. Please reload the page.";
      renderUpload();
      return;
    }

    state.error = "";
    const zone = document.getElementById("upload-zone");
    if (zone) {
      zone.querySelector(".upload-title").textContent = "Reading your table…";
      zone.setAttribute("aria-busy", "true");
    }

    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, {
        type: "array",
        cellDates: false,
      });
      initializeWorkbook(workbook, file.name);
    } catch (error) {
      state.error =
        error && error.message
          ? error.message
          : "This file could not be read as Excel or CSV.";
      renderUpload();
    }
  }

  function loadSample() {
    if (!window.XLSX) {
      state.error = "The spreadsheet reader did not load. Please reload the page.";
      renderUpload();
      return;
    }

    const rows = [
      ["Croatia insects — deliberately messy test data"],
      ["Notes: colours and formatting are not used as scientific data"],
      [],
      [
        "Specimen ID",
        "Scientific name",
        "Fundort",
        "Datum",
        "Time",
        "Collector",
        "Latitude",
        "Longitude",
        "Elevation",
        "Sex",
        "Life stage",
        "Fangmethode",
        "Habitat",
      ],
      [
        " ENT-001 ",
        "Carabus violaceus",
        "Alpine pass  ",
        "12.7.2026",
        "14:00",
        "S. Ronenko",
        "47,3769",
        "8,5417",
        "300 M",
        "Weibchen",
        "Imago",
        "hand",
        "dry meadow",
      ],
      [
        "ENT-002",
        "Pterostichus niger",
        "River bend",
        "07/08/2026",
        "14:00-00:00",
        "S. Ronenko",
        "47.4012 N",
        "8.6018 E",
        "1'000m",
        "m",
        "adult",
        "LT",
        "riparian forest",
      ],
      [
        "ENT-002",
        "Staphylinus olens",
        "River bend",
        "31.02.2026",
        "10.00",
        " A. Lee ",
        "112.3",
        "8.6018",
        "1 000 m",
        "♂",
        "Adult",
        "lighttrap",
        "riparian forest",
      ],
      [
        "",
        "Bombus pascuorum",
        "N/A",
        "2026-07-14",
        "25:70",
        "A. Lee",
        "",
        "",
        "450 metres",
        "female",
        "adult",
        "sweep",
        "flower-rich meadow",
      ],
      [
        "ENT-005",
        "",
        "Pine hollow",
        "2031-08-01",
        "",
        "",
        "47.45",
        "",
        "820",
        "?",
        "larve",
        "pitfall",
        "pine forest",
      ],
      [
        "ENT-006",
        " Cicadella viridis ",
        "Wet meadow",
        "18.07.2026",
        "2:30 PM",
        "S. Ronenko",
        "47°22'36.8\"N",
        "8°32'30.1\"E",
        "520 m a.s.l.",
        "f",
        "nymphe",
        "sweep net",
        " wet   meadow ",
      ],
      [
        "ENT-007",
        "Lucanus cervus",
        "Oak woodland",
        "2026-07-19",
        "09:30:00",
        "S. Ronenko",
        "47.38",
        "8.55",
        "410m",
        "male",
        "adult",
        "hand collected",
        "oak woodland",
      ],
    ];

    const workbook = window.XLSX.utils.book_new();
    const dataSheet = window.XLSX.utils.aoa_to_sheet(rows);
    const readmeSheet = window.XLSX.utils.aoa_to_sheet([
      ["EntoData Doctor sample"],
      ["This sheet is not the specimen table."],
    ]);
    window.XLSX.utils.book_append_sheet(workbook, dataSheet, "Specimens");
    window.XLSX.utils.book_append_sheet(workbook, readmeSheet, "Read me");
    initializeWorkbook(workbook, "EntoData_Doctor_messy_example.xlsx");
  }

  function initializeWorkbook(workbook, fileName) {
    if (!workbook.SheetNames || !workbook.SheetNames.length) {
      throw new Error("No worksheets were found.");
    }
    state.workbook = workbook;
    state.fileName = fileName;
    state.sheetName = chooseInitialSheet(workbook);
    state.changes = [];
    state.exported = false;
    state.filter = "all";
    state.issueView = "grouped";
    state.issueStatus = "open";
    state.ignoredIssueKeys = new Set();
    state.bulkFill = null;
    state.previewPage = 0;
    state.previewScrollLeft = 0;
    state.previewScrollTop = 0;
    prepareSheet(true);
    state.screen = "setup";
    renderSetup();
  }

  function chooseInitialSheet(workbook) {
    let best = workbook.SheetNames[0];
    let bestScore = -1;
    workbook.SheetNames.forEach((name) => {
      const matrix = sheetMatrix(workbook, name);
      const populated = matrix.reduce(
        (total, row) =>
          total + row.filter((value) => String(value).trim() !== "").length,
        0,
      );
      if (populated > bestScore) {
        best = name;
        bestScore = populated;
      }
    });
    return best;
  }

  function sheetMatrix(workbook, sheetName) {
    const worksheet = workbook.Sheets[sheetName];
    return window.XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: true,
    });
  }

  function detectHeaderRow(matrix) {
    const limit = Math.min(matrix.length, 25);
    let bestIndex = 0;
    let bestScore = -Infinity;

    for (let index = 0; index < limit; index += 1) {
      const row = matrix[index] || [];
      const values = row
        .map((value) => String(value).trim())
        .filter((value) => value !== "");
      if (values.length < 2) continue;

      const unique = new Set(values.map(normalizedHeader)).size;
      const textLike = values.filter((value) => /[A-Za-zÀ-ž]/.test(value)).length;
      const nextRows = matrix.slice(index + 1, index + 4);
      const density =
        nextRows.reduce(
          (sum, next) =>
            sum +
            (next || []).filter((value) => String(value).trim() !== "").length,
          0,
        ) / Math.max(1, nextRows.length);
      const fieldMatches = FIELD_DEFINITIONS.reduce((count, field) => {
        const aliases = field.aliases.map(normalizedHeader);
        return (
          count +
          values.filter((value) => aliases.includes(normalizedHeader(value))).length
        );
      }, 0);
      const score =
        values.length * 2 +
        unique +
        textLike * 0.5 +
        density * 0.6 +
        fieldMatches * 5 -
        index * 0.03;

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    }
    return bestIndex;
  }

  function makeUniqueHeaders(rawHeaders) {
    const seen = new Map();
    return rawHeaders.map((raw, index) => {
      const base = cleanWhitespace(String(raw || "")) || `Column ${index + 1}`;
      const key = base.toLowerCase();
      const count = (seen.get(key) || 0) + 1;
      seen.set(key, count);
      return count === 1 ? base : `${base} (${count})`;
    });
  }

  function prepareSheet(autodetectHeader) {
    state.matrix = sheetMatrix(state.workbook, state.sheetName);
    if (!state.matrix.length) {
      state.error = "The selected worksheet is empty.";
      state.screen = "upload";
      renderUpload();
      return;
    }

    if (autodetectHeader) state.headerRow = detectHeaderRow(state.matrix);
    const maxColumns = state.matrix.reduce(
      (max, row) => Math.max(max, (row || []).length),
      0,
    );
    state.rawHeaders = Array.from(
      { length: maxColumns },
      (_, index) => (state.matrix[state.headerRow] || [])[index] || "",
    );
    state.columns = makeUniqueHeaders(state.rawHeaders);
    state.rows = [];
    state.excelRows = [];

    state.matrix.slice(state.headerRow + 1).forEach((row, offset) => {
      const hasValue = (row || []).some(
        (value) => String(value == null ? "" : value).trim() !== "",
      );
      if (!hasValue) return;
      const record = {};
      state.columns.forEach((column, index) => {
        record[column] = row && row[index] != null ? row[index] : "";
      });
      state.rows.push(record);
      state.excelRows.push(state.headerRow + 2 + offset);
    });

    state.originalRows = deepClone(state.rows);
    state.mapping = inferMapping(state.columns);
    state.issues = [];
    state.changes = [];
    state.ignoredIssueKeys = new Set();
    state.bulkFill = null;
  }

  function inferMapping(columns) {
    const mapping = {};
    const normalizedColumns = columns.map((column) => ({
      column,
      normalized: normalizedHeader(column),
    }));

    FIELD_DEFINITIONS.forEach((field) => {
      const aliases = field.aliases.map(normalizedHeader);
      let match = normalizedColumns.find((item) =>
        aliases.includes(item.normalized),
      );
      if (!match) {
        match = normalizedColumns.find((item) =>
          aliases.some(
            (alias) =>
              alias.length >= 4 &&
              (item.normalized.includes(alias) || alias.includes(item.normalized)),
          ),
        );
      }
      mapping[field.key] = match ? match.column : "";
    });
    return mapping;
  }

  function columnOptions(selected) {
    return [
      `<option value="">— Not mapped —</option>`,
      ...state.columns.map(
        (column) =>
          `<option value="${escapeHtml(column)}" ${
            column === selected ? "selected" : ""
          }>${escapeHtml(column)}</option>`,
      ),
    ].join("");
  }

  function rawPreviewMarkup() {
    const previewRows = state.rows.slice(0, 8);
    const previewColumns = state.columns.slice(0, 12);
    return `
      <div class="raw-preview">
        <table class="data-table">
          <thead><tr>${previewColumns
            .map((column) => `<th>${escapeHtml(column)}</th>`)
            .join("")}</tr></thead>
          <tbody>
            ${previewRows
              .map(
                (row) =>
                  `<tr>${previewColumns
                    .map((column) => `<td>${escapeHtml(row[column])}</td>`)
                    .join("")}</tr>`,
              )
              .join("")}
          </tbody>
        </table>
      </div>`;
  }

  function renderSetup() {
    const sheetOptions = state.workbook.SheetNames.map(
      (name) =>
        `<option value="${escapeHtml(name)}" ${
          name === state.sheetName ? "selected" : ""
        }>${escapeHtml(name)}</option>`,
    ).join("");
    const headerOptions = Array.from(
      { length: Math.min(30, state.matrix.length) },
      (_, index) =>
        `<option value="${index}" ${
          index === state.headerRow ? "selected" : ""
        }>Row ${index + 1}</option>`,
    ).join("");

    shell(
      `
      <main class="screen workspace">
        <div class="workspace-head">
          <div>
            <p class="eyebrow">Import check</p>
            <h2>Tell me how to read this table.</h2>
            <p>I found a likely header row and suggested field mappings. Nothing has been changed yet.</p>
          </div>
          <span class="file-pill" title="${escapeHtml(state.fileName)}">▣ ${escapeHtml(
            state.fileName,
          )}</span>
        </div>

        <div class="setup-grid">
          <section class="panel">
            <h3>Table structure</h3>
            <p class="panel-copy">Choose the worksheet and the row containing the real column names.</p>
            <div class="control-stack">
              <div class="form-field">
                <label for="sheet-select">Worksheet</label>
                <select id="sheet-select">${sheetOptions}</select>
              </div>
              <div class="form-field">
                <label for="header-select">Header row</label>
                <select id="header-select">${headerOptions}</select>
              </div>
              <div class="info-banner">
                <span aria-hidden="true">ⓘ</span>
                <span>${state.rows.length} populated records and ${state.columns.length} columns found. Empty rows are ignored.</span>
              </div>
            </div>
            <button class="ghost-button" id="new-file-button" type="button">← Choose another file</button>
          </section>

          <section class="panel">
            <h3>Field mapping</h3>
            <p class="panel-copy">Match your column names to the concepts Doctor understands. Unmapped columns are preserved in the export.</p>
            <div class="mapping-grid">
              ${FIELD_DEFINITIONS.map(
                (field) => `
                <div class="mapping-item">
                  <div class="mapping-label">
                    <strong>${escapeHtml(field.label)}</strong>
                    <span>${escapeHtml(field.hint)}</span>
                  </div>
                  <select class="mapping-select" data-field="${field.key}" aria-label="${escapeHtml(
                    field.label,
                  )} column">${columnOptions(state.mapping[field.key])}</select>
                </div>`,
              ).join("")}
            </div>
            ${rawPreviewMarkup()}
            <div class="mapping-actions">
              <span class="records-note">The original file remains untouched.</span>
              <button class="primary-button" id="diagnose-button" type="button">Diagnose ${state.rows.length} records <span class="button-arrow">→</span></button>
            </div>
          </section>
        </div>
      </main>`,
      1,
    );

    document.getElementById("sheet-select").addEventListener("change", (event) => {
      state.sheetName = event.target.value;
      prepareSheet(true);
      renderSetup();
    });

    document.getElementById("header-select").addEventListener("change", (event) => {
      state.headerRow = Number(event.target.value);
      prepareSheet(false);
      renderSetup();
    });

    document.querySelectorAll(".mapping-select").forEach((select) => {
      select.addEventListener("change", (event) => {
        state.mapping[event.target.dataset.field] = event.target.value;
      });
    });

    document.getElementById("diagnose-button").addEventListener("click", () => {
      state.previewPage = 0;
      state.previewScrollLeft = 0;
      state.previewScrollTop = 0;
      runDiagnosis();
      state.screen = "diagnosis";
      state.exported = false;
      renderDiagnosis();
    });

    document.getElementById("new-file-button").addEventListener("click", resetApp);
  }

  function issueIdentity(issue) {
    return JSON.stringify([
      issue.rowIndex == null ? "table" : issue.rowIndex,
      issue.column || "",
      issue.field || "",
      issue.message || "",
      String(issue.original == null ? "" : issue.original),
      String(issue.proposed == null ? "" : issue.proposed),
    ]);
  }

  function issueGroupKey(issue) {
    return JSON.stringify([
      issue.severity || "warning",
      issue.field || "Table",
      issue.message || "Issue",
    ]);
  }

  function addIssue(issue) {
    const completeIssue = {
      id: `issue-${state.issues.length + 1}`,
      rowIndex: null,
      excelRow: null,
      column: "",
      field: "Table",
      severity: "warning",
      message: "",
      detail: "",
      original: "",
      proposed: null,
      safe: false,
      confirm: false,
      ...issue,
    };
    completeIssue.key = issueIdentity(completeIssue);
    completeIssue.groupKey = issueGroupKey(completeIssue);
    completeIssue.ignored = state.ignoredIssueKeys.has(completeIssue.key);
    state.issues.push(completeIssue);
  }

  function valueAt(row, fieldKey) {
    const column = state.mapping[fieldKey];
    return column ? row[column] : "";
  }

  function isEmpty(value) {
    return value == null || String(value).trim() === "";
  }

  function validCalendarDate(year, month, day) {
    const date = new Date(Date.UTC(year, month - 1, day));
    return (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    );
  }

  function isoDate(year, month, day) {
    return `${String(year).padStart(4, "0")}-${String(month).padStart(
      2,
      "0",
    )}-${String(day).padStart(2, "0")}`;
  }

  function dateAnalysis(value) {
    const text = cleanWhitespace(String(value));
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) {
      const year = Number(match[1]);
      const month = Number(match[2]);
      const day = Number(match[3]);
      return validCalendarDate(year, month, day)
        ? { valid: true, iso: isoDate(year, month, day), safe: true }
        : { valid: false, reason: "This is not a real calendar date." };
    }

    match = text.match(/^(\d{1,2})([./-])(\d{1,2})\2(\d{4})$/);
    if (match) {
      const day = Number(match[1]);
      const separator = match[2];
      const month = Number(match[3]);
      const year = Number(match[4]);
      if (!validCalendarDate(year, month, day)) {
        return { valid: false, reason: "This is not a real calendar date." };
      }
      const ambiguous = separator !== "." && day <= 12 && month <= 12;
      return {
        valid: true,
        iso: isoDate(year, month, day),
        safe: !ambiguous,
        ambiguous,
      };
    }

    if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2}$/.test(text)) {
      return {
        valid: null,
        reason: "The two-digit year is ambiguous.",
      };
    }

    const serial = Number(text);
    if (
      Number.isFinite(serial) &&
      serial > 1 &&
      serial < 80000 &&
      window.XLSX &&
      window.XLSX.SSF
    ) {
      const parsed = window.XLSX.SSF.parse_date_code(serial);
      if (parsed && validCalendarDate(parsed.y, parsed.m, parsed.d)) {
        return {
          valid: true,
          iso: isoDate(parsed.y, parsed.m, parsed.d),
          safe: true,
        };
      }
    }

    if (/[A-Za-zÀ-ž]/.test(text)) {
      const timestamp = Date.parse(text);
      if (!Number.isNaN(timestamp)) {
        const parsed = new Date(timestamp);
        return {
          valid: true,
          iso: isoDate(
            parsed.getUTCFullYear(),
            parsed.getUTCMonth() + 1,
            parsed.getUTCDate(),
          ),
          safe: false,
          ambiguous: true,
        };
      }
    }

    return { valid: false, reason: "The date format could not be understood." };
  }

  function normalizeTimePart(value) {
    let text = cleanWhitespace(String(value)).toUpperCase();
    text = text
      .replace(/\s*UHR$/i, "")
      .replace(/^(\d{1,2})[.](\d{2})(?::(\d{2}))?$/, "$1:$2:$3")
      .replace(/^(\d{1,2})\s*H\s*(\d{2})$/i, "$1:$2")
      .replace(/:$/, "");

    const match = text.match(
      /^(\d{1,2})(?::(\d{2}))?(?::(\d{2}))?\s*(AM|PM)?$/,
    );
    if (!match) {
      return { valid: false, reason: "The time format could not be understood." };
    }

    let hour = Number(match[1]);
    const minute = match[2] == null ? 0 : Number(match[2]);
    const second = match[3] == null ? 0 : Number(match[3]);
    const meridiem = match[4] || "";

    if (minute > 59 || second > 59) {
      return { valid: false, reason: "Minutes and seconds must be below 60." };
    }

    if (meridiem) {
      if (hour < 1 || hour > 12) {
        return { valid: false, reason: "12-hour time must use hours from 1 to 12." };
      }
      if (meridiem === "AM" && hour === 12) hour = 0;
      if (meridiem === "PM" && hour !== 12) hour += 12;
    } else if (hour > 23) {
      return { valid: false, reason: "Hours must be between 0 and 23." };
    }

    const formatted = `${String(hour).padStart(2, "0")}:${String(minute).padStart(
      2,
      "0",
    )}${second ? `:${String(second).padStart(2, "0")}` : ""}`;
    return { valid: true, formatted };
  }

  function timeAnalysis(value) {
    const original = cleanWhitespace(String(value));
    const parts = original.split(/\s*(?:-|–|—|\bto\b|\bbis\b)\s*/i);
    if (parts.length > 2 || parts.some((part) => !part)) {
      return { valid: false, reason: "The time range could not be understood." };
    }

    const analysed = parts.map(normalizeTimePart);
    const invalid = analysed.find((part) => !part.valid);
    if (invalid) return invalid;

    return {
      valid: true,
      formatted: analysed.map((part) => part.formatted).join("–"),
    };
  }

  function altitudeAnalysis(value) {
    const original = cleanWhitespace(String(value));
    let text = original
      .replace(
        /\s*(?:m|meter|meters|metre|metres)(?:\s*(?:a\.?\s*s\.?\s*l\.?|asl|ü\.?\s*m\.?))?\s*$/i,
        "",
      )
      .trim();

    if (!text) {
      return { valid: false, reason: "No numeric altitude remains after the unit." };
    }

    const groupedThousands =
      /^[+-]?\d{1,3}(?:['’\s]\d{3})+(?:[.,]\d+)?$/.test(text) ||
      /^[+-]?\d{1,3}(?:[.,]\d{3})+$/.test(text);

    if (groupedThousands) {
      text = text.replace(/['’\s.,]/g, "");
    } else {
      text = text.replace(/['’\s]/g, "");
      if (text.includes(",") && !text.includes(".")) text = text.replace(",", ".");
    }

    if (!/^[+-]?\d+(?:\.\d+)?$/.test(text)) {
      return {
        valid: false,
        reason: "The value still contains text that cannot be interpreted as metres.",
      };
    }

    const number = Number(text);
    if (!Number.isFinite(number)) {
      return { valid: false, reason: "The altitude is not a finite number." };
    }

    return {
      valid: true,
      formatted: Number(number.toFixed(3)).toString(),
    };
  }

  function coordinateAnalysis(value, axis) {
    const text = cleanWhitespace(String(value)).replace(/,/g, ".");
    const decimal = text.match(/^([+-]?\d+(?:\.\d+)?)\s*°?\s*([NSEW])?$/i);
    if (decimal) {
      let number = Number(decimal[1]);
      const direction = (decimal[2] || "").toUpperCase();
      if (
        (axis === "latitude" && /[EW]/.test(direction)) ||
        (axis === "longitude" && /[NS]/.test(direction))
      ) {
        return { valid: false, reason: `Direction ${direction} does not match ${axis}.` };
      }
      if (direction === "S" || direction === "W") number = -Math.abs(number);
      if (direction === "N" || direction === "E") number = Math.abs(number);
      const limit = axis === "latitude" ? 90 : 180;
      if (Math.abs(number) > limit) {
        return {
          valid: false,
          number,
          reason: `${axis === "latitude" ? "Latitude" : "Longitude"} must be between -${limit} and ${limit}.`,
        };
      }
      return {
        valid: true,
        number,
        formatted: formatCoordinateNumber(number),
      };
    }

    const dms = text.match(
      /^(\d{1,3})\s*[° ]\s*(\d{1,2})\s*['′ ]\s*(\d{1,2}(?:\.\d+)?)\s*["″]?\s*([NSEW])$/i,
    );
    if (dms) {
      const degrees = Number(dms[1]);
      const minutes = Number(dms[2]);
      const seconds = Number(dms[3]);
      const direction = dms[4].toUpperCase();
      if (minutes >= 60 || seconds >= 60) {
        return { valid: false, reason: "Minutes and seconds must be below 60." };
      }
      if (
        (axis === "latitude" && /[EW]/.test(direction)) ||
        (axis === "longitude" && /[NS]/.test(direction))
      ) {
        return { valid: false, reason: `Direction ${direction} does not match ${axis}.` };
      }
      let number = degrees + minutes / 60 + seconds / 3600;
      if (direction === "S" || direction === "W") number *= -1;
      const limit = axis === "latitude" ? 90 : 180;
      if (Math.abs(number) > limit) {
        return { valid: false, reason: `${axis} is outside its valid range.` };
      }
      return {
        valid: true,
        number,
        formatted: formatCoordinateNumber(number),
      };
    }

    return { valid: false, reason: "The coordinate format could not be understood." };
  }

  function formatCoordinateNumber(number) {
    return Number(number.toFixed(6)).toString();
  }

  function mappedFieldLabel(key) {
    const field = FIELD_DEFINITIONS.find((item) => item.key === key);
    return field ? field.label : key;
  }

  function runDiagnosis() {
    state.issues = [];

    [
      ["specimenId", "critical", "Specimen ID is not mapped"],
      ["eventDate", "critical", "Collection date is not mapped"],
      ["locality", "critical", "Locality is not mapped"],
      ["scientificName", "warning", "Scientific name is not mapped"],
      ["collector", "warning", "Collector is not mapped"],
    ].forEach(([key, severity, message]) => {
      if (!state.mapping[key]) {
        addIssue({
          severity,
          field: mappedFieldLabel(key),
          message,
          detail:
            "Return to Field mapping and select the column that contains this information.",
        });
      }
    });

    const rawNormalized = state.rawHeaders
      .map((header) => normalizedHeader(header))
      .filter(Boolean);
    const duplicatedHeaders = rawNormalized.filter(
      (header, index) => rawNormalized.indexOf(header) !== index,
    );
    if (duplicatedHeaders.length) {
      addIssue({
        severity: "warning",
        field: "Header",
        message: "Duplicate column names were found",
        detail:
          "Doctor kept both columns by adding a number. Check that the mapping uses the intended one.",
      });
    }

    state.rows.forEach((row, rowIndex) => {
      const excelRow = state.excelRows[rowIndex];

      state.columns.forEach((column) => {
        const original = row[column];
        if (typeof original !== "string") return;
        const trimmed = cleanWhitespace(original);
        const emptyMarker = EMPTY_MARKERS.has(normalizedText(trimmed));
        if (emptyMarker) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: column,
            severity: "formatting",
            message: "Placeholder can be stored as an empty cell",
            detail: "The original marker remains visible in the change log.",
            original,
            proposed: "",
            safe: true,
          });
        } else if (trimmed !== original) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: column,
            severity: "formatting",
            message: "Extra whitespace",
            detail: "Leading, trailing, or repeated spaces can be removed safely.",
            original,
            proposed: trimmed,
            safe: true,
          });
        }
      });

      [
        ["specimenId", "Specimen ID is missing", "Every collection record should have a stable unique identifier.", "critical"],
        ["eventDate", "Collection date is missing", "A specimen label normally needs the collecting date.", "critical"],
        ["locality", "Locality is missing", "A specimen label normally needs a collecting locality.", "critical"],
        ["scientificName", "Scientific name is missing", "The record can be retained, but determination information is incomplete.", "warning"],
        ["collector", "Collector is missing", "The person who collected the specimen is not recorded.", "warning"],
      ].forEach(([key, message, detail, severity]) => {
        const column = state.mapping[key];
        if (column && isEmpty(valueAt(row, key))) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: mappedFieldLabel(key),
            severity,
            message,
            detail,
          });
        }
      });

      const dateColumn = state.mapping.eventDate;
      const dateValue = valueAt(row, "eventDate");
      if (dateColumn && !isEmpty(dateValue)) {
        const analysis = dateAnalysis(dateValue);
        if (analysis.valid === false) {
          addIssue({
            rowIndex,
            excelRow,
            column: dateColumn,
            field: "Collection date",
            severity: "critical",
            message: "Invalid or unreadable date",
            detail: analysis.reason,
            original: dateValue,
          });
        } else if (analysis.valid === null) {
          addIssue({
            rowIndex,
            excelRow,
            column: dateColumn,
            field: "Collection date",
            severity: "warning",
            message: "Ambiguous year",
            detail: analysis.reason,
            original: dateValue,
          });
        } else {
          const isoToday = new Date().toISOString().slice(0, 10);
          if (analysis.iso > isoToday) {
            addIssue({
              rowIndex,
              excelRow,
              column: dateColumn,
              field: "Collection date",
              severity: "critical",
              message: "Collection date is in the future",
              detail: "Check the year before using this record.",
              original: dateValue,
            });
          } else if (String(dateValue).trim() !== analysis.iso) {
            addIssue({
              rowIndex,
              excelRow,
              column: dateColumn,
              field: "Collection date",
              severity: analysis.safe ? "formatting" : "warning",
              message: analysis.ambiguous
                ? "Date order needs confirmation"
                : "Date can be standardized",
              detail: analysis.ambiguous
                ? "The day and month could be interpreted in more than one way."
                : "ISO dates are unambiguous and sort correctly.",
              original: dateValue,
              proposed: analysis.iso,
              safe: Boolean(analysis.safe),
              confirm: !analysis.safe,
            });
          }
        }
      }

      const timeColumn = state.mapping.eventTime;
      const timeValue = valueAt(row, "eventTime");
      if (timeColumn && !isEmpty(timeValue)) {
        const analysis = timeAnalysis(timeValue);
        if (!analysis.valid) {
          addIssue({
            rowIndex,
            excelRow,
            column: timeColumn,
            field: "Collection time",
            severity: "warning",
            message: "Invalid or unreadable time",
            detail: analysis.reason,
            original: timeValue,
          });
        } else if (String(timeValue).trim() !== analysis.formatted) {
          addIssue({
            rowIndex,
            excelRow,
            column: timeColumn,
            field: "Collection time",
            severity: "formatting",
            message: "Time can be standardized",
            detail:
              "Recognized as a clock time or time range and converted to a consistent 24-hour format.",
            original: timeValue,
            proposed: analysis.formatted,
            safe: true,
          });
        }
      }

      ["latitude", "longitude"].forEach((key) => {
        const column = state.mapping[key];
        const value = valueAt(row, key);
        if (!column || isEmpty(value)) return;
        const analysis = coordinateAnalysis(value, key);
        if (!analysis.valid) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: mappedFieldLabel(key),
            severity: "critical",
            message: "Invalid coordinate",
            detail: analysis.reason,
            original: value,
          });
        } else if (String(value).trim() !== analysis.formatted) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: mappedFieldLabel(key),
            severity: "formatting",
            message: "Coordinate can be standardized",
            detail: "Converted to signed decimal degrees without inventing precision.",
            original: value,
            proposed: analysis.formatted,
            safe: true,
          });
        }
      });

      const latitude = valueAt(row, "latitude");
      const longitude = valueAt(row, "longitude");
      if (state.mapping.latitude && state.mapping.longitude) {
        if (isEmpty(latitude) !== isEmpty(longitude)) {
          const missingKey = isEmpty(latitude) ? "latitude" : "longitude";
          addIssue({
            rowIndex,
            excelRow,
            column: state.mapping[missingKey],
            field: mappedFieldLabel(missingKey),
            severity: "warning",
            message: "Coordinate pair is incomplete",
            detail: "Only one half of the latitude/longitude pair is present.",
          });
        }
      }

      const altitudeColumn = state.mapping.altitude;
      const altitudeValue = valueAt(row, "altitude");
      if (altitudeColumn && !isEmpty(altitudeValue)) {
        const analysis = altitudeAnalysis(altitudeValue);
        if (analysis.valid && String(altitudeValue).trim() !== analysis.formatted) {
          addIssue({
            rowIndex,
            excelRow,
            column: altitudeColumn,
            field: "Altitude",
            severity: "formatting",
            message: "Altitude can be converted to metres",
            detail:
              "Removed the unit and thousands separators while preserving the numeric value.",
            original: altitudeValue,
            proposed: analysis.formatted,
            safe: true,
          });
        } else if (!analysis.valid) {
          addIssue({
            rowIndex,
            excelRow,
            column: altitudeColumn,
            field: "Altitude",
            severity: "warning",
            message: "Altitude is not numeric",
            detail: analysis.reason,
            original: altitudeValue,
          });
        }
      }

      ["sex", "lifeStage", "method"].forEach((key) => {
        const column = state.mapping[key];
        const original = valueAt(row, key);
        if (!column || isEmpty(original)) return;
        const lookup = CONTROLLED_TERMS[key];
        const normalized = normalizedText(original).replace(/\s+/g, " ");
        const proposed = lookup[normalized];
        if (proposed && cleanWhitespace(String(original)) !== proposed) {
          addIssue({
            rowIndex,
            excelRow,
            column,
            field: mappedFieldLabel(key),
            severity: "warning",
            message: `${mappedFieldLabel(key)} can be standardized`,
            detail: "This biological wording change requires your confirmation.",
            original,
            proposed,
            confirm: true,
          });
        }
      });
    });

    const idColumn = state.mapping.specimenId;
    if (idColumn) {
      const occurrences = new Map();
      state.rows.forEach((row, index) => {
        const id = cleanWhitespace(String(row[idColumn] || ""));
        if (!id) return;
        const key = id.toLowerCase();
        if (!occurrences.has(key)) occurrences.set(key, []);
        occurrences.get(key).push(index);
      });
      occurrences.forEach((indices, id) => {
        if (indices.length < 2) return;
        indices.forEach((rowIndex) => {
          addIssue({
            rowIndex,
            excelRow: state.excelRows[rowIndex],
            column: idColumn,
            field: "Specimen ID",
            severity: "critical",
            message: "Duplicate specimen ID",
            detail: `${indices.length} records use “${id}”. Compare them before assigning a new ID.`,
            original: state.rows[rowIndex][idColumn],
          });
        });
      });
    }
  }

  function issueCounts() {
    return state.issues.reduce(
      (counts, issue) => {
        if (issue.ignored) {
          counts.ignored += 1;
          return counts;
        }
        counts.total += 1;
        counts[issue.severity] = (counts[issue.severity] || 0) + 1;
        if (issue.safe) counts.safe += 1;
        if (issue.confirm) counts.confirm += 1;
        return counts;
      },
      {
        total: 0,
        ignored: 0,
        critical: 0,
        warning: 0,
        formatting: 0,
        safe: 0,
        confirm: 0,
      },
    );
  }

  function problemTypeCount(issues) {
    return new Set(issues.map((issue) => issue.groupKey)).size;
  }

  function severityIcon(issue) {
    if (issue.severity === "critical")
      return `<span class="severity-icon icon-critical">!</span>`;
    if (issue.severity === "warning")
      return `<span class="severity-icon icon-warning">△</span>`;
    return `<span class="severity-icon icon-safe">✓</span>`;
  }

  function visibleIssues() {
    const priority = { critical: 0, warning: 1, formatting: 2 };
    return state.issues
      .filter(
        (issue) =>
          (state.issueStatus === "ignored" ? issue.ignored : !issue.ignored) &&
          (state.filter === "all" || issue.severity === state.filter),
      )
      .slice()
      .sort((a, b) => {
        const severityDifference =
          priority[a.severity] - priority[b.severity];
        if (severityDifference) return severityDifference;
        return (a.excelRow || 0) - (b.excelRow || 0);
      });
  }

  function compactNumberRanges(rowNumbers) {
    const rows = [...new Set(rowNumbers.filter(Number.isFinite))].sort(
      (a, b) => a - b,
    );
    if (!rows.length) return "";

    const ranges = [];
    let start = rows[0];
    let end = rows[0];
    for (let index = 1; index <= rows.length; index += 1) {
      const value = rows[index];
      if (value === end + 1) {
        end = value;
        continue;
      }
      ranges.push(start === end ? `${start}` : `${start}–${end}`);
      start = value;
      end = value;
    }
    return ranges.join(", ");
  }

  function compactRowRanges(rowNumbers) {
    const ranges = compactNumberRanges(rowNumbers);
    return ranges ? `Rows ${ranges}` : "Table-level issue";
  }

  function proposalPairs(issues) {
    const pairs = new Map();
    issues.forEach((issue) => {
      if (issue.proposed == null) return;
      const key = JSON.stringify([
        displayValue(issue.original),
        displayValue(issue.proposed),
      ]);
      if (!pairs.has(key)) {
        pairs.set(key, {
          original: displayValue(issue.original),
          proposed: displayValue(issue.proposed),
          count: 0,
        });
      }
      pairs.get(key).count += 1;
    });
    return [...pairs.values()].sort((a, b) => b.count - a.count);
  }

  function canBulkFill(issues) {
    if (state.issueStatus === "ignored") return false;
    const columns = new Set(
      issues
        .filter((issue) => issue.rowIndex != null && issue.column)
        .map((issue) => issue.column),
    );
    if (columns.size !== 1) return false;
    const column = [...columns][0];
    return issues.some(
      (issue) =>
        issue.rowIndex != null &&
        state.rows[issue.rowIndex] &&
        isEmpty(state.rows[issue.rowIndex][column]),
    );
  }

  function emptyIssuesMarkup() {
    const ignoredCopy =
      state.issueStatus === "ignored"
        ? "Nothing has been ignored in this category."
        : "Try another filter or export the cleaned table.";
    return `<div class="empty-state">
      <span class="empty-state-icon">✓</span>
      <strong>No issues in this category</strong>
      <span>${ignoredCopy}</span>
    </div>`;
  }

  function issueRowsMarkup(issues = visibleIssues()) {
    if (!issues.length) {
      return emptyIssuesMarkup();
    }

    return issues
      .slice(0, 400)
      .map(
        (issue) => `
        <article class="issue-row ${issue.ignored ? "is-ignored" : ""}">
          ${severityIcon(issue)}
          <span class="row-number">${issue.excelRow ? `Row ${issue.excelRow}` : "Table"}</span>
          <div class="issue-copy">
            <strong>${escapeHtml(issue.message)}</strong>
            <p>${escapeHtml(issue.field)} · ${escapeHtml(issue.detail)}</p>
            ${
              issue.proposed !== null
                ? `<div class="value-change">
                    <span class="value-chip">${escapeHtml(displayValue(issue.original))}</span>
                    <span>→</span>
                    <span class="value-chip">${escapeHtml(displayValue(issue.proposed))}</span>
                  </div>`
                : ""
            }
          </div>
          <div class="issue-action">
            ${
              !issue.ignored && (issue.safe || issue.confirm)
                ? `<button class="small-button" type="button" data-fix-issue="${issue.id}">${
                    issue.safe ? "Apply fix" : "Use suggestion"
                  }</button>`
                : ""
            }
            ${
              issue.rowIndex != null && issue.column
                ? `<button class="small-button edit-cell-button" type="button" data-edit-issue="${issue.id}">Edit cell</button>`
                : ""
            }
            <button class="small-button ${
              issue.ignored ? "restore-button" : "ignore-button"
            }" type="button" ${
              issue.ignored
                ? `data-restore-issue="${issue.id}">Restore`
                : `data-ignore-issue="${issue.id}">Ignore`
            }</button>
          </div>
        </article>`,
      )
      .join("");
  }

  function issueGroupsMarkup(issues = visibleIssues()) {
    if (!issues.length) {
      return emptyIssuesMarkup();
    }

    const priority = { critical: 0, warning: 1, formatting: 2 };
    const groups = new Map();
    issues.forEach((issue) => {
      if (!groups.has(issue.groupKey)) {
        groups.set(issue.groupKey, {
          key: issue.groupKey,
          severity: issue.severity,
          field: issue.field,
          message: issue.message,
          detail: issue.detail,
          issues: [],
        });
      }
      groups.get(issue.groupKey).issues.push(issue);
    });

    return [...groups.values()]
      .sort((a, b) => {
        const severityDifference =
          priority[a.severity] - priority[b.severity];
        if (severityDifference) return severityDifference;
        const countDifference = b.issues.length - a.issues.length;
        if (countDifference) return countDifference;
        return a.message.localeCompare(b.message);
      })
      .map((group) => {
        const pairs = proposalPairs(group.issues);
        const actionable = group.issues.filter(
          (issue) =>
            !issue.ignored &&
            issue.rowIndex != null &&
            issue.column &&
            issue.proposed != null,
        );
        const firstEditable = group.issues.find(
          (issue) => issue.rowIndex != null && issue.column,
        );
        const groupKey = escapeHtml(group.key);
        const rows = compactRowRanges(
          group.issues.map((issue) => issue.excelRow),
        );
        const actionLabel = actionable.length
          ? actionable.every((issue) => issue.safe)
            ? `Fix all ${actionable.length}`
            : `Apply ${actionable.length} suggestions`
          : "";

        return `
          <details class="issue-group severity-${group.severity} ${
            state.issueStatus === "ignored" ? "is-ignored" : ""
          }" open>
            <summary>
              ${severityIcon(group.issues[0])}
              <span class="group-heading">
                <strong>${escapeHtml(group.message)}</strong>
                <span>${escapeHtml(group.field)} · ${escapeHtml(rows)}</span>
              </span>
              <span class="group-count">${group.issues.length}</span>
              <span class="group-chevron" aria-hidden="true">⌄</span>
            </summary>
            <div class="group-body">
              <p>${escapeHtml(group.detail)}</p>
              ${
                pairs.length
                  ? `<div class="proposal-summary">
                      ${pairs
                        .slice(0, 8)
                        .map(
                          (pair) => `<div class="proposal-row">
                            <span class="value-chip">${escapeHtml(pair.original)}</span>
                            <span>→</span>
                            <span class="value-chip">${escapeHtml(pair.proposed)}</span>
                            <strong>${pair.count}</strong>
                          </div>`,
                        )
                        .join("")}
                      ${
                        pairs.length > 8
                          ? `<span class="more-proposals">+${pairs.length - 8} more value pairs</span>`
                          : ""
                      }
                    </div>`
                  : ""
              }
              <div class="group-actions">
                ${
                  actionable.length
                    ? `<button class="primary-small-button" type="button" data-fix-group="${groupKey}">${actionLabel}</button>`
                    : ""
                }
                ${
                  canBulkFill(group.issues)
                    ? `<button class="primary-small-button fill-button" type="button" data-fill-group="${groupKey}">Fill rows…</button>`
                    : ""
                }
                ${
                  firstEditable
                    ? `<button class="small-button edit-cell-button" type="button" data-edit-issue="${firstEditable.id}">Edit first row</button>`
                    : ""
                }
                <button class="small-button ${
                  state.issueStatus === "ignored"
                    ? "restore-button"
                    : "ignore-button"
                }" type="button" ${
                  state.issueStatus === "ignored"
                    ? `data-restore-group="${groupKey}">Restore all ${group.issues.length}`
                    : `data-ignore-group="${groupKey}">Ignore all ${group.issues.length}`
                }</button>
              </div>
            </div>
          </details>`;
      })
      .join("");
  }

  function issuesMarkup() {
    return state.issueView === "rows"
      ? issueRowsMarkup()
      : issueGroupsMarkup();
  }

  function issueById(issueId) {
    return state.issues.find((issue) => issue.id === issueId);
  }

  function issueGroup(groupKey, includeIgnored = true) {
    return state.issues.filter(
      (issue) =>
        issue.groupKey === groupKey && (includeIgnored || !issue.ignored),
    );
  }

  function issueListHeading() {
    const relevant = visibleIssues();
    return `${relevant.length} ${
      relevant.length === 1 ? "issue" : "issues"
    } · ${problemTypeCount(relevant)} problem ${
      problemTypeCount(relevant) === 1 ? "type" : "types"
    }`;
  }

  function issueCellMap() {
    const map = new Map();
    state.issues.forEach((issue) => {
      if (issue.ignored) return;
      if (issue.rowIndex == null || !issue.column) return;
      const key = `${issue.rowIndex}|${issue.column}`;
      const current = map.get(key);
      if (!current || issue.severity === "critical") map.set(key, issue.severity);
    });
    return map;
  }

  function previewMarkup() {
    const isOriginal = state.previewMode === "original";
    const rows = isOriginal ? state.originalRows : state.rows;
    const columns = state.columns;
    const pageCount = Math.max(
      1,
      Math.ceil(rows.length / state.previewRowsPerPage),
    );
    if (state.previewPage >= pageCount) state.previewPage = pageCount - 1;
    const start = state.previewPage * state.previewRowsPerPage;
    const end = Math.min(start + state.previewRowsPerPage, rows.length);
    const problems = issueCellMap();
    return `
      <div class="preview-table-wrap">
        <table class="data-table preview-table">
          <thead><tr><th>Excel row</th>${columns
            .map((column) => `<th>${escapeHtml(column)}</th>`)
            .join("")}</tr></thead>
          <tbody>
            ${rows
              .slice(start, end)
              .map(
                (row, visibleIndex) => {
                  const rowIndex = start + visibleIndex;
                  return `<tr><td class="excel-row-cell">${state.excelRows[rowIndex]}</td>${columns
                    .map((column, columnIndex) => {
                      const severity = problems.get(`${rowIndex}|${column}`);
                      const className =
                        severity === "critical"
                          ? "has-critical"
                          : severity
                            ? "has-issue"
                            : "";
                      if (isOriginal) {
                        return `<td class="${className}" title="${escapeHtml(
                          row[column],
                        )}">${escapeHtml(row[column])}</td>`;
                      }
                      return `<td class="${className} is-editable">
                        <input
                          class="cell-editor"
                          type="text"
                          value="${escapeHtml(row[column])}"
                          data-edit-cell="${rowIndex}:${columnIndex}"
                          data-row-index="${rowIndex}"
                          data-column-index="${columnIndex}"
                          data-original-value="${escapeHtml(row[column])}"
                          aria-label="Edit row ${state.excelRows[rowIndex]}, ${escapeHtml(column)}"
                        />
                      </td>`;
                    })
                    .join("")}</tr>`;
                },
              )
              .join("")}
          </tbody>
        </table>
      </div>
      <div class="table-pager">
        <button class="small-button" id="preview-previous" type="button" ${
          state.previewPage === 0 ? "disabled" : ""
        }>← Previous</button>
        <span>Rows ${rows.length ? start + 1 : 0}–${end} of ${rows.length} · Page ${
          state.previewPage + 1
        } of ${pageCount}</span>
        <button class="small-button" id="preview-next" type="button" ${
          state.previewPage >= pageCount - 1 ? "disabled" : ""
        }>Next →</button>
      </div>`;
  }

  function parseExcelRowExpression(expression) {
    const text = String(expression || "").trim();
    if (!text) {
      return { rowIndices: [], error: "Enter at least one Excel row number." };
    }

    const parts = text
      .split(/[,;]+/)
      .map((part) => part.trim())
      .filter(Boolean);
    if (!parts.length) {
      return { rowIndices: [], error: "Enter at least one Excel row number." };
    }

    const selected = new Set();
    for (const part of parts) {
      const match = part.match(/^(\d+)(?:\s*[-–—]\s*(\d+))?$/);
      if (!match) {
        return {
          rowIndices: [],
          error: `“${part}” is not a row number or range.`,
        };
      }
      const first = Number(match[1]);
      const second = match[2] ? Number(match[2]) : first;
      const start = Math.min(first, second);
      const end = Math.max(first, second);
      state.excelRows.forEach((excelRow, rowIndex) => {
        if (excelRow >= start && excelRow <= end) selected.add(rowIndex);
      });
    }

    if (!selected.size) {
      return {
        rowIndices: [],
        error: "None of those Excel rows are present in the imported table.",
      };
    }
    return { rowIndices: [...selected].sort((a, b) => a - b), error: "" };
  }

  function bulkFillStats(expression, column) {
    const parsed = parseExcelRowExpression(expression);
    if (parsed.error) {
      return {
        rowIndices: [],
        emptyRowIndices: [],
        protectedCount: 0,
        error: parsed.error,
      };
    }
    const emptyRowIndices = parsed.rowIndices.filter(
      (rowIndex) =>
        state.rows[rowIndex] && isEmpty(state.rows[rowIndex][column]),
    );
    return {
      rowIndices: parsed.rowIndices,
      emptyRowIndices,
      protectedCount: parsed.rowIndices.length - emptyRowIndices.length,
      error: "",
    };
  }

  function bulkFillPreviewCopy(expression, column) {
    const stats = bulkFillStats(expression, column);
    if (stats.error) return stats.error;
    return `${stats.emptyRowIndices.length} empty ${
      stats.emptyRowIndices.length === 1 ? "cell" : "cells"
    } will be filled. ${stats.protectedCount} existing ${
      stats.protectedCount === 1 ? "value is" : "values are"
    } protected.`;
  }

  function bulkFillModalMarkup() {
    if (!state.bulkFill) return "";
    const { field, column, rowExpression, value, error } = state.bulkFill;
    return `
      <div class="modal-backdrop" id="bulk-fill-backdrop">
        <section class="bulk-fill-modal" role="dialog" aria-modal="true" aria-labelledby="bulk-fill-title">
          <div class="modal-head">
            <div>
              <p class="eyebrow">Careful bulk edit</p>
              <h3 id="bulk-fill-title">Fill missing ${escapeHtml(field)}</h3>
            </div>
            <button class="modal-close" id="bulk-fill-close" type="button" aria-label="Close">×</button>
          </div>
          <form id="bulk-fill-form">
            <label class="modal-field">
              <span>Value to insert</span>
              <input id="bulk-fill-value" type="text" value="${escapeHtml(
                value,
              )}" autocomplete="off" required />
            </label>
            <label class="modal-field">
              <span>Excel rows</span>
              <input id="bulk-fill-rows" type="text" value="${escapeHtml(
                rowExpression,
              )}" autocomplete="off" required />
              <small>Examples: 125–138 or 125–138, 142, 148–156</small>
            </label>
            <label class="protected-option">
              <input type="checkbox" checked disabled />
              <span>Fill empty cells only. Existing values are never overwritten.</span>
            </label>
            <div class="bulk-fill-preview ${
              error ? "has-error" : ""
            }" id="bulk-fill-preview">${
              error
                ? escapeHtml(error)
                : escapeHtml(bulkFillPreviewCopy(rowExpression, column))
            }</div>
            <div class="modal-actions">
              <button class="secondary-button" id="bulk-fill-cancel" type="button">Cancel</button>
              <button class="primary-button" type="submit">Apply bulk fill</button>
            </div>
          </form>
        </section>
      </div>`;
  }

  function renderDiagnosis() {
    const counts = issueCounts();
    const activeIndex = state.exported ? 3 : counts.safe || counts.confirm ? 2 : 3;
    const unresolved = counts.critical + counts.warning;
    const statusIssues = state.issues.filter((issue) =>
      state.issueStatus === "ignored" ? issue.ignored : !issue.ignored,
    );
    const statusCounts = statusIssues.reduce(
      (result, issue) => {
        result[issue.severity] += 1;
        return result;
      },
      { critical: 0, warning: 0, formatting: 0 },
    );
    const openProblemTypes = problemTypeCount(
      state.issues.filter((issue) => !issue.ignored),
    );

    shell(
      `
      <main class="screen workspace">
        <div class="workspace-head">
          <div>
            <p class="eyebrow">Diagnosis complete</p>
            <h2>${state.rows.length} records examined.</h2>
            <p>Formatting fixes are reversible. Scientific or ambiguous changes always wait for your confirmation.</p>
          </div>
          <span class="file-pill" title="${escapeHtml(state.fileName)}">▣ ${escapeHtml(
            state.fileName,
          )}</span>
        </div>

        ${
          state.exported
            ? `<div class="info-banner"><span aria-hidden="true">✓</span><span>Your export was prepared. The first worksheet is a flat table ready to reopen in EntoLabel.</span></div>`
            : ""
        }

        <section class="summary-grid" aria-label="Diagnosis summary">
          <article class="summary-card is-overview">
            <span class="summary-label">Table health</span>
            <span class="summary-value">${unresolved === 0 ? "Ready" : "Needs review"}</span>
            <span class="summary-detail">${state.rows.length} records · ${state.columns.length} columns · ${openProblemTypes} problem types · ${state.changes.length} accepted changes</span>
          </article>
          <article class="summary-card is-critical">
            <span class="summary-label">Critical</span>
            <span class="summary-value">${counts.critical}</span>
            <span class="summary-detail">Missing, invalid, or duplicate core data</span>
          </article>
          <article class="summary-card is-warning">
            <span class="summary-label">Review</span>
            <span class="summary-value">${counts.warning}</span>
            <span class="summary-detail">${counts.confirm} suggestions need confirmation</span>
          </article>
          <article class="summary-card">
            <span class="summary-label">Formatting</span>
            <span class="summary-value">${counts.formatting}</span>
            <span class="summary-detail">Machine-readable consistency issues</span>
          </article>
          <article class="summary-card is-safe">
            <span class="summary-label">Safe fixes</span>
            <span class="summary-value">${counts.safe}</span>
            <span class="summary-detail">Can be applied without interpretation</span>
          </article>
        </section>

        <div class="diagnosis-toolbar">
          <div class="toolbar-stack">
            <div class="toolbar-group" aria-label="Issue status">
              <button class="filter-button ${
                state.issueStatus === "open" ? "is-active" : ""
              }" type="button" data-issue-status="open">Open ${counts.total}</button>
              <button class="filter-button ignored-filter ${
                state.issueStatus === "ignored" ? "is-active" : ""
              }" type="button" data-issue-status="ignored">Ignored ${counts.ignored}</button>
            </div>
            <div class="toolbar-group" aria-label="Issue severity filters">
              ${[
                ["all", `All ${statusIssues.length}`],
                ["critical", `Critical ${statusCounts.critical}`],
                ["warning", `Review ${statusCounts.warning}`],
                ["formatting", `Formatting ${statusCounts.formatting}`],
              ]
                .map(
                  ([key, label]) =>
                    `<button class="filter-button ${
                      state.filter === key ? "is-active" : ""
                    }" type="button" data-filter="${key}">${label}</button>`,
                )
                .join("")}
            </div>
          </div>
          <div class="toolbar-group">
            <button class="ghost-button" id="mapping-button" type="button">← Field mapping</button>
            ${
              state.issueStatus === "ignored" && counts.ignored
                ? `<button class="secondary-button" id="restore-all-ignored" type="button">Restore all ignored</button>`
                : ""
            }
            <button class="secondary-button" id="reset-button" type="button" ${
              state.changes.length ? "" : "disabled"
            }>Reset accepted changes</button>
            <button class="primary-button" id="apply-safe-button" type="button" ${
              counts.safe ? "" : "disabled"
            }>Fix ${counts.safe} safe issues</button>
          </div>
        </div>

        <div class="diagnosis-layout">
          <section class="issue-panel">
            <div class="panel-head">
              <div>
                <h3>Check-up notes</h3>
                <p>${issueListHeading()} · Excel row numbers are preserved.</p>
              </div>
              <div class="view-switch" aria-label="Issue view">
                <button class="tab-button ${
                  state.issueView === "grouped" ? "is-active" : ""
                }" type="button" data-issue-view="grouped">Grouped</button>
                <button class="tab-button ${
                  state.issueView === "rows" ? "is-active" : ""
                }" type="button" data-issue-view="rows">By row</button>
              </div>
            </div>
            <div class="issue-list">${issuesMarkup()}</div>
          </section>

          <aside class="preview-panel">
            <div class="panel-head">
              <div>
                <h3>${state.previewMode === "current" ? "Edit table" : "Original table"}</h3>
                <p>${
                  state.previewMode === "current"
                    ? "Click any cell, type the correction, and press Enter. All columns are available."
                    : "Read-only copy of the imported data."
                }</p>
              </div>
              <div class="preview-tabs">
                <button class="tab-button ${
                  state.previewMode === "current" ? "is-active" : ""
                }" type="button" data-preview="current">Edit current</button>
                <button class="tab-button ${
                  state.previewMode === "original" ? "is-active" : ""
                }" type="button" data-preview="original">Original</button>
              </div>
            </div>
            ${previewMarkup()}
            <div class="export-box">
              <h4>EntoLabel-ready export</h4>
              <p>The Excel file contains the cleaned flat table, unresolved Issues, and a complete Change log.</p>
              <div class="export-buttons">
                <button class="primary-button" id="export-xlsx" type="button">Download Excel</button>
                <button class="secondary-button" id="export-csv" type="button">Download CSV</button>
              </div>
            </div>
          </aside>
        </div>
        <p class="footer-note">EntoData Doctor v0.3 · no account · no upload · source columns are preserved</p>
        ${bulkFillModalMarkup()}
      </main>`,
      activeIndex,
    );

    document.querySelectorAll("[data-filter]").forEach((button) => {
      button.addEventListener("click", () => {
        state.filter = button.dataset.filter;
        renderDiagnosis();
      });
    });

    document.querySelectorAll("[data-issue-status]").forEach((button) => {
      button.addEventListener("click", () => {
        state.issueStatus = button.dataset.issueStatus;
        state.filter = "all";
        renderDiagnosis();
      });
    });

    document.querySelectorAll("[data-issue-view]").forEach((button) => {
      button.addEventListener("click", () => {
        state.issueView = button.dataset.issueView;
        renderDiagnosis();
      });
    });

    document.querySelectorAll("[data-fix-issue]").forEach((button) => {
      button.addEventListener("click", () =>
        applyIssue(button.dataset.fixIssue),
      );
    });

    document.querySelectorAll("[data-fix-group]").forEach((button) => {
      button.addEventListener("click", () =>
        applyIssueGroup(button.dataset.fixGroup),
      );
    });

    document.querySelectorAll("[data-edit-issue]").forEach((button) => {
      button.addEventListener("click", () =>
        focusIssueCell(button.dataset.editIssue),
      );
    });

    document.querySelectorAll("[data-ignore-issue]").forEach((button) => {
      button.addEventListener("click", () =>
        setIssueIgnored(button.dataset.ignoreIssue, true),
      );
    });

    document.querySelectorAll("[data-restore-issue]").forEach((button) => {
      button.addEventListener("click", () =>
        setIssueIgnored(button.dataset.restoreIssue, false),
      );
    });

    document.querySelectorAll("[data-ignore-group]").forEach((button) => {
      button.addEventListener("click", () =>
        setIssueGroupIgnored(button.dataset.ignoreGroup, true),
      );
    });

    document.querySelectorAll("[data-restore-group]").forEach((button) => {
      button.addEventListener("click", () =>
        setIssueGroupIgnored(button.dataset.restoreGroup, false),
      );
    });

    document.querySelectorAll("[data-fill-group]").forEach((button) => {
      button.addEventListener("click", () =>
        openBulkFill(button.dataset.fillGroup),
      );
    });

    document.querySelectorAll("[data-preview]").forEach((button) => {
      button.addEventListener("click", () => {
        state.previewMode = button.dataset.preview;
        renderDiagnosis();
      });
    });

    const previewWrap = document.querySelector(".preview-table-wrap");
    if (previewWrap) {
      previewWrap.scrollLeft = state.previewScrollLeft;
      previewWrap.scrollTop = state.previewScrollTop;
      previewWrap.addEventListener("scroll", () => {
        state.previewScrollLeft = previewWrap.scrollLeft;
        state.previewScrollTop = previewWrap.scrollTop;
      });
      previewWrap.addEventListener("keydown", (event) => {
        const editor = event.target.closest(".cell-editor");
        if (!editor) return;
        if (event.key === "Enter") {
          event.preventDefault();
          editor.blur();
        } else if (event.key === "Escape") {
          event.preventDefault();
          editor.value = editor.dataset.originalValue;
          editor.blur();
        }
      });
      previewWrap.addEventListener("focusout", (event) => {
        const editor = event.target.closest(".cell-editor");
        if (!editor) return;
        const oldValue = editor.dataset.originalValue;
        const newValue = editor.value;
        if (oldValue === newValue) return;
        const rowIndex = Number(editor.dataset.rowIndex);
        const columnIndex = Number(editor.dataset.columnIndex);
        window.setTimeout(
          () => saveManualEdit(rowIndex, columnIndex, oldValue, newValue),
          0,
        );
      });
    }

    const previousButton = document.getElementById("preview-previous");
    const nextButton = document.getElementById("preview-next");
    if (previousButton) {
      previousButton.addEventListener("click", () => {
        state.previewPage = Math.max(0, state.previewPage - 1);
        state.previewScrollTop = 0;
        renderDiagnosis();
      });
    }
    if (nextButton) {
      nextButton.addEventListener("click", () => {
        state.previewPage += 1;
        state.previewScrollTop = 0;
        renderDiagnosis();
      });
    }

    const restoreAllIgnored = document.getElementById("restore-all-ignored");
    if (restoreAllIgnored) {
      restoreAllIgnored.addEventListener("click", restoreAllIgnoredIssues);
    }

    const bulkFillForm = document.getElementById("bulk-fill-form");
    if (bulkFillForm && state.bulkFill) {
      const valueInput = document.getElementById("bulk-fill-value");
      const rowsInput = document.getElementById("bulk-fill-rows");
      const preview = document.getElementById("bulk-fill-preview");
      const closeBulkFill = () => {
        state.bulkFill = null;
        renderDiagnosis();
      };
      const updateBulkPreview = () => {
        const copy = bulkFillPreviewCopy(
          rowsInput.value,
          state.bulkFill.column,
        );
        const hasError = Boolean(
          bulkFillStats(rowsInput.value, state.bulkFill.column).error,
        );
        preview.textContent = copy;
        preview.classList.toggle("has-error", hasError);
      };

      rowsInput.addEventListener("input", updateBulkPreview);
      document
        .getElementById("bulk-fill-close")
        .addEventListener("click", closeBulkFill);
      document
        .getElementById("bulk-fill-cancel")
        .addEventListener("click", closeBulkFill);
      document
        .getElementById("bulk-fill-backdrop")
        .addEventListener("click", (event) => {
          if (event.target.id === "bulk-fill-backdrop") closeBulkFill();
        });
      bulkFillForm.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          closeBulkFill();
        }
      });
      bulkFillForm.addEventListener("submit", (event) => {
        event.preventDefault();
        state.bulkFill.value = valueInput.value;
        state.bulkFill.rowExpression = rowsInput.value;
        applyBulkFill();
      });
      window.requestAnimationFrame(() => {
        valueInput.focus();
        valueInput.select();
      });
    }

    document.getElementById("apply-safe-button").addEventListener("click", applySafeFixes);
    document.getElementById("reset-button").addEventListener("click", resetChanges);
    document.getElementById("mapping-button").addEventListener("click", () => {
      state.screen = "setup";
      renderSetup();
    });
    document.getElementById("export-xlsx").addEventListener("click", exportExcel);
    document.getElementById("export-csv").addEventListener("click", exportCsv);
  }

  function recordChange(issue, oldValue, newValue) {
    state.changes.push({
      "Excel row": issue.excelRow || "",
      Field: issue.field,
      Column: issue.column,
      Original: oldValue,
      Cleaned: newValue,
      Reason: issue.message,
      "Change type": issue.safe ? "Safe automatic fix" : "User-confirmed suggestion",
    });
  }

  function setIssueIgnored(issueId, ignored) {
    const issue = issueById(issueId);
    if (!issue) return;
    if (ignored) state.ignoredIssueKeys.add(issue.key);
    else state.ignoredIssueKeys.delete(issue.key);
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function setIssueGroupIgnored(groupKey, ignored) {
    const issues = issueGroup(groupKey);
    issues.forEach((issue) => {
      if (ignored) state.ignoredIssueKeys.add(issue.key);
      else state.ignoredIssueKeys.delete(issue.key);
    });
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function restoreAllIgnoredIssues() {
    state.ignoredIssueKeys.clear();
    state.issueStatus = "open";
    state.filter = "all";
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function applyIssueGroup(groupKey) {
    const issues = issueGroup(groupKey, false).filter(
      (issue) =>
        issue.rowIndex != null &&
        issue.column &&
        issue.proposed != null,
    );
    issues.forEach((issue) => {
      const oldValue = state.rows[issue.rowIndex][issue.column];
      if (String(oldValue) === String(issue.proposed)) return;
      state.rows[issue.rowIndex][issue.column] = issue.proposed;
      recordChange(issue, oldValue, issue.proposed);
    });
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function openBulkFill(groupKey) {
    const issues = issueGroup(groupKey, false);
    const columns = new Set(
      issues
        .filter((issue) => issue.rowIndex != null && issue.column)
        .map((issue) => issue.column),
    );
    if (columns.size !== 1) return;
    const column = [...columns][0];
    const emptyIssues = issues.filter(
      (issue) =>
        issue.rowIndex != null &&
        state.rows[issue.rowIndex] &&
        isEmpty(state.rows[issue.rowIndex][column]),
    );
    if (!emptyIssues.length) return;

    state.bulkFill = {
      groupKey,
      field: emptyIssues[0].field,
      column,
      rowExpression: compactNumberRanges(
        emptyIssues.map((issue) => issue.excelRow),
      ),
      value: "",
      error: "",
    };
    renderDiagnosis();
  }

  function applyBulkFill() {
    if (!state.bulkFill) return;
    const value = cleanWhitespace(String(state.bulkFill.value || ""));
    if (!value) {
      state.bulkFill.error = "Enter the value that should fill the empty cells.";
      renderDiagnosis();
      return;
    }

    const stats = bulkFillStats(
      state.bulkFill.rowExpression,
      state.bulkFill.column,
    );
    if (stats.error) {
      state.bulkFill.error = stats.error;
      renderDiagnosis();
      return;
    }
    if (!stats.emptyRowIndices.length) {
      state.bulkFill.error =
        "Those rows contain no empty cells in this column.";
      renderDiagnosis();
      return;
    }

    const { field, column } = state.bulkFill;
    stats.emptyRowIndices.forEach((rowIndex) => {
      const oldValue = state.rows[rowIndex][column];
      state.rows[rowIndex][column] = value;
      state.changes.push({
        "Excel row": state.excelRows[rowIndex] || "",
        Field: field,
        Column: column,
        Original: oldValue,
        Cleaned: value,
        Reason: `Bulk fill missing ${field}`,
        "Change type": "Bulk fill in EntoData Doctor",
      });
    });

    state.bulkFill = null;
    state.issueStatus = "open";
    state.filter = "all";
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function saveManualEdit(rowIndex, columnIndex, oldValue, newValue) {
    const column = state.columns[columnIndex];
    if (!column || !state.rows[rowIndex]) return;
    const currentValue = state.rows[rowIndex][column];
    if (String(currentValue == null ? "" : currentValue) === newValue) return;

    const pageScroll = window.scrollY;
    state.rows[rowIndex][column] = newValue;
    state.changes.push({
      "Excel row": state.excelRows[rowIndex] || "",
      Field: column,
      Column: column,
      Original: oldValue,
      Cleaned: newValue,
      Reason: "Manual cell edit",
      "Change type": "Manual edit in EntoData Doctor",
    });
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
    window.requestAnimationFrame(() => {
      window.scrollTo({ top: pageScroll, behavior: "auto" });
    });
  }

  function focusIssueCell(issueId) {
    const issue = state.issues.find((item) => item.id === issueId);
    if (!issue || issue.rowIndex == null || !issue.column) return;
    const columnIndex = state.columns.indexOf(issue.column);
    if (columnIndex < 0) return;

    state.previewMode = "current";
    state.previewPage = Math.floor(issue.rowIndex / state.previewRowsPerPage);
    state.previewScrollLeft = 0;
    state.previewScrollTop = 0;
    renderDiagnosis();
    window.requestAnimationFrame(() => {
      const editor = document.querySelector(
        `[data-edit-cell="${issue.rowIndex}:${columnIndex}"]`,
      );
      if (!editor) return;
      editor.focus();
      editor.select();
      editor.scrollIntoView({
        block: "center",
        inline: "center",
        behavior: "smooth",
      });
    });
  }

  function applyIssue(issueId) {
    const issue = state.issues.find((item) => item.id === issueId);
    if (
      !issue ||
      issue.ignored ||
      issue.rowIndex == null ||
      !issue.column ||
      issue.proposed == null
    )
      return;
    const oldValue = state.rows[issue.rowIndex][issue.column];
    state.rows[issue.rowIndex][issue.column] = issue.proposed;
    recordChange(issue, oldValue, issue.proposed);
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function applySafeFixes() {
    const fixes = state.issues.filter(
      (issue) =>
        !issue.ignored &&
        issue.safe &&
        issue.rowIndex != null &&
        issue.column &&
        issue.proposed != null,
    );
    fixes.forEach((issue) => {
      const oldValue = state.rows[issue.rowIndex][issue.column];
      if (oldValue === issue.proposed) return;
      state.rows[issue.rowIndex][issue.column] = issue.proposed;
      recordChange(issue, oldValue, issue.proposed);
    });
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function resetChanges() {
    state.rows = deepClone(state.originalRows);
    state.changes = [];
    state.exported = false;
    runDiagnosis();
    renderDiagnosis();
  }

  function sanitizedBaseName() {
    return (
      state.fileName
        .replace(/\.(xlsx|xls|csv|tsv)$/i, "")
        .replace(/[^\p{L}\p{N}_-]+/gu, "_")
        .replace(/^_+|_+$/g, "") || "specimen_data"
    );
  }

  function exportExcel() {
    const workbook = window.XLSX.utils.book_new();
    const dataRows = state.rows.map((row) => {
      const ordered = {};
      state.columns.forEach((column) => {
        ordered[column] = row[column];
      });
      return ordered;
    });
    const dataSheet = window.XLSX.utils.json_to_sheet(dataRows, {
      header: state.columns,
    });
    const issueRows = state.issues.map((issue) => ({
      "Excel row": issue.excelRow || "",
      Status: issue.ignored ? "Ignored by user" : "Open",
      Severity: issue.severity,
      Field: issue.field,
      Column: issue.column,
      Issue: issue.message,
      Explanation: issue.detail,
      Original: issue.original,
      Suggestion: issue.proposed == null ? "" : issue.proposed,
      "Safe to apply": issue.safe ? "yes" : "no",
    }));
    const issueSheet = window.XLSX.utils.json_to_sheet(
      issueRows.length ? issueRows : [{ Status: "No unresolved issues" }],
    );
    const changeSheet = window.XLSX.utils.json_to_sheet(
      state.changes.length
        ? state.changes
        : [{ Status: "No changes were accepted" }],
    );

    window.XLSX.utils.book_append_sheet(workbook, dataSheet, "Cleaned Data");
    window.XLSX.utils.book_append_sheet(workbook, issueSheet, "Issues");
    window.XLSX.utils.book_append_sheet(workbook, changeSheet, "Change log");
    const bytes = window.XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
      compression: true,
    });
    downloadBlob(
      new Blob([bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      `${sanitizedBaseName()}_cleaned.xlsx`,
    );
    state.exported = true;
    renderDiagnosis();
  }

  function exportCsv() {
    const sheet = window.XLSX.utils.json_to_sheet(state.rows, {
      header: state.columns,
    });
    const csv = window.XLSX.utils.sheet_to_csv(sheet);
    const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
    downloadBlob(blob, `${sanitizedBaseName()}_cleaned.csv`);
    state.exported = true;
    renderDiagnosis();
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function resetApp() {
    Object.assign(state, {
      screen: "upload",
      fileName: "",
      workbook: null,
      sheetName: "",
      matrix: [],
      headerRow: 0,
      rawHeaders: [],
      columns: [],
      rows: [],
      originalRows: [],
      excelRows: [],
      mapping: {},
      issues: [],
      changes: [],
      filter: "all",
      issueView: "grouped",
      issueStatus: "open",
      ignoredIssueKeys: new Set(),
      bulkFill: null,
      previewMode: "current",
      previewPage: 0,
      previewRowsPerPage: 20,
      previewScrollLeft: 0,
      previewScrollTop: 0,
      error: "",
      exported: false,
    });
    renderUpload();
  }

  renderUpload();
})();
