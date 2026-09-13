(function attachVsmExcelExport(root) {
  "use strict";

  const MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const STYLE = Object.freeze({
    normal: 0,
    title: 1,
    section: 2,
    label: 3,
    inputText: 4,
    inputNumber: 5,
    inputInteger: 6,
    inputPercent: 7,
    unit: 8,
    note: 9,
    outputText: 10,
    outputNumber: 11,
    outputInteger: 12,
    outputPercent: 13,
    tableHeader: 14,
    rowIndex: 15,
    dataText: 16,
    dataNumber: 17,
    dataInteger: 18,
    dataPercent: 19,
    calcText: 20,
    calcNumber: 21,
    calcInteger: 22,
    calcPercent: 23,
    statusNormal: 24,
    statusWarning: 25,
    statusDanger: 26,
    banner: 27,
    vsmHeader: 28,
    vsmLabel: 29,
    vsmText: 30,
    vsmNumber: 31,
    vsmInteger: 32,
    vsmPercent: 33,
    connector: 34,
    wip: 35,
    timelineLabel: 36,
    timelineValue: 37,
    subtitle: 38,
    dangerNumber: 39,
    dangerPercent: 40,
    summaryLabel: 41,
    summaryText: 42,
    summaryNumber: 43,
    summaryInteger: 44,
    summaryPercent: 45,
    divider: 46,
    warningPercent: 47,
  });

  function escapeXml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function columnName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name;
  }

  function cellAddress(row, column) {
    return `${columnName(column)}${row}`;
  }

  function finiteNumber(value) {
    const number = Number(value);
    return Number.isFinite(number) ? number : null;
  }

  function displayText(value, fallback = "—") {
    if (value === null || value === undefined || value === "") return fallback;
    return String(value);
  }

  function statusStyle(status) {
    if (status === "產能不足" || status === "超載") return STYLE.statusDanger;
    if (status === "接近滿載" || status === "注意") return STYLE.statusWarning;
    return STYLE.statusNormal;
  }

  function formatExportTime(date) {
    try {
      return date.toLocaleString("zh-TW", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false,
      });
    } catch {
      return date.toISOString().replace("T", " ").slice(0, 19);
    }
  }

  class SheetBuilder {
    constructor(name) {
      this.name = name;
      this.cells = new Map();
      this.merges = [];
      this.columns = [];
      this.rowHeights = new Map();
      this.maxRow = 1;
      this.maxColumn = 0;
      this.freeze = null;
      this.autoFilter = null;
      this.orientation = "landscape";
      this.fitToWidth = 1;
      this.fitToHeight = 0;
    }

    set(row, column, value, style = STYLE.normal) {
      const key = `${row}:${column}`;
      this.cells.set(key, { row, column, value, style });
      this.maxRow = Math.max(this.maxRow, row);
      this.maxColumn = Math.max(this.maxColumn, column);
      return this;
    }

    merge(rowStart, columnStart, rowEnd, columnEnd, value, style = STYLE.normal) {
      for (let row = rowStart; row <= rowEnd; row += 1) {
        for (let column = columnStart; column <= columnEnd; column += 1) {
          this.set(row, column, null, style);
        }
      }
      this.set(rowStart, columnStart, value, style);
      this.merges.push(`${cellAddress(rowStart, columnStart)}:${cellAddress(rowEnd, columnEnd)}`);
      return this;
    }

    setColumn(columnStart, columnEnd, width) {
      this.columns.push({ min: columnStart + 1, max: columnEnd + 1, width });
      this.maxColumn = Math.max(this.maxColumn, columnEnd);
      return this;
    }

    setRowHeight(row, height) {
      this.rowHeights.set(row, height);
      this.maxRow = Math.max(this.maxRow, row);
      return this;
    }

    toXml() {
      const rows = new Map();
      for (const cell of this.cells.values()) {
        if (!rows.has(cell.row)) rows.set(cell.row, []);
        rows.get(cell.row).push(cell);
      }

      const rowXml = [...rows.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([rowNumber, rowCells]) => {
          const height = this.rowHeights.get(rowNumber);
          const heightAttributes = height ? ` ht="${height}" customHeight="1"` : "";
          const cells = rowCells
            .sort((a, b) => a.column - b.column)
            .map((cell) => this.cellXml(cell))
            .join("");
          return `<row r="${rowNumber}"${heightAttributes}>${cells}</row>`;
        })
        .join("");

      const columnsXml = this.columns.length
        ? `<cols>${this.columns.map((column) => `<col min="${column.min}" max="${column.max}" width="${column.width}" customWidth="1"/>`).join("")}</cols>`
        : "";
      const mergesXml = this.merges.length
        ? `<mergeCells count="${this.merges.length}">${this.merges.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>`
        : "";
      const paneXml = this.freeze
        ? `<pane xSplit="${this.freeze.columns}" ySplit="${this.freeze.rows}" topLeftCell="${cellAddress(this.freeze.rows + 1, this.freeze.columns)}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="${cellAddress(this.freeze.rows + 1, this.freeze.columns)}" sqref="${cellAddress(this.freeze.rows + 1, this.freeze.columns)}"/>`
        : `<selection activeCell="A1" sqref="A1"/>`;
      const autoFilterXml = this.autoFilter ? `<autoFilter ref="${this.autoFilter}"/>` : "";
      const dimension = `A1:${cellAddress(this.maxRow, this.maxColumn)}`;

      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
  <dimension ref="${dimension}"/>
  <sheetViews><sheetView showGridLines="0" workbookViewId="0">${paneXml}</sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  ${columnsXml}
  <sheetData>${rowXml}</sheetData>
  ${autoFilterXml}
  ${mergesXml}
  <pageMargins left="0.25" right="0.25" top="0.4" bottom="0.4" header="0.15" footer="0.15"/>
  <pageSetup paperSize="9" orientation="${this.orientation}" fitToWidth="${this.fitToWidth}" fitToHeight="${this.fitToHeight}" horizontalDpi="300" verticalDpi="300"/>
</worksheet>`;
    }

    cellXml(cell) {
      const ref = cellAddress(cell.row, cell.column);
      const style = Number.isInteger(cell.style) ? ` s="${cell.style}"` : "";
      if (cell.value === null || cell.value === undefined || cell.value === "") return `<c r="${ref}"${style}/>`;
      if (typeof cell.value === "number" && Number.isFinite(cell.value)) return `<c r="${ref}"${style}><v>${cell.value}</v></c>`;
      if (typeof cell.value === "boolean") return `<c r="${ref}"${style} t="b"><v>${cell.value ? 1 : 0}</v></c>`;
      const text = escapeXml(cell.value);
      return `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${text}</t></is></c>`;
    }
  }

  function createInstructionsSheet() {
    const sheet = new SheetBuilder("操作說明");
    sheet.setColumn(0, 0, 3).setColumn(1, 1, 15).setColumn(2, 5, 22);
    sheet.merge(1, 0, 2, 5, "VSM 自動產生工具｜匯出報表", STYLE.title);
    sheet.setRowHeight(1, 24).setRowHeight(2, 16);
    sheet.merge(4, 1, 4, 5, "匯出 Excel 使用說明", STYLE.section);
    sheet.setRowHeight(4, 24);

    const steps = [
      ["1", "本活頁簿由網頁當下資料產生，黃色區為原始輸入，淺色區為計算結果。"],
      ["2", "「參數輸入」整理月需求、工作天、班別、工時、OEE 與主要結果。"],
      ["3", "「工站資料」保留各站 C/T、MCT、OCT、換線時間、WIP、產能與負荷率。"],
      ["4", "「VSM圖」以每列最多 4 站呈現流程，超過 4 站時自動接續到下一列，最多 20 站。"],
      ["5", "工站 C/T 超過需求 Takt 時，C/T 與負荷率會以紅色標示。"],
    ];
    steps.forEach(([step, description], index) => {
      const row = 5 + index;
      sheet.set(row, 1, step, STYLE.rowIndex);
      sheet.merge(row, 2, row, 5, description, STYLE.note);
      sheet.setRowHeight(row, 24);
    });

    sheet.merge(12, 1, 12, 5, "計算邏輯", STYLE.section);
    sheet.setRowHeight(12, 24);
    const logicRows = [
      ["需求 Takt", "每日有效可用時間 ÷ 日需求"],
      ["各站月產能", "月有效可用時間 ÷ C/T × 良率 × 並行機台數"],
      ["負荷率", "月需求 ÷ 各站月產能"],
      ["目標 WIP", "向上取整（緩衝分鐘 ÷ 需求 Takt）"],
      ["現況等待", "現況 WIP × 需求 Takt ÷ 60"],
      ["Lead Time", "總現況等待時間 + 全站 C/T 合計 ÷ 60"],
    ];
    logicRows.forEach(([label, description], index) => {
      const row = 13 + index;
      sheet.set(row, 1, label, STYLE.summaryLabel);
      sheet.merge(row, 2, row, 5, description, STYLE.calcText);
    });
    sheet.merge(20, 1, 20, 5, "現場改善與決策仍請以實際量測資料校正。", STYLE.note);
    sheet.setRowHeight(20, 22);
    sheet.orientation = "portrait";
    return sheet;
  }

  function createParametersSheet(state, calc, exportedAt) {
    const sheet = new SheetBuilder("參數輸入");
    sheet.setColumn(0, 0, 3).setColumn(1, 1, 20).setColumn(2, 2, 18).setColumn(3, 3, 14).setColumn(4, 4, 42);
    sheet.merge(1, 0, 2, 4, `VSM 基本參數與即時結果｜${displayText(state.lineName, "未命名產線")}`, STYLE.title);
    sheet.setRowHeight(1, 24).setRowHeight(2, 16);
    sheet.merge(4, 1, 4, 3, "基本輸入", STYLE.section);
    sheet.set(4, 4, "說明", STYLE.section);

    const inputs = [
      ["產線名稱", displayText(state.lineName, ""), "", "本次分析名稱", STYLE.inputText],
      ["月需求", finiteNumber(state.monthlyDemand), "pcs", "每月客戶需求數量", STYLE.inputInteger],
      ["工作天", finiteNumber(state.workDays), "天／月", "當月實際工作天數", STYLE.inputNumber],
      ["班別", finiteNumber(state.shifts), "班／日", "每日生產班數", STYLE.inputNumber],
      ["每班工時", finiteNumber(state.hoursPerShift), "hr／班", "每班排定工時", STYLE.inputNumber],
      ["每班休息", finiteNumber(state.breakMinutes), "min／班", "休息或用餐等非生產時間", STYLE.inputNumber],
      ["計畫稼動率／OEE", finiteNumber(state.availability) === null ? null : finiteNumber(state.availability) / 100, "%", "需求 Takt 的有效時間折減", STYLE.inputPercent],
      ["匯出時間", formatExportTime(exportedAt), "", "網頁匯出當下時間", STYLE.inputText],
    ];
    inputs.forEach(([label, value, unit, note, valueStyle], index) => {
      const row = 5 + index;
      sheet.set(row, 1, label, STYLE.label);
      sheet.set(row, 2, value, valueStyle);
      sheet.set(row, 3, unit, STYLE.unit);
      sheet.set(row, 4, note, STYLE.note);
    });

    sheet.merge(15, 1, 15, 3, "自動計算", STYLE.section);
    sheet.set(15, 4, "單位與說明", STYLE.section);
    const netPerShift = Number.isFinite(calc.netAvailableDay) && finiteNumber(state.shifts) > 0 ? calc.netAvailableDay / Number(state.shifts) : null;
    const demandRate = Number.isFinite(calc.takt) && calc.takt > 0 ? 60 / calc.takt : null;
    const overallStatus = calc.bottleneck?.load > 1 ? "產能不足" : calc.bottleneck?.load >= 0.85 ? "接近滿載" : calc.bottleneck ? "產能正常" : "資料不足";
    const outputs = [
      ["日需求", finiteNumber(calc.dailyDemand), "pcs／日", STYLE.outputNumber],
      ["每班有效時間", finiteNumber(netPerShift), "min／班", STYLE.outputNumber],
      ["每日有效可用時間", finiteNumber(calc.netAvailableDay), "min／日", STYLE.outputNumber],
      ["月有效可用時間", finiteNumber(calc.monthlyMinutes), "min／月", STYLE.outputNumber],
      ["需求 Takt", finiteNumber(calc.takt), "min／pcs", STYLE.outputNumber],
      ["需求速度", finiteNumber(demandRate), "pcs／hr", STYLE.outputNumber],
      ["瓶頸站", calc.bottleneck ? `${displayText(calc.bottleneck.id)}｜${displayText(calc.bottleneck.name, "")}` : "—", "", STYLE.outputText],
      ["瓶頸月產能", finiteNumber(calc.bottleneck?.capacity), "pcs", STYLE.outputInteger],
      ["瓶頸負荷率", finiteNumber(calc.bottleneck?.load), "%", STYLE.outputPercent],
      ["總現況 WIP", finiteNumber(calc.currentWip), "pcs", STYLE.outputInteger],
      ["總目標 WIP", finiteNumber(calc.targetWip), "pcs", STYLE.outputInteger],
      ["總 VA 時間", finiteNumber(calc.vaMinutes), "min", STYLE.outputNumber],
      ["總 NVA 等待", finiteNumber(calc.nvaHours), "hr", STYLE.outputNumber],
      ["現況 Lead Time", finiteNumber(calc.leadTime), "hr", STYLE.outputNumber],
      ["產能判定", overallStatus, "", statusStyle(overallStatus)],
    ];
    outputs.forEach(([label, value, unit, valueStyle], index) => {
      const row = 16 + index;
      sheet.set(row, 1, label, STYLE.label);
      sheet.set(row, 2, value, valueStyle);
      sheet.set(row, 3, unit, STYLE.unit);
      sheet.set(row, 4, label === "需求 Takt" ? "各站 C/T 應不高於此值" : "", STYLE.note);
    });
    sheet.orientation = "portrait";
    sheet.freeze = { rows: 4, columns: 1 };
    return sheet;
  }

  function createStationsSheet(calc) {
    const sheet = new SheetBuilder("工站資料");
    const headers = [
      "順序", "站別", "製程名稱", "C/T\n(min)", "MCT\n(min)", "OCTa\n(min)", "OCTb\n(min)",
      "C/O Changeover Time\n換線／切換時間 (min)", "人數", "並行機台", "現況 WIP\n(pcs)", "緩衝\n(min)",
      "良率", "月產能\n(pcs)", "負荷率", "目標 WIP\n(pcs)", "現況等待\n(hr)", "目標等待\n(hr)", "狀態",
    ];
    const widths = [7, 12, 20, 11, 11, 11, 11, 28, 9, 11, 13, 11, 10, 15, 11, 13, 14, 14, 13];
    widths.forEach((width, index) => sheet.setColumn(index, index, width));
    sheet.merge(1, 0, 2, headers.length - 1, `工站資料｜共 ${calc.stations.length} 站`, STYLE.title);
    headers.forEach((header, index) => sheet.set(4, index, header, STYLE.tableHeader));
    sheet.setRowHeight(1, 24).setRowHeight(2, 16).setRowHeight(4, 40);

    calc.stations.forEach((station, index) => {
      const row = 5 + index;
      const ct = finiteNumber(station.ct);
      const load = finiteNumber(station.load);
      const ctStyle = ct !== null && Number.isFinite(calc.takt) && ct > calc.takt ? STYLE.dangerNumber : STYLE.dataNumber;
      const loadStyle = load !== null && load > 1 ? STYLE.dangerPercent : load !== null && load >= 0.85 ? STYLE.warningPercent : STYLE.calcPercent;
      const values = [
        [index + 1, STYLE.rowIndex],
        [displayText(station.id, ""), STYLE.dataText],
        [displayText(station.name, ""), STYLE.dataText],
        [ct, ctStyle],
        [finiteNumber(station.mct), STYLE.dataNumber],
        [finiteNumber(station.octa), STYLE.dataNumber],
        [finiteNumber(station.octb), STYLE.dataNumber],
        [finiteNumber(station.octc), STYLE.dataNumber],
        [finiteNumber(station.people), STYLE.dataInteger],
        [finiteNumber(station.machines), STYLE.dataInteger],
        [finiteNumber(station.currentWip), STYLE.dataInteger],
        [finiteNumber(station.bufferMin), STYLE.dataNumber],
        [finiteNumber(station.yieldRate) === null ? null : finiteNumber(station.yieldRate) / 100, STYLE.dataPercent],
        [finiteNumber(station.capacity), STYLE.calcInteger],
        [load, loadStyle],
        [finiteNumber(station.targetWip), STYLE.calcInteger],
        [finiteNumber(station.currentWait), STYLE.calcNumber],
        [finiteNumber(station.targetWait), STYLE.calcNumber],
        [displayText(station.status, "資料不足"), statusStyle(station.status)],
      ];
      values.forEach(([value, style], column) => sheet.set(row, column, value, style));
      sheet.setRowHeight(row, 22);
    });
    sheet.autoFilter = `A4:S${4 + calc.stations.length}`;
    sheet.freeze = { rows: 4, columns: 3 };
    return sheet;
  }

  function addSummaryBlock(sheet, startColumn, title, rows) {
    sheet.merge(7, startColumn, 7, startColumn + 5, title, STYLE.section);
    rows.forEach(([label, value, style], index) => {
      const row = 8 + index;
      sheet.merge(row, startColumn, row, startColumn + 2, label, STYLE.summaryLabel);
      sheet.merge(row, startColumn + 3, row, startColumn + 5, value, style);
    });
  }

  function addVsmBlock(sheet, calc, stations, blockIndex) {
    const baseRow = 17 + blockIndex * 17;
    const firstStationNumber = blockIndex * 4 + 1;
    const lastStationNumber = firstStationNumber + stations.length - 1;
    sheet.merge(baseRow - 2, 1, baseRow - 2, 32, `現況 VSM｜第 ${firstStationNumber}–${lastStationNumber} 站`, STYLE.subtitle);

    stations.forEach((station, slot) => {
      const startColumn = 1 + slot * 8;
      const endColumn = startColumn + 5;
      const ct = finiteNumber(station.ct);
      const overTakt = ct !== null && Number.isFinite(calc.takt) && ct > calc.takt;
      sheet.merge(baseRow, startColumn, baseRow, endColumn, `${displayText(station.id)}｜${displayText(station.name, "未命名製程")}`, STYLE.vsmHeader);
      const details = [
        ["C/T", ct, overTakt ? STYLE.dangerNumber : STYLE.vsmNumber, "min"],
        ["MCT", finiteNumber(station.mct), STYLE.vsmNumber, "min"],
        ["OCTa", finiteNumber(station.octa), STYLE.vsmNumber, "min"],
        ["OCTb", finiteNumber(station.octb), STYLE.vsmNumber, "min"],
        ["C/O", finiteNumber(station.octc), STYLE.vsmNumber, "min"],
        ["人數／機台", `${displayText(station.people, "0")}／${displayText(station.machines, "0")}`, STYLE.vsmText, ""],
        ["良率", finiteNumber(station.yieldRate) === null ? null : finiteNumber(station.yieldRate) / 100, STYLE.vsmPercent, ""],
        ["月產能", finiteNumber(station.capacity), STYLE.vsmInteger, "pcs"],
        ["負荷率", finiteNumber(station.load), station.load > 1 ? STYLE.dangerPercent : station.load >= 0.85 ? STYLE.warningPercent : STYLE.vsmPercent, ""],
      ];
      details.forEach(([label, value, style, unit], detailIndex) => {
        const row = baseRow + 1 + detailIndex;
        sheet.merge(row, startColumn, row, startColumn + 2, label, STYLE.vsmLabel);
        sheet.merge(row, startColumn + 3, row, endColumn - 1, value, style);
        sheet.set(row, endColumn, unit, STYLE.vsmText);
      });

      const ctText = finiteNumber(station.ct);
      sheet.merge(baseRow + 12, startColumn, baseRow + 12, endColumn, `${ctText === null ? "—" : ctText.toFixed(1)} min`, STYLE.timelineValue);
      if (slot < stations.length - 1) {
        const gapStart = endColumn + 1;
        const gapEnd = startColumn + 7;
        const currentWait = finiteNumber(station.currentWait);
        sheet.merge(baseRow + 2, gapStart, baseRow + 3, gapEnd, `△\n${finiteNumber(station.currentWip) ?? 0} pcs`, STYLE.wip);
        sheet.merge(baseRow + 5, gapStart, baseRow + 5, gapEnd, "→", STYLE.connector);
        sheet.merge(baseRow + 11, gapStart, baseRow + 11, gapEnd, `${currentWait === null ? "—" : currentWait.toFixed(1)} hr`, STYLE.timelineValue);
      }
    });

    sheet.set(baseRow + 11, 0, "NVA", STYLE.timelineLabel);
    sheet.set(baseRow + 12, 0, "VA", STYLE.timelineLabel);
    sheet.merge(baseRow + 13, 1, baseRow + 13, 32, `區段 VA：${stations.reduce((sum, station) => sum + (finiteNumber(station.ct) ?? 0), 0).toFixed(1)} min｜區段現況 WIP：${stations.reduce((sum, station) => sum + (finiteNumber(station.currentWip) ?? 0), 0).toFixed(0)} pcs`, STYLE.divider);
    for (let row = baseRow; row <= baseRow + 13; row += 1) sheet.setRowHeight(row, row === baseRow ? 24 : 19);
  }

  function createVsmSheet(state, calc) {
    const sheet = new SheetBuilder("VSM圖");
    sheet.setColumn(0, 0, 5).setColumn(1, 32, 4.4);
    sheet.merge(1, 0, 2, 32, `VSM 現況圖｜${displayText(state.lineName, "未命名產線")}`, STYLE.title);
    sheet.setRowHeight(1, 24).setRowHeight(2, 16);
    const overallStatus = calc.bottleneck?.load > 1 ? "產能不足" : calc.bottleneck?.load >= 0.85 ? "接近滿載" : calc.bottleneck ? "產能正常" : "資料不足";
    sheet.merge(4, 5, 5, 27, `需求 Takt：${Number.isFinite(calc.takt) ? calc.takt.toFixed(1) : "—"} min／pcs　｜　瓶頸：${displayText(calc.bottleneck?.id)}　｜　${overallStatus}`, STYLE.banner);
    sheet.setRowHeight(4, 25).setRowHeight(5, 25);

    addSummaryBlock(sheet, 1, "生產條件", [
      ["月需求", finiteNumber(state.monthlyDemand), STYLE.summaryInteger],
      ["工作天", finiteNumber(state.workDays), STYLE.summaryNumber],
      ["班別", finiteNumber(state.shifts), STYLE.summaryNumber],
      ["每日有效時間", finiteNumber(calc.netAvailableDay), STYLE.summaryNumber],
      ["需求 Takt", finiteNumber(calc.takt), STYLE.summaryNumber],
    ]);
    addSummaryBlock(sheet, 11, "即時結果", [
      ["瓶頸站", displayText(calc.bottleneck?.id), STYLE.summaryText],
      ["瓶頸月產能", finiteNumber(calc.bottleneck?.capacity), STYLE.summaryInteger],
      ["瓶頸負荷率", finiteNumber(calc.bottleneck?.load), STYLE.summaryPercent],
      ["總現況 WIP", finiteNumber(calc.currentWip), STYLE.summaryInteger],
      ["總目標 WIP", finiteNumber(calc.targetWip), STYLE.summaryInteger],
    ]);
    addSummaryBlock(sheet, 21, "價值流時間", [
      ["總 VA", finiteNumber(calc.vaMinutes), STYLE.summaryNumber],
      ["總 NVA", finiteNumber(calc.nvaHours), STYLE.summaryNumber],
      ["Lead Time", finiteNumber(calc.leadTime), STYLE.summaryNumber],
      ["站數", calc.stations.length, STYLE.summaryInteger],
      ["產能判定", overallStatus, statusStyle(overallStatus)],
    ]);

    const stationGroups = [];
    for (let index = 0; index < calc.stations.length; index += 4) stationGroups.push(calc.stations.slice(index, index + 4));
    stationGroups.forEach((stations, index) => addVsmBlock(sheet, calc, stations, index));
    sheet.fitToHeight = stationGroups.length === 1 ? 1 : 0;
    sheet.freeze = { rows: 5, columns: 1 };
    return sheet;
  }

  function buildStylesXml() {
    const styles = [
      [0, 0, 0, 0, {}],
      [2, 2, 0, 0, { horizontal: "left", vertical: "center" }],
      [1, 3, 0, 0, { horizontal: "center", vertical: "center", wrapText: true }],
      [0, 0, 0, 0, { horizontal: "left", vertical: "center" }],
      [5, 4, 1, 0, { horizontal: "center", vertical: "center" }],
      [5, 4, 1, 166, { horizontal: "center", vertical: "center" }],
      [5, 4, 1, 3, { horizontal: "center", vertical: "center" }],
      [5, 4, 1, 165, { horizontal: "center", vertical: "center" }],
      [0, 5, 0, 0, { horizontal: "left", vertical: "center" }],
      [6, 0, 0, 0, { horizontal: "left", vertical: "center", wrapText: true }],
      [0, 6, 0, 0, { horizontal: "center", vertical: "center" }],
      [0, 6, 0, 164, { horizontal: "center", vertical: "center" }],
      [0, 6, 0, 3, { horizontal: "center", vertical: "center" }],
      [0, 6, 0, 165, { horizontal: "center", vertical: "center" }],
      [1, 3, 2, 0, { horizontal: "center", vertical: "center", wrapText: true }],
      [0, 5, 2, 3, { horizontal: "center", vertical: "center" }],
      [0, 4, 2, 0, { horizontal: "center", vertical: "center" }],
      [0, 4, 2, 166, { horizontal: "center", vertical: "center" }],
      [0, 4, 2, 3, { horizontal: "center", vertical: "center" }],
      [0, 4, 2, 165, { horizontal: "center", vertical: "center" }],
      [0, 6, 2, 0, { horizontal: "center", vertical: "center" }],
      [0, 6, 2, 164, { horizontal: "center", vertical: "center" }],
      [0, 6, 2, 3, { horizontal: "center", vertical: "center" }],
      [0, 6, 2, 165, { horizontal: "center", vertical: "center" }],
      [3, 10, 2, 0, { horizontal: "center", vertical: "center" }],
      [3, 11, 2, 0, { horizontal: "center", vertical: "center" }],
      [4, 9, 2, 0, { horizontal: "center", vertical: "center" }],
      [1, 7, 0, 0, { horizontal: "center", vertical: "center", wrapText: true }],
      [7, 4, 3, 0, { horizontal: "center", vertical: "center", wrapText: true }],
      [3, 12, 6, 0, { horizontal: "left", vertical: "center" }],
      [0, 12, 6, 0, { horizontal: "center", vertical: "center" }],
      [0, 12, 6, 166, { horizontal: "right", vertical: "center" }],
      [0, 12, 6, 3, { horizontal: "right", vertical: "center" }],
      [0, 12, 6, 165, { horizontal: "right", vertical: "center" }],
      [7, 0, 0, 0, { horizontal: "center", vertical: "center" }],
      [7, 8, 0, 0, { horizontal: "center", vertical: "center", wrapText: true }],
      [3, 0, 0, 0, { horizontal: "center", vertical: "center" }],
      [0, 0, 4, 164, { horizontal: "center", vertical: "center" }],
      [7, 8, 5, 0, { horizontal: "left", vertical: "center" }],
      [4, 9, 2, 166, { horizontal: "center", vertical: "center" }],
      [4, 9, 2, 165, { horizontal: "center", vertical: "center" }],
      [3, 5, 2, 0, { horizontal: "left", vertical: "center" }],
      [0, 8, 2, 0, { horizontal: "center", vertical: "center" }],
      [0, 8, 2, 164, { horizontal: "center", vertical: "center" }],
      [0, 8, 2, 3, { horizontal: "center", vertical: "center" }],
      [0, 8, 2, 165, { horizontal: "center", vertical: "center" }],
      [3, 0, 4, 0, { horizontal: "left", vertical: "center" }],
      [3, 11, 2, 165, { horizontal: "center", vertical: "center" }],
    ];
    const makeAlignment = (alignment) => {
      const attributes = [];
      if (alignment.horizontal) attributes.push(`horizontal="${alignment.horizontal}"`);
      if (alignment.vertical) attributes.push(`vertical="${alignment.vertical}"`);
      if (alignment.wrapText) attributes.push('wrapText="1"');
      return attributes.length ? `<alignment ${attributes.join(" ")}/>` : "";
    };
    const cellXfs = styles.map(([fontId, fillId, borderId, numFmtId, alignment]) => {
      const alignmentXml = makeAlignment(alignment);
      return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"${alignmentXml ? ' applyAlignment="1"' : ""}>${alignmentXml}</xf>`;
    }).join("");

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.0"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="0.0"/></numFmts>
  <fonts count="8">
    <font><sz val="10"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><b/><sz val="10"/><color rgb="FF17365D"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><b/><sz val="10"/><color rgb="FFC00000"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><sz val="10"/><color rgb="FF0070C0"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><i/><sz val="9"/><color rgb="FF666666"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
    <font><b/><sz val="10"/><color rgb="FF0B4F8A"/><name val="Microsoft JhengHei"/><family val="2"/><charset val="136"/></font>
  </fonts>
  <fills count="13">
    <fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0B4F8A"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF7F9FB"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC65911"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="7">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left/><right/><top/><bottom style="thin"><color rgb="FFA6A6A6"/></bottom><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
    <border><left style="medium"><color rgb="FF0B4F8A"/></left><right style="medium"><color rgb="FF0B4F8A"/></right><top style="medium"><color rgb="FF0B4F8A"/></top><bottom style="medium"><color rgb="FF0B4F8A"/></bottom><diagonal/></border>
    <border><left/><right/><top style="thin"><color rgb="FFA6A6A6"/></top><bottom/><diagonal/></border>
    <border><left/><right/><top/><bottom style="medium"><color rgb="FF0B4F8A"/></bottom><diagonal/></border>
    <border><left style="thin"><color rgb="FF0B4F8A"/></left><right style="thin"><color rgb="FF0B4F8A"/></right><top style="thin"><color rgb="FF0B4F8A"/></top><bottom style="thin"><color rgb="FF0B4F8A"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="${styles.length}">${cellXfs}</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
  }

  function buildWorkbookXml(sheets) {
    const printAreas = sheets.map((sheet, index) => `<definedName name="_xlnm.Print_Area" localSheetId="${index}">'${escapeXml(sheet.name)}'!$A$1:$${columnName(sheet.maxColumn)}$${sheet.maxRow}</definedName>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27328"/>
  <workbookPr date1904="0"/>
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16500" activeTab="3"/></bookViews>
  <sheets>${sheets.map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets>
  <definedNames>${printAreas}</definedNames>
  <calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>`;
  }

  function buildWorkbookRelationships(sheets) {
    const sheetRelationships = sheets.map((sheet, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRelationships}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
  }

  function buildContentTypes(sheets) {
    const sheetOverrides = sheets.map((sheet, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${sheetOverrides}
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  }

  function buildAppProperties(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>VSM 產線分析工具</Application><AppVersion>1.0</AppVersion>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheets.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>
  <TitlesOfParts><vt:vector size="${sheets.length}" baseType="lpstr">${sheets.map((sheet) => `<vt:lpstr>${escapeXml(sheet.name)}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts>
</Properties>`;
  }

  function buildCoreProperties(exportedAt) {
    const timestamp = exportedAt.toISOString();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>VSM 產線分析報表</dc:title><dc:creator>VSM 產線分析工具</dc:creator><cp:lastModifiedBy>VSM 產線分析工具</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${timestamp}</dcterms:modified>
</cp:coreProperties>`;
  }

  async function buildWorkbookBytes({ state, calc, exportedAt = new Date() }) {
    const JSZipLibrary = root.JSZip;
    if (!JSZipLibrary) throw new Error("JSZip 尚未載入");
    if (!state || !calc || !Array.isArray(calc.stations)) throw new Error("缺少 VSM 匯出資料");
    const sheets = [
      createInstructionsSheet(),
      createParametersSheet(state, calc, exportedAt),
      createStationsSheet(calc),
      createVsmSheet(state, calc),
    ];
    const zip = new JSZipLibrary();
    zip.file("[Content_Types].xml", buildContentTypes(sheets));
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
    zip.file("docProps/app.xml", buildAppProperties(sheets));
    zip.file("docProps/core.xml", buildCoreProperties(exportedAt));
    zip.file("xl/workbook.xml", buildWorkbookXml(sheets));
    zip.file("xl/_rels/workbook.xml.rels", buildWorkbookRelationships(sheets));
    zip.file("xl/styles.xml", buildStylesXml());
    sheets.forEach((sheet, index) => zip.file(`xl/worksheets/sheet${index + 1}.xml`, sheet.toXml()));
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  }

  const api = { MIME_TYPE, buildWorkbookBytes };
  root.VsmExcelExport = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : window);
