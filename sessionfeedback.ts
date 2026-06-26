type CellValue = string | number | boolean;

interface Metric {
  positive: number;
  answered: number;
}

interface Summary {
  key: string;
  date: Date;
  responses: number;
  comments: string[];
  commentNorms: string[];
  topic: Metric;
  clear: Metric;
  career: Metric;
  skills: Metric;
  confidence: Metric;
  satisfaction: Metric;
  claritySum: number;
  clarityCount: number;
}

interface ManualColumns {
  key: string;
  columnB: CellValue;
  columnC: CellValue;
}

function findSummaryByKey(summaries: Summary[], key: string): Summary | null {
  let i: number;
  for (i = 0; i < summaries.length; i++) {
    if (summaries[i].key === key) {
      return summaries[i];
    }
  }
  return null;
}

function findManualColumnsByKey(manualColumns: ManualColumns[], key: string): ManualColumns | null {
  let i: number;
  for (i = 0; i < manualColumns.length; i++) {
    if (manualColumns[i].key === key) {
      return manualColumns[i];
    }
  }
  return null;
}

function sortSummariesByDate(summaries: Summary[]): void {
  let i: number;
  let j: number;

  for (i = 0; i < summaries.length - 1; i++) {
    for (j = i + 1; j < summaries.length; j++) {
      if (summaries[i].date.getTime() > summaries[j].date.getTime()) {
        const temp: Summary = summaries[i];
        summaries[i] = summaries[j];
        summaries[j] = temp;
      }
    }
  }
}

// =========================================================
// WORKSHEET / OUTPUT HELPERS
// =========================================================

function clearSheet(ws: ExcelScript.Worksheet, totalColumns: number): void {
  const used: ExcelScript.Range | undefined = ws.getUsedRange(true);

  const rowsToClear: number = Math.max(used ? used.getRowCount() : 0, 200);
  const clearRange = ws.getRangeByIndexes(0, 0, rowsToClear, totalColumns);

  clearRange.clearAllConditionalFormats();
  clearRange.clear(ExcelScript.ClearApplyTo.all);
}

function writeHeaders(ws: ExcelScript.Worksheet, headers: string[]): void {
  const headerRange = ws.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.setValues([headers]);
  formatHeader(headerRange);
}

function writeNoDataMessage(ws: ExcelScript.Worksheet, message: string): void {
  ws.getCell(1, 0).setValue(message);
}

function formatHeader(range: ExcelScript.Range): void {
  const format: ExcelScript.RangeFormat = range.getFormat();
  format.getFont().setBold(true);
  format.getFont().setColor("white");
  format.getFill().setColor("#1F4E78");
  format.setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  format.setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  format.setWrapText(true);
}

function finalizeLayout(ws: ExcelScript.Worksheet, totalColumns: number, totalRows: number): void {
  ws.getFreezePanes().freezeRows(1);

  ws.getRange("A:A").getFormat().setColumnWidth(140);
  ws.getRange("B:C").getFormat().setColumnWidth(140);
  ws.getRange("D:D").getFormat().setColumnWidth(600);

  if (totalColumns > 4 && totalRows > 0) {
    ws.getRangeByIndexes(0, 4, totalRows, totalColumns - 4).getFormat().autofitColumns();
  }

  const used: ExcelScript.Range | undefined = ws.getUsedRange(true);
  if (used) {
    used.getFormat().autofitRows();
    used.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);

    const borderTop = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
    const borderBottom = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom);
    const borderLeft = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft);
    const borderRight = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight);
    const insideHorizontal = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal);
    const insideVertical = used.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical);

    borderTop.setStyle(ExcelScript.BorderLineStyle.continuous);
    borderBottom.setStyle(ExcelScript.BorderLineStyle.continuous);
    borderLeft.setStyle(ExcelScript.BorderLineStyle.continuous);
    borderRight.setStyle(ExcelScript.BorderLineStyle.continuous);
    insideHorizontal.setStyle(ExcelScript.BorderLineStyle.continuous);
    insideVertical.setStyle(ExcelScript.BorderLineStyle.continuous);

    borderTop.setColor("#D9E2F3");
    borderBottom.setColor("#D9E2F3");
    borderLeft.setColor("#D9E2F3");
    borderRight.setColor("#D9E2F3");
    insideHorizontal.setColor("#D9E2F3");
    insideVertical.setColor("#D9E2F3");
  }
}

function captureManualColumns(ws: ExcelScript.Worksheet): ManualColumns[] {
  const manualColumns: ManualColumns[] = [];
  const used: ExcelScript.Range | undefined = ws.getUsedRange(true);

  if (!used) {
    return manualColumns;
  }

  const values: CellValue[][] = used.getValues() as CellValue[][];
  let i: number;

  for (i = 1; i < values.length; i++) {
    const row: CellValue[] = values[i];

    const parsedDate: Date | null = parseExcelDate(getCell(row, 0));
    if (!parsedDate) {
      continue;
    }

    const key: string = dateKey(parsedDate);

    const existing: ManualColumns | null = findManualColumnsByKey(manualColumns, key);
    if (existing) {
      existing.columnB = getCell(row, 1);
      existing.columnC = getCell(row, 2);
    } else {
      manualColumns.push({
        key: key,
        columnB: getCell(row, 1),
        columnC: getCell(row, 2)
      });
    }
  }

  return manualColumns;
}

// =========================================================
// CONDITIONAL FORMATTING
// =========================================================

function applyConditionalFormatting(ws: ExcelScript.Worksheet, dataRowCount: number): void {
  if (dataRowCount <= 0) {
    return;
  }

  const lastRow: number = dataRowCount + 1;

  const pctRange: ExcelScript.Range = ws.getRange("G2:M" + String(lastRow));
  pctRange.clearAllConditionalFormats();

  addCustomFillRule(
    pctRange,
    '=AND(G2<>"",G2<0.6)',
    "#FDE9E7",
    "#9C0006",
    false
  );

  addCustomFillRule(
    pctRange,
    '=AND(G2<>"",G2>=0.6,G2<0.8)',
    "#FFF2CC",
    "#7F6000",
    false
  );

  addCustomFillRule(
    pctRange,
    '=AND(G2<>"",G2>=0.8)',
    "#E2F0D9",
    "#215E21",
    true
  );

  const clarityRange: ExcelScript.Range = ws.getRange("N2:N" + String(lastRow));
  clarityRange.clearAllConditionalFormats();

  addCustomFillRule(
    clarityRange,
    '=AND(N2<>"",N2<3)',
    "#FDE9E7",
    "#9C0006",
    false
  );

  addCustomFillRule(
    clarityRange,
    '=AND(N2<>"",N2>=3,N2<4)',
    "#FFF2CC",
    "#7F6000",
    false
  );

  addCustomFillRule(
    clarityRange,
    '=AND(N2<>"",N2>=4)',
    "#E2F0D9",
    "#215E21",
    true
  );

  const responsesRange: ExcelScript.Range = ws.getRange("E2:E" + String(lastRow));
  responsesRange.clearAllConditionalFormats();

  addCustomFillRule(
    responsesRange,
    '=AND(E2<>"",E2<5)',
    "#FCE4D6",
    "#C00000",
    true
  );

  const uniqueCommentsRange: ExcelScript.Range = ws.getRange("F2:F" + String(lastRow));
  uniqueCommentsRange.clearAllConditionalFormats();

  addCustomFillRule(
    uniqueCommentsRange,
    '=AND(F2<>"",E2<5,F2>=3)',
    "#FFF2CC",
    "#7F6000",
    true
  );
}

function addCustomFillRule(
  targetRange: ExcelScript.Range,
  formula: string,
  fillColor: string,
  fontColor: string,
  bold: boolean
): void {
  const conditionalFormat: ExcelScript.ConditionalFormat =
    targetRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);

  const customFormat: ExcelScript.CustomConditionalFormat = conditionalFormat.getCustom();
  customFormat.getRule().setFormula(formula);

  const format: ExcelScript.ConditionalRangeFormat = customFormat.getFormat();
  format.getFill().setColor(fillColor);
  format.getFont().setColor(fontColor);
  format.getFont().setBold(bold);
}

// =========================================================
// DATA ACCESS / DATE HELPERS
// =========================================================

function getCell(row: CellValue[], colIndex: number): CellValue {
  if (colIndex >= 0 && colIndex < row.length) {
    return row[colIndex];
  }
  return "";
}

function parseExcelDate(value: CellValue): Date | null {
  if (typeof value === "number") {
    return excelSerialToDateOnly(value);
  }

  const text: string = String(value).trim();
  if (text === "") {
    return null;
  }

  const parsedExplicit: Date | null = tryParseDateText(text);
  if (parsedExplicit) {
    return parsedExplicit;
  }

  const fallback: Date = new Date(text);
  if (isNaN(fallback.getTime())) {
    return null;
  }

  return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
}

function tryParseDateText(text: string): Date | null {
  const cleaned: string = text.replace("T", " ").trim();
  const firstPart: string = cleaned.split(" ")[0];

  let m: RegExpMatchArray | null = firstPart.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const day: number = parseInt(m[1], 10);
    const month: number = parseInt(m[2], 10);
    const year: number = parseInt(m[3], 10);
    if (isValidYMD(year, month, day)) {
      return new Date(year, month - 1, day);
    }
  }

  m = firstPart.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const day2: number = parseInt(m[1], 10);
    const month2: number = parseInt(m[2], 10);
    const year2: number = parseInt(m[3], 10);
    if (isValidYMD(year2, month2, day2)) {
      return new Date(year2, month2 - 1, day2);
    }
  }

  m = firstPart.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const year3: number = parseInt(m[1], 10);
    const month3: number = parseInt(m[2], 10);
    const day3: number = parseInt(m[3], 10);
    if (isValidYMD(year3, month3, day3)) {
      return new Date(year3, month3 - 1, day3);
    }
  }

  return null;
}

function isValidYMD(year: number, month: number, day: number): boolean {
  if (year < 1900 || year > 2100) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;

  const d: Date = new Date(year, month - 1, day);
  return d.getFullYear() === year && d.getMonth() === (month - 1) && d.getDate() === day;
}

function excelSerialToDateOnly(serial: number): Date {
  const utcMillis: number = Math.round((serial - 25569) * 86400 * 1000);
  const d: Date = new Date(utcMillis);
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}

function dateKey(d: Date): string {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}

function excelSerial(d: Date): number {
  return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000 + 25569;
}

function pad2(n: number): string {
  if (n < 10) {
    return "0" + String(n);
  }
  return String(n);
}

// =========================================================
// COMMENT / TEXT HELPERS
// =========================================================

function getFirst(row: CellValue[], cols: number[]): string {
  let i: number;
  for (i = 0; i < cols.length; i++) {
    const value: string = String(getCell(row, cols[i])).trim();
    if (!isBlankish(value)) {
      return value;
    }
  }
  return "";
}

function extractComments(row: CellValue[], cols: number[]): string[] {
  const results: string[] = [];
  let i: number;

  for (i = 0; i < cols.length; i++) {
    const value: string = String(getCell(row, cols[i])).trim();

    if (value.length <= 3) {
      continue;
    }

    if (isBlankish(value)) {
      continue;
    }

    if (isLikertOnlyText(value)) {
      continue;
    }

    results.push(value);
  }

  return results;
}

function joinCommentsTruncated(comments: string[], maxLength: number): string {
  if (comments.length === 0) {
    return "";
  }

  let result: string = "";
  let i: number;

  for (i = 0; i < comments.length; i++) {
    const nextLine: string = "• " + comments[i];
    const candidate: string = result === "" ? nextLine : result + "\n" + nextLine;

    if (candidate.length > maxLength) {
      if (result === "") {
        return nextLine.substring(0, maxLength);
      }
      return result + "\n" + "[Comments truncated]";
    }

    result = candidate;
  }

  return result;
}

function normalizeText(value: string): string {
  return value.toLowerCase().replace(/\s+/g, " ").trim();
}

function arrayContains(values: string[], target: string): boolean {
  let i: number;
  for (i = 0; i < values.length; i++) {
    if (values[i] === target) {
      return true;
    }
  }
  return false;
}

function isBlankish(value: string): boolean {
  const v: string = value.toLowerCase().trim();
  if (v === "") return true;
  if (v === "undefined") return true;
  if (v === "null") return true;
  if (v === "n/a") return true;
  if (v === "na") return true;
  if (v === "not applicable") return true;
  return false;
}

function isLikertOnlyText(value: string): boolean {
  const v: string = normalizeText(value);

  if (v === "strongly agree") return true;
  if (v === "agree") return true;
  if (v === "neutral") return true;
  if (v === "disagree") return true;
  if (v === "strongly disagree") return true;

  if (v === "1") return true;
  if (v === "2") return true;
  if (v === "3") return true;
  if (v === "4") return true;
  if (v === "5") return true;

  return false;
}

// =========================================================
// METRIC HELPERS
// =========================================================

function updateMetric(metric: Metric, value: string): void {
  const v: string = normalizeText(value);

  if (isBlankish(v)) {
    return;
  }

  metric.answered++;

  if (isPositiveResponse(v)) {
    metric.positive++;
  }
}

function isPositiveResponse(value: string): boolean {
  const v: string = normalizeText(value);

  if (v === "agree") return true;
  if (v === "strongly agree") return true;
  if (v === "4") return true;
  if (v === "5") return true;

  if (startsWithNumber(v, 4)) return true;
  if (startsWithNumber(v, 5)) return true;

  return false;
}

function parseNumericResponse(value: string): number | null {
  const v: string = normalizeText(value);

  if (isBlankish(v)) {
    return null;
  }

  if (/^-?\d+(\.\d+)?$/.test(v)) {
    const parsed: number = parseFloat(v);
    if (!isNaN(parsed)) {
      return parsed;
    }
  }

  const match: RegExpMatchArray | null = v.match(/^-?\d+(\.\d+)?/);
  if (match) {
    const parsed2: number = parseFloat(match[0]);
    if (!isNaN(parsed2)) {
      return parsed2;
    }
  }

  return null;
}

function startsWithNumber(value: string, expected: number): boolean {
  const match: RegExpMatchArray | null = value.match(/^(\d+)/);
  if (!match) {
    return false;
  }

  const parsed: number = parseInt(match[1], 10);
  return parsed === expected;
}

function pctDecimal(metric: Metric): string | number {
  if (metric.answered === 0) {
    return "";
  }

  return metric.positive / metric.answered;
}

function combinedPctDecimal(a: Metric, b: Metric): string | number {
  const totalAnswered: number = a.answered + b.answered;
  if (totalAnswered === 0) {
    return "";
  }

  return (a.positive + b.positive) / totalAnswered;
}

function roundTo1Decimal(value: number): number {
  return Math.round(value * 10) / 10;
}

// =========================================================
// MAIN
// =========================================================

function main(workbook: ExcelScript.Workbook): void {
  console.log("=== SESSION FEEDBACK AGGREGATOR STARTED ===");

  const SAMPLE_DATA_SHEET: string = "Sample Data";
  const OUTPUT_SHEET: string = "Session Feedback";
  const MIN_RESPONSES_TO_OUTPUT: number = 2;

  // Zero-based source column indexes from the Sample Data sheet.
  const START_TIME_COL: number = 1;
  const COLS_COMMENTS: number[] = [8, 22];
  const COLS_TOPIC_RELEVANT: number[] = [13, 24];
  const COLS_SPEAKER_CLEAR: number[] = [17];
  const COLS_CAREER_INTEREST: number[] = [18];
  const COLS_SKILLS_KNOWLEDGE: number[] = [19];
  const COLS_CONFIDENCE_MH: number[] = [20];
  const COLS_PROGRAMME_SATISFACTION: number[] = [21];
  const COLS_CLARITY_RATING: number[] = [12];

  const HEADERS: string[] = [
    "Date",
    "Manual Column B",
    "Manual Column C",
    "Aggregated Comments",
    "Responses",
    "Unique Comments",
    "Topic Relevant %",
    "Speaker Clear %",
    "Combined %",
    "Career Interest %",
    "Skills/Knowledge %",
    "Confidence MH %",
    "Programme Satisfaction %",
    "Avg Clarity",
    "Clarity Responses"
  ];

  const sampleDataWS: ExcelScript.Worksheet | undefined = workbook.getWorksheet(SAMPLE_DATA_SHEET);
  if (!sampleDataWS) {
    console.log("Source worksheet not found: " + SAMPLE_DATA_SHEET);
    return;
  }

  let outputWS: ExcelScript.Worksheet | undefined = workbook.getWorksheet(OUTPUT_SHEET);
  if (!outputWS) {
    outputWS = workbook.addWorksheet(OUTPUT_SHEET);
  }

  const manualColumns: ManualColumns[] = captureManualColumns(outputWS);

  const usedRange: ExcelScript.Range | undefined = sampleDataWS.getUsedRange(true);
  if (!usedRange) {
    clearSheet(outputWS, HEADERS.length);
    writeHeaders(outputWS, HEADERS);
    writeNoDataMessage(outputWS, "No source data found on '" + SAMPLE_DATA_SHEET + "'.");
    finalizeLayout(outputWS, HEADERS.length, 2);
    console.log("No used range found on source worksheet.");
    return;
  }

  const data: CellValue[][] = usedRange.getValues() as CellValue[][];
  if (data.length < 2) {
    clearSheet(outputWS, HEADERS.length);
    writeHeaders(outputWS, HEADERS);
    writeNoDataMessage(outputWS, "No data rows found.");
    finalizeLayout(outputWS, HEADERS.length, 2);
    console.log("No data rows found.");
    return;
  }

  // =========================================================
  // AGGREGATE
  // =========================================================

  const summaries: Summary[] = [];
  let rowIndex: number;

  for (rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row: CellValue[] = data[rowIndex];

    const rawDateValue: CellValue = getCell(row, START_TIME_COL);
    const parsedDate: Date | null = parseExcelDate(rawDateValue);
    if (!parsedDate) {
      continue;
    }

    const key: string = dateKey(parsedDate);
    let summary: Summary | null = findSummaryByKey(summaries, key);

    if (!summary) {
      summary = createSummary(key, parsedDate);
      summaries.push(summary);
    }

    summary.responses++;

    const extractedComments: string[] = extractComments(row, COLS_COMMENTS);
    let j: number;
    for (j = 0; j < extractedComments.length; j++) {
      const comment: string = extractedComments[j];
      const normalized: string = normalizeText(comment);

      if (!arrayContains(summary.commentNorms, normalized)) {
        summary.commentNorms.push(normalized);
        summary.comments.push(comment);
      }
    }

    updateMetric(summary.topic, getFirst(row, COLS_TOPIC_RELEVANT));
    updateMetric(summary.clear, getFirst(row, COLS_SPEAKER_CLEAR));
    updateMetric(summary.career, getFirst(row, COLS_CAREER_INTEREST));
    updateMetric(summary.skills, getFirst(row, COLS_SKILLS_KNOWLEDGE));
    updateMetric(summary.confidence, getFirst(row, COLS_CONFIDENCE_MH));
    updateMetric(summary.satisfaction, getFirst(row, COLS_PROGRAMME_SATISFACTION));

    const clarityText: string = getFirst(row, COLS_CLARITY_RATING);
    const clarityValue: number | null = parseNumericResponse(clarityText);
    if (clarityValue !== null) {
      summary.claritySum += clarityValue;
      summary.clarityCount++;
    }
  }

  sortSummariesByDate(summaries);

  // =========================================================
  // FILTER OUT LOW-RESPONSE SESSIONS
  // =========================================================

  const filteredSummaries: Summary[] = [];
  let i: number;

  for (i = 0; i < summaries.length; i++) {
    if (summaries[i].responses >= MIN_RESPONSES_TO_OUTPUT) {
      filteredSummaries.push(summaries[i]);
    }
  }

  // =========================================================
  // WRITE OUTPUT
  // =========================================================

  clearSheet(outputWS, HEADERS.length);
  writeHeaders(outputWS, HEADERS);

  if (filteredSummaries.length === 0) {
    writeNoDataMessage(
      outputWS,
      "No session feedback data found with at least " + String(MIN_RESPONSES_TO_OUTPUT) + " responses."
    );
    finalizeLayout(outputWS, HEADERS.length, 2);
    console.log("No valid session feedback data found after minimum response filtering.");
    return;
  }

  const outputData: (string | number | boolean)[][] = [];

  for (i = 0; i < filteredSummaries.length; i++) {
    const s: Summary = filteredSummaries[i];

    let avgClarity: string | number = "";
    if (s.clarityCount > 0) {
      avgClarity = roundTo1Decimal(s.claritySum / s.clarityCount);
    }

    const manual: ManualColumns | null = findManualColumnsByKey(manualColumns, s.key);
    const retainedB: CellValue = manual ? manual.columnB : "";
    const retainedC: CellValue = manual ? manual.columnC : "";

    const outputRow: (string | number | boolean)[] = [
      excelSerial(s.date),
      retainedB,
      retainedC,
      joinCommentsTruncated(s.comments, 32000),
      s.responses,
      s.comments.length,
      pctDecimal(s.topic),
      pctDecimal(s.clear),
      combinedPctDecimal(s.topic, s.clear),
      pctDecimal(s.career),
      pctDecimal(s.skills),
      pctDecimal(s.confidence),
      pctDecimal(s.satisfaction),
      avgClarity,
      s.clarityCount
    ];

    outputData.push(outputRow);
  }

  outputWS.getRangeByIndexes(1, 0, outputData.length, HEADERS.length).setValues(outputData);

  // =========================================================
  // FORMAT OUTPUT
  // =========================================================

  outputWS.getRangeByIndexes(1, 0, outputData.length, 1).setNumberFormatLocal("dd mmmm yyyy");
  outputWS.getRangeByIndexes(1, 6, outputData.length, 7).setNumberFormatLocal("0%");
  outputWS.getRangeByIndexes(1, 13, outputData.length, 1).setNumberFormatLocal("0.0");
  outputWS.getRange("D:D").getFormat().setWrapText(true);

  applyConditionalFormatting(outputWS, outputData.length);
  finalizeLayout(outputWS, HEADERS.length, outputData.length + 1);

  console.log("=== SESSION FEEDBACK AGGREGATOR COMPLETE ===");
}

// =========================================================
// SUMMARY HELPERS
// =========================================================

function createSummary(key: string, date: Date): Summary {
  return {
    key: key,
    date: date,
    responses: 0,
    comments: [],
    commentNorms: [],
    topic: initMetric(),
    clear: initMetric(),
    career: initMetric(),
    skills: initMetric(),
    confidence: initMetric(),
    satisfaction: initMetric(),
    claritySum: 0,
    clarityCount: 0
  };
}

function initMetric(): Metric {
  return {
    positive: 0,
    answered: 0
  };
}
