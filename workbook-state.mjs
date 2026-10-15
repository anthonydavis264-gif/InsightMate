import { createHash, randomUUID } from "node:crypto";
import ExcelJS from "exceljs";

function cloneCellValue(value) {
  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (Array.isArray(value)) {
    return value.map(cloneCellValue);
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, entry]) => [
        key,
        cloneCellValue(entry)
      ])
    );
  }

  return value;
}

export function serializeCellValue(value, type) {
  if (value === null || value === undefined) {
    return { type: "blank", value: null };
  }

  if (value instanceof Date) {
    return { type: "date", value: value.toISOString() };
  }

  if (type === ExcelJS.ValueType.Formula && value && typeof value === "object") {
    return {
      type: "formula",
      formula: value.formula ?? value.sharedFormula ?? null,
      result: serializeCellValue(value.result)
    };
  }

  if (type === ExcelJS.ValueType.Error) {
    const errorValue = value && typeof value === "object"
      ? value.error
      : value;
    return { type: "error", value: String(errorValue) };
  }

  if (typeof value === "string") {
    return { type: "string", value };
  }

  if (typeof value === "number") {
    return { type: "number", value };
  }

  if (typeof value === "boolean") {
    return { type: "boolean", value };
  }

  if (value && typeof value === "object") {
    return { type: "unsupported", value: cloneCellValue(value) };
  }

  return { type: "unsupported", value: String(value) };
}

const STANDARD_EXCEL_ERRORS = new Set([
  "#N/A",
  "#REF!",
  "#NAME?",
  "#DIV/0!",
  "#NULL!",
  "#VALUE!",
  "#NUM!",
  "#ERROR!"
]);

const NON_NAME_PARTS = new Set([
  "de", "del", "di", "da", "van", "von", "la", "le", "st",
  "mr", "mrs", "ms", "miss", "dr", "prof", "rev", "fr", "sir",
  "jr", "sr", "ii", "iii", "iv", "v", "phd", "md", "esq"
]);

function personNameHeaderMode(value) {
  if (typeof value !== "string") return null;
  const header = value.trim().replace(/ +/g, " ").toLocaleLowerCase("en-US");
  if (header === "first name" || header === "last name") return "single";
  if (
    header === "person name" ||
    header === "full name" ||
    header === "contact" ||
    header === "contact name" ||
    header === "salesperson" ||
    header === "employee name"
  ) return "full";
  return null;
}

function personNameColumns(worksheet, headerRowNumber) {
  const columns = new Map();
  if (headerRowNumber === null) return columns;
  worksheet.getRow(headerRowNumber).eachCell({ includeEmpty: false }, (cell, column) => {
    const mode = personNameHeaderMode(cell.value);
    if (mode) columns.set(column, mode);
  });
  return columns;
}

function snapshotPersonNameColumns(sheet, headerRowNumber) {
  const columns = new Map();
  if (headerRowNumber === null) return columns;
  for (const cell of sheet.cells) {
    if (cell.row !== headerRowNumber) continue;
    const mode = personNameHeaderMode(cell.value);
    if (mode) columns.set(cell.column, mode);
  }
  return columns;
}

function isProtectedText(value, cell) {
  return Boolean(
    cell?.isMerged ||
    cell?.type !== undefined && cell.type !== ExcelJS.ValueType.String ||
    cell?.hyperlink ||
    /(?:https?:\/\/|www\.)/i.test(value) ||
    /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(value) ||
    /\d/.test(value) ||
    /[-_./:\\]/.test(value)
  );
}

function capitalizePersonName(value, mode) {
  if (!mode || /[^\x20-\x7e]/.test(value)) return value;
  const parts = value.split(" ");
  const hasPotentialMcOrMacCase = parts.some((part) =>
    /^(?:MC|MAC)[A-Z]+$/.test(part) ||
    /^M(?:c|ac)[A-Z]/.test(part)
  );
  if (
    mode === "single" && parts.length !== 1 ||
    mode === "full" && parts.length < 2 ||
    hasPotentialMcOrMacCase ||
    parts.some((part) => !/^[A-Za-z]{2,}$/.test(part) ||
      NON_NAME_PARTS.has(part.toLocaleLowerCase("en-US")))
  ) {
    return value;
  }

  const isTitleCase = (part) => /^[A-Z][a-z]+$/.test(part);
  const hasTitleCaseAnchor = parts.some(isTitleCase);
  const containsAmbiguousAcronym = parts.some((part) => /^[A-Z]{2,4}$/.test(part));
  const looksLikeAllCapsName = mode === "full" &&
    parts.every((part) => /^[A-Z]{4,}$/.test(part));
  const hasMixedSingleName = mode === "single" &&
    parts.some((part) => /[A-Z]/.test(part) && /[a-z]/.test(part)) &&
    parts[0].length >= 5;
  if (
    !hasTitleCaseAnchor &&
    !hasMixedSingleName &&
    !looksLikeAllCapsName
  ) return value;
  if (
    containsAmbiguousAcronym &&
    !looksLikeAllCapsName &&
    parts.some((part) => /^[A-Z]{4}$/.test(part))
  ) return value;
  if (
    containsAmbiguousAcronym &&
    !hasTitleCaseAnchor &&
    !hasMixedSingleName &&
    !looksLikeAllCapsName
  ) {
    return value;
  }

  return parts
    .map((part) => part[0].toLocaleUpperCase("en-US") + part.slice(1).toLocaleLowerCase("en-US"))
    .join(" ");
}

function normalizeTextValue(value, cell, nameMode = null) {
  if (typeof value !== "string" || isProtectedText(value, cell)) return value;

  const normalized = value
    .replace(/^[ ]+|[ ]+$/g, "")
    .replace(/ {2,}/g, " ");
  return capitalizePersonName(normalized, nameMode);
}

export function normalizeComparableText(value) {
  return normalizeTextValue(value, null);
}

function getRecognizedCellError(cell) {
  let error;
  let formula = null;
  if (cell.type === ExcelJS.ValueType.Error) {
    error = typeof cell.value === "string" ? cell.value : cell.value?.error;
  } else if (
    cell.type === ExcelJS.ValueType.Formula &&
    cell.value &&
    typeof cell.value === "object"
  ) {
    const result = cell.value.result;
    error = result?.error;
    formula = cell.value.formula ?? cell.value.sharedFormula ?? null;
  } else if (
    cell.type === ExcelJS.ValueType.String &&
    typeof cell.value === "string"
  ) {
    error = cell.value;
  }
  const normalizedError = typeof error === "string"
    ? error.replace(/^[ ]+|[ ]+$/g, "").toUpperCase()
    : null;
  return normalizedError && STANDARD_EXCEL_ERRORS.has(normalizedError)
    ? { error, formula }
    : null;
}

function columnNumberFromName(name) {
  let column = 0;
  for (const character of name) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column;
}

function mergedRangeCoordinates(range) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    startColumn: columnNumberFromName(match[1]),
    startRow: Number(match[2]),
    endColumn: columnNumberFromName(match[3]),
    endRow: Number(match[4])
  };
}

function mergeHeaderRow(worksheet, preferredHeaderRow) {
  if (preferredHeaderRow !== null) return preferredHeaderRow;
  for (let rowNumber = 1; rowNumber < worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    const labels = [];
    let containsMergedCell = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      containsMergedCell ||= cell.isMerged;
      if (typeof cell.value === "string" && cell.value.trim()) {
        labels.push(cell.value.trim().replace(/ +/g, " ").toLocaleLowerCase("en-US"));
      }
    });
    if (
      !containsMergedCell &&
      labels.length >= 2 &&
      new Set(labels).size === labels.length
    ) return rowNumber;
  }
  return null;
}

function normalizedHeaderLabels(worksheet, headerRowNumber, firstColumn, lastColumn) {
  const labels = [];
  for (let column = firstColumn; column <= lastColumn; column++) {
    const value = worksheet.getCell(headerRowNumber, column).value;
    if (typeof value !== "string" || !value.trim()) return null;
    labels.push(value.trim().replace(/ +/g, " ").toLocaleLowerCase("en-US"));
  }
  return new Set(labels).size === labels.length ? labels : null;
}

function dataRowHasIndependentValues(worksheet, rowNumber, firstColumn, lastColumn) {
  for (let column = firstColumn; column <= lastColumn; column++) {
    const cell = worksheet.getCell(rowNumber, column);
    if (cell.isMerged || !hasMeaningfulCellValue(cell.value)) return false;
  }
  return true;
}

function independentDataContextCount(worksheet, rowNumber, firstColumn, lastColumn, excludedStart, excludedEnd) {
  let count = 0;
  for (let column = firstColumn; column <= lastColumn; column++) {
    if (column >= excludedStart && column <= excludedEnd) continue;
    const cell = worksheet.getCell(rowNumber, column);
    if (!cell.isMerged && hasMeaningfulCellValue(cell.value)) count++;
  }
  return count;
}

function nearbyIndependentDataRows(worksheet, headerRow, mergeRow, lastHeaderColumn, firstColumn, lastColumn) {
  let count = 0;
  const firstRow = Math.max(headerRow + 1, mergeRow - 5);
  const lastRow = Math.min(worksheet.rowCount, mergeRow + 5);
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber++) {
    if (
      rowNumber !== mergeRow &&
      independentDataContextCount(
        worksheet,
        rowNumber,
        1,
        lastHeaderColumn,
        firstColumn,
        lastColumn
      ) > 0 &&
      dataRowHasIndependentValues(
        worksheet,
        rowNumber,
        firstColumn,
        lastColumn
      )
    ) {
      count++;
    }
  }
  return count;
}

function nearbyDataContextRows(worksheet, headerRow, mergeRow, lastHeaderColumn, excludedStart, excludedEnd) {
  let count = 0;
  const firstRow = Math.max(headerRow + 1, mergeRow - 5);
  const lastRow = Math.min(worksheet.rowCount, mergeRow + 5);
  for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber++) {
    if (
      rowNumber !== mergeRow &&
      independentDataContextCount(
        worksheet,
        rowNumber,
        1,
        lastHeaderColumn,
        excludedStart,
        excludedEnd
      ) > 0
    ) {
      count++;
    }
  }
  return count;
}

function trailingMergedColumnsAreUnused(worksheet, firstRow, firstColumn, lastColumn) {
  for (let rowNumber = firstRow; rowNumber <= worksheet.rowCount; rowNumber++) {
    for (let column = firstColumn; column <= lastColumn; column++) {
      const cell = worksheet.getCell(rowNumber, column);
      if (
        hasMeaningfulCellValue(cell.value) &&
        (!cell.isMerged || cell.master.address === cell.address)
      ) {
        return false;
      }
    }
  }
  return true;
}

function mergedAnchorProfile(cell) {
  if (cell.type === ExcelJS.ValueType.Formula) return null;
  if (cell.type === ExcelJS.ValueType.Date) return "date";
  if (cell.type === ExcelJS.ValueType.Number) {
    return /[$€£¥]/.test(cell.numFmt || "") ? "currency" : "number";
  }
  if (cell.type !== ExcelJS.ValueType.String || typeof cell.value !== "string") return null;
  const value = cell.value.trim();
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)) return "email";
  if (/^(?:https?:\/\/|www\.)\S+$/i.test(value)) return "url";
  const digits = value.replace(/\D/g, "").length;
  if (
    digits >= 7 &&
    digits <= 15 &&
    /^[+()\d .-]+$/.test(value)
  ) return "phone-like";
  return "text";
}

function inferMergedAnchorColumn(worksheet, headerRow, mergeRow, firstColumn, lastColumn, anchor) {
  const profile = mergedAnchorProfile(anchor);
  if (!profile || profile === "text") return firstColumn;

  const profileCounts = new Map();
  for (let column = firstColumn; column <= lastColumn; column++) {
    profileCounts.set(column, { matching: 0, total: 0 });
  }
  for (let rowNumber = headerRow + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
    if (rowNumber === mergeRow) continue;
    for (let column = firstColumn; column <= lastColumn; column++) {
      const cell = worksheet.getCell(rowNumber, column);
      if (cell.isMerged || !hasMeaningfulCellValue(cell.value)) continue;
      const counts = profileCounts.get(column);
      counts.total++;
      if (mergedAnchorProfile(cell) === profile) counts.matching++;
    }
  }

  const likelyColumns = [...profileCounts.entries()]
    .filter(([, counts]) => counts.matching >= 2 && counts.matching / counts.total >= 0.8)
    .map(([column]) => column);
  return likelyColumns.length === 1 ? likelyColumns[0] : firstColumn;
}

function normalizeMergedDataFields(worksheet, headerRowNumber) {
  const dataHeaderRow = mergeHeaderRow(worksheet, headerRowNumber);
  const actions = [];
  const ranges = [...(worksheet.model.merges || [])];
  const rangeFrequency = new Map();
  let lastHeaderColumn = 0;
  if (dataHeaderRow !== null) {
    for (let column = 1; column <= worksheet.columnCount; column++) {
      const value = worksheet.getCell(dataHeaderRow, column).value;
      if (typeof value === "string" && value.trim()) lastHeaderColumn = column;
    }
  }
  for (const range of ranges) {
    const coordinates = mergedRangeCoordinates(range);
    if (!coordinates || coordinates.startRow !== coordinates.endRow) continue;
    const key = `${coordinates.startColumn}:${coordinates.endColumn}`;
    rangeFrequency.set(key, (rangeFrequency.get(key) || 0) + 1);
  }

  for (const originalRange of ranges) {
    const coordinates = mergedRangeCoordinates(originalRange);
    let reason = "The range could not be confidently classified as an accidental data-row merge.";
    let canUnmerge = false;
    let extendsPastHeaders = false;
    let independentRows = 0;
    if (!coordinates) reason = "The merged range bounds could not be parsed safely.";
    else if (coordinates.startRow !== coordinates.endRow) {
      reason = "Vertical merged ranges were retained because their row-spanning purpose is ambiguous.";
    } else if (dataHeaderRow === null) {
      reason = "No reliable, unmerged header row was identified.";
    } else if (coordinates.startRow <= dataHeaderRow) {
      reason = coordinates.startRow === dataHeaderRow
        ? "Merged header ranges were retained to preserve worksheet headings."
        : "Title or report-area merges were retained to preserve worksheet layout.";
    } else {
      const frequencyKey = `${coordinates.startColumn}:${coordinates.endColumn}`;
      const headerEndColumn = Math.min(coordinates.endColumn, lastHeaderColumn);
      const hasNamedHeaderSpan =
        coordinates.startColumn <= headerEndColumn &&
        normalizedHeaderLabels(
          worksheet,
          dataHeaderRow,
          coordinates.startColumn,
          headerEndColumn
        );
      extendsPastHeaders = coordinates.endColumn > lastHeaderColumn;
      if (!hasNamedHeaderSpan) {
        reason = "The merged range does not cover distinct, identifiable data headers.";
      } else if (
        extendsPastHeaders &&
        (
          coordinates.startColumn > lastHeaderColumn ||
          !trailingMergedColumnsAreUnused(
            worksheet,
            dataHeaderRow + 1,
            lastHeaderColumn + 1,
            coordinates.endColumn
          )
        )
      ) {
        reason = "The merged range extends into columns with no identifiable headers, but those columns contain independent data.";
      } else {
        const rowContextCount = independentDataContextCount(
          worksheet,
          coordinates.startRow,
          1,
          lastHeaderColumn,
          coordinates.startColumn,
          coordinates.endColumn
        );
        if (extendsPastHeaders) {
          independentRows = nearbyDataContextRows(
            worksheet,
            dataHeaderRow,
            coordinates.startRow,
            lastHeaderColumn,
            coordinates.startColumn,
            coordinates.endColumn
          );
        } else {
          independentRows = nearbyIndependentDataRows(
            worksheet,
            dataHeaderRow,
            coordinates.startRow,
            lastHeaderColumn,
            coordinates.startColumn,
            coordinates.endColumn
          );
        }
        if (
          hasMeaningfulCellValue(
            worksheet.getCell(coordinates.startRow, coordinates.startColumn).value
          ) &&
          rowContextCount > 0 &&
          independentRows >= 2
        ) {
          const anchor = worksheet.getCell(coordinates.startRow, coordinates.startColumn);
          const anchorType = anchor.type === ExcelJS.ValueType.Formula
            ? "formula"
            : anchor.type === ExcelJS.ValueType.Number
              ? "numeric"
              : anchor.type === ExcelJS.ValueType.Date
                ? "date"
                : anchor.type === ExcelJS.ValueType.Error
                  ? "error"
                  : hasMeaningfulCellValue(anchor.value)
                    ? "text or other value"
                    : "blank";
          const destinationColumn = inferMergedAnchorColumn(
            worksheet,
            dataHeaderRow,
            coordinates.startRow,
            coordinates.startColumn,
            coordinates.endColumn,
            anchor
          );
          canUnmerge = true;
          reason = extendsPastHeaders
            ? `This data-row merge extends beyond the identified header span into unused columns. The row has ${rowContextCount} other populated data fields and ${independentRows} nearby records confirm the surrounding data region. The ${anchorType} anchor remains in its original column.`
            : `This horizontal merge spans distinct data fields, and ${independentRows} nearby records contain independent values in those columns. The ${anchorType} anchor ${destinationColumn === coordinates.startColumn ? "remains" : "will be moved"} ${destinationColumn === coordinates.startColumn ? "in its original column" : "to the column with the strongest matching value profile"}.`;
          coordinates.destinationColumn = extendsPastHeaders
            ? coordinates.startColumn
            : destinationColumn;
        } else {
          reason = rangeFrequency.get(frequencyKey) > 1
            ? "These columns are merged repeatedly, but the pattern lacks enough nearby independent data to distinguish data fields from a presentation layout."
            : "Nearby records do not provide enough independent data to prove that the merged cells belong to a data row.";
        }
      }
    }

    if (!canUnmerge) {
      actions.push({
        worksheetName: worksheet.name,
        originalRange,
        reason,
        action: "Retained merged range.",
        unresolvedMissingFields: []
      });
      continue;
    }

    const anchorCell = worksheet.getCell(coordinates.startRow, coordinates.startColumn);
    const anchorValue = cloneCellValue(anchorCell.value);
    const anchorWasBlank = !hasMeaningfulCellValue(anchorValue);
    worksheet.unMergeCells(originalRange);
    const destinationCell = worksheet.getCell(
      coordinates.startRow,
      coordinates.destinationColumn || coordinates.startColumn
    );
    if (destinationCell !== anchorCell) anchorCell.value = null;
    destinationCell.value = anchorValue;
    const unresolvedColumns = anchorWasBlank
      ? Array.from(
          { length: coordinates.endColumn - coordinates.startColumn + 1 },
          (_, index) => coordinates.startColumn + index
        )
      : Array.from(
          { length: coordinates.endColumn - coordinates.startColumn + 1 },
          (_, index) => coordinates.startColumn + index
        ).filter((column) => column !== destinationCell.col);
    actions.push({
      worksheetName: worksheet.name,
      originalRange,
      reason,
      action: extendsPastHeaders
        ? "Unmerged the unused trailing-column merge and preserved the anchor value in its original column."
        : destinationCell === anchorCell
          ? "Unmerged separate data fields and preserved the anchor value in its original column."
          : "Unmerged separate data fields and moved the anchor value to the column with the strongest matching value profile.",
      unresolvedMissingFields: unresolvedColumns.map((column) =>
          worksheet.getCell(dataHeaderRow, column).value ||
          `Column ${columnName(column)}`
        )
    });
  }
  return actions;
}

function getHeaderCandidate(worksheet) {
  let candidate = null;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (!candidate && row.values.slice(1).some((value) => value !== null && value !== undefined && value !== "")) {
      candidate = { row, rowNumber };
    }
  });

  if (!candidate) {
    return {
      status: "confirmation_required",
      candidateRow: null,
      reason: "No non-empty row was found."
    };
  }

  const values = candidate.row.values.slice(1);
  const labels = values.filter((value) => value !== null && value !== undefined && value !== "");
  const normalized = labels.map((value) => String(value).trim().toLocaleLowerCase());
  const hasLaterData = worksheet.rowCount > candidate.rowNumber;
  let hasMergedCells = false;
  candidate.row.eachCell({ includeEmpty: true }, (cell) => {
    hasMergedCells ||= cell.isMerged;
  });
  const reliable =
    labels.length >= 2 &&
    labels.every((value) => typeof value === "string" && value.trim() !== "") &&
    new Set(normalized).size === normalized.length &&
    hasLaterData &&
    !hasMergedCells;

  return reliable
    ? { status: "candidate", candidateRow: candidate.rowNumber, reason: null }
    : {
        status: "confirmation_required",
        candidateRow: candidate.rowNumber,
        reason: "The first non-empty row does not meet the conservative header criteria."
      };
}

function describeWorksheet(worksheet, order) {
  const cells = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      cells.push({
        address: cell.address,
        row: rowNumber,
        column: columnNumber,
        type: cell.type,
        value: cloneCellValue(cell.value),
        serializedValue: serializeCellValue(cell.value, cell.type),
        numberFormat: cell.numFmt || null,
        isMerged: cell.isMerged
      });
    });
  });

  return {
    worksheetKey: `sheet-${order}`,
    name: worksheet.name,
    order,
    visibility: worksheet.state,
    rowCount: worksheet.rowCount,
    columnCount: worksheet.columnCount,
    headerCandidate: getHeaderCandidate(worksheet),
    cells
  };
}

function proposalId(uploadVersionId, reviewId, worksheetKey, target, issueType) {
  return createHash("sha256")
    .update(JSON.stringify([uploadVersionId, reviewId, worksheetKey, target, issueType]))
    .digest("hex");
}

function isBlank(value) {
  return value === null || value === undefined || value === "";
}

function cellFor(sheet, row, column) {
  return sheet.cells.find((cell) => cell.row === row && cell.column === column);
}

function hasMeaningfulCellValue(value) {
  return value !== null && value !== undefined && value !== "";
}

function rowHasMeaningfulValues(row, columnCount) {
  for (let column = 1; column <= columnCount; column++) {
    if (hasMeaningfulCellValue(row.getCell(column).value)) return true;
  }
  return false;
}

function blankRowValueStatus(row, columnCount) {
  let hasWhitespace = false;
  for (let column = 1; column <= columnCount; column++) {
    const value = row.getCell(column).value;
    if (value === null || value === undefined || value === "") continue;
    if (typeof value === "string" && value.trim() === "") {
      hasWhitespace = true;
      continue;
    }
    return { isBlank: false, hasWhitespace: false };
  }
  return { isBlank: true, hasWhitespace };
}

const DISPOSABLE_BLANK_CELL_STYLE_KEYS = new Set([
  "font",
  "border",
  "fill",
  "alignment",
  "numFmt"
]);

function blankRowStructureIssue(row, columnCount) {
  if (
    row.height ||
    row.hidden ||
    row.outlineLevel ||
    row.collapsed ||
    row.pageBreak
  ) {
    return "The blank row has custom row-level metadata.";
  }

  for (const key of Object.keys(row.style || {})) {
    if (!DISPOSABLE_BLANK_CELL_STYLE_KEYS.has(key)) {
      return "The blank row has non-disposable row formatting or protection.";
    }
  }

  for (let column = 1; column <= columnCount; column++) {
    const cell = row.getCell(column);
    if (cell.isMerged) return "The blank row contains merged cells.";
    if (cell.note || cell.hyperlink || cell.dataValidation || cell.protection) {
      return "The blank row contains cell metadata or protection.";
    }
    for (const key of Object.keys(cell.style || {})) {
      if (!DISPOSABLE_BLANK_CELL_STYLE_KEYS.has(key)) {
        return "The blank row has non-disposable cell formatting or protection.";
      }
    }
  }
  return null;
}

function ordinaryDataRowIssue(
  worksheet,
  rowNumber,
  headerRowNumber,
  headerColumnCount,
  headerLabels
) {
  if (headerRowNumber === null || headerColumnCount < 2) {
    return "A reliable multi-column header was not identified for this worksheet.";
  }
  if (rowNumber <= headerRowNumber) {
    return "A neighboring row is at or above the identified data header.";
  }

  const populatedValues = [];
  for (let column = 1; column <= headerColumnCount; column++) {
    const value = worksheet.getCell(rowNumber, column).value;
    if (
      value === null ||
      value === undefined ||
      value === "" ||
      typeof value === "string" && value.trim() === ""
    ) continue;
    populatedValues.push(value);
  }

  const minimumFields = Math.min(3, headerColumnCount);
  if (populatedValues.length < minimumFields) {
    return "A neighboring row has too few populated fields to confirm an ordinary data record.";
  }

  const matchingHeaderLabels = populatedValues.filter((value) =>
    typeof value === "string" &&
    headerLabels.has(value.trim().replace(/ +/g, " ").toLocaleLowerCase("en-US"))
  ).length;
  if (
    matchingHeaderLabels >= 2 &&
    matchingHeaderLabels * 2 >= populatedValues.length
  ) {
    return "A neighboring row resembles a repeated header.";
  }
  return null;
}

function getHeaderColumnCount(row) {
  let lastColumn = 0;
  row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
    if (hasMeaningfulCellValue(cell.value)) lastColumn = columnNumber;
  });
  return lastColumn;
}

function getAmbiguousDateText(value) {
  if (typeof value !== "string") return false;
  const match = value.trim().match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  if (!match) return false;
  const first = Number(match[1]);
  const second = Number(match[2]);
  return first >= 1 && first <= 12 &&
    second >= 1 && second <= 12 &&
    first !== second;
}

function worksheetHasMergedRanges(worksheet) {
  return (worksheet.model.merges || []).length > 0;
}

function formulaReferencesWorksheet(formula, worksheetName) {
  if (typeof formula !== "string" || !formula) return false;
  const escapedName = worksheetName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const quotedReference = `'${escapedName.replace(/'/g, "''")}'!`;
  const simpleReference = `${escapedName}!`;
  return new RegExp(`${quotedReference}|${simpleReference}`, "i").test(formula);
}

function isRowInvariantFormula(formula) {
  if (typeof formula !== "string" || !formula) return false;
  const expression = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = [];
  let offset = 0;

  while (offset < expression.length) {
    if (/\s/.test(expression[offset])) {
      offset++;
      continue;
    }
    const number = expression.slice(offset).match(
      /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/
    );
    if (number) {
      tokens.push("number");
      offset += number[0].length;
      continue;
    }
    if ("()+-*/^%".includes(expression[offset])) {
      tokens.push(expression[offset]);
      offset++;
      continue;
    }
    return false;
  }

  let index = 0;
  function parsePrimary() {
    if (tokens[index] === "number") {
      index++;
      return true;
    }
    if (tokens[index] !== "(") return false;
    index++;
    if (!parseExpression() || tokens[index] !== ")") return false;
    index++;
    return true;
  }
  function parseUnary() {
    if (tokens[index] === "+" || tokens[index] === "-") {
      index++;
      return parseUnary();
    }
    return parsePrimary();
  }
  function parsePower() {
    if (!parseUnary()) return false;
    while (tokens[index] === "%") index++;
    if (tokens[index] === "^") {
      index++;
      return parsePower();
    }
    return true;
  }
  function parseTerm() {
    if (!parsePower()) return false;
    while (tokens[index] === "*" || tokens[index] === "/") {
      index++;
      if (!parsePower()) return false;
    }
    return true;
  }
  function parseExpression() {
    if (!parseTerm()) return false;
    while (tokens[index] === "+" || tokens[index] === "-") {
      index++;
      if (!parseTerm()) return false;
    }
    return true;
  }

  return tokens.length > 0 && parseExpression() && index === tokens.length;
}

function parseCellAddress(address) {
  const match = typeof address === "string"
    ? address.match(/^\$?([A-Z]{1,3})\$?(\d+)$/i)
    : null;
  if (!match) return null;
  const column = [...match[1].toUpperCase()].reduce(
    (value, character) => value * 26 + character.charCodeAt(0) - 64,
    0
  );
  return { column, row: Number(match[2]) };
}

function sharedFormulaRange(worksheet, masterCell) {
  const ref = masterCell.value?.ref;
  if (typeof ref !== "string") return null;
  const endpoints = ref.split(":").map(parseCellAddress);
  if (endpoints.length !== 2 || endpoints.some((point) => !point)) return null;
  const [start, end] = endpoints;
  if (
    masterCell.col < Math.min(start.column, end.column) ||
    masterCell.col > Math.max(start.column, end.column) ||
    masterCell.row < Math.min(start.row, end.row) ||
    masterCell.row > Math.max(start.row, end.row)
  ) {
    return null;
  }
  return {
    master: masterCell.address,
    firstRow: Math.min(start.row, end.row),
    lastRow: Math.max(start.row, end.row),
    firstColumn: Math.min(start.column, end.column),
    lastColumn: Math.max(start.column, end.column)
  };
}

function worksheetFormulaRemovalIssue(workbook, worksheet) {
  const sharedRanges = new Map();
  let issue = null;

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type !== ExcelJS.ValueType.Formula || issue) return;
      const value = cell.value;
      if (!value || typeof value !== "object") {
        issue = "A formula has dependencies that cannot be safely determined.";
        return;
      }

      let formulaValue = value;
      if (typeof value.sharedFormula === "string") {
        const masterCell = worksheet.getCell(value.sharedFormula);
        if (
          masterCell.type !== ExcelJS.ValueType.Formula ||
          typeof masterCell.value?.formula !== "string" ||
          masterCell.value.shareType !== "shared"
        ) {
          issue = "A shared formula dependency cannot be safely determined.";
          return;
        }
        formulaValue = masterCell.value;
        const range = sharedFormulaRange(worksheet, masterCell);
        if (
          !range ||
          cell.row < range.firstRow ||
          cell.row > range.lastRow ||
          cell.col < range.firstColumn ||
          cell.col > range.lastColumn
        ) {
          issue = "A shared formula range cannot be safely determined.";
          return;
        }
        sharedRanges.set(range.master, range);
      } else if (
        value.shareType !== undefined &&
        value.shareType !== "shared"
      ) {
        issue = "A formula uses a shared or array range that cannot be safely adjusted.";
        return;
      } else if (value.shareType === "shared") {
        const range = sharedFormulaRange(worksheet, cell);
        if (!range) {
          issue = "A shared formula range cannot be safely determined.";
          return;
        }
        sharedRanges.set(range.master, range);
      }

      if (!isRowInvariantFormula(formulaValue.formula)) {
        issue = "A formula contains references or syntax whose row dependencies cannot be safely determined.";
      }
    });
  });
  if (issue) return { issue, sharedRanges: [...sharedRanges.values()] };

  let crossSheetDependency = false;
  for (const formulaSheet of workbook.worksheets) {
    if (formulaSheet === worksheet) continue;
    formulaSheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (
          cell.type !== ExcelJS.ValueType.Formula ||
          !cell.value ||
          typeof cell.value !== "object"
        ) {
          return;
        }
        const formula = cell.value.formula ?? cell.value.sharedFormula ?? "";
        crossSheetDependency ||= formulaReferencesWorksheet(formula, worksheet.name) ||
          /\bINDIRECT\s*\(/i.test(formula);
      });
    });
  }
  if (crossSheetDependency) {
    return {
      issue: "A formula on another worksheet may depend on this worksheet's row positions.",
      sharedRanges: [...sharedRanges.values()]
    };
  }

  return { issue: null, sharedRanges: [...sharedRanges.values()] };
}

function worksheetDuplicateRemovalIssue(workbook, worksheet) {
  const formulaSafety = worksheetFormulaRemovalIssue(workbook, worksheet);
  if (formulaSafety.issue) return { issue: formulaSafety.issue, sharedRanges: formulaSafety.sharedRanges };
  if (workbook.definedNames?.model?.length) {
    return { issue: "The workbook contains defined names whose row references cannot be safely verified.", sharedRanges: formulaSafety.sharedRanges };
  }
  if (worksheetHasMergedRanges(worksheet)) {
    return { issue: "This worksheet contains merged cells.", sharedRanges: formulaSafety.sharedRanges };
  }
  if (Object.keys(worksheet.tables || {}).length > 0) {
    return { issue: "This worksheet contains a table.", sharedRanges: formulaSafety.sharedRanges };
  }
  if (worksheet.autoFilter || Object.keys(worksheet.dataValidations?.model || {}).length) {
    return { issue: "This worksheet contains filters or data validation tied to row positions.", sharedRanges: formulaSafety.sharedRanges };
  }
  if (worksheet.conditionalFormattings?.length || worksheet.rowBreaks?.length) {
    return { issue: "This worksheet contains formatting or page-break rules tied to row positions.", sharedRanges: formulaSafety.sharedRanges };
  }
  if (worksheet.getImages?.().length) {
    return { issue: "This worksheet contains images that may be tied to row positions.", sharedRanges: formulaSafety.sharedRanges };
  }
  return { issue: null, sharedRanges: formulaSafety.sharedRanges };
}

function rowStructureSignature(row, columnCount) {
  const cells = [];
  for (let column = 1; column <= columnCount; column++) {
    const cell = row.getCell(column);
    cells.push({
      hyperlink: cell.hyperlink || null,
      note: cell.note || null,
      dataValidation: cell.dataValidation || null
    });
  }
  return JSON.stringify({
    hidden: Boolean(row.hidden),
    outlineLevel: row.outlineLevel || 0,
    collapsed: Boolean(row.collapsed),
    cells
  });
}

function duplicateCandidateId(sourceHash, worksheetKey, rowNumber, keeperRow, key) {
  return createHash("sha256")
    .update(JSON.stringify([
      "insightmate-cleaning-candidate-v1",
      sourceHash,
      worksheetKey,
      rowNumber,
      keeperRow,
      key
    ]))
    .digest("hex");
}

function dateFormatDisplayWidth(numberFormat) {
  if (typeof numberFormat !== "string" || !numberFormat) return null;
  if (numberFormat.includes("F800")) return 30;

  // Common Excel date/time masks are estimated from their tokens. Unusual
  // locale and conditional masks use the maximum width rather than under-size.
  if (/\[\$(?!-F800)|\[(?!h+\]|m+\]|s+\]|black\]|blue\]|cyan\]|green\]|magenta\]|red\]|white\]|yellow\]|[<>=])/i.test(numberFormat)) {
    return /[ymdhs]/i.test(numberFormat) ? 48 : null;
  }

  let section = "";
  let quoted = false;
  let escaped = false;
  for (const character of numberFormat) {
    if (escaped) {
      section += `\\${character}`;
      escaped = false;
    } else if (character === "\\") {
      escaped = true;
    } else if (character === "\"") {
      quoted = !quoted;
      section += character;
    } else if (character === ";" && !quoted) {
      break;
    } else {
      section += character;
    }
  }

  const tokens = [];
  let lastDateToken = null;
  let width = 0;
  for (let index = 0; index < section.length;) {
    const character = section[index];
    if (character === "\"") {
      index++;
      while (index < section.length && section[index] !== "\"") {
        width++;
        index++;
      }
      index++;
      continue;
    }
    if (character === "\\" || character === "_" || character === "*") {
      if (index + 1 < section.length) {
        if (character !== "\\") width++;
        else width++;
        index += 2;
      } else {
        index++;
      }
      continue;
    }
    if (character === "[") {
      const closeIndex = section.indexOf("]", index + 1);
      if (closeIndex < 0) {
        width++;
        index++;
        continue;
      }
      const bracketToken = section.slice(index + 1, closeIndex);
      const elapsed = bracketToken.match(/^(h+|m+|s+)$/i);
      if (elapsed) {
        width += 2;
        tokens.push(elapsed[1][0].toLowerCase());
        lastDateToken = elapsed[1][0].toLowerCase();
      }
      index = closeIndex + 1;
      continue;
    }
    const amPm = section.slice(index).match(/^(AM\/PM|A\/P)/i);
    if (amPm) {
      width += amPm[1].toUpperCase() === "AM/PM" ? 2 : 1;
      lastDateToken = "h";
      index += amPm[1].length;
      continue;
    }
    if (/[ymdhs]/i.test(character)) {
      let end = index + 1;
      while (end < section.length &&
        section[end].toLowerCase() === character.toLowerCase()) {
        end++;
      }
      const token = character.toLowerCase();
      const tokenLength = end - index;
      if (token === "y") {
        width += Math.max(4, tokenLength);
      } else if (token === "d") {
        width += tokenLength >= 4 ? 9 : tokenLength === 3 ? 3 : 2;
      } else if (token === "m") {
        let nextToken = null;
        for (let lookahead = end; lookahead < section.length; lookahead++) {
          if (/[ymdhs]/i.test(section[lookahead])) {
            nextToken = section[lookahead].toLowerCase();
            break;
          }
          if (section[lookahead] === "\"") break;
        }
        const isMinute = lastDateToken === "h" || nextToken === "s";
        width += isMinute
          ? 2
          : tokenLength >= 4 ? 9 : tokenLength === 3 ? 3 : 2;
      } else {
        width += tokenLength === 1 ? 2 : tokenLength;
      }
      lastDateToken = token;
      tokens.push(token);
      index = end;
      continue;
    }
    width++;
    index++;
  }

  const hasDateToken = tokens.some((token) => ["y", "d"].includes(token)) ||
    (tokens.includes("m") && (tokens.includes("h") || tokens.includes("s")));
  const hasTimeToken = tokens.some((token) => ["h", "s"].includes(token));
  if (!hasDateToken && !hasTimeToken) return null;
  return width;
}

function cellDisplayText(cell) {
  const value = cell.value;
  const isDateValue = value instanceof Date ||
    cell.type === ExcelJS.ValueType.Formula &&
      value && typeof value === "object" &&
      (value.result instanceof Date ||
        typeof value.result === "number" &&
          dateFormatDisplayWidth(cell.numFmt) !== null);
  if (isDateValue) {
    const formattedWidth = dateFormatDisplayWidth(cell.numFmt);
    const estimatedWidth = formattedWidth ?? 10;
    return "8".repeat(estimatedWidth);
  }
  if (value && typeof value === "object") {
    if (cell.type === ExcelJS.ValueType.Formula) {
      return String(value.result ?? value.formula ?? "");
    }
    if (Array.isArray(value.richText)) {
      return value.richText.map((entry) => entry.text || "").join("");
    }
    if (value.text) return String(value.text);
    return "";
  }
  return String(value ?? "");
}

function formatWorksheet(worksheet, headerRowNumber) {
  const formattedRows = new Set();
  const headerRow = headerRowNumber === null
    ? null
    : worksheet.getRow(headerRowNumber);

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    let rowFormatted = false;
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!hasMeaningfulCellValue(cell.value)) return;
      cell.font = {
        ...cell.font,
        name: cell.font?.name || "Arial",
        size: cell.font?.size || 11
      };
      cell.alignment = {
        ...cell.alignment,
        horizontal: "center",
        vertical: "middle"
      };

      if (rowNumber === headerRowNumber) {
        cell.font = {
          ...cell.font,
          bold: true,
          color: { argb: "FFFFFFFF" }
        };
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF1F4E78" }
        };
        cell.alignment = {
          ...cell.alignment,
          horizontal: "center",
          vertical: "middle",
          wrapText: true
        };
      }

      const subtleBorder = {
        style: "thin",
        color: { argb: "FFD9E0E8" }
      };
      cell.border = {
        ...cell.border,
        top: cell.border?.top || subtleBorder,
        right: cell.border?.right || subtleBorder,
        bottom: cell.border?.bottom || subtleBorder,
        left: cell.border?.left || subtleBorder
      };
      rowFormatted = true;
    });
    if (rowFormatted) formattedRows.add(rowNumber);
  });

  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber++) {
    const column = worksheet.getColumn(columnNumber);
    let maxLength = 0;
    let hasContent = false;
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      const cell = row.getCell(columnNumber);
      if (!hasMeaningfulCellValue(cell.value)) return;
      hasContent = true;
      maxLength = Math.max(maxLength, cellDisplayText(cell).length);
    });
    if (!hasContent) continue;

    const existingWidth = Number.isFinite(column.width) && column.width > 0
      ? column.width
      : 0;
    column.width = Math.min(48, Math.max(10, existingWidth, maxLength + 2));
  }

  for (const rowNumber of formattedRows) {
    const row = worksheet.getRow(rowNumber);
    let wrappedLines = 1;
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      if (!hasMeaningfulCellValue(cell.value)) return;
      const width = worksheet.getColumn(columnNumber).width || 10;
      const cellLines = Math.max(
        1,
        Math.ceil(cellDisplayText(cell).length / Math.max(1, width - 2))
      );
      wrappedLines = Math.max(wrappedLines, cellLines);
      if (rowNumber !== headerRowNumber && cellLines > 1) {
        cell.alignment = { ...cell.alignment, wrapText: true };
      }
    });

    const minimumHeight = rowNumber === headerRowNumber
      ? Math.max(26, wrappedLines * 16 + 8)
      : Math.max(20, wrappedLines * 15 + 5);
    row.height = Math.max(row.height || 0, minimumHeight);
  }

  if (headerRow && !formattedRows.has(headerRowNumber)) {
    headerRow.height = Math.max(headerRow.height || 0, 26);
  }

  return formattedRows.size;
}

function centerPopulatedCells(worksheet) {
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (!hasMeaningfulCellValue(cell.value)) return;
      cell.alignment = {
        ...cell.alignment,
        horizontal: "center",
        vertical: "middle"
      };
    });
  });
}

function assertCenteredSerializedWorkbook(workbook) {
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (
          hasMeaningfulCellValue(cell.value) &&
          (cell.alignment?.horizontal !== "center" ||
            cell.alignment?.vertical !== "middle")
        ) {
          throw new Error(
            `The serialized alignment for ${worksheet.name}!${cell.address} was not preserved.`
          );
        }
      });
    });
  }
}

function assertSerializedMergedRanges(workbook, expectedRanges) {
  if (workbook.worksheets.length !== expectedRanges.length) {
    throw new Error("The serialized workbook did not preserve the worksheet merge layout.");
  }
  for (const [index, worksheet] of workbook.worksheets.entries()) {
    const expected = expectedRanges[index];
    const actualRanges = [...(worksheet.model.merges || [])].sort();
    if (
      worksheet.name !== expected.name ||
      JSON.stringify(actualRanges) !== JSON.stringify(expected.ranges)
    ) {
      throw new Error(
        `The serialized merged ranges for worksheet ${expected.name} did not match the cleaned workbook.`
      );
    }
  }
}

function columnName(columnNumber) {
  let name = "";
  let remaining = columnNumber;
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return name;
}

export function validateHeaderRow(sheet, rowNumber) {
  if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > sheet.rowCount) {
    return false;
  }

  return sheet.cells.some((cell) => cell.row === rowNumber && !isBlank(cell.value));
}

export async function cleanWorkbook(workbookState) {
  if (!workbookState || typeof workbookState.getSourceBytes !== "function") {
    throw new TypeError("A workbook state with original source bytes is required.");
  }
  const originalBytes = workbookState.getSourceBytes();
  if (!Buffer.isBuffer(originalBytes) || originalBytes.length === 0) {
    throw new TypeError("Original workbook source must be a non-empty Buffer.");
  }
  const sourceHash = createHash("sha256").update(originalBytes).digest("hex");

  const cleanedWorkbook = new ExcelJS.Workbook();
  await cleanedWorkbook.xlsx.load(Buffer.from(originalBytes));
  if (cleanedWorkbook.worksheets.length === 0) {
    throw new Error("The uploaded workbook contains no worksheets.");
  }

  const worksheetSummaries = [];

  for (const worksheet of cleanedWorkbook.worksheets) {
    const whitespaceOnlyRows = new Set();
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      if (blankRowValueStatus(
        worksheet.getRow(rowNumber),
        worksheet.columnCount
      ).hasWhitespace) {
        whitespaceOnlyRows.add(rowNumber);
      }
    }
    const initialHeader = getHeaderCandidate(worksheet);
    const headerRowNumber = initialHeader.status === "candidate"
      ? initialHeader.candidateRow
      : null;
    const headerColumnCount = headerRowNumber === null
      ? 0
      : getHeaderColumnCount(worksheet.getRow(headerRowNumber));
    const headerLabels = new Set();
    if (headerRowNumber !== null) {
      for (let column = 1; column <= headerColumnCount; column++) {
        const value = worksheet.getCell(headerRowNumber, column).value;
        if (typeof value === "string" && value.trim()) {
          headerLabels.add(value.trim().replace(/ +/g, " ").toLocaleLowerCase("en-US"));
        }
      }
    }
    const nameColumns = personNameColumns(worksheet, headerRowNumber);
    const summary = {
      name: worksheet.name,
      order: worksheetSummaries.length,
      visibility: worksheet.state,
      rowsFormatted: 0,
      spacesTrimmed: 0,
      emptyRowsRemoved: 0,
      emptyRowsSkipped: [],
      blankRowsDetected: 0,
      blankRowsRemoved: 0,
      blankRowsRetained: 0,
      blankRowsRetainedDetails: [],
      styledEmptyRowsRetained: [],
      blankCellsDetected: 0,
      duplicateCandidatesFound: 0,
      duplicateCandidates: [],
      duplicateRowsSkipped: [],
      errorReplacements: [],
      removedRows: 0,
      mergedRangeActions: [],
      ambiguousItemsLeftUnchanged: 0,
      ambiguousItems: [],
      itemsSkippedForSafety: 0,
      limitations: [],
      explanation: ""
    };
    summary.mergedRangeActions = normalizeMergedDataFields(
      worksheet,
      headerRowNumber
    );

    if (headerRowNumber === null) {
      summary.limitations.push(
        "A clear header row could not be identified, so missing values and duplicate rows were not assessed."
      );
    }

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        if (cell.type === ExcelJS.ValueType.String && typeof cell.value === "string") {
          const normalized = normalizeTextValue(
            cell.value,
            cell,
            rowNumber === headerRowNumber ? null : nameColumns.get(columnNumber)
          );
          if (normalized !== cell.value) {
            cell.value = normalized;
            summary.spacesTrimmed++;
          }
          return;
        }
        if (cell.isMerged) {
          summary.itemsSkippedForSafety++;
        }
      });
    });

    const eligibleRows = [];
    const duplicateKeepers = new Map();
    const duplicateRemovalSafety = worksheetDuplicateRemovalIssue(
      cleanedWorkbook,
      worksheet
    );
    if (headerRowNumber !== null && headerColumnCount > 0) {
      for (let rowNumber = headerRowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        const duplicateCells = Array.from(
          { length: worksheet.columnCount },
          (_, index) => row.getCell(index + 1)
        );
        const cells = Array.from(
          { length: headerColumnCount },
          (_, index) => row.getCell(index + 1)
        );
        if (!duplicateCells.some((cell) => hasMeaningfulCellValue(cell.value))) continue;

        const key = JSON.stringify(duplicateCells.map((cell, index) =>
          serializeCellValue(
            normalizeTextValue(
              cell.value,
              cell,
              nameColumns.get(index + 1)
            ),
            cell.type
          )
        ));
        const keeper = duplicateKeepers.get(key);
        if (keeper === undefined) {
          duplicateKeepers.set(key, { rowNumber, row });
        } else {
          summary.duplicateCandidatesFound++;
          let unsafeReason = null;
          const rowSafetyIssue = (candidateRow, description) => {
            let issue = null;
            candidateRow.eachCell({ includeEmpty: true }, (cell) => {
              if (cell.type === ExcelJS.ValueType.Formula) {
                issue = `The ${description} row contains a formula.`;
              } else if (cell.isMerged && !issue) {
                issue = `The ${description} row contains merged cells.`;
              }
            });
            return issue;
          };
          unsafeReason =
            rowSafetyIssue(row, "duplicate") ||
            rowSafetyIssue(keeper.row, "matching");
          if (!unsafeReason && duplicateRemovalSafety.issue) {
            unsafeReason = duplicateRemovalSafety.issue;
          }
          if (
            !unsafeReason &&
            duplicateRemovalSafety.sharedRanges.some((range) =>
              rowNumber <= range.lastRow
            )
          ) {
            unsafeReason = "Removing this row could shift a shared formula range.";
          }
          if (
            !unsafeReason &&
            rowStructureSignature(row, worksheet.columnCount) !==
              rowStructureSignature(keeper.row, worksheet.columnCount)
          ) {
            unsafeReason = "The duplicate row has different formatting or cell features.";
          }
          summary.duplicateCandidates.push({
            candidateId: duplicateCandidateId(
              sourceHash,
              `sheet-${summary.order}`,
              rowNumber,
              keeper.rowNumber,
              key
            ),
            originalRow: rowNumber,
            retainedOriginalRow: keeper.rowNumber,
            cleanedRow: null,
            retainedCleanedRow: null,
            safeToRemove: !unsafeReason,
            unsafeReason
          });
          if (unsafeReason) {
            summary.itemsSkippedForSafety++;
            summary.duplicateRowsSkipped.push({
              originalRow: rowNumber,
              retainedOriginalRow: keeper.rowNumber,
              reason: unsafeReason
            });
          }
        }
        eligibleRows.push({ rowNumber, cells });
      }
    }

    const rowsWithBlankedErrors = new Set();
    worksheet.eachRow({ includeEmpty: false }, (row) => {
        row.eachCell({ includeEmpty: false }, (cell) => {
          const recognizedError = getRecognizedCellError(cell);
          if (!recognizedError) return;
          rowsWithBlankedErrors.add(cell.row);
          summary.errorReplacements.push({
            worksheetName: worksheet.name,
            address: cell.address,
            originalError: recognizedError.error,
            originalFormula: recognizedError.formula,
            action: recognizedError.formula
              ? "Replaced cached formula error with a blank cell"
              : "Replaced Excel error with a blank cell"
          });
          cell.value = null;
        });
    });

    if (headerRowNumber !== null && headerColumnCount > 0) {
      const lastOriginalRow = worksheet.rowCount;
      for (let rowNumber = headerRowNumber + 1; rowNumber <= lastOriginalRow; rowNumber++) {
        const row = worksheet.getRow(rowNumber);
        if (!rowHasMeaningfulValues(row, headerColumnCount)) continue;
        for (let columnNumber = 1; columnNumber <= headerColumnCount; columnNumber++) {
          const cell = row.getCell(columnNumber);
          if (!cell.isMerged && !hasMeaningfulCellValue(cell.value)) {
            summary.blankCellsDetected++;
          }
        }
      }

      for (const { rowNumber, cells } of eligibleRows) {
        for (const cell of cells) {
          if (getAmbiguousDateText(cell.value)) {
            summary.ambiguousItemsLeftUnchanged++;
            summary.ambiguousItems.push({
              originalRow: rowNumber,
              address: cell.address,
              value: cell.value
            });
          }
        }
      }
    }

    const emptyRows = [];
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      if (blankRowValueStatus(row, worksheet.columnCount).isBlank) {
        emptyRows.push(rowNumber);
      }
    }
    summary.blankRowsDetected = emptyRows.length;
    let lastPopulatedRow = 0;
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      if (!blankRowValueStatus(
        worksheet.getRow(rowNumber),
        worksheet.columnCount
      ).isBlank) {
        lastPopulatedRow = rowNumber;
      }
    }
    const rowRemovalSafety = emptyRows.length
      ? worksheetDuplicateRemovalIssue(cleanedWorkbook, worksheet)
      : { issue: null, sharedRanges: [] };
    const safelyRemovedEmptyRows = [];
    for (const originalRow of emptyRows) {
      const row = worksheet.getRow(originalRow);
      const hasStyledCells = Array.from(
        { length: worksheet.columnCount },
        (_, index) => row.getCell(index + 1)
      ).some((cell) =>
        Object.keys(cell.style || {}).length > 0
      );
      let reason = null;
      if (whitespaceOnlyRows.has(originalRow)) {
        reason = "The row contained whitespace-only values before text cleanup.";
      } else if (rowsWithBlankedErrors.has(originalRow)) {
        reason = "The row contained recognized Excel errors that were blanked.";
      } else {
        reason = blankRowStructureIssue(row, worksheet.columnCount);
        if (!reason && rowRemovalSafety.issue) {
          reason = rowRemovalSafety.issue;
        }
        if (
          !reason &&
          rowRemovalSafety.sharedRanges.some((range) =>
            originalRow <= range.lastRow
          )
        ) {
          reason = "Removing this row could shift a shared formula range.";
        }
        if (!reason && lastPopulatedRow === 0) {
          reason = "The worksheet has no populated rows, so the blank row's purpose is unclear.";
        }
        if (!reason && originalRow <= lastPopulatedRow) {
          let previousPopulatedRow = originalRow - 1;
          while (
            previousPopulatedRow > 0 &&
            blankRowValueStatus(
              worksheet.getRow(previousPopulatedRow),
              worksheet.columnCount
            ).isBlank
          ) {
            previousPopulatedRow--;
          }
          let nextPopulatedRow = originalRow + 1;
          while (
            nextPopulatedRow <= worksheet.rowCount &&
            blankRowValueStatus(
              worksheet.getRow(nextPopulatedRow),
              worksheet.columnCount
            ).isBlank
          ) {
            nextPopulatedRow++;
          }
          const previousRowIssue = ordinaryDataRowIssue(
            worksheet,
            previousPopulatedRow,
            headerRowNumber,
            headerColumnCount,
            headerLabels
          );
          const nextRowIssue = ordinaryDataRowIssue(
            worksheet,
            nextPopulatedRow,
            headerRowNumber,
            headerColumnCount,
            headerLabels
          );
          if (previousRowIssue || nextRowIssue) {
            const neighborIssue = previousRowIssue || nextRowIssue;
            reason = `A blank row was retained because ${neighborIssue[0].toLocaleLowerCase()}${neighborIssue.slice(1)}`;
          }
        }
      }

      if (reason) {
        summary.itemsSkippedForSafety++;
        const retainedRow = { originalRow, reason };
        summary.emptyRowsSkipped.push(retainedRow);
        summary.blankRowsRetainedDetails.push(retainedRow);
        if (hasStyledCells) {
          summary.styledEmptyRowsRetained.push(retainedRow);
        }
      } else {
        safelyRemovedEmptyRows.push(originalRow);
      }
    }
    for (const rowNumber of safelyRemovedEmptyRows.sort((left, right) => right - left)) {
      worksheet.spliceRows(rowNumber, 1);
    }
    summary.blankRowsRemoved = safelyRemovedEmptyRows.length;
    summary.emptyRowsRemoved = summary.blankRowsRemoved;
    summary.blankRowsRetained = summary.blankRowsRetainedDetails.length;
    const retainedReasons = new Map();
    for (const { reason } of summary.blankRowsRetainedDetails) {
      retainedReasons.set(reason, (retainedReasons.get(reason) || 0) + 1);
    }
    summary.blankRowsRetainedByReason = [...retainedReasons].map(
      ([reason, count]) => ({ reason, count })
    );

    if (emptyRows.length) {
      for (const candidate of summary.duplicateCandidates) {
        candidate.cleanedRow = candidate.originalRow -
          safelyRemovedEmptyRows.filter((rowNumber) => rowNumber < candidate.originalRow).length;
        candidate.retainedCleanedRow = candidate.retainedOriginalRow -
          safelyRemovedEmptyRows.filter((rowNumber) => rowNumber < candidate.retainedOriginalRow).length;
      }
    } else {
      for (const candidate of summary.duplicateCandidates) {
        candidate.cleanedRow = candidate.originalRow;
        candidate.retainedCleanedRow = candidate.retainedOriginalRow;
      }
    }

    const removedBeforeHeader = safelyRemovedEmptyRows.filter(
      (rowNumber) => rowNumber < headerRowNumber
    ).length;
    const adjustedHeaderRow = headerRowNumber === null
      ? null
      : headerRowNumber - removedBeforeHeader;
    summary.rowsFormatted = formatWorksheet(worksheet, adjustedHeaderRow);

    const operations = [];
    if (summary.spacesTrimmed) operations.push(`tidied text in ${summary.spacesTrimmed} cells`);
    if (summary.blankRowsRemoved) operations.push(`removed ${summary.blankRowsRemoved} safe blank rows`);
    if (summary.duplicateCandidatesFound) {
      operations.push(`found ${summary.duplicateCandidatesFound} duplicate rows for automatic removal`);
    }
    if (summary.blankCellsDetected) operations.push(`found ${summary.blankCellsDetected} blank cells`);
    if (summary.ambiguousItemsLeftUnchanged) operations.push(`left ${summary.ambiguousItemsLeftUnchanged} unclear date values unchanged`);
    if (summary.errorReplacements.length) operations.push(`replaced ${summary.errorReplacements.length} Excel errors with blank cells`);
    if (summary.blankRowsRetained) {
      operations.push(`retained ${summary.blankRowsRetained} blank rows because their purpose or structure could not be safely dismissed`);
    }
    if (!operations.length) operations.push("found no safe data changes to make");
    summary.explanation = `Formatted ${summary.rowsFormatted} rows and ${operations.join(", ")}. Missing values were not filled.`;
    if (summary.limitations.length) summary.explanation += ` ${summary.limitations.join(" ")}`;

    worksheetSummaries.push(summary);
  }

  return {
    workbook: cleanedWorkbook,
    summary: {
      worksheetCount: worksheetSummaries.length,
      worksheets: worksheetSummaries,
      explanation: `Processed ${worksheetSummaries.length} worksheets. Duplicate rows are removed automatically when safe; unsafe rows are retained and reported.`
    }
  };
}

function rowValueSignature(row, columnCount) {
  return JSON.stringify(Array.from(
    { length: columnCount },
    (_, index) => {
      const cell = row.getCell(index + 1);
      return serializeCellValue(cell.value, cell.type);
    }
  ));
}

function countRowSignatures(worksheet) {
  const counts = new Map();
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const signature = rowValueSignature(row, worksheet.columnCount);
    counts.set(signature, (counts.get(signature) || 0) + 1);
  });
  return counts;
}

function workbookFormulaSignatures(workbook) {
  const signatures = [];
  for (const worksheet of workbook.worksheets) {
    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (cell.type !== ExcelJS.ValueType.Formula) return;
        signatures.push(JSON.stringify([
          worksheet.name,
          serializeCellValue(cell.value, cell.type),
          cell.value?.shareType ?? null,
          cell.value?.ref ?? null
        ]));
      });
    });
  }
  return signatures.sort();
}

export async function generateCleanedWorkbook(workbookState) {
  const { workbook, summary } = await cleanWorkbook(workbookState);
  const formulaSignaturesBeforeRemoval = workbookFormulaSignatures(workbook);
  const workbooksByWorksheetKey = new Map(
    workbook.worksheets.map((worksheet, order) => [`sheet-${order}`, worksheet])
  );
  const automaticRemovals = [];
  const signaturesBeforeRemoval = new Map();

  for (const worksheetSummary of summary.worksheets) {
    const worksheetKey = `sheet-${worksheetSummary.order}`;
    const worksheet = workbooksByWorksheetKey.get(worksheetKey);
    const rowSignatures = countRowSignatures(worksheet);
    signaturesBeforeRemoval.set(worksheetKey, rowSignatures);

    for (const candidate of worksheetSummary.duplicateCandidates) {
      if (!candidate.safeToRemove) {
        continue;
      }

      automaticRemovals.push({
        worksheetKey,
        worksheet,
        candidate,
        rowSignature: rowValueSignature(
          worksheet.getRow(candidate.cleanedRow),
          worksheet.columnCount
        )
      });
    }
  }

  const byWorksheet = new Map();
  for (const item of automaticRemovals) {
    if (!byWorksheet.has(item.worksheetKey)) byWorksheet.set(item.worksheetKey, []);
    byWorksheet.get(item.worksheetKey).push(item);
  }

  for (const items of byWorksheet.values()) {
    items.sort((left, right) =>
      right.candidate.cleanedRow - left.candidate.cleanedRow
    );
    for (const { worksheet, candidate } of items) {
      worksheet.spliceRows(candidate.cleanedRow, 1);
    }
  }

  for (const worksheet of workbook.worksheets) {
    centerPopulatedCells(worksheet);
  }

  const expectedMergedRanges = workbook.worksheets.map((worksheet) => ({
    name: worksheet.name,
    ranges: [...(worksheet.model.merges || [])].sort()
  }));
  const outputBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  if (outputBytes.length === 0) {
    throw new Error("The cleaned workbook could not be serialized.");
  }

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(Buffer.from(outputBytes));
  assertCenteredSerializedWorkbook(reopened);
  assertSerializedMergedRanges(reopened, expectedMergedRanges);
  if (
    JSON.stringify(workbookFormulaSignatures(reopened)) !==
    JSON.stringify(formulaSignaturesBeforeRemoval)
  ) {
    throw new Error("Duplicate removal changed formula expressions or cached results.");
  }
  if (
    reopened.worksheets.length !== workbook.worksheets.length ||
    reopened.worksheets.some((worksheet, index) =>
      worksheet.name !== workbook.worksheets[index].name ||
      worksheet.state !== workbook.worksheets[index].state
    )
  ) {
    throw new Error("The cleaned workbook did not preserve worksheet order, names, and visibility.");
  }

  const automaticSignatureCounts = new Map();
  for (const { worksheetKey, rowSignature } of automaticRemovals) {
    if (!automaticSignatureCounts.has(worksheetKey)) {
      automaticSignatureCounts.set(worksheetKey, new Map());
    }
    const counts = automaticSignatureCounts.get(worksheetKey);
    counts.set(rowSignature, (counts.get(rowSignature) || 0) + 1);
  }

  for (const worksheetSummary of summary.worksheets) {
    const worksheetKey = `sheet-${worksheetSummary.order}`;
    const beforeCounts = signaturesBeforeRemoval.get(worksheetKey);
    const afterCounts = countRowSignatures(reopened.worksheets[worksheetSummary.order]);
    const expectedRemovals = automaticSignatureCounts.get(worksheetKey) || new Map();
    const expectedRemovedCount = [...expectedRemovals.values()]
      .reduce((total, count) => total + count, 0);
    const rowsBefore = [...beforeCounts.values()]
      .reduce((total, count) => total + count, 0);
    const rowsAfter = [...afterCounts.values()]
      .reduce((total, count) => total + count, 0);
    if (rowsBefore - rowsAfter !== expectedRemovedCount) {
      throw new Error("The cleaned workbook row count did not match automatic duplicate removals.");
    }
    for (const [signature, expectedRemoved] of expectedRemovals) {
      const actualRemoved = (beforeCounts.get(signature) || 0) -
        (afterCounts.get(signature) || 0);
      if (actualRemoved !== expectedRemoved) {
        throw new Error("A detected duplicate row was not removed from the serialized workbook.");
      }
    }
    worksheetSummary.removedRows = expectedRemovedCount;

  }

  summary.explanation = `Processed ${summary.worksheetCount} worksheets. Duplicate rows were removed automatically where safe; unsafe duplicates were retained and reported.`;
  return {
    buffer: outputBytes,
    summary
  };
}

export function generateCleaningReview(workbookState, uploadVersionId, headerRows, existingReview = null) {
  const normalizedHeaders = workbookState.worksheets.map((sheet) => ({
    worksheetKey: sheet.worksheetKey,
    selection: headerRows[sheet.worksheetKey] ?? null
  }));
  const configuration = JSON.stringify(normalizedHeaders);
  if (existingReview?.configuration === configuration) {
    return existingReview;
  }

  const reviewId = randomUUID();
  const proposals = [];
  const findings = [];
  const worksheetResults = [];

  for (const sheet of workbookState.worksheets) {
    const selection = headerRows[sheet.worksheetKey] ?? null;
    if (selection?.action === "skip") {
      worksheetResults.push({
        worksheetKey: sheet.worksheetKey,
        name: sheet.name,
        status: "skipped",
        reason: "The user skipped this worksheet."
      });
      continue;
    }
    const selectedHeader = typeof selection === "number"
      ? selection
      : selection?.row ?? null;
    const headerRow = selectedHeader ??
      (sheet.headerCandidate.status === "candidate" ? sheet.headerCandidate.candidateRow : null);

    if (headerRow === null) {
      worksheetResults.push({
        worksheetKey: sheet.worksheetKey,
        name: sheet.name,
        status: "skipped",
        reason: sheet.headerCandidate.reason || "Header row must be confirmed before proposals can be generated."
      });
      continue;
    }

    if (!validateHeaderRow(sheet, headerRow)) {
      throw new RangeError(`Invalid header row ${headerRow} for worksheet ${sheet.name}.`);
    }
    const nameColumns = snapshotPersonNameColumns(sheet, headerRow);

    worksheetResults.push({
      worksheetKey: sheet.worksheetKey,
      name: sheet.name,
      status: "processed",
      headerRow
    });

    const dataRows = Array.from(
      { length: Math.max(0, sheet.rowCount - headerRow) },
      (_, index) => headerRow + index + 1
    );
    const nonEmptyRows = [];

    for (const rowNumber of dataRows) {
      const cells = Array.from({ length: sheet.columnCount }, (_, index) =>
        cellFor(sheet, rowNumber, index + 1)
      );
      const values = cells.map((cell) => cell?.value ?? null);
      const completelyEmpty = values.every(isBlank);

      if (completelyEmpty) {
        const target = { kind: "row", row: rowNumber };
        proposals.push({
          proposalId: proposalId(uploadVersionId, reviewId, sheet.worksheetKey, target, "empty_row"),
          uploadVersionId,
          reviewId,
          worksheet: { worksheetKey: sheet.worksheetKey, name: sheet.name, order: sheet.order },
          target,
          issueType: "empty_row",
          severity: "low",
          changeKind: "row_removal",
          reason: "Every cell in this original row is empty.",
          decision: "pending",
          createdAt: new Date().toISOString(),
          decidedAt: null
        });
        continue;
      }

      const containsFormulaOrMerge = cells.some((cell) =>
        cell && (cell.type === ExcelJS.ValueType.Formula || cell.isMerged)
      );
      nonEmptyRows.push({
        rowNumber,
        cells,
        values,
        eligibleForDuplicates: !containsFormulaOrMerge
      });

      for (let index = 0; index < cells.length; index++) {
        const cell = cells[index];
        const value = cell?.value;
        if (typeof value === "string" && value.trim() !== value) {
          const target = {
            kind: "cell",
            address: cell?.address || `${columnName(index + 1)}${rowNumber}`,
            row: rowNumber,
            column: index + 1
          };
          const originalValue = serializeCellValue(value, cell.type);
          const proposedValue = serializeCellValue(value.trim(), ExcelJS.ValueType.String);
          proposals.push({
            proposalId: proposalId(uploadVersionId, reviewId, sheet.worksheetKey, target, "leading_trailing_whitespace"),
            uploadVersionId,
            reviewId,
            worksheet: { worksheetKey: sheet.worksheetKey, name: sheet.name, order: sheet.order },
            target,
            issueType: "leading_trailing_whitespace",
            severity: "low",
            originalValue,
            proposedValue,
            reason: "Leading or trailing whitespace was detected.",
            changeKind: "data_correction",
            decision: "pending",
            createdAt: new Date().toISOString(),
            decidedAt: null
          });
        } else if (isBlank(value) && cells.some((entry) => !isBlank(entry?.value))) {
          findings.push({
            worksheetKey: sheet.worksheetKey,
            worksheetName: sheet.name,
            issueType: "missing_value",
            target: {
              kind: "cell",
              address: cell?.address || `${columnName(index + 1)}${rowNumber}`,
              row: rowNumber,
              column: index + 1
            },
            reason: "The cell is blank; no replacement value is inferred."
          });
        }
      }
    }

    const seen = new Map();
    for (const row of nonEmptyRows) {
      if (!row.eligibleForDuplicates) continue;
      const key = JSON.stringify(row.values.map((value, index) =>
        serializeCellValue(
          normalizeTextValue(
            value,
            row.cells[index],
            nameColumns.get(index + 1)
          ),
          row.cells[index]?.type
        )
      ));
      const keeperRow = seen.get(key);
      if (keeperRow === undefined) {
        seen.set(key, row.rowNumber);
        continue;
      }

      const target = { kind: "row", row: row.rowNumber };
      proposals.push({
        proposalId: proposalId(uploadVersionId, reviewId, sheet.worksheetKey, target, "exact_duplicate_row"),
        uploadVersionId,
        reviewId,
        worksheet: { worksheetKey: sheet.worksheetKey, name: sheet.name, order: sheet.order },
        target,
        issueType: "exact_duplicate_row",
        severity: "medium",
        changeKind: "row_removal",
        matches: { retainedRow: keeperRow },
        reason: `This row exactly matches the earliest eligible row ${keeperRow}.`,
        decision: "pending",
        createdAt: new Date().toISOString(),
        decidedAt: null
      });
    }

    for (let column = 1; column <= sheet.columnCount; column++) {
      const populated = nonEmptyRows.flatMap((row) => {
        const cell = row.cells[column - 1];
        return cell && !isBlank(cell.value) ? [cell] : [];
      });
      const formats = new Set(populated.map((cell) => cell.numberFormat || "General"));
      if (formats.size > 1 && populated.some((cell) =>
        cell.type === ExcelJS.ValueType.Number || cell.type === ExcelJS.ValueType.Date
      )) {
        findings.push({
          worksheetKey: sheet.worksheetKey,
          worksheetName: sheet.name,
          issueType: "inconsistent_format",
          target: { kind: "column", column },
          reason: "Cells in this column use different number formats; no format is proposed automatically."
        });
      }
    }
  }

  return {
    reviewId,
    uploadVersionId,
    configuration,
    worksheetResults,
    proposals,
    findings,
    createdAt: new Date().toISOString()
  };
}

export function recordProposalDecision(review, proposalIdValue, decision) {
  if (!["approved", "rejected"].includes(decision)) {
    throw new TypeError("Decision must be approved or rejected.");
  }
  const proposal = review?.proposals.find(
    (entry) => entry.proposalId === proposalIdValue
  );
  if (!proposal) {
    return { status: "not_found" };
  }
  if (proposal.decision !== "pending") {
    return { status: "already_decided", proposal };
  }

  proposal.decision = decision;
  proposal.decidedAt = new Date().toISOString();
  return { status: "updated", proposal };
}

export async function createWorkbookState(uploadedBytes, fileName) {
  if (!Buffer.isBuffer(uploadedBytes) || uploadedBytes.length === 0) {
    throw new TypeError("Workbook source must be a non-empty Buffer.");
  }

  const originalBytes = Buffer.from(uploadedBytes);
  const preservationWorkbook = new ExcelJS.Workbook();

  await preservationWorkbook.xlsx.load(Buffer.from(originalBytes));

  if (preservationWorkbook.worksheets.length === 0) {
    throw new Error("The uploaded workbook contains no worksheets.");
  }

  const worksheets = preservationWorkbook.worksheets.map(
    (worksheet, index) => describeWorksheet(worksheet, index)
  );

  return {
    fileName,
    uploadVersionId: randomUUID(),
    sourceHash: createHash("sha256")
      .update(originalBytes)
      .digest("hex"),
    getSourceBytes() {
      return Buffer.from(originalBytes);
    },
    preservationWorkbook,
    worksheets
  };
}
