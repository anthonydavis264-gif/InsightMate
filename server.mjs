
import express from "express";
import multer from "multer";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { randomUUID, timingSafeEqual } from "node:crypto";
import { fileURLToPath } from "url";
import {
  cleanWorkbook,
  createWorkbookState,
  generateCleanedWorkbook,
  generateCleaningReview,
  normalizeComparableText,
  recordProposalDecision,
  validateHeaderRow
} from "./workbook-state.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3000);
const OLLAMA_URL =
  process.env.OLLAMA_URL || "http://localhost:11434";
const OLLAMA_MODEL = "gemma4:e2b-it-qat";
const SERVER_VERSION = "analysis-engine-v4.4";

const app = express();

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({
  dest: process.env.UPLOAD_DIR || path.join(__dirname, "uploads"),
  limits: {
    fileSize: 25 * 1024 * 1024
  }
});

let currentDataset = {
  fileName: null,
  sheetName: null,
  rows: [],
  columns: []
};
let currentWorkbookState = null;
let currentUploadVersionId = null;
let currentCleaningAccessToken = null;
let currentCleaningOperation = null;
let currentCleaningReview = null;
let currentHeaderRows = {};
let latestStartedUploadSequence = 0;

function markUploadSequence(req, res, next) {
  req.uploadSequence = ++latestStartedUploadSequence;
  next();
}

function authorizeCleaningUpload(uploadVersionId, cleaningToken) {
  if (!currentWorkbookState) {
    return {
      status: 409,
      error: "The active upload is not an Excel workbook."
    };
  }
  if (
    typeof uploadVersionId !== "string" ||
    uploadVersionId !== currentUploadVersionId
  ) {
    return {
      status: 409,
      error: "The upload version is stale."
    };
  }
  if (
    typeof cleaningToken !== "string" ||
    typeof currentCleaningAccessToken !== "string" ||
    cleaningToken.length !== currentCleaningAccessToken.length ||
    !timingSafeEqual(
      Buffer.from(cleaningToken),
      Buffer.from(currentCleaningAccessToken)
    )
  ) {
    return {
      status: 404,
      error: "The cleaning session could not be found."
    };
  }
  return null;
}

function cleaningSummaryForClient(summary, resultReady = false) {
  return {
    worksheetCount: summary.worksheetCount,
    explanation: summary.explanation,
    worksheets: summary.worksheets.map((worksheet) => ({
      name: worksheet.name,
      order: worksheet.order,
      rowsFormatted: worksheet.rowsFormatted,
      spacesTrimmed: worksheet.spacesTrimmed,
      errorReplacements: worksheet.errorReplacements.map((replacement) => ({
        worksheetName: replacement.worksheetName,
        address: replacement.address,
        originalError: replacement.originalError,
        originalFormula: replacement.originalFormula,
        action: replacement.action
      })),
      emptyRowsRemoved: worksheet.emptyRowsRemoved,
      blankRowsDetected: worksheet.blankRowsDetected,
      blankRowsRemoved: resultReady ? worksheet.blankRowsRemoved : 0,
      blankRowsToRemove: resultReady ? 0 : worksheet.blankRowsRemoved,
      blankRowsRetained: worksheet.blankRowsRetained,
      blankRowsRetainedByReason: worksheet.blankRowsRetainedByReason.map((item) => ({
        reason: item.reason,
        count: item.count
      })),
      blankCellsDetected: worksheet.blankCellsDetected,
      duplicateCandidatesFound: worksheet.duplicateCandidatesFound,
      duplicateRowsSkipped: worksheet.duplicateRowsSkipped.map((item) => ({
        originalRow: item.originalRow,
        retainedOriginalRow: item.retainedOriginalRow,
        reason: item.reason
      })),
      removedRows: worksheet.removedRows,
      mergedRangeActions: worksheet.mergedRangeActions.map((item) => ({
        worksheetName: item.worksheetName,
        originalRange: item.originalRange,
        reason: item.reason,
        action: item.action,
        unresolvedMissingFields: item.unresolvedMissingFields
      })),
      ambiguousItemsLeftUnchanged: worksheet.ambiguousItemsLeftUnchanged,
      itemsSkippedForSafety: worksheet.itemsSkippedForSafety,
      duplicateCandidates: worksheet.duplicateCandidates.map((candidate) => ({
        candidateId: candidate.candidateId,
        worksheetKey: `sheet-${worksheet.order}`,
        originalRow: candidate.originalRow,
        retainedOriginalRow: candidate.retainedOriginalRow,
        safeToRemove: candidate.safeToRemove,
        unsafeReason: candidate.unsafeReason
      })),
      explanation: worksheet.explanation,
      limitations: worksheet.limitations
    }))
  };
}

function cleanedDownloadName(fileName) {
  const stem = path.basename(
    String(fileName || "workbook"),
    path.extname(String(fileName || ""))
  );
  const safeStem = stem.replace(/[^\w.-]+/g, "_").replace(/^\.+/, "");
  return `${safeStem || "workbook"}_cleaned.xlsx`;
}

/* ============================================================
   BASIC HELPERS
   ============================================================ */

function normalizeText(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (value === null || value === undefined || value === "") {
    return null;
  }

  let cleaned = String(value)
    .trim()
    .replace(/[₹$€£]/g, "")
    .replace(/,/g, "")
    .replace(/%/g, "")
    .replace(/\s+/g, "");

  if (!cleaned) {
    return null;
  }

  // Support accounting-style negative numbers, e.g. (125.50).
  const isParenthesized = /^\(.*\)$/.test(cleaned);

  if (isParenthesized) {
    cleaned = cleaned.slice(1, -1);
  }

  const number = Number(cleaned);

  if (!Number.isFinite(number)) {
    return null;
  }

  return isParenthesized ? -Math.abs(number) : number;
}

function roundNumber(value, decimals = 2) {
  if (!Number.isFinite(value)) {
    return null;
  }

  const factor = 10 ** decimals;

  return Math.round(
    (value + Number.EPSILON) * factor
  ) / factor;
}

function isEmptyValue(value) {
  return (
    value === null ||
    value === undefined ||
    String(value).trim() === ""
  );
}

function detectDataType(values) {
  const nonEmpty = values.filter(
    (value) => !isEmptyValue(value)
  );

  if (!nonEmpty.length) {
    return "unknown";
  }

  const numericCount = nonEmpty.filter(
    (value) => toNumber(value) !== null
  ).length;

  if (numericCount / nonEmpty.length >= 0.8) {
    return "number";
  }

  return "string";
}

function isRepeatedHeaderRow(row, columns) {
  if (!row || typeof row !== "object" || !columns.length) {
    return false;
  }

  let matches = 0;

  for (const column of columns) {
    const value = normalizeText(row[column]);
    const header = normalizeText(column);

    if (value && header && value === header) {
      matches++;
    }
  }

  return matches >= Math.max(
    1,
    Math.ceil(columns.length * 0.5)
  );
}

function resolveColumnName(requestedName, columns) {
  if (!requestedName || !Array.isArray(columns)) {
    return null;
  }

  const normalizedRequested = normalizeText(requestedName);

  const exact = columns.find(
    (column) => normalizeText(column) === normalizedRequested
  );

  if (exact) {
    return exact;
  }

  const aliases = {
    amount: ["amount", "sales", "revenue", "value", "total"],
    segment: ["segment", "customer segment"],
    order: ["order", "order id", "orderid", "orders"]
  };

  for (const [canonical, possibleNames] of Object.entries(aliases)) {
    if (
      normalizedRequested === canonical ||
      possibleNames.includes(normalizedRequested)
    ) {
      const match = columns.find((column) =>
        possibleNames.includes(normalizeText(column))
      );

      if (match) {
        return match;
      }
    }
  }

  return null;
}

/* ============================================================
   DATASET PROFILE AND DATA QUALITY
   ============================================================ */

function buildDatasetProfile(rows, columns) {
  const profileColumns = columns.map((column) => {
    const values = rows.map((row) => row[column]);

    const nonEmptyValues = values.filter(
      (value) => !isEmptyValue(value)
    );

    const uniqueValues = new Set(
      nonEmptyValues.map((value) => String(value))
    );

    return {
      name: column,
      dataType: detectDataType(values),
      uniqueCount: uniqueValues.size,
      sampleValues: nonEmptyValues.slice(0, 5)
    };
  });

  return {
    rowCount: rows.length,
    columnCount: columns.length,
    columns: profileColumns
  };
}

function buildDataQuality(rows, columns) {
  const quality = {};

  for (const column of columns) {
    let missingCount = 0;

    for (const row of rows) {
      if (isEmptyValue(row[column])) {
        missingCount++;
      }
    }

    quality[column] = {
      missingCount,
      missingPercentage: rows.length > 0
        ? roundNumber(
            (missingCount / rows.length) * 100,
            2
          )
        : 0
    };
  }

  return quality;
}

/* ============================================================
   OLLAMA
   ============================================================ */

async function callOllama(prompt) {
  const response = await fetch(`${OLLAMA_URL}/api/generate`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: OLLAMA_MODEL,
      prompt,
      stream: false,
      format: "json"
    })
  });

  if (!response.ok) {
    const errorText = await response.text();

    throw new Error(
      `Ollama request failed: ${response.status} ${errorText}`
    );
  }

  const data = await response.json();

  return data?.response || "";
}

/* ============================================================
   JSON EXTRACTION
   ============================================================ */

function extractJson(text) {
  if (!text) {
    throw new Error("Ollama returned an empty response.");
  }

  if (typeof text === "object") {
    return text;
  }

  const trimmed = String(text).trim();

  try {
    return JSON.parse(trimmed);
  } catch {
    // Try extracting a JSON object from the response.
  }

  const firstBrace = trimmed.indexOf("{");
  const lastBrace = trimmed.lastIndexOf("}");

  if (firstBrace !== -1 && lastBrace > firstBrace) {
    const jsonText = trimmed.slice(
      firstBrace,
      lastBrace + 1
    );

    try {
      return JSON.parse(jsonText);
    } catch {
      throw new Error("Unable to parse Ollama JSON response.");
    }
  }

  throw new Error("No valid JSON object found in Ollama response.");
}

/* ============================================================
   FILTER REPAIR
   ============================================================ */

function repairFilterPlanFromQuestion(plan, question, columns) {
  const repairedPlan = {
    ...plan,
    filters: Array.isArray(plan?.filters)
      ? [...plan.filters]
      : []
  };

  const normalizedQuestion = normalizeText(question);

  const segmentColumn = resolveColumnName(
    "Segment",
    columns
  );

  if (!segmentColumn) {
    return repairedPlan;
  }

  const knownSegments = [
    "consumer",
    "corporate",
    "home office"
  ];

  const detectedSegment = knownSegments.find((segment) =>
    normalizedQuestion.includes(segment)
  );

  if (detectedSegment) {
    const actualValue =
      detectedSegment === "consumer"
        ? "Consumer"
        : detectedSegment === "corporate"
          ? "Corporate"
          : "Home Office";

    repairedPlan.filters = repairedPlan.filters.filter(
      (filter) =>
        normalizeText(filter?.column) !==
        normalizeText(segmentColumn)
    );

    repairedPlan.filters.push({
      column: segmentColumn,
      operator: "EQUALS",
      value: actualValue
    });
  }

  return repairedPlan;
}

/* ============================================================
   DETERMINISTIC ANALYSIS INTENT REPAIR
   ============================================================ */

function repairAnalysisIntent(plan, question, columns) {
  const repairedPlan = {
    ...plan,
    filters: Array.isArray(plan?.filters)
      ? [...plan.filters]
      : []
  };

  const normalizedQuestion = normalizeText(question);

  const amountColumn = resolveColumnName("Amount", columns);
  const segmentColumn = resolveColumnName("Segment", columns);
  const shipModeColumn = resolveColumnName("Ship Mode", columns);
  const orderIdColumn = resolveColumnName("Order ID", columns);

  const isBySegmentQuestion =
    normalizedQuestion.includes("by segment") ||
    normalizedQuestion.includes("which segment") ||
    normalizedQuestion.includes("what segment") ||
    normalizedQuestion.includes("segment has");

  const isByShipModeQuestion =
    normalizedQuestion.includes("by ship mode") ||
    normalizedQuestion.includes("for each ship mode") ||
    normalizedQuestion.includes("each ship mode") ||
    normalizedQuestion.includes("ship mode has");

  const groupByColumn = isBySegmentQuestion
    ? segmentColumn
    : isByShipModeQuestion
      ? shipModeColumn
      : null;

  if (groupByColumn) {
    repairedPlan.operation = "GROUP_BY";
    repairedPlan.groupBy = groupByColumn;

    const isCountQuestion =
      normalizedQuestion.includes("how many") ||
      normalizedQuestion.includes("count");

    if (isCountQuestion) {
      repairedPlan.metric = orderIdColumn || groupByColumn;
      repairedPlan.aggregation = "COUNT";
      repairedPlan.sort = "DESC";
      repairedPlan.limit = 0;

      return repairedPlan;
    }

    // Amount-based grouping requires a usable metric column.
    if (!amountColumn) {
      return repairedPlan;
    }

    repairedPlan.metric = amountColumn;

    if (
      normalizedQuestion.includes("total") ||
      normalizedQuestion.includes("sum")
    ) {
      repairedPlan.aggregation = "SUM";
    } else if (
      normalizedQuestion.includes("average") ||
      normalizedQuestion.includes("mean")
    ) {
      repairedPlan.aggregation = "AVERAGE";
    } else if (
      normalizedQuestion.includes("maximum") ||
      normalizedQuestion.includes("max")
    ) {
      repairedPlan.aggregation = "MAX";
    } else if (
      normalizedQuestion.includes("minimum") ||
      normalizedQuestion.includes("min")
    ) {
      repairedPlan.aggregation = "MIN";
    } else if (
      normalizedQuestion.includes("highest") ||
      normalizedQuestion.includes("largest")
    ) {
      repairedPlan.aggregation = "MAX";
    } else if (
      normalizedQuestion.includes("lowest") ||
      normalizedQuestion.includes("smallest")
    ) {
      repairedPlan.aggregation = "MIN";
    } else {
      repairedPlan.aggregation = "SUM";
    }

    repairedPlan.sort =
      normalizedQuestion.includes("lowest") ||
      normalizedQuestion.includes("smallest") ||
      normalizedQuestion.includes("minimum")
        ? "ASC"
        : "DESC";

    repairedPlan.limit = 0;

    return repairedPlan;
  }

  if (!amountColumn) {
    return repairedPlan;
  }

  const asksForAmount = normalizedQuestion.includes("amount");

  const hasNoGroupingIntent =
    !normalizedQuestion.includes("by ") &&
    !normalizedQuestion.includes("which segment") &&
    !normalizedQuestion.includes("what segment");

  if (asksForAmount && hasNoGroupingIntent) {
    let aggregation = null;

    if (
      normalizedQuestion.includes("total") ||
      normalizedQuestion.includes("sum")
    ) {
      aggregation = "SUM";
    } else if (
      normalizedQuestion.includes("average") ||
      normalizedQuestion.includes("mean")
    ) {
      aggregation = "AVERAGE";
    } else if (
      normalizedQuestion.includes("maximum") ||
      normalizedQuestion.includes("max") ||
      normalizedQuestion.includes("highest amount")
    ) {
      aggregation = "MAX";
    } else if (
      normalizedQuestion.includes("minimum") ||
      normalizedQuestion.includes("min") ||
      normalizedQuestion.includes("lowest amount")
    ) {
      aggregation = "MIN";
    }

    if (aggregation) {
      repairedPlan.operation = "AGGREGATE";
      repairedPlan.metric = amountColumn;
      repairedPlan.aggregation = aggregation;
      repairedPlan.groupBy = null;
      repairedPlan.sort = null;
      repairedPlan.limit = null;
    }
  }

  return repairedPlan;
}

/* ============================================================
   PLAN VALIDATION
   ============================================================ */

function validateAnalysisPlan(plan, columns) {
  const validOperations = [
    "AGGREGATE",
    "COUNT",
    "GROUP_BY"
  ];

  const validAggregations = [
    "SUM",
    "AVERAGE",
    "MIN",
    "MAX",
    "COUNT"
  ];

  if (!plan || typeof plan !== "object") {
    throw new Error("Invalid analysis plan.");
  }

  if (!validOperations.includes(plan.operation)) {
    throw new Error(
      `Invalid analysis operation: ${plan.operation}`
    );
  }

  if (
    plan.operation === "AGGREGATE" ||
    plan.operation === "GROUP_BY"
  ) {
    if (!resolveColumnName(plan.metric, columns)) {
      throw new Error(
        `Invalid metric column: ${plan.metric}`
      );
    }

    if (!validAggregations.includes(plan.aggregation)) {
      throw new Error(
        `Invalid aggregation: ${plan.aggregation}`
      );
    }

    if (
      plan.operation === "AGGREGATE" &&
      plan.aggregation === "COUNT"
    ) {
      throw new Error(
        "COUNT aggregation must use COUNT or GROUP_BY."
      );
    }
  }

  if (plan.operation === "GROUP_BY") {
    if (!resolveColumnName(plan.groupBy, columns)) {
      throw new Error(
        `Invalid group-by column: ${plan.groupBy}`
      );
    }
  }

  return true;
}

/* ============================================================
   PLAN NORMALIZATION
   ============================================================ */

function normalizeAnalysisPlan(plan, columns) {
  const normalizedPlan = {
    operation: plan?.operation || null,
    groupBy: plan?.groupBy || null,
    metric: plan?.metric || null,
    aggregation: plan?.aggregation || null,
    filters: Array.isArray(plan?.filters)
      ? plan.filters
      : [],
    sort: plan?.sort || null,
    limit: plan?.limit === undefined
      ? null
      : plan.limit
  };

  if (normalizedPlan.metric) {
    normalizedPlan.metric = resolveColumnName(
      normalizedPlan.metric,
      columns
    );
  }

  if (normalizedPlan.groupBy) {
    normalizedPlan.groupBy = resolveColumnName(
      normalizedPlan.groupBy,
      columns
    );
  }

  // Resolve filter columns so they match actual row keys.
  normalizedPlan.filters = normalizedPlan.filters.map(
    (filter) => ({
      ...filter,
      column: resolveColumnName(filter?.column, columns) ||
        filter?.column
    })
  );

  if (normalizedPlan.operation === "GROUP_BY") {
    normalizedPlan.limit = 0;

    if (!normalizedPlan.sort) {
      normalizedPlan.sort = "DESC";
    }
  }

  return normalizedPlan;
}

/* ============================================================
   FILTER APPLICATION
   ============================================================ */

function applyFilters(rows, filters) {
  if (!Array.isArray(filters) || filters.length === 0) {
    return rows;
  }

  return rows.filter((row) =>
    filters.every((filter) => {
      const column = filter?.column;
      const operator = String(
        filter?.operator || "EQUALS"
      ).toUpperCase();

      const expectedValue = filter?.value;

      if (!column || !(column in row)) {
        return false;
      }

      const actualValue = row[column];

      switch (operator) {
        case "EQUALS":
        case "=":
          return normalizeText(actualValue) ===
            normalizeText(expectedValue);

        case "NOT_EQUALS":
        case "!=":
          return normalizeText(actualValue) !==
            normalizeText(expectedValue);

        case "CONTAINS":
          return normalizeText(actualValue).includes(
            normalizeText(expectedValue)
          );

        case "GREATER_THAN": {
          const actual = toNumber(actualValue);
          const expected = toNumber(expectedValue);

          return actual !== null &&
            expected !== null &&
            actual > expected;
        }

        case "LESS_THAN": {
          const actual = toNumber(actualValue);
          const expected = toNumber(expectedValue);

          return actual !== null &&
            expected !== null &&
            actual < expected;
        }

        default:
          return false;
      }
    })
  );
}

/* ============================================================
   AGGREGATION HELPERS
   ============================================================ */

function getNumericValues(rows, metric) {
  return rows
    .map((row) => toNumber(row[metric]))
    .filter((value) => value !== null);
}

function calculateAggregation(rows, metric, aggregation) {
  if (aggregation === "COUNT") {
    return rows.length;
  }

  const values = getNumericValues(rows, metric);

  if (!values.length) {
    return null;
  }

  switch (aggregation) {
    case "SUM":
      return roundNumber(
        values.reduce((sum, value) => sum + value, 0),
        2
      );

    case "AVERAGE":
      return roundNumber(
        values.reduce((sum, value) => sum + value, 0) /
          values.length,
        2
      );

    case "MIN":
      return roundNumber(Math.min(...values), 2);

    case "MAX":
      return roundNumber(Math.max(...values), 2);

    default:
      return null;
  }
}

/* ============================================================
   AGGREGATE
   ============================================================ */

function executeAggregate(rows, plan) {
  const value = calculateAggregation(
    rows,
    plan.metric,
    plan.aggregation
  );

  return {
    type: "aggregate",
    metric: plan.metric,
    aggregation: plan.aggregation,
    value
  };
}

/* ============================================================
   COUNT
   ============================================================ */

function executeCount(rows) {
  return {
    type: "count",
    value: rows.length
  };
}

/* ============================================================
   GROUP BY
   ============================================================ */

function executeGroupBy(rows, plan) {
  const groups = new Map();

  for (const row of rows) {
    const rawGroupValue = row[plan.groupBy];

    const groupValue = isEmptyValue(rawGroupValue)
      ? "Unknown"
      : String(rawGroupValue).trim();

    if (!groups.has(groupValue)) {
      groups.set(groupValue, []);
    }

    groups.get(groupValue).push(row);
  }

  const results = [];

  for (const [groupValue, groupRows] of groups.entries()) {
    const value = calculateAggregation(
      groupRows,
      plan.metric,
      plan.aggregation
    );

    results.push({
      group: groupValue,
      value
    });
  }

  results.sort((a, b) => {
    if (a.value === null) return 1;
    if (b.value === null) return -1;

    return plan.sort === "ASC"
      ? a.value - b.value
      : b.value - a.value;
  });

  const finalResults =
    plan.limit && plan.limit > 0
      ? results.slice(0, plan.limit)
      : results;

  return {
    type: "group_by",
    groupBy: plan.groupBy,
    metric: plan.metric,
    aggregation: plan.aggregation,
    sort: plan.sort,
    rows: finalResults
  };
}

/* ============================================================
   EXECUTE ANALYSIS PLAN
   ============================================================ */

function executeAnalysisPlan(rows, plan) {
  const filteredRows = applyFilters(rows, plan.filters);

  switch (plan.operation) {
    case "COUNT":
      return executeCount(filteredRows);

    case "AGGREGATE":
      return executeAggregate(filteredRows, plan);

    case "GROUP_BY":
      return executeGroupBy(filteredRows, plan);

    default:
      throw new Error(
        `Unsupported analysis operation: ${plan.operation}`
      );
  }
}

/* ============================================================
   CREATE AI ANALYSIS PLAN
   ============================================================ */

async function createAnalysisPlan(question, profile) {
  const columns = profile.columns.map((column) => ({
    name: column.name,
    dataType: column.dataType
  }));

  const prompt = `
You are the analysis planner for InsightMate, a local AI data analyst.

Convert the user's question into a JSON analysis plan.
Do not calculate the answer.
Return only valid JSON.

Available columns:
${JSON.stringify(columns, null, 2)}

User question:
"${question}"

Allowed operations: COUNT, AGGREGATE, GROUP_BY.
Allowed aggregations: SUM, AVERAGE, MIN, MAX, COUNT.

Use this JSON structure:
{
  "operation": "COUNT | AGGREGATE | GROUP_BY",
  "groupBy": "column or null",
  "metric": "column or null",
  "aggregation": "SUM | AVERAGE | MIN | MAX | COUNT | null",
  "filters": [],
  "sort": "ASC | DESC | null",
  "limit": 0
}

Rules:
1. "How many orders are there?" means COUNT.
2. "How many orders are from Consumer?" means COUNT with Segment = Consumer.
3. "How many orders are from Corporate?" means COUNT with Segment = Corporate.
4. "How many orders are from Home Office?" means COUNT with Segment = Home Office.
5. "What is the total amount for Consumer?" means AGGREGATE, Amount, SUM, with Segment = Consumer.
6. "What is the average amount for Corporate?" means AGGREGATE, Amount, AVERAGE, with Segment = Corporate.
7. Maximum amount means AGGREGATE with MAX.
8. Minimum amount means AGGREGATE with MIN.
9. Average amount by Segment means GROUP_BY Segment, Amount, AVERAGE.
10. Total amount by Segment means GROUP_BY Segment, Amount, SUM.
11. Amount questions without grouping use AGGREGATE.
12. Questions about the highest, lowest, largest, or smallest segment use GROUP_BY Segment.
13. "Which segment has the lowest total amount?" means GROUP_BY Segment, Amount, SUM, sort ASC.
14. "Which segment has the highest total amount?" means GROUP_BY Segment, Amount, SUM, sort DESC.
15. Highest or largest means DESC; lowest or smallest means ASC.
16. Grouping questions asking "how many" or "count" use GROUP_BY with aggregation COUNT.
17. "How many orders are there by Ship Mode?" means GROUP_BY Ship Mode, aggregation COUNT.
18. GROUP_BY must return all groups with limit = 0.
19. Use only column names from the available columns.
20. Do not include explanations outside the JSON.
`;

  const response = await callOllama(prompt);

  return extractJson(response);
}

/* ============================================================
   DATA CLEANING HELPERS
   ============================================================ */

function applyCleaningOptions(rows, columns, options = {}) {
  // Work on copies; never modify the active dataset.
  let cleanedRows = rows.map((row) => ({ ...row }));

  const summary = {
    originalRows: rows.length,
    finalRows: rows.length,
    trimmedCells: 0,
    emptyRowsRemoved: 0,
    duplicateRowsRemoved: 0
  };

  if (options.trimWhitespace === true) {
    cleanedRows = cleanedRows.map((row) => {
      const copy = { ...row };

      for (const column of columns) {
        const value = copy[column];

        if (typeof value === "string") {
          const trimmed = value.trim();

          if (trimmed !== value) {
            summary.trimmedCells++;
            copy[column] = trimmed;
          }
        }
      }

      return copy;
    });
  }

  if (options.removeEmptyRows === true) {
    const beforeCount = cleanedRows.length;

    cleanedRows = cleanedRows.filter((row) =>
      columns.some((column) => !isEmptyValue(row[column]))
    );

    summary.emptyRowsRemoved =
      beforeCount - cleanedRows.length;
  }

  const seen = new Set();
  const beforeCount = cleanedRows.length;
  cleanedRows = cleanedRows.filter((row) => {
    if (!columns.some((column) => !isEmptyValue(row[column]))) return true;
    const key = JSON.stringify(columns.map((column) =>
      normalizeComparableText(row[column] ?? "")
    ));
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  summary.duplicateRowsRemoved = beforeCount - cleanedRows.length;

  summary.finalRows = cleanedRows.length;

  return {
    rows: cleanedRows,
    summary
  };
}

function getCleaningOptions(body) {
  const options = body?.options || {};

  return {
    trimWhitespace: options.trimWhitespace === true,
    removeEmptyRows: options.removeEmptyRows === true,
    removeDuplicates: true
  };
}

/* ============================================================
   HEALTH
   ============================================================ */

app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    status: "ok",
    serverVersion: SERVER_VERSION,
    ollamaUrl: OLLAMA_URL,
    ollamaModel: OLLAMA_MODEL,
    datasetLoaded: currentDataset.rows.length > 0
  });
});

/* ============================================================
   DATASET
   ============================================================ */

app.get("/api/dataset", (req, res) => {
  if (!currentDataset.rows.length) {
    return res.json({
      success: true,
      loaded: false
    });
  }

  const profile = buildDatasetProfile(
    currentDataset.rows,
    currentDataset.columns
  );

  const quality = buildDataQuality(
    currentDataset.rows,
    currentDataset.columns
  );

  res.json({
    success: true,
    loaded: true,
    fileName: currentDataset.fileName,
    sheetName: currentDataset.sheetName,
    rows: currentDataset.rows.length,
    columns: currentDataset.columns.length,
    columnNames: currentDataset.columns,
    profile,
    quality,
    preview: currentDataset.rows.slice(0, 10)
  });
});

/* ============================================================
   UPLOAD
   ============================================================ */

app.post(
  "/api/upload",
  markUploadSequence,
  upload.single("file"),
  async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({
          success: false,
          error: "No file uploaded."
        });
      }

      let workbookState = null;

      if (/\.xlsx$/i.test(req.file.originalname)) {
        const uploadedBytes = fs.readFileSync(req.file.path);
        workbookState = await createWorkbookState(
          uploadedBytes,
          req.file.originalname
        );
      }

      const workbook = XLSX.readFile(req.file.path, {
        cellDates: true
      });

      if (!workbook.SheetNames.length) {
        throw new Error(
          "The uploaded workbook contains no sheets."
        );
      }

      // Keep the existing first-worksheet behavior.
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      /*
       * Read headers and values from the same matrix.
       * This keeps each column name connected to the
       * corresponding cell position.
       */
      const matrix = XLSX.utils.sheet_to_json(
        worksheet,
        {
          header: 1,
          defval: "",
          raw: true,
          blankrows: true
        }
      );

      if (!matrix.length) {
        throw new Error(
          "The selected worksheet is empty."
        );
      }

      const rawHeaders = matrix[0] || [];
      const dataRows = matrix.slice(1);

      const columnDefinitions = [];
      const usedNames = new Set();

      for (let index = 0; index < rawHeaders.length; index++) {
        const rawHeader = String(
          rawHeaders[index] ?? ""
        ).trim();

        // Ignore a blank-header column only when every cell is empty.
        const columnHasData = dataRows.some((cells) =>
          !isEmptyValue(cells[index])
        );

        if (!rawHeader && !columnHasData) {
          continue;
        }

        const baseName = rawHeader || `Column ${index + 1}`;

        // Make duplicate names unique without losing their values.
        let columnName = baseName;
        let suffix = 2;

        while (usedNames.has(normalizeText(columnName))) {
          columnName = `${baseName} (${suffix})`;
          suffix++;
        }

        usedNames.add(normalizeText(columnName));

        columnDefinitions.push({
          index,
          name: columnName
        });
      }

      const cleanedColumns = columnDefinitions.map(
        (column) => column.name
      );

      // Rebuild all rows using the same column positions.
      const rawRows = dataRows.map((cells) => {
        const row = {};

        for (const column of columnDefinitions) {
          row[column.name] = cells[column.index] ?? "";
        }

        return row;
      });

      // Remove repeated header rows while retaining other rows.
      const rows = rawRows.filter(
        (row) => !isRepeatedHeaderRow(row, cleanedColumns)
      );

      if (!cleanedColumns.length) {
        throw new Error(
          "No usable column headers or data columns were found."
        );
      }

      const nextDataset = {
        fileName: req.file.originalname,
        sheetName,
        rows,
        columns: cleanedColumns
      };

      const profile = buildDatasetProfile(
        rows,
        cleanedColumns
      );

      const quality = buildDataQuality(
        rows,
        cleanedColumns
      );

      if (req.uploadSequence !== latestStartedUploadSequence) {
        return res.status(409).json({
          success: false,
          error: "This upload was superseded by a newer upload request."
        });
      }

      // Commit both representations only after parsing and profiling succeed.
      currentDataset = nextDataset;
      currentWorkbookState = workbookState;
      currentUploadVersionId =
        workbookState?.uploadVersionId || randomUUID();
      currentCleaningAccessToken = workbookState ? randomUUID() : null;
      currentCleaningOperation = null;
      currentCleaningReview = null;
      currentHeaderRows = {};

      return res.json({
        success: true,
        uploadVersionId: currentUploadVersionId,
        cleaningToken: currentCleaningAccessToken,
        fileName: req.file.originalname,
        sheetName,
        rows: rows.length,
        columns: cleanedColumns.length,
        columnNames: cleanedColumns,
        profile,
        quality,
        preview: rows.slice(0, 10)
      });

    } catch (error) {
      console.error("Upload error:", error);

      if (req.file?.path) {
        try {
          fs.unlinkSync(req.file.path);
        } catch {
          // Ignore temporary file cleanup errors.
        }
      }

      return res.status(500).json({
        success: false,
        error:
          error?.message ||
          "Failed to process uploaded file."
      });
    } finally {
      if (req.file?.path) {
        try {
          fs.unlinkSync(req.file.path);
        } catch (error) {
          if (error?.code !== "ENOENT") {
            console.error("Upload temporary-file cleanup error:", error);
          }
        }
      }
    }
  }
);

app.get("/api/cleaning/workbook", (req, res) => {
  if (!currentWorkbookState) {
    return res.status(409).json({
      success: false,
      error: "The active upload is not an XLSX workbook."
    });
  }

  return res.json({
    success: true,
    uploadVersionId: currentUploadVersionId,
    worksheets: currentWorkbookState.worksheets.map((sheet) => ({
      worksheetKey: sheet.worksheetKey,
      name: sheet.name,
      order: sheet.order,
      visibility: sheet.visibility,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      header: currentHeaderRows[sheet.worksheetKey]
        ? {
            status: currentHeaderRows[sheet.worksheetKey].action === "skip"
              ? "skipped"
              : "confirmed",
            row: currentHeaderRows[sheet.worksheetKey].row ?? null,
            source: "user"
          }
        : sheet.headerCandidate
    }))
  });
});

app.post("/api/cleaning/workbook/headers", (req, res) => {
  if (!currentWorkbookState) {
    return res.status(409).json({
      success: false,
      error: "The active upload is not an XLSX workbook."
    });
  }

  if (req.body?.uploadVersionId !== currentUploadVersionId) {
    return res.status(409).json({
      success: false,
      error: "The upload version is stale."
    });
  }

  if (!Array.isArray(req.body?.headerSelections)) {
    return res.status(400).json({
      success: false,
      error: "headerSelections must be an array."
    });
  }

  const nextHeaderRows = { ...currentHeaderRows };
  for (const selection of req.body.headerSelections) {
    const sheet = currentWorkbookState.worksheets.find(
      (candidate) => candidate.worksheetKey === selection?.worksheetKey
    );
    if (!sheet) {
      return res.status(400).json({
        success: false,
        error: "A worksheet key is invalid."
      });
    }

    if (selection.action === "skip") {
      nextHeaderRows[sheet.worksheetKey] = { action: "skip", row: null };
    } else if (validateHeaderRow(sheet, selection.headerRow)) {
      nextHeaderRows[sheet.worksheetKey] = {
        action: "confirm",
        row: selection.headerRow
      };
    } else {
      return res.status(400).json({
        success: false,
        error: `Header row must be a non-empty original row in worksheet "${sheet.name}".`
      });
    }
  }

  const previousReviewId = currentCleaningReview?.reviewId ?? null;
  const changed =
    JSON.stringify(currentHeaderRows) !== JSON.stringify(nextHeaderRows);
  currentHeaderRows = nextHeaderRows;
  if (changed) {
    currentCleaningReview = null;
  }

  return res.json({
    success: true,
    uploadVersionId: currentUploadVersionId,
    worksheets: currentWorkbookState.worksheets.map((sheet) => {
      const selection = currentHeaderRows[sheet.worksheetKey];
      return {
        worksheetKey: sheet.worksheetKey,
        name: sheet.name,
        header: selection
          ? {
              status: selection.action === "skip" ? "skipped" : "confirmed",
              row: selection.row,
              source: "user"
            }
          : sheet.headerCandidate
      };
    }),
    reviewInvalidated: changed && previousReviewId !== null,
    replacedReviewId: changed ? previousReviewId : null
  });
});

app.post("/api/cleaning/proposals", (req, res) => {
  if (!currentWorkbookState) {
    return res.status(409).json({
      success: false,
      error: "The active upload is not an XLSX workbook."
    });
  }

  if (req.body?.uploadVersionId !== currentUploadVersionId) {
    return res.status(409).json({
      success: false,
      error: "The upload version is stale."
    });
  }

  try {
    currentCleaningReview = generateCleaningReview(
      currentWorkbookState,
      currentUploadVersionId,
      currentHeaderRows,
      currentCleaningReview
    );

    return res.json({
      success: true,
      uploadVersionId: currentUploadVersionId,
      reviewId: currentCleaningReview.reviewId,
      worksheets: currentCleaningReview.worksheetResults,
      proposals: currentCleaningReview.proposals,
      findings: currentCleaningReview.findings
    });
  } catch (error) {
    console.error("Cleaning proposal generation error:", error);
    return res.status(400).json({
      success: false,
      error: error?.message || "Could not generate cleaning proposals."
    });
  }
});

app.post("/api/cleaning/proposals/:proposalId/decision", (req, res) => {
  if (!currentCleaningReview) {
    return res.status(404).json({
      success: false,
      error: "No active cleaning review exists."
    });
  }

  if (
    req.body?.uploadVersionId !== currentUploadVersionId ||
    req.body?.reviewId !== currentCleaningReview.reviewId
  ) {
    return res.status(409).json({
      success: false,
      error: "The upload version or review is stale."
    });
  }

  if (!["approved", "rejected"].includes(req.body?.decision)) {
    return res.status(400).json({
      success: false,
      error: "decision must be approved or rejected."
    });
  }

  const result = recordProposalDecision(
    currentCleaningReview,
    req.params.proposalId,
    req.body.decision
  );
  if (result.status === "not_found") {
    return res.status(404).json({
      success: false,
      error: "The proposal does not exist in the active review."
    });
  }
  if (result.status === "already_decided") {
    return res.status(409).json({
      success: false,
      error: "The proposal has already been decided."
    });
  }
  return res.json({
    success: true,
    uploadVersionId: currentUploadVersionId,
    reviewId: currentCleaningReview.reviewId,
    proposal: result.proposal
  });
});

/* ============================================================
   ANALYZE QUESTION
   ============================================================ */

app.post("/api/cleaning/start", async (req, res) => {
  const authorizationError = authorizeCleaningUpload(
    req.body?.uploadVersionId,
    req.body?.cleaningToken
  );
  if (authorizationError) {
    return res.status(authorizationError.status).json({
      success: false,
      error: authorizationError.error
    });
  }

  const operationUploadVersionId = currentUploadVersionId;
  const operationAccessToken = currentCleaningAccessToken;
  const sourceWorkbookState = currentWorkbookState;
  currentCleaningOperation = null;
  try {
    const { summary } = await cleanWorkbook(sourceWorkbookState);
    if (
      currentUploadVersionId !== operationUploadVersionId ||
      currentCleaningAccessToken !== operationAccessToken
    ) {
      return res.status(409).json({
        success: false,
        error: "The upload changed while cleaning was starting. Please try again."
      });
    }
    const responseSummary = cleaningSummaryForClient(summary);
    currentCleaningOperation = {
      uploadVersionId: operationUploadVersionId,
      summary: responseSummary,
      buffer: null
    };
    return res.json({
      success: true,
      uploadVersionId: operationUploadVersionId,
      summary: responseSummary
    });
  } catch (error) {
    console.error("Excel cleaning start error:", error);
    return res.status(500).json({
      success: false,
      error: "Could not prepare this workbook for cleaning."
    });
  }
});

app.post("/api/cleaning/generate", async (req, res) => {
  const authorizationError = authorizeCleaningUpload(
    req.body?.uploadVersionId,
    req.body?.cleaningToken
  );
  if (authorizationError) {
    return res.status(authorizationError.status).json({
      success: false,
      error: authorizationError.error
    });
  }
  if (
    !currentCleaningOperation ||
    currentCleaningOperation.uploadVersionId !== currentUploadVersionId
  ) {
    return res.status(409).json({
      success: false,
      error: "Start cleaning this workbook before creating a download."
    });
  }
  const operationUploadVersionId = currentUploadVersionId;
  const operationAccessToken = currentCleaningAccessToken;
  const sourceWorkbookState = currentWorkbookState;
  const activeOperation = currentCleaningOperation;
  try {
    const result = await generateCleanedWorkbook(sourceWorkbookState);
    if (
      currentUploadVersionId !== operationUploadVersionId ||
      currentCleaningAccessToken !== operationAccessToken ||
      currentCleaningOperation !== activeOperation
    ) {
      return res.status(409).json({
        success: false,
        error: "The upload changed while the workbook was being created. Please try again."
      });
    }
    const responseSummary = cleaningSummaryForClient(result.summary, true);
    currentCleaningOperation = {
      ...activeOperation,
      summary: responseSummary,
      buffer: Buffer.from(result.buffer)
    };
    return res.json({
      success: true,
      uploadVersionId: operationUploadVersionId,
      downloadReady: true,
      summary: responseSummary
    });
  } catch (error) {
    console.error("Excel workbook generation error:", error);
    return res.status(
      error instanceof TypeError || error instanceof RangeError ? 400 : 500
    ).json({
      success: false,
      error: error instanceof TypeError || error instanceof RangeError
        ? error.message
        : "Could not create the cleaned workbook."
    });
  }
});

app.get("/api/cleaning/download", (req, res) => {
  const authorizationError = authorizeCleaningUpload(
    req.get("x-upload-version-id"),
    req.get("x-cleaning-token")
  );
  if (authorizationError) {
    return res.status(authorizationError.status).json({
      success: false,
      error: authorizationError.error
    });
  }
  if (
    !currentCleaningOperation ||
    currentCleaningOperation.uploadVersionId !== currentUploadVersionId ||
    !Buffer.isBuffer(currentCleaningOperation.buffer)
  ) {
    return res.status(409).json({
      success: false,
      error: "Create a cleaned workbook before downloading it."
    });
  }

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Cache-Control", "no-store");
  res.attachment(cleanedDownloadName(currentDataset.fileName));
  return res.send(currentCleaningOperation.buffer);
});

app.post(
  "/api/analyze-question",
  async (req, res) => {
    try {
      const question = String(
        req.body?.question || ""
      ).trim();

      if (!question) {
        return res.status(400).json({
          success: false,
          error: "Please enter a question."
        });
      }

      if (!currentDataset.rows.length) {
        return res.status(400).json({
          success: false,
          error: "Please upload a dataset first."
        });
      }

      const profile = buildDatasetProfile(
        currentDataset.rows,
        currentDataset.columns
      );

      const aiPlan = await createAnalysisPlan(
        question,
        profile
      );

      let repairedPlan = repairFilterPlanFromQuestion(
        aiPlan,
        question,
        currentDataset.columns
      );

      repairedPlan = repairAnalysisIntent(
        repairedPlan,
        question,
        currentDataset.columns
      );

      const normalizedPlan = normalizeAnalysisPlan(
        repairedPlan,
        currentDataset.columns
      );

      validateAnalysisPlan(
        normalizedPlan,
        currentDataset.columns
      );

      const result = executeAnalysisPlan(
        currentDataset.rows,
        normalizedPlan
      );

      return res.json({
        success: true,
        question,
        analysisPlan: normalizedPlan,
        result
      });

    } catch (error) {
      console.error("Analysis error:", error);

      return res.status(500).json({
        success: false,
        error:
          error?.message ||
          "Failed to analyze the question."
      });
    }
  }
);

/* ============================================================
   CLEANING PREVIEW
   ============================================================ */

app.post("/api/cleaning/preview", (req, res) => {
  try {
    if (!currentDataset.rows.length) {
      return res.status(400).json({
        success: false,
        error: "Please upload a dataset first."
      });
    }

    const options = getCleaningOptions(req.body);

    const { rows, summary } = applyCleaningOptions(
      currentDataset.rows,
      currentDataset.columns,
      options
    );

    return res.json({
      success: true,
      options,
      summary,
      columns: currentDataset.columns,
      preview: rows.slice(0, 10)
    });

  } catch (error) {
    console.error("Cleaning preview error:", error);

    return res.status(500).json({
      success: false,
      error:
        error?.message ||
        "Could not preview cleaning changes."
    });
  }
});

/* ============================================================
   CLEANED EXCEL EXPORT
   ============================================================ */

app.post("/api/cleaning/export", (req, res) => {
  try {
    if (!currentDataset.rows.length) {
      return res.status(400).json({
        success: false,
        error: "Please upload a dataset first."
      });
    }

    if (currentWorkbookState) {
      return res.status(409).json({
        success: false,
        error:
          "Cleaned export is unavailable for this XLSX workbook until feature-preservation checks are implemented."
      });
    }

    const options = getCleaningOptions(req.body);

    const { rows, summary } = applyCleaningOptions(
      currentDataset.rows,
      currentDataset.columns,
      options
    );

    // Build a new workbook; never change the active dataset.
    const exportColumns = currentDataset.columns.filter(
      (column) => {
        const header = String(column ?? "").trim();

        // Exclude old-style empty Excel placeholder columns
        // only when they contain no actual values.
        if (/^__EMPTY(?:_\d+)?$/i.test(header)) {
          return rows.some(
            (row) => !isEmptyValue(row[column])
          );
        }

        return true;
      }
    );

    const exportRows = rows.map((row) => {
      const cleanedRow = {};

      for (const column of exportColumns) {
        cleanedRow[column] = row[column] ?? "";
      }

      return cleanedRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(
      exportRows,
      {
        header: exportColumns,
        skipHeader: false
      }
    );

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Cleaned Data"
    );

    const fileBuffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx"
    });

    const originalBaseName = path.parse(
      currentDataset.fileName || "dataset"
    ).name;

    const safeBaseName = originalBaseName
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
      .slice(0, 100) || "dataset";

    const downloadName = `cleaned-${safeBaseName}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${downloadName}"`
    );

    res.setHeader(
      "X-Cleaning-Summary",
      JSON.stringify(summary)
    );

    return res.send(fileBuffer);

  } catch (error) {
    console.error("Cleaning export error:", error);

    return res.status(500).json({
      success: false,
      error:
        error?.message ||
        "Could not export the cleaned Excel file."
    });
  }
});

/* ============================================================
   START SERVER
   ============================================================ */

app.listen(PORT, () => {
  console.log(
    `InsightMate server running at http://localhost:${PORT}`
  );

  console.log(`Server version: ${SERVER_VERSION}`);
  console.log(`Ollama model: ${OLLAMA_MODEL}`);
});
