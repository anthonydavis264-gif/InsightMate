import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { watch } from "node:fs";
import fs from "node:fs/promises";
import http from "node:http";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { after, before, test } from "node:test";
import ExcelJS from "exceljs";
import XLSX from "xlsx";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  ".."
);

let tempRoot;
let appBaseUrl;
let mockOllama;
let appProcess;
let appOutput = "";
let csvFile;
let xlsxFile;
let offsetHeaderXlsxFile;
let missingRecordsXlsxFile;
let uploadDirectory;

async function getAvailablePort() {
  const server = net.createServer();
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const { port } = server.address();
  await new Promise((resolve, reject) => {
    server.close((error) => error ? reject(error) : resolve());
  });
  return port;
}

function getMockAnalysisPlan(prompt) {
  if (prompt.includes("User question:\n\"How many orders are there?\"")) {
    return {
      operation: "COUNT",
      groupBy: null,
      metric: null,
      aggregation: null,
      filters: [],
      sort: null,
      limit: 0
    };
  }

  if (prompt.includes("User question:\n\"What is the total amount by Category?\"")) {
    return {
      operation: "GROUP_BY",
      groupBy: "Category",
      metric: "Amount",
      aggregation: "SUM",
      filters: [],
      sort: "DESC",
      limit: 0
    };
  }

  return {
    operation: "AGGREGATE",
    groupBy: null,
    metric: "Amount",
    aggregation: "SUM",
    filters: [],
    sort: null,
    limit: 0
  };
}

async function startMockOllama() {
  mockOllama = http.createServer(async (request, response) => {
    if (request.method !== "POST" || request.url !== "/api/generate") {
      response.writeHead(404).end();
      return;
    }

    let body = "";
    for await (const chunk of request) {
      body += chunk;
    }

    let prompt;
    try {
      prompt = JSON.parse(body).prompt;
    } catch {
      response.writeHead(400).end();
      return;
    }

    const plan = getMockAnalysisPlan(String(prompt || ""));
    response.writeHead(200, { "Content-Type": "application/json" });
    response.end(JSON.stringify({ response: JSON.stringify(plan) }));
  });

  mockOllama.listen(0, "127.0.0.1");
  await once(mockOllama, "listening");
  return mockOllama.address().port;
}

async function waitForApplication() {
  const deadline = Date.now() + 10000;

  while (Date.now() < deadline) {
    if (appProcess.exitCode !== null) {
      throw new Error(`Application exited during startup.\n${appOutput}`);
    }

    try {
      const response = await fetch(`${appBaseUrl}/api/health`);
      if (response.ok) {
        return;
      }
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }

  throw new Error(`Application did not become ready.\n${appOutput}`);
}

async function uploadFile(filePath, fileName, contentType) {
  const bytes = await fs.readFile(filePath);
  const form = new FormData();
  form.append("file", new Blob([bytes], { type: contentType }), fileName);

  return fetch(`${appBaseUrl}/api/upload`, {
    method: "POST",
    body: form
  });
}

async function startPausedUpload(fileName, bytes) {
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
  const boundary = `----insightmate-${Date.now()}`;
  const prefix = Buffer.from(
    `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="file"; filename="${fileName}"\r\n` +
      "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n",
    "utf8"
  );
  const suffix = Buffer.from(`\r\n--${boundary}--\r\n`, "utf8");
  const firstChunkLength = Math.min(128, bytes.length);
  const request = http.request(new URL("/api/upload", appBaseUrl), {
    method: "POST",
    headers: {
      "Content-Type": `multipart/form-data; boundary=${boundary}`,
      "Content-Length": prefix.length + bytes.length + suffix.length
    }
  });
  const response = new Promise((resolve, reject) => {
    request.once("error", reject);
    request.once("response", (responseMessage) => {
      const chunks = [];
      responseMessage.on("data", (chunk) => chunks.push(chunk));
      responseMessage.once("error", reject);
      responseMessage.once("end", () => {
        resolve({
          status: responseMessage.statusCode,
          data: JSON.parse(Buffer.concat(chunks).toString("utf8"))
        });
      });
    });
  });

  const fileCreated = new Promise((resolve, reject) => {
    let watcher;
    const timeout = setTimeout(() => {
      const error = new Error("Upload temporary file was not created.");
      watcher.close();
      request.destroy(error);
      reject(error);
    }, 30000);
    watcher = watch(uploadDirectory, { persistent: false }, async () => {
      try {
        const files = await fs.readdir(uploadDirectory);
        if (files.length > 0) {
          clearTimeout(timeout);
          watcher.close();
          resolve();
        }
      } catch (error) {
        clearTimeout(timeout);
        watcher.close();
        reject(error);
      }
    });
    watcher.once("error", (error) => {
      clearTimeout(timeout);
      watcher.close();
      reject(error);
    });
  });

  request.setTimeout(30000, () => {
    request.destroy(new Error("Paused upload request timed out."));
  });
  request.write(prefix);
  request.write(bytes.subarray(0, firstChunkLength));
  await fileCreated;

  return {
    response,
    finish() {
      request.end(Buffer.concat([
        bytes.subarray(firstChunkLength),
        suffix
      ]));
    }
  };
}

async function createWorkbookBytes(sheetName, rows) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(sheetName);
  for (const row of rows) {
    worksheet.addRow(row);
  }
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

async function postJson(endpoint, body) {
  return fetch(`${appBaseUrl}${endpoint}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
}

before(async () => {
  tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "insightmate-tests-"));
  uploadDirectory = path.join(tempRoot, "uploads");
  await fs.mkdir(uploadDirectory);
  const [appPort, ollamaPort] = await Promise.all([
    getAvailablePort(),
    startMockOllama()
  ]);
  appBaseUrl = `http://127.0.0.1:${appPort}`;

  const csv = [
    "Category,Amount,Note",
    "North,10, kept ",
    "South,20,",
    "North,10, kept ",
    ",5,  ",
    ",,"
  ].join("\n");
  csvFile = path.join(tempRoot, "synthetic.csv");
  await fs.writeFile(csvFile, csv, "utf8");

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Category", "Amount", "Amount"],
    ["East", 3, 4],
    ["West", null, 6]
  ]);
  XLSX.utils.book_append_sheet(workbook, sheet, "Synthetic");
  xlsxFile = path.join(tempRoot, "synthetic.xlsx");
  XLSX.writeFile(workbook, xlsxFile);

  const headers = [
    "Date",
    "Client",
    "Contact",
    "Department",
    "Payment",
    "Revenue",
    "Profit",
    "Profit Margin"
  ];
  const records = Array.from({ length: 31 }, (_, index) => {
    const recordNumber = index + 1;
    return [
      "",
      `2023-05-${String(recordNumber).padStart(2, "0")}`,
      `Client ${recordNumber}`,
      `Contact ${recordNumber}`,
      recordNumber === 31
        ? "Final Department"
        : recordNumber % 2 === 1
          ? "North"
          : "South",
      recordNumber === 31
        ? "Final Payment"
        : recordNumber % 2 === 1
          ? "Card"
          : "Cash",
      recordNumber * 10,
      recordNumber,
      0.1
    ];
  });
  const offsetHeaderWorkbook = XLSX.utils.book_new();
  const offsetHeaderSheet = XLSX.utils.aoa_to_sheet([
    [],
    ["", ...headers],
    ...records
  ]);
  offsetHeaderSheet["!ref"] = "A1:I33";
  XLSX.utils.book_append_sheet(
    offsetHeaderWorkbook,
    offsetHeaderSheet,
    "data"
  );
  offsetHeaderXlsxFile = path.join(tempRoot, "offset-header.xlsx");
  XLSX.writeFile(offsetHeaderWorkbook, offsetHeaderXlsxFile);

  const missingRecordsWorkbook = XLSX.utils.book_new();
  const missingRecordsSheet = XLSX.utils.aoa_to_sheet([
    ["Client", "Contact", "Payment", "Revenue", "Profit Margin"],
    ["Alpha", "555-0101", "Card", 10, 0.1],
    ["Beta", "555-0102", "", 20, 0.2],
    ["Gamma", "555-0103", "Cash", "", ""],
    ["Delta", "555-0104", "", 40, 0.4]
  ]);
  XLSX.utils.book_append_sheet(missingRecordsWorkbook, missingRecordsSheet, "data");
  missingRecordsXlsxFile = path.join(tempRoot, "missing-records.xlsx");
  XLSX.writeFile(missingRecordsWorkbook, missingRecordsXlsxFile);

  appProcess = spawn(process.execPath, ["server.mjs"], {
    cwd: projectRoot,
    env: {
      ...process.env,
      PORT: String(appPort),
      OLLAMA_URL: `http://127.0.0.1:${ollamaPort}`,
      UPLOAD_DIR: uploadDirectory
    },
    stdio: ["ignore", "pipe", "pipe"]
  });

  appProcess.stdout.setEncoding("utf8");
  appProcess.stderr.setEncoding("utf8");
  appProcess.stdout.on("data", (data) => { appOutput += data; });
  appProcess.stderr.on("data", (data) => { appOutput += data; });

  await waitForApplication();
});

after(async () => {
  if (appProcess && appProcess.exitCode === null) {
    const exited = once(appProcess, "exit");
    appProcess.kill();
    await exited;
  }

  if (mockOllama?.listening) {
    await new Promise((resolve, reject) => {
      mockOllama.close((error) => error ? reject(error) : resolve());
    });
  }

  if (tempRoot) {
    await fs.rm(tempRoot, { recursive: true, force: true });
  }
});

test("health reports the actual server contract", async () => {
  const response = await fetch(`${appBaseUrl}/api/health`);
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.deepEqual(
    {
      success: data.success,
      status: data.status,
      serverVersion: data.serverVersion,
      ollamaModel: data.ollamaModel,
      datasetLoaded: data.datasetLoaded
    },
    {
      success: true,
      status: "ok",
      serverVersion: "analysis-engine-v4.4",
      ollamaModel: "gemma4:e2b-it-qat",
      datasetLoaded: false
    }
  );
});

test("cleaning preview requires an uploaded dataset", async () => {
  const response = await postJson("/api/cleaning/preview", { options: {} });
  const data = await response.json();

  assert.equal(response.status, 400);
  assert.equal(data.success, false);
  assert.match(data.error, /upload a dataset/i);
});

test("CSV upload preserves headers, numeric values, and empty cells", async () => {
  const response = await uploadFile(
    csvFile,
    "synthetic.csv",
    "text/csv"
  );
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.equal(data.success, true);
  assert.equal(data.fileName, "synthetic.csv");
  assert.deepEqual(data.columnNames, ["Category", "Amount", "Note"]);
  assert.equal(data.rows, 5);
  assert.equal(data.columns, 3);
  assert.equal(data.preview[0].Amount, 10);
  assert.equal(data.preview[1].Note, "");
  assert.equal(data.preview[4].Amount, "");
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
});

test("cleaning preview trims and removes empty rows without changing analysis data", async () => {
  const response = await postJson("/api/cleaning/preview", {
    options: {
      trimWhitespace: true,
      removeEmptyRows: true,
      removeDuplicates: false
    }
  });
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.deepEqual(data.options, {
    trimWhitespace: true,
    removeEmptyRows: true,
    removeDuplicates: true
  });
  assert.equal(data.summary.originalRows, 5);
  assert.equal(data.summary.trimmedCells, 3);
  assert.equal(data.summary.emptyRowsRemoved, 1);
  assert.equal(data.summary.duplicateRowsRemoved, 1);
  assert.equal(data.summary.finalRows, 3);
  assert.equal(data.preview[0].Note, "kept");

  const datasetResponse = await fetch(`${appBaseUrl}/api/dataset`);
  const dataset = await datasetResponse.json();
  assert.equal(dataset.rows, 5);
  assert.equal(dataset.preview[0].Note, " kept ");
});

test("analysis route returns deterministic aggregate, group, and count shapes", async () => {
  const aggregateResponse = await postJson("/api/analyze-question", {
    question: "What is the total amount?"
  });
  const aggregate = await aggregateResponse.json();

  assert.equal(aggregateResponse.status, 200);
  assert.equal(aggregate.success, true);
  assert.equal(aggregate.result.type, "aggregate");
  assert.equal(aggregate.result.metric, "Amount");
  assert.equal(aggregate.result.aggregation, "SUM");
  assert.equal(aggregate.result.value, 45);
  assert.equal(aggregate.analysisPlan.operation, "AGGREGATE");

  const groupResponse = await postJson("/api/analyze-question", {
    question: "What is the total amount by Category?"
  });
  const grouped = await groupResponse.json();

  assert.equal(groupResponse.status, 200);
  assert.equal(grouped.result.type, "group_by");
  assert.deepEqual(grouped.result.rows, [
    { group: "North", value: 20 },
    { group: "South", value: 20 },
    { group: "Unknown", value: 5 }
  ]);

  const countResponse = await postJson("/api/analyze-question", {
    question: "How many orders are there?"
  });
  const count = await countResponse.json();

  assert.equal(countResponse.status, 200);
  assert.equal(count.result.type, "count");
  assert.equal(count.result.value, 5);
});

test("missing-value questions return one missing count per column without changing record counts", async () => {
  const uploadResponse = await uploadFile(csvFile, "synthetic.csv", "text/csv");
  assert.equal(uploadResponse.status, 200);

  const questions = [
    "How many missing values are there in each column?",
    "For each column, count the blank cells and show the column name and missing-value count in a table."
  ];

  for (const question of questions) {
    const response = await postJson("/api/analyze-question", { question });
    const data = await response.json();

    assert.equal(response.status, 200, question);
    assert.equal(data.analysisPlan.operation, "MISSING_VALUES");
    assert.equal(data.result.type, "missing_values");
    assert.equal(data.result.rows.length, 3);
    assert.deepEqual(data.result.rows, [
      { column: "Category", missingCount: 2 },
      { column: "Amount", missingCount: 1 },
      { column: "Note", missingCount: 3 }
    ]);
  }

  const countResponse = await postJson("/api/analyze-question", {
    question: "How many orders are there?"
  });
  const count = await countResponse.json();

  assert.equal(countResponse.status, 200);
  assert.equal(count.analysisPlan.operation, "COUNT");
  assert.equal(count.result.type, "count");
  assert.equal(count.result.value, 5);
});

test("missing-record question returns affected record ordinals and blank column names only", async () => {
  const uploadResponse = await uploadFile(
    missingRecordsXlsxFile,
    "missing-records.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(uploadResponse.status, 200);

  const response = await postJson("/api/analyze-question", {
    question: "Which records have missing values, and which columns are blank?"
  });
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.equal(data.analysisPlan.operation, "MISSING_RECORDS");
  assert.equal(data.result.type, "missing_records");
  assert.deepEqual(data.result.rows, [
    { recordOrdinal: 2, blankColumns: ["Payment"] },
    { recordOrdinal: 3, blankColumns: ["Revenue", "Profit Margin"] },
    { recordOrdinal: 4, blankColumns: ["Payment"] }
  ]);
  assert.ok(data.result.rows.every((row) =>
    Object.keys(row).sort().join(",") === "blankColumns,recordOrdinal"
  ));
  assert.equal(JSON.stringify(data.result).includes("555-010"), false);
  assert.equal(JSON.stringify(data.result).includes("Alpha"), false);
});

test("XLSX upload normalizes duplicate headers and retains numeric and missing values", async () => {
  const response = await uploadFile(
    xlsxFile,
    "synthetic.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.equal(data.success, true);
  assert.equal(data.sheetName, "Synthetic");
  assert.deepEqual(data.columnNames, ["Category", "Amount", "Amount (2)"]);
  assert.equal(data.rows, 2);
  assert.equal(data.preview[0]["Amount (2)"], 4);
  assert.equal(data.preview[1].Amount, "");
  assert.equal(data.preview[1]["Amount (2)"], 6);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
});

test("XLSX upload skips leading blank rows and preserves offset header indexes for analysis", async () => {
  const response = await uploadFile(
    offsetHeaderXlsxFile,
    "offset-header.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.deepEqual(data.columnNames, [
    "Date",
    "Client",
    "Contact",
    "Department",
    "Payment",
    "Revenue",
    "Profit",
    "Profit Margin"
  ]);
  assert.equal(data.rows, 31);
  assert.deepEqual(
    {
      Date: data.preview[0].Date,
      Client: data.preview[0].Client,
      Contact: data.preview[0].Contact,
      Department: data.preview[0].Department,
      Payment: data.preview[0].Payment,
      Revenue: data.preview[0].Revenue,
      Profit: data.preview[0].Profit,
      "Profit Margin": data.preview[0]["Profit Margin"]
    },
    {
      Date: "2023-05-01",
      Client: "Client 1",
      Contact: "Contact 1",
      Department: "North",
      Payment: "Card",
      Revenue: 10,
      Profit: 1,
      "Profit Margin": 0.1
    }
  );
  const analysisResponse = await postJson("/api/analyze-question", {
    question: "What is the total revenue by department?"
  });
  const analysis = await analysisResponse.json();

  assert.equal(analysisResponse.status, 200);
  assert.deepEqual(
    {
      operation: analysis.analysisPlan.operation,
      groupBy: analysis.analysisPlan.groupBy,
      metric: analysis.analysisPlan.metric,
      aggregation: analysis.analysisPlan.aggregation
    },
    {
      operation: "GROUP_BY",
      groupBy: "Department",
      metric: "Revenue",
      aggregation: "SUM"
    }
  );
  assert.deepEqual(analysis.result.rows, [
    { group: "South", value: 2400 },
    { group: "North", value: 2250 },
    { group: "Final Department", value: 310 }
  ]);

  const lastRecordChecks = [
    ["What is the total revenue by date?", "Revenue", "2023-05-31", 310],
    ["What is the total revenue by client?", "Revenue", "Client 31", 310],
    ["What is the total revenue by contact?", "Revenue", "Contact 31", 310],
    ["What is the total revenue by payment?", "Revenue", "Final Payment", 310],
    ["What is the total profit by department?", "Profit", "Final Department", 31],
    [
      "What is the total profit margin by department?",
      "Profit Margin",
      "Final Department",
      0.1
    ]
  ];

  for (const [question, metric, group, value] of lastRecordChecks) {
    const lastRecordResponse = await postJson("/api/analyze-question", {
      question
    });
    const lastRecordAnalysis = await lastRecordResponse.json();

    assert.equal(lastRecordResponse.status, 200, question);
    assert.equal(lastRecordAnalysis.analysisPlan.metric, metric);
    assert.ok(lastRecordAnalysis.result.rows.some((row) =>
      row.group === group && row.value === value
    ), question);
  }

  const restoreResponse = await uploadFile(
    xlsxFile,
    "synthetic.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(restoreResponse.status, 200);
});

test("Excel cleaning is upload-bound, removes duplicates automatically, and downloads only the cleaned result", async () => {
  const workbook = new ExcelJS.Workbook();
  const data = workbook.addWorksheet("Orders");
  const hidden = workbook.addWorksheet("Hidden Notes", { state: "hidden" });
  const contacts = workbook.addWorksheet("Contacts");
  const audit = workbook.addWorksheet("Formula Audit");
  data.addRow(["Item", "Amount"]);
  data.addRow(["Same", 10]);
  data.addRow([]);
  data.addRow(["Same", 10]);
  data.addRow(["Same", 10]);
  data.addRow(["Same", 10]);
  data.getCell("B4").font = { color: { argb: "FFFF0000" } };
  data.addRow(["Tail", 20]);
  data.getColumn(1).width = 24;
  for (const rowNumber of [2, 4, 5, 6]) {
    data.getCell(`B${rowNumber}`).numFmt = "$#,##0.00";
  }
  hidden.addRow(["Note", "Value"]);
  hidden.addRow(["Hidden row", 3]);
  hidden.getCell("C1").value = "Result";
  hidden.getCell("C2").value = { error: "#DIV/0!" };
  contacts.addRows([
    ["Name", "Phone", "Salesperson"],
    ["Customer A", "+1 (555) 111-1111", "Alex"],
    ["Customer B", "+1 (555) 222-2222", "Sam"],
    ["Customer C", "+1 (555) 123-4567", null]
  ]);
  contacts.mergeCells("B4:C4");
  audit.addRows([
    ["Expression"],
    [],
    [{ formula: "1+1", result: 2 }]
  ]);
  audit.getCell("A2").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };
  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const sourceFile = path.join(tempRoot, "cleaner-source.xlsx");
  await fs.writeFile(sourceFile, originalBytes);

  const uploadResponse = await uploadFile(
    sourceFile,
    "cleaner-source.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const upload = await uploadResponse.json();
  assert.equal(uploadResponse.status, 200);
  assert.equal(typeof upload.uploadVersionId, "string");
  assert.equal(typeof upload.cleaningToken, "string");

  const invalidStart = await postJson("/api/cleaning/start", {
    uploadVersionId: upload.uploadVersionId
  });
  assert.equal(invalidStart.status, 404);
  assert.equal((await invalidStart.json()).success, false);

  const credentials = {
    uploadVersionId: upload.uploadVersionId,
    cleaningToken: upload.cleaningToken
  };
  const startResponse = await postJson("/api/cleaning/start", credentials);
  const started = await startResponse.json();
  assert.equal(startResponse.status, 200);
  assert.equal(started.summary.worksheetCount, 4);
  assert.deepEqual(started.summary.worksheets.map(({ name, order }) => ({
    name,
    order
  })), [
    { name: "Orders", order: 0 },
    { name: "Hidden Notes", order: 1 },
    { name: "Contacts", order: 2 },
    { name: "Formula Audit", order: 3 }
  ]);
  assert.equal(started.summary.worksheets[0].spacesTrimmed, 0);
  assert.equal(started.summary.worksheets[0].emptyRowsRemoved, 1);
  assert.equal(started.summary.worksheets[0].blankRowsDetected, 1);
  assert.equal(started.summary.worksheets[0].blankRowsRemoved, 0);
  assert.equal(started.summary.worksheets[0].blankRowsToRemove, 1);
  assert.equal(started.summary.worksheets[0].blankRowsRetained, 0);
  assert.equal(started.summary.worksheets[0].duplicateCandidatesFound, 3);
  assert.equal(started.summary.worksheets[1].duplicateCandidatesFound, 0);
  assert.deepEqual(started.summary.worksheets[2].mergedRangeActions, [{
    worksheetName: "Contacts",
    originalRange: "B4:C4",
    reason: "This horizontal merge spans distinct data fields, and 2 nearby records contain independent values in those columns. The text or other value anchor remains in its original column.",
    action: "Unmerged separate data fields and preserved the anchor value in its original column.",
    unresolvedMissingFields: ["Salesperson"]
  }]);
  assert.equal(started.summary.worksheets[0].duplicateCandidates[0].safeToRemove, true);
  assert.equal(started.summary.worksheets[0].duplicateCandidates[0].originalRow, 4);
  assert.equal(started.summary.worksheets[0].duplicateCandidates[0].retainedOriginalRow, 2);
  assert.equal(started.summary.worksheets[3].blankRowsDetected, 1);
  assert.equal(started.summary.worksheets[3].blankRowsRemoved, 0);
  assert.equal(started.summary.worksheets[3].blankRowsToRemove, 0);
  assert.equal(started.summary.worksheets[3].blankRowsRetained, 1);
  assert.deepEqual(started.summary.worksheets[3].blankRowsRetainedByReason, [{
    reason: "A blank row was retained because a reliable multi-column header was not identified for this worksheet.",
    count: 1
  }]);

  const cleanerScriptResponse = await fetch(`${appBaseUrl}/app.js`);
  assert.equal(cleanerScriptResponse.status, 200);
  const cleanerScript = await cleanerScriptResponse.text();
  assert.match(cleanerScript, /Blank rows detected:/);
  assert.match(cleanerScript, /worksheet\.blankRowsRetainedByReason/);
  assert.match(cleanerScript, /Blank rows retained:/);

  const prematureDownload = await fetch(`${appBaseUrl}/api/cleaning/download`, {
    headers: {
      "X-Upload-Version-Id": upload.uploadVersionId,
      "X-Cleaning-Token": upload.cleaningToken
    }
  });
  assert.equal(prematureDownload.status, 409);
  const prematureError = await prematureDownload.json();
  assert.equal(prematureError.success, false);
  assert.equal(JSON.stringify(prematureError).includes("Same"), false);

  const analysisBefore = await (await postJson("/api/analyze-question", {
    question: "What is the total amount?"
  })).json();
  const generateResponse = await postJson("/api/cleaning/generate", {
    ...credentials
  });
  const generated = await generateResponse.json();
  assert.equal(generateResponse.status, 200);
  assert.equal(generated.downloadReady, true);
  assert.equal(generated.summary.worksheets[0].removedRows, 3);
  assert.equal(generated.summary.worksheets[0].blankRowsDetected, 1);
  assert.equal(generated.summary.worksheets[0].blankRowsRemoved, 1);
  assert.equal(generated.summary.worksheets[0].blankRowsToRemove, 0);
  assert.equal(generated.summary.worksheets[0].blankRowsRetained, 0);
  assert.deepEqual(generated.summary.worksheets[0].duplicateRowsSkipped, []);
  assert.deepEqual(generated.summary.worksheets[1].errorReplacements, [
    {
      worksheetName: "Hidden Notes",
      address: "C2",
      originalError: "#DIV/0!",
      originalFormula: null,
      action: "Replaced Excel error with a blank cell"
    }
  ]);

  const downloadResponse = await fetch(`${appBaseUrl}/api/cleaning/download`, {
    headers: {
      "X-Upload-Version-Id": upload.uploadVersionId,
      "X-Cleaning-Token": upload.cleaningToken
    }
  });
  assert.equal(downloadResponse.status, 200);
  assert.equal(
    downloadResponse.headers.get("content-type"),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.match(downloadResponse.headers.get("content-disposition"), /cleaner-source_cleaned\.xlsx/);
  const downloadedBytes = Buffer.from(await downloadResponse.arrayBuffer());
  const downloadedWorkbook = new ExcelJS.Workbook();
  await downloadedWorkbook.xlsx.load(downloadedBytes);
  assert.deepEqual(downloadedWorkbook.worksheets.map(({ name, state }) => ({
    name,
    state
  })), [
    { name: "Orders", state: "visible" },
    { name: "Hidden Notes", state: "hidden" },
    { name: "Contacts", state: "visible" },
    { name: "Formula Audit", state: "visible" }
  ]);
  const downloadedData = downloadedWorkbook.getWorksheet("Orders");
  assert.equal(downloadedData.rowCount, 3);
  assert.deepEqual(
    [2, 3].map((row) => downloadedData.getRow(row).values.slice(1)),
    [["Same", 10], ["Tail", 20]]
  );
  assert.equal(downloadedData.getColumn(1).width, 24);
  assert.ok(downloadedData.getColumn(2).width >= 10);
  assert.ok(downloadedData.getColumn(2).width <= 48);
  assert.equal(downloadedData.getCell("A1").font.bold, true);
  assert.equal(downloadedData.getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(downloadedData.getCell("A1").font.color.argb, "FFFFFFFF");
  assert.equal(downloadedData.getCell("A1").alignment.wrapText, true);
  assert.equal(downloadedData.getCell("A1").alignment.horizontal, "center");
  assert.equal(downloadedData.getCell("A1").alignment.vertical, "middle");
  assert.equal(downloadedData.getCell("A2").alignment.horizontal, "center");
  assert.equal(downloadedData.getCell("A2").alignment.vertical, "middle");
  assert.equal(downloadedData.getCell("B2").alignment.horizontal, "center");
  assert.equal(downloadedData.getCell("B2").alignment.vertical, "middle");
  assert.ok(downloadedData.getRow(1).height >= 26);
  assert.equal(downloadedData.getCell("B2").numFmt, "$#,##0.00");
  const downloadedHidden = downloadedWorkbook.getWorksheet("Hidden Notes");
  assert.equal(downloadedHidden.getCell("A1").font.bold, true);
  assert.equal(downloadedHidden.getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(downloadedHidden.getCell("A1").alignment.horizontal, "center");
  assert.equal(downloadedHidden.getCell("A2").alignment.horizontal, "center");
  assert.equal(downloadedHidden.getCell("B2").alignment.horizontal, "center");
  assert.equal(downloadedHidden.getCell("C2").value, null);
  const downloadedContacts = downloadedWorkbook.getWorksheet("Contacts");
  assert.equal(downloadedContacts.getCell("B4").value, "+1 (555) 123-4567");
  assert.equal(downloadedContacts.getCell("B4").isMerged, false);
  assert.equal(downloadedContacts.getCell("C4").isMerged, false);
  assert.equal(downloadedContacts.getCell("C4").value, null);
  assert.equal(downloadedWorkbook.getWorksheet("Formula Audit").getCell("A2").value, null);
  assert.equal(downloadedWorkbook.getWorksheet("Formula Audit").getCell("A2").fill.fgColor.argb, "FFE0F0FF");
  assert.deepEqual(downloadedWorkbook.getWorksheet("Formula Audit").getCell("A3").value, {
    formula: "1+1",
    result: 2
  });
  assert.deepEqual(await fs.readFile(sourceFile), originalBytes);
  assert.ok(downloadedBytes.length > 0);

  const analysisAfter = await (await postJson("/api/analyze-question", {
    question: "What is the total amount?"
  })).json();
  assert.deepEqual(analysisAfter.result, analysisBefore.result);

  const otherWorkbook = new ExcelJS.Workbook();
  otherWorkbook.addWorksheet("Private Sheet").addRows([
    ["Secret", "Amount"],
    ["Only this upload", 999]
  ]);
  const otherBytes = Buffer.from(await otherWorkbook.xlsx.writeBuffer());
  const otherFile = path.join(tempRoot, "other-cleaner-source.xlsx");
  await fs.writeFile(otherFile, otherBytes);
  const otherUploadResponse = await uploadFile(
    otherFile,
    "other-cleaner-source.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const otherUpload = await otherUploadResponse.json();
  assert.equal(otherUploadResponse.status, 200);

  const oldSessionStart = await postJson("/api/cleaning/start", credentials);
  assert.equal(oldSessionStart.status, 409);
  const stolenDownload = await fetch(`${appBaseUrl}/api/cleaning/download`, {
    headers: {
      "X-Upload-Version-Id": upload.uploadVersionId,
      "X-Cleaning-Token": upload.cleaningToken
    }
  });
  assert.equal(stolenDownload.status, 409);
  const newSessionDownload = await fetch(`${appBaseUrl}/api/cleaning/download`, {
    headers: {
      "X-Upload-Version-Id": otherUpload.uploadVersionId,
      "X-Cleaning-Token": otherUpload.cleaningToken
    }
  });
  assert.equal(newSessionDownload.status, 409);
  const unavailableDownload = await newSessionDownload.json();
  assert.equal(JSON.stringify(unavailableDownload).includes("Only this upload"), false);

  const secondStart = await (await postJson("/api/cleaning/start", {
    uploadVersionId: otherUpload.uploadVersionId,
    cleaningToken: otherUpload.cleaningToken
  })).json();
  const staleCandidate = await postJson("/api/cleaning/generate", {
    ...credentials
  });
  assert.equal(staleCandidate.status, 409);
  assert.equal(secondStart.summary.worksheetCount, 1);
  assert.deepEqual(await fs.readFile(sourceFile), originalBytes);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);

  const restoreUpload = await uploadFile(
    xlsxFile,
    "synthetic.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(restoreUpload.status, 200);
});

test("cleaned export stays unavailable for XLSX until feature preservation is verified", async () => {
  const response = await postJson("/api/cleaning/export", {
    options: {
      trimWhitespace: true,
      removeEmptyRows: false,
      removeDuplicates: false
    }
  });
  const data = await response.json();

  assert.equal(response.status, 409);
  assert.equal(data.success, false);
  assert.match(data.error, /feature-preservation checks/i);
});

test("analysis calculations use the uploaded XLSX dataset", async () => {
  const response = await postJson("/api/analyze-question", {
    question: "What is the total amount?"
  });
  const data = await response.json();

  assert.equal(response.status, 200);
  assert.equal(data.result.type, "aggregate");
  assert.equal(data.result.metric, "Amount");
  assert.equal(data.result.value, 3);
});

test("invalid workbook returns a meaningful error and leaves active data intact", async () => {
  const invalidFile = path.join(tempRoot, "invalid.xlsx");
  await fs.writeFile(invalidFile, Buffer.from([0x50, 0x4b, 0x03, 0x04, 0xff, 0x00]));

  const response = await uploadFile(
    invalidFile,
    "invalid.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  const data = await response.json();

  assert.equal(response.status, 500);
  assert.equal(data.success, false);
  assert.match(data.error, /file|zip|workbook|format/i);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);

  const datasetResponse = await fetch(`${appBaseUrl}/api/dataset`);
  const dataset = await datasetResponse.json();
  assert.equal(dataset.fileName, "synthetic.xlsx");

  const exportResponse = await postJson("/api/cleaning/export", {
    options: {}
  });
  assert.equal(exportResponse.status, 409);
});

test("analysis route reports missing question as a client error", async () => {
  const response = await postJson("/api/analyze-question", { question: " " });
  const data = await response.json();

  assert.equal(response.status, 400);
  assert.equal(data.success, false);
  assert.match(data.error, /enter a question/i);
});

test("workbook proposals are version-bound, individually decidable, and XLSX export stays blocked", async () => {
  const workbook = new ExcelJS.Workbook();
  const dataSheet = workbook.addWorksheet("Approval Data");
  const uncertainSheet = workbook.addWorksheet("Uncertain");
  dataSheet.addRow(["Name", "Amount"]);
  dataSheet.addRow(["Alpha", 10]);
  dataSheet.addRow(["Alpha", 10]);
  dataSheet.addRow([" Beta ", 20]);
  uncertainSheet.addRow([123]);
  const bytes = await workbook.xlsx.writeBuffer();
  const proposalFile = path.join(tempRoot, "synthetic-proposals.xlsx");
  await fs.writeFile(proposalFile, bytes);

  const uploadResponse = await uploadFile(
    proposalFile,
    "synthetic-proposals.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(uploadResponse.status, 200);

  const workbookResponse = await fetch(`${appBaseUrl}/api/cleaning/workbook`);
  const workbookData = await workbookResponse.json();
  assert.equal(workbookData.success, true);
  assert.ok(workbookData.uploadVersionId);
  assert.deepEqual(
    workbookData.worksheets.map(({ worksheetKey, name, order }) => ({
      worksheetKey,
      name,
      order
    })),
    [
      { worksheetKey: "sheet-0", name: "Approval Data", order: 0 },
      { worksheetKey: "sheet-1", name: "Uncertain", order: 1 }
    ]
  );
  assert.equal(
    workbookData.worksheets[1].header.status,
    "confirmation_required"
  );

  const generatedResponse = await postJson("/api/cleaning/proposals", {
    uploadVersionId: workbookData.uploadVersionId
  });
  const generated = await generatedResponse.json();
  assert.equal(generatedResponse.status, 200);
  assert.equal(generated.worksheets[0].status, "processed");
  assert.equal(generated.worksheets[1].status, "skipped");

  const duplicate = generated.proposals.find(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.ok(duplicate);
  assert.equal(duplicate.target.row, 3);
  assert.equal(duplicate.matches.retainedRow, 2);
  assert.equal(duplicate.decision, "pending");

  const datasetBeforeDecision = await (await fetch(
    `${appBaseUrl}/api/dataset`
  )).json();
  const analysisBeforeDecision = await (await postJson(
    "/api/analyze-question",
    { question: "What is the total amount?" }
  )).json();
  const decisionResponse = await postJson(
    `/api/cleaning/proposals/${duplicate.proposalId}/decision`,
    {
      uploadVersionId: generated.uploadVersionId,
      reviewId: generated.reviewId,
      decision: "approved"
    }
  );
  assert.equal(decisionResponse.status, 200);
  assert.equal((await decisionResponse.json()).proposal.decision, "approved");
  assert.deepEqual(await fs.readFile(proposalFile), bytes);

  const failedUploadFile = path.join(tempRoot, "invalid-proposal-upload.xlsx");
  await fs.writeFile(failedUploadFile, Buffer.from([0x50, 0x4b, 0x03, 0x04, 0xff]));
  const failedUpload = await uploadFile(
    failedUploadFile,
    "invalid-proposal-upload.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(failedUpload.status, 500);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
  const afterFailure = await (await fetch(
    `${appBaseUrl}/api/cleaning/workbook`
  )).json();
  assert.equal(afterFailure.uploadVersionId, generated.uploadVersionId);
  const datasetAfterFailure = await (await fetch(
    `${appBaseUrl}/api/dataset`
  )).json();
  assert.deepEqual(datasetAfterFailure, datasetBeforeDecision);
  const analysisAfterFailure = await (await postJson(
    "/api/analyze-question",
    { question: "What is the total amount?" }
  )).json();
  assert.deepEqual(analysisAfterFailure.result, analysisBeforeDecision.result);
  assert.deepEqual(await fs.readFile(proposalFile), bytes);

  const repeatedResponse = await postJson("/api/cleaning/proposals", {
    uploadVersionId: generated.uploadVersionId
  });
  const repeated = await repeatedResponse.json();
  assert.equal(repeated.reviewId, generated.reviewId);
  assert.equal(
    repeated.proposals.find((proposal) => proposal.proposalId === duplicate.proposalId).decision,
    "approved"
  );

  const conflictingDecision = await postJson(
    `/api/cleaning/proposals/${duplicate.proposalId}/decision`,
    {
      uploadVersionId: generated.uploadVersionId,
      reviewId: generated.reviewId,
      decision: "rejected"
    }
  );
  assert.equal(conflictingDecision.status, 409);

  const staleVersion = await postJson("/api/cleaning/proposals", {
    uploadVersionId: "stale-upload"
  });
  assert.equal(staleVersion.status, 409);

  const staleDecision = await postJson(
    `/api/cleaning/proposals/${duplicate.proposalId}/decision`,
    {
      uploadVersionId: "stale-upload",
      reviewId: generated.reviewId,
      decision: "rejected"
    }
  );
  assert.equal(staleDecision.status, 409);

  const unknownProposal = await postJson(
    "/api/cleaning/proposals/unknown-proposal/decision",
    {
      uploadVersionId: generated.uploadVersionId,
      reviewId: generated.reviewId,
      decision: "approved"
    }
  );
  assert.equal(unknownProposal.status, 404);

  const exportResponse = await postJson("/api/cleaning/export", {
    options: {}
  });
  assert.equal(exportResponse.status, 409);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
});

test("changing a confirmed header invalidates the review and rejects old decisions", async () => {
  const bytes = await createWorkbookBytes("Header Changes", [
    ["Name", "Value"],
    ["Alpha", 1],
    ["Alpha", 1]
  ]);
  const sourceFile = path.join(tempRoot, "header-change.xlsx");
  await fs.writeFile(sourceFile, bytes);
  const upload = await uploadFile(
    sourceFile,
    "header-change.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(upload.status, 200);

  const workbook = await (await fetch(
    `${appBaseUrl}/api/cleaning/workbook`
  )).json();
  const selections = {
    uploadVersionId: workbook.uploadVersionId,
    headerSelections: [{ worksheetKey: "sheet-0", headerRow: 1 }]
  };
  const firstHeaderResponse = await postJson(
    "/api/cleaning/workbook/headers",
    selections
  );
  assert.equal(firstHeaderResponse.status, 200);

  const firstReviewResponse = await postJson("/api/cleaning/proposals", {
    uploadVersionId: workbook.uploadVersionId
  });
  const firstReview = await firstReviewResponse.json();
  const duplicate = firstReview.proposals.find(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.ok(duplicate);

  const repeatedHeaderResponse = await postJson(
    "/api/cleaning/workbook/headers",
    selections
  );
  const repeatedHeaders = await repeatedHeaderResponse.json();
  assert.equal(repeatedHeaders.reviewInvalidated, false);
  const repeatedReview = await (await postJson(
    "/api/cleaning/proposals",
    { uploadVersionId: workbook.uploadVersionId }
  )).json();
  assert.equal(repeatedReview.reviewId, firstReview.reviewId);

  const changedHeaderResponse = await postJson(
    "/api/cleaning/workbook/headers",
    {
      uploadVersionId: workbook.uploadVersionId,
      headerSelections: [{ worksheetKey: "sheet-0", headerRow: 2 }]
    }
  );
  const changedHeaders = await changedHeaderResponse.json();
  assert.equal(changedHeaderResponse.status, 200);
  assert.equal(changedHeaders.reviewInvalidated, true);
  assert.equal(changedHeaders.replacedReviewId, firstReview.reviewId);

  const staleReviewDecision = await postJson(
    `/api/cleaning/proposals/${duplicate.proposalId}/decision`,
    {
      uploadVersionId: workbook.uploadVersionId,
      reviewId: firstReview.reviewId,
      decision: "approved"
    }
  );
  assert.equal(staleReviewDecision.status, 404);

  const activeReview = await (await postJson("/api/cleaning/proposals", {
    uploadVersionId: workbook.uploadVersionId
  })).json();
  assert.notEqual(activeReview.reviewId, firstReview.reviewId);
  const unknownProposalDecision = await postJson(
    "/api/cleaning/proposals/not-in-review/decision",
    {
      uploadVersionId: workbook.uploadVersionId,
      reviewId: activeReview.reviewId,
      decision: "approved"
    }
  );
  assert.equal(unknownProposalDecision.status, 404);
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
});

test("a superseded overlapping upload cannot replace the committed review", async () => {
  const initialBytes = await createWorkbookBytes("Initial", [
    ["Name", "Amount"],
    ["Initial", 1],
    ["Initial", 1]
  ]);
  const initialFile = path.join(tempRoot, "initial-overlap.xlsx");
  await fs.writeFile(initialFile, initialBytes);
  const initialUpload = await uploadFile(
    initialFile,
    "initial-overlap.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(initialUpload.status, 200);

  const olderBytes = await createWorkbookBytes("Older", [
    ["Name", "Amount"],
    ["Older", 10]
  ]);
  const pausedUpload = await startPausedUpload("older-overlap.xlsx", olderBytes);

  const newerBytes = await createWorkbookBytes("Newer", [
    ["Name", "Amount"],
    ["Newer", 90],
    ["Newer", 90]
  ]);
  const newerFile = path.join(tempRoot, "newer-overlap.xlsx");
  await fs.writeFile(newerFile, newerBytes);
  const newerUpload = await uploadFile(
    newerFile,
    "newer-overlap.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  assert.equal(newerUpload.status, 200);

  const newerWorkbook = await (await fetch(
    `${appBaseUrl}/api/cleaning/workbook`
  )).json();
  const newerReviewResponse = await postJson("/api/cleaning/proposals", {
    uploadVersionId: newerWorkbook.uploadVersionId
  });
  const newerReview = await newerReviewResponse.json();
  const duplicate = newerReview.proposals.find(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.ok(duplicate);
  const approval = await postJson(
    `/api/cleaning/proposals/${duplicate.proposalId}/decision`,
    {
      uploadVersionId: newerWorkbook.uploadVersionId,
      reviewId: newerReview.reviewId,
      decision: "approved"
    }
  );
  assert.equal(approval.status, 200);

  pausedUpload.finish();
  const olderResponse = await pausedUpload.response;
  assert.equal(olderResponse.status, 409);
  assert.match(olderResponse.data.error, /superseded/i);

  const committedWorkbook = await (await fetch(
    `${appBaseUrl}/api/cleaning/workbook`
  )).json();
  assert.equal(committedWorkbook.uploadVersionId, newerWorkbook.uploadVersionId);
  const committedDataset = await (await fetch(
    `${appBaseUrl}/api/dataset`
  )).json();
  assert.equal(committedDataset.fileName, "newer-overlap.xlsx");
  assert.equal(committedDataset.preview[0].Name, "Newer");
  const committedReview = await (await postJson(
    "/api/cleaning/proposals",
    { uploadVersionId: newerWorkbook.uploadVersionId }
  )).json();
  assert.equal(committedReview.reviewId, newerReview.reviewId);
  assert.equal(
    committedReview.proposals.find(
      (proposal) => proposal.proposalId === duplicate.proposalId
    ).decision,
    "approved"
  );
  assert.deepEqual(await fs.readdir(uploadDirectory), []);
});
