import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { once } from "node:events";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { after, before, test } from "node:test";
import ExcelJS from "exceljs";
import {
  cleanWorkbook,
  createWorkbookState,
  generateCleaningReview,
  generateCleanedWorkbook,
  recordProposalDecision,
  validateHeaderRow
} from "../workbook-state.mjs";

let tempRoot;
let sourcePath;

function hash(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function createSyntheticWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  const sales = workbook.addWorksheet("Sales & Detail");
  const notes = workbook.addWorksheet("Hidden Notes", {
    state: "hidden"
  });

  sales.getCell("A1").value = "Synthetic Header";
  sales.mergeCells("A1:B1");
  sales.getCell("A1").font = {
    name: "Arial",
    bold: true,
    color: { argb: "FF112233" }
  };
  sales.getCell("A1").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };
  sales.getCell("A1").alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true
  };
  sales.getCell("A1").border = {
    bottom: {
      style: "thin",
      color: { argb: "FF334455" }
    }
  };

  sales.getCell("A3").value = "Synthetic item";
  sales.getCell("B3").value = 12.5;
  sales.getCell("B3").numFmt = "$#,##0.00";
  sales.getCell("C3").value = new Date(Date.UTC(2024, 0, 2, 12));
  sales.getCell("C3").numFmt = "yyyy-mm-dd";
  sales.getCell("E3").value = "After blank column D";
  sales.getCell("A4").value = "Formula row";
  sales.getCell("B4").value = { formula: "B3*2", result: 25 };
  sales.getCell("A5").value = "Hidden synthetic row";
  sales.getCell("A7").value = "After blank row 6";
  sales.getColumn(1).width = 27;
  sales.getColumn(2).width = 16;
  sales.getColumn(3).width = 22;
  sales.getColumn(4).width = 13;
  sales.getColumn(4).hidden = true;
  sales.getColumn(5).width = 31;
  sales.getRow(1).height = 32;
  sales.getRow(5).hidden = true;
  notes.getCell("A1").value = "Synthetic hidden worksheet";
  notes.getCell("C3").value = 99;

  await workbook.xlsx.writeFile(filePath);
}

async function createCleaningWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  const data = workbook.addWorksheet("Proposal Data");
  const hidden = workbook.addWorksheet("Review Notes", { state: "hidden" });
  data.addRow(["Name", "Value", "When"]);
  data.addRow(["Alpha", 1, new Date(Date.UTC(2025, 0, 1))]);
  data.addRow(["Alpha", 1, new Date(Date.UTC(2025, 0, 1))]);
  data.addRow([]);
  data.addRow([" Beta ", 2, null]);
  data.addRow(["Beta", 3, null]);
  hidden.addRow(["Note", "Code"]);
  hidden.addRow(["Synthetic", "X"]);
  await workbook.xlsx.writeFile(filePath);
}

function workbookSnapshot(workbook) {
  return workbook.worksheets.map((worksheet) => {
    const cells = [];
    for (let row = 1; row <= worksheet.rowCount; row++) {
      for (let column = 1; column <= worksheet.columnCount; column++) {
        const cell = worksheet.getCell(row, column);
        cells.push({
          address: cell.address,
          type: cell.type,
          value: cell.value instanceof Date
            ? cell.value.toISOString()
            : cell.value,
          style: JSON.stringify(cell.style),
          numFmt: cell.numFmt,
          isMerged: cell.isMerged
        });
      }
    }
    return {
      name: worksheet.name,
      state: worksheet.state,
      cells
    };
  });
}

before(async () => {
  tempRoot = await fs.mkdtemp(
    path.join(os.tmpdir(), "insightmate-workbook-state-")
  );
  sourcePath = path.join(tempRoot, "synthetic-source.xlsx");
  await createSyntheticWorkbook(sourcePath);
});

after(async () => {
  if (tempRoot) {
    await fs.rm(tempRoot, { recursive: true, force: true });
  }
});

test("workbook state records worksheet order, visibility, and original cell coordinates", async () => {
  const uploadedBytes = await fs.readFile(sourcePath);
  const state = await createWorkbookState(
    uploadedBytes,
    "synthetic-source.xlsx"
  );

  assert.deepEqual(
    state.worksheets.map(({ name, order, visibility }) => ({
      name,
      order,
      visibility
    })),
    [
      { name: "Sales & Detail", order: 0, visibility: "visible" },
      { name: "Hidden Notes", order: 1, visibility: "hidden" }
    ]
  );

  const sales = state.worksheets[0];
  assert.equal(sales.rowCount, 7);
  assert.equal(sales.columnCount, 5);
  assert.deepEqual(
    sales.cells.find((cell) => cell.address === "B4"),
    {
      address: "B4",
      row: 4,
      column: 2,
      type: ExcelJS.ValueType.Formula,
      value: { formula: "B3*2", result: 25 },
      serializedValue: {
        type: "formula",
        formula: "B3*2",
        result: { type: "number", value: 25 }
      },
      numberFormat: null,
      isMerged: false
    }
  );
  assert.deepEqual(
    sales.cells.find((cell) => cell.address === "E3"),
    {
      address: "E3",
      row: 3,
      column: 5,
      type: ExcelJS.ValueType.String,
      value: "After blank column D",
      serializedValue: {
        type: "string",
        value: "After blank column D"
      },
      numberFormat: null,
      isMerged: false
    }
  );
});

test("source bytes stay separate from the preservation workbook and unchanged through round trip", async () => {
  const uploadedBytes = await fs.readFile(sourcePath);
  const originalHash = hash(uploadedBytes);
  const state = await createWorkbookState(
    uploadedBytes,
    "synthetic-source.xlsx"
  );

  assert.notEqual(state.getSourceBytes(), uploadedBytes);
  assert.equal(
    state.preservationWorkbook
      .getWorksheet("Sales & Detail")
      .getCell("A3").value,
    "Synthetic item"
  );
  uploadedBytes[0] ^= 0xff;
  assert.equal(hash(state.getSourceBytes()), originalHash);

  const exposedCopy = state.getSourceBytes();
  exposedCopy[0] ^= 0xff;
  assert.equal(hash(state.getSourceBytes()), originalHash);
  assert.equal(
    state.preservationWorkbook
      .getWorksheet("Sales & Detail")
      .getCell("A3").value,
    "Synthetic item"
  );

  const roundTripPath = path.join(tempRoot, "synthetic-round-trip.xlsx");
  const roundTripBytes = await state.preservationWorkbook.xlsx.writeBuffer();
  await fs.writeFile(roundTripPath, roundTripBytes);

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.readFile(roundTripPath);
  const sales = reopened.getWorksheet("Sales & Detail");
  const notes = reopened.getWorksheet("Hidden Notes");

  assert.deepEqual(
    reopened.worksheets.map((worksheet) => worksheet.name),
    ["Sales & Detail", "Hidden Notes"]
  );
  assert.equal(notes.state, "hidden");
  assert.equal(sales.getCell("A3").value, "Synthetic item");
  assert.equal(sales.getCell("B3").value, 12.5);
  assert.equal(sales.getCell("C3").value.getTime(), Date.UTC(2024, 0, 2, 12));
  assert.equal(sales.getCell("A2").value, null);
  assert.equal(sales.getCell("D3").value, null);
  assert.equal(sales.getCell("A6").value, null);
  assert.equal(sales.getCell("B4").value.formula, "B3*2");
  assert.equal(sales.getCell("B4").value.result, 25);
  assert.equal(sales.getCell("B3").numFmt, "$#,##0.00");
  assert.equal(sales.getCell("C3").numFmt, "yyyy-mm-dd");
  assert.equal(sales.getCell("A1").font.bold, true);
  assert.equal(sales.getCell("A1").fill.fgColor.argb, "FFE0F0FF");
  assert.equal(sales.getCell("A1").alignment.horizontal, "center");
  assert.equal(sales.getCell("A1").border.bottom.style, "thin");
  assert.equal(sales.getCell("B1").isMerged, true);
  assert.equal(sales.getColumn(1).width, 27);
  assert.equal(sales.getColumn(4).hidden, true);
  assert.equal(sales.getRow(1).height, 32);
  assert.equal(sales.getRow(5).hidden, true);
  assert.equal(notes.getCell("C3").value, 99);

  assert.equal(hash(await fs.readFile(sourcePath)), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
});

test("generated workbook formatting survives XLSX serialization across visible and hidden worksheets", async () => {
  const workbook = new ExcelJS.Workbook();
  const metrics = workbook.addWorksheet("Metrics");
  const audit = workbook.addWorksheet("Audit", { state: "hidden" });
  metrics.addRow([
    "Description",
    "Count",
    "Unit Cost",
    "Recorded On",
    "Total",
    "Formula Date"
  ]);
  metrics.addRow([
    "A fairly long product description",
    3,
    12.5,
    new Date(Date.UTC(2025, 5, 7)),
    { formula: "B2*C2", result: 37.5 },
    { formula: "DATE(2025,6,7)", result: 45815 }
  ]);
  metrics.getColumn(1).width = 10;
  metrics.getColumn(2).width = 14;
  metrics.getColumn(4).width = 6;
  metrics.getColumn(6).width = 6;
  metrics.getCell("A1").alignment = { horizontal: "left", vertical: "top" };
  metrics.getCell("A2").alignment = { horizontal: "left", vertical: "top" };
  metrics.getCell("B2").alignment = { horizontal: "right", vertical: "bottom" };
  metrics.getCell("D2").alignment = { horizontal: "left", vertical: "top" };
  metrics.getCell("E2").alignment = { horizontal: "right", vertical: "bottom" };
  metrics.getCell("F2").alignment = { horizontal: "left", vertical: "top" };
  metrics.getCell("C2").numFmt = "$#,##0.00";
  metrics.getCell("D2").numFmt = "dddd, mmmm d, yyyy h:mm AM/PM";
  metrics.getCell("E2").numFmt = "$#,##0.00";
  metrics.getCell("F2").numFmt = "m/d/yyyy h:mm:ss AM/PM";
  audit.addRow(["Review Note", "Status"]);
  audit.addRow(["Checked by the review team", "Complete"]);

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(originalBytes);
  const state = await createWorkbookState(originalBytes, "formatting-source.xlsx");
  const { buffer } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);

  assert.deepEqual(reopened.worksheets.map(({ name, state: visibility }) => ({
    name,
    visibility
  })), [
    { name: "Metrics", visibility: "visible" },
    { name: "Audit", visibility: "hidden" }
  ]);

  const outputMetrics = reopened.getWorksheet("Metrics");
  const outputAudit = reopened.getWorksheet("Audit");
  const header = outputMetrics.getCell("A1");
  assert.equal(header.font.bold, true);
  assert.equal(header.font.color.argb, "FFFFFFFF");
  assert.equal(header.fill.type, "pattern");
  assert.equal(header.fill.fgColor.argb, "FF1F4E78");
  assert.equal(header.alignment.wrapText, true);
  assert.equal(header.alignment.horizontal, "center");
  assert.equal(header.alignment.vertical, "middle");
  assert.ok(outputMetrics.getRow(1).height >= 26);

  for (const address of ["A1", "A2", "B2", "C2", "D2", "E2", "F2"]) {
    assert.equal(outputMetrics.getCell(address).alignment.horizontal, "center");
    assert.equal(outputMetrics.getCell(address).alignment.vertical, "middle");
  }
  assert.equal(outputMetrics.getCell("A2").font.name, outputMetrics.getCell("B2").font.name);
  assert.equal(outputMetrics.getCell("A2").border.bottom.style, "thin");
  assert.equal(outputMetrics.getCell("A2").border.right.style, "thin");
  assert.ok(outputMetrics.getRow(2).height >= 20);
  assert.ok(outputMetrics.getColumn(1).width >= 33);
  assert.ok(outputMetrics.getColumn(1).width <= 48);
  assert.ok(outputMetrics.getColumn(2).width >= 14);
  assert.ok(outputMetrics.getColumn(2).width <= 48);
  assert.ok(outputMetrics.getColumn(4).width >= 40);
  assert.ok(outputMetrics.getColumn(4).width <= 48);
  assert.ok(outputMetrics.getColumn(6).width >= 24);
  assert.ok(outputMetrics.getColumn(6).width <= 48);

  assert.equal(outputMetrics.getCell("A2").value, "A fairly long product description");
  assert.equal(outputMetrics.getCell("B2").value, 3);
  assert.equal(outputMetrics.getCell("D2").value.getTime(), Date.UTC(2025, 5, 7));
  assert.deepEqual(outputMetrics.getCell("E2").value, {
    formula: "B2*C2",
    result: 37.5
  });
  assert.deepEqual(outputMetrics.getCell("F2").value, {
    formula: "DATE(2025,6,7)",
    result: new Date(Date.UTC(2025, 5, 7))
  });
  assert.equal(outputMetrics.getCell("C2").numFmt, "$#,##0.00");
  assert.equal(outputMetrics.getCell("D2").numFmt, "dddd, mmmm d, yyyy h:mm AM/PM");
  assert.equal(outputMetrics.getCell("E2").numFmt, "$#,##0.00");
  assert.equal(outputMetrics.getCell("F2").numFmt, "m/d/yyyy h:mm:ss AM/PM");

  assert.equal(outputAudit.getCell("A1").font.bold, true);
  assert.equal(outputAudit.getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(outputAudit.getCell("A1").alignment.horizontal, "center");
  assert.equal(outputAudit.getCell("A2").alignment.horizontal, "center");
  assert.equal(outputAudit.getCell("A2").alignment.vertical, "middle");
  assert.ok(outputAudit.getColumn(1).width >= 10);
  assert.ok(outputAudit.getColumn(1).width <= 48);
  assert.ok(outputAudit.getRow(1).height >= 26);

  assert.equal(hash(originalBytes), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
});

test("single-sheet generated workbooks receive centered formatting", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("One Sheet");
  sheet.addRows([
    ["Label", "Amount"],
    ["Example", 4]
  ]);
  const sourceBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const sourceHash = hash(sourceBytes);
  const state = await createWorkbookState(sourceBytes, "single-sheet.xlsx");
  const { buffer } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);

  assert.equal(reopened.worksheets.length, 1);
  for (const address of ["A1", "B1", "A2", "B2"]) {
    assert.equal(reopened.worksheets[0].getCell(address).alignment.horizontal, "center");
    assert.equal(reopened.worksheets[0].getCell(address).alignment.vertical, "middle");
  }
  assert.equal(hash(sourceBytes), sourceHash);
  assert.equal(hash(state.getSourceBytes()), sourceHash);
});

test("worksheet identities remain stable and header rows are validated in original coordinates", async () => {
  const bytes = await fs.readFile(sourcePath);
  const state = await createWorkbookState(bytes, "synthetic-source.xlsx");
  assert.deepEqual(
    state.worksheets.map(({ worksheetKey, order }) => ({ worksheetKey, order })),
    [{ worksheetKey: "sheet-0", order: 0 }, { worksheetKey: "sheet-1", order: 1 }]
  );
  assert.equal(validateHeaderRow(state.worksheets[0], 1), true);
  assert.equal(validateHeaderRow(state.worksheets[0], 0), false);
  assert.equal(validateHeaderRow(state.worksheets[0], 8), false);
  assert.equal(validateHeaderRow(state.worksheets[1], 3), true);
});

test("proposal generation matches typed duplicates independently and remains read-only", async () => {
  const proposalPath = path.join(tempRoot, "synthetic-proposals.xlsx");
  await createCleaningWorkbook(proposalPath);
  const sourceBytes = await fs.readFile(proposalPath);
  const sourceHash = hash(sourceBytes);
  const state = await createWorkbookState(sourceBytes, "synthetic-proposals.xlsx");
  const before = workbookSnapshot(state.preservationWorkbook);
  const version = state.uploadVersionId;

  let review = generateCleaningReview(state, version, {});
  const sameReview = generateCleaningReview(state, version, {}, review);
  assert.equal(sameReview.reviewId, review.reviewId);
  assert.deepEqual(sameReview.proposals, review.proposals);

  const duplicateProposals = review.proposals.filter(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.equal(duplicateProposals.length, 1);
  assert.equal(duplicateProposals[0].target.row, 3);
  assert.equal(duplicateProposals[0].matches.retainedRow, 2);
  assert.equal(duplicateProposals[0].decision, "pending");

  assert.equal(
    review.proposals.filter((proposal) => proposal.issueType === "empty_row").length,
    1
  );
  assert.deepEqual(
    review.proposals.find((proposal) => proposal.issueType === "empty_row").target,
    { kind: "row", row: 4 }
  );
  assert.equal(
    review.proposals.filter((proposal) => proposal.issueType === "leading_trailing_whitespace").length,
    1
  );
  const whitespaceTrimProposal = review.proposals.find(
    (proposal) => proposal.issueType === "leading_trailing_whitespace"
  );
  assert.deepEqual(whitespaceTrimProposal.originalValue, {
    type: "string",
    value: " Beta "
  });
  assert.deepEqual(whitespaceTrimProposal.proposedValue, {
    type: "string",
    value: "Beta"
  });
  assert.ok(review.proposals.every((proposal) => proposal.decision === "pending"));
  assert.ok(review.findings.some((finding) => finding.issueType === "missing_value"));

  const decision = recordProposalDecision(
    review,
    duplicateProposals[0].proposalId,
    "approved"
  );
  assert.equal(decision.status, "updated");
  review = generateCleaningReview(state, version, {}, review);
  assert.equal(review.proposals.find(
    (proposal) => proposal.proposalId === duplicateProposals[0].proposalId
  ).decision, "approved");
  assert.equal(recordProposalDecision(
    review,
    duplicateProposals[0].proposalId,
    "rejected"
  ).status, "already_decided");
  const whitespaceProposal = review.proposals.find(
    (proposal) => proposal.issueType === "leading_trailing_whitespace"
  );
  assert.equal(recordProposalDecision(
    review,
    whitespaceProposal.proposalId,
    "rejected"
  ).status, "updated");

  assert.deepEqual(workbookSnapshot(state.preservationWorkbook), before);
  assert.equal(hash(await fs.readFile(proposalPath)), sourceHash);
  assert.equal(hash(state.getSourceBytes()), sourceHash);
});

test("ambiguous date text and missing values are not inferred or changed by review generation", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Ambiguous Values");
  sheet.addRow(["Date Text", "Amount", "Notes"]);
  sheet.addRow(["03/04/2025", null, "unchanged"]);

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const sourceHash = hash(bytes);
  const state = await createWorkbookState(bytes, "ambiguous-values.xlsx");
  const before = workbookSnapshot(state.preservationWorkbook);
  const review = generateCleaningReview(state, state.uploadVersionId, {
    "sheet-0": { action: "confirm", row: 1 }
  });

  assert.equal(
    state.worksheets[0].cells.find((cell) => cell.address === "A2").type,
    ExcelJS.ValueType.String
  );
  assert.equal(
    state.worksheets[0].cells.find((cell) => cell.address === "A2").value,
    "03/04/2025"
  );
  assert.equal(
    state.worksheets[0].cells.find((cell) => cell.address === "B2"),
    undefined
  );
  assert.ok(review.findings.some((finding) =>
    finding.issueType === "missing_value" &&
    finding.target.address === "B2"
  ));
  assert.ok(review.proposals.every((proposal) =>
    !["date_standardization", "missing_value"].includes(proposal.issueType)
  ));
  assert.deepEqual(workbookSnapshot(state.preservationWorkbook), before);
  assert.equal(hash(bytes), sourceHash);
  assert.equal(hash(state.getSourceBytes()), sourceHash);
});

test("safe workbook cleaning formats every sheet and preserves the source workbook", async () => {
  const workbook = new ExcelJS.Workbook();
  const sales = workbook.addWorksheet("Sales & Detail");
  const notes = workbook.addWorksheet("Hidden Notes", { state: "hidden" });
  sales.addRow(["Customer", "Amount", "Date Text", "Notes"]);
  sales.addRow(["  North   Sales  ", 10, "03/04/2025", "first"]);
  sales.addRow(["  North   Sales  ", 10, "03/04/2025", "first"]);
  sales.addRow(["South", null, "13/04/2025", ""]);
  sales.addRow([]);
  sales.addRow(["Tail", 7, "04/13/2025", "last"]);
  sales.getColumn(1).width = 27;
  sales.getRow(1).height = 30;
  sales.getCell("A1").font = { name: "Calibri", color: { argb: "FF112233" } };
  sales.getCell("A1").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };
  sales.getCell("B2").numFmt = "$#,##0.00";
  notes.addRow(["Note", "Value"]);
  notes.addRow(["  keep  internal spaces  ", 42]);

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(originalBytes);
  const sourceState = await createWorkbookState(originalBytes, "safe-cleaning.xlsx");
  const preservationSnapshot = workbookSnapshot(sourceState.preservationWorkbook);
  const { workbook: cleanedWorkbook, summary } = await cleanWorkbook(sourceState);
  const cleanedSales = cleanedWorkbook.getWorksheet("Sales & Detail");
  const cleanedNotes = cleanedWorkbook.getWorksheet("Hidden Notes");
  const salesSummary = summary.worksheets[0];
  const notesSummary = summary.worksheets[1];

  assert.equal(summary.worksheetCount, 2);
  assert.deepEqual(
    cleanedWorkbook.worksheets.map((worksheet) => ({
      name: worksheet.name,
      state: worksheet.state
    })),
    [
      { name: "Sales & Detail", state: "visible" },
      { name: "Hidden Notes", state: "hidden" }
    ]
  );
  assert.equal(salesSummary.spacesTrimmed, 2);
  assert.equal(salesSummary.emptyRowsRemoved, 0);
  assert.equal(salesSummary.blankRowsDetected, 1);
  assert.equal(salesSummary.blankRowsRemoved, 0);
  assert.equal(salesSummary.blankRowsRetained, 1);
  assert.match(salesSummary.emptyRowsSkipped[0].reason, /too few populated fields/);
  assert.equal(salesSummary.duplicateCandidatesFound, 1);
  assert.equal(salesSummary.duplicateCandidates[0].originalRow, 3);
  assert.equal(salesSummary.duplicateCandidates[0].retainedOriginalRow, 2);
  assert.equal(salesSummary.duplicateCandidates[0].cleanedRow, 3);
  assert.match(salesSummary.duplicateCandidates[0].candidateId, /^[a-f0-9]{64}$/);
  assert.equal(salesSummary.duplicateCandidates[0].safeToRemove, true);
  assert.equal(salesSummary.duplicateCandidates[0].unsafeReason, null);
  assert.equal(salesSummary.blankCellsDetected, 2);
  assert.equal(salesSummary.ambiguousItemsLeftUnchanged, 2);
  assert.equal(cleanedSales.getCell("A2").value, "North Sales");
  assert.equal(cleanedSales.getCell("A3").value, "North Sales");
  assert.equal(cleanedSales.getCell("C2").value, "03/04/2025");
  assert.equal(cleanedSales.getCell("C3").value, "03/04/2025");
  assert.equal(cleanedSales.getCell("A4").value, "South");
  assert.equal(cleanedSales.getCell("A5").value, null);
  assert.equal(cleanedSales.getCell("A6").value, "Tail");
  assert.equal(cleanedSales.getRow(7).hasValues, false);
  assert.equal(cleanedNotes.getCell("A2").value, "keep internal spaces");
  assert.ok(notesSummary.spacesTrimmed > 0);
  assert.equal(cleanedSales.getCell("B2").numFmt, "$#,##0.00");
  assert.equal(cleanedSales.getColumn(1).width, 27);
  assert.ok(cleanedSales.getColumn(2).width >= 10);
  assert.ok(cleanedSales.getColumn(2).width <= 48);
  assert.equal(cleanedSales.getRow(1).height, 30);
  assert.equal(cleanedSales.getCell("A1").font.name, "Calibri");
  assert.equal(cleanedSales.getCell("A1").font.bold, true);
  assert.equal(cleanedSales.getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(cleanedSales.getCell("B2").alignment.horizontal, "center");
  assert.equal(cleanedSales.getCell("A2").alignment.horizontal, "center");
  assert.equal(cleanedSales.getCell("B2").alignment.vertical, "middle");
  assert.equal(cleanedSales.getCell("A2").alignment.vertical, "middle");
  assert.ok(cleanedSales.getCell("A2").border.bottom);
  assert.ok(salesSummary.rowsFormatted > 0);
  assert.ok(notesSummary.rowsFormatted > 0);
  assert.equal(hash(originalBytes), originalHash);
  assert.equal(hash(sourceState.getSourceBytes()), originalHash);
  assert.deepEqual(
    workbookSnapshot(sourceState.preservationWorkbook),
    preservationSnapshot
  );
});

test("formula and merged-cell worksheets retain unsafe empty rows and formula definitions", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Structured");
  sheet.addRow(["Key", "Value"]);
  sheet.addRow(["First", 1]);
  sheet.addRow([]);
  sheet.addRow(["Formula", { formula: "B2*2", result: 2 }]);
  sheet.addRow(["Merged", 3]);
  sheet.mergeCells("C5:D5");
  sheet.getColumn(1).width = 24;
  sheet.getRow(1).height = 28;
  sheet.getCell("B2").numFmt = "$#,##0.00";
  sheet.getCell("A1").font = {
    name: "Arial",
    bold: true,
    color: { argb: "FF112233" }
  };
  sheet.getCell("A1").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "structured-cleaning.xlsx");
  const { workbook: cleanedWorkbook, summary } = await cleanWorkbook(state);
  const cleanedSheet = cleanedWorkbook.getWorksheet("Structured");

  assert.equal(cleanedSheet.getCell("A3").value, null);
  assert.deepEqual(cleanedSheet.getCell("B4").value, {
    formula: "B2*2",
    result: 2
  });
  assert.equal(cleanedSheet.getCell("C5").isMerged, true);
  assert.equal(cleanedSheet.getCell("B2").numFmt, "$#,##0.00");
  assert.equal(cleanedSheet.getColumn(1).width, 24);
  assert.equal(cleanedSheet.getRow(1).height, 28);
  assert.equal(cleanedSheet.getCell("A1").font.bold, true);
  assert.equal(cleanedSheet.getCell("A1").font.color.argb, "FFFFFFFF");
  assert.equal(cleanedSheet.getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(summary.worksheets[0].emptyRowsRemoved, 0);
  assert.deepEqual(summary.worksheets[0].emptyRowsSkipped, [
    {
      originalRow: 3,
      reason: "A formula contains references or syntax whose row dependencies cannot be safely determined."
    }
  ]);
  assert.equal(summary.worksheets[0].blankRowsDetected, 1);
  assert.equal(summary.worksheets[0].blankRowsRetained, 1);
  assert.ok(summary.worksheets[0].itemsSkippedForSafety >= 2);
});

test("safe trailing blank rows are removed and ambiguous or structurally dependent rows are retained", async () => {
  const workbook = new ExcelJS.Workbook();
  const safeTail = workbook.addWorksheet("Safe Tail");
  safeTail.addRows([
    ["Name", "Value"],
    ["Alpha", 1],
    ["Independent formula", { formula: "1+1", result: 2 }]
  ]);
  for (let rowNumber = 4; rowNumber <= 5; rowNumber++) {
    for (let columnNumber = 1; columnNumber <= 2; columnNumber++) {
      const cell = safeTail.getCell(rowNumber, columnNumber);
      cell.font = { bold: true };
      cell.border = { bottom: { style: "thin" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE0F0FF" }
      };
    }
  }
  safeTail.addRow([]);

  const separators = workbook.addWorksheet("Separators");
  separators.addRows([
    ["Name", "Value", "Region"],
    ["Before", 1, "North"],
    [],
    ["West Region"],
    [],
    ["After", 2, "West"],
    ["   ", null, null]
  ]);
  for (const rowNumber of [3, 5]) {
    separators.getCell(`A${rowNumber}`).font = { bold: true };
    separators.getCell(`A${rowNumber}`).border = { bottom: { style: "thin" } };
    separators.getCell(`A${rowNumber}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE0F0FF" }
    };
  }

  const ambiguous = workbook.addWorksheet("Ambiguous Neighbors");
  ambiguous.addRows([
    ["Name", "Value", "Region", "Status"],
    ["Before", 1, null, null],
    [],
    ["After", 2, null, null]
  ]);
  ambiguous.getCell("A3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };

  const repeatedHeader = workbook.addWorksheet("Repeated Header");
  repeatedHeader.addRows([
    ["Name", "Value", "Region"],
    ["Before", 1, "North"],
    [],
    ["Name", "Value", "Region"],
    [],
    ["After", 2, "South"]
  ]);
  for (const rowNumber of [3, 5]) {
    repeatedHeader.getCell(`A${rowNumber}`).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE0F0FF" }
    };
  }

  const sameSheet = workbook.addWorksheet("Same Sheet References");
  sameSheet.addRows([
    ["Name", "Value"],
    ["Data", 1],
    ["Dependent formula", { formula: "B2*2", result: 2 }],
    []
  ]);
  sameSheet.getCell("A4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };

  const crossSheet = workbook.addWorksheet("Cross Sheet Target");
  crossSheet.addRows([
    ["Name"],
    ["Data"],
    []
  ]);
  crossSheet.getCell("A3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };
  workbook.addWorksheet("Audit").getCell("A1").value = {
    formula: "'Cross Sheet Target'!A2",
    result: "Data"
  };

  const tableSheet = workbook.addWorksheet("Table");
  tableSheet.addTable({
    name: "RecordsTable",
    ref: "A1",
    headerRow: true,
    columns: [{ name: "Name" }],
    rows: [["Data"]]
  });
  tableSheet.getRow(3).getCell(1).font = { italic: true };

  const merged = workbook.addWorksheet("Merged");
  merged.addRows([["Name"], ["Data"]]);
  merged.mergeCells("B3:C3");

  const rowMetadata = workbook.addWorksheet("Row Metadata");
  rowMetadata.addRows([["Name"], ["Data"], []]);
  rowMetadata.getRow(3).height = 24;

  const sharedFormula = workbook.addWorksheet("Shared Formula");
  sharedFormula.addRows([
    ["Name", "Calculation"],
    ["Alpha", null],
    ["Beta", null],
    []
  ]);
  sharedFormula.getCell("B2").value = {
    formula: "1+1",
    result: 2,
    shareType: "shared",
    ref: "B2:B4"
  };
  sharedFormula.getCell("B3").value = {
    sharedFormula: "B2",
    result: 2
  };
  sharedFormula.getCell("A4").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };

  const autoFilterSheet = workbook.addWorksheet("Auto Filter");
  autoFilterSheet.addRows([
    ["Name", "Value", "Region"],
    ["Before", 1, "North"],
    [],
    ["After", 2, "South"]
  ]);
  autoFilterSheet.autoFilter = "A1:C4";
  autoFilterSheet.getCell("A3").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "blank-row-safety.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const outputPath = path.join(tempRoot, "blank-row-safety-cleaned.xlsx");
  await fs.writeFile(outputPath, buffer);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.readFile(outputPath);

  const safeSummary = summary.worksheets[0];
  const safeOutput = reopened.getWorksheet("Safe Tail");
  assert.equal(safeSummary.blankRowsDetected, 2);
  assert.equal(safeSummary.blankRowsRemoved, 2);
  assert.equal(safeSummary.blankRowsRetained, 0);
  assert.equal(safeOutput.rowCount, 3);
  assert.equal(safeOutput.getCell("A2").value, "Alpha");
  assert.deepEqual(safeOutput.getCell("B3").value, {
    formula: "1+1",
    result: 2
  });

  const separatorSummary = summary.worksheets[1];
  const separatorOutput = reopened.getWorksheet("Separators");
  assert.equal(separatorSummary.blankRowsDetected, 3);
  assert.equal(separatorSummary.blankRowsRemoved, 0);
  assert.equal(separatorSummary.blankRowsRetained, 3);
  assert.equal(separatorOutput.getCell("A2").value, "Before");
  assert.equal(separatorOutput.getRow(3).hasValues, false);
  assert.equal(separatorOutput.getCell("A4").value, "West Region");
  assert.equal(separatorOutput.getRow(5).hasValues, false);
  assert.equal(separatorOutput.getCell("A6").value, "After");
  assert.equal(separatorOutput.getCell("A7").value, "");
  assert.equal(separatorOutput.getCell("A3").font.bold, true);
  assert.equal(separatorOutput.getCell("A5").border.bottom.style, "thin");
  assert.deepEqual(separatorSummary.blankRowsRetainedByReason, [
    {
      reason: "A blank row was retained because a neighboring row has too few populated fields to confirm an ordinary data record.",
      count: 2
    },
    {
      reason: "The row contained whitespace-only values before text cleanup.",
      count: 1
    }
  ]);

  const ambiguousSummary = summary.worksheets[2];
  assert.equal(ambiguousSummary.blankRowsDetected, 1);
  assert.equal(ambiguousSummary.blankRowsRemoved, 0);
  assert.equal(ambiguousSummary.blankRowsRetained, 1);
  assert.match(
    ambiguousSummary.blankRowsRetainedDetails[0].reason,
    /too few populated fields/
  );
  assert.equal(reopened.getWorksheet("Ambiguous Neighbors").getRow(3).hasValues, false);

  const repeatedHeaderSummary = summary.worksheets[3];
  assert.equal(repeatedHeaderSummary.blankRowsDetected, 2);
  assert.equal(repeatedHeaderSummary.blankRowsRemoved, 0);
  assert.equal(repeatedHeaderSummary.blankRowsRetained, 2);
  assert.ok(repeatedHeaderSummary.blankRowsRetainedDetails.every((item) =>
    /resembles a repeated header/.test(item.reason)
  ));
  assert.deepEqual(
    reopened.getWorksheet("Repeated Header").getRow(4).values.slice(1),
    ["Name", "Value", "Region"]
  );

  const sameSheetSummary = summary.worksheets[4];
  assert.equal(sameSheetSummary.blankRowsDetected, 1);
  assert.equal(sameSheetSummary.blankRowsRemoved, 0);
  assert.match(
    sameSheetSummary.blankRowsRetainedDetails[0].reason,
    /formula contains references/
  );
  assert.deepEqual(
    reopened.getWorksheet("Same Sheet References").getCell("B3").value,
    { formula: "B2*2", result: 2 }
  );

  const crossSheetSummary = summary.worksheets[5];
  assert.equal(crossSheetSummary.blankRowsDetected, 1);
  assert.equal(crossSheetSummary.blankRowsRemoved, 0);
  assert.match(
    crossSheetSummary.blankRowsRetainedDetails[0].reason,
    /another worksheet may depend/
  );
  assert.equal(reopened.getWorksheet("Cross Sheet Target").getCell("A2").value, "Data");

  assert.match(
    summary.worksheets[7].blankRowsRetainedDetails[0].reason,
    /contains a table/
  );
  assert.equal(reopened.getWorksheet("Table").rowCount, 3);
  assert.equal(reopened.getWorksheet("Table").getCell("A3").font.italic, true);
  assert.match(
    summary.worksheets[8].blankRowsRetainedDetails[0].reason,
    /merged cells/
  );
  assert.equal(reopened.getWorksheet("Merged").getCell("B3").isMerged, true);
  assert.match(
    summary.worksheets[9].blankRowsRetainedDetails[0].reason,
    /custom row-level metadata/
  );
  assert.equal(reopened.getWorksheet("Row Metadata").getRow(3).height, 24);
  assert.equal(summary.worksheets[10].blankRowsDetected, 1);
  assert.equal(summary.worksheets[10].blankRowsRemoved, 0);
  assert.match(
    summary.worksheets[10].blankRowsRetainedDetails[0].reason,
    /shared formula range/
  );
  assert.deepEqual(reopened.getWorksheet("Shared Formula").getCell("B2").value, {
    formula: "1+1",
    result: 2,
    ref: "B2:B4",
    shareType: "shared"
  });
  assert.deepEqual(reopened.getWorksheet("Shared Formula").getCell("B3").value, {
    result: 2,
    sharedFormula: "B2"
  });
  assert.match(
    summary.worksheets[11].blankRowsRetainedDetails[0].reason,
    /filters or data validation/
  );
  assert.equal(reopened.getWorksheet("Auto Filter").getRow(3).hasValues, false);

  const definedNameWorkbook = new ExcelJS.Workbook();
  const definedNameSheet = definedNameWorkbook.addWorksheet("Defined Name");
  definedNameSheet.addRows([
    ["Name", "Value", "Region"],
    ["Before", 1, "North"],
    [],
    ["After", 2, "South"]
  ]);
  definedNameSheet.getCell("A3").name = "NamedBlankRow";
  const definedNameState = await createWorkbookState(
    Buffer.from(await definedNameWorkbook.xlsx.writeBuffer()),
    "defined-name-blank-row.xlsx"
  );
  const definedNameResult = await generateCleanedWorkbook(definedNameState);
  const definedNameOutput = new ExcelJS.Workbook();
  await definedNameOutput.xlsx.load(definedNameResult.buffer);
  assert.equal(definedNameResult.summary.worksheets[0].blankRowsRemoved, 0);
  assert.match(
    definedNameResult.summary.worksheets[0].blankRowsRetainedDetails[0].reason,
    /defined names/
  );
  assert.equal(definedNameOutput.getWorksheet("Defined Name").getRow(3).hasValues, false);
});

test("literal Excel error placeholders are blanked while styled rows and formula duplicates survive serialization", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Error Placeholders");
  sheet.addRow(["Kind", "Value"]);
  const placeholders = [
    "#N/A",
    "#REF!",
    "#NAME?",
    "#DIV/0!",
    "#NULL!",
    "#VALUE!",
    "#NUM!",
    "#ERROR!"
  ];
  for (const placeholder of placeholders) {
    sheet.addRow(["Literal placeholder", placeholder]);
  }
  sheet.addRow(["Ordinary prose", "The text #ERROR! is part of this sentence."]);
  sheet.addRow(["Typed Excel error", { error: "#REF!" }]);
  sheet.addRow(["Cached formula error", { formula: "1/0", result: { error: "#DIV/0!" } }]);
  sheet.addRow(["Formula duplicate", { formula: "1+1", result: 2 }]);
  sheet.addRow(["Formula duplicate", { formula: "1+1", result: 2 }]);

  const firstStyledRow = sheet.rowCount + 1;
  const styledBlankRowCount = 59;
  for (let rowNumber = firstStyledRow; rowNumber < firstStyledRow + styledBlankRowCount; rowNumber++) {
    for (let columnNumber = 1; columnNumber <= 3; columnNumber++) {
      const cell = sheet.getCell(rowNumber, columnNumber);
      cell.font = { name: "Arial", bold: true, color: { argb: "FF112233" } };
      cell.border = {
        bottom: { style: "thin", color: { argb: "FF334455" } }
      };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFE0F0FF" }
      };
    }
  }
  sheet.addRow(["#ERROR!"]);

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(originalBytes);
  const state = await createWorkbookState(originalBytes, "literal-errors.xlsx");
  const preservationSnapshot = workbookSnapshot(state.preservationWorkbook);
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const outputPath = path.join(tempRoot, "literal-errors-cleaned.xlsx");
  await fs.writeFile(outputPath, buffer);

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.readFile(outputPath);
  const output = reopened.getWorksheet("Error Placeholders");
  const worksheetSummary = summary.worksheets[0];

  for (let index = 0; index < placeholders.length; index++) {
    assert.equal(output.getCell(index + 2, 2).value, null);
  }
  assert.equal(
    output.getCell(placeholders.length + 2, 2).value,
    "The text #ERROR! is part of this sentence."
  );
  assert.equal(output.getCell(placeholders.length + 3, 2).value, null);
  assert.equal(output.getCell(placeholders.length + 4, 2).value, null);
  assert.deepEqual(
    output.getCell(placeholders.length + 5, 2).value,
    { formula: "1+1", result: 2 }
  );
  assert.deepEqual(
    output.getCell(placeholders.length + 6, 2).value,
    { formula: "1+1", result: 2 }
  );
  assert.equal(worksheetSummary.errorReplacements.length, placeholders.length + 3);
  assert.ok(worksheetSummary.errorReplacements.some((replacement) =>
    replacement.originalError === "#ERROR!" &&
    replacement.address === "B9"
  ));
  assert.ok(worksheetSummary.errorReplacements.some((replacement) =>
    replacement.originalFormula === "1/0" &&
    replacement.action === "Replaced cached formula error with a blank cell"
  ));
  assert.ok(worksheetSummary.errorReplacements.some((replacement) =>
    replacement.address === "A74" &&
    replacement.originalError === "#ERROR!"
  ));
  assert.equal(worksheetSummary.duplicateCandidatesFound, 1);
  assert.equal(worksheetSummary.removedRows, 0);
  assert.equal(worksheetSummary.duplicateRowsSkipped.length, 1);
  assert.equal(
    worksheetSummary.duplicateRowsSkipped[0].reason,
    "The duplicate row contains a formula."
  );
  assert.equal(output.rowCount, firstStyledRow);
  assert.equal(output.getCell(firstStyledRow, 1).value, null);
  assert.equal(worksheetSummary.blankRowsDetected, styledBlankRowCount + 1);
  assert.equal(worksheetSummary.blankRowsRemoved, styledBlankRowCount);
  assert.equal(worksheetSummary.blankRowsRetained, 1);
  assert.deepEqual(worksheetSummary.blankRowsRetainedDetails, [
    {
      originalRow: firstStyledRow + styledBlankRowCount,
      reason: "The row contained recognized Excel errors that were blanked."
    }
  ]);
  assert.equal(worksheetSummary.styledEmptyRowsRetained.length, 0);
  assert.equal(hash(await fs.readFile(outputPath)), hash(buffer));
  assert.equal(hash(originalBytes), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
  assert.deepEqual(workbookSnapshot(state.preservationWorkbook), preservationSnapshot);
});

test("cleaning engine rejects missing and invalid original workbook bytes", async () => {
  await assert.rejects(
    cleanWorkbook(null),
    { name: "TypeError" }
  );
  await assert.rejects(
    cleanWorkbook({ getSourceBytes: () => Buffer.alloc(0) }),
    { name: "TypeError" }
  );
  await assert.rejects(
    cleanWorkbook({ getSourceBytes: () => Buffer.from("not an XLSX file") })
  );
});

test("exact duplicate comparison preserves primitive types and normalizes outer whitespace", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Typed Values");
  sheet.addRow(["Name", "Value"]);
  sheet.addRow([" A", 12]);
  sheet.addRow(["A", "12"]);
  sheet.addRow([" A", 12]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "typed-values.xlsx");
  const review = generateCleaningReview(state, state.uploadVersionId, {
    "sheet-0": { action: "confirm", row: 1 }
  });
  const duplicates = review.proposals.filter(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.equal(duplicates.length, 1);
  assert.equal(duplicates[0].target.row, 4);
  assert.equal(duplicates[0].matches.retainedRow, 2);
});

test("later exact duplicates each retain the earliest eligible data row", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Repeated Rows");
  sheet.addRow(["Same", "1"]);
  sheet.addRow(["Same", "1"]);
  sheet.addRow(["Same", "1"]);
  sheet.addRow(["Same", "1"]);
  sheet.addRow(["Same", "1"]);

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "repeated-rows.xlsx");
  const review = generateCleaningReview(state, state.uploadVersionId, {});
  const duplicates = review.proposals.filter(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );

  assert.deepEqual(
    duplicates.map(({ target, matches }) => ({
      row: target.row,
      retainedRow: matches.retainedRow
    })),
    [
      { row: 3, retainedRow: 2 },
      { row: 4, retainedRow: 2 },
      { row: 5, retainedRow: 2 }
    ]
  );
});

test("formula and merged-cell rows are not duplicate keepers", async () => {
  const workbook = new ExcelJS.Workbook();
  const formulaSheet = workbook.addWorksheet("Formula Rows");
  formulaSheet.addRow(["Name", "Value"]);
  formulaSheet.addRow(["Keeper", 0]);
  formulaSheet.addRow(["Formula", { formula: "1+1", result: 2 }]);
  formulaSheet.addRow(["Formula", { formula: "1+1", result: 2 }]);
  formulaSheet.addRow(["Formula", 2]);
  formulaSheet.addRow(["Formula", 2]);

  const mergedSheet = workbook.addWorksheet("Merged Rows");
  mergedSheet.addRow(["Name", "Value", "Extra"]);
  mergedSheet.addRow(["Keeper", 0, null]);
  mergedSheet.addRow(["Merged", 2, null]);
  mergedSheet.mergeCells("B3:C3");
  mergedSheet.addRow(["Merged", 2, null]);
  mergedSheet.addRow(["Merged", 2, null]);

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "ineligible-rows.xlsx");
  const review = generateCleaningReview(state, state.uploadVersionId, {});
  const duplicates = review.proposals.filter(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );

  assert.deepEqual(
    duplicates.map(({ worksheet, target, matches }) => ({
      worksheet: worksheet.name,
      row: target.row,
      retainedRow: matches.retainedRow
    })),
    [
      { worksheet: "Formula Rows", row: 6, retainedRow: 5 },
      { worksheet: "Merged Rows", row: 5, retainedRow: 4 }
    ]
  );
});

test("Excel error cells remain distinct from identical literal text", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Error Values");
  sheet.addRow(["Name", "Value"]);
  sheet.addRow(["Alpha", { error: "#N/A" }]);
  sheet.addRow(["Alpha", "#N/A"]);

  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "error-values.xlsx");
  const errorCell = state.worksheets[0].cells.find(
    (cell) => cell.address === "B2"
  );

  assert.equal(errorCell.type, ExcelJS.ValueType.Error);
  assert.deepEqual(errorCell.serializedValue, {
    type: "error",
    value: "#N/A"
  });

  const review = generateCleaningReview(state, state.uploadVersionId, {});
  assert.equal(
    review.proposals.filter(
      (proposal) => proposal.issueType === "exact_duplicate_row"
    ).length,
    0
  );
});

test("generated workbook conservatively tidies names, normalizes spaces, and blanks recognized Excel errors", async () => {
  const workbook = new ExcelJS.Workbook();
  const people = workbook.addWorksheet("People");
  const hidden = workbook.addWorksheet("Hidden Errors", { state: "hidden" });
  people.addRow([
    "First Name",
    "Person Name",
    "Notes",
    "Email",
    "Website",
    "Code",
    "Amount",
    "Unaffected Formula",
    "Cached Error",
    "No Cached Error",
    "Literal Error Text",
    "Standalone Error",
    "Formula Text Result",
    "Standalone #ERROR!",
    "Cached #ERROR!",
    "Literal #ERROR!"
  ]);
  people.addRow([
    "BrENdAN",
    "Bill SmITH",
    "David     Rasmussen",
    "person@example.com",
    "https://example.com/a  b",
    "AB  123",
    12.5,
    { formula: "G2*2", result: 25 },
    { formula: "1/0", result: { error: "#N/A" } },
    { formula: "1/0" },
    "Literal #DIV/0! remains text",
    { error: "#DIV/0!" },
    { formula: '"#DIV/0!"', result: "#DIV/0!" },
    { error: "#ERROR!" },
    { formula: "NA()", result: { error: "#ERROR!" } },
    "#ERROR!"
  ]);
  people.addRow([
    "BRENDAN",
    "BRENDAN Wallace",
    "\n  Keep     these spaces\nand\t    this tab",
    "ALREADY@EXAMPLE.COM",
    "www.example.com/a    b",
    "SKU-AB  123",
    7,
    { formula: "G3*3", result: 21 },
    null,
    null,
    "#VALUE! as literal text",
    null
  ]);
  people.addRow(["NASA", "NASA Smith", "Dr.  Who", "A.B+tag@example.com", "https://x.test", "AB-CD", 4]);
  people.addRow([null, "McDONALD Smith"]);
  people.addRow([null, "O'NEIL Smith"]);
  people.addRow([null, "Anne-Marie Smith"]);
  people.addRow([null, "Bill Smith Jr"]);
  people.getCell("G2").numFmt = "$#,##0.00";
  people.getCell("G3").numFmt = "$#,##0.00";
  people.getCell("I2").numFmt = "0.000";

  hidden.addRow(["Standalone Error", "Cached Formula Error", "Ordinary Text"]);
  hidden.addRow([
    { error: "#REF!" },
    { formula: "NA()", result: { error: "#VALUE!" } },
    "Do not change #DIV/0! in this text"
  ]);

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(originalBytes);
  const state = await createWorkbookState(originalBytes, "conservative-cleaning.xlsx");
  const preservationBefore = workbookSnapshot(state.preservationWorkbook);
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);

  const outputPeople = reopened.getWorksheet("People");
  const outputHidden = reopened.getWorksheet("Hidden Errors");
  assert.deepEqual(reopened.worksheets.map(({ name, state: visibility }) => ({
    name,
    visibility
  })), [
    { name: "People", visibility: "visible" },
    { name: "Hidden Errors", visibility: "hidden" }
  ]);

  assert.equal(outputPeople.getCell("A2").value, "Brendan");
  assert.equal(outputPeople.getCell("B2").value, "Bill Smith");
  assert.equal(outputPeople.getCell("A3").value, "BRENDAN");
  assert.equal(outputPeople.getCell("B3").value, "Brendan Wallace");
  assert.equal(outputPeople.getCell("A4").value, "NASA");
  assert.equal(outputPeople.getCell("B4").value, "NASA Smith");
  assert.equal(outputPeople.getCell("B5").value, "McDONALD Smith");
  assert.equal(outputPeople.getCell("B6").value, "O'NEIL Smith");
  assert.equal(outputPeople.getCell("B7").value, "Anne-Marie Smith");
  assert.equal(outputPeople.getCell("B8").value, "Bill Smith Jr");
  assert.equal(outputPeople.getCell("C2").value, "David Rasmussen");
  assert.equal(outputPeople.getCell("C3").value, "\n Keep these spaces\nand\t this tab");
  assert.equal(outputPeople.getCell("C4").value, "Dr.  Who");
  assert.equal(outputPeople.getCell("D2").value, "person@example.com");
  assert.equal(outputPeople.getCell("D3").value, "ALREADY@EXAMPLE.COM");
  assert.equal(outputPeople.getCell("D4").value, "A.B+tag@example.com");
  assert.equal(outputPeople.getCell("E2").value, "https://example.com/a  b");
  assert.equal(outputPeople.getCell("E3").value, "www.example.com/a    b");
  assert.equal(outputPeople.getCell("F2").value, "AB  123");
  assert.equal(outputPeople.getCell("F3").value, "SKU-AB  123");
  assert.equal(outputPeople.getCell("F4").value, "AB-CD");

  assert.deepEqual(outputPeople.getCell("H2").value, {
    formula: "G2*2",
    result: 25
  });
  assert.deepEqual(outputPeople.getCell("H3").value, {
    formula: "G3*3",
    result: 21
  });
  assert.equal(outputPeople.getCell("I2").value, null);
  assert.deepEqual(outputPeople.getCell("J2").value, { formula: "1/0" });
  assert.equal(outputPeople.getCell("K2").value, "Literal #DIV/0! remains text");
  assert.equal(outputPeople.getCell("K3").value, "#VALUE! as literal text");
  assert.deepEqual(outputPeople.getCell("M2").value, {
    formula: '"#DIV/0!"',
    result: "#DIV/0!"
  });
  assert.equal(outputPeople.getCell("N2").value, null);
  assert.equal(outputPeople.getCell("O2").value, null);
  assert.equal(outputPeople.getCell("P2").value, null);
  assert.equal(outputPeople.getCell("I2").numFmt, "0.000");
  assert.equal(outputPeople.getCell("L2").value, null);
  assert.equal(outputPeople.getCell("G2").numFmt, "$#,##0.00");
  assert.equal(outputPeople.getCell("G2").value, 12.5);
  assert.equal(outputPeople.getCell("H2").alignment.horizontal, "center");
  assert.equal(outputPeople.getCell("H2").alignment.vertical, "middle");

  assert.equal(outputHidden.getCell("A2").value, null);
  assert.equal(outputHidden.getCell("B2").value, null);
  assert.equal(outputHidden.getCell("C2").value, "Do not change #DIV/0! in this text");
  assert.deepEqual(summary.worksheets[0].errorReplacements, [
    {
      worksheetName: "People",
      address: "I2",
      originalError: "#N/A",
      originalFormula: "1/0",
      action: "Replaced cached formula error with a blank cell"
    },
    {
      worksheetName: "People",
      address: "L2",
      originalError: "#DIV/0!",
      originalFormula: null,
      action: "Replaced Excel error with a blank cell"
    },
    {
      worksheetName: "People",
      address: "N2",
      originalError: "#ERROR!",
      originalFormula: null,
      action: "Replaced Excel error with a blank cell"
    },
    {
      worksheetName: "People",
      address: "O2",
      originalError: "#ERROR!",
      originalFormula: "NA()",
      action: "Replaced cached formula error with a blank cell"
    },
    {
      worksheetName: "People",
      address: "P2",
      originalError: "#ERROR!",
      originalFormula: null,
      action: "Replaced Excel error with a blank cell"
    }
  ]);
  assert.deepEqual(summary.worksheets[1].errorReplacements, [
    {
      worksheetName: "Hidden Errors",
      address: "A2",
      originalError: "#REF!",
      originalFormula: null,
      action: "Replaced Excel error with a blank cell"
    },
    {
      worksheetName: "Hidden Errors",
      address: "B2",
      originalError: "#VALUE!",
      originalFormula: "NA()",
      action: "Replaced cached formula error with a blank cell"
    }
  ]);

  assert.equal(hash(originalBytes), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
  assert.deepEqual(workbookSnapshot(state.preservationWorkbook), preservationBefore);
});

test("duplicate comparison uses the same normalized text values in cleaning and review", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("People");
  sheet.addRows([
    ["Person Name", "Notes"],
    ["Bill Smith", "David Rasmussen"],
    ["Bill SmITH", "David     Rasmussen"],
    ["Brendan Wallace", "Different"]
  ]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(bytes);
  const state = await createWorkbookState(bytes, "normalized-duplicates.xlsx");

  const { summary } = await cleanWorkbook(state);
  assert.deepEqual(
    summary.worksheets[0].duplicateCandidates.map(({ originalRow, retainedOriginalRow }) => ({
      originalRow,
      retainedOriginalRow
    })),
    [{ originalRow: 3, retainedOriginalRow: 2 }]
  );

  const review = generateCleaningReview(state, state.uploadVersionId, {});
  const duplicateProposals = review.proposals.filter(
    (proposal) => proposal.issueType === "exact_duplicate_row"
  );
  assert.deepEqual(
    duplicateProposals.map(({ target, matches }) => ({
      row: target.row,
      retainedRow: matches.retainedRow
    })),
    [{ row: 3, retainedRow: 2 }]
  );
  const { buffer, summary: cleanedSummary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  assert.equal(cleaned.getWorksheet("People").rowCount, 3);
  assert.equal(cleaned.getWorksheet("People").getCell("A2").value, "Bill Smith");
  assert.equal(cleanedSummary.worksheets[0].removedRows, 1);
  assert.equal(hash(bytes), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
});

test("person-name headers normalize the six requested names without changing non-name text", async () => {
  const workbook = new ExcelJS.Workbook();
  const people = workbook.addWorksheet("People");
  people.addRows([
    [
      "Full Name", "Contact", "Salesperson", "Employee Name",
      "Person Name", "Contact Name", "Notes", "Email", "Website",
      "Code", "Formula", "Unrelated"
    ],
    [
      "Bill SmITH", "KEN Singh", "JONha Ma", "BRENDAN Wallace",
      "IVAN HINEY", "Steven MIChael", "NASA Smith", "person@example.com",
      "https://example.com/Path", "AB  123", { formula: "1+1", result: 2 },
      "UNRELATED TEXT"
    ]
  ]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "person-names.xlsx");
  const { buffer } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  const sheet = cleaned.getWorksheet("People");
  assert.deepEqual(
    ["A2", "B2", "C2", "D2", "E2", "F2"].map((address) => sheet.getCell(address).value),
    ["Bill Smith", "Ken Singh", "Jonha Ma", "Brendan Wallace", "Ivan Hiney", "Steven Michael"]
  );
  assert.equal(sheet.getCell("G2").value, "NASA Smith");
  assert.equal(sheet.getCell("H2").value, "person@example.com");
  assert.equal(sheet.getCell("I2").value, "https://example.com/Path");
  assert.equal(sheet.getCell("J2").value, "AB  123");
  assert.deepEqual(sheet.getCell("K2").value, { formula: "1+1", result: 2 });
  assert.equal(sheet.getCell("L2").value, "UNRELATED TEXT");
});

test("safe Phone/Salesperson data merges are unmerged without inventing missing names", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Contacts");
  sheet.addRows([
    ["Name", "Phone", "Salesperson", "Notes"],
    ["Customer A", "+1 (555) 111-1111", "Alex", "Call back"],
    ["Customer B", "+1 (555) 222-2222", "Sam", "Call back"],
    ["Customer C", "+1 (555) 123-4567", null, "Call back"]
  ]);
  sheet.mergeCells("B4:C4");
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "merged-phone-salesperson.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  const cleanedSheet = cleaned.getWorksheet("Contacts");
  assert.equal(cleanedSheet.getCell("B4").value, "+1 (555) 123-4567");
  assert.equal(cleanedSheet.getCell("B4").isMerged, false);
  assert.equal(cleanedSheet.getCell("C4").isMerged, false);
  assert.equal(cleanedSheet.getCell("C4").value, null);
  assert.equal(summary.worksheets[0].mergedRangeActions[0].originalRange, "B4:C4");
  assert.match(summary.worksheets[0].mergedRangeActions[0].reason, /distinct data fields/);
  assert.deepEqual(summary.worksheets[0].mergedRangeActions[0].unresolvedMissingFields, [
    "Salesperson"
  ]);
});

test("intentional title merges are retained and reported", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Report");
  sheet.getCell("A1").value = "Quarterly Report";
  sheet.mergeCells("A1:D1");
  sheet.addRow(["Name", "Phone", "Salesperson", "Notes"]);
  sheet.addRow(["Customer", "+1 555 123 4567", null, "Call back"]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "intentional-title-merge.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  assert.equal(cleaned.getWorksheet("Report").getCell("A1").isMerged, true);
  assert.equal(summary.worksheets[0].mergedRangeActions[0].originalRange, "A1:D1");
  assert.match(summary.worksheets[0].mergedRangeActions[0].reason, /Title or report-area/);
});

test("generic isolated data merges across unknown headers are corrected without moving or inventing values", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Miscellaneous");
  sheet.addRows([
    ["Record", "Description", "Merged_info", "Price", "Status"],
    ["A", "Widget", "Scratch", 12, "Open"],
    ["B", "Gadget", "Dent", 18, "Closed"],
    ["C", "Part", "Missing price", null, "Open"]
  ]);
  sheet.mergeCells("C4:D4");
  const numeric = workbook.addWorksheet("Numeric Merge");
  numeric.addRows([
    ["Record", "Metric", "Cost", "State"],
    ["A", 3, 5, "Open"],
    ["B", 7, 9, "Closed"],
    ["C", 11, null, "Open"]
  ]);
  numeric.mergeCells("B4:C4");
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "generic-data-merges.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);

  const details = cleaned.getWorksheet("Miscellaneous");
  assert.equal(details.getCell("C4").value, "Missing price");
  assert.equal(details.getCell("D4").value, null);
  assert.equal(details.getCell("C4").isMerged, false);
  assert.equal(details.getCell("D4").isMerged, false);
  assert.deepEqual(summary.worksheets[0].mergedRangeActions[0].unresolvedMissingFields, [
    "Price"
  ]);

  const numericOutput = cleaned.getWorksheet("Numeric Merge");
  assert.equal(numericOutput.getCell("B4").value, 11);
  assert.equal(numericOutput.getCell("C4").value, null);
  assert.equal(numericOutput.getCell("B4").isMerged, false);
});

test("a uniquely matching data profile locates an anchor in the correct generic field", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Unlabeled Fields");
  sheet.addRows([
    ["Record", "Field X", "Field Y", "Status"],
    ["A", "Alice", "+1 (555) 111-1111", "Open"],
    ["B", "Bob", "+1 (555) 222-2222", "Closed"],
    ["C", "+1 (555) 333-4444", null, "Open"]
  ]);
  sheet.mergeCells("B4:C4");
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "profiled-merge.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  const output = cleaned.getWorksheet("Unlabeled Fields");

  assert.equal(output.getCell("B4").value, null);
  assert.equal(output.getCell("C4").value, "+1 (555) 333-4444");
  assert.equal(output.getCell("B4").isMerged, false);
  assert.equal(output.getCell("C4").isMerged, false);
  assert.match(summary.worksheets[0].mergedRangeActions[0].action, /moved the anchor value/);
  assert.deepEqual(summary.worksheets[0].mergedRangeActions[0].unresolvedMissingFields, [
    "Field X"
  ]);
});

test("repeated and vertical data merges are retained with explicit reasons", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Structured");
  sheet.addRows([
    ["ID", "Field A", "Field B"],
    ["A", "same", null],
    ["B", "same", null],
    ["C", "same", null],
    ["D", "vertical", null]
  ]);
  sheet.mergeCells("B2:C2");
  sheet.mergeCells("B3:C3");
  sheet.mergeCells("B5:B6");
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "repeated-merges.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  const output = cleaned.getWorksheet("Structured");
  assert.equal(output.getCell("B2").isMerged, true);
  assert.equal(output.getCell("B3").isMerged, true);
  assert.equal(output.getCell("B5").isMerged, true);
  assert.match(summary.worksheets[0].mergedRangeActions[0].reason, /merged repeatedly/);
  assert.match(summary.worksheets[0].mergedRangeActions[2].reason, /Vertical merged/);
});

test("data-region merges are corrected across worksheets, ambiguous merges are retained, and serialized ranges are verified", async () => {
  const workbook = new ExcelJS.Workbook();
  const details = workbook.addWorksheet("Details");
  details.addRows([
    ["Record", "Field X", "Field Y", "Status"],
    ["A", "First", "Owner A", "Open"],
    ["B", "Second", "Owner B", "Closed"],
    ["C", "Anchor stays here", null, "Open"],
    ["D", "Another anchor", null, "Closed"],
    ["E", "Fifth", "Owner E", "Open"],
    ["F", "Sixth", "Owner F", "Closed"]
  ]);
  details.mergeCells("B4:C4");
  details.mergeCells("B5:C5");

  const payments = workbook.addWorksheet("Payments");
  payments.addRows([
    ["Record", "Category", "Description", "Amount", "Date", "Status", "Payment_Method"],
    ["A", "One", "First order", 10, "2025-01-01", "Open", "Card"],
    ["B", "Two", "Second order", 20, "2025-01-02", "Closed", "Cash"],
    ["C", "Three", "Third order", 30, "2025-01-03", "Open", "UPI"],
    ["D", "Four", "Fourth order", 40, "2025-01-04", "Closed", "Card"],
    ["E", "Five", "Fifth order", 50, "2025-01-05", "Open", "Cash"]
  ]);
  payments.mergeCells("G4:H4");

  const ambiguous = workbook.addWorksheet("Ambiguous");
  ambiguous.addRows([
    ["Record", "Field A", "Field B", "Status"],
    ["A", "Repeated label", null, "Open"],
    ["B", "Repeated label", null, "Closed"],
    ["C", "Repeated label", null, "Open"]
  ]);
  ambiguous.mergeCells("B2:C2");
  ambiguous.mergeCells("B3:C3");

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(originalBytes, "data-region-merges.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);

  assert.deepEqual(reopened.worksheets.map((sheet) => sheet.name), [
    "Details",
    "Payments",
    "Ambiguous"
  ]);

  const outputDetails = reopened.getWorksheet("Details");
  assert.deepEqual(outputDetails.model.merges, []);
  assert.equal(outputDetails.getCell("B4").value, "Anchor stays here");
  assert.equal(outputDetails.getCell("C4").value, null);
  assert.equal(outputDetails.getCell("B4").isMerged, false);
  assert.equal(outputDetails.getCell("C4").isMerged, false);
  assert.deepEqual(
    summary.worksheets[0].mergedRangeActions.map((item) => ({
      range: item.originalRange,
      action: item.action
    })),
    [
      {
        range: "B4:C4",
        action: "Unmerged separate data fields and preserved the anchor value in its original column."
      },
      {
        range: "B5:C5",
        action: "Unmerged separate data fields and preserved the anchor value in its original column."
      }
    ]
  );
  assert.ok(summary.worksheets[0].mergedRangeActions.every((item) =>
    /nearby records contain independent values/.test(item.reason)
  ));

  const outputPayments = reopened.getWorksheet("Payments");
  assert.deepEqual(outputPayments.model.merges, []);
  assert.equal(outputPayments.getCell("G4").value, "UPI");
  assert.equal(outputPayments.getCell("H4").value, null);
  assert.equal(outputPayments.getCell("G4").isMerged, false);
  assert.equal(outputPayments.getCell("G4").alignment.horizontal, "center");
  assert.equal(outputPayments.getCell("G4").alignment.vertical, "middle");
  assert.equal(summary.worksheets[1].mergedRangeActions[0].originalRange, "G4:H4");
  assert.equal(
    summary.worksheets[1].mergedRangeActions[0].action,
    "Unmerged the unused trailing-column merge and preserved the anchor value in its original column."
  );
  assert.match(summary.worksheets[1].mergedRangeActions[0].reason, /extends beyond the identified header span/);

  const outputAmbiguous = reopened.getWorksheet("Ambiguous");
  assert.deepEqual(outputAmbiguous.model.merges, ["B2:C2", "B3:C3"]);
  assert.deepEqual(
    summary.worksheets[2].mergedRangeActions.map((item) => item.action),
    ["Retained merged range.", "Retained merged range."]
  );
  assert.ok(summary.worksheets[2].mergedRangeActions.every((item) =>
    /merged repeatedly.*lacks enough nearby independent data/.test(item.reason)
  ));
});

test("duplicate removal ignores presentation differences and unrelated formulas but preserves cross-sheet references", async () => {
  const workbook = new ExcelJS.Workbook();
  const orders = workbook.addWorksheet("Orders");
  orders.addRows([
    ["Order", "Customer", "Amount"],
    [1003, "Charlie", 150],
    [],
    [1003, "Charlie", 150]
  ]);
  orders.getCell("A2").font = { color: { argb: "FFFF0000" } };
  orders.getCell("A4").font = { color: { argb: "FF0000FF" } };
  const audit = workbook.addWorksheet("Audit");
  audit.addRows([
    ["Expression"],
    [{ formula: "1+1", result: 2 }]
  ]);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "styled-duplicate-with-formula.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const output = new ExcelJS.Workbook();
  await output.xlsx.load(buffer);
  assert.deepEqual(output.getWorksheet("Orders").getRow(2).values.slice(1), [
    1003, "Charlie", 150
  ]);
  assert.equal(output.getWorksheet("Orders").rowCount, 2);
  assert.deepEqual(output.getWorksheet("Orders").getRow(2).values.slice(1), [
    1003, "Charlie", 150
  ]);
  assert.deepEqual(output.getWorksheet("Audit").getCell("A2").value, {
    formula: "1+1",
    result: 2
  });
  assert.equal(summary.worksheets[0].removedRows, 1);
  assert.equal(summary.worksheets[0].blankRowsDetected, 1);
  assert.equal(summary.worksheets[0].blankRowsRemoved, 1);
  assert.equal(summary.worksheets[0].duplicateRowsSkipped.length, 0);

  const referencingWorkbook = new ExcelJS.Workbook();
  const referenced = referencingWorkbook.addWorksheet("Orders");
  referenced.addRows([
    ["Order", "Customer", "Amount"],
    [1003, "Charlie", 150],
    [1003, "Charlie", 150]
  ]);
  referencingWorkbook.addWorksheet("Audit").getCell("A1").value = {
    formula: "Orders!C3",
    result: 150
  };
  const referencedBytes = Buffer.from(await referencingWorkbook.xlsx.writeBuffer());
  const referencedState = await createWorkbookState(referencedBytes, "referenced-duplicate.xlsx");
  const referencedResult = await generateCleanedWorkbook(referencedState);
  const reopenedReferenced = new ExcelJS.Workbook();
  await reopenedReferenced.xlsx.load(referencedResult.buffer);
  assert.equal(reopenedReferenced.getWorksheet("Orders").rowCount, 3);
  assert.match(referencedResult.summary.worksheets[0].duplicateRowsSkipped[0].reason, /another worksheet may depend/);
});

test("row-invariant formulas allow exact duplicate removal and preserve the first occurrence", async () => {
  const workbook = new ExcelJS.Workbook();
  const records = workbook.addWorksheet("Records");
  records.addRows([
    ["Name", "Value", "Calculation"],
    ["First", 1, null],
    ["Same", 10, null],
    ["Same", 10, null],
    ["Same", 10, null],
    ["Tail", 20, { formula: "1+2", result: 3 }]
  ]);
  const sourceBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(sourceBytes, "row-invariant-formula.xlsx");
  const { buffer, summary } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);

  const output = reopened.getWorksheet("Records");
  const resultSummary = summary.worksheets[0];
  assert.equal(output.rowCount, 4);
  assert.deepEqual(output.getRow(2).values.slice(1), ["First", 1]);
  assert.deepEqual(output.getRow(3).values.slice(1), ["Same", 10]);
  assert.equal(output.getCell("A3").value, "Same");
  assert.deepEqual(output.getCell("C4").value, {
    formula: "1+2",
    result: 3
  });
  assert.equal(resultSummary.duplicateCandidatesFound, 2);
  assert.equal(resultSummary.removedRows, 2);
  assert.deepEqual(resultSummary.duplicateRowsSkipped, []);
});

test("formula rows remain protected and formula expressions and results distinguish duplicates", async () => {
  const formulaRows = new ExcelJS.Workbook();
  formulaRows.addWorksheet("Records").addRows([
    ["Name", "Value"],
    ["Same", { formula: "1+1", result: 2 }],
    ["Same", { formula: "1+1", result: 2 }]
  ]);
  const formulaBytes = Buffer.from(await formulaRows.xlsx.writeBuffer());
  const formulaState = await createWorkbookState(formulaBytes, "formula-duplicate.xlsx");
  const formulaResult = await generateCleanedWorkbook(formulaState);
  const reopenedFormulaRows = new ExcelJS.Workbook();
  await reopenedFormulaRows.xlsx.load(formulaResult.buffer);
  const formulaSummary = formulaResult.summary.worksheets[0];
  assert.equal(reopenedFormulaRows.getWorksheet("Records").rowCount, 3);
  assert.deepEqual(reopenedFormulaRows.getWorksheet("Records").getCell("B2").value, {
    formula: "1+1",
    result: 2
  });
  assert.deepEqual(reopenedFormulaRows.getWorksheet("Records").getCell("B3").value, {
    formula: "1+1",
    result: 2
  });
  assert.equal(formulaSummary.duplicateCandidatesFound, 1);
  assert.equal(formulaSummary.removedRows, 0);
  assert.equal(formulaSummary.duplicateRowsSkipped.length, 1);
  assert.equal(
    formulaSummary.duplicateRowsSkipped[0].reason,
    "The duplicate row contains a formula."
  );

  const distinctFormulas = new ExcelJS.Workbook();
  distinctFormulas.addWorksheet("Records").addRows([
    ["Name", "Calculation"],
    ["Expression", { formula: "1+1", result: 2 }],
    ["Expression", { formula: "2+0", result: 2 }],
    ["Cached result", { formula: "1+1", result: 2 }],
    ["Cached result", { formula: "1+1", result: 3 }]
  ]);
  const distinctBytes = Buffer.from(await distinctFormulas.xlsx.writeBuffer());
  const distinctState = await createWorkbookState(distinctBytes, "distinct-formulas.xlsx");
  const distinctResult = await generateCleanedWorkbook(distinctState);
  const reopenedDistinct = new ExcelJS.Workbook();
  await reopenedDistinct.xlsx.load(distinctResult.buffer);
  const distinctSummary = distinctResult.summary.worksheets[0];
  assert.equal(reopenedDistinct.getWorksheet("Records").rowCount, 5);
  assert.equal(distinctSummary.duplicateCandidatesFound, 0);
  assert.equal(distinctSummary.removedRows, 0);
  assert.deepEqual(distinctSummary.duplicateRowsSkipped, []);
});

test("same-sheet row references and row-sensitive formulas protect duplicate rows", async () => {
  for (const formula of ["A3", "ROW()"]) {
    const workbook = new ExcelJS.Workbook();
    const records = workbook.addWorksheet("Records");
    records.addRows([
      ["Name", "Value", "Audit"],
      ["Same", 1, null],
      ["Same", 1, null],
      ["Tail", 2, { formula, result: formula === "ROW()" ? 4 : "Same" }]
    ]);
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    const state = await createWorkbookState(bytes, "row-dependent-formula.xlsx");
    const { buffer, summary } = await generateCleanedWorkbook(state);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(buffer);
    const output = reopened.getWorksheet("Records");
    const resultSummary = summary.worksheets[0];

    assert.equal(output.rowCount, 4);
    assert.equal(output.getCell("A2").value, "Same");
    assert.equal(output.getCell("A3").value, "Same");
    assert.deepEqual(output.getCell("C4").value, {
      formula,
      result: formula === "ROW()" ? 4 : "Same"
    });
    assert.equal(resultSummary.duplicateCandidatesFound, 1);
    assert.equal(resultSummary.removedRows, 0);
    assert.equal(resultSummary.duplicateRowsSkipped.length, 1);
    assert.match(
      resultSummary.duplicateRowsSkipped[0].reason,
      /formula contains references or syntax/
    );
  }
});

test("cross-sheet references and INDIRECT dependencies protect duplicate rows", async () => {
  for (const formula of ["Records!A3", 'INDIRECT("Records!A3")']) {
    const workbook = new ExcelJS.Workbook();
    const records = workbook.addWorksheet("Records");
    records.addRows([
      ["Name", "Value"],
      ["Same", 1],
      ["Same", 1]
    ]);
    workbook.addWorksheet("Audit").getCell("A1").value = {
      formula,
      result: "Same"
    };
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    const state = await createWorkbookState(bytes, "cross-sheet-reference.xlsx");
    const { buffer, summary } = await generateCleanedWorkbook(state);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(buffer);
    const output = reopened.getWorksheet("Records");
    const resultSummary = summary.worksheets[0];

    assert.equal(output.rowCount, 3);
    assert.equal(output.getCell("A2").value, "Same");
    assert.equal(output.getCell("A3").value, "Same");
    assert.equal(resultSummary.duplicateCandidatesFound, 1);
    assert.equal(resultSummary.removedRows, 0);
    assert.equal(resultSummary.duplicateRowsSkipped.length, 1);
    assert.match(
      resultSummary.duplicateRowsSkipped[0].reason,
      /another worksheet may depend/
    );
  }
});

test("serialized populated cells are centered after structural edits, including Payment Method", async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Payments");
  sheet.addRows([
    ["Payment Method", "Total", "Formula"],
    ["Card", 10, { formula: "B2*2", result: 20 }],
    ["Cash", 15, { formula: "B3*2", result: 30 }],
    ["Card", 10, { formula: "B4*2", result: 20 }]
  ]);
  for (const address of ["A1", "A2", "B2", "C2", "A4", "B4", "C4"]) {
    sheet.getCell(address).alignment = {
      horizontal: "right",
      vertical: "top",
      wrapText: true
    };
  }
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const state = await createWorkbookState(bytes, "payment-alignment.xlsx");
  const { buffer } = await generateCleanedWorkbook(state);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);
  const output = reopened.getWorksheet("Payments");
  for (let row = 1; row <= output.rowCount; row++) {
    for (let column = 1; column <= output.columnCount; column++) {
      const cell = output.getCell(row, column);
      if (cell.value !== null && cell.value !== undefined && cell.value !== "") {
        assert.equal(cell.alignment.horizontal, "center", cell.address);
        assert.equal(cell.alignment.vertical, "middle", cell.address);
      }
    }
  }
  assert.equal(output.getCell("A2").value, "Card");
  assert.deepEqual(output.getCell("C2").value, { formula: "B2*2", result: 20 });
});

test("duplicate rows are removed automatically in every worksheet while keeping first occurrences", async () => {
  const workbook = new ExcelJS.Workbook();
  const records = workbook.addWorksheet("Records");
  const hidden = workbook.addWorksheet("Hidden", { state: "hidden" });
  records.addRow(["Name", "Value"]);
  records.addRow(["Same", 1]);
  records.addRow([]);
  records.addRow([" Same ", 1]);
  records.addRow(["Same ", 1]);
  records.addRow([" Same", 1]);
  records.addRow(["Same", 1]);
  records.getColumn(1).width = 25;
  records.getColumn(2).width = 15;
  records.getRow(1).height = 28;
  records.getCell("A1").font = { bold: true, color: { argb: "FF112233" } };
  records.getCell("A1").fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE0F0FF" }
  };
  for (const rowNumber of [2, 4, 5, 6, 7]) {
    records.getRow(rowNumber).height = 20;
    records.getCell(`B${rowNumber}`).numFmt = "$#,##0.00";
  }
  hidden.addRow(["Name", "Value"]);
  hidden.addRow(["Same", 1]);
  hidden.addRow(["Same", 1]);

  const originalBytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const originalHash = hash(originalBytes);
  const state = await createWorkbookState(originalBytes, "automatic-duplicates.xlsx");
  const preservationBefore = workbookSnapshot(state.preservationWorkbook);
  const { summary: candidatesSummary } = await cleanWorkbook(state);
  const candidates = candidatesSummary.worksheets[0].duplicateCandidates;
  assert.deepEqual(candidates.map(({ originalRow, cleanedRow }) => ({
    originalRow,
    cleanedRow
  })), [
    { originalRow: 4, cleanedRow: 3 },
    { originalRow: 5, cleanedRow: 4 },
    { originalRow: 6, cleanedRow: 5 },
    { originalRow: 7, cleanedRow: 6 }
  ]);
  assert.ok(candidates.every(({ retainedOriginalRow }) => retainedOriginalRow === 2));

  const { buffer, summary } = await generateCleanedWorkbook(state);
  const result = new ExcelJS.Workbook();
  await result.xlsx.load(buffer);

  assert.deepEqual(result.worksheets.map(({ name, state: visibility }) => ({
    name,
    visibility
  })), [
    { name: "Records", visibility: "visible" },
    { name: "Hidden", visibility: "hidden" }
  ]);
  assert.equal(result.getWorksheet("Records").rowCount, 2);
  assert.deepEqual(result.getWorksheet("Records").getRow(2).values.slice(1), ["Same", 1]);
  assert.equal(result.getWorksheet("Hidden").rowCount, 2);
  assert.deepEqual(result.getWorksheet("Hidden").getRow(2).values.slice(1), ["Same", 1]);
  assert.equal(result.getWorksheet("Records").getColumn(1).width, 25);
  assert.equal(result.getWorksheet("Records").getColumn(2).width, 15);
  assert.equal(result.getWorksheet("Records").getRow(1).height, 28);
  assert.equal(result.getWorksheet("Records").getRow(2).height, 20);
  assert.equal(result.getWorksheet("Records").getCell("A1").font.bold, true);
  assert.equal(result.getWorksheet("Records").getCell("A1").fill.fgColor.argb, "FF1F4E78");
  assert.equal(result.getWorksheet("Records").getCell("B2").numFmt, "$#,##0.00");
  assert.equal(result.getWorksheet("Records").getCell("B2").value, 1);
  const recordSummary = summary.worksheets[0];
  assert.equal(recordSummary.duplicateCandidatesFound, 4);
  assert.equal(recordSummary.removedRows, 4);
  assert.deepEqual(recordSummary.duplicateRowsSkipped, []);
  assert.equal(recordSummary.emptyRowsRemoved, 1);
  assert.equal(recordSummary.blankRowsDetected, 1);
  assert.equal(recordSummary.blankRowsRemoved, 1);
  assert.equal(recordSummary.blankRowsRetained, 0);
  assert.equal(summary.worksheets[1].removedRows, 1);
  assert.equal(hash(originalBytes), originalHash);
  assert.equal(hash(state.getSourceBytes()), originalHash);
  assert.deepEqual(workbookSnapshot(state.preservationWorkbook), preservationBefore);
});

test("automatic duplicate removal keeps original upload bytes and the preserved workbook unchanged", async () => {
  async function duplicateWorkbook(name, marker) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Rows");
    sheet.addRow(["Name", "Value"]);
    sheet.addRow(["Same", 1]);
    sheet.addRow(["Same", 1]);
    sheet.addRow([marker, 2]);
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    return {
      bytes,
      state: await createWorkbookState(bytes, name)
    };
  }

  const first = await duplicateWorkbook("first.xlsx", "first");
  const second = await duplicateWorkbook("second.xlsx", "second");
  const originalHash = hash(first.bytes);
  const preservationBefore = workbookSnapshot(first.state.preservationWorkbook);
  const { buffer, summary } = await generateCleanedWorkbook(first.state);
  const cleaned = new ExcelJS.Workbook();
  await cleaned.xlsx.load(buffer);
  assert.equal(cleaned.getWorksheet("Rows").rowCount, 3);
  assert.equal(summary.worksheets[0].removedRows, 1);
  assert.equal(hash(first.bytes), originalHash);
  assert.equal(hash(first.state.getSourceBytes()), originalHash);
  assert.deepEqual(workbookSnapshot(first.state.preservationWorkbook), preservationBefore);
});

test("unsafe duplicate rows on formula, merged, and table worksheets are reported and kept safe", async () => {
  async function assertSkippedRemoval(workbook, expectedReason, verifyOutput = () => {}) {
    const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
    const state = await createWorkbookState(bytes, "unsafe-removal.xlsx");
    const { summary: beforeSummary } = await cleanWorkbook(state);
    const worksheetSummary = beforeSummary.worksheets[0];
    assert.equal(worksheetSummary.duplicateCandidates.length, 1);
    assert.equal(worksheetSummary.duplicateCandidates[0].safeToRemove, false);
    assert.match(worksheetSummary.duplicateCandidates[0].unsafeReason, expectedReason);

    const { buffer, summary } = await generateCleanedWorkbook(state);
    const reopened = new ExcelJS.Workbook();
    await reopened.xlsx.load(buffer);
    const resultSummary = summary.worksheets[0];
    assert.equal(resultSummary.removedRows, 0);
    assert.equal(resultSummary.duplicateRowsSkipped.length, 1);
    assert.match(resultSummary.duplicateRowsSkipped[0].reason, expectedReason);
    const candidateRow = worksheetSummary.duplicateCandidates[0].cleanedRow;
    assert.equal(reopened.worksheets[0].getCell(`A${candidateRow}`).value, "Same");
    assert.equal(reopened.worksheets[0].getCell(`A${candidateRow - 1}`).value, "Same");
    if (expectedReason.source === "merged cells") {
      assert.equal(reopened.worksheets[0].getCell("C5").isMerged, true);
    }
    if (expectedReason.source === "table") {
      assert.ok(reopened.worksheets[0].getTable("RecordsTable"));
    }
    verifyOutput(reopened.worksheets[0]);
  }

  const formulaBook = new ExcelJS.Workbook();
  const formulaSheet = formulaBook.addWorksheet("Formula");
  formulaSheet.addRows([
    ["Name", "Value", "Extra"],
    ["Same", 1, { formula: "1+1", result: 2 }],
    ["Same", 1, { formula: "1+1", result: 2 }]
  ]);
  await assertSkippedRemoval(
    formulaBook,
    /duplicate row contains a formula/,
    (sheet) => {
      assert.deepEqual(sheet.getCell("C2").value, {
        formula: "1+1",
        result: 2
      });
      assert.deepEqual(sheet.getCell("C3").value, {
        formula: "1+1",
        result: 2
      });
    }
  );

  const mergedBook = new ExcelJS.Workbook();
  const mergedSheet = mergedBook.addWorksheet("Merged");
  mergedSheet.addRow(["Name", "Value", "Extra", "More"]);
  mergedSheet.addRow(["Same", 1, null, null]);
  mergedSheet.addRow(["Same", 1, null, null]);
  mergedSheet.mergeCells("C5:D5");
  await assertSkippedRemoval(mergedBook, /merged cells/);

  const tableBook = new ExcelJS.Workbook();
  const tableSheet = tableBook.addWorksheet("Table");
  tableSheet.addTable({
    name: "RecordsTable",
    ref: "A1:B4",
    headerRow: true,
    columns: [{ name: "Name" }, { name: "Value" }],
    rows: [["Same", 1], ["Same", 1], ["Tail", 2]]
  });
  await assertSkippedRemoval(tableBook, /table/);
});
