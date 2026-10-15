
// ============================================================
// INSIGHTMATE - FRONTEND APPLICATION
// ============================================================

let currentDataset = null;
let resultChart = null;
let excelCleaningSession = null;
const EXCEL_CLEANING_SESSION_KEY = "insightmate.excelCleaningSession";

// ============================================================
// DOM HELPERS
// ============================================================

function getElement(id) {
  return document.getElementById(id);
}


function showElement(id) {
  const element = getElement(id);

  if (element) {
    element.hidden = false;
    element.classList.remove("hidden");
    element.style.display = "";
  }
}

function hideElement(id) {
  const element = getElement(id);

  if (element) {
    element.hidden = true;
    element.classList.add("hidden");
  }
}

// ============================================================
// FORMATTING HELPERS
// ============================================================

function escapeHtml(value) {
  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatNumber(value) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    return escapeHtml(value);
  }

  return number.toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

// ============================================================
// CLEAR ANALYSIS VISUALS
// ============================================================

function clearAnalysisVisuals() {
  if (resultChart) {
    resultChart.destroy();
    resultChart = null;
  }

  const insightContainer =
    getElement("businessInsight");

  if (insightContainer) {
    insightContainer.innerHTML = "";
  }
}

// ============================================================
// FILE NAME DISPLAY
// ============================================================

function updateSelectedFileName(fileName) {
  const possibleElements = [
    getElement("selectedFileName"),
    getElement("selectedFile"),
    getElement("fileNameDisplay")
  ];

  const displayElement =
    possibleElements.find(
      (element) => element
    );

  if (displayElement) {
    displayElement.textContent =
      fileName || "No file chosen";
  }
}

// ============================================================
// RESPONSE DATA HELPERS
// ============================================================

function extractColumnNames(data) {
  const candidates = [
    data?.columns,
    data?.profile?.columns,
    data?.dataset?.columns,
    data?.dataset?.profile?.columns,
    data?.overview?.columns
  ];

  for (const candidate of candidates) {
    if (!Array.isArray(candidate)) {
      continue;
    }

    const names = candidate
      .map((column) => {
        if (
          typeof column === "object" &&
          column !== null
        ) {
          return (
            column.name ??
            column.column ??
            column.key ??
            null
          );
        }

        return column;
      })
      .filter(
        (value) =>
          value !== null &&
          value !== undefined &&
          String(value).trim() !== ""
      );

    if (names.length > 0) {
      return names;
    }
  }

  // ----------------------------------------------------------
  // Final fallback:
  //
  // If preview rows exist, derive the column names directly
  // from the first row.
  // ----------------------------------------------------------

  const previewCandidates = [
    data?.preview,
    data?.previewRows,
    data?.dataset?.preview,
    data?.dataset?.previewRows
  ];

  for (const candidate of previewCandidates) {
    if (
      Array.isArray(candidate) &&
      candidate.length > 0 &&
      candidate[0] &&
      typeof candidate[0] === "object"
    ) {
      return Object.keys(candidate[0]);
    }
  }

  return [];
}

function extractPreviewRows(data) {
  const candidates = [
    data?.preview,
    data?.previewRows,
    data?.dataset?.preview,
    data?.dataset?.previewRows,
    data?.data?.preview
  ];

  for (const candidate of candidates) {
    if (
      Array.isArray(candidate) &&
      candidate.length > 0
    ) {
      return candidate;
    }
  }

  return [];
}

function extractRowCount(data, previewRows) {
  const candidates = [
    data?.rows,
    data?.rowCount,
    data?.totalRows,
    data?.overview?.rows,
    data?.overview?.rowCount,
    data?.overview?.totalRows,
    data?.profile?.rowCount,
    data?.profile?.totalRows,
    data?.profile?.rows,
    data?.dataset?.rows,
    data?.dataset?.rowCount,
    data?.dataset?.totalRows,
    data?.dataset?.profile?.rowCount,
    data?.dataset?.profile?.totalRows
  ];

  for (const candidate of candidates) {
    const number = Number(candidate);

    if (
      Number.isFinite(number) &&
      number > 0
    ) {
      return number;
    }
  }

  // Preview is normally only a small sample, so this is only
  // a fallback when no actual row count was returned.
  if (
    Array.isArray(previewRows) &&
    previewRows.length > 0
  ) {
    return previewRows.length;
  }

  return 0;
}

function extractProfile(data) {
  return (
    data?.profile ||
    data?.dataset?.profile ||
    {}
  );
}

function extractQuality(data) {
  return (
    data?.quality ||
    data?.dataset?.quality ||
    data?.profile?.quality ||
    {}
  );
}

function extractSheetName(data) {
  return (
    data?.sheetName ||
    data?.sheet ||
    data?.worksheet ||
    data?.dataset?.sheetName ||
    data?.dataset?.sheet ||
    ""
  );
}

// ============================================================
// INITIALIZE APPLICATION
// ============================================================

document.addEventListener(
  "DOMContentLoaded",
  () => {
    hideElement("qualitySection");
    hideElement("previewSection");
    hideElement("questionSection");
    hideElement("resultSection");
    hideElement("chartSection");
    hideElement("insightSection");

    const fileInput =
      getElement("fileInput");

    const uploadButton =
      getElement("uploadButton");

    const askButton =
      getElement("askButton");

    const questionInput =
      getElement("questionInput");

    // --------------------------------------------------------
    // Upload button
    // --------------------------------------------------------

    if (
      uploadButton &&
      fileInput
    ) {
      uploadButton.addEventListener(
        "click",
        () => {
          fileInput.click();
        }
      );
    }

    // --------------------------------------------------------
    // File selection
    // --------------------------------------------------------

    if (fileInput) {
      fileInput.addEventListener(
        "change",
        handleFileUpload
      );
    }

    // --------------------------------------------------------
    // Ask button
    // --------------------------------------------------------

    if (askButton) {
      askButton.addEventListener(
        "click",
        handleQuestion
      );
    }

    // --------------------------------------------------------
    // Enter key in question box
    // --------------------------------------------------------

    if (questionInput) {
      questionInput.addEventListener(
        "keydown",
        (event) => {
          if (
            event.key === "Enter" &&
            !event.shiftKey
          ) {
            event.preventDefault();
            handleQuestion();
          }
        }
      );
    }

    initializeExcelCleanerUi();
    initializeExistingExcelCleaner();

    console.log(
      "InsightMate frontend initialized."
    );
  }
);

// ============================================================
// FILE UPLOAD
// ============================================================

async function handleFileUpload() {
  const input =
    getElement("fileInput");

  const uploadButton =
    getElement("uploadButton");

  if (
    !input ||
    !input.files ||
    !input.files.length
  ) {
    updateSelectedFileName(
      "No file chosen"
    );

    return;
  }

  const file =
    input.files[0];

  updateSelectedFileName(
    file.name
  );

  resetExcelCleanerUi();

  if (uploadButton) {
    uploadButton.disabled = true;
    uploadButton.textContent =
      "Uploading...";
  }

  try {
    const formData =
      new FormData();

    formData.append(
      "file",
      file
    );

    console.log(
      "Uploading dataset:",
      file.name
    );

    const response =
      await fetch(
        "/api/upload",
        {
          method: "POST",
          body: formData
        }
      );

    console.log(
      "Upload response status:",
      response.status
    );

    let data;

    try {
      data =
        await response.json();
    } catch (jsonError) {
      throw new Error(
        `Server returned an invalid response (${response.status}).`
      );
    }

    console.log(
      "Upload response:",
      data
    );

    if (!response.ok) {
      throw new Error(
        data?.error ||
        data?.message ||
        `Upload failed with status ${response.status}.`
      );
    }

    // --------------------------------------------------------
    // Extract response data safely.
    // --------------------------------------------------------

    const columnNames =
      extractColumnNames(data);

    const previewRows =
      extractPreviewRows(data);

    const rowCount =
      extractRowCount(
        data,
        previewRows
      );

    const profile =
      extractProfile(data);

    const quality =
      extractQuality(data);

    const sheetName =
      extractSheetName(data);

    // --------------------------------------------------------
    // Store current dataset.
    // --------------------------------------------------------

    currentDataset = {
      fileName:
        data?.fileName ||
        data?.filename ||
        data?.originalName ||
        data?.dataset?.fileName ||
        file.name,

      sheetName,

      rows:
        rowCount,

      columns:
        columnNames,

      profile,

      quality,

      preview:
        previewRows
    };

    console.log(
      "Processed InsightMate dataset:",
      currentDataset
    );

    // --------------------------------------------------------
    // We only stop if we genuinely have no columns.
    //
    // Row count is allowed to be zero here because the server
    // may expose the count under another structure that we can
    // investigate without preventing the rest of the UI from
    // appearing.
    // --------------------------------------------------------

    if (
      !Array.isArray(
        currentDataset.columns
      ) ||
      currentDataset.columns.length === 0
    ) {
      throw new Error(
        "The dataset was uploaded successfully, but no columns could be detected."
      );
    }

    updateSelectedFileName(
      currentDataset.fileName
    );

    clearAnalysisVisuals();

    // --------------------------------------------------------
    // Render the dataset sections.
    // --------------------------------------------------------

    renderDatasetOverview(
      currentDataset
    );

    renderQuality(
      currentDataset.quality,
      currentDataset.profile
    );

    renderPreview(
      currentDataset.preview,
      currentDataset.columns
    );


    // --------------------------------------------------------
    // Show the sections after a successful upload.
    // --------------------------------------------------------

    showElement("qualitySection");
    showElement("previewSection");
    showElement("questionSection");
    if (/\.xlsx$/i.test(String(currentDataset.fileName || ""))) {
      hideElement("cleaningSection");
      showElement("excelCleaningSection");
      setExcelCleaningSessionFromUpload(data, currentDataset.fileName);
    } else {
      showElement("cleaningSection");
      hideElement("excelCleaningSection");
    }

    hideElement("resultSection");
    hideElement("chartSection");
    hideElement("insightSection");

    const resultContent = getElement("resultContent");

    if (resultContent) {
      resultContent.innerHTML = "";
    }


    console.log(
      "Dataset successfully loaded."
    );

    console.log(
      "Detected columns:",
      currentDataset.columns
    );

    console.log(
      "Detected rows:",
      currentDataset.rows
    );

  } catch (error) {
    console.error(
      "InsightMate upload error:",
      error
    );

    alert(
      `Could not upload the dataset.\n\n${error.message}`
    );

    updateSelectedFileName(
      file.name
    );

  } finally {
    if (uploadButton) {
      uploadButton.disabled = false;
      uploadButton.textContent =
        "Upload Dataset";
    }
  }
}

// ============================================================
// DATASET OVERVIEW
// ============================================================

function renderDatasetOverview(
  dataset
) {
  const fileNameElement =
    getElement("fileName");

  const rowCountElement =
    getElement("rowCount");

  const columnCountElement =
    getElement("columnCount");

  const sheetNameElement =
    getElement("sheetName");

  const missingCountElement =
    getElement("missingCount");

  const duplicateCountElement =
    getElement("duplicateCount");

  if (fileNameElement) {
    fileNameElement.textContent =
      dataset.fileName ||
      "Unknown";
  }

  if (rowCountElement) {
    rowCountElement.textContent =
      Number(
        dataset.rows || 0
      ).toLocaleString(
        "en-IN"
      );
  }

  if (columnCountElement) {
    columnCountElement.textContent =
      Number(
        dataset.columns?.length ||
        0
      ).toLocaleString(
        "en-IN"
      );
  }

  if (sheetNameElement) {
    sheetNameElement.textContent =
      dataset.sheetName ||
      "Default";
  }

  const profile =
    dataset.profile ||
    {};

  const quality =
    dataset.quality ||
    {};

  const missingValues =
    Number(
      quality.missingValues ??
      quality.missingCount ??
      profile.missingValues ??
      0
    );

  const duplicateRows =
    Number(
      quality.duplicateRows ??
      quality.duplicates ??
      profile.duplicateRows ??
      0
    );

  if (missingCountElement) {
    missingCountElement.textContent =
      missingValues.toLocaleString(
        "en-IN"
      );
  }

  if (duplicateCountElement) {
    duplicateCountElement.textContent =
      duplicateRows.toLocaleString(
        "en-IN"
      );
  }
}

// ============================================================
// DATA QUALITY
// ============================================================

function renderQuality(
  quality,
  profile
) {
  const container =
    getElement("qualityContent");

  if (!container) {
    return;
  }

  const totalRows =
    Number(
      profile?.rowCount ??
      profile?.totalRows ??
      currentDataset?.rows ??
      0
    );

  const totalColumns =
    Number(
      profile?.columnCount ??
      profile?.totalColumns ??
      currentDataset?.columns?.length ??
      0
    );

  const totalCells =
    totalRows *
    totalColumns;

  const missingValues =
    Number(
      quality?.missingValues ??
      quality?.missingCount ??
      profile?.missingValues ??
      0
    );

  const duplicateRows =
    Number(
      quality?.duplicateRows ??
      quality?.duplicates ??
      profile?.duplicateRows ??
      0
    );

  const totalIssues =
    missingValues +
    duplicateRows;

  let completeness = 100;

  if (totalCells > 0) {
    completeness =
      ((totalCells -
        missingValues) /
        totalCells) *
      100;
  }

  completeness =
    Math.max(
      0,
      Math.min(
        100,
        completeness
      )
    );

  const isHealthy = totalIssues === 0;
  const summaryElement = getElement("qualitySummary");

  if (summaryElement) {
    summaryElement.textContent = isHealthy
      ? "Your data looks good. No cleaning is needed."
      : "Your data needs attention. Review the items below.";
  }

  container.innerHTML = `
    <div class="quality-status ${
      isHealthy ? "quality-status-good" : "quality-status-warning"
    }">
      <strong>${isHealthy ? "Your data looks good" : "Your data needs attention"}</strong>
      <span>${
        isHealthy
          ? "No missing values or duplicate rows were found."
          : `${totalIssues.toLocaleString("en-IN")} issue${
              totalIssues === 1 ? "" : "s"
            } found. Use Data Cleaning & Export when you are ready.`
      }</span>
    </div>

    <div class="quality-grid">
      <div class="quality-card quality-card-good">
        <div class="quality-label">Data completeness</div>
        <div class="quality-value">${completeness.toFixed(1)}%</div>
        <div class="quality-help">How much of your data is filled in</div>
      </div>

      <div class="quality-card ${
        missingValues > 0 ? "quality-card-warning" : "quality-card-good"
      }">
        <div class="quality-label">Missing values</div>
        <div class="quality-value">${missingValues.toLocaleString("en-IN")}</div>
        <div class="quality-help">${
          missingValues === 0 ? "Nothing to fix" : "Some cells are empty"
        }</div>
      </div>

      <div class="quality-card ${
        duplicateRows > 0 ? "quality-card-warning" : "quality-card-good"
      }">
        <div class="quality-label">Duplicate rows</div>
        <div class="quality-value">${duplicateRows.toLocaleString("en-IN")}</div>
        <div class="quality-help">${
          duplicateRows === 0 ? "Nothing to fix" : "Repeated rows were found"
        }</div>
      </div>
    </div>

    <details class="quality-details">
      <summary>See dataset details</summary>
      <div class="quality-details-grid">
        <span>Rows <strong>${totalRows.toLocaleString("en-IN")}</strong></span>
        <span>Columns <strong>${totalColumns.toLocaleString("en-IN")}</strong></span>
        <span>Total issues <strong>${totalIssues.toLocaleString("en-IN")}</strong></span>
      </div>
    </details>
  `;
}

// ============================================================
// DATA PREVIEW
// ============================================================

function renderPreview(
  rows,
  columns
) {
  const container =
    getElement(
      "previewContent"
    );

  if (!container) {
    return;
  }

  if (
    !Array.isArray(rows) ||
    rows.length === 0
  ) {
    container.innerHTML = `
      <div class="empty-state">
        No preview data available.
      </div>
    `;

    return;
  }

  if (
    !Array.isArray(columns) ||
    columns.length === 0
  ) {
    container.innerHTML = `
      <div class="empty-state">
        No columns available.
      </div>
    `;

    return;
  }

  const headerHtml =
    columns
      .map(
        (column) =>
          `<th>${escapeHtml(
            column
          )}</th>`
      )
      .join("");

  const bodyHtml =
    rows
      .map((row) => {
        const cells =
          columns
            .map((column) => {
              const value =
                row?.[column];

              return `
                <td>
                  ${escapeHtml(
                    value
                  )}
                </td>
              `;
            })
            .join("");

        return `<tr>${cells}</tr>`;
      })
      .join("");

  container.innerHTML = `
    <div class="table-wrapper">

      <table class="data-table">

        <thead>
          <tr>
            ${headerHtml}
          </tr>
        </thead>

        <tbody>
          ${bodyHtml}
        </tbody>

      </table>

    </div>
  `;
}

// ============================================================
// ASK QUESTION
// ============================================================

async function handleQuestion() {
  const questionInput =
    getElement(
      "questionInput"
    );

  const askButton =
    getElement(
      "askButton"
    );

  const resultContent =
    getElement(
      "resultContent"
    );

  const question =
    questionInput?.value?.trim() ||
    "";

  if (!question) {
    alert(
      "Please enter a question about your dataset."
    );

    return;
  }

  if (
    !currentDataset ||
    !Array.isArray(
      currentDataset.columns
    ) ||
    currentDataset.columns.length === 0
  ) {
    alert(
      "Please upload a dataset first."
    );

    return;
  }

  if (askButton) {
    askButton.disabled = true;
    askButton.textContent =
      "Analyzing...";
  }

  showElement(
    "resultSection"
  );

  if (resultContent) {
    resultContent.innerHTML = `
      <div class="analysis-loading">
        <p>
          Analyzing your question...
        </p>
      </div>
    `;
  }

  clearAnalysisVisuals();

  try {
    const response =
      await fetch(
        "/api/analyze-question",
        {
          method: "POST",

          headers: {
            "Content-Type":
              "application/json"
          },

          body: JSON.stringify({
            question
          })
        }
      );

    let data;

    try {
      data =
        await response.json();
    } catch (jsonError) {
      throw new Error(
        `Server returned an invalid response (${response.status}).`
      );
    }

    if (!response.ok) {
      throw new Error(
        data?.error ||
        data?.message ||
        `Analysis failed with status ${response.status}.`
      );
    }

    displayAnalysisResult(
      data
    );

  } catch (error) {
    console.error(
      "InsightMate analysis error:",
      error
    );

    if (resultContent) {
      resultContent.innerHTML = `
        <div class="analysis-error">

          <h3>
            Analysis Error
          </h3>

          <p>
            ${escapeHtml(
              error.message
            )}
          </p>

        </div>
      `;
    }

  } finally {
    if (askButton) {
      askButton.disabled = false;
      askButton.textContent =
        "Ask InsightMate";
    }
  }
}

// ============================================================
// DISPLAY ANALYSIS RESULT
// ============================================================

function displayAnalysisResult(
  data
) {
  const resultContent =
    getElement(
      "resultContent"
    );

  if (!resultContent) {
    return;
  }

  if (
    data?.guardrailTriggered ||
    data?.blocked ||
    data?.status === "blocked"
  ) {
    displayGuardrailMessage(
      data
    );

    return;
  }

  const question =
    data?.question ||
    data?.analysis?.question ||
    "";

  const plan = {
    ...(data?.plan ||
      data?.analysisPlan ||
      data?.analysis?.plan ||
      {}),

    question:
      data?.question ||
      data?.analysis?.question ||
      ""
  };

  const analysis =
    data?.result ||
    data?.analysis ||
    {};

  const html = `
    <div class="analysis-complete">

      <h2>
        Analysis Complete
      </h2>

      <div class="analysis-question">

        <strong>
          Question
        </strong>

        <p>
          ${escapeHtml(
            question
          )}
        </p>

      </div>

      <details class="analysis-plan-details">
        <summary>How was this calculated?</summary>
        <div class="analysis-plan-section">
          ${renderAnalysisPlan(plan)}
        </div>
      </details>

      <div class="analysis-result-section">

        <h3>
          Result
        </h3>

        ${renderAnalysisOutput(
          analysis,
          plan
        )}

      </div>

    </div>
  `;

  resultContent.innerHTML =
    html;

  renderChart(
    analysis,
    plan
  );

  generateBusinessInsight(
    analysis,
    plan
  );
}

// ============================================================
// GUARDRAIL MESSAGE
// ============================================================

function displayGuardrailMessage(
  data
) {
  const resultContent =
    getElement(
      "resultContent"
    );

  if (!resultContent) {
    return;
  }

  const message =
    data?.message ||
    data?.error ||
    "This question cannot be answered safely from the available dataset.";

  resultContent.innerHTML = `
    <div class="analysis-guardrail">

      <h2>
        Unable to Analyze
      </h2>

      <p>
        ${escapeHtml(
          message
        )}
      </p>

    </div>
  `;
}

// ============================================================
// ANALYSIS PLAN
// ============================================================

function renderAnalysisPlan(
  plan
) {
  const operation =
    plan?.operation ||
    "N/A";

  const groupBy =
    plan?.groupColumn ||
    plan?.groupBy ||
    "N/A";

  const metric =
    plan?.metric ||
    "N/A";

  const aggregation =
    plan?.aggregation ||
    "N/A";

  const sort =
    plan?.sort ||
    "N/A";

  const limit =
    plan?.limit ??
    "N/A";

  return `
    <div class="analysis-plan">

      <div class="plan-row">

        <span class="plan-label">
          Operation
        </span>

        <span class="plan-value">
          ${escapeHtml(
            operation
          )}
        </span>

      </div>

      ${
        operation ===
        "GROUP_BY"
          ? `
            <div class="plan-row">

              <span class="plan-label">
                Group By
              </span>

              <span class="plan-value">
                ${escapeHtml(
                  groupBy
                )}
              </span>

            </div>
          `
          : ""
      }

      <div class="plan-row">

        <span class="plan-label">
          Metric
        </span>

        <span class="plan-value">
          ${escapeHtml(
            metric
          )}
        </span>

      </div>

      <div class="plan-row">

        <span class="plan-label">
          Aggregation
        </span>

        <span class="plan-value">
          ${escapeHtml(
            aggregation
          )}
        </span>

      </div>

      ${
        operation ===
        "GROUP_BY"
          ? `
            <div class="plan-row">

              <span class="plan-label">
                Sort
              </span>

              <span class="plan-value">
                ${escapeHtml(
                  sort
                )}
              </span>

            </div>

            <div class="plan-row">

              <span class="plan-label">
                Limit
              </span>

              <span class="plan-value">
                ${escapeHtml(
                  limit
                )}
              </span>

            </div>
          `
          : ""
      }

    </div>
  `;
}

// ============================================================
// ANALYSIS OUTPUT ROUTER
// ============================================================

function renderAnalysisOutput(
  analysis,
  plan
) {
  const operation =
    String(
      plan?.operation || ""
    ).toUpperCase();

  if (
    operation ===
    "GROUP_BY"
  ) {
    return renderGroupByResult(
      analysis,
      plan
    );
  }

  if (
    operation ===
    "AGGREGATE"
  ) {
    return renderAggregateResult(
      analysis,
      plan
    );
  }

  if (
    operation ===
    "COUNT"
  ) {
    return renderCountResult(
      analysis,
      plan
    );
  }

  return renderGenericRows(
    analysis
  );
}

// ============================================================
// GROUP BY RESULT
// ============================================================

function getGroupByRows(
  analysis
) {
  if (
    Array.isArray(
      analysis?.results
    )
  ) {
    return analysis.results;
  }

  if (
    Array.isArray(
      analysis?.rows
    )
  ) {
    return analysis.rows;
  }

  if (
    Array.isArray(
      analysis?.allRows
    )
  ) {
    return analysis.allRows;
  }

  return [];
}

function renderGroupByResult(
  analysis,
  plan
) {
  const rows =
    getGroupByRows(
      analysis
    );

  if (rows.length === 0) {
    return `
      <div class="empty-result">
        No grouped results were returned.
      </div>
    `;
  }

  const groupColumn =
    plan?.groupColumn ||
    plan?.groupBy ||
    "Group";

  const metric =
    plan?.metric ||
    "Value";

  const question =
    String(
      plan?.question ||
      analysis?.question ||
      ""
    ).toLowerCase();

  const wantsLowest =
    question.includes(
      "lowest"
    ) ||
    question.includes(
      "smallest"
    ) ||
    question.includes(
      "least"
    ) ||
    question.includes(
      "minimum"
    );

  const wantsHighest =
    question.includes(
      "highest"
    ) ||
    question.includes(
      "largest"
    ) ||
    question.includes(
      "most"
    ) ||
    question.includes(
      "maximum"
    );

  const wantsRanking =
    wantsLowest ||
    wantsHighest;

  const validRows =
    rows
      .map((row) => {
        const group =
          row?.group ??
          row?.[groupColumn] ??
          row?.label ??
          "";

        const value =
          row?.value ??
          row?.[metric] ??
          0;

        return {
          original: row,
          group,
          value: Number(value)
        };
      })
      .filter(
        (row) =>
          row.group !==
            undefined &&
          row.group !==
            null &&
          Number.isFinite(
            row.value
          )
      );

  if (
    validRows.length === 0
  ) {
    return `
      <div class="empty-result">
        No valid grouped results were returned.
      </div>
    `;
  }

  const tableRows =
    rows
      .map((row) => {
        const group =
          row?.group ??
          row?.[groupColumn] ??
          row?.label ??
          "";

        const value =
          row?.value ??
          row?.[metric] ??
          0;

        return `
          <tr>

            <td>
              ${escapeHtml(
                group
              )}
            </td>

            <td>
              ${formatNumber(
                value
              )}
            </td>

          </tr>
        `;
      })
      .join("");

  const tableHtml = `
    <div class="table-wrapper">

      <table class="result-table">

        <thead>

          <tr>

            <th>
              ${escapeHtml(
                groupColumn
              )}
            </th>

            <th>
              ${escapeHtml(
                metric
              )}
            </th>

          </tr>

        </thead>

        <tbody>
          ${tableRows}
        </tbody>

      </table>

    </div>
  `;

  // Normal GROUP_BY question:
  // table only.

  if (!wantsRanking) {
    return `
      <div class="group-by-result">
        ${tableHtml}
      </div>
    `;
  }

  let highlightedRow;

  if (wantsLowest) {
    highlightedRow =
      validRows.reduce(
        (best, row) =>
          row.value <
          best.value
            ? row
            : best,
        validRows[0]
      );
  } else {
    highlightedRow =
      validRows.reduce(
        (best, row) =>
          row.value >
          best.value
            ? row
            : best,
        validRows[0]
      );
  }

  const resultLabel =
    wantsLowest
      ? "Lowest result"
      : "Highest result";

  return `
    <div class="group-by-result">

      <div class="highlight-result">

        <div class="result-label">
          ${escapeHtml(
            resultLabel
          )}
        </div>

        <div class="result-main">

          ${escapeHtml(
            highlightedRow.group
          )}

          —
          ${formatNumber(
            highlightedRow.value
          )}

        </div>

      </div>

      ${tableHtml}

    </div>
  `;
}


// ============================================================
// AGGREGATE RESULT
// ============================================================

function renderAggregateResult(
  analysis,
  plan
) {
  const metric =
    plan?.metric ||
    "Value";

  const aggregation =
    plan?.aggregation ||
    "AGGREGATE";

  const value =
    analysis?.value ??
    analysis?.result ??
    analysis?.calculatedValue ??
    0;

  return `
    <div class="aggregate-result">

      <div class="highlight-result">

        <div class="result-label">
          Calculated value
        </div>

        <div class="result-main">
          ${formatNumber(value)}
        </div>

      </div>

      <div class="result-meta">

        <div>
          Metric:
          <strong>${escapeHtml(metric)}</strong>
          <span> | </span>
          Aggregation:
          <strong>${escapeHtml(aggregation)}</strong>
        </div>

      </div>

    </div>
  `;
}


// ============================================================
// COUNT RESULT
// ============================================================

function renderCountResult(
  analysis,
  plan
) {
  const value =
    analysis?.count ??
    analysis?.value ??
    analysis?.result ??
    analysis?.totalCount ??
    0;

  return `
    <div class="count-result">

      <div class="highlight-result">

        <div class="result-label">
          Total count
        </div>

        <div class="result-main">
          ${Number(
            value
          ).toLocaleString(
            "en-IN"
          )}
        </div>

      </div>

    </div>
  `;
}

// ============================================================
// GENERIC RESULT
// ============================================================

function renderGenericRows(
  analysis
) {
  if (
    Array.isArray(
      analysis?.rows
    ) &&
    analysis.rows.length > 0
  ) {
    const rows =
      analysis.rows;

    const columns =
      Object.keys(
        rows[0] || {}
      );

    const headerHtml =
      columns
        .map(
          (column) =>
            `<th>${escapeHtml(
              column
            )}</th>`
        )
        .join("");

    const bodyHtml =
      rows
        .map((row) => {
          const cells =
            columns
              .map(
                (column) =>
                  `<td>${escapeHtml(
                    row?.[column]
                  )}</td>`
              )
              .join("");

          return `<tr>${cells}</tr>`;
        })
        .join("");

    return `
      <div class="table-wrapper">

        <table class="result-table">

          <thead>

            <tr>
              ${headerHtml}
            </tr>

          </thead>

          <tbody>
            ${bodyHtml}
          </tbody>

        </table>

      </div>
    `;
  }

  const value =
    analysis?.value ??
    analysis?.result ??
    "No result available.";

  return `
    <div class="generic-result">
      ${escapeHtml(value)}
    </div>
  `;
}

// ============================================================
// CHART
// ============================================================

function renderChart(
  analysis,
  plan
) {
  const chartSection =
    getElement(
      "chartSection"
    );

  const chartCanvas =
    getElement(
      "resultChart"
    );

  if (!chartCanvas) {
    return;
  }

  const operation =
    String(
      plan?.operation || ""
    ).toUpperCase();

  if (
    operation !==
    "GROUP_BY"
  ) {
    hideElement(
      "chartSection"
    );

    return;
  }

  const rows =
    getGroupByRows(
      analysis
    );

  if (rows.length === 0) {
    hideElement(
      "chartSection"
    );

    return;
  }

  const groupColumn =
    plan?.groupColumn ||
    plan?.groupBy ||
    "Group";

  const labels = [];
  const values = [];

  rows.forEach(
    (row) => {
      const group =
        row?.group ??
        row?.[groupColumn] ??
        row?.label;

      const value =
        row?.value ??
        row?.[plan?.metric];

      const numericValue =
        Number(value);

      if (
        group !==
          undefined &&
        group !== null &&
        Number.isFinite(
          numericValue
        )
      ) {
        labels.push(
          String(group)
        );

        values.push(
          numericValue
        );
      }
    }
  );

  if (
    labels.length === 0
  ) {
    hideElement(
      "chartSection"
    );

    return;
  }

  if (
    typeof Chart ===
    "undefined"
  ) {
    console.warn(
      "Chart.js is not available."
    );

    hideElement(
      "chartSection"
    );

    return;
  }

  if (resultChart) {
    resultChart.destroy();
    resultChart = null;
  }

  resultChart =
    new Chart(
      chartCanvas,
      {
        type: "bar",

        data: {
          labels,

          datasets: [
            {
              label:
                plan?.metric ||
                "Value",

              data: values
            }
          ]
        },

        options: {
          responsive: true,

          maintainAspectRatio:
            false,

          plugins: {
            legend: {
              display: true
            }
          },

          scales: {
            y: {
              beginAtZero: true
            }
          }
        }
      }
    );

  if (chartSection) {
    showElement(
      "chartSection"
    );
  }
}

// ============================================================
// BUSINESS INSIGHT
// ============================================================

function generateBusinessInsight(
  analysis,
  plan
) {
  const insightSection =
    getElement(
      "insightSection"
    );

  const insightContainer =
    getElement(
      "businessInsight"
    );

  if (!insightContainer) {
    return;
  }

  const operation =
    String(
      plan?.operation || ""
    ).toUpperCase();

  if (
    operation !==
    "GROUP_BY"
  ) {
    hideElement(
      "insightSection"
    );

    return;
  }

  const rows =
    getGroupByRows(
      analysis
    );

  if (rows.length === 0) {
    hideElement(
      "insightSection"
    );

    return;
  }

  const groupColumn =
    plan?.groupColumn ||
    plan?.groupBy ||
    "Group";

  const metric =
    plan?.metric ||
    "value";

  const aggregation =
    String(
      plan?.aggregation || ""
    ).toUpperCase();

  const validRows =
    rows
      .map((row) => {
        const group =
          row?.group ??
          row?.[groupColumn] ??
          row?.label;

        const value =
          row?.value ??
          row?.[metric];

        return {
          group,
          value: Number(value)
        };
      })
      .filter(
        (row) =>
          row.group !==
            undefined &&
          row.group !==
            null &&
          Number.isFinite(
            row.value
          )
      );

  if (
    validRows.length === 0
  ) {
    hideElement(
      "insightSection"
    );

    return;
  }

  const highest =
    validRows.reduce(
      (best, row) =>
        row.value >
        best.value
          ? row
          : best,
      validRows[0]
    );

  const lowest =
    validRows.reduce(
      (best, row) =>
        row.value <
        best.value
          ? row
          : best,
      validRows[0]
    );

  const total =
    validRows.reduce(
      (sum, row) =>
        sum + row.value,
      0
    );

  let highestShare = 0;

  if (total !== 0) {
    highestShare =
      (highest.value /
        total) *
      100;
  }

  let metricPhrase =
    "value";

  if (
    aggregation ===
    "AVERAGE"
  ) {
    metricPhrase =
      "average amount";
  } else if (
    aggregation ===
    "SUM"
  ) {
    metricPhrase =
      "total amount";
  } else if (
    aggregation ===
    "COUNT"
  ) {
    metricPhrase =
      "count";
  } else if (
    aggregation ===
    "MIN"
  ) {
    metricPhrase =
      "minimum amount";
  } else if (
    aggregation ===
    "MAX"
  ) {
    metricPhrase =
      "maximum amount";
  }

  insightContainer.innerHTML = `
    <div class="business-insight-card">

      <p>

        <strong>
          ${escapeHtml(
            highest.group
          )}
        </strong>

        has the highest

        <strong>
          ${escapeHtml(
            metricPhrase
          )}
        </strong>

        at

        <strong>
          ${formatNumber(
            highest.value
          )}
        </strong>.

      </p>

      <p>

        <strong>
          ${escapeHtml(
            lowest.group
          )}
        </strong>

        has the lowest

        <strong>
          ${escapeHtml(
            metricPhrase
          )}
        </strong>

        at

        <strong>
          ${formatNumber(
            lowest.value
          )}
        </strong>.

      </p>

      <p>

        The highest segment represents approximately

        <strong>
          ${highestShare.toFixed(
            1
          )}%
        </strong>

        of the total across the displayed groups.

      </p>

    </div>
  `;

  if (insightSection) {
    showElement(
      "insightSection"
    );
  }
}

/* ============================================================
   DATA CLEANING & EXPORT
   ============================================================ */

document.addEventListener("DOMContentLoaded", () => {
  const cleaningSection = document.getElementById("cleaningSection");
  const previewButton = document.getElementById("previewCleaningButton");
  const exportButton = document.getElementById("exportCleanedButton");
  const statusElement = document.getElementById("cleaningStatus");
  const summaryElement = document.getElementById("cleaningSummary");
  const previewSection = document.getElementById("cleaningPreview");
  const previewCountElement = document.getElementById("cleaningPreviewCount");
  const previewContentElement = document.getElementById("cleaningPreviewContent");

  if (
    !cleaningSection ||
    !previewButton ||
    !exportButton ||
    !statusElement ||
    !summaryElement ||
    !previewSection ||
    !previewCountElement ||
    !previewContentElement
  ) {
    console.error("InsightMate cleaning interface is incomplete.");
    return;
  }

  const optionInputs = [
    document.getElementById("trimWhitespace"),
    document.getElementById("removeEmptyRows")
  ];

  if (optionInputs.some((input) => !input)) {
    console.error("One or more cleaning options are missing.");
    return;
  }

  let previewIsCurrent = false;

  function getSelectedOptions() {
    return {
      trimWhitespace: optionInputs[0].checked,
      removeEmptyRows: optionInputs[1].checked,
      removeDuplicates: true
    };
  }

  function invalidatePreview() {
    previewIsCurrent = false;
    exportButton.disabled = true;
    summaryElement.classList.add("hidden");
    previewSection.classList.add("hidden");
    statusElement.textContent =
      "Your selections changed. Preview the changes again before exporting.";
  }

  function renderCleaningTable(columns, rows) {
    if (!Array.isArray(columns) || !Array.isArray(rows)) {
      return "<p>Preview data is unavailable.</p>";
    }

    if (columns.length === 0) {
      return "<p>No columns are available to preview.</p>";
    }

    const headerHtml = columns
      .map((column) => `<th>${escapeHtml(column)}</th>`)
      .join("");

    const bodyHtml = rows.length
      ? rows.map((row) => `
          <tr>
            ${columns.map((column) => {
              const value = row?.[column];
              return `<td>${escapeHtml(
                value === null || value === undefined ? "" : String(value)
              )}</td>`;
            }).join("")}
          </tr>
        `).join("")
      : `<tr><td colspan="${columns.length}">No rows remain after the selected operations.</td></tr>`;

    return `
      <div class="table-wrapper">
        <table class="preview-table">
          <thead><tr>${headerHtml}</tr></thead>
          <tbody>${bodyHtml}</tbody>
        </table>
      </div>
    `;
  }

  async function readErrorResponse(response) {
    try {
      const data = await response.json();
      return data?.error || "The request could not be completed.";
    } catch {
      return "The request could not be completed.";
    }
  }

  previewButton.addEventListener("click", async () => {
    previewButton.disabled = true;
    exportButton.disabled = true;
    previewIsCurrent = false;
    statusElement.textContent = "Preparing cleaning preview...";
    summaryElement.classList.add("hidden");
    previewSection.classList.add("hidden");

    try {
      const response = await fetch("/api/cleaning/preview", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          options: getSelectedOptions()
        })
      });

      const data = await response.json();

      if (!response.ok || !data.success) {
        throw new Error(data.error || "Could not preview cleaning changes.");
      }

      const summary = data.summary || {};

      summaryElement.innerHTML = `
        <h3>Cleaning Preview Summary</h3>
        <ul>
          <li>Original rows: <strong>${Number(summary.originalRows) || 0}</strong></li>
          <li>Rows after cleaning: <strong>${Number(summary.finalRows) || 0}</strong></li>
          <li>Text cells trimmed: <strong>${Number(summary.trimmedCells) || 0}</strong></li>
          <li>Empty rows removed: <strong>${Number(summary.emptyRowsRemoved) || 0}</strong></li>
          <li>Duplicate rows removed: <strong>${Number(summary.duplicateRowsRemoved) || 0}</strong></li>
        </ul>
        <p>Duplicate rows are always removed automatically. Other selected operations are included in this preview.</p>
      `;

      previewCountElement.textContent =
        `Showing ${Array.isArray(data.preview) ? data.preview.length : 0} of ${Number(summary.finalRows) || 0} cleaned rows.`;

      previewContentElement.innerHTML = renderCleaningTable(
        data.columns,
        data.preview
      );

      summaryElement.classList.remove("hidden");
      previewSection.classList.remove("hidden");

      previewIsCurrent = true;
      exportButton.disabled = false;
      statusElement.textContent =
        "Preview ready. Review the summary and sample rows before exporting.";
    } catch (error) {
      statusElement.textContent =
        error?.message || "Could not preview cleaning changes.";
    } finally {
      previewButton.disabled = false;
    }
  });

  exportButton.addEventListener("click", async () => {
    if (!previewIsCurrent) {
      statusElement.textContent =
        "Please preview your selected cleaning operations first.";
      exportButton.disabled = true;
      return;
    }

    exportButton.disabled = true;
    statusElement.textContent = "Preparing your Excel file...";

    try {
      const response = await fetch("/api/cleaning/export", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          options: getSelectedOptions()
        })
      });

      if (!response.ok) {
        throw new Error(await readErrorResponse(response));
      }

      const blob = await response.blob();

      if (!blob.size) {
        throw new Error("The exported file is empty.");
      }

      const disposition = response.headers.get("Content-Disposition") || "";
      const filenameMatch = disposition.match(/filename="?([^";]+)"?/i);
      const filename = filenameMatch
        ? filenameMatch[1]
        : "cleaned-dataset.xlsx";

      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");

      link.href = downloadUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(downloadUrl);

      statusElement.textContent =
        "Download started. Your original uploaded dataset has not been changed.";
    } catch (error) {
      statusElement.textContent =
        error?.message || "Could not export the cleaned Excel file.";
    } finally {
      exportButton.disabled = !previewIsCurrent;
    }
  });

  optionInputs.forEach((input) => {
    input.addEventListener("change", invalidatePreview);
  });
});

/* ============================================================
   EXCEL WORKBOOK PROPOSAL REVIEW
   ============================================================ */

let workbookReviewState = null;
let workbookReviewBusy = false;

function safeReviewText(value, fallback = "") {
  if (value === null || value === undefined) {
    return fallback;
  }

  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    const text = String(value);
    return ["undefined", "nan", "[object object]"].includes(text.trim().toLowerCase())
      ? fallback
      : text;
  }

  return fallback;
}

function readableProposalValue(value) {
  if (value === null || value === undefined) {
    return "Not provided";
  }

  if (typeof value !== "object") {
    return safeReviewText(value, "Not provided");
  }

  const type = safeReviewText(value.type, "");
  if (type === "blank") return "(blank)";
  if (type === "error") return `Excel error: ${safeReviewText(value.value, "Unknown error")}`;
  if (type === "date") return safeReviewText(value.value, "Date value");
  if (type === "formula") {
    const formula = safeReviewText(value.formula, "formula");
    const result = value.result === undefined
      ? ""
      : ` (result: ${readableProposalValue(value.result)})`;
    return `Formula: ${formula}${result}`;
  }
  if (Object.prototype.hasOwnProperty.call(value, "value")) {
    return readableProposalValue(value.value);
  }

  try {
    const serialized = JSON.stringify(value);
    return serialized && serialized !== "{}" ? serialized : "Structured value";
  } catch {
    return "Structured value";
  }
}

function workbookReviewMessage(message) {
  const status = getElement("workbookReviewStatus");
  if (status) {
    status.textContent = safeReviewText(message, "The request could not be completed.");
  }
}

function reviewRequestError(data, response) {
  const message = typeof data?.error === "string"
    ? data.error.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim()
    : "";
  if (
    message &&
    !/\b(?:stack|syntaxerror|typeerror|referenceerror)\b|\bat\s+\S+:\d+:\d+/i.test(message) &&
    !/^(?:undefined|nan|\[object object\])$/i.test(message)
  ) {
    return message;
  }
  if (response.status === 409) return "The workbook or review has changed. Refreshing the current review is needed.";
  if (response.status === 404) return "The review or proposal is no longer available.";
  return "The request could not be completed. Please try again.";
}

async function sendWorkbookReviewRequest(url, body) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  let data;
  try {
    data = await response.json();
  } catch {
    throw new Error("The server returned an unreadable response.");
  }
  if (!response.ok || !data?.success) {
    const error = new Error(reviewRequestError(data, response));
    error.status = response.status;
    throw error;
  }
  return data;
}

function renderWorkbookWorksheets() {
  const container = getElement("workbookWorksheets");
  if (!container || !workbookReviewState) return;

  container.innerHTML = workbookReviewState.worksheets.map((worksheet, index) => {
    const name = safeReviewText(worksheet.name, "Unnamed worksheet");
    const header = worksheet.header || {};
    const rowCount = Number.isInteger(worksheet.rowCount) ? worksheet.rowCount : 0;
    const columnCount = Number.isInteger(worksheet.columnCount) ? worksheet.columnCount : 0;
    const headerStatus = safeReviewText(header.status, "confirmation_required");
    const confirmed = headerStatus === "confirmed";
    const skipped = headerStatus === "skipped";
    const candidate = headerStatus === "candidate";
    const currentRow = Number.isInteger(header.row) ? header.row : null;
    const suggestedRow = Number.isInteger(header.candidateRow)
      ? header.candidateRow
      : null;
    const action = skipped
      ? "skip"
      : (confirmed || candidate ? "confirm" : "");
    const selectedRow = currentRow ?? (candidate ? suggestedRow : null);
    const statusLabel = skipped
      ? "Skipped by user"
      : confirmed
        ? `Header confirmed on original row ${currentRow ?? "?"}`
        : candidate
          ? `Suggested header row ${suggestedRow ?? "?"} — confirm before review`
          : "Not safe to process automatically";
    const worksheetOrder = Number.isInteger(worksheet.order)
      ? worksheet.order + 1
      : index + 1;
    const reason = safeReviewText(
      header.reason,
      candidate
        ? "The first row appears suitable as a header, but please confirm it."
        : ""
    );
    const processingResult = Array.isArray(workbookReviewState.worksheetProcessingResults)
      ? workbookReviewState.worksheetProcessingResults.find(
          (result) => result?.worksheetKey === worksheet.worksheetKey
        )
      : null;
    const processingStatus = processingResult?.status
      ? safeReviewText(processingResult.status, "unavailable")
          .replace(/_/g, " ")
          .replace(/^\w/, (character) => character.toUpperCase())
      : workbookReviewState.reviewId
        ? "Unavailable"
        : "Pending";
    const proposals = Array.isArray(workbookReviewState.proposals)
      ? workbookReviewState.proposals
      : null;
    const findings = Array.isArray(workbookReviewState.findings)
      ? workbookReviewState.findings
      : null;
    const proposalCount = processingResult && proposals
      ? proposals.filter(
          (proposal) => proposal.worksheet?.worksheetKey === worksheet.worksheetKey
        ).length
      : null;
    const findingCount = processingResult && findings
      ? findings.filter(
          (finding) => finding.worksheetKey === worksheet.worksheetKey
        ).length
      : null;
    const processingReason = processingResult?.status === "skipped"
      ? safeReviewText(processingResult.reason, "")
      : "";
    const choiceUnsaved = workbookReviewState.unsavedWorksheetKeys?.has(worksheet.worksheetKey);

    return `
      <article class="workbook-worksheet-card" data-worksheet-card="${escapeHtml(worksheet.worksheetKey)}">
        <h3>${escapeHtml(`${worksheetOrder}. ${name}`)}</h3>
        <p class="worksheet-meta">
          <span data-worksheet-header-status>${escapeHtml(choiceUnsaved ? "Choice: Unsaved" : statusLabel)}</span><br>
          ${escapeHtml(String(rowCount))} original rows · ${escapeHtml(String(columnCount))} columns · Visibility: ${escapeHtml(safeReviewText(worksheet.visibility, "unknown"))}
        </p>
        <p class="worksheet-processing" data-worksheet-processing>
          ${choiceUnsaved
            ? "Unsaved changes — save choices to refresh results"
            : `Processing: <strong>${escapeHtml(processingStatus)}</strong><br>
              Proposals: ${proposalCount === null ? "Unavailable" : escapeHtml(String(proposalCount))} ·
              Findings: ${findingCount === null ? "Unavailable" : escapeHtml(String(findingCount))}
              ${processingReason ? `<br>${escapeHtml(processingReason)}` : ""}`}
        </p>
        ${!choiceUnsaved && headerStatus === "confirmation_required" ? `
          <p class="worksheet-safety-note">
            InsightMate cannot safely choose a header for this worksheet. Confirm a non-empty original row or skip it. The server will validate the selected row.
            ${reason ? `<br>${escapeHtml(reason)}` : ""}
          </p>
        ` : ""}
        <div class="worksheet-header-controls">
          <label>
            Worksheet choice
            <select data-header-action="${escapeHtml(worksheet.worksheetKey)}">
              <option value="" ${action === "" ? "selected" : ""}>Choose an action</option>
              <option value="confirm" ${action === "confirm" ? "selected" : ""}>Confirm a header row</option>
              <option value="skip" ${action === "skip" ? "selected" : ""}>Skip this worksheet</option>
            </select>
          </label>
          <label>
            Original header row
            <input
              type="number"
              min="1"
              max="${Math.max(1, rowCount)}"
              step="1"
              value="${selectedRow === null ? "" : escapeHtml(String(selectedRow))}"
              data-header-row="${escapeHtml(worksheet.worksheetKey)}"
              ${action === "skip" ? "disabled" : ""}
            >
          </label>
        </div>
      </article>
    `;
  }).join("");

  container.querySelectorAll("[data-header-action]").forEach((select) => {
    select.addEventListener("change", () => {
      const key = select.dataset.headerAction;
      const rowInput = container.querySelector(`[data-header-row="${CSS.escape(key)}"]`);
      if (rowInput) {
        rowInput.disabled = select.value === "skip";
        if (select.value === "skip") rowInput.value = "";
      }
      updateWorksheetChoiceState(key);
    });
  });
  container.querySelectorAll("[data-header-row]").forEach((input) => {
    input.addEventListener("input", () => {
      updateWorksheetChoiceState(input.dataset.headerRow);
    });
  });
}

function normalizeWorksheetChoice(action, row) {
  const rowText = String(row ?? "").trim();
  const numericRow = Number(rowText);
  return {
    action,
    row: action === "confirm" && rowText
      ? (Number.isInteger(numericRow) ? numericRow : rowText)
      : null
  };
}

function worksheetChoiceFromMetadata(worksheet) {
  const header = worksheet.header || {};
  if (header.status === "confirmed") {
    return normalizeWorksheetChoice("confirm", header.row);
  }
  if (header.status === "skipped") {
    return normalizeWorksheetChoice("skip", "");
  }
  if (header.status === "candidate") {
    return normalizeWorksheetChoice("confirm", header.candidateRow);
  }
  return normalizeWorksheetChoice("", "");
}

function worksheetChoiceFromControls(key) {
  const container = getElement("workbookWorksheets");
  const action = container?.querySelector(
    `[data-header-action="${CSS.escape(key)}"]`
  )?.value;
  const row = container?.querySelector(
    `[data-header-row="${CSS.escape(key)}"]`
  )?.value;
  return normalizeWorksheetChoice(action || "", row);
}

function updateWorksheetChoiceState(key) {
  if (!workbookReviewState || !key) return;
  const baseline = workbookReviewState.worksheetChoiceBaselines?.get(key);
  const current = worksheetChoiceFromControls(key);
  const unsaved = !baseline || baseline.action !== current.action || baseline.row !== current.row;
  if (unsaved) {
    workbookReviewState.unsavedWorksheetKeys.add(key);
  } else {
    workbookReviewState.unsavedWorksheetKeys.delete(key);
  }

  const worksheet = workbookReviewState.worksheets.find(
    (entry) => entry.worksheetKey === key
  );
  const card = getElement("workbookWorksheets")?.querySelector(
    `[data-worksheet-card="${CSS.escape(key)}"]`
  );
  if (!worksheet || !card) return;

  const header = worksheet.header || {};
  const headerStatus = safeReviewText(header.status, "confirmation_required");
  const headerLabel = headerStatus === "skipped"
    ? "Skipped by user"
    : headerStatus === "confirmed"
      ? `Header confirmed on original row ${Number.isInteger(header.row) ? header.row : "?"}`
      : headerStatus === "candidate"
        ? `Suggested header row ${Number.isInteger(header.candidateRow) ? header.candidateRow : "?"} — confirm before review`
        : "Not safe to process automatically";
  const headerStatusElement = card.querySelector("[data-worksheet-header-status]");
  if (headerStatusElement) {
    headerStatusElement.textContent = unsaved ? "Choice: Unsaved" : headerLabel;
  }

  const processingElement = card.querySelector("[data-worksheet-processing]");
  if (!processingElement) return;
  if (unsaved) {
    processingElement.textContent = "Unsaved changes — save choices to refresh results";
    card.querySelector(".worksheet-safety-note")?.remove();
  } else {
    renderWorkbookWorksheets();
  }
}

function updateProposalFilters() {
  if (!workbookReviewState) return;
  const worksheetFilter = getElement("proposalWorksheetFilter");
  const issueFilter = getElement("proposalIssueFilter");
  if (!worksheetFilter || !issueFilter) return;

  const previousWorksheet = worksheetFilter.value;
  const previousIssue = issueFilter.value;
  const proposals = Array.isArray(workbookReviewState.proposals)
    ? workbookReviewState.proposals
    : [];
  const worksheets = new Map(
    workbookReviewState.worksheets.map((worksheet) => [
      safeReviewText(worksheet.worksheetKey, ""),
      safeReviewText(worksheet.name, "Unnamed worksheet")
    ])
  );
  const issues = new Set();
  proposals.forEach((proposal) => {
    issues.add(safeReviewText(proposal.issueType, "Other"));
  });

  worksheetFilter.innerHTML = `<option value="">All worksheets</option>${
    [...worksheets.entries()].map(([key, name]) =>
      `<option value="${escapeHtml(key)}">${escapeHtml(name)}</option>`
    ).join("")
  }`;
  issueFilter.innerHTML = `<option value="">All issue types</option>${
    [...issues].sort().map((issue) =>
      `<option value="${escapeHtml(issue)}">${escapeHtml(issue.replace(/_/g, " "))}</option>`
    ).join("")
  }`;
  if (worksheets.has(previousWorksheet)) worksheetFilter.value = previousWorksheet;
  if (issues.has(previousIssue)) issueFilter.value = previousIssue;
}

function renderWorkbookProposals() {
  const area = getElement("workbookProposalArea");
  const list = getElement("workbookProposalList");
  if (!area || !list || !workbookReviewState) return;

  const proposals = Array.isArray(workbookReviewState.proposals)
    ? workbookReviewState.proposals
    : [];
  const counts = { pending: 0, approved: 0, rejected: 0 };
  proposals.forEach((proposal) => {
    if (Object.prototype.hasOwnProperty.call(counts, proposal.decision)) {
      counts[proposal.decision]++;
    }
  });
  getElement("pendingProposalCount").textContent = String(counts.pending);
  getElement("approvedProposalCount").textContent = String(counts.approved);
  getElement("rejectedProposalCount").textContent = String(counts.rejected);
  updateProposalFilters();

  const worksheetValue = getElement("proposalWorksheetFilter")?.value || "";
  const issueValue = getElement("proposalIssueFilter")?.value || "";
  const decisionValue = getElement("proposalDecisionFilter")?.value || "";
  const filtered = proposals.filter((proposal) =>
    (!worksheetValue || proposal.worksheet?.worksheetKey === worksheetValue) &&
    (!issueValue || proposal.issueType === issueValue) &&
    (!decisionValue || proposal.decision === decisionValue)
  );

  if (!filtered.length) {
    list.innerHTML = `<div class="proposal-empty">${
      proposals.length
        ? "No proposals match these filters."
        : "No cleaning proposals were found for the selected worksheets."
    }</div>`;
  } else {
    list.innerHTML = filtered.map((proposal) => {
      const worksheet = safeReviewText(proposal.worksheet?.name, "Unnamed worksheet");
      const issue = safeReviewText(proposal.issueType, "Cleaning proposal").replace(/_/g, " ");
      const decision = ["pending", "approved", "rejected"].includes(proposal.decision)
        ? proposal.decision
        : "unknown";
      const target = proposal.target || {};
      const coordinates = target.kind === "cell"
        ? `${safeReviewText(target.address, "Cell coordinate unavailable")} (row ${safeReviewText(target.row, "unknown")}, column ${safeReviewText(target.column, "unknown")})`
        : Number.isInteger(target.row)
          ? `Original row ${target.row}`
          : "Original coordinates unavailable";
      const location = `${worksheet} · ${coordinates}`;
      const originalValue = readableProposalValue(proposal.originalValue);
      const proposedValue = readableProposalValue(proposal.proposedValue);
      const retainedRow = Number.isInteger(proposal.matches?.retainedRow)
        ? ` The earliest eligible matching row ${proposal.matches.retainedRow} is retained.`
        : "";
      const values = proposal.originalValue !== undefined || proposal.proposedValue !== undefined
        ? `<div class="proposal-values">
            ${proposal.originalValue !== undefined ? `<div class="proposal-value"><strong>Original value</strong><span>${escapeHtml(originalValue)}</span></div>` : ""}
            ${proposal.proposedValue !== undefined ? `<div class="proposal-value"><strong>Proposed value</strong><span>${escapeHtml(proposedValue)}</span></div>` : ""}
          </div>`
        : "";

      return `
        <article class="proposal-card">
          <div class="proposal-meta">${escapeHtml(location)}</div>
          <h3>${escapeHtml(issue)}</h3>
          <p>${escapeHtml(safeReviewText(proposal.reason, "No additional explanation was provided."))}${escapeHtml(retainedRow)}</p>
          ${values}
          <p><span class="proposal-status ${escapeHtml(decision)}">${escapeHtml(decision)}</span></p>
          ${decision === "pending" ? `
            <div class="proposal-actions">
              <button type="button" class="approve-proposal" data-decision="approved" data-proposal-id="${escapeHtml(proposal.proposalId)}">Approve</button>
              <button type="button" class="reject-proposal" data-decision="rejected" data-proposal-id="${escapeHtml(proposal.proposalId)}">Reject</button>
            </div>
          ` : ""}
        </article>
      `;
    }).join("");
  }

  area.classList.remove("hidden");
  list.querySelectorAll("[data-decision][data-proposal-id]").forEach((button) => {
    button.addEventListener("click", () => {
      decideWorkbookProposal(button.dataset.proposalId, button.dataset.decision);
    });
  });
}

function mergeWorksheetHeaders(responseWorksheets) {
  if (!workbookReviewState || !Array.isArray(responseWorksheets)) return;
  const updated = new Map(responseWorksheets.map((worksheet) => [worksheet.worksheetKey, worksheet]));
  workbookReviewState.worksheets = workbookReviewState.worksheets.map((worksheet) => ({
    ...worksheet,
    ...(updated.get(worksheet.worksheetKey) || {})
  }));
  renderWorkbookWorksheets();
}

function worksheetResultsMatchSelections(results, selections) {
  if (!Array.isArray(results) || results.length !== selections.length) return false;
  const resultsByKey = new Map();
  for (const result of results) {
    if (!result?.worksheetKey || resultsByKey.has(result.worksheetKey)) return false;
    resultsByKey.set(result.worksheetKey, result);
  }
  return selections.every((selection) => {
    const result = resultsByKey.get(selection.worksheetKey);
    const expectedStatus = selection.action === "skip" ? "skipped" : "processed";
    return result?.status === expectedStatus;
  });
}

function worksheetHeadersMatchSelections(worksheets, selections) {
  if (!Array.isArray(worksheets) || worksheets.length !== selections.length) return false;
  const worksheetsByKey = new Map();
  for (const worksheet of worksheets) {
    if (!worksheet?.worksheetKey || worksheetsByKey.has(worksheet.worksheetKey)) return false;
    worksheetsByKey.set(worksheet.worksheetKey, worksheet);
  }
  return selections.every((selection) => {
    const worksheet = worksheetsByKey.get(selection.worksheetKey);
    if (selection.action === "skip") return worksheet?.header?.status === "skipped";
    return worksheet?.header?.status === "confirmed" &&
      worksheet.header.row === selection.headerRow;
  });
}

async function fetchAuthoritativeProposalReview(expectedSelections = null) {
  if (!workbookReviewState?.uploadVersionId) {
    throw new Error("Upload an Excel workbook before generating proposals.");
  }
  const review = await sendWorkbookReviewRequest(
    "/api/cleaning/proposals",
    { uploadVersionId: workbookReviewState.uploadVersionId }
  );
  if (
    review.uploadVersionId !== workbookReviewState.uploadVersionId ||
    typeof review.reviewId !== "string" ||
    !Array.isArray(review.proposals)
  ) {
    throw new Error("The server returned a review that does not match the active workbook.");
  }
  if (
    expectedSelections &&
    (!Array.isArray(review.findings) ||
      !worksheetResultsMatchSelections(review.worksheets, expectedSelections))
  ) {
    throw new Error("The server did not return authoritative results for the saved worksheet choices.");
  }
  const { worksheets: worksheetProcessingResults, ...reviewData } = review;
  workbookReviewState = {
    ...workbookReviewState,
    ...reviewData,
    findings: Array.isArray(review.findings) ? review.findings : null,
    worksheetProcessingResults: Array.isArray(worksheetProcessingResults)
      ? worksheetProcessingResults
      : []
  };
  renderWorkbookWorksheets();
  renderWorkbookProposals();
}

function resetExcelCleanerUi() {
  excelCleaningSession = null;
  try {
    sessionStorage.removeItem(EXCEL_CLEANING_SESSION_KEY);
  } catch {
    // The current page can still use the in-memory upload token.
  }
  hideElement("excelCleaningSection");
  const status = getElement("excelCleaningStatus");
  if (status) status.textContent = "Choose Clean & Format to review the summary. Duplicate rows are removed automatically when safe.";
  const summary = getElement("excelCleaningSummary");
  if (summary) {
    summary.innerHTML = "";
    summary.classList.add("hidden");
  }
  const generateButton = getElement("generateExcelButton");
  if (generateButton) generateButton.disabled = true;
  const downloadButton = getElement("downloadExcelButton");
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.hidden = true;
    downloadButton.classList.add("hidden");
  }
}

function setExcelCleaningSessionFromUpload(data, fileName) {
  if (
    typeof data?.uploadVersionId !== "string" ||
    typeof data?.cleaningToken !== "string" ||
    !data.uploadVersionId ||
    !data.cleaningToken
  ) {
    excelCleaningSession = null;
    const status = getElement("excelCleaningStatus");
    if (status) {
      status.textContent = "The secure cleaning session could not be started. Please upload this workbook again.";
    }
    getElement("cleanExcelButton").disabled = true;
    return;
  }
  excelCleaningSession = {
    uploadVersionId: data.uploadVersionId,
    cleaningToken: data.cleaningToken,
    fileName: safeReviewText(fileName, "workbook")
  };
  try {
    sessionStorage.setItem(
      EXCEL_CLEANING_SESSION_KEY,
      JSON.stringify(excelCleaningSession)
    );
  } catch {
    // Keep the token in memory when browser storage is unavailable.
  }
  const cleanButton = getElement("cleanExcelButton");
  if (cleanButton) cleanButton.disabled = false;
}

function displayCount(value) {
  const count = Number(value);
  return Number.isFinite(count) && count > 0 ? count : 0;
}

function renderExcelCleaningSummary(summary, resultReady = false) {
  const element = getElement("excelCleaningSummary");
  if (!element || !Array.isArray(summary?.worksheets)) {
    throw new Error("The cleaning summary could not be displayed.");
  }
  element.innerHTML = `
    <h3>Cleaning summary</h3>
    ${summary.worksheets.map((worksheet) => `
      <article class="excel-cleaning-summary-card">
        <h3>${escapeHtml(worksheet.name || "Worksheet")}</h3>
        <p>${escapeHtml(worksheet.explanation || "This worksheet was checked.")}</p>
        <p>
          Rows formatted: <strong>${displayCount(worksheet.rowsFormatted)}</strong> ·
          Text cells tidied: <strong>${displayCount(worksheet.spacesTrimmed)}</strong> ·
          Blank rows detected: <strong>${displayCount(worksheet.blankRowsDetected)}</strong> ·
          ${resultReady ? "Removed" : "To remove"}: <strong>${displayCount(
            resultReady ? worksheet.blankRowsRemoved : worksheet.blankRowsToRemove
          )}</strong> ·
          Retained: <strong>${displayCount(worksheet.blankRowsRetained)}</strong> ·
          Blank cells found: <strong>${displayCount(worksheet.blankCellsDetected)}</strong>
        </p>
        ${Array.isArray(worksheet.blankRowsRetainedByReason) && worksheet.blankRowsRetainedByReason.length
          ? `<p>Blank rows retained:</p><ul>${worksheet.blankRowsRetainedByReason.map((item) => `
              <li>${displayCount(item.count)} row${Number(item.count) === 1 ? "" : "s"}: ${escapeHtml(item.reason || "Could not be safely removed.")}</li>
            `).join("")}</ul>`
          : ""}
        <p>
          Repeated rows found: <strong>${displayCount(worksheet.duplicateCandidatesFound)}</strong> ·
          Unclear date values left unchanged: <strong>${displayCount(worksheet.ambiguousItemsLeftUnchanged)}</strong> ·
          Items kept for safety: <strong>${displayCount(worksheet.itemsSkippedForSafety)}</strong>
        </p>
        <p>
          Duplicate rows kept for safety: <strong>${Array.isArray(worksheet.duplicateRowsSkipped)
            ? worksheet.duplicateRowsSkipped.length
            : 0}</strong>
        </p>
        ${Array.isArray(worksheet.duplicateRowsSkipped) && worksheet.duplicateRowsSkipped.length
          ? `<ul>${worksheet.duplicateRowsSkipped.map((item) => `
              <li>Row ${displayCount(item.originalRow)} was kept: ${escapeHtml(item.reason || "It could not be safely removed.")}</li>
            `).join("")}</ul>`
          : ""}
        ${Array.isArray(worksheet.errorReplacements) && worksheet.errorReplacements.length
          ? `<p>Excel errors replaced with blank cells:</p><ul>${worksheet.errorReplacements.map((item) => `
              <li>${escapeHtml(item.address)}: ${escapeHtml(item.originalError)} — ${escapeHtml(item.action)}</li>
            `).join("")}</ul>`
          : ""}
        ${Array.isArray(worksheet.mergedRangeActions) && worksheet.mergedRangeActions.length
          ? `<p>Merged ranges:</p><ul>${worksheet.mergedRangeActions.map((item) => `
              <li>${escapeHtml(item.originalRange)}: ${escapeHtml(item.action)} ${escapeHtml(item.reason || "")}${Array.isArray(item.unresolvedMissingFields) &&
                item.unresolvedMissingFields.length
                ? ` Unresolved blank fields: ${item.unresolvedMissingFields.map(escapeHtml).join(", ")}.`
                : ""}</li>
            `).join("")}</ul>`
          : ""}
        ${resultReady ? `
          <p>
            The cleaned workbook is ready. Repeated rows removed automatically: <strong>${displayCount(worksheet.removedRows)}</strong>.
          </p>
        ` : `<p>Safe duplicate rows will be removed automatically when you create the cleaned workbook.</p>`}
      </article>
    `).join("")}
  `;
  element.classList.remove("hidden");
}

async function parseExcelCleanerResponse(response) {
  let data;
  try {
    data = await response.json();
  } catch {
    throw new Error(`The cleaning service returned an unreadable response (${response.status}).`);
  }
  if (!response.ok || data?.success !== true) {
    throw new Error(data?.error || `The cleaning request failed (${response.status}).`);
  }
  return data;
}

async function startExcelCleaning() {
  if (!excelCleaningSession) {
    throw new Error("Upload an Excel workbook before cleaning it.");
  }
  const cleanButton = getElement("cleanExcelButton");
  const generateButton = getElement("generateExcelButton");
  const downloadButton = getElement("downloadExcelButton");
  const status = getElement("excelCleaningStatus");
  if (cleanButton) cleanButton.disabled = true;
  if (generateButton) generateButton.disabled = true;
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.hidden = true;
    downloadButton.classList.add("hidden");
  }
  if (status) status.textContent = "Checking each worksheet and preparing your summary...";
  getElement("excelCleaningSummary")?.classList.add("hidden");

  try {
    const response = await fetch("/api/cleaning/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(excelCleaningSession)
    });
    const data = await parseExcelCleanerResponse(response);
    if (
      data.uploadVersionId !== excelCleaningSession.uploadVersionId ||
      !Array.isArray(data.summary?.worksheets)
    ) {
      throw new Error("The cleaning summary did not match the active workbook.");
    }
    excelCleaningSession.summary = data.summary;
    excelCleaningSession.downloadReady = false;
    renderExcelCleaningSummary(data.summary);
    if (generateButton) generateButton.disabled = false;
    if (status) status.textContent = "Review the summary, then create the cleaned workbook. Duplicate rows are removed automatically where safe.";
  } catch (error) {
    if (status) status.textContent = error?.message || "Could not prepare the workbook.";
  } finally {
    if (cleanButton) cleanButton.disabled = false;
  }
}

async function generateExcelCleaningResult() {
  if (!excelCleaningSession?.summary) return;
  const generateButton = getElement("generateExcelButton");
  const downloadButton = getElement("downloadExcelButton");
  const status = getElement("excelCleaningStatus");
  if (generateButton) generateButton.disabled = true;
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.hidden = true;
    downloadButton.classList.add("hidden");
  }
  excelCleaningSession.downloadReady = false;
  if (status) status.textContent = "Creating a separate cleaned workbook...";
  try {
    const response = await fetch("/api/cleaning/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        uploadVersionId: excelCleaningSession.uploadVersionId,
        cleaningToken: excelCleaningSession.cleaningToken
      })
    });
    const data = await parseExcelCleanerResponse(response);
    if (
      data.uploadVersionId !== excelCleaningSession.uploadVersionId ||
      data.downloadReady !== true ||
      !Array.isArray(data.summary?.worksheets)
    ) {
      throw new Error("The cleaned workbook is not ready for download.");
    }
    excelCleaningSession.summary = data.summary;
    excelCleaningSession.downloadReady = true;
    renderExcelCleaningSummary(data.summary, true);
    if (downloadButton) {
      downloadButton.disabled = false;
      downloadButton.hidden = false;
      downloadButton.classList.remove("hidden");
    }
    if (status) status.textContent = "Your cleaned workbook is ready. The uploaded original has not been changed.";
  } catch (error) {
    if (status) status.textContent = error?.message || "Could not create the cleaned workbook.";
  } finally {
    if (generateButton) generateButton.disabled = false;
  }
}

async function downloadExcelCleaningResult() {
  if (!excelCleaningSession?.downloadReady) return;
  const button = getElement("downloadExcelButton");
  const status = getElement("excelCleaningStatus");
  if (button) button.disabled = true;
  if (status) status.textContent = "Preparing your download...";
  try {
    const response = await fetch("/api/cleaning/download", {
      headers: {
        "X-Upload-Version-Id": excelCleaningSession.uploadVersionId,
        "X-Cleaning-Token": excelCleaningSession.cleaningToken
      }
    });
    if (!response.ok) {
      let message = `The download failed (${response.status}).`;
      try {
        const data = await response.json();
        message = data?.error || message;
      } catch {
        // Keep the clear status message when the server sends no JSON error.
      }
      throw new Error(message);
    }
    if (!String(response.headers.get("content-type") || "").includes(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )) {
      throw new Error("The server did not return an Excel workbook.");
    }
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = `${String(excelCleaningSession.fileName || "workbook")
      .replace(/\.[^.]+$/, "")
      .replace(/[^\w.-]+/g, "_") || "workbook"}_cleaned.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
    if (status) status.textContent = "Your cleaned workbook has been downloaded.";
  } catch (error) {
    if (status) status.textContent = error?.message || "Could not download the cleaned workbook.";
  } finally {
    if (button) button.disabled = false;
  }
}

function initializeExcelCleanerUi() {
  getElement("cleanExcelButton")?.addEventListener("click", startExcelCleaning);
  getElement("generateExcelButton")?.addEventListener("click", generateExcelCleaningResult);
  getElement("downloadExcelButton")?.addEventListener("click", downloadExcelCleaningResult);
}

async function initializeExistingExcelCleaner() {
  try {
    const response = await fetch("/api/dataset");
    if (!response.ok) {
      throw new Error(`Dataset status request failed (${response.status}).`);
    }
    const dataset = await response.json();
    if (
      dataset?.success !== true ||
      dataset.loaded !== true ||
      !/\.xlsx$/i.test(String(dataset.fileName || ""))
    ) {
      workbookReviewState = null;
      hideElement("excelCleaningSection");
      return;
    }

    hideElement("cleaningSection");
    showElement("excelCleaningSection");
    let savedSession = null;
    try {
      savedSession = JSON.parse(
        sessionStorage.getItem(EXCEL_CLEANING_SESSION_KEY) || "null"
      );
    } catch {
      savedSession = null;
    }
    if (
      savedSession?.fileName === dataset.fileName &&
      typeof savedSession.uploadVersionId === "string" &&
      typeof savedSession.cleaningToken === "string"
    ) {
      excelCleaningSession = savedSession;
      await startExcelCleaning();
    } else {
      const status = getElement("excelCleaningStatus");
      if (status) {
        status.textContent = "For security, upload this workbook again to start a cleaning session.";
      }
    }
  } catch (error) {
    workbookReviewState = null;
    console.warn("Could not check for an active Excel workbook:", error?.message || "Dataset status could not be loaded.");
  }
}

async function loadWorkbookReview() {
  workbookReviewState = null;
}

async function generateWorkbookProposals() {
  if (!workbookReviewState || workbookReviewBusy) return;
  const worksheetContainer = getElement("workbookWorksheets");
  const selections = [];

  for (const worksheet of workbookReviewState.worksheets) {
    const actionInput = worksheetContainer.querySelector(
      `[data-header-action="${CSS.escape(worksheet.worksheetKey)}"]`
    );
    const rowInput = worksheetContainer.querySelector(
      `[data-header-row="${CSS.escape(worksheet.worksheetKey)}"]`
    );
    if (!actionInput || !rowInput || !["confirm", "skip"].includes(actionInput.value)) {
      workbookReviewMessage(`Choose whether to confirm a header or skip "${safeReviewText(worksheet.name, "this worksheet")}".`);
      return;
    }
    if (actionInput.value === "skip") {
      selections.push({
        worksheetKey: worksheet.worksheetKey,
        action: "skip"
      });
      continue;
    }

    const headerRow = Number(rowInput.value);
    if (
      !Number.isInteger(headerRow) ||
      headerRow < 1 ||
      headerRow > worksheet.rowCount
    ) {
      workbookReviewMessage(`Enter a whole-number header row from 1 to ${worksheet.rowCount} for "${safeReviewText(worksheet.name, "this worksheet")}".`);
      return;
    }
    selections.push({
      worksheetKey: worksheet.worksheetKey,
      headerRow
    });
  }

  setWorkbookReviewBusy(true);
  workbookReviewMessage("Saving worksheet choices and generating proposals...");
  try {
    const headers = await sendWorkbookReviewRequest(
      "/api/cleaning/workbook/headers",
      {
        uploadVersionId: workbookReviewState.uploadVersionId,
        headerSelections: selections
      }
    );
    if (headers.reviewInvalidated) {
      workbookReviewState.reviewId = null;
      getElement("workbookProposalArea").classList.add("hidden");
    }
    if (!worksheetHeadersMatchSelections(headers.worksheets, selections)) {
      throw new Error("The server did not confirm the saved worksheet choices.");
    }
    await fetchAuthoritativeProposalReview(selections);
    mergeWorksheetHeaders(headers.worksheets);
    workbookReviewState.worksheetChoiceBaselines = new Map(
      headers.worksheets.map((worksheet) => [
        worksheet.worksheetKey,
        worksheetChoiceFromMetadata(worksheet)
      ])
    );
    workbookReviewState.unsavedWorksheetKeys.clear();
    renderWorkbookWorksheets();
    workbookReviewMessage("Proposals are ready. Approve or reject each pending item individually; no workbook data has been changed.");
  } catch (error) {
    workbookReviewMessage(error?.message || "Could not generate cleaning proposals.");
    if (error?.status === 409) {
      await refreshWorkbookReviewAfterStale();
    }
  } finally {
    setWorkbookReviewBusy(false);
  }
}

async function refreshWorkbookReviewAfterStale() {
  const fileName = workbookReviewState?.fileName || currentDataset?.fileName;
  workbookReviewMessage("This workbook or review is out of date. Refreshing the active review...");
  await loadWorkbookReview(fileName);
  if (!workbookReviewState) return;
  try {
    await fetchAuthoritativeProposalReview();
    workbookReviewMessage("The active workbook review has been refreshed. Check the current decisions before continuing.");
  } catch (error) {
    workbookReviewMessage(error?.message || "The current workbook review could not be refreshed.");
  }
}

async function refreshWorkbookProposals() {
  if (!workbookReviewState || workbookReviewBusy) return;
  setWorkbookReviewBusy(true);
  workbookReviewMessage("Refreshing the authoritative proposal list...");
  try {
    await fetchAuthoritativeProposalReview();
    workbookReviewMessage("Proposal list refreshed from the active workbook review.");
  } catch (error) {
    workbookReviewMessage(error?.message || "Could not refresh proposals.");
    if (error?.status === 409) {
      await refreshWorkbookReviewAfterStale();
    }
  } finally {
    setWorkbookReviewBusy(false);
  }
}

async function decideWorkbookProposal(proposalId, decision) {
  if (
    !workbookReviewState?.reviewId ||
    !["approved", "rejected"].includes(decision) ||
    workbookReviewBusy
  ) {
    return;
  }
  setWorkbookReviewBusy(true);
  workbookReviewMessage(`Recording ${decision} decision...`);
  try {
    const result = await sendWorkbookReviewRequest(
      `/api/cleaning/proposals/${encodeURIComponent(proposalId)}/decision`,
      {
        uploadVersionId: workbookReviewState.uploadVersionId,
        reviewId: workbookReviewState.reviewId,
        decision
      }
    );
    if (
      result.uploadVersionId !== workbookReviewState.uploadVersionId ||
      result.reviewId !== workbookReviewState.reviewId ||
      result.proposal?.proposalId !== proposalId
    ) {
      throw new Error("The server response did not match this proposal. Refresh the review before continuing.");
    }
    workbookReviewState.proposals = workbookReviewState.proposals.map((proposal) =>
      proposal.proposalId === proposalId ? result.proposal : proposal
    );
    renderWorkbookProposals();
    workbookReviewMessage(`Decision recorded as ${decision}. The workbook itself remains unchanged.`);
  } catch (error) {
    workbookReviewMessage(error?.message || "The decision could not be recorded.");
    if (error?.status === 409 || error?.status === 404) {
      await refreshWorkbookReviewAfterStale();
    }
  } finally {
    setWorkbookReviewBusy(false);
  }
}

function setWorkbookReviewBusy(busy) {
  workbookReviewBusy = busy;
  [
    "generateWorkbookProposalsButton",
    "refreshWorkbookProposalsButton"
  ].forEach((id) => {
    const button = getElement(id);
    if (button) button.disabled = busy;
  });
  getElement("workbookProposalList")
    ?.querySelectorAll("button")
    .forEach((button) => { button.disabled = busy; });
}

document.addEventListener("DOMContentLoaded", () => {
  getElement("generateWorkbookProposalsButton")
    ?.addEventListener("click", generateWorkbookProposals);
  getElement("refreshWorkbookProposalsButton")
    ?.addEventListener("click", refreshWorkbookProposals);
  [
    "proposalWorksheetFilter",
    "proposalIssueFilter",
    "proposalDecisionFilter"
  ].forEach((id) => {
    getElement(id)?.addEventListener("change", renderWorkbookProposals);
  });
});

/* ============================================================
   INSIGHTMATE — CONSISTENT NUMERIC COLUMN ALIGNMENT
   Aligns numeric values vertically in all HTML tables.
   Does not change calculations or data.
   ============================================================ */

(function enableNumericColumnAlignment() {
  function isNumericTableValue(value) {
    const text = String(value ?? "").trim();

    if (!text) {
      return false;
    }

    const normalized = text
      .replace(/[₹,\s%]/g, "")
      .replace(/^\((.*)\)$/, "-$1");

    return (
      normalized !== "" &&
      Number.isFinite(Number(normalized))
    );
  }

  function addNumericAlignmentStyles() {
    if (
      document.getElementById(
        "insightMateNumericAlignmentStyles"
      )
    ) {
      return;
    }

    const style = document.createElement("style");

    style.id = "insightMateNumericAlignmentStyles";

    style.textContent = `
      table th.numeric-cell,
      table td.numeric-cell {
        text-align: right !important;
        font-variant-numeric: tabular-nums;
        white-space: nowrap;
      }
    `;

    document.head.appendChild(style);
  }

  function alignNumericColumns(table) {
    if (!(table instanceof HTMLTableElement)) {
      return;
    }

    const bodyRows = Array.from(
      table.querySelectorAll("tbody tr")
    );

    const usableRows = bodyRows.filter((row) => {
      const cells = Array.from(row.cells);

      return (
        cells.length > 0 &&
        cells.every((cell) => cell.colSpan === 1)
      );
    });

    if (usableRows.length === 0) {
      return;
    }

    const columnCount = Math.max(
      0,
      ...usableRows.map((row) => row.cells.length)
    );

    for (
      let columnIndex = 0;
      columnIndex < columnCount;
      columnIndex += 1
    ) {
      const values = usableRows
        .map(
          (row) =>
            row.cells[columnIndex]
              ?.textContent?.trim() ?? ""
        )
        .filter((value) => value !== "");

      const isNumericColumn =
        values.length > 0 &&
        values.every(isNumericTableValue);

      if (!isNumericColumn) {
        continue;
      }

      usableRows.forEach((row) => {
        const cell = row.cells[columnIndex];

        if (cell) {
          cell.classList.add("numeric-cell");
        }
      });

      const headerRows = Array.from(
        table.querySelectorAll("thead tr")
      );

      headerRows.forEach((row) => {
        const headerCell = row.cells[columnIndex];

        if (headerCell && headerCell.colSpan === 1) {
          headerCell.classList.add("numeric-cell");
        }
      });
    }
  }

  function alignAllNumericColumns() {
    addNumericAlignmentStyles();

    document
      .querySelectorAll("table")
      .forEach(alignNumericColumns);
  }

  function startAlignment() {
    alignAllNumericColumns();

    const observer = new MutationObserver(() => {
      alignAllNumericColumns();
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener(
      "DOMContentLoaded",
      startAlignment,
      { once: true }
    );
  } else {
    startAlignment();
  }
})();
