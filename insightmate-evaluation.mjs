import fs from "fs";
import path from "path";

const BASE_URL = "http://localhost:3000";
const DATASET_PATH = path.join(process.cwd(), "Cleaned_Organized_Data.xlsx");

const tests = [
  {
    name: "GROUP_BY + SUM — Total amount by Segment",
    question: "What is the total amount by Segment?",
    expected: {
      type: "GROUP_BY",
      aggregation: "SUM",
      groups: {
        Consumer: 195580.97,
        Corporate: 121885.93,
        "Home Office": 74255.00
      }
    }
  },
  {
    name: "GROUP_BY + AVERAGE — Average amount by Segment",
    question: "What is the average amount by Segment?",
    expected: {
      type: "GROUP_BY",
      aggregation: "AVERAGE",
      groups: {
        "Home Office": 566.83,
        Corporate: 493.47,
        Consumer: 440.50
      }
    }
  },
  {
    name: "GROUP_BY + MAX — Maximum amount by Segment",
    question: "What is the maximum amount by Segment?",
    expected: {
      type: "GROUP_BY",
      aggregation: "MAX",
      groups: {
        "Home Office": 23661.23,
        Consumer: 8805.04,
        Corporate: 8167.42
      }
    }
  },
  {
    name: "GROUP_BY + MIN — Minimum amount by Segment",
    question: "What is the minimum amount by Segment?",
    expected: {
      type: "GROUP_BY",
      aggregation: "MIN",
      groups: {
        "Home Office": 3.44,
        Consumer: 2.06,
        Corporate: 1.17
      }
    }
  },
  {
    name: "AGGREGATE + SUM — Overall total",
    question: "What is the total amount?",
    expected: {
      type: "AGGREGATE",
      aggregation: "SUM",
      value: 391721.91
    }
  },
  {
    name: "AGGREGATE + AVERAGE — Overall average",
    question: "What is the average amount?",
    expected: {
      type: "AGGREGATE",
      aggregation: "AVERAGE",
      value: 476.55
    }
  },
  {
    name: "AGGREGATE + MAX — Overall maximum",
    question: "What is the maximum order amount?",
    expected: {
      type: "AGGREGATE",
      aggregation: "MAX",
      value: 23661.23
    }
  },
  {
    name: "AGGREGATE + MIN — Overall minimum",
    question: "What is the minimum order amount?",
    expected: {
      type: "AGGREGATE",
      aggregation: "MIN",
      value: 1.17
    }
  },
  {
    name: "COUNT — Consumer",
    question: "How many orders are from Consumer?",
    expected: {
      type: "COUNT",
      value: 444
    }
  },
  {
    name: "COUNT — Corporate",
    question: "How many orders are from Corporate?",
    expected: {
      type: "COUNT",
      value: 247
    }
  },
  {
    name: "COUNT — Home Office",
    question: "How many orders are from Home Office?",
    expected: {
      type: "COUNT",
      value: 131
    }
  },
  {
    name: "COUNT — All orders",
    question: "How many orders are there?",
    expected: {
      type: "COUNT",
      value: 822
    }
  },
  {
    name: "FILTER + SUM — Consumer total",
    question: "What is the total amount for Consumer orders?",
    expected: {
      type: "AGGREGATE",
      aggregation: "SUM",
      value: 195580.97
    }
  },
  {
    name: "FILTER + AVERAGE — Corporate average",
    question: "What is the average amount for Corporate orders?",
    expected: {
      type: "AGGREGATE",
      aggregation: "AVERAGE",
      value: 493.47
    }
  },
  {
    name: "FILTER + MAX — Consumer maximum",
    question: "What is the maximum order amount for Consumer orders?",
    expected: {
      type: "AGGREGATE",
      aggregation: "MAX",
      value: 8805.04
    }
  },
  {
    name: "FILTER + MIN — Home Office minimum",
    question: "What is the minimum order amount for Home Office orders?",
    expected: {
      type: "AGGREGATE",
      aggregation: "MIN",
      value: 3.44
    }
  }
];

function nearlyEqual(actual, expected) {
  if (
    typeof actual !== "number" ||
    typeof expected !== "number"
  ) {
    return actual === expected;
  }

  return Math.abs(actual - expected) < 0.01;
}

function getGroupResults(result) {
  if (!result || !Array.isArray(result.rows)) {
    return {};
  }

  const output = {};

  for (const row of result.rows) {
    output[row.group] = row.value;
  }

  return output;
}

function assertEqual(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(
      `${label}: expected ${expected}, got ${actual}`
    );
  }
}

function assertNumber(actual, expected, label) {
  if (!nearlyEqual(actual, expected)) {
    throw new Error(
      `${label}: expected ${expected}, got ${actual}`
    );
  }
}

function assertGroups(actual, expected) {
  const actualKeys = Object.keys(actual).sort();
  const expectedKeys = Object.keys(expected).sort();

  if (
    JSON.stringify(actualKeys) !==
    JSON.stringify(expectedKeys)
  ) {
    throw new Error(
      `Group names differ. Expected: ${expectedKeys.join(", ")}. Actual: ${actualKeys.join(", ")}`
    );
  }

  for (const group of expectedKeys) {
    assertNumber(
      actual[group],
      expected[group],
      `Group "${group}"`
    );
  }
}

async function checkHealth() {
  const response = await fetch(
    `${BASE_URL}/api/health`
  );

  if (!response.ok) {
    throw new Error(
      `Health check failed with HTTP ${response.status}`
    );
  }

  const data = await response.json();

  assertEqual(
    data.status,
    "ok",
    "Health status"
  );

  assertEqual(
    data.serverVersion,
    "analysis-engine-v4.4",
    "Server version"
  );

  assertEqual(
    data.ollamaModel,
    "gemma4:e2b-it-qat",
    "Ollama model"
  );

  return data;
}

async function uploadDataset() {
  if (!fs.existsSync(DATASET_PATH)) {
    throw new Error(
      `Dataset not found: ${DATASET_PATH}`
    );
  }

  const fileBuffer =
    fs.readFileSync(DATASET_PATH);

  const blob = new Blob(
    [fileBuffer],
    {
      type:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
  );

  const formData = new FormData();

  formData.append(
    "file",
    blob,
    path.basename(DATASET_PATH)
  );

  const response = await fetch(
    `${BASE_URL}/api/upload`,
    {
      method: "POST",
      body: formData
    }
  );

  if (!response.ok) {
    const text = await response.text();

    throw new Error(
      `Dataset upload failed with HTTP ${response.status}: ${text}`
    );
  }

  const data =
    await response.json();

  if (!data.success) {
    throw new Error(
      "Dataset upload returned success=false."
    );
  }

  assertEqual(
    data.rows,
    822,
    "Dataset row count"
  );

  assertEqual(data.columns, 7, "Dataset column count");
  assertEqual(data.columnNames.length, 7, "Dataset column names");

  if (!data.quality || typeof data.quality !== "object") {
    throw new Error("Dataset quality profile is missing.");
  }

  return data;
}

async function runTest(test) {
  const response = await fetch(
    `${BASE_URL}/api/analyze-question`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        question: test.question
      })
    }
  );

  if (!response.ok) {
    const text =
      await response.text();

    throw new Error(
      `API returned HTTP ${response.status}: ${text}`
    );
  }

  const data =
    await response.json();

  if (!data.success) {
    throw new Error(
      "API returned success=false."
    );
  }

  const result =
    data.result;

  if (!result) {
    throw new Error(
      "API response does not contain result."
    );
  }

  assertEqual(
    result.type,
    test.expected.type.toLowerCase(),
    "Result type"
  );

  if (test.expected.aggregation) {
    assertEqual(
      result.aggregation,
      test.expected.aggregation,
      "Aggregation"
    );
  }

  if (test.expected.value !== undefined) {
    assertNumber(
      result.value,
      test.expected.value,
      "Calculated value"
    );
  }

  if (test.expected.groups) {
    const actualGroups =
      getGroupResults(result);

    assertGroups(
      actualGroups,
      test.expected.groups
    );
  }

  return {
    analysisPlan: data.analysisPlan,
    result
  };
}

async function main() {
  console.log("");
  console.log(
    "=================================================="
  );
  console.log(
    "        INSIGHTMATE FINAL EVALUATION"
  );
  console.log(
    "=================================================="
  );
  console.log("");

  console.log(`Server: ${BASE_URL}`);
  console.log(`Dataset: ${DATASET_PATH}`);
  console.log(`Tests: ${tests.length}`);
  console.log("");

  try {
    const health =
      await checkHealth();

    console.log("✓ HEALTH CHECK PASS");
    console.log(`  Model: ${health.model}`);
    console.log(
      `  Server version: ${health.serverVersion}`
    );
  } catch (error) {
    console.error("");
    console.error("✗ HEALTH CHECK FAILED");
    console.error(`  ${error.message}`);
    console.error("");
    console.error(
      "Make sure server.mjs is running:"
    );
    console.error("node server.mjs");
    console.error("");
    process.exit(1);
  }

  console.log("");

  try {
    const dataset =
      await uploadDataset();

    console.log("✓ DATASET UPLOAD PASS");
    console.log(`  File: ${dataset.fileName}`);
    console.log(`  Rows: ${dataset.rows}`);
    console.log(`  Columns: ${dataset.columns}`);
  } catch (error) {
    console.error("");
    console.error("✗ DATASET UPLOAD FAILED");
    console.error(`  ${error.message}`);
    console.error("");
    process.exit(1);
  }

  console.log("");

  let passed = 0;
  let failed = 0;

  const failures = [];

  for (
    let i = 0;
    i < tests.length;
    i++
  ) {
    const test = tests[i];

    const number =
      String(i + 1).padStart(2, "0");

    process.stdout.write(
      `[${number}/${tests.length}] ${test.name} ... `
    );

    try {
      const output =
        await runTest(test);

      passed++;

      console.log("PASS");

      if (
        output.result.type ===
        "group_by"
      ) {
        const groups =
          getGroupResults(
            output.result
          );

        console.log(
          `      Groups: ${Object.entries(groups)
            .map(
              ([key, value]) =>
                `${key}=${value}`
            )
            .join(" | ")}`
        );
      } else {
        console.log(
          `      Value: ${output.result.value}`
        );
      }
    } catch (error) {
      failed++;

      console.log("FAIL");

      console.log(
        `      ${error.message}`
      );

      failures.push({
        test: test.name,
        question: test.question,
        error: error.message
      });
    }

    console.log("");
  }

  console.log(
    "=================================================="
  );
  console.log(
    "                 FINAL REPORT"
  );
  console.log(
    "=================================================="
  );

  console.log(
    `Total tests : ${tests.length}`
  );

  console.log(
    `Passed      : ${passed}`
  );

  console.log(
    `Failed      : ${failed}`
  );

  console.log(
    `Score       : ${passed}/${tests.length}`
  );

  console.log("");

  if (failed === 0) {
    console.log(
      "FINAL EVALUATION: ALL TESTS PASSED"
    );

    console.log(
      "InsightMate analysis engine passed the complete evaluation suite."
    );

    console.log("");
    process.exit(0);
  }

  console.log(
    "FINAL EVALUATION: FAILURES DETECTED"
  );

  console.log("");

  for (const failure of failures) {
    console.log(`- ${failure.test}`);
    console.log(
      `  Question: ${failure.question}`
    );
    console.log(
      `  Error: ${failure.error}`
    );
    console.log("");
  }

  process.exit(1);
}

main().catch(
  (error) => {
    console.error("");
    console.error(
      "Unexpected evaluation error:"
    );
    console.error(error);
    console.error("");
    process.exit(1);
  }
);