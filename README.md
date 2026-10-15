# InsightMate — AI Data Analyst

**Ask your data. Get clear answers.**

InsightMate is a local AI-powered data analysis application designed to help users explore datasets, understand business performance, and clean Excel workbooks. It combines AI-assisted analysis with deterministic calculations and workbook-processing safeguards.

## Key Features

### 1. AI-Powered Data Analysis
- Ask questions about your uploaded dataset in natural language.
- Get summaries and business insights from your data.
- Use deterministic calculations for supported aggregation and grouping operations.
- View results in tables and charts where appropriate.

### 2. Excel Cleaner
- Upload an Excel workbook and generate a separate cleaned workbook.
- Improve worksheet readability with consistent formatting, alignment, and column widths.
- Remove eligible duplicate records while retaining rows that fail safety checks.
- Remove eligible blank rows when the surrounding data and workbook structure support safe removal.
- Normalize ordinary text and handle recognized Excel error placeholders.
- Preserve worksheets and protect formulas and workbook structures when safe modification cannot be established.

### 3. Local AI Processing
- Connect to Ollama running on the local computer.
- Configure the AI model used by the application.
- Avoid dependence on a paid AI API for the configured local-model workflow.

### 4. Testing and Data Protection
- Automated tests for API behavior and workbook preservation.
- Regression tests for cleaning and safety rules.
- Separate cleaned workbook generation so the source workbook is not intentionally overwritten.
- Conservative handling of formulas, merged cells, defined names, filters, and other workbook structures.

## Technology Stack

- **Backend:** Node.js and Express
- **Frontend:** HTML, CSS, and JavaScript
- **AI:** Ollama with a configurable local language model
- **Excel processing:** ExcelJS and SheetJS
- **Charts:** Chart.js
- **Testing:** Node.js built-in test runner

## How It Works

1. Upload an Excel or CSV dataset for analysis, or upload an Excel workbook for cleaning.
2. Ask a question about the dataset or start the cleaning workflow.
3. For analysis, use the available calculations, tables, charts, and AI-generated explanations.
4. For cleaning, generate a separate workbook with eligible corrections and formatting improvements.
5. Open the resulting workbook in Excel and verify the output.

## Requirements

- Windows, macOS, or Linux with a compatible Node.js installation.
- npm.
- Ollama and a compatible local language model for local AI features.
- A modern web browser.

## Run Locally

Open a terminal in the project folder.

Install the dependencies:

```bash
npm install
```

Start the application:

```bash
node server.mjs
```

Open the application in your browser at:

`http://localhost:3000`

If the application uses a different configured port, use the address shown by the server.

For AI-powered features, ensure Ollama is running and the configured model is available locally.

## Run the Tests

Execute the automated test suite:

```bash
npm test
```

The Excel Cleaner development checkpoint was verified with **48 passing automated tests**, along with real-workbook upload, cleaning, generation, download, and reopen checks.

Additional manual verification covered single-sheet and multi-sheet workbooks.

## Safety Notes and Limitations

- Workbook cleaning is designed to generate a separate output rather than intentionally overwrite the original file.
- Rows or duplicate records may be retained when formula dependencies or workbook structures cannot be verified safely.
- Workbooks with certain defined names or autofilters may receive conservative treatment to protect workbook integrity.
- Always inspect important financial or business workbooks after cleaning.
- Local AI processing depends on the configured local model and available system resources.
- External frontend assets, if loaded from a CDN, may require an internet connection.

## Project Status

The current working version has passed the recorded automated test suite and manual workbook checks. The Excel Cleaner has been tested with both single-sheet and multi-sheet workbooks.

This project is intended as a practical portfolio demonstration of local AI integration, data analysis, Excel automation, and defensive data-processing practices.

## Future Improvements

Potential future improvements include:

- More comprehensive data-quality reporting.
- Additional automated tests with diverse real-world workbooks.
- Clearer documentation of cleaning decisions and retained rows.
- Improved analysis summaries and visualization options.

---

**InsightMate — turning raw data into clearer answers.**