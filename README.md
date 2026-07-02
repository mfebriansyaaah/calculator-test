# Floor Panel Properties — Web Calculator

This project is a Vite (Vanilla JS) web calculator that mirrors the calculations from the provided Excel workbook.

The UI renders the same output blocks as the workbook, and the calculation uses the workbook as the source of truth.

## Prerequisites

- Node.js (LTS recommended)
- Yarn (recommended)

## Step-by-step: Run Locally

1. Clone the repository

   ```bash
   git clone <YOUR_REPO_URL>
   cd calculator-test
   ```

2. Install dependencies

   ```bash
   yarn install
   ```

3. Start the dev server

   ```bash
   yarn dev
   ```

4. Open the app in your browser

   Vite will print a local URL in the terminal (for example `http://127.0.0.1:5173/`).

## Step-by-step: Test the Calculator

### 1) Verify input constraints

- Analytical Method = **Shear Analogy**
  - Supported layers: **3–9**
  - Must be **symmetric from top to bottom**
- Analytical Method = **Gamma**
  - Supported layers: **3 or 5 only**

If an invalid configuration is selected, the app will show an error message instead of calculating.

### 2) Verify hide/show behavior

- When **Shear Analogy** is selected, the UI shows:
  - **SECTION PROPERTIES (T1–T7)**
  - **Shear Analogy Method (T1–T7 + SEiIi XX)**
  - The Gamma block is hidden
- When **Gamma** is selected, the UI shows:
  - **SECTION PROPERTIES (T1–T7)**
  - **Gamma Method (full table)**
  - The Shear Analogy block is hidden

### 3) Quick sanity test (known input)

Use the following input:

- Grade: `MGP10`
- Total Layers: `3`
- Thickness Each Layer (mm): `35`
- Length (m): `5`
- b_eff (mm): `1000`
- Analytical Method: `Shear Analogy`

Expected: the **EI eff** value should match the workbook output for the same inputs.

## How to Compare With the Excel Workbook (Repo Version)

The workbook used by the app is stored at:

- `public/floor-panel-properties.xlsx`

To compare manually:

1. Open `public/floor-panel-properties.xlsx` in Excel
2. Enter the same inputs in the workbook input cells (Grade, layers, thickness, length, b_eff, method)
3. Compare the three output blocks displayed in the web UI against the corresponding blocks in the workbook

## Build & Preview (Production)

```bash
yarn build
yarn preview
```

## Troubleshooting

### The app does not load the workbook

- Ensure the file exists at `public/floor-panel-properties.xlsx`.
- If you renamed or moved the workbook, keep the same path and filename or update the fetch path in `calculator.js`.

### The dev server starts on a different port

Vite will automatically pick another port if the default is already in use. Always use the URL printed in the terminal (for example `http://127.0.0.1:5174/`).

### The output values do not match Excel

- Make sure you are comparing against the same workbook file: `public/floor-panel-properties.xlsx`.
- Enter the exact same input values (grade, layer count, thickness, length, b_eff, and method).
- If Excel shows `#REF!` / `#VALUE!`, it usually indicates a broken workbook, missing links, or calculation settings in that Excel file. The recommended comparison target is the workbook version stored in this repository.

### I cannot install dependencies

- Verify Node.js is installed: `node -v`
- Verify Yarn is installed: `yarn -v`
- If you use npm instead of Yarn:

  ```bash
  npm install
  npm run dev
  ```

## Notes

- Do not commit `node_modules/` or `dist/` (standard frontend practice).
- If you replace the workbook file, keep the same filename and location: `public/floor-panel-properties.xlsx`.
