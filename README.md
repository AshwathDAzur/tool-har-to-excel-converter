# UI Performance Analysis Toolkit

A Node.js toolkit for extracting, analyzing, and reporting UI performance metrics from **HAR files** and **Lighthouse audit reports**. Converts raw performance data into professionally formatted, multi-sheet Excel workbooks with statistical baselines, timing breakdowns, and comprehensive helper documentation.

## Overview

This toolkit consists of two complementary tools that together provide a complete picture of UI performance:

| Tool | Input | Measures | Output |
|------|-------|----------|--------|
| **HAR Parser** (`har-to-excel.js`) | HAR files from browser DevTools | Network-level performance (API response times, TTFB, transfer sizes, timing breakdown) | `output.xlsx` |
| **Lighthouse Parser** (`lighthouseParcer.js`) | Lighthouse JSON exports | Rendering performance (Core Web Vitals, main thread activity, JS execution, resource efficiency) | `lighthouse-output.xlsx` |

Together they answer:
- **HAR**: How fast are individual network requests? Where is time spent in the request lifecycle?
- **Lighthouse**: How fast does the page render, become interactive, and stabilize visually?

## Features

### HAR Parser (`har-to-excel.js`)

- **Auto-Discovery**: Scans `harRepo/` directory and processes all `.har` files
- **One Sheet Per HAR File**: Each file creates a dedicated worksheet (sheet name from filename, truncated to 31 chars)
- **14 Data Columns Per Request**:
  - #, Method, Endpoint URL, Category, Status, Response (ms), Size (KB)
  - Timing Breakdown: Blocked, DNS, Connect, SSL, Send, Wait/TTFB, Receive (all in ms)
- **Automatic URL Categorization**: Extracts meaningful path segments, filters common prefixes (`/api`, `/v1`, `/rest`, etc.), identifies static assets
- **Observed Baseline Statistics**: Total Requests, Average, Median, Min, Max, P90, P95, Total Transfer Size (for both Response time and TTFB)
- **Page Load Timings**: DOMContentLoaded and Page Load (onLoad) per page
- **Helper Sheet**: Documents all metrics with descriptions and units

### Lighthouse Parser (`lighthouseParcer.js`)

- **Auto-Discovery**: Scans `lighthouseRepo/` directory and processes all `.json` files
- **One Sheet Per JSON File**: Each Lighthouse export creates a dedicated worksheet
- **8 Sections Per Page**:
  1. **Audit Information** — URL, fetch time, Lighthouse version, gather mode, benchmark index
  2. **Category Scores** — Performance, Accessibility, Best Practices (color-coded: green/orange/red)
  3. **Performance Metrics** — FCP, LCP, TBT, CLS, Speed Index, TTI, Max Potential FID (with raw values, scores, and ratings)
  4. **Server & Network** — Server Response Time (TTFB), Total Byte Weight
  5. **Main Thread Breakdown** — Script Evaluation, Parsing, Style & Layout, HTML/CSS Parsing, Rendering, GC (with totals)
  6. **JavaScript Execution (Top 10)** — Heaviest scripts by CPU time, scripting time, parse/compile time
  7. **Resource Summary** — Breakdown by type (Script, Stylesheet, Image, Document, etc.) with request count and transfer size
  8. **Diagnostics & Opportunities** — Unused JS/CSS, render-blocking resources, missing compression, with potential savings
- **Score Color Coding**: Green (90-100 Good), Orange (50-89 Needs Work), Red (0-49 Poor)
- **Helper Sheet**: Documents all Lighthouse metrics, scoring thresholds, and interpretation guidance

### Shared Styling

Both tools use consistent professional Excel styling:
- **Header row**: Dark blue background (#394B67) with bold white text
- **Section headers**: Dark navy (#2C3E50) with white text
- **Sub-headers**: Blue-grey (#394B67) with white text
- **Data rows**: Zebra striping with light grey (#F2F4F7) for even rows
- **URLs/Metric names**: Consolas font for monospace readability
- **All other text**: Arial 10pt
- **Thin borders** around all cells for structure

## Prerequisites

- **Node.js**: Version 12.0 or higher (14+ recommended)
- **npm**: Included with Node.js

## Installation

1. Navigate to the project directory (where `package.json` is located):
   ```bash
   cd /path/to/project
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

   This installs the following packages:
   - `exceljs`: Excel workbook creation, styling, and cell manipulation
   - `fs-extra`: Enhanced file system operations

## Usage

### HAR Parser

1. **Capture HAR files** from browser DevTools (see [How to Capture HAR Files](#how-to-capture-har-files))
2. **Place `.har` files** in the `harRepo/` directory
3. **Run**:
   ```bash
   node har-to-excel.js
   ```
4. **Output**: `output.xlsx` in the project root

```
# Example output:
Found 2 HAR file(s) in /path/to/project/harRepo

Reading HAR file: /path/to/project/harRepo/Landing.har
Parsed 87 entries from Landing.har
Added sheet: "Landing" with 87 rows

Reading HAR file: /path/to/project/harRepo/WellProgram.har
Parsed 87 entries from WellProgram.har
Added sheet: "WellProgram" with 87 rows

Added sheet: "Helper"
Excel file saved: output.xlsx
Conversion completed successfully!
```

### Lighthouse Parser

1. **Run Lighthouse audits** from Chrome DevTools (see [How to Capture Lighthouse Reports](#how-to-capture-lighthouse-reports))
2. **Place `.json` exports** in the `lighthouseRepo/` directory
3. **Run**:
   ```bash
   node lighthouseParcer.js
   ```
4. **Output**: `lighthouse-output.xlsx` in the project root

```
# Example output:
Found 2 Lighthouse JSON file(s) in /path/to/project/lighthouseRepo

Reading Lighthouse JSON: /path/to/project/lighthouseRepo/landing.json
Parsed metrics for: https://your-app-url/landing-page
Added sheet: "landing"

Reading Lighthouse JSON: /path/to/project/lighthouseRepo/wellprogram.json
Parsed metrics for: https://your-app-url/well-program
Added sheet: "wellprogram"

Added sheet: "Helper"
Excel file saved: lighthouse-output.xlsx
Conversion completed successfully!
```

## How to Capture HAR Files

### Chrome / Edge / Chromium Browsers

1. Open DevTools (`F12` or `Ctrl + Shift + I`)
2. Navigate to the **Network** tab
3. Enable **Preserve Logs** and **Disable Cache** for accurate captures
4. Load the page and wait until it becomes fully static/idle
5. Right-click in the Network table and select **Save all as HAR with content**
6. Save with a descriptive name (e.g., `landing.har`) and move to `harRepo/`

### Firefox

1. Open DevTools (`F12`)
2. Navigate to the **Network** tab
3. Load the page and wait for completion
4. Right-click on any entry and select **Save All As HAR**
5. Move the file to `harRepo/`

### Safari

1. Enable Develop Menu: Preferences > Advanced > "Show Develop menu in menu bar"
2. Open Web Inspector via Develop menu
3. Navigate to the **Network** tab, reload the page
4. Right-click and select **Export HAR**
5. Move the file to `harRepo/`

### Best Practices for HAR Capture

- Enable **Preserve Logs** to retain entries across navigations
- Enable **Disable Cache** to simulate first-time user experience
- Wait until the page is **fully static/idle** before exporting
- Perform **3-5 runs per page** with cache cleared between iterations to account for DNS, network jitter, and cold cache variability
- Repeat captures **periodically** to establish trend data and detect regressions

## How to Capture Lighthouse Reports

### Chrome DevTools (Recommended)

1. Open DevTools (`F12` or `Ctrl + Shift + I`)
2. Navigate to the **Lighthouse** tab
3. Configure the audit:
   - **Mode**: Navigation
   - **Device**: Desktop (or Mobile depending on target)
   - **Categories**: Check Performance, Accessibility, Best Practices
4. Click **Analyze page load** and wait for completion (do not interact with the page)
5. Click the **export icon** (top-right of report) and select **Save as JSON**
6. Save with a descriptive name (e.g., `landing.json`) and move to `lighthouseRepo/`

### CLI (For Automation)

```bash
npx lighthouse https://your-app-url --output=json --output-path=./lighthouseRepo/landing.json --preset=desktop
```

### Best Practices for Lighthouse Capture

- Run **3-5 audits per page** — Lighthouse scores fluctuate due to CPU load and network conditions
- **Close other tabs** and background applications to reduce noise
- Use consistent **device and throttling settings** across runs
- Export as **JSON** (not HTML) for machine-readable processing

## Output Structure

### HAR Output (`output.xlsx`)

Each sheet contains three sections:

#### 1. Request Data Table (14 columns)

| Column | Description |
|--------|-------------|
| # | Sequential request number |
| Method | HTTP method (GET, POST, etc.) |
| Endpoint URL | Full request URL |
| Category | Auto-derived from URL path |
| Status | HTTP response status code |
| Response (ms) | Total round-trip time |
| Size (KB) | Response body size |
| Blocked (ms) | Time waiting for network slot |
| DNS (ms) | Domain resolution time |
| Connect (ms) | TCP connection time |
| SSL (ms) | TLS/SSL handshake time |
| Send (ms) | Request transmission time |
| Wait/TTFB (ms) | Server processing time |
| Receive (ms) | Response download time |

#### 2. Observed Baseline

| Metric | Description |
|--------|-------------|
| Total Requests | Count of HTTP requests |
| Average | Mean response time and size |
| Median | 50th percentile (typical performance) |
| Min | Fastest observed value |
| Max | Slowest observed value |
| P90 | 90th percentile |
| P95 | 95th percentile (industry-standard SLA metric) |
| Total Transfer Size | Aggregate data downloaded |

#### 3. Page Load Timings

| Metric | Description |
|--------|-------------|
| DOMContentLoaded (ms) | Time until DOM is ready |
| Page Load (ms) | Time until all resources loaded |

### Lighthouse Output (`lighthouse-output.xlsx`)

Each sheet contains eight sections:

#### 1. Audit Information
URL, fetch timestamp, Lighthouse version, gather mode, benchmark index.

#### 2. Category Scores
Overall scores (0-100) for Performance, Accessibility, Best Practices — color-coded by rating.

#### 3. Performance Metrics (Core Web Vitals)

| Metric | Unit | Good | Poor |
|--------|------|------|------|
| First Contentful Paint (FCP) | ms | < 1.8s | > 3.0s |
| Largest Contentful Paint (LCP) | ms | < 2.5s | > 4.0s |
| Total Blocking Time (TBT) | ms | < 200ms | > 600ms |
| Cumulative Layout Shift (CLS) | unitless | < 0.1 | > 0.25 |
| Speed Index (SI) | ms | < 3.4s | > 5.8s |
| Time to Interactive (TTI) | ms | — | — |
| Max Potential FID | ms | — | — |

#### 4. Server & Network
Server response time (TTFB) and total byte weight.

#### 5. Main Thread Breakdown
Time distribution across: Script Evaluation, Script Parsing & Compilation, Style & Layout, Parse HTML & CSS, Rendering, Garbage Collection, Other.

#### 6. JavaScript Execution (Top 10 Scripts)
Heaviest scripts ranked by total CPU time, with scripting and parse/compile breakdown.

#### 7. Resource Summary
Request count and transfer size grouped by resource type (Script, Stylesheet, Image, Document, etc.).

#### 8. Diagnostics & Opportunities
Flagged audits (unused JS/CSS, render-blocking resources, missing compression) with scores and potential savings.

### Helper Sheets

Both tools include a dedicated **Helper** worksheet documenting every metric with descriptions, units, and interpretation guidance.

## Performance Coverage Matrix

| Performance Layer | HAR Parser | Lighthouse Parser |
|-------------------|:----------:|:-----------------:|
| API Response Times | Yes | — |
| Request Timing Breakdown | Yes | — |
| Transfer Size per Request | Yes | — |
| Statistical Baselines (P90/P95) | Yes | — |
| Core Web Vitals (FCP, LCP, CLS) | — | Yes |
| Main Thread Activity | — | Yes |
| JS Execution Profiling | — | Yes |
| Resource Efficiency | — | Yes |
| Diagnostics & Opportunities | — | Yes |
| Page Load Timings | Yes | — |
| Server Response Time | Yes | Yes |
| Category Scores | — | Yes |

## Project Structure

```
project-root/
├── har-to-excel.js             # HAR parser (HARtoExcelConverter class)
├── lighthouseParcer.js          # Lighthouse parser (LighthouseToExcelConverter class)
├── package.json                 # Project metadata and dependencies
├── package-lock.json            # Locked dependency versions
├── README.md                    # This file
├── harRepo/                     # Input: place .har files here
│   ├── Landing.har
│   └── WellProgram.har
├── lighthouseRepo/              # Input: place Lighthouse .json exports here
│   ├── landing.json
│   └── wellprogram.json
├── output.xlsx                  # Generated HAR analysis workbook
├── lighthouse-output.xlsx       # Generated Lighthouse analysis workbook
└── node_modules/                # Installed dependencies (auto-generated)
```

## Tech Stack

| Package | Purpose |
|---------|---------|
| exceljs | Excel workbook creation, styling, and cell manipulation |
| fs-extra | Enhanced file system operations |

**Runtime**: Node.js 12.0+ (14+ recommended), Windows / macOS / Linux

## Architecture

### HAR Parser (`HARtoExcelConverter`)
| Method | Description |
|--------|-------------|
| `processHARFile(filePath)` | Parses HAR JSON, extracts 14 columns per request plus page timings |
| `categorize(url)` | Auto-categorizes URLs by path segment extraction |
| `addSheet(sheetName, rows, pageTimings)` | Creates styled worksheet with data table, baseline, and page timings |
| `addHelperSheet()` | Creates reference documentation worksheet |
| `median(arr)` / `percentile(arr, p)` | Statistical calculation helpers |
| `save(outputPath)` | Writes workbook to disk |

### Lighthouse Parser (`LighthouseToExcelConverter`)
| Method | Description |
|--------|-------------|
| `processLighthouseFile(filePath)` | Parses Lighthouse JSON, extracts all metrics across 8 categories |
| `getScoreStyle(score)` / `getScoreLabel(score)` | Color-code and label scores (Good/Needs Work/Poor) |
| `addSectionTitle(ws, title, totalCols)` | Adds styled section header row |
| `addSubHeader(ws, headers)` | Adds styled sub-header row |
| `addSheet(sheetName, report)` | Creates full worksheet with all 8 sections |
| `addHelperSheet()` | Creates Lighthouse metric reference worksheet |
| `save(outputPath)` | Writes workbook to disk |

## Baseline Metrics Collection Process

- Enable **Preserve Logs** and **Disable Cache** in browser DevTools to ensure complete, uncached network capture.
- Record the Networks tab activity and export the HAR file once the page reaches a fully static/idle state.
- Run **Lighthouse audits** from the Lighthouse tab with Navigation mode and Desktop preset; export as JSON once the report is generated.
- Perform **3-5 separate runs per page** with cache cleared between each iteration to minimize skew from cold cache, DNS resolution, and network jitter.
- Repeat captures **periodically over time** to establish trend data and identify performance regressions.
- Feed collected HAR files and Lighthouse JSON exports into the respective parser tools to extract metrics and compute baselines.

## Interpretation Tips

- **Focus on P90/P95, not Average**: Percentile metrics represent real user experience; averages are skewed by outliers
- **TTFB is Key**: High Wait/TTFB indicates server-side bottlenecks (database queries, computation, network latency)
- **LCP drives user perception**: If users feel the page is "slow", LCP is usually the culprit
- **TBT predicts interactivity**: High TBT means the page looks loaded but doesn't respond to clicks
- **Main Thread Breakdown reveals root cause**: Script Evaluation dominating? Optimize bundle size. Style & Layout? Reduce DOM complexity
- **Resource Summary identifies waste**: Large script bundles, unused CSS, uncompressed assets are quick wins
- **Compare categories separately**: API endpoints and static assets have different performance profiles

## Common Issues and Troubleshooting

### No files found

**Problem**: "No .har files found" or "No .json files found"

**Solution**:
1. Verify the `harRepo/` or `lighthouseRepo/` directory exists in the project root
2. Check that files have the correct extension (`.har` or `.json`)
3. Ensure files are placed directly in the directory (not in subdirectories)

### Parse error

**Problem**: "SyntaxError: Unexpected token..." or "Invalid JSON"

**Solution**:
1. Verify the file is valid JSON
2. For HAR: re-export from DevTools using "Save all as HAR with content"
3. For Lighthouse: re-export using "Save as JSON" from the export menu

### Excel file is locked

**Problem**: "Error: Cannot write file - file is locked"

**Solution**:
1. Close the output file if open in Excel
2. Ensure no other process is using the file
3. Rename the existing file and run again

### Out of memory

**Problem**: "JavaScript heap out of memory"

**Solution**:
1. Increase heap: `node --max-old-space-size=4096 har-to-excel.js`
2. Split large files into smaller chunks
3. Upgrade to Node.js 16+ for improved memory management

## FAQs

### Can I use HAR files from any browser?
Yes. HAR is a standardized format (HAR 1.2 specification), so files from Chrome, Firefox, Safari, or any compliant tool will work.

### Can I use Lighthouse JSON from the CLI?
Yes. JSON exports from `npx lighthouse --output=json` have the same structure as DevTools exports.

### How often should I capture baselines?
Before each major release, after infrastructure changes, or monthly for continuous monitoring. More frequent captures provide better trend data.

### Can I edit the Excel output?
Yes. Both output files are standard `.xlsx` files. However, re-running the tool overwrites the file.

### Why are my Lighthouse scores different each run?
Lighthouse scores fluctuate due to CPU load, network conditions, and background processes. Run 3-5 times and compare medians for reliable baselines.

### How does URL categorization work in the HAR parser?
1. Identifies static assets by file extension (.js, .css, .png, etc.) as "Static Asset"
2. Extracts path segments, skipping common prefixes (/api, /v1, /v2, /rest, /web, /app, etc.)
3. Uses the first meaningful segment as category (e.g., `/api/v1/users/123` becomes "Users")
4. Converts to Title Case for readability
5. Defaults to "General" if no meaningful segment found

---

**Last Updated**: February 2026
**Version**: 2.0.0
**Toolkit**: UI Performance Analysis Toolkit
