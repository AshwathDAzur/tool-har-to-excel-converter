# HAR-to-Excel Performance Analysis Tool

A powerful Node.js utility for converting HTTP Archive (HAR) files into professionally formatted, multi-sheet Excel workbooks. This tool is designed to streamline API performance baselining and analysis by automatically organizing network data into clear, actionable spreadsheets with comprehensive timing breakdowns and statistical baselines.

## Overview

The HAR-to-Excel Performance Analysis Tool processes raw HAR (HTTP Archive) format files—JSON records of web browser interactions with web sites—and transforms them into beautifully formatted Excel spreadsheets. Each HAR file becomes a dedicated sheet in the output workbook, with automatic URL categorization, detailed timing analysis, statistical baselines, and page load metrics.

This tool is essential for performance engineers, QA teams, and DevOps professionals who need to establish performance baselines, compare performance metrics across environments, and track optimization efforts over time.

## Features

- **Auto-Discovery of HAR Files**: Automatically scans the `harRepo/` directory and processes all `.har` files without manual configuration
- **One Sheet Per HAR File**: Each HAR file creates a dedicated worksheet, with the sheet name derived from the filename (truncated to 31 characters to meet Excel limits)
- **Comprehensive Data Columns**: Each request is logged with:
  - Request number (#)
  - HTTP method (GET, POST, PUT, DELETE, etc.)
  - Full endpoint URL with protocol, domain, path, and query parameters
  - Auto-categorized URL category based on path segments
  - HTTP response status code
  - Total response time in milliseconds
  - Response body size in kilobytes
  - Complete timing breakdown (13 detailed metrics)

- **Detailed Timing Breakdown**: Separates request time into distinct phases:
  - Blocked (ms): Time waiting for network connection availability
  - DNS (ms): Domain name resolution time
  - Connect (ms): TCP connection establishment time
  - SSL (ms): TLS/SSL handshake time
  - Send (ms): Time to transmit HTTP request
  - Wait/TTFB (ms): Time to First Byte; server processing time
  - Receive (ms): Time to download response body

- **Automatic URL Categorization**: Intelligently categorizes URLs by extracting meaningful path segments while filtering out common prefixes like `/api`, `/v1`, `/v2`, `/rest`, `/web`, `/app`, and others. Static assets (JS, CSS, images, fonts) are automatically identified as 'Static Asset'

- **Observed Baseline Statistics**: Automatically calculated for each sheet:
  - Total Requests: Number of HTTP requests captured
  - Average: Arithmetic mean of response times and sizes
  - Median: Middle value (50th percentile) for typical performance representation
  - Min: Best-case (fastest) performance
  - Max: Worst-case (slowest) performance
  - 95th Percentile (P95): Performance threshold capturing 95% of requests
  - Total Transfer Size: Sum of all response sizes

- **Page Load Timings**: Captures browser navigation and page load metrics:
  - DOMContentLoaded (ms): Time until HTML is parsed and DOM is ready
  - Page Load/onLoad (ms): Time until all resources are fully loaded

- **Helper Reference Sheet**: Dedicated "Helper" worksheet documenting all column headers, metrics, and their descriptions with proper units

- **Professional Excel Styling**:
  - Header row: Dark blue background (#394B67) with white text
  - Section headers: Dark gray background (#2C3E50) with white text
  - Data rows: Zebra striping with light gray background (#F2F4F7) for even rows
  - Endpoint URLs: Consolas font (8.5pt) for easy URL parsing
  - All other text: Arial font (10pt) for clarity
  - Thin borders around all cells for structure and readability
  - Properly sized columns with text wrapping on long URLs

## Prerequisites

- **Node.js**: Version 12.0 or higher
- **npm**: Included with Node.js

## Installation

1. Clone or download the project to your local machine

2. Navigate to the project directory:
   ```bash
   cd c:\OrgProjects\Halliburton
   ```

3. Install dependencies:
   ```bash
   npm install
   ```

   This will install the following packages:
   - `exceljs` (4.4.0): Excel workbook creation and manipulation
   - `fs-extra` (11.3.3): Enhanced file system operations
   - `xlsx` (0.18.5): Additional spreadsheet functionality

## Usage

### Basic Workflow

1. **Prepare HAR Files**: Place all `.har` files you want to analyze in the `harRepo/` directory within the project root

2. **Run the Tool**:
   ```bash
   node har-to-excel.js
   ```

3. **Review Output**: Open the generated `output.xlsx` file in Microsoft Excel or compatible spreadsheet applications

### Example

```bash
# From the project root directory
node har-to-excel.js

# Expected output:
# Found 3 HAR file(s) in c:\OrgProjects\Halliburton\harRepo
#
# Reading HAR file: c:\OrgProjects\Halliburton\harRepo\homepage.har
# Parsed 42 entries from homepage.har
# Added sheet: "homepage" with 42 rows
#
# Reading HAR file: c:\OrgProjects\Halliburton\harRepo\api-test.har
# Parsed 128 entries from api-test.har
# Added sheet: "api-test" with 128 rows
#
# ...
#
# Added sheet: "Helper"
# Excel file saved: output.xlsx
# Conversion completed successfully!
```

### Output File

The tool generates `output.xlsx` in the project root directory. This Excel workbook contains:
- One worksheet per HAR file (named after the HAR filename)
- One "Helper" worksheet with metric reference documentation

The output file is ready for immediate use—open it in Excel, Google Sheets, LibreOffice Calc, or any compatible spreadsheet application.

## How to Capture HAR Files

To collect HAR files from your browser sessions, follow these steps:

### Chrome / Edge / Chromium Browsers

1. **Open DevTools**: Press `F12` or right-click on the page and select "Inspect"

2. **Navigate to Network Tab**: Click the "Network" tab in the DevTools panel

3. **Load Your Page**: If the Network tab wasn't open when you started, reload the page (Ctrl+R or Cmd+R) to capture network activity

4. **Export HAR**:
   - Right-click anywhere in the Network activity table
   - Select "Save all as HAR with content" (or similar, depending on browser version)
   - Choose a save location and filename (e.g., `homepage.har`)

5. **Move to harRepo**: Move or copy the `.har` file to the `harRepo/` directory

### Firefox

1. **Open DevTools**: Press `F12` or right-click on the page and select "Inspect"

2. **Navigate to Network Tab**: Click the "Network" tab

3. **Load Your Page**: Reload the page if needed to capture activity

4. **Export HAR**:
   - Right-click on any entry in the Network table
   - Select "Save All As HAR"
   - Choose a location and filename

5. **Move to harRepo**: Move or copy the `.har` file to the `harRepo/` directory

### Safari

1. **Enable Develop Menu**: Go to Preferences > Advanced and check "Show Develop menu in menu bar"

2. **Open Develop Menu**: Click "Develop" in the menu bar and select "Show Web Inspector"

3. **Network Tab**: Click the "Network" tab

4. **Load Your Page**: Reload the page to capture activity

5. **Export HAR**:
   - Right-click in the Network panel
   - Select "Export HAR"
   - Choose a location and filename

6. **Move to harRepo**: Move or copy the `.har` file to the `harRepo/` directory

## Output Structure

### Data Sheets (One Per HAR File)

Each sheet corresponding to a HAR file contains three main sections:

#### 1. Request Data Table
A detailed table of all HTTP requests with the following columns:

| Column | Description |
|--------|-------------|
| # | Sequential request number |
| Method | HTTP method (GET, POST, etc.) |
| Endpoint URL | Full request URL |
| Category | Auto-derived category from URL path |
| Status | HTTP response status code |
| Response (ms) | Total request time in milliseconds |
| Size (KB) | Response body size in kilobytes |
| Blocked (ms) | Time waiting for network slot |
| DNS (ms) | Domain resolution time |
| Connect (ms) | TCP connection time |
| SSL (ms) | TLS/SSL handshake time |
| Send (ms) | HTTP request transmission time |
| Wait/TTFB (ms) | Time to first byte / server processing time |
| Receive (ms) | Response download time |

#### 2. Observed Baseline Section
Statistical summary metrics calculated from all requests:

- **Total Requests**: Count of all HTTP requests captured
- **Average**: Mean value across all requests
- **Median**: Middle value (50th percentile); often more representative than average
- **Min**: Minimum/fastest observed value
- **Max**: Maximum/slowest observed value
- **95th Percentile**: Threshold capturing performance for 95% of requests
- **Total Transfer Size**: Aggregate downloaded data in KB

#### 3. Page Load Timings Section (If Available)
Browser navigation metrics extracted from the HAR file:

- **Page Title**: The page being loaded
- **DOMContentLoaded (ms)**: Time until HTML parsing and DOM readiness
- **Page Load (ms)**: Time until all resources fully loaded (onLoad event)

### Helper Sheet

The "Helper" worksheet provides comprehensive documentation of all metrics:

- **Data Columns Section**: Description of each column in the request table
- **Timing Breakdown Section**: Detailed explanation of each timing phase
- **Observed Baseline Metrics Section**: Statistical measures and their use cases
- **Page Load Timings Section**: Navigation metrics and their significance

This sheet is essential for new team members and serves as a reference for interpreting baseline metrics.

## Column Reference

### Main Data Columns

| Column | Type | Units | Description |
|--------|------|-------|-------------|
| # | Integer | N/A | Sequential row number for each network request |
| Method | String | N/A | HTTP method (GET, POST, PUT, DELETE, PATCH, OPTIONS, HEAD, etc.) |
| Endpoint URL | String | N/A | Complete URL including protocol, domain, path, and query parameters |
| Category | String | N/A | Auto-derived from URL path; helps group related endpoints |
| Status | Integer | HTTP Code | HTTP response status code (200, 301, 404, 500, etc.) |
| Response (ms) | Decimal | Milliseconds | Total round-trip time for the complete request/response cycle |
| Size (KB) | Decimal | Kilobytes | Uncompressed response body size |

### Timing Breakdown Columns

| Column | Type | Units | Description |
|--------|------|-------|-------------|
| Blocked (ms) | Decimal | Milliseconds | Time request waited in browser queue for available connection |
| DNS (ms) | Decimal | Milliseconds | Time to resolve domain name to IP address (0 if cached) |
| Connect (ms) | Decimal | Milliseconds | Time to establish TCP connection to server (0 if reused) |
| SSL (ms) | Decimal | Milliseconds | Time for TLS/SSL handshake (0 if HTTP or connection reused) |
| Send (ms) | Decimal | Milliseconds | Time to transmit HTTP request headers and body to server |
| Wait/TTFB (ms) | Decimal | Milliseconds | Time waiting for server response; indicates server-side processing time |
| Receive (ms) | Decimal | Milliseconds | Time to download complete response body from server |

**Note**: Sum of all timing phases approximates the total Response (ms), though there may be minor variations due to rounding and parallel processing of some phases.

## Observed Baseline Metrics

Understanding baseline metrics is crucial for performance analysis and improvement tracking.

### Total Requests
The total number of HTTP requests captured in the HAR file for a given page or flow. This helps establish the scope of analysis and can highlight unexpectedly large request volumes indicating potential optimization opportunities.

### Average (Mean)
The arithmetic mean of all values (response times, sizes, or TTFB). Useful as a general indicator but can be skewed by extreme outliers. For performance analysis, compare average with median to identify outlier impact.

**Use Case**: General performance indicator; good for trend tracking

### Median (50th Percentile)
The middle value when all measurements are sorted. More representative of typical user experience than average because outliers have less influence. If average and median differ significantly, investigate outliers.

**Use Case**: Better representation of typical performance than average

### Min (Minimum)
The fastest or smallest observed value. Represents best-case performance under ideal conditions. Generally less relevant than median or P95 for baseline establishment.

**Use Case**: Reference point for best-case scenarios; identify potential optimization targets

### Max (Maximum)
The slowest or largest observed value. Represents worst-case performance. Can highlight performance problems, but single outliers shouldn't drive decisions—investigate P95 instead.

**Use Case**: Identify performance problems and anomalies; track regressions

### 95th Percentile (P95)
The value below which 95% of observations fall. This is the most important metric for performance baselining. Indicates real user experience for the vast majority while ignoring extreme outliers. Industry standard for SLA definitions.

**Use Case**: Performance SLAs; primary metric for performance baselines and comparisons

### Total Transfer Size (KB)
The sum of all response body sizes in kilobytes. Helps assess overall bandwidth consumption and identifies pages that download excessive data. Consider comparing across different network conditions or optimizing large resources.

**Use Case**: Network optimization; compare across device classes (mobile vs. desktop)

## Best Practices for Performance Baselining

Effective performance baselining requires consistent methodology to ensure meaningful comparisons over time and across environments.

### Recording Best Practices

#### Multiple Captures
- Record each page/flow **3-5 times** in sequence to capture natural variation
- Average results across captures for more reliable baselines
- This accounts for temporary network fluctuations and transient delays

#### Cold Cache vs. Warm Cache
- **Cold Cache**: Clear browser cache before recording to simulate first-time visitors (Ctrl+Shift+Delete or browser settings)
- **Warm Cache**: Record multiple times in succession for cached resource performance
- Document which approach you used and maintain consistency

#### Network Conditions
- Record under consistent network conditions (wired connection, no heavy traffic)
- Use browser throttling if testing specific network conditions (3G, 4G, etc.)
- Document network conditions used for future comparison

#### Page Load Completeness
- Ensure all page interactions are complete before stopping the HAR capture
- Wait for dynamic content loading if using SPAs (Single Page Applications)
- Confirm page load progress indicators show 100% completion

### Baseline Comparison

#### Environment Comparison
- Capture HAR files from different environments (dev, staging, production)
- Compare P95 response times and TTFB metrics across environments
- Investigate significant differences (>5-10% variation)

#### Pre/Post Release Comparison
- Establish baselines before deploying changes
- Capture HAR files immediately after deployment
- Compare key metrics (P95, Total Transfer Size, TTFB)
- Flag regressions (increases) in P95 metrics immediately

#### Category Analysis
- Focus on API endpoints (Category != 'Static Asset') for server performance
- Track Static Asset performance separately for CDN/caching effectiveness
- Identify slow categories and investigate specific endpoints

#### Time-Series Tracking
- Maintain a spreadsheet or database of baseline metrics over time
- Track P95 and median for each major page/flow
- Create graphs to visualize performance trends
- Set alerts for regressions exceeding 5-10% thresholds

### Interpretation Tips

- **Focus on P95, not Average**: P95 represents performance for 95% of users; average is skewed by outliers
- **TTFB is Key**: Wait/TTFB (ms) indicates server-side performance; high values suggest backend issues
- **Size Matters**: Total Transfer Size directly impacts load time on slow networks; prioritize large resources
- **Blocked Time**: High blocked values suggest browser connection limits (typically 6 concurrent connections)
- **Compare Categories Separately**: API endpoints and static assets have different performance characteristics

### Tools for Tracking

- Maintain a log of baseline HAR captures with timestamps and environment info
- Version control baseline Excel files to track historical changes
- Use conditional formatting to highlight regressions in baseline metrics
- Create summary dashboards comparing key metrics across releases

## Project Structure

```
c:\OrgProjects\Halliburton\
├── har-to-excel.js          # Main converter script (HARtoExcelConverter class)
├── package.json             # Project metadata and dependencies
├── package-lock.json        # Locked dependency versions (do not edit)
├── README.md                # This file - project documentation
├── harRepo/                 # Input directory - place .har files here
│   ├── homepage.har
│   ├── api-test.har
│   └── checkout.har
├── node_modules/            # Installed dependencies (auto-generated)
│   ├── exceljs/
│   ├── fs-extra/
│   └── xlsx/
└── output.xlsx              # Generated Excel workbook (overwritten on each run)
```

### Directory Descriptions

- **har-to-excel.js**: The main application script. Contains the `HARtoExcelConverter` class which handles HAR parsing, categorization, statistical calculations, Excel generation, and styling.

- **harRepo/**: Input directory where all source HAR files should be placed. The tool automatically discovers and processes all `.har` files in this directory.

- **output.xlsx**: The generated Excel workbook containing processed HAR data. Overwritten each time the tool runs, so save previous versions if needed.

## Tech Stack

### Core Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| exceljs | 4.4.0 | Excel workbook creation, styling, and cell manipulation |
| fs-extra | 11.3.3 | Enhanced file system operations (cross-platform compatibility) |
| xlsx | 0.18.5 | Additional spreadsheet processing functionality |

### Runtime

- **Node.js**: 12.0+ (recommend 14+ for best compatibility)
- **Platform**: Windows, macOS, Linux

### Architecture

The tool uses a class-based architecture:

- **HARtoExcelConverter**: Main class managing the conversion process
  - `processHARFile(filePath)`: Parses HAR JSON and extracts request data
  - `categorize(url)`: Auto-categorizes URLs based on path segments
  - `addSheet(sheetName, rows, pageTimings)`: Creates styled worksheet with data
  - `addHelperSheet()`: Creates reference documentation worksheet
  - `save(outputPath)`: Writes workbook to disk
  - `median(arr)` / `percentile(arr, p)`: Statistical calculation helpers

## Common Issues and Troubleshooting

### No .har files found

**Problem**: "No .har files found in harRepo/ directory."

**Solution**:
1. Verify the `harRepo/` directory exists in the project root
2. Check that .har files are in the correct directory (case-sensitive on some systems)
3. Ensure filenames end with `.har` (not `.HAR` or `.txt`)

### HAR file parse error

**Problem**: "SyntaxError: Unexpected token..." or "Invalid JSON"

**Solution**:
1. Verify the HAR file is valid JSON (open in a text editor and check structure)
2. Ensure the file was exported correctly from DevTools
3. Check that the file isn't corrupted (re-export from browser)

### Excel file is locked/in use

**Problem**: "Error: Cannot write file - file is locked"

**Solution**:
1. Close the `output.xlsx` file if open in Excel
2. Ensure no other process is using the file
3. Try renaming the existing file and running again

### Out of memory with large HAR files

**Problem**: "JavaScript heap out of memory" or similar

**Solution**:
1. Increase Node.js heap memory: `node --max-old-space-size=4096 har-to-excel.js`
2. Split large HAR files into smaller chunks and process separately
3. Upgrade to Node.js 16+ for improved memory management

## FAQs

### Can I use HAR files from other sources?

Yes. HAR is a standardized format, so files from any source compatible with the standard HAR 1.2 specification will work.

### How often should I capture baseline HAR files?

Establish baselines before each major release, after significant infrastructure changes, or monthly for continuous monitoring. More frequent captures provide better trend data.

### Can I edit the Excel output after generation?

Yes. The Excel file is a standard .xlsx file that can be edited in any spreadsheet application. However, if you re-run the tool, your edits will be lost (the file is overwritten).

### Why is my TTFB so high?

High TTFB (Wait/TTFB (ms)) typically indicates:
- Server-side processing delays (database queries, computation)
- Network latency between client and server
- Server resource constraints (CPU, memory)
- Geographic distance between client and server

Investigate server logs and consider moving services closer to users or optimizing backend code.

### How does URL categorization work?

The tool:
1. Identifies static assets by file extension (.js, .css, .png, etc.) → labeled "Static Asset"
2. Extracts path segments and skips common prefixes (/api, /v1, /v2, /rest, /web, /app, etc.)
3. Uses the first meaningful segment as category (e.g., /api/v1/users/123 → "Users")
4. Converts to Title Case for readability
5. Defaults to "General" if no meaningful segment found

### Can I automate this process?

Yes. Schedule the script with system task scheduler:

**Windows (Task Scheduler)**:
```batch
schtasks /create /tn "HAR-to-Excel Daily" /tr "node c:\OrgProjects\Halliburton\har-to-excel.js" /sc daily /st 09:00
```

**macOS/Linux (cron)**:
```bash
0 9 * * * cd /path/to/project && node har-to-excel.js
```

### What versions of Node.js are supported?

Node.js 12.0 and higher are supported. Version 14+ is recommended for best performance and security. Check your version with `node --version`.

## Performance Considerations

### Large HAR Files (>10MB)

Processing very large HAR files may take several seconds. The tool processes files sequentially, so expect:
- 1-5 MB HAR file: < 1 second processing
- 5-20 MB HAR file: 1-5 seconds processing
- 20+ MB HAR file: 5-30 seconds processing

### Memory Usage

Memory usage is proportional to the number of requests. Typical usage:
- 50 requests: ~2-5 MB
- 500 requests: ~10-15 MB
- 5000+ requests: 50+ MB

If processing multiple large HAR files, increase Node.js heap with `--max-old-space-size`.

### Excel File Size

Output file size depends on:
- Number of requests (100 requests ≈ 50-100 KB)
- Number of HAR files (each adds 5-15 KB baseline)
- URL length and complexity

Typical output.xlsx file sizes:
- 5 HAR files, 50 requests each: 300-500 KB
- 10 HAR files, 100 requests each: 800 KB - 1.5 MB

## License

ISC License. See package.json for details.

## Support and Contribution

For bug reports, feature requests, or contributions, please review the project repository or contact the development team.

---

**Last Updated**: February 2025
**Version**: 1.0.0
**Tool**: HAR-to-Excel Performance Analysis Tool
