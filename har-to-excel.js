const ExcelJS = require('exceljs');
const fs = require('fs-extra');
const path = require('path');

const FONT_DEFAULT = { name: 'Arial', size: 10 };
const FONT_ENDPOINT = { name: 'Consolas', size: 8.5 };
const FONT_HEADER_WHITE = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const FILL_HEADER = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF394B67' } };
const FONT_SECTION_TITLE = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
const FILL_SECTION_TITLE = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2C3E50' } };
const FONT_SUB_HEADER = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
const FILL_SUB_HEADER = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF394B67' } };
const FILL_ROW_EVEN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F4F7' } };
const BORDER_THIN = {
    top: { style: 'thin', color: { argb: 'FFD0D5DD' } },
    bottom: { style: 'thin', color: { argb: 'FFD0D5DD' } },
    left: { style: 'thin', color: { argb: 'FFD0D5DD' } },
    right: { style: 'thin', color: { argb: 'FFD0D5DD' } }
};

class HARtoExcelConverter {
    constructor() {
        this.workbook = new ExcelJS.Workbook();
    }

    async processHARFile(filePath) {
        console.log(`Reading HAR file: ${filePath}`);

        const rawData = await fs.readFile(filePath, 'utf8');
        const harData = JSON.parse(rawData);
        const entries = harData.log.entries;

        console.log(`Parsed ${entries.length} entries from ${path.basename(filePath)}`);

        // Extract page load timings
        const pageTimings = (harData.log.pages || []).map(page => ({
            title: page.title || 'Unknown',
            onContentLoad: page.pageTimings?.onContentLoad ?? -1,
            onLoad: page.pageTimings?.onLoad ?? -1
        }));

        const rows = entries.map((entry, index) => {
            const request = entry.request;
            const response = entry.response;
            const url = new URL(request.url);
            const category = this.categorize(url);
            const sizeKB = (response.content.size / 1024).toFixed(2);
            const timeMs = entry.time.toFixed(2);
            const status = response.status;
            const time = parseFloat(timeMs);
            const size = parseFloat(sizeKB);
            const t = entry.timings || {};
            const val = (v) => (v && v >= 0) ? parseFloat(v.toFixed(2)) : 0;

            return {
                '#': index + 1,
                'Method': request.method,
                'Endpoint URL': request.url,
                'Category': category,
                'Status': status,
                'Response (ms)': time,
                'Size (KB)': size,
                'Blocked (ms)': val(t.blocked),
                'DNS (ms)': val(t.dns),
                'Connect (ms)': val(t.connect),
                'SSL (ms)': val(t.ssl),
                'Send (ms)': val(t.send),
                'Wait/TTFB (ms)': val(t.wait),
                'Receive (ms)': val(t.receive)
            };
        });

        return { rows, pageTimings };
    }

    categorize(url) {
        const pathname = url.pathname;
        const mimeHint = url.pathname.split('.').pop().toLowerCase();

        // Static assets by file extension
        const staticExts = ['js', 'css', 'png', 'jpg', 'jpeg', 'gif', 'svg', 'ico', 'woff', 'woff2', 'ttf', 'eot', 'map'];
        if (staticExts.includes(mimeHint)) return 'Static Asset';

        // Extract the first meaningful path segment after stripping common prefixes
        const segments = pathname.split('/').filter(Boolean);

        // Find the first segment that isn't a version prefix or generic wrapper
        const skipSegments = new Set(['api', 'v1', 'v2', 'v3', 'v4', 'rest', 'web', 'app', 'public', 'private', 'internal', 'external']);
        let category = null;
        for (const seg of segments) {
            if (!skipSegments.has(seg.toLowerCase())) {
                category = seg;
                break;
            }
        }

        if (!category) return 'General';

        // Clean up: remove file extensions, camelCase/kebab-case to Title Case
        category = category.replace(/\.[^.]+$/, '');
        category = category
            .replace(/[-_]/g, ' ')
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/\b\w/g, c => c.toUpperCase())
            .trim();

        return category || 'General';
    }

    addSheet(sheetName, rows, pageTimings) {
        if (rows.length === 0) return;

        const ws = this.workbook.addWorksheet(sheetName);
        const totalCols = 14;

        // Column definitions
        const columns = [
            { header: '#', key: 'num', width: 5 },
            { header: 'Method', key: 'method', width: 8 },
            { header: 'Endpoint URL', key: 'url', width: 120 },
            { header: 'Category', key: 'category', width: 20 },
            { header: 'Status', key: 'status', width: 8 },
            { header: 'Response (ms)', key: 'time', width: 15 },
            { header: 'Size (KB)', key: 'size', width: 12 },
            { header: 'Blocked (ms)', key: 'blocked', width: 13 },
            { header: 'DNS (ms)', key: 'dns', width: 10 },
            { header: 'Connect (ms)', key: 'connect', width: 13 },
            { header: 'SSL (ms)', key: 'ssl', width: 10 },
            { header: 'Send (ms)', key: 'send', width: 10 },
            { header: 'Wait/TTFB (ms)', key: 'wait', width: 15 },
            { header: 'Receive (ms)', key: 'receive', width: 13 }
        ];
        ws.columns = columns;

        // Style header row
        const headerRow = ws.getRow(1);
        headerRow.eachCell(cell => {
            cell.font = FONT_HEADER_WHITE;
            cell.fill = FILL_HEADER;
            cell.border = BORDER_THIN;
        });

        // Add data rows
        rows.forEach(row => {
            const dataRow = ws.addRow([
                row['#'],
                row['Method'],
                row['Endpoint URL'],
                row['Category'],
                row['Status'],
                row['Response (ms)'],
                row['Size (KB)'],
                row['Blocked (ms)'],
                row['DNS (ms)'],
                row['Connect (ms)'],
                row['SSL (ms)'],
                row['Send (ms)'],
                row['Wait/TTFB (ms)'],
                row['Receive (ms)']
            ]);

            // Apply fonts: Consolas 8.5 for Endpoint URL (col 3), Arial 10 for rest
            dataRow.eachCell((cell, colNumber) => {
                cell.font = colNumber === 3 ? FONT_ENDPOINT : FONT_DEFAULT;
            });
        });

        // Calculate observed baseline stats
        const times = rows.map(r => r['Response (ms)']);
        const sizes = rows.map(r => r['Size (KB)']);
        const waits = rows.map(r => r['Wait/TTFB (ms)']);
        const totalRequests = rows.length;
        const avgTime = (times.reduce((a, b) => a + b, 0) / totalRequests).toFixed(2);
        const minTime = Math.min(...times).toFixed(2);
        const maxTime = Math.max(...times).toFixed(2);
        const medianTime = this.median(times).toFixed(2);
        const p95Time = this.percentile(times, 95).toFixed(2);
        const avgSize = (sizes.reduce((a, b) => a + b, 0) / totalRequests).toFixed(2);
        const totalSize = sizes.reduce((a, b) => a + b, 0).toFixed(2);
        const avgTTFB = (waits.reduce((a, b) => a + b, 0) / totalRequests).toFixed(2);
        const p95TTFB = this.percentile(waits, 95).toFixed(2);

        // 2 blank rows gap
        ws.addRow([]);
        ws.addRow([]);

        // Baseline section title
        const baselineTitleRow = ws.addRow(['OBSERVED BASELINE']);
        const baselineTitleRowNum = baselineTitleRow.number;
        ws.mergeCells(baselineTitleRowNum, 1, baselineTitleRowNum, totalCols);
        baselineTitleRow.getCell(1).font = FONT_SECTION_TITLE;
        baselineTitleRow.getCell(1).fill = FILL_SECTION_TITLE;
        baselineTitleRow.getCell(1).border = BORDER_THIN;

        // Baseline sub-header
        const baselineHeader = ['Metric', '', '', '', '', 'Response (ms)', 'Size (KB)', '', '', '', '', '', 'Wait/TTFB (ms)'];
        const baselineHeaderRow = ws.addRow(baselineHeader);
        baselineHeaderRow.eachCell(cell => {
            cell.font = FONT_SUB_HEADER;
            cell.fill = FILL_SUB_HEADER;
            cell.border = BORDER_THIN;
        });

        // Baseline data rows
        const baselineRows = [
            ['Total Requests', '', '', '', '', totalRequests, '', '', '', '', '', '', ''],
            ['Average', '', '', '', '', avgTime, avgSize, '', '', '', '', '', avgTTFB],
            ['Median', '', '', '', '', medianTime, '', '', '', '', '', '', ''],
            ['Min', '', '', '', '', minTime, '', '', '', '', '', '', ''],
            ['Max', '', '', '', '', maxTime, '', '', '', '', '', '', ''],
            ['95th Percentile', '', '', '', '', p95Time, '', '', '', '', '', '', p95TTFB],
            ['Total Transfer Size', '', '', '', '', '', totalSize, '', '', '', '', '', ''],
        ];

        baselineRows.forEach((data, i) => {
            const row = ws.addRow(data);
            row.eachCell(cell => {
                cell.font = FONT_DEFAULT;
                cell.border = BORDER_THIN;
                if (i % 2 === 0) cell.fill = FILL_ROW_EVEN;
            });
        });

        // Page load timings section
        if (pageTimings && pageTimings.length > 0) {
            ws.addRow([]);
            ws.addRow([]);

            const pageTitle = ws.addRow(['PAGE LOAD TIMINGS']);
            const pageTitleRowNum = pageTitle.number;
            ws.mergeCells(pageTitleRowNum, 1, pageTitleRowNum, totalCols);
            pageTitle.getCell(1).font = FONT_SECTION_TITLE;
            pageTitle.getCell(1).fill = FILL_SECTION_TITLE;
            pageTitle.getCell(1).border = BORDER_THIN;

            const pageHeader = ws.addRow(['Page', '', '', '', '', 'DOMContentLoaded (ms)', 'Page Load (ms)']);
            pageHeader.eachCell(cell => {
                cell.font = FONT_SUB_HEADER;
                cell.fill = FILL_SUB_HEADER;
                cell.border = BORDER_THIN;
            });

            pageTimings.forEach((pt, i) => {
                const row = ws.addRow([
                    pt.title, '', '', '', '',
                    pt.onContentLoad >= 0 ? pt.onContentLoad.toFixed(2) : 'N/A',
                    pt.onLoad >= 0 ? pt.onLoad.toFixed(2) : 'N/A'
                ]);
                row.eachCell(cell => {
                    cell.font = FONT_DEFAULT;
                    cell.border = BORDER_THIN;
                    if (i % 2 === 0) cell.fill = FILL_ROW_EVEN;
                });
            });
        }

        console.log(`Added sheet: "${sheetName}" with ${rows.length} rows`);
    }

    median(arr) {
        const sorted = [...arr].sort((a, b) => a - b);
        const mid = Math.floor(sorted.length / 2);
        return sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }

    percentile(arr, p) {
        const sorted = [...arr].sort((a, b) => a - b);
        const index = Math.ceil((p / 100) * sorted.length) - 1;
        return sorted[index];
    }

    addHelperSheet() {
        const ws = this.workbook.addWorksheet('Helper');

        ws.columns = [
            { header: 'Metric / Header', key: 'metric', width: 30 },
            { header: 'Description', key: 'desc', width: 90 }
        ];

        // Style header row
        const headerRow = ws.getRow(1);
        headerRow.eachCell(cell => {
            cell.font = FONT_HEADER_WHITE;
            cell.fill = FILL_HEADER;
            cell.border = BORDER_THIN;
        });

        const entries = [
            // Main data columns
            ['', 'DATA COLUMNS'],
            ['#', 'Sequential row number for each network request captured in the HAR file.'],
            ['Method', 'HTTP method used (GET, POST, PUT, DELETE, PATCH, OPTIONS, etc.).'],
            ['Endpoint URL', 'The full URL of the network request including protocol, domain, path and query parameters.'],
            ['Category', 'Auto-derived category based on the first meaningful path segment of the URL.'],
            ['Status', 'HTTP response status code (e.g. 200 OK, 304 Not Modified, 404 Not Found, 500 Server Error).'],
            ['Response (ms)', 'Total round-trip time for the request in milliseconds (ms). Includes all timing phases combined.'],
            ['Size (KB)', 'Response body size in kilobytes (KB). Represents the uncompressed content size.'],

            // Timing breakdown
            ['', ''],
            ['', 'TIMING BREAKDOWN'],
            ['Blocked (ms)', 'Time (ms) the request spent queued in the browser, waiting for a network connection to become available.'],
            ['DNS (ms)', 'Time (ms) spent resolving the domain name to an IP address. 0 if cached or connection reused.'],
            ['Connect (ms)', 'Time (ms) to establish a TCP connection to the server. 0 if connection was reused.'],
            ['SSL (ms)', 'Time (ms) for the TLS/SSL handshake. 0 if HTTP or connection was reused.'],
            ['Send (ms)', 'Time (ms) to send the HTTP request (headers + body) to the server.'],
            ['Wait/TTFB (ms)', 'Time to First Byte (ms). Time waiting for the server to process and begin sending the response. This is the primary indicator of server-side processing time.'],
            ['Receive (ms)', 'Time (ms) to download the full response body from the server.'],

            // Observed baseline
            ['', ''],
            ['', 'OBSERVED BASELINE METRICS'],
            ['Total Requests', 'Total number of HTTP requests captured in the HAR file for this page.'],
            ['Average', 'Arithmetic mean of all values. Useful as a general indicator but can be skewed by outliers.'],
            ['Median', 'The middle value (50th percentile) when sorted. More representative of typical performance than average.'],
            ['Min', 'The fastest/smallest observed value. Represents best-case performance.'],
            ['Max', 'The slowest/largest observed value. Represents worst-case performance.'],
            ['95th Percentile (P95)', 'Value below which 95% of requests fall. Indicates the performance experienced by most users, excluding extreme outliers.'],
            ['Total Transfer Size (KB)', 'Sum of all response sizes in kilobytes (KB). Represents total data downloaded for the page.'],

            // Page load timings
            ['', ''],
            ['', 'PAGE LOAD TIMINGS'],
            ['DOMContentLoaded (ms)', 'Time (ms) from navigation start until the HTML document is fully parsed and the DOM is ready. Scripts marked "defer" have finished. Stylesheets, images, and subframes may still be loading.'],
            ['Page Load (ms)', 'Time (ms) from navigation start until the entire page is fully loaded, including all stylesheets, images, scripts, and subframes. This is the "onLoad" event.'],
        ];

        entries.forEach((entry, i) => {
            const row = ws.addRow({ metric: entry[0], desc: entry[1] });

            // Section headers (entries with empty metric and uppercase description)
            if (entry[0] === '' && entry[1] && entry[1] === entry[1].toUpperCase() && entry[1].length > 0) {
                row.eachCell(cell => {
                    cell.font = FONT_SECTION_TITLE;
                    cell.fill = FILL_SECTION_TITLE;
                    cell.border = BORDER_THIN;
                });
            } else if (entry[0] === '' && entry[1] === '') {
                // blank separator, no styling
            } else {
                row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
                row.getCell(1).border = BORDER_THIN;
                row.getCell(2).font = FONT_DEFAULT;
                row.getCell(2).border = BORDER_THIN;
                if (i % 2 === 0) {
                    row.getCell(1).fill = FILL_ROW_EVEN;
                    row.getCell(2).fill = FILL_ROW_EVEN;
                }
            }
        });

        // Enable word wrap on description column
        ws.getColumn(2).alignment = { wrapText: true, vertical: 'top' };
        ws.getColumn(1).alignment = { vertical: 'top' };

        console.log('Added sheet: "Helper"');
    }

    async save(outputPath) {
        await this.workbook.xlsx.writeFile(outputPath);
        console.log(`Excel file saved: ${outputPath}`);
    }
}

async function main() {
    const outputFile = 'output.xlsx';
    const dir = path.join(process.cwd(), 'harRepo');

    // Auto-discover all .har files in the current directory
    const allFiles = await fs.readdir(dir);
    const harFiles = allFiles.filter(f => f.toLowerCase().endsWith('.har')).sort();

    if (harFiles.length === 0) {
        console.log('No .har files found in harRepo/ directory.');
        process.exit(1);
    }

    console.log(`Found ${harFiles.length} HAR file(s) in ${dir}\n`);

    const converter = new HARtoExcelConverter();

    try {
        for (const fileName of harFiles) {
            const filePath = path.join(dir, fileName);
            const { rows, pageTimings } = await converter.processHARFile(filePath);
            let sheetName = path.basename(fileName, '.har');
            if (sheetName.length > 31) {
                sheetName = sheetName.substring(0, 31);
            }
            converter.addSheet(sheetName, rows, pageTimings);
        }

        converter.addHelperSheet();
        await converter.save(outputFile);
        console.log('\nConversion completed successfully!');
    } catch (error) {
        console.error('Error:', error.message);
        process.exit(1);
    }
}

if (require.main === module) {
    main();
}

module.exports = HARtoExcelConverter;
