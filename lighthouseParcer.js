const ExcelJS = require('exceljs');
const fs = require('fs-extra');
const path = require('path');

// ── Styling Constants (consistent with har-to-excel.js) ──
const FONT_DEFAULT = { name: 'Arial', size: 10 };
const FONT_METRIC_NAME = { name: 'Consolas', size: 9 };
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

// Score color coding
const FILL_SCORE_GOOD = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F4EA' } };
const FONT_SCORE_GOOD = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF137333' } };
const FILL_SCORE_AVG = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF7E0' } };
const FONT_SCORE_AVG = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFE37400' } };
const FILL_SCORE_POOR = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE8E6' } };
const FONT_SCORE_POOR = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFC5221F' } };

class LighthouseToExcelConverter {
    constructor() {
        this.workbook = new ExcelJS.Workbook();
    }

    processLighthouseFile(filePath) {
        console.log(`Reading Lighthouse JSON: ${filePath}`);

        const rawData = fs.readFileSync(filePath, 'utf8');
        const data = JSON.parse(rawData);
        const audits = data.audits || {};
        const categories = data.categories || {};

        // ── Meta Information ──
        const meta = {
            url: data.finalDisplayedUrl || data.requestedUrl || 'N/A',
            fetchTime: data.fetchTime || 'N/A',
            lighthouseVersion: data.lighthouseVersion || 'N/A',
            userAgent: data.userAgent || 'N/A',
            gatherMode: data.gatherMode || 'N/A',
            benchmarkIndex: data.environment?.benchmarkIndex || 'N/A'
        };

        // ── Category Scores ──
        const categoryScores = [];
        const categoryOrder = ['performance', 'accessibility', 'best-practices', 'seo'];
        for (const catId of categoryOrder) {
            if (categories[catId]) {
                categoryScores.push({
                    category: categories[catId].title,
                    score: categories[catId].score !== null ? Math.round(categories[catId].score * 100) : 'N/A'
                });
            }
        }

        // ── Core Web Vitals & Performance Metrics ──
        const val = (id) => audits[id]?.numericValue ?? null;
        const score = (id) => audits[id]?.score ?? null;
        const display = (id) => audits[id]?.displayValue ?? 'N/A';

        const performanceMetrics = [
            { metric: 'First Contentful Paint (FCP)', value: val('first-contentful-paint'), unit: 'ms', score: score('first-contentful-paint'), display: display('first-contentful-paint') },
            { metric: 'Largest Contentful Paint (LCP)', value: val('largest-contentful-paint'), unit: 'ms', score: score('largest-contentful-paint'), display: display('largest-contentful-paint') },
            { metric: 'Total Blocking Time (TBT)', value: val('total-blocking-time'), unit: 'ms', score: score('total-blocking-time'), display: display('total-blocking-time') },
            { metric: 'Cumulative Layout Shift (CLS)', value: val('cumulative-layout-shift'), unit: '', score: score('cumulative-layout-shift'), display: display('cumulative-layout-shift') },
            { metric: 'Speed Index (SI)', value: val('speed-index'), unit: 'ms', score: score('speed-index'), display: display('speed-index') },
            { metric: 'Time to Interactive (TTI)', value: val('interactive'), unit: 'ms', score: score('interactive'), display: display('interactive') },
            { metric: 'Max Potential FID', value: val('max-potential-fid'), unit: 'ms', score: score('max-potential-fid'), display: display('max-potential-fid') },
        ];

        // ── Server & Network ──
        const serverMetrics = [
            { metric: 'Server Response Time (TTFB)', value: val('server-response-time'), unit: 'ms', display: display('server-response-time') },
            { metric: 'Total Byte Weight', value: audits['total-byte-weight']?.numericValue ?? null, unit: 'bytes', display: display('total-byte-weight') },
        ];

        // ── Main Thread Breakdown ──
        const mainThreadItems = audits['mainthread-work-breakdown']?.details?.items || [];
        const mainThreadBreakdown = mainThreadItems.map(item => ({
            category: item.groupLabel || item.group,
            duration: parseFloat((item.duration || 0).toFixed(2))
        }));
        const mainThreadTotal = audits['mainthread-work-breakdown']?.numericValue
            ? parseFloat(audits['mainthread-work-breakdown'].numericValue.toFixed(2))
            : 0;

        // ── JS Execution (Boot-up Time) ──
        const bootupItems = audits['bootup-time']?.details?.items || [];
        const jsExecution = bootupItems.slice(0, 10).map(item => ({
            url: item.url || 'N/A',
            total: parseFloat((item.total || 0).toFixed(2)),
            scripting: parseFloat((item.scripting || 0).toFixed(2)),
            scriptParseCompile: parseFloat((item.scriptParseCompile || 0).toFixed(2))
        }));
        const jsExecutionTotal = audits['bootup-time']?.numericValue
            ? parseFloat(audits['bootup-time'].numericValue.toFixed(2))
            : 0;

        // ── Resource Summary ──
        const resourceItems = audits['resource-summary']?.details?.items || [];
        const resourceSummary = resourceItems.map(item => ({
            type: item.label || item.resourceType,
            requests: item.requestCount || 0,
            transferSize: parseFloat(((item.transferSize || 0) / 1024).toFixed(2))
        }));

        // ── Diagnostics & Opportunities ──
        const diagnostics = [];
        const diagnosticAudits = [
            'unused-javascript', 'unused-css-rules', 'render-blocking-resources',
            'uses-text-compression', 'uses-optimized-images', 'uses-responsive-images',
            'modern-image-formats', 'efficient-animated-content', 'duplicated-javascript',
            'legacy-javascript', 'unminified-css', 'unminified-javascript'
        ];
        for (const id of diagnosticAudits) {
            if (audits[id] && audits[id].score !== null && audits[id].score < 1) {
                diagnostics.push({
                    audit: audits[id].title || id,
                    score: audits[id].score,
                    displayValue: audits[id].displayValue || 'N/A',
                    savings: audits[id].metricSavings ? JSON.stringify(audits[id].metricSavings) : 'N/A'
                });
            }
        }

        console.log(`Parsed metrics for: ${meta.url}`);

        return {
            meta,
            categoryScores,
            performanceMetrics,
            serverMetrics,
            mainThreadBreakdown,
            mainThreadTotal,
            jsExecution,
            jsExecutionTotal,
            resourceSummary,
            diagnostics
        };
    }

    getScoreStyle(score) {
        if (score === null || score === undefined) return { font: FONT_DEFAULT, fill: null };
        if (score >= 0.9) return { font: FONT_SCORE_GOOD, fill: FILL_SCORE_GOOD };
        if (score >= 0.5) return { font: FONT_SCORE_AVG, fill: FILL_SCORE_AVG };
        return { font: FONT_SCORE_POOR, fill: FILL_SCORE_POOR };
    }

    getScoreLabel(score) {
        if (score === null || score === undefined) return 'N/A';
        const pct = Math.round(score * 100);
        if (score >= 0.9) return `${pct} (Good)`;
        if (score >= 0.5) return `${pct} (Needs Work)`;
        return `${pct} (Poor)`;
    }

    addSectionTitle(ws, title, totalCols) {
        const row = ws.addRow([title]);
        const rowNum = row.number;
        ws.mergeCells(rowNum, 1, rowNum, totalCols);
        row.getCell(1).font = FONT_SECTION_TITLE;
        row.getCell(1).fill = FILL_SECTION_TITLE;
        row.getCell(1).border = BORDER_THIN;
        return row;
    }

    addSubHeader(ws, headers) {
        const row = ws.addRow(headers);
        row.eachCell(cell => {
            cell.font = FONT_SUB_HEADER;
            cell.fill = FILL_SUB_HEADER;
            cell.border = BORDER_THIN;
        });
        return row;
    }

    addDataRow(ws, data, index, options = {}) {
        const row = ws.addRow(data);
        row.eachCell((cell, colNumber) => {
            cell.font = (options.metricCol && colNumber === options.metricCol) ? FONT_METRIC_NAME : FONT_DEFAULT;
            cell.border = BORDER_THIN;
            if (index % 2 === 0) cell.fill = FILL_ROW_EVEN;
        });
        return row;
    }

    addSheet(sheetName, report) {
        const ws = this.workbook.addWorksheet(sheetName);
        const totalCols = 6;

        ws.columns = [
            { header: '', key: 'col1', width: 38 },
            { header: '', key: 'col2', width: 22 },
            { header: '', key: 'col3', width: 18 },
            { header: '', key: 'col4', width: 18 },
            { header: '', key: 'col5', width: 18 },
            { header: '', key: 'col6', width: 30 }
        ];

        // ═══════════════════════════════════════════
        // SECTION 1: Audit Info
        // ═══════════════════════════════════════════
        this.addSectionTitle(ws, 'AUDIT INFORMATION', totalCols);
        this.addSubHeader(ws, ['Property', 'Value', '', '', '', '']);

        const metaRows = [
            ['URL', report.meta.url],
            ['Fetch Time', report.meta.fetchTime],
            ['Lighthouse Version', report.meta.lighthouseVersion],
            ['Gather Mode', report.meta.gatherMode],
            ['Benchmark Index', report.meta.benchmarkIndex],
        ];
        metaRows.forEach((data, i) => {
            const row = ws.addRow([data[0], data[1], '', '', '', '']);
            row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
            row.getCell(1).border = BORDER_THIN;
            row.getCell(2).font = FONT_DEFAULT;
            row.getCell(2).border = BORDER_THIN;
            if (i % 2 === 0) {
                row.getCell(1).fill = FILL_ROW_EVEN;
                row.getCell(2).fill = FILL_ROW_EVEN;
            }
        });

        // ═══════════════════════════════════════════
        // SECTION 2: Category Scores
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'CATEGORY SCORES', totalCols);
        this.addSubHeader(ws, ['Category', 'Score', 'Rating', '', '', '']);

        report.categoryScores.forEach((cat, i) => {
            const scoreVal = cat.score;
            let rating = 'N/A';
            let style = { font: FONT_DEFAULT, fill: null };
            if (typeof scoreVal === 'number') {
                const normalized = scoreVal / 100;
                style = this.getScoreStyle(normalized);
                rating = normalized >= 0.9 ? 'Good' : normalized >= 0.5 ? 'Needs Work' : 'Poor';
            }
            const row = ws.addRow([cat.category, typeof scoreVal === 'number' ? scoreVal : 'N/A', rating, '', '', '']);
            row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
            row.getCell(1).border = BORDER_THIN;
            row.getCell(2).font = style.font;
            row.getCell(2).border = BORDER_THIN;
            if (style.fill) row.getCell(2).fill = style.fill;
            row.getCell(3).font = style.font;
            row.getCell(3).border = BORDER_THIN;
            if (style.fill) row.getCell(3).fill = style.fill;
            if (i % 2 === 0 && !style.fill) {
                row.getCell(1).fill = FILL_ROW_EVEN;
                row.getCell(2).fill = FILL_ROW_EVEN;
                row.getCell(3).fill = FILL_ROW_EVEN;
            }
        });

        // ═══════════════════════════════════════════
        // SECTION 3: Performance Metrics (Core Web Vitals)
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'PERFORMANCE METRICS', totalCols);
        this.addSubHeader(ws, ['Metric', 'Raw Value', 'Unit', 'Display', 'Score', 'Rating']);

        report.performanceMetrics.forEach((m, i) => {
            const rawVal = m.value !== null ? parseFloat(m.value.toFixed(2)) : 'N/A';
            const scoreLabel = this.getScoreLabel(m.score);
            const style = this.getScoreStyle(m.score);

            const row = ws.addRow([m.metric, rawVal, m.unit, m.display, typeof m.score === 'number' ? Math.round(m.score * 100) : 'N/A', scoreLabel]);
            row.getCell(1).font = FONT_METRIC_NAME;
            row.getCell(1).border = BORDER_THIN;
            for (let c = 2; c <= 4; c++) {
                row.getCell(c).font = FONT_DEFAULT;
                row.getCell(c).border = BORDER_THIN;
            }
            row.getCell(5).font = style.font;
            row.getCell(5).border = BORDER_THIN;
            if (style.fill) row.getCell(5).fill = style.fill;
            row.getCell(6).font = style.font;
            row.getCell(6).border = BORDER_THIN;
            if (style.fill) row.getCell(6).fill = style.fill;

            if (i % 2 === 0 && !style.fill) {
                for (let c = 1; c <= 6; c++) {
                    if (!row.getCell(c).fill || !row.getCell(c).fill.fgColor) {
                        row.getCell(c).fill = FILL_ROW_EVEN;
                    }
                }
            }
        });

        // ═══════════════════════════════════════════
        // SECTION 4: Server & Network
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'SERVER & NETWORK', totalCols);
        this.addSubHeader(ws, ['Metric', 'Raw Value', 'Unit', 'Display', '', '']);

        report.serverMetrics.forEach((m, i) => {
            const rawVal = m.value !== null ? (m.unit === 'bytes' ? parseFloat((m.value / 1024).toFixed(2)) : parseFloat(m.value.toFixed(2))) : 'N/A';
            const unitLabel = m.unit === 'bytes' ? 'KB' : m.unit;
            const row = ws.addRow([m.metric, rawVal, unitLabel, m.display, '', '']);
            row.getCell(1).font = FONT_METRIC_NAME;
            row.getCell(1).border = BORDER_THIN;
            for (let c = 2; c <= 4; c++) {
                row.getCell(c).font = FONT_DEFAULT;
                row.getCell(c).border = BORDER_THIN;
            }
            if (i % 2 === 0) {
                for (let c = 1; c <= 4; c++) row.getCell(c).fill = FILL_ROW_EVEN;
            }
        });

        // ═══════════════════════════════════════════
        // SECTION 5: Main Thread Breakdown
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'MAIN THREAD BREAKDOWN', totalCols);
        this.addSubHeader(ws, ['Category', 'Duration (ms)', '', '', '', '']);

        report.mainThreadBreakdown.forEach((item, i) => {
            const row = ws.addRow([item.category, item.duration, '', '', '', '']);
            row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
            row.getCell(1).border = BORDER_THIN;
            row.getCell(2).font = FONT_DEFAULT;
            row.getCell(2).border = BORDER_THIN;
            if (i % 2 === 0) {
                row.getCell(1).fill = FILL_ROW_EVEN;
                row.getCell(2).fill = FILL_ROW_EVEN;
            }
        });

        // Total row
        const totalRow = ws.addRow(['Total Main Thread Time', report.mainThreadTotal, '', '', '', '']);
        totalRow.getCell(1).font = { name: 'Arial', size: 10, bold: true };
        totalRow.getCell(1).border = BORDER_THIN;
        totalRow.getCell(1).fill = FILL_SUB_HEADER;
        totalRow.getCell(1).font = FONT_SUB_HEADER;
        totalRow.getCell(2).font = FONT_SUB_HEADER;
        totalRow.getCell(2).fill = FILL_SUB_HEADER;
        totalRow.getCell(2).border = BORDER_THIN;

        // ═══════════════════════════════════════════
        // SECTION 6: JS Execution (Top Scripts)
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'JAVASCRIPT EXECUTION (TOP 10 SCRIPTS)', totalCols);
        this.addSubHeader(ws, ['Script URL', 'Total CPU (ms)', 'Scripting (ms)', 'Parse/Compile (ms)', '', '']);

        report.jsExecution.forEach((item, i) => {
            const row = ws.addRow([item.url, item.total, item.scripting, item.scriptParseCompile, '', '']);
            row.getCell(1).font = FONT_METRIC_NAME;
            row.getCell(1).border = BORDER_THIN;
            for (let c = 2; c <= 4; c++) {
                row.getCell(c).font = FONT_DEFAULT;
                row.getCell(c).border = BORDER_THIN;
            }
            if (i % 2 === 0) {
                for (let c = 1; c <= 4; c++) row.getCell(c).fill = FILL_ROW_EVEN;
            }
        });

        const jsTotalRow = ws.addRow(['Total JS Execution Time', report.jsExecutionTotal, '', '', '', '']);
        jsTotalRow.getCell(1).font = FONT_SUB_HEADER;
        jsTotalRow.getCell(1).fill = FILL_SUB_HEADER;
        jsTotalRow.getCell(1).border = BORDER_THIN;
        jsTotalRow.getCell(2).font = FONT_SUB_HEADER;
        jsTotalRow.getCell(2).fill = FILL_SUB_HEADER;
        jsTotalRow.getCell(2).border = BORDER_THIN;

        // ═══════════════════════════════════════════
        // SECTION 7: Resource Summary
        // ═══════════════════════════════════════════
        ws.addRow([]);
        ws.addRow([]);
        this.addSectionTitle(ws, 'RESOURCE SUMMARY', totalCols);
        this.addSubHeader(ws, ['Resource Type', 'Requests', 'Transfer Size (KB)', '', '', '']);

        report.resourceSummary.forEach((item, i) => {
            const row = ws.addRow([item.type, item.requests, item.transferSize, '', '', '']);
            row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
            row.getCell(1).border = BORDER_THIN;
            row.getCell(2).font = FONT_DEFAULT;
            row.getCell(2).border = BORDER_THIN;
            row.getCell(3).font = FONT_DEFAULT;
            row.getCell(3).border = BORDER_THIN;
            if (i % 2 === 0) {
                row.getCell(1).fill = FILL_ROW_EVEN;
                row.getCell(2).fill = FILL_ROW_EVEN;
                row.getCell(3).fill = FILL_ROW_EVEN;
            }
        });

        // ═══════════════════════════════════════════
        // SECTION 8: Diagnostics & Opportunities
        // ═══════════════════════════════════════════
        if (report.diagnostics.length > 0) {
            ws.addRow([]);
            ws.addRow([]);
            this.addSectionTitle(ws, 'DIAGNOSTICS & OPPORTUNITIES', totalCols);
            this.addSubHeader(ws, ['Audit', 'Score', 'Details', 'Potential Savings', '', '']);

            report.diagnostics.forEach((item, i) => {
                const scoreLabel = this.getScoreLabel(item.score);
                const style = this.getScoreStyle(item.score);
                const row = ws.addRow([item.audit, scoreLabel, item.displayValue, item.savings, '', '']);
                row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
                row.getCell(1).border = BORDER_THIN;
                row.getCell(2).font = style.font;
                row.getCell(2).border = BORDER_THIN;
                if (style.fill) row.getCell(2).fill = style.fill;
                row.getCell(3).font = FONT_DEFAULT;
                row.getCell(3).border = BORDER_THIN;
                row.getCell(4).font = FONT_DEFAULT;
                row.getCell(4).border = BORDER_THIN;
                if (i % 2 === 0 && !style.fill) {
                    for (let c = 1; c <= 4; c++) row.getCell(c).fill = FILL_ROW_EVEN;
                }
            });
        }

        console.log(`Added sheet: "${sheetName}"`);
    }

    addHelperSheet() {
        const ws = this.workbook.addWorksheet('Helper');

        ws.columns = [
            { header: 'Metric / Header', key: 'metric', width: 40 },
            { header: 'Description', key: 'desc', width: 100 }
        ];

        // Style header row
        const headerRow = ws.getRow(1);
        headerRow.eachCell(cell => {
            cell.font = FONT_HEADER_WHITE;
            cell.fill = FILL_HEADER;
            cell.border = BORDER_THIN;
        });

        const entries = [
            // Category Scores
            ['', 'CATEGORY SCORES'],
            ['Performance', 'Overall performance score (0–100) computed from weighted Core Web Vitals. Scores 90+ are Good, 50–89 Needs Work, below 50 is Poor.'],
            ['Accessibility', 'Score (0–100) measuring how accessible the page is to users with disabilities. Based on automated axe-core checks for ARIA, contrast, labels, etc.'],
            ['Best Practices', 'Score (0–100) for general web development best practices including HTTPS, console errors, deprecated APIs, and image aspect ratios.'],
            ['SEO', 'Score (0–100) for basic search engine optimization checks including meta tags, crawlability, and structured data.'],

            // Core Web Vitals
            ['', ''],
            ['', 'CORE WEB VITALS & PERFORMANCE METRICS'],
            ['First Contentful Paint (FCP)', 'Time (ms) from navigation start to when the browser renders the first piece of DOM content (text, image, SVG, or canvas). Indicates how quickly users see something on screen. Good: < 1.8s, Poor: > 3.0s.'],
            ['Largest Contentful Paint (LCP)', 'Time (ms) until the largest content element (image, video, or text block) in the viewport is fully rendered. Primary metric for perceived load speed. Good: < 2.5s, Poor: > 4.0s.'],
            ['Total Blocking Time (TBT)', 'Total time (ms) between FCP and TTI where the main thread was blocked for more than 50ms. Measures interactivity responsiveness. Correlates with First Input Delay (FID). Good: < 200ms, Poor: > 600ms.'],
            ['Cumulative Layout Shift (CLS)', 'Unitless score measuring unexpected visual movement of page elements during load. Quantifies visual stability. Good: < 0.1, Poor: > 0.25.'],
            ['Speed Index (SI)', 'Time (ms) measuring how quickly content is visually displayed during page load. Calculated from video capture of the loading process. Good: < 3.4s, Poor: > 5.8s.'],
            ['Time to Interactive (TTI)', 'Time (ms) from navigation start until the page is fully interactive — the main thread has been idle for at least 5 seconds with no long tasks. Indicates when users can reliably interact with the page.'],
            ['Max Potential FID', 'Maximum duration (ms) of the longest task on the main thread. Represents the worst-case First Input Delay a user could experience if they interacted during the heaviest processing period.'],

            // Server & Network
            ['', ''],
            ['', 'SERVER & NETWORK'],
            ['Server Response Time (TTFB)', 'Time (ms) for the server to respond to the initial document request. Includes DNS, TCP, TLS handshake, and server processing. Good: < 200ms, Poor: > 600ms. Also known as Time to First Byte for the document.'],
            ['Total Byte Weight', 'Total transfer size (KB) of all resources loaded by the page. Large payloads increase load time and consume user bandwidth. Lighthouse flags pages exceeding 5,000 KB.'],

            // Main Thread
            ['', ''],
            ['', 'MAIN THREAD BREAKDOWN'],
            ['Script Evaluation', 'Time (ms) spent executing JavaScript code on the main thread. The largest contributor to TBT and TTI.'],
            ['Script Parsing & Compilation', 'Time (ms) spent parsing and compiling JavaScript before execution. Reduced by shipping smaller, optimized bundles.'],
            ['Style & Layout', 'Time (ms) spent recalculating CSS styles and computing element layout (reflow). Triggered by DOM changes and CSS modifications.'],
            ['Parse HTML & CSS', 'Time (ms) spent parsing the HTML document and CSS stylesheets into DOM and CSSOM trees.'],
            ['Rendering', 'Time (ms) spent compositing layers and painting pixels to the screen.'],
            ['Garbage Collection', 'Time (ms) spent by the JavaScript engine reclaiming unused memory. Excessive GC indicates memory pressure from large object allocations.'],
            ['Other', 'Time (ms) spent on other main thread activities not categorized above.'],

            // JS Execution
            ['', ''],
            ['', 'JAVASCRIPT EXECUTION'],
            ['Total CPU Time', 'Total time (ms) a script spent on the main thread across all activities (evaluation, parsing, compilation). Top scripts by CPU time are the primary optimization targets.'],
            ['Scripting', 'Time (ms) spent evaluating and executing the JavaScript code within this script file.'],
            ['Parse/Compile', 'Time (ms) spent parsing the source code and compiling it to bytecode for this specific script file.'],

            // Resource Summary
            ['', ''],
            ['', 'RESOURCE SUMMARY'],
            ['Resource Type', 'Category of network resource: Script, Stylesheet, Image, Font, Document, Media, or Other.'],
            ['Requests', 'Number of HTTP requests made for this resource type.'],
            ['Transfer Size (KB)', 'Total compressed transfer size in kilobytes for this resource type. Represents actual bytes downloaded over the network.'],

            // Diagnostics
            ['', ''],
            ['', 'DIAGNOSTICS & OPPORTUNITIES'],
            ['Unused JavaScript', 'JavaScript code that was downloaded but not executed during page load. Reducing unused JS decreases network cost and main thread processing.'],
            ['Unused CSS', 'CSS rules that were downloaded but not applied to any visible elements. Removing unused CSS reduces transfer size and style calculation time.'],
            ['Render-Blocking Resources', 'Scripts and stylesheets in the document head that block the first paint. Deferring or async-loading these resources improves FCP.'],
            ['Text Compression', 'Resources served without gzip/brotli compression. Enabling compression typically reduces transfer size by 60–80%.'],
            ['Potential Savings', 'Estimated improvement in metric values (FCP, LCP, TBT) if the diagnostic issue is resolved. Shown as JSON with metric abbreviations and millisecond savings.'],

            // Scoring
            ['', ''],
            ['', 'SCORE INTERPRETATION'],
            ['Good (90–100)', 'The metric or category is performing well. Shown with green highlighting. Meets the recommended threshold for a good user experience.'],
            ['Needs Work (50–89)', 'The metric or category has room for improvement. Shown with orange/amber highlighting. May impact user experience under certain conditions.'],
            ['Poor (0–49)', 'The metric or category is significantly below recommended thresholds. Shown with red highlighting. Likely degrading user experience and should be prioritized for optimization.'],
        ];

        entries.forEach((entry, i) => {
            const row = ws.addRow({ metric: entry[0], desc: entry[1] });

            if (entry[0] === '' && entry[1] && entry[1] === entry[1].toUpperCase() && entry[1].length > 0) {
                row.eachCell(cell => {
                    cell.font = FONT_SECTION_TITLE;
                    cell.fill = FILL_SECTION_TITLE;
                    cell.border = BORDER_THIN;
                });
            } else if (entry[0] === '' && entry[1] === '') {
                // blank separator
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
    const outputFile = 'lighthouse-output.xlsx';
    const dir = path.join(process.cwd(), 'lighthouseRepo');

    const allFiles = await fs.readdir(dir);
    const jsonFiles = allFiles.filter(f => f.toLowerCase().endsWith('.json')).sort();

    if (jsonFiles.length === 0) {
        console.log('No .json files found in lighthouseRepo/ directory.');
        process.exit(1);
    }

    console.log(`Found ${jsonFiles.length} Lighthouse JSON file(s) in ${dir}\n`);

    const converter = new LighthouseToExcelConverter();

    try {
        for (const fileName of jsonFiles) {
            const filePath = path.join(dir, fileName);
            const report = converter.processLighthouseFile(filePath);
            let sheetName = path.basename(fileName, '.json');
            if (sheetName.length > 31) {
                sheetName = sheetName.substring(0, 31);
            }
            converter.addSheet(sheetName, report);
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

module.exports = LighthouseToExcelConverter;
