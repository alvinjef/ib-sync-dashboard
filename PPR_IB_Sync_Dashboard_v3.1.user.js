// ==UserScript==

// @name         PPR: IB Sync Dashboard

// @namespace    http://tampermonkey.net/

// @version      3.1

// @author       Alvin Jefferson (alvinjef)
// @updateURL    https://github.com/alvinjef/ib-sync-dashboard/raw/refs/heads/main/PPR_IB_Sync_Dashboard_v3.1.user.js
// @downloadURL  https://github.com/alvinjef/ib-sync-dashboard/raw/refs/heads/main/PPR_IB_Sync_Dashboard_v3.1.user.js
// @description  Auto-pull IB metrics: NTP from INTRO APIs, COST from FCLM (LP Rate/Actual Rate), Input Metrics from Vantage, Audits from Apollo (audit_execution_metrics with period filtering). Single script.

// @match        https://fclm-portal.amazon.com/reports/processPathRollup*

// @grant        GM_xmlhttpRequest

// @grant        GM_setClipboard

// @connect      apollo-audit.corp.amazon.com

// @connect      vantage.amazon.com

// @connect      fclm-portal.amazon.com

// @connect      slim.corp.amazon.com

// @connect      atlas.qubit.amazon.dev

// @connect      roboscout.amazon.com

// @connect      api.inbound-flow.na.prod.fmc.aft.amazon.dev

// @connect      outbound-flow-api.na.prod.fmc.aft.amazon.dev

// @connect      intro-trailerplanner-api.na.prod.fmc.aft.amazon.dev

// @connect      intro-shift-planner-api.na.prod.fmc.aft.amazon.dev

// @connect      alps-iad.iad.proxy.amazon.com

// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js

// @require      https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js

// ==/UserScript==



(function() {

    'use strict';



    // ==========================================

    // CONFIGURATION

    // ==========================================

    //

    // ╔══════════════════════════════════════════════════════════════╗

    // ║  SITE CONFIGURATION — Change these values for your building ║

    // ╠══════════════════════════════════════════════════════════════╣

    // ║  ROBOSCOUT_INSTANCE_ID: Find yours in the RoboScout URL     ║

    // ║    when you open it for your site (look for instance_id=)   ║

    // ║  FLOORS: Your building's stow floor names (e.g. A02-A05)    ║

    // ║  TIMEZONE: Your site's IANA timezone string                  ║

    // ║  APOLLO_DEPARTMENT: Usually 4 (ICQA) for all sites           ║

    // ╚══════════════════════════════════════════════════════════════╝



    var SITE_SETTINGS = {

        ROBOSCOUT_INSTANCE_ID: 1932,          // ORD5=1932, CLT4=2210

        ROBOSCOUT_STOW_OBJECT_ID: 19851,      // Same across sites (stow metrics)

        ROBOSCOUT_OOWA_OBJECT_ID: 21628,      // Same across sites (OOWA andons)

        FLOORS: ['A02', 'A03', 'A04', 'A05'], // Stow floor names

        ZONE_PREFIX: 'paKiva',                 // Zone prefix (e.g. paKivaA02)

        TIMEZONE: 'America/Chicago',           // Site timezone for RoboScout

        APOLLO_DEPARTMENT: 4,                  // Apollo department ID (4 = ICQA)

        APOLLO_AUDITS: [                       // Audit names as they appear in Apollo

            { name: "Stow CT", apolloName: "ARS Stow Cycle Time Coaching", goal: 20 },

            { name: "UPF", apolloName: "ORD5 UPF Audit- Stow", goal: 20 },

            { name: "Qty Stow", apolloName: "IDS Quantity Stow - Quality Audit", goal: 5 },

            { name: "FISH", apolloName: "ORD5 FISH Audit", goal: 20 },

            { name: "ASIN Progression", apolloName: "ORD5 Asin Progression Audit 2.0", goal: 20 },

        ],

    };



    // --- Auto-detect warehouse from FCLM URL ---

    function detectWarehouse() {

        var urlParams = new URLSearchParams(window.location.search);

        var warehouseFromUrl = urlParams.get('warehouseId') || urlParams.get('warehouse');

        if (warehouseFromUrl) return warehouseFromUrl.toUpperCase();



        var headerText = document.querySelector('.site-header, .warehouse-name, [class*="warehouse"]');

        if (headerText) {

            var match = headerText.textContent.match(/([A-Z]{3}\d)/);

            if (match) return match[1];

        }



        var bodyText = (document.body && document.body.textContent) ? document.body.textContent.substring(0, 5000) : '';

        var siteMatch = bodyText.match(/\b([A-Z]{3}\d)\b/);

        if (siteMatch) return siteMatch[1];



        var saved = localStorage.getItem('ppr_sync_warehouse');

        if (saved) return saved;



        var input = prompt('IB Sync Dashboard: Enter your site code (e.g., ORD5, DFW7):');

        if (input) {

            localStorage.setItem('ppr_sync_warehouse', input.toUpperCase());

            return input.toUpperCase();

        }

        return 'UNKNOWN';

    }



    var WAREHOUSE = detectWarehouse();

    console.log('[IB Sync] Detected warehouse: ' + WAREHOUSE);

    // Shift/date detection logged after functions are defined (see below)



    // --- Site-specific settings ---

    var SITE_CONFIG_KEY = 'ppr_sync_site_config_' + WAREHOUSE;



    function getDefaultSiteConfig(warehouse) {

        return {

            warehouse: warehouse,

            customer: 'AMZN',

            zones: SITE_SETTINGS.FLOORS.map(function(f) { return SITE_SETTINGS.ZONE_PREFIX + f; }),

            oowaAndonIds: [],

            apolloAudits: [],

        };

    }



    function getSiteConfig() {

        try {

            var stored = localStorage.getItem(SITE_CONFIG_KEY);

            if (stored) return JSON.parse(stored);

        } catch (e) { /* fall through */ }

        return getDefaultSiteConfig(WAREHOUSE);

    }



    function saveSiteConfig(config) {

        localStorage.setItem(SITE_CONFIG_KEY, JSON.stringify(config));

    }



    var SITE = getSiteConfig();



    var CONFIG = {

        // Shift times (24h format)

        dayShift: { start: 6, end: 18 },

        nightShift: { start: 18, end: 6 },



        // Apollo Audit Types

        apolloAudits: SITE_SETTINGS.APOLLO_AUDITS,



        // Goals

        goals: {

            stowCycleTime: 13,

            upf: 12,

            oowa: 2,

            nsta: 15,

        },



        // Vantage

        vantage: {

            warehouse: WAREHOUSE,

            customer: SITE.customer,

            zones: SITE.zones,

            baseUrl: 'https://vantage.amazon.com/fulfillment',

            oowaAndonIds: SITE.oowaAndonIds,

        },



        // INTRO API endpoints (direct JSON calls, no scraper needed)

        introApis: {

            ibFlow: 'https://api.inbound-flow.na.prod.fmc.aft.amazon.dev/v1/',

            trailerPlanner: 'https://intro-trailerplanner-api.na.prod.fmc.aft.amazon.dev/',

            shiftPlanner: 'https://intro-shift-planner-api.na.prod.fmc.aft.amazon.dev/',

            outboundFlow: 'https://outbound-flow-api.na.prod.fmc.aft.amazon.dev:80/',

        },

    };



    // ==========================================

    // UTILITY: Period & Shift Detection

    // ==========================================



    // Detect shift from FCLM URL time selection (not wall clock)

    function detectShiftFromFCLM() {

        try {

            var urlParams = new URLSearchParams(window.location.search);

            var spanType = urlParams.get('spanType');



            if (spanType === 'Intraday') {

                // Intraday has separate hour params

                // Use getAll() to handle multiple params with similar names, take the first match

                var startHourRaw = urlParams.get('startHourIntraday');

                var endHourRaw = urlParams.get('endHourIntraday');

                var startHour = parseInt(startHourRaw);

                var endHour = parseInt(endHourRaw);

                if (isNaN(startHour)) startHour = 0;

                if (isNaN(endHour)) endHour = 0;

                console.log('[IB Sync] FCLM Intraday hours: start=' + startHour + ', end=' + endHour);

                // Day shift: start=6 end=18; Night shift: start=18 end=6

                // Check both start AND end to be sure

                if (startHour === 6 && endHour === 18) return 'Day';

                if (startHour === 18 && (endHour === 6 || endHour === 0)) return 'Night';

                // Fallback: if start is in day range

                if (startHour >= 6 && startHour < 18 && endHour > startHour) return 'Day';

                return 'Night';

            }



            if (spanType === 'Day') {

                // Day view shows full 24h; check if selected date is today

                // If it's today, use wall clock. If past date, default to Day shift.

                var dateStr = urlParams.get('startDateDay') || '';

                dateStr = dateStr.replace(/\//g, '-').substring(0, 10);

                var today = new Date().toISOString().split('T')[0];

                if (dateStr && dateStr !== today) return 'Day'; // Past date defaults to Day

            }

        } catch(e) {}



        // Fallback: use wall clock

        var hour = new Date().getHours();

        return (hour >= 6 && hour < 18) ? 'Day' : 'Night';

    }



    function getCurrentShift() {

        return detectShiftFromFCLM();

    }



    function getCurrentPeriod() {

        var shift = getCurrentShift();

        var hour = new Date().getHours();

        if (shift === 'Day') {

            if (hour >= 6 && hour < 10) return 'P1';

            if (hour >= 10 && hour < 14) return 'P2';

            if (hour >= 14 && hour < 18) return 'P3';

            return 'EOS'; // Outside day shift hours viewing day data = show EOS

        } else {

            if (hour >= 18 && hour < 22) return 'P1';

            if (hour >= 22 || hour < 2) return 'P2';

            if (hour >= 2 && hour < 6) return 'P3';

            return 'EOS'; // Outside night shift hours viewing night data = show EOS

        }

    }



    // Get the date from FCLM URL (for API calls)

    function getSelectedDate() {

        try {

            var urlParams = new URLSearchParams(window.location.search);

            var spanType = urlParams.get('spanType');

            var dateStr = null;

            if (spanType === 'Day') dateStr = urlParams.get('startDateDay');

            else if (spanType === 'Intraday') dateStr = (urlParams.get('startDateIntraday') || '').split(/[\s+T]/)[0];

            else if (spanType === 'Week') dateStr = urlParams.get('startDateWeek');

            if (dateStr) {

                dateStr = dateStr.replace(/\//g, '-');

                if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) return dateStr.substring(0, 10);

            }

        } catch(e) {}

        return new Date().toISOString().split('T')[0];

    }



    function getShiftBounds() {

        var dateStr = getSelectedDate();

        var shift = getCurrentShift();

        var start, end;



        if (shift === 'Day') {

            start = new Date(dateStr + 'T06:00:00');

            end = new Date(dateStr + 'T18:00:00');

        } else {

            start = new Date(dateStr + 'T18:00:00');

            end = new Date(dateStr + 'T06:00:00');

            end.setDate(end.getDate() + 1);

        }

        return { start: start, end: end };

    }



    console.log('[IB Sync] FCLM-detected: date=' + getSelectedDate() + ', shift=' + getCurrentShift() + ', period=' + getCurrentPeriod());



    // ==========================================

    // FCLM TABLE SCRAPER (unchanged from v1)

    // ==========================================

    function scrapeFCLMTable() {

        var data = {

            ibVolumePlan: null,

            percentToLP: null,

            ibVolumeActual: null,

            ibTotalTPH: null,

            ibTotalRate: null,

            ibTotalPlanRate: null,

            decantRate: null,

            decantPlanRate: null,

            stowToprimeRate: null,

            stowToPrimePlanRate: null,

            stowToPrimeSupportRate: null,

            stowToPrimeSupportPlanRate: null,

            stowToPrimeSupportHrs: null,

            stowToPrimeHrs: null,

            transferInRate: null,

            transferInHrs: null,

            transferInSupportRate: null,

            transferInSupportPlanRate: null,

            transferInSupportHrs: null,

            ibLeadPAHrs: null,

            ibLeadPARate: null,

            ibLeadPAPlanRate: null,

            ibProblemSolveHrs: null,

            ibProblemSolveRate: null,

            ibProblemSolvePlanRate: null,

            ibTotalHrs: null,

            ibTotalVol: null,

            stowCycleTime: null,

            stowHC: null,

            decantHC: null,

            problemSolveHC: null,

            indirectHrs: null,

            ibTotalRate: null,

            casePercent: null,

            smallsPercent: null,

            fluidPercent: null,

            ibPercentToPlan: null,

            stowPercentToPlan: null,

            transferInPercentToPlan: null,

        };



        var table = document.querySelector('table.reportLayout, table[class*="report"]');

        if (!table) {

            var tables = document.querySelectorAll('table');

            for (var i = 0; i < tables.length; i++) {

                if (tables[i].querySelector('th') && tables[i].textContent.includes('Main Processes')) {

                    return scrapeFromTable(tables[i], data);

                }

            }

            return scrapeFromAllRows(data);

        }

        return scrapeFromTable(table, data);

    }



    function scrapeFromAllRows(data) {

        var rows = document.querySelectorAll('tr');

        var rowMap = {};



        rows.forEach(function(row) {

            var cells = row.querySelectorAll('td');

            if (cells.length < 7) return;



            var lineItem = '';

            for (var i = 0; i < Math.min(4, cells.length); i++) {

                var text = cells[i] ? cells[i].textContent.trim() : '';

                if (text && text.length > 2 && ['EACH', 'Case', 'Pallet', 'Tote', 'Packages', 'PALLET', 'Hours'].indexOf(text) === -1) {

                    lineItem = text;

                }

            }



            if (lineItem) {

                rowMap[lineItem] = Array.from(cells).map(function(c) { return c.textContent.trim(); });

            }

        });



        data = parseRowData(rowMap, data);

        return data;

    }



    function scrapeFromTable(table, data) {

        var rows = table.querySelectorAll('tr');

        var rowMap = {};

        var rowElements = {}; // Store actual DOM row elements for class-based extraction



        var lineItemPatterns = [

            'Inbound-TOTAL', 'IB Total',

            'Stow to Prime - Total', 'Stow to Prime Support',

            'Each Stow To Prime - Small', 'Each Stow To Prime - Medium', 'Each Stow To Prime - Large',

            'Transfer In Decant',

            'Each Transfer In - Total', 'Each Transfer In - Small', 'Each Transfer In - Medium', 'Each Transfer In - Large',

            'Transfer In - Total',

            'IB Lead/PA', 'IB Problem Solve',

            'Receive - Total', 'Receive Dock',

            'Transfer In Support',

            'Total Inbound', 'THROUGHPUT',

            'Vendor Rec (incl. Stow+Prep+RSR)',

            'Transfer-In',

        ];



        rows.forEach(function(row) {

            var cells = row.querySelectorAll('td, th');

            if (cells.length < 5) return;

            var cellTexts = Array.from(cells).map(function(c) { return c.textContent.trim(); });



            for (var p = 0; p < lineItemPatterns.length; p++) {

                var pattern = lineItemPatterns[p];

                if (cellTexts.some(function(t) { return t === pattern || t.indexOf(pattern) !== -1; })) {

                    rowMap[pattern] = cellTexts;

                    rowElements[pattern] = row; // Store the DOM element

                    break;

                }

            }

        });



        // Class-based extraction helper: find a <td> by CSS class within a row

        function getByClass(rowEl, className) {

            if (!rowEl) return null;

            var td = rowEl.querySelector('td.' + className) || rowEl.querySelector('td[class*="' + className + '"]');

            if (!td) return null;

            var text = td.textContent.trim().replace(/,/g, '').replace('%', '');

            var v = parseFloat(text);

            return isNaN(v) ? null : v;

        }



        // Extract key metrics using class names (immune to column order changes)

        // Classes from FCLM DOM: actualVolume, actualTimeSeconds, actualProductivity, planProductivity, ratioToPlan

        // LP classes: look for cells with class containing 'lpRate' or similar

        function extractFromRow(rowEl) {

            if (!rowEl) return { vol: null, hrs: null, rate: null, planRate: null, pctToPlan: null, lpRate: null, pctToLP: null };

            var cells = rowEl.querySelectorAll('td');

            var result = { vol: null, hrs: null, rate: null, planRate: null, pctToPlan: null, lpRate: null, pctToLP: null };

            cells.forEach(function(td) {

                var cls = td.className || '';

                var text = td.textContent.trim().replace(/,/g, '').replace('%', '');

                var v = parseFloat(text);

                if (isNaN(v)) v = null;

                if (cls.indexOf('actualVolume') !== -1) { result.vol = v; }

                else if (cls.indexOf('actualTimeSeconds') !== -1 && cls.indexOf('priorYear') === -1) { result.hrs = v; }

                else if (cls.indexOf('actualProductivity') !== -1 && cls.indexOf('priorYear') === -1) { result.rate = v; }

                else if (cls.indexOf('planProductivity') !== -1 && cls.indexOf('priorYear') === -1) { result.planRate = v; }

                else if (cls.indexOf('ratioToPlan') !== -1 && cls.indexOf('priorYear') === -1) { result.pctToPlan = v; }

                // LP Rate: PPR VS LP injects LP Rate into .priorYearVolume on visible page

                else if (cls.indexOf('priorYearVolume') !== -1) {

                    var lpV = parseFloat(text);

                    if (!isNaN(lpV) && lpV > 0) result.lpRate = lpV;

                }

                // % to LP: PPR VS LP injects % to LP into .priorYearProductivity on visible page

                else if (cls.indexOf('priorYearProductivity') !== -1) {

                    var pctV = parseFloat(text);

                    if (!isNaN(pctV)) result.pctToLP = pctV;

                }

            });

            return result;

        }



        // Store extracted data on rowElements for use in parseRowData

        data._rowElements = rowElements;

        data._extractFromRow = extractFromRow;



        data = parseRowData(rowMap, data);

        return data;

    }



    function parseRowData(rowMap, data) {

        function num(val) {

            if (!val || val === '---' || val === '' || val === '-') return null;

            var cleaned = val.replace(/,/g, '').replace('%', '');

            var n = parseFloat(cleaned);

            return isNaN(n) ? null : n;

        }



        // --- Detect column positions from header row ---

        // FCLM tables have multi-level headers. We need to find:

        // "Rate" (actual throughput) and "LP Rate" (labor plan rate)

        var colMap = detectFCLMColumns();

        var rateCol = colMap.rateCol;      // actual Rate column index

        var lpRateCol = colMap.lpRateCol;  // LP Rate column index

        var volCol = colMap.volCol;        // actual Volume column index

        var hrsCol = colMap.hrsCol;        // actual Hours column index

        var ptpCol = colMap.ptpCol;        // % to Plan column index



        console.log('[IB Sync] Column indices - Rate:', rateCol, 'LP Rate:', lpRateCol, 'Vol:', volCol, 'Hrs:', hrsCol, 'PTP:', ptpCol);





        // --- Extract data per row using END-COUNTING ---

        // FCLM row cells end with: [...] [LP Rate] [% to LP] [Δ to LP] [checkbox]

        // So for any row: LP Rate = row[length-4], Actual Rate is in the Actual section.

        // Actual section layout (from left): [text cells...] [Unit] [Vol] [Hrs] [Rate(Actual)]

        // Then Plan section: [Rate(Plan)] [Hrs(Plan)] [Δ to Plan] [% to Plan]

        // Then Labor Plan: [LP Rate] [% to LP] [Δ to LP] [checkbox]

        //

        // FCLM table has variable row lengths (15 or 16 cells depending on whether

        // there's an extra icon/expand cell). The rightmost columns are always:

        // ... | LP Rate | % to LP | Δ to LP (Hrs) | checkbox

        // LP Rate is always 4th from end, % to LP is 3rd from end.

        //

        // The Actual Rate column position varies:

        // - 15-cell rows: Actual Rate is 9th from end

        // - 16-cell rows: Actual Rate is 10th from end (extra cell on left)

        // Detect dynamically based on row length.



        function getFromEnd(row, offsetFromEnd) {

            if (!row || row.length < offsetFromEnd) return null;

            return num(row[row.length - offsetFromEnd]);

        }



        // Helper: get Actual Rate and LP Rate for any row

        // LP Rate is always 4th from end

        // Actual Rate offset depends on row length

        function getActualRate(row) {

            if (!row) return null;

            var offset = row.length >= 16 ? 10 : 9;

            return getFromEnd(row, offset);

        }

        function getLPRate(row) { return getFromEnd(row, 4); }

        function getPercentToLP(row) { return getFromEnd(row, 3); }

        function getVol(row) {

            var offset = row.length >= 16 ? 12 : 11;

            return getFromEnd(row, offset);

        }

        function getHrs(row) {

            var offset = row.length >= 16 ? 11 : 10;

            return getFromEnd(row, offset);

        }



        var ibTotal = rowMap['Inbound-TOTAL'] || rowMap['IB Total'];

        if (ibTotal) {

            // Use class-based extraction if available (immune to Prior Year column shifts)

            var ibEl = (data._rowElements || {})['Inbound-TOTAL'] || (data._rowElements || {})['IB Total'];

            var ibEx = (data._extractFromRow && ibEl) ? data._extractFromRow(ibEl) : null;

            if (ibEx && ibEx.rate !== null) {

                data.ibTotalRate = ibEx.rate;

                data.ibTotalVol = ibEx.vol;

                data.ibTotalHrs = ibEx.hrs;

                // LP Rate and % to LP: class-based may return null; always fall back to end-counting (LP is always at fixed end positions)

                data.ibTotalPlanRate = ibEx.lpRate !== null ? ibEx.lpRate : getLPRate(ibTotal);

                data.percentToLP = ibEx.pctToLP !== null ? ibEx.pctToLP : getPercentToLP(ibTotal);

            } else {

                data.ibTotalRate = getActualRate(ibTotal);

                data.ibTotalPlanRate = getLPRate(ibTotal);

                data.ibTotalVol = getVol(ibTotal);

                data.ibTotalHrs = getHrs(ibTotal);

                data.percentToLP = getPercentToLP(ibTotal);

            }

            data.ibTotalTPH = data.ibTotalRate;

            console.log('[IB Sync] FCLM IB Total - Actual Rate:', data.ibTotalRate, 'LP Rate:', data.ibTotalPlanRate, '% to LP:', data.percentToLP, 'row length:', ibTotal.length);

        }



        var stowSupport = rowMap['Stow to Prime Support'];

        if (stowSupport) {

            var stpEl = (data._rowElements || {})['Stow to Prime Support'];

            var stpEx = (data._extractFromRow && stpEl) ? data._extractFromRow(stpEl) : null;

            data.stowToPrimeSupportRate = (stpEx && stpEx.rate !== null) ? stpEx.rate : getActualRate(stowSupport);

            data.stowToPrimeSupportPlanRate = (stpEx && stpEx.lpRate !== null) ? stpEx.lpRate : getLPRate(stowSupport);

            data.stowToPrimePlanRate = data.stowToPrimeSupportPlanRate;

            console.log('[IB Sync] FCLM STP Support - Actual:', data.stowToPrimeSupportRate, 'LP Rate:', data.stowToPrimeSupportPlanRate, 'row length:', stowSupport.length);

        }



        var stow = rowMap['Stow to Prime - Total'] || rowMap['Each Stow to Prime - Total'];

        if (stow) {

            var stowEl = (data._rowElements || {})['Stow to Prime - Total'] || (data._rowElements || {})['Each Stow to Prime - Total'];

            var stowEx = (data._extractFromRow && stowEl) ? data._extractFromRow(stowEl) : null;

            data.stowToprimeRate = (stowEx && stowEx.rate !== null) ? stowEx.rate : getActualRate(stow);

            if (!data.stowToPrimePlanRate) data.stowToPrimePlanRate = (stowEx && stowEx.lpRate !== null) ? stowEx.lpRate : getLPRate(stow);

            data.stowToPrimeHrs = (stowEx && stowEx.hrs !== null) ? stowEx.hrs : getHrs(stow);

        }



        var decant = rowMap['Transfer In Decant'];

        if (decant) {

            var decEl = (data._rowElements || {})['Transfer In Decant'];

            var decEx = (data._extractFromRow && decEl) ? data._extractFromRow(decEl) : null;

            data.decantRate = (decEx && decEx.rate !== null) ? decEx.rate : getActualRate(decant);

            data.decantPlanRate = (decEx && decEx.lpRate !== null) ? decEx.lpRate : getLPRate(decant);

            console.log('[IB Sync] FCLM Decant - Actual:', data.decantRate, 'LP Rate:', data.decantPlanRate, 'row length:', decant.length);

        }



        var transferIn = rowMap['Each Transfer In - Total'] || rowMap['Transfer In - Total'];

        if (transferIn) {

            var tiEl = (data._rowElements || {})['Each Transfer In - Total'] || (data._rowElements || {})['Transfer In - Total'];

            var tiEx = (data._extractFromRow && tiEl) ? data._extractFromRow(tiEl) : null;

            data.transferInRate = (tiEx && tiEx.rate !== null) ? tiEx.rate : getActualRate(transferIn);

            data.transferInVol = (tiEx && tiEx.vol !== null) ? tiEx.vol : getVol(transferIn);

            data.transferInHrs = (tiEx && tiEx.hrs !== null) ? tiEx.hrs : getHrs(transferIn);

            console.log('[IB Sync] FCLM ETI (Each Transfer In - Total) - Rate:', data.transferInRate, 'Hrs:', data.transferInHrs);

        }



        console.log('[IB Sync] FCLM Hours for OOWA% - ETI (Transfer In) Hrs:', data.transferInHrs, 'STP (Stow to Prime) Hrs:', data.stowToPrimeHrs);



        var leadPA = rowMap['IB Lead/PA'];

        if (leadPA) {

            var lpaEl = (data._rowElements || {})['IB Lead/PA'];

            var lpaEx = (data._extractFromRow && lpaEl) ? data._extractFromRow(lpaEl) : null;

            data.ibLeadPARate = (lpaEx && lpaEx.rate !== null) ? lpaEx.rate : getActualRate(leadPA);

            data.ibLeadPAPlanRate = (lpaEx && lpaEx.lpRate !== null) ? lpaEx.lpRate : getLPRate(leadPA);

            console.log('[IB Sync] FCLM Lead/PA - Actual:', data.ibLeadPARate, 'LP Rate:', data.ibLeadPAPlanRate, 'row length:', leadPA.length);

        }



        var ps = rowMap['IB Problem Solve'];

        if (ps) {

            var psEl = (data._rowElements || {})['IB Problem Solve'];

            var psEx = (data._extractFromRow && psEl) ? data._extractFromRow(psEl) : null;

            data.ibProblemSolveRate = (psEx && psEx.rate !== null) ? psEx.rate : getActualRate(ps);

            data.ibProblemSolvePlanRate = (psEx && psEx.lpRate !== null) ? psEx.lpRate : getLPRate(ps);

            console.log('[IB Sync] FCLM Problem Solve - Actual:', data.ibProblemSolveRate, 'LP Rate:', data.ibProblemSolvePlanRate, 'row length:', ps.length);

        }



        var tiSupport = rowMap['Transfer In Support'];

        if (tiSupport) {

            var tisEl = (data._rowElements || {})['Transfer In Support'];

            var tisEx = (data._extractFromRow && tisEl) ? data._extractFromRow(tisEl) : null;

            data.transferInSupportRate = (tisEx && tisEx.rate !== null) ? tisEx.rate : getActualRate(tiSupport);

            data.transferInSupportPlanRate = (tisEx && tisEx.lpRate !== null) ? tisEx.lpRate : getLPRate(tiSupport);

            console.log('[IB Sync] FCLM TI Support - Actual:', data.transferInSupportRate, 'LP Rate:', data.transferInSupportPlanRate, 'row length:', tiSupport.length);

        }



        return data;

    }



    function detectFCLMColumns() {

        // FCLM has multi-row headers with colspans. Strategy:

        // 1. Find the group header row with "Labor Plan" to get its colspan start position

        // 2. OR find all "Rate" columns in detail header - 1st=Actual, 2nd=Plan, 3rd=LP

        // 3. Validate using the data: LP Rate should be after % to Plan column

        var result = { rateCol: null, lpRateCol: null, volCol: null, hrsCol: null, ptpCol: null };

        

        var table = document.querySelector('table.reportLayout, table[class*="report"]');

        if (!table) {

            var tables = document.querySelectorAll('table');

            for (var t = 0; t < tables.length; t++) {

                if (tables[t].textContent.indexOf('Inbound-TOTAL') !== -1 || tables[t].textContent.indexOf('Stow to Prime') !== -1) {

                    table = tables[t];

                    break;

                }

            }

        }

        if (!table) {

            console.log('[IB Sync] WARN: No FCLM table found for column detection');

            result.useFallback = true;

            return result;

        }

        

        var rows = table.querySelectorAll('tr');

        

        // Strategy A: Find "Labor Plan" group header and calculate absolute column positions

        var lpGroupAbsStart = -1;

        for (var r = 0; r < Math.min(5, rows.length); r++) {

            var cells = rows[r].querySelectorAll('th, td');

            var absPos = 0;

            for (var c = 0; c < cells.length; c++) {

                var text = cells[c].textContent.trim().toLowerCase();

                var colspan = parseInt(cells[c].getAttribute('colspan')) || 1;

                

                if (text.indexOf('labor plan') !== -1 || text === 'lp') {

                    lpGroupAbsStart = absPos;

                    console.log('[IB Sync] Found "Labor Plan" group at absolute position:', absPos, 'colspan:', colspan);

                }

                absPos += colspan;

            }

            if (lpGroupAbsStart !== -1) break;

        }

        

        // Strategy B: Find ALL "Rate" columns in detail rows using absolute positions

        var ratePositions = [];

        var volPositions = [];

        var hrsPositions = [];

        var ptpPositions = [];

        

        for (var r2 = 0; r2 < Math.min(5, rows.length); r2++) {

            var cells2 = rows[r2].querySelectorAll('th, td');

            var absPos2 = 0;

            for (var c2 = 0; c2 < cells2.length; c2++) {

                var text2 = cells2[c2].textContent.trim().toLowerCase();

                var colspan2 = parseInt(cells2[c2].getAttribute('colspan')) || 1;

                

                if (text2 === 'rate') ratePositions.push(absPos2);

                if (text2 === 'volume' || text2 === 'vol') volPositions.push(absPos2);

                if (text2 === 'hours' || text2 === 'hrs') hrsPositions.push(absPos2);

                if (text2.indexOf('% to p') !== -1 || text2 === '% plan' || text2 === 'ptp') ptpPositions.push(absPos2);

                if (text2 === 'lp rate') {

                    result.lpRateCol = absPos2;

                    console.log('[IB Sync] Found explicit "LP Rate" at absolute position:', absPos2);

                }

                

                absPos2 += colspan2;

            }

        }

        

        console.log('[IB Sync] All Rate positions found:', JSON.stringify(ratePositions));

        

        // Assign columns:

        // First Rate = Actual Rate

        if (ratePositions.length >= 1) result.rateCol = ratePositions[0];

        // LP Rate = last Rate position (3rd) OR from Labor Plan group, OR from explicit "LP Rate"

        if (result.lpRateCol === null) {

            if (ratePositions.length >= 3) {

                result.lpRateCol = ratePositions[2]; // Third "Rate" is LP Rate

            } else if (lpGroupAbsStart !== -1) {

                // LP Rate is the first Rate column at or after lpGroupAbsStart

                for (var rp = 0; rp < ratePositions.length; rp++) {

                    if (ratePositions[rp] >= lpGroupAbsStart) {

                        result.lpRateCol = ratePositions[rp];

                        break;

                    }

                }

            } else if (ratePositions.length >= 2) {

                // Fallback: assume last Rate is LP Rate

                result.lpRateCol = ratePositions[ratePositions.length - 1];

            }

        }

        

        // Vol = first Volume position

        if (volPositions.length >= 1) result.volCol = volPositions[0];

        // Hrs = first Hours position  

        if (hrsPositions.length >= 1) result.hrsCol = hrsPositions[0];

        // PTP = % to Plan (should be between Plan and LP sections)

        if (ptpPositions.length >= 1) result.ptpCol = ptpPositions[0];

        

        // Validate: LP Rate should always be > Rate + 3 (there are columns in between)

        if (result.lpRateCol !== null && result.rateCol !== null && result.lpRateCol <= result.rateCol + 2) {

            console.log('[IB Sync] WARN: LP Rate col (' + result.lpRateCol + ') too close to Rate col (' + result.rateCol + '), likely Plan Rate not LP Rate');

            // If we have 3+ rate positions, the last is LP Rate

            if (ratePositions.length >= 3) {

                result.lpRateCol = ratePositions[ratePositions.length - 1];

                console.log('[IB Sync] Corrected LP Rate to last Rate position:', result.lpRateCol);

            } else {

                // Can't determine LP Rate reliably, will need manual offset

                result.lpRateCol = null;

            }

        }

        

        // Final fallback: if LP Rate still not found, estimate from known FCLM layout

        // Typical: Rate is at col 7 (0-indexed), LP Rate is at col 12

        // Offset from Rate: LP Rate = Rate + 5 (Rate, PlanRate, PlanHrs, ΔPlan, %Plan, LPRate)

        if (result.lpRateCol === null && result.rateCol !== null) {

            result.lpRateCol = result.rateCol + 5;

            console.log('[IB Sync] Using estimated LP Rate position:', result.lpRateCol, '(Rate + 5)');

        }

        

        // OVERRIDE: LP Rate is in the "Labor Plan" section (far right of FCLM table)

        // Layout: ... | % to Plan | LP Rate | % to LP | Δ to LP | checkbox

        // LP Rate is 4th cell from end (before % to LP, Δ to LP, checkbox)

        // Also: Actual Rate is the FIRST Rate column, LP Rate is the LAST Rate-like column

        // LP Rate detection removed — using per-row end-counting in extraction instead

        // Actual Rate now uses per-row end-counting in extraction

        console.log('[IB Sync] Final columns - Rate:', result.rateCol, 'LP Rate:', result.lpRateCol, 

                    'Vol:', result.volCol, 'Hrs:', result.hrsCol, 'PTP:', result.ptpCol);

        return result;

    }





    // ==========================================

    // INTRO API FETCHERS (v2: direct JSON calls)

    // ==========================================

    function gmFetch(url, options) {

        options = options || {};

        var method = options.method || 'GET';

        var body = options.body || null;

        var headers = { 'Accept': 'application/json, text/plain, */*' };

        if (body) headers['Content-Type'] = 'application/json; charset=UTF-8';

        // Merge any custom headers passed in options

        if (options.headers) {

            Object.keys(options.headers).forEach(function(key) { headers[key] = options.headers[key]; });

        }



        return new Promise(function(resolve) {

            GM_xmlhttpRequest({

                method: method,

                url: url,

                data: body ? JSON.stringify(body) : undefined,

                responseType: 'json',

                headers: headers,

                onload: function(response) {

                    if (response.status >= 200 && response.status < 300) {

                        try {

                            var data = typeof response.response === 'string'

                                ? JSON.parse(response.response)

                                : response.response;

                            resolve({ ok: true, data: data });

                        } catch (e) {

                            console.error('[IB Sync] JSON parse error for ' + url, e);

                            resolve({ ok: false, error: 'Parse error', data: null });

                        }

                    } else {

                        console.error('[IB Sync] HTTP ' + response.status + ' for ' + url);

                        resolve({ ok: false, error: 'HTTP ' + response.status, data: null });

                    }

                },

                onerror: function(err) {

                    console.error('[IB Sync] Network error for ' + url, err);

                    resolve({ ok: false, error: 'Network error', data: null });

                },

                ontimeout: function() {

                    console.error('[IB Sync] Timeout for ' + url);

                    resolve({ ok: false, error: 'Timeout', data: null });

                }

            });

        });

    }



    function fetchINTROApis() {

        // Use FCLM-selected date, not wall clock

        var dateStr = getSelectedDate();

        var shiftType = getCurrentShift() === 'Day' ? 'DAY' : 'NIGHT';



        var promises = [

            // IB Flow: GetAppointments

            gmFetch(CONFIG.introApis.ibFlow, {

                method: 'POST',

                body: { warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.inboundflowdashboardservice.InboundFlowDashboardService.GetAppointments', 'Content-Encoding': 'amz-1.0' }

            }),

            // Trailer Planner: GetTrailersData

            gmFetch(CONFIG.introApis.trailerPlanner, {

                method: 'POST',

                body: { aggregationId: null, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introtrailerplanner.INTROTrailerPlanner.GetTrailersData', 'Content-Encoding': 'amz-1.0' }

            }),

            // Shift Planner call 1: GetShiftPlan (headcounts/plan)

            gmFetch(CONFIG.introApis.shiftPlanner, {

                method: 'POST',

                body: { date: dateStr, shiftType: shiftType, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introshiftplanner.INTROShiftPlanner.GetShiftPlan', 'Content-Encoding': 'amz-1.0' }

            }),

            // Shift Planner call 2: GetTargetMetrics (volume projections)

            gmFetch(CONFIG.introApis.shiftPlanner, {

                method: 'POST',

                body: { shiftDate: dateStr, shiftType: shiftType, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introshiftplanner.INTROShiftPlanner.GetTargetMetrics', 'Content-Encoding': 'amz-1.0' }

            }),

            // Trailer Planner call 2: GetLatestPublishedPlan (allocation summaries by period)

            gmFetch(CONFIG.introApis.trailerPlanner, {

                method: 'POST',

                body: { date: dateStr, shiftType: shiftType, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introtrailerplanner.INTROTrailerPlanner.GetLatestPublishedPlan', 'Content-Encoding': 'amz-1.0' }

            }),

            // Trailer Planner call 3: GetTargetVolume (Optimus/ALPS site target)

            gmFetch(CONFIG.introApis.trailerPlanner, {

                method: 'POST',

                body: { date: dateStr, shiftType: shiftType, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introtrailerplanner.INTROTrailerPlanner.GetTargetVolume', 'Content-Encoding': 'amz-1.0' }

            }),

            // Trailer Planner call 4: GetTargetVolume for OPPOSITE shift (to compute full-day S1 Goal)

            gmFetch(CONFIG.introApis.trailerPlanner, {

                method: 'POST',

                body: { date: dateStr, shiftType: shiftType === 'DAY' ? 'NIGHT' : 'DAY', warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.introtrailerplanner.INTROTrailerPlanner.GetTargetVolume', 'Content-Encoding': 'amz-1.0' }

            }),

            // IB Flow call 2: GetWarehouseSummary (Stow WIP from VL_STOW)

            gmFetch(CONFIG.introApis.ibFlow, {

                method: 'POST',

                body: { warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.inboundflowdashboardservice.InboundFlowDashboardService.GetWarehouseSummary', 'Content-Encoding': 'amz-1.0' }

            }),

            // Outbound Flow: GetShiftPerformance (IB Volume Actuals by period)

            gmFetch(CONFIG.introApis.outboundFlow, {

                method: 'POST',

                body: { date: dateStr, shift: shiftType, warehouseId: WAREHOUSE },

                headers: { 'X-Amz-Target': 'com.amazon.fmhoutboundflowservice.FMHOutboundFlowService.GetShiftPerformance', 'Content-Encoding': 'amz-1.0' }

            })

        ,

            // ALPs: S1 Capacity Forecast (Weekly plan)

            fetchALPsS1()];



        return Promise.allSettled(promises).then(function(results) {

            var ibFlowResult = results[0].status === 'fulfilled' ? results[0].value : { ok: false, data: null };

            var trailerResult = results[1].status === 'fulfilled' ? results[1].value : { ok: false, data: null };

            var shiftResult = results[2].status === 'fulfilled' ? results[2].value : { ok: false, data: null };

            var shiftVolResult = results[3].status === 'fulfilled' ? results[3].value : { ok: false, data: null };

            var trailerPlanResult = results[4].status === 'fulfilled' ? results[4].value : { ok: false, data: null };

            var targetVolResult = results[5].status === 'fulfilled' ? results[5].value : { ok: false, data: null };

            var targetVolOppositeResult = results[6].status === 'fulfilled' ? results[6].value : { ok: false, data: null };

            var stowWipResult = results[7].status === 'fulfilled' ? results[7].value : { ok: false, data: null };

            var shiftPerfResult = results[8].status === 'fulfilled' ? results[8].value : { ok: false, data: null };

            var alpsS1Result = results[9].status === 'fulfilled' ? results[9].value : null;



            console.log('[IB Sync] IB Flow API:', ibFlowResult.ok ? 'OK' : ibFlowResult.error);

            if (ibFlowResult.ok && ibFlowResult.data) console.log('[IB Sync] IB Flow raw:', JSON.stringify(ibFlowResult.data).substring(0, 300));

            if (ibFlowResult.ok && ibFlowResult.data) console.log('[IB Sync] IB Flow keys:', Object.keys(ibFlowResult.data.Output || ibFlowResult.data));

            console.log('[IB Sync] Trailer Planner API:', trailerResult.ok ? 'OK' : trailerResult.error);

            if (trailerResult.ok && trailerResult.data) console.log('[IB Sync] Trailer raw:', JSON.stringify(trailerResult.data).substring(0, 300));

            if (trailerResult.ok && trailerResult.data) console.log('[IB Sync] Trailer Planner keys:', Object.keys(trailerResult.data.Output || trailerResult.data));

            if (trailerResult.ok && trailerResult.data) {

                var tpRaw = trailerResult.data.Output || trailerResult.data;

                // Log ALL fields and their types to find allocation summaries

                Object.keys(tpRaw).forEach(function(k) {

                    var v = tpRaw[k];

                    if (v === null || typeof v !== 'object') {

                        console.log('[IB Sync] TP.' + k + ':', v);

                    } else if (Array.isArray(v)) {

                        console.log('[IB Sync] TP.' + k + ': Array(' + v.length + ')', v.length > 0 ? JSON.stringify(v[0]).substring(0, 200) : '');

                    } else {

                        console.log('[IB Sync] TP.' + k + ': Object', JSON.stringify(v).substring(0, 300));

                    }

                });

            }

            console.log('[IB Sync] Shift Planner (plan) API:', shiftResult.ok ? 'OK' : shiftResult.error);

            console.log('[IB Sync] Shift Planner (vol) API:', shiftVolResult.ok ? 'OK' : shiftVolResult.error);

            if (shiftResult.ok && shiftResult.data) console.log('[IB Sync] SP Plan raw:', JSON.stringify(shiftResult.data).substring(0, 500));

            if (shiftVolResult.ok && shiftVolResult.data) console.log('[IB Sync] SP Vol raw:', JSON.stringify(shiftVolResult.data).substring(0, 500));

            if (shiftResult.ok && shiftResult.data) {

                var spRaw = shiftResult.data.Output;

                console.log('[IB Sync] SP Output type:', typeof spRaw, spRaw ? (typeof spRaw === 'string' ? 'STRING: ' + spRaw.substring(0,100) : '') : 'NULL');

                var spData = spRaw || shiftResult.data;

                if (typeof spData === 'string') { try { spData = JSON.parse(spData); } catch(e) {} }

                console.log('[IB Sync] Shift Planner keys:', Object.keys(spData));

                console.log('[IB Sync] Shift Planner has recommendedShiftPlan:', !!spData.recommendedShiftPlan);

                console.log('[IB Sync] Shift Planner has combined:', !!spData.combined);

                if (spData.recommendedShiftPlan) console.log('[IB Sync] SP plan keys:', Object.keys(spData.recommendedShiftPlan));

                if (spData.recommendedShiftPlan && spData.recommendedShiftPlan.rateAndThroughput) {

                    console.log('[IB Sync] SP rateAndThroughput:', JSON.stringify(spData.recommendedShiftPlan.rateAndThroughput).substring(0, 500));

                }

                if (spData.combined) console.log('[IB Sync] SP combined[0]:', JSON.stringify(spData.combined[0]));

            }

            if (shiftVolResult.ok && shiftVolResult.data) {

                var svRaw = shiftVolResult.data.Output || shiftVolResult.data;

                if (typeof svRaw === 'string') { try { svRaw = JSON.parse(svRaw); } catch(e) {} }

                console.log('[IB Sync] SP Vol keys:', Object.keys(svRaw));

                console.log('[IB Sync] SP Vol has combined:', !!svRaw.combined);

                if (svRaw.combined) console.log('[IB Sync] SP Vol combined[0]:', JSON.stringify(svRaw.combined[0]));

                // Log all top-level scalar values to find Optimus Target

                Object.keys(svRaw).forEach(function(k) {

                    if (typeof svRaw[k] !== 'object' || svRaw[k] === null) console.log('[IB Sync] SP Vol.' + k + ':', svRaw[k]);

                });

            }

            if (shiftResult.ok && shiftResult.data) {

                var planRaw = shiftResult.data.Output || shiftResult.data;

                if (typeof planRaw === 'string') { try { planRaw = JSON.parse(planRaw); } catch(e) {} }

                // Log all top-level scalar values to find Optimus Target

                Object.keys(svRaw).forEach(function(k) {

                    if (typeof svRaw[k] !== 'object' || svRaw[k] === null) console.log('[IB Sync] SP Vol.' + k + ':', svRaw[k]);

                });

            }

            if (shiftResult.ok && shiftResult.data) {

                var planRaw = shiftResult.data.Output || shiftResult.data;

                if (typeof planRaw === 'string') { try { planRaw = JSON.parse(planRaw); } catch(e) {} }

                // Log all top-level keys from plan response to find optimusTarget/siteTarget

                console.log('[IB Sync] SP Plan ALL keys:', Object.keys(planRaw));

                // Check for target-like fields

                ['optimusTarget', 'siteTarget', 'alpsSiteTarget', 'target', 'plannedVolume', 'shiftTarget'].forEach(function(field) {

                    if (planRaw[field] !== undefined) console.log('[IB Sync] SP Plan.' + field + ':', planRaw[field]);

                });

            }



            return {

                ibFlow: ibFlowResult.ok ? (function() {

                    var d = ibFlowResult.data.Output || ibFlowResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                trailerPlanner: trailerResult.ok ? (function() {

                    var d = trailerResult.data.Output || trailerResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                shiftPlanner: (function() {

                    // Merge both Shift Planner responses: plan (headcounts) + volumes

                    var plan = null;

                    var vol = null;

                    if (shiftResult.ok && shiftResult.data) {

                        plan = shiftResult.data.Output || shiftResult.data;

                        if (typeof plan === 'string') { try { plan = JSON.parse(plan); } catch(e) {} }

                    }

                    if (shiftVolResult.ok && shiftVolResult.data) {

                        vol = shiftVolResult.data.Output || shiftVolResult.data;

                        if (typeof vol === 'string') { try { vol = JSON.parse(vol); } catch(e) {} }

                    }

                    // If plan has recommendedShiftPlan and vol has combined, merge them

                    if (plan && vol && vol.combined) {

                        plan.combined = vol.combined;

                        plan.stow = vol.stow;

                        plan.decant = vol.decant;

                        plan.indirect = vol.indirect;

                        plan.stowToPrime = vol.stowToPrime;

                        return plan;

                    }

                    // If only one worked, return whichever has data

                    if (plan && plan.recommendedShiftPlan) return plan;

                    if (vol && vol.combined) return vol;

                    // Fallback: return plan or vol or null

                    return plan || vol || null;

                })(),

                // Trailer allocation plan (period summaries)

                trailerAllocationPlan: trailerPlanResult.ok ? (function() {

                    var d = trailerPlanResult.data.Output || trailerPlanResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                // Target volume (Optimus/ALPS target)

                targetVolume: targetVolResult.ok ? (function() {

                    var d = targetVolResult.data.Output || targetVolResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                // Target volume for OPPOSITE shift (for full-day S1 Goal)

                targetVolumeOpposite: targetVolOppositeResult.ok ? (function() {

                    var d = targetVolOppositeResult.data.Output || targetVolOppositeResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                // Stow WIP from GetWarehouseSummary

                stowWipData: stowWipResult.ok ? (function() {

                    var d = stowWipResult.data.Output || stowWipResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                // Shift Performance (IB Volume Actuals)

                shiftPerformance: shiftPerfResult.ok ? (function() {

                    var d = shiftPerfResult.data.Output || shiftPerfResult.data;

                    if (typeof d === 'string') { try { d = JSON.parse(d); } catch(e) {} }

                    return d;

                })() : null,

                alpsS1: alpsS1Result,

            };

        });

    }



    // ==========================================

    // INTRO DATA PARSERS

    // ==========================================

    function parseIBFlowData(raw, stowWipData) {

        var result = {

            wip: null,

            psPiles: 0,

            remainingCUTs: 0,

            cutCompliance: null,

            totalTrailersOnShift: 0,

            completedOnShift: 0,

        };



        if (!raw) return result;

        var appointments = raw.appointmentList || (raw.data && raw.data.appointmentList) || [];

        if (!appointments.length) return result;



        var shiftBounds = getShiftBounds();

        var startMs = shiftBounds.start.getTime() / 1000;

        var endMs = shiftBounds.end.getTime() / 1000;



        var stowWipItems = 0;

        var psPiles = 0;

        var cutsDue = 0;

        var cutsCompleted = 0;



        appointments.forEach(function(appt) {



            // Problem Solve piles: items at PS* locations

            if (appt.location && appt.location.indexOf('PS') === 0) {

                psPiles++;

            }



            // CUT trailers due on shift

            var cutTime = appt.criticalUnloadTime || null;

            if (cutTime && cutTime >= startMs && cutTime <= endMs) {

                cutsDue++;

                if (appt.status === 'COMPLETED') {

                    cutsCompleted++;

                }

            }

        });



        // Stow WIP from GetChildLocationInfo (VL_STOW location ITEM count)

        if (stowWipData && stowWipData.virtualLocationsSummary) {

            stowWipData.virtualLocationsSummary.forEach(function(loc) {

                if (loc.locationId === 'VL_STOW' && loc.quantities) {

                    loc.quantities.forEach(function(q) {

                        if (q.unit === 'ITEM') stowWipItems = q.count;

                    });

                }

            });

        }

        result.wip = stowWipItems;

        result.psPiles = psPiles;

        result.totalTrailersOnShift = cutsDue;

        result.completedOnShift = cutsCompleted;

        result.remainingCUTs = cutsDue - cutsCompleted;

        result.cutCompliance = cutsDue > 0 ? Math.round((cutsCompleted / cutsDue) * 100) : null;



        return result;

    }



    function parseTrailerPlannerData(raw, allocationPlan, targetVolumeData, targetVolumeOpposite) {

        var result = {

            trailerPlanVolume: null,

            trailerPlanByPeriod: { P1: null, P2: null, P3: null },

            optimusTarget: null,

            s1Goal: null,

            fluidPercent: null,

            casePercent: null,

            smallsPercent: null,

            fluidByPeriod: { P1: null, P2: null, P3: null },

            caseByPeriod: { P1: null, P2: null, P3: null },

            smallsByPeriod: { P1: null, P2: null, P3: null },

        };



        // Get Optimus/ALPS target from GetTargetVolume response

        if (targetVolumeData && targetVolumeData.targetVolume) {

            result.optimusTarget = Math.round(targetVolumeData.targetVolume);

            console.log('[IB Sync] Optimus Target:', result.optimusTarget);

            // S1 Goal = current shift target + opposite shift target (full day)

            var oppositeTarget = (targetVolumeOpposite && targetVolumeOpposite.targetVolume) ? Math.round(targetVolumeOpposite.targetVolume) : 0;

            result.s1Goal = result.optimusTarget + oppositeTarget;

            console.log('[IB Sync] S1 Goal (full day):', result.s1Goal, '= current:', result.optimusTarget, '+ opposite:', oppositeTarget);

            

        }



        // Get allocation summary from GetLatestPublishedPlan response

        // Response structure: { statusCode, trailerAllocationPlan: { allocationSummaryByPeriodList: [...] } }

        var allocData = allocationPlan && allocationPlan.trailerAllocationPlan ? allocationPlan.trailerAllocationPlan : allocationPlan;

        if (allocData && allocData.allocationSummaryByPeriodList) {

            var summaries = allocData.allocationSummaryByPeriodList;

            summaries.forEach(function(s) {

                var as = s.allocationSummary || s;

                var ums = as.unitMixSummary || {};

                var totalUnits = Math.round(ums.totalUnitCount || 0);

                var smallUnits = Math.round(ums.smallUnitCount || 0);

                var mediumUnits = Math.round(ums.mediumUnitCount || 0);

                var cs = as.casesSummary || {};

                var ts = as.totesSummary || {};



                // Fluid % = total fluid / (total fluid + total palletized)

                // Fluid = casesSummary.fluidUnitCount + totesSummary.fluidUnitCount

                // Palletized = totesSummary.palletizedCount + casesSummary.palletizedMixedCount + casesSummary.palletizedSapCount

                var caseFluid = cs.fluidUnitCount || 0;

                var toteFluid = ts.fluidUnitCount || 0;

                var totalFluid = Math.round(caseFluid + toteFluid);

                var totePalletized = ts.palletizedCount || 0;

                var casePalletized = (cs.palletizedMixedCount || 0) + (cs.palletizedSapCount || 0);

                var totalPalletized = Math.round(totePalletized + casePalletized);

                var fluidDenom = totalFluid + totalPalletized;

                var fluidPct = fluidDenom > 0 ? Math.round((totalFluid / fluidDenom) * 1000) / 10 : null;



                // Case % from containerSummary.containerPercentages.cases (pre-calculated by INTRO)

                var contSum = as.containerSummary || {};

                var contPct = contSum.containerPercentages || {};

                var casePct = contPct.cases !== undefined ? Math.round(contPct.cases) : null;



                // Smalls Mix % = small / (small + medium + large)

                var mixTotal = smallUnits + mediumUnits;

                var smallsPct = mixTotal > 0 ? Math.round((smallUnits / mixTotal) * 100) : null;

                

                var periodName = s.periodName || '';

                if (periodName === 'Period_1') {

                    result.trailerPlanByPeriod.P1 = totalUnits;

                    result.fluidByPeriod = result.fluidByPeriod || {};

                    result.fluidByPeriod.P1 = fluidPct;

                    result.caseByPeriod = result.caseByPeriod || {};

                    result.caseByPeriod.P1 = casePct;

                    result.smallsByPeriod = result.smallsByPeriod || {};

                    result.smallsByPeriod.P1 = smallsPct;

                } else if (periodName === 'Period_2') {

                    result.trailerPlanByPeriod.P2 = totalUnits;

                    result.fluidByPeriod = result.fluidByPeriod || {};

                    result.fluidByPeriod.P2 = fluidPct;

                    result.caseByPeriod = result.caseByPeriod || {};

                    result.caseByPeriod.P2 = casePct;

                    result.smallsByPeriod = result.smallsByPeriod || {};

                    result.smallsByPeriod.P2 = smallsPct;

                } else if (periodName === 'Period_3') {

                    result.trailerPlanByPeriod.P3 = totalUnits;

                    result.fluidByPeriod = result.fluidByPeriod || {};

                    result.fluidByPeriod.P3 = fluidPct;

                    result.caseByPeriod = result.caseByPeriod || {};

                    result.caseByPeriod.P3 = casePct;

                    result.smallsByPeriod = result.smallsByPeriod || {};

                    result.smallsByPeriod.P3 = smallsPct;

                } else if (periodName === 'Shift_Summary') {

                    result.trailerPlanVolume = totalUnits;

                    result.fluidPercent = fluidPct;

                    result.casePercent = casePct;

                    result.smallsPercent = smallsPct;

                }

            });

            console.log('[IB Sync] Trailer Plan Volume:', result.trailerPlanVolume, 'P1:', result.trailerPlanByPeriod.P1, 'P2:', result.trailerPlanByPeriod.P2, 'P3:', result.trailerPlanByPeriod.P3);

            console.log('[IB Sync] Fluid%:', result.fluidPercent, 'Case%:', result.casePercent, 'Smalls%:', result.smallsPercent);

        } else {

            console.log('[IB Sync] No allocation summary data found');

        }



        // Fluid/Case % from allocated trailers in plan (17 trailers) or raw data

        var trailers = (allocData && allocData.trailersData) || (raw && raw.trailersData) || [];

        if (trailers.length > 0) {

            var fluidUnits = 0, caseUnits = 0, totalUnits = 0;

            trailers.forEach(function(t) {

                totalUnits += t.totalUnits || 0;

                fluidUnits += (t.fluidCaseUnits || 0) + (t.fluidToteUnits || 0);

                caseUnits += t.totalCaseUnits || 0;

            });

            if (totalUnits > 0) {

                result.fluidPercent = Math.round((fluidUnits / totalUnits) * 100);

                result.casePercent = Math.round((caseUnits / totalUnits) * 100);

            }

        }



        console.log('[IB Sync] Trailer Plan Volume:', result.trailerPlanVolume, 'P1:', result.trailerPlanByPeriod.P1, 'P2:', result.trailerPlanByPeriod.P2, 'P3:', result.trailerPlanByPeriod.P3);

        return result;

    }



    function parseShiftPerformance(raw) {

        var result = { P1: null, P2: null, P3: null, EOS: null, rates: { P1: null, P2: null, P3: null, EOS: null } };

        if (!raw || !raw.ibData || !raw.ibData.periodBreakdown) return result;



        var breakdown = raw.ibData.periodBreakdown;

        var shiftTotal = 0;

        var shiftHours = 0;



        // First, log all process paths from first hour of PERIOD_1 to identify the right one

        var firstPeriod = breakdown['PERIOD_1'];

        if (firstPeriod) {

            var firstHourKey = Object.keys(firstPeriod)[0];

            if (firstHourKey && firstPeriod[firstHourKey] && firstPeriod[firstHourKey].directLabor) {

                var processIds = firstPeriod[firstHourKey].directLabor.map(function(l) {

                    return (l.id ? l.id.process + ' | ' + l.id.processPath : 'unknown');

                });

                console.log('[IB Sync] ShiftPerf directLabor processes:', processIds.join(', '));

            }

        }



        ['PERIOD_1', 'PERIOD_2', 'PERIOD_3'].forEach(function(periodKey) {

            var periodData = breakdown[periodKey];

            if (!periodData) return;



            var periodTotal = 0;

            var periodHours = 0;

            // periodData is an object keyed by hourly timestamps

            Object.keys(periodData).forEach(function(hourKey) {

                var hourEntry = periodData[hourKey];

                if (!hourEntry || !hourEntry.directLabor) return;

                // Only sum entries where processPath contains "TRANSFER_IN" or "STOW_TO_PRIME" at top level

                // OR look for a process that matches "INBOUND" / "IB_TOTAL" / "TRANSFER_IN (Total)"

                hourEntry.directLabor.forEach(function(labor) {

                    if (!labor.id || !labor.metrics || !labor.metrics.volume) return;

                    var proc = labor.id.process || '';

                    var path = labor.id.processPath || '';

                    // IB_TOTAL is the top-level inbound total (avoids double-counting sub-processes)

                    if (proc === 'IB_TOTAL') {

                        if (labor.metrics.volume.actual) {

                            periodTotal += labor.metrics.volume.actual;

                        }

                        if (labor.metrics.hours && labor.metrics.hours.actual) {

                            periodHours += labor.metrics.hours.actual;

                        }

                    }

                });

            });



            var pKey = periodKey === 'PERIOD_1' ? 'P1' : periodKey === 'PERIOD_2' ? 'P2' : 'P3';

            result[pKey] = Math.round(periodTotal);

            shiftTotal += periodTotal;

            shiftHours += periodHours;

            // Rate = volume / hours for the period

            result.rates[pKey] = periodHours > 0 ? Math.round((periodTotal / periodHours) * 100) / 100 : null;

        });



        result.EOS = Math.round(shiftTotal);

        result.rates.EOS = shiftHours > 0 ? Math.round((shiftTotal / shiftHours) * 100) / 100 : null;

        console.log('[IB Sync] ShiftPerf rates - P1:', result.rates.P1, 'P2:', result.rates.P2, 'P3:', result.rates.P3, 'EOS:', result.rates.EOS);

        console.log('[IB Sync] Volume Actuals - P1:', result.P1, 'P2:', result.P2, 'P3:', result.P3, 'EOS:', result.EOS);

        return result;

    }



        function parseShiftPlannerData(raw) {

        var result = {

            // Headcounts by period

            stowHC: { P1: null, P2: null, P3: null },

            // Optimus target (ALPS/Site target)

            optimusTarget: null,

            stowHCByFloor: {

                A02: { P1: null, P2: null, P3: null },

                A03: { P1: null, P2: null, P3: null },

                A04: { P1: null, P2: null, P3: null },

                A05: { P1: null, P2: null, P3: null },

            },

            decantHC: { P1: null, P2: null, P3: null },

            problemSolveHC: { P1: null, P2: null, P3: null, EOS: null },

            indirectHC: { P1: null, P2: null, P3: null },

            waterSpiderHC: { P1: null, P2: null, P3: null },

            waterSpiderHC: { P1: null, P2: null, P3: null, EOS: null },

            // Volume projections

            ibVolumePlan: null,

            volumeByPeriod: { P1: null, P2: null, P3: null },

            volumeActualByPeriod: { P1: null, P2: null, P3: null, EOS: null },

            volumePlannedByPeriod: { P1: null, P2: null, P3: null, EOS: null },

            // TPH

            ibTPH: { P1: null, P2: null, P3: null },

            decantTPH: { P1: null, P2: null, P3: null },

            // Hours (indirect breakdown)

            transferInSupportHrs: { P1: null, P2: null, P3: null },

            stowToPrimeSupportHrs: { P1: null, P2: null, P3: null },

            ibLeadPAHrs: { P1: null, P2: null, P3: null },

            ibProblemSolveHrs: { P1: null, P2: null, P3: null },

            // Smalls mix

            smallsMixPercent: { P1: null, P2: null, P3: null },

            // Indirect spend

            indirectSpendPercent: { P1: null, P2: null, P3: null },

            // Rate overrides

            rateOverrides: {},

            // Plan creation timestamp

            planCreatedAt: null,

        };



        if (!raw) return result;



        var floorMap = { 'FLOOR_2': 'A02', 'FLOOR_3': 'A03', 'FLOOR_4': 'A04', 'FLOOR_5': 'A05' };

        var periodMap = { 'PERIOD_1': 'P1', 'PERIOD_2': 'P2', 'PERIOD_3': 'P3', 'SHIFT_TOTAL': 'EOS' };



        // Parse headcounts from recommendedShiftPlan (search at various depths)

        var plan = raw.recommendedShiftPlan;

        if (!plan && raw.data && raw.data.recommendedShiftPlan) plan = raw.data.recommendedShiftPlan;

        if (!plan && raw.body && raw.body.recommendedShiftPlan) plan = raw.body.recommendedShiftPlan;

        if (!plan) {

            // Search all top-level keys for recommendedShiftPlan

            Object.keys(raw).forEach(function(key) {

                if (raw[key] && typeof raw[key] === 'object' && raw[key].recommendedShiftPlan) {

                    plan = raw[key].recommendedShiftPlan;

                }

            });

        }

        if (plan && plan.headcounts) {

            // Direct headcounts

            if (plan.headcounts.direct) {

                plan.headcounts.direct.forEach(function(entry) {

                    if (entry.processName === 'STOW' && entry.subProcessName === 'STOW') {

                        // Stow HC per floor per period

                        (entry.headcounts || []).forEach(function(hc) {

                            var floor = floorMap[hc.floorName];

                            var period = periodMap[hc.periodName];

                            if (floor && period) {

                                result.stowHCByFloor[floor][period] = hc.headcountValue;

                            }

                        });

                        // Sum across floors for total stow HC per period

                        ['P1', 'P2', 'P3', 'EOS'].forEach(function(p) {

                            var total = 0;

                            var hasVal = false;

                            Object.keys(result.stowHCByFloor).forEach(function(f) {

                                if (result.stowHCByFloor[f][p] !== null) {

                                    total += result.stowHCByFloor[f][p];

                                    hasVal = true;

                                }

                            });

                            if (hasVal) result.stowHC[p] = total;

                        });

                    }

                    if (entry.processName === 'DECANT') {

                        // Sum all DECANT sub-processes (PREDICANT + SAP_DECANT + TASRS_DECANT)

                        (entry.headcounts || []).forEach(function(hc) {

                            var period = periodMap[hc.periodName];

                            if (period && hc.headcountValue) {

                                result.decantHC[period] = (result.decantHC[period] || 0) + hc.headcountValue;

                            }

                        });

                    }

                });

            }



            // Indirect headcounts

            if (plan.headcounts.indirect) {

                plan.headcounts.indirect.forEach(function(entry) {

                    (entry.headcounts || []).forEach(function(hc) {

                        var period = periodMap[hc.periodName];

                        if (period && hc.headcountValue) {

                            result.indirectHC[period] = (result.indirectHC[period] || 0) + hc.headcountValue;

                        }

                    });

                    // Water Spider specifically

                    if (entry.roleName === 'Stow Water Spider') {

                        (entry.headcounts || []).forEach(function(hc) {

                            var period = periodMap[hc.periodName];

                            if (period) result.waterSpiderHC[period] = hc.headcountValue;

                        });

                    }

                    // Stow Problem Solve

                    if (entry.roleName === 'Stow Problem Solve') {

                        (entry.headcounts || []).forEach(function(hc) {

                            var period = periodMap[hc.periodName];

                            if (period) result.problemSolveHC[period] = hc.headcountValue;

                        });

                    }

                });

            }

        }



        // Parse wipBuffers for Planned volumes per period (processedStowVolume)

        if (plan && plan.wipBuffers) {

            var shiftPlannedTotal = 0;

            plan.wipBuffers.forEach(function(wb) {

                var pName = periodMap[wb.periodName];

                if (pName && wb.processedStowVolume) {

                    result.volumePlannedByPeriod[pName] = Math.round(wb.processedStowVolume);

                    shiftPlannedTotal += wb.processedStowVolume;

                }

            });

            result.volumePlannedByPeriod.EOS = Math.round(shiftPlannedTotal);

            console.log('[IB Sync] Planned Volumes from wipBuffers - P1:', result.volumePlannedByPeriod.P1, 'P2:', result.volumePlannedByPeriod.P2, 'P3:', result.volumePlannedByPeriod.P3, 'EOS:', result.volumePlannedByPeriod.EOS);

        }



        // Parse rate overrides

        if (raw.planOverrides && raw.planOverrides.ratesOverrides && raw.planOverrides.ratesOverrides.rateOverridesByProcess) {

            raw.planOverrides.ratesOverrides.rateOverridesByProcess.forEach(function(ro) {

                result.rateOverrides[ro.functionName] = {};

                (ro.periodOverrides || []).forEach(function(po) {

                    var period = periodMap[po.period.periodName];

                    if (period) result.rateOverrides[ro.functionName][period] = po.overrideValue;

                });

            });

        }



        // Parse volume projections (Response B structure - may be in same or separate response)

        // Search for combined array at various nesting levels

        var volData = null;

        if (raw.combined) {

            volData = raw;

        } else if (raw.volumeProjections) {

            volData = raw.volumeProjections;

        } else if (raw.data && raw.data.combined) {

            volData = raw.data;

        } else if (raw.body && raw.body.combined) {

            volData = raw.body;

        } else if (raw.result && raw.result.combined) {

            volData = raw.result;

        } else {

            // Deep search: look for any key containing combined array

            Object.keys(raw).forEach(function(key) {

                if (raw[key] && typeof raw[key] === 'object' && raw[key].combined) {

                    volData = raw[key];

                }

            });

        }

        if (!volData) volData = raw;



        if (volData && volData.combined) {

            volData.combined.forEach(function(entry) {

                if (entry.processPath === 'INBOUND_TOTAL') {

                    if (entry.interval === 'SHIFT_TOTAL') {

                        result.ibVolumePlan = Math.round(entry.volume);

                    }

                    var period = periodMap[entry.interval];

                    if (period) {

                        result.volumeByPeriod[period] = Math.round(entry.volume);

                        result.ibTPH[period] = Math.round(entry.tph);

                    }

                }

            });

        }



        if (volData && volData.decant) {

            volData.decant.forEach(function(entry) {

                if (entry.processPath === 'TRANSFER_IN_DECANT') {

                    var period = periodMap[entry.interval];

                    if (period) {

                        result.decantTPH[period] = Math.round(entry.tph);

                    }

                }

            });

        }



        if (volData && volData.stow) {

            // Smalls mix: EACH_TRANSFER_IN_SMALL / EACH_TRANSFER_IN

            var smallVol = {};

            var totalVol = {};

            volData.stow.forEach(function(entry) {

                var period = periodMap[entry.interval];

                if (!period) return;

                if (entry.processPath === 'EACH_TRANSFER_IN_SMALL') {

                    smallVol[period] = entry.volume;

                }

                if (entry.processPath === 'EACH_TRANSFER_IN') {

                    totalVol[period] = entry.volume;

                }

            });

            ['P1', 'P2', 'P3', 'EOS'].forEach(function(p) {

                if (smallVol[p] && totalVol[p]) {

                    result.smallsMixPercent[p] = Math.round((smallVol[p] / totalVol[p]) * 100);

                }

            });

        }



        if (volData && volData.indirect) {

            volData.indirect.forEach(function(entry) {

                var period = periodMap[entry.interval];

                if (!period) return;



                if (entry.processPath === 'TRANSFER_IN_SUPPORT') {

                    result.transferInSupportHrs[period] = Math.round(entry.hours * 10) / 10;

                }

                if (entry.processPath === 'STOW_TO_PRIME_SUPPORT') {

                    result.stowToPrimeSupportHrs[period] = Math.round(entry.hours * 10) / 10;

                }

                if (entry.processPath === 'IB_LEAD_PA') {

                    result.ibLeadPAHrs[period] = Math.round(entry.hours * 10) / 10;

                }

                if (entry.processPath === 'IB_PROBLEM_SOLVE') {

                    result.ibProblemSolveHrs[period] = Math.round(entry.hours * 10) / 10;

                }

            });



            // Indirect spend % = indirect hours / total hours per period

            if (volData.combined) {

                volData.combined.forEach(function(entry) {

                    if (entry.processPath === 'INBOUND_TOTAL') {

                        var period = periodMap[entry.interval];

                        if (period && entry.hours > 0) {

                            var indirectHrs = (result.transferInSupportHrs[period] || 0) +

                                (result.stowToPrimeSupportHrs[period] || 0) +

                                (result.ibLeadPAHrs[period] || 0) +

                                (result.ibProblemSolveHrs[period] || 0);

                            result.indirectSpendPercent[period] = Math.round((indirectHrs / entry.hours) * 100);

                        }

                    }

                });

            }

        }



        // Extract planCreatedAt from the raw response

        var createdAt = raw.createdAt || (raw.data && raw.data.createdAt) || null;

        if (!createdAt && plan) createdAt = raw.planCreatedAt || null;

        if (createdAt) {

            result.planCreatedAt = createdAt;

            console.log('[IB Sync] Shift Plan createdAt:', createdAt);

        }



        return result;

    }



    // ==========================================

    // APOLLO DATA FETCHER

    // ==========================================



    // Get the FCLM-selected date from the URL (reuses shared utility)

    function getFCLMDate() {

        return getSelectedDate();

    }



    // Build period time boundaries for Apollo filtering

    function getApolloPeriodRanges() {

        var dateStr = getFCLMDate();

        var shift = getCurrentShift();



        if (shift === 'Day') {

            return {

                P1: { start: dateStr + '+06%3A00', end: dateStr + '+10%3A00' },

                P2: { start: dateStr + '+10%3A00', end: dateStr + '+14%3A00' },

                P3: { start: dateStr + '+14%3A00', end: dateStr + '+18%3A00' },

                EOS: { start: dateStr + '+06%3A00', end: dateStr + '+18%3A00' }

            };

        } else {

            // Night shift spans two calendar days

            var nextDate = new Date(dateStr + 'T00:00:00');

            nextDate.setDate(nextDate.getDate() + 1);

            var nextDateStr = nextDate.toISOString().split('T')[0];

            return {

                P1: { start: dateStr + '+18%3A00', end: dateStr + '+22%3A00' },

                P2: { start: dateStr + '+22%3A00', end: nextDateStr + '+02%3A00' },

                P3: { start: nextDateStr + '+02%3A00', end: nextDateStr + '+06%3A00' },

                EOS: { start: dateStr + '+18%3A00', end: nextDateStr + '+06%3A00' }

            };

        }

    }



    // Fetch Apollo audit_execution_metrics HTML for a given time range

    function fetchApolloForPeriod(startDate, endDate) {

        return new Promise(function(resolve) {

            var url = 'https://apollo-audit.corp.amazon.com/reporting/audit_execution_metrics?utf8=%E2%9C%93&department=' + SITE_SETTINGS.APOLLO_DEPARTMENT + '&start_date=' + startDate + '&end_date=' + endDate + '&commit=Search';



            console.log('[IB Sync] Apollo fetch:', url);



            GM_xmlhttpRequest({

                method: 'GET',

                url: url,

                onload: function(response) {

                    if (response.status === 200) {

                        var results = parseApolloHTML(response.responseText);

                        resolve(results);

                    } else {

                        console.log('[IB Sync] Apollo error status:', response.status);

                        resolve({});

                    }

                },

                onerror: function() {

                    console.log('[IB Sync] Apollo network error');

                    resolve({});

                }

            });

        });

    }



    // Parse Apollo audit_execution_metrics HTML table

    function parseApolloHTML(html) {

        var parser = new DOMParser();

        var doc = parser.parseFromString(html, 'text/html');

        var results = {};



        // Find the audit table (has headers: Audit Name, Req. Frequency, Suggested, Completed by Suggested, Total Completed, NC's, Overdue, Hours)

        var tables = doc.querySelectorAll('table');

        var auditTable = null;

        for (var t = 0; t < tables.length; t++) {

            if (tables[t].textContent.indexOf('Total Completed') !== -1) {

                auditTable = tables[t];

                break;

            }

        }



        if (!auditTable) {

            console.log('[IB Sync] Apollo: No audit table found in HTML response');

            return results;

        }



        // Find "Total Completed" column index

        var headerRow = auditTable.querySelector('tr');

        var headers = headerRow ? headerRow.querySelectorAll('th, td') : [];

        var totalCompletedCol = -1;

        for (var h = 0; h < headers.length; h++) {

            if (headers[h].textContent.trim() === 'Total Completed') {

                totalCompletedCol = h;

                break;

            }

        }



        if (totalCompletedCol === -1) {

            console.log('[IB Sync] Apollo: Could not find "Total Completed" column');

            return results;

        }



        // Parse each row

        var rows = auditTable.querySelectorAll('tr');

        for (var r = 1; r < rows.length; r++) {

            var cells = rows[r].querySelectorAll('td');

            if (cells.length <= totalCompletedCol) continue;

            var auditName = cells[0] ? cells[0].textContent.trim() : '';

            var totalCompleted = cells[totalCompletedCol] ? parseInt(cells[totalCompletedCol].textContent.trim()) || 0 : 0;

            if (auditName) results[auditName] = totalCompleted;

        }



        console.log('[IB Sync] Apollo parsed audits:', JSON.stringify(results));

        return results;

    }



    function fetchAllApolloData() {

        var periods = getApolloPeriodRanges();



        // Fetch all 4 periods in parallel (P1, P2, P3, EOS)

        return Promise.allSettled([

            fetchApolloForPeriod(periods.P1.start, periods.P1.end),

            fetchApolloForPeriod(periods.P2.start, periods.P2.end),

            fetchApolloForPeriod(periods.P3.start, periods.P3.end),

            fetchApolloForPeriod(periods.EOS.start, periods.EOS.end)

        ]).then(function(results) {

            var p1Data = results[0].status === 'fulfilled' ? results[0].value : {};

            var p2Data = results[1].status === 'fulfilled' ? results[1].value : {};

            var p3Data = results[2].status === 'fulfilled' ? results[2].value : {};

            var eosData = results[3].status === 'fulfilled' ? results[3].value : {};



            // Map Apollo audit names to board names with period data

            return CONFIG.apolloAudits.map(function(audit) {

                var apolloName = audit.apolloName;

                if (!apolloName) return { name: audit.name, p1: null, p2: null, p3: null, eos: null, count: 0 };

                return {

                    name: audit.name,

                    p1: p1Data[apolloName] !== undefined ? p1Data[apolloName] : null,

                    p2: p2Data[apolloName] !== undefined ? p2Data[apolloName] : null,

                    p3: p3Data[apolloName] !== undefined ? p3Data[apolloName] : null,

                    eos: eosData[apolloName] !== undefined ? eosData[apolloName] : null,

                    count: eosData[apolloName] || 0

                };

            });

        });

    }



    // ==========================================

    // SLIM DATA FETCHER

    // ==========================================

    function fetchSLIMData() {

        return new Promise(function(resolve) {

            var url = 'https://slim.corp.amazon.com/dashboard_loaded?' + WAREHOUSE;



            console.log('[IB Sync] SLIM fetch:', url);



            GM_xmlhttpRequest({

                method: 'GET',

                url: url,

                onload: function(response) {

                    if (response.status === 200) {

                        var result = parseSLIMHTML(response.responseText);

                        resolve(result);

                    } else {

                        console.log('[IB Sync] SLIM error status:', response.status);

                        resolve({ receiveAutoDecant: null, receiveManualDecant: null });

                    }

                },

                onerror: function() {

                    console.log('[IB Sync] SLIM network error');

                    resolve({ receiveAutoDecant: null, receiveManualDecant: null });

                }

            });

        });

    }



    function parseSLIMHTML(html) {

        var result = { receiveAutoDecant: null, receiveManualDecant: null };



        // Try multiple regex patterns to match SLIM profile scores

        // Pattern 1: "Receive Auto Decant - 100.0%" (plain text in Profiles header)

        // Pattern 2: ">Receive Auto Decant</a>...100.0%" (in linked profile entries)

        // Pattern 3: "Receive Auto Decant" followed later by a percentage

        var autoMatch = html.match(/Receive\s*Auto\s*Decant\s*[-–—]\s*([\d.]+)\s*%/i);

        if (!autoMatch) autoMatch = html.match(/Receive\s*Auto\s*Decant[^%]*?([\d.]+)\s*%/i);

        if (autoMatch) result.receiveAutoDecant = parseFloat(autoMatch[1]);



        var manualMatch = html.match(/Receive\s*Manual\s*Decant\s*[-–—]\s*([\d.]+)\s*%/i);

        if (!manualMatch) manualMatch = html.match(/Receive\s*Manual\s*Decant[^%]*?([\d.]+)\s*%/i);

        if (manualMatch) result.receiveManualDecant = parseFloat(manualMatch[1]);



        // Debug: dump a snippet around "Receive Auto Decant" to see format

        var idx = html.indexOf('Receive Auto Decant');

        if (idx !== -1) {

            console.log('[IB Sync] SLIM raw near Auto Decant:', html.substring(idx, idx + 200));

        } else {

            console.log('[IB Sync] SLIM: "Receive Auto Decant" not found in HTML');

        }



        console.log('[IB Sync] SLIM parsed - Auto Decant:', result.receiveAutoDecant, '%, Manual Decant:', result.receiveManualDecant, '%');

        return result;

    }



    // ==========================================

    // ATLAS DATA FETCHER (Quality DPMOs)

    // ==========================================

    function fetchAtlasData() {

        var dateStr = getSelectedDate();

        var shift = getCurrentShift();

        var startTime, endTime;



        if (shift === 'Day') {

            startTime = dateStr + 'T06:00:00';

            endTime = dateStr + 'T18:00:00';

        } else {

            startTime = dateStr + 'T18:00:00';

            var nextDate = new Date(dateStr + 'T00:00:00');

            nextDate.setDate(nextDate.getDate() + 1);

            var nextDateStr = nextDate.toISOString().split('T')[0];

            endTime = nextDateStr + 'T06:00:00';

        }



        // Convert to epoch seconds

        var startEpoch = Math.floor(new Date(startTime).getTime() / 1000);

        var endEpoch = Math.floor(new Date(endTime).getTime() / 1000);



        var body = JSON.stringify({

            extensions: {

                persistedQuery: {

                    sha256Hash: "743a06934850c780d0251262189e145dabc81be58857f25271e2b26a4fd4f4ed",

                    version: 1

                }

            },

            variables: {

                department: "inbound",

                timeRanges: [{ startTime: startEpoch, endTime: endEpoch }],

                warehouseId: WAREHOUSE

            }

        });



        console.log('[IB Sync] ATLAS fetch: startEpoch=' + startEpoch + ', endEpoch=' + endEpoch);



        return new Promise(function(resolve) {

            GM_xmlhttpRequest({

                method: 'POST',

                url: 'https://atlas.qubit.amazon.dev/graphql',

                headers: {

                    'Content-Type': 'application/json',

                    'Accept': '*/*'

                },

                data: body,

                onload: function(response) {

                    if (response.status === 200) {

                        try {

                            var json = JSON.parse(response.responseText);

                            console.log('[IB Sync] ATLAS raw response keys:', Object.keys(json));

                            if (!json.data) {

                                console.log('[IB Sync] ATLAS response (no data):', JSON.stringify(json).substring(0, 500));

                                resolve({});

                                return;

                            }

                            var reportData = json.data.getReportingByWarehouseId || json.data;

                            var reports = reportData.totalReports || reportData.totalsReports || [];

                            var result = {};

                            reports.forEach(function(r) {

                                result[r.defectType] = {

                                    dpmo: r.metricValue,

                                    threshold: r.threshold,

                                    defectCount: r.defectCount,

                                    opportunities: r.opportunities,

                                    metricType: r.metricType

                                };

                            });

                            console.log('[IB Sync] ATLAS parsed:', Object.keys(result).length, 'defect types');

                            resolve(result);

                        } catch (e) {

                            console.log('[IB Sync] ATLAS parse error:', e.message);

                            resolve({});

                        }

                    } else {

                        console.log('[IB Sync] ATLAS error status:', response.status);

                        resolve({});

                    }

                },

                onerror: function() {

                    console.log('[IB Sync] ATLAS network error');

                    resolve({});

                }

            });

        });

    }



    // ==========================================

    // ROBOSCOUT DATA FETCHER (CT, UPF, NSTA, OOWA)

    // ==========================================

    function toUTCHour(localHour) { return localHour + 5; } // CDT = UTC-5



    function roboNextDay(dateStr) {

        var d = new Date(dateStr + 'T12:00:00');

        d.setDate(d.getDate() + 1);

        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');

    }



    function buildRoboScoutStowUrl(dateStr, startHour, endHour) {

        var utcStart = toUTCHour(startHour);

        var utcEnd = toUTCHour(endHour);

        var startDate = utcStart >= 24 ? roboNextDay(dateStr) : dateStr;

        var endDate = utcEnd >= 24 ? roboNextDay(dateStr) : dateStr;

        var startH = utcStart % 24;

        var endH = utcEnd % 24;

        return 'https://roboscout.amazon.com/view_plot_data/?'

            + 'current_day=false'

            + '&startDateTime=' + startDate + '+' + String(startH).padStart(2, '0') + ':00:00'

            + '&endDateTime=' + endDate + '+' + String(endH).padStart(2, '0') + ':00:00'

            + '&viz=nvd3Table'

            + '&mom_ids=394,321,362,379,426'

            + '&ofm_ids=775'

            + '&osm_ids='

            + '&oxm_ids=444'

            + '&sites=(' + WAREHOUSE + ')'

            + '&instance_id=' + SITE_SETTINGS.ROBOSCOUT_INSTANCE_ID

            + '&object_id=' + SITE_SETTINGS.ROBOSCOUT_STOW_OBJECT_ID

            + '&BrowserTZ=' + SITE_SETTINGS.TIMEZONE

            + '&app_name=RoboScout';

    }



    function buildRoboScoutOOWAUrl(dateStr, startHour, endHour) {

        var utcStart = toUTCHour(startHour);

        var utcEnd = toUTCHour(endHour);

        var startDate = utcStart >= 24 ? roboNextDay(dateStr) : dateStr;

        var endDate = utcEnd >= 24 ? roboNextDay(dateStr) : dateStr;

        var startH = utcStart % 24;

        var endH = utcEnd % 24;

        return 'https://roboscout.amazon.com/view_plot_data/?'

            + 'sites=(' + WAREHOUSE + ')'

            + '&startDateTime=' + startDate + '+' + String(startH).padStart(2, '0') + ':00:00'

            + '&endDateTime=' + endDate + '+' + String(endH).padStart(2, '0') + ':00:00'

            + '&mom_ids=2170,2168'

            + '&osm_ids='

            + '&oxm_ids=2593'

            + '&ofm_ids=1017'

            + '&instance_id=0'

            + '&object_id=' + SITE_SETTINGS.ROBOSCOUT_OOWA_OBJECT_ID

            + '&BrowserTZ=' + SITE_SETTINGS.TIMEZONE;

    }



    // ==========================================

    // FCLM PER-PERIOD METRICS FETCH (ETI + COST)

    // ==========================================

    function fetchFCLMByPeriod() {

        var dateStr = getSelectedDate();

        var dateSlash = dateStr.replace(/-/g, '/');

        var shift = getCurrentShift();

        var periodHours;

        if (shift === 'Day') {

            periodHours = { P1: [6, 10], P2: [10, 14], P3: [14, 18] };

        } else {

            periodHours = { P1: [18, 22], P2: [22, 26], P3: [26, 30] };

        }



        function fetchPeriodFCLM(startH, endH) {

            var nextDateSlash = (function() { var d = new Date(dateStr + 'T12:00:00'); d.setDate(d.getDate() + 1); return d.getFullYear() + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + String(d.getDate()).padStart(2,'0'); })();

            var startDate = dateSlash;

            var sH = startH;

            if (startH >= 24) { sH = startH - 24; startDate = nextDateSlash; }

            var endDate = dateSlash;

            var eH = endH;

            if (endH >= 24) { eH = endH - 24; endDate = nextDateSlash; }

            var url = 'https://fclm-portal.amazon.com/reports/processPathRollup?reportFormat=HTML'

                + '&warehouseId=' + WAREHOUSE + '&spanType=Intraday'

                + '&startDateIntraday=' + startDate + '&startHourIntraday=' + sH + '&startMinuteIntraday=0'

                + '&endDateIntraday=' + endDate + '&endHourIntraday=' + eH + '&endMinuteIntraday=0'

                + '&includeOnlyWarehouse=on&employmentType=AllEmployees';

            return new Promise(function(resolve) {

                GM_xmlhttpRequest({

                    method: 'GET', url: url,

                    onload: function(resp) {

                        if (resp.status === 200) {

                            try {

                                var doc = new DOMParser().parseFromString(resp.responseText, 'text/html');

                                var rows = doc.querySelectorAll('tr');

                                var result = { etiRate: null, decantRate: null, transferInSupportRate: null, stowToPrimeRate: null, ibLeadPARate: null, ibProblemSolveRate: null, ibTotalRate: null, percentToLP: null, etiHrs: null, stpHrs: null };



                                // Build rowMap with DOM elements for class-based extraction

                                var rowMap = {}; // text arrays (legacy)

                                var rowEls = {}; // DOM row elements

                                rows.forEach(function(row) {

                                    var cells = row.querySelectorAll('td, th');

                                    if (cells.length < 7) return;

                                    var cellTexts = Array.from(cells).map(function(c) { return c.textContent.trim(); });

                                    var targets = ['Each Transfer In - Total', 'Transfer In Decant', 'Transfer In Support', 'Stow to Prime Support', 'IB Lead/PA', 'IB Problem Solve', 'Inbound-TOTAL', 'IB Total'];

                                    for (var t = 0; t < targets.length; t++) {

                                        if (cellTexts.indexOf(targets[t]) !== -1 && !rowMap[targets[t]]) {

                                            rowMap[targets[t]] = cellTexts;

                                            rowEls[targets[t]] = row;

                                            break;

                                        }

                                    }

                                });



                                function num(str) { var v = parseFloat((str || '').replace(/[,%]/g, '')); return isNaN(v) ? null : v; }



                                // Class-based extraction from fetched report DOM rows

                                function extractRow(rowEl) {

                                    if (!rowEl) return { rate: null, hrs: null, vol: null, lpRate: null, pctToLP: null };

                                    var cells = rowEl.querySelectorAll('td');

                                    var r = { rate: null, hrs: null, vol: null, lpRate: null, pctToLP: null };

                                    cells.forEach(function(td) {

                                        var cls = td.className || '';

                                        var text = td.textContent.trim().replace(/,/g, '').replace('%', '');

                                        var v = parseFloat(text);

                                        if (isNaN(v)) v = null;

                                        if (cls.indexOf('actualVolume') !== -1) { r.vol = v; }

                                        else if (cls.indexOf('actualTimeSeconds') !== -1 && cls.indexOf('priorYear') === -1) { r.hrs = v; }

                                        else if (cls.indexOf('actualProductivity') !== -1 && cls.indexOf('priorYear') === -1) { r.rate = v; }

                                    });

                                    // Note: On fetched pages, PPR VS LP has NOT injected LP Rate.

                                    // .priorYearVolume contains Prior Year Volume (not LP Rate).

                                    // LP Rate for % to LP calc comes from the visible page instead.

                                    return r;

                                }



                                // Extract all metrics using class-based approach

                                var etiEx = extractRow(rowEls['Each Transfer In - Total']);

                                var decEx = extractRow(rowEls['Transfer In Decant']);

                                var tisEx = extractRow(rowEls['Transfer In Support']);

                                var stpEx = extractRow(rowEls['Stow to Prime Support']);

                                var lpaEx = extractRow(rowEls['IB Lead/PA']);

                                var psEx = extractRow(rowEls['IB Problem Solve']);

                                var ibEx = extractRow(rowEls['Inbound-TOTAL'] || rowEls['IB Total']);



                                result.etiRate = etiEx.rate;

                                result.decantRate = decEx.rate;

                                result.transferInSupportRate = tisEx.rate;

                                result.stowToPrimeRate = stpEx.rate;

                                result.ibLeadPARate = lpaEx.rate;

                                result.ibProblemSolveRate = psEx.rate;

                                result.ibTotalRate = ibEx.rate;

                                result.ibTotalVol = ibEx.vol;

                                result.ibTotalLPRate = ibEx.lpRate;

                                result.percentToLP = ibEx.pctToLP;

                                result.etiHrs = etiEx.hrs;

                                result.stpHrs = stpEx.hrs;



                                console.log('[IB Sync] FCLM period class-based extraction - ETI rate:', etiEx.rate, 'IB Total rate:', ibEx.rate, 'vol:', ibEx.vol, 'hrs:', ibEx.hrs, 'LP:', ibEx.lpRate);



                                resolve(result);

                                console.log('[IB Sync] FCLM period fetch result:', JSON.stringify(result));

                                console.log('[IB Sync] FCLM period rowMap keys:', Object.keys(rowMap).join(', '), '| URL:', url.substring(url.indexOf('startHour')));

                            } catch(e) { resolve(null); }

                        } else {

                            console.log('[IB Sync] FCLM period fetch FAILED status:', resp.status, 'URL:', url.substring(url.indexOf('startHour')));

                            resolve(null);

                        }

                    },

                    onerror: function(err) { console.log('[IB Sync] FCLM period fetch ERROR:', err); resolve(null); }

                });

            });

        }



        return Promise.allSettled([

            fetchPeriodFCLM(periodHours.P1[0], periodHours.P1[1]),

            fetchPeriodFCLM(periodHours.P2[0], periodHours.P2[1]),

            fetchPeriodFCLM(periodHours.P3[0], periodHours.P3[1])

        ]).then(function(results) {

            var p1 = results[0].status === 'fulfilled' ? results[0].value : null;

            var p2 = results[1].status === 'fulfilled' ? results[1].value : null;

            var p3 = results[2].status === 'fulfilled' ? results[2].value : null;

            return {

                etiRate: { P1: p1 && p1.etiRate, P2: p2 && p2.etiRate, P3: p3 && p3.etiRate },

                decantRate: { P1: p1 && p1.decantRate, P2: p2 && p2.decantRate, P3: p3 && p3.decantRate },

                transferInSupportRate: { P1: p1 && p1.transferInSupportRate, P2: p2 && p2.transferInSupportRate, P3: p3 && p3.transferInSupportRate },

                stowToPrimeRate: { P1: p1 && p1.stowToPrimeRate, P2: p2 && p2.stowToPrimeRate, P3: p3 && p3.stowToPrimeRate },

                ibLeadPARate: { P1: p1 && p1.ibLeadPARate, P2: p2 && p2.ibLeadPARate, P3: p3 && p3.ibLeadPARate },

                ibProblemSolveRate: { P1: p1 && p1.ibProblemSolveRate, P2: p2 && p2.ibProblemSolveRate, P3: p3 && p3.ibProblemSolveRate },

                ibTotalRate: { P1: p1 && p1.ibTotalRate, P2: p2 && p2.ibTotalRate, P3: p3 && p3.ibTotalRate },

                percentToLP: { P1: p1 && p1.percentToLP, P2: p2 && p2.percentToLP, P3: p3 && p3.percentToLP },

                ibTotalVol: { P1: p1 && p1.ibTotalVol, P2: p2 && p2.ibTotalVol, P3: p3 && p3.ibTotalVol },

                etiHrs: { P1: p1 && p1.etiHrs, P2: p2 && p2.etiHrs, P3: p3 && p3.etiHrs },

                stpHrs: { P1: p1 && p1.stpHrs, P2: p2 && p2.stpHrs, P3: p3 && p3.stpHrs },

                ibTotalLPRate: (p1 && p1.ibTotalLPRate) || (p2 && p2.ibTotalLPRate) || (p3 && p3.ibTotalLPRate) || null,

            };

        });

    }



    function fetchRoboScoutData() {

        var dateStr = getSelectedDate();

        var shift = getCurrentShift();

        var startHour = shift === 'Day' ? 6 : 18;

        var endHour = shift === 'Day' ? 18 : 30; // 30 = 6am next day



        var stowUrl = buildRoboScoutStowUrl(dateStr, startHour, endHour);

        var oowaUrl = buildRoboScoutOOWAUrl(dateStr, startHour, endHour);



        // Period-level stow URLs for P1/P2/P3

        var periodHours;

        if (shift === 'Day') {

            periodHours = { P1: [6, 10], P2: [10, 14], P3: [14, 18] };

        } else {

            periodHours = { P1: [18, 22], P2: [22, 26], P3: [26, 30] }; // 22=10pm, 26=2am, 30=6am next day

        }

        var stowP1Url = buildRoboScoutStowUrl(dateStr, periodHours.P1[0], periodHours.P1[1]);

        var stowP2Url = buildRoboScoutStowUrl(dateStr, periodHours.P2[0], periodHours.P2[1]);

        var stowP3Url = buildRoboScoutStowUrl(dateStr, periodHours.P3[0], periodHours.P3[1]);

        // Per-period OOWA URLs

        var oowaP1Url = buildRoboScoutOOWAUrl(dateStr, periodHours.P1[0], periodHours.P1[1]);

        var oowaP2Url = buildRoboScoutOOWAUrl(dateStr, periodHours.P2[0], periodHours.P2[1]);

        var oowaP3Url = buildRoboScoutOOWAUrl(dateStr, periodHours.P3[0], periodHours.P3[1]);



        // Per-floor stow URL (uses osm_ids=31 for floor breakdown) — uses CURRENT PERIOD time range

        var currentPeriod = getCurrentPeriod();

        var floorPeriodHrs = (currentPeriod === 'P1') ? periodHours.P1 : (currentPeriod === 'P2') ? periodHours.P2 : (currentPeriod === 'P3') ? periodHours.P3 : [startHour, endHour];

        var utcStart = toUTCHour(floorPeriodHrs[0]);

        var utcEnd = toUTCHour(floorPeriodHrs[1]);

        var floorStartDate = utcStart >= 24 ? roboNextDay(dateStr) : dateStr;

        var floorEndDate = utcEnd >= 24 ? roboNextDay(dateStr) : dateStr;

        var floorStartH = utcStart % 24;

        var floorEndH = utcEnd % 24;

        var floorUrl = 'https://roboscout.amazon.com/view_plot_data/?sites=(' + WAREHOUSE + ')&current_day=false'

            + '&startDateTime=' + floorStartDate + '+' + String(floorStartH).padStart(2, '0') + ':00:00'

            + '&endDateTime=' + floorEndDate + '+' + String(floorEndH).padStart(2, '0') + ':00:00'

            + '&mom_ids=394,321,362,379,426&osm_ids=31&oxm_ids=435&ofm_ids=&viz=nvd3Table'

            + '&instance_id=' + SITE_SETTINGS.ROBOSCOUT_INSTANCE_ID + '&object_id=' + SITE_SETTINGS.ROBOSCOUT_STOW_OBJECT_ID + '&BrowserTZ=' + SITE_SETTINGS.TIMEZONE + '&app_name=RoboScout';



        // Per-floor OOWA URL (different params: oxm_ids=2594, object_id=21628)

        var floorOowaUrl = 'https://roboscout.amazon.com/view_plot_data/?sites=(' + WAREHOUSE + ')'

            + '&startDateTime=' + floorStartDate + '+' + String(floorStartH).padStart(2, '0') + ':00:00'

            + '&endDateTime=' + floorEndDate + '+' + String(floorEndH).padStart(2, '0') + ':00:00'

            + '&mom_ids=2170,2168&osm_ids=&oxm_ids=2594&ofm_ids=1017&instance_id=0&object_id=' + SITE_SETTINGS.ROBOSCOUT_OOWA_OBJECT_ID + '&BrowserTZ=' + SITE_SETTINGS.TIMEZONE + '&app_name=RoboScout&viz=nvd3Table';



        console.log('[IB Sync] RoboScout stow URL:', stowUrl.substring(0, 150) + '...');



        return Promise.allSettled([

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: stowUrl, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: oowaUrl, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: floorUrl, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: floorOowaUrl, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: stowP1Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: stowP2Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: stowP3Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: oowaP1Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: oowaP2Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            }),

            new Promise(function(resolve) {

                GM_xmlhttpRequest({ method: 'GET', url: oowaP3Url, onload: function(r) {

                    if (r.status === 200) { try { resolve(JSON.parse(r.responseText)); } catch(e) { resolve(null); } } else { resolve(null); }

                }, onerror: function() { resolve(null); } });

            })

        ]).then(function(results) {

            var stowJson = results[0].status === 'fulfilled' ? results[0].value : null;

            var oowaJson = results[1].status === 'fulfilled' ? results[1].value : null;

            var floorJson = results[2].status === 'fulfilled' ? results[2].value : null;

            var floorOowaJson = results[3].status === 'fulfilled' ? results[3].value : null;

            var stowP1Json = results[4].status === 'fulfilled' ? results[4].value : null;

            var stowP2Json = results[5].status === 'fulfilled' ? results[5].value : null;

            var stowP3Json = results[6].status === 'fulfilled' ? results[6].value : null;

            var oowaP1Json = results[7].status === 'fulfilled' ? results[7].value : null;

            var oowaP2Json = results[8].status === 'fulfilled' ? results[8].value : null;

            var oowaP3Json = results[9].status === 'fulfilled' ? results[9].value : null;



            var get = function(json, key) { if (!json || !json.data) return null; var item = json.data.find(function(d) { return d.key === key && d.type !== 'string' && d.type !== 'date'; }); if (!item) return null; var v = parseFloat(item.yValue); return isNaN(v) ? null : v; };



            // OOWA per-floor getter: uses xValue as floor identifier (e.g., "paKivaA02")

            var getByFloorX = function(json, key, zoneName) { if (!json || !json.data) return null; var item = json.data.find(function(d) { return d.key === key && d.xValue === zoneName && d.type !== 'string'; }); if (!item) return null; var v = parseFloat(item.yValue); return isNaN(v) ? null : v; };



            // Debug: log all OOWA keys

            if (oowaJson && oowaJson.data) { console.log('[IB Sync] RoboScout OOWA site keys:', JSON.stringify(oowaJson.data.map(function(d) { return {key: d.key, yValue: d.yValue, xValue: d.xValue, type: d.type}; }))); }

            if (floorOowaJson && floorOowaJson.data) { console.log('[IB Sync] RoboScout OOWA floor keys:', JSON.stringify(floorOowaJson.data.slice(0, 10).map(function(d) { return {key: d.key, yValue: d.yValue, xValue: d.xValue, type: d.type}; }))); }



            // Parse per-floor data: keys are like "Units_Per_Face  A02"

            var byFloor = {};

            SITE_SETTINGS.FLOORS.forEach(function(floor) {

                var zoneName = SITE_SETTINGS.ZONE_PREFIX + floor;

                byFloor[floor] = {

                    upf: get(floorJson, 'Units_Per_Face  ' + floor),

                    stowCycleTime: get(floorJson, 'Stow_Takt_Time  ' + floor),

                    nsta: get(floorJson, 'Turnaway_Percentage  ' + floor),

                    stowRate: get(floorJson, 'Stow_Rate  ' + floor),

                    etiTotal: get(floorJson, 'Total_Stowed_Units  ' + floor),

                    oowaCount: getByFloorX(floorOowaJson, 'Andon Count', zoneName),

                    oowaDwell: getByFloorX(floorOowaJson, 'Andon Dwell Time Average', zoneName),

                };

            });

            console.log('[IB Sync] RoboScout by floor:', JSON.stringify(byFloor));



            var result = {

                stowCycleTime: get(stowJson, 'Stow_Takt_Time'),

                nsta: get(stowJson, 'Turnaway_Percentage'),

                upf: get(stowJson, 'Units_Per_Face'),

                secPerTA: get(stowJson, 'Seconds_Per_Turnaway'),

                oowaCount: get(oowaJson, 'Andon Count'),

                oowaDwell: get(oowaJson, 'Andon Dwell Time Average'),

                byFloor: byFloor,

                // Per-period metrics

                byPeriod: {

                    P1: { stowCycleTime: get(stowP1Json, 'Stow_Takt_Time'), upf: get(stowP1Json, 'Units_Per_Face'), nsta: get(stowP1Json, 'Turnaway_Percentage'), oowaCount: get(oowaP1Json, 'Andon Count'), oowaDwell: get(oowaP1Json, 'Andon Dwell Time Average') },

                    P2: { stowCycleTime: get(stowP2Json, 'Stow_Takt_Time'), upf: get(stowP2Json, 'Units_Per_Face'), nsta: get(stowP2Json, 'Turnaway_Percentage'), oowaCount: get(oowaP2Json, 'Andon Count'), oowaDwell: get(oowaP2Json, 'Andon Dwell Time Average') },

                    P3: { stowCycleTime: get(stowP3Json, 'Stow_Takt_Time'), upf: get(stowP3Json, 'Units_Per_Face'), nsta: get(stowP3Json, 'Turnaway_Percentage'), oowaCount: get(oowaP3Json, 'Andon Count'), oowaDwell: get(oowaP3Json, 'Andon Dwell Time Average') },

                }

            };

            console.log('[IB Sync] RoboScout results - CT:', result.stowCycleTime, 'UPF:', result.upf, 'NSTA:', result.nsta, 'OOWA count:', result.oowaCount, 'dwell:', result.oowaDwell,

                'P1 CT:', result.byPeriod.P1.stowCycleTime, 'P2 CT:', result.byPeriod.P2.stowCycleTime, 'P3 CT:', result.byPeriod.P3.stowCycleTime);

            return result;

        });

    }



    // ==========================================

    // VANTAGE DATA FETCHER

    // ==========================================

    function fetchVantageMetrics() {

        var shiftBounds = getShiftBounds();

        var startHHMM = ('0' + shiftBounds.start.getHours()).slice(-2) + ('0' + shiftBounds.start.getMinutes()).slice(-2);

        var zones = SITE.zones || SITE_SETTINGS.FLOORS.map(function(f) { return SITE_SETTINGS.ZONE_PREFIX + f; });



        // Correct Vantage URL format (from Network tab): /fulfillment?dataset=...

        var baseUrl = 'https://vantage.amazon.com/fulfillment';

        var params = '?dataset=stow_metrics%2Fstow_performance_metrics';

        params += '&customer=AMZN&warehouse=' + WAREHOUSE;

        zones.forEach(function(z, i) { params += '&zones%5B' + i + '%5D=' + z; });

        params += '&startTime=' + startHHMM;



        var url = baseUrl + params;

        console.log('[IB Sync] Vantage URL:', url.substring(0, 200) + '...');



        return new Promise(function(resolve) {

            GM_xmlhttpRequest({

                method: 'GET',

                url: url,

                responseType: 'json',

                headers: { 'Accept': 'application/json' },

                onload: function(response) {

                    console.log('[IB Sync] Vantage status:', response.status);

                    if (response.status === 200) {

                        try {

                            var raw = response.response || response.responseText;

                            var data = typeof raw === 'string' ? JSON.parse(raw) : raw;

                            console.log('[IB Sync] Vantage data keys:', data ? Object.keys(data) : 'null');



                            // Parse site-level metrics from building or etiData

                            var result = {

                                stowCycleTime: null,

                                upf: null,

                                oowa: null,

                                nsta: null,

                                stowRate: null,

                                etiTotal: null,

                                byFloor: {},

                                zones: {}

                            };



                            if (data) {

                                // Site-level from 'building' or 'etiData'

                                var bldg = data.building || data.etiData || {};

                                if (bldg.stow_cycle_time !== undefined) result.stowCycleTime = parseFloat(bldg.stow_cycle_time) || null;

                                if (bldg.units_per_face !== undefined) result.upf = parseFloat(bldg.units_per_face) || null;

                                if (bldg.turnaway_percentage !== undefined) result.nsta = parseFloat(bldg.turnaway_percentage) || null;

                                if (bldg.percentage_of_time_in_no_stow_turnaway !== undefined) result.oowa = parseFloat(bldg.percentage_of_time_in_no_stow_turnaway) || null;

                                if (bldg.stow_rate !== undefined) result.stowRate = parseFloat(bldg.stow_rate) || null;

                                if (bldg.total_stowed_units !== undefined) result.etiTotal = parseInt(bldg.total_stowed_units) || null;



                                // Per-floor from 'floors' array

                                var floors = data.floors || [];

                                if (Array.isArray(floors)) {

                                    floors.forEach(function(f) {

                                        var floorId = f.floor || f.floor_list || f.name || '';

                                        // Map floor identifiers to zone names (for zones{}) and floor keys

                                        var zoneName = null;

                                        if (floorId.indexOf('02') !== -1 || floorId === '2' || floorId === 'FLOOR_2') zoneName = 'paKivaA02';

                                        else if (floorId.indexOf('03') !== -1 || floorId === '3' || floorId === 'FLOOR_3') zoneName = 'paKivaA03';

                                        else if (floorId.indexOf('04') !== -1 || floorId === '4' || floorId === 'FLOOR_4') zoneName = 'paKivaA04';

                                        else if (floorId.indexOf('05') !== -1 || floorId === '5' || floorId === 'FLOOR_5') zoneName = 'paKivaA05';



                                        if (zoneName) {

                                            var floorData = {

                                                stowCycleTime: parseFloat(f.stow_cycle_time) || null,

                                                upf: parseFloat(f.units_per_face) || null,

                                                nsta: parseFloat(f.turnaway_percentage) || null,

                                                oowa: parseFloat(f.percentage_of_time_in_no_stow_turnaway) || null,

                                                stowRate: parseFloat(f.stow_rate) || null,

                                                etiTotal: parseInt(f.total_stowed_units) || null,

                                            };

                                            result.zones[zoneName] = floorData;

                                            result.byFloor[zoneName.replace(SITE_SETTINGS.ZONE_PREFIX, '')] = floorData;

                                        }

                                    });

                                }



                                console.log('[IB Sync] Vantage parsed - CT:', result.stowCycleTime, 'UPF:', result.upf, 'NSTA:', result.nsta, 'OOWA:', result.oowa, 'ETI:', result.etiTotal);

                                console.log('[IB Sync] Vantage floors:', Object.keys(result.byFloor));

                            }

                            resolve(result);

                        } catch (e) {

                            console.error('[IB Sync] Vantage parse error:', e);

                            resolve(null);

                        }

                    } else {

                        console.log('[IB Sync] Vantage HTTP ' + response.status);

                        resolve(null);

                    }

                },

                onerror: function(e) {

                    console.log('[IB Sync] Vantage network error:', e);

                    resolve(null);

                }

            });

        });

    }

    function fetchVantageOOWA(startISO, endISO) {

        var zones = CONFIG.vantage.zones.join('%2C');

        var shiftStart = getShiftBounds().start;

        var startHHMM = ('0' + shiftStart.getHours()).slice(-2) + ('0' + shiftStart.getMinutes()).slice(-2);

        var url = CONFIG.vantage.baseUrl + '?dataset=pick_metrics%2Fpick_performance_metrics&customer=' + CONFIG.vantage.customer + '&warehouse=' + CONFIG.vantage.warehouse + '&zone=' + zones + '&startTime=' + startHHMM;



        return new Promise(function(resolve) {

            GM_xmlhttpRequest({

                method: 'GET',

                url: url,

                responseType: 'json',

                headers: { 'Accept': 'application/json' },

                onload: function(response) {

                    if (response.status === 200) {

                        try {

                            var data = typeof response.response === 'string' ? JSON.parse(response.response) : response.response;

                            var metrics = Array.isArray(data) ? data[0] : data;

                            var oowa = metrics.oowa_percentage || metrics.oowa_percent || metrics.oowa || metrics.out_of_work_area_percentage || null;

                            resolve(oowa);

                        } catch (e) {

                            resolve(null);

                        }

                    } else {

                        resolve(null);

                    }

                },

                onerror: function() { resolve(null); }

            });

        });

    }



    // ==========================================

    // UI: FORMATTING HELPERS

    // ==========================================

    function formatNum(val) {

        if (val === null || val === undefined) return '';

        if (typeof val === 'number') {

            return val >= 1000 ? val.toLocaleString() : val.toString();

        }

        return val.toString();

    }



    function getAuditGoal(name) {

        var goals = { 'Stow CT': '20', 'UPF': '20', 'Qty Stow': '5+', 'FISH': '20', 'ASIN Progression': '20' };

        return goals[name] || '';

    }



    // ==========================================

    // UI: SYNC DASHBOARD PANEL

    // ==========================================

    function createSyncPanel(fclmData, apolloData, vantageData, introData, slimData, atlasData, sourceStatus, fclmByPeriod) {

        var period = getCurrentPeriod();

        var shift = getCurrentShift();



        // Null guard for per-period FCLM data

        var emptyPeriod = { P1: null, P2: null, P3: null };

        if (!fclmByPeriod) fclmByPeriod = {};

        if (!fclmByPeriod.etiRate) fclmByPeriod.etiRate = emptyPeriod;

        if (!fclmByPeriod.decantRate) fclmByPeriod.decantRate = emptyPeriod;

        if (!fclmByPeriod.transferInSupportRate) fclmByPeriod.transferInSupportRate = emptyPeriod;

        if (!fclmByPeriod.stowToPrimeRate) fclmByPeriod.stowToPrimeRate = emptyPeriod;

        if (!fclmByPeriod.ibLeadPARate) fclmByPeriod.ibLeadPARate = emptyPeriod;

        if (!fclmByPeriod.ibProblemSolveRate) fclmByPeriod.ibProblemSolveRate = emptyPeriod;

        if (!fclmByPeriod.percentToLP) fclmByPeriod.percentToLP = emptyPeriod;

        if (!fclmByPeriod.ibTotalRate) fclmByPeriod.ibTotalRate = emptyPeriod;

        if (!fclmByPeriod.ibTotalVol) fclmByPeriod.ibTotalVol = emptyPeriod;



        // Load saved snapshots for past periods

        var snapshots = {};

        try {

            var storedSnaps = localStorage.getItem(SNAPSHOT_KEY);

            if (storedSnaps) {

                snapshots = JSON.parse(storedSnaps);

                if (snapshots._shiftDate !== getShiftDate()) snapshots = {};

            }

        } catch(e) { snapshots = {}; }



        // Build per-period SLIM data from snapshots + current

        var periodSlim = {

            receiveAutoDecant: { P1: null, P2: null, P3: null, EOS: null },

            receiveManualDecant: { P1: null, P2: null, P3: null, EOS: null },

        };

        ['P1', 'P2', 'P3'].forEach(function(p) {

            if (snapshots[p] && snapshots[p].slim) {

                periodSlim.receiveAutoDecant[p] = snapshots[p].slim.receiveAutoDecant;

                periodSlim.receiveManualDecant[p] = snapshots[p].slim.receiveManualDecant;

            }

        });

        // Current period always uses live SLIM data

        if (slimData) {

            periodSlim.receiveAutoDecant[period] = slimData.receiveAutoDecant;

            periodSlim.receiveManualDecant[period] = slimData.receiveManualDecant;

            periodSlim.receiveAutoDecant.EOS = slimData.receiveAutoDecant;

            periodSlim.receiveManualDecant.EOS = slimData.receiveManualDecant;

        }



        // Build period rates from snapshots + current FCLM data

        // P1/P2 = snapshots from those periods, EOS = current (or shift total when available)

        var periodRates = {

            ibTotalRate: { P1: null, P2: null, P3: null, EOS: null },

            decantRate: { P1: null, P2: null, P3: null, EOS: null },

            stowToprimeRate: { P1: null, P2: null, P3: null, EOS: null },

            transferInSupportRate: { P1: null, P2: null, P3: null, EOS: null },

            ibLeadPARate: { P1: null, P2: null, P3: null, EOS: null },

            ibProblemSolveRate: { P1: null, P2: null, P3: null, EOS: null },

            stowToPrimeSupportRate: { P1: null, P2: null, P3: null, EOS: null },

            transferInRate: { P1: null, P2: null, P3: null, EOS: null },

        };



        // Fill from snapshots

        ['P1', 'P2', 'P3', 'EOS'].forEach(function(p) {

            if (snapshots[p] && snapshots[p].fclm) {

                var sf = snapshots[p].fclm;

                periodRates.ibTotalRate[p] = sf.ibTotalRate || sf.ibTotalTPH;

                periodRates.decantRate[p] = sf.decantRate;

                periodRates.stowToprimeRate[p] = sf.stowToprimeRate;

                periodRates.transferInSupportRate[p] = sf.transferInSupportRate;

                periodRates.ibLeadPARate[p] = sf.ibLeadPARate;

                periodRates.ibProblemSolveRate[p] = sf.ibProblemSolveRate;

                periodRates.stowToPrimeSupportRate[p] = sf.stowToPrimeSupportRate;

                periodRates.transferInRate[p] = sf.transferInRate;

            }

        });



        // Current period always uses live FCLM data (overwrite snapshot with fresh)

        periodRates.ibTotalRate[period] = fclmData.ibTotalRate || fclmData.ibTotalTPH;

        periodRates.decantRate[period] = fclmData.decantRate;

        periodRates.stowToprimeRate[period] = fclmData.stowToprimeRate;

        periodRates.transferInSupportRate[period] = fclmData.transferInSupportRate;

        periodRates.ibLeadPARate[period] = fclmData.ibLeadPARate;

        periodRates.ibProblemSolveRate[period] = fclmData.ibProblemSolveRate;

        periodRates.stowToPrimeSupportRate[period] = fclmData.stowToPrimeSupportRate;

        periodRates.transferInRate[period] = fclmData.transferInRate;



        // EOS = current live rates (always show current totals for EOS)

        periodRates.ibTotalRate.EOS = fclmData.ibTotalRate || fclmData.ibTotalTPH;

        periodRates.decantRate.EOS = fclmData.decantRate;

        periodRates.stowToprimeRate.EOS = fclmData.stowToprimeRate;

        periodRates.transferInSupportRate.EOS = fclmData.transferInSupportRate;

        periodRates.ibLeadPARate.EOS = fclmData.ibLeadPARate;

        periodRates.ibProblemSolveRate.EOS = fclmData.ibProblemSolveRate;

        periodRates.stowToPrimeSupportRate.EOS = fclmData.stowToPrimeSupportRate;

        periodRates.transferInRate.EOS = fclmData.transferInRate;



        console.log('[IB Sync] Period rates (FCLM Actual):', periodRates);



        // Build period Vantage metrics from snapshots

        var periodVantage = {

            stowCycleTime: { P1: null, P2: null, P3: null, EOS: null },

            upf: { P1: null, P2: null, P3: null, EOS: null },

            oowa: { P1: null, P2: null, P3: null, EOS: null },

            nsta: { P1: null, P2: null, P3: null, EOS: null },

            stowRate: { P1: null, P2: null, P3: null, EOS: null },

        };



        ['P1', 'P2', 'P3', 'EOS'].forEach(function(p) {

            if (snapshots[p] && snapshots[p].vantage) {

                var sv = snapshots[p].vantage;

                periodVantage.stowCycleTime[p] = sv.stowCycleTime;

                periodVantage.upf[p] = sv.upf;

                periodVantage.oowa[p] = sv.oowa;

                periodVantage.nsta[p] = sv.nsta;

                periodVantage.stowRate[p] = sv.stowRate;

            }

        });



        // Override with RoboScout per-period data if available (more reliable than snapshots)

        if (vantageData.byPeriod) {

            ['P1', 'P2', 'P3'].forEach(function(p) {

                var bp = vantageData.byPeriod[p];

                if (bp) {

                    if (bp.stowCycleTime) periodVantage.stowCycleTime[p] = bp.stowCycleTime;

                    if (bp.upf) periodVantage.upf[p] = bp.upf;

                    if (bp.nsta) periodVantage.nsta[p] = bp.nsta;

                    if (bp.oowa !== null && bp.oowa !== undefined) periodVantage.oowa[p] = bp.oowa;

                }

            });

        }

        // EOS = site-level (full shift)

        if (vantageData.stowCycleTime) periodVantage.stowCycleTime.EOS = vantageData.stowCycleTime;

        if (vantageData.upf) periodVantage.upf.EOS = vantageData.upf;

        if (vantageData.nsta) periodVantage.nsta.EOS = vantageData.nsta;

        if (vantageData.oowa) periodVantage.oowa.EOS = vantageData.oowa;



        // Current period uses live data

        // Only set current period from site-level if per-period data wasn't already set by RoboScout

        if (!periodVantage.stowCycleTime[period]) periodVantage.stowCycleTime[period] = vantageData.stowCycleTime;

        if (!periodVantage.upf[period]) periodVantage.upf[period] = vantageData.upf;

        if (!periodVantage.oowa[period]) periodVantage.oowa[period] = vantageData.oowa;

        if (!periodVantage.nsta[period]) periodVantage.nsta[period] = vantageData.nsta;

        if (!periodVantage.stowRate[period]) periodVantage.stowRate[period] = vantageData.stowRate;



        // EOS = current

        periodVantage.stowCycleTime.EOS = vantageData.stowCycleTime;

        periodVantage.upf.EOS = vantageData.upf;

        periodVantage.oowa.EOS = vantageData.oowa;

        periodVantage.nsta.EOS = vantageData.nsta;

        periodVantage.stowRate.EOS = vantageData.stowRate;



        var existing = document.getElementById('ib-sync-dashboard');

        if (existing) existing.remove();



        var container = document.createElement('div');

        container.id = 'ib-sync-dashboard';

        container.style.cssText = 'width:100%;max-width:1440px;margin:20px auto;font-family:Arial,sans-serif;border:2px solid #232f3e;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.15);position:relative;z-index:2147483647;background:white;isolation:isolate;';



        // Header

        var s1Goal = '';

        if (introData.trailerPlanner.s1Goal) {

            s1Goal = ' | S1 Goal: ' + formatNum(introData.trailerPlanner.s1Goal);

            var alps = introData.trailerPlanner.alpsS1;

            if (alps && alps.dayCapacity && alps.nightCapacity) {

                s1Goal += ' (DS: ' + formatNum(alps.dayCapacity) + ' | NS: ' + formatNum(alps.nightCapacity) + ')';

            }

        }



        // Build source status indicators

        var ss = sourceStatus || {};

        function srcDot(name, ok) { return '<span title="' + name + (ok ? ' \u2714' : ' \u2718 (open ' + name + ' in browser)') + '" style="font-size:9px;margin:0 2px;color:' + (ok ? '#4caf50' : '#f44336') + ';">' + (ok ? '\u25CF' : '\u25CF') + '</span>'; }

        var statusHtml = '<span style="font-size:9px;color:#aaa;margin-left:8px;" title="Data sources: green=connected, red=failed">' +

            srcDot('FCLM', ss.fclm) + srcDot('INTRO', ss.intro) + srcDot('Apollo', ss.apollo) + srcDot('ATLAS', ss.atlas) + srcDot('SLIM', ss.slim) + srcDot('RoboScout', ss.roboscout) +

            '</span>';



        var header = document.createElement('div');

        header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:#232f3e;color:white;cursor:pointer;';

        header.innerHTML = '<div style="display:flex;align-items:center;gap:12px;">' +

            '<span style="font-size:16px;font-weight:bold;">\u{1F4CB} ' + WAREHOUSE + ' Inbound \u2014 Leadership Sync Tracker \u00b7 ' + shift.toUpperCase() + ' SHIFT</span>' +

            '<span style="font-size:11px;background:#ff9900;color:#232f3e;padding:2px 8px;border-radius:4px;font-weight:bold;">' + period + s1Goal + '</span>' +

            '<span style="font-size:10px;color:#aaa;">v3.1</span>' + statusHtml +

            '</div>' +

            '<div style="display:flex;align-items:center;gap:10px;">' +

            '<button id="sync-refresh-btn" style="padding:4px 12px;font-size:11px;background:#527FFF;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">\u{1F504} Refresh</button>' +

            '<button id="sync-copy-btn" style="padding:4px 12px;font-size:11px;background:#ff9900;color:#232f3e;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">\u{1F4CB} Copy</button>' +

            '<button id="sync-pdf-btn" style="padding:4px 12px;font-size:11px;background:#28a745;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">\u{1F4C4} PDF</button>' +

            '<button id="sync-excel-btn" style="padding:4px 12px;font-size:11px;background:#217346;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">\u{1F4CA} Excel</button>' +

            '<button id="sync-clear-btn" style="padding:4px 12px;font-size:11px;background:#dc3545;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">\u{1F5D1} Clear</button>' +

            '<span id="sync-toggle" style="font-size:14px;">\u25B6</span>' +

            '</div>';



        // Quick-links bar for failed sources (only shows if any source is red)

        var quickLinks = {

            INTRO: 'https://na.prod.fmc.aft.amazon.dev/' + WAREHOUSE + '/intro-shift-planner/shift-planner',

            Apollo: 'https://apollo-audit.corp.amazon.com/reporting/audit_execution_metrics',

            ATLAS: 'https://atlas.qubit.amazon.dev/reporting?aggregateType=WAREHOUSE_ID&queryType=SHIFT&targetProcess=inbound&warehouseId=' + WAREHOUSE,

            SLIM: 'https://slim.corp.amazon.com/dashboard',

            RoboScout: 'https://roboscout.amazon.com/home/?sites=(' + WAREHOUSE + ')'

        };

        var failedSources = Object.keys(quickLinks).filter(function(k) { return !(ss[k.toLowerCase()]); });

        var quickLinkBar = '';

        if (failedSources.length > 0) {

            var links = failedSources.map(function(name) {

                return '<a href="' + quickLinks[name] + '" target="_blank" style="color:#ff6b6b;font-size:9px;text-decoration:underline;margin:0 4px;">' + name + '</a>';

            }).join(' | ');

            quickLinkBar = '<div style="background:#1a1a2e;padding:3px 16px;font-size:9px;color:#aaa;">Open to connect: ' + links + ' \u2022 then click Refresh</div>';

        }



        // Content area - 3 column grid

        var content = document.createElement('div');

        content.id = 'sync-dashboard-content';

        content.style.cssText = 'padding:12px;background:#fafafa;display:none;grid-template-columns:1fr 1fr 1fr;gap:10px;align-items:start;';



        var sp = introData.shiftPlanner;

        var tp = introData.trailerPlanner;

        var ibf = introData.ibFlow;



        // ===================== COLUMN 1 (LEFT) =====================

        var col1 = document.createElement('div');

        col1.style.cssText = 'display:flex;flex-direction:column;gap:10px;';



        // SAFETY

        col1.appendChild(createSection('SAFETY', [

            { metric: 'Safety Incidents', goal: '0' },

            { metric: 'Austin Action Completion', goal: '100%' },

            { metric: 'ICare Completion', goal: '100%' },

            { metric: 'Investigation uploaded to #ord5-incident-notification?', goal: 'Y if incident' },

            { metric: 'DragonFly Completion', goal: '1' },

            { metric: 'Working Well Huddles/CAST', goal: '1' },

            { metric: 'RBI / ARC completed', goal: '100%' },

        ]));



        // PEOPLE

        col1.appendChild(createSection('PEOPLE', [

            { metric: 'Attendance % (show rate)', goal: '\u2265 plan' },

            { metric: 'Campaign Engagements', goal: '100% on site' },

            { metric: 'Pending ADAPT Feedbacks', goal: '100% on Site' },

            { metric: 'Time Not Logged (TNL)', goal: '0' },

            { metric: 'Unscheduled Hours', goal: '0' },

        ]));



        // TRAINING

        col1.appendChild(createSection('TRAINING', [

            { metric: 'Classes today (NH / CT / HITS)', goal: '' },

            { metric: 'Function (Stow / Decant / Indirect)', goal: '' },

            { metric: 'Total HC trained', goal: '' },

            { metric: 'Ambassadors used / hrs consumed', goal: '' },

        ]));



        // EQUIPMENT / SEV

        col1.appendChild(createSection('EQUIPMENT / SEV', [

            { metric: 'SEV events (Y/N)', goal: '0' },

            { metric: 'SEV \u2014 hrs impact / vol impact', goal: '' },

            { metric: 'Non-SEV equipment impacts', goal: '' },

            { metric: 'ARSTOW stations down', goal: '0' },

            { metric: 'Universal stations down', goal: '0' },

        ]));



        // CI PROJECTS

        col1.appendChild(createCISection());



        content.appendChild(col1);



        // ===================== COLUMN 2 (MIDDLE) =====================

        var col2 = document.createElement('div');

        col2.style.cssText = 'display:flex;flex-direction:column;gap:10px;';



        // NAIL THE PLAN

        var ntpRows = [

            (function() {

                // Plan = Optimus Target, EOS = Shift Total Planned (P1+P2+P3 sum)

                // If the sum of period volumes differs from Optimus, the plan was updated mid-shift — add asterisk

                var planVal = tp.optimusTarget;

                var periodSum = (sp.volumePlannedByPeriod.P1 || sp.volumeByPeriod.P1 || 0) + (sp.volumePlannedByPeriod.P2 || sp.volumeByPeriod.P2 || 0) + (sp.volumePlannedByPeriod.P3 || sp.volumeByPeriod.P3 || 0);

                var planChanged = planVal && periodSum && Math.abs(periodSum - planVal) > 100;

                var metricLabel = planChanged ? 'IB Volume Plan *' : 'IB Volume Plan';

                return { metric: metricLabel, goal: formatNum(planVal), p1: formatNum(sp.volumePlannedByPeriod.P1 || sp.volumeByPeriod.P1), p2: formatNum(sp.volumePlannedByPeriod.P2 || sp.volumeByPeriod.P2), p3: formatNum(sp.volumePlannedByPeriod.P3 || sp.volumeByPeriod.P3), eos: formatNum(periodSum || planVal) };

            })(),

            (function() {

                var ibPlan = { P1: sp.volumePlannedByPeriod.P1 || sp.volumeByPeriod.P1, P2: sp.volumePlannedByPeriod.P2 || sp.volumeByPeriod.P2, P3: sp.volumePlannedByPeriod.P3 || sp.volumeByPeriod.P3, EOS: tp.optimusTarget };

                function chk(actual, plan) {

                    if (actual === null || actual === undefined || plan === null || plan === undefined) return '';

                    return actual >= plan ? ' \u2705' : ' \u274C';

                }

                return { metric: 'Trailer Plan Volume',

                    goal: formatNum(tp.trailerPlanVolume) + chk(tp.trailerPlanVolume, tp.optimusTarget),

                    p1: formatNum(tp.trailerPlanByPeriod.P1) + chk(tp.trailerPlanByPeriod.P1, ibPlan.P1),

                    p2: formatNum(tp.trailerPlanByPeriod.P2) + chk(tp.trailerPlanByPeriod.P2, ibPlan.P2),

                    p3: formatNum(tp.trailerPlanByPeriod.P3) + chk(tp.trailerPlanByPeriod.P3, ibPlan.P3),

                    eos: formatNum(tp.trailerPlanVolume) + chk(tp.trailerPlanVolume, tp.optimusTarget) };

            })(),

            (function() {

                var ibPlan = { P1: sp.volumePlannedByPeriod.P1 || sp.volumeByPeriod.P1, P2: sp.volumePlannedByPeriod.P2 || sp.volumeByPeriod.P2, P3: sp.volumePlannedByPeriod.P3 || sp.volumeByPeriod.P3, EOS: tp.optimusTarget };

                var ibVol = fclmByPeriod.ibTotalVol || {};

                function chk(actual, plan) {

                    if (actual === null || actual === undefined || plan === null || plan === undefined) return '';

                    return actual >= plan ? ' \u2705' : ' \u274C';

                }

                return { metric: 'IB Volume Actuals',

                    goal: formatNum(tp.trailerPlanVolume),

                    p1: ibVol.P1 ? formatNum(ibVol.P1) + chk(ibVol.P1, ibPlan.P1) : '',

                    p2: ibVol.P2 ? formatNum(ibVol.P2) + chk(ibVol.P2, ibPlan.P2) : '',

                    p3: ibVol.P3 ? formatNum(ibVol.P3) + chk(ibVol.P3, ibPlan.P3) : '',

                    eos: formatNum(fclmData.ibTotalVol) + chk(fclmData.ibTotalVol, tp.optimusTarget) };

            })(),

            { metric: 'WIP', goal: '> 20K', eos: formatNum(ibf.wip) },

            { metric: 'Fluid Unload %', goal: tp.fluidPercent !== null ? tp.fluidPercent + '%' : '', p1: tp.fluidByPeriod.P1 !== null ? tp.fluidByPeriod.P1 + '%' : '', p2: tp.fluidByPeriod.P2 !== null ? tp.fluidByPeriod.P2 + '%' : '', p3: tp.fluidByPeriod.P3 !== null ? tp.fluidByPeriod.P3 + '%' : '', eos: tp.fluidPercent !== null ? tp.fluidPercent + '%' : '' },

            { metric: 'Case %', goal: tp.casePercent !== null ? tp.casePercent + '%' : '', p1: tp.caseByPeriod.P1 !== null ? tp.caseByPeriod.P1 + '%' : '', p2: tp.caseByPeriod.P2 !== null ? tp.caseByPeriod.P2 + '%' : '', p3: tp.caseByPeriod.P3 !== null ? tp.caseByPeriod.P3 + '%' : '', eos: tp.casePercent !== null ? tp.casePercent + '%' : '' },

            { metric: 'Smalls Mix %', goal: tp.smallsPercent !== null ? tp.smallsPercent + '%' : '', p1: tp.smallsByPeriod.P1 !== null ? tp.smallsByPeriod.P1 + '%' : '', p2: tp.smallsByPeriod.P2 !== null ? tp.smallsByPeriod.P2 + '%' : '', p3: tp.smallsByPeriod.P3 !== null ? tp.smallsByPeriod.P3 + '%' : '', eos: tp.smallsPercent !== null ? tp.smallsPercent + '%' : '' },

            { metric: 'Stow HC', goal: formatNum(sp.stowHC.P1), p1: formatNum(sp.stowHC.P1), p2: formatNum(sp.stowHC.P2), p3: formatNum(sp.stowHC.P3), eos: '' },

            { metric: '  A02', p1: formatNum(sp.stowHCByFloor.A02.P1), p2: formatNum(sp.stowHCByFloor.A02.P2), p3: formatNum(sp.stowHCByFloor.A02.P3), eos: '' },

            { metric: '  A03', p1: formatNum(sp.stowHCByFloor.A03.P1), p2: formatNum(sp.stowHCByFloor.A03.P2), p3: formatNum(sp.stowHCByFloor.A03.P3), eos: '' },

            { metric: '  A04', p1: formatNum(sp.stowHCByFloor.A04.P1), p2: formatNum(sp.stowHCByFloor.A04.P2), p3: formatNum(sp.stowHCByFloor.A04.P3), eos: '' },

            { metric: '  A05', p1: formatNum(sp.stowHCByFloor.A05.P1), p2: formatNum(sp.stowHCByFloor.A05.P2), p3: formatNum(sp.stowHCByFloor.A05.P3), eos: '' },

            { metric: 'Decant HC', goal: formatNum((sp.decantHC.P1 || 0) + (sp.decantHC.P2 || 0) + (sp.decantHC.P3 || 0)), p1: formatNum(sp.decantHC.P1), p2: formatNum(sp.decantHC.P2), p3: formatNum(sp.decantHC.P3), eos: formatNum((sp.decantHC.P1 || 0) + (sp.decantHC.P2 || 0) + (sp.decantHC.P3 || 0)) },

            { metric: 'Problem Solve HC', goal: formatNum((sp.problemSolveHC.P1 || 0) + (sp.problemSolveHC.P2 || 0) + (sp.problemSolveHC.P3 || 0)), p1: formatNum(sp.problemSolveHC.P1), p2: formatNum(sp.problemSolveHC.P2), p3: formatNum(sp.problemSolveHC.P3), eos: formatNum((sp.problemSolveHC.P1 || 0) + (sp.problemSolveHC.P2 || 0) + (sp.problemSolveHC.P3 || 0)) },

            { metric: 'Indirect Spend %', goal: '', p1: sp.indirectSpendPercent.P1 ? sp.indirectSpendPercent.P1 + '%' : '', p2: sp.indirectSpendPercent.P2 ? sp.indirectSpendPercent.P2 + '%' : '', p3: sp.indirectSpendPercent.P3 ? sp.indirectSpendPercent.P3 + '%' : '', eos: sp.indirectSpendPercent.EOS ? sp.indirectSpendPercent.EOS + '%' : '' },

        ];

        placeAutoValues(ntpRows, period);

        col2.appendChild(createSection('NAIL THE PLAN  (Volume, Mix & Staffing)', ntpRows, 'Plan'));



        // COST

        // For COST rates: EOS >= Plan = good (meeting or exceeding labor plan)

        function costChk(actual, plan) {

            if (actual === null || actual === undefined || plan === null || plan === undefined || !actual || !plan) return '';

            return parseFloat(actual) >= parseFloat(plan) ? ' \u2705' : ' \u274C';

        }

        var costRows = [

            (function() {

                function lpFmt(v) { if (!v) return ''; var n = parseFloat(v); return n.toFixed(2) + '%' + (n >= 100 ? ' \u2705' : (n >= 98 ? ' \u26A0\uFE0F' : ' \u274C')); }

                // % to LP per period: calculate from period IB Total rate / LP Rate

                var ibr = fclmByPeriod.ibTotalRate || {};

                // Use live page LP Rate if available; otherwise fall back to per-period fetch LP Rate

                var lpRate = fclmData.ibTotalPlanRate || fclmByPeriod.ibTotalLPRate;

                function calcLP(periodRate) {

                    if (!periodRate || !lpRate) return '';

                    return lpFmt((periodRate / lpRate) * 100);

                }

                var eosLP = fclmData.percentToLP || (fclmData.ibTotalRate && lpRate ? (fclmData.ibTotalRate / lpRate) * 100 : null);

                return { metric: '% to Labor Plan Shift Total', goal: '\u2265 100%', p1: calcLP(ibr.P1), p2: calcLP(ibr.P2), p3: calcLP(ibr.P3), eos: lpFmt(eosLP) };

            })(),

            (function() {

                var ibr = fclmByPeriod.ibTotalRate || {};

                var ibLpRate = fclmData.ibTotalPlanRate || fclmByPeriod.ibTotalLPRate;

                return { metric: 'IB Total TPH', goal: formatNum(ibLpRate),

                  p1: ibr.P1 ? formatNum(ibr.P1) + costChk(ibr.P1, ibLpRate) : '', p2: ibr.P2 ? formatNum(ibr.P2) + costChk(ibr.P2, ibLpRate) : '', p3: ibr.P3 ? formatNum(ibr.P3) + costChk(ibr.P3, ibLpRate) : '', eos: formatNum(fclmData.ibTotalRate) + costChk(fclmData.ibTotalRate, ibLpRate) };

            })(),

            (function() {

                var dr = fclmByPeriod.decantRate || {};

                return { metric: 'Decant Rate', goal: formatNum(fclmData.decantPlanRate), p1: dr.P1 ? formatNum(dr.P1) + costChk(dr.P1, fclmData.decantPlanRate) : '', p2: dr.P2 ? formatNum(dr.P2) + costChk(dr.P2, fclmData.decantPlanRate) : '', p3: dr.P3 ? formatNum(dr.P3) + costChk(dr.P3, fclmData.decantPlanRate) : '', eos: formatNum(fclmData.decantRate) + costChk(fclmData.decantRate, fclmData.decantPlanRate) };

            })(),

            (function() {

                var tis = fclmByPeriod.transferInSupportRate || {};

                return { metric: 'Transfer In Support', goal: formatNum(fclmData.transferInSupportPlanRate), p1: tis.P1 ? formatNum(tis.P1) + costChk(tis.P1, fclmData.transferInSupportPlanRate) : '', p2: tis.P2 ? formatNum(tis.P2) + costChk(tis.P2, fclmData.transferInSupportPlanRate) : '', p3: tis.P3 ? formatNum(tis.P3) + costChk(tis.P3, fclmData.transferInSupportPlanRate) : '', eos: formatNum(fclmData.transferInSupportRate) + costChk(fclmData.transferInSupportRate, fclmData.transferInSupportPlanRate) };

            })(),

            (function() {

                var stp = fclmByPeriod.stowToPrimeRate || {};

                var stpPlan = fclmData.stowToPrimeSupportPlanRate || fclmData.stowToPrimePlanRate;

                return { metric: 'Stow To Prime Support', goal: formatNum(stpPlan), p1: stp.P1 ? formatNum(stp.P1) + costChk(stp.P1, stpPlan) : '', p2: stp.P2 ? formatNum(stp.P2) + costChk(stp.P2, stpPlan) : '', p3: stp.P3 ? formatNum(stp.P3) + costChk(stp.P3, stpPlan) : '', eos: formatNum(fclmData.stowToPrimeSupportRate || fclmData.stowToprimeRate) + costChk(fclmData.stowToPrimeSupportRate || fclmData.stowToprimeRate, stpPlan) };

            })(),

            (function() {

                var lpa = fclmByPeriod.ibLeadPARate || {};

                return { metric: 'IB Lead/PA', goal: formatNum(fclmData.ibLeadPAPlanRate), p1: lpa.P1 ? formatNum(lpa.P1) + costChk(lpa.P1, fclmData.ibLeadPAPlanRate) : '', p2: lpa.P2 ? formatNum(lpa.P2) + costChk(lpa.P2, fclmData.ibLeadPAPlanRate) : '', p3: lpa.P3 ? formatNum(lpa.P3) + costChk(lpa.P3, fclmData.ibLeadPAPlanRate) : '', eos: formatNum(fclmData.ibLeadPARate) + costChk(fclmData.ibLeadPARate, fclmData.ibLeadPAPlanRate) };

            })(),

            (function() {

                var ps = fclmByPeriod.ibProblemSolveRate || {};

                return { metric: 'IB Problem Solve', goal: formatNum(fclmData.ibProblemSolvePlanRate), p1: ps.P1 ? formatNum(ps.P1) + costChk(ps.P1, fclmData.ibProblemSolvePlanRate) : '', p2: ps.P2 ? formatNum(ps.P2) + costChk(ps.P2, fclmData.ibProblemSolvePlanRate) : '', p3: ps.P3 ? formatNum(ps.P3) + costChk(ps.P3, fclmData.ibProblemSolvePlanRate) : '', eos: formatNum(fclmData.ibProblemSolveRate) + costChk(fclmData.ibProblemSolveRate, fclmData.ibProblemSolvePlanRate) };

            })(),

        ];

        placeAutoValues(costRows, period);

        col2.appendChild(createSection('COST', costRows, 'Plan'));



        // INPUT METRICS (site-level)

        function inputChk(val, goal, lowerIsBetter, yellowMax) {

            if (!val) return '';

            var num = parseFloat(val);

            if (isNaN(num)) return '';

            if (lowerIsBetter) {

                if (num <= goal) return ' \u2705';

                if (yellowMax !== undefined && num <= yellowMax) return ' \u26A0\uFE0F';

                return ' \u274C';

            }

            if (num >= goal) return ' \u2705';

            if (yellowMax !== undefined && num >= yellowMax) return ' \u26A0\uFE0F';

            return ' \u274C';

        }

        function addInputChks(pv, goal, lowerIsBetter, suffix, yellowMax) {

            suffix = suffix || '';

            return {

                p1: pv.P1 ? pv.P1 + suffix + inputChk(pv.P1, goal, lowerIsBetter, yellowMax) : '',

                p2: pv.P2 ? pv.P2 + suffix + inputChk(pv.P2, goal, lowerIsBetter, yellowMax) : '',

                p3: pv.P3 ? pv.P3 + suffix + inputChk(pv.P3, goal, lowerIsBetter, yellowMax) : '',

                eos: pv.EOS ? pv.EOS + suffix + inputChk(pv.EOS, goal, lowerIsBetter, yellowMax) : ''

            };

        }

        // Yellow bands: CT 13.1-15, UPF 9-11.99, OOWA 2.01-2.5, NSTA 15.01-15.5

        var ctChk = addInputChks(periodVantage.stowCycleTime, CONFIG.goals.stowCycleTime, true, '', 15);

        var upfChk = addInputChks(periodVantage.upf, CONFIG.goals.upf, false, '', 9);

        var oowaChk = addInputChks(periodVantage.oowa, CONFIG.goals.oowa, true, '%', 2.5);

        var nstaChk = addInputChks(periodVantage.nsta, CONFIG.goals.nsta, true, '%', 15.5);



        var inputRows = [

            (function() {

                var ep = (fclmByPeriod && fclmByPeriod.etiRate) || { P1: null, P2: null, P3: null };

                function etiFmt(val) {

                    if (!val) return '';

                    var v = parseFloat(val).toFixed(1);

                    var n = parseFloat(v);

                    return v + (n >= 210 ? ' \u2705' : (n >= 200 ? ' \u26A0\uFE0F' : ' \u274C'));

                }

                var eosRate = fclmData.transferInRate;

                return { metric: 'ETI Total', goal: '>210',

                    p1: etiFmt(ep.P1), p2: etiFmt(ep.P2), p3: etiFmt(ep.P3), eos: etiFmt(eosRate) };

            })(),

            { metric: 'Stow Cycle Time (s)', goal: '\u2264 13s', p1: ctChk.p1, p2: ctChk.p2, p3: ctChk.p3, eos: ctChk.eos },

            { metric: 'UPF (Units Per Pod Face)', goal: '> 12', p1: upfChk.p1, p2: upfChk.p2, p3: upfChk.p3, eos: upfChk.eos },

            { metric: 'OOWA %', goal: '< 2%', p1: oowaChk.p1, p2: oowaChk.p2, p3: oowaChk.p3, eos: oowaChk.eos },

            { metric: 'NSTA %', goal: '< 15%', p1: nstaChk.p1, p2: nstaChk.p2, p3: nstaChk.p3, eos: nstaChk.eos },

        ];

        placeAutoValues(inputRows, period);

        col2.appendChild(createSection('INPUT METRICS', inputRows));



        // QUALITY

        var atlasOrNull = atlasData || {};

        function getAtlasDPMO(defectType) {

            var entry = atlasOrNull[defectType];

            if (!entry) return { val: '', threshold: '' };

            var val = entry.dpmo !== null && entry.dpmo !== undefined ? formatNum(entry.dpmo) : '';

            var threshold = entry.threshold ? formatNum(entry.threshold) : '';

            return { val: val, threshold: threshold };

        }

        var meDpmo = getAtlasDPMO('Nike Each Multiple Events');

        var qtyMeDpmo = getAtlasDPMO('Nike Quantity Multiple Events');

        var sipsDpmo = getAtlasDPMO('Sips Over And Short');

        var pc99Dpmo = getAtlasDPMO('PC99 to DropZone');

        var bfvDpmo = getAtlasDPMO('Bin Filter Violation');



        var qualityRows = [

            { metric: 'ME DPMO', goal: meDpmo.threshold, eos: meDpmo.val, warn: atlasOrNull['Nike Each Multiple Events'] && atlasOrNull['Nike Each Multiple Events'].dpmo > atlasOrNull['Nike Each Multiple Events'].threshold },

            { metric: 'Quantity Stow ME DPMO', goal: qtyMeDpmo.threshold, eos: qtyMeDpmo.val, warn: atlasOrNull['Nike Quantity Multiple Events'] && atlasOrNull['Nike Quantity Multiple Events'].dpmo > atlasOrNull['Nike Quantity Multiple Events'].threshold },

            { metric: 'SIPS DPMO', goal: sipsDpmo.threshold, eos: sipsDpmo.val, warn: atlasOrNull['Sips Over And Short'] && atlasOrNull['Sips Over And Short'].dpmo > atlasOrNull['Sips Over And Short'].threshold },

            { metric: 'PC99 DPMO', goal: pc99Dpmo.threshold, eos: pc99Dpmo.val, warn: atlasOrNull['PC99 to DropZone'] && atlasOrNull['PC99 to DropZone'].dpmo > atlasOrNull['PC99 to DropZone'].threshold },

            { metric: 'BFV DPMO', goal: bfvDpmo.threshold, eos: bfvDpmo.val, warn: atlasOrNull['Bin Filter Violation'] && atlasOrNull['Bin Filter Violation'].dpmo > atlasOrNull['Bin Filter Violation'].threshold },

            { metric: 'Problem Solve Piles', goal: '' },

            { metric: 'IOLs', goal: '' },

        ];

        placeAutoValues(qualityRows, period);

        col2.appendChild(createSection('QUALITY', qualityRows, 'Threshold'));



        content.appendChild(col2);



        // ===================== COLUMN 3 (RIGHT) =====================

        var col3 = document.createElement('div');

        col3.style.cssText = 'display:flex;flex-direction:column;gap:10px;';



        // SWCM

        var swcmRows = [

            { metric: 'CUT Compliance', goal: '100%', eos: ibf.cutCompliance !== null ? ibf.cutCompliance + '%' + (ibf.cutCompliance >= 100 ? ' \u2705' : ' \u274C') : '' },

            { metric: 'Remaining CUT Trailers Due on Shift', eos: ibf.remainingCUTs !== null ? formatNum(ibf.remainingCUTs) + (ibf.remainingCUTs === 0 ? ' \u2705' : ' \u274C') : '' },

            { metric: 'SLIM Perfect Station Score \u2014 Receive Auto Decant', goal: '> 90%',

              p1: periodSlim.receiveAutoDecant.P1 !== null ? periodSlim.receiveAutoDecant.P1 + '%' : '',

              p2: periodSlim.receiveAutoDecant.P2 !== null ? periodSlim.receiveAutoDecant.P2 + '%' : '',

              p3: periodSlim.receiveAutoDecant.P3 !== null ? periodSlim.receiveAutoDecant.P3 + '%' : '',

              eos: periodSlim.receiveAutoDecant.EOS !== null ? periodSlim.receiveAutoDecant.EOS + '%' + (periodSlim.receiveAutoDecant.EOS >= 90 ? ' \u2705' : ' \u274C') : '' },

            { metric: 'SLIM Perfect Station Score \u2014 Manual Decant', goal: '> 90%',

              p1: periodSlim.receiveManualDecant.P1 !== null ? periodSlim.receiveManualDecant.P1 + '%' : '',

              p2: periodSlim.receiveManualDecant.P2 !== null ? periodSlim.receiveManualDecant.P2 + '%' : '',

              p3: periodSlim.receiveManualDecant.P3 !== null ? periodSlim.receiveManualDecant.P3 + '%' : '',

              eos: periodSlim.receiveManualDecant.EOS !== null ? periodSlim.receiveManualDecant.EOS + '%' + (periodSlim.receiveManualDecant.EOS >= 90 ? ' \u2705' : ' \u274C') : '' },

            (function() {

                // Check if plan was promoted on time

                // Day: must be by 06:30, Night: must be by 17:30

                var sp = introData.shiftPlanner;

                var planTime = sp && sp.planCreatedAt ? sp.planCreatedAt : null;

                var promoted = '';

                if (planTime) {

                    var dt = new Date(typeof planTime === 'number' ? planTime * 1000 : planTime);

                    var h = dt.getHours();

                    var m = dt.getMinutes();

                    var totalMin = h * 60 + m;

                    var deadline = shift === 'Day' ? (6 * 60 + 30) : (17 * 60 + 30); // 06:30 or 17:30

                    var timeStr = ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);

                    promoted = totalMin <= deadline ? 'Y (' + timeStr + ') \u2705' : 'N (' + timeStr + ') \u274C';

                }

                return { metric: 'INTRO Shift Planner promoted 30 min before start of Period', goal: 'Y', eos: promoted };

            })(),

            { metric: 'Stow Coaching Academy Audit', goal: '5+ by AM' },

            { metric: 'STOW GCA Compliance', goal: '1' },

        ];

        placeAutoValues(swcmRows, period);

        col3.appendChild(createSection('SWCM  (Standard Work Compliance)', swcmRows));



        // STANDARD WORK AUDITS (Apollo)

        var auditRows = CONFIG.apolloAudits.map(function(a) {

            var result = apolloData.find(function(r) { return r.name === a.name; });

            var goalNum = parseInt(getAuditGoal(a.name)) || 0;

            var p1Val = result && result.p1 !== null ? '' + result.p1 : '';

            var p2Val = result && result.p2 !== null ? '' + result.p2 : '';

            var p3Val = result && result.p3 !== null ? '' + result.p3 : '';

            var eosVal = result && result.eos !== null ? '' + result.eos : '';

            var eosNum = result ? (result.eos || 0) : 0;

            var warn = goalNum > 0 && eosNum < goalNum;

            // Add check/X to EOS based on goal comparison

            if (eosVal && goalNum > 0) eosVal = eosVal + (eosNum >= goalNum ? ' \u2705' : ' \u274C');

            return { metric: a.name, goal: getAuditGoal(a.name), p1: p1Val, p2: p2Val, p3: p3Val, eos: eosVal, warn: warn };

        });

        auditRows.push({ metric: 'IRDR', goal: '100% on site' });

        placeAutoValues(auditRows, period);

        col3.appendChild(createSection('STANDARD WORK AUDITS', auditRows));



        // INPUT METRICS per-floor

        col3.appendChild(createFloorMetricsSection(vantageData, sp));



        // FLOOR WALK CHECK section (manual entry, per-floor columns)

        (function() {

            var floorCheckMetrics = [

                { metric: 'Are all stairwells (S, E, W) free and clear?', goal: 'Y' },

                { metric: 'Break room lights on, cover in place, fan plugged in?', goal: 'Y' },

                { metric: 'How many chairs removed from E and W runs?', goal: '0' },

                { metric: 'Is the VRC area clear of totes, chairs, and garbage?', goal: 'Y' },

                { metric: 'How many stations are down?', goal: '0' },

                { metric: 'Is West side cleared of bad work, good work, excess totes and UPP?', goal: 'Y' },

            ];

            var fcSection = document.createElement('div');

            fcSection.style.cssText = 'border:1px solid #ddd;border-radius:6px;overflow:hidden;background:white;';

            var fcHeader = document.createElement('div');

            fcHeader.style.cssText = 'padding:6px 10px;background:#232f3e;color:white;font-size:11px;font-weight:bold;';

            fcHeader.textContent = 'FLOOR WALK CHECK';

            fcSection.appendChild(fcHeader);



            var fcTable = document.createElement('table');

            fcTable.style.cssText = 'width:100%;border-collapse:collapse;font-size:10px;background:white;';

            var floors = SITE_SETTINGS.FLOORS;

            var thHtml = '<tr style="background:#B9C9FE;"><th style="padding:3px 5px;text-align:left;border:1px solid #aaa;font-size:9px;">Metric</th><th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:35px;">Goal</th>';

            floors.forEach(function(f) { thHtml += '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:40px;">' + f + '</th>'; });

            thHtml += '</tr>';



            var tbHtml = '';

            floorCheckMetrics.forEach(function(m, idx) {

                tbHtml += '<tr style="background:' + (idx % 2 === 0 ? 'white' : '#f7f7f7') + ';">';

                tbHtml += '<td style="padding:3px 5px;border:1px solid #eee;font-weight:500;white-space:normal;max-width:160px;font-size:9px;">' + m.metric + '</td>';

                tbHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;color:#666;font-size:9px;">' + m.goal + '</td>';

                floors.forEach(function(f) {

                    var saved = getManualValue('FLOOR WALK CHECK', m.metric, f);

                    tbHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;font-weight:bold;cursor:text;min-width:30px;outline:none;" contenteditable="true" data-section="FLOOR WALK CHECK" data-metric="' + m.metric + '" data-col="' + f + '">' + (saved || '') + '</td>';

                });

                tbHtml += '</tr>';

            });



            fcTable.innerHTML = '<thead>' + thHtml + '</thead><tbody>' + tbHtml + '</tbody>';

            fcTable.addEventListener('focusout', function(e) {

                var td = e.target;

                if (td.hasAttribute('contenteditable') && td.dataset.section) {

                    setManualValue(td.dataset.section, td.dataset.metric, td.dataset.col, td.textContent.trim());

                }

            });

            fcSection.appendChild(fcTable);

            col3.appendChild(fcSection);

        })();



        content.appendChild(col3);



        container.appendChild(header);

        if (quickLinkBar) {

            var qlDiv = document.createElement('div');

            qlDiv.innerHTML = quickLinkBar;

            container.appendChild(qlDiv.firstChild);

        }

        // Timestamp bar

        var now = new Date();

        var timeStr = now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true, timeZone: 'America/Chicago' });

        var timestampBar = document.createElement('div');

        timestampBar.style.cssText = 'padding:2px 16px;background:#1a1a2e;color:#aaa;font-size:9px;text-align:right;';

        timestampBar.textContent = 'Data as of: ' + timeStr + ' CDT';

        container.appendChild(timestampBar);



        container.appendChild(content);



        // Toggle

        header.onclick = function(e) {

            if (e.target.id === 'sync-copy-btn' || e.target.id === 'sync-refresh-btn' || e.target.id === 'sync-pdf-btn' || e.target.id === 'sync-clear-btn') return;

            var isHidden = content.style.display === 'none';

            content.style.display = isHidden ? 'grid' : 'none';

            document.getElementById('sync-toggle').textContent = isHidden ? '\u25BC' : '\u25B6';

        };



        // Button handlers (delayed to ensure DOM ready)

        setTimeout(function() {

            var copyBtn = document.getElementById('sync-copy-btn');

            if (copyBtn) {

                copyBtn.onclick = function(e) {

                    e.stopPropagation();

                    copyBtn.textContent = '\u23F3 Capturing...';

                    // Ensure content is visible for capture

                    var contentEl = document.getElementById('sync-dashboard-content');

                    var wasHidden = contentEl && contentEl.style.display === 'none';

                    if (wasHidden) contentEl.style.display = 'grid';

                    // Small delay to let browser render before capture

                    setTimeout(function() {

                    html2canvas(container, {

                        backgroundColor: '#fafafa',

                        scale: 2,

                        useCORS: true,

                        logging: true,

                        allowTaint: true,

                        foreignObjectRendering: false,

                        removeContainer: true

                    }).then(function(canvas) {

                        canvas.toBlob(function(blob) {

                            if (!blob) {

                                copyBtn.textContent = '\u274C No image';

                                setTimeout(function() { copyBtn.textContent = '\u{1F4CB} Copy'; }, 2000);

                                if (wasHidden) contentEl.style.display = 'none';

                                return;

                            }

                            navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]).then(function() {

                                copyBtn.textContent = '\u2705 Copied!';

                                setTimeout(function() { copyBtn.textContent = '\u{1F4CB} Copy'; }, 2000);

                                if (wasHidden) contentEl.style.display = 'none';

                            }).catch(function(err) {

                                console.log('[IB Sync] Clipboard write failed:', err);

                                // Fallback: open image in new tab

                                var url = canvas.toDataURL('image/png');

                                window.open(url, '_blank');

                                copyBtn.textContent = '\u2705 Opened in tab';

                                setTimeout(function() { copyBtn.textContent = '\u{1F4CB} Copy'; }, 2000);

                                if (wasHidden) contentEl.style.display = 'none';

                            });

                        }, 'image/png');

                    }).catch(function(err) {

                        console.log('[IB Sync] html2canvas error:', err);

                        copyBtn.textContent = '\u274C Failed';

                        setTimeout(function() { copyBtn.textContent = '\u{1F4CB} Copy'; }, 2000);

                        if (wasHidden) contentEl.style.display = 'none';

                    });

                    }, 300);

                };

            }

            var refreshBtn = document.getElementById('sync-refresh-btn');

            if (refreshBtn) {

                refreshBtn.onclick = function(e) {

                    e.stopPropagation();

                    refreshBtn.textContent = '\u23F3 Loading...';

                    // Reload the page to refresh FCLM data (same URL with current params)

                    window.location.reload();

                };

            }

            var pdfBtn = document.getElementById('sync-pdf-btn');

            if (pdfBtn) {

                pdfBtn.onclick = function(e) {

                    e.stopPropagation();

                    pdfBtn.textContent = '\u23F3 Generating...';

                    exportToPDF(container);

                    setTimeout(function() { pdfBtn.textContent = '\u{1F4C4} PDF'; }, 2000);

                };

            }



            var excelBtn = document.getElementById('sync-excel-btn');

            if (excelBtn) {

                excelBtn.onclick = function(e) {

                    e.stopPropagation();

                    excelBtn.textContent = '\u23F3 Exporting...';

                    exportToExcel(container, shift, period);

                    setTimeout(function() { excelBtn.textContent = '\u{1F4CA} Excel'; }, 2000);

                };

            }

            var clearBtn = document.getElementById('sync-clear-btn');

            if (clearBtn) {

                clearBtn.onclick = function(e) {

                    e.stopPropagation();

                    if (confirm('Clear all manually entered data? This cannot be undone.')) {

                        localStorage.removeItem(MANUAL_DATA_KEY);

                        // Immediately wipe all contenteditable cells in the dashboard

                        container.querySelectorAll('[contenteditable="true"]').forEach(function(cell) {

                            if (cell.dataset.section) cell.textContent = '';

                        });

                        clearBtn.textContent = '\u2705 Cleared!';

                        setTimeout(function() { clearBtn.textContent = '\u{1F5D1} Clear'; }, 2000);

                    }

                };

            }

        }, 100);



        return container;

    }



    // Helper: place auto values into the current period column

    function placeAutoValues(rows, currentPeriod) {

        var col = currentPeriod.toLowerCase();

        rows.forEach(function(r) {

            if (r.auto && !r[col]) {

                r[col] = r.auto;

            }

        });

    }



    // --- Section builder ---

    var MANUAL_DATA_KEY = 'ppr_sync_manual_' + WAREHOUSE;



    function getManualData() {

        try {

            var stored = localStorage.getItem(MANUAL_DATA_KEY);

            if (stored) return JSON.parse(stored);

        } catch(e) {}

        return {};

    }



    function saveManualData(data) {

        localStorage.setItem(MANUAL_DATA_KEY, JSON.stringify(data));

    }



    function getManualValue(sectionTitle, metric, col) {

        var data = getManualData();

        var key = sectionTitle + '|' + metric + '|' + col;

        return data[key] || '';

    }



    function setManualValue(sectionTitle, metric, col, value) {

        var data = getManualData();

        var key = sectionTitle + '|' + metric + '|' + col;

        if (value) {

            data[key] = value;

        } else {

            delete data[key];

        }

        saveManualData(data);

    }



    function createSection(title, rows, goalLabel) {

        goalLabel = goalLabel || 'Goal';

        var section = document.createElement('div');

        section.style.cssText = 'border:1px solid #ddd;border-radius:6px;overflow:hidden;background:white;';



        var sectionHeader = document.createElement('div');

        sectionHeader.style.cssText = 'padding:6px 10px;background:#232f3e;color:white;font-size:11px;font-weight:bold;';

        sectionHeader.textContent = title;

        section.appendChild(sectionHeader);



        var table = document.createElement('table');

        table.style.cssText = 'width:100%;border-collapse:collapse;font-size:10px;background:white;';



        var thead = document.createElement('thead');

        thead.innerHTML = '<tr style="background:#B9C9FE;">' +

            '<th style="padding:3px 5px;text-align:left;border:1px solid #aaa;font-size:9px;">Metric</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:50px;">' + goalLabel + '</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:45px;">P1</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:45px;">P2</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:45px;">P3</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;width:45px;font-weight:bold;">EOS</th>' +

            '</tr>';

        table.appendChild(thead);



        var tbody = document.createElement('tbody');

        rows.forEach(function(row, idx) {

            var tr = document.createElement('tr');

            tr.style.backgroundColor = idx % 2 === 0 ? 'white' : '#f7f7f7';

            var valColor = row.warn ? 'red' : '#333';



            var cols = [

                { key: 'goal', val: row.goal || '', style: 'color:#666;font-size:9px;', editable: false },

                { key: 'p1', val: row.p1 || '', style: 'color:' + valColor + ';font-weight:bold;', editable: true },

                { key: 'p2', val: row.p2 || '', style: 'color:' + valColor + ';font-weight:bold;', editable: true },

                { key: 'p3', val: row.p3 || '', style: 'color:' + valColor + ';font-weight:bold;', editable: true },

                { key: 'eos', val: row.eos || '', style: 'color:' + valColor + ';font-weight:bold;background:#f8f8f0;', editable: true }

            ];



            var metricTd = '<td style="padding:3px 5px;border:1px solid #eee;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px;" title="' + (row.metric || '') + '">' + (row.metric || '') + '</td>';

            var colsHtml = '';

            cols.forEach(function(col) {

                var cellVal = col.val;

                var isEditable = col.editable && !cellVal;

                // Load saved manual value if cell is blank

                if (isEditable) {

                    var saved = getManualValue(title, row.metric, col.key);

                    if (saved) cellVal = saved;

                }

                var editAttr = isEditable ? ' contenteditable="true"' : '';

                var editStyle = isEditable ? 'cursor:text;min-width:30px;outline:none;' : '';

                var editDataAttr = isEditable ? ' data-section="' + title + '" data-metric="' + (row.metric || '') + '" data-col="' + col.key + '"' : '';

                colsHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;' + col.style + editStyle + '"' + editAttr + editDataAttr + '>' + cellVal + '</td>';

            });



            tr.innerHTML = metricTd + colsHtml;

            tbody.appendChild(tr);

        });

        table.appendChild(tbody);



        // Save on blur for editable cells

        table.addEventListener('focusout', function(e) {

            var td = e.target;

            if (td.hasAttribute('contenteditable') && td.dataset.section) {

                setManualValue(td.dataset.section, td.dataset.metric, td.dataset.col, td.textContent.trim());

            }

        });



        section.appendChild(table);

        return section;

    }



    // --- CI Projects section ---

    function createCISection() {

        var section = document.createElement('div');

        section.style.cssText = 'border:1px solid #ddd;border-radius:6px;overflow:hidden;background:white;';

        var sectionHeader = document.createElement('div');

        sectionHeader.style.cssText = 'padding:6px 10px;background:#232f3e;color:white;font-size:11px;font-weight:bold;';

        sectionHeader.textContent = 'CI PROJECTS / ACTION ITEMS';

        section.appendChild(sectionHeader);



        var table = document.createElement('table');

        table.style.cssText = 'width:100%;border-collapse:collapse;font-size:10px;background:white;';



        var ciTitle = 'CI PROJECTS / ACTION ITEMS';

        var ciRows = [

            { metric: 'Action #1', cols: ['Owner', 'ECD', 'Status', 'Notes'] },

            { metric: 'Action #2', cols: ['Owner', 'ECD', 'Status', 'Notes'] },

            { metric: 'Action #3', cols: ['Owner', 'ECD', 'Status', 'Notes'] },

        ];



        var headerHtml = '<thead><tr style="background:#B9C9FE;">' +

            '<th style="padding:3px 5px;text-align:left;border:1px solid #aaa;font-size:9px;">Metric</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">Owner</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">ECD</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">Status</th>' +

            '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">Notes</th>' +

            '</tr></thead>';



        var bodyHtml = '<tbody>';

        ciRows.forEach(function(row, idx) {

            var bg = idx % 2 === 0 ? 'white' : '#f7f7f7';

            var savedMetric = getManualValue(ciTitle, row.metric, 'metric') || row.metric;

            bodyHtml += '<tr style="background:' + bg + ';">';

            bodyHtml += '<td style="padding:3px 5px;border:1px solid #eee;cursor:text;outline:none;white-space:pre-wrap;word-wrap:break-word;vertical-align:top;" contenteditable="true" data-section="' + ciTitle + '" data-metric="' + row.metric + '" data-col="metric">' + savedMetric + '</td>';

            row.cols.forEach(function(col) {

                var savedVal = getManualValue(ciTitle, row.metric, col);

                bodyHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;cursor:text;outline:none;white-space:pre-wrap;word-wrap:break-word;vertical-align:top;" contenteditable="true" data-section="' + ciTitle + '" data-metric="' + row.metric + '" data-col="' + col + '">' + savedVal + '</td>';

            });

            bodyHtml += '</tr>';

        });

        bodyHtml += '</tbody>';



        table.innerHTML = headerHtml + bodyHtml;



        // Save on blur

        table.addEventListener('focusout', function(e) {

            var td = e.target;

            if (td.hasAttribute('contenteditable') && td.dataset.section) {

                setManualValue(td.dataset.section, td.dataset.metric, td.dataset.col, td.textContent.trim());

            }

        });



        section.appendChild(table);

        return section;

    }



    // --- Per-floor INPUT METRICS section ---

    function createFloorMetricsSection(vantageData, shiftPlannerData) {

        var section = document.createElement('div');

        section.style.cssText = 'border:1px solid #ddd;border-radius:6px;overflow:hidden;background:white;';

        var sectionHeader = document.createElement('div');

        sectionHeader.style.cssText = 'padding:6px 10px;background:#232f3e;color:white;font-size:11px;font-weight:bold;';

        sectionHeader.textContent = 'INPUT METRICS (By Floor)';

        section.appendChild(sectionHeader);



        var zones = vantageData.zones || {};

        var zoneNames = CONFIG.vantage.zones;

        var labels = SITE_SETTINGS.FLOORS;

        var period = getCurrentPeriod();



        var table = document.createElement('table');

        table.style.cssText = 'width:100%;border-collapse:collapse;font-size:10px;background:white;';



        var headerHtml = '<tr style="background:#B9C9FE;"><th style="padding:3px 5px;text-align:left;border:1px solid #aaa;font-size:9px;">Metric</th><th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">Goal</th>';

        labels.forEach(function(l) { headerHtml += '<th style="padding:3px 5px;text-align:center;border:1px solid #aaa;font-size:9px;">' + l + '</th>'; });

        headerHtml += '</tr>';



        var metricsMap = [

            { metric: 'ETI Total', field: 'stowRate', goal: '>210', goalVal: 210, lowerBetter: false, yellowMax: 200 },

            { metric: 'OOWA%', field: 'oowa', goal: '< 2%', goalVal: CONFIG.goals.oowa, lowerBetter: true, yellowMax: 2.5 },

            { metric: 'UPF', field: 'upf', goal: '> 12', goalVal: CONFIG.goals.upf, lowerBetter: false, yellowMax: 9 },

            { metric: 'Stow Cycle Time (s)', field: 'stowCycleTime', goal: '\u2264 13s', goalVal: CONFIG.goals.stowCycleTime, lowerBetter: true, yellowMax: 15 },

            { metric: 'NSTA%', field: 'nsta', goal: '< 15%', goalVal: CONFIG.goals.nsta, lowerBetter: true, yellowMax: 15.5 },

            { metric: 'Stow HC (Plan)', field: null, fromPlan: true, planField: 'stowHCByFloor' },

            { metric: 'Stow HC (Actual)', field: null, editable: true },

            { metric: 'WS HC (Plan)', field: null, fromPlan: true, planField: 'waterSpiderHC' },

            { metric: 'WS HC (Actual)', field: null, editable: true },

            { metric: 'PS HC (Plan)', field: null, fromPlan: true, planField: 'problemSolveHC' },

            { metric: 'PS HC (Actual)', field: null, editable: true },

            { metric: 'PS Piles', field: null },

        ];



        var bodyHtml = '';

        metricsMap.forEach(function(m, idx) {

            var bg = idx % 2 === 0 ? 'white' : '#f7f7f7';

            bodyHtml += '<tr style="background:' + bg + ';"><td style="padding:3px 5px;border:1px solid #eee;font-weight:500;">' + m.metric + '</td>';

            bodyHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;font-size:9px;color:#555;">' + (m.goal || '') + '</td>';



            // Determine 1st/2nd place trophies for numeric metrics with field data

            var trophyMap = {}; // lIdx -> '🥇' or '🥈'

            if (m.field && !m.fromPlan) {

                var floorVals = [];

                labels.forEach(function(label, lIdx) {

                    var zoneName = zoneNames[lIdx];

                    var v = (zones[zoneName] && zones[zoneName][m.field]) ? parseFloat(zones[zoneName][m.field]) : null;

                    if (v !== null && !isNaN(v)) floorVals.push({ idx: lIdx, val: v });

                });

                if (floorVals.length >= 2) {

                    // Sort: lowerBetter = ascending (lowest wins), else descending (highest wins)

                    floorVals.sort(function(a, b) {

                        return m.lowerBetter ? a.val - b.val : b.val - a.val;

                    });

                    // ETI Total: higher is better (no lowerBetter flag = default higher is better)

                    if (m.lowerBetter === undefined) floorVals.sort(function(a, b) { return b.val - a.val; });

                    trophyMap[floorVals[0].idx] = ' \uD83E\uDD47';

                    trophyMap[floorVals[1].idx] = ' \uD83E\uDD48';

                }

            }



            labels.forEach(function(label, lIdx) {

                var val = '';

                var color = '#333';

                var zoneName = zoneNames[lIdx];



                if (m.field && zones[zoneName]) {

                    val = zones[zoneName][m.field] || '';

                    if (val && m.warn && m.warn(val)) color = 'red';

                }



                // Stow HC from shift planner per floor

                if (m.fromPlan && m.planField === 'stowHCByFloor' && shiftPlannerData) {

                    // Show P1/P2/P3 per floor as "#/#/#"

                    var floorObj = shiftPlannerData.stowHCByFloor[label];

                    if (floorObj) {

                        var sP1 = floorObj.P1 !== null ? floorObj.P1 : 0;

                        var sP2 = floorObj.P2 !== null ? floorObj.P2 : 0;

                        var sP3 = floorObj.P3 !== null ? floorObj.P3 : 0;

                        if (floorObj.P1 !== null || floorObj.P2 !== null || floorObj.P3 !== null) {

                            val = 'P1:' + sP1 + ' P2:' + sP2 + ' P3:' + sP3;

                        }

                    }

                }

                if (m.fromPlan && m.planField === 'waterSpiderHC' && shiftPlannerData) {

                    // Water spider HC: divide evenly across 4 floors, remainder goes to A02

                    // Show P1/P2/P3 per floor as "#/#/#"

                    var wsP1 = shiftPlannerData.waterSpiderHC.P1;

                    var wsP2 = shiftPlannerData.waterSpiderHC.P2;

                    var wsP3 = shiftPlannerData.waterSpiderHC.P3;

                    var parts = [];

                    var hasRed = false;

                    [wsP1, wsP2, wsP3].forEach(function(wsTotal) {

                        if (wsTotal !== null && wsTotal > 0) {

                            var perFloor = Math.floor(wsTotal / 4);

                            var remainder = wsTotal % 4;

                            var floorVal = lIdx < remainder ? (perFloor + 1) : perFloor;

                            if (lIdx < remainder) hasRed = true;

                            parts.push(floorVal);

                        } else {

                            parts.push(0);

                        }

                    });

                    if (wsP1 !== null || wsP2 !== null || wsP3 !== null) {

                        val = 'P1:' + parts[0] + ' P2:' + parts[1] + ' P3:' + parts[2];

                        if (hasRed) color = 'red';

                    }

                }

                if (m.fromPlan && m.planField === 'problemSolveHC' && shiftPlannerData) {

                    // Problem Solve HC: same logic as WS HC — divide evenly, show P1/P2/P3 as "#/#/#"

                    var psP1 = shiftPlannerData.problemSolveHC.P1;

                    var psP2 = shiftPlannerData.problemSolveHC.P2;

                    var psP3 = shiftPlannerData.problemSolveHC.P3;

                    var psParts = [];

                    var psHasRed = false;

                    [psP1, psP2, psP3].forEach(function(psTotal) {

                        if (psTotal !== null && psTotal > 0) {

                            var perFloor = Math.floor(psTotal / 4);

                            var remainder = psTotal % 4;

                            var floorVal = lIdx < remainder ? (perFloor + 1) : perFloor;

                            if (lIdx < remainder) psHasRed = true;

                            psParts.push(floorVal);

                        } else {

                            psParts.push(0);

                        }

                    });

                    if (psP1 !== null || psP2 !== null || psP3 !== null) {

                        val = 'P1:' + psParts[0] + ' P2:' + psParts[1] + ' P3:' + psParts[2];

                        if (psHasRed) color = 'red';

                    }

                }



                var isEditable = !val;



                // Add check/X for metrics with goals (only for auto-populated values, not editable)

                if (val && m.goalVal !== undefined && !isEditable) {

                    var numVal = parseFloat(val);

                    if (!isNaN(numVal)) {

                        if (m.lowerBetter) {

                            if (numVal <= m.goalVal) val = val + ' \u2705';

                            else if (m.yellowMax !== undefined && numVal <= m.yellowMax) val = val + ' \u26A0\uFE0F';

                            else val = val + ' \u274C';

                        } else {

                            if (numVal >= m.goalVal) val = val + ' \u2705';

                            else if (m.yellowMax !== undefined && numVal >= m.yellowMax) val = val + ' \u26A0\uFE0F';

                            else val = val + ' \u274C';

                        }

                    }

                }



                // Add trophy for 1st/2nd place floors

                if (trophyMap[lIdx] && val && !isEditable) val = trophyMap[lIdx] + ' ' + val;



                var floorSectionTitle = 'INPUT METRICS (By Floor)';

                if (isEditable) {

                    var saved = getManualValue(floorSectionTitle, m.metric, label);

                    if (saved) val = saved;

                }

                var editAttr = isEditable ? ' contenteditable="true"' : '';

                var editStyle = isEditable ? 'cursor:text;min-width:30px;outline:none;' : '';

                var editDataAttr = isEditable ? ' data-section="' + floorSectionTitle + '" data-metric="' + m.metric + '" data-col="' + label + '"' : '';

                bodyHtml += '<td style="padding:3px 5px;text-align:center;border:1px solid #eee;color:' + color + ';font-weight:bold;' + editStyle + '"' + editAttr + editDataAttr + '>' + val + '</td>';

            });

            bodyHtml += '</tr>';

        });



        table.innerHTML = '<thead>' + headerHtml + '</thead><tbody>' + bodyHtml + '</tbody>';



        // Save on blur for editable cells

        table.addEventListener('focusout', function(e) {

            var td = e.target;

            if (td.hasAttribute('contenteditable') && td.dataset.section) {

                setManualValue(td.dataset.section, td.dataset.metric, td.dataset.col, td.textContent.trim());

            }

        });



        section.appendChild(table);

        return section;

    }



    // ==========================================

    // PDF EXPORT

    // ==========================================

    

    // ==========================================

    // EXPORT TO EXCEL

    // ==========================================

    function exportToExcel(container, shift, period) {

        try {

            // === STYLE DEFINITIONS ===

            var navyFill = { fgColor: { rgb: '232F3E' } };

            var lavenderFill = { fgColor: { rgb: 'B9C9FE' } };

            var lightGrayFill = { fgColor: { rgb: 'F7F7F7' } };

            var whiteFill = { fgColor: { rgb: 'FFFFFF' } };

            var whiteFont = { color: { rgb: 'FFFFFF' }, bold: true, sz: 12 };

            var sectionFont = { color: { rgb: 'FFFFFF' }, bold: true, sz: 10 };

            var colHdrFont = { bold: true, sz: 9, color: { rgb: '333333' } };

            var metricFont = { bold: true, sz: 9, color: { rgb: '333333' } };

            var valFont = { bold: true, sz: 9, color: { rgb: '333333' } };

            var redFont = { bold: true, sz: 9, color: { rgb: 'DC3545' } };

            var greenFont = { bold: true, sz: 9, color: { rgb: '28A745' } };

            var goalFont = { sz: 9, color: { rgb: '666666' } };

            var thinBdr = { top:{style:'thin',color:{rgb:'CCCCCC'}}, bottom:{style:'thin',color:{rgb:'CCCCCC'}}, left:{style:'thin',color:{rgb:'CCCCCC'}}, right:{style:'thin',color:{rgb:'CCCCCC'}} };

            var hdrBdr = { top:{style:'thin',color:{rgb:'AAAAAA'}}, bottom:{style:'thin',color:{rgb:'AAAAAA'}}, left:{style:'thin',color:{rgb:'AAAAAA'}}, right:{style:'thin',color:{rgb:'AAAAAA'}} };



            var wb = XLSX.utils.book_new();

            var ws = {};

            var merges = [];



            // Column offsets (0-indexed): Col1=0, Col2=7, Col3=14

            var COL_OFFSETS = [0, 7, 14];

            var COL_WIDTH = 6; // Each section is 6 columns wide



            // === MAIN HEADER (spans full width) ===

            var headerText = WAREHOUSE + ' Inbound - Leadership Sync Tracker - ' + shift.toUpperCase() + ' SHIFT';



            // Get operational date from FCLM URL (not today's date)

            var xlParams = new URLSearchParams(window.location.search);

            var xlSpan = xlParams.get('spanType') || '';

            var xlDateRaw = '';

            if (xlSpan === 'Intraday') xlDateRaw = (xlParams.get('startDateIntraday') || '').split(/[\s+T]/)[0];

            else if (xlSpan === 'Week') xlDateRaw = xlParams.get('startDateWeek') || '';

            else xlDateRaw = xlParams.get('startDateDay') || '';

            var opDate = xlDateRaw ? new Date(xlDateRaw.replace(/\//g, '-') + 'T12:00:00') : new Date();

            var dateStr = opDate.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });



            // Get S1 Goal from the board header (already rendered)

            var s1Text = '';

            var headerEl = container.querySelector('[id*="sync"] div, .sync-header');

            if (!headerEl) headerEl = container.querySelector('div');

            var headerContent = container.textContent || '';

            var s1Match = headerContent.match(/S1 Goal:[^\)]+\)/);

            if (s1Match) s1Text = ' | ' + s1Match[0];



            var periodText = period + ' | ' + dateStr + s1Text;



            // Row 0: Main header

            for (var hc = 0; hc < 20; hc++) {

                var hRef = XLSX.utils.encode_cell({r:0, c:hc});

                ws[hRef] = { v: hc === 0 ? headerText : '', t: 's', s: { fill: navyFill, font: whiteFont, alignment: {vertical:'center'} } };

            }

            merges.push({ s:{r:0,c:0}, e:{r:0,c:19} });



            // Row 1: Period + date

            for (var pc = 0; pc < 20; pc++) {

                var pRef = XLSX.utils.encode_cell({r:1, c:pc});

                ws[pRef] = { v: pc === 0 ? periodText : '', t: 's', s: { fill: navyFill, font: { color:{rgb:'FF9900'}, bold:true, sz:10 }, alignment:{vertical:'center'} } };

            }

            merges.push({ s:{r:1,c:0}, e:{r:1,c:19} });



            // Row 2: blank spacer

            var startRow = 3;



            // === GATHER SECTION DATA FROM DOM ===

            // The dashboard content div has 3 child column divs

            var dashContent = container.querySelector('#sync-dashboard-content') || container;

            var columnDivs = dashContent.children ? Array.from(dashContent.children) : [];



            // If we can identify the 3 column divs

            var columns = [[], [], []]; // Each column stores array of section objects



            // Process each column div

            columnDivs.forEach(function(colDiv, colIdx) {

                if (colIdx >= 3) return;

                // Each section is a div with: header div + table

                var sectionDivs = colDiv.querySelectorAll(':scope > div');

                sectionDivs.forEach(function(secDiv) {

                    var tbl = secDiv.querySelector('table');

                    if (!tbl) return;



                    // Section header is the first child div (before the table)

                    var secTitle = '';

                    var children = secDiv.children;

                    for (var chi = 0; chi < children.length; chi++) {

                        var child = children[chi];

                        if (child.tagName === 'TABLE') break;

                        if (child.tagName === 'DIV' && child.textContent.trim()) {

                            secTitle = child.textContent.trim();

                            break;

                        }

                    }

                    var tableRows = [];



                    tbl.querySelectorAll('tr').forEach(function(tr) {

                        var cells = tr.querySelectorAll('td, th');

                        var isHdr = tr.querySelectorAll('th').length > 0;

                        var rowCells = [];

                        cells.forEach(function(cell) {

                            var text = cell.textContent.trim();

                            var color = cell.style.color || (cell.parentElement ? cell.parentElement.style.color : '') || '';

                            var bgColor = cell.style.backgroundColor || '';

                            rowCells.push({ text: text, color: color, bg: bgColor });

                        });

                        if (rowCells.length > 0) tableRows.push({ cells: rowCells, isHeader: isHdr });

                    });



                    columns[colIdx].push({ title: secTitle, rows: tableRows });

                });

            });



            // === WRITE SECTIONS TO WORKSHEET ===

            var colRowCursors = [startRow, startRow, startRow]; // Track current row per column



            columns.forEach(function(colSections, colIdx) {

                var colOffset = COL_OFFSETS[colIdx];

                var curRow = startRow;



                colSections.forEach(function(sec) {

                    // Section header row

                    if (sec.title) {

                        for (var sh = 0; sh < COL_WIDTH; sh++) {

                            var shRef = XLSX.utils.encode_cell({r:curRow, c:colOffset+sh});

                            ws[shRef] = { v: sh === 0 ? sec.title : '', t: 's', s: { fill: navyFill, font: sectionFont, border: hdrBdr, alignment:{vertical:'center'} } };

                        }

                        merges.push({ s:{r:curRow, c:colOffset}, e:{r:curRow, c:colOffset+COL_WIDTH-1} });

                        curRow++;

                    }



                    // Table rows

                    sec.rows.forEach(function(row, rIdx) {

                        var isAlt = rIdx % 2 === 1;

                        row.cells.forEach(function(cell, cIdx) {

                            if (cIdx >= COL_WIDTH) return;

                            var cellRef = XLSX.utils.encode_cell({r:curRow, c:colOffset+cIdx});

                            var cellVal = cell.text;

                            var cellType = 's';



                            // Parse numbers

                            // Detect performance markers

                            var hasCheck = cellVal.indexOf('\u2705') !== -1;

                            var hasX = cellVal.indexOf('\u274C') !== -1;

                            var hasWarn = cellVal.indexOf('\u26A0') !== -1;



                            var numStr = cellVal.replace(/,/g, '').replace('%', '').replace(/[\u2705\u274C\u26A0\uFE0F\s]/g, '').trim();

                            var num = parseFloat(numStr);

                            if (!isNaN(num) && numStr.length > 0 && cellVal.match(/^[\d,\.\-%]/)) {

                                // Keep markers with the value as text so Excel shows them

                                if (hasCheck || hasX || hasWarn) {

                                    cellVal = cellVal.trim(); // Keep as text with emoji

                                    cellType = 's';

                                } else {

                                    cellVal = num;

                                    cellType = 'n';

                                }

                            }



                            // Determine style

                            var cellStyle = {};

                            if (row.isHeader) {

                                cellStyle = { fill: lavenderFill, font: colHdrFont, border: hdrBdr, alignment: { horizontal: cIdx === 0 ? 'left' : 'center', vertical:'center', wrapText:true } };

                            } else {

                                var bgFill = isAlt ? lightGrayFill : whiteFill;

                                var font = valFont;



                                // Red/green detection from inline styles AND markers

                                var clr = cell.color.toLowerCase();

                                if (clr.indexOf('red') !== -1 || clr.indexOf('dc3545') !== -1 || hasX) font = redFont;

                                else if (clr.indexOf('green') !== -1 || clr.indexOf('28a745') !== -1 || hasCheck) font = greenFont;

                                else if (hasWarn) font = { bold: true, sz: 9, color: { rgb: 'FFC107' } };



                                if (cIdx === 0) {

                                    cellStyle = { fill: bgFill, font: metricFont, border: thinBdr, alignment:{horizontal:'left', vertical:'center', wrapText:true} };

                                } else if (cIdx === 1) {

                                    cellStyle = { fill: bgFill, font: goalFont, border: thinBdr, alignment:{horizontal:'center', vertical:'center', wrapText:true} };

                                } else {

                                    cellStyle = { fill: bgFill, font: font, border: thinBdr, alignment:{horizontal:'center', vertical:'center', wrapText:true} };

                                }



                                // Check cell background for green/red highlighting

                                if (cell.bg.indexOf('green') !== -1 || cell.bg.indexOf('28a745') !== -1) {

                                    cellStyle.fill = { fgColor: { rgb: 'D4EDDA' } };

                                } else if (cell.bg.indexOf('red') !== -1 || cell.bg.indexOf('dc3545') !== -1) {

                                    cellStyle.fill = { fgColor: { rgb: 'F8D7DA' } };

                                }

                            }



                            ws[cellRef] = { v: cellVal, t: cellType, s: cellStyle };

                        });

                        curRow++;

                    });



                    // Blank row between sections in same column

                    curRow++;

                });

                colRowCursors[colIdx] = curRow;

            });



            // === SET RANGE AND FORMATTING ===

            var maxRow = Math.max.apply(null, colRowCursors);

            ws['!ref'] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:maxRow, c:19}});

            ws['!merges'] = merges;



            // Column widths

            ws['!cols'] = [];

            for (var ci = 0; ci < 20; ci++) {

                var colMod = ci % 7;

                if (colMod === 6) { ws['!cols'].push({wch: 2}); } // spacer column

                else if (colMod === 0) { ws['!cols'].push({wch: 28}); } // metric name

                else if (colMod === 1) { ws['!cols'].push({wch: 11}); } // goal/plan

                else { ws['!cols'].push({wch: 10}); } // P1/P2/P3/EOS

            }



            XLSX.utils.book_append_sheet(wb, ws, 'IB Sync');



            var filename = WAREHOUSE + '_IB_Sync_' + shift + '_' + period + '_' + new Date().toISOString().slice(0,10) + '.xlsx';

            XLSX.writeFile(wb, filename);

            console.log('[IB Sync] Excel exported:', filename);

        } catch(e) {

            console.error('[IB Sync] Excel export error:', e);

            alert('Excel export failed: ' + e.message);

        }

    }



    function exportToPDF(container) {

        // Create a new window with just the dashboard content in landscape

        var printWindow = window.open('', '_blank', 'width=1400,height=900');

        var dashContent = container.querySelector('[style*="grid"]') || container;



        printWindow.document.write('<!DOCTYPE html><html><head><title>' + WAREHOUSE + ' IB Sync Tracker</title>');

        printWindow.document.write('<style>');

        printWindow.document.write('@page { size: landscape; margin: 0.15in; }');

        printWindow.document.write('body { margin: 0; padding: 5px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }');

        printWindow.document.write('.header { background: #232f3e; color: white; padding: 8px 15px; font-size: 14px; font-weight: bold; margin-bottom: 10px; border-radius: 6px; display: flex; justify-content: space-between; align-items: center; }');

        printWindow.document.write('.header .badge { background: #28a745; padding: 3px 10px; border-radius: 12px; font-size: 11px; }');

        printWindow.document.write('.content { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; }');

        printWindow.document.write('.content > div { display: flex; flex-direction: column; gap: 4px; }');

        printWindow.document.write('table { width: 100%; border-collapse: collapse; font-size: 7px; }');

        printWindow.document.write('th, td { padding: 1px 2px; border: 1px solid #ccc; }');

        printWindow.document.write('th { background: #B9C9FE; font-size: 6.5px; }');

        printWindow.document.write('div[style*="border-radius"] { margin-bottom: 3px; overflow: hidden; }');

        printWindow.document.write('div[style*="font-size: 11px"] { font-size: 9px !important; padding: 4px 8px !important; }');

        printWindow.document.write('@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } html, body { width: 100%; height: 100%; overflow: hidden; } }');

        printWindow.document.write('</style></head><body>');



        // Header

        var shift = getCurrentShift();

        var period = getCurrentPeriod();

        var dateStr = getSelectedDate();

        printWindow.document.write('<div class="header"><span>' + WAREHOUSE + ' Inbound \u2014 Leadership Sync Tracker \u00b7 ' + shift.toUpperCase() + ' SHIFT</span>');

        printWindow.document.write('<span class="badge">' + period + ' | ' + dateStr + '</span></div>');



        // Dashboard content

        printWindow.document.write('<div class="content">');

        printWindow.document.write(dashContent.innerHTML);

        printWindow.document.write('</div>');



        printWindow.document.write('</body></html>');

        printWindow.document.close();



        // Trigger print dialog (Save as PDF)

        setTimeout(function() {

            printWindow.print();

        }, 500);

    }



    // ==========================================

    // CLIPBOARD EXPORT

    // ==========================================

    function copyToClipboard(fclmData, apolloData, vantageData, introData) {

        var period = getCurrentPeriod();

        var shift = getCurrentShift();

        var now = new Date().toLocaleString();

        var sp = introData.shiftPlanner;

        var tp = introData.trailerPlanner;

        var ibf = introData.ibFlow;



        var text = WAREHOUSE + ' IB SYNC \u2014 ' + shift + ' Shift \u00b7 ' + period + ' \u00b7 S1 Goal: ' + (formatNum(sp.ibVolumePlan) || 'TBD') + ' \u00b7 ' + now + '\n';

        text += '============================================================\n\n';



        // NTP

        text += 'NAIL THE PLAN\n' + '----------------------------------------\n';

        text += 'IB Volume Plan: ' + (formatNum(sp.ibVolumePlan) || '--') + '\n';

        text += 'Trailer Plan Volume: ' + (formatNum(tp.trailerPlanVolume) || '--') + '\n';

        text += 'WIP: ' + (formatNum(ibf.wip) || '--') + '\n';

        text += 'Fluid%: ' + (tp.fluidPercent !== null ? tp.fluidPercent + '%' : '--') + ' | Case%: ' + (tp.casePercent !== null ? tp.casePercent + '%' : '--') + ' | Smalls%: ' + (sp.smallsMixPercent[period] !== null ? sp.smallsMixPercent[period] + '%' : '--') + '\n';

        text += 'Stow HC: ' + (formatNum(sp.stowHC[period]) || '--') + ' [A02:' + (formatNum(sp.stowHCByFloor.A02[period]) || '--') + ' A03:' + (formatNum(sp.stowHCByFloor.A03[period]) || '--') + ' A04:' + (formatNum(sp.stowHCByFloor.A04[period]) || '--') + ' A05:' + (formatNum(sp.stowHCByFloor.A05[period]) || '--') + ']\n';

        text += 'Decant HC: ' + (formatNum(sp.decantHC[period]) || '--') + '\n';

        text += 'Indirect Spend: ' + (sp.indirectSpendPercent[period] !== null ? sp.indirectSpendPercent[period] + '%' : '--') + '\n';

        text += 'Remaining CUT Trailers: ' + (formatNum(ibf.remainingCUTs) || '--') + ' | CUT Compliance: ' + (ibf.cutCompliance !== null ? ibf.cutCompliance + '%' : '--') + '\n\n';



        // COST

        text += 'COST\n' + '----------------------------------------\n';

        text += 'IB Total TPH: ' + (formatNum(fclmData.ibTotalRate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.ibTotalPlanRate) || '--') + ')\n';

        text += 'Decant Rate: ' + (formatNum(fclmData.decantRate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.decantPlanRate) || '--') + ')\n';

        text += 'TI Support: ' + (formatNum(fclmData.transferInSupportRate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.transferInSupportPlanRate) || '--') + ')\n';

        text += 'STP Support: ' + (formatNum(fclmData.stowToPrimeSupportRate || fclmData.stowToprimeRate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.stowToPrimeSupportPlanRate || fclmData.stowToPrimePlanRate) || '--') + ')\n';

        text += 'IB Lead/PA: ' + (formatNum(fclmData.ibLeadPARate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.ibLeadPAPlanRate) || '--') + ')\n';

        text += 'IB Problem Solve: ' + (formatNum(fclmData.ibProblemSolveRate) || '--') + ' (LP Rate: ' + (formatNum(fclmData.ibProblemSolvePlanRate) || '--') + ')\n\n';



        // INPUT METRICS

        text += 'INPUT METRICS (Vantage)\n' + '----------------------------------------\n';

        text += 'Stow CT: ' + (vantageData.stowCycleTime || '--') + 's (goal \u2264 13s)\n';

        text += 'UPF: ' + (vantageData.upf || '--') + ' (goal > 12)\n';

        text += 'OOWA%: ' + (vantageData.oowa || '--') + '% (goal < 2%)\n';

        text += 'NSTA%: ' + (vantageData.nsta || '--') + '% (goal < 15%)\n';

        text += 'Per-Zone: ';

        Object.keys(vantageData.zones || {}).forEach(function(zone) {

            var zd = vantageData.zones[zone];

            text += zone.replace(SITE_SETTINGS.ZONE_PREFIX, '') + '[CT:' + (zd.stowCycleTime || '--') + ' UPF:' + (zd.upf || '--') + ' NSTA:' + (zd.nsta || '--') + '%] ';

        });

        text += '\n\n';



        // AUDITS

        text += 'STANDARD WORK AUDITS (Apollo)\n' + '----------------------------------------\n';

        apolloData.forEach(function(a) { text += a.name + ': ' + a.count + ' (goal: ' + (getAuditGoal(a.name) || 'N/A') + ')\n'; });

        text += '\n';



        // QUALITY / SAFETY

        text += 'QUALITY: (manual entry)\n';

        text += 'PS Piles: ' + (formatNum(ibf.psPiles) || '--') + '\n\n';

        text += 'SAFETY / PEOPLE / EQUIPMENT / TRAINING: (manual entry)\n';



        if (typeof GM_setClipboard !== 'undefined') {

            GM_setClipboard(text, 'text');

        } else {

            navigator.clipboard.writeText(text).catch(function() {

                var ta = document.createElement('textarea');

                ta.value = text;

                document.body.appendChild(ta);

                ta.select();

                document.execCommand('copy');

                ta.remove();

            });

        }

    }



    // ==========================================

    // PERIOD SNAPSHOT PERSISTENCE

    // ==========================================

    var SNAPSHOT_KEY = 'ppr_sync_snapshots_' + WAREHOUSE;



    function getShiftDate() {

        var dateStr = getSelectedDate();

        var shift = getCurrentShift().toLowerCase();

        return dateStr + '_' + shift;

    }



    function savePeriodSnapshot(period, fclmData, vantageData, apolloData, introData, slimData) {

        var shiftDate = getShiftDate();

        var snapshots = {};



        try {

            var stored = localStorage.getItem(SNAPSHOT_KEY);

            if (stored) {

                snapshots = JSON.parse(stored);

                if (snapshots._shiftDate !== shiftDate) {

                    snapshots = {};

                }

            }

        } catch (e) { snapshots = {}; }



        snapshots._shiftDate = shiftDate;

        snapshots[period] = {

            timestamp: Date.now(),

            fclm: {

                ibTotalRate: fclmData.ibTotalRate,

                ibTotalTPH: fclmData.ibTotalTPH || fclmData.ibTotalRate,

                ibTotalVol: fclmData.ibTotalVol,

                decantRate: fclmData.decantRate,

                stowToprimeRate: fclmData.stowToprimeRate,

                transferInSupportRate: fclmData.transferInSupportRate,

                transferInRate: fclmData.transferInRate,

                ibLeadPARate: fclmData.ibLeadPARate,

                ibLeadPAHrs: fclmData.ibLeadPAHrs,

                ibProblemSolveRate: fclmData.ibProblemSolveRate,

                ibProblemSolveHrs: fclmData.ibProblemSolveHrs,

                stowToPrimeSupportRate: fclmData.stowToPrimeSupportRate,

            },

            vantage: {

                stowCycleTime: vantageData.stowCycleTime,

                upf: vantageData.upf,

                nsta: vantageData.nsta,

                oowa: vantageData.oowa,

                stowRate: vantageData.stowRate,

                zones: vantageData.zones,

            },

            apollo: apolloData.map(function(a) { return { name: a.name, count: a.count }; }),

            slim: {

                receiveAutoDecant: slimData ? slimData.receiveAutoDecant : null,

                receiveManualDecant: slimData ? slimData.receiveManualDecant : null,

            },

        };



        localStorage.setItem(SNAPSHOT_KEY, JSON.stringify(snapshots));

        console.log('[IB Sync] Saved snapshot for ' + period);

    }



    // ==========================================

    // INITIALIZE

    // ==========================================

    function initSyncDashboard() {

        console.log('[IB Sync v3.1] Initializing dashboard...');



        var attempts = 0;

        var waitForTable = setInterval(function() {

            attempts++;

            var rows = document.querySelectorAll('tr');

            var hasData = false;

            rows.forEach(function(r) {

                if (r.textContent.indexOf('Inbound-TOTAL') !== -1) {

                    hasData = true;

                }

            });



            if (hasData && attempts >= 8) {

                // Wait at least 8 seconds after detecting table before scraping

                // FCLM re-renders data multiple times on load; early scrape gets stale values

                clearInterval(waitForTable);

                buildDashboard();

            }

        }, 1000);



        setTimeout(function() { clearInterval(waitForTable); buildDashboard(); }, 25000);



        // Sync with FCLM "HTML" button — when clicked, auto-refresh dashboard after table reloads

        var htmlBtn = document.querySelector('input[value="HTML"], button[value="HTML"]');

        if (!htmlBtn) {

            // Try finding by text content

            document.querySelectorAll('input[type="submit"], button').forEach(function(btn) {

                if (btn.value === 'HTML' || btn.textContent.trim() === 'HTML') htmlBtn = btn;

            });

        }

        if (htmlBtn) {

            htmlBtn.addEventListener('click', function() {

                console.log('[IB Sync] FCLM HTML button clicked — will rebuild dashboard after table loads');

                setTimeout(function() { buildDashboard(); }, 10000); // Wait 10s for FCLM to reload

            });

        }

    }



    // ==========================================

    // ALPs S1 CAPACITY FORECAST (Auto-fetch)

    // ==========================================

    // Fetches S1 Capacity from ALPs API using the Weekly-tagged plan.

    // Uses existing Midway session. Portable across all sites.

    function fetchALPsS1() {
        return new Promise(function(resolve) {
            var alpsBase = 'https://alps-iad.iad.proxy.amazon.com/api';
            var FRIDAY_PLAN_FALLBACK = '8409bf2c-422b-420b-9086-1b85a820c31e';

            var tagUrl = alpsBase + '/site/' + WAREHOUSE + '/latest-completed-plan-by-tag?tagName=Weekly&siteType=FULFILLMENT_CENTER&polling=false';
            console.log('[IB Sync] ALPs S1: Fetching Weekly plan for', WAREHOUSE);

            GM_xmlhttpRequest({
                method: 'GET', url: tagUrl, headers: { 'Accept': 'application/json' },
                onload: function(resp) {
                    if (resp.status !== 200) { fetchCapacityData(FRIDAY_PLAN_FALLBACK); return; }
                    try {
                        var tagData = JSON.parse(resp.responseText);
                        var planId = tagData.planId || tagData.id || null;
                        if (!planId) {
                            var keys = Object.keys(tagData);
                            for (var k = 0; k < keys.length; k++) {
                                if (typeof tagData[keys[k]] === 'string' && tagData[keys[k]].match(/^[0-9a-f]{8}-/)) { planId = tagData[keys[k]]; break; }
                            }
                        }
                        if (!planId) { fetchCapacityData(FRIDAY_PLAN_FALLBACK); return; }
                        console.log('[IB Sync] ALPs S1: Tag returned plan:', planId);
                        validatePlan(planId);
                    } catch(e) { fetchCapacityData(FRIDAY_PLAN_FALLBACK); }
                },
                onerror: function() { fetchCapacityData(FRIDAY_PLAN_FALLBACK); }
            });

            function validatePlan(planId) {
                var metaUrl = alpsBase + '/site/' + WAREHOUSE + '/plan/' + planId + '/metadata';
                GM_xmlhttpRequest({
                    method: 'GET', url: metaUrl, headers: { 'Accept': 'application/json' },
                    onload: function(metaResp) {
                        if (metaResp.status !== 200) { fetchCapacityData(FRIDAY_PLAN_FALLBACK); return; }
                        try {
                            var meta = JSON.parse(metaResp.responseText);
                            var planMeta = meta.planMetaData || meta;
                            var createdAt = planMeta.createdAt || planMeta.planDate || '';
                            var createdDate = new Date(createdAt);
                            var dayOfWeek = createdDate.getDay();
                            var dayName = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dayOfWeek];
                            console.log('[IB Sync] ALPs S1: Plan created', dayName, createdAt);
                            if (dayOfWeek === 5) {
                                console.log('[IB Sync] ALPs S1: VALIDATED Friday plan');
                                fetchCapacityData(planId);
                            } else {
                                console.log('[IB Sync] ALPs S1: REJECTED (not Friday). Using fallback.');
                                fetchCapacityData(FRIDAY_PLAN_FALLBACK);
                            }
                        } catch(e) { fetchCapacityData(FRIDAY_PLAN_FALLBACK); }
                    },
                    onerror: function() { fetchCapacityData(FRIDAY_PLAN_FALLBACK); }
                });
            }

            function fetchCapacityData(planId) {
                var fclmUrlParams = new URLSearchParams(window.location.search);
                var fclmSpanType = fclmUrlParams.get('spanType') || '';
                var fclmDateStr = '';
                if (fclmSpanType === 'Intraday') fclmDateStr = (fclmUrlParams.get('startDateIntraday') || '').split(/[\s+T]/)[0];
                else if (fclmSpanType === 'Week') fclmDateStr = fclmUrlParams.get('startDateWeek') || '';
                else fclmDateStr = fclmUrlParams.get('startDateDay') || '';
                var centerDate = fclmDateStr ? new Date(fclmDateStr.replace(/\//g, '-')) : new Date();
                if (isNaN(centerDate.getTime())) centerDate = new Date();
                var sd = new Date(centerDate); sd.setDate(centerDate.getDate() - 7);
                var ed = new Date(centerDate); ed.setDate(centerDate.getDate() + 7);
                var startStr = sd.getFullYear() + '-' + String(sd.getMonth()+1).padStart(2,'0') + '-' + String(sd.getDate()).padStart(2,'0');
                var endStr = ed.getFullYear() + '-' + String(ed.getMonth()+1).padStart(2,'0') + '-' + String(ed.getDate()).padStart(2,'0');

                var dataUrl = alpsBase + '/report/FULFILLMENT_CENTER/' + WAREHOUSE + '/getPlanSelectionData'
                    + '?view=dailyView&selection=inbound-joint&planId=' + planId
                    + '&withUserOverrides=true&startDate=' + startStr + '&endDate=' + endStr + '&withComputedValues=true';
                console.log('[IB Sync] ALPs S1: Fetching capacity, plan:', planId, 'dates:', startStr, 'to', endStr);

                GM_xmlhttpRequest({
                    method: 'GET', url: dataUrl, headers: { 'Accept': 'application/json' },
                    onload: function(resp) {
                        if (resp.status !== 200) { console.log('[IB Sync] ALPs S1: Data fetch failed:', resp.status); resolve(null); return; }
                        try {
                            var planData = JSON.parse(resp.responseText);
                            var s1 = parseALPsCapacity(planData);
                            console.log('[IB Sync] ALPs S1: Parsed =', JSON.stringify(s1));
                            resolve(s1);
                        } catch(e) { console.log('[IB Sync] ALPs S1: Parse error:', e.message); resolve(null); }
                    },
                    onerror: function() { resolve(null); }
                });
            }
        });
    }



    function parseALPsCapacity(planData) {

        // Parse ALPs getPlanSelectionData response (with withComputedValues=true)

        // Structure: response[0].Metric[0]=Day, [1]=Night, [2]=Local

        // Path to capacity: Metric[n].subRows[0](Volume).subRows[0](Capacity).subRows[1](Forecast).date.value

        // Use the FCLM page's start date (the operational date being viewed)

        var fclmParams = new URLSearchParams(window.location.search);

        var fclmSpan = fclmParams.get('spanType') || '';

        var fclmDateRaw = '';

        if (fclmSpan === 'Intraday') {

            fclmDateRaw = (fclmParams.get('startDateIntraday') || '').split(/[\s+T]/)[0];

        } else if (fclmSpan === 'Week') {

            fclmDateRaw = fclmParams.get('startDateWeek') || '';

        } else {

            fclmDateRaw = fclmParams.get('startDateDay') || '';

        }

        var todayStr;

        if (fclmDateRaw) {

            // FCLM format: 2026/08/25 → convert to 2026-08-25

            todayStr = fclmDateRaw.replace(/\//g, '-').substring(0, 10);

            console.log('[IB Sync] ALPs S1: Using FCLM page date (' + fclmSpan + '):', todayStr);

        } else {

            // Fallback to current date if URL param not available

            var today = new Date();

            todayStr = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');

            console.log('[IB Sync] ALPs S1: No FCLM date param, using today:', todayStr);

        }

        var result = { dayCapacity: null, nightCapacity: null, localCapacity: null, s1Goal: null };



        try {

            console.log('[IB Sync] ALPs S1: Looking for date:', todayStr);



            var metrics = null;

            // Handle both array and object response formats

            if (Array.isArray(planData) && planData[0] && planData[0].Metric) {

                metrics = planData[0].Metric;

            } else if (planData && planData['0'] && planData['0'].Metric) {

                metrics = planData['0'].Metric;

            }



            if (!metrics) {

                console.log('[IB Sync] ALPs S1: Could not find Metric array. Keys:', Object.keys(planData || {}));

                return result;

            }



            // Convert metrics to array if it's an object with numeric keys

            var metricArray = Array.isArray(metrics) ? metrics : Object.keys(metrics).map(function(k) { return metrics[k]; });



            metricArray.forEach(function(shift) {

                if (!shift || !shift.header) return;

                var shiftName = shift.header;



                try {

                    // Navigate: subRows[0] (Volume) > subRows[0] (Capacity) > subRows[1] (Forecast) > date

                    var volume = shift.subRows && shift.subRows[0];

                    if (!volume || !volume.subRows) return;

                    var capacity = volume.subRows[0]; // Capacity section

                    if (!capacity || !capacity.subRows) return;



                    // Find the Forecast row (usually index 1, but search by name to be safe)

                    var forecast = null;

                    for (var i = 0; i < capacity.subRows.length; i++) {

                        if (capacity.subRows[i].header === 'Forecast') {

                            forecast = capacity.subRows[i];

                            break;

                        }

                    }

                    if (!forecast) {

                        // Fallback: try index 1

                        forecast = capacity.subRows[1];

                    }

                    if (!forecast || !forecast[todayStr]) return;



                    var val = forecast[todayStr].value;

                    if (val != null && val !== 0) {

                        val = Math.round(val);

                        if (shiftName === 'Day') result.dayCapacity = val;

                        else if (shiftName === 'Night') result.nightCapacity = val;

                        else if (shiftName === 'Local Time-zone') result.localCapacity = val;

                    }

                } catch(e) {

                    console.log('[IB Sync] ALPs S1: Error parsing shift', shiftName, e.message);

                }

            });



            // Calculate S1 Goal

            if (result.localCapacity) {

                result.s1Goal = result.localCapacity;

            } else if (result.dayCapacity && result.nightCapacity) {

                result.s1Goal = result.dayCapacity + result.nightCapacity;

            }



            if (result.s1Goal) {

                console.log('[IB Sync] ALPs S1: SUCCESS - Day:', result.dayCapacity, 'Night:', result.nightCapacity, 'S1:', result.s1Goal);

            } else {

                console.log('[IB Sync] ALPs S1: No capacity found for', todayStr, '- Day:', result.dayCapacity, 'Night:', result.nightCapacity);

            }

        } catch(e) {

            console.log('[IB Sync] ALPs S1: parseALPsCapacity error:', e.message);

        }

        return result;

    }





    function findTodayValue(valuesArray, datesArray, todayStr) {

        if (!valuesArray || !datesArray) return null;

        var idx = datesArray.indexOf(todayStr);

        if (idx >= 0 && idx < valuesArray.length) return valuesArray[idx];

        return null;

    }







    function buildDashboard() {

        console.log('[IB Sync v3.1] Table detected, fetching all data...');

        var fclmData = scrapeFCLMTable();

        console.log('[IB Sync] FCLM data:', fclmData);

            console.log('[IB Sync] FCLM LP Rates - IB:', fclmData.ibTotalPlanRate, 'Decant:', fclmData.decantPlanRate, 'STP:', fclmData.stowToPrimePlanRate, 'TIS:', fclmData.transferInSupportPlanRate, 'Lead/PA:', fclmData.ibLeadPAPlanRate, 'PS:', fclmData.ibProblemSolvePlanRate);



        var shiftBounds = getShiftBounds();

        var startISO = shiftBounds.start.toISOString();

        var endISO = new Date().toISOString();



        // Fetch ALL data in parallel: Apollo + Vantage + OOWA + INTRO APIs + SLIM

        Promise.allSettled([

            fetchAllApolloData(),

            fetchVantageMetrics(),

            fetchVantageOOWA(startISO, endISO),

            fetchINTROApis(),

            fetchSLIMData(),

            fetchAtlasData(),

            fetchRoboScoutData(),

            fetchFCLMByPeriod()

        ]).then(function(results) {

            var apolloData = results[0].status === 'fulfilled' ? results[0].value : [];

            var vantageData = (results[1].status === 'fulfilled' && results[1].value) ? results[1].value : { stowCycleTime: null, upf: null, nsta: null, oowa: null, stowRate: null, etiTotal: null, zones: {}, byFloor: {} };

            var oowaValue = results[2].status === 'fulfilled' ? results[2].value : null;

            var introRaw = results[3].status === 'fulfilled' ? results[3].value : { ibFlow: null, trailerPlanner: null, shiftPlanner: null };

            var slimData = results[4].status === 'fulfilled' ? results[4].value : { receiveAutoDecant: null, receiveManualDecant: null };

            var atlasData = results[5].status === 'fulfilled' ? results[5].value : {};

            var roboData = results[6].status === 'fulfilled' ? results[6].value : { stowCycleTime: null, upf: null, nsta: null, oowaCount: null, oowaDwell: null, byFloor: {} };

            var fclmByPeriod = results[7].status === 'fulfilled' ? results[7].value : { etiRate: {}, decantRate: {}, transferInSupportRate: {}, stowToPrimeRate: {}, ibLeadPARate: {}, ibProblemSolveRate: {} };

            console.log('[IB Sync] FCLM by period:', fclmByPeriod);



            console.log('[IB Sync] Apollo:', apolloData.length, 'audits');

            console.log('[IB Sync] Vantage:', vantageData);

            console.log('[IB Sync] OOWA:', oowaValue);

            console.log('[IB Sync] INTRO raw:', introRaw);

            console.log('[IB Sync] SLIM:', slimData);

            console.log('[IB Sync] ATLAS:', atlasData);

            console.log('[IB Sync] RoboScout:', roboData);



            // Track data source status for header indicator

            var sourceStatus = {

                intro: !!(introRaw.shiftPlanner || introRaw.trailerPlanner),

                apollo: apolloData.length > 0 && apolloData[0].eos !== null,

                atlas: Object.keys(atlasData).length > 0,

                slim: !!(slimData.receiveAutoDecant !== null || slimData.receiveManualDecant !== null),

                fclm: !!(fclmData.ibTotalRate || fclmData.decantRate),

                roboscout: !!(roboData.stowCycleTime || roboData.upf || roboData.nsta),

            };



            // Merge RoboScout data into vantage (RoboScout replaces Vantage as source for CT/UPF/NSTA/OOWA)

            if (roboData) {

                if (roboData.stowCycleTime) vantageData.stowCycleTime = roboData.stowCycleTime.toFixed(1);

                if (roboData.upf) vantageData.upf = roboData.upf.toFixed(2);

                if (roboData.nsta) vantageData.nsta = roboData.nsta.toFixed(1);

                // OOWA% = (andon_count * avg_dwell_minutes / 60) / (ETI_hrs + STP_hrs) * 100

                // Site-level: sum OOW hours from all floors (floor-level has Andon Count + Dwell)

                var totalOowHours = 0;

                var hasFloorOowa = false;

                if (roboData.byFloor) {

                    SITE_SETTINGS.FLOORS.forEach(function(f) {

                        var fd = roboData.byFloor[f];

                        if (fd && fd.oowaCount !== null && fd.oowaDwell !== null) {

                            totalOowHours += (fd.oowaCount * fd.oowaDwell) / 60;

                            hasFloorOowa = true;

                        }

                    });

                }

                // Fallback to site-level if floor data not available

                if (!hasFloorOowa && roboData.oowaCount !== null && roboData.oowaDwell !== null) {

                    totalOowHours = (roboData.oowaCount * roboData.oowaDwell) / 60;

                    hasFloorOowa = true;

                }

                if (hasFloorOowa) {

                    var etiHrs = fclmData.transferInHrs || 0;

                    var stpHrs = fclmData.stowToPrimeHrs || 0;

                    var productiveHrs = etiHrs + stpHrs;

                    var oowaPct = productiveHrs > 0 ? (totalOowHours / productiveHrs) * 100 : null;

                    if (oowaPct !== null) vantageData.oowa = oowaPct.toFixed(2);

                    console.log('[IB Sync] OOWA% calc: totalOowHrs=' + totalOowHours.toFixed(3) + ' etiHrs=' + etiHrs + ' stpHrs=' + stpHrs + ' productiveHrs=' + productiveHrs + ' => ' + (oowaPct !== null ? oowaPct.toFixed(2) : 'null') + '%');

                }

                // Merge per-period data into vantageData for INPUT METRICS P1/P2/P3

                if (roboData.byPeriod) {

                    if (!vantageData.byPeriod) vantageData.byPeriod = { P1: {}, P2: {}, P3: {} };

                    ['P1', 'P2', 'P3'].forEach(function(p) {

                        var pd = roboData.byPeriod[p];

                        if (pd) {

                            if (pd.stowCycleTime) vantageData.byPeriod[p].stowCycleTime = pd.stowCycleTime.toFixed(1);

                            if (pd.upf) vantageData.byPeriod[p].upf = pd.upf.toFixed(2);

                            if (pd.nsta) vantageData.byPeriod[p].nsta = pd.nsta.toFixed(1);

                            // OOWA%: (count * dwell_minutes / 60) / (etiHrs + stpHrs) * 100

                            if (pd.oowaCount !== null && pd.oowaDwell !== null && fclmByPeriod && fclmByPeriod.etiHrs) {

                                var pEtiHrs = fclmByPeriod.etiHrs[p] || 0;

                                var pStpHrs = fclmByPeriod.stpHrs ? (fclmByPeriod.stpHrs[p] || 0) : 0;

                                var pProductiveHrs = pEtiHrs + pStpHrs;

                                if (pProductiveHrs > 0) {

                                    var pOowaPct = ((pd.oowaCount * pd.oowaDwell) / 60) / pProductiveHrs * 100;

                                    vantageData.byPeriod[p].oowa = pOowaPct.toFixed(2);

                                }

                            }

                        }

                    });

                }

                // Merge per-floor data

                if (roboData.byFloor) {

                    SITE_SETTINGS.FLOORS.forEach(function(floor) {

                        var floorData = roboData.byFloor[floor];

                        if (floorData) {

                            var zoneName = SITE_SETTINGS.ZONE_PREFIX + floor;

                            if (!vantageData.zones[zoneName]) vantageData.zones[zoneName] = {};

                            if (!vantageData.byFloor[zoneName]) vantageData.byFloor[zoneName] = {};

                            if (floorData.upf) { vantageData.zones[zoneName].upf = floorData.upf.toFixed(2); vantageData.byFloor[zoneName].upf = floorData.upf.toFixed(2); }

                            if (floorData.stowCycleTime) { vantageData.zones[zoneName].stowCycleTime = floorData.stowCycleTime.toFixed(1); vantageData.byFloor[zoneName].stowCycleTime = floorData.stowCycleTime.toFixed(1); }

                            if (floorData.nsta) { vantageData.zones[zoneName].nsta = floorData.nsta.toFixed(1); vantageData.byFloor[zoneName].nsta = floorData.nsta.toFixed(1); }

                            if (floorData.stowRate) { vantageData.zones[zoneName].stowRate = floorData.stowRate.toFixed(2); vantageData.byFloor[zoneName].stowRate = floorData.stowRate.toFixed(2); }

                            if (floorData.etiTotal) { vantageData.zones[zoneName].etiTotal = Math.round(floorData.etiTotal); vantageData.byFloor[zoneName].etiTotal = Math.round(floorData.etiTotal); }

                            // OOWA% per floor = (count * dwell_min / 60) / (productive_hrs / 4) * 100

                            if (floorData.oowaCount !== null && floorData.oowaDwell !== null) {

                                var floorOowHrs = (floorData.oowaCount * floorData.oowaDwell) / 60;

                                var floorProductiveHrs = ((fclmData.transferInHrs || 0) + (fclmData.stowToPrimeHrs || 0)) / 4;

                                var floorOowaPct = floorProductiveHrs > 0 ? (floorOowHrs / floorProductiveHrs) * 100 : null;

                                if (floorOowaPct !== null) {

                                    vantageData.zones[zoneName].oowa = floorOowaPct.toFixed(2);

                                    vantageData.byFloor[zoneName].oowa = floorOowaPct.toFixed(2);

                                }

                            }

                        }

                    });

                }

            }



            // Parse INTRO API responses

            var introData = {

                ibFlow: parseIBFlowData(introRaw.ibFlow, introRaw.stowWipData),

                trailerPlanner: parseTrailerPlannerData(introRaw.trailerPlanner, introRaw.trailerAllocationPlan, introRaw.targetVolume, introRaw.targetVolumeOpposite),

                shiftPlanner: parseShiftPlannerData(introRaw.shiftPlanner),

                volumeActuals: parseShiftPerformance(introRaw.shiftPerformance),

            };





            // ALPs S1 override: use ALPs capacity if available, otherwise blank/editable

            var alpsS1Data = introRaw.alpsS1 || null;

            if (alpsS1Data && alpsS1Data.s1Goal) {

                console.log('[IB Sync] ALPs S1: using capacity', alpsS1Data.s1Goal);

                introData.trailerPlanner.s1Goal = alpsS1Data.s1Goal;

                introData.trailerPlanner.alpsS1 = alpsS1Data;

            } else {

                console.log('[IB Sync] ALPs S1: call failed or no data - leaving S1 blank for manual entry');

                introData.trailerPlanner.s1Goal = null;

            }



            console.log('[IB Sync] Parsed INTRO:', introData);



            // Save period snapshot

            var period = getCurrentPeriod();

            savePeriodSnapshot(period, fclmData, vantageData, apolloData, introData, slimData);



            // Build panel

            var panel = createSyncPanel(fclmData, apolloData, vantageData, introData, slimData, atlasData, sourceStatus, fclmByPeriod);



            // Neutralize FCLM sticky headers that overlap the dashboard

            var stickyFix = document.createElement('style');

            stickyFix.id = 'ib-sync-sticky-fix';

            stickyFix.textContent = 

                'table.reportLayout th, table[class*="report"] th, ' +

                '.rt-thead, .rt-th, ' +

                'table th[style*="sticky"], table th[style*="fixed"] { ' +

                '  position: static !important; ' +

                '} ' +

                '#ib-sync-dashboard, #ib-sync-dashboard * { ' +

                '  position: static; ' +

                '} ' +

                '#ib-sync-dashboard { ' +

                '  position: relative !important; ' +

                '  z-index: 2147483647 !important; ' +

                '  isolation: isolate !important; ' +

                '}';

            if (!document.getElementById('ib-sync-sticky-fix')) {

                document.head.appendChild(stickyFix);

            }



            // Insert into page

            var controlPanel = document.getElementById('control-panel');

            if (controlPanel && controlPanel.parentNode) {

                var wrapper = document.createElement('div');

                wrapper.style.cssText = 'width:100%;display:flex;justify-content:center;';

                wrapper.appendChild(panel);

                controlPanel.parentNode.insertBefore(wrapper, controlPanel.nextSibling);

            } else {

                document.body.insertBefore(panel, document.body.firstChild);

            }



            console.log('[IB Sync v3.1] Dashboard rendered successfully.');

        });

    }



    // Start

    if (document.readyState === 'loading') {

        document.addEventListener('DOMContentLoaded', initSyncDashboard);

    } else {

        initSyncDashboard();

    }



})();

