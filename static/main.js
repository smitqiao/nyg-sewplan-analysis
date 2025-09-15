// project_nytg/sewing_plan_analysis/static/main.js
        function formatNumberThousandSep(val) {
            if (val === '' || val === null || val === undefined || isNaN(val)) return '';
            return Number(val).toLocaleString();
        }

        let allData = [];
        let factories = new Set();
        let styleRefs = new Set();
        let styleColorMap = {}; // styleRef -> color
        let minDate = null, maxDate = null;
        let useSamIeCalc = false;
        let groupBy = 'day'; // <-- Add this line
        let currentView = 'table'; // 'table' or 'graph'
        let planPcsChart = null;
        let expandLineStyle = false; // Track expand/collapse state
        // Add: subprocess state
        let showSubprocess = false;
        let selectedSubprocesses = [];
        let usePoints = false;
        let selectedLines = []; // <-- Add this global variable
        let availableLines = []; // <-- add global variable
        let selectedDelayStatuses = ['Delay', 'Can be delay', 'Too early', 'Ok']; // <-- Add this line

        // Helper: get ISO week string (YYYY-Www)
        function getISOWeekString(dateObj) {
            const d = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
            const dayNum = d.getUTCDay() || 7;
            d.setUTCDate(d.getUTCDate() + 4 - dayNum);
            const yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
            const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1)/7);
            return `${d.getUTCFullYear()}-W${String(weekNum).padStart(2,'0')}`;
        }
        // Helper: get Month string (YYYY-MM)
        function getMonthString(dateObj) {
            return `${dateObj.getFullYear()}-${String(dateObj.getMonth()+1).padStart(2,'0')}`;
        }

        function updateDateInputs() {
            // Find min/max date from allData
            const allDates = allData
                .map(row => row.SEW_DATE || row.sew_date)
                .filter(Boolean)
                .map(d => {
                    // Always parse and format as yyyy-MM-dd
                    const dateObj = new Date(d);
                    if (!isNaN(dateObj)) {
                        return dateObj.toISOString().split('T')[0];
                    }
                    return '';
                })
                .filter(Boolean);
            if (allDates.length === 0) return;
            minDate = allDates.reduce((a, b) => a < b ? a : b);
            maxDate = allDates.reduce((a, b) => a > b ? a : b);
            const dateStart = document.getElementById('dateStart');
            const dateEnd = document.getElementById('dateEnd');
            dateStart.min = minDate;
            dateStart.max = maxDate;
            dateEnd.min = minDate;
            dateEnd.max = maxDate;
            dateStart.value = minDate;
            dateEnd.value = maxDate;
        }

        // Helper: generate highly distinct color for large sets
        function getDistinctColor(idx, total) {
            // Use golden angle to spread hues, and cycle saturation/lightness for extra distinction
            const goldenAngle = 137.508;
            const hue = Math.floor((idx * goldenAngle) % 360);
            // Cycle saturation and lightness for extra distinction
            const sat = 60 + (idx % 3) * 15; // 60, 75, 90
            const light = 65 + (idx % 4) * 7; // 65, 72, 79, 86
            return `hsl(${hue}, ${sat}%, ${light}%)`;
        }

        // Assign highly distinct colors to each styleRef
        function assignStyleColors() {
            styleColorMap = {};
            const styles = Array.from(new Set(allData.map(row => row.STYLE_REF || row.style_ref).filter(Boolean)));
            styles.forEach((style, idx) => {
                // If more than 360 styles, fallback to random color for overflow
                if (styles.length > 360) {
                    styleColorMap[style] = `hsl(${Math.floor(Math.random()*360)}, ${60 + Math.floor(Math.random()*30)}%, ${65 + Math.floor(Math.random()*20)}%)`;
                } else {
                    styleColorMap[style] = getDistinctColor(idx, styles.length);
                }
            });
        }

        function renderDashboard(filteredData) {
            // --- Pre-aggregate sums for each (factory, line, groupKey) pair if multi-factory or all factories, else line ---
            let linesSet = new Set();
            let stylesSet = new Set();
            let groupKeySet = new Set();
            let aggMap = {}; // { rowKey: { groupKey: { pcsSum, samIeSum, rows: [], groupLabel } } }
            let useFactoryLine = false;
            // Determine if multiple factories are selected or all factories (no filter)
            const selectedFactories = getSelectedValues('factorySelect');
            useFactoryLine = selectedFactories.length !== 1; // group by factory-line if 0 or >1 selected

            if (expandLineStyle) {
                // Expanded: group by line+style
                filteredData.forEach(row => {
                    const line = row.PROD_LINE || row.prod_line || 'Unknown';
                    const style = row.STYLE_REF || row.style_ref || 'Unknown';
                    const dateRaw = row.SEW_DATE || row.sew_date || '';
                    if (!dateRaw) return;
                    const dateObj = new Date(dateRaw);
                    if (isNaN(dateObj)) return;
                    let groupKey, groupLabel;
                    if (groupBy === 'day') {
                        groupKey = dateObj.toISOString().split('T')[0];
                        groupLabel = groupKey;
                    } else if (groupBy === 'week') {
                        groupKey = getISOWeekString(dateObj);
                        groupLabel = groupKey;
                    } else if (groupBy === 'month') {
                        groupKey = getMonthString(dateObj);
                        groupLabel = groupLabel;
                    }
                    linesSet.add(line);
                    stylesSet.add(style);
                    groupKeySet.add(groupKey);

                    if (!aggMap[line]) aggMap[line] = {};
                    if (!aggMap[line][style]) aggMap[line][style] = {};
                    if (!aggMap[line][style][groupKey]) aggMap[line][style][groupKey] = { pcsSum: 0, samIeSum: 0, rows: [], groupLabel };
                    const pcs = Number(row.PLAN_PCS || row.plan_pcs) || 0;
                    const sam = Number(row.SAM_IE || row.sam_ie) || 0;
                    aggMap[line][style][groupKey].pcsSum += pcs;
                    aggMap[line][style][groupKey].samIeSum += pcs * sam;
                    aggMap[line][style][groupKey].rows.push(row);
                });
            } else {
                // Collapsed: group by (factory-line) if multi-factory, else by line
                filteredData.forEach(row => {
                    const factory = row.FACTORY || row.factory || 'Unknown';
                    const line = row.PROD_LINE || row.prod_line || 'Unknown';
                    const rowKey = useFactoryLine ? `${factory} - ${line}` : line;
                    const dateRaw = row.SEW_DATE || row.sew_date || '';
                    if (!dateRaw) return;
                    const dateObj = new Date(dateRaw);
                    if (isNaN(dateObj)) return;
                    let groupKey, groupLabel;
                    if (groupBy === 'day') {
                        groupKey = dateObj.toISOString().split('T')[0];
                        groupLabel = groupKey;
                    } else if (groupBy === 'week') {
                        groupKey = getISOWeekString(dateObj);
                        groupLabel = groupKey;
                    } else if (groupBy === 'month') {
                        groupKey = getMonthString(dateObj);
                        groupLabel = groupLabel;
                    }
                    linesSet.add(rowKey);
                    groupKeySet.add(groupKey);

                    if (!aggMap[rowKey]) aggMap[rowKey] = {};
                    if (!aggMap[rowKey][groupKey]) aggMap[rowKey][groupKey] = { pcsSum: 0, samIeSum: 0, rows: [], groupLabel };
                    const pcs = Number(row.PLAN_PCS || row.plan_pcs) || 0;
                    const sam = Number(row.SAM_IE || row.sam_ie) || 0;
                    aggMap[rowKey][groupKey].pcsSum += pcs;
                    aggMap[rowKey][groupKey].samIeSum += pcs * sam;
                    aggMap[rowKey][groupKey].rows.push(row);
                });
            }

            const lines = Array.from(linesSet).sort();
            // Sort group keys
            let groupKeys = Array.from(groupKeySet);
            if (groupBy === 'day') {
                groupKeys.sort((a, b) => new Date(a) - new Date(b));
            } else if (groupBy === 'week' || groupBy === 'month') {
                groupKeys.sort();
            }
            const displayLines = lines;
            const displayGroups = groupKeys;
            window._pivotTableDisplayDates = displayGroups; // Expose for click handler

            // --- Build header grouping for week/month ---
            let headerGroups = [];
            if (groupBy === 'day') {
                // Group by month for header
                let lastMonth = '', count = 0;
                displayGroups.forEach((date, i) => {
                    const d = new Date(date);
                    const month = d.toLocaleString('en-US', { month: 'short' });
                    if (month !== lastMonth) {
                        if (count > 0) headerGroups.push({ label: lastMonth, count });
                        lastMonth = month;
                        count = 1;
                    } else {
                        count++;
                    }
                    if (i === displayGroups.length - 1) {
                        headerGroups.push({ label: lastMonth, count });
                    }
                });
            } else if (groupBy === 'week') {
                // Group by year for header
                let lastYear = '', count = 0;
                displayGroups.forEach((wk, i) => {
                    const year = wk.split('-W')[0];
                    if (year !== lastYear) {
                        if (count > 0) headerGroups.push({ label: lastYear, count });
                        lastYear = year;
                        count = 1;
                    } else {
                        count++;
                    }
                    if (i === displayGroups.length - 1) {
                        headerGroups.push({ label: lastYear, count });
                    }
                });
            } else if (groupBy === 'month') {
                // Group by year for header
                let lastYear = '', count = 0;
                displayGroups.forEach((m, i) => {
                    const year = m.split('-')[0];
                    if (year !== lastYear) {
                        if (count > 0) headerGroups.push({ label: lastYear, count });
                        lastYear = year;
                        count = 1;
                    } else {
                        count++;
                    }
                    if (i === displayGroups.length - 1) {
                        headerGroups.push({ label: lastYear, count });
                    }
                });
            }

            // --- Build HTML table with two header rows ---
            let html = '<table id="pivotTable"><colgroup><col></colgroup><thead>';

            // --- Add "Total" and "Line count" rows above header ---
            // --- Calculate Total row sums for each column ---
            let totalRowSums = [];
            html += '<tr>';
            html += '<th style="background:#e5e7eb;font-weight:bold;">Total</th>';
            displayGroups.forEach((groupKey, idx) => {
                let total = 0;
                if (expandLineStyle) {
                    Object.keys(aggMap).forEach(line => {
                        Object.keys(aggMap[line]).forEach(style => {
                            const cell = aggMap[line][style][groupKey];
                            if (cell) {
                                let cellRows = cell.rows || [];
                                const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                                const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                                let includeCell = true;
                                // --- FIX: Only sum rows matching selected delay statuses ---
                                if (delayColorChecked && selectedDelayStatuses.length > 0) {
                                    function getStatus(delivery, sew) {
                                        const dDate = new Date(delivery);
                                        const sDate = new Date(sew);
                                        if (isNaN(dDate) || isNaN(sDate)) return '';
                                        const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                                        if (diff < 0) return 'Delay';
                                        if (diff < 4 && diff > 0) return 'Can be delay';
                                        if (diff > 20) return 'Too early';
                                        return 'Ok';
                                    }
                                    // Filter rows by selected delay statuses
                                    cellRows = cellRows.filter(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                                    includeCell = cellRows.length > 0;
                                }
                                if (includeCell) {
                                    if (showSubprocess && selectedSubprocesses.length > 0) {
                                        cellRows.forEach(r => {
                                            total += sumSubprocess(r, selectedSubprocesses, usePoints);
                                        });
                                    } else {
                                        if (delayColorChecked && selectedDelayStatuses.length > 0) {
                                            // Sum only filtered rows
                                            total += useSamIeCalc
                                                ? cellRows.reduce((sum, r) => sum + (Number(r.PLAN_PCS || r.plan_pcs) || 0) * (Number(r.SAM_IE || r.sam_ie) || 0), 0)
                                                : cellRows.reduce((sum, r) => sum + (Number(r.PLAN_PCS || r.plan_pcs) || 0), 0);
                                        } else {
                                            total += useSamIeCalc ? cell.samIeSum : cell.pcsSum;
                                        }
                                    }
                                }
                            }
                        });
                    });
                } else {
                    Object.keys(aggMap).forEach(line => {
                        const cell = aggMap[line][groupKey];
                        if (cell) {
                            let cellRows = cell.rows || [];
                            const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                            const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                            let includeCell = true;
                            // --- FIX: Only sum rows matching selected delay statuses ---
                            if (delayColorChecked && selectedDelayStatuses.length > 0) {
                                function getStatus(delivery, sew) {
                                    const dDate = new Date(delivery);
                                    const sDate = new Date(sew);
                                    if (isNaN(dDate) || isNaN(sDate)) return '';
                                    const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                                    if (diff < 0) return 'Delay';
                                    if (diff < 4 && diff > 0) return 'Can be delay';
                                    if (diff > 20) return 'Too early';
                                    return 'Ok';
                                }
                                // Filter rows by selected delay statuses
                                cellRows = cellRows.filter(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                                includeCell = cellRows.length > 0;
                            }
                            if (includeCell) {
                                if (showSubprocess && selectedSubprocesses.length > 0) {
                                    cellRows.forEach(r => {
                                        total += sumSubprocess(r, selectedSubprocesses, usePoints);
                                    });
                                } else {
                                    if (delayColorChecked && selectedDelayStatuses.length > 0) {
                                        // Sum only filtered rows
                                        total += useSamIeCalc
                                            ? cellRows.reduce((sum, r) => sum + (Number(r.PLAN_PCS || r.plan_pcs) || 0) * (Number(r.SAM_IE || r.sam_ie) || 0), 0)
                                            : cellRows.reduce((sum, r) => sum + (Number(r.PLAN_PCS || r.plan_pcs) || 0), 0);
                                    } else {
                                        total += useSamIeCalc ? cell.samIeSum : cell.pcsSum;
                                    }
                                }
                            }
                        }
                    });
                }
                totalRowSums[idx] = total;
                html += `<td style="background:#dbeafe;color:#111;font-size:12px">${total ? formatNumberThousandSep(Math.round(total)) : ''}</td>`;
            });
            html += '</tr>';

            html += '<tr>';
            html += '<th style="background:#eff6ff;font-weight:bold;">Line count</th>';
            displayGroups.forEach(groupKey => {
                let count = 0;
                if (expandLineStyle) {
                    Object.keys(aggMap).forEach(line => {
                        Object.keys(aggMap[line]).forEach(style => {
                            const cell = aggMap[line][style][groupKey];
                            if (cell && cell.rows.length > 0) {
                                let cellRows = cell.rows || [];
                                const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                                const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                                let includeCell = true;
                                if (delayColorChecked && cellRows.length > 0) {
                                    const statusPriority = selectedDelayStatuses.length > 0 ? selectedDelayStatuses : ['Delay', 'Can be delay', 'Too early', 'Ok'];
                                    function getStatus(delivery, sew) {
                                        const dDate = new Date(delivery);
                                        const sDate = new Date(sew);
                                        if (isNaN(dDate) || isNaN(sDate)) return '';
                                        const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                                        if (diff < 0) return 'Delay';
                                        if (diff < 4 && diff > 0) return 'Can be delay';
                                        if (diff > 20) return 'Too early';
                                        return 'Ok';
                                    }
                                    includeCell = cellRows.some(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                                }
                                if (includeCell) {
                                    if (showSubprocess && selectedSubprocesses.length > 0) {
                                        // Only count if sum of subprocess > 0
                                        let spSum = 0;
                                        cell.rows.forEach(r => {
                                            spSum += sumSubprocess(r, selectedSubprocesses, usePoints);
                                        });
                                        if (spSum > 0) count++;
                                    } else {
                                        count++;
                                    }
                                }
                            }
                        });
                    });
                } else {
                    Object.keys(aggMap).forEach(line => {
                        const cell = aggMap[line][groupKey];
                        if (cell && cell.rows.length > 0) {
                            let cellRows = cell.rows || [];
                            const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                            const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                            let includeCell = true;
                            if (delayColorChecked && cellRows.length > 0) {
                                const statusPriority = selectedDelayStatuses.length > 0 ? selectedDelayStatuses : ['Delay', 'Can be delay', 'Too early', 'Ok'];
                                function getStatus(delivery, sew) {
                                    const dDate = new Date(delivery);
                                    const sDate = new Date(sew);
                                    if (isNaN(dDate) || isNaN(sDate)) return '';
                                    const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                                    if (diff < 0) return 'Delay';
                                    if (diff < 4 && diff > 0) return 'Can be delay';
                                    if (diff > 20) return 'Too early';
                                    return 'Ok';
                                }
                                includeCell = cellRows.some(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                            }
                            if (includeCell) {
                                if (showSubprocess && selectedSubprocesses.length > 0) {
                                    // Only count if sum of subprocess > 0
                                    let spSum = 0;
                                    cell.rows.forEach(r => {
                                        spSum += sumSubprocess(r, selectedSubprocesses, usePoints);
                                    });
                                    if (spSum > 0) count++;
                                } else {
                                    count++;
                                }
                            }
                        }
                    });
                }
                html += `<td style="background:#dbeafe;color:#111;font-size:12px">${count ? formatNumberThousandSep(count) : ''}</td>`;
            });
            html += '</tr>';

            // First header row: PROD_LINE (+ expand/collapse button) + headerGroups
            html += '<tr>';
            html += `<th rowspan="2" style="position:sticky;left:0;z-index:10;background:#fff;vertical-align:top;" id="lineHeaderCell">
        <button id="lineHeaderBtn" class="expand-collapse-btn" style="width:55px;text-align:left;">
            Line
        </button><br>`;
    if (expandLineStyle) {
        html += '<button id="collapseLineStyleBtn" class="expand-collapse-btn" style="margin-top:6px;font-size:1em;padding:2px 10px;">Collapse</button>';
    } else {
        html += '<button id="expandLineStyleBtn" class="expand-collapse-btn" style="margin-top:6px;font-size:1em;padding:2px 10px;">Expand</button>';
    }
    html += '</th>';
    headerGroups.forEach(g => {
        html += `<th class="date-col" colspan="${g.count}">${g.label}</th>`;
    });
    html += '</tr>';

    // Second header row: group labels
    html += '<tr>';
    displayGroups.forEach(groupKey => {
        let label = groupKey;
        if (groupBy === 'day') {
            const d = new Date(groupKey);
            label = String(d.getDate()).padStart(2, '0');
        } else if (groupBy === 'week') {
            label = groupKey.split('-W')[1];
        } else if (groupBy === 'month') {
            label = groupKey.split('-')[1];
        }
        html += `<th class="date-col" data-date="${groupKey}">${label}</th>`;
    });
    html += '</tr>';

    html += '</thead><tbody>';

    // Helper: sum selected subprocesses for a row
    // Modified: If usePointsMode, return value for each subprocess (for max logic), else sum as before
    function sumSubprocess(row, subprocesses, usePointsMode, soNoDocFilter) {
        if (usePointsMode && soNoDocFilter) {
            // Return an object of {subprocess: value} for max logic
            let result = {};
            subprocesses.forEach(sp => {
                let val = 0;
                if (sp === 'EMBROIDERY') val = Number(row.EMBROIDERY || row.embroidery) || 0;
                if (sp === 'HEAT') val = Number(row.HEAT || row.heat) || 0;
                if (sp === 'PAD_PRINT') val = Number(row.PAD_PRINT || row.pad_print) || 0;
                if (sp === 'PRINT') val = Number(row.PRINT || row.print) || 0;
                if (sp === 'BOND') val = Number(row.BOND || row.bond) || 0;
                if (sp === 'LASER') val = Number(row.LASER || row.laser) || 0;
                result[sp] = val;
            });
            return result;
        } else {
            let sum = 0;
            subprocesses.forEach(sp => {
                let val = 0;
                if (sp === 'EMBROIDERY') val = Number(row.EMBROIDERY || row.embroidery) || 0;
                if (sp === 'HEAT') val = Number(row.HEAT || row.heat) || 0;
                if (sp === 'PAD_PRINT') val = Number(row.PAD_PRINT || row.pad_print) || 0;
                if (sp === 'PRINT') val = Number(row.PRINT || row.print) || 0;
                if (sp === 'BOND') val = Number(row.BOND || row.bond) || 0;
                if (sp === 'LASER') val = Number(row.LASER || row.laser) || 0;
                const planPcs = Number(row.PLAN_PCS || row.plan_pcs) || 0;
                sum += val * planPcs;
            });
            return sum;
        }
    }

    if (expandLineStyle) {
        // Expanded: show line + style rows
        // Alternate row background for each line (just like collapsed mode)
        const lineBgColors = ['#f7f9fb', '#e6e9ed'];
        let lineIdx = 0;
        Object.keys(aggMap).sort().forEach(line => {
            const bgColor = lineBgColors[lineIdx % 2];
            Object.keys(aggMap[line]).sort().forEach(style => {
                html += `<tr><td style="background:${bgColor};">${line} <span style="color:#888;">/</span> <span style="font-weight:bold;background:none;">${style}</span></td>`;
                displayGroups.forEach(groupKey => {
                    let cellValue = '';
                    let cellRows = aggMap[line][style][groupKey]?.rows || [];
                    let cellStyle = '';
                    let cellClass = '';
                    if (aggMap[line][style][groupKey]) {
                        const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                        const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                        // --- Only sum PCS for rows matching selected delay statuses if delay filter is active ---
                        if (delayColorChecked && selectedDelayStatuses.length > 0) {
                            function getStatus(delivery, sew) {
                                const dDate = new Date(delivery);
                                const sDate = new Date(sew);
                                if (isNaN(dDate) || isNaN(sDate)) return '';
                                const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                                if (diff < 0) return 'Delay';
                                if (diff < 4 && diff > 0) return 'Can be delay';
                                if (diff > 20) return 'Too early';
                                return 'Ok';
                            }
                            // Filter rows by selected delay statuses
                            cellRows = cellRows.filter(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                        }
                        if (showSubprocess && selectedSubprocesses.length > 0) {
                            if (usePoints) {
                                // Only if exactly one SO No Doc is selected
                                const soNoDocSelected = getSelectedValues('soNoDocSelect');
                                if (soNoDocSelected.length === 1) {
                                    // For each subprocess, find the max value among all rows for the selected SO No Doc, then sum those max values
                                    let maxBySub = {};
                                    selectedSubprocesses.forEach(sp => { maxBySub[sp] = 0; });
                                    cellRows.forEach(r => {
                                        if ((r.SO_NO_DOC || r.so_no_doc) === soNoDocSelected[0]) {
                                            const vals = sumSubprocess(r, selectedSubprocesses, true, soNoDocSelected[0]);
                                            selectedSubprocesses.forEach(sp => {
                                                if (vals[sp] > maxBySub[sp]) maxBySub[sp] = vals[sp];
                                            });
                                        }
                                    });
                                    let sumMax = 0;
                                    selectedSubprocesses.forEach(sp => { sumMax += maxBySub[sp]; });
                                    cellValue = sumMax ? formatNumberThousandSep(sumMax) : '';
                                } else {
                                    cellValue = '';
                                }
                            } else {
                                let spSum = 0;
                                cellRows.forEach(r => {
                                    spSum += sumSubprocess(r, selectedSubprocesses, false);
                                });
                                cellValue = spSum ? formatNumberThousandSep(spSum) : '';
                            }
                        } else {
                            // --- Sum PCS only for filtered rows ---
                            if (useSamIeCalc) {
                                let samIeSum = 0;
                                cellRows.forEach(r => {
                                    const pcs = Number(r.PLAN_PCS || r.plan_pcs) || 0;
                                    const sam = Number(r.SAM_IE || r.sam_ie) || 0;
                                    samIeSum += pcs * sam;
                                });
                                cellValue = samIeSum ? formatNumberThousandSep(Math.round(samIeSum)) : '';
                            } else {
                                let pcsSum = 0;
                                cellRows.forEach(r => {
                                    pcsSum += Number(r.PLAN_PCS || r.plan_pcs) || 0;
                                });
                                cellValue = pcsSum ? formatNumberThousandSep(pcsSum) : '';
                            }
                        }
                    }
                    const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                    const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                    if (styleColorChecked && cellRows.length > 0) {
                        // Show style ref in cell, color by style
                        if (style && styleColorMap[style]) {
                            cellStyle = `background:${styleColorMap[style]} !important;color:#111 !important;`;
                            cellClass = 'cell-row-style-color';
                            cellValue = style;
                        } else {
                            cellValue = '';
                        }
                    } else if (delayColorChecked && cellRows.length > 0) {
                        const statusPriority = selectedDelayStatuses.length > 0 ? selectedDelayStatuses : ['Delay', 'Can be delay', 'Too early', 'Ok'];
                        function getStatus(delivery, sew) {
                            const dDate = new Date(delivery);
                            const sDate = new Date(sew);
                            if (isNaN(dDate) || isNaN(sDate)) return '';
                            const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                            if (diff < 0) return 'Delay';
                            if (diff < 4 && diff > 0) return 'Can be delay';
                            if (diff > 20) return 'Too early';
                            return 'Ok';
                        }
                        // Only show value if at least one row matches selectedDelayStatuses
                        const hasSelectedStatus = cellRows.some(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                        if (!hasSelectedStatus) {
                            cellValue = '';
                            cellStyle = 'background:#CECECE !important;color:#111 !important;';
                        } else {
                            let foundStatus = '';
                            for (const prio of statusPriority) {
                                if (cellRows.some(r => getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date) === prio)) {
                                    foundStatus = prio;
                                    break;
                                }
                            }
                            if (foundStatus === 'Delay') {
                                cellStyle = 'background:#fee2e2 !important;color:#111 !important;';
                                cellClass = 'cell-row-delay';
                            } else if (foundStatus === 'Can be delay') {
                                cellStyle = 'background:#fef9c3 !important;color:#111 !important;';
                                cellClass = 'cell-row-can-delay';
                            } else if (foundStatus === 'Too early') {
                                cellStyle = 'background:#e0e7ff !important;color:#111 !important;';
                                cellClass = 'cell-row-too-early';
                            } else if (foundStatus === 'Ok') {
                                cellStyle = 'background:#dcfce7 !important;color:#111 !important;';
                                cellClass = 'cell-row-ok';
                            }
                        }
                    }
                    // --- Add blank cell coloring ---
                    if (!cellValue) {
                        cellStyle = 'background:#CECECE !important;color:#111 !important;';
                    }
                    html += `<td class="date-col${cellClass ? ' ' + cellClass : ''}" data-full-date="${groupKey}"${cellStyle ? ` style="${cellStyle}"` : ''}>${cellValue}</td>`;
                });
                html += '</tr>';
            });
            lineIdx++;
        });
    } else {
        // Normal: show line rows
        Object.keys(aggMap).sort().forEach(line => {
            html += `<tr><td>${line}</td>`;
            displayGroups.forEach(groupKey => {
                let cellValue = '';
                let cellStyle = '';
                let cellRows = aggMap[line][groupKey]?.rows || [];
                let cellClass = '';
                if (aggMap[line][groupKey]) {
                    const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                    const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                    // --- Only sum PCS for rows matching selected delay statuses if delay filter is active ---
                    if (delayColorChecked && selectedDelayStatuses.length > 0) {
                        function getStatus(delivery, sew) {
                            const dDate = new Date(delivery);
                            const sDate = new Date(sew);
                            if (isNaN(dDate) || isNaN(sDate)) return '';
                            const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                            if (diff < 0) return 'Delay';
                            if (diff < 4 && diff > 0) return 'Can be delay';
                            if (diff > 20) return 'Too early';
                            return 'Ok';
                        }
                        // Filter rows by selected delay statuses
                        cellRows = cellRows.filter(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                    }
                    if (showSubprocess && selectedSubprocesses.length > 0) {
                        if (usePoints) {
                            // Only if exactly one SO No Doc is selected
                            const soNoDocSelected = getSelectedValues('soNoDocSelect');
                            if (soNoDocSelected.length === 1) {
                                // For each subprocess, find the max value among all rows for the selected SO No Doc, then sum those max values
                                let maxBySub = {};
                                selectedSubprocesses.forEach(sp => { maxBySub[sp] = 0; });
                                cellRows.forEach(r => {
                                    if ((r.SO_NO_DOC || r.so_no_doc) === soNoDocSelected[0]) {
                                        const vals = sumSubprocess(r, selectedSubprocesses, true, soNoDocSelected[0]);
                                        selectedSubprocesses.forEach(sp => {
                                            if (vals[sp] > maxBySub[sp]) maxBySub[sp] = vals[sp];
                                        });
                                    }
                                });
                                let sumMax = 0;
                                selectedSubprocesses.forEach(sp => { sumMax += maxBySub[sp]; });
                                cellValue = sumMax ? formatNumberThousandSep(sumMax) : '';
                            } else {
                                cellValue = '';
                            }
                        } else {
                            let spSum = 0;
                            cellRows.forEach(r => {
                                spSum += sumSubprocess(r, selectedSubprocesses, false);
                            });
                            cellValue = spSum ? formatNumberThousandSep(spSum) : '';
                        }
                    } else {
                        // --- Sum PCS only for filtered rows ---
                        if (useSamIeCalc) {
                            let samIeSum = 0;
                            cellRows.forEach(r => {
                                const pcs = Number(r.PLAN_PCS || r.plan_pcs) || 0;
                                const sam = Number(r.SAM_IE || r.sam_ie) || 0;
                                samIeSum += pcs * sam;
                            });
                            cellValue = samIeSum ? formatNumberThousandSep(Math.round(samIeSum)) : '';
                        } else {
                            let pcsSum = 0;
                            cellRows.forEach(r => {
                                pcsSum += Number(r.PLAN_PCS || r.plan_pcs) || 0;
                            });
                            cellValue = pcsSum ? formatNumberThousandSep(pcsSum) : '';
                        }
                    }
                }
                const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
                if (styleColorChecked && cellRows.length > 0) {
                    // Find style with highest sum PCS in cellRows
                    const styleSums = {};
                    cellRows.forEach(r => {
                        const style = r.STYLE_REF || r.style_ref;
                        const pcs = Number(r.PLAN_PCS || r.plan_pcs) || 0;
                        if (!style) return;
                        styleSums[style] = (styleSums[style] || 0) + pcs;
                    });
                    let maxStyle = '';
                    let maxSum = -1;
                    Object.entries(styleSums).forEach(([style, sum]) => {
                        if (sum > maxSum) {
                            maxSum = sum;
                            maxStyle = style;
                        }
                    });
                    if (maxStyle && styleColorMap[maxStyle]) {
                        cellStyle = `background:${styleColorMap[maxStyle]} !important;color:#111 !important;`;
                        cellClass = 'cell-row-style-color';
                        cellValue = maxStyle;
                    } else {
                        cellValue = '';
                    }
                } else if (delayColorChecked && cellRows.length > 0) {
                    // Color code logic for cell
                    const statusPriority = selectedDelayStatuses.length > 0 ? selectedDelayStatuses : ['Delay', 'Can be delay', 'Too early', 'Ok'];
                    function getStatus(delivery, sew) {
                        const dDate = new Date(delivery);
                        const sDate = new Date(sew);
                        if (isNaN(dDate) || isNaN(sDate)) return '';
                        const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                        if (diff < 0) return 'Delay';
                        if (diff < 4 && diff > 0) return 'Can be delay';
                        if (diff > 20) return 'Too early';
                        return 'Ok';
                    }
                    // Only show value if at least one row matches selectedDelayStatuses
                    const hasSelectedStatus = cellRows.some(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
                    if (!hasSelectedStatus) {
                        cellValue = '';
                        cellStyle = 'background:#CECECE !important;color:#111 !important;';
                    } else {
                        // Color code logic for cell
                        let foundStatus = '';
                        for (const prio of statusPriority) {
                            if (cellRows.some(r => getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date) === prio)) {
                                foundStatus = prio;
                                break;
                            }
                        }
                        if (foundStatus === 'Delay') {
                            cellStyle = 'background:#fee2e2 !important;color:#111 !important;';
                            cellClass = 'cell-row-delay';
                        } else if (foundStatus === 'Can be delay') {
                            cellStyle = 'background:#fef9c3 !important;color:#111 !important;';
                            cellClass = 'cell-row-can-delay';
                        } else if (foundStatus === 'Too early') {
                            cellStyle = 'background:#e0e7ff !important;color:#111 !important;';
                            cellClass = 'cell-row-too-early';
                        } else if (foundStatus === 'Ok') {
                            cellStyle = 'background:#dcfce7 !important;color:#111 !important;';
                            cellClass = 'cell-row-ok';
                        }
                    }
                }
                // --- Add blank cell coloring ---
                if (!cellValue) {
                    cellStyle = 'background:#CECECE !important;color:#111 !important;';
                }
                html += `<td class="date-col${cellClass ? ' ' + cellClass : ''}" data-full-date="${groupKey}"${cellStyle ? ` style="${cellStyle}"` : ''}>${cellValue}</td>`;
            });
            html += '</tr>';
        });
    }
    html += '</tbody></table>';

    // --- After table is rendered ---
    document.getElementById('pivotTableContainer').innerHTML = html;
    addPivotTableCellClickHandlers(filteredData);

    // --- Add/update the "Total Sum" card next to By Points ---
    setTimeout(() => {
        // Find the By Points container
        const byPointsContainer = document.getElementById('byPointsContainer');
        // If not found, try to find the parent of the togglePoints checkbox
        let insertAfterElem = byPointsContainer || document.getElementById('togglePoints')?.parentElement;
        if (insertAfterElem) {
            // Remove previous card if exists
            let totalCard = document.getElementById('totalSumCard');
            if (totalCard) totalCard.remove();
            // Calculate the grand total (sum of all columns in Total row)
            const grandTotal = totalRowSums.reduce((a, b) => a + b, 0);
            // Create the card element
            totalCard = document.createElement('div');
            totalCard.id = 'totalSumCard';
            totalCard.style.cssText = `
                display: inline-flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                min-width: 150px;
                height: 40px;
                background: #f8fafc;
                border-radius: 6px;
                margin-left: 18px;
                border: 1px solid #d1d5db;
                font-size: 1.15em;
                font-weight: bold;
                color: #2563eb;
            `;
            totalCard.innerHTML = `
                <span style="font-size:1.25em;color:#2563eb;font-weight:bold;">${grandTotal ? formatNumberThousandSep(Math.round(grandTotal)) : ''}</span>
            `;
            // Insert after By Points container
            insertAfterElem.parentNode.insertBefore(totalCard, insertAfterElem.nextSibling);
        }
        const expandBtn = document.getElementById('expandLineStyleBtn');
        const collapseBtn = document.getElementById('collapseLineStyleBtn');
        // --- Use availableLines for modal options ---
        if (expandBtn) {
            expandBtn.className = 'expand-collapse-btn';
            expandBtn.onclick = function(e) {
                e.stopPropagation();
                expandLineStyle = true;
                showLoadingSpinnerOverlay();
                setTimeout(() => renderDashboard(filteredData), 0);
            };
        }
        if (collapseBtn) {
            collapseBtn.className = 'expand-collapse-btn';
            collapseBtn.onclick = function(e) {
                e.stopPropagation();
                expandLineStyle = false;
                showLoadingSpinnerOverlay();
                setTimeout(() => renderDashboard(filteredData), 0);
            };
        }
        const lineHeaderBtn = document.getElementById('lineHeaderBtn');
        if (lineHeaderBtn) {
            lineHeaderBtn.onclick = function(e) {
                e.stopPropagation();
                showLineModal(availableLines); // <-- always show all lines in filteredData before selectedLines filter
            };
        }
    }, 0);
}

// --- Modal for line selection ---
function showLineModal(linesArr) {
    let modalId = 'lineModal';
    if (!document.getElementById(modalId)) {
        const modal = document.createElement('div');
        modal.id = modalId;
        modal.className = 'modal-overlay';
        modal.style.zIndex = '10001';
        modal.innerHTML = `
            <div class="modal-content" style="min-width:320px;max-width:96vw;">
                <div style="font-weight:bold;font-size:1.1em;margin-bottom:8px;">Select lines to display</div>
                <div style="margin-bottom:10px;">
                    <button type="button" id="lineModalSelectAllBtn" style="background:#2563eb;color:#fff;border:none;border-radius:5px;padding:4px 12px;font-weight:500;margin-right:8px;">Select All</button>
                    <button type="button" id="lineModalDeselectAllBtn" style="background:#e5e7eb;color:#333;border:none;border-radius:5px;padding:4px 12px;font-weight:500;">Deselect All</button>
                </div>
                <form id="lineForm" style="display:flex;flex-direction:column;gap:8px;max-height:50vh;overflow:auto;">
                </form>
                <div style="margin-top:16px;display:flex;justify-content:flex-end;gap:12px;">
                    <button type="button" id="lineModalOkBtn" style="background:#219a0b;color:#fff;border:none;border-radius:5px;padding:6px 18px;font-weight:600;">OK</button>
                    <button type="button" id="lineModalCancelBtn" style="background:#eee;color:#333;border:none;border-radius:5px;padding:6px 18px;">Cancel</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }
    const modal = document.getElementById(modalId);
    const form = document.getElementById('lineForm');
    // Populate checkboxes for all lines, keep unchecked lines as options
    form.innerHTML = linesArr.map(line =>
        `<label><input type="checkbox" value="${line}" ${selectedLines.length === 0 || selectedLines.includes(line) ? 'checked' : ''}> ${line}</label>`
    ).join('');
    modal.style.display = 'flex';

    // Select All/Deselect All logic
    document.getElementById('lineModalSelectAllBtn').onclick = function() {
        Array.from(form.elements).forEach(el => {
            if (el.type === 'checkbox') el.checked = true;
        });
    };
    document.getElementById('lineModalDeselectAllBtn').onclick = function() {
        Array.from(form.elements).forEach(el => {
            if (el.type === 'checkbox') el.checked = false;
        });
    };

    // OK/Cancel logic
    document.getElementById('lineModalOkBtn').onclick = function() {
        selectedLines = Array.from(form.elements)
            .filter(el => el.type === 'checkbox' && el.checked)
            .map(el => el.value);
        modal.style.display = 'none';
        showLoadingSpinnerOverlay();
        setTimeout(filterAndRender, 0);
    };
    document.getElementById('lineModalCancelBtn').onclick = function() {
        modal.style.display = 'none';
    };
    modal.onclick = function(e) {
        if (e.target === modal) modal.style.display = 'none';
    };
}

// --- Patch filterAndRender to use selectedLines ---
function filterAndRender() {
    showLoadingSpinnerOverlay();
    setTimeout(() => {
        updateAllDropdownsExceptFactory();
        const filters = getCurrentFilters();
        let filteredData = allData.filter(row => {
            const factory = row.FACTORY || row.factory;
            const styleRef = row.STYLE_REF || row.style_ref;
            const customerName = row.CUSTOMER_NAME || row.customer_name;
            const soNoDoc = row.SO_NO_DOC || row.so_no_doc;
            const productType = row.PRODUCT_TYPE || row.product_type;
            const subNo = row.SUB_NO || row.sub_no;
            const colorFc = row.COLOR_FC || row.color_fc;
            const dateRaw = row.SEW_DATE || row.sew_date;
            let date = '';
            if (dateRaw) {
                const dateObj = new Date(dateRaw);
                if (!isNaN(dateObj)) {
                    date = dateObj.toISOString().split('T')[0];
                }
            }
            return (
                (filters.factory.length === 0 || filters.factory.includes(factory)) &&
                (filters.styleRef.length === 0 || filters.styleRef.includes(styleRef)) &&
                (filters.customerName.length === 0 || filters.customerName.includes(customerName)) &&
                (filters.soNoDoc.length === 0 || filters.soNoDoc.includes(soNoDoc)) &&
                (filters.productType.length === 0 || filters.productType.includes(productType)) &&
                (filters.subNo.length === 0 || filters.subNo.includes(subNo)) &&
                (filters.colorFc.length === 0 || filters.colorFc.includes(colorFc)) &&
                (!filters.dateStart || date >= filters.dateStart) &&
                (!filters.dateEnd || date <= filters.dateEnd)
            );
        });

        // --- SO No Doc type filtering for table ---
        // Use window.soNoDocTypeSelections set by modal logic, or fallback to both
        let soNoDocTypeSelections = window.soNoDocTypeSelections || [];
        if (soNoDocTypeSelections.length === 1) {
            const type = soNoDocTypeSelections[0];
            filteredData = filteredData.filter(row => {
                const soNoDocVal = String(row.SO_NO_DOC || row.so_no_doc || '');
                const digits = soNoDocVal.replace(/\D/g, '');
                if (type === 'Early') return digits.length === 7 && /^[478]/.test(digits);
                if (type === 'Bulk') return digits.length === 7 && !/^[478]/.test(digits);
                if (type === 'Sample') return digits.length === 8 || (digits.length === 7 && /^2/.test(digits));
                return true;
            });
        }
        // --- Save all lines present in filteredData before selectedLines filter ---
        availableLines = Array.from(new Set(filteredData.map(row => row.PROD_LINE || row.prod_line || 'Unknown'))).sort();

        // --- Filter by selectedLines ---
        let filteredByLines = filteredData;
        if (selectedLines.length > 0) {
            filteredByLines = filteredData.filter(row => {
                const line = row.PROD_LINE || row.prod_line || 'Unknown';
                return selectedLines.includes(line);
            });
        }
        // --- Determine if By Points should be enabled ---
        const soNoDocSelected = getSelectedValues('soNoDocSelect');
        const togglePoints = document.getElementById('togglePoints');
        if (togglePoints) {
            if (soNoDocSelected.length === 1) {
                togglePoints.disabled = false;
            } else {
                togglePoints.checked = false;
                usePoints = false;
                togglePoints.disabled = true;
            }
        }
        window._lastFilteredData = filteredByLines; // Save for graph
        renderDashboard(filteredByLines);
        if (currentView === 'graph') {
            renderPlanPcsGraph(filteredByLines);
        }
        hideLoadingSpinnerOverlay();
    }, 0);
}

if (typeof window.soNoDocTypeSelections === 'undefined') {
    window.soNoDocTypeSelections = ['Bulk', 'Sample', 'Early'];
}

// Choices.js instances
        let choicesInstances = {};

        // Helper to get selected values (array)
        function getSelectedValues(id) {
            const select = document.getElementById(id);
            const selected = Array.from(select.selectedOptions).map(opt => opt.value);
            if (id === 'factorySelect') {
                // If "All" is selected or nothing is selected, treat as no filter
                if (selected.length === 0 || selected.includes('')) return [];
                return selected;
            }
            return selected.filter(v => v !== '' );
        }

        // Initialize Choices.js for all dropdowns (including factorySelect)
        function initChoicesDropdowns() {
            [
                'factorySelect',
                'styleRefSelect','customerNameSelect','soNoDocSelect',
                'productTypeSelect','subNoSelect','colorFcSelect'
            ].forEach(id => {
                if (!choicesInstances[id]) {
                    choicesInstances[id] = new Choices(document.getElementById(id), {
                        removeItemButton: true,
                        searchEnabled: true,
                        placeholderValue: 'All',
                        shouldSort: false,
                        itemSelectText: ''
                    });
                }
            });
        }

        // Update dropdown options using Choices.js API, preserving selection
        function updateDropdown(selectId, values) {
            const choices = choicesInstances[selectId];
            if (!choices) return;

            // Get currently selected values
            const selectedValues = getSelectedValues(selectId);

            // --- Filter SO No Doc options by type ---
            if (selectId === 'soNoDocSelect') {
                let showBulk = soNoDocTypeSelections.includes('Bulk');
                let showSample = soNoDocTypeSelections.includes('Sample');
                let showEarly = soNoDocTypeSelections.includes('Early');
                values = values.filter(val => {
                    const digits = String(val).replace(/\D/g, '');
                    if (showEarly && digits.length === 7 && /^[478]/.test(digits)) return true;
                    if (showBulk && digits.length === 7 && !/^[478]/.test(digits)) return true;
                    if (showSample && digits.length === 8) return true;
                    if (showBulk && showSample && showEarly) return true;
                    return false;
                });
            }

            // Only show unselected values in dropdown choices
            const unselectedValues = values.filter(val => !selectedValues.includes(val));

            // Clear previous choices to avoid duplication
            choices.clearChoices();

            // If nothing is selected, show "All" as placeholder and as the only option
            if (selectedValues.length === 0) {
                choices.setChoices(
                    values.map(val => ({
                        value: val,
                        label: val,
                        disabled: false
                    })),
                    'value', 'label', false
                );
                choices.setValue([]);
                choices._store.placeholder = true;
                choices._store.placeholderValue = 'All';
                choices.input.element.placeholder = 'All';
            } else {
                // Only show unselected values in the dropdown
                choices.setChoices(
                    unselectedValues.map(val => ({
                        value: val,
                        label: val,
                        disabled: false
                    })),
                    'value', 'label', false
                );
                choices._store.placeholder = false;
                choices.input.element.placeholder = '';
            }
        }

        // Update all dropdowns including factorySelect
        function updateAllDropdowns() {
            updateDropdown(
                'factorySelect',
                Array.from(new Set(getFilteredForDropdown('factory').map(row => row.FACTORY || row.factory).filter(Boolean))).sort()
            );
            updateDropdown(
                'styleRefSelect',
                Array.from(new Set(getFilteredForDropdown('styleRef').map(row => row.STYLE_REF || row.style_ref).filter(Boolean))).sort()
            );
            updateDropdown(
                'customerNameSelect',
                Array.from(new Set(getFilteredForDropdown('customerName').map(row => row.CUSTOMER_NAME || row.customer_name).filter(Boolean))).sort()
            );
            updateDropdown(
                'soNoDocSelect',
                Array.from(new Set(getFilteredForDropdown('soNoDoc').map(row => row.SO_NO_DOC || row.so_no_doc).filter(Boolean))).sort()
            );
            updateDropdown(
                'productTypeSelect',
                Array.from(new Set(getFilteredForDropdown('productType').map(row => row.PRODUCT_TYPE || row.product_type).filter(Boolean))).sort()
            );
            updateDropdown(
                'subNoSelect',
                Array.from(new Set(getFilteredForDropdown('subNo').map(row => row.SUB_NO || row.sub_no).filter(Boolean))).sort()
            );
            updateDropdown(
                'colorFcSelect',
                Array.from(new Set(getFilteredForDropdown('colorFc').map(row => row.COLOR_FC || row.color_fc).filter(Boolean))).sort()
            );
        }

        function getCurrentFilters() {
            return {
                factory: getSelectedValues('factorySelect'),
                styleRef: getSelectedValues('styleRefSelect'),
                customerName: getSelectedValues('customerNameSelect'),
                soNoDoc: getSelectedValues('soNoDocSelect'),
                productType: getSelectedValues('productTypeSelect'),
                subNo: getSelectedValues('subNoSelect'),
                colorFc: getSelectedValues('colorFcSelect'),
                dateStart: document.getElementById('dateStart').value,
                dateEnd: document.getElementById('dateEnd').value
            };
        }

        function getFilteredForDropdown(dropdownId) {
            const filters = getCurrentFilters();
            let tempFilters = { ...filters };
            tempFilters[dropdownId] = [];
            return allData.filter(row => {
                const factory = row.FACTORY || row.factory;
                const styleRef = row.STYLE_REF || row.style_ref;
                const customerName = row.CUSTOMER_NAME || row.customer_name;
                const soNoDoc = row.SO_NO_DOC || row.so_no_doc;
                const productType = row.PRODUCT_TYPE || row.product_type;
                const subNo = row.SUB_NO || row.sub_no;
                const colorFc = row.COLOR_FC || row.color_fc;
                const dateRaw = row.SEW_DATE || row.sew_date;
                let date = '';
                if (dateRaw) {
                    const dateObj = new Date(dateRaw);
                    if (!isNaN(dateObj)) {
                        date = dateObj.toISOString().split('T')[0];
                    }
                }
                return (
                    (tempFilters.factory.length === 0 || tempFilters.factory.includes(factory)) &&
                    (tempFilters.styleRef.length === 0 || tempFilters.styleRef.includes(styleRef)) &&
                    (tempFilters.customerName.length === 0 || tempFilters.customerName.includes(customerName)) &&
                    (tempFilters.soNoDoc.length === 0 || tempFilters.soNoDoc.includes(soNoDoc)) &&
                    (tempFilters.productType.length === 0 || tempFilters.productType.includes(productType)) &&
                    (tempFilters.subNo.length === 0 || tempFilters.subNo.includes(subNo)) &&
                    (tempFilters.colorFc.length === 0 || tempFilters.colorFc.includes(colorFc)) &&
                    (!tempFilters.dateStart || date >= tempFilters.dateStart) &&
                    (!tempFilters.dateEnd || date <= tempFilters.dateEnd)
                );
            });
        }

        function updateAllDropdownsExceptFactory() {
            updateDropdown(
                'factorySelect',
                Array.from(new Set(getFilteredForDropdown('factory').map(row => row.FACTORY || row.factory).filter(Boolean))).sort()
            );
            updateDropdown(
                'styleRefSelect',
                Array.from(new Set(getFilteredForDropdown('styleRef').map(row => row.STYLE_REF || row.style_ref).filter(Boolean))).sort()
            );
            updateDropdown(
                'customerNameSelect',
                Array.from(new Set(getFilteredForDropdown('customerName').map(row => row.CUSTOMER_NAME || row.customer_name).filter(Boolean))).sort()
            );
            updateDropdown(
                'soNoDocSelect',
                Array.from(new Set(getFilteredForDropdown('soNoDoc').map(row => row.SO_NO_DOC || row.so_no_doc).filter(Boolean))).sort()
            );
            updateDropdown(
                'productTypeSelect',
                Array.from(new Set(getFilteredForDropdown('productType').map(row => row.PRODUCT_TYPE || row.product_type).filter(Boolean))).sort()
            );
            updateDropdown(
                'subNoSelect',
                Array.from(new Set(getFilteredForDropdown('subNo').map(row => row.SUB_NO || row.sub_no).filter(Boolean))).sort()
            );
            updateDropdown(
                'colorFcSelect',
                Array.from(new Set(getFilteredForDropdown('colorFc').map(row => row.COLOR_FC || row.color_fc).filter(Boolean))).sort()
            );
        }

        function updateFactoryDropdown() {
            const select = document.getElementById('factorySelect');
            select.innerHTML = '';
            const factoryArr = Array.from(factories).sort();
            factoryArr.forEach(val => {
                const opt = document.createElement('option');
                opt.value = val;
                opt.textContent = val;
                select.appendChild(opt);
            });
            // Select all by default (means "all factories")
            Array.from(select.options).forEach(opt => opt.selected = true);
            updateAllDropdownsExceptFactory();
        }

        // --- Factory selection persistence ---
        function saveFactorySelection() {
            try {
                const select = document.getElementById('factorySelect');
                const selected = Array.from(select.selectedOptions).map(opt => opt.value);
                localStorage.setItem('factorySelect', JSON.stringify(selected));
            } catch (e) {}
        }
        function loadFactorySelection() {
            try {
                const val = localStorage.getItem('factorySelect');
                if (!val) return [];
                return JSON.parse(val);
            } catch (e) { return []; }
        }

        // Patch updateFactoryDropdown to restore selection from localStorage
        function updateFactoryDropdown() {
            const select = document.getElementById('factorySelect');
            const prev = loadFactorySelection();
            select.innerHTML = '';
            // Add factory options
            const factoryArr = Array.from(factories).sort();
            factoryArr.forEach(val => {
                const opt = document.createElement('option');
                opt.value = val;
                opt.textContent = val;
                select.appendChild(opt);
            });
            // Restore previous selection if possible, else select all
            let found = false;
            prev.forEach(v => {
                if (v && factoryArr.includes(v)) {
                    select.querySelector(`option[value="${v}"]`).selected = true;
                    found = true;
                }
            });
            if (!found) {
                // Select all by default (means "all factories")
                Array.from(select.options).forEach(opt => opt.selected = true);
            }
            updateAllDropdownsExceptFactory();
        }

        // Save factory selection on change
        document.addEventListener('DOMContentLoaded', function() {
            const factorySelect = document.getElementById('factorySelect');
            if (factorySelect) {
                factorySelect.addEventListener('change', saveFactorySelection);
            }
        });

        function showLoadingSpinnerOverlay() {
            let container = document.getElementById('pivotTableContainer');
            let overlay = container.querySelector('.table-loading-overlay');
            if (!overlay) {
                overlay = document.createElement('div');
                overlay.className = 'table-loading-overlay';
                overlay.innerHTML = '<div class="loading-spinner"></div>';
                container.appendChild(overlay);
            }
            overlay.style.display = 'flex';
        }
        function hideLoadingSpinnerOverlay() {
            let container = document.getElementById('pivotTableContainer');
            let overlay = container.querySelector('.table-loading-overlay');
            if (overlay) overlay.style.display = 'none';
        }

        function filterAndRender() {
            showLoadingSpinnerOverlay();
            setTimeout(() => {
                updateAllDropdownsExceptFactory();
                const filters = getCurrentFilters();
                let filteredData = allData.filter(row => {
                    const factory = row.FACTORY || row.factory;
                    const styleRef = row.STYLE_REF || row.style_ref;
                    const customerName = row.CUSTOMER_NAME || row.customer_name;
                    const soNoDoc = row.SO_NO_DOC || row.so_no_doc;
                    const productType = row.PRODUCT_TYPE || row.product_type;
                    const subNo = row.SUB_NO || row.sub_no;
                    const colorFc = row.COLOR_FC || row.color_fc;
                    const dateRaw = row.SEW_DATE || row.sew_date;
                    let date = '';
                    if (dateRaw) {
                        const dateObj = new Date(dateRaw);
                        if (!isNaN(dateObj)) {
                            date = dateObj.toISOString().split('T')[0];
                        }
                    }
                    return (
                        (filters.factory.length === 0 || filters.factory.includes(factory)) &&
                        (filters.styleRef.length === 0 || filters.styleRef.includes(styleRef)) &&
                        (filters.customerName.length === 0 || filters.customerName.includes(customerName)) &&
                        (filters.soNoDoc.length === 0 || filters.soNoDoc.includes(soNoDoc)) &&
                        (filters.productType.length === 0 || filters.productType.includes(productType)) &&
                        (filters.subNo.length === 0 || filters.subNo.includes(subNo)) &&
                        (filters.colorFc.length === 0 || filters.colorFc.includes(colorFc)) &&
                        (!filters.dateStart || date >= filters.dateStart) &&
                        (!filters.dateEnd || date <= filters.dateEnd)
                    );
                });

                // --- SO No Doc type filtering for table ---
                // Use window.soNoDocTypeSelections set by modal logic, or fallback to both
                let soNoDocTypeSelections = window.soNoDocTypeSelections || [];
                if (soNoDocTypeSelections.length === 1) {
                    const type = soNoDocTypeSelections[0];
                    filteredData = filteredData.filter(row => {
                        const soNoDocVal = String(row.SO_NO_DOC || row.so_no_doc || '');
                        const digits = soNoDocVal.replace(/\D/g, '');
                        if (type === 'Early') return digits.length === 7 && /^[478]/.test(digits);
                        if (type === 'Bulk') return digits.length === 7 && !/^[478]/.test(digits);
                        if (type === 'Sample') return digits.length === 8 || (digits.length === 7 && /^2/.test(digits));
                        return true;
                    });
                }
                // --- Save all lines present in filteredData before selectedLines filter ---
                availableLines = Array.from(new Set(filteredData.map(row => row.PROD_LINE || row.prod_line || 'Unknown'))).sort();

                // --- Filter by selectedLines ---
                let filteredByLines = filteredData;
                if (selectedLines.length > 0) {
                    filteredByLines = filteredData.filter(row => {
                        const line = row.PROD_LINE || row.prod_line || 'Unknown';
                        return selectedLines.includes(line);
                    });
                }
                // --- Determine if By Points should be enabled ---
                const soNoDocSelected = getSelectedValues('soNoDocSelect');
                const togglePoints = document.getElementById('togglePoints');
                if (togglePoints) {
                    if (soNoDocSelected.length === 1) {
                        togglePoints.disabled = false;
                    } else {
                        togglePoints.checked = false;
                        usePoints = false;
                        togglePoints.disabled = true;
                    }
                }
                window._lastFilteredData = filteredByLines; // Save for graph
                renderDashboard(filteredByLines);
                if (currentView === 'graph') {
                    renderPlanPcsGraph(filteredByLines);
                }
                hideLoadingSpinnerOverlay();
            }, 0);
        }

        function resetFilters() {
            showLoadingSpinnerOverlay();
            // Clear all selections for Choices.js dropdowns
            [
                'factorySelect','styleRefSelect','customerNameSelect','soNoDocSelect',
                'productTypeSelect','subNoSelect','colorFcSelect'
            ].forEach(id => {
                if (choicesInstances[id]) {
                    choicesInstances[id].removeActiveItems();
                    // Explicitly reset Choices.js input value and placeholder
                    setTimeout(() => {
                        choicesInstances[id].setValue([]);
                        choicesInstances[id]._store.placeholder = true;
                        choicesInstances[id]._store.placeholderValue = 'All';
                        choicesInstances[id].input.element.placeholder = 'All';
                        choicesInstances[id].input.element.value = ''; // <-- Force input to empty for placeholder
                        choicesInstances[id].input.element.style.width = '100px'; // <-- Ensure enough width
                    }, 0);
                }
            });
            document.getElementById('dateStart').value = minDate;
            document.getElementById('dateEnd').value = maxDate;
            updateAllDropdownsExceptFactory();
            filterAndRender();
        }

        showLoadingSpinnerOverlay();
        fetch('/api/plan_data')
            .then(res => {
                if (!res.ok) throw new Error("API error: " + res.status);
                return res.json();
            })
            .then(data => {
                allData = data;
                factories = new Set(data.map(row => row.FACTORY || row.factory).filter(Boolean));
                assignStyleColors();
                initChoicesDropdowns();
                updateAllDropdowns(); // <-- update all dropdowns including factory
                updateDateInputs();
                filterAndRender();
                [
                    // Remove factorySelect from Choices.js event binding
                    'styleRefSelect','customerNameSelect','soNoDocSelect',
                    'productTypeSelect','subNoSelect','colorFcSelect','dateStart','dateEnd'
                ].forEach(id => {
                    document.getElementById(id).addEventListener('change', function() {
                        showLoadingSpinnerOverlay();
                        setTimeout(filterAndRender, 0);
                    });
                });
                // Add event for factorySelect (single select)
                document.getElementById('factorySelect').addEventListener('change', function() {
                    showLoadingSpinnerOverlay();
                    setTimeout(filterAndRender, 0);
                });
                document.getElementById('resetFiltersBtn').addEventListener('click', resetFilters);
                document.getElementById('toggleSamIeCalc').addEventListener('change', function() {
                    useSamIeCalc = this.checked;
                    showLoadingSpinnerOverlay();
                    setTimeout(filterAndRender, 0);
                });
                document.getElementById('toggleCellColors').addEventListener('change', function() {
                    showLoadingSpinnerOverlay();
                    setTimeout(filterAndRender, 0);
                });
                document.getElementById('toggleStyleColors').addEventListener('change', function() {
                    showLoadingSpinnerOverlay();
                    setTimeout(filterAndRender, 0);
                });
                // Add event for groupBySelect
                document.getElementById('groupBySelect').addEventListener('change', function() {
                    groupBy = this.value;
                    showLoadingSpinnerOverlay();
                    setTimeout(filterAndRender, 0);
                });
                hideLoadingSpinnerOverlay();
            })
            .catch(err => {
                document.getElementById('pivotTableContainer').innerHTML =
                    `<div style="color:red;">Error fetching data: ${err.message}</div>`;
                console.error("Fetch error:", err);
            });

        // Show spinner while loading
        showLoadingSpinnerOverlay();

        // Modal logic
        function openCellDetailModal(rows, line, date) {
            // --- Apply delay filter to modal rows if active ---
            const delayColorChecked = document.getElementById('toggleCellColors')?.checked;
            if (delayColorChecked && selectedDelayStatuses.length > 0) {
                function getStatus(delivery, sew) {
                    const dDate = new Date(delivery);
                    const sDate = new Date(sew);
                    if (isNaN(dDate) || isNaN(sDate)) return '';
                    const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                    if (diff < 0) return 'Delay';
                    if (diff < 4 && diff > 0) return 'Can be delay';
                    if (diff > 20) return 'Too early';
                    return 'Ok';
                }
                rows = rows.filter(r => selectedDelayStatuses.includes(getStatus(r.DELIVERY_DATE || r.delivery_date, r.SEW_DATE || r.sew_date)));
            }
            const modal = document.getElementById('cellDetailModal');
            const body = document.getElementById('cellDetailModalBody');
            function formatDateOnly(dt) {
                if (!dt) return '';
                const d = new Date(dt);
                if (!isNaN(d)) return d.toISOString().split('T')[0];
                if (typeof dt === 'string' && dt.includes(' ')) return dt.split(' ')[0];
                return dt;
            }
            function getStatus(delivery, sew) {
                const dDate = new Date(delivery);
                const sDate = new Date(sew);
                if (isNaN(dDate) || isNaN(sDate)) return '';
                const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                if (diff < 0) return 'Delay';
                if (diff < 4 && diff > 0) return 'Can be delay';
                if (diff > 20) return 'Too early';
                return 'Ok';
            }
            function getRowClass(status) {
                if (status === 'Delay') return 'modal-row-delay';
                if (status === 'Can be delay') return 'modal-row-can-delay';
                if (status === 'Too early') return 'modal-row-too-early';
                if (status === 'Ok') return 'modal-row-ok';
                return '';
            }
            // Legend HTML (color circle + text)
            const legendHtml = `
                <div style="position:absolute;top:8px;right:18px;display:flex;gap:14px;align-items:center;">
                    <span style="display:flex;align-items:center;gap:4px;">
                        <span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:#fee2e2;border:2px solid #f87171;"></span>
                        <span style="font-size:0.95em;">Delay</span>
                    </span>
                    <span style="display:flex;align-items:center;gap:4px;">
                        <span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:#fef9c3;border:2px solid #facc15;"></span>
                        <span style="font-size:0.95em;">Can be delay</span>
                    </span>
                    <span style="display:flex;align-items:center;gap:4px;">
                        <span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:#e0e7ff;border:2px solid #6366f1;"></span>
                        <span style="font-size:0.95em;">Too early</span>
                    </span>
                    <span style="display:flex;align-items:center;gap:4px;">
                        <span style="display:inline-block;width:14px;height:14px;border-radius:50%;background:#dcfce7;border:2px solid #4ade80;"></span>
                        <span style="font-size:0.95em;">Ok</span>
                    </span>
                </div>
            `;
            // Add modal row highlight styles (inject once)
            if (!document.getElementById('modal-row-highlight-style')) {
                const style = document.createElement('style');
                style.id = 'modal-row-highlight-style';
                style.innerHTML = `
                    .modal-row-delay { 
                        background: #fee2e2 !important; 
                        color: #111 !important; 
                        border: 2px solid #f87171 !important;
                    }
                    .modal-row-can-delay { 
                        background: #fef9c3 !important; 
                        color: #111 !important; 
                        border: 2px solid #facc15 !important;
                    }
                    .modal-row-too-early { 
                        background: #e0e7ff !important; 
                        color: #111 !important; 
                        border: 2px solid #6366f1 !important;
                    }
                    .modal-row-ok { 
                        background: #dcfce7 !important; 
                        color: #111 !important; 
                        border: 2px solid #4ade80 !important;
                    }
                    .modal-table td, .modal-table th { background: unset !important; }
                    .modal-table .extra-col, .modal-table .extra-col-header { display: none; }
                    .modal-table.show-extra-cols .extra-col, 
                    .modal-table.show-extra-cols .extra-col-header { display: table-cell; }
                `;
                document.head.appendChild(style);
            }
            // Add cell color styles for table cells (not just modal)
            if (!document.getElementById('cell-row-highlight-style')) {
                const style = document.createElement('style');
                style.id = 'cell-row-highlight-style';
                style.innerHTML = `
                    .cell-row-delay { background: #fee2e2 !important; color: #111 !important; }
                    .cell-row-can-delay { background: #fef9c3 !important; color: #111 !important; }
                    .cell-row-too-early { background: #e0e7ff !important; color: #111 !important; }
                    .cell-row-ok { background: #dcfce7 !important; color: #111 !important; }
                `;
                document.head.appendChild(style);
            }
            // Toggle icon HTML
            const toggleIconHtml = `
                <span id="modalExtraColsToggle" title="Show/hide extra columns" style="cursor:pointer;user-select:none;margin-left:8px;font-size:1.1em;vertical-align:middle;">
                    <svg width="18" height="18" style="vertical-align:middle;" viewBox="0 0 20 20"><polyline points="6,8 10,12 14,8" fill="none" stroke="#68A4DA" stroke-width="2"/></svg>
                </span>
            `;
            if (!rows || rows.length === 0) {
                body.innerHTML = '<div>No details found.</div>';
            } else {
                body.innerHTML = `
                    ${legendHtml}
                    <div style="font-weight:bold;margin-bottom:8px;">
                        Details for <span style="color:#68A4DA">${line}</span> on <span style="color:#68A4DA">${date}</span>
                    </div>
                    <div style="display:flex;align-items:stretch;gap:0;">
                        <div style="flex:1 1 auto;overflow-x:auto;">
                            <table class="modal-table" id="modalDetailTable" style="margin:0;">
                                <thead>
                                    <tr>
                                        <th>SO No</th>
                                        <th>Customer</th>
                                        <th>Style Ref</th>
                                        <th>Sub No</th>
                                        <th>Color FC</th>
                                        <th>SAM</th>
                                        <th>End SEW</th>
                                        <th>Delivery</th>
                                        <th>PCS</th>
                                        <th class="extra-col-header">EMB</th>
                                        <th class="extra-col-header">HEAT</th>
                                        <th class="extra-col-header">PAD</th>
                                        <th class="extra-col-header">PRINT</th>
                                        <th class="extra-col-header">BOND</th>
                                        <th class="extra-col-header">LASER</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    ${rows.map(r => {
                                        const sewDate = r.SEW_DATE || r.sew_date;
                                        const deliveryDate = r.DELIVERY_DATE || r.delivery_date;
                                        const status = getStatus(deliveryDate, sewDate);
                                        const rowClass = getRowClass(status);
                                        const planPcs = Number(r.PLAN_PCS || r.plan_pcs) || 0;
                                        function mult(val) {
                                            const n = Number(val);
                                            return n && planPcs ? formatNumberThousandSep(n * planPcs) : '';
                                        }
                                        const samIe = r.SAM_IE || r.sam_ie || '';
                                        return `
                                            <tr class="${rowClass}">
                                                <td>${r.SO_NO_DOC || r.so_no_doc || ''}</td>
                                                <td>${r.CUSTOMER_NAME || r.customer_name || ''}</td>
                                                <td>${r.STYLE_REF || r.style_ref || ''}</td>
                                                <td>${r.SUB_NO || r.sub_no || ''}</td>
                                                <td>${r.COLOR_FC || r.color_fc || ''}</td>
                                                <td>${samIe}</td>
                                                <td>${formatDateOnly(sewDate)}</td>
                                                <td>${formatDateOnly(deliveryDate)}</td>
                                                <td>${formatNumberThousandSep(planPcs)}</td>
                                                <td class="extra-col">${mult(r.EMBROIDERY || r.embroidery)}</td>
                                                <td class="extra-col">${mult(r.HEAT || r.heat)}</td>
                                                <td class="extra-col">${mult(r.PAD_PRINT || r.pad_print)}</td>
                                                <td class="extra-col">${mult(r.PRINT || r.print)}</td>
                                                <td class="extra-col">${mult(r.BOND || r.bond)}</td>
                                                <td class="extra-col">${mult(r.LASER || r.laser)}</td>
                                            </tr>
                                        `;
                                    }).join('')}
                                </tbody>
                            </table>
                        </div>
                        <div id="modalExtraColsToggleContainer"
                            style="flex:0 0 32px;min-width:32px;width:32px;height:42px;display:flex;align-items:center;justify-content:center;background:#e5e7eb;border-radius:0 6px 6px 0;box-shadow:0 0 2px #bbb;">
                            <span id="modalExtraColsToggle" title="Show/hide extra columns"
                                style="cursor:pointer;user-select:none;display:flex;align-items:center;justify-content:center;width:100%;height:100%;">
                                <svg width="18" height="18" style="vertical-align:middle;" viewBox="0 0 20 20">
                                    <polyline points="8,6 12,10 8,14" fill="none" stroke="#68A4DA" stroke-width="2"/>
                                </svg>
                            </span>
                        </div>
                    </div>
                `;
            }
            modal.style.display = 'flex';
            // Add expand/collapse logic
            setTimeout(() => {
                const toggle = document.getElementById('modalExtraColsToggle');
                const table = document.getElementById('modalDetailTable');
                let expanded = false;
                if (toggle && table) {
                    table.classList.remove('show-extra-cols');
                    toggle.onclick = function() {
                        expanded = !expanded;
                        if (expanded) {
                            table.classList.add('show-extra-cols');
                            toggle.innerHTML = `<svg width="18" height="18" style="vertical-align:middle;" viewBox="0 0 20 20"><polyline points="12,6 8,10 12,14" fill="none" stroke="#68A4DA" stroke-width="2"/></svg>`;
                        } else {
                            table.classList.remove('show-extra-cols');
                            toggle.innerHTML = `<svg width="18" height="18" style="vertical-align:middle;" viewBox="0 0 20 20"><polyline points="8,6 12,10 8,14" fill="none" stroke="#68A4DA" stroke-width="2"/></svg>`;
                        }
                    };
                }
            }, 0);
        }
        function closeCellDetailModal() {
            document.getElementById('cellDetailModal').style.display = 'none';
        }
        window.closeCellDetailModal = closeCellDetailModal;

        // Add click outside modal-content to close modal
        document.addEventListener('DOMContentLoaded', function() {
            const modal = document.getElementById('cellDetailModal');
            if (modal) {
                modal.addEventListener('click', function(e) {
                    if (e.target === modal) {
                        closeCellDetailModal();
                    }
                });
            }
        });

        // --- Add SO No Image Modal ---
// Update modal size and content to show style names
if (!document.getElementById('soNoImageModal')) {
    const modal = document.createElement('div');
    modal.id = 'soNoImageModal';
    modal.className = 'modal-overlay';
    modal.style.zIndex = '10001';
    modal.innerHTML = `
        <div class="modal-content" style="min-width:50vw;max-width:80vw;min-height:50vh;max-height:80vh;overflow-y:auto;">
            <div style="font-weight:bold;font-size:1.3em;margin-bottom:18px;">SO No Details</div>
            <div id="soNoImageModalBody"></div>
            <div style="margin-top:16px;display:flex;justify-content:flex-end;gap:12px;">
                <button type="button" id="soNoImageModalCloseBtn" style="background:#eee;color:#333;border:none;border-radius:5px;padding:8px 24px;font-size:1.1em;">Close</button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
    document.getElementById('soNoImageModalCloseBtn').onclick = function() {
        modal.style.display = 'none';
    };
    modal.onclick = function(e) {
        if (e.target === modal) modal.style.display = 'none';
    };
}

// Helper to format SO No for image link
function formatSoNoForImage(soNo) {
    let digits = String(soNo).replace(/\D/g, '');
    if (digits.length > 6) digits = digits.slice(-6);
    let first = digits.slice(0, 2).replace(/^0+/, '');
    let last = digits.slice(-4).replace(/^0+/, '');
    return `${first}-${last}`;
}

// Show SO No Image Modal (now accepts matchingRows for style names)
function showSoNoImageModal(soNos, matchingRows) {
    const modal = document.getElementById('soNoImageModal');
    const body = document.getElementById('soNoImageModalBody');
    if (!soNos || soNos.length === 0) {
        body.innerHTML = '<div>No SO No found.</div>';
    } else {
        // Calculate total pieces for each SO No
        const soNoPcsMap = {};
        soNos.forEach(soNo => {
            soNoPcsMap[soNo] = matchingRows
                .filter(r => (r.SO_NO_DOC || r.so_no_doc) === soNo)
                .reduce((sum, r) => sum + (Number(r.PLAN_PCS || r.plan_pcs) || 0), 0);
        });

        // Sort SO Nos by total pieces
        const sortedSoNos = soNos.sort((a, b) => soNoPcsMap[b] - soNoPcsMap[a]);

        body.innerHTML = sortedSoNos.map(soNo => {
            // Find all styles for this SO No in matchingRows and calculate their total PCS
            const stylesPcs = {};
            matchingRows
                .filter(r => (r.SO_NO_DOC || r.so_no_doc) === soNo)
                .forEach(r => {
                    const style = r.STYLE_REF || r.style_ref;
                    if (!style) return;
                    stylesPcs[style] = (stylesPcs[style] || 0) + (Number(r.PLAN_PCS || r.plan_pcs) || 0);
                });
            
            // Sort styles by PCS in descending order
            const styles = Object.entries(stylesPcs)
                .sort((a, b) => b[1] - a[1])
                .map(([style]) => style);

            const styleHtml = styles.length
                ? `<div style="color:#2563eb;font-size:1.08em;margin-bottom:4px;">${styles.join(', ')}</div>`
                : '';
            const imgCode = formatSoNoForImage(soNo);
            const imgUrl = `https://ndes-back.nanyangtextile.com/proxy-image/${imgCode}`;
            const fallbackImageUrl = '/static/image-not-found.png'; 
            return `
                <div style="margin-bottom:28px;display:flex;align-items:center;gap:28px;">
                    <div style="min-width:140px;">
                        <div style="font-weight:bold;font-size:1.15em;">${soNo}</div>
                        ${styleHtml}
                    </div>
                    <img src="${imgUrl}" alt="${soNo}" style="max-width:1800px;max-height:1800px;border:1px solid #ccc;border-radius:8px;"
                    onerror="this.onerror=null;this.src='${fallbackImageUrl}';">
                </div>
            `;
        }).join('');
    }
    modal.style.display = 'flex';
}

// --- Patch addPivotTableCellClickHandlers ---
function addPivotTableCellClickHandlers(filteredData) {
    const table = document.getElementById('pivotTable');
    if (!table) return;

    Array.from(table.querySelectorAll('tbody tr')).forEach((rowTr, rowIdx) => {
        // --- Fix: support expanded line+style mode ---
        let line, style, factory;
        if (expandLineStyle) {
            // Expect cell: "Line / Style"
            const td = rowTr.querySelector('td');
            if (td) {
                const txt = td.textContent || '';
                const parts = txt.split('/');
                line = parts[0].trim();
                style = (parts[1] || '').trim();
            }
        } else {
            // Collapsed: support "Factory - Line" or just "Line"
            const td = rowTr.querySelector('td');
            if (td) {
                const txt = td.textContent || '';
                if (txt.includes(' - ')) {
                    // "Factory - Line"
                    [factory, line] = txt.split(' - ').map(s => s.trim());
                } else {
                    factory = null;
                    line = txt.trim();
                }
            }
            style = null;
        }

        Array.from(rowTr.querySelectorAll('td.date-col')).forEach((cellTd, colIdx) => {
            let groupKey = cellTd.getAttribute('data-full-date');
            cellTd.style.cursor = 'pointer';
            cellTd.onclick = function() {
                const styleColorChecked = document.getElementById('toggleStyleColors')?.checked;
                if (styleColorChecked) {
                    let matchingRows;
                    if (expandLineStyle) {
                        // Expanded: group by line+style
                        const td = rowTr.querySelector('td');
                        let line = '', style = '';
                        if (td) {
                            const txt = td.textContent || '';
                            const parts = txt.split('/');
                            line = parts[0].trim();
                            style = (parts[1] || '').trim();
                        }
                        matchingRows = filteredData.filter(r => {
                            const rLine = r.PROD_LINE || r.prod_line || 'Unknown';
                            const rStyle = r.STYLE_REF || r.style_ref || 'Unknown';
                            const dateRaw = r.SEW_DATE || r.sew_date || '';
                            if (!dateRaw) return false;
                            const dateObj = new Date(dateRaw);
                            if (isNaN(dateObj)) return false;
                            let rGroupKey;
                            if (groupBy === 'day') rGroupKey = dateObj.toISOString().split('T')[0];
                            else if (groupBy === 'week') rGroupKey = getISOWeekString(dateObj);
                            else if (groupBy === 'month') rGroupKey = getMonthString(dateObj);
                            return rLine === line && rStyle === style && rGroupKey === groupKey;
                        });
                    } else {
                        // Collapsed: group by (factory-line) if multi-factory, else by line
                        const td = rowTr.querySelector('td');
                        let factory = null, line = '';
                        if (td) {
                            const txt = td.textContent || '';
                            if (txt.includes(' - ')) {
                                // "Factory - Line"
                                [factory, line] = txt.split(' - ').map(s => s.trim());
                            } else {
                                factory = null;
                                line = txt.trim();
                            }
                        }
                        matchingRows = filteredData.filter(r => {
                            const rFactory = r.FACTORY || r.factory || 'Unknown';
                            const rLine = r.PROD_LINE || r.prod_line || 'Unknown';
                            const dateRaw = r.SEW_DATE || r.sew_date || '';
                            if (!dateRaw) return false;
                            const dateObj = new Date(dateRaw);
                            if (isNaN(dateObj)) return false;
                            let rGroupKey;
                            if (groupBy === 'day') rGroupKey = dateObj.toISOString().split('T')[0];
                            else if (groupBy === 'week') rGroupKey = getISOWeekString(dateObj);
                            else if (groupBy === 'month') rGroupKey = getMonthString(dateObj);
                            if (factory) {
                                // Grouped by "Factory - Line"
                                return rFactory === factory && rLine === line && rGroupKey === groupKey;
                            } else {
                                // Grouped by line only
                                return rLine === line && rGroupKey === groupKey;
                            }
                        });
                    }
                    // Get distinct SO No
                    const soNos = Array.from(new Set(matchingRows.map(r => r.SO_NO_DOC || r.so_no_doc).filter(Boolean)));
                    showSoNoImageModal(soNos, matchingRows); // Pass matchingRows for style names
                    return;
                }

                let matchingRows;
                if (expandLineStyle) {
                    matchingRows = filteredData.filter(r => {
                        const rLine = r.PROD_LINE || r.prod_line || 'Unknown';
                        const rStyle = r.STYLE_REF || r.style_ref || 'Unknown';
                        const dateRaw = r.SEW_DATE || r.sew_date || '';
                        if (!dateRaw) return false;
                        const dateObj = new Date(dateRaw);
                        if (isNaN(dateObj)) return false;
                        let rGroupKey;
                        if (groupBy === 'day') rGroupKey = dateObj.toISOString().split('T')[0];
                        else if (groupBy === 'week') rGroupKey = getISOWeekString(dateObj);
                        else if (groupBy === 'month') rGroupKey = getMonthString(dateObj);
                        return rLine === line && rStyle === style && rGroupKey === groupKey;
                    });
                    openCellDetailModal(matchingRows, `${line} / ${style}`, groupKey);
                } else {
                    matchingRows = filteredData.filter(r => {
                        const rFactory = r.FACTORY || r.factory || 'Unknown';
                        const rLine = r.PROD_LINE || r.prod_line || 'Unknown';
                        const dateRaw = r.SEW_DATE || r.sew_date || '';
                        if (!dateRaw) return false;
                        const dateObj = new Date(dateRaw);
                        if (isNaN(dateObj)) return false;
                        let rGroupKey;
                        if (groupBy === 'day') rGroupKey = dateObj.toISOString().split('T')[0];
                        else if (groupBy === 'week') rGroupKey = getISOWeekString(dateObj);
                        else if (groupBy === 'month') rGroupKey = getMonthString(dateObj);
                        if (factory) {
                            // Grouped by "Factory - Line"
                            return rFactory === factory && rLine === line && rGroupKey === groupKey;
                        } else {
                            // Grouped by line only
                            return rLine === line && rGroupKey === groupKey;
                        }
                    });
                    // Show "Factory - Line" or just "Line" in modal
                    openCellDetailModal(matchingRows, factory ? `${factory} - ${line}` : line, groupKey);
                }
            };
        });
    });
}

// --- View switch logic ---
document.addEventListener('DOMContentLoaded', function() {
    const switchBtn = document.getElementById('switchViewBtn');
    const switchIcon = document.getElementById('switchViewIcon');
    if (switchBtn && switchIcon) {
        switchBtn.addEventListener('click', function() {
            if (currentView === 'table') {
                document.getElementById('pivotTableScroll').style.display = 'none';
                document.getElementById('graphContainer').style.display = '';
                switchIcon.innerHTML = '&#8592;'; // left arrow
                switchIcon.style.transform = 'rotate(180deg)';
                currentView = 'graph';
                renderPlanPcsGraph(window._lastFilteredData || []);
            } else {
                document.getElementById('pivotTableScroll').style.display = '';
                document.getElementById('graphContainer').style.display = 'none';
                switchIcon.innerHTML = '&#8594;'; // right arrow
                switchIcon.style.transform = 'rotate(0deg)';
                currentView = 'table';
            }
        });
    }
});

// --- View toggle logic ---
document.addEventListener('DOMContentLoaded', function() {
    // Toggle switch logic
    const toggleView = document.getElementById('toggleView');
    const toggleViewLabel = document.getElementById('toggleViewLabel');
    const styleBox = document.getElementById('toggleStyleColors');
    const toggleSubprocess = document.getElementById('toggleSubprocess');
    const togglePoints = document.getElementById('togglePoints');
    toggleView.addEventListener('change', function() {
        if (toggleView.checked) {
            document.getElementById('pivotTableScroll').style.display = 'none';
            document.getElementById('graphContainer').style.display = '';
            toggleViewLabel.textContent = 'Graph';
            currentView = 'graph';
            styleBox.disabled = true;
            toggleSubprocess.disabled = true;
            togglePoints.disabled = true;
            renderPlanPcsGraph(window._lastFilteredData || []);
        } else {
            document.getElementById('pivotTableScroll').style.display = '';
            document.getElementById('graphContainer').style.display = 'none';
            toggleViewLabel.textContent = 'Table';
            currentView = 'table';
            styleBox.disabled = false;
            toggleSubprocess.disabled = false;
            togglePoints.disabled = false;
        }
    });
});

// --- Graph rendering ---
function renderPlanPcsGraph(filteredData) {
    const delayBox = document.getElementById('toggleCellColors');
    const showDelay = delayBox?.checked && currentView === 'graph';
    // Add: subprocess logic
    const showSubprocessGraph = showSubprocess && selectedSubprocesses.length > 0;

    const groupKeySet = new Set();
    const aggMap = {}; // { groupKey: sum }
    const statusMap = {}; // { groupKey: { Delay: count, CanBeDelay: count, TooEarly: count, Ok: count, total: count } }

    filteredData.forEach(row => {
        const dateRaw = row.SEW_DATE || row.sew_date || '';
        if (!dateRaw) return;
        const dateObj = new Date(dateRaw);
        if (isNaN(dateObj)) return;
        let groupKey;
        if (groupBy === 'day') {
            groupKey = dateObj.toISOString().split('T')[0];
        } else if (groupBy === 'week') {
            groupKey = getISOWeekString(dateObj);
        } else if (groupBy === 'month') {
            groupKey = getMonthString(dateObj);
        }
        groupKeySet.add(groupKey);
        if (!aggMap[groupKey]) aggMap[groupKey] = 0;
        if (!statusMap[groupKey]) statusMap[groupKey] = { Delay: 0, CanBeDelay: 0, TooEarly: 0, Ok:  0, total: 0 };

        if (showSubprocessGraph) {
            aggMap[groupKey] += sumSubprocess(row, selectedSubrocesses, usePoints);
        } else {
            let pcs = Number(row.PLAN_PCS || row.plan_pcs) || 0;
            let sam = Number(row.SAM_IE || row.sam_ie) || 0;
            aggMap[groupKey] += useSamIeCalc ? pcs * sam : pcs;
        }

        // Status counting for delay graph
        if (showDelay) {
            const delivery = row.DELIVERY_DATE || row.delivery_date;
            const sew = row.SEW_DATE || r.sew_date;
            const dDate = new Date(delivery);
            const sDate = new Date(sew);
            let status = '';
            if (!isNaN(dDate) && !isNaN(sDate)) {
                const diff = Math.floor((dDate - sDate) / (1000 * 60 * 60 * 24));
                if (diff < 0) status = 'Delay';
                else if (diff < 4 && diff > 0) status = 'CanBeDelay';
                else if (diff > 20) status = 'TooEarly';
                else status = 'Ok';
            }
            statusMap[groupKey][status] = (statusMap[groupKey][status] || 0) + (showSubprocessGraph ? sumSubprocess(row, selectedSubprocesses, usePoints) : (Number(row.PLAN_PCS || row.plan_pcs) || 0));
            statusMap[groupKey].total += (showSubprocessGraph ? sumSubprocess(row, selectedSubprocesses, usePoints) : (Number(row.PLAN_PCS || row.plan_pcs) || 0));
        }
    });

    let groupKeys = Array.from(groupKeySet);
    if (groupBy === 'day') {
        groupKeys.sort((a, b) => new Date(a) - new Date(b));
    } else {
        groupKeys.sort();
    }

    const labels = groupKeys;

    let datasets;
    if (showDelay) {
        // Stacked bar: percentage of each status
        const statusOrder = ['Delay', 'CanBeDelay', 'TooEarly', 'Ok'];
        const statusLabels = {
            Delay: 'Delay',
            CanBeDelay: 'Can be delay',
            TooEarly: 'Too early',
            Ok: 'Ok'
        };
        // Use darker colors for graph bars
        const statusColors = {
            Delay: '#f87171',        // red-400
            CanBeDelay: '#facc15',   // yellow-400
            TooEarly: '#6366f1',     // indigo-500
            Ok: '#22c55e'            // green-500
        };
        datasets = statusOrder.map(status => ({
            label: statusLabels[status],
            data: groupKeys.map(k => {
                const total = statusMap[k]?.total || 0;
                const val = statusMap[k]?.[status] || 0;
                return total ? (val / total * 100) : 0;
            }),
            backgroundColor: statusColors[status],
            stack: 'status'
        }));
    } else {
        datasets = [{
            label: showSubprocessGraph
                ? (usePoints
                    ? 'Sum of points (' + selectedSubprocesses.map(sp => {
                        if (sp === 'EMBROIDERY') return 'EMB';
                        if (sp === 'HEAT') return 'HEAT';
                        if (sp === 'PAD_PRINT') return 'PAD';
                        if (sp === 'PRINT') return 'PRINT';
                        if (sp === 'BOND') return 'BOND';
                        if (sp === 'LASER') return 'LASER';
                        return sp;
                    }).join(', ') + ')'
                    : 'Sum of ' + selectedSubprocesses.map(sp => {
                        if (sp === 'EMBROIDERY') return 'EMB';
                        if (sp === 'HEAT') return 'HEAT';
                        if (sp === 'PAD_PRINT') return 'PAD';
                        if (sp === 'PRINT') return 'PRINT';
                        if (sp === 'BOND') return 'BOND';
                        if (sp === 'LASER') return 'LASER';
                        return sp;
                    }).join(', '))
                : (useSamIeCalc ? 'Sum of time minutes' : 'Sum of pieces'),
            data: groupKeys.map(k => aggMap[k] || 0),
            backgroundColor: '#7fb4e3',
            maxBarThickness: 18
        }];
    }

    // Destroy previous chart if exists
    if (planPcsChart) {
        planPcsChart.destroy();
    }
    const ctx = document.getElementById('planPcsChart').getContext('2d');
    planPcsChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true },
                title: {
                    display: true,
                    text: showDelay
                        ? 'Delay status % by ' + (groupBy === 'day' ? 'Day' : groupBy === 'week' ? 'Week' : 'Month')
                        : (
                            showSubprocessGraph
                                ? (usePoints
                                    ? 'Sum of points'
                                    : 'Sum of subprocess')
                                : (useSamIeCalc ? 'Sum of time minutes' : 'Sum of pieces')
                        ) + ' by ' + (groupBy === 'day' ? 'Day' : groupBy === 'week' ? 'Week' : 'Month'),
                    font: { size: 15 }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            if (showDelay) {
                                // Show percentage and absolute value
                                const groupKey = context.label;
                                const status = context.dataset.label;
                                const percent = context.parsed.y;
                                const total = statusMap[groupKey]?.total || 0;
                                const statusOrder = { 'Delay': 'Delay', 'Can be delay': 'CanBeDelay', 'Too early': 'TooEarly', 'Ok': 'Ok' };
                                const statusKey = statusOrder[status] || status;
                                const abs = statusMap[groupKey]?.[statusKey] || 0;

                                return `${status}: ${percent.toFixed(1)}% (${abs})`;
                            } else {
                                let val = context.parsed.y;
                                if (typeof val === 'number' && !Number.isInteger(val)) {
                                    val = val.toFixed(2);
                                }
                                if (showSubprocessGraph && usePoints) {
                                    return 'Points: ' + val;
                                } else if (showSubprocessGraph) {
                                    return 'Subprocess sum: ' + val;
                                } else if (useSamIeCalc) {
                                    return 'Time minutes: ' + val;
                                } else {
                                    return 'Pieces: ' + val;
                                }
                            }
                        }
                    }
                }
            },
            layout: { padding: { left: 8, right: 8, top: 8, bottom: 8 } },
            scales: showDelay
                ? {
                    x: {
                        title: { display: true, text: groupBy === 'day' ? 'Date' : groupBy === 'week' ? 'Week' : 'Month' },
                        stacked: true,
                        ticks: { autoSkip: true, maxTicksLimit: 20, font: { size: 11 } }
                    },
                    y: {
                        title: { display: true, text: 'Status %' },
                        beginAtZero: true,
                        max: 100,
                        stacked: true,
                        ticks: { font: { size: 11 }, callback: v => v + '%' }
                    }
                }
                : {
                    x: {
                        title: { display: true, text: groupBy === 'day' ? 'Date' : groupBy === 'week' ? 'Week' : 'Month' },
                        ticks: { autoSkip: true, maxTicksLimit: 20, font: { size: 11 } }
                    },
                    y: {
                        title: { display: true, text: useSamIeCalc ? 'Time minutes' : 'Pieces' },
                        beginAtZero: true,
                        ticks: { font: { size: 11 } }
                    }
                }
        },
        });
    }

    // Ensure graph updates when useSamIeCalc changes
    document.addEventListener('DOMContentLoaded', function() {
        document.getElementById('toggleSamIeCalc').addEventListener('change', function() {
            useSamIeCalc = this.checked;
            showLoadingSpinnerOverlay();
            setTimeout(function() {
                filterAndRender();
                // If graph is visible, update it immediately
                if (currentView === 'graph') {
                    renderPlanPcsGraph(window._lastFilteredData || []);
                }
            }, 0);
        });
    });

    // --- Checkbox mutual exclusivity (uncheck, not disable) ---
    document.addEventListener('DOMContentLoaded', function() {
        const delayBox = document.getElementById('toggleCellColors');
        const styleBox = document.getElementById('toggleStyleColors');
        const toggleSubprocess = document.getElementById('toggleSubprocess');
        const togglePoints = document.getElementById('togglePoints');
        const toggleSamIeCalc = document.getElementById('toggleSamIeCalc');
        // Mutual exclusivity: style <-> minutes <-> subprocess
        toggleSamIeCalc.addEventListener('change', function() {
            useSamIeCalc = this.checked;
            if (toggleSamIeCalc.checked) {
                styleBox.checked = false;
                // Deselect subprocess if minutes is checked
                toggleSubprocess.checked = false;
                showSubprocess = false;
                togglePoints.disabled = true;
            }
            showLoadingSpinnerOverlay();
            setTimeout(function() {
                filterAndRender();
                if (currentView === 'graph') {
                    renderPlanPcsGraph(window._lastFilteredData || []);
                }
            }, 0);
        });
        styleBox.addEventListener('change', function() {
            if (styleBox.checked) {
                toggleSamIeCalc.checked = false;
                useSamIeCalc = false;
                // Deselect subprocess if style is checked
                toggleSubprocess.checked = false;
                showSubprocess = false;
                togglePoints.disabled = true;
            }
            delayBox.checked = false;
            document.getElementById('subprocessSelect').style.display = 'none';
            document.getElementById('toggleSamIeCalc').disabled = false;
            delayBox.disabled = false;
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        });
        delayBox.addEventListener('change', function() {
            // Only uncheck style, not subprocess
            if (delayBox.checked) {
                styleBox.checked = false;
                styleBox.disabled = false;
            }
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        });
        toggleSubprocess.addEventListener('change', function() {
            // When subprocess is checked, deselect style and minutes, but do not disable them
            if (toggleSubprocess.checked) {
                styleBox.checked = false;
                toggleSamIeCalc.checked = false;
                useSamIeCalc = false;
                showSubprocess = true;
            } else {
                showSubprocess = false;
            }
            // Enable/disable points checkbox only
            togglePoints.disabled = !toggleSubprocess.checked;
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        });
        togglePoints.addEventListener('change', function() {
            usePoints = this.checked;
            showLoadingSpinnerOverlay();
            setTimeout(function() {
                filterAndRender();
                if (currentView === 'graph') {
                    renderPlanPcsGraph(window._lastFilteredData || []);
                }
            }, 0);
        });
    });

    // --- Show by subprocess dropdown/modal logic ---
    document.addEventListener('DOMContentLoaded', function() {
        const toggleSubprocess = document.getElementById('toggleSubprocess');
        const togglePoints = document.getElementById('togglePoints');
        // Only disable "Show by time minutes" and "Show by style"
        toggleSubprocess.addEventListener('change', function() {
            showSubprocess = toggleSubprocess.checked;
            // Only enable points if exactly one SO No Doc is selected
            const soNoDocSelected = getSelectedValues('soNoDocSelect');
            togglePoints.disabled = !(showSubprocess && soNoDocSelected.length === 1);
            if (!showSubprocess) {
                togglePoints.checked = false;
                usePoints = false;
            }
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        });
        // Modal for subprocess selection
        const modalId = 'subprocessModal';
        if (!document.getElementById(modalId)) {
            const modal = document.createElement('div');
            modal.id = modalId;
            modal.className = 'modal-overlay';
            modal.style.zIndex = '10001';
            modal.innerHTML = `
                <div class="modal-content" style="min-width:320px;max-width:96vw;">
                    <div style="font-weight:bold;font-size:1.1em;margin-bottom:8px;">Select subprocess columns to sum</div>
                    <form id="subprocessForm" style="display:flex;flex-direction:column;gap:8px;">
                        <!-- Removed PCS option -->
                        <label><input type="checkbox" value="EMBROIDERY"> EMB</label>
                        <label><input type="checkbox" value="HEAT"> HEAT</label>
                        <label><input type="checkbox" value="PAD_PRINT"> PAD</label>
                        <label><input type="checkbox" value="PRINT"> PRINT</label>
                        <label><input type="checkbox" value="BOND"> BOND</label>
                        <label><input type="checkbox" value="LASER"> LASER</label>
                    </form>
                    <div style="margin-top:16px;display:flex;justify-content:flex-end;gap:12px;">
                        <button type="button" id="subprocessModalOkBtn" style="background:#219a0b;color:#fff;border:none;border-radius:5px;padding:6px 18px;font-weight:600;">OK</button>
                        <button type="button" id="subprocessModalCancelBtn" style="background:#eee;color:#333;border:none;border-radius:5px;padding:6px 18px;">Cancel</button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);
        }

        // Open modal when clicking label text
        document.getElementById('showBySubprocessLabel').addEventListener('click', function(e) {
            e.preventDefault();
            const modal = document.getElementById(modalId);
            const form = document.getElementById('subprocessForm');
            // Set checked boxes based on selectedSubprocesses
            Array.from(form.elements).forEach(el => {
                if (el.type === 'checkbox') {
                    el.checked = selectedSubprocesses.includes(el.value);
                }
            });
            modal.style.display = 'flex';
        });

        // Modal OK/Cancel logic
        document.getElementById('subprocessModalOkBtn').addEventListener('click', function() {
            const form = document.getElementById('subprocessForm');
            selectedSubprocesses = Array.from(form.elements)
                .filter(el => el.type === 'checkbox' && el.checked)
                .map(el => el.value);
            // If any selected, check the main box and show dropdown
            if (selectedSubprocesses.length > 0) {
                document.getElementById('togglePoints').disabled = false;
            } else {
                // If none selected, disable "Show by points"
                document.getElementById('togglePoints').disabled = true;
            }
            showSubprocess = toggleSubprocess.checked;
            document.getElementById(modalId).style.display = 'none';
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        });
        document.getElementById('subprocessModalCancelBtn').addEventListener('click', function() {
            document.getElementById(modalId).style.display = 'none';
        });
        // Close modal on click outside
        document.getElementById(modalId).addEventListener('click', function(e) {
            if (e.target === this) this.style.display = 'none';
        });
    });

    // --- Delay Status Modal ---
    const delayModalId = 'delayStatusModal';
    if (!document.getElementById(delayModalId)) {
        const modal = document.createElement('div');
        modal.id = delayModalId;
        modal.className = 'modal-overlay';
        modal.style.zIndex = '10001';
        modal.innerHTML = `
            <div class="modal-content" style="min-width:320px;max-width:96vw;">
                <div style="font-weight:bold;font-size:1.1em;margin-bottom:8px;">Select delay statuses to show</div>
                <form id="delayStatusForm" style="display:flex;flex-direction:column;gap:8px;">
                    <label><input type="checkbox" value="Delay"> Delay</label>
                    <label><input type="checkbox" value="Can be delay"> Can be delay</label>
                    <label><input type="checkbox" value="Too early"> Too early</label>
                    <label><input type="checkbox" value="Ok"> Ok</label>
                </form>
                <div style="margin-top:16px;display:flex;justify-content:flex-end;gap:12px;">
                    <button type="button" id="delayModalOkBtn" style="background:#219a0b;color:#fff;border:none;border-radius:5px;padding:6px 18px;font-weight:600;">OK</button>
                    <button type="button" id="delayModalCancelBtn" style="background:#eee;color:#333;border:none;border-radius:5px;padding:6px 18px;">Cancel</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }

    // Open modal when clicking "By Delay" label
    const delayLabel = document.querySelector('label[for="toggleCellColors"]');
    if (delayLabel) {
        delayLabel.style.cursor = 'pointer';
        // Add hover effect
        delayLabel.style.transition = 'background 0.15s, color 0.15s';
        delayLabel.addEventListener('mouseenter', function() {
            delayLabel.style.background = '#e0e7ff';
            delayLabel.style.color = '#2563eb';
        });
        delayLabel.addEventListener('mouseleave', function() {
            delayLabel.style.background = '';
            delayLabel.style.color = '';
        });
        delayLabel.addEventListener('click', function(e) {
            e.preventDefault();
            const modal = document.getElementById(delayModalId);
            const form = document.getElementById('delayStatusForm');
            // Set checked boxes based on selectedDelayStatuses
            Array.from(form.elements).forEach(el => {
                if (el.type === 'checkbox') {
                    el.checked = selectedDelayStatuses.includes(el.value);
                }
            });
            modal.style.display = 'flex';
        });
    }

    // OK/Cancel logic
    document.getElementById('delayModalOkBtn').onclick = function() {
        const form = document.getElementById('delayStatusForm');
        selectedDelayStatuses = Array.from(form.elements)
            .filter(el => el.type === 'checkbox' && el.checked)
            .map(el => el.value);
        document.getElementById(delayModalId).style.display = 'none';
        showLoadingSpinnerOverlay();
        setTimeout(filterAndRender, 0);
    };
    document.getElementById('delayModalCancelBtn').onclick = function() {
        document.getElementById(delayModalId).style.display = 'none';
    };
    document.getElementById(delayModalId).addEventListener('click', function(e) {
        if (e.target === this) this.style.display = 'none';
    });

    // --- SO No Doc Type Modal ---
    if (!document.getElementById('soNoDocTypeModal')) {
        const modal = document.createElement('div');
        modal.id = 'soNoDocTypeModal';
        modal.className = 'modal-overlay';
        modal.style.zIndex = '10001';
        modal.innerHTML = `
            <div class="modal-content" style="min-width:320px;max-width:96vw;">
                <div style="font-weight:bold;font-size:1.1em;margin-bottom:8px;">Select SO No Doc type(s)</div>
                <form id="soNoDocTypeForm" style="display:flex;flex-direction:column;gap:8px;">
                    <label><input type="checkbox" value="Bulk"> Bulk</label>
                    <label><input type="checkbox" value="Sample"> Sample</label>
                    <label><input type="checkbox" value="Early"> Early buy-stock</label>
                </form>
                <div style="margin-top:16px;display:flex;justify-content:flex-end;gap:12px;">
                    <button type="button" id="soNoDocTypeModalOkBtn" style="background:#219a0b;color:#fff;border:none;border-radius:5px;padding:6px 18px;font-weight:600;">OK</button>
                    <button type="button" id="soNoDocTypeModalCancelBtn" style="background:#eee;color:#333;border:none;border-radius:5px;padding:6px 18px;">Cancel</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    }

    // Show modal when clicking SO No Doc label
    document.addEventListener('DOMContentLoaded', function() {
        const soNoDocLabel = document.querySelector('.filter-group label[for="soNoDocSelect"]');
        if (soNoDocLabel) {
            soNoDocLabel.style.cursor = 'pointer';
            soNoDocLabel.addEventListener('click', function(e) {
                e.preventDefault();
                const modal = document.getElementById('soNoDocTypeModal');
                const form = document.getElementById('soNoDocTypeForm');
                // Set checked boxes based on window.soNoDocTypeSelections, default to all if none selected
                let sel = window.soNoDocTypeSelections && window.soNoDocTypeSelections.length > 0 
                    ? window.soNoDocTypeSelections 
                    : ['Bulk', 'Sample', 'Early'];
                Array.from(form.elements).forEach(el => {
                    if (el.type === 'checkbox') {
                        el.checked = sel.includes(el.value);
                    }
                });
                modal.style.display = 'flex';
            });
        }
        // Modal OK/Cancel logic
        document.getElementById('soNoDocTypeModalOkBtn').onclick = function() {
            const form = document.getElementById('soNoDocTypeForm');
            window.soNoDocTypeSelections = Array.from(form.elements)
                .filter(el => el.type === 'checkbox' && el.checked)
                .map(el => el.value);
            document.getElementById('soNoDocTypeModal').style.display = 'none';
            showLoadingSpinnerOverlay();
            setTimeout(filterAndRender, 0);
        };
        document.getElementById('soNoDocTypeModalCancelBtn').onclick = function() {
            document.getElementById('soNoDocTypeModal').style.display = 'none';
        };
        document.getElementById('soNoDocTypeModal').addEventListener('click', function(e) {
            if (e.target === this) this.style.display = 'none';
        });
    });