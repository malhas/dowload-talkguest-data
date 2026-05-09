// ==UserScript==
// @name         Talkguest Table Downloader
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Download reservations table to CSV including tooltip details
// @author       You
// @match        https://owner.talkguest.com/Bookings/OwnerArea.aspx*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Helper to escape CSV values
    function escapeCSV(val) {
        if (val === null || val === undefined) return '""';
        let str = String(val).replace(/"/g, '""').trim();
        return `"${str}"`;
    }

    async function downloadCSV() {
        const table = document.querySelector('table.TableRecords');
        if (!table) {
            alert('Table not found!');
            return;
        }

        const btn = document.getElementById('talkguest-csv-btn');
        const originalText = btn.innerText;
        btn.innerText = 'Loading details...';
        btn.disabled = true;

        let csvContent = "";

        // Headers
        const headers = Array.from(table.querySelectorAll('thead th'));
        const headerNames = headers.map(th => th.innerText.trim());
        headerNames.push('IVA (6%)'); // Add IVA column

        const allRowsData = [];
        const extraColumnNames = new Set();

        // Rows
        const rows = Array.from(table.querySelectorAll('tbody > tr'));
        for (const row of rows) {
            const cells = Array.from(row.querySelectorAll('td'));

            // Skip rows that don't match the expected column count (e.g. hidden or nested tables)
            if (cells.length < headers.length) continue;

            const rowData = [];
            const rowInfo = { baseData: rowData, extraData: {} };

            for (let index = 0; index < cells.length; index++) {
                const cell = cells[index];

                // Total column is usually the last one (index 7) or has the tooltip container
                if (index === headers.length - 1 || cell.querySelector('.tooltip-container')) {
                    // Extract visible total
                    const visibleTotalElem = cell.querySelector('.tooltip-widget .OSInline:first-child') ||
                                             Array.from(cell.querySelectorAll('.OSInline')).find(el => !el.classList.contains('InfoIcon'));
                    const visibleTotal = visibleTotalElem ? visibleTotalElem.innerText.trim() : cell.innerText.trim();
                    rowData.push(visibleTotal);

                    // Simulate hover to load data
                    const tooltipWidget = cell.querySelector('.tooltip-widget');
                    if (tooltipWidget) {
                        let items = cell.querySelectorAll('.tooltip .ListRecords > div');
                        if (items.length === 0) {
                            tooltipWidget.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));

                            // Wait for AJAX to load tooltip content (up to ~3 seconds)
                            let retries = 0;
                            while (retries < 30) {
                                items = cell.querySelectorAll('.tooltip .ListRecords > div');
                                if (items.length > 0 && items[0].innerText.trim() !== '') {
                                    await new Promise(r => setTimeout(r, 100)); // allow rendering to settle
                                    break;
                                }
                                await new Promise(r => setTimeout(r, 100));
                                retries++;
                            }
                        }
                    }

                    // Extract tooltip details
                    let baseForIva = 0;
                    let foundIncomeItem = false;

                    const tooltipItems = cell.querySelectorAll('.tooltip .ListRecords > div');
                    tooltipItems.forEach(item => {
                        const inlines = item.querySelectorAll('.OSInline');
                        if (inlines.length >= 2) {
                            const desc = inlines[0].innerText.trim().replace(/\s*-\s*$/, '');
                            const val = inlines[1].innerText.trim();

                            let keyName = desc;
                            const descLower = desc.toLowerCase();

                            // Normalize the 'Reservation at...' key so it doesn't create unique columns for each reservation ID
                            if (descLower.startsWith('reservation')) {
                                keyName = 'Valor da Reserva';
                            }

                            rowInfo.extraData[keyName] = val;
                            extraColumnNames.add(keyName);

                            if (descLower.startsWith('reservation') || descLower.includes('cleaning fee') || descLower.includes('limpeza')) {
                                const numVal = parseFloat(val.replace(/[€\s]/g, '').replace(/\./g, '').replace(',', '.'));
                                if (!isNaN(numVal)) {
                                    baseForIva += numVal;
                                    foundIncomeItem = true;
                                }
                            }
                        } else {
                            const keyName = item.innerText.trim();
                            if (keyName) {
                                rowInfo.extraData[keyName] = "Sim";
                                extraColumnNames.add(keyName);
                            }
                        }
                    });

                    // Hide tooltip again
                    if (tooltipWidget) {
                         tooltipWidget.dispatchEvent(new MouseEvent('mouseleave', { bubbles: true }));
                    }

                    if (!foundIncomeItem) {
                        baseForIva = parseFloat(visibleTotal.replace(/[€\s]/g, '').replace(/\./g, '').replace(',', '.'));
                        if (isNaN(baseForIva)) baseForIva = 0;
                    }

                    // Calculate IVA (included in the base value at 6%)
                    let ivaValueStr = "";
                    let ivaNum = 0;
                    if (baseForIva > 0) {
                        ivaNum = baseForIva - (baseForIva / 1.06);
                        ivaValueStr = `€ ${ivaNum.toFixed(2).replace('.', ',')}`;
                    }

                    // Calculate Profit and Comissão de Gestão
                    let valorDaReservaNum = 0;
                    let comissoesETaxasNum = 0;
                    let cleaningFeeNum = 0;

                    for (const [key, valStr] of Object.entries(rowInfo.extraData)) {
                        const valNum = parseFloat(valStr.replace(/[€\s]/g, '').replace(/\./g, '').replace(',', '.'));
                        if (!isNaN(valNum)) {
                            const keyLower = key.toLowerCase();
                            if (keyLower === 'valor da reserva') {
                                valorDaReservaNum += valNum;
                            } else if (keyLower.includes('comissão do canal') || keyLower.includes('comissao do canal') ||
                                       keyLower.includes('comissão de pagamento') || keyLower.includes('comissao de pagamento')) {
                                comissoesETaxasNum += valNum;
                            } else if (keyLower.includes('cleaning fee') || keyLower.includes('limpeza')) {
                                cleaningFeeNum += valNum;
                            }
                        }
                    }

                    let ivaCleaningFeeNum = 0;
                    if (cleaningFeeNum !== 0) {
                        ivaCleaningFeeNum = cleaningFeeNum - (cleaningFeeNum / 1.06);
                    }

                    // Profit = Valor da Reserva - (Comissions + Taxes from tooltip + calculated IVA) + IVA da Cleaning Fee
                    // (porque a base de comissão de gestão não desconta o IVA da taxa de limpeza)
                    const profitNum = valorDaReservaNum - comissoesETaxasNum - ivaNum + ivaCleaningFeeNum;
                    const comissaoGestaoNum = profitNum * 0.30;

                    const profitStr = `€ ${profitNum.toFixed(2).replace('.', ',')}`;
                    const comissaoGestaoStr = `€ ${comissaoGestaoNum.toFixed(2).replace('.', ',')}`;

                    const totalFinalNum = profitNum - comissaoGestaoNum;
                    const totalFinalStr = `€ ${totalFinalNum.toFixed(2).replace('.', ',')}`;
                    rowInfo.totalFinalNum = totalFinalNum;
                    rowInfo.profitStr = profitStr;
                    rowInfo.comissaoGestaoStr = comissaoGestaoStr;
                    rowInfo.totalFinalStr = totalFinalStr;

                    // Add new columns to the end of the base row
                    rowData.push(ivaValueStr);
                } else {
                    rowData.push(cell.innerText.trim());
                }
            }

            allRowsData.push(rowInfo);
        }

        // Build CSV content
        const extraColumnsArray = Array.from(extraColumnNames);
        const finalHeaders = [...headerNames, ...extraColumnsArray, 'Lucro / Profit', 'Comissão de Gestão (30%)', 'Total Final'];

        // Add UTF-8 BOM so Excel opens it correctly with accents
        csvContent += '\uFEFF';
        csvContent += finalHeaders.map(escapeCSV).join(',') + '\n';

        let grandTotal = 0;
        for (const row of allRowsData) {
            if (row.totalFinalNum) grandTotal += row.totalFinalNum;
            const finalRow = [...row.baseData];
            for (const col of extraColumnsArray) {
                finalRow.push(row.extraData[col] || "");
            }
            finalRow.push(row.profitStr || "");
            finalRow.push(row.comissaoGestaoStr || "");
            finalRow.push(row.totalFinalStr || "");
            csvContent += finalRow.map(escapeCSV).join(',') + '\n';
        }

        // Add Totals row at the bottom
        const totalsRow = new Array(finalHeaders.length).fill("");
        totalsRow[0] = "TOTAL GERAL";
        totalsRow[finalHeaders.length - 1] = `€ ${grandTotal.toFixed(2).replace('.', ',')}`;
        csvContent += totalsRow.map(escapeCSV).join(',') + '\n';

        // Restore button state
        btn.innerText = originalText;
        btn.disabled = false;

        // Download
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // Generate filename with current date
        const dateStr = new Date().toISOString().split('T')[0];
        a.download = `talkguest_reservations_${dateStr}.csv`;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function createButton() {
        if (document.getElementById('talkguest-csv-btn')) return;

        const btn = document.createElement('button');
        btn.id = 'talkguest-csv-btn';
        btn.innerText = 'Download CSV';
        btn.style.cssText = `
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 9999;
            padding: 12px 24px;
            background-color: #007bff;
            color: white;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            font-family: Arial, sans-serif;
            font-size: 14px;
            font-weight: bold;
            transition: background-color 0.2s;
        `;

        btn.addEventListener('mouseenter', () => btn.style.backgroundColor = '#0056b3');
        btn.addEventListener('mouseleave', () => btn.style.backgroundColor = '#007bff');
        btn.addEventListener('click', downloadCSV);

        document.body.appendChild(btn);
    }

    // Try to create the button periodically as the page might be SPA or slow to load
    setInterval(createButton, 2000);

})();
