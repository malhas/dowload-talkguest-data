// ==UserScript==
// @name         Talkguest Table Downloader
// @namespace    http://tampermonkey.net/
// @version      0.2
// @description  Download reservations table to XLSX with formulas
// @author       You
// @match        https://owner.talkguest.com/Bookings/OwnerArea.aspx*
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    async function downloadXLSX() {
        const table = document.querySelector('table.TableRecords');
        if (!table) {
            alert('Table not found!');
            return;
        }

        const btn = document.getElementById('talkguest-xlsx-btn');
        const originalText = btn.innerText;
        btn.innerText = 'Loading details...';
        btn.disabled = true;

        const headers = Array.from(table.querySelectorAll('thead th'));
        const baseHeaderNames = headers.map(th => th.innerText.trim());

        const allRowsData = [];
        const extraColumnNames = new Set();

        const rows = Array.from(table.querySelectorAll('tbody > tr'));
        for (const row of rows) {
            const cells = Array.from(row.querySelectorAll('td'));
            if (cells.length < headers.length) continue;

            const rowData = [];
            const rowInfo = { baseData: rowData, extraData: {}, foundIncomeItem: false, totalColIndex: -1 };

            for (let index = 0; index < cells.length; index++) {
                const cell = cells[index];

                if (index === headers.length - 1 || cell.querySelector('.tooltip-container')) {
                    rowInfo.totalColIndex = index;
                    const visibleTotalElem = cell.querySelector('.tooltip-widget .OSInline:first-child') ||
                                             Array.from(cell.querySelectorAll('.OSInline')).find(el => !el.classList.contains('InfoIcon'));
                    const visibleTotal = visibleTotalElem ? visibleTotalElem.innerText.trim() : cell.innerText.trim();
                    rowData.push(visibleTotal);

                    const tooltipWidget = cell.querySelector('.tooltip-widget');
                    if (tooltipWidget) {
                        let items = cell.querySelectorAll('.tooltip .ListRecords > div');
                        if (items.length === 0) {
                            tooltipWidget.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));

                            let retries = 0;
                            while (retries < 30) {
                                items = cell.querySelectorAll('.tooltip .ListRecords > div');
                                if (items.length > 0 && items[0].innerText.trim() !== '') {
                                    await new Promise(r => setTimeout(r, 100));
                                    break;
                                }
                                await new Promise(r => setTimeout(r, 100));
                                retries++;
                            }
                        }

                        const tooltipItems = cell.querySelectorAll('.tooltip .ListRecords > div');
                        tooltipItems.forEach(item => {
                            const inlines = item.querySelectorAll('.OSInline');
                            if (inlines.length >= 2) {
                                const desc = inlines[0].innerText.trim().replace(/\s*-\s*$/, '');
                                const val = inlines[1].innerText.trim();

                                let keyName = desc;
                                const descLower = desc.toLowerCase();

                                if (descLower.startsWith('reservation')) {
                                    keyName = 'Valor da Reserva';
                                }

                                rowInfo.extraData[keyName] = val;
                                extraColumnNames.add(keyName);

                                if (descLower.startsWith('reservation') || descLower.includes('cleaning fee') || descLower.includes('limpeza')) {
                                    rowInfo.foundIncomeItem = true;
                                }
                            } else {
                                const keyName = item.innerText.trim();
                                if (keyName) {
                                    rowInfo.extraData[keyName] = "Sim";
                                    extraColumnNames.add(keyName);
                                }
                            }
                        });

                        tooltipWidget.dispatchEvent(new MouseEvent('mouseleave', { bubbles: true }));
                    }
                } else {
                    rowData.push(cell.innerText.trim());
                }
            }

            allRowsData.push(rowInfo);
        }

        const extraColumnsArray = Array.from(extraColumnNames);
        if (!extraColumnsArray.includes('Valor da Reserva')) {
            extraColumnsArray.unshift('Valor da Reserva');
        }
        const finalHeaders = [...baseHeaderNames, 'IVA (6%)', ...extraColumnsArray, 'Lucro / Profit', 'Comissão de Gestão (30%)', 'Total Final'];

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Reservations');

        worksheet.addRow(finalHeaders);
        worksheet.getRow(1).font = { bold: true };

        function getColLetter(colIndex) { // 0-based
            let temp, letter = '';
            while (colIndex >= 0) {
                temp = colIndex % 26;
                letter = String.fromCharCode(temp + 65) + letter;
                colIndex = (colIndex - temp) / 26 - 1;
            }
            return letter;
        }

        function parseNum(val) {
            if (!val) return 0;
            if (typeof val === 'number') return val;
            const num = parseFloat(String(val).replace(/[€\s]/g, '').replace(/\./g, '').replace(',', '.'));
            return isNaN(num) ? 0 : num;
        }

        for (let i = 0; i < allRowsData.length; i++) {
            const rowInfo = allRowsData[i];
            const rowNum = i + 2;
            const rowValues = [];

            // Base columns
            for (let j = 0; j < baseHeaderNames.length; j++) {
                let val = rowInfo.baseData[j];
                if (j === rowInfo.totalColIndex) {
                    val = parseNum(val);
                } else if (typeof val === 'string' && val.includes('€')) {
                    val = parseNum(val);
                }
                rowValues.push(val);
            }

            // IVA (6%)
            let ivaFormula;
            const totalColLetter = getColLetter(rowInfo.totalColIndex);
            if (rowInfo.foundIncomeItem) {
                const incomeCols = [];
                for (let c = 0; c < extraColumnsArray.length; c++) {
                    const colName = extraColumnsArray[c].toLowerCase();
                    if (colName === 'valor da reserva' || colName.includes('cleaning fee') || colName.includes('limpeza')) {
                        incomeCols.push(getColLetter(baseHeaderNames.length + 1 + c));
                    }
                }
                const sumStr = incomeCols.length > 0 ? `SUM(${incomeCols.map(c => c+rowNum).join(',')})` : '0';
                ivaFormula = `=${sumStr} - (${sumStr})/1.06`;
            } else {
                ivaFormula = `=${totalColLetter}${rowNum} - ${totalColLetter}${rowNum}/1.06`;
            }
            rowValues.push({ formula: ivaFormula });

            // Extra columns
            for (let c = 0; c < extraColumnsArray.length; c++) {
                const colName = extraColumnsArray[c];

                if (colName === 'Valor da Reserva') {
                    const totalColLetter = getColLetter(rowInfo.totalColIndex);
                    const taxaCols = [];
                    for (let k = 0; k < extraColumnsArray.length; k++) {
                        const lowerName = extraColumnsArray[k].toLowerCase();
                        if (lowerName.includes('taxa municipal') || lowerName.includes('city tax') || lowerName.includes('taxa turística') || lowerName.includes('taxa turistica')) {
                            taxaCols.push(getColLetter(baseHeaderNames.length + 1 + k));
                        }
                    }
                    if (taxaCols.length > 0) {
                        const taxaSumExpr = taxaCols.length > 1 ? `SUM(${taxaCols.map(tc => tc+rowNum).join(',')})` : `${taxaCols[0]}${rowNum}`;
                        rowValues.push({ formula: `=${totalColLetter}${rowNum} - ${taxaSumExpr}` });
                    } else {
                        rowValues.push({ formula: `=${totalColLetter}${rowNum}` });
                    }
                } else {
                    let val = rowInfo.extraData[colName];
                    if (val === undefined || val === null) {
                        const lowerName = colName.toLowerCase();
                        if (lowerName.includes('taxa municipal') || lowerName.includes('city tax') || lowerName.includes('taxa turística') || lowerName.includes('taxa turistica')) {
                            val = 0;
                        } else {
                            val = "";
                        }
                    } else if (typeof val === 'string') {
                        if (val === 'Sim') val = "Sim";
                        else val = parseNum(val);
                    }
                    rowValues.push(val);
                }
            }

            // Formulas for Profit, Comissão, Total
            const colValorReserva = [];
            const colComissoes = [];
            const colCleaningFee = [];
            for (let c = 0; c < extraColumnsArray.length; c++) {
                const colName = extraColumnsArray[c].toLowerCase();
                const letter = getColLetter(baseHeaderNames.length + 1 + c);
                if (colName === 'valor da reserva') colValorReserva.push(letter);
                else if (colName.includes('comissão do canal') || colName.includes('comissao do canal') || colName.includes('comissão de pagamento') || colName.includes('comissao de pagamento')) colComissoes.push(letter);
                else if (colName.includes('cleaning fee') || colName.includes('limpeza')) colCleaningFee.push(letter);
            }

            const sumExpr = (cols) => cols.length > 0 ? `SUM(${cols.map(c => c+rowNum).join(',')})` : '0';
            const vrExpr = sumExpr(colValorReserva);
            const comExpr = sumExpr(colComissoes);
            const ivaCell = getColLetter(baseHeaderNames.length) + rowNum;
            const cfExpr = sumExpr(colCleaningFee);
            const ivaCfExpr = cfExpr === '0' ? '0' : `(${cfExpr} - (${cfExpr})/1.06)`;

            const profitFormula = `=${vrExpr} - ${comExpr} - SUM(${ivaCell}) + ${ivaCfExpr}`;
            rowValues.push({ formula: profitFormula });

            const profitColLetter = getColLetter(baseHeaderNames.length + 1 + extraColumnsArray.length);
            const comissaoFormula = `=${profitColLetter}${rowNum}*0.30`;
            rowValues.push({ formula: comissaoFormula });

            const comissaoColLetter = getColLetter(baseHeaderNames.length + 1 + extraColumnsArray.length + 1);
            const totalFormula = `=${profitColLetter}${rowNum} - ${comissaoColLetter}${rowNum}`;
            rowValues.push({ formula: totalFormula });

            const wsRow = worksheet.addRow(rowValues);

            wsRow.eachCell((cell, colNumber) => {
                const colIndex = colNumber - 1;
                if (colIndex === rowInfo.totalColIndex ||
                    colIndex === baseHeaderNames.length ||
                    colIndex >= baseHeaderNames.length + 1 + extraColumnsArray.length) {
                    cell.numFmt = '#,##0.00" €"';
                } else if (colIndex > baseHeaderNames.length && colIndex < baseHeaderNames.length + 1 + extraColumnsArray.length) {
                    if (typeof cell.value === 'number') {
                        cell.numFmt = '#,##0.00" €"';
                    }
                }
            });
        }

        // Add totals row
        const totalsRowNum = allRowsData.length + 2;
        const totalsValues = new Array(finalHeaders.length).fill("");
        totalsValues[0] = "TOTAL GERAL";
        const totalFinalColLetter = getColLetter(finalHeaders.length - 1);
        totalsValues[finalHeaders.length - 1] = { formula: `=SUM(${totalFinalColLetter}2:${totalFinalColLetter}${totalsRowNum-1})` };

        const totalsRow = worksheet.addRow(totalsValues);
        totalsRow.getCell(finalHeaders.length).numFmt = '#,##0.00" €"';
        totalsRow.font = { bold: true };

        btn.innerText = originalText;
        btn.disabled = false;

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        let fileName = "";
        const parts = [];
        const filterLabels = Array.from(document.querySelectorAll('label'));
        filterLabels.forEach(label => {
            const text = label.innerText.trim().toUpperCase();
            if (['CHECK-IN', 'ALOJAMENTO', 'ESTADO', 'CANAL', 'ORIGEM'].includes(text)) {
                const inputId = label.getAttribute('for');
                if (inputId) {
                    const el = document.getElementById(inputId);
                    if (el) {
                        let val = el.tagName === 'SELECT' && el.selectedIndex >= 0 ? el.options[el.selectedIndex].text : el.value;
                        val = (val || '').trim();
                        if (val && val !== '-' && val.toLowerCase() !== 'all' && val.toLowerCase() !== 'todos') {
                            val = val.replace(/\s*-\s*/g, '_').replace(/\//g, '-').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9\-_\u00C0-\u017F]/g, '');
                            if (val) parts.push(val);
                        }
                    }
                }
            }
        });

        if (parts.length === 0) {
            const selects = Array.from(document.querySelectorAll('select'));
            const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
            const filterElements = [...inputs.slice(0, 5), ...selects.slice(0, 5)];
            filterElements.forEach(el => {
                let val = el.tagName === 'SELECT' && el.selectedIndex >= 0 ? el.options[el.selectedIndex].text : el.value;
                val = (val || '').trim();
                if (val && val !== '-' && val.toLowerCase() !== 'all' && val.toLowerCase() !== 'todos' && !val.includes('Pesquisar')) {
                    val = val.replace(/\s*-\s*/g, '_').replace(/\//g, '-').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9\-_\u00C0-\u017F]/g, '');
                    if (val) parts.push(val);
                }
            });
        }

        const uniqueParts = [...new Set(parts)];
        if (uniqueParts.length > 0) {
            fileName = `talkguest_${uniqueParts.join('_')}.xlsx`;
        } else {
            const dateStr = new Date().toISOString().split('T')[0];
            fileName = `talkguest_reservations_${dateStr}.xlsx`;
        }

        a.download = fileName;

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function createButton() {
        if (document.getElementById('talkguest-xlsx-btn')) return;

        const btn = document.createElement('button');
        btn.id = 'talkguest-xlsx-btn';
        btn.innerText = 'Download XLSX';
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
        btn.addEventListener('click', downloadXLSX);

        document.body.appendChild(btn);
    }

    setInterval(createButton, 2000);

})();
