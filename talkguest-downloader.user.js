// ==UserScript==
// @name         Talkguest Table Downloader
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  Download reservations table to XLSX with formulas
// @author       You
// @match        https://owner.talkguest.com/Bookings/OwnerArea.aspx*
// @require      https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js
// @updateURL    https://raw.githubusercontent.com/malhas/dowload-talkguest-data/main/talkguest-downloader.user.js
// @downloadURL  https://raw.githubusercontent.com/malhas/dowload-talkguest-data/main/talkguest-downloader.user.js
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
                                    keyName = 'Vendas';
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

        const extraColumnsArray = Array.from(extraColumnNames).filter(name => name !== 'Vendas');
        if (!extraColumnsArray.includes('Vendas')) {
            extraColumnsArray.unshift('Vendas');
        }
        const finalHeaders = [
            ...baseHeaderNames,
            'IVA (6%)',
            ...extraColumnsArray,
            'Comissão de Gestão (30%)',
            'Total Final',
            'Despesas Proprietário',
            'Despesas ORM',
            'Total a Transferir'
        ];

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

        function parseDate(val) {
            if (!val) return new Date(8640000000000000);
            const parts = String(val).trim().match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
            if (parts) {
                return new Date(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1]));
            }
            const parsed = new Date(val);
            return isNaN(parsed.getTime()) ? new Date(8640000000000000) : parsed;
        }

        const checkoutColIndex = baseHeaderNames.findIndex(name => {
            const normalized = name.toLowerCase().replace(/\s+/g, '');
            return normalized === 'check-out' || normalized === 'checkout';
        });

        if (checkoutColIndex >= 0) {
            allRowsData.sort((a, b) => parseDate(a.baseData[checkoutColIndex]) - parseDate(b.baseData[checkoutColIndex]));
        }

        const formula = (expression, result) => {
            const value = { formula: expression };
            if (typeof result === 'number' && isFinite(result)) {
                value.result = result;
            }
            return value;
        };
        const totalsByColumn = new Array(finalHeaders.length).fill(0);

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

            const totalColLetter = getColLetter(rowInfo.totalColIndex);

            const getExtraCols = matcher => {
                const cols = [];
                for (let c = 0; c < extraColumnsArray.length; c++) {
                    if (matcher(extraColumnsArray[c].toLowerCase())) {
                        cols.push(getColLetter(baseHeaderNames.length + 1 + c));
                    }
                }
                return cols;
            };

            const sumExpr = cols => cols.length > 0 ? `SUM(${cols.map(c => c + rowNum).join(',')})` : '0';
            const taxaCols = getExtraCols(name => name.includes('taxa municipal') || name.includes('city tax') || name.includes('taxa turística') || name.includes('taxa turistica'));
            const cleaningCols = getExtraCols(name => name.includes('cleaning fee') || name.includes('limpeza'));
            const channelCommissionCols = getExtraCols(name => name.includes('comissão do canal') || name.includes('comissao do canal'));
            const paymentCommissionCols = getExtraCols(name => name.includes('comissão de pagamento') || name.includes('comissao de pagamento'));
            const vendasCols = getExtraCols(name => name === 'vendas');

            const taxaExpr = sumExpr(taxaCols);
            const cleaningExpr = sumExpr(cleaningCols);
            const revenueBaseExpr = `(${totalColLetter}${rowNum} - ${taxaExpr} - ${cleaningExpr})`;

            const sumExtraValues = matcher => extraColumnsArray.reduce((sum, colName) => {
                return matcher(colName.toLowerCase()) ? sum + parseNum(rowInfo.extraData[colName]) : sum;
            }, 0);
            const totalValue = parseNum(rowInfo.baseData[rowInfo.totalColIndex]);
            const taxaValue = sumExtraValues(name => name.includes('taxa municipal') || name.includes('city tax') || name.includes('taxa turística') || name.includes('taxa turistica'));
            const cleaningValue = sumExtraValues(name => name.includes('cleaning fee') || name.includes('limpeza'));
            const channelCommissionValue = sumExtraValues(name => name.includes('comissão do canal') || name.includes('comissao do canal'));
            const paymentCommissionValue = sumExtraValues(name => name.includes('comissão de pagamento') || name.includes('comissao de pagamento'));
            const revenueBaseValue = totalValue - taxaValue - cleaningValue;
            const ivaValue = revenueBaseValue - revenueBaseValue / 1.06;
            const vendasValue = revenueBaseValue / 1.06 - channelCommissionValue - paymentCommissionValue;
            const managementCommissionValue = vendasValue * 0.30;
            const totalFinalValue = vendasValue - managementCommissionValue;

            // IVA (6%) included in Vendas base, excluding city tax and cleaning fees.
            rowValues.push(formula(`${revenueBaseExpr} - (${revenueBaseExpr})/1.06`, ivaValue));

            // Extra columns
            for (let c = 0; c < extraColumnsArray.length; c++) {
                const colName = extraColumnsArray[c];

                if (colName === 'Vendas') {
                    const channelCommissionExpr = sumExpr(channelCommissionCols);
                    const paymentCommissionExpr = sumExpr(paymentCommissionCols);
                    rowValues.push(formula(`${revenueBaseExpr}/1.06 - ${channelCommissionExpr} - ${paymentCommissionExpr}`, vendasValue));
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

            // Formulas for Comissão de Gestão and Total, based on Vendas.
            const vendasExpr = vendasCols.length > 0 ? `${vendasCols[0]}${rowNum}` : '0';
            rowValues.push(formula(`${vendasExpr}*0.30`, managementCommissionValue));

            const comissaoColLetter = getColLetter(baseHeaderNames.length + 1 + extraColumnsArray.length);
            rowValues.push(formula(`${vendasExpr} - ${comissaoColLetter}${rowNum}`, totalFinalValue));

            rowValues.push("");
            rowValues.push("");
            rowValues.push("");

            const wsRow = worksheet.addRow(rowValues);

            rowValues.forEach((value, index) => {
                if (typeof value === 'number') {
                    totalsByColumn[index] += value;
                } else if (value && typeof value.result === 'number') {
                    totalsByColumn[index] += value.result;
                }
            });

            wsRow.eachCell((cell, colNumber) => {
                const colIndex = colNumber - 1;
                if (colIndex === rowInfo.totalColIndex ||
                    colIndex === baseHeaderNames.length ||
                    finalHeaders[colIndex] === 'Vendas' ||
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
        for (let colIndex = 0; colIndex < finalHeaders.length; colIndex++) {
            const header = String(finalHeaders[colIndex]).toLowerCase();
            const shouldTotal = header === 'vendas' ||
                                header.includes('comissão') ||
                                header.includes('comissao') ||
                                header === 'total final';
            if (shouldTotal) {
                const colLetter = getColLetter(colIndex);
                totalsValues[colIndex] = formula(`SUM(${colLetter}2:${colLetter}${totalsRowNum-1})`, totalsByColumn[colIndex]);
            }
        }
        const totalFinalIndex = finalHeaders.indexOf('Total Final');
        const despesasProprietarioIndex = finalHeaders.indexOf('Despesas Proprietário');
        const despesasOrmIndex = finalHeaders.indexOf('Despesas ORM');
        const totalTransferirIndex = finalHeaders.indexOf('Total a Transferir');
        if (despesasProprietarioIndex >= 0) {
            totalsValues[despesasProprietarioIndex] = 0;
        }
        if (despesasOrmIndex >= 0) {
            totalsValues[despesasOrmIndex] = 0;
        }
        if (totalFinalIndex >= 0 && despesasProprietarioIndex >= 0 && despesasOrmIndex >= 0 && totalTransferirIndex >= 0) {
            const totalFinalCell = `${getColLetter(totalFinalIndex)}${totalsRowNum}`;
            const despesasProprietarioCell = `${getColLetter(despesasProprietarioIndex)}${totalsRowNum}`;
            const despesasOrmCell = `${getColLetter(despesasOrmIndex)}${totalsRowNum}`;
            totalsValues[totalTransferirIndex] = formula(
                `${totalFinalCell} - ${despesasProprietarioCell} - ${despesasOrmCell}`,
                totalsByColumn[totalFinalIndex]
            );
        }

        const totalsRow = worksheet.addRow(totalsValues);
        for (let colIndex = 0; colIndex < finalHeaders.length; colIndex++) {
            if (typeof totalsValues[colIndex] === 'number' ||
                (totalsValues[colIndex] && typeof totalsValues[colIndex].result === 'number')) {
                totalsRow.getCell(colIndex + 1).numFmt = '#,##0.00" €"';
            }
        }
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
