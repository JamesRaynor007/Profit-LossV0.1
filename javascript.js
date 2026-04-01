// script.js

const dataContainer = document.getElementById('dataContainer');
const fileInput = document.getElementById('fileInput');
const ratiosDiv = document.getElementById('ratios');

let currentData = [];
let revenueCurrent = 0;
let revenueLast = 0;

// Variables para métricas importantes
let grossProfitCurr = 0, grossProfitLast = 0;
let operatingProfitCurr = 0, operatingProfitLast = 0;
let netIncomeCurr = 0, netIncomeLast = 0;
let incomeBeforeIncomeTaxesCurr = 0, incomeBeforeIncomeTaxesLast = 0;
let salesCurr = 0, salesLast = 0;

// Encabezados para la tabla
const HEADER_TITLES = ['ACCOUNT', 'CURRENT', 'LAST'];

// Función para formatear números con puntos como separador de miles y 2 decimales
function formatNumber(num) {
    if (isNaN(num)) return '';
    return num.toLocaleString('de-DE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Función para parsear valores de cadenas con formato de miles y decimales
function parseFormattedNumber(str) {
    if (typeof str !== 'string') return 0;
    const cleanStr = str.replace(/[^0-9,\.-]/g, '').replace(',', '.');
    const parsed = parseFloat(cleanStr);
    return isNaN(parsed) ? 0 : parsed;
}

// Función auxiliar para determinar si la cuenta es una de las que deben tener color verde en % VAR
function isPositiveInImportantAccounts(label) {
    const importantAccounts = [
        'SALES', 'SALES RETURNS', 'SALES DISCOUNTS', 'NET SALES',
        'GROSS PROFIT', 'OPERATING PROFIT', 'NET INCOME'
    ];
    if (!label || typeof label !== 'string') return false;
    const labelUpper = label.toUpperCase();
    return importantAccounts.some(acc => labelUpper.includes(acc));
}

// ------------------ Funciones para la tabla de carga de template ------------------

// Encabezado para la plantilla
document.getElementById('downloadTemplate').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    const ws_data = [
        ['Item', 'Current', 'Last'], // encabezado
        ['Sales', 0, 0],
        ['Sales Returns', 0, 0],
        ['Sales Discounts', 0, 0],
        ['Net Sales', 0, 0],
        ['Costo of Goods Sold', 0, 0],
        ['Gross Profit', 0, 0],
        ['Operating Expenses', 0, 0],
        ['Salaries & Wages', 0, 0],
        ['Depreciation Expenses', 0, 0],
        ['Office Expenses', 0, 0],
        ['Rent Expenses', 0, 0],
        ['Travel Expenses', 0, 0],
        ['Repair and Maintenance Expenses', 0, 0],
        ['Advertising Expenses', 0, 0],
        ['Utilities Expenses', 0, 0],
        ['Bank Fees and Charges', 0, 0],
        ['Professional Fees', 0, 0],
        ['Insurance Expenses', 0, 0],
        ['Total Operating Expenses', 0, 0],
        ['Operating Profit', 0, 0],
        ['Interest Income (Expense)', 0, 0],
        ['Income Before Income Taxes', 0, 0],
        ['Income Tax Expense', 0, 0],
        ['Net Income', 0, 0],
    ];
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, 'Profit & Loss');
    XLSX.writeFile(wb, 'template_profit_loss.xlsx');
});

// Evento para cargar archivo XLSX
document.getElementById('loadTemplate').addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, {header:1});
        currentData = jsonData;
        resetMetrics();
        displayData();
    };
    reader.readAsArrayBuffer(file);
});

function resetMetrics() {
    revenueCurrent = 0;
    revenueLast = 0;
    grossProfitCurr = 0; grossProfitLast = 0;
    operatingProfitCurr = 0; operatingProfitLast = 0;
    netIncomeCurr = 0; netIncomeLast = 0;
    incomeBeforeIncomeTaxesCurr = 0; incomeBeforeIncomeTaxesLast = 0;
    salesCurr = 0; salesLast = 0;
}

// Función para mostrar datos en la tabla con encabezados
function displayData() {
    dataContainer.innerHTML = '';

    if (!currentData.length) return;

    const table = document.createElement('table');

    // Crear encabezado de acuerdo a HEADER_TITLES + columna "% VAR"
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    HEADER_TITLES.forEach(title => {
        const th = document.createElement('th');
        th.textContent = title;
        th.style.textAlign = 'center';
        headerRow.appendChild(th);
    });
    // Añadimos columna "% VAR" solo aquí
    const thVar = document.createElement('th');
    thVar.textContent = '% VAR';
    thVar.style.textAlign = 'center';
    headerRow.appendChild(thVar);

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    let revenueSum = 0;
    let revenueSumLast = 0;

    currentData.forEach((row, index) => {
        const label = row[0];
        let currentValue = row[1];
        let lastValue = row[2];

        // Convertir valores a números
        let numCurrent = parseFloat(currentValue);
        if (isNaN(numCurrent)) numCurrent = 0;

        let numLast = parseFloat(lastValue);
        if (isNaN(numLast) ) numLast = 0;

        // Filtrar filas con saldo 0 excluyendo headers y sumas
        const isHeader = typeof label === 'string' && label.toLowerCase().includes('header');
        const isSum = typeof label === 'string' && label.toLowerCase().includes('sum');

        if (!isHeader && !isSum && numCurrent === 0 && numLast === 0) {
            return; // saltar filas con saldo cero
        }

        const tr = document.createElement('tr');

        if (isHeader) {
            tr.className = 'header';
        } else if (isSum) {
            tr.className = 'sum';
        }

        // Crear celdas en orden según HEADER_TITLES
        HEADER_TITLES.forEach(title => {
            const td = document.createElement('td');

            if (title === 'ACCOUNT') {
                td.textContent = label;
                td.style.textAlign = 'left';
            } else if (title === 'CURRENT') {
                td.contentEditable = true;
                // Formatear con "$" y 2 decimales con separador de miles
                td.textContent = isNaN(numCurrent) ? '' : `$${formatNumber(numCurrent)}`;
                td.style.textAlign = 'right';

                // Evento para actualizar en tiempo real
                td.addEventListener('input', () => {
                    const raw = td.textContent.replace(/[^0-9,\.-]/g, '');
                    const val = parseFormattedNumber(raw);
                    tr.dataset.current = val;
                });
            } else if (title === 'LAST') {
                td.contentEditable = true;
                // Formatear con "$" y 2 decimales con separador de miles
                td.textContent = isNaN(numLast) ? '' : `$${formatNumber(numLast)}`;
                td.style.textAlign = 'right';

                // Evento para actualizar en tiempo real
                td.addEventListener('input', () => {
                    const raw = td.textContent.replace(/[^0-9,\.-]/g, '');
                    const val = parseFormattedNumber(raw);
                    tr.dataset.last = val;
                });
            }
            tr.appendChild(td);
        });

        // Añadir columna "% VAR" solo en visualización
        let percentVar = 0;
        if (numLast !== 0) {
            percentVar = ((numCurrent / numLast) - 1) * 100; // variación porcentual
        } else if (numCurrent !== 0 && numLast === 0) {
            percentVar = 100;
        } else {
            percentVar = 0;
        }

        const tdVar = document.createElement('td');
        tdVar.textContent = `${formatNumber(percentVar)}%`;
        tdVar.style.textAlign = 'right';

        // Aplicar color condicional según la cuenta
        if (isPositiveInImportantAccounts(label)) {
            // Si importante, verde si % VAR positivo, rojo si negativo
            if (percentVar > 0) {
                tdVar.style.color = 'green';
            } else if (percentVar < 0) {
                tdVar.style.color = 'red';
            } else {
                tdVar.style.color = 'black';
            }
        } else {
            // En otras cuentas, positivo en rojo, negativo en verde
            if (percentVar > 0) {
                tdVar.style.color = 'red';
            } else if (percentVar < 0) {
                tdVar.style.color = 'green';
            } else {
                tdVar.style.color = 'black';
            }
        }

        tr.appendChild(tdVar);

        // Guardar los valores en dataset
        tr.dataset.label = label;
        tr.dataset.current = numCurrent;
        tr.dataset.last = numLast;

        tbody.appendChild(tr);

        // Sumar ingresos para cálculos de ratios
        if (typeof label === 'string' && 
            (label.toLowerCase().includes('sales') || label.toLowerCase().includes('net sales') || label.toLowerCase().includes('income'))) {
            if (!isNaN(numCurrent)) {
                revenueSum += numCurrent;
            }
            if (!isNaN(numLast)) {
                revenueSumLast += numLast;
            }
        }
    });

    table.appendChild(tbody);
    dataContainer.appendChild(table);

    // Actualizar revenue para período actual y pasado
    revenueCurrent = revenueSum || 1; // evitar división por cero
    revenueLast = revenueSumLast || 1;

    // Actualizar métricas principales
    updateMetrics();
}

// ------------------ Funciones para métricas principales ------------------
function updateMetrics() {
    grossProfitCurr = 0; grossProfitLast = 0;
    operatingProfitCurr = 0; operatingProfitLast = 0;
    netIncomeCurr = 0; netIncomeLast = 0;
    incomeBeforeIncomeTaxesCurr = 0; incomeBeforeIncomeTaxesLast = 0;
    salesCurr = 0; salesLast = 0;

    currentData.forEach(row => {
        const label = row[0];
        const currentVal = parseFloat(row[1]);
        const lastVal = parseFloat(row[2]);
        if (isNaN(currentVal) || isNaN(lastVal)) return;

        switch (label.toLowerCase()) {
            case 'gross profit':
                grossProfitCurr = currentVal;
                grossProfitLast = lastVal;
                break;
            case 'operating profit':
                operatingProfitCurr = currentVal;
                operatingProfitLast = lastVal;
                break;
            case 'net income':
                netIncomeCurr = currentVal;
                netIncomeLast = lastVal;
                break;
            case 'income before income taxes':
                incomeBeforeIncomeTaxesCurr = currentVal;
                incomeBeforeIncomeTaxesLast = lastVal;
                break;
            case 'sales':
            case 'net sales':
                salesCurr = currentVal;
                salesLast = lastVal;
                break;
            default:
                break;
        }
    });
}

// ------------------ Función para calcular ratios y mostrar resultados ------------------
document.getElementById('calculateRatios').addEventListener('click', () => {
    const safeRevenueCurr = revenueCurrent === 0 ? 1 : revenueCurrent;
    const safeRevenueLast = revenueLast === 0 ? 1 : revenueLast;

    const grossProfitMarginCurr = (grossProfitCurr / safeRevenueCurr) * 100;
    const operatingProfitMarginCurr = (operatingProfitCurr / safeRevenueCurr) * 100;
    const netProfitMarginCurr = (netIncomeCurr / safeRevenueCurr) * 100;
    const ROS_curr = (netIncomeCurr / safeRevenueCurr) * 100;
    const EBIT_margin_curr = (incomeBeforeIncomeTaxesCurr / safeRevenueCurr) * 100;

    const grossProfitMarginLast = (grossProfitLast / safeRevenueLast) * 100;
    const operatingProfitMarginLast = (operatingProfitLast / safeRevenueLast) * 100;
    const netProfitMarginLast = (netIncomeLast / safeRevenueLast) * 100;
    const ROS_last = (netIncomeLast / safeRevenueLast) * 100;
    const EBIT_margin_last = (incomeBeforeIncomeTaxesLast / safeRevenueLast) * 100;

    const varGrossProfitMargin = grossProfitMarginCurr - grossProfitMarginLast;
    const varOperatingProfitMargin = operatingProfitMarginCurr - operatingProfitMarginLast;
    const varNetProfitMargin = netProfitMarginCurr - netProfitMarginLast;
    const varROS = ROS_curr - ROS_last;
    const varEBIT = EBIT_margin_curr - EBIT_margin_last;

    // Crear matriz de ratios
    const ratiosMatrix = [
        ['Gross Profit Margin (%)', grossProfitMarginCurr, grossProfitMarginLast, varGrossProfitMargin],
        ['Operating Profit Margin (%)', operatingProfitMarginCurr, operatingProfitMarginLast, varOperatingProfitMargin],
        ['Net Profit Margin (%)', netProfitMarginCurr, netProfitMarginLast, varNetProfitMargin],
        ['Return On Sales (ROS) (%)', ROS_curr, ROS_last, varROS],
        ['EBIT Margin (%)', EBIT_margin_curr, EBIT_margin_last, varEBIT]
    ];

    // Crear tabla para ratios
    ratiosDiv.innerHTML = ''; // limpiar
    const table = document.createElement('table');
    table.style.borderCollapse = 'collapse';
    table.style.width = '100%';

    // Encabezados
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    ['METRIC', 'CURRENT', 'LAST', 'VARIATION'].forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        th.style.border = '1px solid #000';
        th.style.padding = '8px';
        th.style.textAlign = 'center';
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Añadir filas con formato porcentaje y color condicional
    ratiosMatrix.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach((cell, index) => {
            const td = document.createElement('td');
            td.style.border = '1px solid #000';
            td.style.padding = '8px';
            td.style.textAlign = index === 0 ? 'left' : 'right';

            if (index === 0) {
                td.textContent = cell; // nombre de la métrica
            } else {
                // Mostrar número formateado con porcentaje
                td.textContent = `${formatNumber(cell)}%`;
                // Color condicional en la columna VARIATION (última columna)
                if (index === 3) {
                    // determinar si la métrica es importante
                    if (isPositiveInImportantAccounts(row[0])) {
                        // si importante: verde si % VAR positivo, rojo si negativo
                        if (cell > 0) {
                            td.style.color = 'green';
                        } else if (cell < 0) {
                            td.style.color = 'red';
                        } else {
                            td.style.color = 'black';
                        }
                    } else {
                        // en otras: rojo si % VAR positivo, verde si negativo
                        if (cell > 0) {
                            td.style.color = 'red';
                        } else if (cell < 0) {
                            td.style.color = 'green';
                        } else {
                            td.style.color = 'black';
                        }
                    }
                }
            }
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    ratiosDiv.appendChild(table);
});
