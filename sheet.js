let excelData = []; // Placeholder for Excel data
let filteredData = []; // Placeholder for filtered data

// Load the Google Sheets file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelSheet(fileUrl);
    }
});

// Function to load Excel sheet data
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        excelData = XLSX.utils.sheet_to_json(sheet, { defval: null });
        displaySheet(excelData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Display Sheet
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null || cell === "" ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply Operation
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operation = document.getElementById('operation').value;
    const contentType = document.getElementById('content-type').value;

    const rowRangeFrom = parseInt(document.getElementById('row-range-from').value, 10);
    const rowRangeTo = parseInt(document.getElementById('row-range-to').value, 10);

    if (!primaryColumn || !operationColumnsInput) {
        alert('Please enter the primary column and columns to operate on.');
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim());
    filteredData = excelData.filter((row, index) => {
        // Check if the current row index is within the specified range
        if (index < rowRangeFrom - 1 || index > rowRangeTo - 1) return false;

        const isPrimaryNull = row[primaryColumn] === null || row[primaryColumn] === "";

        if (operation === 'not-null') {
            if (contentType === 'filled') {
                return !isPrimaryNull;
            } else if (contentType === 'words') {
                return typeof row[primaryColumn] === 'string' && row[primaryColumn].match(/^[a-zA-Z]+$/);
            } else if (contentType === 'numbers') {
                return typeof row[primaryColumn] === 'number';
            } else if (contentType === 'links') {
                return typeof row[primaryColumn] === 'string' && row[primaryColumn].match(/https?:\/\/[^\s]+/);
            }
        } else if (operation === 'null') {
            return isPrimaryNull;
        }

        return false;
    });

    // Optionally, you can highlight the selected rows and columns here
    highlightRowsAndColumns(rowRangeFrom, rowRangeTo, operationColumns);

    // Display the filtered data
    displaySheet(filteredData);
}

// Highlight selected rows and columns
function highlightRowsAndColumns(fromRow, toRow, operationColumns) {
    const sheetContentDiv = document.getElementById('sheet-content');
    const tableRows = sheetContentDiv.querySelectorAll('tr');

    // Clear previous highlights
    tableRows.forEach(row => row.classList.remove('highlight'));

    // Highlight the specified rows
    for (let i = fromRow; i <= toRow; i++) {
        if (tableRows[i]) {
            tableRows[i].classList.add('highlight');
        }
    }

    // Highlight the specified columns
    operationColumns.forEach(col => {
        const colIndex = Array.from(tableRows[0].cells).findIndex(th => th.textContent === col);
        if (colIndex !== -1) {
            for (let i = 1; i < tableRows.length; i++) { // Start from 1 to skip header
                if (tableRows[i]) {
                    tableRows[i].cells[colIndex].classList.add('highlight');
                }
            }
        }
    });
}

// Event listener for Apply Operation button
document.getElementById('apply-operation').addEventListener('click', applyOperation);

// Additional download functionality (if needed)
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});

document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value;
    const format = document.getElementById('file-format').value;

    if (!filename) {
        alert('Please enter a filename.');
        return;
    }

    if (format === 'xlsx') {
        downloadAsExcel(filename);
    } else if (format === 'csv') {
        downloadAsCSV(filename);
    }
});

document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Download functions
function downloadAsExcel(filename) {
    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Data');
    XLSX.writeFile(workbook, `${filename}.xlsx`);
}

function downloadAsCSV(filename) {
    const csvContent = filteredData.map(row => Object.values(row).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.setAttribute('href', URL.createObjectURL(blob));
    link.setAttribute('download', `${filename}.csv`);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
