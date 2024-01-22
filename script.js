document.addEventListener('DOMContentLoaded', function() {
    createSpreadsheet(10, 10); 

    document.getElementById('spreadsheet').addEventListener('click', function(event) {
        if (event.target.tagName === 'TD') {
            highlightHeaders(event.target);
            updateCurrentCellAddress(event.target);
        }
    });

    document.getElementById('exportBtn').addEventListener('click', exportToExcel);
});

function createSpreadsheet(rows, cols) {
    const table = document.getElementById('spreadsheet');
    let headerRow = table.insertRow();
    for (let i = 0; i <= cols; i++) {
        let cell = headerRow.insertCell();
        cell.textContent = i === 0 ? '' : String.fromCharCode('A'.charCodeAt(0) + i - 1);
    }

    for (let i = 1; i <= rows; i++) {
        let row = table.insertRow();
        for (let j = 0; j <= cols; j++) {
            let cell = row.insertCell();
            if (j === 0) {
                cell.textContent = i;
                cell.classList.add('header');
            } else {
                cell.contentEditable = true;
            }
        }
    }
}

function highlightHeaders(cell) {
    resetHighlight();
    const rowIndex = cell.parentNode.rowIndex;
    const colIndex = cell.cellIndex;
    const table = document.getElementById('spreadsheet');
    table.rows[rowIndex].cells[0].classList.add('highlight');
    table.rows[0].cells[colIndex].classList.add('highlight');
}

function resetHighlight() {
    document.querySelectorAll('.highlight').forEach(cell => cell.classList.remove('highlight'));
}

function updateCurrentCellAddress(cell) {
    const rowIndex = cell.parentNode.rowIndex;
    const colIndex = cell.cellIndex;
    const columnName = String.fromCharCode('A'.charCodeAt(0) + colIndex - 1);
    document.getElementById('currentCell').textContent = `Cell: ${columnName}${rowIndex}`;
}

// function exportToExcel() {
//     let workbook = XLSX.utils.book_new();
//     let worksheet = XLSX.utils.table_to_sheet(document.getElementById('spreadsheet'));
//     XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//     XLSX.writeFile(workbook, 'spreadsheet.xlsx');
// }
//---> 헤더행 헤더열이 저장할 때 포함되는 문제 생김

function exportToExcel() {
    let workbook = XLSX.utils.book_new();
    let table = document.getElementById('spreadsheet');
    let data = [];

    // 테이블의 각 행을 순회합니다.
    for (let i = 1; i < table.rows.length; i++) { // 첫 번째 행(헤더)는 건너뜁니다.
        let rowData = [];
        for (let j = 1; j < table.rows[i].cells.length; j++) { // 첫 번째 열(헤더)는 건너뜁니다.
            rowData.push(table.rows[i].cells[j].innerText);
        }
        data.push(rowData);
    }

    let worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, 'spreadsheet.xlsx');
}
