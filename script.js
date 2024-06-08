
let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0],
tableExists = false;

const generateTable = () => {
let rowsNumber = parseInt(rows.value),
    columnsNumber = parseInt(columns.value);

if (isNaN(rowsNumber) || isNaN(columnsNumber) || rowsNumber <= 0 || columnsNumber <= 0) {
    Swal.fire({
        title: 'Error!',
        text: 'Please enter valid number of rows and columns',
        icon: 'error',
        confirmButtonText: 'OK'
    });
    return;
}

table.innerHTML = "";
for (let i = 0; i < rowsNumber; i++) {
    var tableRow = "<tr>";
    for (let j = 0; j < columnsNumber; j++) {
        tableRow += `<td contenteditable></td>`;
    }
    tableRow += "</tr>";
    table.innerHTML += tableRow;
}
tableExists = true;
}

const ExportToExcel = (type, fn, dl) => {
if (!tableExists) {
    Swal.fire({
        title: 'Error!',
        text: 'No table generated to export',
        icon: 'error',
        confirmButtonText: 'OK'
    });
    return;
}
var elt = table;
var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
return dl ?
    XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' }) :
    XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')));
}