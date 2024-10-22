// Fungsi untuk membaca file Excel
document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    var files = e.target.files;
    var file = files[0];

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        // Ambil sheet pertama
        var firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // Convert sheet ke JSON
        var sheetData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // Render ke dalam tabel HTML
        renderTable(sheetData);
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk menampilkan data di tabel HTML
function renderTable(data) {
    var table = document.getElementById('attendanceTable');
    var thead = table.querySelector('thead tr');
    var tbody = table.querySelector('tbody');

    // Hapus semua baris di tabel sebelum render ulang
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Tambahkan header dari baris pertama Excel
    data[0].forEach(function (col) {
        var th = document.createElement('th');
        th.innerText = col;
        thead.appendChild(th);
    });

    // Tambahkan baris data
    for (var i = 1; i < data.length; i++) {
        var row = document.createElement('tr');
        data[i].forEach(function (cell) {
            var td = document.createElement('td');
            td.innerText = cell;
            row.appendChild(td);
        });
        tbody.appendChild(row);
    }
}

// Fungsi untuk men-download file Excel yang sudah di-update
function exportToExcel() {
    var table = document.getElementById('attendanceTable');
    var wb = XLSX.utils.table_to_book(table, { sheet: "Attendance" });
    XLSX.writeFile(wb, 'updated_attendance.xlsx');
}
