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
let attendanceData = [];
let currentSheetName = 'Sheet1'; // Default sheet

// Toggle menu
function toggleMenu() {
    const menu = document.getElementById('navMenu');
    menu.classList.toggle('nav-hidden');
}

// Select sheet
function selectSheet(sheetName) {
    currentSheetName = sheetName;
    // Logic to load the selected sheet data (optional)
    alert(`Switched to ${sheetName}`);
}

// Read Excel file
document.getElementById('upload').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        loadSheet(workbook, currentSheetName);
    };

    reader.readAsArrayBuffer(file);
});

// Load selected sheet data
function loadSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet not found!');
        return;
    }

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    updateAttendanceData(jsonData);
    renderTable();
}

// Update attendance data with random data
function updateAttendance() {
    const rawData = document.getElementById('randomData').value.trim();
    const lines = rawData.split('\n');
    const newEntries = [];

    // Process the random data
    lines.forEach(line => {
        const parts = line.split(' ');
        const name = parts[0];
        const className = parts[1]; // Assuming class is the second word
        if (name && className) {
            newEntries.push([name, className]);
        }
    });

    // Add new entries to attendance data
    newEntries.forEach(row => {
        const existingIndex = attendanceData.findIndex(entry => entry[0] === row[0]);
        if (existingIndex !== -1) {
            attendanceData[existingIndex] = row; // Update existing entry
        } else {
            attendanceData.push(row); // Add new entry
        }
    });

    renderTable(); // Render updated table
}

// Other existing functions (updateAttendanceData, renderTable, exportToExcel)...
