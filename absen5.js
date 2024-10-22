let attendanceData = [];
let currentSheetName = 'Kelas3'; // Default sheet

// Fungsi untuk toggle menu
function toggleMenu() {
    const menu = document.getElementById('navMenu');
    menu.classList.toggle('nav-hidden');
}

// Fungsi untuk memilih sheet
function selectSheet(sheetName) {
    currentSheetName = sheetName;
    renderTable(); // Render ulang tabel untuk sheet yang dipilih
}

// Fungsi untuk membaca file Excel
document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    var files = e.target.files;
    var file = files[0];

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        // Ambil sheet berdasarkan pilihan
        loadSheet(workbook, currentSheetName);
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk memuat sheet berdasarkan kelas
function loadSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet tidak ditemukan!');
        return;
    }

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    updateAttendanceData(jsonData);
    renderTable();
}

// Fungsi untuk mengupdate data kehadiran berdasarkan input acak
function updateAttendance() {
    const rawData = document.getElementById('randomData').value.trim();
    const lines = rawData.split('\n');
    const newEntries = [];

    lines.forEach(line => {
        const parts = line.split(' ');
        const name = parts.slice(1, parts.length - 1).join(' '); // Menggabungkan nama
        const className = parts[parts.length - 1]; // Bagian terakhir adalah kelas

        if (name && className) {
            newEntries.push([parts[0], name, className, 'P']); // Tanda ceklis Windings 2
        }
    });

    // Tambahkan data baru ke attendanceData
    newEntries.forEach(row => {
        const existingIndex = attendanceData.findIndex(entry => entry[1] === row[1]);
        if (existingIndex !== -1) {
            attendanceData[existingIndex] = row; // Update data yang sudah ada
        } else {
            attendanceData.push(row); // Tambahkan data baru
        }
    });

    renderTable(); // Render ulang tabel
}

// Fungsi untuk merender tabel
function renderTable() {
    const table = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
    table.innerHTML = ''; // Kosongkan tabel sebelum merender ulang

    attendanceData
        .filter(row => row[2].startsWith(currentSheetName.slice(5))) // Hanya tampilkan data untuk kelas yang dipilih
        .forEach(row => {
            const newRow = table.insertRow();
            row.forEach(cell => {
                const newCell = newRow.insertCell();
                newCell.textContent = cell;
            });
        });
}

// Fungsi untuk men-download file Excel yang sudah di-update
function exportToExcel() {
    const table = document.getElementById('attendanceTable');
    const wb = XLSX.utils.table_to_book(table, { sheet: currentSheetName });
    XLSX.writeFile(wb, 'updated_attendance.xlsx');
}
