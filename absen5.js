let attendanceData = [];
let workbookGlobal; // Menyimpan workbook secara global
let currentSheetName = ''; // Nama sheet yang dipilih

// Fungsi untuk toggle burger menu
function toggleMenu() {
    const menu = document.getElementById('navMenu');
    menu.classList.toggle('nav-hidden');
}

// Fungsi untuk memilih sheet dan merender data
function selectSheet(sheetName) {
    currentSheetName = sheetName;
    loadSheet(workbookGlobal, currentSheetName); // Muat sheet yang dipilih
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

        // Simpan workbook di variabel global
        workbookGlobal = workbook;

        // Tampilkan nama-nama sheet di burger menu
        loadSheetNames(workbook);

        // Muat sheet pertama sebagai default
        currentSheetName = workbook.SheetNames[0];
        loadSheet(workbook, currentSheetName);
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk menampilkan nama sheet di burger menu
function loadSheetNames(workbook) {
    const sheetNames = workbook.SheetNames;
    const navMenu = document.getElementById('navMenu');
    const ul = navMenu.querySelector('ul');
    ul.innerHTML = ''; // Bersihkan daftar menu sebelum ditambahkan

    // Tambahkan nama sheet ke menu
    sheetNames.forEach(sheetName => {
        const li = document.createElement('li');
        li.innerText = sheetName;
        li.onclick = function () {
            selectSheet(sheetName); // Pilih sheet saat diklik
        };
        ul.appendChild(li); // Tambahkan nama sheet ke menu
    });
}

// Fungsi untuk memuat data dari sheet yang dipilih
function loadSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet tidak ditemukan!');
        return;
    }

    // Convert sheet ke JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    updateAttendanceData(jsonData);
    renderTable();
}

// Fungsi untuk mengupdate data absensi
function updateAttendanceData(data) {
    attendanceData = data; // Simpan data dari sheet ke variabel
}

// Fungsi untuk merender tabel
function renderTable() {
    const table = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
    table.innerHTML = ''; // Kosongkan tabel sebelum dirender ulang

    // Buat header dari baris pertama
    const headerRow = document.createElement('tr');
    attendanceData[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Tambahkan data ke tabel
    for (let i = 1; i < attendanceData.length; i++) {
        const row = document.createElement('tr');
        attendanceData[i].forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            row.appendChild(td);
        });
        table.appendChild(row);
    }
}

// Fungsi untuk update kehadiran berdasarkan input acak
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

// Fungsi untuk men-download file Excel yang sudah di-update
function exportToExcel() {
    const table = document.getElementById('attendanceTable');
    const wb = XLSX.utils.table_to_book(table, { sheet: currentSheetName });
    XLSX.writeFile(wb, 'updated_attendance.xlsx');
}
