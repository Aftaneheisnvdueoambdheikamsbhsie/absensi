// Variabel untuk menyimpan data dan sheet yang dipilih
let attendanceData = [];
let currentSheetName = ''; // Nama sheet yang dipilih
let randomData = []; // Data acak yang akan digunakan untuk kehadiran

// Fungsi untuk membaca file Excel
document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    var files = e.target.files;
    var file = files[0];

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        // Tampilkan nama-nama sheet yang ada
        loadSheetNames(workbook);

        // Muat sheet pertama secara default
        currentSheetName = workbook.SheetNames[0];
        loadSheet(workbook, currentSheetName);
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk menampilkan daftar sheet di dalam menu
function loadSheetNames(workbook) {
    var sheetNames = workbook.SheetNames;
    var navMenu = document.getElementById('navMenu');
    var ul = navMenu.querySelector('ul');
    ul.innerHTML = ''; // Kosongkan daftar menu

    // Buat daftar sheet di menu
    sheetNames.forEach(function (sheetName) {
        var li = document.createElement('li');
        li.innerText = sheetName;
        li.onclick = function () {
            selectSheet(sheetName, workbook); // Muat sheet saat diklik
        };
        ul.appendChild(li); // Tambahkan ke daftar menu
    });
}

// Fungsi untuk memuat data dari sheet yang dipilih
function selectSheet(sheetName, workbook) {
    currentSheetName = sheetName;
    loadSheet(workbook, sheetName);
}

// Fungsi untuk memuat dan merender data sheet ke tabel HTML
function loadSheet(workbook, sheetName) {
    var sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet not found!');
        return;
    }

    // Convert sheet ke JSON
    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Update attendance data dengan data dari sheet
    updateAttendanceData(jsonData);

    // Render data ke dalam tabel HTML
    renderTable();
}

// Fungsi untuk mengupdate attendance data
function updateAttendanceData(data) {
    attendanceData = data; // Simpan data dari sheet ke variabel
}

// Fungsi untuk menampilkan data di tabel HTML
function renderTable() {
    var table = document.getElementById('attendanceTable');
    var thead = table.querySelector('thead tr');
    var tbody = table.querySelector('tbody');

    // Hapus semua baris di tabel sebelum render ulang
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Tambahkan header dari baris pertama Excel
    attendanceData[0].forEach(function (col) {
        var th = document.createElement('th');
        th.innerText = col;
        thead.appendChild(th);
    });

    // Tambahkan kolom baru untuk kehadiran
    var th = document.createElement('th');
    th.innerText = "Kehadiran";
    thead.appendChild(th);

    // Tambahkan baris data
    for (var i = 1; i < attendanceData.length; i++) {
        var row = document.createElement('tr');
        attendanceData[i].forEach(function (cell) {
            var td = document.createElement('td');
            td.innerText = cell;
            row.appendChild(td);
        });

        // Tambahkan status kehadiran
        var statusTd = document.createElement('td');
        const name = attendanceData[i][0];
        const className = attendanceData[i][1];
        const isPresent = checkAttendance(name, className);
        statusTd.innerText = isPresent ? 'P' : ''; // Gunakan karakter 'P' atau simbol Wingdings
        row.appendChild(statusTd);

        tbody.appendChild(row);
    }
}

// Fungsi untuk memeriksa kehadiran berdasarkan data acak
function checkAttendance(name, className) {
    return randomData.some(entry => entry[0] === name && entry[1] === className);
}

// Fungsi untuk mengupdate data absensi secara acak
function updateAttendance() {
    const rawData = document.getElementById('randomData').value.trim();
    const lines = rawData.split('\n');
    randomData = [];

    // Proses data acak yang diinput
    lines.forEach(line => {
        const parts = line.split(' ');
        const name = parts[0];
        const className = parts[1]; // Anggap kelas adalah kata kedua
        if (name && className) {
            randomData.push([name, className]);
        }
    });

    renderTable(); // Render tabel yang sudah diupdate
}

// Fungsi untuk men-download file Excel yang sudah di-update
function exportToExcel() {
    var table = document.getElementById('attendanceTable');
    var wb = XLSX.utils.table_to_book(table, { sheet: "Attendance" });
    XLSX.writeFile(wb, 'updated_attendance.xlsx');
}

// Fungsi untuk toggle menu
function toggleMenu() {
    const menu = document.getElementById('navMenu');
    menu.classList.toggle('nav-hidden');
}
