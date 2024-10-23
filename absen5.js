// Variabel untuk menyimpan data dan sheet yang dipilih
let attendanceData = [];
let currentSheetName = ''; // Nama sheet yang dipilih

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
// Fungsi untuk memuat dan merender data sheet ke tabel HTML dengan mempertimbangkan merge cell
function loadSheet(workbook, sheetName) {
    var sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet tidak ditemukan!');
        return;
    }

    // Ambil data sheet dan informasi merge
    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    var merges = sheet['!merges'] || []; // Ambil informasi merge jika ada

    // Update attendance data dengan data dari sheet
    updateAttendanceData(jsonData);

    // Render data ke dalam tabel HTML dengan merge cell
    renderTableWithMerge(merges);
}

// Fungsi untuk menampilkan data di tabel HTML dengan merge cell
function renderTableWithMerge(merges) {
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

    // Tambahkan baris data dan terapkan merge
    for (var i = 1; i < attendanceData.length; i++) {
        var row = document.createElement('tr');
        attendanceData[i].forEach(function (cell, index) {
            var td = document.createElement('td');
            td.innerText = cell;

            // Cek apakah kolom ini merupakan bagian dari merge
            merges.forEach(function (merge) {
                if (merge.s.r === i && merge.s.c === index) {
                    td.setAttribute('rowspan', merge.e.r - merge.s.r + 1);
                    td.setAttribute('colspan', merge.e.c - merge.s.c + 1);
                }
            });

            row.appendChild(td);
        });
        tbody.appendChild(row);
    }
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

    // Tambahkan baris data
    for (var i = 1; i < attendanceData.length; i++) {
        var row = document.createElement('tr');
        attendanceData[i].forEach(function (cell) {
            var td = document.createElement('td');
            td.innerText = cell;
            row.appendChild(td);
        });
        tbody.appendChild(row);
    }
}

// Fungsi untuk mencocokkan string secara fuzzy
function fuzzyMatch(str1, str2) {
    return str1.toLowerCase().includes(str2.toLowerCase()) || str2.toLowerCase().includes(str1.toLowerCase());
}

// Fungsi untuk mengupdate data absensi secara acak dengan pencocokan nama dan kelas
document.getElementById('updateAttendanceBtn').addEventListener('click', updateAttendance);

function updateAttendance() {
    console.log('Update Attendance button clicked');
    const rawData = document.getElementById('randomData').value.trim();
    if (!rawData) {
        alert("Tidak ada data yang dimasukkan!");
        return;
    }
    const lines = rawData.split('\n');
    const newEntries = [];

    // Proses data acak yang diinput
    lines.forEach(line => {
        const parts = line.trim().split(/\s+/);
        if (parts.length < 2) {
            console.error('Data tidak lengkap:', line);
            return;
        }
        const className = parts.pop(); // Kelas selalu di bagian akhir
        const name = parts.join(' '); // Gabungkan kembali nama siswa yang terpecah
        if (name && className) {
            newEntries.push([name, className]);
        }
    });

    if (newEntries.length === 0) {
        alert('Data yang dimasukkan tidak valid');
        return;
    }

    updateAttendanceDataWithNewEntries(newEntries);
}

function updateAttendanceDataWithNewEntries(newEntries) {
    const currentDate = new Date().toLocaleDateString(); // Mendapatkan tanggal hari ini

    newEntries.forEach(row => {
        let updated = false;
        for (let i = 1; i < attendanceData.length; i++) {
            if (fuzzyMatch(attendanceData[i][0], row[0]) && fuzzyMatch(attendanceData[i][1], row[1])) {
                attendanceData[i].push('P', currentDate);
                updated = true;
                break;
            }
        }
        if (!updated) {
            let newRow = Array(attendanceData[0].length - 3).fill(''); 
            newRow.unshift(row[0], row[1], 'P', currentDate); // Tambahkan tanggal
            attendanceData.push(newRow);
        }
    });

    renderTable();
    alert('Update attendance berhasil!');

    if (confirm('Ingin mendownload file terupdate?')) {
        exportToExcel();
    }
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
