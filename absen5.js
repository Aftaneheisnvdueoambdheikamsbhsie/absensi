// Variabel untuk menyimpan data dan sheet yang dipilih
let attendanceData = [];
let currentSheetName = ''; // Nama sheet yang dipilih
let currentWorkbook = null; // Variabel global untuk menyimpan workbook yang diupload

// Fungsi untuk membaca file Excel
document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    var files = e.target.files;
    var file = files[0];

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });

        // Tampilkan nama-nama sheet yang ada
        loadSheetNames(currentWorkbook);

        // Muat sheet pertama secara default
        currentSheetName = currentWorkbook.SheetNames[0];
        loadSheet(currentWorkbook, currentSheetName);
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

// Fungsi untuk memuat dan merender data sheet ke tabel HTML dengan mempertimbangkan merge cell
function loadSheet(workbook, sheetName) {
    var sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        alert('Sheet tidak ditemukan!');
        return;
    }

    // Ambil data sheet dan informasi merge
    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' }); // Tambahkan defval untuk menangani sel kosong
    var merges = sheet['!merges'] || []; // Ambil informasi merge jika ada

    // Update attendance data dengan data dari sheet
    updateAttendanceData(jsonData);

    // Render data ke dalam tabel HTML dengan merge cell jika perlu
    renderTableWithMerge(jsonData, merges);
}

// Fungsi untuk menampilkan data di tabel HTML dengan merge cell
function renderTableWithMerge(data, merges) {
    var table = document.getElementById('attendanceTable');
    var thead = table.querySelector('thead');
    var tbody = table.querySelector('tbody');

    // Hapus semua baris di tabel sebelum render ulang
    thead.innerHTML = '';
    tbody.innerHTML = '';

    // Buat elemen baris header
    var headerRow = document.createElement('tr');
    data[0].forEach(function (col) {
        var th = document.createElement('th');
        th.innerText = col;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // Render data tabel dan terapkan merge cell
    for (var i = 1; i < data.length; i++) {
        var row = document.createElement('tr');
        data[i].forEach(function (cell, colIndex) {
            var td = document.createElement('td');
            td.innerText = cell;

            // Cek apakah sel ini merupakan bagian dari merge
            merges.forEach(function (merge) {
                if (merge.s.r === i && merge.s.c === colIndex) {
                    td.setAttribute('rowspan', merge.e.r - merge.s.r + 1);
                    td.setAttribute('colspan', merge.e.c - merge.s.c + 1);
                }
            });

            row.appendChild(td);
        });
        tbody.appendChild(row);
    }
}

// Fungsi untuk mencocokkan string secara fuzzy
function fuzzyMatch(studentName, className) {
    return studentName.toLowerCase().includes(className.toLowerCase()) || className.toLowerCase().includes(studentName.toLowerCase());
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
        // Memanggil fuzzyMatch dengan nama dan kelas yang sesuai
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

    renderTableWithMerge([]); // Render ulang tabel setelah update
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
