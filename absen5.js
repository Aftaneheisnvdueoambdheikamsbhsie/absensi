// Variabel untuk menyimpan data dan sheet yang dipilih
let attendanceData = [];
let currentSheetName = ''; 
let currentWorkbook = null;

// Fungsi untuk membaca file Excel
document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    var files = e.target.files;
    var file = files[0];

    var reader = new FileReader();
    reader.onload = function (event) {
        var data = new Uint8Array(event.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });
        loadSheetNames(currentWorkbook);

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
    ul.innerHTML = '';

    sheetNames.forEach(function (sheetName) {
        var li = document.createElement('li');
        li.innerText = sheetName;
        li.onclick = function () {
            selectSheet(sheetName, workbook);
        };
        ul.appendChild(li);
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

    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    var merges = sheet['!merges'] || [];

    attendanceData = jsonData;
    renderTableWithMerge(jsonData, merges);
}

// Fungsi untuk menampilkan data di tabel HTML dengan merge cell
function renderTableWithMerge(data, merges) {
    var table = document.getElementById('attendanceTable');
    var thead = table.querySelector('thead');
    var tbody = table.querySelector('tbody');

    thead.innerHTML = '';
    tbody.innerHTML = '';

    var headerRow = document.createElement('tr');
    data[0].forEach(function (col, colIndex) {
        var th = document.createElement('th');
        th.innerText = col;
        headerRow.appendChild(th);

        merges.forEach(function (merge) {
            if (merge.s.r === 1 && merge.s.c === colIndex) {
                th.setAttribute('rowspan', merge.e.r - merge.s.r + 1);
                th.setAttribute('colspan', merge.e.c - merge.s.c + 1);
            }
        });
    });
    thead.appendChild(headerRow);

    for (var i = 1; i < data.length; i++) {
        var row = document.createElement('tr');
        data[i].forEach(function (cell, colIndex) {
            var td = document.createElement('td');
            td.innerText = cell;

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

// Fungsi Levenshtein Distance untuk mencocokkan nama dan kelas
function levenshteinDistance(a, b) {
    const matrix = Array.from({ length: a.length + 1 }, (_, i) =>
        Array.from({ length: b.length + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
    );

    for (let i = 1; i <= a.length; i++) {
        for (let j = 1; j <= b.length; j++) {
            if (a[i - 1] === b[j - 1]) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j] + 1,
                    matrix[i][j - 1] + 1,
                    matrix[i - 1][j - 1] + 1
                );
            }
        }
    }
    return matrix[a.length][b.length];
}

// Fungsi untuk mencocokkan nama dan kelas berdasarkan Levenshtein Distance
function levenshteinMatch(studentName, className, inputName, inputClass) {
    const nameDistance = levenshteinDistance(studentName.toLowerCase(), inputName.toLowerCase());
    const classDistance = levenshteinDistance(className.toLowerCase(), inputClass.toLowerCase());
    
    const nameThreshold = 3; 
    const classThreshold = 1;

    return nameDistance <= nameThreshold && classDistance <= classThreshold;
}

// Fungsi untuk memperbarui data absensi berdasarkan pencocokan nama dan kelas
document.getElementById('updateAttendanceBtn').addEventListener('click', updateAttendance);

function updateAttendance() {
    const rawData = document.getElementById('randomData').value.trim();
    if (!rawData) {
        alert("Tidak ada data yang dimasukkan!");
        return;
    }

    const lines = rawData.split('\n');
    const newEntries = [];

    lines.forEach(line => {
        const parts = line.trim().split(/\s+/);
        if (parts.length < 2) {
            console.error('Data tidak lengkap:', line);
            return;
        }
        const className = parts.pop(); 
        const name = parts.join(' '); 
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

// Fungsi untuk menambahkan kolom tanggal dan memperbarui "P" di baris yang cocok
function updateAttendanceDataWithNewEntries(newEntries) {
    const currentDate = document.getElementById('datePicker').value;
    if (!currentDate) {
        alert("Tanggal tidak valid! Silakan pilih tanggal.");
        return;
    }

    let dateIndex = attendanceData[0].indexOf(currentDate);

    if (dateIndex === -1) {
        attendanceData[0].push(currentDate);
        dateIndex = attendanceData[0].length - 1;
        attendanceData.forEach((row, rowIndex) => {
            if (rowIndex > 0) row.push('');
        });
    }

    newEntries.forEach(row => {
        for (let i = 1; i < attendanceData.length; i++) {
            if (levenshteinMatch(attendanceData[i][0], attendanceData[i][1], row[0], row[1])) {
                attendanceData[i][dateIndex] = 'P';
            }
        }
    });

    renderTableWithMerge(attendanceData, []); 
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
