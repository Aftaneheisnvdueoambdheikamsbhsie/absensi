let attendanceData = [];
let currentSheetName = ' '; // Default sheet

// Fungsi untuk toggle menu
function toggleMenu() {
    const menu = document.getElementById('navMenu');
    menu.classList.toggle('nav-hidden');
}

// Fungsi untuk memilih sheet
function selectSheet(sheetName) {
    currentSheetName = sheetName;
    loadSheet(currentWorkbook, currentSheetName); // Muat sheet yang dipilih
}

// Fungsi untuk membaca file Excel
let currentWorkbook; // Menyimpan workbook saat ini

document.getElementById('upload').addEventListener('change', handleFile, false);

function handleFile(e) {
    const files = e.target.files;
    const file = files[0];

    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        currentWorkbook = XLSX.read(data, { type: 'array' });

        // Ambil semua nama sheet secara dinamis
        const sheetNames = currentWorkbook.SheetNames;
        populateSheetMenu(sheetNames); // Memperbarui menu dengan nama sheet

        // Muat sheet default atau sheet pertama
        if (sheetNames.length > 0) {
            currentSheetName = sheetNames[0]; // Set sheet default ke sheet pertama
            loadSheet(currentWorkbook, currentSheetName);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Memperbarui menu dengan nama sheet yang ada
function populateSheetMenu(sheetNames) {
    const navMenu = document.getElementById('navMenu');
    navMenu.innerHTML = ''; // Kosongkan menu

    sheetNames.forEach(sheetName => {
        const li = document.createElement('li');
        li.textContent = sheetName;
        li.onclick = () => selectSheet(sheetName); // Mengatur fungsi onclick
        navMenu.appendChild(li);
    });

    toggleMenu(); // Tampilkan menu setelah diupdate
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
    const dateInput = document.getElementById('inputDate').value;
    
    if (!dateInput) {
        alert('Silakan pilih tanggal!');
        return;
    }

    const dateParts = dateInput.split('-'); // Memecah tanggal menjadi [YYYY, MM, DD]
    const newEntries = [];
    const lines = rawData.split('\n');

    lines.forEach(line => {
        const parts = line.split(' ');
        const name = parts.slice(1, parts.length - 1).join(' ');
        const className = parts[parts.length - 1];

        if (name && className) {
            newEntries.push([parts[0], name, className, 'P', dateParts[1], dateParts[2]]); // Menambahkan bulan dan tanggal
        }
    });

    // Tambahkan data baru ke attendanceData
    newEntries.forEach(row => {
        const existingIndex = attendanceData.findIndex(entry => entry[1] === row[1] && entry[2] === row[2]);
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
        .forEach((row, index) => {
            const newRow = table.insertRow();
            newRow.insertCell().textContent = index + 1; // Nomor
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
