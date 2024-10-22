let attendanceData = {
    Kelas3: [],
    Kelas4: [],
    Kelas5: [],
    Kelas6: [],
}; // Memisahkan data berdasarkan kelas
let currentSheetName = 'Kelas3'; // Default sheet
let sheetNames = []; // Array untuk menyimpan nama sheet

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
    const files = e.target.files;
    const file = files[0];

    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Ambil semua nama sheet secara dinamis
        sheetNames = workbook.SheetNames;
        populateSheetMenu(sheetNames); // Memperbarui menu dengan nama sheet

        // Muat sheet default atau sheet pertama
        loadSheet(workbook, sheetNames[0]); 
    };
    reader.readAsArrayBuffer(file);
}

// Fungsi untuk memuat sheet berdasarkan nama
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

    // Tampilkan menu setelah diupdate
    toggleMenu();
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

    // Tambahkan data baru ke attendanceData sesuai kelasnya
    newEntries.forEach(row => {
        const classKey = row[2]; // Mendapatkan kelas dari data yang baru

        if (attendanceData[classKey]) {
            const existingIndex = attendanceData[classKey].findIndex(entry => entry[1] === row[1]);
            if (existingIndex !== -1) {
                attendanceData[classKey][existingIndex] = row; // Update data yang sudah ada
            } else {
                attendanceData[classKey].push(row); // Tambahkan data baru
            }
        }
    });

    renderTable(); // Render ulang tabel
}

// Fungsi untuk merender tabel
function renderTable() {
    const table = document.getElementById('attendanceTable').getElementsByTagName('tbody')[0];
    table.innerHTML = ''; // Kosongkan tabel sebelum merender ulang

    // Mendapatkan data dari kelas yang dipilih
    const dataToRender = attendanceData[currentSheetName] || [];

    dataToRender.forEach((row, index) => {
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
