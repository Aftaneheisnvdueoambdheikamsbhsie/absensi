// Menangani tampilan menu burger
document.getElementById('burgerMenu').addEventListener('click', function() {
    const sideMenu = document.getElementById('sideMenu');
    sideMenu.style.display = sideMenu.style.display === 'block' ? 'none' : 'block';
});

// Fungsi untuk memproses data absensi
function prosesAbsensi() {
    const inputText = document.getElementById('inputNama').value;
    const lines = inputText.split('\n');
    
    // Array untuk menyimpan data per kelas
    const dataKelas = {
        '3': [],
        '4': [],
        '5': [],
        '6': []
    };

    // Memproses setiap baris input
    lines.forEach(function(line) {
        const [nama, kelas] = line.trim().split(' ');
        if (dataKelas[kelas[0]]) {
            dataKelas[kelas[0]].push({ nama, kelas });
        }
    });

    // Menampilkan data di tabel HTML
    tampilkanData(dataKelas);
}

// Fungsi untuk menampilkan data di tabel HTML
function tampilkanData(dataKelas) {
    Object.keys(dataKelas).forEach(function(kelas) {
        const table = document.getElementById(`tableClass${kelas}`);
        table.innerHTML = ''; // Menghapus isi tabel sebelumnya

        // Header tabel
        let headerRow = '<tr><th>Tanggal</th><th>Nama</th><th>Status</th></tr>';
        table.innerHTML += headerRow;

        // Menambahkan data ke tabel
        dataKelas[kelas].forEach(function(item) {
            const status = '<span style="font-family: Wingdings 2;">p</span>'; // Ceklis
            const row = `<tr><td>${getCurrentDate()}</td><td>${item.nama}</td><td>${status}</td></tr>`;
            table.innerHTML += row;
        });
    });
}

// Fungsi untuk mendapatkan tanggal saat ini
function getCurrentDate() {
    const date = new Date();
    return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`; // Format DD/MM/YYYY
}

// Fungsi untuk mengekspor data ke Excel
function exportToExcel() {
    const workbook = XLSX.utils.book_new();

    // Mengambil data dari setiap tabel dan menambahkannya ke workbook
    for (let kelas = 3; kelas <= 6; kelas++) {
        const table = document.getElementById(`tableClass${kelas}`);
        const worksheet = XLSX.utils.table_to_sheet(table);
        XLSX.utils.book_append_sheet(workbook, worksheet, `Class ${kelas}`);
    }

    // Menyimpan file Excel
    XLSX.writeFile(workbook, "attendance.xlsx");
}

// Menambahkan event listener untuk tombol proses dan ekspor
document.getElementById('btnProses').addEventListener('click', prosesAbsensi);
document.getElementById('btnExport').addEventListener('click', exportToExcel);
