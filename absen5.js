document.addEventListener("DOMContentLoaded", function () {
    let attendanceData = [];
    let currentSheetName = ''; 
    let currentWorkbook = null; 

    document.getElementById('upload').addEventListener('change', handleFile, false);

    function handleFile(e) {
        const files = e.target.files;
        const file = files[0];

        const reader = new FileReader();
        reader.onload = function (event) {
            const data = new Uint8Array(event.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });

            loadSheetNames(currentWorkbook);
            currentSheetName = currentWorkbook.SheetNames[0];
            loadSheet(currentWorkbook, currentSheetName);
        };
        reader.readAsArrayBuffer(file);
    }

    function loadSheetNames(workbook) {
        const sheetNames = workbook.SheetNames;
        const navMenu = document.getElementById('navMenu');
        const ul = navMenu.querySelector('ul');
        ul.innerHTML = '';

        sheetNames.forEach(function (sheetName) {
            const li = document.createElement('li');
            li.innerText = sheetName;
            li.onclick = function () {
                selectSheet(sheetName, workbook);
            };
            ul.appendChild(li);
        });
    }

    function selectSheet(sheetName, workbook) {
        currentSheetName = sheetName;
        loadSheet(workbook, sheetName);
    }

    function loadSheet(workbook, sheetName) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) {
            alert('Sheet tidak ditemukan!');
            return;
        }

        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
        const merges = sheet['!merges'] || [];

        attendanceData = jsonData; 
        renderTableWithMerge(jsonData, merges); 
    }

    function renderTableWithMerge(data, merges) {
        const table = document.getElementById('attendanceTable');
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');

        thead.innerHTML = '';
        tbody.innerHTML = '';

        const headerRow = document.createElement('tr');
        data[0].forEach(function (col, colIndex) {
            const th = document.createElement('th');
            th.innerText = col;
            headerRow.appendChild(th);

            merges.forEach(function (merge) {
                if (merge.s.r === 0 && merge.s.c === colIndex) {
                    th.setAttribute('rowspan', merge.e.r - merge.s.r + 1);
                    th.setAttribute('colspan', merge.e.c - merge.s.c + 1);
                }
            });
        });
        thead.appendChild(headerRow);

        for (let i = 1; i < data.length; i++) {
            const row = document.createElement('tr');
            data[i].forEach(function (cell, colIndex) {
                const td = document.createElement('td');
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

    function normalizeString(str) {
        return str.trim().toLowerCase(); 
    }

    function fuzzyMatch(studentName, className) {
        const normalizedName = normalizeString(studentName);
        const normalizedClassName = normalizeString(className);
        return normalizedName.includes(normalizedClassName);
    }

    document.getElementById('updateAttendanceBtn').addEventListener('click', function () {
        updateAttendance();
    });

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
                console.warn('Data tidak lengkap:', line); 
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
            let updated = false;
            for (let i = 1; i < attendanceData.length; i++) { 
                if (fuzzyMatch(attendanceData[i][0], row[0]) && fuzzyMatch(attendanceData[i][1], row[1])) {
                    attendanceData[i][dateIndex] = 'P'; 
                    updated = true;
                    break;
                }
            }

            if (!updated) {
                console.log("Tidak ditemukan kecocokan untuk:", row);
            }
        });

        renderTableWithMerge(attendanceData, []);
    }

    function exportToExcel() {
        if (!currentWorkbook) {
            alert("Tidak ada file yang diunggah!");
            return;
        }

        const sheet = XLSX.utils.aoa_to_sheet(attendanceData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, sheet, currentSheetName);

        XLSX.writeFile(newWorkbook, 'updated_attendance.xlsx');
    }

    function toggleMenu() {
        const navMenu = document.getElementById('navMenu');
        navMenu.classList.toggle('nav-hidden');
    }

    document.querySelector('.burger-menu').addEventListener('click', toggleMenu);
});
