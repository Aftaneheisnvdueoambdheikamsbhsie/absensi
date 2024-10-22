// Burger Menu Toggle
document.getElementById('burger-btn').addEventListener('click', function () {
    document.querySelector('.burger-menu').classList.toggle('active');
});

// Close menu when clicking outside
document.addEventListener('click', function (e) {
    if (!e.target.closest('.burger-menu')) {
        document.querySelector('.burger-menu').classList.remove('active');
    }
});

// Process Data Functionality
function processData() {
    const input = document.getElementById('studentInput').value;
    const month = document.getElementById('month').value;
    const date = document.getElementById('date').value;

    if (!input || !month || !date) {
        alert('Please provide student data, month, and date.');
        return;
    }

    const lines = input.trim().split('\n');
    const sortedData = {
        class3: [],
        class4: [],
        class5: [],
        class6: []
    };

    // Sort data based on class
    lines.forEach((line) => {
        const [name, classInfo] = line.split(' ');
        const classNumber = classInfo.charAt(0);
        if (classNumber === '3') sortedData.class3.push({ name, classInfo });
        if (classNumber === '4') sortedData.class4.push({ name, classInfo });
        if (classNumber === '5') sortedData.class5.push({ name, classInfo });
        if (classNumber === '6') sortedData.class6.push({ name, classInfo });
    });

    // Generate table
    displayData(sortedData, month, date);
}

// Display sorted data in a table
function displayData(sortedData, month, date) {
    let tableHTML = `<table><tr><th>No</th><th>Name</th><th>Class</th><th>${month} - ${date}</th></tr>`;
    let index = 1;

    Object.keys(sortedData).forEach((className) => {
        sortedData[className].forEach((student) => {
            tableHTML += `<tr><td>${index++}</td><td>${student.name}</td><td>${student.classInfo}</td><td>&#x2713;</td></tr>`;
        });
    });

    tableHTML += '</table>';
    document.getElementById('outputTable').innerHTML = tableHTML;
    document.getElementById('outputTable').style.display = 'block';
}

// Export to Excel
function exportToExcel() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.table_to_sheet(document.querySelector('table'));
    XLSX.utils.book_append_sheet(wb, ws, 'Attendance');
    XLSX.writeFile(wb, 'attendance.xlsx');
}
