document.getElementById('inputForm').addEventListener('submit', function(event) {
  event.preventDefault();

  const dataInput = document.getElementById('dataInput').value.trim();
  const attendanceDate = document.getElementById('attendanceDate').value;
  
  if (!dataInput || !attendanceDate) {
    alert('Please enter student data and attendance date.');
    return;
  }

  const students = dataInput.split('\n').map(row => {
    const parts = row.trim().split(' ');
    const className = parts.pop();
    const name = parts.join(' ');
    return { name, className };
  });

  students.sort((a, b) => a.className.localeCompare(b.className));

  const workbook = XLSX.utils.book_new();

  const groupedByClass = students.reduce((acc, curr) => {
    if (!acc[curr.className]) acc[curr.className] = [];
    acc[curr.className].push(curr.name);
    return acc;
  }, {});

  for (let className in groupedByClass) {
    let sheetData = [];
    const classStudents = groupedByClass[className];

    sheetData.push(['No', 'Name', 'Class', 'Attendance']);
    classStudents.forEach((student, index) => {
      sheetData.push([index + 1, student, className, 'p']); // 'p' for Wingdings checkmark
    });

    let sheet = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(workbook, sheet, className);
  }

  XLSX.writeFile(workbook, `Attendance_${attendanceDate}.xlsx`);
  document.getElementById('outputMessage').innerText = 'File successfully generated!';
});
