document.getElementById('inputForm').addEventListener('submit', function(event) {
  event.preventDefault();
  const dataInput = document.getElementById('dataInput').value.trim();
  if (!dataInput) return;

  const rows = dataInput.split('\n');
  const students = [];

  rows.forEach(row => {
    const parts = row.split(' ');
    const className = parts.pop(); // Last part is the class
    const name = parts.join(' '); // Rest is the name
    students.push({ name, className });
  });

  students.sort((a, b) => a.className.localeCompare(b.className));

  const tableBody = document.querySelector('#resultTable tbody');
  tableBody.innerHTML = '';

  students.forEach(student => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${student.name}</td>
      <td>${student.className}</td>
      <td><input type="checkbox" checked></td>
    `;
    tableBody.appendChild(row);
  });

  document.getElementById('tableContainer').style.display = 'block';
});

document.getElementById('downloadBtn').addEventListener('click', function() {
  const table = document.getElementById('resultTable');
  const rows = table.querySelectorAll('tr');
  let csvContent = 'data:text/csv;charset=utf-8,';
  
  rows.forEach(row => {
    const cells = row.querySelectorAll('th, td');
    const rowData = [];
    cells.forEach(cell => rowData.push(cell.innerText || cell.querySelector('input')?.checked ? '✔' : ''));
    csvContent += rowData.join(',') + '\n';
  });

  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', 'attendance.csv');
  document.body.appendChild(link);

  link.click();
  document.body.removeChild(link);
});
