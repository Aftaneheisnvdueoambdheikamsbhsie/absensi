function toggleMenu() {
    const menuItems = document.getElementById('menuItems');
    menuItems.style.display = menuItems.style.display === 'block' ? 'none' : 'block';
}

document.addEventListener('click', function(event) {
    const menuItems = document.getElementById('menuItems');
    const burgerButton = document.getElementById('burgerButton');
    if (!burgerButton.contains(event.target)) {
        menuItems.style.display = 'none';
    }
});

function processData() {
    const dataInput = document.getElementById('dataInput').value;
    const date = document.getElementById('updateDate').value;
    const outputTable = document.getElementById('outputTable');
    
    if (!dataInput || !date) {
        alert("Please paste the data and select a date.");
        return;
    }

    const rows = dataInput.split('\n').filter(row => row.trim() !== '');
    const sortedData = rows.map(row => {
        const [name, classInfo] = row.split(' ').reverse();
        return { name: row.split(' ').slice(0, -1).join(' '), classInfo };
    }).sort((a, b) => a.classInfo.localeCompare(b.classInfo));

    let table = `
        <table>
            <thead>
                <tr>
                    <th>No.</th>
                    <th>Name</th>
                    <th>Class</th>
                    <th>Month</th>
                    <th>Date</th>
                    <th>Checkmark</th>
                </tr>
            </thead>
            <tbody>
    `;

    sortedData.forEach((item, index) => {
        table += `
            <tr>
                <td>${index + 1}</td>
                <td>${item.name}</td>
                <td>${item.classInfo}</td>
                <td>${new Date(date).toLocaleString('default', { month: 'long' })}</td>
                <td>${new Date(date).getDate()}</td>
                <td class="checkmark">p</td>
            </tr>
        `;
    });

    table += `</tbody></table>`;
    outputTable.innerHTML = table;
}
