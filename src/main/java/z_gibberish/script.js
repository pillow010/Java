function searchLogs() {
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;
    // Assuming you have a function that retrieves logs from the server/database
    // and populates them in the logsContainer tbody of the logs-table
    // Replace this with your actual implementation to retrieve logs
    retrieveLogs(startDate, endDate);
}

function retrieveLogs(startDate, endDate) {
    // Generate dummy data for logs within the selected date range
    const logs = [
        {
            noreg: '0201xxxx',
            name: 'Satria',
            rm: '00269766',
            resume: 'Ada Belum Final',
            cppt: 'Ada',
            operasi: 'Tidak Ada',
            laboratorium: 'Ada',
            radiologi: 'Ada',
            dpjp: 'Dadang Herdiana dr. Sp.Pd.',
            ruangan: 'Lily',
        },
        {
            noreg: '0301xxxx',
            name: 'Bagus',
            rm: '00269766',
            resume: 'Tidak Ada',
            cppt: 'Ada',
            operasi: 'Tidak Ada',
            laboratorium: 'Ada',
            radiologi: 'Ada',
            dpjp: 'Anindya Anggara L dr.',
            ruangan: 'Lily',
        }
        // Add more log entries here based on the selected date range
    ];

    // Display the logs in the logs-container tbody
    const logsContainer = document.getElementById('logs-container');
    logsContainer.innerHTML = ''; // Clear previous logs

    logs.forEach((log, index) => {
        const logEntry = document.createElement('tr');
        logEntry.innerHTML = `
      <td>${index + 1}</td>
      <td>${log.noreg}</td>
      <td>${log.name}</td>
      <td>${log.rm}</td>
      <td>${log.resume}</td>
      <td>${log.cppt}</td>
      <td>${log.operasi}</td>
      <td>${log.laboratorium}</td>
      <td>${log.radiologi}</td>
      <td>${log.dpjp}</td>
      <td>${log.ruangan}</td>
    `;
        logsContainer.appendChild(logEntry);
    });
}
