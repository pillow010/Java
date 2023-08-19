function searchLogs() {
    const startDate = document.getElementById('start-date').value;
    const endDate = document.getElementById('end-date').value;
    const logs = retrieveLogs(startDate, endDate);
    populateLogsTable(logs);
    populateSummaryTable(logs);
}

function retrieveLogs(startDate, endDate) {
    // Generate dummy data for logs within the selected date range
    const logs = [
        {
            Doctor: 'Beriman',
            successSign: 100,
            failedSign: 2,
            Total: 102
        },
        {
            Doctor: 'Dadang',
            successSign: 200,
            failedSign: 10,
            Total: 210
        },
        {
            Doctor: 'Rama',
            successSign: 50,
            failedSign: 5,
            Total: 55
        },
        // Add more log entries here based on the selected date range
    ];

    return logs; // Return the generated logs array
}

function populateLogsTable(logs) {
    const logsContainer = document.getElementById('logs-container');

    // Clear previous logs
    logsContainer.innerHTML = '';

    logs.forEach((log, index) => {
        const logEntry = document.createElement('tr');
        logEntry.innerHTML = `
            <td>${index + 1}</td>
            <td>${log.Doctor}</td>
            <td>${log.successSign}</td>
            <td>${log.failedSign}</td>
            <td>${log.Total}</td>
        `;
        logsContainer.appendChild(logEntry);
    });
}


function populateSummaryTable(logs) {
    const summaryContainer = document.getElementById('summary-container');

    // Calculate the total success, total fail, and total count (accumulation)
    let totalSuccess = 0;
    let totalFail = 0;
    let totalCount = 0;

    logs.forEach((log) => {
        totalSuccess += log.successSign;
        totalFail += log.failedSign;
        totalCount += log.Total;
    });

    // Display the summary data in the summary table
    summaryContainer.innerHTML = ''; // Clear previous data

    const summaryRow = document.createElement('tr');
    summaryRow.innerHTML = `
        <td>${totalSuccess}</td>
        <td>${totalFail}</td>
        <td>${totalCount}</td>
    `;
    summaryContainer.appendChild(summaryRow);
}
