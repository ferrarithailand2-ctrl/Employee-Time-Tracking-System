// Store the processed data
let employeeData = {};
let manualEntries = [];
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();

// Define work schedules
const workSchedules = {
    'PANYAPHON': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'ALINE': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'JIRAPAT': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'KANATIP': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'Silikul': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'EITHUZARSOE': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'PONGSAKORN': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'MANEERAT': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'SARAWUT': { start: '08:00', end: '17:00', mondayStart: '07:30', mondayEnd: '16:30' },
    'PORNWANDEE': { start: '08:30', end: '17:30', mondayStart: '08:30', mondayEnd: '17:30' },
    'CHATNARIN': { start: '08:30', end: '17:30', mondayStart: '08:30', mondayEnd: '17:30' },
    'CHANIDA': { start: '08:30', end: '17:30', mondayStart: '08:30', mondayEnd: '17:30' }
};

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Initialize date selects
    document.getElementById('monthSelect').value = currentMonth;
    document.getElementById('employeeMonthSelect').value = currentMonth;
    document.getElementById('yearSelect').value = currentYear;
    document.getElementById('employeeYearSelect').value = currentYear;
    
    // Initialize with today's date for manual entry
    document.getElementById('manualDate').valueAsDate = new Date();
    
    // Set up event listeners
    setupEventListeners();
}

function setupEventListeners() {
    // Process Excel file
    document.getElementById('processBtn').addEventListener('click', processExcelFile);
    
    // Manual entry
    document.getElementById('addManualEntry').addEventListener('click', addManualEntryHandler);
    
    // Export buttons
    document.getElementById('exportSummary').addEventListener('click', exportSummary);
    document.getElementById('exportEmployee').addEventListener('click', exportEmployeeData);
    
    // Tab functionality - FIXED
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', handleTabClick);
    });
    
    // Report controls
    document.getElementById('reportType').addEventListener('change', updateSummaryTable);
    document.getElementById('monthSelect').addEventListener('change', updateSummaryTable);
    document.getElementById('yearSelect').addEventListener('change', updateSummaryTable);
    
    document.getElementById('employeeReportType').addEventListener('change', updateEmployeeDetailView);
    document.getElementById('employeeMonthSelect').addEventListener('change', updateEmployeeDetailView);
    document.getElementById('employeeYearSelect').addEventListener('change', updateEmployeeDetailView);
    
    // Confirmation dialog
    document.getElementById('cancelConfirm').addEventListener('click', function() {
        document.getElementById('confirmationDialog').style.display = 'none';
    });
}

// Process the uploaded Excel file
function processExcelFile() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) {
        alert('Please select an Excel file first.');
        return;
    }
    
    const file = fileInput.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Assuming the first sheet contains the data
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // Process the data (skip header row)
            processEmployeeData(jsonData.slice(1));
            
            // Update the UI
            updateSummaryTable();
            updateRawDataTable();
            updateEmployeeSelect();
            updateEmployeeDetailSelect();
            
            alert('Data processed successfully!');
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing file. Please make sure it\'s a valid Excel file.');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Process employee data from Excel
function processEmployeeData(data) {
    employeeData = {};
    
    data.forEach(row => {
        if (row.length < 5) return; // Skip incomplete rows
        
        const name = row[0];
        const equipment = row[1];
        const id = row[2];
        const time = row[3];
        const date = row[4];
        
        if (!name || !time || !date) return; // Skip rows with missing data
        
        // Initialize employee if not exists
        if (!employeeData[name]) {
            employeeData[name] = {
                id: id,
                equipment: equipment,
                entries: []
            };
        }
        
        // Add entry
        employeeData[name].entries.push({
            date: date,
            time: time
        });
    });
    
    // Sort entries by date and time for each employee
    Object.keys(employeeData).forEach(name => {
        employeeData[name].entries.sort((a, b) => {
            const dateA = new Date(a.date);
            const dateB = new Date(b.date);
            
            if (dateA.getTime() !== dateB.getTime()) {
                return dateA - dateB;
            }
            
            // If same date, sort by time
            const timeA = a.time.split(':').map(Number);
            const timeB = b.time.split(':').map(Number);
            
            return (timeA[0] * 60 + timeA[1]) - (timeB[0] * 60 + timeB[1]);
        });
    });
}

// Calculate late time for an employee on a specific date
function calculateLateTime(employeeName, date, time) {
    // Check if this is a manual entry with confirmed status
    const manualEntry = manualEntries.find(entry => 
        entry.employeeName === employeeName && 
        entry.date === date && 
        entry.time === time
    );
    
    // If it's a manual entry marked as "not late", return 0
    if (manualEntry && manualEntry.confirmed && !manualEntry.isLate) {
        return 0;
    }
    
    // Otherwise, calculate normally
    const schedule = workSchedules[employeeName];
    if (!schedule) return 0;
    
    // Check if it's Monday
    const dayOfWeek = new Date(date).getDay();
    const isMonday = dayOfWeek === 1;
    
    const startTime = isMonday ? schedule.mondayStart : schedule.start;
    
    // Convert times to minutes for comparison
    const [startHour, startMinute] = startTime.split(':').map(Number);
    const [timeHour, timeMinute] = time.split(':').map(Number);
    
    const startTotalMinutes = startHour * 60 + startMinute;
    const timeTotalMinutes = timeHour * 60 + timeMinute;
    
    if (timeTotalMinutes > startTotalMinutes) {
        return timeTotalMinutes - startTotalMinutes;
    }
    
    return 0;
}

// Check if time is in the afternoon (after 12:00)
function isAfternoon(time) {
    const [hour] = time.split(':').map(Number);
    return hour >= 12;
}

// Update the summary table
function updateSummaryTable() {
    const summaryBody = document.getElementById('summaryBody');
    summaryBody.innerHTML = '';
    
    const reportType = document.getElementById('reportType').value;
    const month = parseInt(document.getElementById('monthSelect').value);
    const year = parseInt(document.getElementById('yearSelect').value);
    
    Object.keys(employeeData).forEach(name => {
        const employee = employeeData[name];
        
        // Filter entries based on report type
        let filteredEntries = [...employee.entries];
        let filteredManualEntries = manualEntries.filter(entry => entry.employeeName === name);
        
        if (reportType === 'monthly') {
            filteredEntries = filteredEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getMonth() === month && entryDate.getFullYear() === year;
            });
            
            filteredManualEntries = filteredManualEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getMonth() === month && entryDate.getFullYear() === year;
            });
        } else if (reportType === 'yearly') {
            filteredEntries = filteredEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getFullYear() === year;
            });
            
            filteredManualEntries = filteredManualEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getFullYear() === year;
            });
        }
        
        // Group entries by date
        const entriesByDate = {};
        filteredEntries.forEach(entry => {
            if (!entriesByDate[entry.date]) {
                entriesByDate[entry.date] = [];
            }
            entriesByDate[entry.date].push(entry);
        });
        
        // Add manual entries
        filteredManualEntries.forEach(entry => {
            if (!entriesByDate[entry.date]) {
                entriesByDate[entry.date] = [];
            }
            entriesByDate[entry.date].push({
                date: entry.date,
                time: entry.time,
                isManual: true
            });
        });
        
        // Calculate statistics
        let totalLateMinutes = 0;
        let lateDays = 0;
        let workDays = Object.keys(entriesByDate).length;
        
        Object.keys(entriesByDate).forEach(date => {
            // Find the first arrival time for the day
            const firstEntry = entriesByDate[date][0];
            const lateMinutes = calculateLateTime(name, date, firstEntry.time);
            
            if (lateMinutes > 0) {
                totalLateMinutes += lateMinutes;
                lateDays++;
            }
        });
        
        // Convert to hours
        const totalLateHours = (totalLateMinutes / 60).toFixed(2);
        
        // Create table row
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${name}</td>
            <td>${employee.id}</td>
            <td>${workDays}</td>
            <td>${lateDays}</td>
            <td class="late-time">${totalLateMinutes}</td>
            <td class="late-time">${totalLateHours}</td>
        `;
        
        summaryBody.appendChild(row);
    });
}

// Update employee detail view
function updateEmployeeDetailView() {
    const employeeName = document.getElementById('employeeDetailSelect').value;
    if (!employeeName) {
        document.getElementById('employeeDetailContent').innerHTML = '<p>Please select an employee to view details.</p>';
        return;
    }
    
    const reportType = document.getElementById('employeeReportType').value;
    const month = parseInt(document.getElementById('employeeMonthSelect').value);
    const year = parseInt(document.getElementById('employeeYearSelect').value);
    
    // Update employee name in header
    document.getElementById('employeeDetailName').textContent = `${employeeName} - Details`;
    
    // Update monthly summary
    updateEmployeeMonthlySummary(employeeName, reportType, month, year);
    
    // Update daily details
    updateEmployeeDailyDetails(employeeName, reportType, month, year);
}

// Update employee monthly summary
function updateEmployeeMonthlySummary(employeeName, reportType, month, year) {
    const monthlySummaryBody = document.getElementById('employeeMonthlySummary');
    monthlySummaryBody.innerHTML = '';
    
    if (reportType === 'yearly') {
        // Show monthly breakdown for the year
        for (let m = 0; m < 12; m++) {
            const monthEntries = getEmployeeEntriesForPeriod(employeeName, 'monthly', m, year);
            const monthStats = calculateEmployeeStats(employeeName, monthEntries);
            
            const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                               'July', 'August', 'September', 'October', 'November', 'December'];
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${monthNames[m]} ${year}</td>
                <td>${monthStats.workDays}</td>
                <td>${monthStats.lateDays}</td>
                <td class="late-time">${monthStats.totalLateMinutes}</td>
                <td class="late-time">${monthStats.totalLateHours}</td>
            `;
            
            monthlySummaryBody.appendChild(row);
        }
    } else {
        // Show just the selected month
        const monthEntries = getEmployeeEntriesForPeriod(employeeName, 'monthly', month, year);
        const monthStats = calculateEmployeeStats(employeeName, monthEntries);
        
        const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                           'July', 'August', 'September', 'October', 'November', 'December'];
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${monthNames[month]} ${year}</td>
            <td>${monthStats.workDays}</td>
            <td>${monthStats.lateDays}</td>
            <td class="late-time">${monthStats.totalLateMinutes}</td>
            <td class="late-time">${monthStats.totalLateHours}</td>
        `;
        
        monthlySummaryBody.appendChild(row);
    }
}

// Update employee daily details
function updateEmployeeDailyDetails(employeeName, reportType, month, year) {
    const dailyTable = document.getElementById('employeeDailyTable');
    const dailyBody = document.getElementById('employeeDailyBody');
    
    // Clear previous content
    dailyBody.innerHTML = '';
    
    // Get days in the month
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    // Create header row with days
    let headerRow = '<tr><th class="calendar-day">Date</th>';
    for (let day = 1; day <= daysInMonth; day++) {
        headerRow += `<th class="calendar-day">${day}</th>`;
    }
    headerRow += '</tr>';
    dailyTable.querySelector('thead').innerHTML = headerRow;
    
    // Get ALL entries for the employee
    const allEntries = [...employeeData[employeeName].entries];
    const manualEntriesForEmployee = manualEntries.filter(entry => entry.employeeName === employeeName);
    
    // Combine with manual entries
    const combinedEntries = [...allEntries];
    manualEntriesForEmployee.forEach(entry => {
        combinedEntries.push({
            date: entry.date,
            time: entry.time,
            isManual: true
        });
    });
    
    // Group entries by date
    const entriesByDate = {};
    combinedEntries.forEach(entry => {
        const entryDate = new Date(entry.date);
        // Only include entries from the selected month and year
        if (entryDate.getMonth() === month && entryDate.getFullYear() === year) {
            const dateKey = entryDate.toISOString().split('T')[0]; // YYYY-MM-DD format
            if (!entriesByDate[dateKey]) {
                entriesByDate[dateKey] = [];
            }
            entriesByDate[dateKey].push(entry);
        }
    });
    
    // Sort entries within each date by time
    Object.keys(entriesByDate).forEach(date => {
        entriesByDate[date].sort((a, b) => {
            const timeA = a.time.split(':').map(Number);
            const timeB = b.time.split(':').map(Number);
            return (timeA[0] * 60 + timeA[1]) - (timeB[0] * 60 + timeB[1]);
        });
    });
    
    // Create rows for different data
    const firstArrivalRow = document.createElement('tr');
    firstArrivalRow.innerHTML = '<td>First Arrival Time</td>';
    
    const lateMinutesRow = document.createElement('tr');
    lateMinutesRow.innerHTML = '<td>Late Minutes</td>';
    
    const statusRow = document.createElement('tr');
    statusRow.innerHTML = '<td>Status</td>';
    
    // Fill data for each day
    for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${(month + 1).toString().padStart(2, '0')}-${day.toString().padStart(2, '0')}`;
        const dayEntries = entriesByDate[dateStr] || [];
        
        if (dayEntries.length > 0) {
            const firstEntry = dayEntries[0];
            
            // Check if this is a manual entry with special status
            const isManualEntry = firstEntry.isManual;
            const manualEntryData = isManualEntry ? 
                manualEntries.find(entry => 
                    entry.employeeName === employeeName && 
                    entry.date === dateStr && 
                    entry.time === firstEntry.time
                ) : null;
            
            const lateMinutes = manualEntryData && manualEntryData.confirmed && !manualEntryData.isLate 
                ? 0 
                : calculateLateTime(employeeName, dateStr, firstEntry.time);
            
            const isAfternoonArrival = isAfternoon(firstEntry.time);
            
            // First arrival time cell
            let firstArrivalCell = document.createElement('td');
            firstArrivalCell.textContent = firstEntry.time;
            
            if (isManualEntry) {
                if (manualEntryData.entryType === 'missing') {
                    firstArrivalCell.innerHTML += ' <span class="status-badge status-confirmed">Missing Clock In</span>';
                } else {
                    firstArrivalCell.innerHTML += ' <span class="status-badge">Manual</span>';
                }
            }
            
            firstArrivalRow.appendChild(firstArrivalCell);
            
            // Late minutes cell
            let lateMinutesCell = document.createElement('td');
            if (lateMinutes > 0) {
                lateMinutesCell.textContent = lateMinutes;
                lateMinutesCell.classList.add('late-time');
            } else {
                lateMinutesCell.textContent = manualEntryData && manualEntryData.confirmed && !manualEntryData.isLate 
                    ? '0 (Confirmed)' 
                    : '0';
            }
            lateMinutesRow.appendChild(lateMinutesCell);
            
            // Status cell
            let statusCell = document.createElement('td');
            if (lateMinutes > 0) {
                statusCell.innerHTML = '<span class="status-badge status-late">Late</span>';
            } else {
                if (manualEntryData && manualEntryData.confirmed && !manualEntryData.isLate) {
                    statusCell.innerHTML = '<span class="status-badge status-confirmed">Work-Related</span>';
                } else {
                    statusCell.innerHTML = '<span class="status-badge status-ontime">On Time</span>';
                }
            }
            statusRow.appendChild(statusCell);
        } else {
            // No record for this day
            const noRecordCell = document.createElement('td');
            noRecordCell.textContent = 'No record';
            noRecordCell.classList.add('no-record');
            firstArrivalRow.appendChild(noRecordCell);
            
            const emptyCell = document.createElement('td');
            emptyCell.textContent = '-';
            lateMinutesRow.appendChild(emptyCell);
            
            const statusCell = document.createElement('td');
            statusCell.innerHTML = '<span class="status-badge">Absent?</span>';
            statusRow.appendChild(statusCell);
        }
    }
    
    dailyBody.appendChild(firstArrivalRow);
    dailyBody.appendChild(lateMinutesRow);
    dailyBody.appendChild(statusRow);
}

// Get employee entries for a specific period
function getEmployeeEntriesForPeriod(employeeName, reportType, month, year) {
    let entries = [...employeeData[employeeName].entries];
    let manualEntriesForEmployee = manualEntries.filter(entry => entry.employeeName === employeeName);
    
    if (reportType === 'monthly') {
        entries = entries.filter(entry => {
            const entryDate = new Date(entry.date);
            return entryDate.getMonth() === month && entryDate.getFullYear() === year;
        });
        
        manualEntriesForEmployee = manualEntriesForEmployee.filter(entry => {
            const entryDate = new Date(entry.date);
            return entryDate.getMonth() === month && entryDate.getFullYear() === year;
        });
    } else if (reportType === 'yearly') {
        entries = entries.filter(entry => {
            const entryDate = new Date(entry.date);
            return entryDate.getFullYear() === year;
        });
        
        manualEntriesForEmployee = manualEntriesForEmployee.filter(entry => {
            const entryDate = new Date(entry.date);
            return entryDate.getFullYear() === year;
        });
    }
    
    // Combine with manual entries and mark them
    const combinedEntries = [...entries];
    manualEntriesForEmployee.forEach(entry => {
        combinedEntries.push({
            date: entry.date,
            time: entry.time,
            isManual: true
        });
    });
    
    // Sort by date
    combinedEntries.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    return combinedEntries;
}

// Calculate employee statistics
function calculateEmployeeStats(employeeName, entries) {
    // Group entries by date
    const entriesByDate = {};
    entries.forEach(entry => {
        if (!entriesByDate[entry.date]) {
            entriesByDate[entry.date] = [];
        }
        entriesByDate[entry.date].push(entry);
    });
    
    // Calculate statistics
    let totalLateMinutes = 0;
    let lateDays = 0;
    let workDays = Object.keys(entriesByDate).length;
    
    Object.keys(entriesByDate).forEach(date => {
        // Find the first arrival time for the day
        const firstEntry = entriesByDate[date][0];
        const lateMinutes = calculateLateTime(employeeName, date, firstEntry.time);
        
        if (lateMinutes > 0) {
            totalLateMinutes += lateMinutes;
            lateDays++;
        }
    });
    
    const totalLateHours = (totalLateMinutes / 60).toFixed(2);
    
    return {
        workDays,
        lateDays,
        totalLateMinutes,
        totalLateHours
    };
}

// Update the raw data table
function updateRawDataTable() {
    const rawDataBody = document.getElementById('rawDataBody');
    rawDataBody.innerHTML = '';
    
    Object.keys(employeeData).forEach(name => {
        const employee = employeeData[name];
        
        employee.entries.forEach(entry => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${name}</td>
                <td>${employee.equipment}</td>
                <td>${employee.id}</td>
                <td>${entry.time}</td>
                <td>${entry.date}</td>
            `;
            
            rawDataBody.appendChild(row);
        });
    });
}

// Update employee select for manual entry
function updateEmployeeSelect() {
    const employeeSelect = document.getElementById('employeeSelect');
    employeeSelect.innerHTML = '<option value="">Select Employee</option>';
    
    Object.keys(employeeData).forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        employeeSelect.appendChild(option);
    });
}

// Update employee detail select
function updateEmployeeDetailSelect() {
    const employeeDetailSelect = document.getElementById('employeeDetailSelect');
    employeeDetailSelect.innerHTML = '<option value="">-- Select Employee --</option>';
    
    Object.keys(employeeData).forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        employeeDetailSelect.appendChild(option);
    });
    
    // Add event listener
    employeeDetailSelect.addEventListener('change', updateEmployeeDetailView);
}

// Manual entry handler
function addManualEntryHandler() {
    const employeeName = document.getElementById('employeeSelect').value;
    const date = document.getElementById('manualDate').value;
    const time = document.getElementById('manualTime').value;
    const entryType = document.querySelector('input[name="entryType"]:checked').value;
    
    if (!employeeName || !date || !time) {
        alert('Please fill in all fields.');
        return;
    }

    const isAfternoonArrival = isAfternoon(time);
    const lateMinutes = calculateLateTime(employeeName, date, time);
    
    // For work-related late clock ins, show confirmation dialog
    if (isAfternoonArrival && entryType === 'work-related') {
        showConfirmationDialog(employeeName, date, time, lateMinutes);
    } else {
        // For missing clock in entries, add directly
        addManualEntry(employeeName, date, time, 'missing', lateMinutes > 0);
    }
}

// Show confirmation dialog for work-related late arrivals
function showConfirmationDialog(employeeName, date, time, lateMinutes) {
    const dialog = document.getElementById('confirmationDialog');
    const confirmationText = document.getElementById('confirmationText');
    
    confirmationText.innerHTML = `
        <strong>${employeeName}</strong> - ${date} at ${time}<br><br>
        This is a late clock in (${lateMinutes} minutes after scheduled start time).<br>
        Was this a work-related arrival (not late) or should it be counted as late?
    `;
    
    dialog.style.display = 'flex';
    
    // Set up event listeners for confirmation buttons
    document.getElementById('confirmLate').onclick = function() {
        addManualEntry(employeeName, date, time, 'work-related', true);
        dialog.style.display = 'none';
    };
    
    document.getElementById('confirmNotLate').onclick = function() {
        addManualEntry(employeeName, date, time, 'work-related', false);
        dialog.style.display = 'none';
    };
}

// Add manual entry with status
function addManualEntry(employeeName, date, time, entryType, isLate) {
    manualEntries.push({
        employeeName: employeeName,
        date: date,
        time: time,
        entryType: entryType,
        isLate: isLate,
        confirmed: true
    });
    
    updateManualEntriesTable();
    updateSummaryTable();
    
    if (document.getElementById('employeeDetailSelect').value === employeeName) {
        updateEmployeeDetailView();
    }
    
    // Clear the form
    document.getElementById('manualDate').value = '';
    document.getElementById('manualTime').value = '';
}

// Enhanced manual entries table
function updateManualEntriesTable() {
    const manualEntriesBody = document.getElementById('manualEntriesBody');
    manualEntriesBody.innerHTML = '';
    
    manualEntries.forEach((entry, index) => {
        const row = document.createElement('tr');
        
        let typeBadge = entry.entryType === 'missing' 
            ? '<span class="status-badge status-confirm">Missing Clock In Time</span>'
            : '<span class="status-badge">Late Clock In due to Work Related</span>';
        
        let lateStatus = entry.isLate 
            ? '<span class="status-badge status-late">Late</span>'
            : '<span class="status-badge status-ontime">Not Late</span>';
        
        row.innerHTML = `
            <td>${entry.employeeName}</td>
            <td>${entry.date}</td>
            <td>${entry.time}</td>
            <td>${typeBadge}</td>
            <td>${lateStatus}</td>
            <td>
                <button class="toggle-late" data-index="${index}">Toggle Late</button>
                <button class="delete-entry" data-index="${index}">Delete</button>
            </td>
        `;
        
        manualEntriesBody.appendChild(row);
    });
    
    // Add event listeners
    document.querySelectorAll('.delete-entry').forEach(button => {
        button.addEventListener('click', function() {
            const index = parseInt(this.getAttribute('data-index'));
            const employeeName = manualEntries[index].employeeName;
            manualEntries.splice(index, 1);
            updateManualEntriesTable();
            updateSummaryTable();
            if (document.getElementById('employeeDetailSelect').value === employeeName) {
                updateEmployeeDetailView();
            }
        });
    });
    
    document.querySelectorAll('.toggle-late').forEach(button => {
        button.addEventListener('click', function() {
            const index = parseInt(this.getAttribute('data-index'));
            manualEntries[index].isLate = !manualEntries[index].isLate;
            updateManualEntriesTable();
            updateSummaryTable();
            if (document.getElementById('employeeDetailSelect').value === manualEntries[index].employeeName) {
                updateEmployeeDetailView();
            }
        });
    });
}

// Tab functionality
function handleTabClick() {
    // Remove active class from all tabs and contents
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    // Add active class to clicked tab and corresponding content
    this.classList.add('active');
    const tabId = this.getAttribute('data-tab') + '-tab';
    document.getElementById(tabId).classList.add('active');
    
    // If switching to employee tab, update the view if an employee is selected
    if (tabId === 'employee-tab' && document.getElementById('employeeDetailSelect').value) {
        updateEmployeeDetailView();
    }
}

// Export summary data
function exportSummary() {
    const reportType = document.getElementById('reportType').value;
    const month = parseInt(document.getElementById('monthSelect').value);
    const year = parseInt(document.getElementById('yearSelect').value);
    
    let csvContent = "Employee Name,Employee ID,Work Days,Late Days,Total Late (Minutes),Total Late (Hours)\n";
    
    Object.keys(employeeData).forEach(name => {
        const employee = employeeData[name];
        
        // Filter entries based on report type
        let filteredEntries = [...employee.entries];
        let filteredManualEntries = manualEntries.filter(entry => entry.employeeName === name);
        
        if (reportType === 'monthly') {
            filteredEntries = filteredEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getMonth() === month && entryDate.getFullYear() === year;
            });
            
            filteredManualEntries = filteredManualEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getMonth() === month && entryDate.getFullYear() === year;
            });
        } else if (reportType === 'yearly') {
            filteredEntries = filteredEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getFullYear() === year;
            });
            
            filteredManualEntries = filteredManualEntries.filter(entry => {
                const entryDate = new Date(entry.date);
                return entryDate.getFullYear() === year;
            });
        }
        
        // Group entries by date
        const entriesByDate = {};
        filteredEntries.forEach(entry => {
            if (!entriesByDate[entry.date]) {
                entriesByDate[entry.date] = [];
            }
            entriesByDate[entry.date].push(entry);
        });
        
        // Add manual entries
        filteredManualEntries.forEach(entry => {
            if (!entriesByDate[entry.date]) {
                entriesByDate[entry.date] = [];
            }
            entriesByDate[entry.date].push({
                date: entry.date,
                time: entry.time,
                isManual: true
            });
        });
        
        // Calculate statistics
        let totalLateMinutes = 0;
        let lateDays = 0;
        let workDays = Object.keys(entriesByDate).length;
        
        Object.keys(entriesByDate).forEach(date => {
            const firstEntry = entriesByDate[date][0];
            const lateMinutes = calculateLateTime(name, date, firstEntry.time);
            
            if (lateMinutes > 0) {
                totalLateMinutes += lateMinutes;
                lateDays++;
            }
        });
        
        const totalLateHours = (totalLateMinutes / 60).toFixed(2);
        
        csvContent += `"${name}","${employee.id}",${workDays},${lateDays},${totalLateMinutes},${totalLateHours}\n`;
    });
    
    downloadCSV(csvContent, `employee_summary_${reportType}_${month + 1}_${year}.csv`);
}

// Export employee data
function exportEmployeeData() {
    const employeeName = document.getElementById('employeeDetailSelect').value;
    if (!employeeName) {
        alert('Please select an employee first.');
        return;
    }
    
    const reportType = document.getElementById('employeeReportType').value;
    const month = parseInt(document.getElementById('employeeMonthSelect').value);
    const year = parseInt(document.getElementById('employeeYearSelect').value);
    
    let csvContent = "Date,First Arrival Time,Late Minutes,Status,Entry Type\n";
    
    const entries = getEmployeeEntriesForPeriod(employeeName, reportType, month, year);
    
    // Group entries by date
    const entriesByDate = {};
    entries.forEach(entry => {
        if (!entriesByDate[entry.date]) {
            entriesByDate[entry.date] = [];
        }
        entriesByDate[entry.date].push(entry);
    });
    
    Object.keys(entriesByDate).forEach(date => {
        const firstEntry = entriesByDate[date][0];
        const lateMinutes = calculateLateTime(employeeName, date, firstEntry.time);
        
        let status = lateMinutes > 0 ? 'Late' : 'On Time';
        let entryType = firstEntry.isManual ? 'Manual Entry' : 'Auto Recorded';
        
        csvContent += `"${date}","${firstEntry.time}",${lateMinutes},"${status}","${entryType}"\n`;
    });
    
    downloadCSV(csvContent, `${employeeName}_details_${reportType}_${month + 1}_${year}.csv`);
}

// Helper function to download CSV
function downloadCSV(csvContent, filename) {
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}