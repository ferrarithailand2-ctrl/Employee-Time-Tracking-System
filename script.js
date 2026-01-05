// 1. DATA STORAGE
let employeeData = JSON.parse(localStorage.getItem('employeeData')) || {};
let workSchedules = JSON.parse(localStorage.getItem('workSchedules')) || {
    'PANYAPHON': { start: '08:00', mondayStart: '07:30' }
};

const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// 2. INITIALIZATION
document.addEventListener('DOMContentLoaded', () => {
    initDropdowns();
    setupEventListeners();
    refreshUI();
});

function initDropdowns() {
    const months = monthNames.map((m, i) => `<option value="${i}">${m}</option>`).join('');
    
    // Initialize all month dropdowns
    const monthDropdowns = ['monthSelect', 'employeeMonthSelect', 'rawMonthFilter', 'monthToClear'];
    monthDropdowns.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            if (id === 'rawMonthFilter') {
                element.innerHTML = '<option value="">All Months</option>' + months;
            } else {
                element.innerHTML = '<option value="">-- Select --</option>' + months;
            }
        }
    });
    
    // Set current month and year
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    // Set default values
    document.getElementById('monthSelect').value = currentMonth;
    document.getElementById('employeeMonthSelect').value = currentMonth;
    document.getElementById('yearSelect').value = currentYear;
    document.getElementById('employeeYearSelect').value = currentYear;
    document.getElementById('rawYearFilter').value = currentYear;
    document.getElementById('yearToClear').value = currentYear;
}

function saveData() {
    localStorage.setItem('employeeData', JSON.stringify(employeeData));
    localStorage.setItem('workSchedules', JSON.stringify(workSchedules));
}

function setupEventListeners() {
    document.getElementById('processBtn').addEventListener('click', processExcelFile);
    document.getElementById('saveEmpBtn').addEventListener('click', saveSchedule);
    document.getElementById('applyRawFilter').addEventListener('click', updateRawDataTable);
    
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.tab, .tab-content').forEach(el => el.classList.remove('active'));
            tab.classList.add('active');
            document.getElementById(`${tab.dataset.tab}-tab`).classList.add('active');
            if(tab.dataset.tab === 'settings') updateScheduleTable();
        });
    });

    ['reportType', 'monthSelect', 'yearSelect'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.addEventListener('change', updateSummaryTable);
    });
    
    ['employeeDetailSelect', 'employeeReportType', 'employeeMonthSelect', 'employeeYearSelect'].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.addEventListener('change', updateEmployeeDetailView);
    });

    document.getElementById('exportSummary').addEventListener('click', exportSummaryToExcel);
    document.getElementById('exportEmployee').addEventListener('click', exportEmployeeDetailToExcel);
    document.getElementById('clearData').addEventListener('click', clearAllData);
}

// 3. LOGIC (Cycle 26th-25th)
function isInCycle(dateStr, targetM, targetY) {
    const d = new Date(dateStr);
    const day = d.getDate();
    const m = d.getMonth();
    const y = d.getFullYear();
    if (day >= 26) return m === (parseInt(targetM) - 1) && y === parseInt(targetY);
    return m === parseInt(targetM) && y === parseInt(targetY);
}

function calcLate(name, date, time) {
    const s = workSchedules[name];
    if (!s || !time || time === '-') return 0;
    const start = (new Date(date).getDay() === 1) ? s.mondayStart : s.start;
    const [sh, sm] = start.split(':').map(Number);
    const [th, tm] = time.split(':').map(Number);
    const diff = (th * 60 + tm) - (sh * 60 + sm);
    return diff > 0 ? diff : 0;
}

// 4. INTERACTIVE MANUAL ENTRY
function addManualTime(name, dateStr) {
    const newTime = prompt(`Enter Clock-In time for ${name} on ${dateStr} (Format HH:mm):`);
    if (newTime === null) return;
    
    // Simple validation
    if (!/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(newTime)) {
        alert("Invalid format! Use HH:mm (e.g. 08:05)");
        return;
    }

    if (!employeeData[name]) employeeData[name] = { id: '-', entries: [] };
    
    // Remove existing entry for that day if any to avoid duplicates
    employeeData[name].entries = employeeData[name].entries.filter(e => e.date !== dateStr);
    
    employeeData[name].entries.push({ date: dateStr, time: newTime });
    saveData();
    updateEmployeeDetailView();
    updateSummaryTable();
    updateRawDataTable();
}

// 5. RENDERING
function updateEmployeeDetailView() {
    const name = document.getElementById('employeeDetailSelect').value;
    if (!name || !employeeData[name]) return;
    const m = parseInt(document.getElementById('employeeMonthSelect').value);
    const y = parseInt(document.getElementById('employeeYearSelect').value);
    const type = document.getElementById('employeeReportType').value;

    document.getElementById('employeeDetailName').textContent = name;

    // Daily Calendar Part
    const dailyBody = document.getElementById('employeeDailyBody');
    const dailyHead = document.getElementById('dailyHeader');
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    
    dailyHead.innerHTML = '<th>Row</th>' + Array.from({length: daysInMonth}, (_, i) => `<th>${i+1}</th>`).join('');
    
    const cycleEntries = employeeData[name].entries.filter(e => isInCycle(e.date, m, y));
    const dayMap = {};
    cycleEntries.forEach(e => { 
        if(!dayMap[e.date]) dayMap[e.date] = e.time; 
    });

    let rTime = '<td><strong>Time</strong></td>';
    let rLate = '<td><strong>Late</strong></td>';

    for(let i = 1; i <= daysInMonth; i++) {
        // Properly format the date to match what's in employeeData
        const monthNum = m + 1; // Convert 0-based month to 1-based
        const dStr = `${y}-${String(monthNum).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
        const time = dayMap[dStr] || '-';
        const late = calcLate(name, dStr, time);
        
        // Make time cells clickable
        rTime += `<td class="calendar-cell" onclick="addManualTime('${name}', '${dStr}')">${time}</td>`;
        rLate += `<td class="${late > 0 ? 'late-time' : ''}">${late > 0 ? late : '-'}</td>`;
    }
    dailyBody.innerHTML = `<tr>${rTime}</tr><tr>${rLate}</tr>`;

    // Monthly/Yearly Summary Part
    const sumBody = document.getElementById('employeeMonthlySummary');
    sumBody.innerHTML = '';
    if (type === 'yearly') {
        for(let i = 0; i < 12; i++) {
            const stats = getStats(name, employeeData[name].entries.filter(e => isInCycle(e.date, i, y)));
            if(stats.days > 0) sumBody.innerHTML += `<tr><td>Ending 25-${monthNames[i]}</td><td>${stats.days}</td><td>${stats.late}</td><td>${stats.mins}</td><td>${(stats.mins/60).toFixed(2)}</td></tr>`;
        }
    } else {
        const stats = getStats(name, cycleEntries);
        sumBody.innerHTML = `<tr><td>${monthNames[m]} Cycle</td><td>${stats.days}</td><td>${stats.late}</td><td>${stats.mins}</td><td>${(stats.mins/60).toFixed(2)}</td></tr>`;
    }
}

// 6. EXCEL PROCESSING
function processExcelFile() {
    const file = document.getElementById('fileInput').files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to JSON with headers
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
        
        if (jsonData.length === 0) {
            alert("Empty file or no data found!");
            return;
        }
        
        console.log("First few rows of parsed data:", jsonData.slice(0, 5));
        
        let processedCount = 0;
        let skippedCount = 0;
        
        // Process each row
        jsonData.forEach((row, index) => {
            // Get column names from the first row
            const keys = Object.keys(row);
            
            // Find the correct columns
            let name = '', date = '', time = '', empId = '-';
            
            // Look for name in different possible column names
            const nameKeys = ['ชื่อพนักงาน', 'ชื่อ', 'Name', 'Employee', 'Employee Name', 'พนักงาน'];
            for (const key of nameKeys) {
                if (row[key] !== undefined) {
                    name = String(row[key]).trim();
                    break;
                }
            }
            
            // If not found by key name, try to find by position
            if (!name) {
                // Try first column that's not empty
                for (const key of keys) {
                    const value = String(row[key] || '').trim();
                    if (value && value.length > 0 && !value.match(/^\d{1,2}[-/]\w{3,}[-/]\d{4}$/) && !value.match(/^\d{1,2}:\d{2}$/)) {
                        name = value;
                        break;
                    }
                }
            }
            
            if (!name) {
                skippedCount++;
                return;
            }
            
            // Find date - look for date-like values
            const dateKeys = ['วันที่', 'Date', 'DATE', 'วันเดือนปี'];
            for (const key of dateKeys) {
                if (row[key] !== undefined) {
                    date = row[key];
                    break;
                }
            }
            
            // If not found by key, look for date pattern
            if (!date) {
                for (const key of keys) {
                    const value = row[key];
                    if (value && (value instanceof Date || String(value).match(/\d{1,2}[-/]\w{3,}[-/]\d{4}/))) {
                        date = value;
                        break;
                    }
                }
            }
            
            // Find time - look for time-like values
            const timeKeys = ['เวลา', 'Time', 'TIME', 'Clock'];
            for (const key of timeKeys) {
                if (row[key] !== undefined) {
                    time = row[key];
                    break;
                }
            }
            
            // If not found by key, look for time pattern
            if (!time) {
                for (const key of keys) {
                    const value = row[key];
                    if (value && (value instanceof Date || String(value).match(/\d{1,2}:\d{2}/))) {
                        time = value;
                        break;
                    }
                }
            }
            
            // Parse date
            let dateObj;
            if (date instanceof Date) {
                dateObj = date;
            } else if (typeof date === 'string') {
                // Parse DD-MMM-YYYY format
                const match = date.match(/^(\d{1,2})[-/](\w{3,})[-/](\d{4})$/i);
                if (match) {
                    const day = parseInt(match[1], 10);
                    const monthStr = match[2].toLowerCase();
                    const year = parseInt(match[3], 10);
                    
                    const monthMap = {
                        'jan': 0, 'feb': 1, 'mar': 2, 'apr': 3, 'may': 4, 'jun': 5,
                        'jul': 6, 'aug': 7, 'sep': 8, 'oct': 9, 'nov': 10, 'dec': 11
                    };
                    
                    const month = monthMap[monthStr.substring(0, 3)];
                    if (month !== undefined) {
                        dateObj = new Date(year, month, day);
                    } else {
                        dateObj = new Date(date);
                    }
                } else {
                    dateObj = new Date(date);
                }
            } else if (typeof date === 'number') {
                // Excel serial date
                const excelEpoch = new Date(1899, 11, 30);
                dateObj = new Date(excelEpoch.getTime() + date * 86400000);
            }
            
            if (!dateObj || isNaN(dateObj.getTime())) {
                console.warn(`Skipping row ${index+1}: Invalid date -`, date);
                skippedCount++;
                return;
            }
            
            // Format date as YYYY-MM-DD
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const day = String(dateObj.getDate()).padStart(2, '0');
            const dStr = `${year}-${month}-${day}`;
            
            // Parse time
            let timeStr = '-';
            if (time instanceof Date) {
                const hours = String(time.getHours()).padStart(2, '0');
                const minutes = String(time.getMinutes()).padStart(2, '0');
                timeStr = `${hours}:${minutes}`;
            } else if (typeof time === 'number') {
                const totalSeconds = Math.round(time * 86400);
                const hours = Math.floor(totalSeconds / 3600);
                const minutes = Math.floor((totalSeconds % 3600) / 60);
                timeStr = `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
            } else if (typeof time === 'string') {
                timeStr = time.trim();
                if (!/^\d{1,2}:\d{2}$/.test(timeStr)) {
                    timeStr = '-';
                }
            }
            
            // Get employee ID if available
            const idKeys = ['รหัสพนักงาน', 'รหัส', 'ID', 'Employee ID', 'Emp ID'];
            for (const key of idKeys) {
                if (row[key] !== undefined) {
                    empId = String(row[key]).trim();
                    break;
                }
            }
            
            // Initialize employee data if not exists
            if (!employeeData[name]) {
                employeeData[name] = { id: empId, entries: [] };
            }
            
            // Keep the first ID we find for this employee
            if (employeeData[name].id === '-' && empId !== '-') {
                employeeData[name].id = empId;
            }
            
            // Check if we already have an entry for this date
            const existingEntry = employeeData[name].entries.find(e => e.date === dStr);
            if (!existingEntry) {
                employeeData[name].entries.push({ date: dStr, time: timeStr });
                processedCount++;
            } else if (existingEntry.time === '-' && timeStr !== '-') {
                existingEntry.time = timeStr;
                processedCount++;
            }
        });
        
        // Sort entries by date for each employee
        Object.keys(employeeData).forEach(name => {
            employeeData[name].entries.sort((a, b) => new Date(a.date) - new Date(b.date));
        });
        
        saveData();
        refreshUI();
        
        let message = `Data processed successfully!`;
        message += `\nProcessed: ${processedCount} entries`;
        message += `\nTotal employees: ${Object.keys(employeeData).length}`;
        if (skippedCount > 0) {
            message += `\nSkipped: ${skippedCount} rows (invalid or incomplete data)`;
        }
        alert(message);
    };
    reader.readAsArrayBuffer(file);
}

function getStats(name, entries) {
    const daily = {};
    entries.forEach(e => { 
        if(!daily[e.date]) daily[e.date] = e.time; 
    });
    let mins = 0, late = 0;
    Object.keys(daily).forEach(d => {
        const m = calcLate(name, d, daily[d]);
        if(m > 0) { 
            mins += m; 
            late++; 
        }
    });
    return { 
        days: Object.keys(daily).length, 
        late, 
        mins 
    };
}

function updateSummaryTable() {
    const body = document.getElementById('summaryBody');
    body.innerHTML = '';
    const type = document.getElementById('reportType').value;
    const m = parseInt(document.getElementById('monthSelect').value);
    const y = parseInt(document.getElementById('yearSelect').value);

    Object.keys(employeeData).forEach(name => {
        const entries = employeeData[name].entries.filter(e => 
            type === 'yearly' 
                ? new Date(e.date).getFullYear() == y 
                : isInCycle(e.date, m, y)
        );
        const stats = getStats(name, entries);
        const row = body.insertRow();
        row.innerHTML = `<td>${name}</td><td>${employeeData[name].id}</td><td>${stats.days}</td><td>${stats.late}</td><td>${stats.mins}</td><td>${(stats.mins/60).toFixed(2)}</td>`;
    });
}

function updateRawDataTable() {
    const body = document.getElementById('rawDataBody');
    body.innerHTML = '';
    const empF = document.getElementById('rawEmployeeFilter').value;
    const mF = document.getElementById('rawMonthFilter').value;
    const yF = parseInt(document.getElementById('rawYearFilter').value);
    
    Object.keys(employeeData).forEach(name => {
        if (empF && name !== empF) return;
        employeeData[name].entries.forEach(e => {
            const d = new Date(e.date);
            if (mF !== "" && d.getMonth() != mF) return;
            if (yF && d.getFullYear() != yF) return;
            body.innerHTML += `<tr><td>${name}</td><td>${employeeData[name].id}</td><td>${e.time}</td><td>${e.date}</td></tr>`;
        });
    });
}

function updateEmployeeSelects() {
    const options = '<option value="">-- Select --</option>' + Object.keys(employeeData).map(n => `<option value="${n}">${n}</option>`).join('');
    
    // Update employee dropdowns
    const employeeDropdowns = ['employeeDetailSelect', 'rawEmployeeFilter', 'employeeToClear'];
    employeeDropdowns.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.innerHTML = options;
        }
    });
}

function refreshUI() {
    updateSummaryTable();
    updateEmployeeSelects();
    updateRawDataTable();
}

// 7. SCHEDULE MANAGEMENT
function saveSchedule() {
    const n = document.getElementById('newEmpName').value;
    const s = document.getElementById('newEmpStart').value;
    const m = document.getElementById('newEmpMonday').value;
    if(!n || !s || !m) return alert("Fill all fields");
    workSchedules[n] = { start: s, mondayStart: m };
    saveData(); 
    updateScheduleTable(); 
    alert("Schedule saved!");
}

function updateScheduleTable() {
    const body = document.getElementById('scheduleBody');
    body.innerHTML = Object.keys(workSchedules).map(n => 
        `<tr>
            <td>${n}</td>
            <td>${workSchedules[n].start}</td>
            <td>${workSchedules[n].mondayStart}</td>
            <td><button onclick="deleteSched('${n}')" class="danger">X</button></td>
        </tr>`
    ).join('');
}

window.deleteSched = (n) => { 
    delete workSchedules[n]; 
    saveData(); 
    updateScheduleTable(); 
};

// 8. DATA MANAGEMENT FUNCTIONS
function clearAllData() {
    if(confirm("⚠️ WARNING: Are you sure you want to delete ALL data?\n\nThis will delete:\n• All employee time records\n• All employee schedules (except default)\n• All settings\n\nThis action cannot be undone!")) {
        localStorage.clear();
        employeeData = {};
        workSchedules = {
            'PANYAPHON': { start: '08:00', mondayStart: '07:30' }
        };
        alert("All data has been cleared. The page will now reload.");
        location.reload();
    }
}

// Clear specific employee data
window.clearSelectedEmployee = function() {
    const name = document.getElementById('employeeToClear').value;
    if (!name) {
        alert("Please select an employee first.");
        return;
    }
    
    if(confirm(`Are you sure you want to delete all time records for ${name}?\n\nThis will remove all clock-in data but keep their schedule.`)) {
        if (employeeData[name]) {
            // Keep the employee record but clear entries
            employeeData[name].entries = [];
            saveData();
            refreshUI();
            alert(`All time records for ${name} have been deleted.`);
        }
    }
}

// Clear specific month data for all employees
window.clearSelectedMonth = function() {
    const month = document.getElementById('monthToClear').value;
    const year = document.getElementById('yearToClear').value;
    
    if (!month || !year) {
        alert("Please select both month and year.");
        return;
    }
    
    const monthName = monthNames[month];
    if(confirm(`Are you sure you want to delete all time records for ${monthName} ${year}?\n\nThis will remove data for ALL employees for this month.`)) {
        let deletedCount = 0;
        
        Object.keys(employeeData).forEach(name => {
            const originalLength = employeeData[name].entries.length;
            employeeData[name].entries = employeeData[name].entries.filter(e => {
                const d = new Date(e.date);
                return !(d.getMonth() === parseInt(month) && d.getFullYear() === parseInt(year));
            });
            deletedCount += (originalLength - employeeData[name].entries.length);
        });
        
        saveData();
        refreshUI();
        alert(`Deleted ${deletedCount} time records for ${monthName} ${year}.`);
    }
}

// 9. EXCEL EXPORT FUNCTIONS
function exportSummaryToExcel() {
    try {
        const table = document.getElementById('summaryTable');
        if (!table || table.rows.length <= 1) {
            alert("No summary data to export!");
            return;
        }
        
        // Get report parameters
        const type = document.getElementById('reportType').value;
        const m = parseInt(document.getElementById('monthSelect').value);
        const y = parseInt(document.getElementById('yearSelect').value);
        
        // Create data array
        let data = [];
        
        // Add title and metadata
        data.push(["EMPLOYEE TIME TRACKING SUMMARY REPORT"]);
        data.push([`Report Type: ${type === 'yearly' ? 'Yearly' : 'Monthly (26th-25th)'}`]);
        data.push([`Period: ${type === 'yearly' ? `Year ${y}` : `${monthNames[m]} ${y}`}`]);
        data.push([`Generated: ${new Date().toLocaleString()}`]);
        data.push([]); // Empty row
        
        // Add table headers
        const headers = [];
        for (let i = 0; i < table.rows[0].cells.length; i++) {
            headers.push(table.rows[0].cells[i].textContent);
        }
        data.push(headers);
        
        // Add table data
        for (let i = 1; i < table.rows.length; i++) {
            const row = table.rows[i];
            const rowData = [];
            for (let j = 0; j < row.cells.length; j++) {
                rowData.push(row.cells[j].textContent);
            }
            data.push(rowData);
        }
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths
        const colWidths = [
            { wch: 25 }, // Name
            { wch: 10 }, // ID
            { wch: 12 }, // Work Days
            { wch: 12 }, // Late Days
            { wch: 15 }, // Total Mins
            { wch: 15 }  // Total Hours
        ];
        ws['!cols'] = colWidths;
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Summary Report");
        
        // Generate file name
        const fileName = `Time_Tracking_Summary_${type === 'yearly' ? `Year_${y}` : `${monthNames[m]}_${y}`}.xlsx`;
        
        // Save file
        XLSX.writeFile(wb, fileName);
        
    } catch (error) {
        console.error("Error exporting summary:", error);
        alert("Error exporting summary to Excel. Please try again.");
    }
}

function exportEmployeeDetailToExcel() {
    try {
        const name = document.getElementById('employeeDetailSelect').value;
        if (!name || !employeeData[name]) {
            alert("Please select an employee first!");
            return;
        }
        
        // Get report parameters
        const m = parseInt(document.getElementById('employeeMonthSelect').value);
        const y = parseInt(document.getElementById('employeeYearSelect').value);
        const type = document.getElementById('employeeReportType').value;
        
        // Create data array for ONLY the selected employee
        let data = [];
        
        // Add title and metadata
        data.push(["EMPLOYEE TIME TRACKING DETAILS"]);
        data.push([`Employee: ${name}`]);
        data.push([`Employee ID: ${employeeData[name].id || '-'}`]);
        data.push([`Report Type: ${type === 'yearly' ? 'Yearly' : 'Monthly (26th-25th)'}`]);
        data.push([`Period: ${type === 'yearly' ? `Year ${y}` : `${monthNames[m]} ${y}`}`]);
        data.push([`Generated: ${new Date().toLocaleString()}`]);
        data.push([]); // Empty row
        
        if (type === 'monthly') {
            // MONTHLY VIEW
            
            // Get entries for this cycle FOR THE SELECTED EMPLOYEE ONLY
            const cycleEntries = employeeData[name].entries.filter(e => isInCycle(e.date, m, y));
            
            // Create daily table
            const daysInMonth = new Date(y, m + 1, 0).getDate();
            const dayMap = {};
            cycleEntries.forEach(e => { 
                if(!dayMap[e.date]) dayMap[e.date] = e.time; 
            });
            
            // Add daily table header
            data.push(["DAILY TIME RECORD"]);
            data.push(["Date", "Day", "Clock-In Time", "Late (minutes)", "Status", "Remarks"]);
            
            // Add each day's data
            for(let i = 1; i <= daysInMonth; i++) {
                const monthNum = m + 1;
                const dStr = `${y}-${String(monthNum).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
                const dateObj = new Date(dStr);
                const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'short' });
                const displayDate = dateObj.toLocaleDateString('en-GB', { 
                    day: '2-digit', 
                    month: 'short', 
                    year: 'numeric' 
                });
                
                const time = dayMap[dStr] || '-';
                const late = calcLate(name, dStr, time);
                let status = 'On Time';
                let remarks = '';
                
                if (time === '-') {
                    status = 'No Record';
                    remarks = 'No clock-in data';
                } else if (late > 0) {
                    status = 'Late';
                    remarks = `${late} minutes late`;
                }
                
                data.push([
                    displayDate,
                    dayName,
                    time,
                    late > 0 ? late : '-',
                    status,
                    remarks
                ]);
            }
            
            data.push([]); // Empty row
            
            // Add monthly summary
            const stats = getStats(name, cycleEntries);
            data.push(["MONTHLY SUMMARY"]);
            data.push(["Metric", "Value"]);
            data.push(["Work Days", stats.days]);
            data.push(["Late Days", stats.late]);
            data.push(["Total Late Minutes", stats.mins]);
            data.push(["Total Late Hours", (stats.mins/60).toFixed(2)]);
            data.push(["Late Rate", stats.days > 0 ? `${((stats.late/stats.days)*100).toFixed(1)}%` : '0%']);
            
        } else {
            // YEARLY VIEW
            
            // Add yearly summary header
            data.push(["YEARLY SUMMARY BY MONTH"]);
            data.push(["Month", "Cycle Period", "Work Days", "Late Days", "Total Late Minutes", "Total Late Hours", "Late Rate"]);
            
            let yearlyWorkDays = 0;
            let yearlyLateDays = 0;
            let yearlyLateMins = 0;
            
            // Add monthly summaries FOR THE SELECTED EMPLOYEE ONLY
            for(let i = 0; i < 12; i++) {
                const monthlyEntries = employeeData[name].entries.filter(e => isInCycle(e.date, i, y));
                const stats = getStats(name, monthlyEntries);
                
                if(stats.days > 0) {
                    const cyclePeriod = `26 ${i === 0 ? 'Dec' : monthNames[i-1].substring(0,3)} - 25 ${monthNames[i].substring(0,3)}`;
                    const lateRate = stats.days > 0 ? `${((stats.late/stats.days)*100).toFixed(1)}%` : '0%';
                    
                    data.push([
                        monthNames[i],
                        cyclePeriod,
                        stats.days,
                        stats.late,
                        stats.mins,
                        (stats.mins/60).toFixed(2),
                        lateRate
                    ]);
                    
                    yearlyWorkDays += stats.days;
                    yearlyLateDays += stats.late;
                    yearlyLateMins += stats.mins;
                }
            }
            
            data.push([]); // Empty row
            
            // Add yearly total
            data.push(["YEARLY TOTAL"]);
            data.push(["Work Days", "Late Days", "Total Late Minutes", "Total Late Hours", "Late Rate"]);
            
            const yearlyLateRate = yearlyWorkDays > 0 ? `${((yearlyLateDays/yearlyWorkDays)*100).toFixed(1)}%` : '0%';
            data.push([
                yearlyWorkDays,
                yearlyLateDays,
                yearlyLateMins,
                (yearlyLateMins/60).toFixed(2),
                yearlyLateRate
            ]);
        }
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths based on content type
        let colWidths;
        if (type === 'monthly') {
            colWidths = [
                { wch: 15 }, // Date
                { wch: 10 }, // Day
                { wch: 15 }, // Time
                { wch: 15 }, // Late
                { wch: 12 }, // Status
                { wch: 20 }  // Remarks
            ];
        } else {
            colWidths = [
                { wch: 15 }, // Month
                { wch: 20 }, // Cycle Period
                { wch: 12 }, // Work Days
                { wch: 12 }, // Late Days
                { wch: 18 }, // Late Minutes
                { wch: 15 }, // Late Hours
                { wch: 12 }  // Late Rate
            ];
        }
        
        ws['!cols'] = colWidths;
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Employee Details");
        
        // Generate file name
        const safeName = name.replace(/[^a-zA-Z0-9ก-ฮ]/g, '_').substring(0, 30);
        const fileName = `${safeName}_${type === 'yearly' ? `Year_${y}` : `${monthNames[m].substring(0,3)}_${y}`}.xlsx`;
        
        // Save file
        XLSX.writeFile(wb, fileName);
        
        alert(`Excel file for ${name} has been exported successfully!`);
        
    } catch (error) {
        console.error("Error exporting employee details:", error);
        alert("Error exporting employee details to Excel. Please try again.");
    }
}