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
    ['monthSelect', 'employeeMonthSelect', 'rawMonthFilter', 'monthToClear'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = (id === 'rawMonthFilter' || id === 'monthToClear' ? '<option value="">All Months</option>' : '') + months;
    });
    const now = new Date();
    document.getElementById('monthSelect').value = now.getMonth();
    document.getElementById('employeeMonthSelect').value = now.getMonth();
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

    ['reportType', 'monthSelect', 'yearSelect'].forEach(id => document.getElementById(id).addEventListener('change', updateSummaryTable));
    ['employeeDetailSelect', 'employeeReportType', 'employeeMonthSelect', 'employeeYearSelect'].forEach(id => document.getElementById(id).addEventListener('change', updateEmployeeDetailView));

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
    
    if (!/^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(newTime)) {
        alert("Invalid format! Use HH:mm (e.g. 08:05)");
        return;
    }

    if (!employeeData[name]) employeeData[name] = { id: '-', entries: [] };
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
    const dailyBody = document.getElementById('employeeDailyBody');
    const dailyHead = document.getElementById('dailyHeader');
    const daysInMonth = new Date(y, m + 1, 0).getDate();
    
    dailyHead.innerHTML = '<th>Row</th>' + Array.from({length: daysInMonth}, (_, i) => `<th>${i+1}</th>`).join('');
    
    const cycleEntries = employeeData[name].entries.filter(e => isInCycle(e.date, m, y));
    const dayMap = {};
    cycleEntries.forEach(e => { if(!dayMap[e.date]) dayMap[e.date] = e.time; });

    let rTime = '<td><strong>Time</strong></td>';
    let rLate = '<td><strong>Late</strong></td>';

    for(let i = 1; i <= daysInMonth; i++) {
        const monthNum = m + 1;
        const dStr = `${y}-${String(monthNum).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
        const time = dayMap[dStr] || '-';
        const late = calcLate(name, dStr, time);
        rTime += `<td class="calendar-cell" onclick="addManualTime('${name}', '${dStr}')">${time}</td>`;
        rLate += `<td class="${late > 0 ? 'late-time' : ''}">${late > 0 ? late : '-'}</td>`;
    }
    dailyBody.innerHTML = `<tr>${rTime}</tr><tr>${rLate}</tr>`;

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

function processExcelFile() {
    const file = document.getElementById('fileInput').files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
        
        jsonData.forEach(row => {
            let name = row['ชื่อพนักงาน'] || row['Name'] || row['Employee'];
            let date = row['วันที่'] || row['Date'];
            let time = row['เวลา'] || row['Time'];
            let empId = row['รหัสพนักงาน'] || row['ID'] || '-';

            if (!name || !date) return;

            // Simplified date parsing for standard Excel formats
            let dObj = new Date(date);
            const dStr = dObj.toISOString().split('T')[0];
            
            let timeStr = '-';
            if (time instanceof Date) {
                timeStr = `${String(time.getHours()).padStart(2, '0')}:${String(time.getMinutes()).padStart(2, '0')}`;
            }

            if (!employeeData[name]) employeeData[name] = { id: empId, entries: [] };
            if (!employeeData[name].entries.find(e => e.date === dStr)) {
                employeeData[name].entries.push({ date: dStr, time: timeStr });
            }
        });
        saveData();
        refreshUI();
        alert("Data processed successfully!");
    };
    reader.readAsArrayBuffer(file);
}

function getStats(name, entries) {
    let mins = 0, late = 0;
    entries.forEach(e => {
        const m = calcLate(name, e.date, e.time);
        if(m > 0) { mins += m; late++; }
    });
    return { days: entries.length, late, mins };
}

function updateSummaryTable() {
    const body = document.getElementById('summaryBody');
    if (!body) return;
    body.innerHTML = '';
    const type = document.getElementById('reportType').value;
    const m = parseInt(document.getElementById('monthSelect').value);
    const y = parseInt(document.getElementById('yearSelect').value);

    Object.keys(employeeData).forEach(name => {
        const entries = employeeData[name].entries.filter(e => 
            type === 'yearly' ? new Date(e.date).getFullYear() == y : isInCycle(e.date, m, y)
        );
        const stats = getStats(name, entries);
        const row = body.insertRow();
        row.innerHTML = `<td>${name}</td><td>${employeeData[name].id}</td><td>${stats.days}</td><td>${stats.late}</td><td>${stats.mins}</td><td>${(stats.mins/60).toFixed(2)}</td>`;
    });
}

function updateRawDataTable() {
    const body = document.getElementById('rawDataBody');
    if (!body) return;
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
    const options = '<option value="">-- Select --</option>' + Object.keys(employeeData).sort().map(n => `<option value="${n}">${n}</option>`).join('');
    ['employeeDetailSelect', 'rawEmployeeFilter', 'employeeToClear'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = options;
    });
}

function refreshUI() {
    updateSummaryTable();
    updateEmployeeSelects();
    updateRawDataTable();
}

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
        `<tr><td>${n}</td><td>${workSchedules[n].start}</td><td>${workSchedules[n].mondayStart}</td><td><button onclick="deleteSched('${n}')" class="danger">X</button></td></tr>`
    ).join('');
}

window.deleteSched = (n) => { 
    delete workSchedules[n]; 
    saveData(); 
    updateScheduleTable(); 
};

// --- DATA MANAGEMENT FUNCTIONS ---

function clearAllData() {
    if(confirm("Are you sure you want to delete ALL data? This cannot be undone!")) {
        localStorage.clear();
        employeeData = {};
        workSchedules = { 'PANYAPHON': { start: '08:00', mondayStart: '07:30' } };
        location.reload();
    }
}

// FIX: New function to clear a specific employee
window.clearSelectedEmployee = () => {
    const name = document.getElementById('employeeToClear').value;
    if (!name) return alert("Please select an employee first.");
    
    if (confirm(`Are you sure you want to delete all records for ${name}?`)) {
        delete employeeData[name];
        saveData();
        refreshUI();
        alert(`Data for ${name} has been removed.`);
    }
};

// FIX: New function to clear a specific month
window.clearSelectedMonth = () => {
    const month = document.getElementById('monthToClear').value;
    const year = document.getElementById('yearToClear').value;
    
    if (month === "" || !year) return alert("Please select both month and year.");

    if (confirm(`Delete all records for ${monthNames[month]} ${year}?`)) {
        Object.keys(employeeData).forEach(name => {
            employeeData[name].entries = employeeData[name].entries.filter(e => {
                const d = new Date(e.date);
                return !(d.getMonth() == month && d.getFullYear() == year);
            });
        });
        saveData();
        refreshUI();
        alert(`Records for ${monthNames[month]} ${year} have been cleared.`);
    }
};

// EXCEL EXPORT FUNCTIONS - FIXED VERSION
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
            
            data.push([]); // Empty row
            
            // Add detailed attendance summary
            const yearlyEntries = employeeData[name].entries.filter(e => new Date(e.date).getFullYear() == y);
            const dayMapYearly = {};
            yearlyEntries.forEach(e => { 
                if(!dayMapYearly[e.date]) dayMapYearly[e.date] = e.time; 
            });
            
            let onTimeDays = 0;
            let lateDays = 0;
            let noRecordDays = 0;
            
            Object.keys(dayMapYearly).forEach(dateStr => {
                const time = dayMapYearly[dateStr];
                if (time === '-') {
                    noRecordDays++;
                } else {
                    const late = calcLate(name, dateStr, time);
                    if (late > 0) {
                        lateDays++;
                    } else {
                        onTimeDays++;
                    }
                }
            });
            
            data.push(["ATTENDANCE ANALYSIS"]);
            data.push(["Status", "Days", "Percentage"]);
            data.push(["On Time", onTimeDays, yearlyWorkDays > 0 ? `${((onTimeDays/yearlyWorkDays)*100).toFixed(1)}%` : '0%']);
            data.push(["Late", lateDays, yearlyWorkDays > 0 ? `${((lateDays/yearlyWorkDays)*100).toFixed(1)}%` : '0%']);
            data.push(["No Record", noRecordDays, yearlyWorkDays > 0 ? `${((noRecordDays/yearlyWorkDays)*100).toFixed(1)}%` : '0%']);
            data.push(["Total", yearlyWorkDays, "100%"]);
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
        
        // Style the header rows (optional - makes Excel look better)
        if (!ws['!merges']) ws['!merges'] = [];
        
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