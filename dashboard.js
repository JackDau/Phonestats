// YourGP Phone Dashboard JavaScript - Enhanced Version

// ========================================
// CONFIGURATION (Extracted from hardcoded values)
// ========================================
const CONFIG = {
    // OneDrive/Azure configuration
    ONEDRIVE_CLIENT_ID: "bdf5829d-49a6-4bed-aa55-89bf6ef866bc",

    // Opening hours (minutes from midnight)
    OPENING_HOURS: {
        weekday: { start: 450, end: 1050 },  // Mon-Fri: 7:30am (450) - 5:30pm (1050)
        saturday: { start: 540, end: 750 },   // Sat: 9:00am (540) - 12:30pm (750)
        sunday: null                          // Closed
    },

    // Heatmap time slots
    HEATMAP: {
        startSlot: 15,  // 7:30 AM
        endSlot: 35,    // 5:30 PM
        days: [1, 2, 3, 4, 5, 6]  // Mon-Sat
    },

    // Internal extensions to exclude from incoming call statistics
    INTERNAL_EXTENSIONS: [
        'Nurse 1', 'Crace', 'Lyneham - Rec 1', 'Crace - Rec 1',
        'Crace - Rec 2', 'Crace Office', 'Nurse 5 (TR1)', 'Nurse 2',
        'Nurse Consult', 'Lyneham - Nurse', 'Nurse 3 (TR2)',
        'Denman - Nurse', 'Nurse 4 (TR2)', 'Denman - Rec 1'
    ],

    // Queue name mappings
    QUEUE_MAP: {
        'appointments': 'Appointments',
        'vasectomy': 'Canberra Vasectomy',
        'general': 'General Enquiries',
        'health': 'Health Professionals'
    },

    // Location patterns for filtering
    LOCATION_PATTERNS: {
        'crace': ['crace'],
        'denman': ['denman'],
        'lyneham': ['lyneham'],
        'practice': ['practice support'],
        'management': ['management / support', 'management/support']
    },

    // Targets for benchmarks
    TARGETS: {
        missedPct: 5,      // Target: under 5%
        serviceLevel: 80   // Target: 80%
    }
};

// ========================================
// GLOBAL STATE
// ========================================
let rawData = [];
let queueData = {};
let currentQueueFilter = 'all';
let currentDailyDirection = 'all';
let serviceLevelTarget = 90;
let currentLocationFilter = 'all';
let currentGlobalLocation = 'all';
let currentHeatmapLocation = 'all';
let availableWeeks = [];
let currentWeekFilter = 'all';
let showWeeklyAverages = false;
let dateRangeStart = null;
let dateRangeEnd = null;
let dataMinDate = null;
let dataMaxDate = null;
let hourlyChart = null;
let weekTrendVolumeChart = null;
let weekTrendQualityChart = null;
let callbackWindowHours = 24;
let currentHeatmapTab = 'volume';

// ========================================
// ERROR HANDLING (Replaces alerts)
// ========================================
function showError(message) {
    const errorDisplay = document.getElementById('errorDisplay');
    const errorMessage = document.getElementById('errorMessage');
    if (errorDisplay && errorMessage) {
        errorMessage.textContent = message;
        errorDisplay.style.display = 'block';
    }
    console.error('Dashboard Error:', message);
}

function hideError() {
    const errorDisplay = document.getElementById('errorDisplay');
    if (errorDisplay) {
        errorDisplay.style.display = 'none';
    }
}

// ========================================
// LOADING STATE
// ========================================
function showLoading(step = '') {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('noData').style.display = 'none';
    document.getElementById('dashboard').style.display = 'none';
    const loadingStep = document.getElementById('loadingStep');
    if (loadingStep) {
        loadingStep.textContent = step;
    }
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

// ========================================
// ONEDRIVE FILE PICKER
// ========================================
function launchOneDrivePicker() {
    const btn = document.getElementById('oneDriveBtn');
    btn.textContent = 'Connecting...';

    OneDrive.open({
        clientId: CONFIG.ONEDRIVE_CLIENT_ID,
        action: "download",
        multiSelect: true,
        advanced: { filter: ".csv" },
        success: handleOneDriveFiles,
        cancel: function() {
            btn.textContent = 'Load from SharePoint';
        },
        error: function(err) {
            btn.textContent = 'Load from SharePoint';
            showError('OneDrive Error: ' + (err.message || 'Unknown error'));
        }
    });
}

async function handleOneDriveFiles(response) {
    if (!response.value || response.value.length === 0) {
        console.log('No files selected from OneDrive');
        return;
    }

    showLoading('Connecting to SharePoint...');

    try {
        queueData = {};
        let mainFileData = null;

        for (const file of response.value) {
            const url = file['@microsoft.graph.downloadUrl'] || file['@content.downloadUrl'];

            if (!url) {
                console.error('No download URL for file:', file.name);
                continue;
            }

            showLoading(`Downloading ${file.name}...`);
            console.log('Downloading from OneDrive:', file.name);

            const res = await fetch(url);
            const blob = await res.blob();

            showLoading(`Parsing ${file.name}...`);
            const rows = await readCsvFromBlob(blob);

            if (file.name.toLowerCase().startsWith('callqueue')) {
                const filenameQueueName = extractQueueName(file.name);
                rows.forEach(row => {
                    if (row.CallGUID) {
                        const queueName = row.CallQueueName || filenameQueueName;
                        queueData[row.CallGUID] = queueName;
                    }
                });
                const sampleQueue = rows[0]?.CallQueueName || filenameQueueName;
                console.log(`Loaded ${rows.length} records from queue: ${sampleQueue}`);
            } else {
                mainFileData = rows;
            }
        }

        if (!mainFileData) {
            throw new Error('No main export file found. Please select a file starting with "Export".');
        }

        showLoading('Processing data...');

        dataMinDate = null;
        dataMaxDate = null;
        dateRangeStart = null;
        dateRangeEnd = null;

        rawData = mainFileData.map(row => ({
            ...row,
            CallDuration: parseFloat(row.CallDuration) || 0,
            TimeToAnswer: parseFloat(row.TimeToAnswer) || 0,
            BillableTime: parseFloat(row.BillableTime) || 0,
            queueName: queueData[row.CallGUID] || null
        }));

        const queuedCount = rawData.filter(r => r.queueName).length;
        console.log(`Loaded ${rawData.length} records from OneDrive, ${queuedCount} matched to queues`);

        processAndDisplay();

    } catch (err) {
        console.error('Error processing OneDrive files:', err);
        showError('Error processing files from OneDrive: ' + err.message);
        hideLoading();
        document.getElementById('noData').style.display = 'block';
    }

    document.getElementById('oneDriveBtn').textContent = 'Load from SharePoint';
}

function readCsvFromBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', raw: true });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { raw: true });
                resolve(rows);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = function() {
            reject(new Error('Failed to read file'));
        };
        reader.readAsArrayBuffer(blob);
    });
}

// ========================================
// LOCAL FILE UPLOAD
// ========================================
document.getElementById('fileInput').addEventListener('change', async function(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    showLoading('Reading files...');

    try {
        let mainFile = null;
        const queueFiles = [];

        files.forEach(file => {
            if (file.name.startsWith('Export') || file.name.startsWith('export')) {
                mainFile = file;
            } else if (file.name.startsWith('CallQueue') || file.name.startsWith('callqueue')) {
                queueFiles.push(file);
            }
        });

        if (!mainFile && files.length === 1) {
            mainFile = files[0];
        }

        if (!mainFile) {
            throw new Error('No main export file found. Please include a file starting with "Export".');
        }

        queueData = {};

        for (const queueFile of queueFiles) {
            showLoading(`Loading ${queueFile.name}...`);
            const queueRows = await readCsvFile(queueFile);
            const filenameQueueName = extractQueueName(queueFile.name);

            queueRows.forEach(row => {
                if (row.CallGUID) {
                    const queueName = row.CallQueueName || filenameQueueName;
                    queueData[row.CallGUID] = queueName;
                }
            });

            const sampleQueue = queueRows[0]?.CallQueueName || filenameQueueName;
            console.log(`Loaded ${queueRows.length} records from queue: ${sampleQueue}`);
        }

        showLoading('Processing main export...');

        dataMinDate = null;
        dataMaxDate = null;
        dateRangeStart = null;
        dateRangeEnd = null;

        rawData = await readCsvFile(mainFile);

        rawData = rawData.map(row => ({
            ...row,
            CallDuration: parseFloat(row.CallDuration) || 0,
            TimeToAnswer: parseFloat(row.TimeToAnswer) || 0,
            BillableTime: parseFloat(row.BillableTime) || 0,
            queueName: queueData[row.CallGUID] || null
        }));

        const queuedCount = rawData.filter(r => r.queueName).length;
        console.log(`Loaded ${rawData.length} records from main CSV, ${queuedCount} matched to queues`);

        processAndDisplay();
    } catch (err) {
        console.error('Error reading files:', err);
        showError('Error reading CSV files: ' + err.message);
        hideLoading();
        document.getElementById('noData').style.display = 'block';
    }
});

function readCsvFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', raw: true });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { raw: true });
                resolve(rows);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file: ' + file.name));
        reader.readAsArrayBuffer(file);
    });
}

// ========================================
// SAMPLE DATA (For demo/training)
// ========================================
function loadSampleData() {
    showLoading('Generating sample data...');

    // Generate realistic sample data
    const sampleData = generateSampleData();
    rawData = sampleData;
    queueData = {};

    dataMinDate = null;
    dataMaxDate = null;
    dateRangeStart = null;
    dateRangeEnd = null;

    setTimeout(() => {
        processAndDisplay();
    }, 500);
}

function generateSampleData() {
    const data = [];
    const queues = ['Appointments', 'General Enquiries', 'Health Professionals', 'Canberra Vasectomy'];
    const locations = ['Crace', 'Denman Prospect', 'Lyneham'];
    const staff = ['Sarah M', 'James T', 'Emily R', 'Michael B', 'Jessica L', 'David K'];

    // Generate 2 weeks of data
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 14);

    for (let day = 0; day < 14; day++) {
        const date = new Date(startDate);
        date.setDate(date.getDate() + day);

        if (date.getDay() === 0) continue; // Skip Sundays

        const isSaturday = date.getDay() === 6;
        const callCount = isSaturday ? Math.floor(Math.random() * 30) + 20 : Math.floor(Math.random() * 80) + 60;

        for (let i = 0; i < callCount; i++) {
            const hour = isSaturday
                ? Math.floor(Math.random() * 3) + 9
                : Math.floor(Math.random() * 10) + 8;
            const minute = Math.floor(Math.random() * 60);

            const callDate = new Date(date);
            callDate.setHours(hour, minute, 0, 0);

            const isIncoming = Math.random() > 0.3;
            const isAnswered = Math.random() > 0.08;
            const hasQueue = isIncoming && Math.random() > 0.2;

            const record = {
                CallGUID: `SAMPLE-${day}-${i}`,
                CallDateTime: callDate.toLocaleString('en-AU'),
                Direction: isIncoming ? 'In' : 'Out',
                OfficeName: locations[Math.floor(Math.random() * locations.length)],
                UserName: staff[Math.floor(Math.random() * staff.length)],
                OriginNumber: `04${Math.floor(Math.random() * 100000000).toString().padStart(8, '0')}`,
                CallDuration: isAnswered ? Math.floor(Math.random() * 300) + 30 : Math.floor(Math.random() * 60),
                TimeToAnswer: isAnswered ? Math.floor(Math.random() * 120) + 5 : 0,
                BillableTime: isAnswered ? Math.floor(Math.random() * 300) + 30 : 0,
                queueName: hasQueue ? queues[Math.floor(Math.random() * queues.length)] : null
            };

            data.push(record);
        }
    }

    return data;
}

// ========================================
// UTILITY FUNCTIONS
// ========================================
function extractQueueName(filename) {
    const withoutExt = filename.replace(/\.csv$/i, '');
    const parts = withoutExt.split('_');
    let queueNameParts = [];
    let foundDates = 0;
    for (let i = 0; i < parts.length; i++) {
        if (/^\d{8}$/.test(parts[i])) {
            foundDates++;
        } else if (foundDates >= 2) {
            queueNameParts.push(parts[i]);
        }
    }
    return queueNameParts.join(' ') || 'Unknown';
}

function isInternalCall(row) {
    const userName = row.UserName || '';
    return CONFIG.INTERNAL_EXTENSIONS.some(ext =>
        userName.toLowerCase() === ext.toLowerCase()
    );
}

function isWithinOpeningHours(dateValue) {
    const date = getDateObj(dateValue);
    if (!date) return false;

    const dayOfWeek = date.getDay();
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const timeInMinutes = hours * 60 + minutes;

    if (dayOfWeek === 0) return false; // Sunday closed

    if (dayOfWeek === 6) {
        const sat = CONFIG.OPENING_HOURS.saturday;
        return timeInMinutes >= sat.start && timeInMinutes <= sat.end;
    }

    const weekday = CONFIG.OPENING_HOURS.weekday;
    return timeInMinutes >= weekday.start && timeInMinutes <= weekday.end;
}

function formatTime(seconds) {
    if (seconds === null || seconds === undefined || isNaN(seconds)) return '-';
    seconds = Math.round(seconds);
    if (seconds < 60) return `${seconds}s`;
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    if (mins < 60) return `${mins}m ${secs}s`;
    const hrs = Math.floor(mins / 60);
    return `${hrs}h ${mins % 60}m`;
}

function formatTimeShort(seconds) {
    if (seconds === null || seconds === undefined || isNaN(seconds)) return '-';
    seconds = Math.round(seconds);
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins}:${secs.toString().padStart(2, '0')}`;
}

function getDateObj(dateValue) {
    if (dateValue instanceof Date) return dateValue;
    if (typeof dateValue === 'number') {
        const excelEpoch = new Date(1899, 11, 30);
        const days = Math.floor(dateValue);
        const timeFraction = dateValue - days;
        const date = new Date(excelEpoch.getTime() + days * 86400 * 1000);
        const totalSeconds = Math.round(timeFraction * 86400);
        date.setHours(Math.floor(totalSeconds / 3600));
        date.setMinutes(Math.floor((totalSeconds % 3600) / 60));
        date.setSeconds(totalSeconds % 60);
        return date;
    }
    if (typeof dateValue === 'string') {
        const ddmmyyyyMatch = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?)?/i);
        if (ddmmyyyyMatch) {
            const day = parseInt(ddmmyyyyMatch[1]);
            const month = parseInt(ddmmyyyyMatch[2]) - 1;
            const year = parseInt(ddmmyyyyMatch[3]);
            let hours = ddmmyyyyMatch[4] ? parseInt(ddmmyyyyMatch[4]) : 0;
            const minutes = ddmmyyyyMatch[5] ? parseInt(ddmmyyyyMatch[5]) : 0;
            const seconds = ddmmyyyyMatch[6] ? parseInt(ddmmyyyyMatch[6]) : 0;
            const ampm = ddmmyyyyMatch[7] ? ddmmyyyyMatch[7].toUpperCase() : null;

            if (ampm === 'PM' && hours !== 12) hours += 12;
            else if (ampm === 'AM' && hours === 12) hours = 0;

            return new Date(year, month, day, hours, minutes, seconds);
        }
        return new Date(dateValue);
    }
    return null;
}

function getDayOfWeek(dateValue) {
    const date = getDateObj(dateValue);
    return date ? date.getDay() : null;
}

function getHour(dateValue) {
    const date = getDateObj(dateValue);
    return date ? date.getHours() : null;
}

function getTimeSlot(dateValue) {
    const date = getDateObj(dateValue);
    if (!date) return null;
    const hours = date.getHours();
    const minutes = date.getMinutes();
    return hours * 2 + (minutes >= 30 ? 1 : 0);
}

function formatTimeSlot(slot) {
    const hours = Math.floor(slot / 2);
    const minutes = (slot % 2) * 30;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// ========================================
// FILTER FUNCTIONS
// ========================================
function setQueueFilter(queue) {
    currentQueueFilter = queue;
    const buttons = ['queueAll', 'queueAppointments', 'queueVasectomy', 'queueGeneral', 'queueHealth', 'queueNone'];
    const values = ['all', 'appointments', 'vasectomy', 'general', 'health', 'noqueue'];
    buttons.forEach((btnId, idx) => {
        const btn = document.getElementById(btnId);
        if (btn) btn.classList.toggle('active', values[idx] === queue);
    });
    if (rawData.length > 0) processAndDisplay();
}

function setDailyDirection(direction) {
    currentDailyDirection = direction;
    document.getElementById('tabInOut').classList.toggle('active', direction === 'all');
    document.getElementById('tabIn').classList.toggle('active', direction === 'in');
    document.getElementById('tabOut').classList.toggle('active', direction === 'out');
    if (rawData.length > 0) updateDailyTable();
}

function updateServiceLevelTarget() {
    serviceLevelTarget = parseInt(document.getElementById('slTarget').value);
    if (rawData.length > 0) {
        const filteredData = getGlobalFilteredData();
        const incomingCalls = filteredData.filter(row => row.Direction === 'In');
        updateSummaryMetrics(incomingCalls);
    }
}

function updateCallbackWindow() {
    callbackWindowHours = parseInt(document.getElementById('callbackWindow').value);
    const windowText = callbackWindowHours === 24 ? '24h' : (callbackWindowHours / 24) + ' day';
    document.getElementById('followupHeader').textContent = `Missed Call Follow-up (${windowText} Window)`;

    if (rawData.length > 0) {
        const filteredData = getGlobalFilteredData();
        const incomingCalls = filteredData.filter(row => row.Direction === 'In');
        updateSummaryMetrics(incomingCalls);
        updateMissedCallFollowup(incomingCalls);
    }
}

function setGlobalLocation() {
    currentGlobalLocation = document.getElementById('globalLocationFilter').value;
    if (rawData.length > 0) processAndDisplay();
}

function updateStaffTableFilter() {
    currentLocationFilter = document.getElementById('locationFilter').value;
    if (rawData.length > 0) {
        let filteredData = getGlobalFilteredData();
        updateStaffTable(filteredData);
    }
}

function filterByLocation(data, location) {
    if (location === 'all') return data;

    const patterns = CONFIG.LOCATION_PATTERNS[location.toLowerCase()];
    if (!patterns) return data;

    return data.filter(row => {
        const office = (row.OfficeName || '').toLowerCase();
        return patterns.some(pattern => office.includes(pattern));
    });
}

function setWeekFilter() {
    currentWeekFilter = document.getElementById('weekSelector').value;
    processAndDisplay();
}

function setDateRange() {
    const fromValue = document.getElementById('dateFrom').value;
    const toValue = document.getElementById('dateTo').value;

    if (fromValue) dateRangeStart = new Date(fromValue + 'T00:00:00');
    if (toValue) dateRangeEnd = new Date(toValue + 'T23:59:59');

    if (dateRangeStart && dateRangeEnd && dateRangeStart > dateRangeEnd) {
        [dateRangeStart, dateRangeEnd] = [dateRangeEnd, dateRangeStart];
        populateDateInputs();
    }

    processAndDisplay();
}

function toggleWeeklyAverages() {
    showWeeklyAverages = document.getElementById('avgToggle').checked;
    const filteredData = getGlobalFilteredData();
    const incomingCalls = filteredData.filter(row => row.Direction === 'In');
    updateSummaryMetrics(incomingCalls);
}

function resetAllFilters() {
    // Reset all filter states
    currentQueueFilter = 'all';
    currentGlobalLocation = 'all';
    currentWeekFilter = 'all';
    showWeeklyAverages = false;
    serviceLevelTarget = 90;
    callbackWindowHours = 24;

    // Reset date range to full data range
    dateRangeStart = dataMinDate;
    dateRangeEnd = dataMaxDate;

    // Reset UI elements
    document.getElementById('globalLocationFilter').value = 'all';
    document.getElementById('weekSelector').value = 'all';
    document.getElementById('avgToggle').checked = false;
    document.getElementById('slTarget').value = '90';
    document.getElementById('callbackWindow').value = '24';

    // Reset queue buttons
    const buttons = ['queueAll', 'queueAppointments', 'queueVasectomy', 'queueGeneral', 'queueHealth', 'queueNone'];
    buttons.forEach((btnId, idx) => {
        const btn = document.getElementById(btnId);
        if (btn) btn.classList.toggle('active', idx === 0);
    });

    populateDateInputs();

    if (rawData.length > 0) processAndDisplay();
}

function toggleAdvancedFilters() {
    const toggle = document.getElementById('advancedToggle');
    const filters = document.getElementById('advancedFilters');
    toggle.classList.toggle('expanded');
    filters.classList.toggle('visible');

    const isExpanded = toggle.classList.contains('expanded');
    toggle.innerHTML = isExpanded
        ? '<span class="arrow">&#9660;</span> Less Filters'
        : '<span class="arrow">&#9660;</span> More Filters';
}

function switchHeatmapTab(tab) {
    currentHeatmapTab = tab;

    // Update tab buttons
    document.querySelectorAll('.heatmap-tab').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tab);
    });

    // Update tab content
    document.querySelectorAll('.heatmap-tab-content').forEach(content => {
        content.classList.remove('active');
    });

    const tabMap = { 'volume': 'heatmapTabVolume', 'wait': 'heatmapTabWait', 'missed': 'heatmapTabMissed' };
    const activeContent = document.getElementById(tabMap[tab]);
    if (activeContent) activeContent.classList.add('active');
}

function updateHeatmapLocation() {
    currentHeatmapLocation = document.getElementById('heatmapLocationFilter').value;
    const filteredData = getGlobalFilteredData();
    updateHeatmaps(filteredData);
}

// ========================================
// DATA FILTERING
// ========================================
function getGlobalFilteredData() {
    let filteredData = rawData;

    if (dateRangeStart && dateRangeEnd) {
        filteredData = filteredData.filter(row => {
            const date = getDateObj(row.CallDateTime);
            if (!date) return false;
            return date >= dateRangeStart && date <= dateRangeEnd;
        });
    }

    if (currentWeekFilter !== 'all') {
        const weekIdx = parseInt(currentWeekFilter);
        const week = availableWeeks[weekIdx];
        if (week) {
            filteredData = filteredData.filter(row => {
                const date = getDateObj(row.CallDateTime);
                if (!date) return false;
                return date >= week.start && date <= week.end;
            });
        }
    }

    filteredData = filteredData.filter(row => row.Direction !== 'Int');
    filteredData = filteredData.filter(row => isWithinOpeningHours(row.CallDateTime));
    filteredData = filteredData.filter(row => !isInternalCall(row));
    filteredData = filterByLocation(filteredData, currentGlobalLocation);

    if (currentQueueFilter !== 'all') {
        if (currentQueueFilter === 'noqueue') {
            filteredData = filteredData.filter(row => !row.queueName);
        } else {
            const targetQueue = CONFIG.QUEUE_MAP[currentQueueFilter];
            filteredData = filteredData.filter(row => row.queueName === targetQueue);
        }
    }

    return filteredData;
}

// ========================================
// WEEK/DATE DETECTION
// ========================================
function detectWeeksInData(data) {
    const dates = data.map(row => getDateObj(row.CallDateTime)).filter(d => d && !isNaN(d));
    if (dates.length === 0) return [];

    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    const weeks = [];
    let weekStart = new Date(minDate);
    const dayOfWeek = weekStart.getDay();
    const daysToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    weekStart.setDate(weekStart.getDate() + daysToMonday);
    weekStart.setHours(0, 0, 0, 0);

    while (weekStart <= maxDate) {
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        weekEnd.setHours(23, 59, 59, 999);

        const label = `Week of ${weekStart.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' })}`;
        weeks.push({ start: new Date(weekStart), end: new Date(weekEnd), label });

        weekStart.setDate(weekStart.getDate() + 7);
    }
    return weeks;
}

function populateWeekSelector() {
    const select = document.getElementById('weekSelector');
    if (!select) return;

    select.innerHTML = '<option value="all">All Weeks</option>';
    availableWeeks.forEach((week, idx) => {
        select.innerHTML += `<option value="${idx}">${week.label}</option>`;
    });

    if (currentWeekFilter !== 'all') {
        const idx = parseInt(currentWeekFilter);
        if (idx >= 0 && idx < availableWeeks.length) {
            select.value = currentWeekFilter;
        } else {
            currentWeekFilter = 'all';
        }
    }
}

function detectDateRange(data) {
    const dates = data.map(row => getDateObj(row.CallDateTime)).filter(d => d && !isNaN(d));
    if (dates.length === 0) return { min: null, max: null };
    return {
        min: new Date(Math.min(...dates)),
        max: new Date(Math.max(...dates))
    };
}

function formatDateForInput(date) {
    if (!date) return '';
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function populateDateInputs() {
    const fromInput = document.getElementById('dateFrom');
    const toInput = document.getElementById('dateTo');
    if (!fromInput || !toInput) return;

    if (dataMinDate) {
        fromInput.min = formatDateForInput(dataMinDate);
        fromInput.max = formatDateForInput(dataMaxDate);
        fromInput.value = formatDateForInput(dateRangeStart);
    }
    if (dataMaxDate) {
        toInput.min = formatDateForInput(dataMinDate);
        toInput.max = formatDateForInput(dataMaxDate);
        toInput.value = formatDateForInput(dateRangeEnd);
    }
}

// ========================================
// MAIN PROCESSING
// ========================================
function processAndDisplay() {
    hideLoading();
    document.getElementById('dashboard').style.display = 'block';

    availableWeeks = detectWeeksInData(rawData);
    populateWeekSelector();

    if (!dataMinDate || !dataMaxDate) {
        const range = detectDateRange(rawData);
        dataMinDate = range.min;
        dataMaxDate = range.max;
        dateRangeStart = range.min;
        dateRangeEnd = range.max;
    }
    populateDateInputs();

    let filteredData = getGlobalFilteredData();

    updateWeekInfo(filteredData);

    const incomingCalls = filteredData.filter(row => row.Direction === 'In');
    updateSummaryMetrics(incomingCalls);
    updateInsightsPanel(incomingCalls, filteredData);
    updateAbandonmentAnalysis(incomingCalls);
    updateMissedCallFollowup(filteredData);
    updateHourlyChart(filteredData);
    updateWeekTrendCharts();
    updateSiteBreakdown(incomingCalls);
    updateDailyTable();
    updateHeatmaps(filteredData);
    updateStaffTable(filteredData);
}

function updateWeekInfo(data) {
    if (data.length === 0) {
        document.getElementById('weekInfo').textContent = 'No data';
        return;
    }

    let dates = data.map(row => getDateObj(row.CallDateTime)).filter(d => d && !isNaN(d));

    if (dates.length === 0) {
        document.getElementById('weekInfo').textContent = 'No valid dates';
        return;
    }

    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    const formatDate = (d) => d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short' });
    document.getElementById('weekInfo').textContent =
        `${formatDate(minDate)} - ${formatDate(maxDate)} (${data.length} calls)`;
}

// ========================================
// AUTO-GENERATED INSIGHTS
// ========================================
function updateInsightsPanel(incomingCalls, allCalls) {
    const insightsGrid = document.getElementById('insightsGrid');
    if (!insightsGrid) return;

    const insights = [];

    // Find busiest time slot
    const slotCounts = {};
    incomingCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null) {
            const key = `${day}-${slot}`;
            slotCounts[key] = (slotCounts[key] || 0) + 1;
        }
    });

    const busiestSlot = Object.entries(slotCounts).sort((a, b) => b[1] - a[1])[0];
    if (busiestSlot) {
        const [daySlot, count] = busiestSlot;
        const [day, slot] = daySlot.split('-').map(Number);
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        insights.push({
            icon: '📞',
            label: 'Busiest Time',
            value: `${dayNames[day]} ${formatTimeSlot(slot)} (${count} calls)`
        });
    }

    // Find highest missed rate time
    const missedBySlot = {};
    const totalBySlot = {};
    incomingCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null) {
            const key = `${day}-${slot}`;
            totalBySlot[key] = (totalBySlot[key] || 0) + 1;
            if (!call.TimeToAnswer || call.TimeToAnswer === 0) {
                missedBySlot[key] = (missedBySlot[key] || 0) + 1;
            }
        }
    });

    let worstSlot = null;
    let worstRate = 0;
    Object.keys(totalBySlot).forEach(key => {
        if (totalBySlot[key] >= 5) { // Only consider slots with at least 5 calls
            const rate = (missedBySlot[key] || 0) / totalBySlot[key];
            if (rate > worstRate) {
                worstRate = rate;
                worstSlot = key;
            }
        }
    });

    if (worstSlot && worstRate > 0.05) {
        const [day, slot] = worstSlot.split('-').map(Number);
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        insights.push({
            icon: '⚠️',
            label: 'Highest Miss Rate',
            value: `${dayNames[day]} ${formatTimeSlot(slot)} (${(worstRate * 100).toFixed(0)}%)`
        });
    }

    // Best performing site
    const siteStats = {};
    incomingCalls.forEach(call => {
        const office = call.OfficeName || 'Unknown';
        const site = office.toLowerCase().includes('crace') ? 'Crace' :
                    office.toLowerCase().includes('denman') ? 'Denman' :
                    office.toLowerCase().includes('lyneham') ? 'Lyneham' : null;
        if (site) {
            if (!siteStats[site]) siteStats[site] = { total: 0, answered: 0 };
            siteStats[site].total++;
            if (call.TimeToAnswer > 0) siteStats[site].answered++;
        }
    });

    let bestSite = null;
    let bestRate = 0;
    Object.entries(siteStats).forEach(([site, stats]) => {
        if (stats.total >= 10) {
            const rate = stats.answered / stats.total;
            if (rate > bestRate) {
                bestRate = rate;
                bestSite = site;
            }
        }
    });

    if (bestSite) {
        insights.push({
            icon: '🏆',
            label: 'Best Performer',
            value: `${bestSite} - ${(bestRate * 100).toFixed(0)}% answered`
        });
    }

    // Average wait time insight
    const answeredCalls = incomingCalls.filter(c => c.TimeToAnswer > 0);
    if (answeredCalls.length > 0) {
        const avgWait = answeredCalls.reduce((sum, c) => sum + c.TimeToAnswer, 0) / answeredCalls.length;
        const status = avgWait < 30 ? 'Excellent' : avgWait < 60 ? 'Good' : avgWait < 90 ? 'Fair' : 'Needs attention';
        insights.push({
            icon: '⏱️',
            label: 'Avg Wait Time',
            value: `${formatTime(avgWait)} - ${status}`
        });
    }

    // Render insights
    insightsGrid.innerHTML = insights.map(insight => `
        <div class="insight-card">
            <div class="insight-icon">${insight.icon}</div>
            <div class="insight-content">
                <div class="insight-label">${insight.label}</div>
                <div class="insight-value">${insight.value}</div>
            </div>
        </div>
    `).join('');
}

// ========================================
// SUMMARY METRICS
// ========================================
function updateSummaryMetrics(calls) {
    const total = calls.length;
    const answered = calls.filter(c => c.TimeToAnswer > 0).length;
    const missed = calls.filter(c => c.queueName && (!c.TimeToAnswer || c.TimeToAnswer === 0)).length;
    const missedPct = total > 0 ? ((missed / total) * 100).toFixed(1) : 0;

    const answeredCalls = calls.filter(c => c.TimeToAnswer > 0);
    const avgWait = answeredCalls.length > 0
        ? answeredCalls.reduce((sum, c) => sum + (c.TimeToAnswer || 0), 0) / answeredCalls.length
        : 0;
    const maxWait = answeredCalls.length > 0
        ? Math.max(...answeredCalls.map(c => c.TimeToAnswer || 0))
        : 0;

    const callLengths = answeredCalls
        .map(c => (c.CallDuration || 0) - (c.TimeToAnswer || 0))
        .filter(l => l > 0);
    const avgCallLength = callLengths.length > 0
        ? callLengths.reduce((a, b) => a + b, 0) / callLengths.length
        : 0;

    const withinTarget = answeredCalls.filter(c => c.TimeToAnswer <= serviceLevelTarget).length;
    const serviceLevel = total > 0 ? ((withinTarget / total) * 100).toFixed(1) : 0;

    const { fcrRate, callbackRate } = calculateCallbackMetrics(calls);

    const numWeeks = currentWeekFilter === 'all' ? Math.max(1, availableWeeks.length) : 1;

    let displayTotal = total;
    let displayAnswered = answered;
    let displayMissed = missed;

    if (showWeeklyAverages && currentWeekFilter === 'all' && numWeeks > 1) {
        displayTotal = Math.round(total / numWeeks);
        displayAnswered = Math.round(answered / numWeeks);
        displayMissed = Math.round(missed / numWeeks);
    }

    document.getElementById('totalCalls').textContent = displayTotal;
    document.getElementById('answeredCalls').textContent = displayAnswered;
    document.getElementById('missedCalls').textContent = displayMissed;
    document.getElementById('missedPercent').textContent = missedPct + '%';
    document.getElementById('serviceLevel').textContent = serviceLevel + '%';
    document.getElementById('fcrRate').textContent = fcrRate + '%';
    document.getElementById('callbackRate').textContent = callbackRate + '%';
    document.getElementById('avgWait').textContent = formatTime(avgWait);
    document.getElementById('maxWait').textContent = formatTime(maxWait);
    document.getElementById('avgCallLength').textContent = formatTime(avgCallLength);

    const outOfHours = countOutOfHoursCalls();
    document.getElementById('outOfHoursCalls').textContent = outOfHours;
}

function countOutOfHoursCalls() {
    let data = rawData;
    data = data.filter(row => row.Direction !== 'Int');
    data = data.filter(row => !isInternalCall(row));
    data = filterByLocation(data, currentGlobalLocation);

    if (currentQueueFilter !== 'all') {
        if (currentQueueFilter === 'noqueue') {
            data = data.filter(row => !row.queueName);
        } else {
            const targetQueue = CONFIG.QUEUE_MAP[currentQueueFilter];
            data = data.filter(row => row.queueName === targetQueue);
        }
    }

    return data.filter(row =>
        row.Direction === 'In' &&
        !isWithinOpeningHours(row.CallDateTime)
    ).length;
}

function calculateCallbackMetrics(calls) {
    const validCalls = calls.filter(c =>
        c.OriginNumber && c.OriginNumber !== '0' && c.OriginNumber !== 0
    );

    if (validCalls.length === 0) {
        return { fcrRate: 0, callbackRate: 0 };
    }

    const callsByNumber = {};
    validCalls.forEach(call => {
        const num = String(call.OriginNumber);
        if (!callsByNumber[num]) callsByNumber[num] = [];
        const date = getDateObj(call.CallDateTime);
        if (date) {
            callsByNumber[num].push({
                call,
                date,
                isAnswered: call.TimeToAnswer > 0
            });
        }
    });

    let uniqueCallersWithCallback = 0;
    let totalUniqueCallers = Object.keys(callsByNumber).length;

    Object.values(callsByNumber).forEach(callList => {
        if (callList.length <= 1) return;
        callList.sort((a, b) => a.date - b.date);

        let hasCallback = false;
        for (let i = 0; i < callList.length - 1; i++) {
            const timeDiff = callList[i + 1].date - callList[i].date;
            const hoursDiff = timeDiff / (1000 * 60 * 60);
            if (hoursDiff <= callbackWindowHours) {
                hasCallback = true;
                break;
            }
        }

        if (hasCallback) uniqueCallersWithCallback++;
    });

    const callbackRate = totalUniqueCallers > 0
        ? ((uniqueCallersWithCallback / totalUniqueCallers) * 100).toFixed(1)
        : 0;

    const fcrRate = totalUniqueCallers > 0
        ? (100 - parseFloat(callbackRate)).toFixed(1)
        : 0;

    return { fcrRate, callbackRate };
}

// ========================================
// ABANDONMENT ANALYSIS
// ========================================
function updateAbandonmentAnalysis(calls) {
    const missedCalls = calls.filter(c => !c.TimeToAnswer || c.TimeToAnswer === 0);

    const totalAbandoned = missedCalls.length;
    document.getElementById('totalAbandoned').textContent = totalAbandoned;

    if (totalAbandoned === 0) {
        document.getElementById('avgAbandonWait').textContent = '-';
        document.getElementById('abandonmentGrid').innerHTML =
            '<p style="color: #95a5a6; font-size: 12px;">No abandoned calls in this period</p>';
        return;
    }

    const waitTimes = missedCalls.map(c => c.CallDuration || 0);
    const avgWait = waitTimes.reduce((a, b) => a + b, 0) / waitTimes.length;
    document.getElementById('avgAbandonWait').textContent = formatTime(avgWait);

    const buckets = [
        { label: '<30s', min: 0, max: 30, count: 0 },
        { label: '30-60s', min: 30, max: 60, count: 0 },
        { label: '1-2min', min: 60, max: 120, count: 0 },
        { label: '2-5min', min: 120, max: 300, count: 0 },
        { label: '>5min', min: 300, max: Infinity, count: 0 }
    ];

    missedCalls.forEach(call => {
        const waitTime = call.CallDuration || 0;
        for (const bucket of buckets) {
            if (waitTime >= bucket.min && waitTime < bucket.max) {
                bucket.count++;
                break;
            }
        }
    });

    let html = '';
    buckets.forEach(bucket => {
        const pct = totalAbandoned > 0 ? ((bucket.count / totalAbandoned) * 100).toFixed(0) : 0;
        html += `
            <div class="abandonment-bucket">
                <div class="count">${bucket.count}</div>
                <div class="range">${bucket.label}</div>
                <div class="pct">${pct}%</div>
            </div>
        `;
    });

    document.getElementById('abandonmentGrid').innerHTML = html;
}

// ========================================
// MISSED CALL FOLLOWUP
// ========================================
function toggleFollowupDetails() {
    const btn = document.getElementById('followupExpandBtn');
    const details = document.getElementById('followupDetails');
    btn.classList.toggle('expanded');
    details.classList.toggle('visible');
    btn.innerHTML = btn.classList.contains('expanded')
        ? '<span class="arrow">&#9660;</span> Hide Details'
        : '<span class="arrow">&#9660;</span> Show Details';
}

function updateMissedCallFollowup(calls) {
    const validCalls = calls.filter(c =>
        c.Direction === 'In' &&
        c.OriginNumber &&
        c.OriginNumber !== '0' &&
        c.OriginNumber !== 0
    );

    if (validCalls.length === 0) {
        document.getElementById('lostOpportunities').textContent = '0';
        document.getElementById('persistentCallers').textContent = '0';
        document.getElementById('lostPeakHour').textContent = '';
        document.getElementById('avgAttempts').textContent = '';
        document.getElementById('attemptsGrid').innerHTML = '<p style="color: #95a5a6; font-size: 12px;">No data</p>';
        document.getElementById('lostPeakHours').innerHTML = '<p style="color: #95a5a6; font-size: 12px;">No data</p>';
        document.getElementById('avgWaitBeforeHangup').textContent = '-';
        return;
    }

    const callsByNumber = {};
    validCalls.forEach(call => {
        const num = String(call.OriginNumber);
        if (!callsByNumber[num]) callsByNumber[num] = [];
        const date = getDateObj(call.CallDateTime);
        if (date) {
            callsByNumber[num].push({
                call,
                date,
                hour: date.getHours(),
                isMissed: !call.TimeToAnswer || call.TimeToAnswer === 0,
                isAnswered: call.TimeToAnswer > 0,
                waitTime: call.CallDuration || 0
            });
        }
    });

    let lostCount = 0;
    let persistentCount = 0;
    const lostByHour = {};
    const attemptCounts = { 1: 0, 2: 0, '3+': 0 };
    let totalMissedWaitTime = 0;
    let missedWaitCount = 0;
    const persistentAttempts = [];

    Object.values(callsByNumber).forEach(callList => {
        callList.sort((a, b) => a.date - b.date);

        const firstMissedIdx = callList.findIndex(c => c.isMissed);
        if (firstMissedIdx === -1) return;

        const firstMissed = callList[firstMissedIdx];

        const within24h = callList.filter(c => {
            const hoursDiff = (c.date - firstMissed.date) / (1000 * 60 * 60);
            return hoursDiff >= 0 && hoursDiff <= callbackWindowHours;
        });

        within24h.filter(c => c.isMissed).forEach(c => {
            totalMissedWaitTime += c.waitTime;
            missedWaitCount++;
        });

        const gotAnswered = within24h.some(c => c.isAnswered);

        if (within24h.length === 1 && !gotAnswered) {
            lostCount++;
            const hour = firstMissed.hour;
            lostByHour[hour] = (lostByHour[hour] || 0) + 1;
        } else if (!gotAnswered && within24h.length > 1) {
            lostCount++;
            const hour = firstMissed.hour;
            lostByHour[hour] = (lostByHour[hour] || 0) + 1;
        } else if (gotAnswered) {
            persistentCount++;

            const answeredIdx = within24h.findIndex(c => c.isAnswered);
            const attempts = within24h.slice(0, answeredIdx + 1).filter(c => c.isMissed || c.isAnswered).length;

            persistentAttempts.push(attempts);

            if (attempts === 1) attemptCounts[1]++;
            else if (attempts === 2) attemptCounts[2]++;
            else attemptCounts['3+']++;
        }
    });

    document.getElementById('lostOpportunities').textContent = lostCount;
    document.getElementById('persistentCallers').textContent = persistentCount;

    const peakHour = Object.entries(lostByHour).sort((a, b) => b[1] - a[1])[0];
    if (peakHour) {
        const hour = parseInt(peakHour[0]);
        const hourStr = hour > 12 ? `${hour - 12}-${hour - 11}pm` : (hour === 12 ? '12-1pm' : `${hour}-${hour + 1}am`);
        document.getElementById('lostPeakHour').textContent = `Peak: ${hourStr} (${peakHour[1]})`;
    } else {
        document.getElementById('lostPeakHour').textContent = '';
    }

    if (persistentAttempts.length > 0) {
        const avgAttempts = (persistentAttempts.reduce((a, b) => a + b, 0) / persistentAttempts.length).toFixed(1);
        document.getElementById('avgAttempts').textContent = `Avg ${avgAttempts} attempts`;
    } else {
        document.getElementById('avgAttempts').textContent = '';
    }

    let attemptsHtml = '';
    Object.entries(attemptCounts).forEach(([attempts, count]) => {
        attemptsHtml += `
            <div class="detail-item">
                <div class="count">${count}</div>
                <div class="label">${attempts} attempt${attempts === '1' ? '' : 's'}</div>
            </div>
        `;
    });
    document.getElementById('attemptsGrid').innerHTML = attemptsHtml || '<p style="color: #95a5a6; font-size: 12px;">No persistent callers</p>';

    const sortedHours = Object.entries(lostByHour).sort((a, b) => b[1] - a[1]).slice(0, 5);

    let peakHoursHtml = '';
    sortedHours.forEach(([hour, count]) => {
        const h = parseInt(hour);
        const hourStr = h > 12 ? `${h - 12}-${h - 11}pm` : (h === 12 ? '12-1pm' : `${h}-${h + 1}am`);
        peakHoursHtml += `<span class="peak-hour-tag">${hourStr}: ${count} lost</span>`;
    });
    document.getElementById('lostPeakHours').innerHTML = peakHoursHtml || '<p style="color: #95a5a6; font-size: 12px;">No lost opportunities</p>';

    const avgWait = missedWaitCount > 0 ? totalMissedWaitTime / missedWaitCount : 0;
    document.getElementById('avgWaitBeforeHangup').textContent = formatTime(avgWait);
}

// ========================================
// CHARTS
// ========================================
function updateHourlyChart(data) {
    const inCalls = data.filter(row => row.Direction === 'In');
    const outCalls = data.filter(row => row.Direction === 'Out');

    const hours = [];
    for (let h = 7; h <= 18; h++) hours.push(h);

    const inCountsPerHour = hours.map(h => inCalls.filter(c => getHour(c.CallDateTime) === h).length);
    const outCountsPerHour = hours.map(h => outCalls.filter(c => getHour(c.CallDateTime) === h).length);

    const labels = hours.map(h => {
        const suffix = h >= 12 ? 'PM' : 'AM';
        const hour12 = h > 12 ? h - 12 : (h === 0 ? 12 : h);
        return `${hour12}${suffix}`;
    });

    if (hourlyChart) hourlyChart.destroy();

    const ctx = document.getElementById('hourlyChart').getContext('2d');
    hourlyChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Calls In',
                    data: inCountsPerHour,
                    borderColor: '#1565c0',
                    backgroundColor: 'rgba(21, 101, 192, 0.08)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointHoverRadius: 8,
                    pointBackgroundColor: '#1565c0',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    borderWidth: 3
                },
                {
                    label: 'Calls Out',
                    data: outCountsPerHour,
                    borderColor: '#7cb342',
                    backgroundColor: 'rgba(124, 179, 66, 0.08)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 5,
                    pointHoverRadius: 8,
                    pointBackgroundColor: '#7cb342',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    borderWidth: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top', labels: { usePointStyle: true, padding: 20 } }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: 'rgba(0, 0, 0, 0.06)' } },
                x: { grid: { display: false } }
            }
        }
    });
}

function updateWeekTrendCharts() {
    const section = document.getElementById('weekTrendsSection');
    if (!section) return;

    if (availableWeeks.length < 2) {
        section.style.display = 'none';
        return;
    }
    section.style.display = 'block';

    const weeklyData = availableWeeks.map(week => {
        let weekCalls = rawData.filter(row => {
            const date = getDateObj(row.CallDateTime);
            if (!date) return false;
            return date >= week.start && date <= week.end &&
                   isWithinOpeningHours(row.CallDateTime) &&
                   !isInternalCall(row) &&
                   row.Direction !== 'Int';
        });

        weekCalls = filterByLocation(weekCalls, currentGlobalLocation);

        if (currentQueueFilter !== 'all') {
            if (currentQueueFilter === 'noqueue') {
                weekCalls = weekCalls.filter(row => !row.queueName);
            } else {
                const targetQueue = CONFIG.QUEUE_MAP[currentQueueFilter];
                weekCalls = weekCalls.filter(row => row.queueName === targetQueue);
            }
        }

        const inCalls = weekCalls.filter(c => c.Direction === 'In');
        const total = inCalls.length;
        const missed = inCalls.filter(c => c.queueName && (!c.TimeToAnswer || c.TimeToAnswer === 0)).length;
        const missedPct = total > 0 ? (missed / total * 100) : 0;

        const answeredCalls = inCalls.filter(c => c.TimeToAnswer > 0);
        const avgWait = answeredCalls.length > 0
            ? answeredCalls.reduce((sum, c) => sum + (c.TimeToAnswer || 0), 0) / answeredCalls.length
            : 0;

        const withinTarget = answeredCalls.filter(c => c.TimeToAnswer <= serviceLevelTarget).length;
        const serviceLevel = total > 0 ? (withinTarget / total * 100) : 0;

        return { total, missed, missedPct, avgWait, serviceLevel };
    });

    const labels = availableWeeks.map(w => w.label.replace('Week of ', ''));

    // Volume Chart
    if (weekTrendVolumeChart) weekTrendVolumeChart.destroy();

    const volumeCtx = document.getElementById('weekTrendVolumeChart').getContext('2d');
    weekTrendVolumeChart = new Chart(volumeCtx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Total Calls',
                    data: weeklyData.map(w => w.total),
                    borderColor: '#1565c0',
                    backgroundColor: 'rgba(21, 101, 192, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    borderWidth: 3
                },
                {
                    label: 'Missed Calls',
                    data: weeklyData.map(w => w.missed),
                    borderColor: '#c62828',
                    backgroundColor: 'rgba(198, 40, 40, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    borderWidth: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top', labels: { usePointStyle: true } }
            },
            scales: {
                y: { beginAtZero: true, title: { display: true, text: 'Call Count' } }
            }
        }
    });

    // Quality Chart
    if (weekTrendQualityChart) weekTrendQualityChart.destroy();

    const qualityCtx = document.getElementById('weekTrendQualityChart').getContext('2d');
    weekTrendQualityChart = new Chart(qualityCtx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Service Level %',
                    data: weeklyData.map(w => parseFloat(w.serviceLevel.toFixed(1))),
                    borderColor: '#1565c0',
                    backgroundColor: 'rgba(21, 101, 192, 0.1)',
                    tension: 0.4,
                    pointRadius: 6,
                    borderWidth: 3,
                    yAxisID: 'y'
                },
                {
                    label: 'Missed %',
                    data: weeklyData.map(w => parseFloat(w.missedPct.toFixed(1))),
                    borderColor: '#c62828',
                    tension: 0.4,
                    pointRadius: 6,
                    borderWidth: 3,
                    borderDash: [5, 5],
                    yAxisID: 'y'
                },
                {
                    label: 'Avg Wait (sec)',
                    data: weeklyData.map(w => Math.round(w.avgWait)),
                    borderColor: '#f57c00',
                    tension: 0.4,
                    pointRadius: 6,
                    borderWidth: 3,
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'top', labels: { usePointStyle: true } }
            },
            scales: {
                y: {
                    type: 'linear',
                    position: 'left',
                    beginAtZero: true,
                    max: 100,
                    title: { display: true, text: 'Percentage' }
                },
                y1: {
                    type: 'linear',
                    position: 'right',
                    beginAtZero: true,
                    grid: { drawOnChartArea: false },
                    title: { display: true, text: 'Seconds' }
                }
            }
        }
    });
}

// ========================================
// TABLES
// ========================================
function updateSiteBreakdown(calls) {
    const sites = [
        { name: 'Crace', patterns: ['crace'] },
        { name: 'Denman', patterns: ['denman'] },
        { name: 'Lyneham', patterns: ['lyneham'] }
    ];

    function calcSiteMetrics(siteCalls) {
        const total = siteCalls.length;
        const answered = siteCalls.filter(c => c.TimeToAnswer > 0).length;
        const missed = siteCalls.filter(c => c.queueName && (!c.TimeToAnswer || c.TimeToAnswer === 0)).length;
        const missedPct = total > 0 ? ((missed / total) * 100).toFixed(1) : 0;

        const answeredCalls = siteCalls.filter(c => c.TimeToAnswer > 0);
        const avgWait = answeredCalls.length > 0
            ? answeredCalls.reduce((sum, c) => sum + (c.TimeToAnswer || 0), 0) / answeredCalls.length
            : null;
        const maxWait = answeredCalls.length > 0
            ? Math.max(...answeredCalls.map(c => c.TimeToAnswer || 0))
            : null;

        const callLengths = answeredCalls
            .map(c => (c.CallDuration || 0) - (c.TimeToAnswer || 0))
            .filter(l => l > 0);
        const avgCallLength = callLengths.length > 0
            ? callLengths.reduce((a, b) => a + b, 0) / callLengths.length
            : null;

        return { total, answered, missed, missedPct, avgWait, maxWait, avgCallLength };
    }

    let html = '';

    sites.forEach(site => {
        const siteCalls = calls.filter(c => {
            const office = (c.OfficeName || '').toLowerCase();
            return site.patterns.some(p => office.includes(p));
        });

        const metrics = calcSiteMetrics(siteCalls);

        html += '<tr>';
        html += `<td style="text-align: left; font-weight: 500;">${site.name}</td>`;
        html += `<td>${metrics.total}</td>`;
        html += `<td>${metrics.answered}</td>`;
        html += `<td style="color: ${metrics.missed > 0 ? '#e74c3c' : 'inherit'};">${metrics.missed}</td>`;
        html += `<td style="color: ${parseFloat(metrics.missedPct) > 5 ? '#e74c3c' : 'inherit'};">${metrics.missedPct}%</td>`;
        html += `<td>${formatTime(metrics.avgWait)}</td>`;
        html += `<td>${formatTime(metrics.maxWait)}</td>`;
        html += `<td>${formatTime(metrics.avgCallLength)}</td>`;
        html += '</tr>';
    });

    const totalMetrics = calcSiteMetrics(calls);
    html += '<tr style="font-weight: 600; background: #f8f9fa;">';
    html += '<td style="text-align: left;">Total</td>';
    html += `<td>${totalMetrics.total}</td>`;
    html += `<td>${totalMetrics.answered}</td>`;
    html += `<td style="color: ${totalMetrics.missed > 0 ? '#e74c3c' : 'inherit'};">${totalMetrics.missed}</td>`;
    html += `<td style="color: ${parseFloat(totalMetrics.missedPct) > 5 ? '#e74c3c' : 'inherit'};">${totalMetrics.missedPct}%</td>`;
    html += `<td>${formatTime(totalMetrics.avgWait)}</td>`;
    html += `<td>${formatTime(totalMetrics.maxWait)}</td>`;
    html += `<td>${formatTime(totalMetrics.avgCallLength)}</td>`;
    html += '</tr>';

    document.getElementById('siteTableBody').innerHTML = html;
}

function updateDailyTable() {
    let filteredData = getGlobalFilteredData();

    if (currentDailyDirection === 'in') {
        filteredData = filteredData.filter(row => row.Direction === 'In');
    } else if (currentDailyDirection === 'out') {
        filteredData = filteredData.filter(row => row.Direction === 'Out');
    }

    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dayData = {};
    dayNames.forEach((name, idx) => dayData[idx] = []);

    filteredData.forEach(row => {
        const day = getDayOfWeek(row.CallDateTime);
        if (day !== null && dayData[day]) dayData[day].push(row);
    });

    function calcMetrics(calls) {
        const total = calls.length;
        const answered = calls.filter(c => c.TimeToAnswer > 0).length;
        const missed = total - answered;
        const missedPct = total > 0 ? ((missed / total) * 100).toFixed(1) : '-';

        const answeredCalls = calls.filter(c => c.TimeToAnswer > 0);
        const avgWait = answeredCalls.length > 0
            ? answeredCalls.reduce((sum, c) => sum + (c.TimeToAnswer || 0), 0) / answeredCalls.length
            : null;
        const maxWait = answeredCalls.length > 0
            ? Math.max(...answeredCalls.map(c => c.TimeToAnswer || 0))
            : null;

        const callLengths = answeredCalls
            .map(c => (c.CallDuration || 0) - (c.TimeToAnswer || 0))
            .filter(l => l > 0);
        const avgCallLength = callLengths.length > 0
            ? callLengths.reduce((a, b) => a + b, 0) / callLengths.length
            : null;

        return { total, answered, missed, missedPct, avgWait, maxWait, avgCallLength };
    }

    const metrics = {};
    const displayOrder = [1, 2, 3, 4, 5, 6, 0];
    displayOrder.forEach(day => metrics[day] = calcMetrics(dayData[day]));
    const weekMetrics = calcMetrics(filteredData);

    const rows = [
        { label: 'Total Calls', key: 'total', format: v => v },
        { label: 'Answered', key: 'answered', format: v => v },
        { label: 'Missed', key: 'missed', format: v => v },
        { label: 'Missed %', key: 'missedPct', format: v => v === '-' ? v : v + '%' },
        { label: 'Avg Wait', key: 'avgWait', format: formatTimeShort },
        { label: 'Max Wait', key: 'maxWait', format: formatTimeShort },
        { label: 'Avg Call', key: 'avgCallLength', format: formatTimeShort }
    ];

    let html = '';
    rows.forEach(row => {
        html += '<tr>';
        html += `<td style="text-align: left; font-weight: 500;">${row.label}</td>`;
        displayOrder.forEach(day => {
            const val = metrics[day][row.key];
            html += `<td>${row.format(val)}</td>`;
        });
        html += `<td style="font-weight: 600; background: #f8f9fa;">${row.format(weekMetrics[row.key])}</td>`;
        html += '</tr>';
    });

    document.getElementById('dailyTableBody').innerHTML = html;
}

function updateStaffTable(data) {
    const locationFilteredData = filterByLocation(data, currentLocationFilter);

    const staffStats = {};

    locationFilteredData.forEach(row => {
        const name = row.UserName;
        if (!name || name === '0' || name === 0) return;

        if (!staffStats[name]) {
            staffStats[name] = {
                name: name,
                callsIn: 0,
                callsOut: 0,
                totalPickupTime: 0,
                pickupCount: 0,
                totalCallLengthIn: 0,
                callLengthCountIn: 0,
                totalCallLengthOut: 0,
                callLengthCountOut: 0
            };
        }

        const stats = staffStats[name];

        if (row.Direction === 'In') {
            stats.callsIn++;
            if (row.TimeToAnswer > 0) {
                stats.totalPickupTime += row.TimeToAnswer;
                stats.pickupCount++;
                const callLength = (row.CallDuration || 0) - (row.TimeToAnswer || 0);
                if (callLength > 0) {
                    stats.totalCallLengthIn += callLength;
                    stats.callLengthCountIn++;
                }
            }
        } else if (row.Direction === 'Out') {
            stats.callsOut++;
            if (row.CallDuration > 0) {
                stats.totalCallLengthOut += row.CallDuration;
                stats.callLengthCountOut++;
            }
        }
    });

    const staffArray = Object.values(staffStats)
        .map(s => ({
            ...s,
            totalCalls: s.callsIn + s.callsOut,
            avgPickup: s.pickupCount > 0 ? s.totalPickupTime / s.pickupCount : null,
            avgCallLengthIn: s.callLengthCountIn > 0 ? s.totalCallLengthIn / s.callLengthCountIn : null,
            avgCallLengthOut: s.callLengthCountOut > 0 ? s.totalCallLengthOut / s.callLengthCountOut : null
        }))
        .sort((a, b) => b.totalCalls - a.totalCalls);

    let html = '';
    staffArray.forEach(staff => {
        html += '<tr>';
        html += `<td style="text-align: left;">${staff.name}</td>`;
        html += `<td>${staff.callsIn}</td>`;
        html += `<td>${staff.callsOut}</td>`;
        html += `<td style="font-weight: 600;">${staff.totalCalls}</td>`;
        html += `<td>${formatTime(staff.avgPickup)}</td>`;
        html += `<td>${formatTime(staff.avgCallLengthIn)}</td>`;
        html += `<td>${formatTime(staff.avgCallLengthOut)}</td>`;
        html += '</tr>';
    });

    document.getElementById('staffTableBody').innerHTML = html;
}

// ========================================
// HEATMAPS (Refactored to reduce duplication)
// ========================================
function updateHeatmaps(data) {
    let heatmapData = data;
    if (currentHeatmapLocation !== 'all') {
        heatmapData = filterByLocation(data, currentHeatmapLocation);
    }

    const inCalls = heatmapData.filter(row => row.Direction === 'In');
    const outCalls = heatmapData.filter(row => row.Direction === 'Out');

    renderHeatmap('heatmapIn', inCalls);
    renderHeatmap('heatmapOut', outCalls);

    renderWaitTimeHeatmap('heatmapMaxWait', inCalls, 'max');
    renderWaitTimeHeatmap('heatmapAvgWait', inCalls, 'avg');

    const missedCalls = inCalls.filter(c => !c.TimeToAnswer || c.TimeToAnswer === 0);
    renderMissedHeatmap('heatmapMissed', missedCalls, inCalls, 'count');
    renderMissedHeatmap('heatmapMissedRate', missedCalls, inCalls, 'rate');

    updateMissedByQueue(inCalls);
}

function renderHeatmap(elementId, calls) {
    const { startSlot, endSlot, days } = CONFIG.HEATMAP;
    const dayNames = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    const counts = {};
    let maxCount = 0;

    for (let slot = startSlot; slot <= endSlot; slot++) {
        counts[slot] = {};
        days.forEach(day => counts[slot][day] = 0);
    }

    calls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            counts[slot][day]++;
            maxCount = Math.max(maxCount, counts[slot][day]);
        }
    });

    let html = '<div class="heatmap-cell heatmap-header"></div>';
    days.forEach(day => html += `<div class="heatmap-cell heatmap-header">${dayNames[day]}</div>`);

    for (let slot = startSlot; slot <= endSlot; slot++) {
        html += `<div class="heatmap-cell heatmap-time">${formatTimeSlot(slot)}</div>`;
        days.forEach(day => {
            const count = counts[slot][day];
            const heatLevel = maxCount > 0 ? Math.min(7, Math.ceil((count / maxCount) * 7)) : 0;
            html += `<div class="heatmap-cell heat-${heatLevel}" title="${count} calls">${count || ''}</div>`;
        });
    }

    document.getElementById(elementId).innerHTML = html;
}

function renderWaitTimeHeatmap(elementId, calls, mode) {
    const { startSlot, endSlot, days } = CONFIG.HEATMAP;
    const dayNames = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    const answeredCalls = calls.filter(c => c.TimeToAnswer > 0);

    const waitTimes = {};
    for (let slot = startSlot; slot <= endSlot; slot++) {
        waitTimes[slot] = {};
        days.forEach(day => waitTimes[slot][day] = []);
    }

    answeredCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            waitTimes[slot][day].push(call.TimeToAnswer);
        }
    });

    const values = {};
    for (let slot = startSlot; slot <= endSlot; slot++) {
        values[slot] = {};
        days.forEach(day => {
            const times = waitTimes[slot][day];
            if (times.length === 0) {
                values[slot][day] = null;
            } else if (mode === 'max') {
                values[slot][day] = Math.max(...times);
            } else {
                values[slot][day] = times.reduce((a, b) => a + b, 0) / times.length;
            }
        });
    }

    function getWaitHeatLevel(seconds) {
        if (seconds === null) return 0;
        if (seconds < 30) return 1;
        if (seconds < 60) return 2;
        if (seconds < 90) return 3;
        if (seconds < 120) return 4;
        if (seconds < 180) return 5;
        if (seconds < 300) return 6;
        return 7;
    }

    let html = '<div class="heatmap-cell heatmap-header"></div>';
    days.forEach(day => html += `<div class="heatmap-cell heatmap-header">${dayNames[day]}</div>`);

    for (let slot = startSlot; slot <= endSlot; slot++) {
        html += `<div class="heatmap-cell heatmap-time">${formatTimeSlot(slot)}</div>`;
        days.forEach(day => {
            const val = values[slot][day];
            const heatLevel = getWaitHeatLevel(val);
            const displayVal = val !== null ? formatTimeShort(val) : '';
            const titleText = val !== null ? `${mode === 'max' ? 'Max' : 'Avg'}: ${formatTime(val)}` : 'No calls';
            html += `<div class="heatmap-cell wait-heat-${heatLevel}" title="${titleText}">${displayVal}</div>`;
        });
    }

    document.getElementById(elementId).innerHTML = html;
}

function renderMissedHeatmap(elementId, missedCalls, allCalls, type) {
    const { startSlot, endSlot, days } = CONFIG.HEATMAP;
    const dayNames = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    const missedCounts = {};
    const totalCounts = {};

    for (let slot = startSlot; slot <= endSlot; slot++) {
        missedCounts[slot] = {};
        totalCounts[slot] = {};
        days.forEach(day => {
            missedCounts[slot][day] = 0;
            totalCounts[slot][day] = 0;
        });
    }

    missedCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            missedCounts[slot][day]++;
        }
    });

    allCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            totalCounts[slot][day]++;
        }
    });

    let maxVal = 0;
    const values = {};
    for (let slot = startSlot; slot <= endSlot; slot++) {
        values[slot] = {};
        days.forEach(day => {
            if (type === 'count') {
                values[slot][day] = missedCounts[slot][day];
            } else {
                const total = totalCounts[slot][day];
                values[slot][day] = total > 0 ? (missedCounts[slot][day] / total) * 100 : null;
            }
            if (values[slot][day] !== null) {
                maxVal = Math.max(maxVal, values[slot][day]);
            }
        });
    }

    function getMissedHeatLevel(value, isRate) {
        if (value === null || value === 0) return 0;
        if (isRate) {
            if (value < 10) return 1;
            if (value < 20) return 2;
            if (value < 30) return 3;
            if (value < 40) return 4;
            if (value < 50) return 5;
            if (value < 70) return 6;
            return 7;
        } else {
            if (maxVal === 0) return 0;
            return Math.min(7, Math.ceil((value / maxVal) * 7));
        }
    }

    let html = '<div class="heatmap-cell heatmap-header"></div>';
    days.forEach(day => html += `<div class="heatmap-cell heatmap-header">${dayNames[day]}</div>`);

    for (let slot = startSlot; slot <= endSlot; slot++) {
        html += `<div class="heatmap-cell heatmap-time">${formatTimeSlot(slot)}</div>`;
        days.forEach(day => {
            const val = values[slot][day];
            const heatLevel = getMissedHeatLevel(val, type === 'rate');
            let displayVal = '';
            let titleText = '';

            if (type === 'count') {
                displayVal = val > 0 ? val : '';
                titleText = `${val} missed calls`;
            } else {
                displayVal = val !== null && val > 0 ? val.toFixed(0) + '%' : '';
                titleText = val !== null ? `${val.toFixed(1)}% missed (${missedCounts[slot][day]}/${totalCounts[slot][day]})` : 'No calls';
            }

            html += `<div class="heatmap-cell missed-heat-${heatLevel}" title="${titleText}">${displayVal}</div>`;
        });
    }

    document.getElementById(elementId).innerHTML = html;
}

function updateMissedByQueue(calls) {
    const inCalls = calls.filter(c => c.Direction === 'In');

    const queues = [
        { key: 'appointments', name: 'Appointments', filter: 'Appointments' },
        { key: 'vasectomy', name: 'Canberra Vasectomy', filter: 'Canberra Vasectomy' },
        { key: 'general', name: 'General Enquiries', filter: 'General Enquiries' },
        { key: 'health', name: 'Health Professionals', filter: 'Health Professionals' },
        { key: 'noqueue', name: 'Direct Calls', filter: null }
    ];

    let html = '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">';
    html += '<thead><tr style="background: #f8f9fa;">';
    html += '<th style="text-align: left; padding: 8px; border-bottom: 2px solid #dee2e6;">Queue</th>';
    html += '<th style="text-align: center; padding: 8px; border-bottom: 2px solid #dee2e6;">Total Calls</th>';
    html += '<th style="text-align: center; padding: 8px; border-bottom: 2px solid #dee2e6;">Answered</th>';
    html += '<th style="text-align: center; padding: 8px; border-bottom: 2px solid #dee2e6;">Missed</th>';
    html += '<th style="text-align: center; padding: 8px; border-bottom: 2px solid #dee2e6;">Miss Rate</th>';
    html += '</tr></thead><tbody>';

    let totalAll = 0, answeredAll = 0, missedAll = 0;

    queues.forEach(queue => {
        let queueCalls;
        if (queue.filter === null) {
            queueCalls = inCalls.filter(c => !c.queueName);
        } else {
            queueCalls = inCalls.filter(c => c.queueName === queue.filter);
        }

        const total = queueCalls.length;
        const missed = queueCalls.filter(c => !c.TimeToAnswer || c.TimeToAnswer === 0).length;
        const answered = total - missed;
        const missRate = total > 0 ? ((missed / total) * 100).toFixed(1) : 0;

        totalAll += total;
        answeredAll += answered;
        missedAll += missed;

        const rateColor = parseFloat(missRate) > 10 ? '#e53e3e' : (parseFloat(missRate) > 5 ? '#f57c00' : 'inherit');

        html += '<tr style="border-bottom: 1px solid #dee2e6;">';
        html += `<td style="text-align: left; padding: 8px;">${queue.name}</td>`;
        html += `<td style="text-align: center; padding: 8px;">${total}</td>`;
        html += `<td style="text-align: center; padding: 8px; color: #1565c0;">${answered}</td>`;
        html += `<td style="text-align: center; padding: 8px; color: ${missed > 0 ? '#e53e3e' : 'inherit'};">${missed}</td>`;
        html += `<td style="text-align: center; padding: 8px; font-weight: 600; color: ${rateColor};">${missRate}%</td>`;
        html += '</tr>';
    });

    const totalMissRate = totalAll > 0 ? ((missedAll / totalAll) * 100).toFixed(1) : 0;
    html += '<tr style="background: #f8f9fa; font-weight: 600;">';
    html += '<td style="text-align: left; padding: 8px;">TOTAL</td>';
    html += `<td style="text-align: center; padding: 8px;">${totalAll}</td>`;
    html += `<td style="text-align: center; padding: 8px; color: #1565c0;">${answeredAll}</td>`;
    html += `<td style="text-align: center; padding: 8px; color: ${missedAll > 0 ? '#e53e3e' : 'inherit'};">${missedAll}</td>`;
    html += `<td style="text-align: center; padding: 8px;">${totalMissRate}%</td>`;
    html += '</tr>';

    html += '</tbody></table>';

    document.getElementById('missedByQueue').innerHTML = html;
}

// ========================================
// EXPORT FUNCTIONS
// ========================================
function exportToCSV() {
    if (rawData.length === 0) {
        showError('No data to export. Please load data first.');
        return;
    }

    const filteredData = getGlobalFilteredData();
    const incomingCalls = filteredData.filter(row => row.Direction === 'In');

    // Calculate summary stats
    const total = incomingCalls.length;
    const answered = incomingCalls.filter(c => c.TimeToAnswer > 0).length;
    const missed = incomingCalls.filter(c => c.queueName && (!c.TimeToAnswer || c.TimeToAnswer === 0)).length;

    // Create CSV content
    let csv = 'YourGP Phone Dashboard Export\n';
    csv += `Date Range,${document.getElementById('weekInfo').textContent}\n`;
    csv += `Generated,${new Date().toLocaleString()}\n\n`;

    csv += 'Summary Metrics\n';
    csv += `Total Calls,${total}\n`;
    csv += `Answered,${answered}\n`;
    csv += `Missed,${missed}\n`;
    csv += `Missed %,${total > 0 ? ((missed/total)*100).toFixed(1) : 0}%\n\n`;

    csv += 'Raw Data\n';
    csv += 'Date,Time,Direction,Queue,Office,Staff,Duration (sec),Wait Time (sec),Answered\n';

    filteredData.forEach(row => {
        const date = getDateObj(row.CallDateTime);
        const dateStr = date ? date.toLocaleDateString('en-AU') : '';
        const timeStr = date ? date.toLocaleTimeString('en-AU') : '';
        csv += `${dateStr},${timeStr},${row.Direction},${row.queueName || 'None'},${row.OfficeName || ''},${row.UserName || ''},${row.CallDuration},${row.TimeToAnswer},${row.TimeToAnswer > 0 ? 'Yes' : 'No'}\n`;
    });

    // Download
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `phone-dashboard-export-${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function exportToPDF() {
    // For PDF export, we'll trigger print which allows saving as PDF
    showError('To export as PDF, use the Print button and select "Save as PDF" as your printer.');
}
