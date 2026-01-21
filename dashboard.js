// YourGP Phone Dashboard JavaScript

// Global state
let rawData = [];
let queueData = {}; // Map of CallGUID -> queue name
let currentQueueFilter = 'all'; // 'all', 'appointments', 'vasectomy', 'general', 'health', 'noqueue'
let currentDailyDirection = 'all'; // 'all', 'in', or 'out'
let serviceLevelTarget = 90; // Default 90 seconds
let currentLocationFilter = 'all'; // 'all', 'crace', 'denman', or 'lyneham' (for staff table only)
let currentGlobalLocation = 'all'; // 'all', 'crace', 'denman', or 'lyneham' (for entire dashboard)
let hourlyChart = null;
let callbackWindowHours = 24; // Default 24 hours for callback/FCR calculations

// OneDrive Configuration - Replace with your Azure App ID
const ONEDRIVE_CLIENT_ID = "bdf5829d-49a6-4bed-aa55-89bf6ef866bc";

// Launch OneDrive File Picker
function launchOneDrivePicker() {
    const btn = document.getElementById('oneDriveBtn');
    btn.textContent = 'Connecting...';

    OneDrive.open({
        clientId: ONEDRIVE_CLIENT_ID,
        action: "download",
        multiSelect: true,
        advanced: { filter: ".csv" },
        success: handleOneDriveFiles,
        cancel: function() {
            btn.textContent = 'Open from OneDrive';
        },
        error: function(err) {
            btn.textContent = 'Open from OneDrive';
            alert('OneDrive error: ' + (err.message || 'Unknown error'));
        }
    });
}

// Handle files selected from OneDrive
async function handleOneDriveFiles(response) {
    if (!response.value || response.value.length === 0) {
        console.log('No files selected from OneDrive');
        return;
    }

    document.getElementById('loading').style.display = 'block';
    document.getElementById('noData').style.display = 'none';
    document.getElementById('dashboard').style.display = 'none';

    try {
        queueData = {};
        let mainFileData = null;

        for (const file of response.value) {
            const url = file['@microsoft.graph.downloadUrl'] || file['@content.downloadUrl'];

            if (!url) {
                console.error('No download URL for file:', file.name);
                continue;
            }

            console.log('Downloading from OneDrive:', file.name);

            const res = await fetch(url);
            const blob = await res.blob();
            const rows = await readCsvFromBlob(blob);

            // Determine if this is a queue file or main export
            if (file.name.toLowerCase().startsWith('callqueue')) {
                const queueName = extractQueueName(file.name);
                rows.forEach(row => {
                    if (row.CallGUID) {
                        queueData[row.CallGUID] = queueName;
                    }
                });
                console.log(`Loaded ${rows.length} records from queue: ${queueName}`);
            } else {
                mainFileData = rows;
            }
        }

        if (!mainFileData) {
            throw new Error('No main export file found. Please select a file starting with "Export".');
        }

        // Process the main data
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
        alert('Error processing files from OneDrive: ' + err.message);
        document.getElementById('loading').style.display = 'none';
        document.getElementById('noData').style.display = 'block';
    }

    document.getElementById('oneDriveBtn').textContent = 'Open from OneDrive';
}

// Helper function to read CSV from Blob (for OneDrive files)
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

// Internal extensions to exclude from incoming call statistics
const INTERNAL_EXTENSIONS = [
    'Nurse 1',
    'Crace',
    'Lyneham - Rec 1',
    'Crace - Rec 1',
    'Crace - Rec 2',
    'Crace Office',
    'Nurse 5 (TR1)',
    'Nurse 2',
    'Nurse Consult',
    'Lyneham - Nurse',
    'Nurse 3 (TR2)'
];

// Check if a call is to an internal extension (should be excluded from incoming stats)
function isInternalCall(row) {
    // Check if UserName matches any internal extension
    const userName = row.UserName || '';
    return INTERNAL_EXTENSIONS.some(ext =>
        userName.toLowerCase() === ext.toLowerCase()
    );
}

// Opening hours filter - excludes calls outside business hours
function isWithinOpeningHours(dateValue) {
    const date = getDateObj(dateValue);
    if (!date) return false;

    const dayOfWeek = date.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
    const hours = date.getHours();
    const minutes = date.getMinutes();
    const timeInMinutes = hours * 60 + minutes;

    // Sunday = closed
    if (dayOfWeek === 0) return false;

    // Saturday: 9:00am (540) - 12:30pm (750)
    if (dayOfWeek === 6) {
        return timeInMinutes >= 540 && timeInMinutes <= 750;
    }

    // Mon-Fri: 7:30am (450) - 5:30pm (1050)
    return timeInMinutes >= 450 && timeInMinutes <= 1050;
}

// File upload handler - supports multiple files (main export + queue CSVs)
document.getElementById('fileInput').addEventListener('change', async function(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    document.getElementById('loading').style.display = 'block';
    document.getElementById('noData').style.display = 'none';
    document.getElementById('dashboard').style.display = 'none';

    try {
        // Separate main export from queue files
        let mainFile = null;
        const queueFiles = [];

        files.forEach(file => {
            if (file.name.startsWith('Export') || file.name.startsWith('export')) {
                mainFile = file;
            } else if (file.name.startsWith('CallQueue') || file.name.startsWith('callqueue')) {
                queueFiles.push(file);
            }
        });

        // If only one file and it's not clearly identifiable, assume it's the main file
        if (!mainFile && files.length === 1) {
            mainFile = files[0];
        }

        if (!mainFile) {
            throw new Error('No main export file found. Please include a file starting with "Export".');
        }

        // Reset queue data
        queueData = {};

        // Load queue files first to build lookup map
        for (const queueFile of queueFiles) {
            const queueRows = await readCsvFile(queueFile);

            // Extract queue name from filename: CallQueue_Detailed_*_<QueueName>.csv
            const queueName = extractQueueName(queueFile.name);

            queueRows.forEach(row => {
                if (row.CallGUID) {
                    queueData[row.CallGUID] = queueName;
                }
            });

            console.log(`Loaded ${queueRows.length} records from queue: ${queueName}`);
        }

        // Load main export file
        rawData = await readCsvFile(mainFile);

        // Convert numeric fields back to numbers (raw:true keeps everything as strings)
        rawData = rawData.map(row => ({
            ...row,
            CallDuration: parseFloat(row.CallDuration) || 0,
            TimeToAnswer: parseFloat(row.TimeToAnswer) || 0,
            BillableTime: parseFloat(row.BillableTime) || 0,
            // Add queue name from lookup
            queueName: queueData[row.CallGUID] || null
        }));

        // Count queue matches
        const queuedCount = rawData.filter(r => r.queueName).length;
        console.log(`Loaded ${rawData.length} records from main CSV, ${queuedCount} matched to queues`);
        console.log('Sample record:', rawData[0]);

        processAndDisplay();
    } catch (err) {
        console.error('Error reading files:', err);
        alert('Error reading CSV files: ' + err.message);
        document.getElementById('loading').style.display = 'none';
        document.getElementById('noData').style.display = 'block';
    }
});

// Helper function to read a CSV file and return rows
function readCsvFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                // Use raw:true to prevent SheetJS from auto-converting DD/MM/YYYY to Excel serial numbers
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

// Extract queue name from filename like "CallQueue_Detailed_20260111_20260118_Appointments.csv"
function extractQueueName(filename) {
    // Remove .csv extension
    const withoutExt = filename.replace(/\.csv$/i, '');
    // Split by underscore and get the last part (queue name)
    const parts = withoutExt.split('_');
    // Find where the date parts end (they're 8 digits)
    let queueNameParts = [];
    let foundDates = 0;
    for (let i = 0; i < parts.length; i++) {
        if (/^\d{8}$/.test(parts[i])) {
            foundDates++;
        } else if (foundDates >= 2) {
            // After two date parts, everything else is the queue name
            queueNameParts.push(parts[i]);
        }
    }
    return queueNameParts.join(' ') || 'Unknown';
}

// Queue filter toggle
function setQueueFilter(queue) {
    currentQueueFilter = queue;
    // Update button active states
    const buttons = ['queueAll', 'queueAppointments', 'queueVasectomy', 'queueGeneral', 'queueHealth', 'queueNone'];
    const values = ['all', 'appointments', 'vasectomy', 'general', 'health', 'noqueue'];
    buttons.forEach((btnId, idx) => {
        const btn = document.getElementById(btnId);
        if (btn) {
            btn.classList.toggle('active', values[idx] === queue);
        }
    });
    if (rawData.length > 0) {
        processAndDisplay();
    }
}

// Daily direction toggle
function setDailyDirection(direction) {
    currentDailyDirection = direction;
    document.getElementById('tabInOut').classList.toggle('active', direction === 'all');
    document.getElementById('tabIn').classList.toggle('active', direction === 'in');
    document.getElementById('tabOut').classList.toggle('active', direction === 'out');
    if (rawData.length > 0) {
        updateDailyTable();
    }
}

// Service level target update
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
    // Update header text
    const windowText = callbackWindowHours === 24 ? '24h' : (callbackWindowHours / 24) + ' day';
    document.getElementById('followupHeader').textContent = `Missed Call Follow-up (${windowText} Window)`;
    // Update tooltips
    const windowDesc = callbackWindowHours === 24 ? '24 hours' : (callbackWindowHours / 24) + ' days';
    document.getElementById('fcrRate').parentElement.title = `First Call Resolution - estimated percentage of callers who didn't need to call back within ${windowDesc}`;
    document.getElementById('callbackRate').parentElement.title = `Percentage of unique callers who called more than once within ${windowDesc}`;
    // Recalculate if data loaded
    if (rawData.length > 0) {
        const filteredData = getGlobalFilteredData();
        const incomingCalls = filteredData.filter(row => row.Direction === 'In');
        updateSummaryMetrics(incomingCalls);
        updateMissedCallFollowup(incomingCalls);
    }
}

// Global location filter - applies to entire dashboard
function setGlobalLocation() {
    currentGlobalLocation = document.getElementById('globalLocationFilter').value;
    if (rawData.length > 0) {
        processAndDisplay();
    }
}

// Location filter for staff table (additional filter on top of global)
function updateStaffTableFilter() {
    currentLocationFilter = document.getElementById('locationFilter').value;
    if (rawData.length > 0) {
        let filteredData = getGlobalFilteredData();
        updateStaffTable(filteredData);
    }
}

// Filter data by location
function filterByLocation(data, location) {
    if (location === 'all') return data;
    const locationLower = location.toLowerCase();
    return data.filter(row => {
        const office = (row.OfficeName || '').toLowerCase();
        return office.includes(locationLower);
    });
}

// Get data filtered by global location and queue filter
function getGlobalFilteredData() {
    let filteredData = rawData;

    // Apply opening hours filter first (Mon-Fri 7:30am-5:30pm, Sat 9am-12:30pm, Sun closed)
    filteredData = filteredData.filter(row => isWithinOpeningHours(row.CallDateTime));

    // Exclude internal calls (calls to nurse stations, reception desks, etc.)
    filteredData = filteredData.filter(row => !isInternalCall(row));

    // Apply global location filter
    filteredData = filterByLocation(filteredData, currentGlobalLocation);

    // Apply queue filter
    if (currentQueueFilter !== 'all') {
        if (currentQueueFilter === 'noqueue') {
            // Show calls that didn't enter any queue
            filteredData = filteredData.filter(row => !row.queueName);
        } else {
            // Map filter values to actual queue names
            const queueMap = {
                'appointments': 'Appointments',
                'vasectomy': 'Canberra Vasectomy',
                'general': 'General Enquiries',
                'health': 'Health Professionals'
            };
            const targetQueue = queueMap[currentQueueFilter];
            filteredData = filteredData.filter(row => row.queueName === targetQueue);
        }
    }

    return filteredData;
}

// Format seconds to mm:ss or hh:mm:ss
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

// Format time for compact display
function formatTimeShort(seconds) {
    if (seconds === null || seconds === undefined || isNaN(seconds)) return '-';
    seconds = Math.round(seconds);
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins}:${secs.toString().padStart(2, '0')}`;
}

// Get date object from various formats
function getDateObj(dateValue) {
    if (dateValue instanceof Date) return dateValue;
    if (typeof dateValue === 'number') {
        // Excel serial number - convert to local date (not UTC)
        // Excel epoch is Dec 30, 1899 (day 0)
        const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
        const days = Math.floor(dateValue);
        const timeFraction = dateValue - days;
        const date = new Date(excelEpoch.getTime() + days * 86400 * 1000);
        // Add time portion
        const totalSeconds = Math.round(timeFraction * 86400);
        date.setHours(Math.floor(totalSeconds / 3600));
        date.setMinutes(Math.floor((totalSeconds % 3600) / 60));
        date.setSeconds(totalSeconds % 60);
        return date;
    }
    if (typeof dateValue === 'string') {
        // Handle DD/MM/YYYY or DD/MM/YYYY HH:MM:SS format (Australian date format)
        const ddmmyyyyMatch = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
        if (ddmmyyyyMatch) {
            const day = parseInt(ddmmyyyyMatch[1]);
            const month = parseInt(ddmmyyyyMatch[2]) - 1; // JS months are 0-indexed
            const year = parseInt(ddmmyyyyMatch[3]);
            const hours = ddmmyyyyMatch[4] ? parseInt(ddmmyyyyMatch[4]) : 0;
            const minutes = ddmmyyyyMatch[5] ? parseInt(ddmmyyyyMatch[5]) : 0;
            const seconds = ddmmyyyyMatch[6] ? parseInt(ddmmyyyyMatch[6]) : 0;
            return new Date(year, month, day, hours, minutes, seconds);
        }
        // Fallback to standard parsing
        return new Date(dateValue);
    }
    return null;
}

// Get day of week (0=Sunday, 1=Monday, etc.)
function getDayOfWeek(dateValue) {
    const date = getDateObj(dateValue);
    return date ? date.getDay() : null;
}

// Get hour of day
function getHour(dateValue) {
    const date = getDateObj(dateValue);
    return date ? date.getHours() : null;
}

// Get time slot (30-min intervals)
function getTimeSlot(dateValue) {
    const date = getDateObj(dateValue);
    if (!date) return null;
    const hours = date.getHours();
    const minutes = date.getMinutes();
    return hours * 2 + (minutes >= 30 ? 1 : 0);
}

// Format time slot to display string
function formatTimeSlot(slot) {
    const hours = Math.floor(slot / 2);
    const minutes = (slot % 2) * 30;
    const h = hours.toString().padStart(2, '0');
    const m = minutes.toString().padStart(2, '0');
    return `${h}:${m}`;
}

// Main processing function
function processAndDisplay() {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';

    // Filter data based on global location and view
    let filteredData = getGlobalFilteredData();

    // Update week info
    updateWeekInfo(filteredData);

    // Calculate and display metrics
    const incomingCalls = filteredData.filter(row => row.Direction === 'In');
    updateSummaryMetrics(incomingCalls);

    // Update abandonment analysis
    updateAbandonmentAnalysis(incomingCalls);

    // Update missed call follow-up analysis
    updateMissedCallFollowup(filteredData);

    // Update hourly chart
    updateHourlyChart(filteredData);

    // Update daily table
    updateDailyTable();

    // Update heatmaps
    updateHeatmaps(filteredData);

    // Update staff table
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

function updateSummaryMetrics(calls) {
    const total = calls.length;
    const answered = calls.filter(c => c.TimeToAnswer > 0).length;
    // Missed = calls that entered a queue but were not answered
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

    // Service Level: % of answered calls within target time
    const withinTarget = answeredCalls.filter(c => c.TimeToAnswer <= serviceLevelTarget).length;
    const serviceLevel = total > 0 ? ((withinTarget / total) * 100).toFixed(1) : 0;

    // Calculate FCR and Callback Rate
    const { fcrRate, callbackRate } = calculateCallbackMetrics(calls);

    document.getElementById('totalCalls').textContent = total;
    document.getElementById('answeredCalls').textContent = answered;
    document.getElementById('missedCalls').textContent = missed;
    document.getElementById('missedPercent').textContent = missedPct + '%';
    document.getElementById('serviceLevel').textContent = serviceLevel + '%';
    document.getElementById('fcrRate').textContent = fcrRate + '%';
    document.getElementById('callbackRate').textContent = callbackRate + '%';
    document.getElementById('avgWait').textContent = formatTime(avgWait);
    document.getElementById('maxWait').textContent = formatTime(maxWait);
    document.getElementById('avgCallLength').textContent = formatTime(avgCallLength);

    // Out of hours calls (uses rawData, not filtered)
    const outOfHours = countOutOfHoursCalls();
    document.getElementById('outOfHoursCalls').textContent = outOfHours;
}

function countOutOfHoursCalls() {
    // Apply global location filter but NOT opening hours filter
    let data = rawData;

    // Exclude internal calls
    data = data.filter(row => !isInternalCall(row));

    // Apply global location filter
    data = filterByLocation(data, currentGlobalLocation);

    // Apply queue filter if needed (same logic as getGlobalFilteredData but without opening hours)
    if (currentQueueFilter !== 'all') {
        if (currentQueueFilter === 'noqueue') {
            data = data.filter(row => !row.queueName);
        } else {
            const queueMap = {
                'appointments': 'Appointments',
                'vasectomy': 'Canberra Vasectomy',
                'general': 'General Enquiries',
                'health': 'Health Professionals'
            };
            const targetQueue = queueMap[currentQueueFilter];
            data = data.filter(row => row.queueName === targetQueue);
        }
    }

    // Count incoming calls OUTSIDE opening hours
    return data.filter(row =>
        row.Direction === 'In' &&
        !isWithinOpeningHours(row.CallDateTime)
    ).length;
}

function calculateCallbackMetrics(calls) {
    // Get all incoming calls with valid OriginNumber
    const validCalls = calls.filter(c =>
        c.OriginNumber &&
        c.OriginNumber !== '0' &&
        c.OriginNumber !== 0
    );

    if (validCalls.length === 0) {
        return { fcrRate: 0, callbackRate: 0 };
    }

    // Group ALL calls by OriginNumber with timestamps
    const callsByNumber = {};
    validCalls.forEach(call => {
        const num = String(call.OriginNumber);
        if (!callsByNumber[num]) {
            callsByNumber[num] = [];
        }
        const date = getDateObj(call.CallDateTime);
        if (date) {
            callsByNumber[num].push({
                call,
                date,
                isAnswered: call.TimeToAnswer > 0
            });
        }
    });

    // Count unique callers who called back within 24 hours
    let uniqueCallersWithCallback = 0;
    let totalUniqueCallers = Object.keys(callsByNumber).length;

    Object.values(callsByNumber).forEach(callList => {
        if (callList.length <= 1) return; // Only one call, no callback

        // Sort by date
        callList.sort((a, b) => a.date - b.date);

        // Check if any two consecutive calls are within the callback window
        let hasCallback = false;
        for (let i = 0; i < callList.length - 1; i++) {
            const timeDiff = callList[i + 1].date - callList[i].date;
            const hoursDiff = timeDiff / (1000 * 60 * 60);
            if (hoursDiff <= callbackWindowHours) {
                hasCallback = true;
                break;
            }
        }

        if (hasCallback) {
            uniqueCallersWithCallback++;
        }
    });

    const callbackRate = totalUniqueCallers > 0
        ? ((uniqueCallersWithCallback / totalUniqueCallers) * 100).toFixed(1)
        : 0;

    // FCR = unique callers without callback = 100% - callback rate
    const fcrRate = totalUniqueCallers > 0
        ? (100 - parseFloat(callbackRate)).toFixed(1)
        : 0;

    return { fcrRate, callbackRate };
}

function updateAbandonmentAnalysis(calls) {
    // Missed calls where TimeToAnswer is 0 or null
    // For missed calls, CallDuration represents how long they waited before hanging up
    const missedCalls = calls.filter(c => !c.TimeToAnswer || c.TimeToAnswer === 0);

    const totalAbandoned = missedCalls.length;
    document.getElementById('totalAbandoned').textContent = totalAbandoned;

    if (totalAbandoned === 0) {
        document.getElementById('avgAbandonWait').textContent = '-';
        document.getElementById('abandonmentGrid').innerHTML =
            '<p style="color: #95a5a6; font-size: 12px;">No abandoned calls in this period</p>';
        return;
    }

    // Calculate average wait time before abandonment
    const waitTimes = missedCalls.map(c => c.CallDuration || 0);
    const avgWait = waitTimes.reduce((a, b) => a + b, 0) / waitTimes.length;
    document.getElementById('avgAbandonWait').textContent = formatTime(avgWait);

    // Distribution buckets
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
    // Get all incoming calls with valid OriginNumber
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

    // Group calls by OriginNumber
    const callsByNumber = {};
    validCalls.forEach(call => {
        const num = String(call.OriginNumber);
        if (!callsByNumber[num]) {
            callsByNumber[num] = [];
        }
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

    // Analyze each unique caller
    let lostCount = 0;
    let persistentCount = 0;
    const lostByHour = {};
    const attemptCounts = { 1: 0, 2: 0, '3+': 0 };
    let totalMissedWaitTime = 0;
    let missedWaitCount = 0;
    const persistentAttempts = [];

    Object.values(callsByNumber).forEach(callList => {
        // Sort by date
        callList.sort((a, b) => a.date - b.date);

        // Find first missed call
        const firstMissedIdx = callList.findIndex(c => c.isMissed);
        if (firstMissedIdx === -1) return; // No missed calls for this number

        const firstMissed = callList[firstMissedIdx];

        // Get all calls within the callback window of the first missed call
        const within24h = callList.filter(c => {
            const hoursDiff = (c.date - firstMissed.date) / (1000 * 60 * 60);
            return hoursDiff >= 0 && hoursDiff <= callbackWindowHours;
        });

        // Track wait times for missed calls
        within24h.filter(c => c.isMissed).forEach(c => {
            totalMissedWaitTime += c.waitTime;
            missedWaitCount++;
        });

        // Check if they ever got answered within 24h
        const gotAnswered = within24h.some(c => c.isAnswered);

        if (within24h.length === 1 && !gotAnswered) {
            // Only one missed call, never called back = lost opportunity
            lostCount++;
            const hour = firstMissed.hour;
            lostByHour[hour] = (lostByHour[hour] || 0) + 1;
        } else if (!gotAnswered && within24h.length > 1) {
            // Called back but never got answered = still lost, but persistent
            lostCount++;
            const hour = firstMissed.hour;
            lostByHour[hour] = (lostByHour[hour] || 0) + 1;
        } else if (gotAnswered) {
            // Eventually got answered = persistent caller
            persistentCount++;

            // Count attempts (missed calls before being answered)
            const answeredIdx = within24h.findIndex(c => c.isAnswered);
            const attempts = within24h.slice(0, answeredIdx + 1).filter(c => c.isMissed || c.isAnswered).length;

            persistentAttempts.push(attempts);

            if (attempts === 1) attemptCounts[1]++;
            else if (attempts === 2) attemptCounts[2]++;
            else attemptCounts['3+']++;
        }
    });

    // Update summary cards
    document.getElementById('lostOpportunities').textContent = lostCount;
    document.getElementById('persistentCallers').textContent = persistentCount;

    // Find peak hour for lost opportunities
    const peakHour = Object.entries(lostByHour).sort((a, b) => b[1] - a[1])[0];
    if (peakHour) {
        const hour = parseInt(peakHour[0]);
        const hourStr = hour > 12 ? `${hour - 12}-${hour - 11}pm` : (hour === 12 ? '12-1pm' : `${hour}-${hour + 1}am`);
        document.getElementById('lostPeakHour').textContent = `Peak: ${hourStr} (${peakHour[1]})`;
    } else {
        document.getElementById('lostPeakHour').textContent = '';
    }

    // Average attempts for persistent callers
    if (persistentAttempts.length > 0) {
        const avgAttempts = (persistentAttempts.reduce((a, b) => a + b, 0) / persistentAttempts.length).toFixed(1);
        document.getElementById('avgAttempts').textContent = `Avg ${avgAttempts} attempts`;
    } else {
        document.getElementById('avgAttempts').textContent = '';
    }

    // Update details section - attempts distribution
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

    // Update details section - peak hours
    const sortedHours = Object.entries(lostByHour)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    let peakHoursHtml = '';
    sortedHours.forEach(([hour, count]) => {
        const h = parseInt(hour);
        const hourStr = h > 12 ? `${h - 12}-${h - 11}pm` : (h === 12 ? '12-1pm' : `${h}-${h + 1}am`);
        peakHoursHtml += `<span class="peak-hour-tag">${hourStr}: ${count} lost</span>`;
    });
    document.getElementById('lostPeakHours').innerHTML = peakHoursHtml || '<p style="color: #95a5a6; font-size: 12px;">No lost opportunities</p>';

    // Average wait before hanging up
    const avgWait = missedWaitCount > 0 ? totalMissedWaitTime / missedWaitCount : 0;
    document.getElementById('avgWaitBeforeHangup').textContent = formatTime(avgWait);
}

function updateHourlyChart(data) {
    const inCalls = data.filter(row => row.Direction === 'In');
    const outCalls = data.filter(row => row.Direction === 'Out');

    // Hours from 7 to 18 (7 AM to 6 PM)
    const hours = [];
    for (let h = 7; h <= 18; h++) {
        hours.push(h);
    }

    // Count calls per hour
    const inCountsPerHour = hours.map(h =>
        inCalls.filter(c => getHour(c.CallDateTime) === h).length
    );
    const outCountsPerHour = hours.map(h =>
        outCalls.filter(c => getHour(c.CallDateTime) === h).length
    );

    const labels = hours.map(h => {
        const suffix = h >= 12 ? 'PM' : 'AM';
        const hour12 = h > 12 ? h - 12 : (h === 0 ? 12 : h);
        return `${hour12}${suffix}`;
    });

    // Destroy existing chart if any
    if (hourlyChart) {
        hourlyChart.destroy();
    }

    const ctx = document.getElementById('hourlyChart').getContext('2d');
    hourlyChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Calls In',
                    data: inCountsPerHour,
                    borderColor: '#27ae60',
                    backgroundColor: 'rgba(39, 174, 96, 0.1)',
                    fill: true,
                    tension: 0.3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                },
                {
                    label: 'Calls Out',
                    data: outCountsPerHour,
                    borderColor: '#3498db',
                    backgroundColor: 'rgba(52, 152, 219, 0.1)',
                    fill: true,
                    tension: 0.3,
                    pointRadius: 4,
                    pointHoverRadius: 6
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Number of Calls'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Hour of Day'
                    }
                }
            }
        }
    });
}

function updateDailyTable() {
    let filteredData = getGlobalFilteredData();

    // Further filter by direction
    if (currentDailyDirection === 'in') {
        filteredData = filteredData.filter(row => row.Direction === 'In');
    } else if (currentDailyDirection === 'out') {
        filteredData = filteredData.filter(row => row.Direction === 'Out');
    }

    // Group by day of week
    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dayData = {};
    dayNames.forEach((name, idx) => {
        dayData[idx] = [];
    });

    filteredData.forEach(row => {
        const day = getDayOfWeek(row.CallDateTime);
        if (day !== null && dayData[day]) {
            dayData[day].push(row);
        }
    });

    // Calculate metrics for each day
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
    // Reorder: Mon(1), Tue(2), Wed(3), Thu(4), Fri(5), Sat(6), Sun(0)
    const displayOrder = [1, 2, 3, 4, 5, 6, 0];
    displayOrder.forEach(day => {
        metrics[day] = calcMetrics(dayData[day]);
    });
    const weekMetrics = calcMetrics(filteredData);

    // Build table HTML
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

function updateHeatmaps(data) {
    const inCalls = data.filter(row => row.Direction === 'In');
    const outCalls = data.filter(row => row.Direction === 'Out');

    renderHeatmap('heatmapIn', inCalls);
    renderHeatmap('heatmapOut', outCalls);

    // Render wait time heatmaps (incoming calls only)
    renderWaitTimeHeatmap('heatmapMaxWait', inCalls, 'max');
    renderWaitTimeHeatmap('heatmapAvgWait', inCalls, 'avg');
}

function renderHeatmap(elementId, calls) {
    // Time slots from 7:30 AM (slot 15) to 5:30 PM (slot 35)
    const startSlot = 15; // 7:30
    const endSlot = 35;   // 17:30
    const days = [1, 2, 3, 4, 5, 6]; // Mon-Sat
    const dayNames = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    // Count calls per slot per day
    const counts = {};
    let maxCount = 0;

    for (let slot = startSlot; slot <= endSlot; slot++) {
        counts[slot] = {};
        days.forEach(day => {
            counts[slot][day] = 0;
        });
    }

    calls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            counts[slot][day]++;
            maxCount = Math.max(maxCount, counts[slot][day]);
        }
    });

    // Build heatmap HTML
    let html = '';

    // Header row
    html += '<div class="heatmap-cell heatmap-header"></div>';
    days.forEach(day => {
        html += `<div class="heatmap-cell heatmap-header">${dayNames[day]}</div>`;
    });

    // Data rows
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
    // Time slots from 7:30 AM (slot 15) to 5:30 PM (slot 35)
    const startSlot = 15; // 7:30
    const endSlot = 35;   // 17:30
    const days = [1, 2, 3, 4, 5, 6]; // Mon-Sat
    const dayNames = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    // Only include answered calls
    const answeredCalls = calls.filter(c => c.TimeToAnswer > 0);

    // Collect wait times per slot per day
    const waitTimes = {};
    for (let slot = startSlot; slot <= endSlot; slot++) {
        waitTimes[slot] = {};
        days.forEach(day => {
            waitTimes[slot][day] = [];
        });
    }

    answeredCalls.forEach(call => {
        const day = getDayOfWeek(call.CallDateTime);
        const slot = getTimeSlot(call.CallDateTime);
        if (day !== null && slot !== null && days.includes(day) && slot >= startSlot && slot <= endSlot) {
            waitTimes[slot][day].push(call.TimeToAnswer);
        }
    });

    // Calculate max or avg for each cell
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

    // Determine heat level based on wait time thresholds (in seconds)
    // 0: no data, 1: <30s, 2: 30-60s, 3: 60-90s, 4: 90-120s, 5: 120-180s, 6: 180-300s, 7: >300s
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

    // Build heatmap HTML
    let html = '';

    // Header row
    html += '<div class="heatmap-cell heatmap-header"></div>';
    days.forEach(day => {
        html += `<div class="heatmap-cell heatmap-header">${dayNames[day]}</div>`;
    });

    // Data rows
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

function updateStaffTable(data) {
    // Apply location filter
    const locationFilteredData = filterByLocation(data, currentLocationFilter);

    // Group by staff member
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

    // Convert to array and sort by total calls
    const staffArray = Object.values(staffStats)
        .map(s => ({
            ...s,
            totalCalls: s.callsIn + s.callsOut,
            avgPickup: s.pickupCount > 0 ? s.totalPickupTime / s.pickupCount : null,
            avgCallLengthIn: s.callLengthCountIn > 0 ? s.totalCallLengthIn / s.callLengthCountIn : null,
            avgCallLengthOut: s.callLengthCountOut > 0 ? s.totalCallLengthOut / s.callLengthCountOut : null
        }))
        .sort((a, b) => b.totalCalls - a.totalCalls);

    // Build table HTML
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
