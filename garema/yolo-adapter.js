// YOLO PBX export adapter for the Garema phone dashboard.
//
// The dashboard logic (copied from the YourGP dashboard) expects one row per
// call in the shape of the old PBX's "ExportCallRecords" file:
//
//   CallGUID       unique call id
//   CallDateTime   Date the call started
//   Direction      'In' | 'Out' | 'Int'
//   OriginNumber   caller's number
//   TargetNumber   number dialled
//   UserName       staff member who answered / made the call
//   TimeToAnswer   seconds until answered (> 0 means answered, 0 means not answered)
//   CallDuration   seconds; for answered calls = TimeToAnswer + talk time,
//                  for unanswered calls = how long the caller waited before hanging up
//   queueName      queue name if the call reached the queue, otherwise null
//
// YOLO's export format was not known when this was written, so columns are
// matched by name against a list of likely aliases (see FIELD_ALIASES). When
// the first real export arrives, add its exact column names to the alias
// lists if anything is reported as "not found" in the Data Check panel.
//
// Works in the browser (window.YoloAdapter) and in Node (module.exports) so it
// can be tested from the command line.

(function (root) {
    'use strict';

    // Name used for the single Garema queue throughout the dashboard
    const QUEUE_NAME = 'Garema';

    // Aliases are compared after lower-casing and removing everything that is
    // not a letter or digit, so "Call Start Time", "call_start_time" and
    // "CallStartTime" all become "callstarttime".
    const FIELD_ALIASES = {
        id: ['callid', 'callguid', 'guid', 'uniqueid', 'linkedid', 'callreference', 'callref', 'reference', 'sessionid', 'calluuid', 'uuid', 'id'],
        start: ['callstart', 'callstarttime', 'callstartdatetime', 'starttime', 'startdatetime', 'calldatetime', 'datetime', 'startedat', 'started', 'timestamp', 'start', 'calltimestamp'],
        dateOnly: ['date', 'calldate', 'startdate'],
        timeOnly: ['time', 'calltime'],
        answer: ['answertime', 'answeredtime', 'answeredat', 'timeanswered', 'answerdatetime', 'connecttime', 'connectedat', 'connectedtime', 'pickuptime'],
        end: ['endtime', 'callend', 'callendtime', 'enddatetime', 'endedat', 'ended', 'hanguptime', 'finishtime', 'end'],
        wait: ['waittime', 'waitduration', 'waitseconds', 'waitsecs', 'queuetime', 'queueduration', 'queuewaittime', 'timeinqueue', 'queuewait', 'timetoanswer', 'calltimetoanswer', 'ringtime', 'ringduration', 'holdtimebeforeanswer'],
        talk: ['talktime', 'talkduration', 'talkseconds', 'conversationtime', 'conversationduration', 'connectedduration', 'billsec', 'billabletime', 'billableseconds', 'totaltalktimeallusers'],
        duration: ['duration', 'callduration', 'totalduration', 'totaltime', 'calllength', 'length', 'durationseconds', 'durationsecs', 'totalcallduration'],
        direction: ['direction', 'calldirection', 'inout', 'calltype', 'type'],
        from: ['from', 'fromnumber', 'caller', 'callerid', 'callernumber', 'callerphone', 'cli', 'callingnumber', 'callingparty', 'originnumber', 'origin', 'source', 'sourcenumber', 'src', 'ani'],
        to: ['to', 'tonumber', 'dialled', 'diallednumber', 'dialednumber', 'callednumber', 'calledparty', 'targetnumber', 'destination', 'destinationnumber', 'dst', 'did', 'dnis'],
        agent: ['agent', 'agentname', 'answeredby', 'answeredbyname', 'username', 'user', 'staff', 'staffmember', 'staffname', 'handledby', 'extensionname', 'firstusername', 'member', 'operator'],
        extension: ['extension', 'ext', 'agentextension', 'userextension', 'extensionnumber'],
        outcome: ['outcome', 'calloutcome', 'status', 'callstatus', 'result', 'callresult', 'disposition', 'hangupreason', 'endreason', 'reason'],
        queue: ['queue', 'queuename', 'callqueue', 'callqueuename', 'huntgroup', 'ringgroup']
    };

    // Fields the dashboard really needs, with a plain-English description for
    // the Data Check panel. A field is satisfied if any of its "anyOf" groups
    // were found.
    const REQUIREMENTS = [
        { label: 'Call start date and time', anyOf: [['start'], ['dateOnly', 'timeOnly']] },
        { label: 'Wait time before answer', anyOf: [['wait'], ['answer', 'start']] },
        { label: 'Call end time or call length', anyOf: [['end'], ['duration'], ['talk']] },
        { label: 'Direction (in / out / internal)', anyOf: [['direction']] },
        { label: 'Caller number', anyOf: [['from']] },
        { label: 'Staff member who answered', anyOf: [['agent'], ['extension']] },
        { label: 'Call outcome (answered / missed)', anyOf: [['outcome'], ['answer'], ['wait'], ['agent']] },
        { label: 'Unique call ID', anyOf: [['id']] }
    ];

    function normaliseHeader(h) {
        return String(h || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    // Work out which column in the file holds each field.
    function detectColumns(headers) {
        const norm = headers.map(normaliseHeader);
        const claimed = new Set();
        const map = {};

        // Pass 1: exact alias matches, in alias priority order
        for (const field of Object.keys(FIELD_ALIASES)) {
            for (const alias of FIELD_ALIASES[field]) {
                const idx = norm.findIndex((n, i) => n === alias && !claimed.has(i));
                if (idx !== -1) {
                    map[field] = headers[idx];
                    claimed.add(idx);
                    break;
                }
            }
        }

        // Pass 2: a header that contains a longer alias (e.g. "Caller Number (CLI)")
        for (const field of Object.keys(FIELD_ALIASES)) {
            if (map[field]) continue;
            for (const alias of FIELD_ALIASES[field]) {
                if (alias.length < 6) continue;
                const idx = norm.findIndex((n, i) => !claimed.has(i) && n.includes(alias));
                if (idx !== -1) {
                    map[field] = headers[idx];
                    claimed.add(idx);
                    break;
                }
            }
        }

        const unused = headers.filter((h, i) => !claimed.has(i));
        return { map, unused };
    }

    function isBlank(v) {
        return v === undefined || v === null || String(v).trim() === '' || String(v).trim() === '-';
    }

    // Convert an Excel serial date number to a local Date
    function excelSerialToDate(serial) {
        const epoch = new Date(1899, 11, 30);
        const days = Math.floor(serial);
        const totalSeconds = Math.round((serial - days) * 86400);
        const d = new Date(epoch.getTime() + days * 86400000);
        d.setHours(Math.floor(totalSeconds / 3600), Math.floor((totalSeconds % 3600) / 60), totalSeconds % 60, 0);
        return d;
    }

    // Parse a date/time value. Ambiguous "a/b/yyyy" dates are read as
    // DD/MM/YYYY (Australian) unless monthFirst is true.
    function parseDateTime(v, monthFirst) {
        if (isBlank(v)) return null;
        if (v instanceof Date) return isNaN(v) ? null : v;
        if (typeof v === 'number') return v > 1000 ? excelSerialToDate(v) : null;

        const s = String(v).trim();

        // DD/MM/YYYY or DD-MM-YYYY, optional time, optional AM/PM
        let m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:[ T,]+(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(AM|PM)?)?$/i);
        if (m) {
            let day = parseInt(m[1], 10);
            let month = parseInt(m[2], 10);
            if (monthFirst) [day, month] = [month, day];
            let year = parseInt(m[3], 10);
            if (year < 100) year += 2000;
            let hours = m[4] ? parseInt(m[4], 10) : 0;
            const ampm = m[7] ? m[7].toUpperCase() : null;
            if (ampm === 'PM' && hours !== 12) hours += 12;
            if (ampm === 'AM' && hours === 12) hours = 0;
            return new Date(year, month - 1, day, hours, m[5] ? parseInt(m[5], 10) : 0, m[6] ? parseInt(m[6], 10) : 0);
        }

        // YYYY-MM-DD with optional time. Without a timezone suffix it is
        // treated as local (Canberra) time; with Z or an offset the browser
        // converts it to local time.
        m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?)?\s*(Z|[+-]\d{2}:?\d{2})?$/i);
        if (m) {
            if (m[7]) {
                const d = new Date(s.replace(' ', 'T'));
                return isNaN(d) ? null : d;
            }
            return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10),
                m[4] ? parseInt(m[4], 10) : 0, m[5] ? parseInt(m[5], 10) : 0, m[6] ? parseInt(m[6], 10) : 0);
        }

        // Numeric string that is an Excel serial
        if (/^\d+(\.\d+)?$/.test(s) && parseFloat(s) > 20000) return excelSerialToDate(parseFloat(s));

        const d = new Date(s);
        return isNaN(d) ? null : d;
    }

    // Parse a time of day ("9:05:30 AM", "13:20", Excel fraction) into seconds after midnight
    function parseTimeOfDay(v) {
        if (isBlank(v)) return null;
        if (typeof v === 'number') return v < 1 ? Math.round(v * 86400) : null;
        const m = String(v).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(AM|PM)?$/i);
        if (!m) return null;
        let h = parseInt(m[1], 10);
        const ampm = m[4] ? m[4].toUpperCase() : null;
        if (ampm === 'PM' && h !== 12) h += 12;
        if (ampm === 'AM' && h === 12) h = 0;
        return h * 3600 + parseInt(m[2], 10) * 60 + (m[3] ? parseInt(m[3], 10) : 0);
    }

    // Parse a duration into seconds: 83, "83", "83s", "1:23", "00:01:23", "1m 23s", or an Excel day fraction
    function parseSeconds(v) {
        if (isBlank(v)) return null;
        if (typeof v === 'number') {
            if (v > 0 && v < 1 && !Number.isInteger(v)) return Math.round(v * 86400);
            return v;
        }
        const s = String(v).trim().toLowerCase();
        if (/^-?\d+(\.\d+)?s?$/.test(s)) return parseFloat(s);
        let m = s.match(/^(\d+):(\d{1,2}):(\d{1,2})(?:\.\d+)?$/);
        if (m) return parseInt(m[1], 10) * 3600 + parseInt(m[2], 10) * 60 + parseInt(m[3], 10);
        m = s.match(/^(\d+):(\d{1,2})(?:\.\d+)?$/);
        if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
        m = s.match(/^(?:(\d+)\s*h(?:ours?|rs?)?)?\s*(?:(\d+)\s*m(?:in(?:ute)?s?)?)?\s*(?:(\d+)\s*s(?:ec(?:ond)?s?)?)?$/);
        if (m && (m[1] || m[2] || m[3])) {
            return (parseInt(m[1] || 0, 10) * 3600) + (parseInt(m[2] || 0, 10) * 60) + parseInt(m[3] || 0, 10);
        }
        return null;
    }

    function normaliseDirection(v) {
        if (isBlank(v)) return null;
        const s = String(v).trim().toLowerCase();
        if (/^(in|inbound|incoming|received|inward)/.test(s)) return 'In';
        if (/^(out|outbound|outgoing|dialled|dialed|placed)/.test(s)) return 'Out';
        if (/^(int|internal|local|extension|ext|intercom)/.test(s)) return 'Int';
        return null;
    }

    // Map a YOLO outcome/status string onto: answered, missed, menu, or null (unknown)
    function normaliseOutcome(v) {
        if (isBlank(v)) return null;
        const s = String(v).trim().toLowerCase();
        if (/(ivr|menu|auto ?attendant|before queue|not queued|no queue)/.test(s)) return 'menu';
        if (/(unanswered|not answered|no answer|noanswer|missed|abandon|timeout|timed out|overflow|voicemail|voice mail|no agent|busy|failed|declined|rejected|cancel)/.test(s)) return 'missed';
        if (/(answered|connected|completed|normal_clearing|normal clearing|success|handled|talking|accepted)/.test(s)) return 'answered';
        return null;
    }

    function cleanNumber(v) {
        if (isBlank(v)) return '';
        return String(v).trim();
    }

    // Decide Direction when the export has no direction column. Extensions are
    // short (3 or 4 digits); external numbers are long.
    function guessDirection(from, to) {
        const isExt = n => /^\d{2,5}$/.test(String(n || '').replace(/\D/g, '')) && String(n).replace(/\D/g, '').length <= 5;
        if (isExt(from) && isExt(to)) return 'Int';
        if (isExt(from)) return 'Out';
        return 'In';
    }

    // True if any a/b/yyyy date in the column has a second part greater than 12,
    // which means the file is month-first (US style).
    function detectMonthFirst(rows, cols) {
        let firstOver12 = false, secondOver12 = false;
        for (const row of rows) {
            for (const f of ['start', 'dateOnly', 'answer', 'end']) {
                const col = cols[f];
                if (!col) continue;
                const m = String(row[col] ?? '').match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-]\d{2,4}/);
                if (!m) continue;
                if (parseInt(m[1], 10) > 12) firstOver12 = true;
                if (parseInt(m[2], 10) > 12) secondOver12 = true;
            }
        }
        return secondOver12 && !firstOver12;
    }

    // Convert one raw YOLO row into an intermediate record
    function convertRow(row, cols, monthFirst) {
        const get = f => (cols[f] ? row[cols[f]] : undefined);

        let start = parseDateTime(get('start'), monthFirst);
        if (!start && cols.dateOnly) {
            const d = parseDateTime(get('dateOnly'), monthFirst);
            const t = parseTimeOfDay(get('timeOnly'));
            if (d) {
                start = new Date(d);
                if (t !== null) start.setHours(Math.floor(t / 3600), Math.floor((t % 3600) / 60), t % 60, 0);
            }
        }
        // Time-only answer/end columns are combined with the start date
        const withStartDate = v => {
            const full = parseDateTime(v, monthFirst);
            if (full && !(typeof v === 'string' && /^\d{1,2}:\d{2}/.test(v.trim()))) return full;
            const t = parseTimeOfDay(v);
            if (t === null || !start) return full;
            const d = new Date(start);
            d.setHours(Math.floor(t / 3600), Math.floor((t % 3600) / 60), t % 60, 0);
            if (d < start) d.setDate(d.getDate() + 1);
            return d;
        };
        const answer = withStartDate(get('answer'));
        const end = withStartDate(get('end'));

        const from = cleanNumber(get('from'));
        const to = cleanNumber(get('to'));
        // Fall back to guessing from number lengths when there is no direction
        // column or its value is not recognised
        const direction = normaliseDirection(get('direction')) || guessDirection(from, to);

        return {
            id: isBlank(get('id')) ? null : String(get('id')).trim(),
            start,
            answer,
            end,
            wait: parseSeconds(get('wait')),
            talk: parseSeconds(get('talk')),
            duration: parseSeconds(get('duration')),
            direction,
            from,
            to,
            agent: isBlank(get('agent')) ? '' : String(get('agent')).trim(),
            extension: isBlank(get('extension')) ? '' : String(get('extension')).trim(),
            outcomeRaw: isBlank(get('outcome')) ? '' : String(get('outcome')).trim(),
            outcome: normaliseOutcome(get('outcome')),
            queue: isBlank(get('queue')) ? '' : String(get('queue')).trim(),
            raw: row
        };
    }

    // Merge several legs of the same call (transfers, multiple queue attempts)
    function mergeLegs(legs) {
        if (legs.length === 1) return legs[0];
        const byStart = legs.slice().sort((a, b) => (a.start || 0) - (b.start || 0));
        const first = byStart[0];
        const answeredLeg = byStart.find(l => l.outcome === 'answered' || l.answer || (l.agent && (l.talk || 0) > 0));
        const ends = legs.map(l => l.end).filter(Boolean);
        const merged = Object.assign({}, first);
        merged.end = ends.length ? new Date(Math.max(...ends)) : first.end;
        merged.queue = legs.map(l => l.queue).find(Boolean) || '';
        merged.talk = legs.reduce((s, l) => s + (l.talk || 0), 0) || first.talk;
        if (answeredLeg) {
            merged.answer = answeredLeg.answer || first.answer;
            merged.agent = answeredLeg.agent || first.agent;
            merged.extension = answeredLeg.extension || first.extension;
            merged.outcome = 'answered';
            if (answeredLeg.wait !== null && answeredLeg !== first) merged.wait = (first.wait || 0) + answeredLeg.wait;
        } else if (legs.some(l => l.outcome === 'missed')) {
            merged.outcome = 'missed';
        }
        merged.legs = legs.length;
        return merged;
    }

    // Turn an intermediate record into the row shape the dashboard expects
    function toDashboardRow(r, cols, index) {
        const hasOutcome = !!cols.outcome && r.outcome !== null;
        const startMs = r.start ? r.start.getTime() : null;
        const secsBetween = (a, b) => (a && b ? Math.max(0, (b.getTime() - a.getTime()) / 1000) : null);

        let answered;
        if (hasOutcome) answered = r.outcome === 'answered';
        else answered = !!(r.answer || (r.agent && ((r.talk || 0) > 0 || (r.wait !== null && r.duration > r.wait))));
        if (r.direction === 'Out') answered = (r.talk || 0) > 0 || !!r.answer || (hasOutcome ? r.outcome === 'answered' : (r.duration || 0) > 0);

        // Seconds from start to answer (or to hang-up if never answered)
        let wait = r.wait;
        if (wait === null) wait = answered ? secsBetween(r.start, r.answer) : secsBetween(r.start, r.end);

        // Talk time
        let talk = r.talk;
        if (talk === null && answered) talk = secsBetween(r.answer, r.end);
        if (talk === null && answered && r.duration !== null && wait !== null) talk = Math.max(0, r.duration - wait);

        // Total duration
        let duration = r.duration;
        if (duration === null) duration = secsBetween(r.start, r.end);
        if (duration === null) duration = (wait || 0) + (talk || 0);

        let timeToAnswer = 0;
        let callDuration;
        if (r.direction === 'In') {
            if (answered) {
                // 0.01 keeps instantly-answered calls counted as answered (dashboard uses TimeToAnswer > 0)
                timeToAnswer = Math.max(wait || 0, 0.01);
                callDuration = timeToAnswer + (talk !== null ? talk : Math.max(0, duration - timeToAnswer));
            } else {
                // For unanswered calls the dashboard reads CallDuration as "waited before hanging up"
                callDuration = wait !== null ? wait : duration;
            }
        } else {
            timeToAnswer = 0;
            callDuration = talk !== null && talk > 0 ? talk : duration;
        }

        // Did the inbound call reach the Garema queue?
        let queueName = null;
        if (r.direction === 'In') {
            if (r.outcome === 'menu') queueName = null;
            else if (cols.queue) queueName = r.queue ? QUEUE_NAME : (answered ? QUEUE_NAME : null);
            else queueName = QUEUE_NAME;
        }

        const userName = r.agent || r.extension || '';

        return {
            CallGUID: r.id || ('row-' + index),
            CallDateTime: r.start,
            Direction: r.direction,
            OriginNumber: r.from,
            TargetNumber: r.to,
            UserName: userName,
            UserExtension: r.extension,
            OfficeName: 'Garema',
            TimeToAnswer: timeToAnswer,
            CallDuration: callDuration || 0,
            TalkTime: talk || 0,
            BillableTime: talk || 0,
            queueName,
            YoloQueue: r.queue,
            YoloOutcome: r.outcomeRaw,
            Legs: r.legs || 1,
            startMs
        };
    }

    // Main entry point. rows = array of objects (header -> value) from one or
    // more export files. Returns { calls, report }.
    function normaliseRows(rows) {
        const report = {
            totalRows: rows.length,
            columns: {},
            unusedColumns: [],
            requirements: [],
            warnings: [],
            skippedNoDate: 0,
            mergedLegs: 0,
            monthFirst: false,
            outcomeValues: {}
        };
        if (!rows.length) return { calls: [], report };

        const headerSet = new Set();
        rows.forEach(r => Object.keys(r).forEach(k => headerSet.add(k)));
        const headers = Array.from(headerSet);
        const { map: cols, unused } = detectColumns(headers);
        report.columns = cols;
        report.unusedColumns = unused;
        report.monthFirst = detectMonthFirst(rows.slice(0, 2000), cols);

        report.requirements = REQUIREMENTS.map(req => {
            const met = req.anyOf.find(group => group.every(f => cols[f]));
            return { label: req.label, found: !!met, columns: met ? met.map(f => cols[f]) : [] };
        });

        const converted = [];
        rows.forEach(row => {
            const r = convertRow(row, cols, report.monthFirst);
            if (!r.start) { report.skippedNoDate++; return; }
            if (r.outcomeRaw) report.outcomeValues[r.outcomeRaw] = (report.outcomeValues[r.outcomeRaw] || 0) + 1;
            converted.push(r);
        });

        // Group legs that share a call id
        let records = converted;
        if (cols.id) {
            const groups = new Map();
            const loose = [];
            converted.forEach(r => {
                if (!r.id) { loose.push(r); return; }
                if (!groups.has(r.id)) groups.set(r.id, []);
                groups.get(r.id).push(r);
            });
            records = loose.slice();
            groups.forEach(legs => {
                if (legs.length > 1) report.mergedLegs += legs.length - 1;
                records.push(mergeLegs(legs));
            });
        }

        const calls = records
            .map((r, i) => toDashboardRow(r, cols, i))
            .sort((a, b) => a.startMs - b.startMs);

        // Unrecognised outcome strings
        if (cols.outcome) {
            const unknown = Object.keys(report.outcomeValues).filter(v => normaliseOutcome(v) === null);
            if (unknown.length) {
                report.warnings.push(`Outcome values not recognised (treated as unknown, answered/missed worked out from other columns): ${unknown.slice(0, 10).join(', ')}`);
            }
        }
        if (report.skippedNoDate) report.warnings.push(`${report.skippedNoDate} rows skipped because no call start date/time could be read.`);
        if (report.monthFirst) report.warnings.push('Dates appear to be month-first (MM/DD/YYYY) and were read that way.');
        if (!cols.direction) report.warnings.push('No direction column found; direction was guessed from number lengths (short numbers are treated as extensions).');

        return { calls, report };
    }

    const api = {
        QUEUE_NAME,
        FIELD_ALIASES,
        detectColumns,
        parseDateTime,
        parseSeconds,
        normaliseOutcome,
        normaliseDirection,
        normaliseRows
    };

    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else root.YoloAdapter = api;
})(typeof window !== 'undefined' ? window : this);
