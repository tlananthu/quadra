import { state } from './state.js';

export function roundToQuarterHour(hoursVal) {
    return Math.round(hoursVal * 4) / 4;
}

export function showToast(message) {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerText = message;
    container.appendChild(toast);
    
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

export function getTzOffset(tzString) {
    if(!tzString || tzString === 'local') return -new Date().getTimezoneOffset() / 60;
    const d = new Date();
    const utcStr = d.toLocaleString('en-US', { timeZone: 'UTC' });
    const tzStr = d.toLocaleString('en-US', { timeZone: tzString });
    return (new Date(tzStr) - new Date(utcStr)) / 3600000;
}

export function getTzTime(date, tzString) {
    if (!tzString || tzString === 'local') return { h: date.getHours(), m: date.getMinutes(), s: date.getSeconds() };
    const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: tzString,
        hour: 'numeric', minute: 'numeric', second: 'numeric',
        hour12: false
    }).formatToParts(date);
    
    let h = 0, m = 0, s = 0;
    for (let p of parts) {
        if (p.type === 'hour') h = parseInt(p.value, 10);
        if (p.type === 'minute') m = parseInt(p.value, 10);
        if (p.type === 'second') s = parseInt(p.value, 10);
    }
    if (h === 24) h = 0;
    return { h, m, s };
}

export function formatCurrentTimeBadge(date) {
    let primaryOpts = { hour: '2-digit', minute:'2-digit' };
    if (state.appConfig.primaryTz !== 'local') primaryOpts.timeZone = state.appConfig.primaryTz;
    let text = date.toLocaleTimeString([], primaryOpts);

    if (state.appConfig.secondaryTz && state.appConfig.secondaryTz !== 'none') {
        let secOpts = { hour: '2-digit', minute:'2-digit', timeZoneName: 'short' };
        secOpts.timeZone = state.appConfig.secondaryTz;
        let secText = date.toLocaleTimeString([], secOpts);
        text += ` (${secText})`;
    }
    return text;
}

export function decToTime(dec) {
    let roundedDec = roundToQuarterHour(dec);
    let normalizedDec = roundedDec % 24;
    let h = Math.floor(normalizedDec);
    let m = Math.round((normalizedDec - h) * 60);
    if (m === 60) { h++; m = 0; }
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

export function addDays(dateStr, days) {
    const [y, m, d] = dateStr.split('-').map(Number);
    const date = new Date(y, m - 1, d + days);
    const localY = date.getFullYear();
    const localM = String(date.getMonth() + 1).padStart(2, '0');
    const localD = String(date.getDate()).padStart(2, '0');
    return `${localY}-${localM}-${localD}`;
}

export function timeStrToDecimal(timeStr) {
    if (!timeStr) return null;
    const [h, m] = timeStr.split(':').map(Number);
    return h + (m / 60);
}

export function decimalToTimeStr(dec) {
    let norm = dec % 24;
    let h = Math.floor(norm);
    let m = Math.round((norm - h) * 60);
    if (m === 60) { h++; m = 0; }
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

export function escapeHTML(str) { 
    return str.replace(/[&<>'"]/g, tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag)); 
}

import { state } from './state.js';

export function roundToQuarterHour(hoursVal) {
    return Math.round(hoursVal * 4) / 4;
}

export function showToast(message) {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerText = message;
    container.appendChild(toast);
    
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

export function getTzOffset(tzString) {
    if(!tzString || tzString === 'local') return -new Date().getTimezoneOffset() / 60;
    const d = new Date();
    const utcStr = d.toLocaleString('en-US', { timeZone: 'UTC' });
    const tzStr = d.toLocaleString('en-US', { timeZone: tzString });
    return (new Date(tzStr) - new Date(utcStr)) / 3600000;
}

export function getTzTime(date, tzString) {
    if (!tzString || tzString === 'local') return { h: date.getHours(), m: date.getMinutes(), s: date.getSeconds() };
    const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: tzString,
        hour: 'numeric', minute: 'numeric', second: 'numeric',
        hour12: false
    }).formatToParts(date);
    
    let h = 0, m = 0, s = 0;
    for (let p of parts) {
        if (p.type === 'hour') h = parseInt(p.value, 10);
        if (p.type === 'minute') m = parseInt(p.value, 10);
        if (p.type === 'second') s = parseInt(p.value, 10);
    }
    if (h === 24) h = 0;
    return { h, m, s };
}

export function formatCurrentTimeBadge(date) {
    let primaryOpts = { hour: '2-digit', minute:'2-digit' };
    if (state.appConfig.primaryTz !== 'local') primaryOpts.timeZone = state.appConfig.primaryTz;
    let text = date.toLocaleTimeString([], primaryOpts);

    if (state.appConfig.secondaryTz && state.appConfig.secondaryTz !== 'none') {
        let secOpts = { hour: '2-digit', minute:'2-digit', timeZoneName: 'short' };
        secOpts.timeZone = state.appConfig.secondaryTz;
        let secText = date.toLocaleTimeString([], secOpts);
        text += ` (${secText})`;
    }
    return text;
}

export function decToTime(dec) {
    let roundedDec = roundToQuarterHour(dec);
    let normalizedDec = roundedDec % 24;
    let h = Math.floor(normalizedDec);
    let m = Math.round((normalizedDec - h) * 60);
    if (m === 60) { h++; m = 0; }
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

export function addDays(dateStr, days) {
    const [y, m, d] = dateStr.split('-').map(Number);
    const date = new Date(y, m - 1, d + days);
    const localY = date.getFullYear();
    const localM = String(date.getMonth() + 1).padStart(2, '0');
    const localD = String(date.getDate()).padStart(2, '0');
    return `${localY}-${localM}-${localD}`;
}

export function timeStrToDecimal(timeStr) {
    if (!timeStr) return null;
    const [h, m] = timeStr.split(':').map(Number);
    return h + (m / 60);
}

export function decimalToTimeStr(dec) {
    let norm = dec % 24;
    let h = Math.floor(norm);
    let m = Math.round((norm - h) * 60);
    if (m === 60) { h++; m = 0; }
    return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}`;
}

export function escapeHTML(str) { 
    return str.replace(/[&<>'"]/g, tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag)); 
}

export function parseTags(str) { 
    let html = escapeHTML(str);
    html = html.replace(/(#[a-zA-Z0-9_]+)/g, '<span class="hashtag" onclick="searchTag(\'$1\', event)">$1</span>');
    html = html.replace(/(@[a-zA-Z0-9_]+)/g, '<span class="person-tag" onclick="searchTag(\'$1\', event)">$1</span>');
    return html;
}

export function parseICS(icsText, targetDateStr, ignoreKeywords) {
    let events = [];
    let vTimezones = {};
    let lines = icsText.split(/\r\n|\n|\r/);
    let currentEvent = null;
    let currentTzId = null;
    let currentTzProp = null;

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        if (line.startsWith('BEGIN:VTIMEZONE')) {
            currentTzId = null;
        } else if (line.startsWith('TZID:')) {
            let colonIdx = line.indexOf(':');
            if (colonIdx !== -1) {
                currentTzId = line.substring(colonIdx + 1).trim();
                if (!vTimezones[currentTzId]) vTimezones[currentTzId] = {};
            }
        } else if (line.startsWith('BEGIN:STANDARD') || line.startsWith('BEGIN:DAYLIGHT')) {
            currentTzProp = line.includes('DAYLIGHT') ? 'DAYLIGHT' : 'STANDARD';
        } else if (line.startsWith('TZOFFSETFROM:') || line.startsWith('TZOFFSETTO:')) {
            if (currentTzId && currentTzProp) {
                let colonIdx = line.indexOf(':');
                let offsetStr = line.substring(colonIdx + 1).trim();
                let sign = offsetStr.startsWith('-') ? -1 : 1;
                let hrs = parseInt(offsetStr.substring(1, 3), 10) || 0;
                let mins = parseInt(offsetStr.substring(3, 5), 10) || 0;
                let totalHours = sign * (hrs + (mins / 60));
                
                if (!vTimezones[currentTzId][currentTzProp]) {
                    vTimezones[currentTzId][currentTzProp] = {};
                }
                if (line.startsWith('TZOFFSETTO:')) {
                    vTimezones[currentTzId][currentTzProp].offsetTo = totalHours;
                }
            }
        } else if (line.startsWith('DTSTART:')) {
            if (currentTzId && currentTzProp) {
                let colonIdx = line.indexOf(':');
                let ds = line.substring(colonIdx + 1).trim();
                if (ds.length >= 8) {
                    let ruleMonth = parseInt(ds.substring(4, 6), 10);
                    if (!vTimezones[currentTzId][currentTzProp]) {
                        vTimezones[currentTzId][currentTzProp] = {};
                    }
                    vTimezones[currentTzId][currentTzProp].startMonth = ruleMonth;
                }
            }
        } else if (line.startsWith('END:VTIMEZONE')) {
            currentTzId = null;
            currentTzProp = null;
        }
    }

    const fallbackOffsets = {
        'INDIA STANDARD TIME': 5.5, 'IST': 5.5,
        'AUS EASTERN STANDARD TIME': 10, 'E. AUSTRALIA STANDARD TIME': 10, 'AEST': 10, 'AEDT': 11,
        'GMT': 0, 'UTC': 0, 'PACIFIC STANDARD TIME': -8, 'EASTERN STANDARD TIME': -5, 'CENTRAL STANDARD TIME': -6
    };

    for (let line of lines) {
        if (line === 'BEGIN:VEVENT') {
            currentEvent = {};
        } else if (line === 'END:VEVENT' && currentEvent) {
            if (currentEvent.SUMMARY && Object.keys(currentEvent).some(k => k.startsWith('DTSTART'))) {
                events.push(currentEvent);
            }
            currentEvent = null;
        } else if (currentEvent) {
            let colonIdx = line.indexOf(':');
            if (colonIdx !== -1) {
                let fullKey = line.substring(0, colonIdx);
                let val = line.substring(colonIdx + 1);
                currentEvent[fullKey] = val;
            }
        }
    }

    let parsedEvents = [];
    events.forEach(ev => {
        let summary = ev.SUMMARY || 'Meeting';
        let lowerSummary = summary.toLowerCase();
        let shouldIgnore = ignoreKeywords.some(kw => lowerSummary.includes(kw));
        if (shouldIgnore) return;

        let dtStartKey = Object.keys(ev).find(k => k.startsWith('DTSTART'));
        let dtEndKey = Object.keys(ev).find(k => k.startsWith('DTEND'));

        if (!dtStartKey || !ev[dtStartKey]) return;

        let dtStartVal = ev[dtStartKey];
        let dtEndVal = dtEndKey ? ev[dtEndKey] : dtStartVal;

        let parseUniversalDate = (keyName, valStr) => {
            let clean = valStr.trim();
            let match = clean.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2}))?Z?$/);
            if (!match) return null;

            let y = parseInt(match[1], 10), m = parseInt(match[2], 10) - 1, d = parseInt(match[3], 10);
            let h = match[4] ? parseInt(match[4], 10) : 0, min = match[5] ? parseInt(match[5], 10) : 0, s = match[6] ? parseInt(match[6], 10) : 0;

            let dateObj;
            if (clean.endsWith('Z')) {
                dateObj = new Date(Date.UTC(y, m, d, h, min, s));
            } else {
                let tzidMatch = keyName.match(/TZID=([^;:]+)/i);
                let sourceOffset = undefined;

                if (tzidMatch) {
                    let tzidKey = tzidMatch[1].trim();
                    if (vTimezones[tzidKey]) {
                        let std = vTimezones[tzidKey]['STANDARD'];
                        let dst = vTimezones[tzidKey]['DAYLIGHT'];
                        
                        if (std && dst && std.startMonth && dst.startMonth) {
                            let eventMonth = m + 1; 
                            if (std.startMonth < dst.startMonth) {
                                sourceOffset = (eventMonth >= std.startMonth && eventMonth < dst.startMonth) ? std.offsetTo : dst.offsetTo;
                            } else {
                                sourceOffset = (eventMonth >= dst.startMonth && eventMonth < std.startMonth) ? dst.offsetTo : std.offsetTo;
                            }
                        } else {
                            let tzData = std || dst;
                            if (tzData) sourceOffset = tzData.offsetTo;
                        }
                    }
                    if (sourceOffset === undefined) sourceOffset = fallbackOffsets[tzidKey.toUpperCase()];
                }

                if (sourceOffset !== undefined) {
                    let baseOffsetHours = getTzOffset(state.appConfig.primaryTz);
                    let netOffsetDiff = baseOffsetHours - sourceOffset;
                    let eventUTC = Date.UTC(y, m, d, h, min, s);
                    let actualUTC = eventUTC + (netOffsetDiff * 3600000);
                    dateObj = new Date(actualUTC);
                } else {
                    dateObj = new Date(y, m, d, h, min, s);
                }
            }

            return { dateStr: dateObj.toLocaleDateString('en-CA').split('T')[0], hour: dateObj.getHours() + (dateObj.getMinutes() / 60) };
        };

        try {
            let startParsed = parseUniversalDate(dtStartKey, dtStartVal);
            let endParsed = parseUniversalDate(dtEndKey || dtStartKey, dtEndVal);

            if (startParsed && endParsed && startParsed.dateStr === targetDateStr) {
                let startHour = roundToQuarterHour(startParsed.hour);
                let endHour = roundToQuarterHour(endParsed.hour);
                if (endHour <= startHour) endHour += 24;
                let duration = Math.max(0.25, roundToQuarterHour(endHour - startHour));

                parsedEvents.push({ id: ev.UID || (Math.random().toString()), summary: summary, startHour: startHour, duration: duration });
            }
        } catch(e) {}
    });

    return parsedEvents;
}

export function parseICS(icsText, targetDateStr, ignoreKeywords) {
    let events = [];
    let vTimezones = {};
    let lines = icsText.split(/\r\n|\n|\r/);
    let currentEvent = null;
    let currentTzId = null;
    let currentTzProp = null;

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        if (line.startsWith('BEGIN:VTIMEZONE')) {
            currentTzId = null;
        } else if (line.startsWith('TZID:')) {
            let colonIdx = line.indexOf(':');
            if (colonIdx !== -1) {
                currentTzId = line.substring(colonIdx + 1).trim();
                if (!vTimezones[currentTzId]) vTimezones[currentTzId] = {};
            }
        } else if (line.startsWith('BEGIN:STANDARD') || line.startsWith('BEGIN:DAYLIGHT')) {
            currentTzProp = line.includes('DAYLIGHT') ? 'DAYLIGHT' : 'STANDARD';
        } else if (line.startsWith('TZOFFSETFROM:') || line.startsWith('TZOFFSETTO:')) {
            if (currentTzId && currentTzProp) {
                let colonIdx = line.indexOf(':');
                let offsetStr = line.substring(colonIdx + 1).trim();
                let sign = offsetStr.startsWith('-') ? -1 : 1;
                let hrs = parseInt(offsetStr.substring(1, 3), 10) || 0;
                let mins = parseInt(offsetStr.substring(3, 5), 10) || 0;
                let totalHours = sign * (hrs + (mins / 60));
                
                if (!vTimezones[currentTzId][currentTzProp]) {
                    vTimezones[currentTzId][currentTzProp] = {};
                }
                if (line.startsWith('TZOFFSETTO:')) {
                    vTimezones[currentTzId][currentTzProp].offsetTo = totalHours;
                }
            }
        } else if (line.startsWith('DTSTART:')) {
            if (currentTzId && currentTzProp) {
                let colonIdx = line.indexOf(':');
                let ds = line.substring(colonIdx + 1).trim();
                if (ds.length >= 8) {
                    let ruleMonth = parseInt(ds.substring(4, 6), 10);
                    if (!vTimezones[currentTzId][currentTzProp]) {
                        vTimezones[currentTzId][currentTzProp] = {};
                    }
                    vTimezones[currentTzId][currentTzProp].startMonth = ruleMonth;
                }
            }
        } else if (line.startsWith('END:VTIMEZONE')) {
            currentTzId = null;
            currentTzProp = null;
        }
    }

    const fallbackOffsets = {
        'INDIA STANDARD TIME': 5.5, 'IST': 5.5,
        'AUS EASTERN STANDARD TIME': 10, 'E. AUSTRALIA STANDARD TIME': 10, 'AEST': 10, 'AEDT': 11,
        'GMT': 0, 'UTC': 0, 'PACIFIC STANDARD TIME': -8, 'EASTERN STANDARD TIME': -5, 'CENTRAL STANDARD TIME': -6
    };

    for (let line of lines) {
        if (line === 'BEGIN:VEVENT') {
            currentEvent = {};
        } else if (line === 'END:VEVENT' && currentEvent) {
            if (currentEvent.SUMMARY && Object.keys(currentEvent).some(k => k.startsWith('DTSTART'))) {
                events.push(currentEvent);
            }
            currentEvent = null;
        } else if (currentEvent) {
            let colonIdx = line.indexOf(':');
            if (colonIdx !== -1) {
                let fullKey = line.substring(0, colonIdx);
                let val = line.substring(colonIdx + 1);
                currentEvent[fullKey] = val;
            }
        }
    }

    let parsedEvents = [];
    events.forEach(ev => {
        let summary = ev.SUMMARY || 'Meeting';
        let lowerSummary = summary.toLowerCase();
        let shouldIgnore = ignoreKeywords.some(kw => lowerSummary.includes(kw));
        if (shouldIgnore) return;

        let dtStartKey = Object.keys(ev).find(k => k.startsWith('DTSTART'));
        let dtEndKey = Object.keys(ev).find(k => k.startsWith('DTEND'));

        if (!dtStartKey || !ev[dtStartKey]) return;

        let dtStartVal = ev[dtStartKey];
        let dtEndVal = dtEndKey ? ev[dtEndKey] : dtStartVal;

        let parseUniversalDate = (keyName, valStr) => {
            let clean = valStr.trim();
            let match = clean.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2}))?Z?$/);
            if (!match) return null;

            let y = parseInt(match[1], 10), m = parseInt(match[2], 10) - 1, d = parseInt(match[3], 10);
            let h = match[4] ? parseInt(match[4], 10) : 0, min = match[5] ? parseInt(match[5], 10) : 0, s = match[6] ? parseInt(match[6], 10) : 0;

            let dateObj;
            if (clean.endsWith('Z')) {
                dateObj = new Date(Date.UTC(y, m, d, h, min, s));
            } else {
                let tzidMatch = keyName.match(/TZID=([^;:]+)/i);
                let sourceOffset = undefined;

                if (tzidMatch) {
                    let tzidKey = tzidMatch[1].trim();
                    if (vTimezones[tzidKey]) {
                        let std = vTimezones[tzidKey]['STANDARD'];
                        let dst = vTimezones[tzidKey]['DAYLIGHT'];
                        
                        if (std && dst && std.startMonth && dst.startMonth) {
                            let eventMonth = m + 1; 
                            if (std.startMonth < dst.startMonth) {
                                sourceOffset = (eventMonth >= std.startMonth && eventMonth < dst.startMonth) ? std.offsetTo : dst.offsetTo;
                            } else {
                                sourceOffset = (eventMonth >= dst.startMonth && eventMonth < std.startMonth) ? dst.offsetTo : std.offsetTo;
                            }
                        } else {
                            let tzData = std || dst;
                            if (tzData) sourceOffset = tzData.offsetTo;
                        }
                    }
                    if (sourceOffset === undefined) sourceOffset = fallbackOffsets[tzidKey.toUpperCase()];
                }

                if (sourceOffset !== undefined) {
                    let baseOffsetHours = getTzOffset(state.appConfig.primaryTz);
                    let netOffsetDiff = baseOffsetHours - sourceOffset;
                    let eventUTC = Date.UTC(y, m, d, h, min, s);
                    let actualUTC = eventUTC + (netOffsetDiff * 3600000);
                    dateObj = new Date(actualUTC);
                } else {
                    dateObj = new Date(y, m, d, h, min, s);
                }
            }

            return { dateStr: dateObj.toLocaleDateString('en-CA').split('T')[0], hour: dateObj.getHours() + (dateObj.getMinutes() / 60) };
        };

        try {
            let startParsed = parseUniversalDate(dtStartKey, dtStartVal);
            let endParsed = parseUniversalDate(dtEndKey || dtStartKey, dtEndVal);

            if (startParsed && endParsed && startParsed.dateStr === targetDateStr) {
                let startHour = roundToQuarterHour(startParsed.hour);
                let endHour = roundToQuarterHour(endParsed.hour);
                if (endHour <= startHour) endHour += 24;
                let duration = Math.max(0.25, roundToQuarterHour(endHour - startHour));

                parsedEvents.push({ id: ev.UID || (Math.random().toString()), summary: summary, startHour: startHour, duration: duration });
            }
        } catch(e) {}
    });

    return parsedEvents;
}