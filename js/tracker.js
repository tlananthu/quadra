import { state, saveNotes } from './state.js';
import { roundToQuarterHour, getTzOffset, getTzTime, formatCurrentTimeBadge, decToTime, addDays, showToast, parseTags } from './utils.js';
import { openTaskModal, openEditBlockModal } from './ui.js';
import { matchesSearchQuery } from './board.js';

export function setTrackerMode(mode) {
    state.currentTrackerMode = mode;
    document.getElementById('btnTrackerDay').classList.toggle('active', mode === 'day');
    document.getElementById('btnTrackerWeek').classList.toggle('active', mode === 'week');
    renderTrackerTimeline();
}

export function startLiveClock() {
    if(state.clockIntervalId) clearInterval(state.clockIntervalId);
    state.clockIntervalId = setInterval(() => {
        if (state.currentLayout === 'tracker' && state.currentTrackerMode === 'day' && document.getElementById('tracker-view').style.display === 'block') {
            const line = document.querySelector('.current-time-line');
            const badge = document.querySelector('.current-time-badge');
            if (line && badge) {
                const hourPx = 60 * state.timelineZoom;
                const now = new Date();
                const primaryTime = getTzTime(now, state.appConfig.primaryTz);
                const currentHour = primaryTime.h + (primaryTime.m / 60) + (primaryTime.s / 3600);
                
                line.style.top = `${currentHour * hourPx}px`;
                badge.innerText = formatCurrentTimeBadge(now);
            }
        }
    }, 30000); 
}

export function stopLiveClock() {
    if(state.clockIntervalId) { clearInterval(state.clockIntervalId); state.clockIntervalId = null; }
}

export function changeTrackerDay(offset) {
    const dateInput = document.getElementById('trackerDate');
    const d = new Date(dateInput.value);
    d.setDate(d.getDate() + offset);
    dateInput.value = d.toLocaleDateString('en-CA').split('T')[0];
    renderTrackerTimeline();
}

export function goToToday() {
    const dateInput = document.getElementById('trackerDate');
    dateInput.value = new Date().toLocaleDateString('en-CA').split('T')[0];
    renderTrackerTimeline();
}

export function adjustTimelineZoom(amount) {
    state.timelineZoom = Math.max(0.5, Math.min(3, roundToQuarterHour(state.timelineZoom + amount)));
    localStorage.setItem('quadra_zoom', state.timelineZoom);
    updateZoomDisplay();
    renderTrackerTimeline();
}

export function updateZoomDisplay() {
    const display = document.getElementById('zoomLevelDisplay');
    if (display) display.innerText = `${state.timelineZoom}x`;
}

export function exportTimesheet() {
    const baseDateStr = document.getElementById('trackerDate').value;
    const [y, m, d] = baseDateStr.split('-');
    const baseDate = new Date(y, m - 1, d);
    let datesToExport = [];
    
    if (state.currentTrackerMode === 'day') {
        datesToExport.push(baseDateStr);
    } else {
        const dayOfWeek = baseDate.getDay();
        const startOfWeek = new Date(baseDate);
        startOfWeek.setDate(baseDate.getDate() - dayOfWeek);
        for(let i=0; i<7; i++) {
            let dateIter = new Date(startOfWeek);
            dateIter.setDate(startOfWeek.getDate() + i);
            const localY = dateIter.getFullYear();
            const localM = String(dateIter.getMonth() + 1).padStart(2, '0');
            const localD = String(dateIter.getDate()).padStart(2, '0');
            datesToExport.push(`${localY}-${localM}-${localD}`);
        }
    }

    let plainTextLog = `Timesheet Report\n`;
    let overallTotal = 0;

    datesToExport.forEach(dateStr => {
        let dailyHrs = 0;
        let dailyLog = "";
        state.notes.forEach(note => {
            if (note.deleted || !note.timeBlocks) return;
            note.timeBlocks.forEach(tb => { 
                if(tb.date === dateStr) {
                    let dur = roundToQuarterHour(tb.duration);
                    dailyHrs += dur;
                    const tags = note.text.match(/(#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+)/g);
                    let identifier = tags ? tags.join(' ') : note.text.substring(0, 40).replace(/\n/g, ' ') + '...';
                    dailyLog += `- [${identifier}] ${dur}h\n`;
                }
            });
        });
        
        if (dailyHrs > 0 || state.currentTrackerMode === 'day') {
             plainTextLog += `\nDate: ${dateStr}\n------------------------\n`;
             if(dailyHrs > 0) plainTextLog += dailyLog;
             else plainTextLog += "No time logged.\n";
             plainTextLog += `Total: ${roundToQuarterHour(dailyHrs)}h\n`;
        }
        overallTotal += dailyHrs;
    });
    
    if (state.currentTrackerMode === 'week') {
         plainTextLog += `\n========================\nWeekly Total: ${roundToQuarterHour(overallTotal)}h\n`;
    }

    navigator.clipboard.writeText(plainTextLog).then(() => showToast("Timesheet copied to clipboard!")).catch(err => showToast("Copy failed"));
}

export function startBlockDrag(e, noteId, blockId, isResize) {
    e.stopPropagation();
    const note = state.notes.find(n => n.id === noteId);
    const block = note.timeBlocks.find(b => b.id === blockId);
    
    state.dragState = {
        noteId, blockId, isResize,
        startY: e.clientY,
        originalStart: roundToQuarterHour(block.startHour),
        originalDuration: roundToQuarterHour(block.duration),
        hasMoved: false
    };
    
    document.addEventListener('mousemove', onBlockDrag);
    document.addEventListener('mouseup', stopBlockDrag);
}

export function onBlockDrag(e) {
    if (!state.dragState) return;
    state.dragState.hasMoved = true;
    const hourPx = 60 * state.timelineZoom;
    
    const dy = e.clientY - state.dragState.startY;
    const dHours = roundToQuarterHour(dy / hourPx); 
    
    const note = state.notes.find(n => n.id === state.dragState.noteId);
    const block = note.timeBlocks.find(b => b.id === state.dragState.blockId);
    
    if (state.dragState.isResize) {
        block.duration = Math.max(0.25, roundToQuarterHour(state.dragState.originalDuration + dHours));
    } else {
        block.startHour = Math.max(0, Math.min(30 - block.duration, roundToQuarterHour(state.dragState.originalStart + dHours)));
    }
    
    renderTrackerTimeline();
}

export function stopBlockDrag(e) {
    if (state.dragState) {
        document.removeEventListener('mousemove', onBlockDrag);
        document.removeEventListener('mouseup', stopBlockDrag);
        
        if (state.dragState.hasMoved) {
            const note = state.notes.find(n => n.id === state.dragState.noteId);
            note.dirty = true; saveNotes(); renderTrackerTimeline(); 
        }
        state.dragState = null;
    }
}

export function handleTimelineClick(ev, dateStr) {
    if (ev.target.closest('.logged-block')) return; 
    const hourPx = 60 * state.timelineZoom;
    const rect = ev.currentTarget.getBoundingClientRect();
    const y = ev.clientY - rect.top; 
    let dropHour = roundToQuarterHour(y / hourPx);
    openTaskModal('inbox', null, ev, { date: dateStr, startHour: dropHour });
}

export function allowTrackerDrop(ev) { ev.preventDefault(); ev.currentTarget.classList.add('drag-over'); }
export function dragLeaveTracker(ev) { ev.currentTarget.classList.remove('drag-over'); }

export function dropToTracker(ev, dateStr) {
    ev.preventDefault(); ev.currentTarget.classList.remove('drag-over');
    const noteId = ev.dataTransfer.getData("text/plain");
    const note = state.notes.find(n => n.id === noteId);
    const hourPx = 60 * state.timelineZoom;
    
    if (note) {
        const rect = ev.currentTarget.getBoundingClientRect();
        const y = ev.clientY - rect.top; 
        let dropHour = roundToQuarterHour(y / hourPx);

        if (!note.timeBlocks) note.timeBlocks = [];
        note.timeBlocks.push({ id: Date.now().toString(), date: dateStr, startHour: dropHour, duration: 1.0 });
        note.dirty = true; saveNotes(); 
        if (window.handleSearch) window.handleSearch(); 
    }
}

export function renderTrackerTimeline() {
    const searchInput = document.getElementById('searchInput');
    const query = searchInput ? searchInput.value.toLowerCase() : '';
    const canvas = document.getElementById('timelineCanvas');
    const paletteList = document.getElementById('tracker-palette-list');
    const hourPx = 60 * state.timelineZoom;
    
    paletteList.innerHTML = '';
    const filteredNotes = state.notes.filter(n => !n.deleted && matchesSearchQuery(n.text, query) && !n.eventId);
    
    filteredNotes.forEach(note => {
        const el = document.createElement('div');
        el.className = 'palette-note';
        el.draggable = true;
        el.ondragstart = (e) => e.dataTransfer.setData('text/plain', note.id);
        el.onclick = (e) => openTaskModal(null, note.id, e);
        
        let title = note.text.split('\n')[0].substring(0, 50);
        el.innerHTML = `<strong>${note.status === 'closed' ? '✓ ' : ''}</strong> ${parseTags(title)}`;
        if (note.status === 'closed') el.style.opacity = '0.6';
        paletteList.appendChild(el);
    });

    canvas.innerHTML = '';
    canvas.style.height = `${24 * hourPx}px`;

    const bgLines = document.createElement('div');
    bgLines.className = 'timeline-bg-lines';
    
    let hasSecTz = state.appConfig.secondaryTz && state.appConfig.secondaryTz !== 'none';
    let secOffsetDiff = 0;
    if (hasSecTz) {
        let primOff = getTzOffset(state.appConfig.primaryTz);
        let secOff = getTzOffset(state.appConfig.secondaryTz);
        secOffsetDiff = secOff - primOff;
    }

    for (let i = 0; i <= 24; i++) {
        const row = document.createElement('div');
        row.className = 'time-row'; row.style.top = `${i * hourPx}px`;
        
        let labelText = `${i.toString().padStart(2, '0')}:00`;
        if (hasSecTz) {
            let secHourRaw = i + secOffsetDiff;
            let sH = Math.floor(secHourRaw);
            let sM = Math.round((secHourRaw - sH) * 60);
            if (sM < 0) { sM += 60; sH -= 1; }
            if (sM === 60) { sH += 1; sM = 0; }
            let dispH = sH % 24;
            if (dispH < 0) dispH += 24;
            labelText += ` (${dispH.toString().padStart(2, '0')}:${sM.toString().padStart(2, '0')})`;
        }
        
        row.innerHTML = `<span class="time-label">${labelText}</span>`;
        bgLines.appendChild(row);
    }
    canvas.appendChild(bgLines);

    const baseDateStr = document.getElementById('trackerDate').value;
    localStorage.setItem('quadra_tracker_date', baseDateStr);
    
    const [y, m, d] = baseDateStr.split('-');
    const baseDate = new Date(y, m - 1, d);
    
    const daySpan = document.getElementById('trackerDayOfWeek');
    if (daySpan) {
        const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        daySpan.innerText = dayNames[baseDate.getDay()];
    }

    const weekStart = new Date(baseDate);
    const dayOfWeek = baseDate.getDay();
    weekStart.setDate(baseDate.getDate() - dayOfWeek);
    const weekDateKeys = new Set();
    for (let i = 0; i < 7; i++) {
        const dateIter = new Date(weekStart);
        dateIter.setDate(weekStart.getDate() + i);
        const localY = dateIter.getFullYear();
        const localM = String(dateIter.getMonth() + 1).padStart(2, '0');
        const localD = String(dateIter.getDate()).padStart(2, '0');
        weekDateKeys.add(`${localY}-${localM}-${localD}`);
    }

    let datesToRender = [];
    
    if (state.currentTrackerMode === 'day') {
        datesToRender.push({ date: baseDateStr, label: '' });
    } else {
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        for(let i=0; i<7; i++) {
            let dateIter = new Date(weekStart);
            dateIter.setDate(weekStart.getDate() + i);
            const localY = dateIter.getFullYear();
            const localM = String(dateIter.getMonth() + 1).padStart(2, '0');
            const localD = String(dateIter.getDate()).padStart(2, '0');
            datesToRender.push({ date: `${localY}-${localM}-${localD}`, label: `${dayNames[i]} ${dateIter.getDate()}` });
        }
    }

    const colsContainer = document.createElement('div');
    colsContainer.className = 'timeline-cols-container';
    const timeGutterWidth = hasSecTz ? 120 : 55;
    const overlayPadding = hasSecTz ? 130 : 65;
    colsContainer.style.left = `${timeGutterWidth}px`;
    canvas.appendChild(colsContainer);

    let totalTimeRendered = 0;
    let weeklyTotalRendered = 0;

    state.notes.forEach(note => {
        if (note.deleted || !note.timeBlocks) return;
        const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
        if (!isCalendarEvent && !matchesSearchQuery(note.text, query)) return;
        
        note.timeBlocks.forEach(tb => {
            const duration = roundToQuarterHour(tb.duration);
            if (!duration) return;
            const blockStart = roundToQuarterHour(tb.startHour);
            const blockEnd = blockStart + duration;

            if (weekDateKeys.has(tb.date)) weeklyTotalRendered += duration;

            if (tb.date === baseDateStr) {
                totalTimeRendered += duration;
            } else if (tb.date === addDays(baseDateStr, -1) && blockStart < 24 && blockEnd > 24) {
                totalTimeRendered += roundToQuarterHour(blockEnd - 24);
            }
        });
    });

    datesToRender.forEach(dtObj => {
        const dateStr = dtObj.date;
        const col = document.createElement('div');
        col.className = 'time-col';
        col.ondragover = allowTrackerDrop;
        col.ondragleave = dragLeaveTracker;
        col.ondrop = (e) => dropToTracker(e, dateStr);
        col.onclick = (e) => handleTimelineClick(e, dateStr);

        if (state.currentTrackerMode === 'week') {
            const header = document.createElement('div');
            header.className = 'col-header';
            const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0]; 
            if (dateStr === todayStr) header.classList.add('today');
            header.innerText = dtObj.label;
            col.appendChild(header);
        }

        if (state.currentTrackerMode === 'day') {
            const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];
            if (dateStr === todayStr) {
                const now = new Date();
                const primaryTime = getTzTime(now, state.appConfig.primaryTz);
                const currentHour = primaryTime.h + (primaryTime.m / 60) + (primaryTime.s / 3600);
                
                const timeLine = document.createElement('div');
                timeLine.className = 'current-time-line';
                timeLine.style.top = `${currentHour * hourPx}px`;
                timeLine.innerHTML = `<span class="current-time-badge">${formatCurrentTimeBadge(now)}</span>`;
                col.appendChild(timeLine);
            }
        }

        state.appSchedule.forEach(block => {
            let start = roundToQuarterHour(block.startHour);
            let end = roundToQuarterHour(block.endHour);

            if (start < 24) {
                let renderEnd = Math.min(end, 24);
                const overlay = document.createElement('div');
                overlay.className = 'schedule-overlay';
                overlay.style.top = `${start * hourPx}px`;
                overlay.style.height = `${(renderEnd - start) * hourPx}px`;
                overlay.style.left = `-${timeGutterWidth}px`;
                overlay.style.paddingLeft = `${overlayPadding}px`;
                if (state.currentTrackerMode === 'day') overlay.innerText = block.title;
                col.appendChild(overlay);
            }

            state.appSchedule.forEach(prevBlock => {
                if (prevBlock.endHour > 24 && dtObj.date === dateStr) {
                    let wrappedSpan = roundToQuarterHour(prevBlock.endHour - 24);
                    if (wrappedSpan > 0) {
                        const overlay = document.createElement('div');
                        overlay.className = 'schedule-overlay';
                        overlay.style.top = `0px`;
                        overlay.style.height = `${Math.min(wrappedSpan, 24) * hourPx}px`;
                        overlay.style.left = `-${timeGutterWidth}px`;
                        overlay.style.paddingLeft = `${overlayPadding}px`;
                        if (state.currentTrackerMode === 'day') overlay.innerText = prevBlock.title;
                        col.appendChild(overlay);
                    }
                }
            });
        });

        state.notes.forEach(note => {
            if (note.deleted || !note.timeBlocks) return;
            const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
            if (!isCalendarEvent && !matchesSearchQuery(note.text, query)) return;

            note.timeBlocks.forEach(tb => {
                let blockStart = roundToQuarterHour(tb.startHour);
                let duration = roundToQuarterHour(tb.duration);
                let blockEnd = blockStart + duration;
                if (blockEnd < blockStart) { blockEnd += 24; }
                let actualDuration = roundToQuarterHour(blockEnd - blockStart);

                if (tb.date === dateStr) {
                    let renderStart = blockStart % 24;
                    let renderDuration = duration;

                    if (blockStart < 24 && blockEnd > 24) {
                        renderDuration = 24 - blockStart; 
                    } else if (blockStart >= 24) {
                        return; 
                    }

                    const blockEl = document.createElement('div');
                    const isMeeting = note.eventId || note.text.includes('#meeting');
                    const quadClass = note.quadrant || 'q2';
                    blockEl.className = 'logged-block' + (isMeeting ? ' is-meeting' : ` ${quadClass}`);
                    blockEl.id = `block-${tb.id}`;
                    blockEl.style.top = `${renderStart * hourPx}px`;
                    blockEl.style.height = `${Math.max(15, renderDuration * hourPx)}px`;
                    
                    let displayTitle = note.text.split('\n')[0].substring(0, 40);
                    const actualEndHour = (blockStart + actualDuration) % 24;
                    const timeStr = `${decToTime(blockStart)} - ${decToTime(actualEndHour)} (${actualDuration}h)`;

                    blockEl.innerHTML = `
                        <div class="block-info">
                            <div class="block-title">${displayTitle}</div>
                            <div class="block-meta">${timeStr}</div>
                        </div>
                        <button class="block-edit-icon" onclick="openEditBlockModal('${note.id}', '${tb.id}', event)" title="Edit Time Block Hours">📝</button>
                        <div class="resize-handle" onmousedown="startBlockDrag(event, '${note.id}', '${tb.id}', true)"></div>
                    `;
                    
                    blockEl.onclick = (e) => {
                        if(e.target.closest('.block-edit-icon') || e.target.closest('.resize-handle')) return;
                        openTaskModal(null, note.id, e);
                    };

                    blockEl.onmousedown = (e) => {
                        if(e.target.closest('.block-edit-icon') || e.target.closest('.resize-handle')) return;
                        startBlockDrag(e, note.id, tb.id, false);
                    };
                    col.appendChild(blockEl);
                } 
                else if (tb.date === addDays(dateStr, -1)) {
                    if (blockStart < 24 && blockEnd > 24) {
                        let overflowDuration = roundToQuarterHour(blockEnd - 24);
                        const blockEl = document.createElement('div');
                        const isMeeting = note.eventId || note.text.includes('#meeting');
                        const quadClass = note.quadrant || 'q2';
                        blockEl.className = 'logged-block' + (isMeeting ? ' is-meeting' : ` ${quadClass}`);
                        blockEl.id = `block-overflow-${tb.id}`;
                        blockEl.style.top = `0px`;
                        blockEl.style.height = `${Math.max(15, overflowDuration * hourPx)}px`;
                        
                        let displayTitle = note.text.split('\n')[0].substring(0, 40);
                        const actualEndHour = (blockStart + actualDuration) % 24;
                        const timeStr = `${decToTime(blockStart)} - ${decToTime(actualEndHour)} (${actualDuration}h)`;

                        blockEl.innerHTML = `
                            <div class="block-info">
                                <div class="block-title">${displayTitle} (cont.)</div>
                                <div class="block-meta">${timeStr}</div>
                            </div>
                            <button class="block-edit-icon" onclick="openEditBlockModal('${note.id}', '${tb.id}', event)" title="Edit Time Block Hours">📝</button>
                        `;
                        blockEl.onclick = (e) => {
                            if(e.target.closest('.block-edit-icon')) return;
                            openTaskModal(null, note.id, e);
                        };
                        col.appendChild(blockEl);
                    }
                }
            });
        });

        colsContainer.appendChild(col);
    });

    const dayTotal = roundToQuarterHour(totalTimeRendered);
    const weekTotal = roundToQuarterHour(weeklyTotalRendered);
    const dailyTotalEl = document.getElementById('trackerDailyTotal');
    if(dailyTotalEl) dailyTotalEl.innerText = `Total: ${dayTotal}h/${weekTotal}h`;

    const scrollArea = document.getElementById('timelineScrollArea');
    if (scrollArea && scrollArea.scrollTop === 0) scrollArea.scrollTop = 7 * hourPx; 
}