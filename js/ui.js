import { state, saveNotes, saveConfig, saveSchedule } from './state.js';
import { showToast, timeStrToDecimal, decimalToTimeStr, roundToQuarterHour } from './utils.js';
import { renderTrackerTimeline, startLiveClock, stopLiveClock, updateZoomDisplay } from './tracker.js';
import { renderOverdueTasksPage, handleSearch } from './board.js';
import { loadCalendars, checkConfigState } from './api.js';

export function applyViewVisibility() {
    const btnGrid = document.getElementById('btnGrid');
    const btnKanban = document.getElementById('btnKanban');
    const btnTracker = document.getElementById('btnTracker');
    const btnOverdue = document.getElementById('btnOverdue');
    const btnGcal = document.getElementById('btnGcal');
    
    if (btnGrid) btnGrid.style.display = state.appConfig.viewsEnabled.grid ? '' : 'none';
    if (btnKanban) btnKanban.style.display = state.appConfig.viewsEnabled.kanban ? '' : 'none';
    if (btnTracker) btnTracker.style.display = state.appConfig.viewsEnabled.tracker ? '' : 'none';
    if (btnOverdue) btnOverdue.style.display = state.appConfig.viewsEnabled.overdue ? '' : 'none';
    if (btnGcal) btnGcal.style.display = state.appConfig.viewsEnabled.gcal ? '' : 'none';
    
    if (!state.appConfig.viewsEnabled[state.currentLayout]) {
        setLayout(state.appConfig.defaultView);
    }
}

export function setLayout(layout) {
    if (!state.appConfig.viewsEnabled[layout]) return;
    
    state.currentLayout = layout;
    localStorage.setItem('quadra_layout', layout);
    
    const matrix = document.getElementById('matrix');
    const tracker = document.getElementById('tracker-view');
    const overdue = document.getElementById('overdue-view');
    const gcalView = document.getElementById('gcal-view');
    const settingsView = document.getElementById('settings-view');
    
    const btnGrid = document.getElementById('btnGrid');
    const btnKanban = document.getElementById('btnKanban');
    const btnOverdue = document.getElementById('btnOverdue');
    const btnGcal = document.getElementById('btnGcal');
    const btnTracker = document.getElementById('btnTracker');

    if (btnGrid) btnGrid.classList.toggle('active', layout === 'grid');
    if (btnKanban) btnKanban.classList.toggle('active', layout === 'kanban');
    if (btnOverdue) btnOverdue.classList.toggle('active', layout === 'overdue');
    if (btnGcal) btnGcal.classList.toggle('active', layout === 'gcal');
    if (btnTracker) btnTracker.classList.toggle('active', layout === 'tracker');
    
    if(settingsView) settingsView.style.display = 'none';
    if(matrix) matrix.style.display = 'none';
    if(tracker) tracker.style.display = 'none';
    if(overdue) overdue.style.display = 'none';
    if(gcalView) gcalView.style.display = 'none';

    if (layout === 'tracker') {
        if(tracker) tracker.style.display = 'block'; 
        document.body.classList.remove('sidebar-open');
        updateZoomDisplay();
        renderTrackerTimeline();
        startLiveClock();
    } else if (layout === 'overdue') {
        if(overdue) overdue.style.display = 'block';
        document.body.classList.remove('sidebar-open');
        renderOverdueTasksPage();
        stopLiveClock();
    } else if (layout === 'gcal') {
        if(gcalView) gcalView.style.display = 'block';
        document.body.classList.remove('sidebar-open');
        stopLiveClock();
        
        const iframe = document.getElementById('gcal-iframe');
        const defaultUrl = 'https://calendar.google.com/calendar/embed?src=primary';
        const targetUrl = state.appConfig.gcalEmbedUrl || defaultUrl;
        
        if (iframe && iframe.src !== targetUrl) {
            iframe.src = targetUrl;
        }
    } else {
        if(matrix) {
            matrix.style.display = ''; 
            matrix.classList.remove('layout-grid', 'layout-kanban');
            matrix.classList.add(`layout-${layout}`);
        }
        stopLiveClock();
    }
    document.getElementById('quick-tags-bar').style.display = 'flex';
    document.getElementById('searchHeaderContainer').style.display = 'block';
    document.getElementById('viewToggleGroup').style.display = 'flex';
    document.getElementById('settingsNavBtn').style.display = 'inline-block';
    document.getElementById('backNavBtn').style.display = 'none';
}

export function openSettingsPage() {
    document.getElementById('matrix').style.display = 'none';
    document.getElementById('tracker-view').style.display = 'none';
    document.getElementById('overdue-view').style.display = 'none';
    document.getElementById('gcal-view').style.display = 'none';
    document.getElementById('settings-view').style.display = 'block';
    document.body.classList.remove('sidebar-open');
    document.getElementById('quick-tags-bar').style.display = 'none';
    document.getElementById('searchHeaderContainer').style.display = 'none';
    document.getElementById('viewToggleGroup').style.display = 'none';
    document.getElementById('settingsNavBtn').style.display = 'none';
    document.getElementById('backNavBtn').style.display = 'inline-block';

    document.getElementById('configClientId').value = state.appConfig.clientId || '';
    document.getElementById('configApiKey').value = state.appConfig.apiKey || '';
    document.getElementById('configIgnoreKeywords').value = state.appConfig.ignoreKeywords || '';
    document.getElementById('configCalSource').value = state.appConfig.calSource || 'google';
    document.getElementById('configIcsUrl').value = state.appConfig.icsUrl || '';
    document.getElementById('configGcalEmbedUrl').value = state.appConfig.gcalEmbedUrl || '';
    
    document.getElementById('configPrimaryTz').value = state.appConfig.primaryTz || 'local';
    document.getElementById('configSecondaryTz').value = state.appConfig.secondaryTz || 'none';
    
    document.getElementById('configDefaultView').value = state.appConfig.defaultView || 'grid';
    document.getElementById('configViewGrid').checked = state.appConfig.viewsEnabled.grid !== false;
    document.getElementById('configViewKanban').checked = state.appConfig.viewsEnabled.kanban !== false;
    document.getElementById('configViewTracker').checked = state.appConfig.viewsEnabled.tracker !== false;
    document.getElementById('configViewOverdue').checked = state.appConfig.viewsEnabled.overdue !== false;
    document.getElementById('configViewGcal').checked = state.appConfig.viewsEnabled.gcal !== false;
    
    toggleCalSourceFields(state.appConfig.calSource || 'google');
    loadCalendars();
    renderScheduleSettings();
}

export function toggleCalSourceFields(source) {
    if (source === 'outlook') {
        document.getElementById('googleCalGroup').style.display = 'none';
        document.getElementById('outlookIcsGroup').style.display = 'block';
    } else {
        document.getElementById('googleCalGroup').style.display = 'block';
        document.getElementById('outlookIcsGroup').style.display = 'none';
    }
}

export function closeSettingsPage() {
    setLayout(state.currentLayout);
}

export function saveSettings() {
    state.appConfig.clientId = document.getElementById('configClientId').value.trim();
    state.appConfig.apiKey = document.getElementById('configApiKey').value.trim();
    state.appConfig.ignoreKeywords = document.getElementById('configIgnoreKeywords').value.trim();
    state.appConfig.calSource = document.getElementById('configCalSource').value;
    state.appConfig.icsUrl = document.getElementById('configIcsUrl').value.trim();
    state.appConfig.gcalEmbedUrl = document.getElementById('configGcalEmbedUrl').value.trim();
    
    state.appConfig.primaryTz = document.getElementById('configPrimaryTz').value;
    state.appConfig.secondaryTz = document.getElementById('configSecondaryTz').value;
    
    const calendarSelect = document.getElementById('configCalendar');
    if (calendarSelect && calendarSelect.value) {
        state.appConfig.calendarId = calendarSelect.value;
    }

    state.appConfig.defaultView = document.getElementById('configDefaultView').value;
    state.appConfig.viewsEnabled = {
        grid: document.getElementById('configViewGrid').checked,
        kanban: document.getElementById('configViewKanban').checked,
        tracker: document.getElementById('configViewTracker').checked,
        gcal: document.getElementById('configViewGcal').checked,
        overdue: document.getElementById('configViewOverdue').checked
    };
    
    if (!state.appConfig.viewsEnabled[state.appConfig.defaultView]) {
        state.appConfig.viewsEnabled[state.appConfig.defaultView] = true;
    }
    
    saveConfig();
    
    const rows = document.querySelectorAll('.schedule-row');
    state.appSchedule = Array.from(rows).map(row => {
        const startStr = row.querySelector('.sched-start').value.split(':');
        const endStr = row.querySelector('.sched-end').value.split(':');
        
        let startH = parseInt(startStr[0] || 0) + (parseInt(startStr[1] || 0) / 60);
        let endH = parseInt(endStr[0] || 0) + (parseInt(endStr[1] || 0) / 60);
        
        if (endH <= startH) endH += 24;

        return {
            title: row.querySelector('.sched-title').value,
            startHour: roundToQuarterHour(startH),
            endHour: roundToQuarterHour(endH)
        };
    });
    saveSchedule();
    
    applyViewVisibility();
    closeSettingsPage(); 
    showToast("Configuration saved!");
    checkConfigState();
    if (state.currentLayout === 'tracker') renderTrackerTimeline();
}

export function renderScheduleSettings() {
    const container = document.getElementById('scheduleConfigList');
    container.innerHTML = '';
    state.appSchedule.forEach(block => addScheduleRow(block));
}

export function formatTimeForInput(decimalHour) {
    let normalized = decimalHour % 24;
    const hrs = Math.floor(normalized).toString().padStart(2, '0');
    const mins = Math.round((normalized % 1) * 60).toString().padStart(2, '0');
    return `${hrs}:${mins}`;
}

export function addScheduleRow(block = {title: '', startHour: 14, endHour: 29}) {
    const div = document.createElement('div');
    div.className = 'schedule-row';
    div.style.display = 'flex'; div.style.gap = '8px'; div.style.marginBottom = '8px';
    div.innerHTML = `
        <input type="text" class="sched-title" value="${block.title}" placeholder="Label" style="flex:2; padding:8px;">
        <input type="time" class="sched-start" value="${formatTimeForInput(block.startHour)}" style="padding:8px;">
        <input type="time" class="sched-end" value="${formatTimeForInput(block.endHour)}" style="padding:8px;">
        <button class="btn btn-outline" style="color:red; padding:8px 12px;" onclick="this.parentElement.remove()">×</button>
    `;
    document.getElementById('scheduleConfigList').appendChild(div);
}

export function clearCalendarCache() {
    if (confirm("Are you sure you want to clear the calendar event cache? This will allow re-importing all meetings when you click Sync Meetings.")) {
        state.notes = state.notes.filter(n => !n.eventId);
        saveNotes();
        handleSearch();
        showToast("Calendar event cache cleared!");
    }
}

export function exportData() { 
    const dataStr = JSON.stringify(state.notes, null, 2); 
    const blob = new Blob([dataStr], { type: "application/json" }); 
    const url = URL.createObjectURL(blob); 
    const a = document.createElement('a'); 
    a.href = url; 
    a.download = `quadra_backup_${new Date().toISOString().split('T')[0]}.json`; 
    a.click(); 
    URL.revokeObjectURL(url); 
}

export function triggerImport() { document.getElementById('importFile').click(); }

export function importData(event) { 
    const file = event.target.files[0]; 
    if (!file) return; 
    const reader = new FileReader(); 
    reader.onload = function(e) { 
        try { 
            const importedNotes = JSON.parse(e.target.result); 
            if (Array.isArray(importedNotes)) { 
                const noteMap = new Map(state.notes.map(n => [n.id, n])); 
                importedNotes.forEach(inNote => { inNote.dirty = true; noteMap.set(inNote.id, inNote); }); 
                state.notes = Array.from(noteMap.values()); 
                saveNotes(); 
                handleSearch(); 
                showToast("Tasks merged!"); 
                closeSettingsPage(); 
            } else showToast("Invalid format."); 
        } catch (err) { showToast("Error reading file."); } 
        event.target.value = ''; 
    }; 
    reader.readAsText(file); 
}

export function onTimesChanged() {
    const startVal = document.getElementById('taskStartTime').value;
    const endVal = document.getElementById('taskEndTime').value;
    if (startVal && endVal) {
        let startDec = timeStrToDecimal(startVal);
        let endDec = timeStrToDecimal(endVal);
        if (endDec <= startDec) endDec += 24;
        let duration = roundToQuarterHour(endDec - startDec);
        document.getElementById('taskHoursWorked').value = duration;
    }
}

export function onHoursChanged() {
    const hoursVal = parseFloat(document.getElementById('taskHoursWorked').value);
    const startVal = document.getElementById('taskStartTime').value;
    if (!isNaN(hoursVal) && startVal) {
        let startDec = timeStrToDecimal(startVal);
        let endDec = startDec + hoursVal;
        document.getElementById('taskEndTime').value = decimalToTimeStr(endDec);
    }
}

export function openTaskModal(quadrant = null, noteId = null, event = null, timelineContext = null) {
    if (event) event.stopPropagation();
    const modal = document.getElementById('taskModal');
    const titleInput = document.getElementById('taskTitleInput');
    const infoInput = document.getElementById('taskInfoInput');
    const dueDateInput = document.getElementById('taskDueDate');
    const selectedDateInput = document.getElementById('taskSelectedDate');
    const hoursWorkedInput = document.getElementById('taskHoursWorked');
    const startTimeInput = document.getElementById('taskStartTime');
    const endTimeInput = document.getElementById('taskEndTime');

    const trackerSection = document.getElementById('trackerInputSection');
    if (trackerSection) {
        trackerSection.style.display = (state.appConfig.viewsEnabled && state.appConfig.viewsEnabled.tracker === false) ? 'none' : 'block';
    }

    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0]; 

    if (noteId) {
        state.currentEditingId = noteId; 
        const note = state.notes.find(n => n.id === noteId); 
        document.getElementById('taskModalTitle').innerText = 'Task Details'; 

        let lines = note.text.split('\n');
        titleInput.value = lines[0] || '';
        infoInput.value = lines.slice(1).join('\n') || '';

        dueDateInput.value = note.dueDate ? note.dueDate.split('T')[0] : '';
        
        let activeBlock = (note.timeBlocks && note.timeBlocks.length > 0) ? note.timeBlocks[0] : null;
        selectedDateInput.value = activeBlock ? activeBlock.date : todayStr;
        hoursWorkedInput.value = activeBlock ? roundToQuarterHour(activeBlock.duration) : '';

        if (activeBlock && activeBlock.startHour !== undefined) {
            startTimeInput.value = decimalToTimeStr(activeBlock.startHour);
            endTimeInput.value = decimalToTimeStr(activeBlock.startHour + activeBlock.duration);
        } else {
            startTimeInput.value = '';
            endTimeInput.value = '';
        }

    } else { 
        state.currentEditingId = null; 
        state.currentAddingQuadrant = quadrant || 'inbox'; 
        document.getElementById('taskModalTitle').innerText = 'Add Task'; 
        titleInput.value = '';
        infoInput.value = '';
        dueDateInput.value = timelineContext ? timelineContext.date : '';
        selectedDateInput.value = timelineContext ? timelineContext.date : todayStr;
        hoursWorkedInput.value = timelineContext ? '1.0' : '';
        
        if (timelineContext && timelineContext.startHour !== undefined) {
            startTimeInput.value = decimalToTimeStr(timelineContext.startHour);
            endTimeInput.value = decimalToTimeStr(timelineContext.startHour + 1.0);
        } else {
            startTimeInput.value = '';
            endTimeInput.value = '';
        }

        state.pendingTimelineContext = timelineContext || null;
    }
    modal.style.display = 'flex'; 
    setTimeout(() => titleInput.focus(), 100);
}

export function closeTaskModal() { 
    document.getElementById('taskModal').style.display = 'none'; 
    state.pendingTimelineContext = null; 
}

export function saveTaskModal() {
    const titleText = document.getElementById('taskTitleInput').value.trim(); 
    const infoText = document.getElementById('taskInfoInput').value.trim(); 
    const dueDate = document.getElementById('taskDueDate').value;
    const selectedDate = document.getElementById('taskSelectedDate').value;
    let hoursVal = parseFloat(document.getElementById('taskHoursWorked').value);
    if (!isNaN(hoursVal)) hoursVal = roundToQuarterHour(hoursVal);
    const startVal = document.getElementById('taskStartTime').value;

    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0]; 

    if (!titleText && !infoText) return closeTaskModal();

    const fullText = titleText + (infoText ? ('\n' + infoText) : '');

    let calculatedStartHour = 9;
    if (startVal) {
        calculatedStartHour = roundToQuarterHour(timeStrToDecimal(startVal));
    } else if (state.pendingTimelineContext) {
        calculatedStartHour = roundToQuarterHour(state.pendingTimelineContext.startHour);
    }

    if (state.currentEditingId) { 
        const note = state.notes.find(n => n.id === state.currentEditingId); 
        if (note) { 
            note.text = fullText; 
            note.dueDate = dueDate || null;
            
            if (!isNaN(hoursVal) && hoursVal > 0) {
                let existingBlock = (note.timeBlocks && note.timeBlocks.length > 0) ? note.timeBlocks[0] : null;
                if (existingBlock) {
                    existingBlock.date = selectedDate || todayStr;
                    existingBlock.duration = hoursVal;
                    existingBlock.startHour = calculatedStartHour;
                } else {
                    if (!note.timeBlocks) note.timeBlocks = [];
                    note.timeBlocks.push({
                        id: Date.now().toString() + Math.random(),
                        date: selectedDate || todayStr,
                        startHour: calculatedStartHour,
                        duration: hoursVal
                    });
                }
            }

            note.dirty = true; 
            saveNotes(); 
            handleSearch(); 
        } 
    } else { 
        let newBlocks = [];
        let targetQuad = state.currentAddingQuadrant || 'inbox';
        
        if (!isNaN(hoursVal) && hoursVal > 0) {
            newBlocks.push({
                id: Date.now().toString() + Math.random(),
                date: selectedDate || (state.pendingTimelineContext ? state.pendingTimelineContext.date : todayStr),
                startHour: calculatedStartHour,
                duration: hoursVal
            });
        }

        state.notes.push({ 
            id: Date.now().toString(), 
            text: fullText, 
            quadrant: targetQuad, 
            status: 'active', 
            dueDate: dueDate || null,
            timeBlocks: newBlocks, 
            dirty: true, 
            deleted: false, 
            eventId: null 
        }); 
        saveNotes(); 
        handleSearch(); 
    }
    closeTaskModal();
}

export function openEditBlockModal(noteId, blockId, event) {
    if (event) event.stopPropagation();
    const note = state.notes.find(n => n.id === noteId);
    const block = note.timeBlocks.find(b => b.id === blockId);
    state.editingBlockState = { noteId, blockId };
    
    document.getElementById('editBlockTaskName').innerText = note.text.split('\n')[0];
    document.getElementById('editBlockHours').value = roundToQuarterHour(block.duration);
    document.getElementById('editTimeBlockModal').style.display = 'flex';
}

export function closeEditBlockModal() { 
    document.getElementById('editTimeBlockModal').style.display = 'none'; 
    state.editingBlockState = null;
}

export function saveTimeBlock() {
    if(!state.editingBlockState) return;
    const newHours = roundToQuarterHour(parseFloat(document.getElementById('editBlockHours').value));
    const note = state.notes.find(n => n.id === state.editingBlockState.noteId);
    const block = note.timeBlocks.find(b => b.id === state.editingBlockState.blockId);
    
    if (newHours > 0) { block.duration = newHours; note.dirty = true; saveNotes(); handleSearch(); }
    closeEditBlockModal();
}

export function deleteTimeBlock() {
    if(!state.editingBlockState) return;
    const note = state.notes.find(n => n.id === state.editingBlockState.noteId);
    note.timeBlocks = note.timeBlocks.filter(b => b.id !== state.editingBlockState.blockId);
    note.dirty = true; saveNotes(); handleSearch(); closeEditBlockModal();
}

export function openShortcutsModal() {
    const modal = document.getElementById('shortcutsModal');
    if (modal) modal.style.display = 'flex';
}

export function closeShortcutsModal() {
    const modal = document.getElementById('shortcutsModal');
    if (modal) modal.style.display = 'none';
}