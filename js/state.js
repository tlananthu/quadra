export const state = {
    appConfig: JSON.parse(localStorage.getItem('quadra_config')) || {},
    appSchedule: JSON.parse(localStorage.getItem('quadra_schedule')) || [
        { title: 'Out of office hours', startHour: 14, endHour: 29 }, 
        { title: 'Lunch', startHour: 13, endHour: 14 }
    ],
    notes: [],
    timelineZoom: parseFloat(localStorage.getItem('quadra_zoom')) || 1,
    currentLayout: localStorage.getItem('quadra_layout') || 'grid',
    currentTrackerMode: 'day',
    tokenClient: null,
    isGoogleSynced: false,
    clockIntervalId: null,
    dragState: null,
    currentEditingId: null,
    currentAddingQuadrant: null,
    pendingTimelineContext: null,
    editingBlockState: null,
    quadResizeTimeout: null
};

// Initialize configuration defaults
if (!state.appConfig.ignoreKeywords) state.appConfig.ignoreKeywords = 'out of office, ooo, away, vacation, holiday';
if (!state.appConfig.calSource) state.appConfig.calSource = 'google';
if (!state.appConfig.icsUrl) state.appConfig.icsUrl = '';
if (!state.appConfig.gcalEmbedUrl) state.appConfig.gcalEmbedUrl = '';
if (!state.appConfig.viewsEnabled) state.appConfig.viewsEnabled = { grid: true, kanban: true, overdue: true, gcal: true, tracker: true };
if (state.appConfig.viewsEnabled.gcal === undefined) state.appConfig.viewsEnabled.gcal = true;
if (!state.appConfig.defaultView) state.appConfig.defaultView = 'grid';
if (!state.appConfig.primaryTz) state.appConfig.primaryTz = 'local';
if (!state.appConfig.secondaryTz) state.appConfig.secondaryTz = 'none';
if (!state.appConfig.quadrantOrder) state.appConfig.quadrantOrder = ['q1', 'q2', 'q3', 'q4', 'tray-inbox', 'tray-calendar', 'tray-closed'];
if (!state.appConfig.quadrantWidths) state.appConfig.quadrantWidths = {};

// Initialize Notes State
let loadedNotes = JSON.parse(localStorage.getItem('quadra_notes')) || [];
state.notes = loadedNotes.map(note => {
    let blocks = note.timeBlocks || [];
    if (note.timeLogs && Object.keys(note.timeLogs).length > 0) {
        Object.keys(note.timeLogs).forEach(date => {
            blocks.push({ id: Date.now().toString() + Math.random(), date: date, startHour: 9, duration: Math.round((note.timeLogs[date]) * 4) / 4 });
        });
        delete note.timeLogs; 
    }
    blocks = blocks.map(b => ({ ...b, startHour: Math.round(b.startHour * 4) / 4, duration: Math.round(b.duration * 4) / 4 }));
    return { ...note, status: note.status || 'active', dirty: note.dirty || false, deleted: note.deleted || false, timeBlocks: blocks, eventId: note.eventId || null, quadrant: note.quadrant || 'q2' };
});

export function saveNotes() { 
    localStorage.setItem('quadra_notes', JSON.stringify(state.notes)); 
}

export function saveConfig() { 
    localStorage.setItem('quadra_config', JSON.stringify(state.appConfig)); 
}

export function saveSchedule() { 
    localStorage.setItem('quadra_schedule', JSON.stringify(state.appSchedule)); 
}

export function saveQuadrantState() {
    const container = document.getElementById('matrix');
    const quads = [...container.querySelectorAll('.quadrant')];
    state.appConfig.quadrantOrder = quads.map(q => q.id);
    saveConfig();
}