let version = '4.23';
let appConfig = JSON.parse(localStorage.getItem('quadra_config')) || {};
let isDocMode = false;
let tokenHeartbeatId = null;
let currentNotebookLayout = 'grid';

if (!appConfig.ignoreKeywords) appConfig.ignoreKeywords = 'out of office, ooo, away, vacation, holiday';
if (!appConfig.calSource) appConfig.calSource = 'google';
if (!appConfig.icsUrl) appConfig.icsUrl = '';
if (!appConfig.viewsEnabled) appConfig.viewsEnabled = { grid: true, kanban: true, overdue: true, tracker: true, notebook: true };
if (appConfig.viewsEnabled.notebook === undefined) appConfig.viewsEnabled.notebook = true;
if (!appConfig.defaultView) appConfig.defaultView = 'grid';
if (!appConfig.primaryTz) appConfig.primaryTz = 'local';
if (!appConfig.secondaryTz) appConfig.secondaryTz = 'none';
if (!appConfig.quadrantOrder) {
    appConfig.quadrantOrder = ['q1', 'q2', 'q3', 'q4', 'tray-inbox', 'tray-calendar', 'tray-closed'];
} else if (!appConfig.quadrantOrder.includes('notes')) {
    // Force inject 'notes' after 'tray-inbox' for existing saved layouts
    const inboxIdx = appConfig.quadrantOrder.indexOf('tray-inbox');
    if (inboxIdx !== -1) {
        appConfig.quadrantOrder.splice(inboxIdx + 1, 0, 'notes');
    } else {
        appConfig.quadrantOrder.push('notes');
    }
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
}
if (!appConfig.oooDates) {
    appConfig.oooDates = [];
}
if (!appConfig.quadrantWidths) appConfig.quadrantWidths = {};

if (!appConfig.projects || appConfig.projects.length === 0) {
    appConfig.projects = [{ id: 'p_default', name: 'Default', visible: true }];
}

let tokenClient;
let isGoogleSynced = false;
let clockIntervalId = null;
let timelineZoom = parseFloat(localStorage.getItem('quadra_zoom')) || 1;
let autoSaveTimerId = null;
let currentEditingId = null; 
let currentAddingQuadrant = null;
let pendingTimelineContext = null;
let dragState = null;
let isDraggingBlock = false;

// --- SQLite Database Engine ---
let db;
let SQL;
let driveFileId = null; // Will store the Google Drive file ID

async function initSQLite(binaryData = null) {
    if (!SQL) {
        SQL = await initSqlJs({
            // Fetch the WebAssembly file from the CDN
            locateFile: file => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${file}`
        });
    }
    
    if (binaryData) {
        // Load existing database from Google Drive
        db = new SQL.Database(new Uint8Array(binaryData));
        console.log("Loaded existing SQLite database from Google Drive.");
    } else {
        // Create a fresh database and schema
        db = new SQL.Database();
        db.run(`
            CREATE TABLE IF NOT EXISTS tasks (
                id TEXT PRIMARY KEY,
                text TEXT,
                quadrant TEXT,
                status TEXT,
                dueDate TEXT,
                timeBlocks TEXT,
                deleted INTEGER DEFAULT 0,
                projectId TEXT DEFAULT 'p_default'
            );
        `);
        try { db.run(`ALTER TABLE tasks ADD COLUMN projectId TEXT DEFAULT 'p_default'`); } catch (e) {}
        console.log("Created fresh SQLite database in memory.");
    }
}

function syncNotesToSQLite() {
    if (!db) return;

    // Ensure the table exists
    db.run(`
        CREATE TABLE IF NOT EXISTS tasks (
            id TEXT PRIMARY KEY,
            text TEXT,
            quadrant TEXT,
            status TEXT,
            dueDate TEXT,
            timeBlocks TEXT,
            deleted INTEGER DEFAULT 0,
            projectId TEXT DEFAULT 'p_default'
        );
    `);
    try { db.run(`ALTER TABLE tasks ADD COLUMN projectId TEXT DEFAULT 'p_default'`); } catch (e) {}

    // Prepare a statement to insert or replace task records
    const stmt = db.prepare(`
        INSERT OR REPLACE INTO tasks (id, text, quadrant, status, dueDate, timeBlocks, deleted, projectId)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?);
    `);

    notes.forEach(note => {
        stmt.run([
            note.id.toString(),
            note.text || '',
            note.quadrant || 'inbox',
            note.status || 'active',
            note.dueDate || null,
            JSON.stringify(note.timeBlocks || []),
            note.deleted ? 1 : 0,
            JSON.stringify(note.projectIds || [note.projectId || 'p_default'])
        ]);
    });

    stmt.free();
}

// --- Google Drive AppData Sync ---
async function downloadDatabaseFromDrive() {
    // --- FIXED: True Promise-based await for the anti-race condition ---
    if (typeof gapi === 'undefined' || !gapi.client || !gapi.client.drive) {
        console.log("Waiting for Google Drive API to initialize...");
        await new Promise(resolve => setTimeout(resolve, 500));
        return await downloadDatabaseFromDrive(); // Recursively try again and block execution
    }

    try {
        // 1. Search the user's visible Drive for the database file
        const response = await gapi.client.drive.files.list({
            q: "name='quadra.sqlite' and trashed=false", 
            fields: 'files(id, name)',
            orderBy: 'createdTime desc' // Always grab the newest one if duplicates exist
        });
        
        const files = response.result.files;
        if (files && files.length > 0) {
            driveFileId = files[0].id;
            
            // 2. If found, download the binary contents
            const fileRes = await fetch(`https://www.googleapis.com/drive/v3/files/${driveFileId}?alt=media`, {
                headers: { 'Authorization': `Bearer ${gapi.client.getToken().access_token}` }
            });
            const arrayBuffer = await fileRes.arrayBuffer();
            
            // 3. Boot SQLite with the downloaded data
            await initSQLite(arrayBuffer);
            setCloudSyncIcon('saved');
        } else {
            // No file exists yet in Drive, boot a fresh database
            await initSQLite(null);
        }
    } catch (e) {
        console.error("Failed to load DB from Drive:", e);
        await initSQLite(null); // Fallback to fresh DB
    }
}

async function uploadDatabaseToDrive() {
    if (!isGoogleSynced) {
        setCloudSyncIcon('error');
        return;
    }
    
    // 1. Change UI to "Saving" state immediately
    setCloudSyncIcon('saving');
    
    if (!db) {
        await initSQLite(null);
    }
    syncNotesToSQLite();

    try {
        const binaryData = db.export();
        const blob = new Blob([binaryData], { type: 'application/x-sqlite3' });
        const token = gapi.client.getToken().access_token;
        
        // --- FIX: Prevent duplicates by double-checking Drive if driveFileId is missing ---
        if (!driveFileId && gapi.client.drive) {
            const searchRes = await gapi.client.drive.files.list({
                q: "name='quadra.sqlite' and trashed=false",
                fields: 'files(id, name)',
                orderBy: 'createdTime desc'
            });
            if (searchRes.result.files && searchRes.result.files.length > 0) {
                driveFileId = searchRes.result.files[0].id;
            }
        }
        
        let url;
        let method;
        let metadata = { name: 'quadra.sqlite' };

        if (driveFileId) {
            url = `https://www.googleapis.com/upload/drive/v3/files/${driveFileId}?uploadType=multipart`;
            method = 'PATCH';
        } else {
            url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
            method = 'POST';
        }
        
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
        form.append('file', blob);
        
        const res = await fetch(url, {
            method: method,
            headers: { 'Authorization': `Bearer ${token}` },
            body: form
        });
        
        if (!res.ok) {
            const errData = await res.json();
            throw new Error(errData.error?.message || "Upload request failed");
        }
        
        const result = await res.json();
        if (result.id) driveFileId = result.id;
        
        // 2. Change UI to "Saved" state on success
        setCloudSyncIcon('saved');
        
    } catch (e) {
        console.error("Failed to upload DB to Drive:", e);
        // 3. Change UI to "Error" state on failure
        setCloudSyncIcon('error');
    }
}

function toggleDocMode() {
    const modalContent = document.querySelector('#taskModal .modal-content');
    const toggleBtn = document.getElementById('docModeToggleBtn');
    
    isDocMode = !isDocMode;
    
    if (isDocMode) {
        modalContent.classList.add('doc-mode');
        toggleBtn.innerText = '🗗'; // Window restore icon
        toggleBtn.title = "Exit Doc Mode";
        
        // Focus the main body text automatically when entering Doc Mode
        setTimeout(() => document.getElementById('taskInfoInput').focus(), 250);
    } else {
        modalContent.classList.remove('doc-mode');
        toggleBtn.innerText = '⛶'; // Maximize icon
        toggleBtn.title = "Enter Doc Mode";
    }
}

const SCOPES = 'https://www.googleapis.com/auth/tasks https://www.googleapis.com/auth/calendar https://www.googleapis.com/auth/drive.file';

const defaultSchedule = [
    { title: 'Out of office hours', startHour: 14, endHour: 29 }, 
    { title: 'Lunch', startHour: 13, endHour: 14 }
];
let appSchedule = JSON.parse(localStorage.getItem('quadra_schedule')) || defaultSchedule;

// Data Migration
// --- Data Migration (Upgrade to Multi-Day Architecture) ---
let notes = JSON.parse(localStorage.getItem('quadra_notes')) || [];

notes = notes.map(note => {
    // 1. Initialize the new timeBlocks array if it doesn't exist
    if (!note.timeBlocks) {
        note.timeBlocks = [];
    }

    // 2. Safely port existing scheduled times into the new array format
    if (note.dueDate && note.dueTime !== undefined && note.dueDuration !== undefined) {
        
        // Prevent duplicates if the migration runs multiple times
        const alreadyMigrated = note.timeBlocks.some(b => b.date === note.dueDate && b.startHour === note.dueTime);
        
        if (!alreadyMigrated) {
            note.timeBlocks.push({
                blockId: 'b_' + Date.now().toString() + Math.floor(Math.random() * 1000),
                date: note.dueDate,
                startHour: note.dueTime,
                duration: note.dueDuration
            });
        }
        
        // 3. Clean up the outdated root-level scheduling variables
        delete note.dueTime;
        delete note.dueDuration;
    }

    // Ensure standard properties exist
    if (note.syncFailed === undefined) note.syncFailed = false;
    delete note.timeLogs; // Remove legacy unused property
    
    return { 
        ...note, 
        status: note.status || 'active', 
        dirty: note.dirty || false, 
        deleted: note.deleted || false, 
        eventId: note.eventId || null, 
        quadrant: note.quadrant || 'q2' 
    };
});

const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0]; 
let savedDate = localStorage.getItem('quadra_tracker_date');
document.getElementById('trackerDate').value = savedDate || todayStr;

document.getElementById('taskTitleInput')?.addEventListener('keydown', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
    }
});

// --- Saving Quadrant Order and Width States ---
function saveQuadrantState() {
    const container = document.getElementById('matrix');
    const quads = [...container.querySelectorAll('.quadrant')];
    appConfig.quadrantOrder = quads.map(q => q.id);
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
}

let quadResizeTimeout;
const quadResizeObserver = new ResizeObserver(entries => {
    clearTimeout(quadResizeTimeout);
    quadResizeTimeout = setTimeout(() => {
        let changed = false;
        entries.forEach(entry => {
            const el = entry.target;
            if (el.style.width && appConfig.quadrantWidths[el.id] !== el.style.width) {
                appConfig.quadrantWidths[el.id] = el.style.width;
                changed = true;
            }
        });
        if (changed) {
            localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        }
    }, 500);
});

// --- Edge Splitter Pull Resizing Engine ---
function initQuadResize(event, quadId) {
    event.preventDefault();
    event.stopPropagation();
    const quadEl = document.getElementById(quadId);
    if (!quadEl) return;
    const startX = event.clientX;
    const startWidth = quadEl.getBoundingClientRect().width;
    document.body.style.cursor = 'col-resize';
    function onMouseMove(e) {
        const deltaX = e.clientX - startX;
        const newWidth = Math.max(260, Math.min(1000, Math.round(startWidth + deltaX)));
        quadEl.style.width = newWidth + 'px';
        appConfig.quadrantWidths[quadId] = newWidth + 'px';
    }
    function onMouseUp() {
        document.body.style.cursor = '';
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
    }
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
}

function formatCurrentTimeBadge(date) {
    let primaryOpts = { hour: '2-digit', minute:'2-digit' };
    if (appConfig.primaryTz !== 'local') primaryOpts.timeZone = appConfig.primaryTz;
    let text = date.toLocaleTimeString([], primaryOpts);

    if (appConfig.secondaryTz && appConfig.secondaryTz !== 'none') {
        let secOpts = { hour: '2-digit', minute:'2-digit', timeZoneName: 'short' };
        secOpts.timeZone = appConfig.secondaryTz;
        let secText = date.toLocaleTimeString([], secOpts);
        text += ` (${secText})`;
    }
    return text;
}

// --- Layout Management ---
let currentLayout = localStorage.getItem('quadra_layout') || appConfig.defaultView;
let currentTrackerMode = 'day'; 

if (!appConfig.viewsEnabled[currentLayout]) {
    const firstEnabled = Object.keys(appConfig.viewsEnabled).find(k => appConfig.viewsEnabled[k]);
    currentLayout = firstEnabled || 'grid';
}

function applyViewVisibility() {
    const btnGrid = document.getElementById('btnGrid') || document.getElementById('btn-layout-grid');
    const btnKanban = document.getElementById('btnKanban') || document.getElementById('btn-layout-kanban');
    const btnTracker = document.getElementById('btnTracker') || document.getElementById('btn-layout-tracker');
    const btnOverdue = document.getElementById('btnOverdue') || document.getElementById('btn-layout-overdue');
    const btnNotebook = document.getElementById('btnNotebook') || document.getElementById('btn-layout-notebook');
    
    if (btnGrid) btnGrid.style.display = appConfig.viewsEnabled.grid ? '' : 'none';
    if (btnKanban) btnKanban.style.display = appConfig.viewsEnabled.kanban ? '' : 'none';
    if (btnTracker) btnTracker.style.display = appConfig.viewsEnabled.tracker ? '' : 'none';
    if (btnOverdue) btnOverdue.style.display = appConfig.viewsEnabled.overdue ? '' : 'none';
    if (btnNotebook) btnNotebook.style.display = appConfig.viewsEnabled.notebook ? '' : 'none';
    
    if (!appConfig.viewsEnabled[currentLayout]) {
        const firstEnabled = Object.keys(appConfig.viewsEnabled).find(k => appConfig.viewsEnabled[k]);
        setLayout(firstEnabled || 'grid');
    }
}

function setLayout(layout) {
    if (!appConfig.viewsEnabled[layout]) return; 
    
    currentLayout = layout;
    localStorage.setItem('quadra_layout', layout);
    
    const matrix = document.getElementById('matrix');
    const tracker = document.getElementById('tracker-view');
    const overdue = document.getElementById('overdue-view');
    const settingsView = document.getElementById('settings-view');
    const notebookView = document.getElementById('notebook-view'); 
    
    const btnGrid = document.getElementById('btnGrid') || document.getElementById('btn-layout-grid');
    const btnKanban = document.getElementById('btnKanban') || document.getElementById('btn-layout-kanban');
    const btnOverdue = document.getElementById('btnOverdue') || document.getElementById('btn-layout-overdue');
    const btnTracker = document.getElementById('btnTracker') || document.getElementById('btn-layout-tracker');
    const btnNotebook = document.getElementById('btnNotebook') || document.getElementById('btn-layout-notebook');

    if (btnGrid) btnGrid.classList.toggle('active', layout === 'grid');
    if (btnKanban) btnKanban.classList.toggle('active', layout === 'kanban');
    if (btnOverdue) btnOverdue.classList.toggle('active', layout === 'overdue');
    if (btnTracker) btnTracker.classList.toggle('active', layout === 'tracker');
    if (btnNotebook) btnNotebook.classList.toggle('active', layout === 'notebook');
    
    // ENSURE ALL VIEWS ARE HIDDEN FIRST
    if(settingsView) settingsView.style.display = 'none';
    if(matrix) matrix.style.display = 'none';
    if(tracker) tracker.style.display = 'none';
    if(overdue) overdue.style.display = 'none';
    if(notebookView) notebookView.style.display = 'none'; 

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
    } else if (layout === 'notebook') {
        if(notebookView) notebookView.style.display = 'flex';
        document.body.classList.remove('sidebar-open');
        renderNotebookView();
        stopLiveClock();
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

function switchLayout(layout) {
    setLayout(layout);
}

function adjustTimelineZoom(amount) {
    timelineZoom = Math.max(0.5, Math.min(3, roundToQuarterHour(timelineZoom + amount)));
    localStorage.setItem('quadra_zoom', timelineZoom);
    updateZoomDisplay();
    renderTrackerTimeline();
}

function updateZoomDisplay() {
    const display = document.getElementById('zoomLevelDisplay');
    if (display) display.innerText = `${timelineZoom}x`;
}

function openSettingsPage() {
    document.getElementById('matrix').style.display = 'none';
    document.getElementById('tracker-view').style.display = 'none';
    document.getElementById('overdue-view').style.display = 'none';
    document.getElementById('settings-view').style.display = 'block';
    document.body.classList.remove('sidebar-open');
    document.getElementById('quick-tags-bar').style.display = 'none';
    document.getElementById('searchHeaderContainer').style.display = 'none';
    document.getElementById('viewToggleGroup').style.display = 'none';
    document.getElementById('settingsNavBtn').style.display = 'none';
    document.getElementById('backNavBtn').style.display = 'inline-block';

    document.getElementById('configClientId').value = appConfig.clientId || '';
    document.getElementById('configApiKey').value = appConfig.apiKey || '';
    document.getElementById('configTimesheetUrl').value = appConfig.timesheetUrl || '';
    
    // Load toggles
    if (appConfig.importBehavior === 'palette') {
        document.getElementById('importPalette').checked = true;
    } else {
        document.getElementById('importAuto').checked = true;
    }

    const ignoreEl = document.getElementById('configIgnoreKeywords');
    if (ignoreEl) ignoreEl.value = appConfig.ignoreKeywords || '';
    
    const calSourceEl = document.getElementById('configCalSource');
    if (calSourceEl) calSourceEl.value = appConfig.calSource || 'google';

    document.getElementById('configPrimaryTz').value = appConfig.primaryTz || 'local';
    document.getElementById('configSecondaryTz').value = appConfig.secondaryTz || 'none';
    
    document.getElementById('configDefaultView').value = appConfig.defaultView || 'grid';
    document.getElementById('configViewGrid').checked = appConfig.viewsEnabled.grid !== false;
    document.getElementById('configViewKanban').checked = appConfig.viewsEnabled.kanban !== false;
    document.getElementById('configViewTracker').checked = appConfig.viewsEnabled.tracker !== false;
    document.getElementById('configViewOverdue').checked = appConfig.viewsEnabled.overdue !== false;
    document.getElementById('configViewNotebook').checked = appConfig.viewsEnabled.notebook !== false;
    
    loadCalendars();
    renderScheduleSettings();
    renderArchivedProjects();
}

function toggleCalSourceFields(source) {
    if (source === 'outlook') {
        document.getElementById('googleCalGroup').style.display = 'none';
        document.getElementById('outlookIcsGroup').style.display = 'block';
    } else {
        document.getElementById('googleCalGroup').style.display = 'block';
        document.getElementById('outlookIcsGroup').style.display = 'none';
    }
}

function closeSettingsPage() {
    setLayout(currentLayout);
}

function setTrackerMode(mode) {
    currentTrackerMode = mode;
    document.getElementById('btnTrackerDay').classList.toggle('active', mode === 'day');
    document.getElementById('btnTrackerWeek').classList.toggle('active', mode === 'week');
    renderTrackerTimeline();
}

function saveNotes() { 
    localStorage.setItem('quadra_notes', JSON.stringify(notes)); 
    setCloudSyncIcon('unsaved');
}

function startLiveClock() {
    if(clockIntervalId) clearInterval(clockIntervalId);
    clockIntervalId = setInterval(() => {
        if (currentLayout === 'tracker' && currentTrackerMode === 'day' && document.getElementById('tracker-view').style.display === 'block') {
            const line = document.querySelector('.current-time-line');
            const badge = document.querySelector('.current-time-badge');
            if (line && badge) {
                const hourPx = 60 * timelineZoom;
                const now = new Date();
                const primaryTime = getTzTime(now, appConfig.primaryTz);
                const currentHour = primaryTime.h + (primaryTime.m / 60) + (primaryTime.s / 3600);
                
                line.style.top = `${currentHour * hourPx}px`;
                badge.innerText = formatCurrentTimeBadge(now);
            }
        }
    }, 30000); 
}

function stopLiveClock() {
    if(clockIntervalId) { clearInterval(clockIntervalId); clockIntervalId = null; }
}

// --- Drag/Drop Quadrants (Columns) Logic ---
function dragStartQuad(e) {
    if (e.target.classList.contains('quadrant')) {
        e.dataTransfer.setData('quadrant_id', e.target.id);
        e.dataTransfer.effectAllowed = 'move';
        setTimeout(() => e.target.style.opacity = '0.5', 0);
    }
}

function dragEndQuad(e) {
    if (e.target.classList.contains('quadrant')) {
        e.target.style.opacity = '1';
        document.querySelectorAll('.quadrant').forEach(q => q.classList.remove('quad-drag-over', 'drag-over'));
    }
}

function allowDropQuad(e) {
    e.preventDefault();
    if(e.dataTransfer.types.includes('quadrant_id')) {
        e.currentTarget.classList.add('quad-drag-over');
    } else {
        e.currentTarget.classList.add('drag-over');
    }
}

function dragLeaveQuad(e) {
    e.currentTarget.classList.remove('quad-drag-over', 'drag-over');
}

function dropQuad(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('quad-drag-over', 'drag-over');
    
    const quadId = e.dataTransfer.getData('quadrant_id');
    if (quadId) {
        const draggedQuad = document.getElementById(quadId);
        const targetQuad = e.currentTarget;
        if (draggedQuad && targetQuad && draggedQuad !== targetQuad) {
            const container = document.getElementById('matrix');
            const allQuads = [...container.querySelectorAll('.quadrant')];
            const draggedIdx = allQuads.indexOf(draggedQuad);
            const targetIdx = allQuads.indexOf(targetQuad);
            
            if (draggedIdx < targetIdx) {
                targetQuad.parentNode.insertBefore(draggedQuad, targetQuad.nextSibling);
            } else {
                targetQuad.parentNode.insertBefore(draggedQuad, targetQuad);
            }
            saveQuadrantState();
        }
    } else {
        const noteId = e.dataTransfer.getData("text/plain");
        if(!noteId) return;
        const note = notes.find(n => n.id === noteId);
        let targetKey = e.currentTarget.id; 
        if(targetKey.startsWith('tray-')) targetKey = targetKey.replace('tray-', '');
        
        if (note && note.status === 'active' && note.quadrant !== targetKey) { 
            if (targetKey === 'closed') note.status = 'closed';
            
            if (targetKey === 'notes' && !note.text.includes('#note')) {
                note.text += ' #note';
            }
            
            note.quadrant = targetKey; 
            note.dirty = true; 
            saveNotes();
            syncSingleTask(note.id);
            handleSearch(); 
        }
    }
}

function changeTrackerDay(offset) {
    const dateInput = document.getElementById('trackerDate');
    
    // Split to explicitly force local time instead of UTC midnight
    const [y, m, d] = dateInput.value.split('-'); 
    const dateObj = new Date(y, m - 1, d);
    
    dateObj.setDate(dateObj.getDate() + offset);
    
    const localY = dateObj.getFullYear();
    const localM = String(dateObj.getMonth() + 1).padStart(2, '0');
    const localD = String(dateObj.getDate()).padStart(2, '0');
    
    dateInput.value = `${localY}-${localM}-${localD}`;
    renderTrackerTimeline();
}

function goToToday() {
    const dateInput = document.getElementById('trackerDate');
    dateInput.value = new Date().toLocaleDateString('en-CA').split('T')[0];
    renderTrackerTimeline();
}

function renderTrackerPalette() {
    const globalSearchInput = document.getElementById('searchInput');
    const globalQuery = globalSearchInput ? globalSearchInput.value : '';
    
    const paletteSearchInput = document.getElementById('paletteSearchInput');
    const paletteSearchText = paletteSearchInput ? paletteSearchInput.value : '';
    
    const clearPaletteBtn = document.getElementById('clearPaletteSearchBtn');
    if (clearPaletteBtn) {
        clearPaletteBtn.style.display = paletteSearchText.trim().length > 0 ? 'block' : 'none';
    }
    
    const effectivePaletteQuery = paletteSearchText.trim().length > 0 ? paletteSearchText : globalQuery;

    const paletteList = document.getElementById('tracker-palette-list');
    if (!paletteList) return;
    
    paletteList.innerHTML = '';
    
    const trackerDate = document.getElementById('trackerDate') ? document.getElementById('trackerDate').value : new Date().toLocaleDateString('en-CA').split('T')[0];
    const dueToggle = document.getElementById('dueFilterToggle');
    const isDueFilterOn = dueToggle && dueToggle.checked;

    let paletteNotes = notes.filter(n => !n.deleted && n.status !== 'closed' && matchesSearchQuery(n.text, effectivePaletteQuery) && !n.eventId && isProjectVisible(n));
    
    // 1. If Global Due is ON: Show ONLY tasks due on the selected calendar day
    if (isDueFilterOn) {
        paletteNotes = paletteNotes.filter(n => n.dueDate === trackerDate);
    }

    const quadPriority = { 'q1': 1, 'q2': 2, 'q3': 3, 'q4': 4, 'inbox': 5, 'calendar': 6 };
    
    paletteNotes.sort((a, b) => {
        const aIsDueSelectedDay = a.dueDate === trackerDate;
        const bIsDueSelectedDay = b.dueDate === trackerDate;

        // 2. If Global Due is OFF: Force tasks due on the selected day to the very top
        if (aIsDueSelectedDay && !bIsDueSelectedDay) return -1;
        if (!aIsDueSelectedDay && bIsDueSelectedDay) return 1;

        // Standard sorting for the rest
        if (a.dueDate && b.dueDate) {
            const dateCompare = a.dueDate.localeCompare(b.dueDate);
            if (dateCompare !== 0) return dateCompare;
        } 
        else if (a.dueDate && !b.dueDate) return -1;
        else if (!a.dueDate && b.dueDate) return 1;
        
        const pA = quadPriority[a.quadrant] || 99;
        const pB = quadPriority[b.quadrant] || 99;
        return pA - pB;
    });

    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];

    const quadStyles = {
        'q1': { color: 'var(--q1-text)', border: 'var(--q1-border)', bg: 'var(--q1-bg)', label: 'Q1 (Urgent)' },
        'q2': { color: 'var(--q2-text)', border: 'var(--q2-border)', bg: 'var(--q2-bg)', label: 'Q2 (Schedule)' },
        'q3': { color: 'var(--q3-text)', border: 'var(--q3-border)', bg: 'var(--q3-bg)', label: 'Q3 (Delegate)' },
        'q4': { color: 'var(--q4-text)', border: 'var(--q4-border)', bg: 'var(--q4-bg)', label: 'Q4 (Later)' },
        'inbox': { color: 'var(--text-muted)', border: 'var(--border-color)', bg: '#F1F5F9', label: 'Inbox' },
        'calendar': { color: 'var(--cal-text)', border: 'var(--cal-border)', bg: 'var(--cal-bg)', label: 'Calendar' }
    };

    paletteNotes.forEach(note => {
        const el = document.createElement('div');
        el.className = 'note'; 
        el.style.marginBottom = '8px'; 
        el.style.cursor = 'grab';
        el.draggable = true;
        el.ondragstart = (e) => e.dataTransfer.setData('text/plain', note.id);
        
        const qStyle = quadStyles[note.quadrant] || { color: 'var(--text-muted)', border: 'var(--border-color)', bg: '#F1F5F9', label: note.quadrant };

        // --- NEW: Apply a thick left border matching the CSS variable ---
        el.style.borderLeft = `4px solid ${qStyle.border}`;

        const isPlannedOnCalendar = note.timeBlocks && note.timeBlocks.some(block => block.date >= todayStr);
        
        if (isPlannedOnCalendar) {
            el.style.backgroundColor = '#F8FAFC';
            el.style.borderTop = '1px solid var(--border-color)';
            el.style.borderRight = '1px solid var(--border-color)';
            el.style.borderBottom = '1px solid var(--border-color)';
        }
        
        const contentWrapper = document.createElement('div');
        contentWrapper.className = 'note-content-wrapper';
        contentWrapper.style.maxWidth = '100%'; 
        
        let overdueIndicator = '';
        if (note.dueDate) {
            const dueDateStr = note.dueDate.split('T')[0];
            if (dueDateStr < todayStr) {
                overdueIndicator = `<span style="background: #FFF1F2; color: #9F1239; border: 1px solid #FECDD3; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-right: 6px;">OVERDUE</span>`;
            }
        }

        let cleanTextTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
        contentWrapper.innerHTML = `<div class="note-text">${overdueIndicator}${parseTags(cleanTextTitle)}</div>`;
        
        // --- NEW: Inject the Quadrant Badge utilizing the full CSS theme ---
        let metaHTML = `<div style="font-size:11px; margin-top:10px; font-weight:600; display: flex; align-items: center; flex-wrap: wrap; gap: 8px;">`;
        
        metaHTML += `<span style="color: ${qStyle.color}; background-color: ${qStyle.bg}; border: 1px solid ${qStyle.border}; padding: 2px 6px; border-radius: 4px;">${qStyle.label}</span>`;

        if (note.dueDate) {
            metaHTML += `<span style="color: var(--text-muted); display: flex; align-items: center; gap: 4px;">🗓️ ${note.dueDate.split('T')[0]}</span>`;
        }
        
        if (isPlannedOnCalendar) {
            metaHTML += `<span style="margin-left: auto; background: var(--q2-bg); color: var(--q2-text); border: 1px solid var(--q2-border); padding: 2px 6px; border-radius: 4px; font-size: 9px; font-weight: 800; letter-spacing: 0.5px;">🕒 PLANNED</span>`;
        }

        metaHTML += `</div>`;
        contentWrapper.innerHTML += metaHTML;
        
        contentWrapper.onclick = (e) => openTaskModal(null, note.id, e);
        
        el.appendChild(contentWrapper);
        paletteList.appendChild(el);
    });
}

function renderTrackerTimeline() {
    const searchInput = document.getElementById('searchInput');
    const globalQuery = searchInput ? searchInput.value.toLowerCase() : '';
    
    renderTrackerPalette();

    const canvas = document.getElementById('timelineCanvas');
    const hourPx = 60 * timelineZoom;
    
    canvas.innerHTML = '';
    canvas.style.height = `${24 * hourPx}px`;

    const bgLines = document.createElement('div');
    bgLines.className = 'timeline-bg-lines';
    
    let hasSecTz = appConfig.secondaryTz && appConfig.secondaryTz !== 'none';
    let secOffsetDiff = 0;
    if (hasSecTz) {
        let primOff = getTzOffset(appConfig.primaryTz);
        let secOff = getTzOffset(appConfig.secondaryTz);
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
        
        row.innerHTML = `<span class="time-row-label">${labelText}</span>`;
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
    if (currentTrackerMode === 'day') {
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

    datesToRender.forEach(dtObj => {
        const dateStr = dtObj.date;
        const col = document.createElement('div');
        col.className = 'time-col';
        col.ondragover = allowTrackerDrop;
        col.ondragleave = dragLeaveTracker;
        col.ondrop = (e) => dropToTracker(e, dateStr);
        col.onclick = (e) => handleTimelineClick(e, dateStr);

        if (currentTrackerMode === 'week') {
            const header = document.createElement('div');
            header.className = 'col-header';
            if (dateStr === todayStr) header.classList.add('today');
            header.innerText = dtObj.label;
            col.appendChild(header);
        }

        if (currentTrackerMode === 'day' && dateStr === todayStr) {
            const now = new Date();
            const primaryTime = getTzTime(now, appConfig.primaryTz);
            const currentHour = primaryTime.h + (primaryTime.m / 60) + (primaryTime.s / 3600);
            
            const timeLine = document.createElement('div');
            timeLine.className = 'current-time-line';
            timeLine.style.top = `${currentHour * hourPx}px`;
            timeLine.innerHTML = `<span class="current-time-badge">${formatCurrentTimeBadge(now)}</span>`;
            col.appendChild(timeLine);
        }

        appSchedule.forEach(block => {
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
                if (currentTrackerMode === 'day') overlay.innerText = block.title;
                col.appendChild(overlay);
            }

            appSchedule.forEach(prevBlock => {
                if (prevBlock.endHour > 24 && dtObj.date === dateStr) {
                    let wrappedSpan = roundToQuarterHour(prevBlock.endHour - 24);
                    if (wrappedSpan > 0) {
                        const overlay = document.createElement('div');
                        overlay.className = 'schedule-overlay';
                        overlay.style.top = `0px`;
                        overlay.style.height = `${Math.min(wrappedSpan, 24) * hourPx}px`;
                        overlay.style.left = `-${timeGutterWidth}px`;
                        overlay.style.paddingLeft = `${overlayPadding}px`;
                        if (currentTrackerMode === 'day') overlay.innerText = prevBlock.title;
                        col.appendChild(overlay);
                    }
                }
            });
        });

        let dayBlocks = [];

        notes.forEach(note => {
            if (note.deleted) return; 
            
            //if (!isProjectVisible(note)) return; 
                        
            //if (!matchesSearchQuery(note.text, globalQuery)) return;
            
            const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;

            let blocksToProcess = note.timeBlocks || [];
            
            // Dynamic fallback for Imported Google Events
            if (isCalendarEvent && note.dueTime !== undefined) {
                blocksToProcess = [{ blockId: 'cal', date: note.dueDate, startHour: note.dueTime, duration: note.dueDuration }];
            }

            blocksToProcess.forEach(tBlock => {
                let blockStart = roundToQuarterHour(tBlock.startHour);
                let duration = roundToQuarterHour(tBlock.duration);
                let blockEnd = blockStart + duration;
                if (blockEnd < blockStart) { blockEnd += 24; }
                let actualDuration = roundToQuarterHour(blockEnd - blockStart);

                if (tBlock.date === dateStr) {
                    let renderStart = blockStart % 24;
                    let renderDuration = duration;

                    if (blockStart < 24 && blockEnd > 24) {
                        renderDuration = 24 - blockStart; 
                    } else if (blockStart >= 24) {
                        return; 
                    }

                    const blockEl = document.createElement('div');
                    const quadClass = note.quadrant || 'q2';
                    blockEl.className = 'logged-block' + (isCalendarEvent ? ' is-meeting' : ` ${quadClass}`) + (note.status === 'closed' ? ' is-closed' : '');
                    
                    // NEW: ID includes the specific blockId
                    blockEl.id = `block-${note.id}-${tBlock.blockId}`;
                    blockEl.style.top = `${renderStart * hourPx}px`;
                    blockEl.style.height = `${Math.max(15, renderDuration * hourPx)}px`;
                    
                    let pid = note.projectId || note.projectIds?.[0] || 'p_default';
                    let pObj = appConfig.projects.find(p => p.id === pid);
                    let pName = pObj ? `${pObj.name} - ` : '';
                    let cleanTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
                    let displayTitle = (pName + cleanTitle).substring(0, 45);
                    const actualEndHour = (blockStart + actualDuration) % 24;
                    const timeStr = `${decToTime(blockStart)} - ${decToTime(actualEndHour)}`;

                    // NEW: Pass both note.id and tBlock.blockId to startBlockDrag
                    blockEl.innerHTML = `
                        <div class="block-info">
                            <div class="block-title">${displayTitle}</div>
                            <div class="block-meta">${timeStr}</div>
                        </div>
                        ${isCalendarEvent ? '' : `<div class="resize-handle" onmousedown="startBlockDrag(event, '${note.id}', '${tBlock.blockId}', true)"></div>`}
                    `;
                    
                    blockEl.onclick = (e) => {
                        if(isDraggingBlock) return;
                        if(e.target.closest('.resize-handle')) return;
                        openTaskModal(null, note.id, e);
                    };

                    if(!isCalendarEvent) {
                        blockEl.onmousedown = (e) => {
                            if(e.target.closest('.resize-handle')) return;
                            startBlockDrag(e, note.id, tBlock.blockId, false);
                        };
                    }
                    
                    dayBlocks.push({ el: blockEl, start: renderStart, end: renderStart + renderDuration, duration: renderDuration });
                } 
                else if (tBlock.date === addDays(dateStr, -1)) {
                    if (blockStart < 24 && blockEnd > 24) {
                        let overflowDuration = roundToQuarterHour(blockEnd - 24);
                        const blockEl = document.createElement('div');
                        const quadClass = note.quadrant || 'q2';
                        blockEl.className = 'logged-block' + (isCalendarEvent ? ' is-meeting' : ` ${quadClass}`) + (note.status === 'closed' ? ' is-closed' : '');
                        blockEl.id = `block-overflow-${note.id}-${tBlock.blockId}`;
                        blockEl.style.top = `0px`;
                        blockEl.style.height = `${Math.max(15, overflowDuration * hourPx)}px`;
                        
                        let pid = note.projectId || note.projectIds?.[0] || 'p_default';
                        let pObj = appConfig.projects.find(p => p.id === pid);
                        let pName = pObj ? `${pObj.name} - ` : '';
                        let cleanTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
                        let displayTitle = (pName + cleanTitle).substring(0, 45);
                        const actualEndHour = (blockStart + actualDuration) % 24;
                        const timeStr = `${decToTime(blockStart)} - ${decToTime(actualEndHour)}`;

                        blockEl.innerHTML = `
                            <div class="block-info">
                                <div class="block-title">${displayTitle} (cont.)</div>
                                <div class="block-meta">${timeStr}</div>
                            </div>
                        `;
                        blockEl.onclick = (e) => {
                            if(isDraggingBlock) return;
                            openTaskModal(null, note.id, e);
                        };
                        dayBlocks.push({ el: blockEl, start: 0, end: overflowDuration, duration: overflowDuration });
                    }
                }
            });
        });

        // Apply overlapping layout
        let groups = [];
        let currentGroup = [];
        let currentGroupEnd = -1;

        dayBlocks.sort((a, b) => a.start - b.start || b.duration - a.duration);

        dayBlocks.forEach(block => {
            if (currentGroup.length === 0) {
                currentGroup.push(block);
                currentGroupEnd = block.end;
            } else if (block.start < currentGroupEnd) {
                currentGroup.push(block);
                currentGroupEnd = Math.max(currentGroupEnd, block.end);
            } else {
                groups.push(currentGroup);
                currentGroup = [block];
                currentGroupEnd = block.end;
            }
        });
        if (currentGroup.length > 0) groups.push(currentGroup);

        groups.forEach(group => {
            let columns = [];
            group.forEach(block => {
                let placed = false;
                for (let i = 0; i < columns.length; i++) {
                    let lastBlock = columns[i][columns[i].length - 1];
                    if (block.start >= lastBlock.end) {
                        columns[i].push(block);
                        block.colIndex = i;
                        placed = true;
                        break;
                    }
                }
                if (!placed) {
                    block.colIndex = columns.length;
                    columns.push([block]);
                }
            });

            let numCols = columns.length;
            group.forEach(block => {
                block.el.style.left = `calc(${block.colIndex} * (100% / ${numCols}) + 6px)`;
                block.el.style.width = `calc((100% / ${numCols}) - 12px)`;
                block.el.style.right = 'auto';
                col.appendChild(block.el);
            });
        });

        colsContainer.appendChild(col);
    });

    // --- NEW: Daily/Weekly Hour Calculations loop through timeBlocks ---
    let actualWeekly = 0;
    let countedIds = new Set();
    notes.forEach(note => {
        if (note.deleted || note.eventId || note.status === 'closed') return;
        (note.timeBlocks || []).forEach(tb => {
            if (weekDateKeys.has(tb.date) && !countedIds.has(`${note.id}-${tb.blockId}`)) {
                actualWeekly += roundToQuarterHour(tb.duration || 1.0);
                countedIds.add(`${note.id}-${tb.blockId}`);
            }
        });
    });

    let actualDaily = 0;
    notes.forEach(note => {
        if (note.deleted || note.eventId || note.status === 'closed') return;
        (note.timeBlocks || []).forEach(tb => {
            if (tb.date === baseDateStr) {
                actualDaily += roundToQuarterHour(tb.duration || 1.0);
            }
        });
    });

    const dailyTotalEl = document.getElementById('trackerDailyTotal');
    if(dailyTotalEl) dailyTotalEl.innerText = `🎯: ${actualDaily}h/${actualWeekly}h`;

    const scrollArea = document.getElementById('timelineScrollArea');
    if (scrollArea && scrollArea.scrollTop === 0) scrollArea.scrollTop = 7 * hourPx; 
}

function toggleDueFilter() {
    const toggle = document.getElementById('dueFilterToggle');
    if (toggle) {
        localStorage.setItem('quadra_due_filter', toggle.checked);
        handleSearch(); 
    }
}

// --- DUE TASKS (TRIAGE) ENGINE ---
function renderOverdueTasksPage() {
    const backlogList = document.getElementById('backlog-list');
    const horizonContainer = document.getElementById('horizon-container');
    
    if (!backlogList || !horizonContainer) return;
    
    backlogList.innerHTML = '';
    horizonContainer.innerHTML = '';
    
    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];
    const todayObj = new Date();
    
    // --- FIX: Grab the global search query ---
    const searchInput = document.getElementById('searchInput');
    const globalQuery = searchInput ? searchInput.value : '';
    
    // 1. Filter eligible notes (Not closed, deleted, meetings, Notebook notes, and MATCHES SEARCH)
    const activeNotes = notes.filter(n => 
        !n.deleted && 
        n.status !== 'closed' && 
        !n.eventId && 
        n.quadrant !== 'notes' && 
        isProjectVisible(n) &&
        matchesSearchQuery(n.text, globalQuery) // <-- Applied here
    );
    
    // 2. Identify Backlog (Overdue OR Unscheduled)
    const backlogNotes = activeNotes.filter(n => {
        if (!n.dueDate) return true; // Unscheduled
        return n.dueDate < todayStr; // Overdue
    });
    
    // Sort Backlog: Overdue dates first, then unscheduled
    backlogNotes.sort((a, b) => {
        if (a.dueDate && b.dueDate) return a.dueDate.localeCompare(b.dueDate);
        if (a.dueDate) return -1;
        if (b.dueDate) return 1;
        return 0;
    });
    
    document.getElementById('backlog-count').innerText = backlogNotes.length;
    
    backlogNotes.forEach(note => {
        backlogList.appendChild(createTriageCard(note, todayStr));
    });
    
    // 3. Generate Horizon Columns (Today + Next 6 Days)
    for (let i = 0; i < 7; i++) {
        const d = new Date();
        d.setDate(todayObj.getDate() + i);
        
        const localY = d.getFullYear();
        const localM = String(d.getMonth() + 1).padStart(2, '0');
        const localD = String(d.getDate()).padStart(2, '0');
        const dateStr = `${localY}-${localM}-${localD}`;
        
        const dayOfWeek = d.getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isToday = i === 0;
        
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        
        const subLabel = dayNames[dayOfWeek];
        const mainLabel = `${monthNames[d.getMonth()]} ${d.getDate()}` + (isToday ? ' (Today)' : '');
        
        const isOOO = appConfig.oooDates && appConfig.oooDates.includes(dateStr);
        
        const col = document.createElement('div');
        col.className = `day-col ${isWeekend ? 'weekend' : ''} ${isOOO ? 'ooo-day' : ''}`;
        col.innerHTML = `
            <div class="day-header ${isToday ? 'today' : ''}">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
                    <div class="sub">${subLabel}</div>
                    <button class="ooo-btn ${isOOO ? 'active' : ''}" onclick="toggleOOODay('${dateStr}')" title="Mark Out of Office">🌴</button>
                </div>
                <div>${mainLabel}</div>
            </div>
            <div class="day-content" ondragover="allowHorizonDrop(event)" ondragleave="dragLeaveHorizon(event)" ondrop="dropToHorizon(event, '${dateStr}')"></div>
        `;
        
        const contentArea = col.querySelector('.day-content');
        
        // Populate tasks scheduled for this specific day
        const dayNotes = activeNotes.filter(n => n.dueDate === dateStr);
        dayNotes.forEach(note => {
            contentArea.appendChild(createTriageCard(note, todayStr));
        });
        
        horizonContainer.appendChild(col);
    }
}

function createTriageCard(note, todayStr) {
    const el = document.createElement('div');
    el.className = `triage-task ${note.quadrant || 'q2'}`;
    el.draggable = true;
    el.ondragstart = (e) => e.dataTransfer.setData('text/plain', note.id);
    el.onclick = (e) => openTaskModal(null, note.id, e);
    
    let title = cleanHTMLToPlainText(note.text).split('\n')[0];
        
    // --- NEW: Calculate Due Date Metadata ---
    let metaText = '';
    //const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];
    
    if (note.dueDate && note.dueDate < todayStr) {
        metaText = `<span style="color: #EF4444; font-weight: 700;">Overdue (${note.dueDate})</span>`;
    } else if (!note.dueDate) {
        metaText = `Unscheduled`;
    } else {
        metaText = `Due ${note.dueDate}`;
    }
    
    // --- NEW: Render Card without Quadrant Pill ---
    el.innerHTML = `
        <div style="font-weight: 600; color: #334155; margin-bottom: 4px;">${parseTags(title)}</div>
        <div style="font-size: 11px; color: #64748B;">${metaText}</div>
    `;
    return el;
}

// --- Triage Drag & Drop Handlers ---
function toggleOOODay(dateStr) {
    if (!appConfig.oooDates) appConfig.oooDates = [];
    const idx = appConfig.oooDates.indexOf(dateStr);
    
    // Toggle the date in the array
    if (idx > -1) {
        appConfig.oooDates.splice(idx, 1);
    } else {
        appConfig.oooDates.push(dateStr);
    }
    
    // Save and re-render
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
    renderOverdueTasksPage();
}
function allowHorizonDrop(ev) {
    ev.preventDefault();
    ev.currentTarget.parentElement.classList.add('drag-over');
}
function dragLeaveHorizon(ev) {
    ev.currentTarget.parentElement.classList.remove('drag-over');
}
function dropToHorizon(ev, dateStr) {
    ev.preventDefault(); 
    ev.currentTarget.parentElement.classList.remove('drag-over');
    
    const noteId = ev.dataTransfer.getData("text/plain");
    const note = notes.find(n => n.id === noteId);
    
    if (note) {
        note.dueDate = dateStr;
        // Shift existing calendar blocks to match the new day
        if (note.timeBlocks && note.timeBlocks.length > 0) {
            note.timeBlocks.forEach(tb => tb.date = dateStr);
        }
        note.dirty = true;
        saveNotes();
        handleSearch(); 
    }
}

function allowBacklogDrop(ev) {
    ev.preventDefault();
    ev.currentTarget.classList.add('drag-over');
}
function dragLeaveBacklog(ev) {
    ev.currentTarget.classList.remove('drag-over');
}
function dropToBacklog(ev) {
    ev.preventDefault();
    ev.currentTarget.classList.remove('drag-over');
    
    const noteId = ev.dataTransfer.getData("text/plain");
    const note = notes.find(n => n.id === noteId);
    
    if (note) {
        note.dueDate = null; // Unschedule the task completely
        if (note.timeBlocks) {
            note.timeBlocks = []; 
        }
        note.dirty = true;
        saveNotes();
        handleSearch();
    }
}

// --- Drag & Resize Engine (Multi-Day Architecture) ---
function startBlockDrag(e, noteId, blockId, isResize) {
    e.stopPropagation();
    const note = notes.find(n => n.id === noteId);
    if (!note || !note.timeBlocks) return;
    
    let tBlock = note.timeBlocks.find(b => b.blockId === blockId);
    if (!tBlock) return;
    
    isDraggingBlock = true;

    dragState = {
        noteId, blockId, isResize,
        startX: e.clientX,
        startY: e.clientY,
        originalStart: tBlock.startHour,
        originalDuration: tBlock.duration,
        hasMoved: false,
        el: document.getElementById(`block-${noteId}-${blockId}`)
    };

    if (dragState.el) {
        dragState.el.style.zIndex = '100';
        dragState.el.style.transition = 'none';
    }
    
    document.addEventListener('mousemove', onBlockDrag);
    document.addEventListener('mouseup', stopBlockDrag);
}

function onBlockDrag(e) {
    if (!dragState) return;
    
    if (!dragState.hasMoved && (Math.abs(e.clientY - dragState.startY) > 3 || Math.abs(e.clientX - dragState.startX) > 3)) {
        dragState.hasMoved = true;
        isDraggingBlock = true;
    }

    if (!dragState.hasMoved) return;

    const hourPx = 60 * timelineZoom;
    const dy = e.clientY - dragState.startY;
    const dx = e.clientX - dragState.startX;
    const dHours = roundToQuarterHour(dy / hourPx); 
    
    const note = notes.find(n => n.id === dragState.noteId);
    if (!note || !note.timeBlocks) return;
    
    let tBlock = note.timeBlocks.find(b => b.blockId === dragState.blockId);
    if (!tBlock) return;
    
    const paletteEl = document.querySelector('.tracker-palette');
    let isOverPalette = false;
    
    if (paletteEl && !dragState.isResize) {
        const rect = paletteEl.getBoundingClientRect();
        if (e.clientX >= rect.left && e.clientX <= rect.right &&
            e.clientY >= rect.top && e.clientY <= rect.bottom) {
            isOverPalette = true;
            paletteEl.style.background = 'rgba(239, 68, 68, 0.05)';
            paletteEl.style.boxShadow = 'inset 0 0 0 2px #EF4444';
        } else {
            paletteEl.style.background = '';
            paletteEl.style.boxShadow = '';
        }
    }

    if (dragState.isResize) {
        let newDuration = Math.max(0.25, dragState.originalDuration + dHours);
        tBlock.duration = newDuration;
        if (dragState.el) {
            dragState.el.style.height = `${newDuration * hourPx}px`;
            const actualEndHour = (tBlock.startHour + newDuration) % 24;
            const timeStr = `${decToTime(tBlock.startHour)} - ${decToTime(actualEndHour)}`;
            const metaEl = dragState.el.querySelector('.block-meta');
            if (metaEl) metaEl.innerText = timeStr;
        }
    } else {
        let newStart = Math.max(0, Math.min(30 - tBlock.duration, dragState.originalStart + dHours));
        tBlock.startHour = newStart;
        if (dragState.el) {
            dragState.el.style.top = `${(newStart % 24) * hourPx}px`;
            
            if (isOverPalette) {
                dragState.el.style.transform = `translateX(${dx}px)`;
                dragState.el.style.opacity = '0.5';
            } else {
                dragState.el.style.transform = '';
                dragState.el.style.opacity = '1';
            }

            const actualEndHour = (newStart + tBlock.duration) % 24;
            const timeStr = `${decToTime(newStart)} - ${decToTime(actualEndHour)}`;
            const metaEl = dragState.el.querySelector('.block-meta');
            if (metaEl) metaEl.innerText = timeStr;
        }
    }
}

function stopBlockDrag(e) {
    if (dragState) {
        document.removeEventListener('mousemove', onBlockDrag);
        document.removeEventListener('mouseup', stopBlockDrag);
        
        let didMove = dragState.hasMoved;
        let unscheduled = false;

        const paletteEl = document.querySelector('.tracker-palette');
        if (paletteEl && !dragState.isResize) {
            const rect = paletteEl.getBoundingClientRect();
            if (e.clientX >= rect.left && e.clientX <= rect.right &&
                e.clientY >= rect.top && e.clientY <= rect.bottom) {
                
                const note = notes.find(n => n.id === dragState.noteId);
                if (note && note.timeBlocks) {
                    // NEW: Filter out this specific block to un-schedule it
                    note.timeBlocks = note.timeBlocks.filter(b => b.blockId !== dragState.blockId);
                    note.dirty = true;
                    unscheduled = true;
                    saveNotes();
                }
            }
            paletteEl.style.background = '';
            paletteEl.style.boxShadow = '';
        }

        if (didMove && !unscheduled) {
            const note = notes.find(n => n.id === dragState.noteId);
            if (note) { note.dirty = true; saveNotes(); }
        }
        
        if (dragState.el) {
            dragState.el.style.zIndex = '10';
            dragState.el.style.transition = ''; 
            dragState.el.style.transform = ''; 
            dragState.el.style.opacity = '1';
        }
        
        dragState = null;
        if (didMove || unscheduled) { renderTrackerTimeline(); setTimeout(() => { isDraggingBlock = false; }, 50); } 
        else { isDraggingBlock = false; }
    }
}

function dropToTracker(ev, dateStr) {
    ev.preventDefault(); ev.currentTarget.classList.remove('drag-over');
    const noteId = ev.dataTransfer.getData("text/plain");
    const note = notes.find(n => n.id === noteId);
    const hourPx = 60 * timelineZoom;
    
    if (note) {
        const rect = ev.currentTarget.getBoundingClientRect();
        const y = ev.clientY - rect.top; 
        let dropHour = roundToQuarterHour(y / hourPx);

        if (!note.timeBlocks) note.timeBlocks = [];
        
        // NEW: Push a fresh calendar instance instead of overwriting a date
        note.timeBlocks.push({
            blockId: 'b_' + Date.now().toString() + Math.floor(Math.random() * 1000),
            date: dateStr,
            startHour: dropHour,
            duration: 1.0
        });

        // We purposely do NOT overwrite note.dueDate here, allowing the deadline to stay distinct
        note.dirty = true; 
        saveNotes(); 
        handleSearch(); 
    }
}

function handleTimelineClick(ev, dateStr) {
    if (isDraggingBlock) return;
    if (ev.target.closest('.logged-block')) return; 
    const hourPx = 60 * timelineZoom;
    
    const rect = ev.currentTarget.getBoundingClientRect();
    const y = ev.clientY - rect.top; 
    let dropHour = roundToQuarterHour(y / hourPx);

    openTaskModal('calendar', null, ev, { date: dateStr, startHour: dropHour });
}

function allowTrackerDrop(ev) { ev.preventDefault(); ev.currentTarget.classList.add('drag-over'); }
function dragLeaveTracker(ev) { ev.currentTarget.classList.remove('drag-over'); }

function openShortcutsModal() {
    const modal = document.getElementById('shortcutsModal');
    if (modal) modal.style.display = 'flex';
}

function closeShortcutsModal() {
    const modal = document.getElementById('shortcutsModal');
    if (modal) modal.style.display = 'none';
}

document.addEventListener('click', function(e) {
    const li = e.target.closest('li.todo-item');
    if (li) {
        // The padding-left is 28px. If the click is on the far left side, 
        // they clicked the ::before pseudo-element (the checkbox).
        if (e.offsetX >= 0 && e.offsetX <= 26) {
            e.preventDefault();
            e.stopPropagation();
            
            li.classList.toggle('completed');
            triggerAutoSaveInterval(); // Save the state
        }
    }
});

document.addEventListener('keydown', (e) => {
    // --- NEW: Intercept Ctrl+S / Cmd+S globally ---
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault(); // Stop default browser "Save Webpage" dialog
        
        // 1. Save standard state to localStorage first
        saveNotes();
        
        // 2. Trigger SQLite backup to Google Drive AppData
        uploadDatabaseToDrive();
        
        // 3. Trigger Google Tasks Sync
        performBackgroundSync();
        
        return;
    }

    const isEditingText = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName) || document.activeElement.isContentEditable;
    
    if (isEditingText) {
        // --- NEW: Rich Text Formatting Shortcuts ---
        
        // Ctrl+Alt+Shift+S : Code Block
        if ((e.ctrlKey || e.metaKey) && e.altKey && e.shiftKey && e.key.toLowerCase() === 's') {
            e.preventDefault();
            insertCodeBlock();
            return;
        }
        
        // Ctrl+Shift+X : Strikethrough
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'x' && !e.altKey) {
            e.preventDefault();
            document.execCommand('strikeThrough', false, null);
            triggerAutoSaveInterval();
            return;
        }
        
        // Ctrl+B : Bold
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b' && !e.shiftKey && !e.altKey) {
            e.preventDefault();
            document.execCommand('bold', false, null);
            triggerAutoSaveInterval();
            return;
        }
        
        // Ctrl+I : Italics
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'i' && !e.shiftKey && !e.altKey) {
            e.preventDefault();
            document.execCommand('italic', false, null);
            triggerAutoSaveInterval();
            return;
        }

        // Existing custom Ctrl+1 checklist shortcut
        if (e.ctrlKey && e.key === '1' && !e.shiftKey && !e.altKey) {
            e.preventDefault();
            toggleChecklistFormatting();
            return; 
        }
    }

    if (e.key === 'Escape') {
        const taskModal = document.getElementById('taskModal');
        const shortcutsModal = document.getElementById('shortcutsModal');
        const projectModal = document.getElementById('projectModal');

        if (taskModal && taskModal.style.display === 'flex') closeTaskModal();
        if (shortcutsModal && shortcutsModal.style.display === 'flex') closeShortcutsModal();
        if (projectModal && projectModal.style.display === 'flex') closeProjectModal();
    } else if (!isEditingText) {
        // --- NEW: Alt + Up/Down Arrow for Project Traversal ---
        if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
            e.preventDefault();
            const unarchived = appConfig.projects.filter(p => !p.archived);
            
            if (unarchived.length > 1) {
                let currentIndex = unarchived.findIndex(p => p.visible);
                if (currentIndex === -1) currentIndex = 0;
                
                let newIndex;
                if (e.key === 'ArrowDown') {
                    newIndex = (currentIndex + 1) % unarchived.length; // Next project
                } else {
                    newIndex = (currentIndex - 1 + unarchived.length) % unarchived.length; // Previous project
                }
                
                const targetProjectId = unarchived[newIndex].id;
                appConfig.projects.forEach(p => p.visible = (p.id === targetProjectId));
                
                localStorage.setItem('quadra_config', JSON.stringify(appConfig));
                renderProjectTabs();
                handleSearch();
            }
            return;
        }
        if (e.shiftKey && (e.key === '?' || e.key === '/')) {
            e.preventDefault();
            openShortcutsModal();
        } else if (e.key === '/') {
            e.preventDefault();
            const searchInput = document.getElementById('searchInput');
            if (searchInput) {
                searchInput.focus();
                searchInput.select();
            }
        } else if (e.key.toLowerCase() === 'q' && appConfig.viewsEnabled.grid) {
            e.preventDefault();
            setLayout('grid');
        } else if (e.key.toLowerCase() === 'k' && appConfig.viewsEnabled.kanban) {
            e.preventDefault();
            setLayout('kanban');
        } else if (e.key.toLowerCase() === 'c' && appConfig.viewsEnabled.tracker) {
            e.preventDefault();
            setLayout('tracker');
        } else if (e.key.toLowerCase() === 'o' && appConfig.viewsEnabled.overdue) {
            e.preventDefault();
            setLayout('overdue');
        } else if (e.key.toLowerCase() === 'n' && appConfig.viewsEnabled.notebook) {
            // --- NEW: Direct shortcut for Notebook ---
            e.preventDefault();
            setLayout('notebook');
        } else if (e.key === '`' || e.key === '~') {
            e.preventDefault();
            
            // --- FIX: Matched to the new visual order (Grid, Kanban, Overdue, Tracker, Notebook) ---
            const allViews = ['grid', 'kanban', 'overdue', 'tracker', 'notebook'];
            
            const views = allViews.filter(v => appConfig.viewsEnabled[v]);
            if (views.length === 0) return; 
            
            let idx = views.indexOf(currentLayout);
            if (idx === -1) idx = 0;
            if (e.shiftKey) {
                idx = (idx - 1 + views.length) % views.length;
            } else {
                idx = (idx + 1) % views.length;
            }
            setLayout(views[idx]);
        }
    }
});

function toggleTaskCompleteFromModal() {
    if (!currentEditingId) return;
    
    const quadrantSelect = document.getElementById('taskQuadrant');
    
    if (quadrantSelect) {
        // 1. Force the dropdown menu to "Closed" (or back to "Inbox" if restoring)
        if (quadrantSelect.value === 'closed') {
            quadrantSelect.value = 'inbox';
        } else {
            quadrantSelect.value = 'closed';
        }
        saveTaskModal();
    }
}

function triggerAutoSaveInterval() {
    if (autoSaveTimerId) return; 
    autoSaveTimerId = setInterval(() => {
        if (!currentEditingId) return;
        const note = notes.find(n => n.id === currentEditingId);
        if (note) {
            const titleText = document.getElementById('taskTitleInput').innerHTML;
            const infoText = document.getElementById('taskInfoInput').innerHTML;
            
            if ((!titleText || titleText === '<br>') && (!infoText || infoText === '<br>')) return;

            const revisedFullText = titleText + (infoText && infoText !== '<br>' ? ('\n' + infoText) : '');
            if (note.text !== revisedFullText) {
                note.text = revisedFullText;
                note.dirty = true;
                saveNotes();
                console.log("⏰ Snapshot checkpoint autosaved into LocalStorage.");
            }
        }
    }, 30000); 
}

function clearAutoSaveInterval() {
    if (autoSaveTimerId) {
        clearInterval(autoSaveTimerId);
        autoSaveTimerId = null;
    }
}

function openTaskModal(quadrant = null, noteId = null, event = null, timelineContext = null) {
    if (event) event.stopPropagation();

    isDocMode = false;
    document.querySelector('#taskModal .modal-content').classList.remove('doc-mode');
    const toggleBtn = document.getElementById('docModeToggleBtn');
    if (toggleBtn) {
        toggleBtn.innerText = '⛶';
        toggleBtn.title = "Enter Doc Mode";
    }
    
    const modal = document.getElementById('taskModal');
    const titleInput = document.getElementById('taskTitleInput');
    const infoInput = document.getElementById('taskInfoInput');
    const dueDateInput = document.getElementById('taskDueDate');
    const quadrantInput = document.getElementById('taskQuadrant');
    const completeBtn = document.getElementById('taskModalCompleteBtn');

    clearAutoSaveInterval();

    if (noteId) {
        currentEditingId = noteId; 
        const note = notes.find(n => n.id === noteId); 
        document.getElementById('taskModalTitle').innerText = 'Task Details'; 

        let rawText = note.text || "";
        let match = rawText.match(/\n|<br\s*\/?>/i);

        if (!match) {
            titleInput.innerHTML = rawText.trim();
            infoInput.innerHTML = '';
        } else {
            let splitIdx = match.index;
            let skipLen = match[0].length;
            
            titleInput.innerHTML = rawText.substring(0, splitIdx).trim();
            let bodyHTML = rawText.substring(splitIdx + skipLen);
            
            if (!/<[a-z][\s\S]*>/i.test(bodyHTML)) {
                bodyHTML = bodyHTML.replace(/\n/g, '<br>');
            }
            infoInput.innerHTML = bodyHTML;
        }

        formatEditorNodes('taskTitleInput');
        formatEditorNodes('taskInfoInput');

        dueDateInput.value = note.dueDate || '';
        if (quadrantInput) quadrantInput.value = note.quadrant || 'inbox'; 

        completeBtn.style.display = 'inline-block';
        if (note.status === 'closed') {
            completeBtn.innerHTML = '↺ Restore Task';
            completeBtn.style.color = '#3B82F6';
            completeBtn.style.borderColor = '#3B82F6';
        } else {
            completeBtn.innerHTML = '✓ Mark Complete';
            completeBtn.style.color = '#10B981';
            completeBtn.style.borderColor = '#10B981';
        }

    } else { 
        currentEditingId = null; 
        currentAddingQuadrant = quadrant || 'inbox'; 
        document.getElementById('taskModalTitle').innerText = 'Add Task'; 
        titleInput.innerHTML = '';
        infoInput.innerHTML = '';
        dueDateInput.value = timelineContext ? timelineContext.date : '';
        if (quadrantInput) quadrantInput.value = currentAddingQuadrant; 
        completeBtn.style.display = 'none';
        
        pendingTimelineContext = timelineContext || null;
    }
    
    // --- UPDATED: Populate Project Dropdown with "All" Option & Archive Logic ---
    const projectInput = document.getElementById('taskProject');
    if (projectInput) {
        projectInput.innerHTML = '<option value="all">🌐 All Projects (Global)</option>';
        
        let assigned = 'p_default';
        if (noteId) {
            const activeNote = notes.find(n => n.id === noteId);
            assigned = activeNote?.projectIds?.[0] || activeNote?.projectId || 'p_default';
        }

        appConfig.projects.forEach(p => {
            // Only add if it's NOT archived, or if it IS archived but currently assigned to this task
            if (!p.archived || p.id === assigned) {
                const opt = document.createElement('option');
                opt.value = p.id; 
                opt.innerText = p.name + (p.archived ? ' (Archived)' : '');
                projectInput.appendChild(opt);
            }
        });
        
        if (noteId) {
            projectInput.value = assigned;
        } else {
            // If adding a task, default to the currently visible unarchived project
            const visibleProjects = appConfig.projects.filter(p => p.visible && !p.archived);
            projectInput.value = visibleProjects.length === 1 ? visibleProjects[0].id : 'p_default';
        }
    }
    document.getElementById('taskModalContent').classList.remove('time-panel-open');
    document.getElementById('taskModalRightPane').style.display = 'none';
    
    document.getElementById('quickLogDate').value = new Date().toLocaleDateString('en-CA').split('T')[0];
    document.getElementById('quickLogHours').value = '';
    const deleteBtn = document.getElementById('taskModalDeleteBtn');
    if (deleteBtn) deleteBtn.style.display = noteId ? 'inline-block' : 'none';

    renderQuickTimeLogs();
    modal.style.display = 'flex'; 
    setTimeout(() => titleInput.focus(), 100);
}

function closeTaskModal() { 
    clearAutoSaveInterval();
    document.getElementById('taskModal').style.display = 'none'; 
    pendingTimelineContext = null; 
}

function updateModalForQuadrant() {
    const quadrantInput = document.getElementById('taskQuadrant');
    const dueDateInput = document.getElementById('taskDueDate');
    const completeBtn = document.getElementById('taskModalCompleteBtn');
    
    if (!quadrantInput || !dueDateInput) return;
    
    const isNotes = quadrantInput.value === 'notes';
    dueDateInput.style.display = isNotes ? 'none' : 'block';
    
    if (completeBtn && currentEditingId) {
        completeBtn.style.display = isNotes ? 'none' : 'inline-block';
    }
}

// Add the listener right after DOM load or just float it in the global scope:
document.addEventListener('DOMContentLoaded', () => {
    const quadrantInput = document.getElementById('taskQuadrant');
    if (quadrantInput) quadrantInput.addEventListener('change', updateModalForQuadrant);
});

// --- 1. NEW: Save Modal (Multi-Day Architecture & Projects) ---
function saveTaskModal() {
    clearAutoSaveInterval();
    const titleText = document.getElementById('taskTitleInput').innerHTML; 
    const infoText = document.getElementById('taskInfoInput').innerHTML; 
    const dueDate = document.getElementById('taskDueDate').value;
    
    if ((!titleText || titleText === '<br>') && (!infoText || infoText === '<br>')) return closeTaskModal();

    const fullText = titleText + (infoText && infoText !== '<br>' ? ('\n' + infoText) : '');

    const quadrantSelect = document.getElementById('taskQuadrant');
    const selectedQuadrant = quadrantSelect ? quadrantSelect.value : null;
    let targetQuadForNote = selectedQuadrant || currentAddingQuadrant || 'inbox';
    
    let finalPayloadText = fullText;
    if (targetQuadForNote === 'notes' && !finalPayloadText.includes('#note')) {
        finalPayloadText += ' #note';
    }

    let newTimeBlock = null;
    if (pendingTimelineContext && pendingTimelineContext.startHour !== undefined) {
        newTimeBlock = {
            blockId: 'b_' + Date.now().toString() + Math.floor(Math.random() * 1000),
            date: pendingTimelineContext.date,
            startHour: roundToQuarterHour(pendingTimelineContext.startHour),
            duration: 1.0
        };
    }

    // --- UPDATED: Extract Target Projects ---
    const projectSelect = document.getElementById('taskProject');
    const selectedVal = projectSelect ? projectSelect.value : 'p_default';
    const targetProjects = [selectedVal];

    if (currentEditingId) { 
        const note = notes.find(n => n.id === currentEditingId); 
        if (note) { 
            note.projectIds = targetProjects;
            note.projectId = selectedVal; // Backward compatibility
            note.text = finalPayloadText; 
            note.dueDate = dueDate || null;
            
            if (selectedQuadrant) note.quadrant = selectedQuadrant;
            
            if (selectedQuadrant === 'closed') note.status = 'closed';
            else if (note.status === 'closed' && selectedQuadrant !== 'closed') note.status = 'active';

            if (newTimeBlock) {
                if(!note.timeBlocks) note.timeBlocks = [];
                note.timeBlocks.push(newTimeBlock);
            }
            note.dirty = true; 
            saveNotes(); 
            syncSingleTask(note.id);
            handleSearch(); 
        } 
    } else { 
        let targetQuad = selectedQuadrant || currentAddingQuadrant || 'inbox';
        let newNoteId = Date.now().toString(); 
        let newStatus = targetQuad === 'closed' ? 'closed' : 'active';
        
        notes.push({ 
            id: newNoteId, 
            text: finalPayloadText, 
            quadrant: targetQuad, 
            status: newStatus, 
            dueDate: dueDate || null,
            timeBlocks: newTimeBlock ? [newTimeBlock] : [],
            dirty: true, 
            deleted: false, 
            eventId: null,
            syncFailed: false,
            projectId: selectedVal, // Backward compatibility
            projectIds: targetProjects
        }); 
        saveNotes(); 
        syncSingleTask(newNoteId);
        handleSearch(); 
    }
    closeTaskModal();
}

function clearCalendarCache() {
    if (confirm("Are you sure you want to clear the calendar event cache? This will allow re-importing all meetings when you click Sync Meetings.")) {
        notes = notes.filter(n => !n.eventId);
        saveNotes();
        handleSearch();
        showToast("Calendar event cache cleared!");
    }
}

function openTimesheetModal() {
    const baseDateStr = document.getElementById('trackerDate').value;
    const [y, m, d] = baseDateStr.split('-');
    const baseDate = new Date(y, m - 1, d);
    
    // Find the Monday of the current week
    const dayOfWeek = baseDate.getDay();
    const startOfWeek = new Date(baseDate);
    startOfWeek.setDate(baseDate.getDate() - dayOfWeek); // Moves to Sunday
    
    // Collect specific Mon-Fri date strings
    const workWeekDates = new Set();
    for (let i = 1; i <= 5; i++) {
        let dateIter = new Date(startOfWeek);
        dateIter.setDate(startOfWeek.getDate() + i);
        const localY = dateIter.getFullYear();
        const localM = String(dateIter.getMonth() + 1).padStart(2, '0');
        const localD = String(dateIter.getDate()).padStart(2, '0');
        workWeekDates.add(`${localY}-${localM}-${localD}`);
    }

    let projectTotals = {};
    let totalLogged = 0;
    let detailedLogs = []; // NEW: Array to hold line-by-line data

    notes.forEach(note => {
        if (note.deleted) return;
        
        // Get actual assigned project (fallback to generic if none)
        const pid = note.projectId || note.projectIds?.[0] || 'p_default';
        
        let blocksToProcess = note.timeBlocks || [];
        // Catch legacy or Google imported formats
        if (note.eventId && note.dueTime !== undefined) {
            blocksToProcess = [{ date: note.dueDate, startHour: note.dueTime, duration: note.dueDuration }];
        }

        blocksToProcess.forEach(tb => {
            if (workWeekDates.has(tb.date)) {
                const dur = roundToQuarterHour(tb.duration || 0);
                if (dur > 0) {
                    // 1. Tally for the UI Summary
                    if (!projectTotals[pid]) projectTotals[pid] = 0;
                    projectTotals[pid] += dur;
                    totalLogged += dur;
                    
                    // 2. Log for the Google Sheets Payload
                    let cleanTextTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
                    detailedLogs.push({
                        date: tb.date,
                        projectId: pid,
                        title: cleanTextTitle,
                        hours: dur
                    });
                }
            }
        });
    });

    const tbody = document.getElementById('timesheetTableBody');
    tbody.innerHTML = '';
    
    Object.keys(projectTotals).forEach(pid => {
        const pObj = appConfig.projects.find(p => p.id === pid);
        const pName = pObj ? pObj.name : pid;
        
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${pName}</td><td style="text-align: right;">${projectTotals[pid].toFixed(2)}</td>`;
        tbody.appendChild(tr);
    });
    
    document.getElementById('timesheetTotalHours').innerText = totalLogged.toFixed(2);
    
    // Progress Bar Math
    const target = 40;
    const remaining = Math.max(0, target - totalLogged);
    const pct = Math.min(100, (totalLogged / target) * 100);
    
    const remEl = document.getElementById('timesheetRemaining');
    const barEl = document.getElementById('timesheetProgressBar');
    
    remEl.innerText = remaining > 0 ? `${remaining.toFixed(2)} Hours Remaining` : `Target Met!`;
    remEl.style.color = remaining > 0 ? '#F59E0B' : '#10B981';
    
    barEl.style.width = `${pct}%`;
    barEl.className = 'progress-bar-fill'; 
    if (pct < 50) barEl.classList.add('danger');
    else if (pct < 100) barEl.classList.add('warning');
    
    document.getElementById('timesheetModal').style.display = 'flex';
    
    // --- UPDATED: Pass the detailed logs instead of the summary totals ---
    document.getElementById('btnSyncTimesheet').onclick = () => confirmAndSyncTimesheet(detailedLogs, baseDateStr);
}

async function confirmAndSyncTimesheet(payloadData, dateRef) {
    if (!appConfig.timesheetUrl) {
        return showToast("❌ Please add your Timesheet Web App URL in Settings first.");
    }

    const btn = document.getElementById('btnSyncTimesheet');
    const originalText = btn.innerText;
    
    btn.innerText = "Syncing to Sheets...";
    btn.disabled = true;

    const payload = {
        weekOf: dateRef,
        data: payloadData
    };

    try {
        const response = await fetch(appConfig.timesheetUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }, 
            body: JSON.stringify(payload)
        });
        
        const result = await response.json();
        
        if (result.status === 'success') {
            showToast("✅ Timesheet successfully logged to Google Sheets!");
            document.getElementById('timesheetModal').style.display = 'none';
        } else {
            throw new Error(result.message);
        }
    } catch (e) {
        console.error("Timesheet Sync Error:", e);
        showToast("❌ Sync failed. Check the console or verify your Web App URL.");
    } finally {
        btn.innerText = originalText;
        btn.disabled = false;
    }
}

async function loadCalendars() {
    if (typeof gapi === 'undefined' || !gapi.client || !gapi.client.getToken()) return;
    const token = gapi.client.getToken();
    if (!token || !token.access_token) return;
    try {
        const res = await fetch('https://www.googleapis.com/calendar/v3/users/me/calendarList', {
            headers: { 'Authorization': `Bearer ${token.access_token}` }
        });
        const data = await res.json();
        if (data.error) return;

        const calendars = data.items || [];
        
        const sourceSelect = document.getElementById('sourceCalendar');
        const targetSelect = document.getElementById('targetCalendar');
        
        if (!sourceSelect || !targetSelect) return;

        sourceSelect.innerHTML = '<option value="">-- Select Source Calendar --</option>';
        targetSelect.innerHTML = '<option value="">-- Select Target Calendar --</option>';
        
        calendars.forEach(cal => {
            // Populate Source Dropdown
            const opt1 = document.createElement('option');
            opt1.value = cal.id; opt1.innerText = cal.summary;
            if(cal.id === appConfig.sourceCalendar) opt1.selected = true;
            sourceSelect.appendChild(opt1);
            
            // Populate Target Dropdown
            const opt2 = document.createElement('option');
            opt2.value = cal.id; opt2.innerText = cal.summary;
            if(cal.id === appConfig.targetCalendar) opt2.selected = true;
            targetSelect.appendChild(opt2);
        });
    } catch (e) {
        console.error("Auto Calendar Load Error:", e);
    }
}

function renderScheduleSettings() {
    const container = document.getElementById('scheduleConfigList');
    container.innerHTML = '';
    appSchedule.forEach(block => addScheduleRow(block));
}

function formatTimeForInput(decimalHour) {
    let normalized = decimalHour % 24;
    const hrs = Math.floor(normalized).toString().padStart(2, '0');
    const mins = Math.round((normalized % 1) * 60).toString().padStart(2, '0');
    return `${hrs}:${mins}`;
}

function addScheduleRow(block = {title: '', startHour: 14, endHour: 29}) {
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

function formatEditorNodes(editorId) {
    const editor = document.getElementById(editorId);
    if (!editor) return;

    let sel = window.getSelection();
    let hasCaret = false;
    
    // Only insert marker if we are actively editing this specific field
    if (sel.rangeCount > 0 && editor.contains(sel.focusNode)) {
        let r = sel.getRangeAt(0);
        let marker = document.createElement('span');
        marker.id = 'caret-marker';
        r.insertNode(marker);
        hasCaret = true;
    }

    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null, false);
    let nodesToProcess = [];
    let node;
    
    while (node = walker.nextNode()) {
        let p = node.parentNode;
        if (p.tagName === 'A' || p.classList.contains('hashtag') || p.classList.contains('person-tag') || p.id === 'caret-marker') {
            continue;
        }
        // NEW: Safe boundary check test
        if (/(https?:\/\/[^\s]+)|(^|[\s\(\)\[\]\{\}>;"',\.|])([#@][a-zA-Z0-9_]+)/.test(node.nodeValue)) {
            nodesToProcess.push(node);
        }
    }

    nodesToProcess.forEach(n => {
        let span = document.createElement('span');
        let escaped = escapeHTML(n.nodeValue);
        
        // NEW: Uses $1 to preserve the preceding space/bracket, and $2 for the actual tag
        let formatted = escaped
            .replace(/(https?:\/\/[^\s]+)/g, '<a href="$1" target="_blank" style="color:var(--brand-primary); text-decoration:underline;">$1</a>')
            .replace(/(^|[\s\(\)\[\]\{\}>;"',\.|])(#[a-zA-Z0-9_]+)/g, '$1<span class="hashtag">$2</span>')
            .replace(/(^|[\s\(\)\[\]\{\}>;"',\.|])(@[a-zA-Z0-9_]+)/g, '$1<span class="person-tag">$2</span>');
            
        span.innerHTML = formatted;
        n.parentNode.replaceChild(span, n);
        while (span.firstChild) {
            span.parentNode.insertBefore(span.firstChild, span);
        }
        span.parentNode.removeChild(span);
    });

    if (hasCaret) {
        let marker = document.getElementById('caret-marker');
        if (marker) {
            let newRange = document.createRange();
            newRange.setStartAfter(marker);
            newRange.collapse(true);
            sel.removeAllRanges();
            sel.addRange(newRange);
            marker.parentNode.removeChild(marker);
        }
    }
}

function saveSettings() {
    appConfig.clientId = document.getElementById('configClientId').value.trim();
    appConfig.apiKey = document.getElementById('configApiKey').value.trim();
    appConfig.timesheetUrl = document.getElementById('configTimesheetUrl').value.trim();
    
    const ignoreEl = document.getElementById('configIgnoreKeywords');
    if (ignoreEl) appConfig.ignoreKeywords = ignoreEl.value.trim();
    
    const calSourceEl = document.getElementById('configCalSource');
    if (calSourceEl) appConfig.calSource = calSourceEl.value;
    
    const sourceSelect = document.getElementById('sourceCalendar');
    if (sourceSelect) appConfig.sourceCalendar = sourceSelect.value;
    
    const targetSelect = document.getElementById('targetCalendar');
    if (targetSelect) appConfig.targetCalendar = targetSelect.value;
    
    const importBehavior = document.querySelector('input[name="importBehavior"]:checked');
    if (importBehavior) appConfig.importBehavior = importBehavior.value;
    
    appConfig.primaryTz = document.getElementById('configPrimaryTz').value;
    appConfig.secondaryTz = document.getElementById('configSecondaryTz').value;

    appConfig.defaultView = document.getElementById('configDefaultView').value;
    appConfig.viewsEnabled = {
        grid: document.getElementById('configViewGrid').checked,
        kanban: document.getElementById('configViewKanban').checked,
        tracker: document.getElementById('configViewTracker').checked,
        overdue: document.getElementById('configViewOverdue').checked,
        notebook: document.getElementById('configViewNotebook').checked
    };
    
    if (!appConfig.viewsEnabled[appConfig.defaultView]) {
        const firstEnabled = Object.keys(appConfig.viewsEnabled).find(k => appConfig.viewsEnabled[k]);
        if (firstEnabled) {
            appConfig.defaultView = firstEnabled;
            document.getElementById('configDefaultView').value = firstEnabled;
        } else {
            appConfig.viewsEnabled.grid = true;
            appConfig.defaultView = 'grid';
            document.getElementById('configDefaultView').value = 'grid';
            document.getElementById('configViewGrid').checked = true;
        }
    }
    
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
    
    const rows = document.querySelectorAll('.schedule-row');
    appSchedule = Array.from(rows).map(row => {
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
    localStorage.setItem('quadra_schedule', JSON.stringify(appSchedule));
    
    applyViewVisibility();
    closeSettingsPage(); 
    showToast("Configuration saved!");
    checkConfigState();
    if (currentLayout === 'tracker') renderTrackerTimeline();
}

function updateQuickTags() {
    const tagsBar = document.getElementById('quick-tags-bar');
    const searchInput = document.getElementById('searchInput');
    if (!tagsBar || !searchInput) return;
    
    let tagCounts = new Map();
    const hexColorRegex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/i; 
    
    const dueToggle = document.getElementById('dueFilterToggle');
    const isDueFilterOn = dueToggle && dueToggle.checked;
    
    notes.forEach(note => {
        // FIX 1: Ignore deleted, calendar events, AND CLOSED tasks
        if (note.deleted || note.eventId || note.status === 'closed') return;
        
        if (!isProjectVisible(note)) return; 
        
        if (isDueFilterOn && !note.dueDate && note.quadrant !== 'notes') return;
        
        // Add spaces to line breaks and block elements so tags don't get squashed together
        let tempDiv = document.createElement('div');
        let htmlString = (note.text || '').replace(/<br\s*\/?>/gi, ' ').replace(/<\/div>|<\/li>|<\/p>/gi, ' ');
        tempDiv.innerHTML = htmlString;
        let safePlainText = tempDiv.textContent || tempDiv.innerText || '';
        
        const regex = /(^|[\s\(\)\[\]\{\}>;"',\.|])(#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+)/g;
        let matches = [];
        let m;
        
        // Grab every tag found in this specific task
        while ((m = regex.exec(safePlainText)) !== null) {
            matches.push(m[2].toLowerCase());
        }
        
        if (matches.length > 0) {
            // FIX 2: Deduplicate tags within the SAME task so they only count once globally
            let uniqueTagsInTask = [...new Set(matches)];
            
            uniqueTagsInTask.forEach(tag => {
                if (hexColorRegex.test(tag)) return; // Ignore hex color codes
                tagCounts.set(tag, (tagCounts.get(tag) || 0) + 1);
            });
        }
    });
    
    tagsBar.innerHTML = '';
    
    const rawSearchValue = searchInput.value;
    const currentSearch = rawSearchValue.trim();
    
    const words = rawSearchValue.split(/\s+/);
    const lastWord = words[words.length - 1];
    
    let tagFilter = null;
    if (lastWord.startsWith('#') || lastWord.startsWith('@')) {
        tagFilter = lastWord.toLowerCase(); 
    }
    
    if (currentSearch.length > 0) {
        const endsWithOperator = /\b(AND|OR)$/i.test(currentSearch);
        
        if (!endsWithOperator && !tagFilter) {
            ['AND', 'OR'].forEach(op => {
                let opBtn = document.createElement('button');
                opBtn.className = 'filter-tag';
                opBtn.style.backgroundColor = '#E2E8F0'; 
                opBtn.style.color = '#475569';
                opBtn.style.fontWeight = '700';
                opBtn.innerText = op;
                
                opBtn.onclick = () => {
                    searchInput.value = currentSearch + ` ${op} `;
                    searchInput.focus();
                    handleSearch(); 
                };
                tagsBar.appendChild(opBtn);
            });
            
            let divider = document.createElement('div');
            divider.style.width = '1px';
            divider.style.backgroundColor = 'var(--border-color)';
            divider.style.margin = '0 8px';
            tagsBar.appendChild(divider);
        }
    }
    
    // Sort Tags: Highest count first, then alphabetically
    let sortedTags = Array.from(tagCounts.entries()).sort((a, b) => {
        if (b[1] !== a[1]) {
            return b[1] - a[1]; 
        }
        return a[0].localeCompare(b[0]); 
    });
    
    if (tagFilter) {
        sortedTags = sortedTags.filter(([tag, count]) => tag.startsWith(tagFilter));
    }
    
    sortedTags.forEach(([tag, count]) => {
        let btn = document.createElement('button');
        btn.className = 'filter-tag' + (tag.startsWith('@') ? ' person-filter' : '');
        btn.innerText = `${tag} (${count})`;
        
        btn.onclick = () => {
            let currentVal = searchInput.value.replace(/\s+$/, '');
            let words = currentVal ? currentVal.split(/\s+/) : [];
            
            if (words.length > 0 && (words[words.length - 1].startsWith('#') || words[words.length - 1].startsWith('@'))) {
                words.pop();
            }
            
            words.push(tag);
            
            searchInput.value = words.join(' ') + ' ';
            searchInput.focus();
            searchInput.setSelectionRange(searchInput.value.length, searchInput.value.length);
            
            handleSearch();
        };
        tagsBar.appendChild(btn);
    });
}

function handleSearch() {
    const searchInput = document.getElementById('searchInput');
    if(!searchInput) return;
    const query = searchInput.value;
    
    if (query.trim().length > 0) {
        localStorage.setItem('quadra_search', query);
    } else {
        localStorage.removeItem('quadra_search');
    }

    document.getElementById('clearSearchBtn').style.display = query.length > 0 ? 'block' : 'none';

    // --- NEW: Sync Global Search to Local Palette Search ---
    const paletteSearchInput = document.getElementById('paletteSearchInput');
    // Only sync if the user isn't actively typing inside the palette search box
    if (paletteSearchInput && document.activeElement !== paletteSearchInput) {
        paletteSearchInput.value = query;
        const clearPaletteBtn = document.getElementById('clearPaletteSearchBtn');
        if (clearPaletteBtn) {
            clearPaletteBtn.style.display = query.trim().length > 0 ? 'block' : 'none';
        }
    }

    renderNotes(query);
}

// --- 3. UPDATED: Project Tabs Engine (With Rename/Delete) ---
function renderProjectTabs() {
    const container = document.getElementById('project-tabs-container');
    if (!container) return;
    container.innerHTML = '';
    
    // Only render tabs for projects that are not archived
    appConfig.projects.filter(p => !p.archived).forEach(proj => {
        const tab = document.createElement('div');
        tab.className = `project-tab ${proj.visible ? 'active-view' : ''}`;
        
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = proj.visible !== false;
        
        cb.onclick = (e) => {
            e.stopPropagation();
            proj.visible = cb.checked;
            const unarchived = appConfig.projects.filter(p => !p.archived);
            if (!unarchived.some(p => p.visible)) {
                proj.visible = true; 
                showToast("At least one project must be visible.");
            } else {
                localStorage.setItem('quadra_config', JSON.stringify(appConfig));
                renderProjectTabs();
                handleSearch(); 
            }
        };
        
        const label = document.createElement('span');
        label.innerText = proj.name;
        
        tab.appendChild(cb);
        tab.appendChild(label);
        
        tab.onclick = () => {
            appConfig.projects.forEach(p => p.visible = (p.id === proj.id));
            localStorage.setItem('quadra_config', JSON.stringify(appConfig));
            renderProjectTabs();
            handleSearch();
        };
        
        tab.ondblclick = (e) => {
            e.stopPropagation();
            openProjectModal(proj.id);
        };
        
        container.appendChild(tab);
    });
}

function addNewProject() {
    const name = prompt("Enter new project name:");
    if (name && name.trim()) {
        const newId = 'p_' + Date.now();
        // Make the new project the only visible one immediately
        appConfig.projects.forEach(p => p.visible = false);
        appConfig.projects.push({ id: newId, name: name.trim(), visible: true });
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        renderProjectTabs();
        handleSearch();
    }
}

// Global filter check for views
function isProjectVisible(note) {
    if (!note) return false;
    
    let pIds = note.projectIds;
    if (!pIds) {
        pIds = note.projectId ? [note.projectId] : ['p_default'];
    }

    if (pIds.includes('all')) return true;

    return pIds.some(pid => {
        const p = appConfig.projects.find(proj => proj.id === pid);
        return p ? (!p.archived && p.visible !== false) : false;
    });
}


function renderNotes(searchQuery = '') {
    ['q1', 'q2', 'q3', 'q4', 'inbox', 'calendar', 'notes', 'closed'].forEach(q => { 
        const el = document.getElementById(`list-${q}`); 
        if (el) el.innerHTML = ''; 
    });
    
    // 1. Filter out deleted, hidden projects, etc.
    let filteredNotes = notes.filter(note => {
        if (note.deleted) return false;
        if (!isProjectVisible(note)) return false;
        
        const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
        const dueToggle = document.getElementById('dueFilterToggle');
        if (dueToggle && dueToggle.checked && !note.dueDate && note.quadrant !== 'notes' && !isCalendarEvent) {
            return false; 
        }
        return matchesSearchQuery(note.text, searchQuery);
    });

    // 2. Group into Bins & Count
    let filteredCounts = { q1: 0, q2: 0, q3: 0, q4: 0, inbox: 0, calendar: 0, notes: 0, closed: 0 };
    let bins = { q1: [], q2: [], q3: [], q4: [], inbox: [], calendar: [], notes: [], closed: [] };

    filteredNotes.forEach(note => {
        let targetQuad = note.eventId ? 'calendar' : note.quadrant;
        if (bins[targetQuad]) bins[targetQuad].push(note);
        
        const matchesSearch = matchesSearchQuery(note.text, searchQuery);
        if (note.status === 'active' && !note.eventId && filteredCounts[note.quadrant] !== undefined && matchesSearch) {
            filteredCounts[note.quadrant]++;
        }
        if (note.eventId && filteredCounts['calendar'] !== undefined) {
            filteredCounts['calendar']++;
        }
        if (note.quadrant === 'inbox' && filteredCounts['inbox'] !== undefined && matchesSearch) {
            filteredCounts['inbox']++;
        }
    });

    if (!appConfig.sortPrefs) appConfig.sortPrefs = {};

    // 3. Sort each bin individually and render
    ['q1', 'q2', 'q3', 'q4', 'inbox', 'calendar', 'notes', 'closed'].forEach(q => {
        let list = document.getElementById(`list-${q}`);
        if (!list) return;

        // Default: Notes sort newest first, everything else sorts by impending due date
        let pref = appConfig.sortPrefs[q] || (q === 'notes' ? 'created_desc' : 'due_asc');
        let [sortBy, sortDir] = pref.split('_');

        bins[q].sort((a, b) => {
            if (sortBy === 'title') {
                let valA = cleanHTMLToPlainText(a.text).split('\n')[0].toLowerCase();
                let valB = cleanHTMLToPlainText(b.text).split('\n')[0].toLowerCase();
                return sortDir === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
            } else if (sortBy === 'created') {
                // Task IDs are generated via Date.now(), so they natively represent created time
                let valA = parseFloat(a.id) || 0;
                let valB = parseFloat(b.id) || 0;
                return sortDir === 'asc' ? valA - valB : valB - valA;
            } else { 
                // Due Date logic: Tasks without due dates safely sink to the bottom
                if (!a.dueDate && !b.dueDate) return 0;
                if (!a.dueDate) return 1; 
                if (!b.dueDate) return -1;
                return sortDir === 'asc' ? a.dueDate.localeCompare(b.dueDate) : b.dueDate.localeCompare(a.dueDate);
            }
        });

        bins[q].forEach(note => {
            const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
            const noteEl = document.createElement('div'); 
            noteEl.className = 'note' + (note.status === 'closed' ? ' closed-note' : '');
            
            if (note.status === 'active' && !note.eventId) { 
                noteEl.draggable = true; 
                noteEl.ondragstart = (e) => {
                    e.stopPropagation();
                    e.dataTransfer.setData('text/plain', note.id);
                }; 
            }

            const contentWrapper = document.createElement('div'); 
            contentWrapper.className = 'note-content-wrapper';
            
            let overdueIndicator = '';
            if (!isCalendarEvent && note.status === 'active' && note.dueDate) {
                const dueDateStr = note.dueDate.split('T')[0];
                if (dueDateStr < todayStr) {
                    overdueIndicator = `<span style="background: #FFF1F2; color: #9F1239; border: 1px solid #FECDD3; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; margin-right: 6px;">OVERDUE</span>`;
                }
            }

            let cleanTextTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
            contentWrapper.innerHTML = `<div class="note-text">${overdueIndicator}${parseTags(cleanTextTitle)}</div>`;
            
            if (note.dueDate) contentWrapper.innerHTML += `<div style="font-size:12px; color:var(--brand-primary); margin-top:6px; font-weight:500;">🗓️ ${note.dueDate.split('T')[0]}</div>`;
            
            contentWrapper.onclick = (e) => openTaskModal(null, note.id, e);
            noteEl.append(contentWrapper);

            if (!note.eventId) {
                if (note.status !== 'active') {
                    const actionsDiv = document.createElement('div'); 
                    actionsDiv.className = 'note-actions';
                    actionsDiv.innerHTML = `<button class="action-btn restore-btn" onclick="restoreTask('${note.id}')">↺</button><button class="action-btn delete-btn" onclick="deleteTask('${note.id}')">×</button>`;
                    noteEl.append(actionsDiv);
                }
            } else {
                const actionsDiv = document.createElement('div'); 
                actionsDiv.className = 'note-actions';
                actionsDiv.innerHTML = `<button class="action-btn delete-btn" onclick="deleteTask('${note.id}')" title="Remove event">×</button>`;
                noteEl.append(actionsDiv);
            }
            list.appendChild(noteEl);
        });
    });

    document.getElementById('badge-q1').innerText = filteredCounts.q1;
    document.getElementById('badge-q2').innerText = filteredCounts.q2;
    document.getElementById('badge-q3').innerText = filteredCounts.q3;
    document.getElementById('badge-q4').innerText = filteredCounts.q4;
    document.getElementById('badge-closed').innerText = filteredNotes.filter(n => n.status === 'closed' && !n.eventId).length;
    document.getElementById('badge-inbox').innerText = filteredCounts.inbox;
    document.getElementById('badge-calendar').innerText = filteredCounts.calendar;
    
    const notesBadge = document.getElementById('badge-notes');
    if (notesBadge) notesBadge.innerText = filteredCounts.notes;
    
    if (currentLayout === 'tracker') renderTrackerTimeline();
    if (currentLayout === 'notebook') renderNotebookView(); 
    if (currentLayout === 'overdue') renderOverdueTasksPage(); 
    
    updateQuickTags();
}

function completeTask(id) { const note = notes.find(n => n.id === id); if (note) { note.status = 'closed'; note.quadrant = 'closed'; note.dirty = true; saveNotes(); syncSingleTask(id); handleSearch(); } }
function restoreTask(id) { const note = notes.find(n => n.id === id); if (note) { note.status = 'active'; note.quadrant = 'inbox'; note.dirty = true; saveNotes(); syncSingleTask(id); handleSearch(); } }
function deleteTask(id) { const note = notes.find(n => n.id === id); if (note) { note.deleted = true; note.dirty = true; saveNotes(); syncSingleTask(id); handleSearch(); } }

function checkConfigState() {
    const authorizeButton = document.getElementById('authorize_button');
    const signoutButton = document.getElementById('signout_button');

    authorizeButton.style.display = 'none';
    signoutButton.style.display = 'none';

    const savedTokenData = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (savedTokenData && savedTokenData.expires_at > Date.now()) {
        authorizeButton.style.display = 'none';
        signoutButton.style.display = 'inline-block';
        isGoogleSynced = true;
        if (typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: savedTokenData.token });
        
        loadCalendars(); 
        startTokenHeartbeat();
        downloadDatabaseFromDrive(); // <-- NEW: Pre-fetch Drive file ID and DB in background
    } else {
        authorizeButton.style.display = 'inline-block';
        signoutButton.style.display = 'none';
    }
}

function exportData() { const dataStr = JSON.stringify(notes, null, 2); const blob = new Blob([dataStr], { type: "application/json" }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `quadra_backup_${new Date().toISOString().split('T')[0]}.json`; a.click(); URL.revokeObjectURL(url); }
function triggerImport() { document.getElementById('importFile').click(); }
function importData(event) { const file = event.target.files[0]; if (!file) return; const reader = new FileReader(); reader.onload = function(e) { try { const importedNotes = JSON.parse(e.target.result); if (Array.isArray(importedNotes)) { const noteMap = new Map(notes.map(n => [n.id, n])); importedNotes.forEach(inNote => { inNote.dirty = true; noteMap.set(inNote.id, inNote); }); notes = Array.from(noteMap.values()); saveNotes(); handleSearch(); showToast("Tasks merged!"); closeSettingsPage(); } else showToast("Invalid format."); } catch (err) { showToast("Error reading file."); } event.target.value = ''; }; reader.readAsText(file); }

window.addEventListener('load', () => {

    if (!appConfig.sortPrefs) appConfig.sortPrefs = {};
    ['q1', 'q2', 'q3', 'q4', 'inbox', 'calendar', 'notes', 'closed'].forEach(q => {
        const sel = document.getElementById(`sort-${q}`);
        if (sel) {
            sel.value = appConfig.sortPrefs[q] || (q === 'notes' ? 'created_desc' : 'due_asc');
        }
    });
    
    renderProjectTabs();
    
    const matrixContainer = document.getElementById('matrix');
    
    if (appConfig.quadrantOrder && appConfig.quadrantOrder.length > 0) {
        appConfig.quadrantOrder.forEach(id => {
            if (id === 'notes') return;
            const el = document.getElementById(id);
            if (el) matrixContainer.appendChild(el);
        });
    }
    
    document.querySelectorAll('.quadrant').forEach(q => {
        if (appConfig.quadrantWidths && appConfig.quadrantWidths[q.id]) {
            q.style.width = appConfig.quadrantWidths[q.id];
        }
        quadResizeObserver.observe(q);
    });

    if (!localStorage.getItem('quadra_layout')) {
        currentLayout = appConfig.defaultView;
    }
    applyViewVisibility();

    if (typeof gapi !== 'undefined' && appConfig.apiKey) {
        gapi.load('client', () => {
            gapi.client.init({ 
                apiKey: appConfig.apiKey, 
                discoveryDocs: [
                    'https://www.googleapis.com/discovery/v1/apis/tasks/v1/rest',
                    'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
                ] 
            }).catch(() => {});
        });
    }
    if (typeof google !== 'undefined' && google.accounts && appConfig.clientId) {
        tokenClient = google.accounts.oauth2.initTokenClient({ 
            client_id: appConfig.clientId, 
            scope: SCOPES, 
            callback: async (resp) => { 
                if (resp.error !== undefined) { throw (resp); } 
                localStorage.setItem('quadra_gapi_token_v2', JSON.stringify({ token: resp.access_token, expires_at: Date.now() + (resp.expires_in * 1000) })); 
                document.getElementById('auth-overlay').style.display = 'none';
                document.getElementById('authorize_button').style.display = 'none'; 
                document.getElementById('signout_button').style.display = 'inline-block'; 
                isGoogleSynced = true;
                if(typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: resp.access_token });
                loadCalendars(); 
                performBackgroundSync(); 
            }, 
        });
    }
    checkConfigState();

    const searchInput = document.getElementById('searchInput');
    const tagsBar = document.getElementById('quick-tags-bar');

    if (searchInput && tagsBar) {
        searchInput.addEventListener('keydown', function(e) {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                const firstTag = tagsBar.querySelector('.filter-tag');
                if (firstTag) firstTag.focus();
            }
        });

        tagsBar.addEventListener('keydown', function(e) {
            if (e.key === 'ArrowRight') {
                e.preventDefault();
                if (document.activeElement.nextElementSibling) {
                    document.activeElement.nextElementSibling.focus();
                }
            } else if (e.key === 'ArrowLeft') {
                e.preventDefault();
                if (document.activeElement.previousElementSibling) {
                    document.activeElement.previousElementSibling.focus();
                }
            } else if (e.key === 'ArrowUp' || e.key === 'Escape') {
                e.preventDefault();
                searchInput.focus();
            } else if (e.key === 'Enter') {
                e.preventDefault();
                document.activeElement.click();
            }
        });
    }
});

function handleAuthClick() { 
    if (!appConfig.clientId) return openSettingsPage(); 
    if (!tokenClient && typeof google !== 'undefined' && google.accounts) { 
        tokenClient = google.accounts.oauth2.initTokenClient({ 
            client_id: appConfig.clientId, 
            scope: SCOPES, 
            callback: async (resp) => { 
                if (resp.error !== undefined) { throw (resp); } 
                localStorage.setItem('quadra_gapi_token_v2', JSON.stringify({ token: resp.access_token, expires_at: Date.now() + (resp.expires_in * 1000) })); 
                document.getElementById('auth-overlay').style.display = 'none'; 
                document.getElementById('authorize_button').style.display = 'none'; 
                document.getElementById('signout_button').style.display = 'inline-block'; 
                isGoogleSynced = true;
                if(typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: resp.access_token });
                loadCalendars(); 
                performBackgroundSync(); 
            }, 
        }); 
    }
    if (!tokenClient) { showToast("Google services are still loading. Please wait a moment and try again."); return; }
    tokenClient.requestAccessToken({prompt: 'consent'}); 
}

function handleSignoutClick() { 
    const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (savedToken && savedToken.token && typeof google !== 'undefined' && google.accounts && google.accounts.oauth2) { 
        google.accounts.oauth2.revoke(savedToken.token, () => {}); 
    } 
    localStorage.removeItem('quadra_gapi_token_v2'); 
    isGoogleSynced = false; 
    
    if (tokenHeartbeatId) clearInterval(tokenHeartbeatId); // NEW: Kill the heartbeat
    
    document.getElementById('auth-overlay').style.display = 'none';
    document.getElementById('authorize_button').style.display = 'inline-block'; 
    document.getElementById('signout_button').style.display = 'none'; 
    showToast("Signed out successfully.");
}

async function importCalendarEvents() {
    const dateStr = document.getElementById('trackerDate').value;
    const ignoreKeywords = (appConfig.ignoreKeywords || 'out of office, ooo, away, vacation, holiday')
        .toLowerCase()
        .split(',')
        .map(k => k.trim())
        .filter(k => k.length > 0);

    let importedCount = 0;
    let updatedCount = 0;
    let ignoredCount = 0;

    try {
        document.getElementById('sync-banner').style.display = 'block';
        
        const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
        if (!savedToken || !savedToken.token || savedToken.expires_at < Date.now()) {
            document.getElementById('sync-banner').style.display = 'none';
            return showToast("Please sign in to Google first.");
        }
        
        if (!appConfig.sourceCalendar) {
            document.getElementById('sync-banner').style.display = 'none';
            return showToast("Please select a Source Calendar in Settings first.");
        }
        
        const [y, m, day] = dateStr.split('-');
        const timeMin = new Date(y, m-1, day, 0, 0, 0).toISOString();
        const timeMax = new Date(y, m-1, day, 23, 59, 59).toISOString();

        const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(appConfig.sourceCalendar)}/events?timeMin=${timeMin}&timeMax=${timeMax}&singleEvents=true&orderBy=startTime`, {
            headers: { 'Authorization': `Bearer ${savedToken.token}` }
        });
        const data = await res.json();
        if(data.error) throw new Error(data.error.message);

        const autoPlot = appConfig.importBehavior !== 'palette';
        const events = data.items || [];
        
        events.forEach(event => {
            if (!event.start.dateTime) return; // Skip all-day events

            const title = (event.summary || '').toLowerCase();
            const shouldIgnore = ignoreKeywords.some(keyword => title.includes(keyword));

            if (shouldIgnore) {
                ignoredCount++;
                return;
            }
            
            const start = new Date(event.start.dateTime);
            const end = new Date(event.end.dateTime);
            
            const rawStartHour = start.getHours() + (start.getMinutes() / 60);
            let rawEndHour = end.getHours() + (end.getMinutes() / 60);
            if (rawEndHour <= rawStartHour) rawEndHour += 24;
            
            const startHour = roundToQuarterHour(rawStartHour);
            const duration = Math.max(0.25, roundToQuarterHour(rawEndHour - rawStartHour));
            
            const existingNote = notes.find(n => n.eventId === event.id);
            
            // --- Determine routing for timeBlocks ---
            let newTimeBlocks = [];
            if (autoPlot) {
                newTimeBlocks = [{ blockId: 'cal', date: dateStr, startHour: startHour, duration: duration }];
            }
            
            // --- UPDATE EXISTING GOOGLE EVENT ---
            if (existingNote) {
                let changed = false;
                
                if (autoPlot) {
                    if (!existingNote.timeBlocks) existingNote.timeBlocks = [];
                    let calBlock = existingNote.timeBlocks.find(b => b.blockId === 'cal' || b.date === dateStr);
                    
                    if (!calBlock) {
                        existingNote.timeBlocks.push({ blockId: 'cal', date: dateStr, startHour: startHour, duration: duration });
                        changed = true;
                    } else {
                        if (calBlock.startHour !== startHour) { calBlock.startHour = startHour; changed = true; }
                        if (calBlock.duration !== duration) { calBlock.duration = duration; changed = true; }
                    }
                }
                
                const newText = `${event.summary || 'Meeting'} #meeting`;
                if (existingNote.text !== newText) { existingNote.text = newText; changed = true; }
                
                if (changed) { 
                    existingNote.dirty = true; 
                    updatedCount++; 
                }
                return;
            }
            
            // --- INSERT NEW GOOGLE EVENT ---
            notes.push({
                id: Date.now().toString() + Math.random(),
                eventId: event.id, 
                text: `${event.summary || 'Meeting'} #meeting`,
                quadrant: 'q2', 
                status: 'active',
                dirty: false, 
                deleted: false,
                dueDate: dateStr,
                timeBlocks: newTimeBlocks,
                projectId: 'p_default',
                syncFailed: false
            });
            importedCount++;
        });
        
        document.getElementById('sync-banner').style.display = 'none';
        
        if (importedCount > 0 || updatedCount > 0 || ignoredCount > 0) {
            saveNotes();
            renderTrackerTimeline(); 
            showToast(`${importedCount} imported, ${updatedCount} updated, ${ignoredCount} ignored.`);
        } else {
            showToast("No new meetings found or updated for this date.");
        }
        
    } catch(e) {
        console.error("Calendar Sync Error:", e);
        document.getElementById('sync-banner').style.display = 'none';
        showToast("Failed to fetch calendar events.");
    }
}

async function performBackgroundSync() {
    const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (!savedToken || !savedToken.token) return;
    
    try {
        if (typeof gapi === 'undefined' || !gapi.client || !gapi.client.tasks || !gapi.client.tasks.tasklists) return;
        
        const syncBanner = document.getElementById('sync-banner');
        if (syncBanner) syncBanner.style.display = 'block';

        if (!gapi.client.getToken()) gapi.client.setToken({ access_token: savedToken.token });
        
        const response = await gapi.client.tasks.tasklists.list();
        const remoteLists = response.result.items || [];
        const GAPI_LIST_NAMES = { 'inbox': 'Quadra: Inbox', 'q1': 'Quadra: Do First', 'q2': 'Quadra: Schedule', 'q3': 'Quadra: Delegate', 'q4': 'Quadra: Later', 'notes': 'Quadra: Notes', 'closed': 'Quadra: Completed' };
        let gapiListIds = { inbox: null, q1: null, q2: null, q3: null, q4: null, notes: null, closed: null };
        
        for (const quadKey of Object.keys(GAPI_LIST_NAMES)) { 
            const existingList = remoteLists.find(l => l.title === GAPI_LIST_NAMES[quadKey]); 
            if (existingList) { 
                gapiListIds[quadKey] = existingList.id; 
            } else { 
                const newListReq = await gapi.client.tasks.tasklists.insert({ resource: { title: GAPI_LIST_NAMES[quadKey] } }); 
                gapiListIds[quadKey] = newListReq.result.id; 
            } 
        }
        
        localStorage.setItem('quadra_gapi_lists', JSON.stringify(gapiListIds));
        let remoteTaskMap = {};
        const lastSync = localStorage.getItem('quadra_last_sync');

        for (const quadKey of Object.keys(gapiListIds)) { 
            const listId = gapiListIds[quadKey]; 
            let reqOpts = { tasklist: listId, showHidden: true, showDeleted: true, maxResults: 100 };
            if (lastSync) reqOpts.updatedMin = lastSync; 
            
            const tasksReq = await gapi.client.tasks.tasks.list(reqOpts); 
            const rTasks = tasksReq.result.items || []; 
            rTasks.forEach(t => { remoteTaskMap[t.id] = { task: t, listId: listId, quadKey: quadKey }; }); 
        }
        
        const syncSnapshot = JSON.parse(JSON.stringify(notes));
        
        for (let sn of syncSnapshot) {
            if (sn.eventId) continue; 
            
            const remoteObj = remoteTaskMap[sn.id];
            const isStrandedLocal = !isNaN(sn.id) || sn.id.toString().includes('.');
            
            // --- CONFLICT RESOLUTION: GOOGLE WINS ---
            if (remoteObj && currentEditingId !== sn.id) {
                if (remoteObj.task.deleted) {
                    sn.deleted = true;
                    sn.dirty = false;
                } else {
                    let fullText = remoteObj.task.title || '';
                    if (remoteObj.task.notes) fullText += '\n' + remoteObj.task.notes;
                    
                    fullText = fullText.replace(/^\[x\]\s+(.*)$/gm, '<ul class="todo-list"><li class="todo-item completed">$1</li></ul>')
                                    .replace(/^\[ \]\s+(.*)$/gm, '<ul class="todo-list"><li class="todo-item">$1</li></ul>');
                    fullText = fullText.replace(/<\/ul>\s*<ul class="todo-list">/g, '');
                    if (!/<[a-z][\s\S]*>/i.test(fullText)) fullText = fullText.replace(/\n/g, '<br>');
                    
                    sn.text = fullText; 
                    sn.status = remoteObj.task.status === 'completed' ? 'closed' : 'active'; 
                    sn.quadrant = remoteObj.quadKey; 
                    
                    // STRICT DUE DATE SYNC: Updates the badge on the card, completely ignores timeBlocks
                    sn.dueDate = remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null; 

                    sn.dirty = false; 
                    sn.syncFailed = false;
                }
                delete remoteTaskMap[sn.id];
                if (sn.tempId) delete remoteTaskMap[sn.tempId];
                continue; 
            }

            if (sn.deleted && (sn.dirty || isStrandedLocal)) { 
                if (remoteObj) { 
                    try { 
                        await gapi.client.tasks.tasks.delete({ tasklist: remoteObj.listId, task: sn.id }); 
                    } catch(e){} 
                } 
                sn.syncFailed = false;
                sn.dirty = false; 
                continue; 
            }
            
            if ((sn.dirty || isStrandedLocal) && !sn.deleted) {
                const targetListId = gapiListIds[sn.quadrant]; 
                const gStatus = sn.status === 'closed' ? 'completed' : 'needsAction'; 
                
                let plainTextPayload = cleanHTMLToPlainText(sn.text);
                let lines = plainTextPayload.split('\n');
                let tTitle = lines[0].trim() || 'Untitled Task';
                let tNotes = lines.slice(1).join('\n').trim();

                if (tTitle.length > 1000) tTitle = tTitle.substring(0, 1000) + '...';
                if (tNotes.length > 8100) tNotes = tNotes.substring(0, 8100) + '\n\n[...Truncated for Google Tasks]';

                const resourceBody = { title: tTitle, notes: tNotes, status: gStatus }; 
                
                if (sn.dueDate) {
                    const [y, m, d] = sn.dueDate.split('-');
                    resourceBody.due = new Date(Date.UTC(y, m - 1, d, 0, 0, 0)).toISOString();
                }

                if (remoteObj && currentEditingId === sn.id) { 
                    if (remoteObj.listId !== targetListId) { 
                        try { 
                            await gapi.client.tasks.tasks.delete({ tasklist: remoteObj.listId, task: sn.id }); 
                            const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                            sn.tempId = sn.id; sn.id = res.result.id; sn.syncFailed = false;
                        } catch(e){ sn.syncFailed = true; } 
                    } else { 
                        try { 
                            await gapi.client.tasks.tasks.patch({ tasklist: remoteObj.listId, task: sn.id, resource: resourceBody }); 
                            sn.syncFailed = false;
                        } catch(e){ sn.syncFailed = true; } 
                    } 
                    delete remoteTaskMap[sn.id]; 
                    if (sn.tempId) delete remoteTaskMap[sn.tempId]; 
                } else { 
                    if (!isStrandedLocal) {
                        try { 
                            await gapi.client.tasks.tasks.patch({ tasklist: targetListId, task: sn.id, resource: resourceBody }); 
                            sn.syncFailed = false;
                        } catch(e) {
                            if (e.status === 404 || (e.result && e.result.error && e.result.error.code === 404)) {
                                try { 
                                    const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                                    sn.tempId = sn.id; sn.id = res.result.id; sn.syncFailed = false;
                                } catch(err){ sn.syncFailed = true; }
                            } else {
                                sn.syncFailed = true;
                            }
                        }
                    } else {
                        try { 
                            const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                            sn.tempId = sn.id; sn.id = res.result.id; sn.syncFailed = false;
                        } catch(e){ sn.syncFailed = true; } 
                    }
                }
                sn.dirty = false;
            } 
        }
        
        Object.values(remoteTaskMap).forEach(remoteObj => { 
            if (!remoteObj.task.deleted) {
                let fullText = remoteObj.task.title || '';
                if (remoteObj.task.notes) fullText += '\n' + remoteObj.task.notes;
                
                fullText = fullText.replace(/^\[x\]\s+(.*)$/gm, '<ul class="todo-list"><li class="todo-item completed">$1</li></ul>')
                                .replace(/^\[ \]\s+(.*)$/gm, '<ul class="todo-list"><li class="todo-item">$1</li></ul>');
                fullText = fullText.replace(/<\/ul>\s*<ul class="todo-list">/g, '');
                if (!/<[a-z][\s\S]*>/i.test(fullText)) fullText = fullText.replace(/\n/g, '<br>');
                
                // For brand new tasks from Google, initialize with empty timeBlocks array
                syncSnapshot.push({ 
                    id: remoteObj.task.id, text: fullText, quadrant: remoteObj.quadKey, 
                    status: remoteObj.task.status === 'completed' ? 'closed' : 'active', 
                    dueDate: remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null, 
                    timeBlocks: [], 
                    eventId: null, dirty: false, deleted: false, syncFailed: false
                });
            }
        });
        
        let newNotesArray = []; 
        let syncedIds = new Set();
        syncSnapshot.forEach(sn => { 
            if (sn.deleted && !sn.dirty) return; 
            const liveNote = notes.find(n => n.id === sn.id || (sn.tempId && n.id === sn.tempId)); 
            if (liveNote) { 
                sn.eventId = liveNote.eventId || null; 
                if (liveNote.dirty) { 
                    if (sn.tempId) liveNote.id = sn.id; 
                    newNotesArray.push(liveNote); 
                } else { 
                    newNotesArray.push(sn); 
                } 
                syncedIds.add(liveNote.id); 
                if (sn.tempId) syncedIds.add(sn.tempId); 
            } else { 
                newNotesArray.push(sn); 
            } 
        });
        
        notes = newNotesArray; 
        saveNotes(); 
        localStorage.setItem('quadra_last_sync', new Date().toISOString());
        handleSearch(); 
        
        // --- NEW: Also trigger a SQLite backup to Drive when Tasks sync completes ---
        uploadDatabaseToDrive();
        
        if (syncBanner) syncBanner.style.display = 'none';
        showToast("✓ Successfully synced with Google Tasks & Drive!");
        
    } catch (e) {
        console.error("Background sync error:", e);
        const syncBanner = document.getElementById('sync-banner');
        if (syncBanner) syncBanner.style.display = 'none';
        showToast("Tasks sync failed. Check your network or API configuration.");
    }
}

// --- Execute Initial Render logic with Search Preservation ---
setLayout(currentLayout); 

// 1. Restore Due Filter Toggle state BEFORE rendering
const savedDueFilter = localStorage.getItem('quadra_due_filter') !== 'false';
const dueToggleEl = document.getElementById('dueFilterToggle');
if (dueToggleEl) {
    dueToggleEl.checked = savedDueFilter;
    if (savedDueFilter && !localStorage.getItem('quadra_due_filter')) {
        localStorage.setItem('quadra_due_filter', 'true');
    }
}

// 2. Restore Search State
const savedSearch = localStorage.getItem('quadra_search') || '';
const searchInputEl = document.getElementById('searchInput');
if (searchInputEl && savedSearch) {
    searchInputEl.value = savedSearch;
    document.getElementById('clearSearchBtn').style.display = 'block';
}

// 3. Render safely
renderNotes(savedSearch);

// --- Editor Toolbar Logic ---
document.getElementById('editorToolbar')?.addEventListener('click', function(e) {
    let btn = e.target.closest('button');
    if (!btn) return;
    let command = btn.getAttribute('data-command');
    let value = btn.getAttribute('data-value') || null;
    if (command) {
        e.preventDefault();
        document.execCommand(command, false, value);
        document.getElementById('taskInfoInput').focus();
        triggerAutoSaveInterval();
    }
});

function toggleChecklistFormatting() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;

    let container = sel.getRangeAt(0).commonAncestorContainer;
    let el = container.nodeType === 3 ? container.parentNode : container;
    
    // 1. Check if we are already inside a checklist item
    let existingLi = el.closest('li.todo-item');
    if (existingLi) {
        existingLi.classList.toggle('completed');
        triggerAutoSaveInterval();
        return;
    }

    // 2. If not a checklist, safely convert the current line using native commands
    document.execCommand('insertUnorderedList', false, null);
    
    // 3. Immediately upgrade the native list to our custom checklist styles
    setTimeout(() => {
        let currSel = window.getSelection();
        if(!currSel.rangeCount) return;
        let currNode = currSel.getRangeAt(0).commonAncestorContainer;
        let currEl = currNode.nodeType === 3 ? currNode.parentNode : currNode;
        
        let li = currEl.closest('li');
        let ul = currEl.closest('ul');
        
        if (li) li.classList.add('todo-item');
        if (ul) ul.classList.add('todo-list');
        
        triggerAutoSaveInterval();
    }, 10);
}

// --- Live Tag Formatting for Editor ---
// --- Live Tag & Markdown Formatting for Editor ---
['taskTitleInput', 'taskInfoInput'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
        el.addEventListener('keyup', function(e) {
            // Standard tag formatting
            if (e.key === ' ' || e.key === 'Enter') {
                formatEditorNodes(id);
                
                // --- UPDATED: Markdown Auto-formatting (Strict Line Start) ---
                if (e.key === ' ') {
                    const sel = window.getSelection();
                    if (!sel.rangeCount) return;
                    
                    let range = sel.getRangeAt(0);
                    let node = range.startContainer;
                    
                    if (node.nodeType === 3) { 
                        let offset = range.startOffset;
                        let textBeforeCursor = node.textContent.substring(0, offset);
                        
                        // Check if the current text matches our exact triggers
                        if (textBeforeCursor === '* ' || textBeforeCursor === '- ' || textBeforeCursor === '1. ') {
                            
                            // Verify the text node is physically at column 0 of the line
                            let isStartOfLine = false;
                            
                            // Condition 1: It is the very first node in its block container
                            if (!node.previousSibling) {
                                isStartOfLine = true;
                            } 
                            // Condition 2: The element immediately before it is a soft line break
                            else if (node.previousSibling && node.previousSibling.tagName === 'BR') {
                                isStartOfLine = true;
                            }
                            // Condition 3: It's wrapped in a format tag (like a span) that is the first child
                            else if (node.parentNode && node.parentNode !== el && !node.parentNode.previousSibling) {
                                 isStartOfLine = true;
                            }

                            if (isStartOfLine) {
                                range.setStart(node, 0);
                                range.setEnd(node, offset);
                                range.deleteContents(); // Erase the trigger characters
                                
                                if (textBeforeCursor === '1. ') {
                                    document.execCommand('insertOrderedList', false, null);
                                } else {
                                    document.execCommand('insertUnorderedList', false, null);
                                }
                                triggerAutoSaveInterval();
                            }
                        }
                    }
                }
            }
        });
        
        el.addEventListener('paste', function(e) {
            setTimeout(() => {
                formatEditorNodes(id);
            }, 10);
        });
    }
});

// --- Code Block Insertion Engine ---
function insertCodeBlock() {
    const editor = document.getElementById('taskInfoInput');
    if (!editor) return;
    
    editor.focus();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    
    let range = sel.getRangeAt(0);
    
    // Safety check: ensure the user's cursor is actually inside the rich text editor
    let node = range.commonAncestorContainer;
    while (node && node !== editor) {
        if (node.parentNode === editor) break;
        node = node.parentNode;
    }
    
    // If cursor is outside the editor, force it to the end of the editor
    if (!node || (node !== editor && node.parentNode !== editor)) {
        range.selectNodeContents(editor);
        range.collapse(false);
    }

    // 1. Create the fixed-width code block
    const pre = document.createElement('pre');
    pre.className = 'editor-code-block';
    pre.innerHTML = '<br>'; // Gives the block physical height so it can be clicked

    // 2. Create the "escape" div below it so the user isn't trapped inside the formatting
    const escapeDiv = document.createElement('div');
    escapeDiv.innerHTML = '<br>';

    // 3. Inject both into the editor
    range.deleteContents();
    const frag = document.createDocumentFragment();
    frag.appendChild(pre);
    frag.appendChild(escapeDiv);
    range.insertNode(frag);

    // 4. Force the blinking cursor inside the new code block automatically
    range.setStart(pre, 0);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);

    triggerAutoSaveInterval();
}
// --- Google API Token Heartbeat ---
function startTokenHeartbeat() {
    if (tokenHeartbeatId) clearInterval(tokenHeartbeatId);
    
    // Check the token health every 1 minute
    tokenHeartbeatId = setInterval(() => {
        const savedTokenData = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
        
        if (savedTokenData && isGoogleSynced) {
            const timeRemaining = savedTokenData.expires_at - Date.now();
            const fiveMinutes = 5 * 60 * 1000;
            
            // If we have less than 5 minutes left, silently request a new token
            if (timeRemaining > 0 && timeRemaining < fiveMinutes) {
                console.log("Token expiring soon. Attempting silent refresh...");
                attemptSilentTokenRefresh();
            } else if (timeRemaining <= 0) {
                // If it already expired while the computer was asleep, log them out safely
                handleSignoutClick();
                showToast("Google session expired. Please sign in again.");
            }
        }
    }, 60000);
}

function attemptSilentTokenRefresh() {
    if (!tokenClient || typeof google === 'undefined') return;
    
    // The prompt: '' parameter tells Google to skip the consent screen 
    // if the user is already logged into Chrome/Google.
    tokenClient.requestAccessToken({ prompt: '' });
}
// --- Instant Push Engine (Hybrid Sync) ---
async function syncSingleTask(noteId) {
    if (!isGoogleSynced || typeof gapi === 'undefined' || !gapi.client || !gapi.client.tasks) return;
    
    const note = notes.find(n => n.id === noteId);
    if (!note || note.eventId) return;

    let listCache = JSON.parse(localStorage.getItem('quadra_gapi_lists')) || {};
    let targetListId = listCache[note.quadrant];
    if (!targetListId) return;

    try {
        if (note.deleted) {
            await gapi.client.tasks.tasks.delete({ tasklist: targetListId, task: note.id }).catch(()=>{});
            note.dirty = false;
            saveNotes();
            return;
        }

        const gStatus = note.status === 'closed' ? 'completed' : 'needsAction'; 
        let plainTextPayload = cleanHTMLToPlainText(note.text);
        let lines = plainTextPayload.split('\n');
        
        let tTitle = lines[0].trim() || 'Untitled Task';
        let tNotes = lines.slice(1).join('\n').trim();

        // --- PREVENT 400 ERRORS: Graceful Truncation ---
        if (tTitle.length > 1000) tTitle = tTitle.substring(0, 1000) + '...';
        if (tNotes.length > 8100) tNotes = tNotes.substring(0, 8100) + '\n\n[...Truncated for Google Tasks]';

        let resourceBody = { title: tTitle, notes: tNotes, status: gStatus }; 

        if (note.dueDate) {
            const [y, m, d] = note.dueDate.split('-');
            resourceBody.due = new Date(Date.UTC(y, m - 1, d, 0, 0, 0)).toISOString();
        }

        const isNew = !isNaN(note.id) || note.id.toString().includes('.'); 
        
        if (!isNew) {
            try {
                await gapi.client.tasks.tasks.patch({ tasklist: targetListId, task: note.id, resource: resourceBody });
                note.dirty = false;
            } catch(e) {
                // --- PREVENT 404 LOOPS: If task was deleted remotely, recreate it ---
                if (e.status === 404 || (e.result && e.result.error && e.result.error.code === 404)) {
                    console.warn("Task missing on Google. Recreating it...");
                    const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                    if (currentEditingId === note.id) currentEditingId = res.result.id;
                    note.id = res.result.id;
                    note.dirty = false;
                    setTimeout(() => handleSearch(), 100); 
                }
            }
        } else {
            const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
            if (currentEditingId === note.id) currentEditingId = res.result.id;
            note.id = res.result.id; 
            note.dirty = false;
            setTimeout(() => handleSearch(), 100); 
        }
        saveNotes();
    } catch (error) {
        console.error("Instant push failed:", error);
    }
}

// --- Cloud Sync Status UI ---
function setCloudSyncIcon(state) {
    const icon = document.getElementById('cloudSyncIcon');
    if (!icon) return;

    if (state === 'saving') {
        icon.innerHTML = '🌧️'; 
        icon.style.color = '#3B82F6'; // Blue
        icon.title = 'Saving to Google Drive...';
    } else if (state === 'saved') {
        icon.innerHTML = '🌤️';
        icon.style.color = '#10B981'; // Green
        icon.title = 'Saved to Google Drive';
    } else if (state === 'unsaved') {
        icon.innerHTML = '☁️';
        icon.style.color = '#F59E0B'; // Orange
        icon.title = 'Unsaved changes (Press Ctrl+S)';
    } else if (state === 'error') {
        icon.innerHTML = '⛈️'; // Added a storm cloud for errors!
        icon.style.color = '#EF4444'; // Red
        icon.title = 'Error saving to Drive';
    }
}

function toggleNotebookLayout(layout) {
    currentNotebookLayout = layout;
    const container = document.getElementById('notebook-container');
    
    const gridBtn = document.getElementById('notebook-grid-btn');
    const listBtn = document.getElementById('notebook-list-btn');

    if (layout === 'grid') {
        container.className = 'notebook-grid';
        gridBtn.style.borderColor = '#3B82F6'; gridBtn.style.color = '#3B82F6'; gridBtn.style.background = 'white';
        listBtn.style.borderColor = 'transparent'; listBtn.style.color = 'var(--text-muted)'; listBtn.style.background = 'transparent';
    } else {
        container.className = 'notebook-list';
        listBtn.style.borderColor = '#3B82F6'; listBtn.style.color = '#3B82F6'; listBtn.style.background = 'white';
        gridBtn.style.borderColor = 'transparent'; gridBtn.style.color = 'var(--text-muted)'; gridBtn.style.background = 'transparent';
    }
}

function renderNotebookView() {
    const container = document.getElementById('notebook-container');
    if (!container) return;
    container.innerHTML = '';
    
    const searchInput = document.getElementById('searchInput');
    const globalQuery = searchInput ? searchInput.value : '';
    
    // Grab only notes that belong to the 'notes' quadrant
    const notebookNotes = notes.filter(n => !n.deleted && n.quadrant === 'notes' && matchesSearchQuery(n.text, globalQuery) && isProjectVisible(n));
    
    notebookNotes.forEach(note => {
        const card = document.createElement('div');
        card.className = 'notebook-card';
        card.onclick = (e) => openTaskModal(null, note.id, e);
        
        let cleanText = cleanHTMLToPlainText(note.text);
        let lines = cleanText.split('\n');
        let title = lines[0] || 'Untitled Note';
        
        // --- FIX: Parse tags on each line individually, THEN join with actual <br> tags ---
        let bodyText = lines.slice(1).map(line => parseTags(line)).join('<br>') || '';

        card.innerHTML = `
            <div class="notebook-card-title">${parseTags(title)}</div>
            <div class="notebook-card-body">${bodyText}</div>
        `;
        container.appendChild(card);
    });
}


// --- TARGET CALENDAR MIRROR SYNC ---
async function pushWeekToTargetCalendar() {
    if (!appConfig.targetCalendar) {
        return showToast("❌ Please select a Target Calendar in Settings first.");
    }
    
    const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (!savedToken || !savedToken.token || savedToken.expires_at < Date.now()) {
        return showToast("❌ Please sign in to Google first.");
    }

    const btn = document.getElementById('btnSyncTargetCal');
    const originalText = btn ? btn.innerHTML : '';
    if (btn) { btn.innerText = "Syncing..."; btn.disabled = true; }

    try {
        // 1. Determine the current week's dates
        const baseDateStr = document.getElementById('trackerDate').value;
        const [y, m, d] = baseDateStr.split('-');
        const baseDate = new Date(y, m - 1, d);
        const dayOfWeek = baseDate.getDay();
        const startOfWeek = new Date(baseDate);
        startOfWeek.setDate(baseDate.getDate() - dayOfWeek);

        const workWeekDates = new Set();
        for (let i = 0; i < 7; i++) {
            let dateIter = new Date(startOfWeek);
            dateIter.setDate(startOfWeek.getDate() + i);
            const localY = dateIter.getFullYear();
            const localM = String(dateIter.getMonth() + 1).padStart(2, '0');
            const localD = String(dateIter.getDate()).padStart(2, '0');
            workWeekDates.add(`${localY}-${localM}-${localD}`);
        }

        let syncCount = 0;

        // 2. Loop through all active notes and their time blocks
        for (let note of notes) {
            if (note.deleted || !note.timeBlocks) continue;
            
            // Get the project name for the calendar title
            let pid = note.projectId || note.projectIds?.[0] || 'p_default';
            let pObj = appConfig.projects.find(p => p.id === pid);
            let pName = pObj ? pObj.name : pid;
            let cleanTextTitle = cleanHTMLToPlainText(note.text).split('\n')[0];
            
            // Format: "[3094] Architecture Review"
            let eventSummary = `[${pName}] ${cleanTextTitle}`;

            for (let tb of note.timeBlocks) {
                // Only sync blocks that fall in the current week and have a duration
                if (workWeekDates.has(tb.date) && tb.duration > 0) {
                    
                    // Convert Quadra's decimal time into UTC ISO Strings
                    const [tY, tM, tD] = tb.date.split('-');
                    let startH = Math.floor(tb.startHour);
                    let startM = Math.round((tb.startHour - startH) * 60);
                    let startDateObj = new Date(tY, tM - 1, tD, startH, startM, 0);
                    
                    let endDec = tb.startHour + tb.duration;
                    let endH = Math.floor(endDec);
                    let endM = Math.round((endDec - endH) * 60);
                    let endDateObj = new Date(tY, tM - 1, tD, endH, endM, 0);

                    let resourceBody = {
                        summary: eventSummary,
                        start: { dateTime: startDateObj.toISOString() },
                        end: { dateTime: endDateObj.toISOString() }
                    };

                    // 3. Push to Google Calendar
                    try {
                        if (tb.targetEventId) {
                            // Update existing event if it was moved/resized
                            await gapi.client.calendar.events.patch({
                                calendarId: appConfig.targetCalendar,
                                eventId: tb.targetEventId,
                                resource: resourceBody
                            });
                        } else {
                            // Create new event and save the generated Google ID to Quadra
                            let res = await gapi.client.calendar.events.insert({
                                calendarId: appConfig.targetCalendar,
                                resource: resourceBody
                            });
                            tb.targetEventId = res.result.id;
                            note.dirty = true;
                        }
                        syncCount++;
                    } catch (err) {
                        // If the event was manually deleted on Google Calendar, it will 404. Recreate it.
                        if (err.status === 404 || (err.result && err.result.error && err.result.error.code === 404)) {
                            let res = await gapi.client.calendar.events.insert({
                                calendarId: appConfig.targetCalendar,
                                resource: resourceBody
                            });
                            tb.targetEventId = res.result.id;
                            note.dirty = true;
                            syncCount++;
                        } else {
                            console.error("Event Sync Error:", err);
                        }
                    }
                }
            }
        }

        if (syncCount > 0) {
            saveNotes();
            showToast(`✅ Synced ${syncCount} timeline blocks to Target Calendar!`);
        } else {
            showToast("No timeline blocks found to sync for this week.");
        }

    } catch (error) {
        console.error("Target Calendar Sync Failed:", error);
        showToast("❌ Failed to sync to Target Calendar.");
    } finally {
        if (btn) { btn.innerHTML = originalText; btn.disabled = false; }
    }
}

function renderArchivedProjects() {
    const container = document.getElementById('archivedProjectsList');
    if (!container) return;
    container.innerHTML = '';
    
    const archived = appConfig.projects.filter(p => p.archived);
    if (archived.length === 0) {
        container.innerHTML = '<p style="font-size: 13px; color: var(--text-muted);">No archived projects.</p>';
        return;
    }
    
    archived.forEach(proj => {
        const div = document.createElement('div');
        div.className = 'form-row-inline';
        div.style.marginBottom = '8px';
        div.innerHTML = `
            <span style="flex:1; font-weight:600; font-size: 13px; color: var(--text-main);">${proj.name}</span>
            <button class="btn btn-outline" style="padding: 4px 10px; font-size: 12px;" onclick="unarchiveProject('${proj.id}')">Restore</button>
        `;
        container.appendChild(div);
    });
}

function unarchiveProject(id) {
    const proj = appConfig.projects.find(p => p.id === id);
    if (proj) {
        proj.archived = false;
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        renderProjectTabs();
        renderArchivedProjects();
        handleSearch();
        showToast(`Project '${proj.name}' restored.`);
    }
}

let currentEditingProjectId = null;

function openProjectModal(projectId) {
    const proj = appConfig.projects.find(p => p.id === projectId);
    if (!proj) return;
    
    currentEditingProjectId = projectId;
    const modal = document.getElementById('projectModal');
    const input = document.getElementById('editProjectNameInput');
    const btnDelete = document.getElementById('btnDeleteProject');
    const btnArchive = document.getElementById('btnArchiveProject');
    
    input.value = proj.name;
    
    // Hide dangerous actions for the Default project
    if (proj.id === 'p_default') {
        btnDelete.style.display = 'none';
        btnArchive.style.display = 'none';
    } else {
        btnDelete.style.display = 'inline-block';
        btnArchive.style.display = 'inline-block';
    }
    
    modal.style.display = 'flex';
    setTimeout(() => {
        input.focus();
        input.select();
    }, 100);
}

function closeProjectModal() {
    currentEditingProjectId = null;
    document.getElementById('projectModal').style.display = 'none';
}

function saveProjectModal() {
    if (!currentEditingProjectId) return;
    
    const input = document.getElementById('editProjectNameInput');
    const newName = input.value.trim();
    if (!newName) return showToast("Project name cannot be empty.");
    
    const proj = appConfig.projects.find(p => p.id === currentEditingProjectId);
    if (proj) {
        proj.name = newName;
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        renderProjectTabs();
    }
    
    closeProjectModal();
}

function archiveProjectFromModal() {
    if (!currentEditingProjectId || currentEditingProjectId === 'p_default') return;
    
    const proj = appConfig.projects.find(p => p.id === currentEditingProjectId);
    if (proj) {
        proj.archived = true;
        proj.visible = false;
        
        const unarchived = appConfig.projects.filter(p => !p.archived);
        if (!unarchived.some(p => p.visible) && unarchived.length > 0) {
            unarchived[0].visible = true;
        }
        
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        renderProjectTabs();
        handleSearch();
        showToast(`Project '${proj.name}' archived.`);
    }
    closeProjectModal();
}

function deleteProjectFromModal() {
    if (!currentEditingProjectId || currentEditingProjectId === 'p_default') return;
    
    const proj = appConfig.projects.find(p => p.id === currentEditingProjectId);
    if (!proj) return;
    
    if (confirm(`Are you sure you want to delete '${proj.name}'? All associated tasks will be moved to your primary project.`)) {
        notes.forEach(n => { 
            if (n.projectId === proj.id || (n.projectIds && n.projectIds.includes(proj.id))) {
                n.projectId = 'p_default';
                n.projectIds = ['p_default'];
                n.dirty = true;
            }
        });
        saveNotes();
        
        appConfig.projects = appConfig.projects.filter(p => p.id !== proj.id);
        const unarchived = appConfig.projects.filter(p => !p.archived);
        if (!unarchived.some(p => p.visible) && unarchived.length > 0) unarchived[0].visible = true;
        
        localStorage.setItem('quadra_config', JSON.stringify(appConfig));
        renderProjectTabs();
        handleSearch();
        closeProjectModal();
    }
}

// --- SORTING ENGINE ---
function toggleSortMenu(quadrant, event) {
    event.stopPropagation();
    
    // Close other open menus
    document.querySelectorAll('.dropdown-content').forEach(menu => {
        if (menu.id !== `sort-menu-${quadrant}`) {
            menu.classList.remove('show');
        }
    });

    const menu = document.getElementById(`sort-menu-${quadrant}`);
    if (menu) {
        menu.classList.toggle('show');
        
        // Highlight active sort option
        let pref = appConfig.sortPrefs[quadrant] || (quadrant === 'notes' ? 'created_desc' : 'due_asc');
        Array.from(menu.children).forEach(child => {
            if (child.getAttribute('onclick').includes(pref)) {
                child.classList.add('active-sort');
            } else {
                child.classList.remove('active-sort');
            }
        });
    }
}

// Close menu when clicking outside
document.addEventListener('click', (e) => {
    if (!e.target.closest('.custom-dropdown')) {
        document.querySelectorAll('.dropdown-content.show').forEach(m => m.classList.remove('show'));
    }
});

function changeSort(quadrant, val) {
    if (!appConfig.sortPrefs) appConfig.sortPrefs = {};
    appConfig.sortPrefs[quadrant] = val;
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
    
    const menu = document.getElementById(`sort-menu-${quadrant}`);
    if (menu) menu.classList.remove('show');
    
    handleSearch(); // Trigger re-render with new sort
}

// --- Modal Extensions: Time Tracking & Deletion ---
function toggleTimeTrackPanel() {
    const content = document.getElementById('taskModalContent');
    const rightPane = document.getElementById('taskModalRightPane');
    
    if (content.classList.contains('time-panel-open')) {
        content.classList.remove('time-panel-open');
        setTimeout(() => rightPane.style.display = 'none', 300);
    } else {
        rightPane.style.display = 'flex';
        setTimeout(() => content.classList.add('time-panel-open'), 10);
    }
}

function renderQuickTimeLogs() {
    const tbody = document.getElementById('quickLogTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';
    
    if (!currentEditingId) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:var(--text-muted); padding: 12px;">Save task first</td></tr>';
        return;
    }
    
    const note = notes.find(n => n.id === currentEditingId);
    if (!note || !note.timeBlocks || note.timeBlocks.length === 0) {
        tbody.innerHTML = '<tr><td colspan="3" style="text-align:center; color:var(--text-muted); padding: 12px;">No time logged</td></tr>';
        return;
    }
    
    // Sort blocks by date descending
    let sortedBlocks = [...note.timeBlocks].sort((a,b) => b.date.localeCompare(a.date));
    
    sortedBlocks.forEach(tb => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td style="padding: 8px;">${tb.date}</td>
            <td style="padding: 8px; text-align: center; font-weight: 600;">${tb.duration}</td>
            <td style="padding: 8px; text-align: center;">
                <button class="action-btn delete-btn" style="font-size:16px;" onclick="removeTimeBlock('${tb.blockId}')" title="Delete record">×</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function addManualTimeBlock() {
    if (!currentEditingId) return showToast("Please save the task first to log time.");
    
    const dateVal = document.getElementById('quickLogDate').value;
    const hoursVal = parseFloat(document.getElementById('quickLogHours').value);
    
    if (!dateVal || isNaN(hoursVal) || hoursVal <= 0) {
        return showToast("Please enter a valid date and hours.");
    }
    
    const note = notes.find(n => n.id === currentEditingId);
    if (note) {
        if (!note.timeBlocks) note.timeBlocks = [];
        note.timeBlocks.push({
            blockId: 'b_' + Date.now().toString() + Math.floor(Math.random() * 1000),
            date: dateVal,
            startHour: 9, // Native manual entries default to 9am to avoid overlap issues
            duration: hoursVal
        });
        note.dirty = true;
        saveNotes();
        renderQuickTimeLogs();
        document.getElementById('quickLogHours').value = '';
    }
}

function removeTimeBlock(blockId) {
    if (!currentEditingId) return;
    const note = notes.find(n => n.id === currentEditingId);
    if (note && note.timeBlocks) {
        note.timeBlocks = note.timeBlocks.filter(b => b.blockId !== blockId);
        note.dirty = true;
        saveNotes();
        renderQuickTimeLogs();
    }
}

function deleteTaskFromModal() {
    if (currentEditingId && confirm("Are you sure you want to delete this task/event?")) {
        deleteTask(currentEditingId);
        closeTaskModal();
        if (currentLayout === 'tracker') renderTrackerTimeline();
    }
}

// --- PDF EXPORT ENGINE ---
function exportNoteToPDF() {
    if (!currentEditingId) {
        return showToast("Please save the task first before exporting.");
    }

    const note = notes.find(n => n.id === currentEditingId);
    if (!note) return;

    // 1. Extract Data
    let titleText = cleanHTMLToPlainText(note.text).split('\n')[0];
    let bodyHTML = document.getElementById('taskInfoInput').innerHTML;
    let pid = note.projectId || note.projectIds?.[0] || 'p_default';
    let pObj = appConfig.projects.find(p => p.id === pid);
    let projectName = pObj ? pObj.name : 'Default';

    // 2. Build the Print Template (Invisible to the user)
    const printElement = document.createElement('div');
    printElement.style.padding = '30px';
    printElement.style.fontFamily = 'Inter, Helvetica, Arial, sans-serif';
    printElement.style.color = '#0F172A';
    
    printElement.innerHTML = `
        <div style="border-bottom: 3px solid #4F46E5; padding-bottom: 12px; margin-bottom: 24px;">
            <div style="font-size: 10px; font-weight: 700; color: #4F46E5; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px;">Quadra Export</div>
            <h1 style="margin: 0; font-size: 24px; font-weight: 800; color: #1E293B; line-height: 1.3;">${parseTags(titleText)}</h1>
            <div style="display: flex; gap: 16px; font-size: 12px; font-weight: 500; color: #64748B; margin-top: 12px;">
                ${note.dueDate ? `<span><b>Due:</b> ${note.dueDate}</span>` : ''} 
                <span><b>Project:</b> ${projectName}</span>
                <span><b>Status:</b> ${note.status.toUpperCase()}</span>
            </div>
        </div>
        <div style="font-size: 13px; line-height: 1.6; color: #334155;">
            ${bodyHTML !== '<br>' && bodyHTML !== '' ? bodyHTML : '<i>No description provided.</i>'}
        </div>
    `;

    // 3. Configure html2pdf settings
    const opt = {
        margin:       0.5,
        filename:     `${titleText.replace(/[^a-z0-9]/gi, '_').toLowerCase().substring(0, 30)}_export.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, letterRendering: true },
        jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
    };

    // 4. Generate and Download
    showToast("Generating PDF...");
    html2pdf().set(opt).from(printElement).save().then(() => {
        showToast("✓ PDF Downloaded");
    });
}