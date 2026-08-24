let version = '3.30';
let appConfig = JSON.parse(localStorage.getItem('quadra_config')) || {};
let isDocMode = false;
let tokenHeartbeatId = null;

if (!appConfig.ignoreKeywords) appConfig.ignoreKeywords = 'out of office, ooo, away, vacation, holiday';
if (!appConfig.calSource) appConfig.calSource = 'google';
if (!appConfig.icsUrl) appConfig.icsUrl = '';
if (!appConfig.viewsEnabled) appConfig.viewsEnabled = { grid: true, kanban: true, overdue: true, tracker: true };
if (!appConfig.defaultView) appConfig.defaultView = 'grid';
if (!appConfig.primaryTz) appConfig.primaryTz = 'local';
if (!appConfig.secondaryTz) appConfig.secondaryTz = 'none';
if (!appConfig.quadrantOrder) appConfig.quadrantOrder = ['q1', 'q2', 'q3', 'q4', 'tray-inbox', 'tray-calendar', 'tray-closed'];
if (!appConfig.quadrantWidths) appConfig.quadrantWidths = {};

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
                deleted INTEGER DEFAULT 0
            );
        `);
        console.log("Created fresh SQLite database in memory.");
    }
}

// --- Sync Memory Notes into SQLite Database ---
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
            deleted INTEGER DEFAULT 0
        );
    `);

    // Prepare a statement to insert or replace task records
    const stmt = db.prepare(`
        INSERT OR REPLACE INTO tasks (id, text, quadrant, status, dueDate, timeBlocks, deleted)
        VALUES (?, ?, ?, ?, ?, ?, ?);
    `);

    notes.forEach(note => {
        stmt.run([
            note.id.toString(),
            note.text || '',
            note.quadrant || 'inbox',
            note.status || 'active',
            note.dueDate || null,
            JSON.stringify(note.timeBlocks || []),
            note.deleted ? 1 : 0
        ]);
    });

    stmt.free();
}

// --- Google Drive AppData Sync ---
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
        
        let url;
        let method;
        let metadata;

        if (driveFileId) {
            url = `https://www.googleapis.com/upload/drive/v3/files/${driveFileId}?uploadType=multipart`;
            method = 'PATCH';
            metadata = { name: 'quadra.sqlite' };
        } else {
            url = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart';
            method = 'POST';
            metadata = { name: 'quadra.sqlite' };
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

const SCOPES = 'https://www.googleapis.com/auth/tasks https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/drive.file';

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
    const btnGrid = document.getElementById('btnGrid');
    const btnKanban = document.getElementById('btnKanban');
    const btnTracker = document.getElementById('btnTracker');
    const btnOverdue = document.getElementById('btnOverdue');
    
    if (btnGrid) btnGrid.style.display = appConfig.viewsEnabled.grid ? '' : 'none';
    if (btnKanban) btnKanban.style.display = appConfig.viewsEnabled.kanban ? '' : 'none';
    if (btnTracker) btnTracker.style.display = appConfig.viewsEnabled.tracker ? '' : 'none';
    if (btnOverdue) btnOverdue.style.display = appConfig.viewsEnabled.overdue ? '' : 'none';
    
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
    
    const btnGrid = document.getElementById('btnGrid');
    const btnKanban = document.getElementById('btnKanban');
    const btnOverdue = document.getElementById('btnOverdue');
    const btnTracker = document.getElementById('btnTracker');

    if (btnGrid) btnGrid.classList.toggle('active', layout === 'grid');
    if (btnKanban) btnKanban.classList.toggle('active', layout === 'kanban');
    if (btnOverdue) btnOverdue.classList.toggle('active', layout === 'overdue');
    if (btnTracker) btnTracker.classList.toggle('active', layout === 'tracker');
    
    if(settingsView) settingsView.style.display = 'none';
    if(matrix) matrix.style.display = 'none';
    if(tracker) tracker.style.display = 'none';
    if(overdue) overdue.style.display = 'none';

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
    document.getElementById('configIgnoreKeywords').value = appConfig.ignoreKeywords || '';
    document.getElementById('configCalSource').value = appConfig.calSource || 'google';
    document.getElementById('configIcsUrl').value = appConfig.icsUrl || '';
    
    document.getElementById('configPrimaryTz').value = appConfig.primaryTz || 'local';
    document.getElementById('configSecondaryTz').value = appConfig.secondaryTz || 'none';
    
    document.getElementById('configDefaultView').value = appConfig.defaultView || 'grid';
    document.getElementById('configViewGrid').checked = appConfig.viewsEnabled.grid !== false;
    document.getElementById('configViewKanban').checked = appConfig.viewsEnabled.kanban !== false;
    document.getElementById('configViewTracker').checked = appConfig.viewsEnabled.tracker !== false;
    document.getElementById('configViewOverdue').checked = appConfig.viewsEnabled.overdue !== false;
    
    toggleCalSourceFields(appConfig.calSource || 'google');
    loadCalendars();
    renderScheduleSettings();
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
    const d = new Date(dateInput.value);
    d.setDate(d.getDate() + offset);
    dateInput.value = d.toLocaleDateString('en-CA').split('T')[0];
    renderTrackerTimeline();
}

function goToToday() {
    const dateInput = document.getElementById('trackerDate');
    dateInput.value = new Date().toLocaleDateString('en-CA').split('T')[0];
    renderTrackerTimeline();
}

// --- NEW: Dedicated Palette Renderer ---
// --- NEW: Dedicated Palette Renderer (With Planned Indicator) ---
// --- NEW: Dedicated Palette Renderer (Multi-Day Architecture) ---
// --- NEW: Dedicated Palette Renderer (Smart Planned Indicator) ---
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
    
    let paletteNotes = notes.filter(n => !n.deleted && n.status !== 'closed' && matchesSearchQuery(n.text, effectivePaletteQuery) && !n.eventId);
    
    const dueToggle = document.getElementById('dueFilterToggle');
    if (dueToggle && dueToggle.checked) {
        paletteNotes = paletteNotes.filter(n => n.dueDate);
        paletteNotes.sort((a, b) => a.dueDate.localeCompare(b.dueDate));
    }
    
    // Calculate today's date string once to use for both planned & overdue checks
    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];

    paletteNotes.forEach(note => {
        const el = document.createElement('div');
        el.className = 'note'; 
        el.style.marginBottom = '8px'; 
        el.style.cursor = 'grab';
        el.draggable = true;
        el.ondragstart = (e) => e.dataTransfer.setData('text/plain', note.id);
        
        // --- UPGRADED: Only show PLANNED if there is a block scheduled for today or the future ---
        const isPlannedOnCalendar = note.timeBlocks && note.timeBlocks.some(block => block.date >= todayStr);
        
        if (isPlannedOnCalendar) {
            el.style.backgroundColor = '#F8FAFC';
            el.style.borderColor = '#E2E8F0';
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
        
        if (note.dueDate) {
            let plannedBadge = isPlannedOnCalendar 
                ? `<span style="margin-left: auto; background: #E0F2FE; color: #0369A1; border: 1px solid #BAE6FD; padding: 2px 6px; border-radius: 4px; font-size: 9px; font-weight: 800; letter-spacing: 0.5px;">🕒 PLANNED</span>` 
                : '';
                
            contentWrapper.innerHTML += `
                <div style="font-size:12px; color:var(--brand-primary); margin-top:8px; font-weight:500; display: flex; align-items: center;">
                    🗓️ ${note.dueDate.split('T')[0]}
                    ${plannedBadge}
                </div>`;
        }
        
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
            const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
            if (!isCalendarEvent && !matchesSearchQuery(note.text, globalQuery)) return;

            // --- NEW: Loop through timeBlocks array ---
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
                    
                    let displayTitle = cleanHTMLToPlainText(note.text).split('\n')[0].substring(0, 40);
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
                        
                        let displayTitle = cleanHTMLToPlainText(note.text).split('\n')[0].substring(0, 40);
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
        if (note.deleted || note.eventId) return;
        (note.timeBlocks || []).forEach(tb => {
            if (weekDateKeys.has(tb.date) && !countedIds.has(`${note.id}-${tb.blockId}`)) {
                actualWeekly += roundToQuarterHour(tb.duration || 1.0);
                countedIds.add(`${note.id}-${tb.blockId}`);
            }
        });
    });

    let actualDaily = 0;
    notes.forEach(note => {
        if (note.deleted || note.eventId) return;
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
        renderTrackerPalette();
    }
}

function renderOverdueTasksPage() {
    const container = document.getElementById('overdue-list-container');
    container.innerHTML = '';
    
    const overdueNotes = notes.filter(n => {
        if (n.deleted || n.status === 'closed' || !n.dueDate || n.eventId) return false;
        return n.dueDate.split('T')[0] < todayStr;
    });

    overdueNotes.sort((a, b) => a.dueDate.localeCompare(b.dueDate));

    document.getElementById('overdue-count-badge').innerText = overdueNotes.length;

    if (overdueNotes.length === 0) {
        container.innerHTML = `<p style="color: var(--text-muted); text-align: center; padding: 40px;">🎉 No overdue tasks! All caught up.</p>`;
        return;
    }

    overdueNotes.forEach(note => {
        const card = document.createElement('div');
        card.className = 'overdue-card';
        card.onclick = (e) => openTaskModal(null, note.id, e);
        card.style.cursor = 'pointer';

        let title = cleanHTMLToPlainText(note.text).split('\n')[0];
        let dueDate = note.dueDate.split('T')[0];

        card.innerHTML = `
            <div style="display: flex; flex-direction: column; gap: 4px;">
                <strong style="font-size: 15px; color: #9F1239;">${parseTags(title)}</strong>
                <span style="font-size: 12px; color: var(--text-muted);">Due: ${dueDate} (${note.quadrant.toUpperCase()})</span>
            </div>
            <button class="btn" style="padding: 6px 12px; font-size: 12px;" onclick="completeTask('${note.id}'); event.stopPropagation(); renderOverdueTasksPage();">Complete ✓</button>
        `;
        container.appendChild(card);
    });
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

    openTaskModal('inbox', null, ev, { date: dateStr, startHour: dropHour });
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
        if (taskModal && taskModal.style.display === 'flex') {
            closeTaskModal();
        }
        if (shortcutsModal && shortcutsModal.style.display === 'flex') {
            closeShortcutsModal();
        }
    } else if (!isEditingText) {
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
        } else if (e.key === '`' || e.key === '~') {
            e.preventDefault();
            const allViews = ['grid', 'kanban', 'overdue', 'tracker'];
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
    const note = notes.find(n => n.id === currentEditingId);
    if (note) {
        if (note.status === 'closed') {
            note.status = 'active';
            note.quadrant = note.quadrant === 'closed' ? 'inbox' : note.quadrant;
        } else {
            note.status = 'closed';
            note.quadrant = 'closed';
        }
        note.dirty = true;
        saveNotes();
        handleSearch();
        closeTaskModal();
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

    // Reset Doc Mode state when opening a modal
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
        completeBtn.style.display = 'none';
        
        pendingTimelineContext = timelineContext || null;
    }
    modal.style.display = 'flex'; 
    setTimeout(() => titleInput.focus(), 100);
}

function closeTaskModal() { 
    clearAutoSaveInterval();
    document.getElementById('taskModal').style.display = 'none'; 
    pendingTimelineContext = null; 
}

// --- NEW: Save Modal (Multi-Day Architecture) ---
function saveTaskModal() {
    clearAutoSaveInterval();
    const titleText = document.getElementById('taskTitleInput').innerHTML; 
    const infoText = document.getElementById('taskInfoInput').innerHTML; 
    const dueDate = document.getElementById('taskDueDate').value;

    if ((!titleText || titleText === '<br>') && (!infoText || infoText === '<br>')) return closeTaskModal();

    const fullText = titleText + (infoText && infoText !== '<br>' ? ('\n' + infoText) : '');

    let newTimeBlock = null;
    if (pendingTimelineContext && pendingTimelineContext.startHour !== undefined) {
        newTimeBlock = {
            blockId: 'b_' + Date.now().toString() + Math.floor(Math.random() * 1000),
            date: pendingTimelineContext.date,
            startHour: roundToQuarterHour(pendingTimelineContext.startHour),
            duration: 1.0
        };
    }

    if (currentEditingId) { 
        const note = notes.find(n => n.id === currentEditingId); 
        if (note) { 
            note.text = fullText; 
            note.dueDate = dueDate || null;
            if (newTimeBlock) {
                if(!note.timeBlocks) note.timeBlocks = [];
                note.timeBlocks.push(newTimeBlock);
            }
            note.dirty = true; 
            saveNotes(); 
            syncSingleTask(note.id); // <-- NEW
            handleSearch(); 
        } 
    } else { 
        let targetQuad = currentAddingQuadrant || 'inbox';
        let newNoteId = Date.now().toString(); // Extract ID to variable
        notes.push({ 
            id: newNoteId, 
            text: fullText, 
            quadrant: targetQuad, 
            status: 'active', 
            dueDate: dueDate || null,
            timeBlocks: newTimeBlock ? [newTimeBlock] : [],
            dirty: true, 
            deleted: false, 
            eventId: null,
            syncFailed: false
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

function exportTimesheet() {
    const baseDateStr = document.getElementById('trackerDate').value;
    const [y, m, d] = baseDateStr.split('-');
    const baseDate = new Date(y, m - 1, d);
    let datesToExport = [];
    
    if (currentTrackerMode === 'day') {
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

    let plainTextLog = `Calendar Report\n`;
    
    datesToExport.forEach(dateStr => {
        let dailyLog = "";
        let itemFound = false;
        notes.forEach(note => {
            if (note.deleted || note.status === 'closed') return;
            if (note.dueDate === dateStr) {
                itemFound = true;
                let cleanText = cleanHTMLToPlainText(note.text);
                const tags = cleanText.match(/(#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+)/g);
                let identifier = tags ? tags.join(' ') : cleanText.substring(0, 40).replace(/\n/g, ' ') + '...';
                dailyLog += `- [${identifier}] ${decToTime(note.dueTime)} (${note.dueDuration}h)\n`;
            }
        });
        
        if (itemFound || currentTrackerMode === 'day') {
             plainTextLog += `\nDate: ${dateStr}\n------------------------\n`;
             if(itemFound) plainTextLog += dailyLog;
             else plainTextLog += "No tasks scheduled.\n";
        }
    });

    navigator.clipboard.writeText(plainTextLog).then(() => showToast("Calendar copied to clipboard!")).catch(err => showToast("Copy failed"));
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
        const select = document.getElementById('configCalendar');
        const currentVal = appConfig.calendarId;
        select.innerHTML = '<option value="">-- Select a Calendar --</option>';
        calendars.forEach(cal => {
            const opt = document.createElement('option');
            opt.value = cal.id; opt.innerText = cal.summary;
            if(cal.id === currentVal) opt.selected = true;
            select.appendChild(opt);
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
    appConfig.ignoreKeywords = document.getElementById('configIgnoreKeywords').value.trim();
    appConfig.calSource = document.getElementById('configCalSource').value;
    appConfig.icsUrl = document.getElementById('configIcsUrl').value.trim();
    
    appConfig.primaryTz = document.getElementById('configPrimaryTz').value;
    appConfig.secondaryTz = document.getElementById('configSecondaryTz').value;
    
    const calendarSelect = document.getElementById('configCalendar');
    if (calendarSelect && calendarSelect.value) {
        appConfig.calendarId = calendarSelect.value;
    }

    appConfig.defaultView = document.getElementById('configDefaultView').value;
    appConfig.viewsEnabled = {
        grid: document.getElementById('configViewGrid').checked,
        kanban: document.getElementById('configViewKanban').checked,
        tracker: document.getElementById('configViewTracker').checked,
        overdue: document.getElementById('configViewOverdue').checked
    };
    
    if (!appConfig.viewsEnabled[appConfig.defaultView]) {
        const firstEnabled = Object.keys(appConfig.viewsEnabled).find(k => appConfig.viewsEnabled[k]);
        if (firstEnabled) {
            appConfig.defaultView = firstEnabled;
            const selectEl = document.getElementById('configDefaultView');
            if (selectEl) selectEl.value = firstEnabled;
        } else {
            appConfig.viewsEnabled.grid = true;
            appConfig.defaultView = 'grid';
            const selectEl = document.getElementById('configDefaultView');
            if (selectEl) selectEl.value = 'grid';
            const checkEl = document.getElementById('configViewGrid');
            if (checkEl) checkEl.checked = true;
        }
    }
    
    localStorage.setItem('quadra_config', JSON.stringify(appConfig));
    
    const rows = document.querySelectorAll('.schedule-row');
    appSchedule = Array.from(rows).map(row => {
        const startStr = row.querySelector('.sched-start').value.split(':');
        const endStr = row.querySelector('.sched-end').value.split(':');
        
        let startH = parseInt(startStr[0] || 0) + (parseInt(startStr[1] || 0) / 60);
        let endH = parseInt(endStr[0] || 0) + (parseInt(endStr[1] || 0) / 60);
        
        if (endH <= startH) {
            endH += 24;
        }

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

// --- Updated Quick Tags Function (With AND/OR Logic, Counts & Sorting) ---
// --- Updated Quick Tags Function (With Real-Time Filtering & Smart Append) ---
// --- Updated Quick Tags Function (Case-Insensitive Tag Merging) ---
function updateQuickTags() {
    const tagsBar = document.getElementById('quick-tags-bar');
    const searchInput = document.getElementById('searchInput');
    if (!tagsBar || !searchInput) return;
    
    let tagCounts = new Map();
    // Added 'i' flag to regex to ensure lowercase hex colors are also caught safely
    const hexColorRegex = /^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/i; 
    
    notes.forEach(note => {
        if (note.deleted || note.eventId) return;
        
        let tempDiv = document.createElement('div');
        tempDiv.innerHTML = note.text || '';
        let safePlainText = tempDiv.textContent || tempDiv.innerText || '';
        
        // Use a safe boundary regex and extract just the tag using an exec loop
        const regex = /(^|[\s\(\)\[\]\{\}>;"',\.|])(#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+)/g;
        let matches = [];
        let m;
        while ((m = regex.exec(safePlainText)) !== null) {
            matches.push(m[2]);
        }
        
        if (matches.length > 0) {
            // Convert all tags to lowercase BEFORE deduplicating
            const lowerCaseMatches = matches.map(tag => tag.toLowerCase());
            let uniqueMatches = [...new Set(lowerCaseMatches)];
            
            uniqueMatches.forEach(tag => {
                if (hexColorRegex.test(tag)) return;
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
        tagFilter = lastWord.toLowerCase(); // Ensure active typing filter is also lowercase
    }
    
    // --- Inject AND / OR logic buttons if search is active ---
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
    
    // Convert to Array and Sort: Descending by count, then alphabetically
    let sortedTags = Array.from(tagCounts.entries()).sort((a, b) => {
        if (b[1] !== a[1]) {
            return b[1] - a[1]; 
        }
        return a[0].localeCompare(b[0]); 
    });
    
    // Apply the real-time typing filter
    if (tagFilter) {
        sortedTags = sortedTags.filter(([tag, count]) => tag.startsWith(tagFilter));
    }
    
    // Populate the bar with clickable tags
    sortedTags.forEach(([tag, count]) => {
        let btn = document.createElement('button');
        btn.className = 'filter-tag' + (tag.startsWith('@') ? ' person-filter' : '');
        btn.innerText = `${tag} (${count})`;
        
        btn.onclick = () => {
            let currentWords = searchInput.value.split(/\s+/);
            
            if (currentWords[currentWords.length - 1].startsWith('#') || currentWords[currentWords.length - 1].startsWith('@') || currentWords[currentWords.length - 1] === '') {
                currentWords.pop();
            }
            
            currentWords.push(tag);
            searchInput.value = currentWords.join(' ') + ' ';
            
            searchInput.focus();
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


function renderNotes(searchQuery = '') {
    ['q1', 'q2', 'q3', 'q4', 'inbox', 'calendar', 'closed'].forEach(q => { 
        const el = document.getElementById(`list-${q}`); 
        if (el) el.innerHTML = ''; 
    });
    
    let filteredNotes = notes.filter(note => {
        if (note.deleted) return false;
        const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
        if (isCalendarEvent) return true;
        return matchesSearchQuery(note.text, searchQuery);
    });

    filteredNotes.sort((a, b) => { if (!a.dueDate && !b.dueDate) return 0; if (!a.dueDate) return -1; if (!b.dueDate) return 1; return a.dueDate.localeCompare(b.dueDate); });

    let filteredCounts = { q1: 0, q2: 0, q3: 0, q4: 0, inbox: 0, calendar: 0, closed: 0 };

    filteredNotes.forEach(note => {
        let targetListId = `list-${note.quadrant}`;
        if (note.eventId) {
            targetListId = 'list-calendar';
        }
        const list = document.getElementById(targetListId);
        if (list) {
            const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
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

            const noteEl = document.createElement('div'); noteEl.className = 'note' + (note.status === 'closed' ? ' closed-note' : '');
            if (note.status === 'active' && !note.eventId) { 
                noteEl.draggable = true; 
                noteEl.ondragstart = (e) => {
                    e.stopPropagation();
                    e.dataTransfer.setData('text/plain', note.id);
                }; 
            }

            const contentWrapper = document.createElement('div'); contentWrapper.className = 'note-content-wrapper';
            
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

            // --- REVISED: Actions Div Logic ---
            if (!note.eventId) {
                // If it's closed, we still show Restore and Delete buttons
                if (note.status !== 'active') {
                    const actionsDiv = document.createElement('div'); 
                    actionsDiv.className = 'note-actions';
                    actionsDiv.innerHTML = `<button class="action-btn restore-btn" onclick="restoreTask('${note.id}')">↺</button><button class="action-btn delete-btn" onclick="deleteTask('${note.id}')">×</button>`;
                    noteEl.append(actionsDiv);
                }
                // (Active tasks no longer receive an actionsDiv, stripping the complete & sync elements)
            } else {
                // Calendar events still show the Delete button
                const actionsDiv = document.createElement('div'); 
                actionsDiv.className = 'note-actions';
                actionsDiv.innerHTML = `<button class="action-btn delete-btn" onclick="deleteTask('${note.id}')" title="Remove event">×</button>`;
                noteEl.append(actionsDiv);
            }
            // ----------------------------------
            
            list.appendChild(noteEl);
        }
    });

    document.getElementById('badge-closed').innerText = notes.filter(n => !n.deleted && n.status === 'closed' && !n.eventId).length;
    document.getElementById('badge-inbox').innerText = filteredCounts.inbox;
    document.getElementById('badge-calendar').innerText = filteredCounts.calendar;
    
    ['q1', 'q2', 'q3', 'q4'].forEach(q => {
        const badge = document.getElementById(`badge-${q}`);
        if (badge) badge.innerText = filteredCounts[q];
    });
    
    if (currentLayout === 'tracker') renderTrackerTimeline();
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
    const matrixContainer = document.getElementById('matrix');
    
    if (appConfig.quadrantOrder && appConfig.quadrantOrder.length > 0) {
        appConfig.quadrantOrder.forEach(id => {
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

    // --- NEW: Load the saved state for the Due Only toggle ---
    // Using !== 'false' ensures it defaults to true on the very first visit
    const savedDueFilter = localStorage.getItem('quadra_due_filter') !== 'false';
    const dueToggle = document.getElementById('dueFilterToggle');

    if (dueToggle) {
        dueToggle.checked = savedDueFilter;
        
        // If it defaulted to true on the first visit, trigger the filter immediately
        if (savedDueFilter && !localStorage.getItem('quadra_due_filter')) {
            toggleDueFilter();
        }
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
    let ignoredCount = 0;

    try {
        document.getElementById('sync-banner').style.display = 'block';

        if (appConfig.calSource === 'outlook' && appConfig.icsUrl) {
            let fetchUrl = appConfig.icsUrl;
            
            const res = await fetch(fetchUrl);
            if (!res.ok) throw new Error("Failed to fetch ICS feed via Apps Script bridge");
            
            const icsText = await res.text();
            // Relying on global appConfig.primaryTz in lib
            const icsEvents = parseICS(icsText, dateStr, ignoreKeywords, appConfig.primaryTz);

            icsEvents.forEach(ev => {
                if (notes.find(n => n.eventId === ev.id)) return;
                notes.push({
                    id: Date.now().toString() + Math.random(),
                    eventId: ev.id,
                    text: `${ev.summary} #meeting`,
                    quadrant: 'q2',
                    status: 'active',
                    dirty: false,
                    deleted: false,
                    dueDate: dateStr,
                    dueTime: ev.startHour,
                    dueDuration: ev.duration,
                    syncFailed: false
                });
                importedCount++;
            });

        } else {
            const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
            if (!savedToken || !savedToken.token || savedToken.expires_at < Date.now()) {
                document.getElementById('sync-banner').style.display = 'none';
                return showToast("Please sign in to Google first.");
            }
            if (!appConfig.calendarId) {
                document.getElementById('sync-banner').style.display = 'none';
                return showToast("Please select a Google calendar in Settings first.");
            }
            
            const [y, m, day] = dateStr.split('-');
            const timeMin = new Date(y, m-1, day, 0, 0, 0).toISOString();
            const timeMax = new Date(y, m-1, day, 23, 59, 59).toISOString();

            const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(appConfig.calendarId)}/events?timeMin=${timeMin}&timeMax=${timeMax}&singleEvents=true&orderBy=startTime`, {
                headers: { 'Authorization': `Bearer ${savedToken.token}` }
            });
            const data = await res.json();
            if(data.error) throw new Error(data.error.message);

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
                const rawDuration = rawEndHour - rawStartHour;

                const startHour = roundToQuarterHour(rawStartHour);
                const duration = Math.max(0.25, roundToQuarterHour(rawDuration));
                
                if (notes.find(n => n.eventId === event.id)) return; 
                
                const newNote = {
                    id: Date.now().toString() + Math.random(),
                    eventId: event.id, 
                    text: `${event.summary || 'Meeting'} #meeting`,
                    quadrant: 'q2', 
                    status: 'active',
                    dirty: false, 
                    deleted: false,
                    dueDate: dateStr,
                    dueTime: startHour,
                    dueDuration: duration,
                    syncFailed: false
                };
                
                notes.push(newNote);
                importedCount++;
            });
        }
        
        document.getElementById('sync-banner').style.display = 'none';
        if (importedCount > 0 || ignoredCount > 0) {
            saveNotes();
            renderTrackerTimeline(); 
            showToast(`${importedCount} imported, ${ignoredCount} ignored.`);
        } else {
            showToast("No new meetings found for this date.");
        }
        
    } catch(e) {
        console.error("Calendar Sync Error:", e);
        document.getElementById('sync-banner').style.display = 'none';
        showToast("Failed to fetch calendar events. Check your network or Apps Script Web App URL.");
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
        const GAPI_LIST_NAMES = { 'inbox': 'Quadra: Inbox', 'q1': 'Quadra: Do First', 'q2': 'Quadra: Schedule', 'q3': 'Quadra: Delegate', 'q4': 'Quadra: Later', 'closed': 'Quadra: Completed' };
        let gapiListIds = { inbox: null, q1: null, q2: null, q3: null, q4: null, closed: null };
        
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
            // If the task was updated on Google since last sync, pull it and overwrite local
            // UNLESS the user is actively typing in this exact task's modal right now.
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
                    sn.dueDate = remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null; 
                    sn.dirty = false; 
                    sn.syncFailed = false;
                }
                delete remoteTaskMap[sn.id];
                if (sn.tempId) delete remoteTaskMap[sn.tempId];
                continue; // Skip the push logic!
            }

            // --- DELETIONS ---
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
            
            // --- PUSH LOCAL CHANGES ---
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
                    // User is actively editing a conflict. Quadra wins this specific edge case.
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
                
                syncSnapshot.push({ 
                    id: remoteObj.task.id, text: fullText, quadrant: remoteObj.quadKey, 
                    status: remoteObj.task.status === 'completed' ? 'closed' : 'active', 
                    dueDate: remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null, 
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
        
        if (syncBanner) syncBanner.style.display = 'none';
        showToast("✓ Successfully synced with Google Tasks!");
        
    } catch (e) {
        console.error("Background sync error:", e);
        const syncBanner = document.getElementById('sync-banner');
        if (syncBanner) syncBanner.style.display = 'none';
        showToast("Tasks sync failed. Check your network or API configuration.");
    }
}

// --- Execute Initial Render logic with Search Preservation ---
setLayout(currentLayout); 

const savedSearch = localStorage.getItem('quadra_search') || '';
const searchInputEl = document.getElementById('searchInput');
if (searchInputEl && savedSearch) {
    searchInputEl.value = savedSearch;
    document.getElementById('clearSearchBtn').style.display = 'block';
}

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