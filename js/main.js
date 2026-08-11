import { state } from './state.js';
import * as ui from './ui.js';
import * as board from './board.js';
import * as tracker from './tracker.js';
import * as api from './api.js';

// Map HTML inline functions to the global window object
window.handleSearch = board.handleSearch;
window.updateQuickTags = board.updateQuickTags;
window.clearSearch = board.clearSearch;
window.searchTag = board.searchTag;

window.performBackgroundSync = api.performBackgroundSync;
window.setLayout = ui.setLayout;
window.handleAuthClick = api.handleAuthClick;
window.handleSignoutClick = api.handleSignoutClick;

window.openSettingsPage = ui.openSettingsPage;
window.closeSettingsPage = ui.closeSettingsPage;
window.saveSettings = ui.saveSettings;
window.toggleCalSourceFields = ui.toggleCalSourceFields;
window.addScheduleRow = ui.addScheduleRow;
window.clearCalendarCache = ui.clearCalendarCache;
window.exportData = ui.exportData;
window.triggerImport = ui.triggerImport;
window.importData = ui.importData;

window.dragStartQuad = board.dragStartQuad;
window.dragEndQuad = board.dragEndQuad;
window.allowDropQuad = board.allowDropQuad;
window.dragLeaveQuad = board.dragLeaveQuad;
window.dropQuad = board.dropQuad;
window.initQuadResize = board.initQuadResize;

window.openTaskModal = ui.openTaskModal;
window.closeTaskModal = ui.closeTaskModal;
window.saveTaskModal = ui.saveTaskModal;
window.onTimesChanged = ui.onTimesChanged;
window.onHoursChanged = ui.onHoursChanged;

window.openEditBlockModal = ui.openEditBlockModal;
window.closeEditBlockModal = ui.closeEditBlockModal;
window.saveTimeBlock = ui.saveTimeBlock;
window.deleteTimeBlock = ui.deleteTimeBlock;

window.completeTask = board.completeTask;
window.restoreTask = board.restoreTask;
window.deleteTask = board.deleteTask;

window.setTrackerMode = tracker.setTrackerMode;
window.changeTrackerDay = tracker.changeTrackerDay;
window.goToToday = tracker.goToToday;
window.adjustTimelineZoom = tracker.adjustTimelineZoom;
window.renderTrackerTimeline = tracker.renderTrackerTimeline;
window.exportTimesheet = tracker.exportTimesheet;
window.importCalendarEvents = api.importCalendarEvents;

window.startBlockDrag = tracker.startBlockDrag;
window.allowTrackerDrop = tracker.allowTrackerDrop;
window.dragLeaveTracker = tracker.dragLeaveTracker;
window.dropToTracker = tracker.dropToTracker;
window.handleTimelineClick = tracker.handleTimelineClick;

window.closeShortcutsModal = ui.closeShortcutsModal;

// Initialization
window.addEventListener('load', () => {
    const matrixContainer = document.getElementById('matrix');
    
    if (state.appConfig.quadrantOrder && state.appConfig.quadrantOrder.length > 0) {
        state.appConfig.quadrantOrder.forEach(id => {
            const el = document.getElementById(id);
            if (el) matrixContainer.appendChild(el);
        });
    }
    
    document.querySelectorAll('.quadrant').forEach(q => {
        if (state.appConfig.quadrantWidths && state.appConfig.quadrantWidths[q.id]) {
            q.style.width = state.appConfig.quadrantWidths[q.id];
        }
        board.quadResizeObserver.observe(q);
    });

    ui.applyViewVisibility();

    if (typeof gapi !== 'undefined') {
        gapi.load('client', () => {
            let initArgs = { 
                discoveryDocs: [
                    'https://www.googleapis.com/discovery/v1/apis/tasks/v1/rest',
                    'https://www.googleapis.com/discovery/v1/apis/calendar/v3/rest'
                ] 
            };
            if (state.appConfig.apiKey) initArgs.apiKey = state.appConfig.apiKey;
            gapi.client.init(initArgs).catch(() => {});
        });
    }
    
    if (typeof google !== 'undefined' && google.accounts && state.appConfig.clientId) {
        state.tokenClient = google.accounts.oauth2.initTokenClient({ 
            client_id: state.appConfig.clientId, 
            scope: api.SCOPES, 
            callback: async (resp) => { 
                if (resp.error !== undefined) { throw (resp); } 
                localStorage.setItem('quadra_gapi_token_v2', JSON.stringify({ token: resp.access_token, expires_at: Date.now() + (resp.expires_in * 1000) })); 
                document.getElementById('auth-overlay').style.display = 'none';
                document.getElementById('authorize_button').style.display = 'none'; 
                document.getElementById('signout_button').style.display = 'inline-block'; 
                state.isGoogleSynced = true;
                if(typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: resp.access_token });
                api.loadCalendars(); 
                api.performBackgroundSync(); 
            }, 
        });
    }
    api.checkConfigState();

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
                if (document.activeElement.nextElementSibling) document.activeElement.nextElementSibling.focus();
            } else if (e.key === 'ArrowLeft') {
                e.preventDefault();
                if (document.activeElement.previousElementSibling) document.activeElement.previousElementSibling.focus();
            } else if (e.key === 'ArrowUp' || e.key === 'Escape') {
                e.preventDefault();
                searchInput.focus();
            } else if (e.key === 'Enter') {
                e.preventDefault();
                document.activeElement.click();
            }
        });
    }

    ui.setLayout(state.currentLayout); 
    
    const savedSearch = localStorage.getItem('quadra_search') || '';
    if (searchInput && savedSearch) {
        searchInput.value = savedSearch;
        document.getElementById('clearSearchBtn').style.display = 'block';
    }
    board.renderNotes(savedSearch);
});

document.addEventListener('keydown', (e) => {
    const isEditingText = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName) || document.activeElement.isContentEditable;

    if (e.key === 'Escape') {
        const taskModal = document.getElementById('taskModal');
        const editBlockModal = document.getElementById('editTimeBlockModal');
        const shortcutsModal = document.getElementById('shortcutsModal');
        if (taskModal && taskModal.style.display === 'flex') ui.closeTaskModal();
        if (editBlockModal && editBlockModal.style.display === 'flex') ui.closeEditBlockModal();
        if (shortcutsModal && shortcutsModal.style.display === 'flex') ui.closeShortcutsModal();
    } else if (!isEditingText) {
        if (e.shiftKey && (e.key === '?' || e.key === '/')) {
            e.preventDefault();
            ui.openShortcutsModal();
        } else if (e.key === '/') {
            e.preventDefault();
            const searchInput = document.getElementById('searchInput');
            if (searchInput) { searchInput.focus(); searchInput.select(); }
        } else if (e.key.toLowerCase() === 'q' && state.appConfig.viewsEnabled.grid) {
            e.preventDefault(); ui.setLayout('grid');
        } else if (e.key.toLowerCase() === 'k' && state.appConfig.viewsEnabled.kanban) {
            e.preventDefault(); ui.setLayout('kanban');
        } else if (e.key.toLowerCase() === 't' && state.appConfig.viewsEnabled.tracker) {
            e.preventDefault(); ui.setLayout('tracker');
        } else if (e.key.toLowerCase() === 'o' && state.appConfig.viewsEnabled.overdue) {
            e.preventDefault(); ui.setLayout('overdue');
        } else if (e.key.toLowerCase() === 'c' && state.appConfig.viewsEnabled.gcal) {
            e.preventDefault(); ui.setLayout('gcal');
        } else if (e.key === '`' || e.key === '~') {
            e.preventDefault();
            const allViews = ['grid', 'kanban', 'overdue', 'gcal', 'tracker'];
            const views = allViews.filter(v => state.appConfig.viewsEnabled[v]);
            if (views.length === 0) return;
            
            let idx = views.indexOf(state.currentLayout);
            if (idx === -1) idx = 0;
            if (e.shiftKey) {
                idx = (idx - 1 + views.length) % views.length;
            } else {
                idx = (idx + 1) % views.length;
            }
            ui.setLayout(views[idx]);
        }
    }
});