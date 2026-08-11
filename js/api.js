import { state, saveNotes } from './state.js';
import { showToast, parseICS, roundToQuarterHour } from './utils.js';
import { renderTrackerTimeline } from './tracker.js';
import { handleSearch } from './board.js';

export const SCOPES = 'https://www.googleapis.com/auth/tasks https://www.googleapis.com/auth/calendar';

export async function loadCalendars() {
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
        const currentVal = state.appConfig.calendarId;
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

export function handleAuthClick() { 
    if (!state.appConfig.clientId) return window.openSettingsPage(); 
    if (!state.tokenClient && typeof google !== 'undefined' && google.accounts) { 
        state.tokenClient = google.accounts.oauth2.initTokenClient({ 
            client_id: state.appConfig.clientId, 
            scope: SCOPES, 
            callback: async (resp) => { 
                if (resp.error !== undefined) { throw (resp); } 
                localStorage.setItem('quadra_gapi_token_v2', JSON.stringify({ token: resp.access_token, expires_at: Date.now() + (resp.expires_in * 1000) })); 
                document.getElementById('auth-overlay').style.display = 'none'; 
                document.getElementById('authorize_button').style.display = 'none'; 
                document.getElementById('signout_button').style.display = 'inline-block'; 
                state.isGoogleSynced = true;
                if(typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: resp.access_token });
                loadCalendars(); 
                performBackgroundSync(); 
            }, 
        }); 
    }
    if (!state.tokenClient) { showToast("Google services are still loading. Please wait a moment and try again."); return; }
    state.tokenClient.requestAccessToken({prompt: 'consent'}); 
}

export function handleSignoutClick() { 
    const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (savedToken && savedToken.token && typeof google !== 'undefined' && google.accounts && google.accounts.oauth2) { 
        google.accounts.oauth2.revoke(savedToken.token, () => {}); 
    } 
    localStorage.removeItem('quadra_gapi_token_v2'); 
    state.isGoogleSynced = false; 
    document.getElementById('auth-overlay').style.display = 'none';
    document.getElementById('authorize_button').style.display = 'inline-block'; 
    document.getElementById('signout_button').style.display = 'none'; 
    showToast("Signed out successfully.");
}

export function checkConfigState() {
    const authorizeButton = document.getElementById('authorize_button');
    const signoutButton = document.getElementById('signout_button');

    authorizeButton.style.display = 'none';
    signoutButton.style.display = 'none';

    const savedTokenData = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (savedTokenData && savedTokenData.expires_at > Date.now()) {
        authorizeButton.style.display = 'none';
        signoutButton.style.display = 'inline-block';
        state.isGoogleSynced = true;
        if (typeof gapi !== 'undefined' && gapi.client) gapi.client.setToken({ access_token: savedTokenData.token });
        loadCalendars(); 
    } else {
        authorizeButton.style.display = 'inline-block';
        signoutButton.style.display = 'none';
    }
}

export async function importCalendarEvents() {
    const dateStr = document.getElementById('trackerDate').value;
    const ignoreKeywords = (state.appConfig.ignoreKeywords || 'out of office, ooo, away, vacation, holiday')
        .toLowerCase().split(',').map(k => k.trim()).filter(k => k.length > 0);

    let importedCount = 0;
    let ignoredCount = 0;

    try {
        document.getElementById('sync-banner').style.display = 'block';

        if (state.appConfig.calSource === 'outlook' && state.appConfig.icsUrl) {
            const res = await fetch(state.appConfig.icsUrl);
            if (!res.ok) throw new Error("Failed to fetch ICS feed via Apps Script bridge");
            
            const icsText = await res.text();
            const icsEvents = parseICS(icsText, dateStr, ignoreKeywords);

            icsEvents.forEach(ev => {
                if (state.notes.find(n => n.eventId === ev.id)) return;
                state.notes.push({
                    id: Date.now().toString() + Math.random(),
                    eventId: ev.id,
                    text: `${ev.summary} #meeting`,
                    quadrant: 'q2',
                    status: 'active',
                    dirty: false,
                    deleted: false,
                    timeBlocks: [{ id: Date.now().toString() + Math.random(), date: dateStr, startHour: ev.startHour, duration: ev.duration }]
                });
                importedCount++;
            });
        } else {
            const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
            if (!savedToken || !savedToken.token || savedToken.expires_at < Date.now()) {
                document.getElementById('sync-banner').style.display = 'none';
                return showToast("Please sign in to Google first.");
            }
            if (!state.appConfig.calendarId) {
                document.getElementById('sync-banner').style.display = 'none';
                return showToast("Please select a Google calendar in Settings first.");
            }
            
            const [y, m, day] = dateStr.split('-');
            const timeMin = new Date(y, m-1, day, 0, 0, 0).toISOString();
            const timeMax = new Date(y, m-1, day, 23, 59, 59).toISOString();

            const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(state.appConfig.calendarId)}/events?timeMin=${timeMin}&timeMax=${timeMax}&singleEvents=true&orderBy=startTime`, {
                headers: { 'Authorization': `Bearer ${savedToken.token}` }
            });
            const data = await res.json();
            if(data.error) throw new Error(data.error.message);

            const events = data.items || [];
            events.forEach(event => {
                if (!event.start.dateTime) return; // Skip all-day events

                const title = (event.summary || '').toLowerCase();
                const shouldIgnore = ignoreKeywords.some(keyword => title.includes(keyword));

                if (shouldIgnore) { ignoredCount++; return; }
                
                const start = new Date(event.start.dateTime);
                const end = new Date(event.end.dateTime);
                
                const rawStartHour = start.getHours() + (start.getMinutes() / 60);
                let rawEndHour = end.getHours() + (end.getMinutes() / 60);
                if (rawEndHour <= rawStartHour) rawEndHour += 24;
                const rawDuration = rawEndHour - rawStartHour;

                const startHour = roundToQuarterHour(rawStartHour);
                const duration = Math.max(0.25, roundToQuarterHour(rawDuration));
                
                if (state.notes.find(n => n.eventId === event.id)) return; 
                
                state.notes.push({
                    id: Date.now().toString() + Math.random(),
                    eventId: event.id, 
                    text: `${event.summary || 'Meeting'} #meeting`,
                    quadrant: 'q2', 
                    status: 'active',
                    dirty: false, 
                    deleted: false,
                    timeBlocks: [{ id: Date.now().toString() + Math.random(), date: dateStr, startHour: startHour, duration: duration }]
                });
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
        showToast("Failed to fetch calendar events. Check network.");
    }
}

export async function performBackgroundSync() {
    const savedToken = JSON.parse(localStorage.getItem('quadra_gapi_token_v2'));
    if (!savedToken || !savedToken.token) return showToast("Please sign in first to sync with Google Tasks.");
    try {
        document.getElementById('sync-banner').style.display = 'block';

        if (typeof gapi === 'undefined' || !gapi.client || !gapi.client.tasks || !gapi.client.tasks.tasklists || !gapi.client.tasks.tasks) {
            document.getElementById('sync-banner').style.display = 'none';
            return showToast("Google APIs are still loading. Please try again in a moment.");
        }

        if(typeof gapi !== 'undefined' && gapi.client && typeof gapi.client.setToken === 'function' && !gapi.client.getToken()) {
            gapi.client.setToken({ access_token: savedToken.token });
        }
        
        // 1. TASK SYNC LOGIC
        const response = await gapi.client.tasks.tasklists.list();
        const remoteLists = response.result.items || [];
        const GAPI_LIST_NAMES = { 'inbox': 'Quadra: Inbox', 'q1': 'Quadra: Do First', 'q2': 'Quadra: Schedule', 'q3': 'Quadra: Delegate', 'q4': 'Quadra: Later', 'closed': 'Quadra: Completed' };
        let gapiListIds = { inbox: null, q1: null, q2: null, q3: null, q4: null, closed: null };
        for (const quadKey of Object.keys(GAPI_LIST_NAMES)) { const existingList = remoteLists.find(l => l.title === GAPI_LIST_NAMES[quadKey]); if (existingList) { gapiListIds[quadKey] = existingList.id; } else { const newListReq = await gapi.client.tasks.tasklists.insert({ title: GAPI_LIST_NAMES[quadKey] }); gapiListIds[quadKey] = newListReq.result.id; } }
        let remoteTaskMap = {};
        for (const quadKey of Object.keys(gapiListIds)) { const listId = gapiListIds[quadKey]; const tasksReq = await gapi.client.tasks.tasks.list({ tasklist: listId, showHidden: true, maxResults: 100 }); const rTasks = tasksReq.result.items || []; rTasks.forEach(t => { remoteTaskMap[t.id] = { task: t, listId: listId, quadKey: quadKey }; }); }
        
        const syncSnapshot = JSON.parse(JSON.stringify(state.notes));
        for (let sn of syncSnapshot) {
            if (sn.eventId) continue; 
            if (sn.deleted && sn.dirty) { 
                if (remoteTaskMap[sn.id]) { 
                    try { await gapi.client.tasks.tasks.delete({ tasklist: remoteTaskMap[sn.id].listId, task: sn.id }); } catch(e){ console.error("Error deleting remote task", e); } 
                } 
                sn.dirty = false; 
                continue; 
            }
            if (sn.dirty && !sn.deleted) {
                const targetListId = gapiListIds[sn.quadrant]; 
                const remoteObj = remoteTaskMap[sn.id]; 
                const gStatus = sn.status === 'closed' ? 'completed' : 'needsAction'; 
                
                let formattedDue = null;
                if (sn.dueDate) { formattedDue = new Date(sn.dueDate).toISOString(); }
    
                const resourceBody = { title: sn.text, status: gStatus };
                if (formattedDue) { resourceBody.due = formattedDue; } else { resourceBody.due = null; }
    
                if (remoteObj) { 
                    if (remoteObj.listId !== targetListId) { 
                        try { 
                            await gapi.client.tasks.tasks.delete({ tasklist: remoteObj.listId, task: sn.id }); 
                            const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                            sn.tempId = sn.id; 
                            sn.id = res.result.id; 
                        } catch(e){ console.error("Error recreating task for move", e); } 
                    } else { 
                        try { 
                            await gapi.client.tasks.tasks.patch({ tasklist: targetListId, task: sn.id, resource: resourceBody }); 
                        } catch(e){ console.error("Error updating task", e); } 
                    } 
                    delete remoteTaskMap[sn.id]; 
                    if (sn.tempId) delete remoteTaskMap[sn.tempId]; 
                } else { 
                    try { 
                        const res = await gapi.client.tasks.tasks.insert({ tasklist: targetListId, resource: resourceBody }); 
                        sn.tempId = sn.id; 
                        sn.id = res.result.id; 
                    } catch(e){ console.error("Error inserting new task", e); } 
                }
                sn.dirty = false;
            } else if (!sn.dirty && !sn.deleted) { 
                const remoteObj = remoteTaskMap[sn.id]; 
                if (remoteObj) { 
                    sn.text = remoteObj.task.title || ''; 
                    sn.status = remoteObj.task.status === 'completed' ? 'closed' : 'active'; 
                    sn.quadrant = remoteObj.quadKey; 
                    sn.dueDate = remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null; 
                    delete remoteTaskMap[sn.id]; 
                } 
            }
        }
        Object.values(remoteTaskMap).forEach(remoteObj => { syncSnapshot.push({ id: remoteObj.task.id, text: remoteObj.task.title || '', quadrant: remoteObj.quadKey, status: remoteObj.task.status === 'completed' ? 'closed' : 'active', dueDate: remoteObj.task.due ? remoteObj.task.due.split('T')[0] : null, timeBlocks: [], eventId: null, dirty: false, deleted: false }); });
        let newNotesArray = []; let syncedIds = new Set();
        syncSnapshot.forEach(sn => { if (sn.deleted && !sn.dirty) return; const liveNote = state.notes.find(n => n.id === sn.id || (sn.tempId && n.id === sn.tempId)); if (liveNote) { sn.timeBlocks = liveNote.timeBlocks || []; sn.eventId = liveNote.eventId || null; if (liveNote.dirty) { if (sn.tempId) liveNote.id = sn.id; newNotesArray.push(liveNote); } else { newNotesArray.push(sn); } syncedIds.add(liveNote.id); if (sn.tempId) syncedIds.add(sn.tempId); } else { newNotesArray.push(sn); } });
        state.notes = newNotesArray; 
        saveNotes(); 
        handleSearch(); 

        // 2. TRACKER SYNC TO GOOGLE CALENDAR
        let calSyncedCount = 0;
        if (gapi.client.calendar) {
            let quadraCalId = null;
            const calList = await gapi.client.calendar.calendarList.list();
            const existingCal = calList.result.items.find(c => c.summary === 'Quadra');
            if (existingCal) {
                quadraCalId = existingCal.id;
            } else {
                const newCal = await gapi.client.calendar.calendars.insert({
                    resource: { summary: 'Quadra', description: 'Synced time blocks from Quadra Tracker' }
                });
                quadraCalId = newCal.result.id;
            }

            let activeTz = state.appConfig.primaryTz !== 'local' ? state.appConfig.primaryTz : Intl.DateTimeFormat().resolvedOptions().timeZone;

            for (let note of state.notes) {
                if (note.deleted || !note.timeBlocks || note.eventId) continue; 
                for (let tb of note.timeBlocks) {
                    const [y, m, d] = tb.date.split('-');
                    let startH = Math.floor(tb.startHour);
                    let startM = Math.round((tb.startHour - startH) * 60);
                    let endHourDec = tb.startHour + tb.duration;
                    let endH = Math.floor(endHourDec);
                    let endM = Math.round((endHourDec - endH) * 60);
                    
                    let endDate = new Date(y, m - 1, d, endH, endM);
                    if (endH >= 24) { endDate = new Date(y, m - 1, d, endH - 24, endM); endDate.setDate(endDate.getDate() + 1); }
                    let startDate = new Date(y, m - 1, d, startH, startM);
                    
                    let displayTitle = (note.text || '').split('\n')[0];
                    const eventPayload = {
                        summary: displayTitle,
                        description: "Synced from Quadra Time Tracker",
                        start: { dateTime: startDate.toISOString(), timeZone: activeTz },
                        end: { dateTime: endDate.toISOString(), timeZone: activeTz }
                    };

                    if (tb.gcalEventId) {
                        try {
                            await gapi.client.calendar.events.update({ calendarId: quadraCalId, eventId: tb.gcalEventId, resource: eventPayload });
                            calSyncedCount++;
                        } catch(e) {
                            const res = await gapi.client.calendar.events.insert({ calendarId: quadraCalId, resource: eventPayload });
                            tb.gcalEventId = res.result.id;
                            calSyncedCount++;
                        }
                    } else {
                        const res = await gapi.client.calendar.events.insert({ calendarId: quadraCalId, resource: eventPayload });
                        tb.gcalEventId = res.result.id;
                        calSyncedCount++;
                    }
                }
            }
            saveNotes();
        }

        document.getElementById('sync-banner').style.display = 'none';
        showToast(`Successfully synced Tasks & ${calSyncedCount} Tracker Blocks!`);
    } catch (e) {
        document.getElementById('sync-banner').style.display = 'none';
        console.error("Sync Process Error: ", e);
        showToast("Sync failed. Ensure API scopes are updated.");
    }
}