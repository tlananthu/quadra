import { state, saveNotes, saveQuadrantState } from './state.js';
import { escapeHTML, parseTags } from './utils.js';
import { openTaskModal } from './ui.js';
import { renderTrackerTimeline } from './tracker.js';

// Quadrant Drag and Drop
export function dragStartQuad(e) {
    if (e.target.classList.contains('quadrant')) {
        e.dataTransfer.setData('quadrant_id', e.target.id);
        e.dataTransfer.effectAllowed = 'move';
        setTimeout(() => e.target.style.opacity = '0.5', 0);
    }
}

export function dragEndQuad(e) {
    if (e.target.classList.contains('quadrant')) {
        e.target.style.opacity = '1';
        document.querySelectorAll('.quadrant').forEach(q => q.classList.remove('quad-drag-over', 'drag-over'));
    }
}

export function allowDropQuad(e) {
    e.preventDefault();
    if(e.dataTransfer.types.includes('quadrant_id')) {
        e.currentTarget.classList.add('quad-drag-over');
    } else {
        e.currentTarget.classList.add('drag-over');
    }
}

export function dragLeaveQuad(e) {
    e.currentTarget.classList.remove('quad-drag-over', 'drag-over');
}

export function dropQuad(e) {
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
        const note = state.notes.find(n => n.id === noteId);
        let targetKey = e.currentTarget.id; 
        if(targetKey.startsWith('tray-')) targetKey = targetKey.replace('tray-', '');
        
        if (note && note.status === 'active' && note.quadrant !== targetKey) { 
            if (targetKey === 'closed') note.status = 'closed';
            note.quadrant = targetKey; 
            note.dirty = true; 
            saveNotes(); 
            handleSearch(); 
        }
    }
}

export function initQuadResize(e, quadId) {
    e.preventDefault();
    e.stopPropagation();
    
    const quadEl = document.getElementById(quadId);
    if (!quadEl) return;
    
    const startX = e.clientX;
    const startWidth = quadEl.getBoundingClientRect().width;
    const handleEl = e.currentTarget;
    handleEl.classList.add('resizing');
    document.body.style.cursor = 'col-resize';
    
    function onMouseMove(ev) {
        ev.preventDefault();
        const deltaX = ev.clientX - startX;
        const newWidth = Math.max(260, Math.min(1000, Math.round(startWidth + deltaX)));
        quadEl.style.width = newWidth + 'px';
        state.appConfig.quadrantWidths[quadId] = newWidth + 'px';
    }
    
    function onMouseUp(ev) {
        handleEl.classList.remove('resizing');
        document.body.style.cursor = '';
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
        localStorage.setItem('quadra_config', JSON.stringify(state.appConfig));
    }
    
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
}

// Search and Tags
export function searchTag(tag, event) { 
    event.stopPropagation(); 
    document.getElementById('searchInput').value = tag; 
    handleSearch(); 
}

export function clearSearch() { 
    document.getElementById('searchInput').value = ''; 
    handleSearch(); 
}

export function insertSearchSuggestion(suggestion) {
    const searchInput = document.getElementById('searchInput');
    if (!searchInput) return;

    const currentValue = searchInput.value || '';
    const cursorPos = searchInput.selectionStart ?? currentValue.length;
    const beforeCursor = currentValue.slice(0, cursorPos);
    const afterCursor = currentValue.slice(cursorPos);
    const currentTokenMatch = beforeCursor.match(/(?:^|\s)([^\s]*)$/);
    const tokenStart = currentTokenMatch ? beforeCursor.length - currentTokenMatch[1].length : 0;
    const newBefore = beforeCursor.slice(0, tokenStart);

    let replacement = suggestion;
    if (replacement !== 'AND' && replacement !== 'OR') {
        replacement = replacement.toLowerCase();
    }
    
    const suffix = ' ';
    const cleanAfter = afterCursor.trimStart();
    const nextValue = `${newBefore}${replacement}${suffix}${cleanAfter}`;

    searchInput.value = nextValue;
    handleSearch(); 
    
    searchInput.focus();
    const nextCursor = (newBefore + replacement + suffix).length;
    searchInput.setSelectionRange(nextCursor, nextCursor);
}

export function updateQuickTags() {
    const tagsBar = document.getElementById('quick-tags-bar'); 
    if (!tagsBar) return;
    
    const searchInput = document.getElementById('searchInput');
    const currentValue = searchInput ? (searchInput.value || '') : '';
    const cursorPos = searchInput ? (searchInput.selectionStart ?? currentValue.length) : currentValue.length;
    const beforeCursor = currentValue.slice(0, cursorPos);
    
    const lastTokenMatch = beforeCursor.match(/(?:^|\s)([^\s]*)$/);
    const lastToken = lastTokenMatch ? lastTokenMatch[1] : '';
    
    let prefix = null;
    let showBooleans = false;
    
    if (lastToken.startsWith('#') || lastToken.startsWith('@')) {
        prefix = lastToken.toLowerCase();
    } else if (beforeCursor.endsWith(' ') || beforeCursor.trim() !== '') {
        if (beforeCursor.trim() !== '' && !lastToken) showBooleans = true;
    }

    const tagCounts = {};
    state.notes.forEach(note => {
        if (!note.deleted && note.status === 'active' && !note.eventId) {
            const matches = note.text.match(/(#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+)/g);
            if (matches) {
                const uniqueTags = new Set(matches.map(t => t.toLowerCase()));
                uniqueTags.forEach(tag => { tagCounts[tag] = (tagCounts[tag] || 0) + 1; });
            }
        }
    });

    tagsBar.innerHTML = '';

    if (showBooleans) {
        ['AND', 'OR'].forEach(op => {
            const btn = document.createElement('button');
            btn.className = 'filter-tag';
            btn.innerText = op;
            btn.onclick = (e) => { e.preventDefault(); insertSearchSuggestion(op); };
            tagsBar.appendChild(btn);
        });
    }

    const sortedTags = Object.keys(tagCounts).sort((a, b) => {
        if (tagCounts[b] !== tagCounts[a]) return tagCounts[b] - tagCounts[a];
        return a.localeCompare(b);
    });

    sortedTags.forEach(tag => {
        if (prefix && !tag.startsWith(prefix)) return;
        const btn = document.createElement('button');
        const isPerson = tag.startsWith('@');
        btn.className = 'filter-tag' + (isPerson ? ' person-filter' : '');
        btn.innerText = `${tag} (${tagCounts[tag]})`;
        btn.onclick = (e) => { e.preventDefault(); insertSearchSuggestion(tag); };
        tagsBar.appendChild(btn);
    });
}

export function matchesSearchQuery(noteText, query) {
    const normalizedText = (noteText || '').toLowerCase();
    const trimmedQuery = (query || '').trim();
    if (!trimmedQuery) return true;

    const normalizedQuery = trimmedQuery.toLowerCase();
    const hasBooleanOperator = /\b(and|or)\b/.test(normalizedQuery);
    if (!hasBooleanOperator) {
        return normalizedText.includes(normalizedQuery);
    }

    const tokens = normalizedQuery.match(/#[a-zA-Z0-9_]+|@[a-zA-Z0-9_]+|and|or|[^\s]+/g) || [];
    const terms = [];
    let pendingOperator = 'and';

    tokens.forEach(token => {
        if (token === 'and' || token === 'or') {
            pendingOperator = token;
        } else {
            terms.push({ value: token, operator: pendingOperator });
            pendingOperator = 'and';
        }
    });

    if (terms.length === 0) return true;

    let result = normalizedText.includes(terms[0].value);
    for (let i = 1; i < terms.length; i++) {
        const term = terms[i];
        const match = normalizedText.includes(term.value);
        if (term.operator === 'or') {
            result = result || match;
        } else {
            result = result && match;
        }
    }
    return result;
}

export function handleSearch() {
    const searchInput = document.getElementById('searchInput');
    if(!searchInput) return;
    const query = searchInput.value;
    
    if (query.trim().length > 0) {
        localStorage.setItem('quadra_search', query);
    } else {
        localStorage.removeItem('quadra_search');
    }

    const clearBtn = document.getElementById('clearSearchBtn');
    if (clearBtn) clearBtn.style.display = query.length > 0 ? 'block' : 'none';
    renderNotes(query);
}

export function renderNotes(searchQuery = '') {
    ['q1', 'q2', 'q3', 'q4', 'inbox', 'calendar', 'closed'].forEach(q => { 
        const el = document.getElementById(`list-${q}`); 
        if (el) el.innerHTML = ''; 
    });
    
    let filteredNotes = state.notes.filter(note => {
        if (note.deleted) return false;
        const isCalendarEvent = note.eventId !== null && note.eventId !== undefined;
        if (isCalendarEvent) return true;
        return matchesSearchQuery(note.text, searchQuery);
    });

    filteredNotes.sort((a, b) => { 
        if (!a.dueDate && !b.dueDate) return 0; 
        if (!a.dueDate) return -1; 
        if (!b.dueDate) return 1; 
        return a.dueDate.localeCompare(b.dueDate); 
    });

    let filteredCounts = { q1: 0, q2: 0, q3: 0, q4: 0, inbox: 0, calendar: 0, closed: 0 };
    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];

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

            contentWrapper.innerHTML = `<div class="note-text">${overdueIndicator}${parseTags(note.text)}</div>`;
            
            if (note.dueDate) contentWrapper.innerHTML += `<div style="font-size:12px; color:var(--brand-primary); margin-top:6px; font-weight:500;">🗓️ ${note.dueDate.split('T')[0]}</div>`;
            
            contentWrapper.onclick = (e) => openTaskModal(null, note.id, e);

            const actionsDiv = document.createElement('div'); 
            actionsDiv.className = 'note-actions';
            if (!note.eventId) {
                if (note.status === 'active') actionsDiv.innerHTML = `<button class="action-btn complete-btn" onclick="completeTask('${note.id}')">✓</button>`;
                else actionsDiv.innerHTML = `<button class="action-btn restore-btn" onclick="restoreTask('${note.id}')">↺</button><button class="action-btn delete-btn" onclick="deleteTask('${note.id}')">×</button>`;
            } else {
                actionsDiv.innerHTML = `<button class="action-btn delete-btn" onclick="deleteTask('${note.id}')" title="Remove event">×</button>`;
            }
            
            noteEl.append(contentWrapper, actionsDiv); 
            list.appendChild(noteEl);
        }
    });

    const badgeClosed = document.getElementById('badge-closed');
    if (badgeClosed) badgeClosed.innerText = state.notes.filter(n => !n.deleted && n.status === 'closed' && !n.eventId).length;
    
    const badgeInbox = document.getElementById('badge-inbox');
    if (badgeInbox) badgeInbox.innerText = filteredCounts.inbox;
    
    const badgeCal = document.getElementById('badge-calendar');
    if (badgeCal) badgeCal.innerText = filteredCounts.calendar;
    
    ['q1', 'q2', 'q3', 'q4'].forEach(q => {
        const badge = document.getElementById(`badge-${q}`);
        if (badge) badge.innerText = filteredCounts[q];
    });
    
    if (state.currentLayout === 'tracker') renderTrackerTimeline();
    updateQuickTags();
}

export function renderOverdueTasksPage() {
    const container = document.getElementById('overdue-list-container');
    if (!container) return;
    container.innerHTML = '';
    
    const todayStr = new Date().toLocaleDateString('en-CA').split('T')[0];
    const overdueNotes = state.notes.filter(n => {
        if (n.deleted || n.status === 'closed' || !n.dueDate || n.eventId) return false;
        return n.dueDate.split('T')[0] < todayStr;
    });

    overdueNotes.sort((a, b) => a.dueDate.localeCompare(b.dueDate));

    const badge = document.getElementById('overdue-count-badge');
    if (badge) badge.innerText = overdueNotes.length;

    if (overdueNotes.length === 0) {
        container.innerHTML = `<p style="color: var(--text-muted); text-align: center; padding: 40px;">🎉 No overdue tasks! All caught up.</p>`;
        return;
    }

    overdueNotes.forEach(note => {
        const card = document.createElement('div');
        card.className = 'overdue-card';
        card.onclick = (e) => openTaskModal(null, note.id, e);
        card.style.cursor = 'pointer';

        let title = note.text.split('\n')[0];
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

export function completeTask(id) { 
    const note = state.notes.find(n => n.id === id); 
    if (note) { 
        note.status = 'closed'; 
        note.quadrant = 'closed'; 
        note.dirty = true; 
        saveNotes(); 
        handleSearch(); 
    } 
}

export function restoreTask(id) { 
    const note = state.notes.find(n => n.id === id); 
    if (note) { 
        note.status = 'active'; 
        note.quadrant = 'inbox'; 
        note.dirty = true; 
        saveNotes(); 
        handleSearch(); 
    } 
}

export function deleteTask(id) { 
    const note = state.notes.find(n => n.id === id); 
    if (note) { 
        note.deleted = true; 
        note.dirty = true; 
        saveNotes(); 
        handleSearch(); 
    } 
}