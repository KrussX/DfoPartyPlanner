/**
 * DFO Party Planner - Logic
 * Collaborative via Firebase Firestore.
 */
document.addEventListener('DOMContentLoaded', () => {
    // --- Firebase Init ---
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore();
    const auth = firebase.auth();

    // --- Room ID (from URL hash, e.g. site.com/#abc123) ---
    function generateRoomId() {
        return Math.random().toString(36).substring(2, 8) +
            Math.random().toString(36).substring(2, 8);
    }

    let roomId = window.location.hash.replace('#', '');
    let roomRef = null;
    let userName = sessionStorage.getItem('dfoUserName');

    const landingOverlay = document.getElementById('landing-overlay');
    const landingForm = document.getElementById('landing-form');
    const landingNameInput = document.getElementById('landing-name-input');
    const landingSubtitle = document.getElementById('landing-subtitle');
    const landingSubmitBtn = document.getElementById('landing-submit-btn');
    const appLayout = document.querySelector('.app-layout');
    const userDisplay = document.getElementById('user-display');
    const logoutBtn = document.getElementById('logout-btn');

    if (userName && userDisplay) {
        userDisplay.textContent = '👤 ' + userName;
    }

    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            sessionStorage.removeItem('dfoUserName');
            window.location.reload();
        });
    }

    if (!userName) {
        appLayout.style.display = 'none';
        landingOverlay.style.display = 'flex';

        if (!roomId) {
            landingSubtitle.textContent = "Welcome! Please enter your name to create a room.";
            landingSubmitBtn.textContent = "Create Room";
        } else {
            landingSubtitle.textContent = "You've been invited! Enter your name to join the room.";
            landingSubmitBtn.textContent = "Join Room";
        }

        landingForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const name = landingNameInput.value.trim();
            if (!name) return;
            sessionStorage.setItem('dfoUserName', name);
            userName = name;

            if (!roomId) {
                roomId = generateRoomId();
                window.location.hash = roomId;
            }

            if (userDisplay) {
                userDisplay.textContent = '👤 ' + userName;
            }

            landingOverlay.style.display = 'none';
            appLayout.style.display = 'flex';
            initApp();
        });
    } else {
        if (!roomId) {
            roomId = generateRoomId();
            window.location.hash = roomId;
        }
        initApp();
    }

    // --- Shared State ---
    const partyPlan = new Map(); // characterId -> charData
    const hiddenCharacters = new Set(); // local only (search result hide)
    let currentSearchResults = [];
    let contents = [];
    let activeContentId = null;
    let draggedCardId = null;
    let sourceSlotId = null;
    let isAdmin = false;
    let isWriting = false; // Prevents snapshot re-renders from our own writes

    // Pool controls state
    let poolFilter = 'all'; // 'all' | 'dps' | 'buff'
    let poolSort = null;    // null | 'dps' | 'buff'
    let poolSearchQuery = '';


    // --- DOM Elements ---
    const searchForm = document.getElementById('search-form');
    const searchBtn = document.getElementById('search-btn');
    const searchLoader = document.getElementById('search-loader');
    const searchBtnText = searchBtn.querySelector('.btn-text');
    const searchStatus = document.getElementById('search-status');
    const searchResults = document.getElementById('search-results');

    const searchContent = document.getElementById('search-content');
    const toggleSearchBtn = document.getElementById('toggle-search-btn');

    const partyList = document.getElementById('party-list');
    const addRaidBtn = document.getElementById('add-raid-btn');
    const autoPlanBtn = document.getElementById('auto-plan-btn');
    const raidsContainer = document.getElementById('raids-container');
    const clubSummaryContainer = document.getElementById('club-summary-container');
    const clubSummaryList = document.getElementById('club-summary-list');

    // Controls
    const exportJsonBtn = document.getElementById('export-json-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const importJsonBtn = document.getElementById('import-json-btn');
    const importFileInput = document.getElementById('import-file-input');
    const themeToggleBtn = document.getElementById('theme-toggle-btn');
    const refreshScoresBtn = document.getElementById('refresh-scores-btn');
    const copyLinkBtn = document.getElementById('copy-link-btn');

    // Confirm Modal Elements
    const confirmModal = document.getElementById('confirm-modal');
    const modalMessage = document.getElementById('modal-message');
    const modalConfirmBtn = document.getElementById('modal-confirm');
    const modalCancelBtn = document.getElementById('modal-cancel');

    // Manage Users Elements
    const manageUsersBtn = document.getElementById('manage-users-btn');
    const manageUsersModal = document.getElementById('manage-users-modal');
    const closeManageUsersBtn = document.getElementById('close-manage-users-btn');
    const usersListContainer = document.getElementById('users-list-container');

    // Pool Controls
    const poolSearchInput = document.getElementById('pool-search');

    // Content Controls
    const contentSelect = document.getElementById('content-select');
    const addContentBtn = document.getElementById('add-content-btn');
    const editContentBtn = document.getElementById('edit-content-btn');
    const duplicateContentBtn = document.getElementById('duplicate-content-btn');
    const deleteContentBtn = document.getElementById('delete-content-btn');
    // autoPlanBtn — declared above at line 111

    // Content Modal Controls
    const contentModal = document.getElementById('content-modal');
    const contentForm = document.getElementById('content-form');
    const contentNameInput = document.getElementById('content-name-input');
    const contentClubLimitInput = document.getElementById('content-club-limit-input');
    const contentPartySizeInput = document.getElementById('content-party-size-input');
    const contentModalCancel = document.getElementById('content-modal-cancel');
    const contentModalTitle = document.getElementById('content-modal-title');

    function getActiveContent() {
        if (!activeContentId) return null;
        return contents.find(c => c.id === activeContentId) || null;
    }


    const searchToggleHeader = document.getElementById('search-toggle-header');

    // --- Toggle Search ---
    searchToggleHeader.style.cursor = 'pointer';
    searchToggleHeader.addEventListener('click', () => {
        if (searchContent.style.display === 'none') {
            searchContent.style.display = 'block';
            toggleSearchBtn.textContent = 'Hide Search ▾';
        } else {
            searchContent.style.display = 'none';
            toggleSearchBtn.textContent = 'Show Search ▸';
        }
    });

    // --- Theme Toggle (local only) ---
    function applyTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        themeToggleBtn.textContent = theme === 'dark' ? '☀️ Light' : '🌙 Dark';
        localStorage.setItem('dfoTheme', theme);
    }

    const savedTheme = localStorage.getItem('dfoTheme') || 'light';
    applyTheme(savedTheme);

    themeToggleBtn.addEventListener('click', () => {
        const current = document.documentElement.getAttribute('data-theme');
        applyTheme(current === 'dark' ? 'light' : 'dark');
    });

    // --- Copy Link ---
    copyLinkBtn.addEventListener('click', () => {
        navigator.clipboard.writeText(window.location.href).then(() => {
            showToast('Room link copied to clipboard!', 'success');
        }).catch(() => {
            prompt('Copy this link to share the room:', window.location.href);
        });
    });

    // --- Sidebar Toggle ---
    const sidebarToggleBtn = document.getElementById('sidebar-toggle-btn');
    const panelLeft = document.querySelector('.panel-left');

    function applySidebarState(collapsed) {
        if (collapsed) {
            panelLeft.classList.add('collapsed');
            sidebarToggleBtn.textContent = '▶ Show Panel';
        } else {
            panelLeft.classList.remove('collapsed');
            sidebarToggleBtn.textContent = '◀ Hide Panel';
        }
        localStorage.setItem('dfoSidebarCollapsed', collapsed ? '1' : '0');
    }

    const savedSidebarState = localStorage.getItem('dfoSidebarCollapsed') === '1';
    applySidebarState(savedSidebarState);

    sidebarToggleBtn.addEventListener('click', () => {
        const isCollapsed = panelLeft.classList.contains('collapsed');
        applySidebarState(!isCollapsed);
    });

    // --- State Hydration ---
    function loadPartyPlanFromContent() {
        partyPlan.clear();
        const current = getActiveContent();
        if (current && current.partyPlan) {
            Object.entries(current.partyPlan).forEach(([id, charData]) => partyPlan.set(id, charData));
        }
    }

    function syncPartyPlanToContent() {
        const current = getActiveContent();
        if (current) {
            const planObj = {};
            partyPlan.forEach((val, key) => planObj[key] = val);
            current.partyPlan = planObj;
        }
    }

    // --- Firestore: Apply incoming state to local vars + re-render ---
    function applyState(data) {
        if (data.contents) {
            contents = data.contents;
        } else {
            // Legacy Migration
            contents = [];
            if (data.raids && data.raids.length > 0) {
                let partySizeVal = 3;
                if (data.meta && data.meta.raidSize) {
                    partySizeVal = parseInt(data.meta.raidSize) / 4;
                }
                contents.push({
                    id: 'content_default',
                    name: 'Default Content',
                    clubLimit: data.meta ? (data.meta.globalClubLimit || '') : '',
                    partySize: partySizeVal,
                    raidCounter: data.meta ? (data.meta.raidCounter || 0) : 0,
                    raids: data.raids || []
                });
            }
        }

        // Native property check
        contents.forEach(c => {
            if (!c.partyPlan) c.partyPlan = {};
        });

        // Legacy partyPlan Migration (Move root partyPlan to first content)
        if (data.partyPlan && contents.length > 0) {
            const firstContent = contents[0];
            const migratedPlan = {};
            if (Array.isArray(data.partyPlan)) {
                data.partyPlan.forEach(([id, charData]) => migratedPlan[id] = charData);
            } else {
                Object.entries(data.partyPlan).forEach(([id, charData]) => migratedPlan[id] = charData);
            }
            firstContent.partyPlan = { ...migratedPlan, ...firstContent.partyPlan };
        }

        // Restore activeContentId logic safely
        const localActive = localStorage.getItem('dfoActiveContentId');
        if (localActive && contents.some(c => c.id === localActive)) {
            activeContentId = localActive;
        } else if (contents.length > 0) {
            activeContentId = contents[0].id; // Fallback to first
        } else {
            activeContentId = null;
        }

        renderContentSelect();

        if (data.meta) {
            const admins = data.meta.admins || [];
            if (admins.length === 0 && data.meta.adminName) admins.push(data.meta.adminName);

            const newIsAdmin = admins.includes(userName);
            if (newIsAdmin !== isAdmin) {
                isAdmin = newIsAdmin;
                updateAdminUI();
            }

            if (data.meta.users) {
                renderUsersList(data.meta.users, admins);
            }
        }
        if (data.history) {
            renderHistory(data.history);
        }

        loadPartyPlanFromContent();

        updateAllViews();
        updateAdminUI();
    }

    function updateAdminUI() {
        document.querySelectorAll('.admin-only').forEach(el => {
            el.style.display = isAdmin ? '' : 'none';
        });

        // Add Raid / Auto Assign: Visible to everyone if content is selected
        if (addRaidBtn) {
            addRaidBtn.style.display = activeContentId ? '' : 'none';
        }
        if (autoPlanBtn) {
            autoPlanBtn.style.display = (isAdmin && activeContentId) ? '' : 'none';
        }
    }

    async function savePartyPlan() {
        syncPartyPlanToContent();
        await saveContents();
    }

    async function saveContents() {
        isWriting = true;
        try {
            await roomRef.update({ contents: contents });
        } catch (e) { console.error('saveContents failed', e); }
        setTimeout(() => { isWriting = false; }, 500);
    }

    // Legacy wrappers referencing the new system
    async function saveRaids() {
        await saveContents();
    }
    async function saveMeta() {
        await saveContents();
    }

    // Legacy saveState: calls targeted saves. Used by updateAllViews.
    function saveState() {
        savePartyPlan();
        saveRaids();
    }

    // --- Firestore: Subscribe to real-time updates ---
    function subscribeToRoom() {
        roomRef.onSnapshot((doc) => {
            if (!doc.exists) return;
            const data = doc.data();

            // Always render history so it is instantly live for everyone including the writer
            if (data.history) {
                renderHistory(data.history);
            }

            if (isWriting) return; // Skip re-render of the main board from our own writes
            applyState(data);
        }, (err) => {
            console.error('Firestore snapshot error:', err);
        });
    }

    // --- History Logic ---
    async function logAction(action, details) {
        if (!roomRef) return;
        const entry = {
            id: Date.now() + Math.random(),
            time: Date.now(),
            user: userName,
            action,
            details,
            isAdmin
        };
        try {
            await roomRef.update({
                history: firebase.firestore.FieldValue.arrayUnion(entry)
            });
        } catch (e) {
            console.error('History log failed', e);
        }
    }

    function renderHistory(historyArray) {
        const list = document.getElementById('history-list');
        if (!historyArray || historyArray.length === 0) {
            list.innerHTML = '<p class="history-empty">No actions recorded yet.</p>';
            return;
        }

        let html = '';
        // Sort newest first
        const sorted = [...historyArray].sort((a, b) => b.time - a.time);

        sorted.forEach(item => {
            const date = new Date(item.time);
            const timeStr = date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            const adminTag = item.isAdmin ? '<span style="color:var(--accent-blue)">[Admin]</span> ' : '';

            html += `
                <div class="history-item ${item.isAdmin ? 'admin' : ''}">
                    <div class="hist-time">${timeStr}</div>
                    <div><span class="hist-user">${adminTag}${item.user}</span> ${item.details}</div>
                </div>
            `;
        });
        list.innerHTML = html;
    }

    const historySidebar = document.getElementById('history-sidebar');
    const closeHistoryBtn = document.getElementById('close-history-btn');
    const historyBtn = document.getElementById('history-btn');

    historyBtn.addEventListener('click', () => { historySidebar.classList.add('open'); });
    closeHistoryBtn.addEventListener('click', () => { historySidebar.classList.remove('open'); });

    // --- Manage Users Logic ---
    function renderUsersList(users, admins) {
        if (!usersListContainer) return;
        let html = '';
        users.forEach(u => {
            const isUserAdmin = admins.includes(u);
            const isSelf = u === userName;

            html += `
                <div class="user-item">
                    <div class="user-name-wrapper">
                        <span class="user-name">${u}${isSelf ? ' (You)' : ''}</span>
                        ${isUserAdmin ? '<span class="user-admin-badge">[Admin]</span>' : ''}
                    </div>
                    ${(!isUserAdmin && isAdmin) ? `<button class="btn btn-sm btn-primary promote-btn" data-user="${u}">Make Admin</button>` : ''}
                </div>
            `;
        });
        usersListContainer.innerHTML = html;
    }

    async function promoteUser(targetName) {
        if (!isAdmin || !roomRef) return;
        try {
            await roomRef.update({
                'meta.admins': firebase.firestore.FieldValue.arrayUnion(targetName)
            });
            logAction('promote_admin', `promoted ${targetName} to Admin.`);
            showToast(`${targetName} is now an Admin!`, 'success');
        } catch (e) {
            console.error('Promotion failed', e);
            showToast('Failed to promote user.', 'danger');
        }
    }

    manageUsersBtn.addEventListener('click', () => {
        manageUsersModal.style.display = 'flex';
    });

    closeManageUsersBtn.addEventListener('click', () => {
        manageUsersModal.style.display = 'none';
    });

    usersListContainer.addEventListener('click', (e) => {
        if (e.target.classList.contains('promote-btn')) {
            const target = e.target.getAttribute('data-user');
            promoteUser(target);
        }
    });

    // --- Auth + Room Init ---
    function initApp() {
        roomRef = db.collection('rooms').doc(roomId);

        auth.signInAnonymously().then(async (cred) => {
            const uid = cred.user.uid;

            // Check if room exists
            const snap = await roomRef.get();
            if (!snap.exists) {
                // New room — current user is the admin
                const initialState = {
                    meta: {
                        admins: [userName],
                        users: [userName]
                    },
                    contents: [
                        {
                            id: 'content_' + Date.now(),
                            name: 'General Raid',
                            clubLimit: '',
                            partySize: 3,
                            raidCounter: 0,
                            raids: []
                        }
                    ],
                    partyPlan: {},
                    history: [{
                        id: Date.now(), time: Date.now(), user: 'System',
                        action: 'created_room', details: `${userName} created the room.`
                    }]
                };
                await roomRef.set(initialState);
                isAdmin = true;
            } else {
                const data = snap.data();
                let admins = data.meta.admins || [];

                // Persist Migration: If user matches the old adminName, ensure they are in the admins array in Firestore
                if (data.meta.adminName === userName && !admins.includes(userName)) {
                    await roomRef.update({
                        'meta.admins': firebase.firestore.FieldValue.arrayUnion(userName)
                    });
                    // For local immediate state before snapshot reflects
                    if (!admins.includes(userName)) admins.push(userName);
                }

                isAdmin = admins.includes(userName);

                // Add current user to users list idempotently
                await roomRef.update({
                    'meta.users': firebase.firestore.FieldValue.arrayUnion(userName)
                });

                logAction('joined_room', `joined the room.`);
            }

            updateAdminUI();

            // Subscribe to live updates
            subscribeToRoom();
        }).catch(err => console.error('Auth error:', err));
    }

    // --- Content Management Logic ---
    function renderContentSelect() {
        if (!contentSelect) return;
        contentSelect.innerHTML = '';
        contents.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.id;
            opt.textContent = c.name;
            if (c.id === activeContentId) opt.selected = true;
            contentSelect.appendChild(opt);
        });
    }

    contentSelect.addEventListener('change', (e) => {
        activeContentId = e.target.value;
        localStorage.setItem('dfoActiveContentId', activeContentId);
        loadPartyPlanFromContent();
        updateAdminUI();
        updateAllViews();
    });

    let editingContentId = null;

    addContentBtn.addEventListener('click', () => {
        if (!isAdmin) return;
        editingContentId = null;
        contentModalTitle.textContent = 'Add Content';
        contentNameInput.value = '';
        contentClubLimitInput.value = '';
        contentPartySizeInput.value = '3';
        contentModal.style.display = 'flex';
    });

    editContentBtn.addEventListener('click', () => {
        if (!isAdmin) return;
        const current = getActiveContent();
        if (!current) return;
        editingContentId = current.id;
        contentModalTitle.textContent = 'Edit Content';
        contentNameInput.value = current.name || '';
        contentClubLimitInput.value = current.clubLimit || '';
        contentPartySizeInput.value = (current.partySize || 3).toString();
        contentModal.style.display = 'flex';
    });

    duplicateContentBtn.addEventListener('click', () => {
        if (!isAdmin) return;
        const current = getActiveContent();
        if (!current) return;

        showConfirm(`Are you sure you want to duplicate "${current.name}"? This will copy all raids and character assignments.`, () => {
            // Create a deep copy
            const copyContent = JSON.parse(JSON.stringify(current));
            copyContent.id = 'content_' + Date.now();
            copyContent.name = (copyContent.name || 'Content') + ' (Copy)';
            
            contents.push(copyContent);
            activeContentId = copyContent.id;
            localStorage.setItem('dfoActiveContentId', activeContentId);
            
            logAction('duplicate_content', `duplicated content "${current.name}" as "${copyContent.name}".`);
            
            saveContents();
            renderContentSelect();
            updateAdminUI();
            updateAllViews();
            showToast(`✅ Duplicated into "${copyContent.name}"`, 'success');
        }, 'Duplicate Content', 'primary');
    });

    contentModalCancel.addEventListener('click', () => {
        contentModal.style.display = 'none';
    });

    contentForm.addEventListener('submit', (e) => {
        e.preventDefault();
        if (!isAdmin) return;

        const name = contentNameInput.value.trim();
        const clubLimit = contentClubLimitInput.value;
        const partySize = parseInt(contentPartySizeInput.value) || 3;

        if (!name) return;

        if (editingContentId) {
            const current = contents.find(c => c.id === editingContentId);
            if (current) {
                const oldName = current.name;
                current.name = name;
                current.clubLimit = clubLimit;
                current.partySize = partySize;
                logAction('edit_content', `edited content "${oldName}" to "${name}".`);
            }
        } else {
            const newContent = {
                id: 'content_' + Date.now(),
                name,
                clubLimit,
                partySize,
                raidCounter: 0,
                raids: []
            };
            contents.push(newContent);
            activeContentId = newContent.id;
            localStorage.setItem('dfoActiveContentId', activeContentId);
            logAction('add_content', `added a new content "${name}".`);
        }

        contentModal.style.display = 'none';
        saveContents();
        renderContentSelect();
        updateAdminUI();
        updateAllViews();
    });

    deleteContentBtn.addEventListener('click', () => {
        if (!isAdmin) return;
        const current = getActiveContent();
        if (!current) return;

        showConfirm(`Are you sure you want to delete the content "${current.name}"? This removes all raids within it.`, () => {
            contents = contents.filter(c => c.id !== current.id);
            logAction('delete_content', `deleted content "${current.name}".`);

            if (contents.length > 0) {
                activeContentId = contents[0].id;
            } else {
                activeContentId = null;
            }
            localStorage.setItem('dfoActiveContentId', activeContentId || '');

            saveContents();
            renderContentSelect();
            updateAdminUI();
            updateAllViews();
        }, 'Delete Content', 'danger');
    });

    // --- Global Controls (Export / Import) ---
    exportJsonBtn.addEventListener('click', async () => {
        const state = {
            partyPlan: Array.from(partyPlan.entries()),
            hiddenCharacters: Array.from(hiddenCharacters),
            contents: contents
        };
        const jsonStr = JSON.stringify(state, null, 2);

        // Generate formatted date YYYY-MM-DD
        const date = new Date();
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');

        const activeContent = getActiveContent();
        const contentName = activeContent ? activeContent.name.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase() : 'all';
        const filename = `dfo_planner_${contentName}_${yyyy}-${mm}-${dd}.json`;

        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: filename,
                    types: [{
                        description: 'JSON File',
                        accept: { 'application/json': ['.json'] },
                    }],
                });
                const writable = await handle.createWritable();
                await writable.write(jsonStr);
                await writable.close();
                return;
            } catch (err) {
                if (err.name === 'AbortError') return; // User cancelled
                console.error('File System API failed, falling back:', err);
            }
        }

        // Fallback for browsers that do not support showSaveFilePicker
        const blob = new Blob([jsonStr], { type: "application/json" });
        const url = URL.createObjectURL(blob);

        const dlAnchorElem = document.createElement('a');
        dlAnchorElem.style.display = 'none';
        dlAnchorElem.href = url;
        dlAnchorElem.setAttribute("download", filename);
        document.body.appendChild(dlAnchorElem);
        dlAnchorElem.click();

        setTimeout(() => {
            document.body.removeChild(dlAnchorElem);
            URL.revokeObjectURL(url);
        }, 500);
    });

    exportExcelBtn.addEventListener('click', () => {
        const current = getActiveContent();

        if (!current || !current.raids || current.raids.length === 0) {
            alert('No raids to export in the current content.');
            return;
        }

        const wb = XLSX.utils.book_new();
        const ws_data = [];
        const merges = [];

        const numRaids = current.raids.length;
        const leftCount = Math.ceil(numRaids / 2);

        let leftRow = 0;
        let rightRow = 0;

        function setCell(r, c, val, style, type = 's', format = undefined) {
            if (!ws_data[r]) ws_data[r] = [];
            const cell = { v: val, t: type, s: style };
            if (format) cell.z = format;
            ws_data[r][c] = cell;
        }

        current.raids.forEach((raid, rIndex) => {
            const isRight = rIndex >= leftCount;
            let currentRow = isRight ? rightRow : leftRow;
            const startCol = isRight ? 4 : 0; // Col 0 or Col 4

            const raidName = raid.name || `Raid ${rIndex + 1}`;

            // Raid Title Array Merge
            merges.push({ s: { r: currentRow, c: startCol }, e: { r: currentRow, c: startCol + 2 } });

            const titleStyle = {
                font: { bold: true, color: { rgb: "FFFFFF" }, sz: 16 },
                fill: { fgColor: { rgb: "333333" } },
                alignment: { vertical: "center" }
            };
            setCell(currentRow, startCol, raidName, titleStyle);
            setCell(currentRow, startCol + 1, "", titleStyle);
            setCell(currentRow, startCol + 2, "", titleStyle);
            currentRow++;

            // Headers
            const headerStyle = {
                font: { color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "595959" } }
            };
            setCell(currentRow, startCol, "Explorer club", headerStyle);
            setCell(currentRow, startCol + 1, "Char", headerStyle);
            setCell(currentRow, startCol + 2, "DPS or buffer score (dfogang)", headerStyle);
            currentRow++;

            let raidSum = 0;

            // Parties
            const colors = [
                { rgb: "C00000" }, // Red
                { rgb: "FFFF00" }, // Yellow
                { rgb: "00B050" }  // Green
            ];
            const textColors = [
                { rgb: "FFFFFF" }, // White text on Red
                { rgb: "000000" }, // Black text on Yellow
                { rgb: "FFFFFF" }  // White text on Green
            ];

            raid.parties.forEach((party, pIndex) => {
                const bgFill = colors[pIndex] || { rgb: "CCCCCC" };
                const fontColor = textColors[pIndex] || { rgb: "000000" };
                const rowStyle = {
                    fill: { fgColor: bgFill },
                    font: { color: fontColor }
                };

                for (let sIdx = 0; sIdx < 4; sIdx++) {
                    const charId = party.slots[sIdx];
                    let ec = "";
                    let nameJob = "";
                    let scoreVal = null;

                    if (charId && partyPlan.has(charId)) {
                        const char = partyPlan.get(charId);
                        ec = char.adventureName || "?";
                        nameJob = `${char.characterName}`;
                        if (char.total_buff_score != null) {
                            scoreVal = char.total_buff_score;
                        } else if (char.dps && char.dps.normal) {
                            scoreVal = char.dps.normal;
                        }
                        raidSum += (scoreVal || 0);
                    }

                    setCell(currentRow, startCol, ec, rowStyle);
                    setCell(currentRow, startCol + 1, nameJob, rowStyle);

                    if (charId) {
                        setCell(currentRow, startCol + 2, scoreVal, rowStyle, 'n', '#,##0');
                    } else {
                        setCell(currentRow, startCol + 2, "", rowStyle);
                    }
                    currentRow++;
                }
            });

            // Total Row
            const totalStyle = {
                fill: { fgColor: { rgb: "404040" } },
                font: { bold: true, color: { rgb: "FFFFFF" } }
            };
            setCell(currentRow, startCol, "", totalStyle);
            setCell(currentRow, startCol + 1, "", totalStyle);
            setCell(currentRow, startCol + 2, raidSum, totalStyle, 'n', '#,##0');
            currentRow++;

            // Blank Row space
            currentRow++;

            if (isRight) {
                rightRow = currentRow;
            } else {
                leftRow = currentRow;
            }
        });

        // --- UNASSIGNED ROW GENERATOR ---
        let maxRow = Math.max(leftRow, rightRow) + 2;

        const assignedIds = getAssignedChars();
        const unassignedChars = [];

        partyPlan.forEach((charData, charId) => {
            if (!assignedIds.has(charId)) {
                unassignedChars.push(charData);
            }
        });

        if (unassignedChars.length > 0) {
            unassignedChars.sort((a, b) => {
                const getScore = (c) => c.total_buff_score != null ? c.total_buff_score : (c.dps && c.dps.normal ? c.dps.normal : 0);
                return getScore(b) - getScore(a); // Descending score
            });

            const unassignedTitleStyle = {
                font: { bold: true, color: { rgb: "FFFFFF" }, sz: 14 },
                fill: { fgColor: { rgb: "333333" } },
                alignment: { vertical: "center" }
            };

            merges.push({ s: { r: maxRow, c: 0 }, e: { r: maxRow, c: 2 } });
            setCell(maxRow, 0, "Unassigned Pool", unassignedTitleStyle);
            setCell(maxRow, 1, "", unassignedTitleStyle);
            setCell(maxRow, 2, "", unassignedTitleStyle);
            maxRow++;

            const headerStyle = {
                font: { color: { rgb: "FFFFFF" } },
                fill: { fgColor: { rgb: "595959" } }
            };
            setCell(maxRow, 0, "Explorer club", headerStyle);
            setCell(maxRow, 1, "Char", headerStyle);
            setCell(maxRow, 2, "DPS or buffer score (dfogang)", headerStyle);
            maxRow++;

            const emptyStyle = { font: { color: { rgb: "000000" } } };

            unassignedChars.forEach(char => {
                const ec = char.adventureName || "?";
                const name = char.characterName;
                let scoreVal = null;

                if (char.total_buff_score != null) {
                    scoreVal = char.total_buff_score;
                } else if (char.dps && char.dps.normal) {
                    scoreVal = char.dps.normal;
                }

                setCell(maxRow, 0, ec, emptyStyle);
                setCell(maxRow, 1, name, emptyStyle);
                setCell(maxRow, 2, scoreVal, emptyStyle, 'n', '#,##0');
                maxRow++;
            });
        }

        // Ensure all rows are arrays even if empty
        for (let i = 0; i < maxRow; i++) {
            if (!ws_data[i]) ws_data[i] = [];
            // Fill entirely to avoid sparse array issues in parsing
            for (let j = 0; j <= 6; j++) {
                if (!ws_data[i][j]) ws_data[i][j] = { v: "", t: "s" };
            }
        }

        const ws = XLSX.utils.aoa_to_sheet(ws_data);
        ws['!merges'] = merges;
        ws['!cols'] = [
            { wch: 18 }, // A: Explorer club
            { wch: 35 }, // B: Char
            { wch: 30 }, // C: DPS
            { wch: 3 }, // D: Spacer
            { wch: 18 }, // E: Explorer club
            { wch: 35 }, // F: Char
            { wch: 30 }  // G: DPS
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Raids");

        const date = new Date();
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const filename = `raid_roster_${current.name.replace(/[^a-zA-Z0-9_-]/g, '_')}_${yyyy}-${mm}-${dd}.xlsx`;

        XLSX.writeFile(wb, filename);
    });

    importJsonBtn.addEventListener('click', () => {
        importFileInput.click();
    });

    importFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.json') && file.type !== 'application/json') {
            alert('Invalid file format. Please upload a structured .json planner file.');
            importFileInput.value = '';
            return;
        }

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const state = JSON.parse(e.target.result);
                partyPlan.clear();
                if (state.partyPlan) {
                    if (Array.isArray(state.partyPlan)) {
                        state.partyPlan.forEach(([id, data]) => partyPlan.set(id, data));
                    } else {
                        Object.entries(state.partyPlan).forEach(([id, data]) => partyPlan.set(id, data));
                    }
                }

                hiddenCharacters.clear();
                if (state.hiddenCharacters) state.hiddenCharacters.forEach(id => hiddenCharacters.add(id));

                if (state.contents) {
                    contents = state.contents;
                } else if (state.raids) {
                    // Legacy import migration
                    contents = [];
                    const migratedRaids = [];
                    state.raids.forEach(r => {
                        if (r.parties && r.parties.length > 0 && !r.parties[0].slots) {
                            r.parties = r.parties.map(p => ({
                                slots: [p.dps1 || null, p.dps2 || null, p.dps3 || null, p.buff || null]
                            }));
                        }
                        migratedRaids.push(r);
                    });

                    contents.push({
                        id: 'content_imported_' + Date.now(),
                        name: 'Imported Legacy Data',
                        clubLimit: '',
                        partySize: 3,
                        raidCounter: state.raidCounter || 0,
                        raids: migratedRaids
                    });
                }

                if (contents.length > 0) {
                    activeContentId = contents[0].id;
                } else {
                    activeContentId = null;
                }
                localStorage.setItem('dfoActiveContentId', activeContentId || '');

                renderContentSelect();
                updateAllViews();
                saveContents();
                savePartyPlan();
                logAction('import_json', `imported a planner backup.`);
                showToast('Backup imported successfully!', 'success');
            } catch (err) {
                alert('Invalid backup file. Could not import data.');
                console.error(err);
            }
            importFileInput.value = '';
        };
        reader.readAsText(file);
    });

    // --- Raid Controls ---
    addRaidBtn.addEventListener('click', () => {
        const current = getActiveContent();
        if (!current) return;

        current.raidCounter++;
        const partiesCount = current.partySize || 3;
        const size = partiesCount * 4;
        const parties = [];
        for (let i = 0; i < partiesCount; i++) {
            parties.push({ slots: [null, null, null, null] });
        }
        current.raids.push({ id: current.raidCounter, size, parties });
        updateAllViews();
        saveContents();
        logAction('add_raid', `added a new Raid (#${current.raidCounter}) to "${current.name}".`);
    });

    // --- Auto Assign Everything ---
    const autoAssignModal = document.getElementById('auto-assign-modal');
    const autoAssignForm = document.getElementById('auto-assign-form');
    const autoAssignCancel = document.getElementById('auto-assign-cancel');
    const autoAssignPartiesConfig = document.getElementById('auto-assign-parties-config');

    const PARTY_COLORS = ['red', 'yellow', 'green'];
    const PARTY_NAMES = ['Red Party', 'Yellow Party', 'Green Party'];

    if (autoAssignCancel) {
        autoAssignCancel.addEventListener('click', () => {
            autoAssignModal.style.display = 'none';
        });
    }

    function buildAutoAssignPartyFields(partySize) {
        let html = '<div class="auto-assign-parties-grid">';
        for (let p = 0; p < partySize; p++) {
            const color = PARTY_COLORS[p] || 'gray';
            const name = PARTY_NAMES[p] || `Party ${p + 1}`;
            html += `
                <div class="auto-assign-party-card party-${color}">
                    <div class="auto-assign-party-title">${name}</div>
                    <div class="auto-assign-party-fields">
                        <div class="form-field">
                            <label class="form-label">Target DPS <span class="form-hint">(M)</span></label>
                            <input type="number" class="form-input auto-target-dps" data-party="${p}" placeholder="e.g. 150" min="0" step="any">
                        </div>
                        <div class="form-field">
                            <label class="form-label">Target Buff <span class="form-hint">(M)</span></label>
                            <input type="number" class="form-input auto-target-buff" data-party="${p}" placeholder="e.g. 50" min="0" step="any">
                        </div>
                    </div>
                </div>
            `;
        }
        html += '</div>';
        autoAssignPartiesConfig.innerHTML = html;
    }

    if (autoPlanBtn) autoPlanBtn.addEventListener('click', () => {
        if (!isAdmin) return;
        const current = getActiveContent();
        if (!current) return;

        const P = current.partySize || 3;
        buildAutoAssignPartyFields(P);

        // Show configuration modal
        autoAssignModal.style.display = 'flex';
    });

    if (autoAssignForm) autoAssignForm.addEventListener('submit', (e) => {
        e.preventDefault();
        if (!isAdmin) return;

        const current = getActiveContent();
        if (!current) return;

        const P = current.partySize || 3;

        // Read per-party targets
        const partyTargets = [];
        for (let p = 0; p < P; p++) {
            const dpsInput = autoAssignPartiesConfig.querySelector(`.auto-target-dps[data-party="${p}"]`);
            const buffInput = autoAssignPartiesConfig.querySelector(`.auto-target-buff[data-party="${p}"]`);
            const targetDps = (parseFloat(dpsInput?.value) || 0) * 1_000_000;
            const targetBuff = (parseFloat(buffInput?.value) || 0) * 1_000_000;
            partyTargets.push({ dps: targetDps, buff: targetBuff });
        }

        autoAssignModal.style.display = 'none';

        const existingRaidCount = current.raids.length;
        const warningMsg = existingRaidCount > 0
            ? `This will permanently replace all ${existingRaidCount} existing raid(s) in "${current.name}" with auto-generated assignments. This cannot be undone.`
            : `This will auto-generate raid assignments for all characters in the roster for "${current.name}".`;

        showConfirm(warningMsg, () => {
            runAutoAssign(current, partyTargets);
        }, 'Run Auto Assign', 'primary');
    });

    function runAutoAssign(content, partyTargets) {
        // --- Step 1: Gather pools ---
        const poolDPS = [];
        const poolBuffers = [];

        partyPlan.forEach((charData, charId) => {
            const isBuffer = charData.total_buff_score != null;
            if (isBuffer) {
                poolBuffers.push({
                    id: charId,
                    data: charData,
                    adv: charData.adventureName,
                    power: charData.total_buff_score || 0
                });
            } else {
                poolDPS.push({
                    id: charId,
                    data: charData,
                    adv: charData.adventureName,
                    power: charData.dps?.normal || 0
                });
            }
        });

        if (poolDPS.length === 0 && poolBuffers.length === 0) {
            showToast('No characters available to auto-assign!', 'warning');
            return;
        }

        // Sort both pools descending by score
        poolDPS.sort((a, b) => b.power - a.power);
        poolBuffers.sort((a, b) => b.power - a.power);

        const P = content.partySize || 3; // parties per raid
        const DPS_PER_PARTY = 3;
        const BUFF_PER_PARTY = 1;
        const dpsPerRaid = DPS_PER_PARTY * P;
        const size = P * 4;

        const globalLimitStr = content.clubLimit;
        const globalLimit = (globalLimitStr && parseInt(globalLimitStr) > 0) ? parseInt(globalLimitStr) : Infinity;

        // --- Step 2: Determine max raids (driven by DPS count) ---
        let maxN = Math.max(1, Math.floor(poolDPS.length / dpsPerRaid));

        // --- Step 3: Find best N where all DPS fit respecting club limits ---
        let raidBuckets = null;
        let globalClubUsage = new Map();

        for (let N = maxN; N >= 1; N--) {
            const testBuckets = [];
            for (let i = 0; i < N; i++) {
                const parties = [];
                for (let p = 0; p < P; p++) {
                    parties.push({
                        dps: [], buffers: [], advNames: new Set(),
                        dpsSum: 0, buffSum: 0,
                        partyIdx: p // Track which party slot (Red=0, Yellow=1, Green=2)
                    });
                }
                testBuckets.push({ parties, raidAdvNames: new Set() });
            }

            const clubUsage = new Map();

            // Assign DPS via greedy balancing using per-party targets
            for (const char of poolDPS) {
                const currentUsage = clubUsage.get(char.adv) || 0;
                if (currentUsage >= globalLimit) continue;

                let bestParty = null;
                let bestRaidIdx = -1;
                let bestDeviation = Infinity;

                for (let rIdx = 0; rIdx < testBuckets.length; rIdx++) {
                    const raid = testBuckets[rIdx];
                    if (raid.raidAdvNames.has(char.adv)) continue;

                    for (const party of raid.parties) {
                        if (party.dps.length >= DPS_PER_PARTY) continue;

                        // Use this party's specific target
                        const target = partyTargets[party.partyIdx]?.dps || 0;
                        const deviation = party.dpsSum - target;
                        if (deviation < bestDeviation) {
                            bestDeviation = deviation;
                            bestParty = party;
                            bestRaidIdx = rIdx;
                        }
                    }
                }

                if (bestParty) {
                    bestParty.dps.push(char);
                    bestParty.dpsSum += char.power;
                    bestParty.advNames.add(char.adv);
                    testBuckets[bestRaidIdx].raidAdvNames.add(char.adv);
                    clubUsage.set(char.adv, currentUsage + 1);
                }
            }

            // Check if all raids are full (DPS-wise)
            let allAssigned = true;
            for (const bucket of testBuckets) {
                for (const party of bucket.parties) {
                    if (party.dps.length < DPS_PER_PARTY) {
                        allAssigned = false;
                        break;
                    }
                }
                if (!allAssigned) break;
            }

            if (allAssigned) {
                raidBuckets = testBuckets;
                globalClubUsage = clubUsage;
                break;
            }
        }

        // Fallback: if even 1 raid can't fill, use 1 raid anyway
        if (!raidBuckets) {
            raidBuckets = [];
            const parties = [];
            for (let p = 0; p < P; p++) {
                parties.push({
                    dps: [], buffers: [], advNames: new Set(),
                    dpsSum: 0, buffSum: 0, partyIdx: p
                });
            }
            raidBuckets.push({ parties, raidAdvNames: new Set() });
            globalClubUsage = new Map();

            for (const char of poolDPS) {
                const currentUsage = globalClubUsage.get(char.adv) || 0;
                if (currentUsage >= globalLimit) continue;

                let bestParty = null;
                let bestDeviation = Infinity;

                for (const party of raidBuckets[0].parties) {
                    if (party.dps.length >= DPS_PER_PARTY) continue;
                    const target = partyTargets[party.partyIdx]?.dps || 0;
                    const deviation = party.dpsSum - target;
                    if (deviation < bestDeviation) {
                        bestDeviation = deviation;
                        bestParty = party;
                    }
                }
                if (bestParty) {
                    bestParty.dps.push(char);
                    bestParty.dpsSum += char.power;
                    bestParty.advNames.add(char.adv);
                    raidBuckets[0].raidAdvNames.add(char.adv);
                    globalClubUsage.set(char.adv, currentUsage + 1);
                }
            }
        }

        // --- Step 4: Assign Buffers using per-party targets ---
        for (const char of poolBuffers) {
            const currentUsage = globalClubUsage.get(char.adv) || 0;
            if (currentUsage >= globalLimit) continue;

            let bestParty = null;
            let bestRaidIdx = -1;
            let bestDeviation = Infinity;

            for (let rIdx = 0; rIdx < raidBuckets.length; rIdx++) {
                const raid = raidBuckets[rIdx];
                if (raid.raidAdvNames.has(char.adv)) continue;

                for (const party of raid.parties) {
                    if (party.buffers.length >= BUFF_PER_PARTY) continue;

                    const target = partyTargets[party.partyIdx]?.buff || 0;
                    const deviation = party.buffSum - target;
                    if (deviation < bestDeviation) {
                        bestDeviation = deviation;
                        bestParty = party;
                        bestRaidIdx = rIdx;
                    }
                }
            }

            if (bestParty) {
                bestParty.buffers.push(char);
                bestParty.buffSum += char.power;
                bestParty.advNames.add(char.adv);
                raidBuckets[bestRaidIdx].raidAdvNames.add(char.adv);
                globalClubUsage.set(char.adv, currentUsage + 1);
            }
        }

        // --- Step 5: Gap Minimization (consecutive raids per explorer) ---
        const explorerRaidMap = new Map();
        raidBuckets.forEach((raid, rIdx) => {
            raid.parties.forEach(party => {
                [...party.dps, ...party.buffers].forEach(char => {
                    if (!explorerRaidMap.has(char.adv)) explorerRaidMap.set(char.adv, new Set());
                    explorerRaidMap.get(char.adv).add(rIdx);
                });
            });
        });

        const MAX_SWAP_ITERATIONS = 200;
        let swapCount = 0;

        for (const [advName, raidIndices] of explorerRaidMap) {
            if (swapCount >= MAX_SWAP_ITERATIONS) break;

            const indices = Array.from(raidIndices).sort((a, b) => a - b);
            if (indices.length <= 1) continue;

            const span = indices[indices.length - 1] - indices[0];
            const idealSpan = indices.length - 1;
            if (span <= idealSpan) continue;

            for (let i = 0; i < indices.length && swapCount < MAX_SWAP_ITERATIONS; i++) {
                const currentRaidIdx = indices[i];
                const targetStart = indices[0];
                const idealIdx = targetStart + i;

                if (currentRaidIdx === idealIdx) continue;
                if (idealIdx < 0 || idealIdx >= raidBuckets.length) continue;

                const sourceRaid = raidBuckets[currentRaidIdx];
                const targetRaid = raidBuckets[idealIdx];

                let sourceChar = null;
                let sourceParty = null;
                let sourceIsBuffer = false;

                for (const party of sourceRaid.parties) {
                    const dpsMatch = party.dps.find(c => c.adv === advName);
                    if (dpsMatch) {
                        sourceChar = dpsMatch;
                        sourceParty = party;
                        sourceIsBuffer = false;
                        break;
                    }
                    const buffMatch = party.buffers.find(c => c.adv === advName);
                    if (buffMatch) {
                        sourceChar = buffMatch;
                        sourceParty = party;
                        sourceIsBuffer = true;
                        break;
                    }
                }

                if (!sourceChar) continue;
                if (targetRaid.raidAdvNames.has(advName)) continue;

                let bestSwapCandidate = null;
                let bestSwapParty = null;
                let bestSwapScore = Infinity;

                for (const tParty of targetRaid.parties) {
                    const pool = sourceIsBuffer ? tParty.buffers : tParty.dps;
                    for (const candidate of pool) {
                        if (sourceRaid.raidAdvNames.has(candidate.adv) && candidate.adv !== advName) continue;
                        if (candidate.adv !== advName && sourceRaid.raidAdvNames.has(candidate.adv)) continue;

                        const scoreDiff = Math.abs(sourceChar.power - candidate.power);
                        if (scoreDiff < bestSwapScore) {
                            bestSwapScore = scoreDiff;
                            bestSwapCandidate = candidate;
                            bestSwapParty = tParty;
                        }
                    }
                }

                if (!bestSwapCandidate) continue;

                // Tolerance check using per-party targets
                const sourceTargetKey = sourceIsBuffer ? 'buff' : 'dps';
                const sourceTarget = partyTargets[sourceParty.partyIdx]?.[sourceTargetKey] || 0;
                const swapTarget = partyTargets[bestSwapParty.partyIdx]?.[sourceTargetKey] || 0;

                if (sourceTarget > 0 || swapTarget > 0) {
                    const scoreKey = sourceIsBuffer ? 'buffSum' : 'dpsSum';
                    const tolerance = Math.max(sourceTarget, swapTarget) * 0.15;
                    const newSourceScore = sourceParty[scoreKey] - sourceChar.power + bestSwapCandidate.power;
                    const newSwapScore = bestSwapParty[scoreKey] - bestSwapCandidate.power + sourceChar.power;

                    if ((sourceTarget > 0 && Math.abs(newSourceScore - sourceTarget) > tolerance + Math.abs(sourceParty[scoreKey] - sourceTarget)) ||
                        (swapTarget > 0 && Math.abs(newSwapScore - swapTarget) > tolerance + Math.abs(bestSwapParty[scoreKey] - swapTarget))) {
                        continue;
                    }
                }

                // Perform the swap
                const sourcePool = sourceIsBuffer ? sourceParty.buffers : sourceParty.dps;
                const targetPool = sourceIsBuffer ? bestSwapParty.buffers : bestSwapParty.dps;
                const scoreKey = sourceIsBuffer ? 'buffSum' : 'dpsSum';

                const srcIdx = sourcePool.indexOf(sourceChar);
                if (srcIdx !== -1) sourcePool.splice(srcIdx, 1);
                sourceParty[scoreKey] -= sourceChar.power;
                sourceParty.advNames.delete(sourceChar.adv);

                const tgtIdx = targetPool.indexOf(bestSwapCandidate);
                if (tgtIdx !== -1) targetPool.splice(tgtIdx, 1);
                bestSwapParty[scoreKey] -= bestSwapCandidate.power;
                bestSwapParty.advNames.delete(bestSwapCandidate.adv);

                targetPool.push(sourceChar);
                bestSwapParty[scoreKey] += sourceChar.power;
                bestSwapParty.advNames.add(sourceChar.adv);

                sourcePool.push(bestSwapCandidate);
                sourceParty[scoreKey] += bestSwapCandidate.power;
                sourceParty.advNames.add(bestSwapCandidate.adv);

                sourceRaid.raidAdvNames = new Set();
                sourceRaid.parties.forEach(p => {
                    [...p.dps, ...p.buffers].forEach(c => sourceRaid.raidAdvNames.add(c.adv));
                });
                targetRaid.raidAdvNames = new Set();
                targetRaid.parties.forEach(p => {
                    [...p.dps, ...p.buffers].forEach(c => targetRaid.raidAdvNames.add(c.adv));
                });

                swapCount++;
            }
        }

        // --- Step 6: Collect unassigned into a leftover raid ---
        const assignedIds = new Set();
        raidBuckets.forEach(raid => {
            raid.parties.forEach(party => {
                [...party.dps, ...party.buffers].forEach(c => assignedIds.add(c.id));
            });
        });

        const leftoverDPS = poolDPS.filter(c => !assignedIds.has(c.id));
        const leftoverBuffers = poolBuffers.filter(c => !assignedIds.has(c.id));

        if (leftoverDPS.length > 0 || leftoverBuffers.length > 0) {
            const leftoverParties = [];
            for (let p = 0; p < P; p++) {
                leftoverParties.push({
                    dps: [], buffers: [], advNames: new Set(),
                    dpsSum: 0, buffSum: 0, partyIdx: p
                });
            }
            const leftoverRaid = { parties: leftoverParties, raidAdvNames: new Set() };

            for (const char of leftoverDPS) {
                // Check global club limit
                const currentUsage = globalClubUsage.get(char.adv) || 0;
                if (currentUsage >= globalLimit) continue;

                // Check raid-level club uniqueness
                if (leftoverRaid.raidAdvNames.has(char.adv)) continue;

                let bestParty = null;
                let bestDeviation = Infinity;
                for (const party of leftoverParties) {
                    if (party.dps.length >= DPS_PER_PARTY) continue;
                    const deviation = party.dpsSum;
                    if (deviation < bestDeviation) {
                        bestDeviation = deviation;
                        bestParty = party;
                    }
                }
                if (bestParty) {
                    bestParty.dps.push(char);
                    bestParty.dpsSum += char.power;
                    bestParty.advNames.add(char.adv);
                    leftoverRaid.raidAdvNames.add(char.adv);
                    globalClubUsage.set(char.adv, currentUsage + 1);
                }
            }

            for (const char of leftoverBuffers) {
                // Check global club limit
                const currentUsage = globalClubUsage.get(char.adv) || 0;
                if (currentUsage >= globalLimit) continue;

                // Check raid-level club uniqueness
                if (leftoverRaid.raidAdvNames.has(char.adv)) continue;

                let bestParty = null;
                let bestDeviation = Infinity;
                for (const party of leftoverParties) {
                    if (party.buffers.length >= BUFF_PER_PARTY) continue;
                    const deviation = party.buffSum;
                    if (deviation < bestDeviation) {
                        bestDeviation = deviation;
                        bestParty = party;
                    }
                }
                if (bestParty) {
                    bestParty.buffers.push(char);
                    bestParty.buffSum += char.power;
                    bestParty.advNames.add(char.adv);
                    leftoverRaid.raidAdvNames.add(char.adv);
                    globalClubUsage.set(char.adv, currentUsage + 1);
                }
            }

            const hasContent = leftoverParties.some(p => p.dps.length > 0 || p.buffers.length > 0);
            if (hasContent) {
                raidBuckets.push(leftoverRaid);
            }
        }

        if (raidBuckets.length === 0) {
            showToast('No characters could be assigned!', 'warning');
            return;
        }

        // --- Step 7: Build actual raid data ---
        content.raids.length = 0;
        content.raidCounter = 0;

        for (const bucket of raidBuckets) {
            content.raidCounter++;
            const parties = [];

            for (let pIdx = 0; pIdx < P; pIdx++) {
                const slots = [null, null, null, null];
                const partyData = bucket.parties[pIdx];

                partyData.dps.sort((a, b) => b.power - a.power);
                partyData.buffers.sort((a, b) => b.power - a.power);

                for (let s = 0; s < Math.min(partyData.dps.length, 3); s++) {
                    slots[s] = partyData.dps[s].id;
                }
                if (partyData.buffers.length > 0) {
                    slots[3] = partyData.buffers[0].id;
                }

                parties.push({ slots });
            }

            content.raids.push({ id: content.raidCounter, size, parties });
        }

        const filledCount = raidBuckets.length;
        const totalDpsAssigned = raidBuckets.reduce((sum, r) => sum + r.parties.reduce((s, p) => s + p.dps.length, 0), 0);
        const totalBuffAssigned = raidBuckets.reduce((sum, r) => sum + r.parties.reduce((s, p) => s + p.buffers.length, 0), 0);

        updateAllViews();
        saveContents();
        logAction('auto_plan', `auto-assigned ${totalDpsAssigned} DPS + ${totalBuffAssigned} buffers into ${filledCount} raid(s) in "${content.name}".`);
        showToast(`✅ Auto-assign complete! ${filledCount} raid(s) — ${totalDpsAssigned} DPS + ${totalBuffAssigned} buffers assigned.`, 'success');
        if (raidsContainer.offsetTop) {
            window.scrollTo({ top: raidsContainer.offsetTop - 50, behavior: 'smooth' });
        }
    }


    raidsContainer.addEventListener('click', (e) => {
        // Remove raid
        if (e.target.classList.contains('raid-remove-btn')) {
            const current = getActiveContent();
            if (!current) return;

            const rId = parseInt(e.target.dataset.raidId);
            const raid = current.raids.find(r => r.id === rId);
            const raidName = raid ? (raid.name || `#${rId}`) : 'this raid';

            showConfirm(`Are you sure you want to remove ${raidName} from "${current.name}"?`, () => {
                const idx = current.raids.findIndex(r => r.id === rId);
                if (idx !== -1) {
                    current.raids.splice(idx, 1);
                    updateAllViews();
                    saveContents();
                    logAction('remove_raid', `removed ${raidName} from "${current.name}".`);
                }
            }, 'Remove Raid', 'primary');
            return;
        }

        // Remove card from slot via X button
        if (e.target.closest('.remove-btn')) {
            const card = e.target.closest('.result-card');
            const slot = card.closest('.party-slot');
            if (slot) {
                setCharInSlot(slot.dataset.slotId, null);
                updateAllViews();
                saveContents();
            }
        }
    });

    raidsContainer.addEventListener('change', (e) => {
        if (e.target.classList.contains('raid-title-input')) {
            const current = getActiveContent();
            if (!current) return;

            const rId = parseInt(e.target.dataset.raidId);
            const raid = current.raids.find(r => r.id === rId);
            if (raid) {
                const oldName = raid.name || `#${rId}`;
                raid.name = e.target.value.trim();
                saveContents();
                logAction('rename_raid', `renamed raid from "${oldName}" to "${raid.name}" in "${current.name}".`);
            }
        }
    });

    // --- Helpers ---
    function showConfirm(message, onConfirm, confirmText = 'Confirm', type = 'primary') {
        modalMessage.textContent = message;
        modalConfirmBtn.textContent = confirmText;
        modalConfirmBtn.className = `btn btn-${type}`;
        confirmModal.style.display = 'flex';

        const handleConfirm = () => {
            onConfirm();
            close();
        };

        const handleCancel = () => close();

        const close = () => {
            confirmModal.style.display = 'none';
            modalConfirmBtn.removeEventListener('click', handleConfirm);
            modalCancelBtn.removeEventListener('click', handleCancel);
        };

        modalConfirmBtn.addEventListener('click', handleConfirm);
        modalCancelBtn.addEventListener('click', handleCancel);
    }

    function showToast(message, type = 'error') {
        const container = document.getElementById('toast-container');
        if (!container) return;

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;

        const icon = type === 'error' ? '×' : (type === 'warning' ? '!' : '✓');

        toast.innerHTML = `
            <div class="toast-icon">${icon}</div>
            <div class="toast-content">${message}</div>
        `;

        container.appendChild(toast);

        // Trigger animation safely
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                toast.classList.add('show');
            });
        });

        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 400);
        }, 3500);
    }

    function formatMillions(num) {
        if (num == null) return '—';
        return (num / 1_000_000).toFixed(2) + 'M';
    }

    function getAssignedChars() {
        const assigned = new Set();
        const current = getActiveContent();
        if (!current) return assigned;

        current.raids.forEach(r => r.parties.forEach(p => {
            p.slots.forEach(s => { if (s) assigned.add(s); });
        }));
        return assigned;
    }

    function getCharInSlot(slotId) {
        if (slotId === 'pool') return null;
        const current = getActiveContent();
        if (!current) return null;

        const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        if (!parts) return null;
        const raid = current.raids.find(r => r.id === parseInt(parts[1]));
        if (raid) return raid.parties[parseInt(parts[2])].slots[parseInt(parts[3])];
        return null;
    }

    function setCharInSlot(slotId, charId) {
        if (slotId === 'pool') return;
        const current = getActiveContent();
        if (!current) return;

        const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        if (!parts) return;
        const raid = current.raids.find(r => r.id === parseInt(parts[1]));
        if (raid) raid.parties[parseInt(parts[2])].slots[parseInt(parts[3])] = charId;
    }

    function checkAdventureNameConflict(raidId, charAId, targetSlotId) {
        const current = getActiveContent();
        if (!current) return null;

        const raid = current.raids.find(r => r.id === raidId);
        if (!raid) return null;

        const charA = partyPlan.get(charAId);
        if (!charA || !charA.adventureName) return null;
        const advA = charA.adventureName;

        const parts = targetSlotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        const targetPIdx = parts ? parseInt(parts[2]) : -1;

        for (let pIdx = 0; pIdx < raid.parties.length; pIdx++) {
            const party = raid.parties[pIdx];
            for (let sIdx = 0; sIdx < party.slots.length; sIdx++) {
                const currentSlotId = `raid-${raid.id}-party-${pIdx}-slot-${sIdx}`;
                if (currentSlotId === targetSlotId) continue;

                const occupantId = party.slots[sIdx];
                if (occupantId && occupantId !== charAId) {
                    const occupant = partyPlan.get(occupantId);
                    if (occupant && occupant.adventureName === advA) {
                        return pIdx === targetPIdx ? 'party' : 'raid';
                    }
                }
            }
        }
        return null;
    }

    function isGlobalAdventureNameLimitReached(advName, excludeCharId = null) {
        const current = getActiveContent();
        if (!current) return false;

        const limitStr = current.clubLimit;
        if (!limitStr) return false;
        const limit = parseInt(limitStr);
        if (isNaN(limit) || limit <= 0) return false;

        let currentCount = 0;
        current.raids.forEach(r => r.parties.forEach(p => {
            p.slots.forEach(charId => {
                if (charId && charId !== excludeCharId && partyPlan.has(charId)) {
                    if (partyPlan.get(charId).adventureName === advName) {
                        currentCount++;
                    }
                }
            });
        }));

        return currentCount >= limit;
    }

    function renderClubSummary() {
        const allCounts = new Map();
        let totalAssigned = 0;
        const current = getActiveContent();
        if (!current) {
            clubSummaryContainer.style.display = 'none';
            return;
        }

        current.raids.forEach(r => r.parties.forEach(p => {
            p.slots.forEach(charId => {
                if (charId && partyPlan.has(charId)) {
                    const adv = partyPlan.get(charId).adventureName;
                    if (adv) {
                        allCounts.set(adv, (allCounts.get(adv) || 0) + 1);
                        totalAssigned++;
                    }
                }
            });
        }));

        if (totalAssigned === 0) {
            clubSummaryContainer.style.display = 'none';
            return;
        }

        clubSummaryContainer.style.display = 'block';

        const sortedCounts = Array.from(allCounts.entries()).sort((a, b) => b[1] - a[1]);

        let html = '';
        const limitStr = current.clubLimit;
        const limit = (limitStr && parseInt(limitStr) > 0) ? parseInt(limitStr) : null;

        sortedCounts.forEach(([adv, count]) => {
            const isCapped = limit && count >= limit;
            const bg = isCapped ? 'rgba(220, 38, 38, 0.08)' : 'rgba(37, 99, 235, 0.06)';
            const border = isCapped ? 'rgba(220, 38, 38, 0.25)' : 'rgba(37, 99, 235, 0.15)';
            const color = isCapped ? '#dc2626' : '#2563eb';

            html += `<span class="club-badge" style="background: ${bg}; border: 1px solid ${border}; color: ${color};">
                        ${adv} <strong>${count}${limit ? '/' + limit : ''}</strong>
                     </span>`;
        });
        clubSummaryList.innerHTML = html;
    }

    function removeCharFromRaid(charId) {
        const current = getActiveContent();
        if (!current) return;
        current.raids.forEach(r => r.parties.forEach(p => {
            p.slots = p.slots.map(s => s === charId ? null : s);
        }));
    }

    function getPartyTypeCounts(party, excludeSlotIdx = -1) {
        let dpsCount = 0, buffCount = 0;
        party.slots.forEach((charId, idx) => {
            if (idx === excludeSlotIdx || !charId) return;
            const char = partyPlan.get(charId);
            if (!char) return;
            if (char.total_buff_score != null) buffCount++;
            else dpsCount++;
        });
        return { dpsCount, buffCount };
    }

    function calcPartyTotals(party) {
        let dpsSum = 0, buffSum = 0;
        party.slots.forEach(charId => {
            if (!charId) return;
            const char = partyPlan.get(charId);
            if (!char) return;
            if (char.total_buff_score != null) {
                buffSum += char.total_buff_score;
            } else if (char.dps && char.dps.normal) {
                dpsSum += char.dps.normal;
            }
        });
        return {
            dps: dpsSum ? formatMillions(dpsSum) : '—',
            buff: buffSum ? formatMillions(buffSum) : '—'
        };
    }

    function createCardHTML(char, isSelected = false, isDraggable = false, currentSlotId = 'pool') {
        const isBuffer = char.total_buff_score != null;
        const scoreValue = isBuffer ? char.total_buff_score : (char.dps ? char.dps.normal : null);

        let scoreDisplay = formatMillions(scoreValue);

        const removeBtnStr = currentSlotId === 'search' ? '' : '<button class="remove-btn" title="Remove" aria-label="Remove">×</button>';

        return `
            <div class="result-card ${isBuffer ? 'buffer' : 'dealer'} ${isSelected ? 'selected' : ''}" 
                 data-id="${char.characterId}"
                 ${isDraggable ? 'draggable="true"' : ''}>
                ${removeBtnStr}
                <div class="card-info">
                    <div class="card-name">${char.characterName}</div>
                    <div class="card-sub">${char.adventureName || '?'} · ${(char.jobGrowName || char.jobName || '').replace(/^Neo:\s*/i, '')}</div>
                </div>
                <div class="card-stats">
                    <div class="card-score">${scoreDisplay}</div>
                </div>
            </div>
        `;
    }

    // ... [RENDER RAIDS] ...
    function renderRaids() {
        raidsContainer.innerHTML = '';
        const current = getActiveContent();
        
        raidsContainer.className = 'raids-container';
        if (current && current.partySize) {
            raidsContainer.classList.add(`parties-${current.partySize}`);
        }

        if (!current || !current.raids || current.raids.length === 0) {
            raidsContainer.innerHTML = '<div class="empty-raids-msg"><p>No raids yet. Click <strong>Add Raid</strong> or <strong>Auto Assign</strong> to get started.</p></div>';
            return;
        }

        current.raids.forEach((raid, rIndex) => {
            const raidBlock = document.createElement('div');
            raidBlock.className = 'raid-block';
            raidBlock.dataset.raidId = raid.id;

            const colors = ['red', 'yellow', 'green'];
            const names = ['Red', 'Yellow', 'Green'];

            let partiesHtml = '';
            raid.parties.forEach((party, pIndex) => {
                const colorClass = colors[pIndex] || 'gray';
                const partyName = names[pIndex] ? `${names[pIndex]} Party` : `Party ${pIndex + 1}`;

                let slotsHtml = '';
                party.slots.forEach((charId, sIdx) => {
                    const slotId = `raid-${raid.id}-party-${pIndex}-slot-${sIdx}`;
                    let slotContent = '';
                    let emptyClass = 'empty';
                    if (charId && partyPlan.has(charId)) {
                        const char = partyPlan.get(charId);
                        const isBuffer = char.total_buff_score != null;
                        slotContent = createCardHTML(char, false, true, slotId);
                        emptyClass = '';
                    }
                    slotsHtml += `
                        <div class="party-slot ${emptyClass}" data-slot-id="${slotId}">
                            ${slotContent}
                        </div>
                    `;
                });

                const totals = calcPartyTotals(party);

                partiesHtml += `
                    <div class="party-block party-${colorClass}">
                        <div class="party-header">${partyName}</div>
                        ${slotsHtml}
                        <div class="party-footer">
                            <div class="party-footer-row">
                                <span class="party-dmg-label">DPS</span>
                                <span class="party-dmg-value dps-val">${totals.dps}</span>
                            </div>
                            <div class="party-footer-row">
                                <span class="party-dmg-label">Buff</span>
                                <span class="party-dmg-value buff-val">${totals.buff}</span>
                            </div>
                        </div>
                    </div>
                `;
            });

            raidBlock.innerHTML = `
                <div class="raid-header">
                    <input type="text" class="raid-title-input" 
                           value="${raid.name || '#' + (rIndex + 1)}" 
                           placeholder="#${rIndex + 1}" 
                           data-raid-id="${raid.id}"
                           spellcheck="false">
                    <button class="raid-remove-btn" title="Remove Raid" data-raid-id="${raid.id}">×</button>
                </div>
                <div class="raid-parties">
                    ${partiesHtml}
                </div>
            `;
            raidsContainer.appendChild(raidBlock);
        });
    }

    // ... [RENDER RESULTS] ...
    function renderResultCards(characters) {
        searchResults.innerHTML = '';
        currentSearchResults = characters;

        if (characters.length === 0) {
            searchResults.innerHTML = '<p class="no-results">No characters match the requirements.</p>';
            return;
        }

        const visibleChars = characters.filter(c => !hiddenCharacters.has(c.characterId));

        if (visibleChars.length === 0) {
            searchResults.innerHTML = '<p class="no-results">All matching characters have been hidden from results.</p>';
            return;
        }

        let html = '';
        visibleChars.forEach((char) => {
            const isSelected = partyPlan.has(char.characterId);
            html += createCardHTML(char, isSelected, false, 'search');
        });

        searchResults.innerHTML = html;

        Array.from(searchResults.children).forEach((child, i) => {
            child.style.animationDelay = `${(i % 15) * 0.03}s`;
        });
    }

    // ... [UPDATE ALL VIEWS] ...
    function updateAllViews() {
        renderRaids();
        applyRaidSearchHighlight();
        renderClubSummary();

        const assigned = getAssignedChars();

        // Update pool count
        const poolCountEl = document.getElementById('pool-count');
        const poolTotal = partyPlan.size;
        const poolAvailable = poolTotal - assigned.size;
        if (poolCountEl) poolCountEl.textContent = `${poolAvailable}/${poolTotal}`;

        // Maintain current search card highlight state
        if (currentSearchResults && currentSearchResults.length > 0) {
            document.querySelectorAll('#search-results .char-card, #search-results .result-card').forEach(card => {
                const charId = card.getAttribute('data-id');
                if (!charId) return;
                if (partyPlan.has(charId)) {
                    card.classList.add('selected');
                } else {
                    card.classList.remove('selected');
                }
            });
        }

        if (partyPlan.size === 0) {
            partyList.innerHTML = '<p class="no-results" id="empty-party-msg">Click characters from search results to add them.</p>';
            return;
        }

        // Build pool array (unassigned characters only)
        let basePoolChars = [];
        partyPlan.forEach((charData, charId) => {
            if (!assigned.has(charId)) {
                basePoolChars.push({ charId, charData });
            }
        });

        // Apply pool search first so counts reflect the search
        if (poolSearchQuery) {
            const q = poolSearchQuery.toLowerCase();
            basePoolChars = basePoolChars.filter(c => {
                const name = (c.charData.characterName || '').toLowerCase();
                const adv = (c.charData.adventureName || '').toLowerCase();
                return name.includes(q) || adv.includes(q);
            });
        }

        // Count for each filter
        const allCount = basePoolChars.length;
        const dpsCount = basePoolChars.filter(c => c.charData.total_buff_score == null).length;
        const buffCount = basePoolChars.filter(c => c.charData.total_buff_score != null).length;

        // Update filter button texts
        const btnAll = document.querySelector('.pool-filter-btn[data-filter="all"]');
        const btnDps = document.querySelector('.pool-filter-btn[data-filter="dps"]');
        const btnBuff = document.querySelector('.pool-filter-btn[data-filter="buff"]');
        if (btnAll) btnAll.textContent = `All`;
        if (btnDps) btnDps.textContent = `DPS (${dpsCount})`;
        if (btnBuff) btnBuff.textContent = `Buff (${buffCount})`;

        // Apply pool filter
        let poolChars = basePoolChars;
        if (poolFilter === 'dps') {
            poolChars = poolChars.filter(c => c.charData.total_buff_score == null);
        } else if (poolFilter === 'buff') {
            poolChars = poolChars.filter(c => c.charData.total_buff_score != null);
        }

        // Apply pool sort (highest to lowest)
        if (poolSort === 'dps') {
            poolChars.sort((a, b) => {
                const aScore = (a.charData.dps && a.charData.dps.normal) ? a.charData.dps.normal : 0;
                const bScore = (b.charData.dps && b.charData.dps.normal) ? b.charData.dps.normal : 0;
                return bScore - aScore;
            });
        } else if (poolSort === 'buff') {
            poolChars.sort((a, b) => {
                const aScore = a.charData.total_buff_score != null ? a.charData.total_buff_score : 0;
                const bScore = b.charData.total_buff_score != null ? b.charData.total_buff_score : 0;
                return bScore - aScore;
            });
        }

        let poolHtml = '';
        poolChars.forEach(({ charData }) => {
            poolHtml += createCardHTML(charData, false, true, 'pool');
        });

        if (poolChars.length === 0) {
            if (poolSearchQuery || poolFilter !== 'all') {
                poolHtml = '<p class="no-results">No characters match the current filter.</p>';
            } else {
                poolHtml = '<p class="no-results">All characters assigned to raids.</p>';
            }
        }
        partyList.innerHTML = poolHtml;
    }

    // --- Raid search highlight ---
    function applyRaidSearchHighlight() {
        const q = poolSearchQuery ? poolSearchQuery.toLowerCase() : '';
        document.querySelectorAll('#raids-container .party-slot .result-card').forEach(card => {
            if (!q) {
                card.classList.remove('raid-search-match');
                return;
            }
            const charId = card.getAttribute('data-id');
            const char = charId ? partyPlan.get(charId) : null;
            if (!char) { card.classList.remove('raid-search-match'); return; }
            const nameMatch = (char.characterName || '').toLowerCase().includes(q);
            const advMatch = (char.adventureName || '').toLowerCase().includes(q);
            if (nameMatch || advMatch) {
                card.classList.add('raid-search-match');
            } else {
                card.classList.remove('raid-search-match');
            }
        });
    }

    // --- Search Results Events ---
    searchResults.addEventListener('click', (e) => {
        const card = e.target.closest('.result-card');
        if (!card) return;

        const charId = card.dataset.id;
        const charData = currentSearchResults.find(c => c.characterId === charId);
        if (!charData) return;

        if (e.target.closest('.remove-btn')) {
            hiddenCharacters.add(charId);
            partyPlan.delete(charId);
            removeCharFromRaid(charId);
            renderResultCards(currentSearchResults);
            updateAllViews();
            savePartyPlan();
            saveContents();
            logAction('hide_char', `hid ${charData.characterName} from search results.`);
            return;
        }

        if (partyPlan.has(charId)) {
            partyPlan.delete(charId);
            removeCharFromRaid(charId);
            card.classList.remove('selected');
            logAction('remove_pool', `removed ${charData.characterName} from Waiting Room.`);
        } else {
            partyPlan.set(charId, charData);
            card.classList.add('selected');
            logAction('add_pool', `added ${charData.characterName} to Waiting Room.`);
        }
        updateAllViews();
        savePartyPlan();
        saveContents();
    });

    // --- Party List Events ---
    partyList.addEventListener('click', (e) => {
        const card = e.target.closest('.result-card');
        if (!card) return;

        const charId = card.dataset.id;

        // Remove button or click drops it from the roster
        if (e.target.closest('.remove-btn')) {
            const charData = partyPlan.get(charId);
            partyPlan.delete(charId);
            removeCharFromRaid(charId);
            updateAllViews();
            savePartyPlan();
            saveContents();
            if (charData) logAction('remove_pool', `removed ${charData.characterName} from Waiting Room.`);

            const searchCard = searchResults.querySelector(`.result-card[data-id="${charId}"]`);
            if (searchCard) {
                searchCard.classList.remove('selected');
            }
        }
    });

    // --- Pool Controls ---
    if (poolSearchInput) {
        poolSearchInput.addEventListener('input', (e) => {
            poolSearchQuery = e.target.value.trim();
            updateAllViews();
        });
    }

    document.querySelectorAll('.pool-filter-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.pool-filter-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            poolFilter = btn.dataset.filter;
            updateAllViews();
        });
    });

    document.querySelectorAll('.pool-sort-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const sort = btn.dataset.sort;
            if (poolSort === sort) {
                // Toggle off
                poolSort = null;
                btn.classList.remove('active');
            } else {
                // Activate this sort, deactivate others
                document.querySelectorAll('.pool-sort-btn').forEach(b => b.classList.remove('active'));
                poolSort = sort;
                btn.classList.add('active');
            }
            updateAllViews();
        });
    });

    // --- Drag Preview: Highlight valid/invalid raid blocks ---
    function applyRaidDragPreview(charId, fromSlotId) {
        const current = getActiveContent();
        if (!current) return;

        const charData = partyPlan.get(charId);
        if (!charData) return;

        const advName = charData.adventureName;
        const isFromPool = fromSlotId === 'pool';

        // Check global club limit (only relevant when dragging from pool)
        let globalLimitReached = false;
        if (isFromPool && advName) {
            globalLimitReached = isGlobalAdventureNameLimitReached(advName, charId);
        }

        // Get the source raid ID (if dragging between raids)
        const sourceMatch = fromSlotId === 'pool' ? null : fromSlotId.match(/raid-(\d+)/);
        const sourceRaidId = sourceMatch ? parseInt(sourceMatch[1]) : null;

        document.querySelectorAll('.raid-block').forEach(block => {
            const raidId = parseInt(block.dataset.raidId);
            const raid = current.raids.find(r => r.id === raidId);
            if (!raid) return;

            // Same raid as source — always valid (internal moves/swaps)
            if (sourceRaidId === raidId) {
                block.classList.add('raid-drop-valid');
                return;
            }

            // Global limit reached from pool — all other raids invalid
            if (globalLimitReached) {
                block.classList.add('raid-drop-invalid');
                return;
            }

            // Check if the explorer club already exists in this raid
            if (advName) {
                let clubConflict = false;
                const conflictPartyIndices = [];
                const conflictSlotIds = [];
                raid.parties.forEach((party, pIdx) => {
                    party.slots.forEach((slotCharId, sIdx) => {
                        if (slotCharId && slotCharId !== charId) {
                            const occupant = partyPlan.get(slotCharId);
                            if (occupant && occupant.adventureName === advName) {
                                clubConflict = true;
                                if (!conflictPartyIndices.includes(pIdx)) conflictPartyIndices.push(pIdx);
                                conflictSlotIds.push(`raid-${raid.id}-party-${pIdx}-slot-${sIdx}`);
                            }
                        }
                    });
                });

                if (clubConflict) {
                    block.classList.add('raid-drop-invalid');
                    // Highlight the conflicting party, dim the others
                    const partyBlocks = block.querySelectorAll('.party-block');
                    partyBlocks.forEach((pb, idx) => {
                        if (conflictPartyIndices.includes(idx)) {
                            pb.classList.add('party-conflict-highlight');
                        } else {
                            pb.classList.add('party-dimmed');
                        }
                    });
                    // Highlight the specific conflicting character slots
                    conflictSlotIds.forEach(slotId => {
                        const slotEl = block.querySelector(`[data-slot-id="${slotId}"]`);
                        if (slotEl) slotEl.classList.add('slot-conflict-highlight');
                    });
                    return;
                }
            }

            block.classList.add('raid-drop-valid');
        });
    }

    function clearRaidDragPreview() {
        document.querySelectorAll('.raid-drop-valid, .raid-drop-invalid').forEach(el => {
            el.classList.remove('raid-drop-valid', 'raid-drop-invalid');
        });
        document.querySelectorAll('.party-conflict-highlight, .party-dimmed, .slot-conflict-highlight').forEach(el => {
            el.classList.remove('party-conflict-highlight', 'party-dimmed', 'slot-conflict-highlight');
        });
    }

    // --- Drag and Drop Events ---
    document.addEventListener('dragstart', (e) => {
        const card = e.target.closest('.result-card');
        if (!card || !card.hasAttribute('draggable')) return;

        draggedCardId = card.dataset.id;
        const slotEl = card.closest('.party-slot');
        sourceSlotId = slotEl ? slotEl.dataset.slotId : 'pool';

        e.dataTransfer.effectAllowed = 'move';
        setTimeout(() => {
            card.classList.add('dragging');
            applyRaidDragPreview(draggedCardId, sourceSlotId);
        }, 0);
    });

    document.addEventListener('dragend', (e) => {
        const card = e.target.closest('.result-card');
        if (card) card.classList.remove('dragging');

        document.querySelectorAll('.drag-over, .invalid').forEach(el => {
            el.classList.remove('drag-over', 'invalid');
        });
        clearRaidDragPreview();
        draggedCardId = null;
        sourceSlotId = null;
    });

    function validateDrag(dragAId, sourceSlotId, targetSlotId) {
        if (sourceSlotId === targetSlotId) return null;

        const current = getActiveContent();
        if (!current) return 'No active content.';

        const charA = partyPlan.get(dragAId);
        if (!charA) return 'Character not found.';

        const charBId = targetSlotId === 'pool' ? null : getCharInSlot(targetSlotId);
        const charB = charBId ? partyPlan.get(charBId) : null;

        const tMatch = targetSlotId === 'pool' ? null : targetSlotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        const sMatch = sourceSlotId === 'pool' ? null : sourceSlotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);

        const tRId = tMatch ? parseInt(tMatch[1]) : null;
        const tpIdx = tMatch ? parseInt(tMatch[2]) : null;
        const tsIdx = tMatch ? parseInt(tMatch[3]) : null;

        const sRId = sMatch ? parseInt(sMatch[1]) : null;
        const spIdx = sMatch ? parseInt(sMatch[2]) : null;
        const ssIdx = sMatch ? parseInt(sMatch[3]) : null;

        // 1. Global limit (only applies if charA is entering from pool)
        if (sourceSlotId === 'pool' && charA.adventureName && isGlobalAdventureNameLimitReached(charA.adventureName, dragAId)) {
            return 'Global Explorer Club limit reached across all raids combined.';
        }

        // 2. Target Raid limit (charA entering target raid)
        if (tRId) {
            const conflict = checkAdventureNameConflict(tRId, dragAId, targetSlotId);
            if (conflict === 'party') return 'Explorer Club is already in this party.';
            if (conflict === 'raid') return 'Explorer Club is already in this raid.';
        }

        // 3. Source Raid limit (charB entering source raid via swap)
        if (charB && sRId) {
            const conflict = checkAdventureNameConflict(sRId, charBId, sourceSlotId);
            if (conflict === 'party') return 'Swapped character\'s Explorer Club is already in the source party.';
            if (conflict === 'raid') return 'Swapped character\'s Explorer Club is already in the source raid.';
        }

        // 4. Target Party Types
        if (tMatch && (!sMatch || tRId !== sRId || tpIdx !== spIdx)) {
            const raid = current.raids.find(r => r.id === tRId);
            if (raid) {
                const party = raid.parties[tpIdx];
                let buffCount = 0; let dpsCount = 0;
                party.slots.forEach((cId, idx) => {
                    if (idx === tsIdx || !cId) return;
                    const c = partyPlan.get(cId);
                    if (c && c.total_buff_score != null) buffCount++;
                    else if (c && c.total_buff_score == null) dpsCount++;
                });
                if (charA.total_buff_score != null) buffCount++; else dpsCount++;
                if (buffCount > 3) return 'Max 3 buffers per party.';
                if (dpsCount > 3) return 'Max 3 DPS per party.';
            }
        }

        // 5. Source Party Types (swap)
        if (sMatch && charB && (!tMatch || tRId !== sRId || tpIdx !== spIdx)) {
            const raid = current.raids.find(r => r.id === sRId);
            if (raid) {
                const party = raid.parties[spIdx];
                let buffCount = 0; let dpsCount = 0;
                party.slots.forEach((cId, idx) => {
                    if (idx === ssIdx || !cId) return;
                    const c = partyPlan.get(cId);
                    if (c && c.total_buff_score != null) buffCount++;
                    else if (c && c.total_buff_score == null) dpsCount++;
                });
                if (charB.total_buff_score != null) buffCount++; else dpsCount++;
                if (buffCount > 3) return 'Swapping would exceed max 3 buffers in source party.';
                if (dpsCount > 3) return 'Swapping would exceed max 3 DPS in source party.';
            }
        }

        return null;
    }

    document.addEventListener('dragover', (e) => {
        const dropzone = e.target.closest('.party-slot, .party-list');
        if (!dropzone || !draggedCardId) return;

        e.preventDefault();

        if (dropzone.classList.contains('party-slot')) {
            const slotId = dropzone.dataset.slotId;
            const errorMsg = validateDrag(draggedCardId, sourceSlotId, slotId);

            if (errorMsg) {
                dropzone.classList.add('drag-over', 'invalid');
                dropzone.title = errorMsg;
                return;
            }

            dropzone.classList.add('drag-over');
            dropzone.classList.remove('invalid');
            dropzone.title = '';
        } else {
            dropzone.classList.add('drag-over');
        }
    });

    document.addEventListener('dragleave', (e) => {
        const dropzone = e.target.closest('.party-slot, .party-list');
        if (dropzone) {
            dropzone.classList.remove('drag-over', 'invalid');
        }
    });

    document.addEventListener('drop', (e) => {
        const dropzone = e.target.closest('.party-slot, .party-list');
        if (!dropzone || !draggedCardId) return;
        e.preventDefault();
        dropzone.classList.remove('drag-over', 'invalid');

        if (dropzone.classList.contains('party-slot')) {
            const slotId = dropzone.dataset.slotId;
            const errorMsg = validateDrag(draggedCardId, sourceSlotId, slotId);

            if (errorMsg) {
                showToast(errorMsg, 'error');
                return;
            }
        }

        const targetSlotId = dropzone.id === 'party-list' ? 'pool' : dropzone.dataset.slotId;
        if (sourceSlotId === targetSlotId) return;

        const charA = draggedCardId;
        const charB = targetSlotId === 'pool' ? null : getCharInSlot(targetSlotId);

        setCharInSlot(targetSlotId, charA);
        setCharInSlot(sourceSlotId, charB);

        updateAllViews();
        saveContents();

        const charA_Data = partyPlan.get(charA);
        if (charA_Data) {
            let targetName = targetSlotId === 'pool' ? 'Waiting Room' : 'a Raid Slot';
            const tMatch = targetSlotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
            if (tMatch) {
                const trId = parseInt(tMatch[1]);
                const tpIdx = parseInt(tMatch[2]);
                const current = getActiveContent();
                const tRaid = current ? current.raids.find(r => r.id === trId) : null;
                const raidName = tRaid ? (tRaid.name || `#${trId}`) : `#${trId}`;
                const partyColors = ['Red', 'Yellow', 'Green'];
                const partyName = partyColors[tpIdx] || `Party ${tpIdx + 1}`;
                targetName = `${raidName} ${partyName}`;
            }

            if (charB) {
                const charB_Data = partyPlan.get(charB);
                logAction('swap_char', `swapped ${charA_Data.characterName} and ${charB_Data ? charB_Data.characterName : 'someone'}.`);
            } else {
                logAction('move_char', `moved ${charA_Data.characterName} to ${targetName}.`);
            }
        }
    });

    searchForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const clubName = document.getElementById('explorer-club-name').value.trim();
        const minDpsRaw = document.getElementById('min-dps').value;
        const minBuffRaw = document.getElementById('min-buff').value;
        const minDps = minDpsRaw ? Number(minDpsRaw) * 1_000_000 : 0;
        const minBuff = minBuffRaw ? Number(minBuffRaw) * 1_000_000 : 0;

        if (!clubName) {
            searchStatus.textContent = 'Explorer Club Name is required.';
            searchStatus.className = 'search-status error';
            return;
        }

        // Show loading state
        searchBtn.disabled = true;
        searchBtnText.textContent = 'Searching…';
        searchLoader.style.display = 'inline';
        searchStatus.textContent = '';
        searchStatus.className = 'search-status';
        searchResults.innerHTML = '';

        try {
            const response = await fetch('https://api.dfogang.com/search_explorer', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    name: clubName,
                    server: 'explorer',
                    average_set_dmg: false,
                    exact_match: true
                })
            });

            if (!response.ok) {
                throw new Error(`API error: ${response.status} ${response.statusText}`);
            }

            const data = await response.json();
            const allResults = data.results || [];

            // Always treat Priest(M) Crusaders with no buff score as buffers (buff power = 0)
            allResults.forEach(char => {
                if (char.jobGrowName === 'Neo: Crusader' && char.jobName === 'Priest (M)') {
                    if (char.total_buff_score == null) {
                        char.total_buff_score = 0;
                    }
                }
            });

            // Filter: buffers must meet minBuff, DPS must meet minDps
            const clubNameLower = clubName.toLowerCase();
            const filtered = allResults.filter(char => {
                if (!char.adventureName || char.adventureName.toLowerCase() !== clubNameLower) {
                    return false;
                }
                const isBuffer = char.total_buff_score != null;
                if (isBuffer) {
                    return char.total_buff_score >= minBuff;
                } else {
                    return char.dps && char.dps.normal != null && char.dps.normal >= minDps;
                }
            });

            // Sort descending by fame
            filtered.sort((a, b) => (b.fame || 0) - (a.fame || 0));

            renderResultCards(filtered);

            searchStatus.textContent = `Click on the characters you want to add or remove from the Waiting Room`;
            searchStatus.className = 'search-status success';

        } catch (err) {
            console.error('Search failed:', err);
            searchStatus.textContent = `❌ Search failed: ${err.message}`;
            searchStatus.className = 'search-status error';
        } finally {
            searchBtn.disabled = false;
            searchBtnText.textContent = 'Search';
            searchLoader.style.display = 'none';
        }
    });

    // Initialize application state from Firestore on load (already handled by initApp/onSnapshot)

    // --- Refresh Scores ---
    refreshScoresBtn.addEventListener('click', async () => {
        if (partyPlan.size === 0) {
            showToast('No characters in the roster to refresh.', 'warning');
            return;
        }

        // Collect all unique adventureNames
        const clubNames = new Set();
        partyPlan.forEach(charData => {
            if (charData.adventureName) clubNames.add(charData.adventureName);
        });

        if (clubNames.size === 0) {
            showToast('No Explorer Club names found to refresh.', 'warning');
            return;
        }

        refreshScoresBtn.disabled = true;
        refreshScoresBtn.textContent = '⏳ 0/' + clubNames.size;

        let updated = 0;
        let errors = 0;
        let idx = 0;

        for (const clubName of clubNames) {
            idx++;
            refreshScoresBtn.textContent = `⏳ ${idx}/${clubNames.size}`;
            try {
                const response = await fetch('https://api.dfogang.com/search_explorer', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        name: clubName,
                        server: 'explorer',
                        average_set_dmg: false,
                        exact_match: true
                    })
                });

                if (!response.ok) { errors++; continue; }

                const data = await response.json();
                const results = data.results || [];

                // Build a lookup by characterId from API results
                const apiLookup = new Map();
                results.forEach(c => {
                    if (c.adventureName && c.adventureName.toLowerCase() === clubName.toLowerCase()) {
                        // Always treat Priest(M) Crusaders as buffers
                        if (c.jobGrowName === 'Neo: Crusader' && c.jobName === 'Priest (M)') {
                            if (c.total_buff_score == null) {
                                c.total_buff_score = 0;
                            }
                        }
                        apiLookup.set(c.characterId, c);
                    }
                });

                // Update matching characters — only update scores UPWARD
                partyPlan.forEach((charData, charId) => {
                    if (apiLookup.has(charId)) {
                        const fresh = apiLookup.get(charId);

                        // DPS: only update if new value is higher
                        const oldDps = (charData.dps && charData.dps.normal) ? charData.dps.normal : 0;
                        const newDps = (fresh.dps && fresh.dps.normal) ? fresh.dps.normal : 0;
                        if (newDps > oldDps) {
                            charData.dps = fresh.dps;
                        }

                        // Buff: only update if new value is higher, never clear existing
                        const oldBuff = charData.total_buff_score != null ? charData.total_buff_score : 0;
                        const newBuff = fresh.total_buff_score != null ? fresh.total_buff_score : 0;
                        if (newBuff > oldBuff) {
                            charData.total_buff_score = fresh.total_buff_score;
                        }
                        // If char had a buff score and API returns null (switched to DPS), keep existing buff
                        // (charData.total_buff_score stays unchanged)

                        if (fresh.fame != null && fresh.fame > (charData.fame || 0)) charData.fame = fresh.fame;
                        if (fresh.jobGrowName) charData.jobGrowName = fresh.jobGrowName;
                        if (fresh.jobName) charData.jobName = fresh.jobName;
                        partyPlan.set(charId, charData);
                        updated++;
                    }
                });

            } catch (err) {
                console.error(`Refresh failed for club "${clubName}":`, err);
                errors++;
            }
        }

        refreshScoresBtn.disabled = false;
        refreshScoresBtn.textContent = 'Refresh Current Scores';

        // Post-refresh validation: Remove characters if their type changed causing >3 of one role
        let ejectedCount = 0;
        const current = getActiveContent();
        if (current) {
            current.raids.forEach(r => {
                r.parties.forEach(p => {
                    let dpsCount = 0, buffCount = 0;
                    p.slots.forEach((charId, idx) => {
                        if (!charId) return;
                        const char = partyPlan.get(charId);
                        if (!char) return;

                        const isBuffer = char.total_buff_score != null;
                        if (isBuffer) {
                            if (buffCount >= 3) {
                                p.slots[idx] = null;
                                ejectedCount++;
                            } else {
                                buffCount++;
                            }
                        } else {
                            if (dpsCount >= 3) {
                                p.slots[idx] = null;
                                ejectedCount++;
                            } else {
                                dpsCount++;
                            }
                        }
                    });
                });
            });
        }

        updateAllViews();
        if (currentSearchResults.length > 0) renderResultCards(currentSearchResults);
        savePartyPlan();
        if (ejectedCount > 0) saveContents();

        logAction('refresh_scores', `refreshed scores for ${clubNames.size} Explorer Club(s). ${updated} character(s) updated.`);

        let msg = `Refresh complete! ${updated} character(s) updated.`;
        if (errors > 0) msg += `\n${errors} club(s) had errors.`;
        if (ejectedCount > 0) msg += `\n⚠️ ${ejectedCount} character(s) were removed from parties due to the max-3-per-type limit.`;
        showToast(msg, errors > 0 ? 'warning' : 'success');
    });
});
