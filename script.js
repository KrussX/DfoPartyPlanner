/**
 * DFO Party Planner - Logic
 * Minimal client-side JS handling slot selection.
 */
document.addEventListener('DOMContentLoaded', () => {
    // --- State Management ---
    const partyPlan = new Map(); // characterId -> charData
    const hiddenCharacters = new Set(); // set of characterIds
    let currentSearchResults = [];

    let raidCounter = 0;
    const raids = []; // Array of { id, size, parties: [] }
    let draggedCardId = null;
    let sourceSlotId = null;

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
    const raidSizeSelect = document.getElementById('raid-size-select');
    const raidsContainer = document.getElementById('raids-container');
    const globalClubLimitInput = document.getElementById('global-club-limit');
    const clubSummaryContainer = document.getElementById('club-summary-container');
    const clubSummaryList = document.getElementById('club-summary-list');

    // Controls
    const clearBtn = document.getElementById('clear-btn');
    const clearRaidsBtn = document.getElementById('clear-raids-btn');
    const exportJsonBtn = document.getElementById('export-json-btn');
    const exportExcelBtn = document.getElementById('export-excel-btn');
    const importJsonBtn = document.getElementById('import-json-btn');
    const importFileInput = document.getElementById('import-file-input');
    const themeToggleBtn = document.getElementById('theme-toggle-btn');
    const refreshScoresBtn = document.getElementById('refresh-scores-btn');

    // Confirm Modal Elements
    const confirmModal = document.getElementById('confirm-modal');
    const modalMessage = document.getElementById('modal-message');
    const modalConfirmBtn = document.getElementById('modal-confirm');
    const modalCancelBtn = document.getElementById('modal-cancel');

    // --- Toggle Search ---
    toggleSearchBtn.addEventListener('click', () => {
        if (searchContent.style.display === 'none') {
            searchContent.style.display = 'block';
            toggleSearchBtn.textContent = '▾';
        } else {
            searchContent.style.display = 'none';
            toggleSearchBtn.textContent = '▸';
        }
    });

    // --- Theme Toggle ---
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

    // --- State Persistence (LocalStorage) ---
    function saveState() {
        const state = {
            partyPlan: Array.from(partyPlan.entries()),
            hiddenCharacters: Array.from(hiddenCharacters),
            raids: raids,
            raidCounter: raidCounter,
            globalClubLimit: globalClubLimitInput.value
        };
        localStorage.setItem('dfoRaidPlannerState', JSON.stringify(state));
    }

    function loadState() {
        const saved = localStorage.getItem('dfoRaidPlannerState');
        if (saved) {
            try {
                const state = JSON.parse(saved);
                partyPlan.clear();
                if (state.partyPlan) {
                    state.partyPlan.forEach(([id, data]) => partyPlan.set(id, data));
                }
                hiddenCharacters.clear();
                if (state.hiddenCharacters) {
                    state.hiddenCharacters.forEach(id => hiddenCharacters.add(id));
                }
                raids.length = 0;
                if (state.raids) {
                    state.raids.forEach(r => {
                        // Migrate old format { dps1, dps2, dps3, buff } to { slots: [] }
                        if (r.parties && r.parties.length > 0 && !r.parties[0].slots) {
                            r.parties = r.parties.map(p => ({
                                slots: [p.dps1 || null, p.dps2 || null, p.dps3 || null, p.buff || null]
                            }));
                        }
                        raids.push(r);
                    });
                }
                raidCounter = state.raidCounter || 0;
                if (state.globalClubLimit !== undefined) {
                    globalClubLimitInput.value = state.globalClubLimit;
                }
                updateAllViews();
            } catch (e) { console.error('Failed to load state', e); }
        }
    }

    globalClubLimitInput.addEventListener('change', () => {
        saveState();
        renderClubSummary();
    });

    // --- Global Controls (Export / Import / Clear) ---
    clearBtn.addEventListener('click', () => {
        showConfirm('Are you sure you want to clear EVERYTHING? This will remove all characters and all raids.', () => {
            partyPlan.clear();
            hiddenCharacters.clear();
            raids.length = 0;
            raidCounter = 0;
            updateAllViews();
            if (currentSearchResults.length > 0) {
                renderResultCards(currentSearchResults);
            } else {
                document.getElementById('search-results').innerHTML = '';
            }
        }, 'Clear Everything', 'danger');
    });

    clearRaidsBtn.addEventListener('click', () => {
        showConfirm('Are you sure you want to clear all raids?', () => {
            raids.length = 0;
            raidCounter = 0;
            updateAllViews();
        }, 'Clear Raids', 'primary');
    });

    exportJsonBtn.addEventListener('click', async () => {
        const state = {
            partyPlan: Array.from(partyPlan.entries()),
            hiddenCharacters: Array.from(hiddenCharacters),
            raids: raids,
            raidCounter: raidCounter
        };
        const jsonStr = JSON.stringify(state, null, 2);

        // Generate formatted date YYYY-MM-DD
        const date = new Date();
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const filename = `planner_${yyyy}-${mm}-${dd}.json`;

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
        if (raids.length === 0) {
            alert('No raids to export.');
            return;
        }

        const wb = XLSX.utils.book_new();
        const ws_data = [];
        const merges = [];

        const numRaids = raids.length;
        const leftCount = Math.ceil(numRaids / 2);

        let leftRow = 0;
        let rightRow = 0;

        function setCell(r, c, val, style, type = 's', format = undefined) {
            if (!ws_data[r]) ws_data[r] = [];
            const cell = { v: val, t: type, s: style };
            if (format) cell.z = format;
            ws_data[r][c] = cell;
        }

        raids.forEach((raid, rIndex) => {
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
                    let isDpsSader = false;

                    if (charId && partyPlan.has(charId)) {
                        const char = partyPlan.get(charId);
                        ec = char.adventureName || "?";
                        nameJob = `${char.characterName}`;
                        if (char.total_buff_score != null) {
                            scoreVal = char.total_buff_score;
                        } else if (char.dps && char.dps.normal) {
                            scoreVal = char.dps.normal;
                        }
                        
                        if (char._isDpsSader) {
                            isDpsSader = true;
                        }
                        raidSum += (scoreVal || 0);
                    }

                    setCell(currentRow, startCol, ec, rowStyle);
                    setCell(currentRow, startCol + 1, nameJob, rowStyle);
                    
                    if (charId) {
                        if (isDpsSader) {
                            // Show DPS MODE
                            setCell(currentRow, startCol + 2, "DPS MODE", rowStyle);
                        } else {
                            // Format number
                            setCell(currentRow, startCol + 2, scoreVal, rowStyle, 'n', '#,##0');
                        }
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

            const emptyStyle = { font: { color: { rgb: "000000"} } };

            unassignedChars.forEach(char => {
                const ec = char.adventureName || "?";
                const name = char.characterName;
                let scoreVal = null;
                let isDpsSader = false;
                
                if (char.total_buff_score != null) {
                    scoreVal = char.total_buff_score;
                } else if (char.dps && char.dps.normal) {
                    scoreVal = char.dps.normal;
                }
                if (char._isDpsSader) {
                    isDpsSader = true;
                }

                setCell(maxRow, 0, ec, emptyStyle);
                setCell(maxRow, 1, name, emptyStyle);
                
                if (isDpsSader) {
                    setCell(maxRow, 2, "DPS MODE", emptyStyle);
                } else {
                    setCell(maxRow, 2, scoreVal, emptyStyle, 'n', '#,##0');
                }
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
            { wch: 3  }, // D: Spacer
            { wch: 18 }, // E: Explorer club
            { wch: 35 }, // F: Char
            { wch: 30 }  // G: DPS
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Raids");

        const date = new Date();
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const filename = `raid_roster_${yyyy}-${mm}-${dd}.xlsx`;

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
                if (state.partyPlan) state.partyPlan.forEach(([id, data]) => partyPlan.set(id, data));

                hiddenCharacters.clear();
                if (state.hiddenCharacters) state.hiddenCharacters.forEach(id => hiddenCharacters.add(id));

                raids.length = 0;
                if (state.raids) state.raids.forEach(r => {
                    if (r.parties && r.parties.length > 0 && !r.parties[0].slots) {
                        r.parties = r.parties.map(p => ({
                            slots: [p.dps1 || null, p.dps2 || null, p.dps3 || null, p.buff || null]
                        }));
                    }
                    raids.push(r);
                });

                raidCounter = state.raidCounter || 0;
                updateAllViews();
                alert('Backup imported successfully!');
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
        raidCounter++;
        const size = parseInt(raidSizeSelect.value);
        const partiesCount = size / 4;
        const parties = [];
        for (let i = 0; i < partiesCount; i++) {
            parties.push({ slots: [null, null, null, null] });
        }
        raids.push({ id: raidCounter, size, parties });
        updateAllViews();
    });

    // --- Auto Planner ---
    autoPlanBtn.addEventListener('click', () => {
        const poolDPS = [];
        // Buffers are deliberately ignored in Auto Planner

        partyPlan.forEach((charData, charId) => {
            const isBuffer = charData.total_buff_score != null;
            if (!isBuffer) {
                poolDPS.push({
                    id: charId,
                    data: charData,
                    adv: charData.adventureName,
                    power: charData.dps?.normal || 0
                });
            }
        });

        poolDPS.sort((a, b) => b.power - a.power);

        const size = parseInt(raidSizeSelect.value);
        const P = size / 4;
        const R_DPS = 3 * P; // 3 DPS per party

        let maxN = Math.floor(poolDPS.length / R_DPS);

        const globalLimitStr = globalClubLimitInput.value;
        const globalLimit = (globalLimitStr && parseInt(globalLimitStr) > 0) ? parseInt(globalLimitStr) : Infinity;

        let bestRaids = null;
        let globalClubUsage = new Map();

        for (let N = maxN; N >= 1; N--) {
            const testRaids = [];
            for (let i = 0; i < N; i++) {
                testRaids.push({ dps: [], advNames: new Set(), dpsSum: 0 });
            }

            const clubUsageCount = new Map();

            for (const char of poolDPS) {
                const currentUsage = clubUsageCount.get(char.adv) || 0;
                if (currentUsage >= globalLimit) continue; // Skip globally if capped

                let bestBucket = null;
                for (const bucket of testRaids) {
                    if (bucket.dps.length < R_DPS && !bucket.advNames.has(char.adv)) {
                        if (!bestBucket || bucket.dpsSum < bestBucket.dpsSum) {
                            bestBucket = bucket;
                        }
                    }
                }
                if (bestBucket) {
                    bestBucket.dps.push(char);
                    bestBucket.advNames.add(char.adv);
                    bestBucket.dpsSum += char.power;
                    clubUsageCount.set(char.adv, currentUsage + 1);
                }
            }

            let success = true;
            for (const bucket of testRaids) {
                if (bucket.dps.length < R_DPS) {
                    success = false;
                    break;
                }
            }

            if (success) {
                bestRaids = testRaids;
                globalClubUsage = clubUsageCount;
                break;
            }
        }

        if (!bestRaids) {
            bestRaids = [];
        }

        // --- Build Incomplete Extra Raid ---
        const assignedIds = new Set();
        bestRaids.forEach(r => {
            r.dps.forEach(c => assignedIds.add(c.id));
        });

        const leftoverBucket = { dps: [], advNames: new Set(), dpsSum: 0 };
        let addedLeftovers = false;

        for (const char of poolDPS) {
            if (assignedIds.has(char.id)) continue;
            const currentUsage = globalClubUsage.get(char.adv) || 0;
            if (currentUsage >= globalLimit) continue;

            if (leftoverBucket.dps.length < R_DPS && !leftoverBucket.advNames.has(char.adv)) {
                leftoverBucket.dps.push(char);
                leftoverBucket.advNames.add(char.adv);
                leftoverBucket.dpsSum += char.power;
                globalClubUsage.set(char.adv, currentUsage + 1);
                addedLeftovers = true;
            }
        }

        if (addedLeftovers) {
            bestRaids.push(leftoverBucket);
        }

        if (bestRaids.length === 0) {
            alert(`No valid characters available!`);
            return;
        }

        raids.length = 0;
        raidCounter = 0;

        for (const bucket of bestRaids) {
            raidCounter++;
            const parties = [];
            for (let pIdx = 0; pIdx < P; pIdx++) parties.push({ slots: [null, null, null, null] });

            bucket.dps.sort((a, b) => b.power - a.power);

            // Fill slots 0, 1, 2 with DPS using balance logic. (Slot 3 is left empty for buffer)
            for (const dChar of bucket.dps) {
                let lowestDps = Infinity;
                let targetParty = null;
                let targetSlotIdx = null;

                for (const p of parties) {
                    // Only find empty slots among the first 3 (indices 0, 1, 2)
                    let emptyIdx = -1;
                    for (let sIdx = 0; sIdx < 3; sIdx++) {
                        if (p.slots[sIdx] === null) {
                            emptyIdx = sIdx;
                            break;
                        }
                    }

                    if (emptyIdx !== -1) {
                        let dSum = dChar.power;
                        for (let sIdx = 0; sIdx < 3; sIdx++) {
                            const s = p.slots[sIdx];
                            if (s) {
                                const sc = partyPlan.get(s);
                                if (sc && sc.dps && sc.dps.normal) dSum += sc.dps.normal;
                            }
                        }

                        if (dSum < lowestDps) {
                            lowestDps = dSum;
                            targetParty = p;
                            targetSlotIdx = emptyIdx;
                        }
                    }
                }

                if (targetParty && targetSlotIdx !== null) {
                    targetParty.slots[targetSlotIdx] = dChar.id;
                }
            }

            raids.push({ id: raidCounter, size: parseInt(raidSizeSelect.value), parties });
        }

        updateAllViews();
        if (raidsContainer.offsetTop) {
            window.scrollTo({ top: raidsContainer.offsetTop - 50, behavior: 'smooth' });
        }
    });

    raidsContainer.addEventListener('click', (e) => {
        // Remove raid
        if (e.target.classList.contains('raid-remove-btn')) {
            const rId = parseInt(e.target.dataset.raidId);
            const raid = raids.find(r => r.id === rId);
            const raidName = raid ? (raid.name || `#${rId}`) : 'this raid';
            
            showConfirm(`Are you sure you want to remove ${raidName}?`, () => {
                const idx = raids.findIndex(r => r.id === rId);
                if (idx !== -1) {
                    raids.splice(idx, 1);
                    updateAllViews();
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
            }
        }
    });

    raidsContainer.addEventListener('change', (e) => {
        if (e.target.classList.contains('raid-title-input')) {
            const rId = parseInt(e.target.dataset.raidId);
            const raid = raids.find(r => r.id === rId);
            if (raid) {
                raid.name = e.target.value.trim();
                saveState();
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
        raids.forEach(r => r.parties.forEach(p => {
            p.slots.forEach(s => { if (s) assigned.add(s); });
        }));
        return assigned;
    }

    function getCharInSlot(slotId) {
        if (slotId === 'pool') return null;
        const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        if (!parts) return null;
        const raid = raids.find(r => r.id === parseInt(parts[1]));
        if (raid) return raid.parties[parseInt(parts[2])].slots[parseInt(parts[3])];
        return null;
    }

    function setCharInSlot(slotId, charId) {
        if (slotId === 'pool') return;
        const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
        if (!parts) return;
        const raid = raids.find(r => r.id === parseInt(parts[1]));
        if (raid) raid.parties[parseInt(parts[2])].slots[parseInt(parts[3])] = charId;
    }

    function checkAdventureNameConflict(raidId, charAId, targetSlotId) {
        const raid = raids.find(r => r.id === raidId);
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
        const limitStr = globalClubLimitInput.value;
        if (!limitStr) return false;
        const limit = parseInt(limitStr);
        if (isNaN(limit) || limit <= 0) return false;

        let currentCount = 0;
        raids.forEach(r => r.parties.forEach(p => {
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
        raids.forEach(r => r.parties.forEach(p => {
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
        const limitStr = globalClubLimitInput.value;
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
        raids.forEach(r => r.parties.forEach(p => {
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
        if (char._isDpsSader) {
            scoreDisplay = 'DPS MODE';
        }

        const removeBtnStr = currentSlotId === 'search' ? '' : '<button class="remove-btn" title="Remove" aria-label="Remove">×</button>';

        return `
            <div class="result-card ${isBuffer ? 'buffer' : 'dealer'} ${isSelected ? 'selected' : ''}" 
                 data-id="${char.characterId}"
                 ${isDraggable ? 'draggable="true"' : ''}>
                ${removeBtnStr}
                <div class="card-info">
                    <div class="card-name">${char.characterName}</div>
                    <div class="card-sub">${char.adventureName || '?'} · ${char.jobGrowName || char.jobName}</div>
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
        if (raids.length === 0) {
            raidsContainer.innerHTML = '<div class="empty-raids-msg"><p>No raids yet. Click <strong>+ Raid</strong> or <strong>⚡ Auto</strong> to get started.</p></div>';
            return;
        }
        raids.forEach((raid, rIndex) => {
            const raidBlock = document.createElement('div');
            raidBlock.className = 'raid-block';

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
                            <span class="party-dmg-label">DPS</span>
                            <span class="party-dmg-value dps-val">${totals.dps}</span>
                            <span class="party-dmg-label">Buff</span>
                            <span class="party-dmg-value buff-val">${totals.buff}</span>
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
        renderClubSummary();

        const assigned = getAssignedChars();

        // Update pool count
        const poolCountEl = document.getElementById('pool-count');
        const poolTotal = partyPlan.size;
        const poolAvailable = poolTotal - assigned.size;
        if (poolCountEl) poolCountEl.textContent = `${poolAvailable}/${poolTotal}`;

        if (partyPlan.size === 0) {
            partyList.innerHTML = '<p class="no-results" id="empty-party-msg">Click characters from search results to add them.</p>';
            saveState();
            return;
        }

        let poolHtml = '';
        let poolCount = 0;
        partyPlan.forEach((charData, charId) => {
            if (!assigned.has(charId)) {
                poolHtml += createCardHTML(charData, false, true, 'pool');
                poolCount++;
            }
        });

        if (poolCount === 0) {
            poolHtml = '<p class="no-results">All characters assigned to raids.</p>';
        }
        partyList.innerHTML = poolHtml;

        // Auto-save State on any substantial UI update
        saveState();
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
            return;
        }

        if (partyPlan.has(charId)) {
            partyPlan.delete(charId);
            removeCharFromRaid(charId);
            card.classList.remove('selected');
        } else {
            partyPlan.set(charId, charData);
            card.classList.add('selected');
        }
        updateAllViews();
    });

    // --- Party List Events ---
    partyList.addEventListener('click', (e) => {
        const card = e.target.closest('.result-card');
        if (!card) return;

        const charId = card.dataset.id;

        // Remove button or click drops it from the roster
        if (e.target.closest('.remove-btn')) {
            partyPlan.delete(charId);
            removeCharFromRaid(charId);
            updateAllViews();

            const searchCard = searchResults.querySelector(`.result-card[data-id="${charId}"]`);
            if (searchCard) {
                searchCard.classList.remove('selected');
            }
        }
    });

    // --- Drag and Drop Events ---
    document.addEventListener('dragstart', (e) => {
        const card = e.target.closest('.result-card');
        if (!card || !card.hasAttribute('draggable')) return;

        draggedCardId = card.dataset.id;
        const slotEl = card.closest('.party-slot');
        sourceSlotId = slotEl ? slotEl.dataset.slotId : 'pool';

        e.dataTransfer.effectAllowed = 'move';
        setTimeout(() => card.classList.add('dragging'), 0);
    });

    document.addEventListener('dragend', (e) => {
        const card = e.target.closest('.result-card');
        if (card) card.classList.remove('dragging');

        document.querySelectorAll('.drag-over, .invalid').forEach(el => {
            el.classList.remove('drag-over', 'invalid');
        });
        draggedCardId = null;
        sourceSlotId = null;
    });

    document.addEventListener('dragover', (e) => {
        const dropzone = e.target.closest('.party-slot, .party-list');
        if (!dropzone || !draggedCardId) return;

        e.preventDefault();

        if (dropzone.classList.contains('party-slot')) {
            const charData = partyPlan.get(draggedCardId);
            if (!charData) return;
            const isBuffer = charData.total_buff_score != null;

            // Check type count limits (max 3 of either type)
            const slotId = dropzone.dataset.slotId;
            const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
            if (parts) {
                const rId = parseInt(parts[1]);
                const pIdx = parseInt(parts[2]);
                const sIdx = parseInt(parts[3]);
                const raid = raids.find(r => r.id === rId);
                if (raid) {
                    const party = raid.parties[pIdx];
                    const existingOccupant = party.slots[sIdx];
                    // Only check limit if slot is empty or swapping with different type
                    if (!existingOccupant || sourceSlotId === 'pool') {
                        const counts = getPartyTypeCounts(party, sIdx);
                        if (isBuffer && counts.buffCount >= 3) {
                            dropzone.classList.add('drag-over', 'invalid');
                            dropzone.title = 'Max 3 buffers per party.';
                            return;
                        }
                        if (!isBuffer && counts.dpsCount >= 3) {
                            dropzone.classList.add('drag-over', 'invalid');
                            dropzone.title = 'Max 3 DPS per party.';
                            return;
                        }
                    }
                }

                if (rId && isAdventureNameDuplicate(rId, draggedCardId, slotId)) {
                    dropzone.classList.add('drag-over', 'invalid');
                    dropzone.title = 'Cannot have multiple characters from the same Explorer Club in one raid.';
                    return;
                }
            }

            if (sourceSlotId === 'pool' && charData.adventureName && isGlobalAdventureNameLimitReached(charData.adventureName, draggedCardId)) {
                dropzone.classList.add('drag-over', 'invalid');
                dropzone.title = 'Global Explorer Club limit reached across all raids combined.';
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
            const charData = partyPlan.get(draggedCardId);
            if (!charData) return;
            const isBuffer = charData.total_buff_score != null;

            const slotId = dropzone.dataset.slotId;
            const parts = slotId.match(/raid-(\d+)-party-(\d+)-slot-(\d+)/);
            if (parts) {
                const rId = parseInt(parts[1]);
                const pIdx = parseInt(parts[2]);
                const sIdx = parseInt(parts[3]);
                const raid = raids.find(r => r.id === rId);
                if (raid) {
                    const party = raid.parties[pIdx];
                    const existingOccupant = party.slots[sIdx];
                    if (!existingOccupant || sourceSlotId === 'pool') {
                        const counts = getPartyTypeCounts(party, sIdx);
                        if (isBuffer && counts.buffCount >= 3) {
                            showToast('Cannot add a 4th buffer to this party.', 'warning');
                            return;
                        }
                        if (!isBuffer && counts.dpsCount >= 3) {
                            showToast('Cannot add a 4th DPS to this party.', 'warning');
                            return;
                        }
                    }
                }

                if (rId) {
                    const conflict = checkAdventureNameConflict(rId, draggedCardId, slotId);
                    if (conflict === 'party') {
                        showToast('Explorer Club is already in this party.', 'error');
                        return;
                    } else if (conflict === 'raid') {
                        showToast('Explorer Club is already in this raid.', 'error');
                        return;
                    }
                }
            }

            if (sourceSlotId === 'pool' && charData.adventureName) {
                if (isGlobalAdventureNameLimitReached(charData.adventureName, draggedCardId)) {
                    showToast(`Global Club Limit reached for ${charData.adventureName}.`, 'error');
                    return;
                }
            }
        }

        const targetSlotId = dropzone.id === 'party-list' ? 'pool' : dropzone.dataset.slotId;
        if (sourceSlotId === targetSlotId) return;

        const charA = draggedCardId;
        const charB = targetSlotId === 'pool' ? null : getCharInSlot(targetSlotId);

        setCharInSlot(targetSlotId, charA);
        setCharInSlot(sourceSlotId, charB);

        updateAllViews();
    });

    searchForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const clubName = document.getElementById('explorer-club-name').value.trim();
        const minDpsRaw = document.getElementById('min-dps').value;
        const minBuffRaw = document.getElementById('min-buff').value;
        const minDps = minDpsRaw ? Number(minDpsRaw) * 1_000_000 : 0;
        const minBuff = minBuffRaw ? Number(minBuffRaw) * 1_000_000 : 0;
        const treatDpsSaderAsBuffer = document.getElementById('dps-sader-buffer-toggle').checked;

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

            // Apply Buffer Toggle for DPS Crusaders
            allResults.forEach(char => {
                if (treatDpsSaderAsBuffer) {
                    if (char.jobGrowName === 'Neo: Crusader' && char.jobName === 'Priest (M)') {
                        if (char.total_buff_score == null) {
                            char.total_buff_score = 0;
                            char._isDpsSader = true; // Flag for UI badge
                        }
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
                    // Bypass minBuff if it's our artificially injected DPS mode buffer
                    if (char._isDpsSader) return true;
                    return char.total_buff_score >= minBuff;
                } else {
                    return char.dps && char.dps.normal != null && char.dps.normal >= minDps;
                }
            });

            // Sort descending by fame
            filtered.sort((a, b) => (b.fame || 0) - (a.fame || 0));

            renderResultCards(filtered);

            searchStatus.textContent = `✅ ${filtered.length} of ${allResults.length} character(s) meet the requirements.`;
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

    // Initialize application state from local storage on load
    loadState();

    // --- Refresh Scores ---
    refreshScoresBtn.addEventListener('click', async () => {
        if (partyPlan.size === 0) {
            alert('No characters in the roster to refresh.');
            return;
        }

        // Collect all unique adventureNames
        const clubNames = new Set();
        partyPlan.forEach(charData => {
            if (charData.adventureName) clubNames.add(charData.adventureName);
        });

        if (clubNames.size === 0) {
            alert('No Explorer Club names found to refresh.');
            return;
        }

        refreshScoresBtn.disabled = true;
        refreshScoresBtn.textContent = '⏳ 0/' + clubNames.size;

        const treatDpsSaderAsBuffer = document.getElementById('dps-sader-buffer-toggle').checked;

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
                        // Apply Buffer Toggle
                        if (treatDpsSaderAsBuffer) {
                            if (c.jobGrowName === 'Neo: Crusader' && c.jobName === 'Priest (M)') {
                                if (c.total_buff_score == null) {
                                    c.total_buff_score = 0;
                                    c._isDpsSader = true;
                                }
                            }
                        }
                        apiLookup.set(c.characterId, c);
                    }
                });

                // Update matching characters in partyPlan
                partyPlan.forEach((charData, charId) => {
                    if (apiLookup.has(charId)) {
                        const fresh = apiLookup.get(charId);
                        // Update all score fields
                        if (fresh.dps) charData.dps = fresh.dps;
                        charData.total_buff_score = fresh.total_buff_score; // overwrites null, undefined, or 0
                        if (fresh._isDpsSader) charData._isDpsSader = true;
                        else delete charData._isDpsSader;

                        if (fresh.fame != null) charData.fame = fresh.fame;
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
        refreshScoresBtn.textContent = '🔄 Refresh';

        // Post-refresh validation: Remove characters if their type changed causing >3 of one role
        let ejectedCount = 0;
        raids.forEach(r => {
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

        updateAllViews();
        if (currentSearchResults.length > 0) renderResultCards(currentSearchResults);

        let msg = `Refresh complete! ${updated} character(s) updated.`;
        if (errors > 0) msg += `\n${errors} club(s) had errors.`;
        if (ejectedCount > 0) msg += `\n⚠️ ${ejectedCount} character(s) were removed from parties due to the max-3-per-type limit.`;
        alert(msg);
    });
});
