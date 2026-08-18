// ==UserScript==
// @name         Excel Macro Automation Companion V8.5
// @namespace    http://tampermonkey.net
// @version      8.5
// @description  Replaces textarea with a persistent 40-row x 20-stop interactive live UI grid.
// @author       Internal Tooling
// @match        https://*.sharepoint.com/*
// @match        https://*.live.com/*
// @match        https://*.office.com/*
// @allFrames    true
// @noframes     false
// @grant        none
// @run-at       document-idle
// ==/UserScript==

// NOTE: this is a sanitized template. The original @match targeted one
// specific SharePoint workbook URL (embedding an internal site name and
// a file-access token); that's been broadened to match any SharePoint/
// Office/Live document instead. Narrow @match back down to your actual
// workbook URL if you want the script scoped to a single file.

(function() {
    'use strict';

    const currentUrl = window.location.href;
    if (!currentUrl.includes('sharepoint.com') && !currentUrl.includes('://live.com') && !currentUrl.includes('office.com')) {
        return;
    }

    // --- State & Core Structures ---
    let tabsState = {
        activeTabId: "tab_default",
        tabs: {
            "tab_default": { id: "tab_default", name: "Main Matrix", gridData: {} }
        }
    };

    // Hydrate state from storage
    try {
        const savedState = localStorage.getItem('emd_multitab_state_v8');
        if (savedState) {
            tabsState = JSON.parse(savedState);
        }
    } catch(e) {
        console.error("Failed to rehydrate multi-tab state database:", e);
    }

    // Ensure layout structure compatibility
    Object.keys(tabsState.tabs).forEach(tabId => {
        if (!tabsState.tabs[tabId].gridData) {
            tabsState.tabs[tabId].gridData = {};
        }
    });

    function saveStateToStorage() {
        localStorage.setItem('emd_multitab_state_v8', JSON.stringify(tabsState));
    }

    function getActiveTab() {
        if (!tabsState.tabs[tabsState.activeTabId]) {
            tabsState.activeTabId = Object.keys(tabsState.tabs)[0] || "tab_default";
        }
        return tabsState.tabs[tabsState.activeTabId];
    }

    function sanitizeKey(str) {
        return str ? str.toString().replace(/[^A-Z0-9]/gi, '').toUpperCase() : '';
    }

    // --- Core Injector DOM Layout Setup ---
    function injectControlPanel() {
        if (document.getElementById('excel-injector-companion-panel')) {
            updateTabHeadersUI();
            return;
        }

        const panel = document.createElement('div');
        panel.id = 'excel-injector-companion-panel';
        panel.style = `
            position: fixed; bottom: 30px; left: 30px; width: 480px;
            background: #107c41; color: white; border-radius: 6px;
            z-index: 2147483647; padding: 12px; font-family: Segoe UI, Arial, sans-serif;
            box-shadow: 0px 4px 16px rgba(0,0,0,0.5); font-size: 12px;
            text-align: center; border: 1px solid #0b592e;
            display: flex; flex-direction: column; gap: 6px;
            user-select: none; max-height: 88vh;
        `;

        // Generate the 20 columns header HTML using standard loop to prevent template escape bugs
        let stopsHeaderHtml = "";
        for (let i = 1; i <= 20; i++) {
            stopsHeaderHtml += `<th style="padding: 4px; border-right: 1px solid #ddd; text-align: center; min-width: 30px;">${i}</th>`;
        }

        panel.innerHTML = `
            <div id="emd-panel-header" style="font-weight: bold; padding-bottom: 4px; border-bottom: 1px solid #0b592e; cursor: move; flex-shrink: 0; display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <span id="emd-title-text">Excel Matrix Dropper v8.6</span>
                <div style="display: flex; align-items: center; gap: 6px;">
                    <span id="emd-hint-text" style="font-size: 9px; opacity: 0.8; font-weight: normal;">Double-click tabs to rename</span>
                    <button id="emd-min-btn" style="background: none; border: none; font-weight: bold; font-size: 16px; cursor: pointer; color: white; padding: 0 4px; line-height: 10px; outline: none;">−</button>
                </div>
            </div>

            <div id="emd-panel-body" style="display: flex; flex-direction: column; gap: 6px; overflow: hidden; height: auto;">
                <div style="display: flex; align-items: center; background: #0b592e; padding: 3px 3px 0 3px; border-radius: 4px 4px 0 0; gap: 2px; overflow-x: auto; flex-shrink: 0; margin-bottom: -2px;">
                    <div id="emd-tabs-wrapper" style="display: flex; gap: 2px; align-items: center;"></div>
                    <button id="emd-add-tab-btn" style="background: rgba(255,255,255,0.2); color: white; border: none; padding: 2px 7px; border-radius: 3px; cursor: pointer; font-weight: bold; font-size: 11px; outline: none; margin-bottom: 3px;">+</button>
                </div>

                <div style="background: #ffebee; color: #c62828; font-weight: bold; padding: 5px; border-radius: 4px; border: 1px solid #ffcdd2; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; flex-shrink: 0;">
                    ALL DAM MUST BE MANUALLY CHECKED
                </div>

                <div style="text-align: left; font-size: 10px; color: #c8e6c9; font-weight: bold; flex-shrink: 0;">1. Data Matrix Inputs Grid (<span id="active-grid-title-label"></span>):</div>
                <div style="width: 100%; overflow: auto; background: #fff; border-radius: 4px; max-height: 280px; border: 1px solid #0b592e; flex-grow: 1;">
                    <table id="ui-override-grid" style="width: 100%; border-collapse: collapse; font-size: 10px; color: #333; text-align: left;">
                        <thead style="background: #f3f3f3; position: sticky; top: 0; z-index: 10;">
                            <tr style="border-bottom: 1px solid #ddd;">
                                <th style="padding: 4px; border-right: 1px solid #ddd; min-width: 80px; background: #f3f3f3; position: sticky; left: 0; z-index: 11;">Route</th>
                                <th style="padding: 4px; border-right: 1px solid #ddd; text-align: center; min-width: 50px; background: #f3f3f3;">Truck</th>
                                <th style="padding: 4px; border-right: 1px solid #ddd; text-align: center; min-width: 50px; background: #f3f3f3;">Team</th>
                                ` + stopsHeaderHtml + `
                            </tr>
                        </thead>
                        <tbody id="ui-override-grid-body"></tbody>
                    </table>
                </div>

                <div style="display: flex; gap: 4px; width: 100%; flex-shrink: 0;">
                    <button id="save-pic-memory-btn" style="flex: 1; background: #1565c0; color: white; border: none; padding: 6px; border-radius: 3px; cursor: pointer; font-size: 10px; font-weight: bold;">Save Configs</button>
                    <button id="clear-pic-memory-btn" style="flex: 1; background: #c62828; color: white; border: none; padding: 6px; border-radius: 3px; cursor: pointer; font-size: 10px; font-weight: bold;">Reset Active Tab</button>
                </div>

                <div style="text-align: left; font-size: 10px; color: #c8e6c9; font-weight: bold; flex-shrink: 0;">2. Upload EDP CSV:</div>
                <input type="file" id="csv-file-picker" accept=".csv" style="width: 100%; font-size: 10px; background: #0b592e; padding: 4px; border-radius: 3px; border: none; color: white; cursor: pointer; flex-shrink: 0;">

                <button id="inject-excel-data-btn" style="width: 100%; background: white; color: #107c41; border: none; padding: 8px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 11px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); flex-shrink: 0;">
                    Inject CSV Stream Matrix
                </button>
                <div id="injector-status-msg" style="font-size: 10px; color: #e1f5fe; font-style: italic; flex-shrink: 0;">Ready...</div>
            </div>
        `;

        document.body.appendChild(panel);

        const minBtn = document.getElementById('emd-min-btn');
        const bodyContent = document.getElementById('emd-panel-body');
        const headerEl = document.getElementById('emd-panel-header');
        const titleText = document.getElementById('emd-title-text');
        const hintText = document.getElementById('emd-hint-text');

        function applyMinimizeLayout(isMin) {
            if (isMin) {
                bodyContent.style.display = 'none';
                titleText.style.display = 'none';
                hintText.style.display = 'none';
                headerEl.style.margin = '0px';
                headerEl.style.padding = '0px';
                headerEl.style.border = 'none';

                panel.style.width = '24px';
                panel.style.height = '24px';
                panel.style.minWidth = '24px';
                panel.style.minHeight = '24px';
                panel.style.padding = '0px';
                panel.style.background = '#107c41';
                panel.style.border = '2px solid #ffffff';
                panel.style.borderRadius = '50%';
                panel.style.boxShadow = '0px 2px 8px rgba(0,0,0,0.3)';
                panel.style.opacity = '0.25';
                panel.style.cursor = 'pointer';

                minBtn.style.position = 'absolute';
                minBtn.style.top = '0px';
                minBtn.style.left = '0px';
                minBtn.style.width = '100%';
                minBtn.style.height = '100%';
                minBtn.style.fontSize = '12px';
                minBtn.style.color = 'white';
                minBtn.style.display = 'flex';
                minBtn.style.alignItems = 'center';
                minBtn.style.justifyContent = 'center';
                minBtn.textContent = '⛶';
            } else {
                bodyContent.style.display = 'flex';
                titleText.style.display = 'inline';
                hintText.style.display = 'inline';
                headerEl.style.paddingBottom = '4px';
                headerEl.style.borderBottom = '1px solid #0b592e';

                panel.style.width = '480px';
                panel.style.height = 'auto';
                panel.style.padding = '12px';
                panel.style.background = '#107c41';
                panel.style.border = '1px solid #0b592e';
                panel.style.borderRadius = '6px';
                panel.style.boxShadow = '0px 4px 16px rgba(0,0,0,0.5)';
                panel.style.opacity = '1';
                panel.style.cursor = 'default';

                minBtn.style.position = 'static';
                minBtn.style.width = 'auto';
                minBtn.style.height = 'auto';
                minBtn.style.fontSize = '16px';
                minBtn.style.color = 'white';
                minBtn.textContent = '−';
            }
        }

        panel.addEventListener('mouseenter', () => {
            if (localStorage.getItem('emd_panel_minimized') === 'true') panel.style.opacity = '1';
        });
        panel.addEventListener('mouseleave', () => {
            if (localStorage.getItem('emd_panel_minimized') === 'true') panel.style.opacity = '0.25';
        });

        const isMinimized = localStorage.getItem('emd_panel_minimized') === 'true';
        if (isMinimized) applyMinimizeLayout(true);

        minBtn.addEventListener('click', (e) => {
            const currentMin = localStorage.getItem('emd_panel_minimized') === 'true';
            localStorage.setItem('emd_panel_minimized', currentMin ? 'false' : 'true');
            applyMinimizeLayout(!currentMin);
            e.stopPropagation();
        });

        // --- Panel Dragging Mechanics ---
        const header = document.getElementById('emd-panel-header');
        let isDragging = false;
        let offsetX, offsetY;

        header.addEventListener('mousedown', (e) => {
            if (localStorage.getItem('emd_panel_minimized') === 'true' || e.target.closest('button') || e.target.closest('input')) return;
            isDragging = true;
            offsetX = e.clientX - panel.getBoundingClientRect().left;
            offsetY = e.clientY - panel.getBoundingClientRect().top;
            panel.style.bottom = 'auto';
            panel.style.right = 'auto';
            header.style.cursor = 'grabbing';
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            let leftPosition = e.clientX - offsetX;
            let topPosition = e.clientY - offsetY;
            leftPosition = Math.max(0, Math.min(leftPosition, window.innerWidth - panel.offsetWidth));
            topPosition = Math.max(0, Math.min(topPosition, window.innerHeight - panel.offsetHeight));
            panel.style.left = leftPosition + 'px';
            panel.style.top = topPosition + 'px';
        });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                header.style.cursor = 'move';
            }
        });

        document.getElementById('emd-add-tab-btn').addEventListener('click', createNewTabScope);
        document.getElementById('save-pic-memory-btn').addEventListener('click', saveActiveTabConfiguration);
        document.getElementById('clear-pic-memory-btn').addEventListener('click', clearActiveTabMemoryScope);
        document.getElementById('inject-excel-data-btn').addEventListener('click', processAndStreamCsvMatrix);

        switchWorkspaceContext();
    }

    // --- Dynamic Multi-tab Controller Engines ---
    function updateTabHeadersUI() {
        const wrapper = document.getElementById('emd-tabs-wrapper');
        if (!wrapper) return;

        wrapper.innerHTML = "";
        Object.keys(tabsState.tabs).forEach(tabId => {
            const tab = tabsState.tabs[tabId];
            const isActive = tabId === tabsState.activeTabId;

            const tabEl = document.createElement('div');
            tabEl.style = `
                background: ${isActive ? '#107c41' : 'rgba(255,255,255,0.15)'};
                color: ${isActive ? '#fff' : '#c8e6c9'};
                padding: 4px 8px; font-size: 10px; font-weight: bold;
                border-radius: 4px 4px 0 0; cursor: pointer;
                display: flex; align-items: center; gap: 4px;
                border: 1px solid ${isActive ? '#0b592e' : 'transparent'};
                border-bottom: none; max-width: 120px; overflow: hidden;
                white-space: nowrap; margin-bottom: 1px;
            `;
            tabEl.title = tab.name;

            const nameSpan = document.createElement('span');
            nameSpan.style = "overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 70px;";
            nameSpan.textContent = tab.name;
            tabEl.appendChild(nameSpan);

            const renameBtn = document.createElement('span');
            renameBtn.innerHTML = "&#9998;";
            renameBtn.style = "cursor: pointer; font-size: 10px; opacity: 0.6; padding: 0 2px; display: inline-block; vertical-align: middle;";
            renameBtn.title = "Rename Workspace";

            const triggerRename = (e) => {
                e.stopPropagation();
                const newName = prompt(`Rename processing workspace target:`, tab.name);
                if (newName && newName.trim()) {
                    tab.name = newName.trim();
                    saveStateToStorage();
                    updateTabHeadersUI();
                    if(isActive) document.getElementById('active-grid-title-label').textContent = tab.name;
                }
            };

            renameBtn.addEventListener('click', triggerRename);
            tabEl.addEventListener('dblclick', triggerRename);

            tabEl.addEventListener('click', () => {
                if (tabsState.activeTabId !== tabId) {
                    tabsState.activeTabId = tabId;
                    saveStateToStorage();
                    switchWorkspaceContext();
                }
            });

            if (Object.keys(tabsState.tabs).length > 1) {
                const closeBtn = document.createElement('span');
                closeBtn.textContent = "×";
                closeBtn.style = "color: rgba(255,255,255,0.6); font-size: 12px; font-weight: bold; cursor: pointer; padding: 0 2px; border-radius: 2px; margin-left: 2px;";
                closeBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    if (confirm(`Completely delete "${tab.name}" pipeline memory workspace?`)) {
                        delete tabsState.tabs[tabId];
                        if (tabsState.activeTabId === tabId) {
                            tabsState.activeTabId = Object.keys(tabsState.tabs)[0];
                        }
                        saveStateToStorage();
                        switchWorkspaceContext();
                    }
                });
                tabEl.appendChild(closeBtn);
            }

            wrapper.appendChild(tabEl);
        });
    }

    function createNewTabScope() {
        const id = "tab_" + Date.now();
        const count = Object.keys(tabsState.tabs).length + 1;
        tabsState.tabs[id] = {
            id: id,
            name: `Matrix Tab ${count}`,
            gridData: {}
        };
        tabsState.activeTabId = id;
        saveStateToStorage();
        switchWorkspaceContext();
    }

    function switchWorkspaceContext() {
        const activeTab = getActiveTab();
        updateTabHeadersUI();

        document.getElementById('active-grid-title-label').textContent = activeTab.name;
        document.getElementById('injector-status-msg').textContent = `Workspace switched to ${activeTab.name}`;
        document.getElementById('injector-status-msg').style.color = "#e1f5fe";

        renderPersistentGrid();
    }

    function renderPersistentGrid() {
        const activeTab = getActiveTab();
        const tbody = document.getElementById('ui-override-grid-body');
        if (!tbody) return;

        let html = "";

        // Generate persistent 40 rows
        for (let r = 0; r < 40; r++) {
            const rowData = activeTab.gridData[r] || { route: "", truck: "", team: "", stops: {} };

            html += `<tr data-row="${r}" style="border-bottom: 1px solid #eee; background: #fff;">`;

            // Route Name Column (Sticky)
            html += `<td style="padding: 2px; border-right: 1px solid #eee; font-weight: bold; font-family: monospace; background: #fafafa; font-size: 10px; position: sticky; left: 0; z-index: 5; box-shadow: 2px 0 5px -2px rgba(0,0,0,0.15);">
                        <input type="text" data-field="route" data-row="${r}" value="${rowData.route || ''}" placeholder="Route..." style="width: 100%; border: none; font-size: 10px; font-weight: bold; padding: 2px; background: transparent; outline: none; text-transform: uppercase;">
                     </td>`;

            // Truck Column
            html += `<td style="padding: 2px; border-right: 1px solid #eee; text-align: center;">
                        <input type="text" data-field="truck" data-row="${r}" value="${rowData.truck || ''}" placeholder="-" style="width: 100%; border: none; font-size: 10px; text-align: center; padding: 2px; background: transparent; outline: none; text-transform: uppercase;">
                     </td>`;

            // Team Column
            html += `<td style="padding: 2px; border-right: 1px solid #eee; text-align: center;">
                        <input type="text" data-field="team" data-row="${r}" value="${rowData.team || ''}" placeholder="-" style="width: 100%; border: none; font-size: 10px; text-align: center; padding: 2px; background: transparent; outline: none; text-transform: uppercase;">
                     </td>`;

            // Stops 1 to 20 Columns
            for (let stopNum = 1; stopNum <= 20; stopNum++) {
                const stopValue = rowData.stops && rowData.stops[stopNum] ? rowData.stops[stopNum] : "";
                html += `
                    <td style="padding: 2px; border-right: 1px solid #eee; text-align: center;">
                        <input type="text" data-field="stop" data-stop="${stopNum}" data-row="${r}" value="${stopValue}" placeholder="-"
                            style="width: 100%; border: none; text-align: center; font-size: 10px; padding: 2px; outline: none; background: transparent; text-transform: uppercase;">
                    </td>`;
            }
            html += `</tr>`;
        }

        tbody.innerHTML = html;
        setupGridListeners(tbody);
    }

    function setupGridListeners(tbody) {
        const activeTab = getActiveTab();
        const inputs = Array.from(tbody.querySelectorAll('input'));

        inputs.forEach((input) => {
            const rowIndex = input.getAttribute('data-row');
            const fieldType = input.getAttribute('data-field');

            // Synchronize UI edits instantly into state memory
            input.addEventListener('input', (e) => {
                const val = e.target.value.trim().toUpperCase();
                if (!activeTab.gridData[rowIndex]) {
                    activeTab.gridData[rowIndex] = { route: "", truck: "", team: "", stops: {} };
                }

                if (fieldType === 'route') activeTab.gridData[rowIndex].route = val;
                else if (fieldType === 'truck') activeTab.gridData[rowIndex].truck = val;
                else if (fieldType === 'team') activeTab.gridData[rowIndex].team = val;
                else if (fieldType === 'stop') {
                    const stopNum = input.getAttribute('data-stop');
                    activeTab.gridData[rowIndex].stops[stopNum] = val;
                }
                saveStateToStorage();
            });

            // Grid Paste System Configuration
            input.addEventListener('paste', (e) => {
                e.preventDefault();
                e.stopPropagation();

                const clipboardData = e.clipboardData || window.clipboardData;
                const pastedText = clipboardData.getData('text');
                if (!pastedText) return;

                const pasteRows = pastedText.split(/[\r\n]+/);
                const startRow = parseInt(rowIndex, 10);

                // Determine starting column index mapping
                let startColOffset = 0;
                if (fieldType === 'truck') startColOffset = 1;
                else if (fieldType === 'team') startColOffset = 2;
                else if (fieldType === 'stop') startColOffset = 2 + parseInt(input.getAttribute('data-stop'), 10);

                pasteRows.forEach((rowText, rOffset) => {
                    if (!rowText.trim()) return;
                    const targetRowIndex = startRow + rOffset;
                    if (targetRowIndex >= 40) return; // Cap max row bounds

                    const cellValues = rowText.split('\t');
                    if (!activeTab.gridData[targetRowIndex]) {
                        activeTab.gridData[targetRowIndex] = { route: "", truck: "", team: "", stops: {} };
                    }

                    cellValues.forEach((cellVal, cOffset) => {
                        const currentAbsoluteCol = startColOffset + cOffset;
                        const cleanVal = cellVal.trim().toUpperCase();

                        const rowEl = tbody.querySelector(`tr[data-row="${targetRowIndex}"]`);
                        if (!rowEl) return;

                        if (currentAbsoluteCol === 0) {
                            activeTab.gridData[targetRowIndex].route = cleanVal;
                            rowEl.querySelector('input[data-field="route"]').value = cleanVal;
                        } else if (currentAbsoluteCol === 1) {
                            activeTab.gridData[targetRowIndex].truck = cleanVal;
                            rowEl.querySelector('input[data-field="truck"]').value = cleanVal;
                        } else if (currentAbsoluteCol === 2) {
                            activeTab.gridData[targetRowIndex].team = cleanVal;
                            rowEl.querySelector('input[data-field="team"]').value = cleanVal;
                        } else {
                            const stopNum = currentAbsoluteCol - 2;
                            if (stopNum <= 20) {
                                activeTab.gridData[targetRowIndex].stops[stopNum] = cleanVal;
                                const stopInput = rowEl.querySelector(`input[data-field="stop"][data-stop="${stopNum}"]`);
                                if (stopInput) stopInput.value = cleanVal;
                            }
                        }
                    });
                });
                saveStateToStorage();
            });

            // Focus and Visual Navigation Crosshairs Highlight
            input.addEventListener('focus', () => {
                input.style.background = '#ffffff';
                input.style.boxShadow = 'inset 0 0 0 2px #107c41';

                const parentRow = input.closest('tr');
                if (parentRow) parentRow.style.background = '#d1e7dd';
            });

            input.addEventListener('blur', () => {
                input.style.background = 'transparent';
                input.style.boxShadow = 'none';

                const parentRow = input.closest('tr');
                if (parentRow) parentRow.style.background = '#fff';
            });

            // Arrow key grid movement rules
            input.addEventListener('keydown', (e) => {
                const currentIdx = inputs.indexOf(input);
                let targetInput = null;

                // Simple grid mapping offsets (Total inputs per row = 1 route + 1 truck + 1 team + 20 stops = 23 elements)
                if (e.key === 'ArrowRight') targetInput = inputs[currentIdx + 1];
                else if (e.key === 'ArrowLeft') targetInput = inputs[currentIdx - 1];
                else if (e.key === 'ArrowDown') targetInput = inputs[currentIdx + 23];
                else if (e.key === 'ArrowUp') targetInput = inputs[currentIdx - 23];

                if (targetInput) {
                    e.preventDefault();
                    targetInput.focus();
                    targetInput.select();
                }
            });
        });
    }

    function saveActiveTabConfiguration() {
        const activeTab = getActiveTab();
        const statusEl = document.getElementById('injector-status-msg');

        saveStateToStorage();
        statusEl.textContent = `[${activeTab.name}] Grid Saved Successfully! ✓`;
        statusEl.style.color = "#a5d6a7";
    }

    function clearActiveTabMemoryScope() {
        const activeTab = getActiveTab();
        if (confirm(`Reset and clear all 40 data rows inside "${activeTab.name}"?`)) {
            activeTab.gridData = {};
            saveStateToStorage();
            switchWorkspaceContext();

            const statusEl = document.getElementById('injector-status-msg');
            statusEl.textContent = "Active grid memory cleared.";
            statusEl.style.color = "#ffcdd2";
        }
    }

    function processAndStreamCsvMatrix() {
        const activeTab = getActiveTab();
        const fileInput = document.getElementById('csv-file-picker');
        const statusEl = document.getElementById('injector-status-msg');

        if (!fileInput.files || fileInput.files.length === 0) {
            statusEl.textContent = "Error: Select a source CSV stream file!";
            statusEl.style.color = "#ffcdd2";
            return;
        }

        // Remap grid matrix dataset for processing performance optimization
        let routeMapLookup = {};
        Object.keys(activeTab.gridData).forEach(rIdx => {
            const row = activeTab.gridData[rIdx];
            const cleanKey = sanitizeKey(row.route);
            if (cleanKey) {
                routeMapLookup[cleanKey] = row;
            }
        });

        const file = fileInput.files[0];
        const reader = new FileReader();

        reader.onload = function(e) {
            let text = e.target.result;
            if (text.startsWith('\uFEFF')) text = text.substring(1);

            statusEl.textContent = `Processing data against [${activeTab.name}] rules...`;
            statusEl.style.color = "#fff";

            const lines = text.split(/[\r\n]+/);
            let unifiedClipboardRows = [];

            lines.forEach(line => {
                if (!line.trim()) return;

                const columns = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(col => col.replace(/^"|"$/g, '').trim());
                if (columns.length <= 10 || !columns[10]) return;

                const rawRouteName = columns[10];
                const routeKey = sanitizeKey(rawRouteName);
                let reconstructedLineCells = [];

                reconstructedLineCells.push(rawRouteName);

                const matchedData = routeMapLookup[routeKey];
                if (matchedData) {
                    reconstructedLineCells.push(matchedData.truck || "");
                    reconstructedLineCells.push(matchedData.team || "");
                } else {
                    reconstructedLineCells.push("");
                    reconstructedLineCells.push("");
                }

                const stopShortcodes = columns.slice(13);
                stopShortcodes.forEach((code, index) => {
                    const stopNumber = index + 1;
                    let cleanCode = code ? code.toString().trim().toUpperCase() : "";

                    // Prioritize exact manual match overrides found within the UI dataset row matrix
                    if (matchedData && matchedData.stops && matchedData.stops[stopNumber]) {
                        cleanCode = matchedData.stops[stopNumber];
                    } else if (cleanCode === "MC" || cleanCode === "MANUAL CHECK") {
                        cleanCode = "DAM";
                    }

                    reconstructedLineCells.push(cleanCode);
                });

                unifiedClipboardRows.push(reconstructedLineCells.join('\t'));
            });

            if (unifiedClipboardRows.length === 0) {
                statusEl.textContent = "No valid data rows matched grid filters!";
                statusEl.style.color = "#ffcc00";
                return;
            }

            const streamPayload = unifiedClipboardRows.join('\n');
            statusEl.textContent = "Streaming formatted matrix payload...";

            navigator.clipboard.writeText(streamPayload).then(() => {
                statusEl.textContent = `[${activeTab.name}] Matrix Stream Injected! ✓`;
                statusEl.style.color = "#a5d6a7";
            }).catch(err => {
                statusEl.textContent = "Clipboard interaction blocked.";
                console.error(err);
            });
        };

        reader.readAsText(file);
    }

    setInterval(injectControlPanel, 2000);
})();
