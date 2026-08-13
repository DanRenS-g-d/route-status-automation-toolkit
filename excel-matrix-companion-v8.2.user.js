// ==UserScript==
// @name         Excel Matrix Automation Companion (Sanitized Demo)
// @namespace    http://tampermonkey.net
// @version      8.2
// @description  Replaces text area exceptions with an interactive live UI grid for point-and-click cell protection on Excel Online / SharePoint. Portfolio/demo build.
// @author       Danilo Jose Rengifo Sulbaran (check GitHub Repo)
// @match        https://*.sharepoint.com/*
// @match        https://*.live.com/*
// @match        https://*.office.com/*
// @allFrames    true
// @noframes     false
// @grant        none
// @run-at       document-idle
// ==/UserScript==

// NOTE (sanitized build): the original version pinned @match to several
// specific SharePoint/Teams document URLs (site names + doc GUIDs + an
// embedded user email in the query string) across multiple store sites.
// That's been replaced here with broad domain matches; the script's own
// runtime check below narrows it further. Point this at your own
// workbook(s) by tightening @match or the currentUrl check as needed.

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
            "tab_default": { id: "tab_default", name: "Main Matrix", rawInput: "", exceptions: {} }
        }
    };

    // Hydrate persistent multi-tab database safely
    try {
        const savedState = localStorage.getItem('emd_multitab_state_v7');
        if (savedState) {
            tabsState = JSON.parse(savedState);
        }
    } catch(e) {
        console.error("Failed to rehydrate multi-tab state database:", e);
    }

    function saveStateToStorage() {
        localStorage.setItem('emd_multitab_state_v7', JSON.stringify(tabsState));
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
            updateTabHeadersUI(); // Keep header synchronization alive
            return;
        }

        const panel = document.createElement('div');
        panel.id = 'excel-injector-companion-panel';
        panel.style = `
            position: fixed; bottom: 30px; left: 30px; width: 360px;
            background: #107c41; color: white; border-radius: 6px;
            z-index: 2147483647; padding: 12px; font-family: Segoe UI, Arial, sans-serif;
            box-shadow: 0px 4px 16px rgba(0,0,0,0.5); font-size: 12px;
            text-align: center; border: 1px solid #0b592e;
            display: flex; flex-direction: column; gap: 6px;
            user-select: none; max-height: 88vh;
        `;

        panel.innerHTML = `
            <div id="emd-panel-header" style="font-weight: bold; padding-bottom: 4px; border-bottom: 1px solid #0b592e; cursor: move; flex-shrink: 0; display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <span id="emd-title-text">Excel Matrix Dropper v7.0</span>
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

                <div style="text-align: left; font-size: 10px; color: #c8e6c9; font-weight: bold; flex-shrink: 0;">1. Paste Truck/Team Pic List:</div>
                <textarea id="pic-config-input" placeholder="0421RTE01\\tYES\\tYES..." style="width: 100%; height: 50px; box-sizing: border-box; border-radius: 3px; border: none; padding: 4px; font-size: 10px; font-family: monospace; resize: none; color: #333; background: #fff; user-select: text; flex-shrink: 0;"></textarea>

                <div style="text-align: left; font-size: 10px; color: #c8e6c9; font-weight: bold; flex-shrink: 0;">2. Live Exception Overrides Grid (<span id="active-grid-title-label"></span>):</div>
                <div style="width: 100%; overflow: auto; background: #fff; border-radius: 4px; max-height: 160px; border: 1px solid #0b592e; flex-grow: 1;">
                    <table id="ui-override-grid" style="width: 100%; border-collapse: collapse; font-size: 10px; color: #333; text-align: left;">
                        <thead style="background: #f3f3f3; position: sticky; top: 0; z-index: 10;">
                            <tr style="border-bottom: 1px solid #ddd;">
                                <th style="padding: 4px; border-right: 1px solid #ddd; min-width: 90px; background: #f3f3f3;">Route</th>
                                ${Array.from({length: 20}, (_, i) => `<th style="padding: 4px; border-right: 1px solid #ddd; text-align: center; min-width: 35px;">${i + 1}</th>`).join('')}
                            </tr>
                        </thead>
                        <tbody id="ui-override-grid-body"></tbody>
                    </table>
                </div>

                <div style="display: flex; gap: 4px; width: 100%; flex-shrink: 0;">
                    <button id="save-pic-memory-btn" style="flex: 1; background: #1565c0; color: white; border: none; padding: 6px; border-radius: 3px; cursor: pointer; font-size: 10px; font-weight: bold;">Save Configs</button>
                    <button id="clear-pic-memory-btn" style="flex: 1; background: #c62828; color: white; border: none; padding: 6px; border-radius: 3px; cursor: pointer; font-size: 10px; font-weight: bold;">Reset Active Tab</button>
                </div>

                <div style="text-align: left; font-size: 10px; color: #c8e6c9; font-weight: bold; flex-shrink: 0;">3. Upload Route Status CSV:</div>
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

                // Visible, tactical mini-pill layout transformation
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

                panel.style.width = '360px';
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
            if (currentMin) {
                localStorage.setItem('emd_panel_minimized', 'false');
                applyMinimizeLayout(false);
            } else {
                localStorage.setItem('emd_panel_minimized', 'true');
                applyMinimizeLayout(true);
            }
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
        document.getElementById('pic-config-input').addEventListener('input', handleRawInputTextChange);
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
            renameBtn.addEventListener('mouseenter', () => renameBtn.style.opacity = '1');
            renameBtn.addEventListener('mouseleave', () => renameBtn.style.opacity = '0.6');

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
                closeBtn.addEventListener('mouseenter', () => closeBtn.style.color = '#fff');
                closeBtn.addEventListener('mouseleave', () => closeBtn.style.color = 'rgba(255,255,255,0.6)');
                closeBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    if (confirm(`Completely delete "${tab.name}" pipeline memory workspace?`)) {
                        delete tabsState.tabs[tabId];
                        localStorage.removeItem(`emd_shift_map_${tabId}`);
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
            rawInput: "",
            exceptions: {}
        };
        tabsState.activeTabId = id;
        saveStateToStorage();
        switchWorkspaceContext();
    }

    function switchWorkspaceContext() {
        const activeTab = getActiveTab();
        updateTabHeadersUI();

        document.getElementById('pic-config-input').value = activeTab.rawInput || "";
        document.getElementById('active-grid-title-label').textContent = activeTab.name;
        document.getElementById('injector-status-msg').textContent = `Workspace switched to ${activeTab.name}`;
        document.getElementById('injector-status-msg').style.color = "#e1f5fe";

        renderDynamicOverrideGrid();
    }

    function handleRawInputTextChange(e) {
        const activeTab = getActiveTab();
        activeTab.rawInput = e.target.value;
        saveStateToStorage();
        renderDynamicOverrideGrid();
    }

    function renderDynamicOverrideGrid() {
        const activeTab = getActiveTab();
        const tbody = document.getElementById('ui-override-grid-body');
        if (!tbody) return;

        const lines = (activeTab.rawInput || "").split(/[\r\n]+/);
        let html = "";
        let processedRoutes = new Set();
        let rowIndex = 0;

        lines.forEach(line => {
            if (!line.trim()) return;
            const tabs = line.split('\t');
            if (tabs.length > 0) {
                const originalRouteName = tabs[0].trim();
                const cleanRoute = sanitizeKey(originalRouteName);
                if (!cleanRoute || processedRoutes.has(cleanRoute)) return;
                processedRoutes.add(cleanRoute);

                html += `<tr data-route="${cleanRoute}" data-row="${rowIndex}" style="border-bottom: 1px solid #eee; background: #fff;">`;

                // FROZEN ROUTE LAYER: Locked label column sitting dynamically at left: 0
                html += `<td style="padding: 4px; border-right: 1px solid #eee; font-weight: bold; font-family: monospace; background: #fafafa; font-size: 10px; position: sticky; left: 0; z-index: 5; box-shadow: 2px 0 5px -2px rgba(0,0,0,0.15); white-space: nowrap;">${originalRouteName}</td>`;

                // INPUT LAYER: Generates override input boxes from Stop 1 up to Stop 20
                for (let stopNum = 1; stopNum <= 20; stopNum++) {
                    const currentSavedValue = (activeTab.exceptions[cleanRoute] && activeTab.exceptions[cleanRoute][stopNum]) ? activeTab.exceptions[cleanRoute][stopNum] : "";
                    html += `
                        <td style="padding: 2px; border-right: 1px solid #eee; text-align: center; position: relative;">
                            <input type="text" data-route="${cleanRoute}" data-row="${rowIndex}" data-col="${stopNum}" data-stop="${stopNum}" value="${currentSavedValue}" placeholder="-"
                                style="width: 100%; border: 2px solid transparent; text-align: center; font-size: 10px; font-weight: bold; text-transform: uppercase; padding: 2px 0; outline: none; background: transparent; box-sizing: border-box; transition: none; z-index: 2; position: relative;">
                        </td>`;
                }
                html += `</tr>`;
                rowIndex++;
            }
        });

        if (html === "") {
            tbody.innerHTML = `<tr><td colspan="21" style="padding: 10px; text-align: center; color: #999; background: #fff;">Paste your tab Pic list to generate row cells.</td></tr>`;
            return;
        } else {
            tbody.innerHTML = html;
        }

        const inputs = Array.from(tbody.querySelectorAll('input'));

        inputs.forEach((input) => {
            // Live data persistence sync listener
            input.addEventListener('input', (e) => {
                const r = e.target.getAttribute('data-route');
                const s = e.target.getAttribute('data-stop');
                const val = e.target.value.trim().toUpperCase();

                if (!activeTab.exceptions[r]) activeTab.exceptions[r] = {};
                if (val === "") {
                    delete activeTab.exceptions[r][s];
                    if (Object.keys(activeTab.exceptions[r]).length === 0) delete activeTab.exceptions[r];
                } else {
                    activeTab.exceptions[r][s] = val;
                }
                saveStateToStorage();
            });

            // PASTE INTERCEPTOR ENGINE: Allows bulk pasting whole tables straight into the grid
            input.addEventListener('paste', (e) => {
                e.preventDefault();
                e.stopPropagation();

                const clipboardData = e.clipboardData || window.clipboardData;
                const pastedText = clipboardData.getData('text');
                if (!pastedText) return;

                // Split copied rows by line breaks, then split cell content by tabs
                const pasteRows = pastedText.split(/[\r\n]+/);
                const startRowIndex = parseInt(input.getAttribute('data-row'), 10);
                const startColIndex = parseInt(input.getAttribute('data-col'), 10);

                const tableRows = Array.from(tbody.querySelectorAll('tr'));

                pasteRows.forEach((rowText, rowOffset) => {
                    if (!rowText.trim()) return;
                    const cellValues = rowText.split('\t');
                    const targetRowEl = tableRows[startRowIndex + rowOffset];
                    if (!targetRowEl) return;

                    const rKey = targetRowEl.getAttribute('data-route');
                    if (!activeTab.exceptions[rKey]) activeTab.exceptions[rKey] = {};

                    cellValues.forEach((cellVal, colOffset) => {
                        const targetStopNum = startColIndex + colOffset;
                        if (targetStopNum > 20) return; // Cap at stop index 20

                        const cleanVal = cellVal.trim().toUpperCase();
                        const targetInput = targetRowEl.querySelector(`input[data-stop="${targetStopNum}"]`);

                        if (targetInput) {
                            targetInput.value = cleanVal;
                            if (cleanVal === "") {
                                delete activeTab.exceptions[rKey][targetStopNum];
                            } else {
                                activeTab.exceptions[rKey][targetStopNum] = cleanVal;
                            }
                        }
                    });
                });

                if (Object.keys(activeTab.exceptions).forEach(k => { if(Object.keys(activeTab.exceptions[k]).length === 0) delete activeTab.exceptions[k]; }));
                saveStateToStorage();
            });

            // CROSSHAIR HIGHLIGHTERS
            input.addEventListener('focus', () => {
                const activeRow = input.getAttribute('data-row');
                const activeCol = input.getAttribute('data-col');
                input.style.border = '2px solid #107c41';
                input.style.background = '#ffffff';

                const parentRow = input.closest('tr');
                if (parentRow) {
                    const frozenRouteCell = parentRow.querySelector('td:first-child');
                    if (frozenRouteCell) frozenRouteCell.style.background = '#d1e7dd';
                }

                inputs.forEach(otherInput => {
                    const r = otherInput.getAttribute('data-row');
                    const c = otherInput.getAttribute('data-col');
                    if (otherInput !== input) {
                        if (r === activeRow || c === activeCol) {
                            otherInput.style.background = '#d1e7dd';
                            otherInput.style.border = '2px solid transparent';
                        } else {
                            otherInput.style.background = 'transparent';
                            otherInput.style.border = '2px solid transparent';
                        }
                    }
                });
            });

            input.addEventListener('blur', () => {
                input.style.border = '2px solid transparent';
                input.style.background = 'transparent';
                const parentRow = input.closest('tr');
                if (parentRow) {
                    const frozenRouteCell = parentRow.querySelector('td:first-child');
                    if (frozenRouteCell) frozenRouteCell.style.background = '#fafafa';
                }
                inputs.forEach(otherInput => {
                    otherInput.style.background = 'transparent';
                    otherInput.style.border = '2px solid transparent';
                });
            });

            input.addEventListener('keydown', (e) => {
                const currentStop = parseInt(input.getAttribute('data-stop'), 10);
                const currentIdx = inputs.indexOf(input);
                let targetInput = null;

                if (e.key === 'ArrowRight' && currentStop < 20) targetInput = inputs[currentIdx + 1];
                else if (e.key === 'ArrowLeft' && currentStop > 1) targetInput = inputs[currentIdx - 1];
                else if (e.key === 'ArrowDown') targetInput = inputs[currentIdx + 20];
                else if (e.key === 'ArrowUp') targetInput = inputs[currentIdx - 20];

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

        let parsedShiftMap = {};
        const lines = (activeTab.rawInput || "").split(/[\r\n]+/);

        lines.forEach(line => {
            if (!line.trim()) return;
            const tabs = line.split('\t');
            if (tabs.length >= 3) {
                const rKey = sanitizeKey(tabs[0]);
                parsedShiftMap[rKey] = {
                    truck: tabs[1].trim().toUpperCase(),
                    team: tabs[2].trim().toUpperCase()
                };
            }
        });

        localStorage.setItem(`emd_shift_map_${activeTab.id}`, JSON.stringify(parsedShiftMap));
        saveStateToStorage();

        statusEl.textContent = `[${activeTab.name}] Configurations Saved! ✓`;
        statusEl.style.color = "#a5d6a7";
    }

    function clearActiveTabMemoryScope() {
        const activeTab = getActiveTab();
        if (confirm(`Reset all configuration matrix items inside "${activeTab.name}"?`)) {
            activeTab.rawInput = "";
            activeTab.exceptions = {};
            localStorage.removeItem(`emd_shift_map_${activeTab.id}`);
            saveStateToStorage();
            switchWorkspaceContext();

            const statusEl = document.getElementById('injector-status-msg');
            statusEl.textContent = "Active workspace memory wiped clean.";
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

        let tabShiftMap = {};
        const savedShiftStr = localStorage.getItem(`emd_shift_map_${activeTab.id}`);
        if (savedShiftStr) {
            try { tabShiftMap = JSON.parse(savedShiftStr); } catch(e) {}
        }

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

                if (tabShiftMap[routeKey]) {
                    reconstructedLineCells.push(tabShiftMap[routeKey].truck);
                    reconstructedLineCells.push(tabShiftMap[routeKey].team);
                } else {
                    reconstructedLineCells.push("");
                    reconstructedLineCells.push("");
                }

                const stopShortcodes = columns.slice(13);
                stopShortcodes.forEach((code, index) => {
                    const stopNumber = index + 1;
                    let cleanCode = code ? code.toString().trim().toUpperCase() : "";

                    // If a custom override is pasted into the grid, prioritize it completely
                    if (activeTab.exceptions[routeKey] && activeTab.exceptions[routeKey][stopNumber]) {
                        cleanCode = activeTab.exceptions[routeKey][stopNumber];
                    } else if (cleanCode === "MC" || cleanCode === "MANUAL CHECK") {
                        cleanCode = "DAM";
                    }

                    reconstructedLineCells.push(cleanCode);
                });

                unifiedClipboardRows.push(reconstructedLineCells.join('\t'));
            });

            if (unifiedClipboardRows.length === 0) {
                statusEl.textContent = "No valid data rows matched tab filters!";
                statusEl.style.color = "#ffcc00";
                return;
            }

            const streamPayload = unifiedClipboardRows.join('\n');
            statusEl.textContent = "Streaming formatted tab payload...";

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