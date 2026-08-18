// ==UserScript==
// @name         EDP Routing Automation V1.2
// @namespace    http://tampermonkey.net
// @version      1.2
// @description  Automates tracking route stop statuses on an internal route-management portal, with local CSV export pipelines and multi-tab workspace profiles
// @author       Internal Tooling
// @match        https://your-edc-portal.example.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

// NOTE: this is a sanitized template. Set @match above to your actual
// portal URL before use — the CSS selectors and XPath expressions below
// still target that portal's real DOM structure, so they only work
// once pointed at the correct page.

(function() {
    'use strict';

    // --- Core State & Multi-tab Database Engine ---
    let tabsState = {
        activeTabId: "edp_tab_default",
        tabs: {
            "edp_tab_default": {
                id: "edp_tab_default",
                name: "Routing Profile 1",
                routesToProcess: [],
                currentRouteIndex: 0,
                stopHistory: {},
                finalOutputData: {},
                userSkippedRoutes: []
            }
        }
    };

    let timerCountdown = 15 * 60;
    let countdownInterval = null;
    let masterRosterData = {};

    // Valid statuses that don't need a manual check fallback
    const expectedStatuses = ["DELIVERED", "IN TRANSIT", "OUT FOR DELIVERY", "NARROWED", "ARRIVED", "LOADED", "NOT LOADED"];

    // Translation Mapping Database Core
    const transMap = {
        "ARRIVED": "A", "CANCELED": "CA", "CANCELLED": "CA", "CUSTOMER R/S COMPLIANT": "CXRSC",
        "CUSTOMER R/S NOT COMPLIANT": "CXRSNC", "DELIVERED": "D", "DAMAGED ITEMDA": "DAM",
        "DRIVER DAMAGE": "DD", "DELIVERED EARLY": "DE", "DEL": "DEL", "DELETED": "DEL",
        "DELIVERED LATE": "DL", "INVALID NOT AT HOME": "INAH", "IN TRANSIT": "IT",
        "LOADED": "L", "NOT LOADED": "NL", "NARROWED": "NW", "OUT FOR DELIVERY": "OFD",
        "PARTIALLY DELIVERED": "PD", "STORE STOP": "SS", "SAVE THE SALE DAMAGED": "STSD",
        "VALID NOT AT HOME": "VNAH", "WRONG ADDRESS": "WA", "MANUAL CHECK": "MC"
    };

    // Restore working data safely across automated page refreshes
    try {
        const savedRoster = localStorage.getItem('edp_master_roster');
        if (savedRoster) masterRosterData = JSON.parse(savedRoster);

        const savedMultiTabState = localStorage.getItem('edp_multitab_state_v11');
        if (savedMultiTabState) {
            tabsState = JSON.parse(savedMultiTabState);
        }
    } catch(e) {
        console.error("Failed to rehydrate working loop context:", e);
    }

    function saveStateToStorage() {
        localStorage.setItem('edp_multitab_state_v11', JSON.stringify(tabsState));
    }

    function getActiveTab() {
        const keys = Object.keys(tabsState.tabs);
        if (!tabsState.tabs[tabsState.activeTabId]) {
            tabsState.activeTabId = keys[0] || "edp_tab_default";
        }
        return tabsState.tabs[tabsState.activeTabId];
    }

    function updateUIStatus(statusText) {
        const statusEl = document.getElementById('bot-status');
        if (statusEl) statusEl.textContent = statusText;
    }

    // --- Simulated Human Clicks Function ---
    function simulateHumanClick(element) {
        if (!element) return;
        ['mousedown', 'mouseup', 'click'].forEach(eventType => {
            element.dispatchEvent(new MouseEvent(eventType, {
                bubbles: true,
                cancelable: true,
                view: window
            }));
        });
    }

    // --- DOM Utilities ---
    function waitForElement(selector, timeout = 20000) {
        return new Promise((resolve, reject) => {
            const el = document.querySelector(selector);
            if (el) return resolve(el);
            const observer = new MutationObserver(() => {
                const element = document.querySelector(selector);
                if (element) {
                    observer.disconnect();
                    resolve(element);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            setTimeout(() => {
                observer.disconnect();
                reject(new Error(`Timeout waiting for selector: ${selector}`));
            }, timeout);
        });
    }

    function waitForXPath(xpath, timeout = 20000) {
        return new Promise((resolve, reject) => {
            const getEl = () => document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            const el = getEl();
            if (el) return resolve(el);
            const observer = new MutationObserver(() => {
                const element = getEl();
                if (element) {
                    observer.disconnect();
                    resolve(element);
                }
            });
            observer.observe(document.body, { childList: true, subtree: true });
            setTimeout(() => {
                observer.disconnect();
                reject(new Error(`Timeout waiting for XPath: ${xpath}`));
            }, timeout);
        });
    }

    function setNativeValue(element, value) {
        const valueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
        const prototype = Object.getPrototypeOf(element);
        const prototypeValueSetter = Object.getOwnPropertyDescriptor(prototype, "value")?.set;

        if (prototypeValueSetter && prototypeValueSetter !== valueSetter) {
            prototypeValueSetter.call(element, value);
        } else if (valueSetter) {
            valueSetter.call(element, value);
        } else {
            element.value = value;
        }

        element.dispatchEvent(new Event('input', { bubbles: true }));
        element.dispatchEvent(new Event('change', { bubbles: true }));
    }

    function parseExcelData(text) {
        if (!text.trim()) return [];
        const lines = text.split(/\r?\n/);
        const parsedRoutes = [];
        lines.forEach(line => {
            // Updated regex to explicitly isolate the store number parts separate from the asset tags
            const match = line.match(/\b(\d+)(?:RLXBOX|RLXLFB|LFB)(\d*)\b/i);

            if (match) {
                const storeNum = match[1].padStart(4, '0'); // Forces short numbers (like 75) into 4 digits (0075)
                const assetTag = line.match(/(?:RLXBOX|RLXLFB|LFB)/i)[0].toUpperCase();
                const suffix = match[2];
                parsedRoutes.push(`${storeNum}${assetTag}${suffix}`);
            } else if (line.trim().match(/^(\d+)(?:RLXBOX|RLXLFB|LFB)(\d*)$/i)) {
                const directMatch = line.trim().match(/^(\d+)(?:RLXBOX|RLXLFB|LFB)(\d*)$/i);
                const storeNum = directMatch[1].padStart(4, '0');
                const assetTag = line.trim().match(/(?:RLXBOX|RLXLFB|LFB)/i)[0].toUpperCase();
                const suffix = directMatch[2];
                parsedRoutes.push(`${storeNum}${assetTag}${suffix}`);
            }
        });
        return parsedRoutes;
    }

    // --- Core Automation Loop ---
    async function processNextRoute() {
        const activeTab = getActiveTab();

        if (activeTab.currentRouteIndex >= activeTab.routesToProcess.length) {
            updateUIStatus("Complete");
            startHibernationTimer();
            return;
        }

        const fullRoute = activeTab.routesToProcess[activeTab.currentRouteIndex];

        if (!activeTab.userSkippedRoutes) activeTab.userSkippedRoutes = [];
        if (activeTab.userSkippedRoutes.includes(fullRoute)) {
            console.log(`Bypassing route processing sequence for flagged path target: ${fullRoute}`);
            activeTab.finalOutputData[fullRoute] = `${fullRoute}\tSKIPPED BY USER`;
            saveStateToStorage();
            renderOutputList();
            nextIteration();
            return;
        }

        updateUIStatus(`Checking ${activeTab.currentRouteIndex + 1}/${activeTab.routesToProcess.length}`);

        try {
            let storeMatch = fullRoute.match(/^\d+/);
            if (!storeMatch) throw new Error("Could not parse store number from route code.");
            let storeNumber = storeMatch[0];

            // Step 1: Open profile dropdown menu
            updateUIStatus("Opening profile dropdown...");
            const profileBtn = await waitForElement('a[data-testid="profile-dropdown"]');
            simulateHumanClick(profileBtn);
            await new Promise(r => setTimeout(r, 800));

            // Step 2: Click the "Shadowing Store Number:" span option
            updateUIStatus("Selecting Shadowing Store option...");
            const shadowTextSpan = await waitForXPath("//span[contains(normalize-space(.), 'Shadowing Store Number:')]");
            simulateHumanClick(shadowTextSpan);
            await new Promise(r => setTimeout(r, 800));

            // Step 3: Target the input field and securely set the store number
            updateUIStatus("Entering store number...");
            const storeInput = await waitForElement('input#textInput');
            setNativeValue(storeInput, storeNumber);
            await new Promise(r => setTimeout(r, 800));

            // Step 4: Click the primary Select button
            updateUIStatus("Confirming store select...");
            const selectBtn = await waitForXPath("//button[@type='button' and contains(@class, 'btn-primary')][normalize-space(.)='Select']");
            simulateHumanClick(selectBtn);

            // Step 5: Wait for layout configuration
            updateUIStatus("Loading data layers...");
            await new Promise(r => setTimeout(r, 5000));

            await scrapeActiveRoutePanelDetails(fullRoute);

        } catch (error) {
            console.error(error);
            activeTab.finalOutputData[fullRoute] = `${fullRoute}\tERROR: ${error.message}`;
            saveStateToStorage();
            renderOutputList();
            nextIteration();
        }
    }

    // --- Updated Scraper Function ---
    async function scrapeActiveRoutePanelDetails(fullRoute) {
        const activeTab = getActiveTab();
        updateUIStatus(`Processing data for ${fullRoute}...`);

        try {
            // Step 6: Select asset class from Picker dropdown
            updateUIStatus("Picking asset type...");
            const resourcePicker = await waitForElement('select.resource-type-picker');
            const isFlatBed = fullRoute.toUpperCase().includes("RLXLFB");
            const targetAssetClass = isFlatBed ? "Large Flat Bed" : "Box Truck";

            resourcePicker.value = targetAssetClass;
            resourcePicker.dispatchEvent(new Event('change', { bubbles: true }));
            await new Promise(r => setTimeout(r, 2000));

            // Step 7: Click route row matching text content securely via dynamic containing check
            updateUIStatus(`Finding route row for ${fullRoute}...`);

            // Remove leading zeros just for the screen selection match (e.g., changes "0520RLXBOX2" back to "520RLXBOX2")
            const displayRouteName = fullRoute.replace(/^0+/, '');

            const routeRowEl = await waitForXPath(`//span[contains(@class, 'resource-id')][contains(normalize-space(.), '${displayRouteName}')]`);
            simulateHumanClick(routeRowEl);
            await new Promise(r => setTimeout(r, 1000));

            // Step 8: Open Route details panel sidebar element using space normalization
            updateUIStatus("Opening Route Details pane...");
            const detailsPanelBtn = await waitForXPath("//li[contains(@class, 'enabled')][contains(normalize-space(.), 'View Route Details')]");
            simulateHumanClick(detailsPanelBtn);

            // Allow panel animations to finalize
            await new Promise(r => setTimeout(r, 3500));

            // Parse rows inside the explicit table structures inside the sidebar
            const rows = document.querySelectorAll('.all-stops-container table tbody tr');
            if (rows.length === 0) {
                throw new Error("No stop records identified within sidebar structure.");
            }

            let currentRunStopIds = [];
            let stopDataMap = {};
            let highestStopNum = 0;

            rows.forEach((row) => {
                const seqEl = row.querySelector('.sequence-no');
                if (!seqEl) return;

                let stopNum = parseInt(seqEl.textContent.trim(), 10);
                if (isNaN(stopNum)) return;
                if (stopNum > highestStopNum) highestStopNum = stopNum;

                const orderLink = row.querySelector('.order-id a');
                const orderId = orderLink ? orderLink.textContent.trim() : `STOP_${stopNum}`;
                currentRunStopIds.push(orderId);

                const cardTextContent = row.innerText || "";
                // Flags a stop as a store-to-store transfer when the retail
                // brand name appears in the stop text (case-insensitive).
                // Replace RETAIL_BRAND_NAME below with your actual retailer name.
                const RETAIL_BRAND_NAME = "LOWE'S";
                let isStoreSaleStop = new RegExp(`\\b${RETAIL_BRAND_NAME}\\b`, "i").test(cardTextContent);

                // Detect color directly from element custom attributes
                let detectedTimeColor = null;
                const arrivalTimeSpan = row.querySelector('.arrived-time');
                if (arrivalTimeSpan) {
                    const attrColor = arrivalTimeSpan.getAttribute('text-color');
                    if (attrColor === 'yellow') detectedTimeColor = 'yellow';
                    if (attrColor === 'red') detectedTimeColor = 'red';
                }

                // Gather badge text context values
                let rawStatus = "";
                const badgeEl = row.querySelector('.badge');
                if (badgeEl) {
                    rawStatus = badgeEl.textContent.trim().toUpperCase();
                }

                let cleanStatus = "MC";
                if (isStoreSaleStop) {
                    cleanStatus = "SS";
                } else if (rawStatus === "DELIVERED" && detectedTimeColor === "yellow") {
                    cleanStatus = "DE";
                } else if (rawStatus === "DELIVERED" && detectedTimeColor === "red") {
                    cleanStatus = "DL";
                } else if (expectedStatuses.includes(rawStatus)) {
                    cleanStatus = transMap[rawStatus] || rawStatus;
                }

                stopDataMap[stopNum] = cleanStatus;
            });

            let routeRowParts = [fullRoute];
            for (let i = 1; i <= highestStopNum; i++) {
                routeRowParts.push(stopDataMap[i] || "DEL");
            }
            activeTab.stopHistory[fullRoute] = currentRunStopIds;
            activeTab.finalOutputData[fullRoute] = routeRowParts.join('\t');
            saveStateToStorage();
            renderOutputList();
        } catch (error) {
            console.error(error);
            activeTab.finalOutputData[fullRoute] = `${fullRoute}\tERROR: ${error.message}`;
            saveStateToStorage();
            renderOutputList();
        }
        nextIteration();
    }

    function nextIteration() {
        const activeTab = getActiveTab();
        activeTab.currentRouteIndex++;
        saveStateToStorage();
        setTimeout(processNextRoute, 1500);
    }

    function startHibernationTimer() {
        timerCountdown = 15 * 60;
        clearInterval(countdownInterval);
        countdownInterval = setInterval(() => {
            timerCountdown--;
            const mins = Math.floor(timerCountdown / 60);
            const secs = timerCountdown % 60;
            updateUIStatus("Standby (" + mins + ":" + (secs < 10 ? '0' : '') + secs + ")");
            if (timerCountdown <= 0) {
                clearInterval(countdownInterval);
                window.location.reload();
            }
        }, 1000);
    }

    // --- Grid Rendering Logic with Character Translation Mapping ---
    function renderOutputList() {
        const activeTab = getActiveTab();
        const tbody = document.getElementById('bot-grid-body');
        const table = document.getElementById('bot-grid-table');
        if (!tbody || !table) return;
        if (!activeTab.routesToProcess || activeTab.routesToProcess.length === 0) {
            tbody.innerHTML = `<tr><td colspan="2" style="padding: 10px; color: #888; text-align: center;">No data loaded yet.</td></tr>`;
            return;
        }
        let translatedOutputData = {};
        let maxStopsCount = 0;
        activeTab.routesToProcess.forEach(route => {
            if (!route) return;
            const rawLine = activeTab.finalOutputData[route];
            if (rawLine && rawLine.includes('\t')) {
                const parts = rawLine.split('\t');
                const routeName = parts[0];
                const cleanShortCodes = parts.slice(1).map(status => {
                    const lookup = status ? status.trim().toUpperCase() : "";
                    return transMap[lookup] || lookup;
                });
                translatedOutputData[routeName] = [routeName, ...cleanShortCodes].join('\t');
                if (cleanShortCodes.length > maxStopsCount) {
                    maxStopsCount = cleanShortCodes.length;
                }
                if (masterRosterData && typeof masterRosterData === 'object') {
                    masterRosterData[routeName] = cleanShortCodes;
                }
            } else {
                translatedOutputData[route] = rawLine || `${route}`;
                if (masterRosterData && masterRosterData[route] && masterRosterData[route].length > maxStopsCount) {
                    maxStopsCount = masterRosterData[route].length;
                }
            }
        });
        localStorage.setItem('edp_master_roster', JSON.stringify(masterRosterData || {}));
        let headerHtml = `<tr style="background: #002060; color: #ffffff; border-bottom: 2px solid #ccc; position: sticky; top: 0; z-index: 10;"> <th style="padding: 6px; border-right: 1px solid #ddd; min-width: 120px;">Route:</th>`;
        for (let i = 1; i <= Math.max(maxStopsCount, 1); i++) {
            headerHtml += `<th style="padding: 6px; border-right: 1px solid #ddd; text-align: center; min-width: 40px;">${i}</th>`;
        }
        headerHtml += `</tr>`;
        table.querySelector('thead').innerHTML = headerHtml;
        let html = "";
        if (!activeTab.userSkippedRoutes) activeTab.userSkippedRoutes = [];
        activeTab.routesToProcess.forEach(route => {
            if (!route) return;
            const dataRow = translatedOutputData[route];
            const isUserSkipped = activeTab.userSkippedRoutes.includes(route);
            let rowHtml = `<tr style="background: ${isUserSkipped ? '#ffebee' : '#fff'}; opacity: ${isUserSkipped ? '0.5' : '1'}; transition: background 0.2s;">`;
            rowHtml +=  `<td class="edp-route-toggle-btn" data-route="${route}" style="padding: 6px; border-right: 1px solid #ddd; font-weight: bold; background: ${isUserSkipped ? '#ef9a9a' : '#fff'}; color: ${isUserSkipped ? '#b71c1c' : '#000'}; position: sticky; left: 0; border-bottom: 1px solid #eee; cursor: pointer; text-align: left; user-select: none;" title="Click to skip/include this route"> ${isUserSkipped ? '🚫 ' : ''}${route} </td>`;
            const segments = (dataRow && dataRow.includes('\t')) ? dataRow.split('\t').slice(1) : (masterRosterData[route] || []);
            if (isUserSkipped) {
                rowHtml += `<td colspan="${Math.max(maxStopsCount, 1)}" style="padding: 6px; border-bottom: 1px solid #eee; color: #b71c1c; font-style: italic; font-weight: bold; background: #ffebee; text-align: left;">[SKIPPED BY USER] - Click route code to restore tracking</td>`;
            } else if (segments && segments.length > 0) {
                for (let i = 0; i < Math.max(maxStopsCount, 1); i++) {
                    const shortCode = segments[i];
                    if (shortCode) {

                        let bg = "#e0e0e0"; let color = "#333";
                        if (shortCode === "D" || shortCode === "DE" || shortCode === "DL") { bg = "#e8f5e9"; color = "#2e7d32"; }
                        else if (shortCode === "IT") { bg = "#e3f2fd"; color = "#1565c0"; }
                        else if (shortCode === "OFD") { bg = "#fff3e0"; color = "#ef6c00"; }
                        else if (shortCode === "NW") { bg = "#f3e5f5"; color = "#6a1b9a"; }
                        else if (shortCode === "DEL" || shortCode === "DAM" || shortCode === "DD") { bg = "#ffebee"; color = "#c62828"; }
                        else if (shortCode === "SS" || shortCode === "A" || shortCode === "L") { bg = "#eceff1"; color = "#455a64"; }
                        else if (shortCode === "MC") { bg = "#fffde7"; color = "#f57f17"; }
                        rowHtml +=  `<td style="padding: 4px; border-right: 1px solid #eee; border-bottom: 1px solid #eee; text-align: center;"> <span style="background: ${bg}; color: ${color}; padding: 2px 6px; border-radius: 3px; border: 1px solid rgba(0,0,0,0.05); display: inline-block; font-weight: bold; font-size: 11px; min-width: 24px;">${shortCode}</span> </td>`;
                    } else {
                        rowHtml += `<td style="padding: 4px; border-right: 1px solid #eee; border-bottom: 1px solid #eee;"></td>`;
                    }
                }
            } else if (dataRow) {
                const errorMsg = dataRow.replace(route, '').trim();
                rowHtml += `<td colspan="${Math.max(maxStopsCount, 1)}" style="padding: 6px; border-bottom: 1px solid #eee; color: #c62828; font-style: italic;">${errorMsg}</td>`;
            } else {
                rowHtml += `<td colspan="${Math.max(maxStopsCount, 1)}" style="padding: 6px; border-bottom: 1px solid #eee; color: #999;">Pending...</td>`;
            }
            rowHtml += `</tr>`;
            html += rowHtml;
        });
        tbody.innerHTML = html;
        table.querySelectorAll('.edp-route-toggle-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const selectedRoute = e.currentTarget.getAttribute('data-route');
                if (!activeTab.userSkippedRoutes) activeTab.userSkippedRoutes = [];
                const routeIndex = activeTab.userSkippedRoutes.indexOf(selectedRoute);
                if (routeIndex > -1) {
                    activeTab.userSkippedRoutes.splice(routeIndex, 1);
                } else {
                    activeTab.userSkippedRoutes.push(selectedRoute);
                }
                saveStateToStorage();
                renderOutputList();
            });
        });
    }

    // --- Dynamic Multi-Tab Workspace Management UI ---
    function updateTabHeadersUI() {
        const wrapper = document.getElementById('edp-tabs-wrapper');
        if (!wrapper) return;
        wrapper.innerHTML = "";
        Object.keys(tabsState.tabs).forEach(tabId => {
            const tab = tabsState.tabs[tabId];
            const isActive = tabId === tabsState.activeTabId;
            const tabEl = document.createElement('div');
            tabEl.style = `background: ${isActive ? '#002060' : 'rgba(0,114,206,0.15)'}; color: ${isActive ? '#fff' : '#0072CE'}; padding: 4px 8px; font-size: 10px; font-weight: bold; border-radius: 4px 4px 0 0; cursor: pointer; display: flex; align-items: center; gap: 4px; border: 1px solid ${isActive ? '#001030' : 'transparent'}; border-bottom: none; max-width: 110px; overflow: hidden; white-space: nowrap; margin-bottom: 1px;`;
            tabEl.title = tab.name;
            const nameSpan = document.createElement('span');
            nameSpan.style = "overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 65px;";
            nameSpan.textContent = tab.name;
            tabEl.appendChild(nameSpan);
            const renameBtn = document.createElement('span');
            renameBtn.innerHTML = "✏️";
            renameBtn.style = "cursor: pointer; font-size: 10px; opacity: 0.6; padding: 0 2px;";
            renameBtn.title = "Rename Queue";
            const triggerRename = (e) => {
                e.stopPropagation();
                const newName = prompt("Rename automation queue workspace:", tab.name);
                if (newName && newName.trim()) {
                    tab.name = newName.trim();
                    saveStateToStorage();
                    updateTabHeadersUI();
                }
            };
            renameBtn.addEventListener('click', triggerRename);
            tabEl.addEventListener('dblclick', triggerRename);
            tabEl.appendChild(renameBtn);
            if (Object.keys(tabsState.tabs).length > 1) {
                const closeBtn = document.createElement('span');
                closeBtn.textContent = "×";
                closeBtn.style = "color: rgba(255,255,255,0.6); font-size: 12px; font-weight: bold; cursor: pointer; margin-left: 2px;";
                closeBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    if (confirm(`Completely delete "${tab.name}" tracking queue context?`)) {
                        delete tabsState.tabs[tabId];
                        if (tabsState.activeTabId === tabId) {
                            tabsState.activeTabId = Object.keys(tabsState.tabs);
                        }
                        saveStateToStorage();
                        switchWorkspaceContext();
                    }
                });
                tabEl.appendChild(closeBtn);
            }
            tabEl.addEventListener('click', () => {
                if (tabsState.activeTabId !== tabId) {
                    tabsState.activeTabId = tabId;
                    saveStateToStorage();
                    switchWorkspaceContext();
                }
            });
            wrapper.appendChild(tabEl);
        });
    }

    function createNewTabScope() {
        const id = "edp_tab_" + Date.now();
        const count = Object.keys(tabsState.tabs).length + 1;
        tabsState.tabs[id] = {
            id: id,
            name: `Queue Profile ${count}`,
            routesToProcess: [],
            currentRouteIndex: 0,
            stopHistory: {},
            finalOutputData: {},
            userSkippedRoutes: []
        };
        tabsState.activeTabId = id;
        saveStateToStorage();
        switchWorkspaceContext();
    }

    function switchWorkspaceContext() {
        const activeTab = getActiveTab();
        updateTabHeadersUI();
        const inputField = document.getElementById('excel-input');
        if (inputField) {
            inputField.value = activeTab.routesToProcess.join('\n');
        }
        const runBtn = document.getElementById('start-bot-btn');
        if (runBtn) {
            if (activeTab.currentRouteIndex > 0 && activeTab.currentRouteIndex < activeTab.routesToProcess.length) {
                runBtn.textContent = "Resume Automation Engine";
                runBtn.style.background = "#ff9800";
            } else {
                runBtn.textContent = "Load & Run Automation";
                runBtn.style.background = "#0072CE";
            }
        }
        let currentStatus = "Idle";
        if (activeTab.currentRouteIndex > 0 && activeTab.currentRouteIndex < activeTab.routesToProcess.length) {
            currentStatus = "Paused (" + activeTab.currentRouteIndex + "/" + activeTab.routesToProcess.length + ")";
        } else if (activeTab.currentRouteIndex >= activeTab.routesToProcess.length && activeTab.routesToProcess.length > 0) {
            currentStatus = "Complete";
        }
        updateUIStatus(currentStatus);
        renderOutputList();
    }

    // --- Control Panel Construction ---
    function initializePanelWhenReady() {
        if (document.getElementById('edp-automation-panel')) {
            updateTabHeadersUI();
            return;
        }

        const runInterval = setInterval(() => {
            if (document.readyState === "complete" || document.readyState === "interactive") {
                clearInterval(runInterval);
                createAutomationPanel();
            }
        }, 1000);
    }

    async function createAutomationPanel() {
        const panel = document.createElement('div');
        panel.id = 'edp-automation-panel';
        panel.style = `position: fixed; top: 15px; right: 15px; width: 420px; background: #ffffff; border: 2px solid #0072CE; border-radius: 8px; z-index: 999999; padding: 14px; font-family: Arial, sans-serif; box-shadow: 0px 5px 15px rgba(0,0,0,0.2); font-size: 13px; display: flex; flex-direction: column; max-height: 85vh; user-select: none;`;
        panel.innerHTML = `
            <div id="bot-panel-header" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; padding-bottom: 4px; border-bottom: 1px solid #ddd; cursor: move;">
                <div>
                    <strong id="bot-title-text" style="color: #0072CE; font-size: 14px;">EDP Routing Engine</strong>
                    <span id="bot-status" style="margin-left: 8px; background: #e3f2fd; color: #0072CE; padding: 2px 6px; border-radius: 10px; font-size: 11px; font-weight: bold;">Idle</span>
                </div>
                <button id="bot-min-btn" style="background: none; border: none; font-size: 16px; color: #0072CE; cursor: pointer; font-weight: bold; width: 20px; height: 20px; display: flex; align-items: center; justify-content: center;">&minus;</button>
            </div>
            <div id="edp-tabs-bar" style="display: flex; align-items: flex-end; justify-content: space-between; border-bottom: 1px solid #0072CE; margin-bottom: 10px; gap: 4px;">
                <div id="edp-tabs-wrapper" style="display: flex; gap: 3px; overflow-x: auto; flex-grow: 1;"></div>
                <button id="edp-add-tab-btn" style="background: #0072CE; color: #fff; border: none; border-radius: 4px 4px 0 0; padding: 2px 8px; font-weight: bold; cursor: pointer; font-size: 12px; margin-bottom: 1px;" title="New Profile Slot">+</button>
            </div>
            <div id="bot-panel-body" style="display: flex; flex-direction: column; gap: 10px; flex-grow: 1; overflow: hidden;">
                <textarea id="excel-input" placeholder="Paste column rows from Excel here..." style="width: 100%; height: 70px; box-sizing: border-box; resize: none; border: 1px solid #ccc; border-radius: 4px; padding: 6px; font-family: monospace; font-size: 11px;"></textarea>
                <button id="start-bot-btn" style="width: 100%; background: #0072CE; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 13px; transition: background 0.2s;">Load & Run Automation</button>
                <div style="flex-grow: 1; overflow: auto; border: 1px solid #ccc; border-radius: 4px; max-height: 220px; background: #fafafa;">
                    <table id="bot-grid-table" style="width: 100%; border-collapse: collapse; font-size: 12px; text-align: left;">
                        <thead></thead>
                        <tbody id="bot-grid-body"></tbody>
                    </table>
                </div>
                <button id="copy-grid-btn" style="width: 100%; background: #2e7d32; color: white; border: none; padding: 8px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 13px; display: flex; align-items: center; justify-content: center; gap: 6px;">
                    <span>Generate Excel CSV File</span>
                </button>
            </div>
        `;
        document.body.appendChild(panel);
        const minBtn = document.getElementById('bot-min-btn');
        const bodyContent = document.getElementById('bot-panel-body');
        const tabsBar = document.getElementById('edp-tabs-bar');
        const headerEl = document.getElementById('bot-panel-header');
        const titleText = document.getElementById('bot-title-text');
        const statusEl = document.getElementById('bot-status');

        function applyMinimizeLayout(isMin) {
            if (isMin) {
                bodyContent.style.display = 'none';
                tabsBar.style.display = 'none';
                titleText.style.display = 'none';
                statusEl.style.display = 'none';
                headerEl.style.margin = '0px';
                headerEl.style.padding = '0px';
                headerEl.style.border = 'none';
                panel.style.width = '24px';
                panel.style.height = '24px';
                panel.style.minWidth = '24px';
                panel.style.minHeight = '24px';
                panel.style.padding = '0px';
                panel.style.background = '#0072CE';
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
                minBtn.style.color = '#ffffff';
                minBtn.style.display = 'flex';
                minBtn.style.alignItems = 'center';
                minBtn.style.justifyContent = 'center';
                minBtn.textContent = '⚙️';
            } else {
                bodyContent.style.display = 'flex';
                tabsBar.style.display = 'flex';
                titleText.style.display = 'inline';
                statusEl.style.display = 'inline';
                headerEl.style.marginBottom = '6px';
                panel.style.width = '420px';
                panel.style.height = 'auto';
                panel.style.padding = '14px';
                panel.style.background = '#ffffff';
                panel.style.border = '2px solid #0072CE';
                panel.style.borderRadius = '8px';
                panel.style.boxShadow = '0px 5px 15px rgba(0,0,0,0.2)';
                panel.style.opacity = '1';
                panel.style.cursor = 'default';
                minBtn.style.position = 'static';
                minBtn.style.width = 'auto';
                minBtn.style.height = 'auto';
                minBtn.style.fontSize = '16px';
                minBtn.style.color = '#0072CE';
                minBtn.textContent = '−';
            }
        }

        panel.addEventListener('mouseenter', () => {
            if (localStorage.getItem('edp_panel_minimized') === 'true') panel.style.opacity = '1';
        });
        panel.addEventListener('mouseleave', () => {
            if (localStorage.getItem('edp_panel_minimized') === 'true') panel.style.opacity = '0.25';
        });
        const isMinimized = localStorage.getItem('edp_panel_minimized') === 'true';
        if (isMinimized) applyMinimizeLayout(true);
        minBtn.addEventListener('click', (e) => {
            const currentMin = localStorage.getItem('edp_panel_minimized') === 'true';
            localStorage.setItem('edp_panel_minimized', !currentMin ? 'true' : 'false');
            applyMinimizeLayout(!currentMin);
            e.stopPropagation();
        });

        let isDragging = false;
        let offsetX, offsetY;
        headerEl.addEventListener('mousedown', (e) => {
            if (localStorage.getItem('edp_panel_minimized') === 'true' || e.target.closest('button')) return;
            isDragging = true;
            offsetX = e.clientX - panel.getBoundingClientRect().left;
            offsetY = e.clientY - panel.getBoundingClientRect().top;
            panel.style.right = 'auto';
            headerEl.style.cursor = 'grabbing';
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
                headerEl.style.cursor = 'move';
            }
        });

        document.getElementById('edp-add-tab-btn').addEventListener('click', createNewTabScope);
        document.getElementById('start-bot-btn').addEventListener('click', () => {
            const activeTab = getActiveTab();
            const rawText = document.getElementById('excel-input').value;
            const parsedRoutes = parseExcelData(rawText);
            if (parsedRoutes.length === 0) {
                alert("No valid route codes extracted.");
                return;
            }
            activeTab.finalOutputData = {};
            activeTab.stopHistory = {};
            activeTab.routesToProcess = parsedRoutes;
            activeTab.currentRouteIndex = 0;
            saveStateToStorage();
            switchWorkspaceContext();
            processNextRoute();
        });

        document.getElementById('copy-grid-btn').addEventListener('click', () => {
            const activeTab = getActiveTab();
            if (!activeTab.routesToProcess || activeTab.routesToProcess.length === 0) {
                alert("No route data loaded.");
                return;
            }
            let maxCols = 0;
            activeTab.routesToProcess.forEach(route => {
                const rawLine = activeTab.finalOutputData[route];
                if (rawLine && rawLine.includes('\t')) {
                    const len = rawLine.split('\t').slice(1).length;
                    if (len > maxCols) maxCols = len;
                }
            });
            let csvRows = [];
            activeTab.routesToProcess.forEach(route => {
                let cellData = [];
                for (let i = 0; i < 10; i++) cellData.push("");
                cellData.push(`"${route}"`);
                cellData.push("");
                cellData.push("");
                const rawLine = activeTab.finalOutputData[route];
                if (rawLine && rawLine.includes('\t')) {
                    const parts = rawLine.split('\t').slice(1);
                    for (let i = 0; i < maxCols; i++) {
                        const status = parts[i] ? parts[i].trim().toUpperCase() : "";
                        cellData.push(status ? `"${transMap[status] || status}"` : "");
                    }
                } else {
                    for (let i = 0; i < maxCols; i++) cellData.push("");
                }
                csvRows.push(cellData.join(','));
            });
            const csvContent = "\uFEFF" + csvRows.join('\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const dateStr = new Date().toISOString().substring(0,10);
            link.href = URL.createObjectURL(blob);
            link.setAttribute("download", activeTab.name.replace(/[^a-z0-9]/gi, '_') + "Export" + dateStr + ".csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });

        switchWorkspaceContext();
    }

    initializePanelWhenReady();
})();
