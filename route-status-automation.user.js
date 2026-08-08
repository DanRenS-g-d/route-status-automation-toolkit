// ==UserScript==
// @name         Route Status Automation (Sanitized Demo)
// @namespace    http://tampermonkey.net
// @version      8.1
// @description  Automates tracking route stop statuses on a delivery-routing web portal with local CSV export pipelines. Portfolio/demo build — @match and the constants below must be pointed at your own target system before use.
// @author       Danilo Jose Rengifo Sulbaran (check GitHub Repo)
// @match        https://your-routing-portal.example.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

// NOTE (sanitized build): the line above is a placeholder. Point @match at
// whatever internal routing portal you're automating, and update
// ROUTE_ID_PATTERN / STORE_KEYWORD_REGEX below to match your own naming
// conventions before running this against a real site.

(function() {
    'use strict';

    let routesToProcess = [];
    let currentRouteIndex = 0;
    let stopHistory = {};
    let finalOutputData = {};
    let timerCountdown = 15 * 60;
    let countdownInterval = null;
    let masterRosterData = {};

    // Valid statuses that don't need a manual check fallback
    const expectedStatuses = ["DELIVERED", "IN TRANSIT", "OUT FOR DELIVERY", "NARROWED", "ARRIVED", "LOADED", "NOT LOADED"];

    // Customize with your own retailer/company keyword (see Store Stop Detection below)
    const STORE_KEYWORD_REGEX = /\bYourStoreName\b/i;

    // Restore working data safely across automated page refreshes
    try {
        const savedRoster = localStorage.getItem('edp_master_roster');
        if (savedRoster) masterRosterData = JSON.parse(savedRoster);

        const savedRoutes = localStorage.getItem('edp_routes_list');
        if (savedRoutes) routesToProcess = JSON.parse(savedRoutes);

        const savedOutput = localStorage.getItem('edp_final_output_raw');
        if (savedOutput) finalOutputData = JSON.parse(savedOutput);
    } catch(e) {
        console.error("Failed to rehydrate working loop context:", e);
    }

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

    // --- DOM Utilities ---
    function waitForElement(selector, timeout = 10000) {
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

    function setNativeValue(element, value) {
        const lastValue = element.value;
        element.value = value;
        const event = new Event('input', { bubbles: true });
        event.simulated = true;
        const tracker = element._valueTracker;
        if (tracker) tracker.setValue(lastValue);
        element.dispatchEvent(event);
    }

    // Customize this to match your own route-ID naming convention
    // (e.g. depot/truck code prefixes). Example below matches IDs like
    // "0421RTE01" or "0421TRK7".
    const ROUTE_ID_PATTERN = /\b\d+(?:RTE|TRK)\d*\b/i;

    function parseExcelData(text) {
        if (!text.trim()) return [];
        const lines = text.split(/[\r\n]+/);
        const parsedRoutes = [];
        lines.forEach(line => {
            const match = line.match(ROUTE_ID_PATTERN);
            if (match) parsedRoutes.push(match[0].toUpperCase().trim());
        });
        return parsedRoutes;
    }

    // --- Core Automation Loop ---
    async function processNextRoute() {
        if (currentRouteIndex >= routesToProcess.length) {
            updateUIStatus("Complete");
            startHibernationTimer();
            return;
        }

        const fullRoute = routesToProcess[currentRouteIndex];
        updateUIStatus(`Checking ${currentRouteIndex + 1}/${routesToProcess.length}`);

        try {
            // Step 3: Location Switcher Configuration
            let numbersPart = fullRoute.match(/^\d+/);
            if (numbersPart) {
                numbersPart = numbersPart[0];
                if (numbersPart.length === 3) numbersPart = "0" + numbersPart;
                else if (numbersPart.length > 4) numbersPart = numbersPart.substring(0, 4);
            }

            const locBtn = await waitForElement('button[data-testid="location-button"]');
            locBtn.click();

            const editBtn = await waitForElement('button[data-testid="shadow-location-edit"]');
            editBtn.click();

            const inputField = await waitForElement('input[role="combobox"]');
            setNativeValue(inputField, numbersPart);

            // Trigger UI updates for combobox input
            ['keydown', 'keypress', 'keyup'].forEach(type => {
                inputField.dispatchEvent(new KeyboardEvent(type, { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true }));
            });

            // Step 4: Wait 10 seconds, then search for and dismiss the modal 'Close' button
            await new Promise(r => setTimeout(r, 10000));
            const closeBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === 'Close');
            if (closeBtn) closeBtn.click();

            // Step 5 & 6: Wait 5 seconds for background panel rendering
            await new Promise(r => setTimeout(r, 5000));

            // Step 7, 8, 9: Locate the designated Route card block in the visibility panel
            const routeCards = Array.from(document.querySelectorAll('[data-testid^="route-card-"]'));
            const targetCard = routeCards.find(card => {
                const header = card.querySelector('h6');
                return header && header.textContent.trim() === fullRoute;
            });

            if (!targetCard) {
                finalOutputData[fullRoute] = `${fullRoute}\tROUTE NOT FOUND`;
                renderOutputList();
                nextIteration();
                return;
            }

            // Click the card body to populate the stop layout lists
            targetCard.click();
            await new Promise(r => setTimeout(r, 2000));

            // Step 10 & 11: Locate and extract information from the dynamic scroll container box matching this route
            const stopContainer = document.querySelector(`[data-rbd-droppable-id="${fullRoute}"]`);
            if (!stopContainer) {
                finalOutputData[fullRoute] = `${fullRoute}\tCONTAINER NOT FOUND`;
                renderOutputList();
                nextIteration();
                return;
            }

            const stopCards = Array.from(stopContainer.querySelectorAll('[data-testid="stop-card"]'));
            let currentRunStopIds = [];

            // Map validation data explicitly using the blue circle indicator numbers
            let stopDataMap = {};
            let highestStopNum = 0;

            stopCards.forEach((card) => {
                const orderLink = card.querySelector('[data-testid="order-id-link"]');
                const orderId = orderLink ? orderLink.textContent.trim() : "UNKNOWN_ID";
                currentRunStopIds.push(orderId);

                let stopNum = 0;
                const rowContainer = card.closest('[style*="flex-direction: row"]') || card.parentElement;
                if (rowContainer) {
                    const possibleBubble = rowContainer.querySelector('div > div') || rowContainer.firstElementChild;
                    if (possibleBubble && /^\d+$/.test(possibleBubble.textContent.trim())) {
                        stopNum = parseInt(possibleBubble.textContent.trim(), 10);
                    }
                }

                if (!stopNum || isNaN(stopNum)) {
                    const textMatch = card.parentElement?.innerText.match(/^(\d+)/);
                    if (textMatch) stopNum = parseInt(textMatch, 10);
                }

                if (stopNum > highestStopNum) highestStopNum = stopNum;

                // --- Store Stop Detection ---
                // Flags a stop as a "store stop" (vs. a customer delivery) when the
                // location text contains your retailer's own name/keyword. Customize
                // STORE_KEYWORD_REGEX for your organization.
                const locationParagraph = card.querySelector('p[aria-label], .css-1rlsv84');
                const fullLocationText = locationParagraph ?
                    (locationParagraph.getAttribute('aria-label') || locationParagraph.textContent || "") : "";

                // Only matches the configured keyword as a distinct word boundary
                let isStoreSaleStop = STORE_KEYWORD_REGEX.test(fullLocationText);

                // --- Re-engineered Arrived Time Color Evaluation ---
                let detectedTimeColor = null; // Can be "yellow", "red", or null

                // Grabs every typography node across the entire card layout block
                const allCardSpans = Array.from(card.querySelectorAll('span, p, div'));

                // Matches standard layout time signatures like 6:15 AM or 12:30 PM
                const timePattern = /\b\d{1,2}:\d{2}\s*(?:AM|PM)\b/i;

                for (let el of allCardSpans) {
                    const textStr = el.textContent.trim();

                    // Verify if this specific layout element contains a timestamp string
                    if (timePattern.test(textStr)) {
                        const computedColor = window.getComputedStyle(el).color.replace(/\s+/g, '');

                        if (computedColor === 'rgb(255,212,0)') {
                            detectedTimeColor = "yellow";
                            break;
                        } else if (computedColor === 'rgb(255,0,0)') {
                            detectedTimeColor = "red";
                            break;
                        }
                    }
                }

                const chipLabelEl = card.querySelector('.MuiChip-labelSmall');
                const rawStatus = chipLabelEl ? chipLabelEl.textContent.trim().toUpperCase() : "";

                let cleanStatus = "MC";

                // Hierarchy execution rules assignment
                if (isStoreSaleStop) {
                    cleanStatus = "SS";
                } else if (rawStatus === "DELIVERED" && detectedTimeColor === "yellow") {
                    cleanStatus = "DE"; // Yellow arrival time = Delivered Early
                } else if (rawStatus === "DELIVERED" && detectedTimeColor === "red") {
                    cleanStatus = "DL"; // Red arrival time = Delivered Late
                } else if (expectedStatuses.includes(rawStatus)) {
                    cleanStatus = transMap[rawStatus] || rawStatus;
                }

                if (stopNum > 0) {
                    stopDataMap[stopNum] = cleanStatus;
                }
            });

            // Reconstruct row parts sequentially. Missing index positions default to DEL
            let routeRowParts = [fullRoute];
            for (let i = 1; i <= highestStopNum; i++) {
                if (stopDataMap[i]) {
                    routeRowParts.push(stopDataMap[i]);
                } else {
                    routeRowParts.push("DEL");
                }
            }

            // Step 13: Evaluate missing steps against historical entries for tracking deletions (DEL)
            if (stopHistory[fullRoute]) {
                const previousRunStopIds = stopHistory[fullRoute];
                let checkedRowParts = [fullRoute];

                stopCards.forEach((card, index) => {
                    const orderLink = card.querySelector('[data-testid="order-id-link"]');
                    const orderId = orderLink ? orderLink.textContent.trim() : "";

                    if (previousRunStopIds[index] && previousRunStopIds[index] !== orderId && !currentRunStopIds.includes(previousRunStopIds[index])) {
                        checkedRowParts.push("DEL");
                    }

                    const chipLabelEl = card.querySelector('.MuiChip-labelSmall');
                    const rawStatus = chipLabelEl ? chipLabelEl.textContent.trim().toUpperCase() : "";
                    let cleanStatus = expectedStatuses.includes(rawStatus) ? (transMap[rawStatus] || rawStatus) : "MC";
                    checkedRowParts.push(cleanStatus);
                });

                if (routeRowParts.length <= checkedRowParts.length) {
                    routeRowParts = checkedRowParts;
                }
            }

            // Save state context for next sequence comparison
            stopHistory[fullRoute] = currentRunStopIds;
            finalOutputData[fullRoute] = routeRowParts.join('\t');
            renderOutputList();

        } catch (error) {
            console.error(error);
            finalOutputData[fullRoute] = `${fullRoute}\tERROR: ${error.message}`;
            renderOutputList();
        }

        nextIteration();
    }

    function nextIteration() {
        currentRouteIndex++;
        setTimeout(processNextRoute, 1500);
    }

    function startHibernationTimer() {
        timerCountdown = 15 * 60;
        clearInterval(countdownInterval);

        countdownInterval = setInterval(() => {
            timerCountdown--;
            const mins = Math.floor(timerCountdown / 60);
            const secs = timerCountdown % 60;
            updateUIStatus(`Standby (${mins}:${secs < 10 ? '0' : ''}${secs})`);

            if (timerCountdown <= 0) {
                clearInterval(countdownInterval);
                window.location.reload();
            }
        }, 1000);
    }

    // --- Grid Rendering Logic with Character Translation Mapping ---
    function renderOutputList() {
        const tbody = document.getElementById('bot-grid-body');
        const table = document.getElementById('bot-grid-table');
        if (!tbody || !table) return;

        if (!routesToProcess || routesToProcess.length === 0) {
            tbody.innerHTML = `<tr><td colspan="2" style="padding: 10px; color: #888; text-align: center;">No data loaded yet.</td></tr>`;
            return;
        }

        let translatedOutputData = {};
        let maxStopsCount = 0;

        routesToProcess.forEach(route => {
            if (!route) return;
            const rawLine = finalOutputData[route];

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
        localStorage.setItem('edp_final_output_raw', JSON.stringify(finalOutputData || {}));
        localStorage.setItem('edp_routes_list', JSON.stringify(routesToProcess || []));

        let headerHtml = `
            <tr style="background: #002060; color: #ffffff; border-bottom: 2px solid #ccc; position: sticky; top: 0; z-index: 10;">
                <th style="padding: 6px; border-right: 1px solid #ddd; min-width: 120px;">Route:</th>
        `;
        for (let i = 1; i <= Math.max(maxStopsCount, 1); i++) {
            headerHtml += `<th style="padding: 6px; border-right: 1px solid #ddd; text-align: center; min-width: 40px;">${i}</th>`;
        }
        headerHtml += `</tr>`;
        table.querySelector('thead').innerHTML = headerHtml;

        let html = "";
        routesToProcess.forEach(route => {
            if (!route) return;
            const dataRow = translatedOutputData[route];
            let rowHtml = `<tr style="background: #fff;">`;
            rowHtml += `<td style="padding: 6px; border-right: 1px solid #ddd; font-weight: bold; background: #fff; position: sticky; left: 0; border-bottom: 1px solid #eee;">${route}</td>`;

            const segments = (dataRow && dataRow.includes('\t')) ? dataRow.split('\t').slice(1) : (masterRosterData[route] || []);

            if (segments && segments.length > 0) {
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

                        rowHtml += `
                            <td style="padding: 4px; border-right: 1px solid #eee; border-bottom: 1px solid #eee; text-align: center;">
                                <span style="background: ${bg}; color: ${color}; padding: 2px 6px; border-radius: 3px; border: 1px solid rgba(0,0,0,0.05); display: inline-block; font-weight: bold; font-size: 11px; min-width: 24px;">${shortCode}</span>
                            </td>`;
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
    }

    // --- Updated Overlay UI Generation with Invisible Discrete Mode ---
    function createAutomationPanel() {
        const panel = document.createElement('div');
        panel.id = 'route-automation-panel';
        panel.style = `
            position: fixed; top: 15px; right: 15px; width: 420px;
            background: #ffffff; border: 2px solid #0072CE; border-radius: 8px;
            z-index: 999999; padding: 14px; font-family: Arial, sans-serif;
            box-shadow: 0px 5px 15px rgba(0,0,0,0.2); font-size: 13px;
            display: flex; flex-direction: column; max-height: 85vh;
            user-select: none;
        `;

        panel.innerHTML = `
            <div id="bot-panel-header" style="font-weight: bold; margin-bottom: 10px; color: #0072CE; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #eee; padding-bottom: 6px; flex-shrink: 0; cursor: move; width: 100%;">
                <span id="bot-title-text">EDP Routing Engine</span>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span id="bot-status" style="font-weight: bold; color: #ff9800; font-size: 11px;">Idle</span>
                    <button id="bot-min-btn" style="background: none; border: none; font-weight: bold; font-size: 16px; cursor: pointer; color: #0072CE; padding: 0 4px; line-height: 10px; outline: none;">−</button>
                </div>
            </div>
            <div id="bot-panel-body" style="display: flex; flex-direction: column; gap: 8px; overflow: hidden; height: auto;">
                <textarea id="excel-input" placeholder="Paste columns straight from Excel here..." style="width: 100%; height: 60px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; padding: 6px; resize: none; flex-shrink: 0; user-select: text;"></textarea>
                <button id="start-bot-btn" style="width: 100%; background: #0072CE; color: white; border: none; padding: 8px; border-radius: 4px; cursor: pointer; font-weight: bold;">Load & Run Automation</button>

                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div style="font-weight: bold; color: #333;">Live Route Grid:</div>
                    <button id="copy-grid-btn" style="background: #2e7d32; color: white; border: none; padding: 4px 10px; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 11px;">Generate Excel CSV File</button>
                </div>

                <div style="width: 100%; overflow: auto; border: 1px solid #ccc; border-radius: 4px; background: #fafafa; flex-grow: 1; max-height: 40vh;">
                    <table id="bot-grid-table" style="width: 100%; border-collapse: collapse; font-family: monospace; font-size: 11px; text-align: left;">
                        <thead>
                            <tr style="background: #f0f0f0; border-bottom: 2px solid #ccc; position: sticky; top: 0; z-index: 10;">
                                <th style="padding: 6px; border-right: 1px solid #ddd;">Route</th>
                                <th style="padding: 6px;">Stops Statuses</th>
                            </tr>
                        </thead>
                        <tbody id="bot-grid-body">
                            <tr>
                                <td colspan="2" style="padding: 10px; color: #888; text-align: center;">No data loaded yet.</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        `;

        document.body.appendChild(panel);

        const minBtn = document.getElementById('bot-min-btn');
        const bodyContent = document.getElementById('bot-panel-body');
        const headerEl = document.getElementById('bot-panel-header');
        const titleText = document.getElementById('bot-title-text');
        const statusEl = document.getElementById('bot-status');

        function applyMinimizeLayout(isMin) {
            if (isMin) {
                bodyContent.style.display = 'none';
                titleText.style.display = 'none';
                statusEl.style.display = 'none';
                headerEl.style.margin = '0px';
                headerEl.style.padding = '0px';
                headerEl.style.border = 'none';

                // Visible, tactical mini-pill layout transformation
                panel.style.width = '24px';
                panel.style.height = '24px';
                panel.style.minWidth = '24px';
                panel.style.minHeight = '24px';
                panel.style.padding = '0px';
                panel.style.background = '#0072CE';
                panel.style.border = '2px solid #ffffff';
                panel.style.borderRadius = '50%';
                panel.style.boxShadow = '0px 2px 8px rgba(0,0,0,0.3)';
                panel.style.opacity = '0.25'; // Noticeable base visibility
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
                minBtn.textContent = '⛶'; // Maximize icon hint symbol
            } else {
                bodyContent.style.display = 'flex';
                titleText.style.display = 'inline';
                statusEl.style.display = 'inline';
                headerEl.style.marginBottom = '10px';
                headerEl.style.paddingBottom = '6px';
                headerEl.style.borderBottom = '1px solid #eee';

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

        // Snap immediately to full crisp visibility when hovering over it
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
            if (currentMin) {
                localStorage.setItem('edp_panel_minimized', 'false');
                applyMinimizeLayout(false);
            } else {
                localStorage.setItem('edp_panel_minimized', 'true');
                applyMinimizeLayout(true);
            }
            e.stopPropagation();
        });

        // Maintain core functional listeners
        const header = document.getElementById('bot-panel-header');
        let isDragging = false; let offsetX, offsetY;
        header.addEventListener('mousedown', (e) => {
            if (localStorage.getItem('edp_panel_minimized') === 'true' || e.target.closest('button')) return;
            isDragging = true; offsetX = e.clientX - panel.getBoundingClientRect().left; offsetY = e.clientY - panel.getBoundingClientRect().top;
            panel.style.right = 'auto'; header.style.cursor = 'grabbing';
        });
        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            let leftPosition = e.clientX - offsetX; let topPosition = e.clientY - offsetY;
            leftPosition = Math.max(0, Math.min(leftPosition, window.innerWidth - panel.offsetWidth));
            topPosition = Math.max(0, Math.min(topPosition, window.innerHeight - panel.offsetHeight));
            panel.style.left = leftPosition + 'px'; panel.style.top = topPosition + 'px';
        });
        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                header.style.cursor = 'move';
            }
        });

        if (routesToProcess && routesToProcess.length > 0) {
            document.getElementById('excel-input').value = routesToProcess.join('\n');
            renderOutputList();
        }

        document.getElementById('start-bot-btn').addEventListener('click', () => {
            const rawText = document.getElementById('excel-input').value;
            const parsedRoutes = parseExcelData(rawText);

            if (parsedRoutes.length === 0) {
                alert("No valid route codes extracted.");
                return;
            }

            localStorage.removeItem('edp_master_roster');
            localStorage.removeItem('edp_final_output_raw');
            localStorage.removeItem('edp_routes_list');

            masterRosterData = {};
            finalOutputData = {};
            stopHistory = {};

            parsedRoutes.forEach(route => {
                masterRosterData[route] = [];
            });

            routesToProcess = parsedRoutes;
            clearInterval(countdownInterval);
            currentRouteIndex = 0;
            renderOutputList();
            processNextRoute();
        });

        document.getElementById('copy-grid-btn').addEventListener('click', () => {
            if (!routesToProcess || routesToProcess.length === 0) {
                alert("No route data loaded.");
                return;
            }

            let maxCols = 0;
            routesToProcess.forEach(route => {
                const rawLine = finalOutputData[route];
                if (rawLine && rawLine.includes('\t')) {
                    const len = rawLine.split('\t').slice(1).length;
                    if (len > maxCols) maxCols = len;
                }
            });

            let csvRows = [];
            routesToProcess.forEach(route => {
                let cellData = [];
                for (let i = 0; i < 10; i++) cellData.push("");
                cellData.push(`"${route}"`);
                cellData.push("");
                cellData.push("");

                const rawLine = finalOutputData[route];
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
            link.setAttribute("download", `EDP_Route_Export_${dateStr}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    }

    function updateUIStatus(statusText) {
        const statusEl = document.getElementById('bot-status');
        if (statusEl) statusEl.textContent = statusText;
    }

    if (document.readyState === "complete" || document.readyState === "interactive") {
        createAutomationPanel();
    } else {
        window.addEventListener('DOMContentLoaded', createAutomationPanel);
    }
})();