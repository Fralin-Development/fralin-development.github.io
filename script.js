/**
 * Fralin Development - Live Donor Roll
 * Supports:
 * 1. Google Sheets (keyless live CSV export)
 * 2. SharePoint / OneDrive / Excel Online (.xlsx live download & parsing via SheetJS)
 * 3. Local CSV fallback (./donors.csv)
 */

// =============================================================================
// CONFIGURATION
// =============================================================================
const CONFIG = {
    // 1. Paste your full Google Sheet URL here (any share link, view link, or tab link with #gid=...)
    // Example: "https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0"
    sheetUrl: "https://myuva-my.sharepoint.com/:x:/g/personal/dbg5af_virginia_edu/IQBDrHqx3UG3QJppnFzTHoMGATxyrU3UWV3-Irso_3Nirt8?e=wVAlnf",

    // 2. Specific Excel sheet / tab name to read from (e.g., "Donation Screen", "Acks - Edited")
    sheetTabName: "Donation Screen",

    // 3. Local fallback CSV path in case external browser CORS blocks direct download
    fallbackCsvUrl: "./donors.csv",

    // 4. Titles displayed at the top of the scrolling list
    headerTitle: "Thank you to our generous donors who make <em>art together</em> possible",
    subHeaderTitle: "", // Optional subtitle (e.g., "Annual Campaign 2026")

    // 5. Polling interval: How often (in seconds) to check for updates
    refreshIntervalSeconds: 30,

    // 6. Scroll speed in pixels per second (e.g., 35 = slow/relaxed, 50 = standard, 70 = brisk)
    scrollSpeedPixelsPerSecond: 45,

    // 7. Column mappings: Header names (case-insensitive) or 0-based column indexes
    // If set to null or not matched, it auto-detects or defaults to Column A for names.
    columns: {
        name: "Name",       // Look for header with "name" or "donor", or column 0
        amount: "Amount",   // Optional: header with "amount" or "gift"
        message: "Message"  // Optional: header with "message", "note", or "dedication"
    },

    // 8. Show status indicator in top-right corner
    showStatusBadge: true
};

// Fallback sample data displayed when no data source is reachable
const SAMPLE_DONORS = [
    { name: "The Harrison Family Foundation", amount: "$50,000", message: "In honor of Mary & Robert Harrison" },
    { name: "Dr. Arthur & Eleanor Vance", amount: "$25,000", message: "Supporting future generations of leaders" },
    { name: "Commonwealth Trust & Partners", amount: "$15,000", message: "Proud supporters of the Fralin Mission" },
    { name: "Sarah & David Montgomery", amount: "$10,000", message: "" },
    { name: "The Sterling Memorial Fund", amount: "$10,000", message: "In memory of Thomas Sterling, Class of '72" },
    { name: "Dr. Rachel Chen", amount: "$5,000", message: "With deep gratitude" },
    { name: "Marcus & Elena Bennett", amount: "$5,000", message: "" }
];

// =============================================================================
// DOM ELEMENTS & STATE
// =============================================================================
const donorListElement = document.getElementById("donor-list");
const donorItemsContainer = document.getElementById("donor-items");
const mainHeaderElement = document.getElementById("main-header");
const subHeaderElement = document.getElementById("sub-header");
const scrollContainer = document.getElementById("scroll-container");
const statusBadge = document.getElementById("status-badge");
const statusText = document.getElementById("status-text");
const errorBanner = document.getElementById("error-banner");
const errorMessage = document.getElementById("error-message");

let lastDataSignature = "";
let isFetching = false;

// =============================================================================
// URL PARSER & CONVERTERS
// =============================================================================

function isSharePointOrExcelUrl(url) {
    if (!url) return false;
    return url.includes("sharepoint.com") || url.includes("onedrive.live.com") || url.includes(".xlsx");
}

function getSharePointDownloadUrl(url) {
    if (!url) return null;
    const cleanUrl = url.split("?")[0];
    return `${cleanUrl}?download=1`;
}

function getGoogleSheetCsvUrl(rawUrl) {
    if (!rawUrl || typeof rawUrl !== "string") return null;
    const trimmed = rawUrl.trim();
    if (!trimmed) return null;

    if (trimmed.includes("output=csv") || (trimmed.includes("export?format=csv") && !trimmed.includes("/edit"))) {
        return trimmed;
    }

    const pubMatch = trimmed.match(/\/spreadsheets\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (pubMatch) {
        const docId = pubMatch[1];
        const gidMatch = trimmed.match(/[?&#]gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : "0";
        return `https://docs.google.com/spreadsheets/d/e/${docId}/pub?output=csv&gid=${gid}`;
    }

    const idMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) {
        const docId = idMatch[1];
        const gidMatch = trimmed.match(/[?&#]gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : "0";
        return `https://docs.google.com/spreadsheets/d/${docId}/export?format=csv&gid=${gid}`;
    }

    return trimmed;
}

// =============================================================================
// CSV & EXCEL PARSERS
// =============================================================================

function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let currentCell = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];

        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                currentCell += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === "," && !inQuotes) {
            currentRow.push(currentCell.trim());
            currentCell = "";
        } else if ((char === "\r" || char === "\n") && !inQuotes) {
            if (char === "\r" && nextChar === "\n") {
                i++;
            }
            currentRow.push(currentCell.trim());
            if (currentRow.some(cell => cell.length > 0)) {
                rows.push(currentRow);
            }
            currentRow = [];
            currentCell = "";
        } else {
            currentCell += char;
        }
    }

    if (currentCell.length > 0 || currentRow.length > 0) {
        currentRow.push(currentCell.trim());
        if (currentRow.some(cell => cell.length > 0)) {
            rows.push(currentRow);
        }
    }

    return rows;
}

function parseExcelBuffer(arrayBuffer, preferredTabName = "Donation Screen") {
    if (typeof XLSX === "undefined") {
        throw new Error("SheetJS (XLSX) library is not loaded.");
    }

    const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
    
    // Find target sheet tab
    let targetSheetName = workbook.SheetNames[0];
    if (preferredTabName) {
        const found = workbook.SheetNames.find(s => s.toLowerCase().trim() === preferredTabName.toLowerCase().trim())
                   || workbook.SheetNames.find(s => s.toLowerCase().includes(preferredTabName.toLowerCase()));
        if (found) targetSheetName = found;
    }

    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet) return [];

    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    return rawRows.map(row => row.map(cell => String(cell || "").trim()));
}

function mapRowsToDonors(rows) {
    if (!rows || rows.length === 0) return [];

    const firstRow = rows[0].map(c => String(c).toLowerCase().trim());
    let hasHeader = false;
    let nameIdx = 0;
    let amountIdx = -1;
    let messageIdx = -1;

    // Detect header row keywords
    const potentialHeaderKeywords = ["name", "donor", "contributor", "supporter", "amount", "gift", "message", "note", "dedication", "tier"];
    const containsHeaderKeyword = firstRow.some(cell => potentialHeaderKeywords.some(kw => cell.includes(kw)));

    if (containsHeaderKeyword) {
        hasHeader = true;

        const nameQuery = typeof CONFIG.columns.name === "string" ? CONFIG.columns.name.toLowerCase() : "name";
        const foundNameIdx = firstRow.findIndex(c => c.includes(nameQuery) || c.includes("donor") || c.includes("contributor"));
        if (foundNameIdx !== -1) nameIdx = foundNameIdx;

        const amountQuery = typeof CONFIG.columns.amount === "string" ? CONFIG.columns.amount.toLowerCase() : "amount";
        const foundAmountIdx = firstRow.findIndex(c => c.includes(amountQuery) || c.includes("gift") || c.includes("donation") || c.includes("tier"));
        if (foundAmountIdx !== -1) amountIdx = foundAmountIdx;

        const msgQuery = typeof CONFIG.columns.message === "string" ? CONFIG.columns.message.toLowerCase() : "message";
        const foundMsgIdx = firstRow.findIndex(c => c.includes(msgQuery) || c.includes("note") || c.includes("dedication") || c.includes("memory") || c.includes("honor"));
        if (foundMsgIdx !== -1) messageIdx = foundMsgIdx;
    } else {
        if (typeof CONFIG.columns.name === "number") nameIdx = CONFIG.columns.name;
        if (typeof CONFIG.columns.amount === "number") amountIdx = CONFIG.columns.amount;
        if (typeof CONFIG.columns.message === "number") messageIdx = CONFIG.columns.message;
    }

    const dataRows = hasHeader ? rows.slice(1) : rows;

    const donors = [];
    for (const row of dataRows) {
        const name = row[nameIdx] ? String(row[nameIdx]).trim() : "";
        if (!name) continue;

        const amount = (amountIdx >= 0 && row[amountIdx]) ? String(row[amountIdx]).trim() : "";
        const message = (messageIdx >= 0 && row[messageIdx]) ? String(row[messageIdx]).trim() : "";

        donors.push({ name, amount, message });
    }

    return donors;
}

// =============================================================================
// RENDERING & SCROLL TIMING
// =============================================================================

function updateScrollDuration() {
    if (!donorListElement || !scrollContainer) return;

    requestAnimationFrame(() => {
        const contentHeight = donorListElement.offsetHeight;
        const containerHeight = scrollContainer.offsetHeight || window.innerHeight;
        const totalTravelDistance = contentHeight + containerHeight;
        const speed = Math.max(CONFIG.scrollSpeedPixelsPerSecond || 45, 10);
        const durationSeconds = Math.max(totalTravelDistance / speed, 12);

        donorListElement.style.animationDuration = `${durationSeconds.toFixed(1)}s`;
    });
}

function renderDonors(donors) {
    if (mainHeaderElement) mainHeaderElement.innerHTML = CONFIG.headerTitle;
    if (subHeaderElement) {
        subHeaderElement.textContent = CONFIG.subHeaderTitle;
        subHeaderElement.style.display = CONFIG.subHeaderTitle ? "block" : "none";
    }

    if (!donorItemsContainer) return;
    donorItemsContainer.innerHTML = "";

    if (!donors || donors.length === 0) {
        donorItemsContainer.innerHTML = `
            <div class="donor-card">
                <div class="donor-name">No donor records found</div>
            </div>
        `;
        updateScrollDuration();
        return;
    }

    const fragment = document.createDocumentFragment();

    donors.forEach(donor => {
        const card = document.createElement("div");
        card.className = "donor-card";

        const nameEl = document.createElement("div");
        nameEl.className = "donor-name";
        nameEl.textContent = donor.name;
        card.appendChild(nameEl);

        if (donor.amount) {
            const amountEl = document.createElement("div");
            amountEl.className = "donor-amount";
            amountEl.textContent = donor.amount;
            card.appendChild(amountEl);
        }

        if (donor.message) {
            const msgEl = document.createElement("div");
            msgEl.className = "donor-message";
            msgEl.textContent = `“${donor.message}”`;
            card.appendChild(msgEl);
        }

        fragment.appendChild(card);
    });

    donorItemsContainer.appendChild(fragment);
    updateScrollDuration();
}

// =============================================================================
// DATA FETCHING & LIVE SYNC
// =============================================================================

function showBannerError(msg) {
    if (errorBanner) {
        if (errorMessage) errorMessage.innerHTML = msg;
        errorBanner.classList.remove("hidden");
    }
}

function hideBannerError() {
    if (errorBanner) {
        errorBanner.classList.add("hidden");
    }
}

function updateStatus(text) {
    if (!CONFIG.showStatusBadge || !statusBadge) return;
    statusBadge.classList.remove("hidden");
    if (statusText) {
        statusText.textContent = text;
    }
}

async function fetchFromFallbackCsv() {
    if (!CONFIG.fallbackCsvUrl) return false;
    try {
        const res = await fetch(`${CONFIG.fallbackCsvUrl}?_t=${Date.now()}`);
        if (!res.ok) return false;
        const text = await res.text();
        const rows = parseCSV(text);
        const donors = mapRowsToDonors(rows);
        if (donors.length > 0) {
            const signature = JSON.stringify(donors);
            if (signature !== lastDataSignature) {
                lastDataSignature = signature;
                renderDonors(donors);
            }
            hideBannerError();
            updateStatus("Donors Loaded (CSV)");
            return true;
        }
    } catch (e) {
        console.warn("Fallback CSV fetch failed:", e);
    }
    return false;
}

async function fetchDonorData() {
    if (isFetching) return;
    isFetching = true;

    const rawUrl = CONFIG.sheetUrl;

    if (!rawUrl) {
        const loadedFromCsv = await fetchFromFallbackCsv();
        if (!loadedFromCsv) {
            renderDonors(SAMPLE_DONORS);
            updateStatus("Demo Mode");
        }
        isFetching = false;
        return;
    }

    try {
        if (isSharePointOrExcelUrl(rawUrl)) {
            // Fetch Excel file (.xlsx)
            const downloadUrl = getSharePointDownloadUrl(rawUrl);
            const cacheBustedUrl = `${downloadUrl}&_t=${Date.now()}`;
            
            const response = await fetch(cacheBustedUrl);
            if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);

            const arrayBuffer = await response.arrayBuffer();
            const rows = parseExcelBuffer(arrayBuffer, CONFIG.sheetTabName);
            const donors = mapRowsToDonors(rows);

            if (donors.length > 0) {
                const signature = JSON.stringify(donors);
                if (signature !== lastDataSignature) {
                    lastDataSignature = signature;
                    renderDonors(donors);
                }
                hideBannerError();
                const now = new Date();
                const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
                updateStatus(`SharePoint Live (${timeStr})`);
            } else {
                throw new Error(`No donors found in tab "${CONFIG.sheetTabName}".`);
            }
        } else {
            // Fetch Google Sheets CSV
            const csvUrl = getGoogleSheetCsvUrl(rawUrl);
            const cacheBustedUrl = csvUrl.includes("?") 
                ? `${csvUrl}&_t=${Date.now()}` 
                : `${csvUrl}?_t=${Date.now()}`;

            const response = await fetch(cacheBustedUrl, {
                headers: { "Accept": "text/csv, text/plain, */*" }
            });

            if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);

            const csvText = await response.text();
            if (csvText.includes("<!DOCTYPE html>") || csvText.includes("<html")) {
                throw new Error("Received HTML login page instead of CSV.");
            }

            const rows = parseCSV(csvText);
            const donors = mapRowsToDonors(rows);

            const signature = JSON.stringify(donors);
            if (signature !== lastDataSignature) {
                lastDataSignature = signature;
                renderDonors(donors);
            }

            hideBannerError();
            const now = new Date();
            const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
            updateStatus(`Live Sync (${timeStr})`);
        }
    } catch (error) {
        console.warn("Direct live fetch error (falling back to donors.csv):", error);
        const loadedFromCsv = await fetchFromFallbackCsv();
        if (!loadedFromCsv && !lastDataSignature) {
            renderDonors(SAMPLE_DONORS);
            showBannerError(
                `Unable to connect to live document.<br>Displaying sample donor data.`
            );
        }
    } finally {
        isFetching = false;
    }
}

// =============================================================================
// INITIALIZATION & EVENT LISTENERS
// =============================================================================

function applyEmbedParameters() {
    const params = new URLSearchParams(window.location.search);
    const isIframe = window.self !== window.top || params.get("embed") === "true" || params.get("embed") === "1";

    if (isIframe) {
        document.body.classList.add("is-embedded");
    }

    if (params.get("theme") === "dark") {
        document.body.classList.add("theme-dark");
    }
    
    if (params.get("bg") === "transparent" || params.get("transparent") === "true" || params.get("transparent") === "1") {
        document.body.classList.add("theme-transparent");
    }

    if (params.get("header") === "false" || params.get("header") === "0") {
        document.body.classList.add("no-header");
    }

    if (params.has("title")) {
        CONFIG.headerTitle = params.get("title");
    }

    if (params.has("subtitle")) {
        CONFIG.subHeaderTitle = params.get("subtitle");
    }

    if (params.has("speed")) {
        const customSpeed = parseFloat(params.get("speed"));
        if (!isNaN(customSpeed) && customSpeed > 0) {
            CONFIG.scrollSpeedPixelsPerSecond = customSpeed;
        }
    }
}

applyEmbedParameters();
fetchDonorData();

const pollIntervalMs = Math.max((CONFIG.refreshIntervalSeconds || 30) * 1000, 5000);
setInterval(fetchDonorData, pollIntervalMs);

window.addEventListener("resize", () => {
    updateScrollDuration();
});




