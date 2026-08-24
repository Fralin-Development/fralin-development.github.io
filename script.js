/**
 * Fralin Development - Live Donor Roll
 * Keyless Google Sheets CSV integration with live polling and dynamic scrolling.
 */

// =============================================================================
// CONFIGURATION
// =============================================================================
const CONFIG = {
    // 1. Paste your full Google Sheet URL here (any share link, view link, or tab link with #gid=...)
    // Example: "https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0"
    sheetUrl: "https://myuva-my.sharepoint.com/:x:/r/personal/dbg5af_virginia_edu/_layouts/15/Doc.aspx?sourcedoc=%7BB17AAC43-41DD-40B7-9A69-9C5CD31E8306%7D&file=2026%20Fralin%20Acknowledgements.xlsx&action=default&mobileredirect=true&DefaultItemOpen=1&web=1",

    // 2. Titles displayed at the top of the scrolling list
    headerTitle: "Special Thanks to Our Donors",
    subHeaderTitle: "", // Optional subtitle (e.g., "Annual Campaign 2026")

    // 3. Polling interval: How often (in seconds) to check the Google Sheet for new donors
    refreshIntervalSeconds: 30,

    // 4. Scroll speed in pixels per second (e.g., 35 = slow/relaxed, 50 = standard, 70 = brisk)
    scrollSpeedPixelsPerSecond: 45,

    // 5. Column mappings: Header names (case-insensitive) or 0-based column indexes
    // If set to null or not matched, it will auto-detect or default to Column A for names.
    columns: {
        name: "Name",       // Look for header with "name" or "donor", or column 0
        amount: "Amount",   // Optional: header with "amount" or "gift"
        message: "Message"  // Optional: header with "message", "note", or "dedication"
    },

    // 6. Show status indicator in top-right corner
    showStatusBadge: true
};

// Fallback sample data displayed when no Google Sheet URL is configured yet
const SAMPLE_DONORS = [
    { name: "The Harrison Family Foundation", amount: "$50,000", message: "In honor of Mary & Robert Harrison" },
    { name: "Dr. Arthur & Eleanor Vance", amount: "$25,000", message: "Supporting future generations of leaders" },
    { name: "Commonwealth Trust & Partners", amount: "$15,000", message: "Proud supporters of the Fralin Mission" },
    { name: "Sarah & David Montgomery", amount: "$10,000", message: "" },
    { name: "The Sterling Memorial Fund", amount: "$10,000", message: "In memory of Thomas Sterling, Class of '72" },
    { name: "Dr. Rachel Chen", amount: "$5,000", message: "With deep gratitude" },
    { name: "Marcus & Elena Bennett", amount: "$5,000", message: "" },
    { name: "The Blue Ridge Heritage Fund", amount: "$2,500", message: "Celebrating innovation and excellence" },
    { name: "Jessica & Tyler Brooks", amount: "$2,500", message: "" },
    { name: "Professor William C. Hughes", amount: "$1,000", message: "In memory of Dean Catherine Ward" },
    { name: "Friends of the University", amount: "$1,000", message: "" }
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
// URL PARSER & CSV UTILITIES
// =============================================================================

/**
 * Converts any Google Sheet share / browser / tab URL into a direct CSV export endpoint.
 * Requires NO Google Cloud API keys or OAuth setup.
 */
function getGoogleSheetCsvUrl(rawUrl) {
    if (!rawUrl || typeof rawUrl !== "string") return null;
    const trimmed = rawUrl.trim();
    if (!trimmed) return null;

    // Already a direct export / CSV URL
    if (trimmed.includes("output=csv") || (trimmed.includes("export?format=csv") && !trimmed.includes("/edit"))) {
        return trimmed;
    }

    // Published web format: /spreadsheets/d/e/{ID}/pubhtml or /pub
    const pubMatch = trimmed.match(/\/spreadsheets\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (pubMatch) {
        const docId = pubMatch[1];
        const gidMatch = trimmed.match(/[?&#]gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : "0";
        return `https://docs.google.com/spreadsheets/d/e/${docId}/pub?output=csv&gid=${gid}`;
    }

    // Standard spreadsheet format: /spreadsheets/d/{ID}/edit...
    const idMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (idMatch) {
        const docId = idMatch[1];
        const gidMatch = trimmed.match(/[?&#]gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : "0";
        return `https://docs.google.com/spreadsheets/d/${docId}/export?format=csv&gid=${gid}`;
    }

    // Fallback: return as-is
    return trimmed;
}

/**
 * RFC 4180 compliant CSV parser.
 * Handles escaped quotes, commas inside quotes, multiline values, and varying delimiters.
 */
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
                i++; // Skip escaped quote
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

/**
 * Maps CSV rows to structured donor objects based on configured headers or column positions.
 */
function mapRowsToDonors(rows) {
    if (!rows || rows.length === 0) return [];

    const firstRow = rows[0].map(c => c.toLowerCase().trim());
    let hasHeader = false;
    let nameIdx = 0;
    let amountIdx = -1;
    let messageIdx = -1;

    // Detect if first row is a header row
    const potentialHeaderKeywords = ["name", "donor", "contributor", "supporter", "amount", "gift", "message", "note", "dedication", "tier"];
    const containsHeaderKeyword = firstRow.some(cell => potentialHeaderKeywords.some(kw => cell.includes(kw)));

    if (containsHeaderKeyword) {
        hasHeader = true;

        // Name column detection
        const nameQuery = typeof CONFIG.columns.name === "string" ? CONFIG.columns.name.toLowerCase() : "name";
        const foundNameIdx = firstRow.findIndex(c => c.includes(nameQuery) || c.includes("donor") || c.includes("contributor"));
        if (foundNameIdx !== -1) nameIdx = foundNameIdx;

        // Amount column detection
        const amountQuery = typeof CONFIG.columns.amount === "string" ? CONFIG.columns.amount.toLowerCase() : "amount";
        const foundAmountIdx = firstRow.findIndex(c => c.includes(amountQuery) || c.includes("gift") || c.includes("donation") || c.includes("tier"));
        if (foundAmountIdx !== -1) amountIdx = foundAmountIdx;

        // Message column detection
        const msgQuery = typeof CONFIG.columns.message === "string" ? CONFIG.columns.message.toLowerCase() : "message";
        const foundMsgIdx = firstRow.findIndex(c => c.includes(msgQuery) || c.includes("note") || c.includes("dedication") || c.includes("memory") || c.includes("honor"));
        if (foundMsgIdx !== -1) messageIdx = foundMsgIdx;
    } else {
        // Fall back to explicit numerical indexes if provided
        if (typeof CONFIG.columns.name === "number") nameIdx = CONFIG.columns.name;
        if (typeof CONFIG.columns.amount === "number") amountIdx = CONFIG.columns.amount;
        if (typeof CONFIG.columns.message === "number") messageIdx = CONFIG.columns.message;
    }

    const dataRows = hasHeader ? rows.slice(1) : rows;

    const donors = [];
    for (const row of dataRows) {
        const name = row[nameIdx] ? row[nameIdx].trim() : "";
        if (!name) continue;

        const amount = (amountIdx >= 0 && row[amountIdx]) ? row[amountIdx].trim() : "";
        const message = (messageIdx >= 0 && row[messageIdx]) ? row[messageIdx].trim() : "";

        donors.push({ name, amount, message });
    }

    return donors;
}

// =============================================================================
// RENDERING & SCROLL TIMING
// =============================================================================

/**
 * Calculates dynamic animation duration so scroll speed remains perfectly consistent
 * regardless of the number of donors in the sheet.
 */
function updateScrollDuration() {
    if (!donorListElement || !scrollContainer) return;

    // Use requestAnimationFrame to ensure accurate DOM dimensions after paint
    requestAnimationFrame(() => {
        const contentHeight = donorListElement.offsetHeight;
        const containerHeight = scrollContainer.offsetHeight || window.innerHeight;
        const totalTravelDistance = contentHeight + containerHeight;
        const speed = Math.max(CONFIG.scrollSpeedPixelsPerSecond || 45, 10);
        const durationSeconds = Math.max(totalTravelDistance / speed, 12);

        donorListElement.style.animationDuration = `${durationSeconds.toFixed(1)}s`;
    });
}

/**
 * Renders the list of donors into the DOM.
 */
function renderDonors(donors) {
    // Update headers
    if (mainHeaderElement) mainHeaderElement.textContent = CONFIG.headerTitle;
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

function updateStatus(text, isLive = true) {
    if (!CONFIG.showStatusBadge || !statusBadge) return;
    statusBadge.classList.remove("hidden");
    if (statusText) {
        statusText.textContent = text;
    }
}

async function fetchDonorData() {
    if (isFetching) return;
    isFetching = true;

    const csvUrl = getGoogleSheetCsvUrl(CONFIG.sheetUrl);

    // If no URL is provided, display sample demo data
    if (!csvUrl) {
        console.info("ℹ️ No Google Sheet URL configured in CONFIG.sheetUrl. Displaying sample donor data.");
        renderDonors(SAMPLE_DONORS);
        updateStatus("Demo Mode");
        isFetching = false;
        return;
    }

    try {
        // Append timestamp cache-buster to ensure live updates bypass browser cache
        const cacheBustedUrl = csvUrl.includes("?") 
            ? `${csvUrl}&_t=${Date.now()}` 
            : `${csvUrl}?_t=${Date.now()}`;

        const response = await fetch(cacheBustedUrl, {
            method: "GET",
            headers: {
                "Accept": "text/csv, text/plain, */*"
            }
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const csvText = await response.text();

        // Check if Google returned an HTML login / permission page instead of CSV
        if (csvText.includes("<!DOCTYPE html>") || csvText.includes("<html")) {
            throw new Error("Received HTML login page instead of CSV. Ensure sheet is shared with 'Anyone with the link can view'.");
        }

        const rows = parseCSV(csvText);
        const donors = mapRowsToDonors(rows);

        // Check if data actually changed to avoid restarting animation unnecessarily
        const signature = JSON.stringify(donors);
        if (signature !== lastDataSignature) {
            lastDataSignature = signature;
            renderDonors(donors);
        }

        hideBannerError();
        const now = new Date();
        const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        updateStatus(`Live Sync (${timeStr})`);
    } catch (error) {
        console.error("Error fetching Google Sheet CSV:", error);
        showBannerError(
            `Unable to access Google Sheet.<br>Please verify the sheet permissions are set to: <strong>"Anyone with the link can view"</strong>.`
        );
        // If first load failed, show sample donors so screen isn't blank
        if (!lastDataSignature) {
            renderDonors(SAMPLE_DONORS);
        }
        updateStatus("Sync Error");
    } finally {
        isFetching = false;
    }
}

// =============================================================================
// INITIALIZATION & EVENT LISTENERS
// =============================================================================

// Initial load
fetchDonorData();

// Live polling
const pollIntervalMs = Math.max((CONFIG.refreshIntervalSeconds || 30) * 1000, 5000);
setInterval(fetchDonorData, pollIntervalMs);

// Recalculate duration on window resize for responsive display
window.addEventListener("resize", () => {
    updateScrollDuration();
});


