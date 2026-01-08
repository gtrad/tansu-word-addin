/**
 * Tansu Word Add-in - Task Pane JavaScript
 * Communicates with Tansu desktop app via localhost API
 */

// API endpoints - local.tansu.co resolves to 127.0.0.1 via DNS, with valid SSL cert
const API_BASE = 'https://local.tansu.co:5050';
const WS_URL = 'wss://local.tansu.co:5050/ws';
const STORAGE_KEY = 'tansu_welcome_shown';

// State
let variables = [];
let isConnected = false;
let ws = null;
let useHttpFallback = false;
let wsConnectionAttempts = 0;
const MAX_WS_ATTEMPTS = 2;

// Debug logging to visible panel
function debugLog(msg) {
    console.log('[Tansu]', msg);
    const debugEl = document.getElementById('debug-log');
    if (debugEl) {
        const time = new Date().toLocaleTimeString();
        debugEl.innerHTML += `<div>${time}: ${msg}</div>`;
        debugEl.scrollTop = debugEl.scrollHeight;
    }
}

// DOM Elements
let welcomeContainerEl, mainContainerEl, skipWelcomeBtn;
let statusEl, statusTextEl, searchEl, variablesListEl, loadingEl, noResultsEl, errorContainerEl, retryBtn;

/**
 * Initialize the add-in when Office is ready
 */
debugLog('Page loaded, waiting for Office.js...');
Office.onReady((info) => {
    debugLog('Office ready: host=' + info.host);
    if (info.host === Office.HostType.Word) {
        initializeAddin();
    }
});

/**
 * Initialize DOM elements and start the app
 */
function initializeAddin() {
    // Get DOM elements
    welcomeContainerEl = document.getElementById('welcome-container');
    mainContainerEl = document.getElementById('main-container');
    skipWelcomeBtn = document.getElementById('skip-welcome-btn');
    statusEl = document.getElementById('status');
    statusTextEl = statusEl.querySelector('.status-text');
    searchEl = document.getElementById('search');
    variablesListEl = document.getElementById('variables-list');
    loadingEl = document.getElementById('loading');
    noResultsEl = document.getElementById('no-results');
    errorContainerEl = document.getElementById('error-container');
    retryBtn = document.getElementById('retry-btn');

    // Set up event listeners
    searchEl.addEventListener('input', handleSearch);
    retryBtn.addEventListener('click', () => {
        wsConnectionAttempts = 0;
        useHttpFallback = false;
        checkConnectionAndLoad();
    });
    skipWelcomeBtn.addEventListener('click', dismissWelcome);

    // Check if first run
    if (isFirstRun()) {
        showWelcome();
    } else {
        showMainApp();
    }
}

/**
 * Check if this is the first run
 */
function isFirstRun() {
    try {
        return !localStorage.getItem(STORAGE_KEY);
    } catch (e) {
        // localStorage might not be available
        return false;
    }
}

/**
 * Mark welcome as shown
 */
function markWelcomeShown() {
    try {
        localStorage.setItem(STORAGE_KEY, 'true');
    } catch (e) {
        // Ignore if localStorage not available
    }
}

/**
 * Show welcome screen
 */
function showWelcome() {
    welcomeContainerEl.style.display = 'flex';
    mainContainerEl.style.display = 'none';
}

/**
 * Dismiss welcome and show main app
 */
function dismissWelcome() {
    markWelcomeShown();
    showMainApp();
}

/**
 * Show main app
 */
function showMainApp() {
    welcomeContainerEl.style.display = 'none';
    mainContainerEl.style.display = 'flex';

    // Start the app
    checkConnectionAndLoad();

    // Poll for updates every 5 seconds
    setInterval(refreshVariables, 5000);
}

/**
 * Check connection to Tansu - try WebSocket first, fall back to HTTP
 */
function checkConnectionAndLoad() {
    setStatus('checking', 'Connecting...');
    showLoading(true);
    hideError();

    if (useHttpFallback) {
        debugLog('Using HTTP fallback mode');
        loadVariablesViaHttp();
        return;
    }

    // Close existing WebSocket connection
    if (ws) {
        ws.close();
        ws = null;
    }

    wsConnectionAttempts++;
    debugLog(`WebSocket attempt ${wsConnectionAttempts}/${MAX_WS_ATTEMPTS}`);

    try {
        debugLog('Connecting to ' + WS_URL);
        ws = new WebSocket(WS_URL);

        // Set a timeout for WebSocket connection
        const wsTimeout = setTimeout(() => {
            if (ws && ws.readyState !== WebSocket.OPEN) {
                debugLog('WebSocket timeout, trying HTTP fallback');
                ws.close();
                tryHttpFallback();
            }
        }, 3000);

        ws.onopen = () => {
            clearTimeout(wsTimeout);
            debugLog('WebSocket connected!');
            isConnected = true;
            wsConnectionAttempts = 0;
            markWelcomeShown();
            setStatus('connected', 'Connected to Tansu');
            ws.send(JSON.stringify({ type: 'get_variables' }));
        };

        ws.onmessage = (event) => {
            try {
                debugLog('Received: ' + event.data.substring(0, 50) + '...');
                const data = JSON.parse(event.data);
                if (data.type === 'variables') {
                    variables = data.variables || [];
                    renderVariables(variables);
                    showLoading(false);
                } else if (data.type === 'insert_result') {
                    if (!data.success) {
                        alert(`Failed to insert variable: ${data.error}`);
                    }
                }
            } catch (e) {
                debugLog('Parse error: ' + e.message);
            }
        };

        ws.onerror = (error) => {
            clearTimeout(wsTimeout);
            debugLog('WebSocket error: ' + (error.message || 'Connection failed'));
        };

        ws.onclose = (event) => {
            clearTimeout(wsTimeout);
            debugLog('WebSocket closed: code=' + event.code + ' clean=' + event.wasClean);

            if (!isConnected) {
                // Connection never established, try fallback
                tryHttpFallback();
            } else {
                // Was connected but got disconnected
                isConnected = false;
                setStatus('error', 'Disconnected');
            }
            ws = null;
        };

    } catch (error) {
        debugLog('WebSocket exception: ' + error.message);
        tryHttpFallback();
    }
}

/**
 * Try HTTP fallback after WebSocket fails
 */
function tryHttpFallback() {
    if (wsConnectionAttempts >= MAX_WS_ATTEMPTS) {
        debugLog('Switching to HTTP fallback mode');
        useHttpFallback = true;
        loadVariablesViaHttp();
    } else {
        // Try WebSocket again
        setTimeout(checkConnectionAndLoad, 1000);
    }
}

/**
 * Load variables via HTTP fetch (fallback mode)
 */
async function loadVariablesViaHttp() {
    try {
        debugLog('Fetching variables via HTTP...');
        const response = await fetch(`${API_BASE}/variables`);

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        debugLog('HTTP response: ' + JSON.stringify(data).substring(0, 50) + '...');

        variables = data.variables || [];
        isConnected = true;
        markWelcomeShown();
        setStatus('connected', 'Connected to Tansu');
        renderVariables(variables);
        showLoading(false);

    } catch (error) {
        debugLog('HTTP error: ' + error.message);
        isConnected = false;
        setStatus('error', 'Not connected');
        showError();
    }
}

/**
 * Refresh variables
 */
function refreshVariables() {
    if (useHttpFallback) {
        loadVariablesViaHttp();
        return;
    }

    if (!isConnected || !ws || ws.readyState !== WebSocket.OPEN) {
        checkConnectionAndLoad();
        return;
    }
    ws.send(JSON.stringify({ type: 'get_variables' }));
}

/**
 * Handle search input
 */
function handleSearch() {
    const query = searchEl.value.toLowerCase().trim();

    if (!query) {
        renderVariables(variables);
        return;
    }

    const filtered = variables.filter(v =>
        v.name.toLowerCase().includes(query) ||
        String(v.value).toLowerCase().includes(query)
    );

    renderVariables(filtered);
}

/**
 * Render variables list
 */
function renderVariables(vars) {
    if (vars.length === 0) {
        variablesListEl.innerHTML = '';
        noResultsEl.style.display = 'block';
        return;
    }

    noResultsEl.style.display = 'none';

    variablesListEl.innerHTML = vars.map(v => `
        <div class="variable-card" data-name="${escapeHtml(v.name)}">
            <div class="variable-name">${escapeHtml(v.name)}</div>
            <div class="variable-value">
                ${escapeHtml(String(v.value))}
                ${v.unit ? `<span class="variable-unit">${escapeHtml(v.unit)}</span>` : ''}
            </div>
        </div>
    `).join('');

    // Add click handlers
    variablesListEl.querySelectorAll('.variable-card').forEach(card => {
        card.addEventListener('click', () => insertVariable(card.dataset.name));
    });
}

/**
 * Insert a variable into the Word document
 */
async function insertVariable(varName) {
    const card = variablesListEl.querySelector(`[data-name="${varName}"]`);
    if (card) {
        card.classList.add('inserting');
        setTimeout(() => card.classList.remove('inserting'), 500);
    }

    // Find the variable to get its value
    const variable = variables.find(v => v.name === varName);
    if (!variable) {
        alert('Variable not found');
        return;
    }

    // Insert directly into Word using Office.js
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.insertText(String(variable.value), Word.InsertLocation.replace);
            await context.sync();
            debugLog('Inserted: ' + varName + ' = ' + variable.value);
        });
    } catch (error) {
        debugLog('Insert error: ' + error.message);
        alert('Failed to insert variable: ' + error.message);
    }
}

/**
 * Set status indicator
 */
function setStatus(type, text) {
    statusEl.className = `status status-${type}`;
    statusTextEl.textContent = text;
}

/**
 * Show/hide loading state
 */
function showLoading(show) {
    loadingEl.style.display = show ? 'block' : 'none';
    variablesListEl.style.display = show ? 'none' : 'flex';
}

/**
 * Show error state
 */
function showError() {
    document.getElementById('variables-container').style.display = 'none';
    errorContainerEl.style.display = 'flex';
    showLoading(false);
}

/**
 * Hide error state
 */
function hideError() {
    document.getElementById('variables-container').style.display = 'flex';
    errorContainerEl.style.display = 'none';
}

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
