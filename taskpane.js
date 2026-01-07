/**
 * Tansu Word Add-in - Task Pane JavaScript
 * Communicates with Tansu desktop app via localhost API
 */

// Connect to Tansu desktop app via local.tansu.co (resolves to 127.0.0.1)
// This allows HTTPS communication without mixed-content issues
const API_BASE = 'https://local.tansu.co:5050';
const STORAGE_KEY = 'tansu_welcome_shown';

// State
let variables = [];
let isConnected = false;

// DOM Elements
let welcomeContainerEl, mainContainerEl, skipWelcomeBtn;
let statusEl, statusTextEl, searchEl, variablesListEl, loadingEl, noResultsEl, errorContainerEl, retryBtn;

/**
 * Initialize the add-in when Office is ready
 */
Office.onReady((info) => {
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
    retryBtn.addEventListener('click', checkConnectionAndLoad);
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
 * Check connection to Tansu and load variables
 */
async function checkConnectionAndLoad() {
    setStatus('checking', 'Checking connection...');
    showLoading(true);
    hideError();

    try {
        const response = await fetch(`${API_BASE}/ping`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });

        if (response.ok) {
            isConnected = true;
            markWelcomeShown();
            setStatus('connected', 'Connected to Tansu');
            await loadVariables();
            return;
        }
    } catch (error) {
        console.log('Failed to connect to Tansu:', error.message || error);
    }

    // Connection failed
    isConnected = false;
    setStatus('error', 'Not connected');
    showError();
}

/**
 * Load variables from API
 */
async function loadVariables() {
    try {
        const response = await fetch(`${API_BASE}/variables`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
            throw new Error('Failed to load variables');
        }

        const data = await response.json();
        variables = data.variables || [];
        renderVariables(variables);
        showLoading(false);
    } catch (error) {
        console.error('Failed to load variables:', error);
        showError();
    }
}

/**
 * Refresh variables without showing loading state
 */
async function refreshVariables() {
    if (!isConnected) {
        checkConnectionAndLoad();
        return;
    }

    try {
        const response = await fetch(`${API_BASE}/variables`, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });

        if (!response.ok) {
            throw new Error('Failed to refresh variables');
        }

        const data = await response.json();
        variables = data.variables || [];

        // Re-render with current search filter
        handleSearch();
    } catch (error) {
        console.error('Failed to refresh variables:', error);
        isConnected = false;
        setStatus('error', 'Connection lost');
    }
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
    }

    try {
        // Call the Tansu API to insert the variable
        const response = await fetch(`${API_BASE}/insert`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json'
            },
            body: JSON.stringify({
                name: varName,
                with_unit: false
            })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Insert failed');
        }

        // Success - the variable was inserted by Tansu
        console.log(`Inserted variable: ${varName}`);
    } catch (error) {
        console.error('Failed to insert variable:', error);

        // Show error to user
        alert(`Failed to insert variable: ${error.message}`);
    } finally {
        if (card) {
            card.classList.remove('inserting');
        }
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
