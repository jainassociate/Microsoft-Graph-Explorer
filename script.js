/**
 * Microsoft Graph OneDrive Explorer - Main Orchestrator Script
 * Coordinates authentication, folder navigation, file transfers, logs, and Graph requests.
 */

import { auth } from './src/auth.js';
import { logger } from './src/logs.js';
import { dashboardManager } from './src/dashboard.js';
import { driveNavigator } from './src/drive.js';
import { transferManager } from './src/transfer.js';
import { sharingManager } from './src/sharing.js';
import { batchBuilder } from './src/batch.js';
import { apiTester } from './src/explorer.js';
import { generateCodeSnippets } from './src/code-gen.js';

// Global state variables
let selectedItem = null;
let currentView = 'dashboard-view';
let activeTransferModalAction = 'move'; // 'move' or 'copy'

// DOM Elements
const DOM = {
  // Theme & General
  themeToggle: document.getElementById('theme-toggle'),
  connectionBadge: document.getElementById('connection-badge'),
  tokenExpiryTimer: document.getElementById('token-expiry-timer'),
  sidebar: document.getElementById('app-sidebar'),
  sidebarToggle: document.getElementById('sidebar-toggle'),
  sidebarItems: document.querySelectorAll('.sidebar-item'),
  viewPanels: document.querySelectorAll('.view-panel'),
  toastContainer: document.getElementById('toast-messages-container'),

  // Dashboard
  dashTotal: document.getElementById('dash-total'),
  dashUsed: document.getElementById('dash-used'),
  dashRemaining: document.getElementById('dash-remaining'),
  dashDeleted: document.getElementById('dash-deleted'),
  dashState: document.getElementById('dash-state'),
  dashOwner: document.getElementById('dash-owner'),
  dashType: document.getElementById('dash-type'),
  dashId: document.getElementById('dash-id'),
  refreshDashboardBtn: document.getElementById('refresh-dashboard-btn'),
  storageRadialGauge: document.getElementById('storage-radial-gauge'),
  storagePctLabel: document.getElementById('storage-pct-label'),

  // Explorer
  breadcrumbs: document.getElementById('explorer-breadcrumbs'),
  backBtn: document.getElementById('explorer-back-btn'),
  forwardBtn: document.getElementById('explorer-forward-btn'),
  refreshBtn: document.getElementById('explorer-refresh-btn'),
  newFolderBtn: document.getElementById('explorer-new-folder-btn'),
  quickUploadBtn: document.getElementById('explorer-quick-upload-btn'),
  sortSelect: document.getElementById('explorer-sort-select'),
  filterInput: document.getElementById('explorer-filter-input'),
  toggleListView: document.getElementById('toggle-list-view-btn'),
  toggleGridView: document.getElementById('toggle-grid-view-btn'),
  filesViewport: document.getElementById('files-list-viewport'),
  detailsEmpty: document.getElementById('details-panel-empty'),
  detailsContent: document.getElementById('details-panel-content'),
  paneResizer: document.getElementById('explorer-pane-resizer'),
  detailsPane: document.getElementById('explorer-details-pane'),

  // Item Details fields
  detailsIcon: document.getElementById('details-type-icon'),
  detailsName: document.getElementById('details-item-name'),
  detailsSize: document.getElementById('details-item-size'),
  detailsModified: document.getElementById('details-item-modified'),
  detailsOwner: document.getElementById('details-item-owner'),
  detailsItemId: document.getElementById('details-item-id'),
  detailsCheckoutStatus: document.getElementById('details-checkout-status'),
  detailsCheckoutControls: document.getElementById('details-checkout-controls'),
  actionCheckoutBtn: document.getElementById('action-checkout-btn'),
  checkinInputs: document.getElementById('checkin-inputs'),
  checkinComment: document.getElementById('checkin-comment'),
  actionCheckinBtn: document.getElementById('action-checkin-btn'),

  // Details Action buttons
  actionDownload: document.getElementById('action-download-btn'),
  actionRename: document.getElementById('action-rename-btn'),
  actionMove: document.getElementById('action-move-btn'),
  actionCopy: document.getElementById('action-copy-btn'),
  actionSharing: document.getElementById('action-sharing-btn'),
  actionDelete: document.getElementById('action-delete-btn'),

  // Upload Panel
  dropzone: document.getElementById('upload-dropzone'),
  filePicker: document.getElementById('upload-file-picker'),
  uploadQueueEmpty: document.getElementById('upload-queue-empty'),
  uploadQueueContainer: document.getElementById('upload-queue-container'),
  clearCompletedQueue: document.getElementById('clear-completed-queue-btn'),
  clearAllQueue: document.getElementById('clear-all-queue-btn'),

  // Sharing View
  shareItemSelector: document.getElementById('share-item-selector'),
  refreshPermissionsBtn: document.getElementById('refresh-permissions-btn'),
  permissionsEmpty: document.getElementById('permissions-empty'),
  permissionsTableContainer: document.getElementById('permissions-table-container'),
  permissionsRows: document.getElementById('permissions-rows'),
  createLinkForm: document.getElementById('create-link-form'),
  inviteUsersForm: document.getElementById('invite-users-form'),
  inviteMessage: document.getElementById('invite-message'),

  // Versions View
  versionItemSelector: document.getElementById('version-item-selector'),
  refreshVersionsBtn: document.getElementById('refresh-versions-btn'),
  versionsEmpty: document.getElementById('versions-empty'),
  versionsTableContainer: document.getElementById('versions-table-container'),
  versionsRows: document.getElementById('versions-rows'),

  // Advanced Search View
  searchForm: document.getElementById('advanced-search-form'),
  searchQueryInput: document.getElementById('search-query-input'),
  searchExtSelect: document.getElementById('search-ext-select'),
  searchSizeSelect: document.getElementById('search-size-select'),
  searchResultsEmpty: document.getElementById('search-results-empty'),
  searchResultsContainer: document.getElementById('search-results-container'),
  searchResultsRows: document.getElementById('search-results-rows'),

  // Delta Sync API View
  deltaStartBtn: document.getElementById('delta-start-btn'),
  deltaResetBtn: document.getElementById('delta-reset-btn'),
  deltaTokenDisplay: document.getElementById('delta-token-display'),
  deltaEventsEmpty: document.getElementById('delta-events-empty'),
  deltaEventsTableContainer: document.getElementById('delta-events-table-container'),
  deltaEventRows: document.getElementById('delta-event-rows'),

  // Recent & Shared View
  refreshRecentBtn: document.getElementById('refresh-recent-btn'),
  refreshSharedBtn: document.getElementById('refresh-shared-btn'),
  recentEmpty: document.getElementById('recent-empty'),
  recentList: document.getElementById('recent-list'),
  sharedEmpty: document.getElementById('shared-empty'),
  sharedList: document.getElementById('shared-list'),

  // Batch API View
  batchPresetSelect: document.getElementById('batch-preset-select'),
  batchRequestsList: document.getElementById('batch-requests-list'),
  executeBatchBtn: document.getElementById('execute-batch-btn'),
  batchResponsesEmpty: document.getElementById('batch-responses-empty'),
  batchResponsesResults: document.getElementById('batch-responses-results'),

  // Graph Explorer View
  reqMethod: document.getElementById('req-method'),
  reqUrl: document.getElementById('req-url'),
  reqHeaders: document.getElementById('req-headers'),
  reqBody: document.getElementById('req-body'),
  sendRequestBtn: document.getElementById('send-request-btn'),
  saveFavoriteBtn: document.getElementById('save-favorite-btn'),
  resetHeadersBtn: document.getElementById('reset-headers-btn'),
  tabHistoryBtn: document.getElementById('tab-history-btn'),
  tabFavoritesBtn: document.getElementById('tab-favorites-btn'),
  explorerListContainer: document.getElementById('explorer-list-container'),
  resStatusBadge: document.getElementById('res-status-badge'),
  resTime: document.getElementById('res-time'),
  resEmpty: document.getElementById('res-empty'),
  resBodyPre: document.getElementById('res-body-pre'),
  copyResponseBtn: document.getElementById('copy-response-btn'),
  downloadResponseBtn: document.getElementById('download-response-btn'),
  codeSnippetPre: document.getElementById('code-snippet-pre'),
  codeSnippetTabs: document.getElementById('code-snippets-tabs'),

  // Settings View
  settingsForm: document.getElementById('settings-form'),
  setAccessToken: document.getElementById('set-access-token'),
  setBaseUrl: document.getElementById('set-base-url'),
  setUserId: document.getElementById('set-user-id'),
  testConnectionBtn: document.getElementById('test-connection-btn'),
  resetSettingsBtn: document.getElementById('reset-settings-btn'),

  // Console Logs View
  logsExportBtn: document.getElementById('logs-export-btn'),
  logsClearBtn: document.getElementById('logs-clear-btn'),
  logsFilterSelect: document.getElementById('logs-filter-select'),
  logsConsoleViewport: document.getElementById('logs-console-viewport'),

  // Modals & inputs
  createFolderModal: document.getElementById('create-folder-modal'),
  confirmCreateFolderBtn: document.getElementById('confirm-create-folder-btn'),
  newFolderNameInput: document.getElementById('new-folder-name-input'),

  renameModal: document.getElementById('rename-modal'),
  confirmRenameBtn: document.getElementById('confirm-rename-btn'),
  renameItemInput: document.getElementById('rename-item-input'),

  transferModal: document.getElementById('transfer-op-modal'),
  transferModalTitle: document.getElementById('transfer-modal-title'),
  confirmTransferBtn: document.getElementById('confirm-transfer-btn'),
  transferTargetFolderId: document.getElementById('transfer-target-folder-id'),

  // Custom Confirm Modal
  customConfirmModal: document.getElementById('custom-confirm-modal'),
  customConfirmTitle: document.getElementById('custom-confirm-title'),
  customConfirmMessage: document.getElementById('custom-confirm-message'),
  customConfirmCancel: document.getElementById('custom-confirm-cancel'),
  customConfirmClose: document.getElementById('custom-confirm-close'),
  customConfirmBtn: document.getElementById('custom-confirm-btn'),

  // Graph API Guide Elements
  guideClientId: document.getElementById('guide-client-id'),
  guideRedirectUri: document.getElementById('guide-redirect-uri'),
  guideTenantId: document.getElementById('guide-tenant-id'),
  guideGenAuthBtn: document.getElementById('guide-gen-auth-btn'),
  guideGeneratedUrl: document.getElementById('guide-generated-url'),
};

/* --- SYSTEM ALERTS / TOAST FUNCTIONS --- */
function showToast(type, message) {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${type === 'success' ? '✅' : type === 'error' ? '❌' : '⚠️'}</span> <span>${message}</span>`;
  DOM.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.animation = 'slideIn 0.3s reverse forwards';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

/**
 * Reusable Custom Confirm Dialog replacing browser confirm()
 */
function showConfirm(title, message, isDangerous = false) {
  return new Promise((resolve) => {
    DOM.customConfirmTitle.innerText = title;
    DOM.customConfirmMessage.innerText = message;
    
    if (isDangerous) {
      DOM.customConfirmBtn.style.backgroundColor = 'var(--status-error)';
      DOM.customConfirmBtn.style.borderColor = 'var(--status-error)';
    } else {
      DOM.customConfirmBtn.style.backgroundColor = 'var(--accent)';
      DOM.customConfirmBtn.style.borderColor = 'var(--accent)';
    }
    
    const handleYes = () => {
      cleanup();
      resolve(true);
    };
    
    const handleNo = () => {
      cleanup();
      resolve(false);
    };
    
    const cleanup = () => {
      DOM.customConfirmBtn.removeEventListener('click', handleYes);
      DOM.customConfirmCancel.removeEventListener('click', handleNo);
      DOM.customConfirmClose.removeEventListener('click', handleNo);
      DOM.customConfirmModal.classList.remove('active');
    };
    
    DOM.customConfirmBtn.addEventListener('click', handleYes);
    DOM.customConfirmCancel.addEventListener('click', handleNo);
    DOM.customConfirmClose.addEventListener('click', handleNo);
    
    DOM.customConfirmModal.classList.add('active');
  });
}

/* --- THEME CONTROLLER --- */
function initTheme() {
  const isDark = localStorage.getItem('theme-dark') === 'true';
  DOM.themeToggle.checked = isDark;
  document.body.className = isDark ? 'dark-theme' : 'light-theme';

  DOM.themeToggle.addEventListener('change', (e) => {
    const dark = e.target.checked;
    document.body.className = dark ? 'dark-theme' : 'light-theme';
    localStorage.setItem('theme-dark', dark);
    logger.info(`Swapped theme display to ${dark ? 'Dark Mode' : 'Light Mode'}`);
  });
}

/* --- SPLIT PANE RESIZER --- */
function initResizer() {
  let isDragging = false;

  DOM.paneResizer.addEventListener('mousedown', (e) => {
    isDragging = true;
    DOM.paneResizer.classList.add('dragging');
    document.body.style.cursor = 'col-resize';
  });

  document.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    const containerWidth = DOM.paneResizer.parentElement.clientWidth;
    const targetWidth = containerWidth - e.clientX;
    if (targetWidth > 250 && targetWidth < containerWidth - 300) {
      DOM.detailsPane.style.width = `${targetWidth}px`;
    }
  });

  document.addEventListener('mouseup', () => {
    if (isDragging) {
      isDragging = false;
      DOM.paneResizer.classList.remove('dragging');
      document.body.style.cursor = '';
    }
  });
}

/* --- APP GENERAL NAVIGATION & ROUTER --- */
function hasValidToken() {
  return auth.token && auth.token !== 'null' && auth.token !== 'undefined' && auth.token.trim().length > 0;
}

function handleViewSwitch(viewId) {
  currentView = viewId;
  DOM.viewPanels.forEach(panel => {
    panel.classList.toggle('active', panel.id === viewId);
  });
  DOM.sidebarItems.forEach(item => {
    item.classList.toggle('active', item.getAttribute('data-target') === viewId);
  });

  logger.info(`View switched to panel: ${viewId}`);

  // Automatically trigger loading routines when arriving at panels, but only if token is valid
  if (viewId === 'dashboard-view') {
    if (hasValidToken()) {
      loadDashboardStats();
    } else {
      DOM.dashTotal.innerText = 'N/A';
      DOM.dashUsed.innerText = 'N/A';
      DOM.dashRemaining.innerText = 'N/A';
      DOM.dashDeleted.innerText = 'N/A';
      DOM.dashState.innerText = 'N/A';
      DOM.dashOwner.innerText = 'N/A';
      DOM.dashType.innerText = 'N/A';
      DOM.dashId.innerText = 'N/A';
      DOM.storagePctLabel.innerText = '0%';
      DOM.storageRadialGauge.style.background = 'var(--border-color)';
    }
  } else if (viewId === 'explorer-view') {
    if (hasValidToken()) {
      driveNavigator.navigateTo('root', false);
    } else {
      DOM.filesViewport.innerHTML = '<div style="text-align:center; color:var(--text-muted); padding: 40px; line-height: 1.6;">⚠️ <strong>Connection Required</strong><br>Please configure a valid Microsoft Graph Bearer Access Token in the Settings panel to explore your files.</div>';
    }
  } else if (viewId === 'recents-view') {
    if (hasValidToken()) {
      loadRecentAndShared();
    } else {
      DOM.recentList.innerHTML = '<div style="text-align:center; color:var(--text-muted); padding: 20px;">Connection required. Please supply a Bearer Token.</div>';
      DOM.sharedList.innerHTML = '<div style="text-align:center; color:var(--text-muted); padding: 20px;">Connection required. Please supply a Bearer Token.</div>';
    }
  }
}

function initRouter() {
  DOM.sidebarItems.forEach(item => {
    item.addEventListener('click', () => {
      const target = item.getAttribute('data-target');
      handleViewSwitch(target);
    });
  });

  DOM.sidebarToggle.addEventListener('click', () => {
    DOM.sidebar.classList.toggle('collapsed');
  });
}

/* --- AUTH STATE LISTENER --- */
function initAuthTracking() {
  // Sync fields inicialmente on form
  DOM.setAccessToken.value = auth.token;
  DOM.setBaseUrl.value = auth.baseUrl;
  DOM.setUserId.value = auth.userId;

  // Track status badge
  auth.onExpiryUpdate((info) => {
    if (auth.isExpired()) {
      DOM.connectionBadge.innerText = "🔴 Expired / Unauthenticated";
      DOM.connectionBadge.style.backgroundColor = '#fde7e9';
      DOM.connectionBadge.style.color = '#a80000';
      DOM.tokenExpiryTimer.style.display = 'none';
    } else {
      DOM.connectionBadge.innerText = "🟢 Connected";
      DOM.connectionBadge.style.backgroundColor = '#dff6dd';
      DOM.connectionBadge.style.color = '#107c41';
      
      DOM.tokenExpiryTimer.style.display = 'inline';
      DOM.tokenExpiryTimer.innerText = `Expires in: ${info.expiryString}`;
    }
  });

  // Save Settings Form
  DOM.settingsForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const token = DOM.setAccessToken.value;
    const url = DOM.setBaseUrl.value;
    const user = DOM.setUserId.value;

    auth.setCredentials(token, url, user);
    showToast('success', 'Credentials configured successfully!');
    
    try {
      await auth.testConnection();
      showToast('success', 'Microsoft Graph connection verified successfully!');
      handleViewSwitch('dashboard-view');
    } catch (err) {
      showToast('error', `Connection test failed: ${err.message}`);
    }
  });

  DOM.testConnectionBtn.addEventListener('click', async () => {
    try {
      await auth.testConnection();
      showToast('success', 'Graph API connection successfully validated!');
    } catch (err) {
      showToast('error', `Validation check failed: ${err.message}`);
    }
  });

  DOM.resetSettingsBtn.addEventListener('click', () => {
    auth.clearCredentials();
    DOM.setAccessToken.value = '';
    DOM.setBaseUrl.value = auth.defaultBaseUrl;
    DOM.setUserId.value = auth.defaultUserId;
    showToast('warning', 'Default settings cleared and restored.');
  });

  // Microsoft Graph Setup Guide interactive link builder
  const updateGuideUrl = () => {
    const clientId = DOM.guideClientId.value.trim();
    const redirectUri = DOM.guideRedirectUri.value.trim() || 'https://developer.microsoft.com/en-us/graph/graph-explorer';
    const tenant = DOM.guideTenantId.value.trim() || 'common';
    const scopes = encodeURIComponent('Files.ReadWrite.All User.Read offline_access');

    if (!clientId) {
      DOM.guideGeneratedUrl.value = 'Fill in your App ID above to construct the direct authorization request link...';
      return;
    }

    const url = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${scopes}&response_mode=fragment&state=onedrive_explorer`;
    DOM.guideGeneratedUrl.value = url;
  };

  [DOM.guideClientId, DOM.guideRedirectUri, DOM.guideTenantId].forEach(el => {
    if (el) {
      el.addEventListener('input', updateGuideUrl);
      el.addEventListener('change', updateGuideUrl);
    }
  });

  if (DOM.guideGenAuthBtn) {
    DOM.guideGenAuthBtn.addEventListener('click', () => {
      const clientId = DOM.guideClientId.value.trim();
      if (!clientId) {
        showToast('error', 'Please input your Azure Application (Client) ID first.');
        return;
      }
      updateGuideUrl();
      const url = DOM.guideGeneratedUrl.value;
      window.open(url, '_blank');
      showToast('info', 'Auth URL launched! Authorize the application and copy the resulting bearer access token.');
    });
  }

  // Pre-fill some default info for guide URL builder
  updateGuideUrl();
}

/* --- VIEW 1: DASHBOARD --- */
async function loadDashboardStats() {
  try {
    const data = await dashboardManager.loadDriveInfo();
    const q = dashboardManager.getQuotaSummary();

    DOM.dashTotal.innerText = q.total;
    DOM.dashUsed.innerText = q.used;
    DOM.dashRemaining.innerText = q.remaining;
    DOM.dashDeleted.innerText = q.deleted;
    DOM.dashState.innerText = q.state;
    DOM.dashOwner.innerText = data.owner?.user?.displayName || "Organization/Daemon Client";
    DOM.dashType.innerText = data.driveType || "personal";
    DOM.dashId.innerText = data.id || "N/A";

    DOM.storagePctLabel.innerText = `${q.percentUsed}%`;
    DOM.storageRadialGauge.style.background = `conic-gradient(var(--accent) ${q.percentUsed}%, var(--border-color) ${q.percentUsed}%)`;
  } catch (err) {
    showToast('error', 'Unable to retrieve storage stats: ' + err.message);
  }
}

DOM.refreshDashboardBtn.addEventListener('click', loadDashboardStats);

/* --- VIEW 2: EXPLORER / DRIVE FILES --- */
let isGridView = false;

function formatSize(bytes) {
  if (bytes === undefined || bytes === null) return '-';
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function formatDate(dateStr) {
  if (!dateStr) return '-';
  const date = new Date(dateStr);
  return date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}

function getItemIcon(item) {
  if (item.folder) return '📁';
  if (item.file?.mimeType?.startsWith('image/')) return '🖼️';
  if (item.file?.mimeType?.includes('pdf')) return '📄';
  if (item.file?.mimeType?.includes('sheet') || item.name.endsWith('.xlsx')) return '📊';
  if (item.file?.mimeType?.includes('document') || item.name.endsWith('.docx')) return '📝';
  return '📄';
}

function selectItem(item, element) {
  selectedItem = item;
  document.querySelectorAll('.grid-item, .list-item').forEach(el => el.classList.remove('selected'));
  if (element) {
    element.classList.add('selected');
  }

  DOM.detailsEmpty.style.display = 'none';
  DOM.detailsContent.style.display = 'flex';

  DOM.detailsIcon.innerText = getItemIcon(item);
  DOM.detailsName.innerText = item.name;
  DOM.detailsSize.innerText = item.folder ? `${item.folder.childCount || 0} items` : formatSize(item.size);
  DOM.detailsModified.innerText = formatDate(item.lastModifiedDateTime);
  DOM.detailsOwner.innerText = item.lastModifiedBy?.user?.displayName || "N/A";
  DOM.detailsItemId.innerText = item.id;

  // Checkout info
  if (item.file) {
    DOM.detailsCheckoutControls.style.display = 'flex';
    if (item.publication?.level === 'checkout') {
      DOM.detailsCheckoutStatus.innerText = "Checked Out";
      DOM.actionCheckoutBtn.style.display = 'none';
      DOM.checkinInputs.style.display = 'flex';
    } else {
      DOM.detailsCheckoutStatus.innerText = "Checked In";
      DOM.actionCheckoutBtn.style.display = 'block';
      DOM.checkinInputs.style.display = 'none';
    }
  } else {
    DOM.detailsCheckoutControls.style.display = 'none';
  }
}

function renderExplorer(state) {
  DOM.filesViewport.innerHTML = '';
  selectedItem = null;
  DOM.detailsEmpty.style.display = 'flex';
  DOM.detailsContent.style.display = 'none';

  // Sort and filter cached items
  let items = [...state.items];
  
  // Real-time text filter
  const filterVal = DOM.filterInput.value.toLowerCase().trim();
  if (filterVal) {
    items = items.filter(item => item.name.toLowerCase().includes(filterVal));
  }

  // Sorting logic
  const sortVal = DOM.sortSelect.value;
  items.sort((a, b) => {
    if (sortVal === 'name_asc') return a.name.localeCompare(b.name);
    if (sortVal === 'name_desc') return b.name.localeCompare(a.name);
    if (sortVal === 'size_desc') return (b.size || 0) - (a.size || 0);
    if (sortVal === 'size_asc') return (a.size || 0) - (b.size || 0);
    if (sortVal === 'date_desc') return new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime);
    if (sortVal === 'date_asc') return new Date(a.lastModifiedDateTime) - new Date(b.lastModifiedDateTime);
    return 0;
  });

  // Breadcrumbs builder
  DOM.breadcrumbs.innerHTML = '';
  // Since Graph paths can be complex, we represent history simply
  state.canGoBack ? DOM.backBtn.removeAttribute('disabled') : DOM.backBtn.setAttribute('disabled', 'true');
  state.canGoForward ? DOM.forwardBtn.removeAttribute('disabled') : DOM.forwardBtn.setAttribute('disabled', 'true');

  const backLink = document.createElement('span');
  backLink.className = 'breadcrumb-item';
  backLink.innerText = 'Root Drive';
  backLink.addEventListener('click', () => driveNavigator.navigateTo('root'));
  DOM.breadcrumbs.appendChild(backLink);

  if (state.currentFolderId !== 'root') {
    const sep = document.createElement('span');
    sep.className = 'breadcrumb-separator';
    sep.innerText = ' > ';
    DOM.breadcrumbs.appendChild(sep);

    const subLink = document.createElement('span');
    subLink.className = 'breadcrumb-item active';
    subLink.innerText = `Folder ID: ${state.currentFolderId.substring(0, 8)}...`;
    DOM.breadcrumbs.appendChild(subLink);
  }

  // Populate share options selectors on other tabs dynamically
  DOM.shareItemSelector.innerHTML = '<option value="">-- Choose file --</option>';
  DOM.versionItemSelector.innerHTML = '<option value="">-- Choose file --</option>';

  items.forEach(item => {
    // Sharing & version selectors additions
    const opt = document.createElement('option');
    opt.value = item.id;
    opt.innerText = `${item.folder ? '📁' : '📄'} ${item.name}`;
    DOM.shareItemSelector.appendChild(opt.cloneNode(true));
    if (item.file) {
      DOM.versionItemSelector.appendChild(opt.cloneNode(true));
    }

    if (isGridView) {
      // GRID VIEW RENDER
      const gridRow = document.createElement('div');
      gridRow.className = 'grid-item';
      
      const thumbnailImg = item.thumbnails?.[0]?.medium?.url;
      gridRow.innerHTML = `
        <div class="grid-icon">
          ${thumbnailImg ? `<img src="${thumbnailImg}" class="thumb-preview" referrerPolicy="no-referrer">` : getItemIcon(item)}
        </div>
        <div class="grid-name">${item.name}</div>
      `;

      gridRow.addEventListener('click', () => selectItem(item, gridRow));
      gridRow.addEventListener('dblclick', () => {
        if (item.folder) {
          driveNavigator.navigateTo(item.id);
        }
      });
      DOM.filesViewport.appendChild(gridRow);
    } else {
      // LIST VIEW RENDER
      const listRow = document.createElement('div');
      listRow.className = 'list-item';
      listRow.innerHTML = `
        <div class="list-name-col">
          <span>${getItemIcon(item)}</span>
          <span>${item.name}</span>
        </div>
        <div>${item.folder ? `${item.folder.childCount || 0} items` : formatSize(item.size)}</div>
        <div>${formatDate(item.lastModifiedDateTime)}</div>
        <div style="font-family: var(--font-mono); font-size: 11px;">${item.id.substring(0, 8)}...</div>
      `;

      listRow.addEventListener('click', () => selectItem(item, listRow));
      listRow.addEventListener('dblclick', () => {
        if (item.folder) {
          driveNavigator.navigateTo(item.id);
        }
      });
      DOM.filesViewport.appendChild(listRow);
    }
  });

  if (items.length === 0) {
    DOM.filesViewport.innerHTML = '<div style="text-align:center; color:var(--text-muted); padding: 40px;">No folder items matching parameters are present in this folder view.</div>';
  }
}

// Subscribe explorer updates
driveNavigator.subscribe(renderExplorer);

DOM.backBtn.addEventListener('click', () => driveNavigator.goBack());
DOM.forwardBtn.addEventListener('click', () => driveNavigator.goForward());
DOM.refreshBtn.addEventListener('click', () => driveNavigator.loadCurrentFolder());
DOM.sortSelect.addEventListener('change', () => renderExplorer({ items: driveNavigator.itemsCache, currentFolderId: driveNavigator.currentFolderId }));
DOM.filterInput.addEventListener('input', () => renderExplorer({ items: driveNavigator.itemsCache, currentFolderId: driveNavigator.currentFolderId }));

DOM.toggleListView.addEventListener('click', () => {
  isGridView = false;
  DOM.toggleListView.classList.add('active');
  DOM.toggleGridView.classList.remove('active');
  renderExplorer({ items: driveNavigator.itemsCache, currentFolderId: driveNavigator.currentFolderId });
});

DOM.toggleGridView.addEventListener('click', () => {
  isGridView = true;
  DOM.toggleListView.classList.remove('active');
  DOM.toggleGridView.classList.add('active');
  renderExplorer({ items: driveNavigator.itemsCache, currentFolderId: driveNavigator.currentFolderId });
});

// Modals management trigger
DOM.newFolderBtn.addEventListener('click', () => {
  DOM.createFolderModal.classList.add('active');
  DOM.newFolderNameInput.value = '';
});

DOM.confirmCreateFolderBtn.addEventListener('click', async () => {
  const name = DOM.newFolderNameInput.value.trim();
  if (name) {
    try {
      await driveNavigator.createFolder(name);
      showToast('success', `Created folder "${name}" successfully.`);
      DOM.createFolderModal.classList.remove('active');
    } catch (err) {
      showToast('error', `Failed to create folder: ${err.message}`);
    }
  }
});

// Rename item trigger
DOM.actionRename.addEventListener('click', () => {
  if (selectedItem) {
    DOM.renameModal.classList.add('active');
    DOM.renameItemInput.value = selectedItem.name;
  }
});

DOM.confirmRenameBtn.addEventListener('click', async () => {
  const name = DOM.renameItemInput.value.trim();
  if (selectedItem && name) {
    try {
      await driveNavigator.renameItem(selectedItem.id, name);
      showToast('success', `Item successfully renamed to "${name}"`);
      DOM.renameModal.classList.remove('active');
    } catch (err) {
      showToast('error', `Rename failed: ${err.message}`);
    }
  }
});

// Copy & Move triggers
DOM.actionMove.addEventListener('click', () => {
  if (selectedItem) {
    activeTransferModalAction = 'move';
    DOM.transferModalTitle.innerText = "➡️ Move OneDrive Item";
    DOM.transferTargetFolderId.value = 'root';
    DOM.transferModal.classList.add('active');
  }
});

DOM.actionCopy.addEventListener('click', () => {
  if (selectedItem) {
    activeTransferModalAction = 'copy';
    DOM.transferModalTitle.innerText = "📋 Copy OneDrive Item";
    DOM.transferTargetFolderId.value = 'root';
    DOM.transferModal.classList.add('active');
  }
});

DOM.confirmTransferBtn.addEventListener('click', async () => {
  const targetId = DOM.transferTargetFolderId.value.trim();
  if (selectedItem && targetId) {
    try {
      if (activeTransferModalAction === 'move') {
        await driveNavigator.moveItem(selectedItem.id, targetId);
        showToast('success', 'Item moved successfully.');
      } else {
        await driveNavigator.copyItem(selectedItem.id, targetId);
        showToast('success', 'Copy action triggered. OneDrive is processing copy in background.');
      }
      DOM.transferModal.classList.remove('active');
    } catch (err) {
      showToast('error', `Operation failed: ${err.message}`);
    }
  }
});

// Delete trigger
DOM.actionDelete.addEventListener('click', async () => {
  if (!selectedItem) return;
  const confirmed = await showConfirm(
    "🗑️ Delete Item",
    `Are you sure you want to delete "${selectedItem.name}"? This action will move it to your OneDrive Recycle Bin.`,
    true
  );
  if (confirmed) {
    try {
      await driveNavigator.deleteItem(selectedItem.id);
      showToast('success', 'Item deleted successfully.');
    } catch (err) {
      showToast('error', `Delete failed: ${err.message}`);
    }
  }
});

// Checkout action triggers
DOM.actionCheckoutBtn.addEventListener('click', async () => {
  if (selectedItem) {
    try {
      await sharingManager.checkoutFile(selectedItem.id);
      showToast('success', 'File checked out.');
      await driveNavigator.loadCurrentFolder();
    } catch (err) {
      showToast('error', err.message);
    }
  }
});

DOM.actionCheckinBtn.addEventListener('click', async () => {
  const comment = DOM.checkinComment.value;
  if (selectedItem) {
    try {
      await sharingManager.checkinFile(selectedItem.id, 'published', comment);
      showToast('success', 'File checked in successfully.');
      DOM.checkinComment.value = '';
      await driveNavigator.loadCurrentFolder();
    } catch (err) {
      showToast('error', err.message);
    }
  }
});

DOM.actionDownload.addEventListener('click', async () => {
  if (selectedItem) {
    if (selectedItem.folder) {
      showToast('info', `Scanning and downloading folder "${selectedItem.name}"...`);
      try {
        await transferManager.downloadFolderAsZip(selectedItem.id, selectedItem.name);
        showToast('success', `Folder "${selectedItem.name}" downloaded successfully as ZIP.`);
      } catch (err) {
        showToast('error', `Folder download failed: ${err.message}`);
      }
    } else {
      showToast('info', `Starting download for "${selectedItem.name}"...`);
      try {
        await transferManager.downloadFile(selectedItem.id, selectedItem.name);
        showToast('success', `Downloaded "${selectedItem.name}" successfully.`);
      } catch (err) {
        showToast('error', `File download failed: ${err.message}`);
      }
    }
  }
});

DOM.actionSharing.addEventListener('click', () => {
  if (selectedItem) {
    handleViewSwitch('sharing-view');
    DOM.shareItemSelector.value = selectedItem.id;
    loadPermissionsList(selectedItem.id);
  }
});

// Wire modal dismiss
document.querySelectorAll('.modal-close, .modal-cancel').forEach(el => {
  el.addEventListener('click', () => {
    document.querySelectorAll('.modal-overlay').forEach(modal => modal.classList.remove('active'));
  });
});

/* --- VIEW 3: UPLOAD CENTER --- */
DOM.quickUploadBtn.addEventListener('click', () => {
  handleViewSwitch('upload-view');
});

DOM.dropzone.addEventListener('click', () => DOM.filePicker.click());

DOM.filePicker.addEventListener('change', (e) => {
  if (e.target.files.length > 0) {
    transferManager.addFilesToQueue(e.target.files, driveNavigator.currentFolderId);
    DOM.filePicker.value = '';
  }
});

// Drag & drop handlers
DOM.dropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  DOM.dropzone.style.borderColor = 'var(--accent)';
});

DOM.dropzone.addEventListener('dragleave', () => {
  DOM.dropzone.style.borderColor = 'var(--border-color)';
});

DOM.dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  DOM.dropzone.style.borderColor = 'var(--border-color)';
  if (e.dataTransfer.files.length > 0) {
    transferManager.addFilesToQueue(e.dataTransfer.files, driveNavigator.currentFolderId);
  }
});

function renderUploadQueue(queue) {
  DOM.uploadQueueContainer.innerHTML = '';
  if (queue.length === 0) {
    DOM.uploadQueueEmpty.style.display = 'block';
    DOM.uploadQueueContainer.style.display = 'none';
    return;
  }

  DOM.uploadQueueEmpty.style.display = 'none';
  DOM.uploadQueueContainer.style.display = 'flex';

  queue.forEach(item => {
    const row = document.createElement('div');
    row.className = 'transfer-item';

    const progressPct = item.progress || 0;
    const speedFormatted = item.status === 'uploading' ? transferManager.formatSpeed(item.speed) : '';
    const timeFormatted = item.status === 'uploading' ? transferManager.formatTimeRemaining(item.timeRemaining) : '';

    let statusLabel = item.status.toUpperCase();
    if (item.status === 'failed' && item.error) {
      statusLabel += ` (${item.error})`;
    }

    row.innerHTML = `
      <div class="transfer-info">
        <div style="display: flex; justify-content: space-between; font-weight: 600;">
          <span>${item.name}</span>
          <span style="font-size: 11px; color: var(--text-secondary);">${formatSize(item.size)}</span>
        </div>
        <div class="progress-container">
          <div class="progress-bar" style="width: ${progressPct}%;"></div>
        </div>
        <div style="display: flex; justify-content: space-between; font-size: 11px; color: var(--text-muted);">
          <span>Status: <strong style="color: ${item.status === 'completed' ? 'var(--status-success)' : item.status === 'failed' ? 'var(--status-error)' : 'inherit'}">${statusLabel}</strong></span>
          <span>${speedFormatted} ${speedFormatted ? '|' : ''} ${timeFormatted}</span>
        </div>
      </div>
      <div style="display: flex; gap: 6px;">
        ${item.status === 'uploading' ? `<button class="fluent-button pause-btn" data-id="${item.id}" style="padding: 4px 8px;">⏸️</button>` : ''}
        ${item.status === 'paused' ? `<button class="fluent-button resume-btn" data-id="${item.id}" style="padding: 4px 8px;">▶️</button>` : ''}
        ${['uploading', 'paused', 'pending'].includes(item.status) ? `<button class="fluent-button cancel-btn" data-id="${item.id}" style="padding: 4px 8px; color: var(--status-error);">❌</button>` : ''}
        ${['completed', 'failed', 'cancelled'].includes(item.status) ? `<button class="fluent-button delete-queue-btn" data-id="${item.id}" style="padding: 4px 8px;">🗑️</button>` : ''}
      </div>
    `;

    // Hook buttons actions
    row.querySelector('.pause-btn')?.addEventListener('click', () => transferManager.pauseUpload(item.id));
    row.querySelector('.resume-btn')?.addEventListener('click', () => transferManager.startUpload(item.id));
    row.querySelector('.cancel-btn')?.addEventListener('click', () => transferManager.cancelUpload(item.id));
    row.querySelector('.delete-queue-btn')?.addEventListener('click', () => transferManager.removeFromQueue(item.id));

    DOM.uploadQueueContainer.appendChild(row);
  });
}

// Subcribe upload view updates
transferManager.subscribe(renderUploadQueue);

DOM.clearCompletedQueue.addEventListener('click', () => {
  transferManager.uploadQueue.forEach(item => {
    if (['completed', 'failed', 'cancelled'].includes(item.status)) {
      transferManager.removeFromQueue(item.id);
    }
  });
});

DOM.clearAllQueue.addEventListener('click', () => {
  transferManager.clearQueue();
});

/* --- VIEW 4: SHARING CONTROLLER --- */
async function loadPermissionsList(itemId) {
  if (!itemId) {
    DOM.permissionsEmpty.style.display = 'block';
    DOM.permissionsTableContainer.style.display = 'none';
    return;
  }

  try {
    const list = await sharingManager.listPermissions(itemId);
    DOM.permissionsRows.innerHTML = '';
    
    if (list.length === 0) {
      DOM.permissionsEmpty.style.display = 'block';
      DOM.permissionsTableContainer.style.display = 'none';
      return;
    }

    DOM.permissionsEmpty.style.display = 'none';
    DOM.permissionsTableContainer.style.display = 'block';

    list.forEach(perm => {
      const row = document.createElement('tr');
      row.style.borderBottom = '1px solid var(--border-color)';
      
      const identity = perm.grantedTo?.user?.displayName || perm.grantedToIdentities?.[0]?.user?.displayName || 'Anyone / Public Link';
      const type = perm.link?.type || perm.roles?.join(', ') || 'N/A';
      const webUrl = perm.link?.webUrl ? `<a href="${perm.link.webUrl}" target="_blank" style="color: var(--accent);">Web link</a>` : 'Direct API Permission';

      row.innerHTML = `
        <td style="padding: 8px 0; font-weight: 500;">${identity}</td>
        <td style="padding: 8px 0;">${type}</td>
        <td style="padding: 8px 0;">${webUrl}</td>
        <td style="padding: 8px 0; text-align: right;">
          <button class="fluent-button remove-perm-btn" data-id="${perm.id}" style="padding: 3px 6px; font-size: 11px; color: var(--status-error); border-color: var(--status-error);">Revoke</button>
        </td>
      `;

      row.querySelector('.remove-perm-btn').addEventListener('click', async () => {
        const confirmed = await showConfirm(
          "🔗 Revoke Sharing",
          "Are you sure you want to revoke this item's sharing permission or link?",
          true
        );
        if (confirmed) {
          try {
            await sharingManager.deletePermission(itemId, perm.id);
            showToast('success', 'Permission revoked successfully.');
            loadPermissionsList(itemId);
          } catch (err) {
            showToast('error', err.message);
          }
        }
      });

      DOM.permissionsRows.appendChild(row);
    });
  } catch (err) {
    showToast('error', 'Could not load permissions list: ' + err.message);
  }
}

DOM.shareItemSelector.addEventListener('change', (e) => {
  loadPermissionsList(e.target.value);
});

DOM.refreshPermissionsBtn.addEventListener('click', () => {
  loadPermissionsList(DOM.shareItemSelector.value);
});

DOM.createLinkForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const itemId = DOM.shareItemSelector.value;
  if (!itemId) {
    showToast('warning', 'Please select a file or folder above first.');
    return;
  }

  const type = document.getElementById('link-type').value;
  const scope = document.getElementById('link-scope').value;
  const password = document.getElementById('link-password').value;
  const expiry = document.getElementById('link-expiry').value;

  try {
    const data = await sharingManager.createSharingLink(itemId, type, scope, password, expiry);
    showToast('success', `Sharing link created! Link: ${data.link?.webUrl}`);
    loadPermissionsList(itemId);
    DOM.createLinkForm.reset();
  } catch (err) {
    showToast('error', err.message);
  }
});

DOM.inviteUsersForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const itemId = DOM.shareItemSelector.value;
  if (!itemId) {
    showToast('warning', 'Please select a file or folder above first.');
    return;
  }

  const emailsInput = document.getElementById('invite-emails').value;
  const emails = emailsInput.split(',').map(em => em.trim()).filter(em => em.length > 0);
  const role = document.getElementById('invite-role').value;
  const message = DOM.inviteMessage.value;
  const sendEmail = document.getElementById('invite-send-notif').checked;

  try {
    await sharingManager.inviteUsers(itemId, [role], emails, true, sendEmail, message);
    showToast('success', 'Successfully invited users.');
    loadPermissionsList(itemId);
    DOM.inviteUsersForm.reset();
  } catch (err) {
    showToast('error', err.message);
  }
});

/* --- VIEW 5: VERSIONS CONTROL --- */
async function loadVersionsList(itemId) {
  if (!itemId) {
    DOM.versionsEmpty.style.display = 'block';
    DOM.versionsTableContainer.style.display = 'none';
    return;
  }

  try {
    const list = await sharingManager.listVersions(itemId);
    DOM.versionsRows.innerHTML = '';

    if (list.length === 0) {
      DOM.versionsEmpty.style.display = 'block';
      DOM.versionsTableContainer.style.display = 'none';
      return;
    }

    DOM.versionsEmpty.style.display = 'none';
    DOM.versionsTableContainer.style.display = 'block';

    list.forEach(ver => {
      const row = document.createElement('tr');
      row.style.borderBottom = '1px solid var(--border-color)';

      row.innerHTML = `
        <td style="padding: 10px 0; font-family: var(--font-mono);">${ver.id}</td>
        <td style="padding: 10px 0;">${formatDate(ver.lastModifiedDateTime)}</td>
        <td style="padding: 10px 0;">${ver.lastModifiedBy?.user?.displayName || 'N/A'}</td>
        <td style="padding: 10px 0;">${formatSize(ver.size)}</td>
        <td style="padding: 10px 0; text-align: right; display: flex; gap: 4px; justify-content: flex-end;">
          <button class="fluent-button restore-btn" style="padding: 2px 6px; font-size: 11px;">Restore</button>
          <button class="fluent-button delete-ver-btn" style="padding: 2px 6px; font-size: 11px; color: var(--status-error); border-color: var(--status-error);">Delete</button>
        </td>
      `;

      row.querySelector('.restore-btn').addEventListener('click', async () => {
        const confirmed = await showConfirm(
          "🔄 Restore Version",
          `Are you sure you want to restore the file to version ID ${ver.id}?`
        );
        if (confirmed) {
          try {
            await sharingManager.restoreVersion(itemId, ver.id);
            showToast('success', `File successfully rolled back to version ${ver.id}`);
            loadVersionsList(itemId);
          } catch (err) {
            showToast('error', err.message);
          }
        }
      });

      row.querySelector('.delete-ver-btn').addEventListener('click', async () => {
        const confirmed = await showConfirm(
          "🗑️ Delete Version",
          `Permanently delete historical version ID ${ver.id}?`,
          true
        );
        if (confirmed) {
          try {
            await sharingManager.deleteVersion(itemId, ver.id);
            showToast('success', `Historical version ${ver.id} removed.`);
            loadVersionsList(itemId);
          } catch (err) {
            showToast('error', err.message);
          }
        }
      });

      DOM.versionsRows.appendChild(row);
    });
  } catch (err) {
    showToast('error', 'Could not load version history: ' + err.message);
  }
}

DOM.versionItemSelector.addEventListener('change', (e) => {
  loadVersionsList(e.target.value);
});

DOM.refreshVersionsBtn.addEventListener('click', () => {
  loadVersionsList(DOM.versionItemSelector.value);
});

/* --- VIEW 6: ADVANCED SEARCH --- */
DOM.searchForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const query = DOM.searchQueryInput.value.trim();
  const ext = DOM.searchExtSelect.value;
  const size = DOM.searchSizeSelect.value;

  try {
    let list = await driveNavigator.searchFiles(query);
    
    // Client-side post-filters to satisfy size limit & extensions criteria precisely
    if (ext) {
      list = list.filter(item => item.name.toLowerCase().endsWith(`.${ext}`));
    }
    if (size) {
      if (size === 'small') list = list.filter(item => (item.size || 0) < 1 * 1024 * 1024);
      if (size === 'medium') list = list.filter(item => (item.size || 0) >= 1 * 1024 * 1024 && (item.size || 0) <= 25 * 1024 * 1024);
      if (size === 'large') list = list.filter(item => (item.size || 0) > 25 * 1024 * 1024);
    }

    DOM.searchResultsRows.innerHTML = '';
    
    if (list.length === 0) {
      DOM.searchResultsEmpty.style.display = 'block';
      DOM.searchResultsContainer.style.display = 'none';
      return;
    }

    DOM.searchResultsEmpty.style.display = 'none';
    DOM.searchResultsContainer.style.display = 'grid';

    list.forEach(item => {
      const row = document.createElement('div');
      row.className = 'list-item';
      row.style.gridTemplateColumns = '2.5fr 1fr 1fr';
      
      row.innerHTML = `
        <div class="list-name-col">
          <span>${getItemIcon(item)}</span>
          <span>${item.name}</span>
        </div>
        <div>${formatSize(item.size)}</div>
        <div style="text-align: right;">
          <button class="fluent-button nav-parent-btn" style="padding: 2px 6px; font-size: 11px;">Locate Parent</button>
        </div>
      `;

      row.querySelector('.nav-parent-btn').addEventListener('click', () => {
        if (item.parentReference?.id) {
          handleViewSwitch('explorer-view');
          driveNavigator.navigateTo(item.parentReference.id);
        }
      });

      DOM.searchResultsRows.appendChild(row);
    });
  } catch (err) {
    showToast('error', 'Search operation failed: ' + err.message);
  }
});

/* --- VIEW 7: DELTA SYNC API --- */
DOM.deltaStartBtn.addEventListener('click', async () => {
  try {
    const data = await driveNavigator.runDeltaSync();
    DOM.deltaTokenDisplay.value = data.token || "Sync fully caught up (Last iteration completed).";
    DOM.deltaEventRows.innerHTML = '';

    if (data.items.length === 0) {
      DOM.deltaEventsEmpty.style.display = 'block';
      DOM.deltaEventsTableContainer.style.display = 'none';
      showToast('info', 'Delta Sync executed: No modifications detected since last state.');
      return;
    }

    DOM.deltaEventsEmpty.style.display = 'none';
    DOM.deltaEventsTableContainer.style.display = 'block';

    data.items.forEach(item => {
      const row = document.createElement('tr');
      row.style.borderBottom = '1px solid var(--border-color)';

      const isDeleted = item.deleted ? '🔴 REMOVED' : '🟢 ADDED/EDITED';
      const category = item.folder ? 'Folder' : 'File';

      row.innerHTML = `
        <td style="padding: 8px 0; font-weight: 600;">${isDeleted}</td>
        <td style="padding: 8px 0;">${item.name || `ID: ${item.id.substring(0, 12)}...`}</td>
        <td style="padding: 8px 0;">${category}</td>
        <td style="padding: 8px 0;">${formatDate(item.lastModifiedDateTime || new Date())}</td>
      `;

      DOM.deltaEventRows.appendChild(row);
    });
    
    showToast('success', `Delta Sync loaded ${data.items.length} changes successfully.`);
  } catch (err) {
    showToast('error', err.message);
  }
});

DOM.deltaResetBtn.addEventListener('click', async () => {
  try {
    const data = await driveNavigator.runDeltaSync(true);
    DOM.deltaTokenDisplay.value = data.token || "Sync Track Refreshed.";
    showToast('success', 'Delta sync state token successfully reset to initial baseline.');
  } catch (err) {
    showToast('error', err.message);
  }
});

/* --- VIEW 8: RECENTS & SHARED --- */
async function loadRecentAndShared() {
  // Load Recents
  try {
    const list = await driveNavigator.loadRecentFiles();
    DOM.recentList.innerHTML = '';
    if (list.length === 0) {
      DOM.recentEmpty.style.display = 'block';
    } else {
      DOM.recentEmpty.style.display = 'none';
      list.slice(0, 8).forEach(item => {
        const div = document.createElement('div');
        div.className = 'list-item';
        div.style.gridTemplateColumns = '3fr 1fr';
        div.innerHTML = `
          <div class="list-name-col">
            <span>📄</span>
            <span>${item.name}</span>
          </div>
          <div style="font-size: 11px; text-align: right; color: var(--text-muted);">${formatDate(item.lastModifiedDateTime)}</div>
        `;
        DOM.recentList.appendChild(div);
      });
    }
  } catch (err) {
    showToast('error', 'Failed loading recents: ' + err.message);
  }

  // Load Shared
  try {
    const list = await driveNavigator.loadSharedWithMe();
    DOM.sharedList.innerHTML = '';
    if (list.length === 0) {
      DOM.sharedEmpty.style.display = 'block';
    } else {
      DOM.sharedEmpty.style.display = 'none';
      list.slice(0, 8).forEach(item => {
        const div = document.createElement('div');
        div.className = 'list-item';
        div.style.gridTemplateColumns = '3fr 1fr';
        div.innerHTML = `
          <div class="list-name-col">
            <span>${item.folder ? '📁' : '📄'}</span>
            <span>${item.name}</span>
          </div>
          <div style="font-size: 11px; text-align: right; color: var(--text-muted);">Owner: ${item.remoteItem?.owner?.user?.displayName || 'External'}</div>
        `;
        DOM.sharedList.appendChild(div);
      });
    }
  } catch (err) {
    showToast('error', 'Failed loading shared list: ' + err.message);
  }
}

DOM.refreshRecentBtn.addEventListener('click', loadRecentAndShared);
DOM.refreshSharedBtn.addEventListener('click', loadRecentAndShared);

/* --- VIEW 9: BATCH API CONTROLLER --- */
function initBatchView() {
  const templates = batchBuilder.getTemplates();
  DOM.batchPresetSelect.innerHTML = '<option value="">-- Choose Preset --</option>';
  templates.forEach((temp, idx) => {
    const opt = document.createElement('option');
    opt.value = idx.toString();
    opt.innerText = temp.name;
    DOM.batchPresetSelect.appendChild(opt);
  });

  DOM.batchPresetSelect.addEventListener('change', (e) => {
    const idx = parseInt(e.target.value, 10);
    if (!isNaN(idx)) {
      const temp = templates[idx];
      temp.build();
      renderBatchSequence();
    } else {
      DOM.batchRequestsList.innerHTML = '<span class="text-muted">No active batch structured. Choose a preset template above.</span>';
      DOM.executeBatchBtn.setAttribute('disabled', 'true');
    }
  });

  DOM.executeBatchBtn.addEventListener('click', async () => {
    try {
      const response = await batchBuilder.executeBatch();
      const payload = JSON.parse(response.body);
      DOM.batchResponsesEmpty.style.display = 'none';
      DOM.batchResponsesResults.style.display = 'flex';
      DOM.batchResponsesResults.innerHTML = '';

      payload.responses.forEach(res => {
        const card = document.createElement('div');
        card.className = 'fluent-card';
        card.innerHTML = `
          <div style="display: flex; justify-content: space-between; align-items: center;">
            <strong>Request Task ID: #${res.id}</strong>
            <span class="fluent-button" style="padding: 2px 6px; font-size: 11px; background-color: ${res.status < 300 ? '#dff6dd' : '#fde7e9'}; color: ${res.status < 300 ? '#107c41' : '#a80000'}; pointer-events: none;">
              Status: ${res.status}
            </span>
          </div>
          <pre style="max-height: 120px; overflow-y: auto; font-size: 11px;">${JSON.stringify(res.body, null, 2)}</pre>
        `;
        DOM.batchResponsesResults.appendChild(card);
      });
      showToast('success', 'Batch API completed successfully.');
    } catch (err) {
      showToast('error', err.message);
    }
  });
}

function renderBatchSequence() {
  DOM.batchRequestsList.innerHTML = '';
  batchBuilder.requests.forEach(req => {
    const card = document.createElement('div');
    card.style.padding = '8px';
    card.style.border = '1px solid var(--border-color)';
    card.style.backgroundColor = 'var(--bg-hover)';
    card.style.borderRadius = '4px';
    card.innerHTML = `
      <div style="display: flex; justify-content: space-between;">
        <strong>Task ID: #${req.id} - ${req.method}</strong>
        <span style="font-size: 11px; color: var(--text-muted);">${req.url}</span>
      </div>
      ${req.dependsOn ? `<div style="font-size: 10px; color: var(--status-warning);">Depends on tasks: #${req.dependsOn.join(', #')}</div>` : ''}
    `;
    DOM.batchRequestsList.appendChild(card);
  });
  DOM.executeBatchBtn.removeAttribute('disabled');
}

/* --- VIEW 10: GRAPH EXPLORER REST COMPOSER --- */
let activeSnippetLang = 'curl';

function renderCurrentCodeSnippets() {
  const method = DOM.reqMethod.value;
  const url = DOM.reqUrl.value || 'https://graph.microsoft.com/v1.0/me';
  
  let headers = {};
  try {
    headers = JSON.parse(DOM.reqHeaders.value);
  } catch {
    headers = { "Error": "Invalid Headers JSON" };
  }

  const body = DOM.reqBody.value;
  const snippets = generateCodeSnippets(method, url, headers, body);
  DOM.codeSnippetPre.innerText = snippets[activeSnippetLang] || 'Code snippet generation error';
}

function initGraphExplorerView() {
  DOM.reqUrl.value = `${auth.baseUrl}/${auth.getDrivePathPrefix()}`;

  // Update snippets on every input
  ['input', 'change'].forEach(evt => {
    DOM.reqUrl.addEventListener(evt, renderCurrentCodeSnippets);
    DOM.reqMethod.addEventListener(evt, renderCurrentCodeSnippets);
    DOM.reqHeaders.addEventListener(evt, renderCurrentCodeSnippets);
    DOM.reqBody.addEventListener(evt, renderCurrentCodeSnippets);
  });

  // Snippet tabs click
  DOM.codeSnippetTabs.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      DOM.codeSnippetTabs.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      activeSnippetLang = btn.getAttribute('data-lang');
      renderCurrentCodeSnippets();
    });
  });

  // Reset default headers
  DOM.resetHeadersBtn.addEventListener('click', () => {
    DOM.reqHeaders.value = JSON.stringify({
      "ConsistencyLevel": "eventual",
      "Content-Type": "application/json"
    }, null, 2);
    renderCurrentCodeSnippets();
  });

  // Send request action
  DOM.sendRequestBtn.addEventListener('click', async () => {
    const method = DOM.reqMethod.value;
    const url = DOM.reqUrl.value;
    
    let headers = {};
    try {
      headers = JSON.parse(DOM.reqHeaders.value);
    } catch (err) {
      showToast('error', 'Invalid Headers JSON structure!');
      return;
    }

    const body = DOM.reqBody.value;

    DOM.resEmpty.style.display = 'none';
    DOM.resBodyPre.style.display = 'block';
    DOM.resBodyPre.innerText = 'Executing request in progress...';
    DOM.resStatusBadge.style.display = 'none';
    DOM.resTime.innerText = '';

    try {
      const response = await apiTester.executeRequest(method, url, headers, body);
      
      DOM.resStatusBadge.style.display = 'inline-block';
      DOM.resStatusBadge.innerText = `Status ${response.status}`;
      DOM.resStatusBadge.style.backgroundColor = response.status < 300 ? '#dff6dd' : '#fde7e9';
      DOM.resStatusBadge.style.color = response.status < 300 ? '#107c41' : '#a80000';

      DOM.resTime.innerText = `(${response.duration}ms)`;
      DOM.resBodyPre.innerText = response.body;
      
      renderExplorerLogs();
    } catch (err) {
      DOM.resBodyPre.innerText = `Error: ${err.message}`;
    }
  });

  // Save Favorite
  DOM.saveFavoriteBtn.addEventListener('click', () => {
    const method = DOM.reqMethod.value;
    const url = DOM.reqUrl.value;
    const headers = DOM.reqHeaders.value;
    const body = DOM.reqBody.value;

    const name = prompt("Name this favorite Graph request:", `${method} ${new URL(url).pathname}`);
    if (name) {
      apiTester.addFavorite(name, method, url, JSON.parse(headers), body);
    }
  });

  // Initial snippets render
  renderCurrentCodeSnippets();
  renderExplorerLogs();
}

let activeExplorerTab = 'history';

DOM.tabHistoryBtn.addEventListener('click', () => {
  activeExplorerTab = 'history';
  DOM.tabHistoryBtn.classList.add('active');
  DOM.tabFavoritesBtn.classList.remove('active');
  renderExplorerLogs();
});

DOM.tabFavoritesBtn.addEventListener('click', () => {
  activeExplorerTab = 'favorites';
  DOM.tabHistoryBtn.classList.remove('active');
  DOM.tabFavoritesBtn.classList.add('active');
  renderExplorerLogs();
});

function renderExplorerLogs() {
  DOM.explorerListContainer.innerHTML = '';
  
  if (activeExplorerTab === 'history') {
    if (apiTester.history.length === 0) {
      DOM.explorerListContainer.innerHTML = '<span style="color:var(--text-muted); font-size:12px; display:block; text-align:center;">No history recorded yet.</span>';
      return;
    }
    apiTester.history.forEach(item => {
      const row = document.createElement('div');
      row.style.padding = '6px';
      row.style.borderBottom = '1px solid var(--border-color)';
      row.style.cursor = 'pointer';
      row.style.display = 'flex';
      row.style.justifyContent = 'space-between';
      row.style.fontSize = '12px';
      
      row.innerHTML = `
        <span style="font-weight:600; color:var(--accent);">${item.method}</span>
        <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:200px;">${item.url}</span>
        <span style="color:${item.status < 300 ? 'var(--status-success)' : 'var(--status-error)'};">${item.status}</span>
      `;

      row.addEventListener('click', () => {
        DOM.reqMethod.value = item.method;
        DOM.reqUrl.value = item.url;
        renderCurrentCodeSnippets();
      });

      DOM.explorerListContainer.appendChild(row);
    });
  } else {
    if (apiTester.favorites.length === 0) {
      DOM.explorerListContainer.innerHTML = '<span style="color:var(--text-muted); font-size:12px; display:block; text-align:center;">No favorites saved yet.</span>';
      return;
    }
    apiTester.favorites.forEach(item => {
      const row = document.createElement('div');
      row.style.padding = '6px';
      row.style.borderBottom = '1px solid var(--border-color)';
      row.style.cursor = 'pointer';
      row.style.display = 'flex';
      row.style.justifyContent = 'space-between';
      row.style.fontSize = '12px';

      row.innerHTML = `
        <span style="font-weight:600; max-width:180px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">⭐ ${item.name}</span>
        <button class="remove-fav-btn" style="border:none; background:none; cursor:pointer; color:var(--status-error);">Delete</button>
      `;

      row.addEventListener('click', (e) => {
        if (e.target.className === 'remove-fav-btn') {
          apiTester.removeFavorite(item.id);
          renderExplorerLogs();
          return;
        }
        DOM.reqMethod.value = item.method;
        DOM.reqUrl.value = item.url;
        DOM.reqHeaders.value = JSON.stringify(item.headers || {}, null, 2);
        DOM.reqBody.value = item.body || '';
        renderCurrentCodeSnippets();
      });

      DOM.explorerListContainer.appendChild(row);
    });
  }
}

// Copy & download actions
DOM.copyResponseBtn.addEventListener('click', () => {
  if (DOM.resBodyPre.innerText) {
    navigator.clipboard.writeText(DOM.resBodyPre.innerText);
    showToast('success', 'Copied raw response content to clipboard!');
  }
});

DOM.downloadResponseBtn.addEventListener('click', () => {
  if (DOM.resBodyPre.innerText) {
    const blob = new Blob([DOM.resBodyPre.innerText], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `graph_response_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }
});

/* --- VIEW 11: LOGGER INTEGRATION --- */
function initLogsView() {
  logger.subscribe((entry, logs) => {
    renderLogsConsole(logs);
  });

  DOM.logsClearBtn.addEventListener('click', () => logger.clear());
  DOM.logsFilterSelect.addEventListener('change', () => renderLogsConsole(logger.logs));
  DOM.logsExportBtn.addEventListener('click', () => {
    const text = logger.exportJSON();
    const blob = new Blob([text], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `onedrive_explorer_logs_${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  });

  renderLogsConsole(logger.logs);
}

function renderLogsConsole(logs) {
  DOM.logsConsoleViewport.innerHTML = '';
  const filterVal = DOM.logsFilterSelect.value;
  
  let list = [...logs];
  if (filterVal !== 'all') {
    list = list.filter(l => l.type === filterVal);
  }

  list.forEach(item => {
    const el = document.createElement('div');
    el.className = `log-entry ${item.type}`;
    el.innerText = `[${new Date(item.timestamp).toLocaleTimeString()}] [${item.type.toUpperCase()}] ${item.message}`;
    DOM.logsConsoleViewport.appendChild(el);
  });

  DOM.logsConsoleViewport.scrollTop = DOM.logsConsoleViewport.scrollHeight;
}

/* --- SYSTEM INITIALIZATION HOOK --- */
function init() {
  logger.info("Initializing Microsoft Graph OneDrive Explorer workspace...");
  initTheme();
  initResizer();
  initRouter();
  initAuthTracking();
  initBatchView();
  initGraphExplorerView();
  initLogsView();

  // Try auto authentication on launch
  if (hasValidToken()) {
    auth.testConnection()
      .then(() => {
        showToast('success', 'Successfully restored Microsoft Graph connection!');
        loadDashboardStats();
      })
      .catch((err) => {
        showToast('warning', 'Session token has expired or is invalid. Please configure settings.');
        handleViewSwitch('settings-view');
      });
  } else {
    handleViewSwitch('settings-view');
    showToast('info', 'Please supply a Microsoft Graph Bearer Access Token to initiate the connection.');
  }
}

// Fire launch!
init();
