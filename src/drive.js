/**
 * Microsoft Graph OneDrive Explorer - Drive Navigator & File Operations Module
 * Manages folder trees, file operations (move, copy, rename, delete), delta syncs, shared items, and sorting.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';
import { apiTester } from './explorer.js';

class DriveNavigator {
  constructor() {
    this.currentFolderId = 'root'; // 'root' or specific UUID
    this.history = ['root'];
    this.historyIndex = 0;
    
    this.itemsCache = []; // Current folder contents
    this.deltaToken = null; // Stored delta query token
    this.listeners = [];
  }

  /**
   * Register drive change listeners (e.g. for automatic list updates)
   */
  subscribe(callback) {
    this.listeners.push(callback);
    return () => {
      this.listeners = this.listeners.filter(l => l !== callback);
    };
  }

  notifyListeners() {
    this.listeners.forEach(listener => {
      try {
        listener({
          currentFolderId: this.currentFolderId,
          items: this.itemsCache,
          canGoBack: this.historyIndex > 0,
          canGoForward: this.historyIndex < this.history.length - 1
        });
      } catch (err) {
        console.error(err);
      }
    });
  }

  /**
   * Set folder navigation history
   */
  async navigateTo(folderId, recordHistory = true) {
    logger.info(`Navigating to folder: ${folderId}`);
    
    if (recordHistory) {
      // Cut off any "forward" history we navigated away from
      this.history = this.history.slice(0, this.historyIndex + 1);
      this.history.push(folderId);
      this.historyIndex = this.history.length - 1;
    }

    this.currentFolderId = folderId;
    await this.loadCurrentFolder();
  }

  async goBack() {
    if (this.historyIndex > 0) {
      this.historyIndex--;
      await this.navigateTo(this.history[this.historyIndex], false);
    }
  }

  async goForward() {
    if (this.historyIndex < this.history.length - 1) {
      this.historyIndex++;
      await this.navigateTo(this.history[this.historyIndex], false);
    }
  }

  /**
   * Load files & subfolders for the current folder
   */
  async loadCurrentFolder() {
    logger.info(`Fetching items inside folder: ${this.currentFolderId}`);
    
    let path = `root/children`;
    if (this.currentFolderId !== 'root') {
      path = `items/${this.currentFolderId}/children`;
    }

    // Append $expand=thumbnails to generate automatic previews
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/${path}?$expand=thumbnails`;

    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        this.itemsCache = data.value || [];
        logger.success(`Loaded ${this.itemsCache.length} items from folder.`);
        this.notifyListeners();
        return this.itemsCache;
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Failed to load folder contents: ${err.message}`);
      this.itemsCache = [];
      this.notifyListeners();
      throw err;
    }
  }

  /**
   * Create a new folder
   */
  async createFolder(name) {
    logger.info(`Creating folder: "${name}" inside ${this.currentFolderId}`);
    
    let path = 'root/children';
    if (this.currentFolderId !== 'root') {
      path = `items/${this.currentFolderId}/children`;
    }

    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/${path}`;
    const body = {
      name: name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename"
    };

    try {
      const result = await apiTester.executeRequest('POST', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if (result.status === 201) {
        logger.success(`Folder "${name}" created successfully.`);
        await this.loadCurrentFolder();
        return JSON.parse(result.body);
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Folder creation failed: ${err.message}`);
      throw err;
    }
  }

  /**
   * Delete a file or folder (sends to OneDrive Recycle Bin)
   */
  async deleteItem(itemId) {
    logger.info(`Deleting OneDrive item: ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}`;
    
    try {
      const result = await apiTester.executeRequest('DELETE', url);
      if (result.status === 204 || result.status === 200 || result.status === 202) {
        logger.success(`Item ${itemId} deleted successfully.`);
        await this.loadCurrentFolder();
        return true;
      }
      throw new Error(`Delete failed. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Failed to delete item: ${err.message}`);
      throw err;
    }
  }

  /**
   * Rename a file or folder
   */
  async renameItem(itemId, newName) {
    logger.info(`Renaming item ${itemId} to: "${newName}"`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}`;
    const body = { name: newName };

    try {
      const result = await apiTester.executeRequest('PATCH', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if (result.status === 200) {
        logger.success(`Item renamed to "${newName}" successfully.`);
        await this.loadCurrentFolder();
        return JSON.parse(result.body);
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Rename failed: ${err.message}`);
      throw err;
    }
  }

  /**
   * Move an item to another folder
   */
  async moveItem(itemId, targetFolderId) {
    logger.info(`Moving item ${itemId} to folder: ${targetFolderId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}`;
    
    const body = {
      parentReference: {
        id: targetFolderId
      }
    };

    try {
      const result = await apiTester.executeRequest('PATCH', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if (result.status === 200) {
        logger.success("Item moved successfully.");
        await this.loadCurrentFolder();
        return JSON.parse(result.body);
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Move failed: ${err.message}`);
      throw err;
    }
  }

  /**
   * Copy an item to another folder (Asynchronous long-running Graph operation)
   */
  async copyItem(itemId, targetFolderId, newName = null) {
    logger.info(`Copying item ${itemId} to folder: ${targetFolderId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/copy`;
    
    const body = {
      parentReference: {
        id: targetFolderId
      }
    };

    if (newName) {
      body.name = newName;
    }

    try {
      const result = await apiTester.executeRequest('POST', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      
      if (result.status === 202) {
        const monitorUrl = result.headers['location'] || '';
        logger.success("Copy operation accepted by OneDrive. Running asynchronously in background.", { monitorUrl });
        
        // Return immediately; copying takes time and returns location status header
        return { async: true, monitorUrl };
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Copy failed: ${err.message}`);
      throw err;
    }
  }

  /**
   * Fetch Recycle Bin children
   */
  async loadRecycleBin() {
    logger.info("Loading Recycle Bin items...");
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/special/recyclebin/children`;
    
    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        this.itemsCache = data.value || [];
        logger.success(`Loaded ${this.itemsCache.length} recycled items.`);
        this.notifyListeners();
        return this.itemsCache;
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Failed to load Recycle Bin: ${err.message}`);
      throw err;
    }
  }

  /**
   * Restore item from Recycle Bin
   */
  async restoreRecycledItem(itemId) {
    logger.info(`Restoring item ${itemId} from Recycle Bin...`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/restore`;
    
    try {
      const result = await apiTester.executeRequest('POST', url);
      if ([200, 201, 204].includes(result.status)) {
        logger.success("Item restored successfully from Recycle Bin.");
        return true;
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Failed to restore item: ${err.message}`);
      throw err;
    }
  }

  /**
   * Fetch Recent Files
   */
  async loadRecentFiles() {
    logger.info("Loading recent files list...");
    
    // Attempt /me/drive/recent first if getDrivePathPrefix is not me/drive, as MS Graph only supports recent/shared on signed-in user
    let urlsToTry = [];
    if (auth.getDrivePathPrefix() !== 'me/drive') {
      urlsToTry.push(`${auth.baseUrl}/me/drive/recent`);
    }
    urlsToTry.push(`${auth.baseUrl}/${auth.getDrivePathPrefix()}/recent`);

    let lastError = null;
    for (const url of urlsToTry) {
      try {
        const result = await apiTester.executeRequest('GET', url);
        if (result.status === 200) {
          const data = JSON.parse(result.body);
          this.itemsCache = data.value || [];
          logger.success(`Loaded ${this.itemsCache.length} recent files.`);
          this.notifyListeners();
          return this.itemsCache;
        }
        lastError = new Error(`Status ${result.status}: ${result.body}`);
      } catch (err) {
        lastError = err;
      }
    }

    // Graceful fallback for environments/tokens that do not support recent files (e.g. Daemon/Application-only permissions or non-me target drives)
    logger.warn(`Recent files endpoint is not supported by Microsoft Graph for the current credentials context: ${lastError?.message || 'Unsupported'}. Displaying empty recent files list.`);
    this.itemsCache = [];
    this.notifyListeners();
    return [];
  }

  /**
   * Fetch items shared with the user
   */
  async loadSharedWithMe() {
    logger.info("Loading items shared with me...");
    
    // Attempt /me/drive/sharedWithMe first if getDrivePathPrefix is not me/drive, as MS Graph only supports recent/shared on signed-in user
    let urlsToTry = [];
    if (auth.getDrivePathPrefix() !== 'me/drive') {
      urlsToTry.push(`${auth.baseUrl}/me/drive/sharedWithMe`);
    }
    urlsToTry.push(`${auth.baseUrl}/${auth.getDrivePathPrefix()}/sharedWithMe`);

    let lastError = null;
    for (const url of urlsToTry) {
      try {
        const result = await apiTester.executeRequest('GET', url);
        if (result.status === 200) {
          const data = JSON.parse(result.body);
          this.itemsCache = data.value || [];
          logger.success(`Loaded ${this.itemsCache.length} shared items.`);
          this.notifyListeners();
          return this.itemsCache;
        }
        lastError = new Error(`Status ${result.status}: ${result.body}`);
      } catch (err) {
        lastError = err;
      }
    }

    // Graceful fallback for environments/tokens that do not support shared items (e.g. Daemon/Application-only permissions or non-me target drives)
    logger.warn(`Shared items endpoint is not supported by Microsoft Graph for the current credentials context: ${lastError?.message || 'Unsupported'}. Displaying empty shared items list.`);
    this.itemsCache = [];
    this.notifyListeners();
    return [];
  }

  /**
   * Search files matching criteria
   */
  async searchFiles(query) {
    logger.info(`Searching files for query: "${query}"`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/root/search(q='${encodeURIComponent(query)}')`;
    
    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        this.itemsCache = data.value || [];
        logger.success(`Search completed. Found ${this.itemsCache.length} matches.`);
        this.notifyListeners();
        return this.itemsCache;
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Search failed: ${err.message}`);
      throw err;
    }
  }

  /**
   * Delta Sync operations (track additions, changes, and deletions over time)
   */
  async runDeltaSync(reset = false) {
    logger.info("Running Delta Sync operation...");
    
    let url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/root/delta`;
    if (this.deltaToken && !reset) {
      url = this.deltaToken;
    }

    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        
        // Save the link for next delta call (it can be deltaLink or nextLink)
        this.deltaToken = data['@odata.deltaLink'] || data['@odata.nextLink'] || null;
        this.itemsCache = data.value || [];
        
        logger.success(`Delta sync completed. Retrieved ${this.itemsCache.length} delta event records.`, {
          nextDeltaUrl: this.deltaToken
        });
        this.notifyListeners();
        return { items: this.itemsCache, token: this.deltaToken };
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Delta API Sync failed: ${err.message}`);
      throw err;
    }
  }
}

export const driveNavigator = new DriveNavigator();
