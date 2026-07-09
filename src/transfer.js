/**
 * Microsoft Graph OneDrive Explorer - File Transfer Module
 * Manages small/large (chunked) file uploads, queue management, speed, pause/resume, and downloads.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';
import JSZip from 'jszip';

class TransferManager {
  constructor() {
    this.uploadQueue = [];
    this.activeTransfers = new Map(); // id -> { abortController, file, xhr, speed, etc. }
    this.listeners = [];
  }

  /**
   * Subscribe to transfer queue updates
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
        listener(this.uploadQueue);
      } catch (err) {
        console.error(err);
      }
    });
  }

  /**
   * Add files to the upload queue
   */
  addFilesToQueue(files, parentId = 'root') {
    const addedItems = [];
    Array.from(files).forEach(file => {
      const item = {
        id: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9),
        name: file.name,
        size: file.size,
        type: file.type || 'application/octet-stream',
        parentId: parentId,
        progress: 0,
        status: 'pending', // 'pending', 'uploading', 'paused', 'completed', 'cancelled', 'failed'
        speed: 0, // bytes per second
        timeRemaining: null, // seconds
        bytesUploaded: 0,
        fileObject: file,
        uploadUrl: null, // Set for large uploads
        error: null
      };
      this.uploadQueue.push(item);
      addedItems.push(item);
      logger.info(`Added file "${file.name}" to upload queue.`);
    });
    this.notifyListeners();
    this.processQueue();
    return addedItems;
  }

  /**
   * Process pending uploads (maximum 2 parallel uploads)
   */
  processQueue() {
    const activeCount = Array.from(this.activeTransfers.values()).filter(t => t.status === 'uploading').length;
    if (activeCount >= 2) return;

    const nextItem = this.uploadQueue.find(item => item.status === 'pending');
    if (!nextItem) return;

    this.startUpload(nextItem.id);
  }

  /**
   * Start or resume an upload from the queue
   */
  async startUpload(id) {
    const item = this.uploadQueue.find(q => q.id === id);
    if (!item) return;

    item.status = 'uploading';
    item.error = null;
    this.notifyListeners();

    const abortController = new AbortController();
    this.activeTransfers.set(id, {
      id: id,
      abortController: abortController,
      status: 'uploading',
      startTime: Date.now(),
      lastUploadedBytes: item.bytesUploaded,
      lastUpdateTime: Date.now()
    });

    try {
      if (item.size <= 4 * 1024 * 1024) {
        // Small file upload (<4MB)
        await this.uploadSmallFile(item, abortController);
      } else {
        // Large file upload session
        await this.uploadLargeFile(item, abortController);
      }
    } catch (err) {
      if (err.name === 'AbortError' || item.status === 'paused') {
        logger.info(`Upload of "${item.name}" was paused or cancelled.`);
      } else {
        item.status = 'failed';
        item.error = err.message;
        logger.error(`Upload of "${item.name}" failed: ${err.message}`);
        this.notifyListeners();
      }
    } finally {
      this.activeTransfers.delete(id);
      this.processQueue();
    }
  }

  /**
   * Upload tiny files in one PUT request
   */
  async uploadSmallFile(item, abortController) {
    const file = item.fileObject;
    const cleanName = encodeURIComponent(file.name);
    
    let pathPart = `root:/${cleanName}:/content`;
    if (item.parentId && item.parentId !== 'root') {
      pathPart = `items/${item.parentId}:/${cleanName}:/content`;
    }

    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/${pathPart}`;
    
    const headers = auth.getHeaders({
      'Content-Type': file.type || 'application/octet-stream',
    });

    logger.info(`Initiating simple upload for "${file.name}"...`);

    const response = await fetch(url, {
      method: 'PUT',
      headers: headers,
      body: file,
      signal: abortController.signal
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Upload returned status ${response.status}: ${errText}`);
    }

    item.progress = 100;
    item.bytesUploaded = item.size;
    item.status = 'completed';
    logger.success(`Successfully uploaded simple file: "${file.name}"`);
    this.notifyListeners();
  }

  /**
   * Orchestrates large file chunked upload sessions
   */
  async uploadLargeFile(item, abortController) {
    const file = item.fileObject;
    
    // Create upload session if we don't have one
    if (!item.uploadUrl) {
      logger.info(`Creating chunked upload session for "${file.name}"...`);
      const cleanName = encodeURIComponent(file.name);
      let pathPart = `root:/${cleanName}:/createUploadSession`;
      if (item.parentId && item.parentId !== 'root') {
        pathPart = `items/${item.parentId}:/${cleanName}:/createUploadSession`;
      }

      const sessionUrl = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/${pathPart}`;
      const sessionHeaders = auth.getHeaders({ 'Content-Type': 'application/json' });
      
      const sessionBody = {
        item: {
          "@microsoft.graph.conflictBehavior": "rename",
          name: file.name
        }
      };

      const sessionResponse = await fetch(sessionUrl, {
        method: 'POST',
        headers: sessionHeaders,
        body: JSON.stringify(sessionBody),
        signal: abortController.signal
      });

      if (!sessionResponse.ok) {
        const errText = await sessionResponse.text();
        throw new Error(`Failed to create upload session: ${errText}`);
      }

      const sessionData = await sessionResponse.json();
      item.uploadUrl = sessionData.uploadUrl;
      logger.success(`Upload session initialized for large file: "${file.name}"`);
    }

    // Now upload chunks
    const chunkSize = 320 * 1024 * 4; // 1.25 MB chunk (must be multiple of 320 KB)
    const totalSize = file.size;

    while (item.bytesUploaded < totalSize) {
      if (abortController.signal.aborted || item.status === 'paused') {
        return;
      }

      const start = item.bytesUploaded;
      const end = Math.min(start + chunkSize, totalSize);
      const chunkSlice = file.slice(start, end);
      const chunkLength = end - start;

      logger.info(`Uploading chunk: ${start}-${end - 1} / ${totalSize} for "${file.name}"`);

      const headers = {
        'Content-Length': chunkLength.toString(),
        'Content-Range': `bytes ${start}-${end - 1}/${totalSize}`
      };

      const startTime = performance.now();

      const response = await fetch(item.uploadUrl, {
        method: 'PUT',
        headers: headers,
        body: chunkSlice,
        signal: abortController.signal
      });

      const endTime = performance.now();
      const chunkDuration = (endTime - startTime) / 1000; // in seconds

      if (response.status === 200 || response.status === 201) {
        // Complete!
        item.progress = 100;
        item.bytesUploaded = totalSize;
        item.status = 'completed';
        logger.success(`Large file fully uploaded: "${file.name}"`);
        this.notifyListeners();
        return;
      } else if (response.status === 202) {
        // Chunk accepted, continue
        item.bytesUploaded = end;
        item.progress = Math.round((end / totalSize) * 100);
        
        // Speed & Time remaining calculations
        const speed = chunkDuration > 0 ? Math.round(chunkLength / chunkDuration) : 0;
        item.speed = speed;
        
        if (speed > 0) {
          const remainingBytes = totalSize - end;
          item.timeRemaining = Math.round(remainingBytes / speed);
        } else {
          item.timeRemaining = null;
        }

        this.notifyListeners();
      } else {
        const errorText = await response.text();
        throw new Error(`Session upload failed chunk PUT (${response.status}): ${errorText}`);
      }
    }
  }

  /**
   * Pause an active upload
   */
  pauseUpload(id) {
    const item = this.uploadQueue.find(q => q.id === id);
    const active = this.activeTransfers.get(id);

    if (item && item.status === 'uploading') {
      item.status = 'paused';
      if (active && active.abortController) {
        active.abortController.abort();
      }
      logger.info(`Paused upload: "${item.name}"`);
      this.notifyListeners();
      this.processQueue();
    }
  }

  /**
   * Cancel an active or pending upload
   */
  async cancelUpload(id) {
    const item = this.uploadQueue.find(q => q.id === id);
    const active = this.activeTransfers.get(id);

    if (item) {
      const oldStatus = item.status;
      item.status = 'cancelled';
      item.progress = 0;
      item.bytesUploaded = 0;

      if (active && active.abortController) {
        active.abortController.abort();
      }

      // If we have an active large upload session, cancel it on the server
      if (item.uploadUrl && oldStatus !== 'completed') {
        try {
          logger.info(`Deleting server upload session for cancelled item: "${item.name}"`);
          await fetch(item.uploadUrl, { method: 'DELETE' });
        } catch (err) {
          console.warn("Could not cancel session on Graph server:", err.message);
        }
        item.uploadUrl = null;
      }

      logger.info(`Cancelled upload: "${item.name}"`);
      this.notifyListeners();
      this.processQueue();
    }
  }

  /**
   * Remove a completed/failed upload from the queue list
   */
  removeFromQueue(id) {
    this.cancelUpload(id); // Ensure cancelled first to close handles
    this.uploadQueue = this.uploadQueue.filter(q => q.id !== id);
    this.notifyListeners();
  }

  /**
   * Clear all items from queue
   */
  clearQueue() {
    this.uploadQueue.forEach(item => {
      this.cancelUpload(item.id);
    });
    this.uploadQueue = [];
    this.notifyListeners();
    logger.info("Upload queue fully cleared.");
  }

  /**
   * Formats transfer speed values
   */
  formatSpeed(bytesPerSecond) {
    if (!bytesPerSecond) return '0 B/s';
    const k = 1024;
    const sizes = ['B/s', 'KB/s', 'MB/s', 'GB/s'];
    const i = Math.floor(Math.log(bytesPerSecond) / Math.log(k));
    return parseFloat((bytesPerSecond / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  /**
   * Formats remaining time
   */
  formatTimeRemaining(seconds) {
    if (seconds === null || seconds === undefined) return 'Calculating...';
    if (seconds < 60) return `${seconds}s remaining`;
    const mins = Math.floor(seconds / 60);
    const secs = seconds % 60;
    return `${mins}m ${secs}s remaining`;
  }

  /**
   * Downloads a file using standard Web Blobs
   */
  async downloadFile(itemId, filename) {
    logger.info(`Initiating file download: "${filename}"`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/content`;
    
    let headers = {};
    try {
      headers = auth.getHeaders();
    } catch (err) {
      logger.error("Authentication Error. Cannot start download.");
      throw err;
    }

    try {
      const response = await fetch(url, { headers });
      if (!response.ok) {
        throw new Error(`Failed to download. Graph API status: ${response.status}`);
      }

      // Check content-length
      const totalBytes = parseInt(response.headers.get('content-length') || '0', 10);
      
      const reader = response.body.getReader();
      let receivedBytes = 0;
      const chunks = [];

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        chunks.push(value);
        receivedBytes += value.length;

        if (totalBytes > 0) {
          const progress = Math.round((receivedBytes / totalBytes) * 100);
          // Broadcast progress (optional logging)
          if (progress % 25 === 0) {
            logger.info(`Downloading "${filename}": ${progress}% downloaded`);
          }
        }
      }

      const blob = new Blob(chunks, { type: response.headers.get('content-type') || 'application/octet-stream' });
      const blobUrl = URL.createObjectURL(blob);
      
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);

      logger.success(`Downloaded file successfully: "${filename}"`);
    } catch (err) {
      logger.error(`Download failed for file "${filename}": ${err.message}`);
      throw err;
    }
  }

  /**
   * Downloads a folder recursively and packages it as a ZIP archive
   */
  async downloadFolderAsZip(folderId, folderName) {
    logger.info(`Starting recursive download of folder "${folderName}"...`);
    
    const zip = new JSZip();

    try {
      const headers = auth.getHeaders();
      
      const traverse = async (currentId, currentZipFolder) => {
        let path = `root/children`;
        if (currentId !== 'root') {
          path = `items/${currentId}/children`;
        }
        const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/${path}`;
        
        const response = await fetch(url, { headers });
        if (!response.ok) {
          throw new Error(`Failed to load folder. Graph API status: ${response.status}`);
        }
        
        const data = await response.json();
        const items = data.value || [];

        for (const item of items) {
          if (item.folder) {
            const subFolder = currentZipFolder.folder(item.name);
            await traverse(item.id, subFolder);
          } else {
            logger.info(`Downloading file inside folder: "${item.name}"`);
            const fileUrl = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${item.id}/content`;
            
            const fileRes = await fetch(fileUrl, { headers });
            if (!fileRes.ok) {
              logger.error(`Skipping file "${item.name}": failed to download content (status ${fileRes.status})`);
              continue;
            }
            
            const fileData = await fileRes.arrayBuffer();
            currentZipFolder.file(item.name, fileData);
          }
        }
      };

      await traverse(folderId, zip);

      logger.info(`Compiling ZIP file for folder "${folderName}"...`);
      const content = await zip.generateAsync({ type: 'blob' });
      
      const blobUrl = URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = `${folderName}.zip`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(blobUrl);

      logger.success(`Folder downloaded and archived successfully: "${folderName}.zip"`);
    } catch (err) {
      logger.error(`Failed to download folder "${folderName}": ${err.message}`);
      throw err;
    }
  }
}

export const transferManager = new TransferManager();
