/**
 * Microsoft Graph OneDrive Explorer - Dashboard Module
 * Handles loading, formatting, and visualizing OneDrive storage quota & properties.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';

class DashboardManager {
  constructor() {
    this.driveInfo = null;
  }

  /**
   * Fetch core drive metadata and storage usage from Microsoft Graph
   */
  async loadDriveInfo() {
    logger.info("Loading OneDrive details for dashboard...");
    
    let url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}`;
    let headers = {};
    
    try {
      headers = auth.getHeaders();
    } catch (err) {
      logger.error("Authentication check failed. Cannot load dashboard data.");
      throw err;
    }

    try {
      let response = await fetch(url, { headers });
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API Status ${response.status}: ${errorText}`);
      }

      const data = await response.json();
      this.driveInfo = data;
      logger.success("OneDrive information fetched successfully.", data);
      return data;
    } catch (err) {
      logger.error("Failed to fetch OneDrive info: " + err.message);
      throw err;
    }
  }

  /**
   * Formats file size bytes to human-readable strings (KB, MB, GB, TB)
   */
  formatBytes(bytes, decimals = 2) {
    if (bytes === 0 || !bytes) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  }

  /**
   * Generates quota status details
   */
  getQuotaSummary() {
    if (!this.driveInfo || !this.driveInfo.quota) {
      return {
        total: "N/A",
        used: "N/A",
        remaining: "N/A",
        deleted: "N/A",
        state: "Unknown",
        percentUsed: 0
      };
    }

    const q = this.driveInfo.quota;
    const total = q.total || 0;
    const used = q.used || 0;
    const remaining = q.remaining || 0;
    const deleted = q.deleted || 0;
    const state = q.state || "Normal";
    const percentUsed = total > 0 ? ((used / total) * 100).toFixed(1) : 0;

    return {
      total: this.formatBytes(total),
      used: this.formatBytes(used),
      remaining: this.formatBytes(remaining),
      deleted: this.formatBytes(deleted),
      state: state,
      percentUsed: percentUsed
    };
  }
}

export const dashboardManager = new DashboardManager();
