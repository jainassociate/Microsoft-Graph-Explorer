/**
 * Microsoft Graph OneDrive Explorer - Authentication Module
 * Manages Access Token storage, expiration tracking, User ID, and Base URL configuration.
 */

import { logger } from './logs.js';

class AuthManager {
  constructor() {
    this.tokenKey = 'mgraph_onedrive_token';
    this.baseUrlKey = 'mgraph_onedrive_baseurl';
    this.userIdKey = 'mgraph_onedrive_userid';
    
    // Default config values
    this.defaultBaseUrl = 'https://graph.microsoft.com/v1.0';
    this.defaultUserId = 'drive@tzm06.onmicrosoft.com';

    this.token = localStorage.getItem(this.tokenKey) || '';
    this.baseUrl = localStorage.getItem(this.baseUrlKey) || this.defaultBaseUrl;
    this.userId = localStorage.getItem(this.userIdKey) || this.defaultUserId;
    this.useMeEndpoint = localStorage.getItem('mgraph_onedrive_use_me_endpoint') !== 'false';

    this.decodedToken = null;
    this.expiryTimer = null;
    this.expiryListeners = [];
    
    if (this.token) {
      this.decodeToken();
    }
  }

  getDrivePathPrefix() {
    if (this.useMeEndpoint || this.userId === 'me') {
      return 'me/drive';
    }
    return `users/${this.userId}/drive`;
  }

  /**
   * Save credentials to LocalStorage
   */
  setCredentials(token, baseUrl, userId) {
    this.token = token.trim();
    this.baseUrl = (baseUrl || this.defaultBaseUrl).trim();
    this.userId = (userId || this.defaultUserId).trim();

    localStorage.setItem(this.tokenKey, this.token);
    localStorage.setItem(this.baseUrlKey, this.baseUrl);
    localStorage.setItem(this.userIdKey, this.userId);

    this.decodeToken();
    logger.success("Auth credentials saved and updated.", { 
      baseUrl: this.baseUrl, 
      userId: this.userId,
      tokenLength: this.token.length,
      expiresIn: this.getExpiryString()
    });

    this.startExpiryTracker();
  }

  /**
   * Clear all stored credentials
   */
  clearCredentials() {
    this.token = '';
    this.baseUrl = this.defaultBaseUrl;
    this.userId = this.defaultUserId;
    this.decodedToken = null;

    localStorage.removeItem(this.tokenKey);
    localStorage.removeItem(this.baseUrlKey);
    localStorage.removeItem(this.userIdKey);

    if (this.expiryTimer) {
      clearInterval(this.expiryTimer);
      this.expiryTimer = null;
    }

    logger.info("Credentials cleared from local storage.");
    this.notifyExpiryListeners();
  }

  /**
   * Parse the JWT Access Token if possible to get expiration and user info
   */
  decodeToken() {
    if (!this.token) {
      this.decodedToken = null;
      return;
    }

    try {
      const parts = this.token.split('.');
      if (parts.length === 3) {
        // Base64Url decode helper
        const base64Url = parts[1];
        const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const jsonPayload = decodeURIComponent(
          atob(base64)
            .split('')
            .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
            .join('')
        );

        this.decodedToken = JSON.parse(jsonPayload);
        this.startExpiryTracker();
      } else {
        // Not a standard 3-part JWT, might be a custom reference token
        this.decodedToken = { customToken: true };
      }
    } catch (err) {
      this.decodedToken = null;
      logger.warn("Unable to parse JWT token structure. Token tracking is disabled.", { error: err.message });
    }
  }

  /**
   * Start a timer checking token expiry in background
   */
  startExpiryTracker() {
    if (this.expiryTimer) {
      clearInterval(this.expiryTimer);
    }

    if (!this.decodedToken || !this.decodedToken.exp) return;

    this.expiryTimer = setInterval(() => {
      this.notifyExpiryListeners();
    }, 10000); // Check every 10s
    this.notifyExpiryListeners();
  }

  /**
   * Subscribes to token expiration updates (for real-time counter rendering)
   */
  onExpiryUpdate(callback) {
    this.expiryListeners.push(callback);
    return () => {
      this.expiryListeners = this.expiryListeners.filter(l => l !== callback);
    };
  }

  notifyExpiryListeners() {
    const info = {
      isExpired: this.isExpired(),
      expiresInSeconds: this.getExpiresInSeconds(),
      expiryString: this.getExpiryString(),
      decoded: this.decodedToken
    };
    this.expiryListeners.forEach(listener => {
      try {
        listener(info);
      } catch (err) {
        console.error(err);
      }
    });
  }

  isExpired() {
    if (!this.token) return true;
    if (!this.decodedToken || !this.decodedToken.exp) return false; // Assume non-expired if parsing fails
    const now = Math.floor(Date.now() / 1000);
    return now >= this.decodedToken.exp;
  }

  getExpiresInSeconds() {
    if (!this.decodedToken || !this.decodedToken.exp) return -1;
    const now = Math.floor(Date.now() / 1000);
    return Math.max(0, this.decodedToken.exp - now);
  }

  getExpiryString() {
    if (!this.token) return "No token";
    if (!this.decodedToken || !this.decodedToken.exp) return "Unknown (Custom format)";
    
    const secondsLeft = this.getExpiresInSeconds();
    if (secondsLeft === 0) return "Expired";

    const minutes = Math.floor(secondsLeft / 60);
    const seconds = secondsLeft % 60;
    return `${minutes}m ${seconds}s`;
  }

  /**
   * Generates authorization & client-request headers for Microsoft Graph API requests
   */
  getHeaders(customHeaders = {}) {
    if (!this.token || this.token === 'null' || this.token === 'undefined' || !this.token.trim()) {
      throw new Error("No Microsoft Graph Access Token provided. Please enter a valid token.");
    }

    return {
      'Authorization': `Bearer ${this.token}`,
      'Content-Type': 'application/json',
      ...customHeaders
    };
  }

  /**
   * Check if credentials look valid with a test call to Microsoft Graph
   */
  async testConnection() {
    if (!this.token || this.token === 'null' || this.token === 'undefined' || !this.token.trim()) {
      throw new Error("No Access Token provided.");
    }

    logger.info("Testing connection to Microsoft Graph...");
    const url = `${this.baseUrl}/me`;
    const headers = this.getHeaders();

    try {
      const response = await fetch(url, { headers });
      if (!response.ok) {
        // Fall back to target user/drive check if /me is restricted (as with daemon app registrations)
        logger.info("Could not fetch /me, checking user drive access instead...");
        const backupUrl = `${this.baseUrl}/users/${this.userId}/drive`;
        const backupRes = await fetch(backupUrl, { headers });
        if (!backupRes.ok) {
          const errMsg = await backupRes.text();
          throw new Error(`Graph API returned Status ${backupRes.status}: ${errMsg}`);
        }
        const data = await backupRes.json();
        logger.success("Connection verified successfully via /users/{id}/drive.", data);
        
        this.useMeEndpoint = false;
        localStorage.setItem('mgraph_onedrive_use_me_endpoint', 'false');
        
        return { success: true, method: 'user_drive', data };
      }
      
      const meData = await response.json();
      logger.success("Connection verified successfully via /me.", meData);
      
      this.useMeEndpoint = true;
      localStorage.setItem('mgraph_onedrive_use_me_endpoint', 'true');
      
      return { success: true, method: 'me', data: meData };
    } catch (err) {
      logger.error("Microsoft Graph connection test failed.", { message: err.message });
      throw err;
    }
  }
}

export const auth = new AuthManager();
