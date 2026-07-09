/**
 * Microsoft Graph OneDrive Explorer - API Testing Tool Module (Graph Explorer)
 * Supports full custom queries, headers, body, logs, request execution history, and favorites.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';
import { generateCodeSnippets } from './code-gen.js';

class ApiTestingTool {
  constructor() {
    this.historyKey = 'mgraph_explorer_history';
    this.favoritesKey = 'mgraph_explorer_favorites';

    this.history = JSON.parse(localStorage.getItem(this.historyKey)) || [];
    this.favorites = JSON.parse(localStorage.getItem(this.favoritesKey)) || [];
    this.listeners = [];
  }

  /**
   * Execute an arbitrary request against Microsoft Graph
   * @param {string} method - HTTP Verb (GET, POST, etc.)
   * @param {string} url - Target URL
   * @param {Object} [customHeaders={}] - Additional headers
   * @param {string} [body=''] - Request Payload
   * @returns {Promise<Object>} Response object containing status, statusText, headers, body, timeTaken
   */
  async executeRequest(method, url, customHeaders = {}, body = '') {
    const startTime = performance.now();
    let finalHeaders = {};
    
    try {
      finalHeaders = auth.getHeaders(customHeaders);
    } catch (err) {
      logger.error("Authentication Error: " + err.message);
      throw err;
    }

    const requestOptions = {
      method: method.toUpperCase(),
      headers: finalHeaders,
    };

    if (['POST', 'PUT', 'PATCH'].includes(requestOptions.method) && body) {
      requestOptions.body = body;
    }

    logger.info(`Sending Request: ${method} ${url}`);

    try {
      const response = await fetch(url, requestOptions);
      const endTime = performance.now();
      const duration = Math.round(endTime - startTime);

      // Extract headers from the response
      const responseHeaders = {};
      response.headers.forEach((value, key) => {
        responseHeaders[key] = value;
      });

      // Parse body (text or json)
      let responseBody = '';
      const contentType = response.headers.get('content-type') || '';
      
      if (contentType.includes('application/json')) {
        try {
          const jsonVal = await response.json();
          responseBody = JSON.stringify(jsonVal, null, 2);
        } catch {
          responseBody = await response.text();
        }
      } else {
        responseBody = await response.text();
      }

      const result = {
        status: response.status,
        statusText: response.statusText,
        headers: responseHeaders,
        body: responseBody,
        duration: duration,
        method: method,
        url: url,
        timestamp: new Date().toISOString()
      };

      // Add to history
      this.addToHistory(result);

      if (response.ok) {
        logger.success(`Request succeeded. Status ${response.status} in ${duration}ms.`, { method, url });
      } else {
        logger.error(`Request failed. Status ${response.status} in ${duration}ms.`, { 
          method, 
          url, 
          error: responseBody 
        });
      }

      return result;
    } catch (err) {
      const endTime = performance.now();
      const duration = Math.round(endTime - startTime);
      
      const errorResult = {
        status: 0,
        statusText: "Network Error / CORS Block",
        headers: {},
        body: JSON.stringify({
          error: {
            code: "NetworkError",
            message: err.message,
            innerError: {
              date: new Date().toISOString(),
              request_id: "N/A"
            }
          }
        }, null, 2),
        duration: duration,
        method: method,
        url: url,
        timestamp: new Date().toISOString()
      };

      this.addToHistory(errorResult);
      logger.error(`Network request failed: ${err.message}`, { method, url, duration });
      return errorResult;
    }
  }

  addToHistory(item) {
    // Keep raw request info in history to allow re-running
    const historyItem = {
      id: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9),
      method: item.method,
      url: item.url,
      timestamp: item.timestamp,
      status: item.status,
      duration: item.duration
    };

    this.history.unshift(historyItem);
    if (this.history.length > 50) {
      this.history.pop();
    }
    localStorage.setItem(this.historyKey, JSON.stringify(this.history));
    this.notifyListeners();
  }

  clearHistory() {
    this.history = [];
    localStorage.removeItem(this.historyKey);
    logger.info("Request history cleared.");
    this.notifyListeners();
  }

  addFavorite(name, method, url, headers = {}, body = '') {
    const favorite = {
      id: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9),
      name: name || `${method} ${new URL(url).pathname}`,
      method: method,
      url: url,
      headers: headers,
      body: body,
      timestamp: new Date().toISOString()
    };

    this.favorites.unshift(favorite);
    localStorage.setItem(this.favoritesKey, JSON.stringify(this.favorites));
    logger.success(`Saved request "${favorite.name}" to Favorites.`);
    this.notifyListeners();
  }

  removeFavorite(id) {
    this.favorites = this.favorites.filter(fav => fav.id !== id);
    localStorage.setItem(this.favoritesKey, JSON.stringify(this.favorites));
    logger.info("Request removed from Favorites.");
    this.notifyListeners();
  }

  exportFavorites() {
    return JSON.stringify(this.favorites, null, 2);
  }

  importFavorites(favoritesJson) {
    try {
      const imported = JSON.parse(favoritesJson);
      if (Array.isArray(imported)) {
        this.favorites = [...imported, ...this.favorites];
        localStorage.setItem(this.favoritesKey, JSON.stringify(this.favorites));
        logger.success(`Successfully imported ${imported.length} favorites.`);
        this.notifyListeners();
        return true;
      }
      throw new Error("Invalid format: Must be a JSON array.");
    } catch (err) {
      logger.error("Failed to import favorites: " + err.message);
      return false;
    }
  }

  subscribe(callback) {
    this.listeners.push(callback);
    return () => {
      this.listeners = this.listeners.filter(l => l !== callback);
    };
  }

  notifyListeners() {
    this.listeners.forEach(listener => {
      try {
        listener({ history: this.history, favorites: this.favorites });
      } catch (err) {
        console.error(err);
      }
    });
  }
}

export const apiTester = new ApiTestingTool();
