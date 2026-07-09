/**
 * Microsoft Graph OneDrive Explorer - Logger Module
 * Manages enterprise-grade application logging, console tracing, and export.
 */

class Logger {
  constructor() {
    this.logs = [];
    this.listeners = [];
  }

  /**
   * Log an event
   * @param {'info' | 'warning' | 'error' | 'success'} type 
   * @param {string} message 
   * @param {Object} [details=null] - Optional details/metadata
   */
  log(type, message, details = null) {
    const entry = {
      id: crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).substring(2, 9),
      timestamp: new Date().toISOString(),
      type,
      message,
      details: details ? JSON.parse(JSON.stringify(details)) : null
    };

    this.logs.unshift(entry); // Newest first

    // Limit log size to 1000 items
    if (this.logs.length > 1000) {
      this.logs.pop();
    }

    // Trigger listeners
    this.listeners.forEach(listener => {
      try {
        listener(entry, this.logs);
      } catch (err) {
        console.error("Logger listener error:", err);
      }
    });

    // Console output with colors
    const prefix = `[OneDrive Explorer] [${type.toUpperCase()}] [${new Date().toLocaleTimeString()}]`;
    if (type === 'error') {
      console.error(prefix, message, details || '');
    } else if (type === 'warning') {
      console.warn(prefix, message, details || '');
    } else {
      console.log(prefix, message, details || '');
    }
  }

  info(message, details = null) {
    this.log('info', message, details);
  }

  warn(message, details = null) {
    this.log('warning', message, details);
  }

  error(message, details = null) {
    this.log('error', message, details);
  }

  success(message, details = null) {
    this.log('success', message, details);
  }

  /**
   * Clear all logs
   */
  clear() {
    this.logs = [];
    this.info("Application logs cleared.");
    this.listeners.forEach(listener => listener(null, this.logs));
  }

  /**
   * Register a callback for log additions
   */
  subscribe(callback) {
    this.listeners.push(callback);
    return () => {
      this.listeners = this.listeners.filter(l => l !== callback);
    };
  }

  /**
   * Export logs as a JSON string
   */
  exportJSON() {
    return JSON.stringify(this.logs, null, 2);
  }
}

export const logger = new Logger();
