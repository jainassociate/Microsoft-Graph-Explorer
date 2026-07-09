/**
 * Microsoft Graph OneDrive Explorer - Batch API Module (/$batch)
 * Builds and processes multiple requests, dependencies, and complex batches.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';
import { apiTester } from './explorer.js';

class BatchBuilder {
  constructor() {
    this.requests = [];
  }

  /**
   * Reset the current batch
   */
  clearBatch() {
    this.requests = [];
    logger.info("Batch builder reset.");
  }

  /**
   * Add a request to the batch sequence
   * @param {string} method - GET, POST, PUT, DELETE, PATCH
   * @param {string} relativeUrl - Relative URL from graph base (e.g. /me/drive/root/children)
   * @param {Object} [headers={}] - Custom headers for this specific request
   * @param {string|Object} [body=''] - Optional request payload
   * @param {string[]} [dependsOn=[]] - IDs of requests that must complete first
   * @returns {string} The assigned ID for the request
   */
  addRequest(method, relativeUrl, headers = {}, body = '', dependsOn = []) {
    const id = (this.requests.length + 1).toString();
    
    const requestItem = {
      id: id,
      method: method.toUpperCase(),
      url: relativeUrl.startsWith('/') ? relativeUrl : `/${relativeUrl}`,
      headers: headers
    };

    if (body) {
      requestItem.body = typeof body === 'string' ? JSON.parse(body) : body;
    }

    if (dependsOn && dependsOn.length > 0) {
      requestItem.dependsOn = dependsOn;
    }

    this.requests.push(requestItem);
    logger.info(`Added request #${id} to batch: ${method} ${relativeUrl}`);
    return id;
  }

  /**
   * Executes the batch request to /$batch
   */
  async executeBatch() {
    if (this.requests.length === 0) {
      throw new Error("Batch is empty. Add at least one request.");
    }

    const batchUrl = `${auth.baseUrl}/$batch`;
    const payload = {
      requests: this.requests
    };

    logger.info(`Executing Batch API call with ${this.requests.length} requests...`);
    
    try {
      const response = await apiTester.executeRequest(
        'POST', 
        batchUrl, 
        { 'Content-Type': 'application/json' }, 
        JSON.stringify(payload)
      );

      return response;
    } catch (err) {
      logger.error("Batch execution failed: " + err.message);
      throw err;
    }
  }

  /**
   * Returns standard preset templates for demonstration/testing
   */
  getTemplates() {
    return [
      {
        name: "Retrieve User Details & Root Drive Items",
        description: "Fetch current user info and list top-level OneDrive files simultaneously.",
        build: () => {
          this.clearBatch();
          this.addRequest('GET', '/me');
          this.addRequest('GET', `/${auth.getDrivePathPrefix()}`);
          this.addRequest('GET', `/${auth.getDrivePathPrefix()}/root/children`);
        }
      },
      {
        name: "Create Folder then Check Membership",
        description: "Creates a 'BatchTest' folder and immediately queries the files, depending on creation.",
        build: () => {
          this.clearBatch();
          const createId = this.addRequest(
            'POST', 
            `/${auth.getDrivePathPrefix()}/root/children`, 
            { 'Content-Type': 'application/json' }, 
            {
              name: `BatchTest_${Math.floor(Math.random() * 1000)}`,
              folder: {},
              "@microsoft.graph.conflictBehavior": "rename"
            }
          );
          
          this.addRequest(
            'GET', 
            `/${auth.getDrivePathPrefix()}/root/children`, 
            {}, 
            '', 
            [createId]
          );
        }
      }
    ];
  }
}

export const batchBuilder = new BatchBuilder();
