/**
 * Microsoft Graph OneDrive Explorer - Sharing & Permissions Module
 * Manages links, invites, permissions, version histories, and checkout states.
 */

import { auth } from './auth.js';
import { logger } from './logs.js';
import { apiTester } from './explorer.js';

class SharingManager {
  /**
   * List sharing links & permissions of a specific file/folder
   */
  async listPermissions(itemId) {
    logger.info(`Fetching permissions for item: ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/permissions`;
    
    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        return data.value || [];
      }
      throw new Error(`Failed to list permissions. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Error loading permissions for item ${itemId}: ${err.message}`);
      throw err;
    }
  }

  /**
   * Create a sharing link for a file or folder
   */
  async createSharingLink(itemId, type = 'view', scope = 'anonymous', password = '', expiration = '') {
    logger.info(`Creating sharing link for item ${itemId}. Type: ${type}, Scope: ${scope}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/createLink`;
    
    const body = {
      type: type, // 'view', 'edit', 'embed'
      scope: scope // 'anonymous', 'organization', 'users'
    };

    if (password) {
      body.password = password;
    }

    if (expiration) {
      body.expirationDateTime = new Date(expiration).toISOString();
    }

    try {
      const result = await apiTester.executeRequest('POST', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if ([200, 201].includes(result.status)) {
        const data = JSON.parse(result.body);
        logger.success(`Sharing link created successfully. Link: ${data.link.webUrl}`);
        return data;
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Failed to create sharing link: ${err.message}`);
      throw err;
    }
  }

  /**
   * Delete / Remove a permission or link
   */
  async deletePermission(itemId, permissionId) {
    logger.info(`Removing permission ${permissionId} from item ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/permissions/${permissionId}`;
    
    try {
      const result = await apiTester.executeRequest('DELETE', url);
      if ([204, 200].includes(result.status)) {
        logger.success(`Permission ${permissionId} removed successfully.`);
        return true;
      }
      throw new Error(`Failed to delete permission. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Failed to delete permission: ${err.message}`);
      throw err;
    }
  }

  /**
   * Invite users to share a file or folder
   */
  async inviteUsers(itemId, roles = ['read'], recipients = [], requireSignIn = true, sendInvitation = true, message = '') {
    logger.info(`Sending share invites for item ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/invite`;
    
    const body = {
      recipients: recipients.map(email => ({ email: email.trim() })),
      roles: roles, // ['read'] or ['write']
      requireSignIn: requireSignIn,
      sendInvitation: sendInvitation,
      message: message
    };

    try {
      const result = await apiTester.executeRequest('POST', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if ([200, 201].includes(result.status)) {
        const data = JSON.parse(result.body);
        logger.success("Users invited and permissions granted successfully.");
        return data.value || [];
      }
      throw new Error(result.body);
    } catch (err) {
      logger.error(`Failed to invite users: ${err.message}`);
      throw err;
    }
  }

  /**
   * List versions of a OneDrive file
   */
  async listVersions(itemId) {
    logger.info(`Fetching version history for file: ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/versions`;
    
    try {
      const result = await apiTester.executeRequest('GET', url);
      if (result.status === 200) {
        const data = JSON.parse(result.body);
        return data.value || [];
      }
      throw new Error(`Failed to fetch version history. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Error loading versions for item ${itemId}: ${err.message}`);
      throw err;
    }
  }

  /**
   * Restore a previous file version
   */
  async restoreVersion(itemId, versionId) {
    logger.info(`Restoring version ${versionId} of file ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/versions/${versionId}/restoreVersion`;
    
    try {
      const result = await apiTester.executeRequest('POST', url);
      if ([200, 204].includes(result.status)) {
        logger.success(`Version ${versionId} restored successfully.`);
        return true;
      }
      throw new Error(`Failed to restore version. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Error restoring version: ${err.message}`);
      throw err;
    }
  }

  /**
   * Delete a previous file version
   */
  async deleteVersion(itemId, versionId) {
    logger.info(`Deleting version ${versionId} of file ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/versions/${versionId}`;
    
    try {
      const result = await apiTester.executeRequest('DELETE', url);
      if ([200, 204].includes(result.status)) {
        logger.success(`Version ${versionId} deleted successfully.`);
        return true;
      }
      throw new Error(`Failed to delete version. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Error deleting version: ${err.message}`);
      throw err;
    }
  }

  /**
   * Checkout a file
   */
  async checkoutFile(itemId) {
    logger.info(`Checking out file: ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/checkout`;
    
    try {
      const result = await apiTester.executeRequest('POST', url);
      if ([200, 204].includes(result.status)) {
        logger.success("File checked out successfully.");
        return true;
      }
      throw new Error(`Checkout failed. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Failed to checkout file: ${err.message}`);
      throw err;
    }
  }

  /**
   * Checkin a file
   */
  async checkinFile(itemId, checkInAs = 'published', comment = '') {
    logger.info(`Checking in file: ${itemId}`);
    const url = `${auth.baseUrl}/${auth.getDrivePathPrefix()}/items/${itemId}/checkin`;
    
    const body = {
      comment: comment,
      checkInAs: checkInAs // 'minor' or 'published' or 'unspecified'
    };

    try {
      const result = await apiTester.executeRequest('POST', url, { 'Content-Type': 'application/json' }, JSON.stringify(body));
      if ([200, 204].includes(result.status)) {
        logger.success("File checked in successfully.");
        return true;
      }
      throw new Error(`Checkin failed. Status: ${result.status}`);
    } catch (err) {
      logger.error(`Failed to checkin file: ${err.message}`);
      throw err;
    }
  }
}

export const sharingManager = new SharingManager();
