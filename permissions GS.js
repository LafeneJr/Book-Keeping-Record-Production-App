// ===========================================
// SETTINGS PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check settings permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanEdit')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkSettingsPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has settings permissions
    const settingsPerms = session.permissions && session.permissions["settings"];
    if (!settingsPerms) {
      throw new Error("You don't have permission to access settings.");
    }
    
    // Check specific action permission
    const hasPermission = settingsPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} settings.`);
    }
    
    return session;
  } catch (error) {
    console.error('Settings permission check error:', error);
    throw error;
  }
}

// ===========================================
// SETTINGS FUNCTIONS (For Settings.html)
// ===========================================

/**
 * Get all roles for settings
 */
function getSettingsRoles(token) {
  try {
    // Check view permission
    checkSettingsPermission(token, "CanView");
    
    const db = new Database();
    const roles = db.getAllRoles();
    
    // Return just role names for dropdown
    return roles.map(role => role.RoleName || role.Role || '').filter(r => r);
  } catch (error) {
    console.error('Error getting roles for settings:', error);
    throw error;
  }
}

/**
 * Get permissions by role
 */
function getPermissionsByRole(token, roleName) {
  try {
    // Check view permission
    checkSettingsPermission(token, "CanView");
    
    const db = new Database();
    const roleId = roleName.toLowerCase().replace(/\s+/g, '_');
    return db.getPermissionsByRole(roleId);
  } catch (error) {
    console.error('Error getting permissions:', error);
    throw error;
  }
}

/**
 * Update permissions
 */
function updatePermissions(token, permissions) {
  try {
    // Check edit permission
    checkSettingsPermission(token, "CanEdit");
    
    const db = new Database();
    return db.updatePermissions(permissions);
  } catch (error) {
    console.error('Error updating permissions:', error);
    throw error;
  }
}