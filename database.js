// ===========================================
// DATABASE.GS - Complete Database Management
// ===========================================

class Database {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.cache = {};
  }

  // ====================
  // CORE METHODS
  // ====================

  /**
   * Get sheet by name
   */
  getSheet(sheetName) {
    return this.ss.getSheetByName(sheetName);
  }

  /**
   * Get all data from a sheet as objects
   */
  getAll(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return [];

    const headers = data[0];
    const result = [];

    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((header, index) => {
        let value = data[i][index];
        
        // Convert Date objects to ISO strings for safe serialization
        if (value instanceof Date) {
          row[header] = value.toISOString();
        } else {
          row[header] = value || '';
        }
      });
      result.push(row);
    }

    return result;
  }

  /**
   * Get all data with headers
   */
  getAllWithHeader(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return { headers: [], rows: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return { headers: [], rows: [] };

    const headers = data[0];
    const rows = [];

    for (let i = 1; i < data.length; i++) {
      const row = {};
      headers.forEach((header, index) => {
        row[header] = data[i][index] || '';
      });
      rows.push(row);
    }

    return { headers, rows };
  }

  /**
   * Insert a new record
   */
  insert(sheetName, data) {
    try {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newRow = headers.map(header => data[header] || '');

      sheet.appendRow(newRow);
      return true;
    } catch (error) {
      console.error(`Error inserting into ${sheetName}:`, error);
      return false;
    }
  }

  /**
   * Update a record
   */
  update(sheetName, keyColumn, keyValue, updates) {
    try {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const keyIndex = headers.indexOf(keyColumn);

      if (keyIndex === -1) return false;

      for (let i = 1; i < data.length; i++) {
        if (data[i][keyIndex] == keyValue) {
          const updatedRow = headers.map((header, index) => {
            if (updates.hasOwnProperty(header)) {
              return updates[header];
            }
            return data[i][index];
          });

          sheet.getRange(i + 1, 1, 1, updatedRow.length).setValues([updatedRow]);
          return true;
        }
      }
      return false;
    } catch (error) {
      console.error(`Error updating ${sheetName}:`, error);
      return false;
    }
  }

  /**
   * Delete a record
   */
  delete(sheetName, keyColumn, keyValue) {
    try {
      const sheet = this.getSheet(sheetName);
      if (!sheet) return false;

      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const keyIndex = headers.indexOf(keyColumn);

      if (keyIndex === -1) return false;

      for (let i = data.length - 1; i > 0; i--) {
        if (data[i][keyIndex] == keyValue) {
          sheet.deleteRow(i + 1);
          return true;
        }
      }
      return false;
    } catch (error) {
      console.error(`Error deleting from ${sheetName}:`, error);
      return false;
    }
  }

  // ====================
  // USER METHODS
  // ====================

  /**
   * Get user by email
   */
  getUserByEmail(email) {
    const users = this.getAll('Users');
    return users.find(user => user.Email === email) || null;
  }

  /**
   * Get user by ID
   */
  getUserById(userId) {
  console.log('getUserById called with:', userId);
  
  const users = this.getAll('Users');
  console.log('Total users in database:', users.length);
  
  const user = users.find(user => user.UserID == userId) || null;
  
  if (user) {
    console.log('Found user details:', {
      UserID: user.UserID,
      Email: user.Email,
      FirstName: user.FirstName,
      LastName: user.LastName
    });
  } else {
    console.log('User not found. Available UserIDs:', users.map(u => u.UserID));
  }
  
  return user;
}

  /**
   * Get user by reset token
   */
  getUserByToken(token) {
    const users = this.getAll('Users');
    return users.find(user => user.Token === token) || null;
  }

  /**
   * Get all users
   */
  getAllUsers() {
    return this.getAll('Users');
  }

  /**
   * Add new user
   */
  addUser(userData) {
    return this.insert('Users', userData);
  }

  /**
   * Update user
   */
  updateUser(userId, updates) {
    return this.update('Users', 'UserID', userId, updates);
  }

  /**
   * Update user last login
   */
  updateUserLastLogin(userId) {
    return this.updateUser(userId, {
      LastLogin: new Date().toISOString()
    });
  }

  /**
   * Update user password
   */
  updateUserPassword(userId, hashedPassword) {
    return this.updateUser(userId, {
      Password: hashedPassword,
      Token: '' // Clear reset token
    });
  }

  /**
   * Update user status
   */
  updateUserStatus(userId, status) {
    return this.updateUser(userId, { Status: status });
  }

  /**
   * Delete user
   */
  deleteUser(userId) {
    return this.delete('Users', 'UserID', userId);
  }

  /**
   * Generate user ID
   */
  generateUserId(firstName, lastName) {
    const users = this.getAll('Users');
    
    // Format names
    const formatName = (name) => name.trim().replace(/\s+/g, '_').toLowerCase();
    const firstNameFormatted = formatName(firstName);
    const lastNameFormatted = formatName(lastName);
    
    const baseID = `${firstNameFormatted}_${lastNameFormatted}_`;
    
    // Find highest existing number
    let lastNumber = 0;
    users.forEach(user => {
      if (user.UserID && user.UserID.startsWith(baseID)) {
        const numPart = parseInt(user.UserID.split('_').pop(), 10);
        if (!isNaN(numPart) && numPart > lastNumber) {
          lastNumber = numPart;
        }
      }
    });
    
    return `${baseID}${(lastNumber + 1).toString().padStart(4, '0')}`;
  }

  // ====================
  // ROLE METHODS
  // ====================

  /**
 * Get all roles
 */
getAllRoles() {
  const roles = this.getAll('Roles');
  return roles.map(role => {
    // If RoleID is in old format (ROLE001), transform it
    let roleId = role.RoleID;
    let roleName = role.RoleName || role.Role;
    
    // If RoleID is in ROLE001 format, use RoleName for the ID
    if (roleId && roleId.startsWith('ROLE') && roleName) {
      roleId = roleName.toLowerCase().replace(/\s+/g, '_');
    }
    
    return {
      RoleID: roleId,
      Role: roleName,
      RoleName: roleName,
      Description: role.Description || ''
    };
  });
}

  /**
   * Get role by name
   */
  getRoleByName(roleName) {
    const roles = this.getAllRoles();
    return roles.find(role => role.RoleName === roleName) || null;
  }

  /**
   * Add new role
   */
  addRole(roleData) {
    return this.insert('Roles', roleData);
  }

  // ====================
  // DEPARTMENT METHODS
  // ====================

 /**
 * Get all departments
 */
getAllDepartments() {
  const departments = this.getAll('Department');
  return departments.map(dept => {
    // If DepartmentID is in old format (DEPT001), transform it
    let deptId = dept.DepartmentID;
    let deptName = dept.DepartmentName || dept.Department;
    
    // If DepartmentID is in DEPT001 format, use DepartmentName for the ID
    if (deptId && deptId.startsWith('DEPT') && deptName) {
      deptId = deptName.toLowerCase().replace(/\s+/g, '_');
    }
    
    return {
      DepartmentID: deptId,
      Department: deptName,
      Description: dept.Description || ''
    };
  });
}

  /**
   * Add new department
   */
  addDepartment(deptData) {
    return this.insert('Department', deptData);
  }

  // ====================
  // PERMISSION METHODS
  // ====================

  /**
   * Get permissions by role
   */
  getPermissionsByRole(roleId) {
    console.log('Getting permissions for role:', roleId);
    const permissions = this.getAll('Permissions');
    const menus = this.getAll('Menus');
    
    // Filter permissions for this role
    const rolePermissions = permissions.filter(p => p.RoleID === roleId);
    console.log('Found permissions:', rolePermissions.length);
    
    // Create a map of existing permissions
    const permissionMap = {};
    rolePermissions.forEach(p => {
      permissionMap[p.MenuID] = {
        RoleID: p.RoleID,
        MenuID: p.MenuID,
        CanView: p.CanView === true || p.CanView === 'TRUE' || p.CanView === 'true',
        CanAdd: p.CanAdd === true || p.CanAdd === 'TRUE' || p.CanAdd === 'true',
        CanEdit: p.CanEdit === true || p.CanEdit === 'TRUE' || p.CanEdit === 'true',
        CanDelete: p.CanDelete === true || p.CanDelete === 'TRUE' || p.CanDelete === 'true'
      };
    });
    
    // Ensure all menus have permission entries
    const allPermissions = [];
    menus.forEach(menu => {
      const menuId = menu.MenuID || '';
      if (permissionMap[menuId]) {
        allPermissions.push(permissionMap[menuId]);
      } else {
        allPermissions.push({
          RoleID: roleId,
          MenuID: menuId,
          CanView: false,
          CanAdd: false,
          CanEdit: false,
          CanDelete: false
        });
      }
    });
    
    return allPermissions;
  }

  /**
   * Update permissions
   */
  updatePermissions(permissions) {
    try {
      const sheet = this.getSheet('Permissions');
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      
      for (const permission of permissions) {
        let found = false;
        
        // Check if permission already exists
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === permission.RoleID && data[i][1] === permission.MenuID) {
            // Update existing
            sheet.getRange(i + 1, 3).setValue(permission.CanView ? "TRUE" : "FALSE");
            sheet.getRange(i + 1, 4).setValue(permission.CanAdd ? "TRUE" : "FALSE");
            sheet.getRange(i + 1, 5).setValue(permission.CanEdit ? "TRUE" : "FALSE");
            sheet.getRange(i + 1, 6).setValue(permission.CanDelete ? "TRUE" : "FALSE");
            found = true;
            break;
          }
        }
        
        if (!found) {
          // Insert new
          sheet.appendRow([
            permission.RoleID,
            permission.MenuID,
            permission.CanView ? "TRUE" : "FALSE",
            permission.CanAdd ? "TRUE" : "FALSE",
            permission.CanEdit ? "TRUE" : "FALSE",
            permission.CanDelete ? "TRUE" : "FALSE"
          ]);
        }
      }
      
      return { success: true };
    } catch (error) {
      console.error('Error updating permissions:', error);
      return { success: false, message: error.message };
    }
  }

  // ====================
  // MENU METHODS
  // ====================

  /**
   * Get all menus
   */
  getAllMenus() {
    return this.getAll('Menus');
  }

  /**
   * Get user menus based on role
   */
  getUserMenus(roleId) {
    try {
      console.log('Getting menus for role:', roleId);
      
      // Get all menus
      const allMenus = this.getAllMenus();
      console.log('Total menus in system:', allMenus.length);
      
      // Get permissions for this role
      const permissions = this.getPermissionsByRole(roleId);
      console.log('Permissions array length:', permissions.length);
      
      // Filter menus where user has view permission
      const accessibleMenus = [];
      
      // Create a map of view permissions
      const viewPermissionMap = {};
      permissions.forEach(p => {
        const menuId = p.MenuID || '';
        if (menuId) {
          viewPermissionMap[menuId] = p.CanView;
        }
      });
      
      console.log('View permission map:', viewPermissionMap);
      
      // Check each menu
      allMenus.forEach(menu => {
        const menuId = menu.MenuID || '';
        
        // Check if user has view permission for this menu
        const hasPermission = viewPermissionMap[menuId];
        
        if (hasPermission) {
          console.log('Menu granted:', menuId, menu.MenuName);
          accessibleMenus.push({
            MenuID: menu.MenuID,
            MenuName: menu.MenuName,
            Icon: menu.Icon,
            URL: menu.URL || menu.MenuID,
            DisplayOrder: parseInt(menu.DisplayOrder) || 0
          });
        } else {
          console.log('Menu denied:', menuId, menu.MenuName);
        }
      });
      
      // Sort by display order
      accessibleMenus.sort((a, b) => a.DisplayOrder - b.DisplayOrder);
      
      console.log('Final accessible menus:', accessibleMenus.length);
      accessibleMenus.forEach(m => console.log('  -', m.MenuName));
      
      return accessibleMenus;
    } catch (error) {
      console.error('Error in getUserMenus:', error);
      return [];
    }
  }

  /**
   * Get user permissions for all menus
   */
  getPermissionsMapByRole(roleId) { 
    try {
      const permissions = this.getPermissionsByRole(roleId);
      const permissionMap = {};
      
      permissions.forEach(p => {
        const menuId = p.MenuID || '';
        if (menuId) {
          permissionMap[menuId] = {
            RoleID: p.RoleID,
            MenuID: p.MenuID,
            PageID: p.MenuID, // For compatibility
            CanView: p.CanView === 'TRUE' || p.CanView === true,
            CanAdd: p.CanAdd === 'TRUE' || p.CanAdd === true,
            CanEdit: p.CanEdit === 'TRUE' || p.CanEdit === true,
            CanDelete: p.CanDelete === 'TRUE' || p.CanDelete === true
          };
        }
      });
      
      console.log('Permissions map created with', Object.keys(permissionMap).length, 'entries');
      return permissionMap;
    } catch (error) {
      console.error('Error in getPermissionsMapByRole:', error);
      return {};
    }
  }

  /**
   * Create default permissions for a new role (all false)
   */
  createDefaultPermissionsForRole(roleId, roleName) {
    try {
      const allMenus = this.getAllMenus();
      const sheet = this.getSheet('Permissions');
      
      if (!sheet) {
        console.error('Permissions sheet not found');
        return false;
      }
      
      // Create all false permissions for each menu
      allMenus.forEach(menu => {
        const menuId = menu.MenuID || '';
        if (menuId) {
          sheet.appendRow([
            roleId,
            menuId,
            "FALSE", // CanView
            "FALSE", // CanAdd
            "FALSE", // CanEdit
            "FALSE"  // CanDelete
          ]);
        }
      });
      
      console.log(`Created default permissions for role: ${roleName}`);
      return true;
      
    } catch (error) {
      console.error('Error creating default permissions:', error);
      return false;
    }
  }

  /**
   * Log user activity
   */
  logUserActivity(userId, activity, details = '') {
    try {
      let sheet = this.getSheet('UserActivity');
      if (!sheet) {
        // Create activity sheet if it doesn't exist
        sheet = this.ss.insertSheet('UserActivity');
        sheet.appendRow(['Timestamp', 'UserID', 'Activity', 'Details', 'IPAddress']);
      }
      
      const timestamp = new Date().toISOString();
      sheet.appendRow([timestamp, userId, activity, details, '']);
      
      return true;
    } catch (error) {
      console.error('Error logging activity:', error);
      return false;
    }
  }

  /**
   * Initialize activity tracking sheet
   */
  initializeActivitySheet() {
    try {
      const ss = this.ss;
      
      // Check if sheet exists
      let sheet = this.getSheet('UserActivity');
      
      if (!sheet) {
        // Create new sheet
        sheet = ss.insertSheet('UserActivity');
        
        // Add headers
        sheet.appendRow(['Timestamp', 'UserID', 'Activity', 'Details', 'IPAddress']);
        
        // Format headers
        const headers = sheet.getRange(1, 1, 1, 5);
        headers.setBackground('#2c3e50');
        headers.setFontColor('white');
        headers.setFontWeight('bold');
        
        // Set column widths
        sheet.setColumnWidth(1, 180); // Timestamp
        sheet.setColumnWidth(2, 150); // UserID
        sheet.setColumnWidth(3, 120); // Activity
        sheet.setColumnWidth(4, 250); // Details
        sheet.setColumnWidth(5, 120); // IPAddress
        
        console.log('Created UserActivity sheet');
      }
      
      return sheet;
    } catch (error) {
      console.error('Error initializing activity sheet:', error);
      return null;
    }
  }
}