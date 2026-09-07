// ===========================================
// USERS PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check users permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanAdd', 'CanEdit', 'CanDelete')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkUsersPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has users permissions
    const usersPerms = session.permissions && session.permissions["users"];
    if (!usersPerms) {
      throw new Error("You don't have permission to access user management.");
    }
    
    // Check specific action permission
    const hasPermission = usersPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} users.`);
    }
    
    return session;
  } catch (error) {
    console.error('Users permission check error:', error);
    throw error;
  }
}


// ===========================================
// USER MANAGEMENT FUNCTIONS (For Users.html)
// ===========================================

/**
 * Get all users (for Users page)
 */
function getUsers(token) {
  try {
    // Check view permission
    checkUsersPermission(token, "CanView");
        
    const db = new Database();
    const users = db.getAllUsers();
    
    if (!users || !Array.isArray(users)) {
      return [];
    }
    
    // Format for frontend - CRITICAL: Convert Date objects to strings
    const formattedUsers = users.map(user => {
      // Handle joinDate - convert Date to string
      let joinDate = user.JoinDate || '';
      if (joinDate instanceof Date) {
        joinDate = joinDate.toISOString().split('T')[0]; // YYYY-MM-DD format
      }
      
      // Handle lastLogin - ensure it's a string
      let lastLogin = user.LastLogin || '';
      if (lastLogin instanceof Date) {
        lastLogin = lastLogin.toISOString();
      }
      
      return {
        id: user.UserID || '',
        email: user.Email || '',
        firstName: user.FirstName || '',
        lastName: user.LastName || '',
        role: user.Role || '',
        department: user.Department || '',
        joinDate: joinDate, // Now a string, not a Date object
        status: user.Status || 'Active',
        phone: user.Phone || '',
        address: user.Address || '',
        lastLogin: lastLogin || 'Never'
      };
    });
    
    console.log('Final formatted users to return:', formattedUsers);
    return formattedUsers;
  } catch (error) {
    console.error('Error getting users:', error);
    throw error;
  }
}

/**
 * Get all roles (for Users page)
 */
function getRoles() {
  try {
    const db = new Database();
    const roles = db.getAllRoles();
    
    // Extract role names
    return roles.map(role => role.RoleName || role.Role || '').filter(r => r);
  } catch (error) {
    console.error('Error getting roles:', error);
    return [];
  }
}

/**
 * Get all departments (for Users page)
 */
function getDepartments() {
  try {
    const db = new Database();
    const departments = db.getAllDepartments();
    
    // Extract department names
    return departments.map(dept => dept.DepartmentName || dept.Department || '').filter(d => d);
  } catch (error) {
    console.error('Error getting departments:', error);
    return [];
  }
}

/**
 * Generate user ID
 */
function generateUserId(firstName, lastName) {
  try {
    const db = new Database();
    return db.generateUserId(firstName, lastName);
  } catch (error) {
    console.error('Error generating user ID:', error);
    return 'USER_' + Math.floor(Math.random() * 10000);
  }
}

/**
 * Create new user
 */
function createUser(token, userData) {
  try {
    // Check add permission
    checkUsersPermission(token, "CanAdd");
    
    return Authentication.createUser(userData, 'system');
  } catch (error) {
    console.error('Error creating user:', error);
    throw error;
  }
}

/**
 * Update user
 */
function updateUser(token, userData) {
  try {
    // Check edit permission
    checkUsersPermission(token, "CanEdit");

    const db = new Database();
    const updates = {
      FirstName: userData.firstName,
      LastName: userData.lastName,
      Role: userData.role,
      Department: userData.department || '',
      Status: userData.status || 'Active',
      Phone: userData.phone || '',
      Address: userData.address || '',
      JoinDate: userData.joinDate || ''
    };
    
    const success = db.updateUser(userData.id, updates);
    
    if (success) {
      return { success: true, message: 'User updated successfully' };
    } else {
      return { success: false, message: 'User not found or update failed' };
    }
  } catch (error) {
    console.error('Error updating user:', error);
    throw error;
  }
}

/**
 * Delete user
 */
function deleteUser(token, userId) {
  try {
    // Check delete permission
    checkUsersPermission(token, "CanDelete");

    const db = new Database();
    const user = db.getUserById(userId);
    
    if (!user) {
      return 'not_found';
    }
    
    // Get current user from session
    const session = Authentication.validateSession(token);
    if (session && user.Email === session.userEmail) {
      return 'self_error';
    }
    
    const success = db.deleteUser(userId);
    
    if (success) {
      return 'success';
    } else {
      return 'error';
    }
  } catch (error) {
    console.error('Error deleting user:', error);
    throw error;
  }
}

/**
 * Send password reset email
 */
function sendPasswordReset(token, userId) {
  try {
    // Check edit permission (resetting password is an edit action)
    checkUsersPermission(token, "CanEdit");
    
    const db = new Database();
    const user = db.getUserById(userId);
    
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    // Get current user from session to ensure they're not trying to reset their own password
    const session = Authentication.validateSession(token);
    if (session && user.Email === session.userEmail) {
      return { 
        success: false, 
        message: 'You cannot send password reset to yourself. Use "Change Password" in your profile instead.' 
      };
    }
    
    return Authentication.requestPasswordReset(user.Email);
  } catch (error) {
    console.error('Error sending password reset:', error);
    return { success: false, message: error.message };
  }
}


/**
 * Add new role
 */
function addNewRole(roleName) {
  try {
    const db = new Database();
    
    // Generate role ID from role name (lowercase with underscores)
    const roleId = roleName.toLowerCase().replace(/\s+/g, '_');
    
    // Check if role already exists (case-insensitive)
    const roles = db.getAll('Roles');
    const existingRole = roles.find(role => 
      role.RoleName.toLowerCase() === roleName.toLowerCase() || 
      role.RoleID.toLowerCase() === roleId
    );
    
    if (existingRole) {
      return { 
        success: false, 
        message: `Role "${roleName}" already exists` 
      };
    }
    
    // Create role data
    const roleData = {
      RoleID: roleId,
      RoleName: roleName,
      Description: `Added on ${new Date().toLocaleDateString()}`
    };
    
    // Add role to Roles sheet
    db.insert("Roles", roleData);
    
    // Get all menus
    const menus = db.getAll('Menus');
    
    // Create default permissions (all false)
    if (menus.length > 0) {
      const permissions = menus.map(menu => ({
        RoleID: roleId,
        MenuID: menu.MenuID,
        CanView: false,
        CanAdd: false,
        CanEdit: false,
        CanDelete: false
      }));
      
      // Insert all permissions at once
      permissions.forEach(permission => {
        db.insert("Permissions", {
          RoleID: permission.RoleID,
          MenuID: permission.MenuID,
          CanView: permission.CanView ? "TRUE" : "FALSE",
          CanAdd: permission.CanAdd ? "TRUE" : "FALSE",
          CanEdit: permission.CanEdit ? "TRUE" : "FALSE",
          CanDelete: permission.CanDelete ? "TRUE" : "FALSE"
        });
      });
    }
    
    return { 
      success: true, 
      message: `Role "${roleName}" added with default permissions (all false)`,
      roleId: roleId
    };
    
  } catch (error) {
    console.error('Error adding role:', error);
    return { 
      success: false, 
      message: 'Error adding role: ' + error.message 
    };
  }
}

/**
 * Add new department
 */
function addNewDepartment(deptName) {
  try {
    const db = new Database();
    
    // Generate department ID from department name
    const deptId = deptName.toLowerCase().replace(/\s+/g, '_');
    
    // Check if department already exists (case-insensitive)
    const departments = db.getAll('Department');
    const existingDept = departments.find(dept => 
      dept.DepartmentName && dept.DepartmentName.toLowerCase() === deptName.toLowerCase()
    );
    
    if (existingDept) {
      return { 
        success: false, 
        message: 'Department already exists' 
      };
    }
    
    // Create department data
    const deptData = {
      DepartmentID: deptId,
      DepartmentName: deptName,
      Description: ""
    };
    
    db.insert("Department", deptData);
    
    return { 
      success: true, 
      message: 'Department added successfully',
      deptId: deptId
    };
    
  } catch (error) {
    console.error('Error adding department:', error);
    return { 
      success: false, 
      message: 'Error adding department: ' + error.message 
    };
  }
}