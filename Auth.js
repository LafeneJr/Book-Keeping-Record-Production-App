// ===========================================
// AUTH.GS - Authentication Module
// ===========================================

const Authentication = (function() {
  // Cache expiration time (3 hours)
  const CACHE_EXPIRATION = 10800;

    /**
     * Login user
     */
    function login(email, password) {
      try {
        console.log('Login attempt for email:', email);
        const db = new Database();
        
        // Find user by email
        const user = db.getUserByEmail(email);
        console.log('User found:', user ? 'Yes' : 'No');
        
        if (!user) {
          return { success: false, message: 'User not found' };
        }
        
        // Check if user is active
        console.log('User status:', user.Status);
        if (user.Status !== 'Active') {
          return { success: false, message: 'Account has been deactivated. Please contact administrator.' };
        }
        
        // Verify password
        console.log('Password field exists:', 'Password' in user);
        
        if (!verifyPassword(password, user.Password)) {
          return { success: false, message: 'Invalid password' };
        }
        
        // Generate session token
        const sessionToken = generateUUID();
        
        // Get user permissions and menus
        const roleId = user.Role.toLowerCase().replace(/\s+/g, '_');
        console.log('Role:', user.Role, 'Role ID:', roleId);
        
        // FIXED: Use the correct method names
        const userPermissions = db.getPermissionsMapByRole(roleId); // Changed from getUserPermissions
        const userMenus = db.getUserMenus(roleId);
        
        console.log('Permissions loaded:', userPermissions ? 'Yes' : 'No');
        console.log('Number of permissions:', Object.keys(userPermissions || {}).length);
        console.log('Menus loaded:', userMenus ? 'Yes' : 'No');
        console.log('Number of menus:', (userMenus || []).length);
        
        // Create session data
        const sessionData = {
          userId: user.UserID,
          userEmail: user.Email,
          userFirstName: user.FirstName,
          userLastName: user.LastName,
          userRole: user.Role,
          roleId: roleId,
          permissions: userPermissions, // Changed variable name
          menus: userMenus, // Changed variable name
          profileImage: formatProfileImageUrl(user.ProfileImage || ''),
          timestamp: new Date().getTime()
        };
        
        // Save session to cache
        const cache = CacheService.getUserCache();
        cache.put(sessionToken, JSON.stringify(sessionData), CACHE_EXPIRATION);
        
        // Update last login
        db.updateUserLastLogin(user.UserID);
        
        // Log the login activity
        db.logUserActivity(user.UserID, 'Login', 'Successful login from web');
        
        // Return success with user info
        return {
          success: true,
          token: sessionToken,
          user: {
            id: user.UserID,
            email: user.Email,
            firstName: user.FirstName,
            lastName: user.LastName,
            role: user.Role,
            profileImage: user.ProfileImage || ''
          },
          menus: userMenus, // Changed variable name
          permissions: userPermissions // Changed variable name
        };
        
      } catch (error) {
        console.error('Login error details:', {
          message: error.message,
          stack: error.stack,
          lineNumber: error.lineNumber
        });
        return { success: false, message: 'An error occurred during login: ' + error.message };
      }
    }

  /**
   * Validate session token
   */
    function validateSession(token) {
    if (!token) return null;
    
    try {
      const cache = CacheService.getUserCache();
      const sessionData = cache.get(token); 
      
      if (!sessionData) return null;
      
      const session = JSON.parse(sessionData);
      const currentTime = new Date().getTime();
      
      // Check if session has expired (3 hours)
      if (currentTime - session.timestamp > 3 * 60 * 60 * 1000) {
        cache.remove(token);
        return null;
      }
      
      // Validate session structure
      if (!session.userId || !session.userEmail || !session.userRole) {
        console.error('Invalid session structure:', session);
        cache.remove(token);
        return null;
      }
      
      // Update session timestamp
      session.timestamp = currentTime;
      cache.put(token, JSON.stringify(session), CACHE_EXPIRATION);
      
      return session;
    } catch (error) {
      console.error('Session validation error:', error);
      return null;
    }
  }

  /**
   * Logout user
   */
  function logout(token) {
    try {
      if (!token) return { success: true };
      
      const cache = CacheService.getUserCache();
      cache.remove(token);
      return { success: true };
    } catch (error) {
      console.error('Logout error:', error);
      return { success: false, message: 'Logout failed' };
    }
  }

  /**
   * Request password reset
   */
  function requestPasswordReset(email) {
    try {
      const db = new Database();
      const user = db.getUserByEmail(email);
      
      if (!user) {
        return { success: false, message: 'Email not found in our records' };
      }
      
      // Generate reset token
      const resetToken = generateRandomToken();
      
      // Save token to user record
      db.updateUser(user.UserID, { Token: resetToken });
      
      // Send reset email
      const scriptUrl = getScriptUrl();
      const resetUrl = `${scriptUrl}?page=reset-password&email=${email}&token=${resetToken}`;
      
      MailApp.sendEmail({
        to: email,
        subject: "Password Reset Request - Book Keeping Record",
        body: `Hello ${user.FirstName},\n\nYou have requested to reset your password. Please click the link below to reset your password:\n\n${resetUrl}\n\nThis link will expire in 24 hours.\n\nIf you did not request this, please ignore this email.`
      });
      
      return { 
        success: true, 
        message: 'If the email exists in our system, a password reset link has been sent' 
      };
    } catch (error) {
      console.error('Reset request error:', error);
      return { success: false, message: 'An error occurred while processing your request' };
    }
  }

  /**
   * Reset password with token
   */
  function resetPassword(email, newPassword, token) {
    try {
      const db = new Database();
      const user = db.getUserByEmail(email);
      
      if (!user) {
        return { success: false, message: 'User not found' };
      }
      
      // Verify token
      if (user.Token !== token) {
        return { success: false, message: 'Invalid or expired reset token' };
      }
      
      // Validate new password
      if (!isValidPassword(newPassword)) {
        return { 
          success: false, 
          message: 'Password must be at least 8 characters long and contain uppercase, lowercase, and numbers' 
        };
      }
      
      // Hash new password
      const hashedPassword = hashPassword(newPassword);
      
      // Update user password and clear token
      db.updateUserPassword(user.UserID, hashedPassword);
      
      return { success: true, message: 'Password has been reset successfully' };
    } catch (error) {
      console.error('Password reset error:', error);
      return { success: false, message: 'An error occurred while resetting your password' };
    }
  }

  /**
   * Change password (when already logged in)
   */
  function changePassword(userId, currentPassword, newPassword) {
    try {
      const db = new Database();
      const user = db.getUserById(userId);
      
      if (!user) {
        return { success: false, message: 'User not found' };
      }
      
      // Verify current password
      if (!verifyPassword(currentPassword, user.Password)) {
        return { success: false, message: 'Current password is incorrect' };
      }
      
      // Validate new password
      if (!isValidPassword(newPassword)) {
        return { 
          success: false, 
          message: 'Password must be at least 8 characters long and contain uppercase, lowercase, and numbers' 
        };
      }
      
      // Hash new password
      const hashedPassword = hashPassword(newPassword);
      
      // Update user password
      db.updateUserPassword(userId, hashedPassword);
      
      return { success: true, message: 'Password changed successfully' };
    } catch (error) {
      console.error('Password change error:', error);
      return { success: false, message: 'An error occurred while changing your password' };
    }
  }

  /**
   * Create new user (admin function)
   */
  function createUser(userData, createdBy) {
    try {
      const db = new Database();
      
      // Check if email already exists
      const existingUser = db.getUserByEmail(userData.email);
      if (existingUser) {
        return { success: false, message: 'Email already in use' };
      }
      
      // Generate user ID
      const userId = db.generateUserId(userData.firstName, userData.lastName);
      
      // Generate reset token for password setup
      const resetToken = generateRandomToken();
      
      // Create new user object
      const newUser = {
        UserID: userId,
        Email: userData.email,
        FirstName: userData.firstName,
        LastName: userData.lastName,
        Role: userData.role,
        Department: userData.department || '',
        Status: userData.status || 'Active',
        Phone: userData.phone || '',
        Address: userData.address || '',
        JoinDate: userData.joinDate || new Date().toISOString().split('T')[0],
        Token: resetToken,
        LastLogin: ''
        // Password will be set when user sets up password
      };
      
      // Add user to database
      const success = db.addUser(newUser);
      
      if (!success) {
        return { success: false, message: 'Failed to create user' };
      }
      
      // Send password setup email
      const scriptUrl = getScriptUrl();
      const setupUrl = `${scriptUrl}?page=reset-password&email=${userData.email}&token=${resetToken}&mode=setup`;
      
      MailApp.sendEmail({
        to: userData.email,
        subject: "Account Setup - Book Keeping Record",
        body: `Hello ${userData.firstName},\n\nYour account has been created. Please click the link below to set up your password:\n\n${setupUrl}\n\nThis link will expire in 24 hours.`
      });
      
      return {
        success: true,
        message: 'User created successfully. Password setup email sent.',
        user: {
          id: userId,
          email: userData.email,
          firstName: userData.firstName,
          lastName: userData.lastName,
          role: userData.role
        }
      };
    } catch (error) {
      console.error('Create user error:', error);
      return { success: false, message: 'An error occurred while creating the user' };
    }
  }

  /**
   * Activate user account
   */
  function activateUser(userId) {
    try {
      const db = new Database();
      const success = db.updateUserStatus(userId, 'Active');
      
      if (success) {
        return { success: true, message: 'User activated successfully' };
      } else {
        return { success: false, message: 'User not found' };
      }
    } catch (error) {
      console.error('Activate user error:', error);
      return { success: false, message: 'An error occurred while activating the user' };
    }
  }

  /**
   * Deactivate user account
   */
  function deactivateUser(userId) {
    try {
      const db = new Database();
      const success = db.updateUserStatus(userId, 'Inactive');
      
      if (success) {
        return { success: true, message: 'User deactivated successfully' };
      } else {
        return { success: false, message: 'User not found' };
      }
    } catch (error) {
      console.error('Deactivate user error:', error);
      return { success: false, message: 'An error occurred while deactivating the user' };
    }
  }

  /**
   * Check if user has permission
   */
  function hasPermission(session, menuId, action = 'view') {
    if (!session || !session.permissions) return false;
    
    const permission = session.permissions[menuId];
    if (!permission) return false;
    
    switch (action) {
      case 'view': return permission.CanView;
      case 'add': return permission.CanAdd;
      case 'edit': return permission.CanEdit;
      case 'delete': return permission.CanDelete;
      default: return false;
    }
  }

  // Return public API
  return {
    login: login,
    validateSession: validateSession,
    logout: logout,
    requestPasswordReset: requestPasswordReset,
    resetPassword: resetPassword,
    changePassword: changePassword,
    createUser: createUser,
    activateUser: activateUser,
    deactivateUser: deactivateUser,
    hasPermission: hasPermission
  };
})();