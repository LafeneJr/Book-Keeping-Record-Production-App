// ===========================================
// TEMPLATE.GS - Main Application Controller
// ===========================================

// Global constants
const CACHE_EXPIRATION_TIME = 10800; // 3 hours in seconds

/**
 * Main request handler - Routes to appropriate pages
 */
function doGet(e) {
  if (!e) {
    e = { parameter: {} };
  }
  try {
    const page = e.parameter.page || 'login';
    const token = e.parameter.token || '';
    
    // Handle public pages (no authentication required)
    if (page === 'login' || page === 'reset-password') {
      return handlePublicPage(page, e);
    }
        
    if (!token) {
      console.error('No token provided for protected page:', page);
      return HtmlService.createTemplateFromFile('login')
        .evaluate()
        .setTitle('Book Keeping Record - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // For protected pages, validate session
    const sessionData = Authentication.validateSession(token);
        
    if (!sessionData || !sessionData.userId) {
      console.error('Invalid session or missing userId. Token:', token, 'Session data:', sessionData);
      // Invalid session, redirect to login
      return HtmlService.createTemplateFromFile('login')
        .evaluate()
        .setTitle('Book Keeping Record - Login')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Check if user has permission for this page
    if (!checkPagePermission(sessionData, page)) {
      // No permission, redirect to dashboard
      return createMainTemplate('dashboard', sessionData, token);
    }
    
    // User has permission, show requested page
    return createMainTemplate(page, sessionData, token);
    
  } catch (error) {
    console.error('doGet error:', error);
    // Fallback to login on error
    return HtmlService.createTemplateFromFile('login')
      .evaluate()
      .setTitle('Book Keeping Record - Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Handle public pages (login, reset-password)
 */
function handlePublicPage(page, e) {
  if (page === 'login') {
    return HtmlService.createTemplateFromFile('login')
      .evaluate()
      .setTitle('Book Keeping Record - Login')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  if (page === 'reset-password') {
    const template = HtmlService.createTemplateFromFile('reset-password');
    template.token = e.parameter.token || '';
    template.email = e.parameter.email || '';
    template.mode = e.parameter.mode || 'reset';
    return template
      .evaluate()
      .setTitle('Book Keeping Record - Reset Password')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Default to login
  return HtmlService.createTemplateFromFile('login')
    .evaluate()
    .setTitle('Book Keeping Record - Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Create main template for authenticated users
 */
function createMainTemplate(page, sessionData, token) {
  // Map pages to their template files
  const pageTemplates = {
    'dashboard': 'index',
    'inventory': 'inventory',
    'suppliers': 'suppliers',
    'customers': 'customers',
    'purchases': 'purchases',
    'sales': 'sales',
    'productions': 'productions',
    'payments': 'payments',
    'users': 'users',
    'settings': 'settings',
    'profile': 'profile'
  };
  
  const contentTemplate = pageTemplates[page] || 'index';
  
  // Get fresh user data with profile image
  const db = new Database();
  const user = db.getUserById(sessionData.userId);

  const roleId = user.Role.toLowerCase().replace(/\s+/g, '_');
  const permissions = db.getPermissionsMapByRole(roleId); // Fresh permissions
  const menus = db.getUserMenus(roleId); // Fresh menus

   // Update sessionData with fresh data
  sessionData.permissions = permissions;
  sessionData.menus = menus;
  
  // Update cache with fresh data
  const cache = CacheService.getUserCache();
  sessionData.timestamp = new Date().getTime();
  cache.put(token, JSON.stringify(sessionData), CACHE_EXPIRATION_TIME);

  // Get specific page permissions
  const pagePermissions = sessionData.permissions && sessionData.permissions[page] || {};
  
  
  // Update sessionData with properly formatted profile image
  let profileImage = formatProfileImageUrl(user.ProfileImage);
  if (!profileImage) {
    profileImage = getDefaultAvatar(user.Email, user.FirstName);
  }
  sessionData.profileImage = profileImage;
  
  // Create main template
  const template = HtmlService.createTemplateFromFile('Template');
  template.contentTemplate = contentTemplate;
  template.currentPage = page;
  template.userData = sessionData;
  template.token = token;
  // Pass page permissions to template
  template.pagePermissions = pagePermissions;
  template.canView = pagePermissions.CanView || false;
  template.canAdd = pagePermissions.CanAdd || false;
  template.canEdit = pagePermissions.CanEdit || false;
  template.canDelete = pagePermissions.CanDelete || false;
  
  return template
    .evaluate()
    .setTitle('Book Keeping Record - ' + page.charAt(0).toUpperCase() + page.slice(1))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Check if user has permission for a page
 */
function checkPagePermission(sessionData, page) {
  try {
    // For settings, allow both Admin and Super Admin
    if (page === 'settings') {
      const allowedRoles = ['Super Admin', 'Admin'];
      if (!allowedRoles.includes(sessionData.userRole)) {
        return false;
      }
    }
    
    // Default: allow access to all other pages
    return true;
  } catch (error) {
    console.error('Permission check error:', error);
    return false;
  }
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Get script URL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}



// ===========================================
// AUTHENTICATION FUNCTIONS (For Login.html)
// ===========================================

/**
 * Login function (wrapper for Authentication.login)
 */
function login(email, password) {
  return Authentication.login(email, password);
}

/**
 * Logout function (wrapper for Authentication.logout)
 */
function logout(sessionToken) {
  let authResult;
  
  try {
    // Call original logout
    authResult = Authentication.logout(sessionToken);
    
    // Handle different return types
    if (authResult === undefined || authResult === null) {
      // Authentication.logout doesn't return anything
      authResult = { success: true };
    } else if (typeof authResult === 'boolean') {
      // Authentication.logout returns boolean
      authResult = { success: authResult };
    }
    // If it's already an object, keep it as is
    
  } catch (error) {
    console.error('Authentication.logout error:', error);
    authResult = { 
      success: false, 
      message: 'Authentication logout failed: ' + error.message 
    };
  }
  
  // Always clear cache regardless of Authentication.logout result
  if (sessionToken) {
    try {
      const cache = CacheService.getUserCache();
      cache.remove(sessionToken);
      
      const scriptCache = CacheService.getScriptCache();
      scriptCache.remove('session_' + sessionToken);
      
      console.log('Cache cleared for token:', sessionToken);
    } catch (cacheError) {
      console.error('Cache clearing error:', cacheError);
    }
  }
  
  // Return combined result
  return {
    success: authResult.success || false,
    message: authResult.message || (authResult.success ? 'Logged out successfully' : 'Logout completed with issues'),
    timestamp: new Date().toISOString(),
    authResult: authResult // Include original result for debugging
  };
}

/**
 * Request password reset (wrapper)
 */
function requestPasswordReset(email) {
  return Authentication.requestPasswordReset(email);
}

/**
 * Reset user password (wrapper)
 */
function resetUserPassword(email, newPassword, token) {
  return Authentication.resetPassword(email, newPassword, token);
}

// ===========================================
// HELPER FUNCTIONS
// ===========================================

/**
 * Validate session (wrapper)
 */
function validateSession(token) {
  return Authentication.validateSession(token);
}

/**
 * Get current user info
 */
function getCurrentUser(sessionToken) {
  const session = Authentication.validateSession(sessionToken);
  if (!session) {
    return { valid: false, message: 'Invalid session' };
  }
  
  return {
    valid: true,
    user: {
      id: session.userId,
      email: session.userEmail,
      firstName: session.userFirstName,
      lastName: session.userLastName,
      role: session.userRole
    }
  };
}

/**
 * Get user permissions for frontend
 */
function getUserPermissions(sessionToken) {
  const session = Authentication.validateSession(sessionToken);
  if (!session) {
    return { valid: false, message: 'Invalid session' };
  }
  
  return {
    valid: true,
    permissions: session.permissions || {}
  };
}

/**
 * Get user menus for frontend
 */
function getUserMenus(sessionToken) {
  const session = Authentication.validateSession(sessionToken);
  if (!session) {
    return { valid: false, message: 'Invalid session' };
  }
  
  return {
    valid: true,
    menus: session.menus || []
  };
}

// ===========================================
// PROFILE FUNCTIONS (For Profile.html)
// ===========================================

/**
 * Get current user profile (wrapper)
 */
function getCurrentUserProfile(sessionToken) {
  return Profile.getCurrentUserProfile(sessionToken);
}

/**
 * Update user profile (wrapper)
 */
function updateUserProfile(sessionToken, profileData) {
  return Profile.updateUserProfile(sessionToken, profileData);
}

/**
 * Upload profile picture (wrapper)
 */
function uploadProfilePicture(sessionToken, imageData) {
  return Profile.uploadProfilePicture(sessionToken, imageData);
}

/**
 * Change user password (wrapper)
 */
function changeUserPassword(sessionToken, currentPassword, newPassword) {
  return Profile.changeUserPassword(sessionToken, currentPassword, newPassword);
}

/**
 * Format profile image URL for display
 */
function formatProfileImageUrl(imageData) {
  if (!imageData) return '';
  
  // If it's already a full URL with thumbnail
  if (imageData.includes('drive.google.com/thumbnail')) {
    return imageData + '&t=' + Date.now(); // Add timestamp to prevent caching
  }
  
  // If it's a base64 image
  if (imageData.startsWith('data:image')) {
    return imageData;
  }
  
  // If it's a Drive file ID or URL
  const fileId = extractFileIdFromUrl(imageData);
  if (fileId) {
    return `https://drive.google.com/thumbnail?id=${fileId}&sz=200&t=${Date.now()}`;
  }
  
  return imageData;
}

/**
 * Generate default avatar based on name/email
 */
function getDefaultAvatar(email, firstName) {
  // Use a simple avatar service or generate initials
  const initials = firstName ? firstName.charAt(0).toUpperCase() : email.charAt(0).toUpperCase();
  const colors = ['#1abc9c', '#3498db', '#9b59b6', '#e74c3c', '#f39c12'];
  const colorIndex = email.length % colors.length;
  
  return `data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><circle cx="50" cy="50" r="50" fill="${colors[colorIndex]}"/><text x="50" y="60" font-family="Arial" font-size="40" fill="white" text-anchor="middle" dominant-baseline="central">${initials}</text></svg>`;
}