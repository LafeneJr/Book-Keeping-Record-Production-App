// ===========================================
// PROFILE.GS - User Profile Management
// ===========================================

/**
 * Get current user's full profile data
 */
function getCurrentUserProfile(sessionToken) {
  try {
    const session = Authentication.validateSession(sessionToken);
    if (!session) {
      return { success: false, message: 'Invalid session' };
    }
    
    const db = new Database();
    const user = db.getUserById(session.userId);
    
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    // Handle profile image - convert file ID to display URL
    let profileImage = '';
    const imageData = user.ProfileImage || '';
    
    if (imageData) {
      if (imageData.includes('drive.google.com')) {
        // Old format: full URL - extract file ID
        const fileId = extractFileIdFromUrl(imageData);
        if (fileId) {
          profileImage = `https://drive.google.com/thumbnail?id=${fileId}&sz=s500`;
        } else {
          profileImage = imageData;
        }
      } else if (imageData.length > 20) {
        // Likely a file ID (not a URL)
        profileImage = `https://drive.google.com/thumbnail?id=${imageData}&sz=s500`;
      } else if (imageData.startsWith('data:image')) {
        // Base64 image
        profileImage = imageData;
      } else {
        // Unknown format
        profileImage = imageData;
      }
    } else {
      // Generate default avatar
      profileImage = getDefaultAvatar(user.Email, user.FirstName);
    }
    
    // Get user stats
    const stats = getUserStats(session.userId);
    
    return {
      success: true,
      profile: {
        id: user.UserID,
        email: user.Email,
        firstName: user.FirstName,
        lastName: user.LastName,
        fullName: `${user.FirstName} ${user.LastName}`,
        role: user.Role,
        department: user.Department || '',
        joinDate: user.JoinDate || '',
        status: user.Status || 'Active',
        phone: user.Phone || '',
        address: user.Address || '',
        profileImage: profileImage,
        lastLogin: user.LastLogin || 'Never',
        token: sessionToken
      },
      stats: stats
    };
  } catch (error) {
    console.error('Error getting user profile:', error);
    return { success: false, message: 'Error loading profile' };
  }
}

/**
 * Update user profile information
 */
function updateUserProfile(sessionToken, profileData) {
  try {
    const session = Authentication.validateSession(sessionToken);
    if (!session) {
      return { success: false, message: 'Invalid session' };
    }
    
    const db = new Database();
    const userId = session.userId;
    
    // Prepare updates
    const updates = {
      FirstName: profileData.firstName,
      LastName: profileData.lastName,
      Phone: profileData.phone || '',
      Address: profileData.address || ''
    };
    
    // Update in database
    const success = db.updateUser(userId, updates);
    
    if (success) {
      // Update session data
      const updatedUser = db.getUserById(userId);
      const cache = CacheService.getUserCache();
      const sessionData = {
        userId: updatedUser.UserID,
        userEmail: updatedUser.Email,
        userFirstName: updatedUser.FirstName,
        userLastName: updatedUser.LastName,
        userRole: updatedUser.Role,
        roleId: updatedUser.Role.toLowerCase().replace(/\s+/g, '_'),
        permissions: session.permissions,
        menus: session.menus,
        timestamp: new Date().getTime()
      };
      
      // Update cache
      cache.put(sessionToken, JSON.stringify(sessionData), 10800);
      
      return {
        success: true,
        message: 'Profile updated successfully',
        user: {
          firstName: updatedUser.FirstName,
          lastName: updatedUser.LastName,
          fullName: `${updatedUser.FirstName} ${updatedUser.LastName}`
        }
      };
    } else {
      return { success: false, message: 'Failed to update profile' };
    }
  } catch (error) {
    console.error('Error updating profile:', error);
    return { success: false, message: 'Error updating profile' };
  }
}

/**
 * Upload profile picture (base64 to Google Drive)
 */

function uploadProfilePicture(sessionToken, imageData) {
  try {
    const session = Authentication.validateSession(sessionToken);
    if (!session) {
      return { success: false, message: 'Invalid session' };
    }
    
    const db = new Database();
    const userId = session.userId;
    
    if (!imageData || !imageData.startsWith('data:image')) {
      return { success: false, message: 'Invalid image data' };
    }
    
    // Extract base64 data
    const base64Data = imageData.split(',')[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'image/png',
      `profile_${userId}_${Date.now()}.png`
    );
    
    // Save to Google Drive
    const folder = getOrCreateProfileImagesFolder();
    const file = folder.createFile(blob);
    
    // Set sharing permissions
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Get ONLY the file ID (not the full URL)
    const fileId = file.getId();
    console.log('File created with ID:', fileId);
    
    // Store ONLY the file ID in the database
    const success = db.updateUser(userId, { 
      ProfileImage: fileId, // Store just the ID, not the URL
      ProfileImageType: 'drive_id'
    });
    
    if (success) {
      // Return display URL to frontend
      const displayUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=s500`;
      
      return {
        success: true,
        message: 'Profile picture updated',
        fileId: fileId,
        imageUrl: displayUrl,
        timestamp: Date.now()
      };
    } else {
      return { success: false, message: 'Failed to update profile in database' };
    }
  } catch (error) {
    console.error('Error uploading profile picture:', error);
    return { success: false, message: 'Error uploading image: ' + error.message };
  }
}

/**
 * Change user password
 */
function changeUserPassword(sessionToken, currentPassword, newPassword) {
  return Authentication.changePassword(
    Authentication.validateSession(sessionToken).userId,
    currentPassword,
    newPassword
  );
}

/**
 * Get user activity stats
 */
function getUserStats(userId) {
  try {
    const db = new Database();
    
    // Get user's recent activity (you can expand this)
    const today = new Date();
    const thirtyDaysAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
    
    // Example stats - customize based on your actual activity tracking
    return {
      lastLogin: new Date().toLocaleDateString(),
      daysActive: Math.floor(Math.random() * 30) + 1, // Placeholder
      tasksCompleted: Math.floor(Math.random() * 50) + 10, // Placeholder
      upcomingTasks: Math.floor(Math.random() * 5) + 1 // Placeholder
    };
  } catch (error) {
    console.error('Error getting user stats:', error);
    return {};
  }
}

/**
 * Get or create profile images folder in Drive
 */
function getOrCreateProfileImagesFolder() {
  const folderName = 'BookKeepingApp_ProfileImages';
  
  try {
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
      // Set folder sharing permissions
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    
    return folder;
  } catch (error) {
    console.error('Error with profile images folder:', error);
    throw error;
  }
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

// Make functions available to frontend
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getCurrentUserProfile,
    updateUserProfile,
    uploadProfilePicture,
    changeUserPassword,
    getUserStats
  };
}

/**
 * Extract file ID from Drive URL
 */
function extractFileIdFromUrl(url) {
  if (!url || typeof url !== 'string') return '';
  
  // If it's already just an ID (not a URL)
  if (!url.includes('drive.google.com') && !url.includes('http')) {
    return url;
  }
  
  // Try different patterns
  const patterns = [
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  // If no pattern matched, try to extract from common Drive URL formats
  if (url.includes('drive.google.com')) {
    const parts = url.split('/');
    for (const part of parts) {
      if (part.length > 20 && part.length < 50 && !part.includes('=') && !part.includes('?')) {
        return part;
      }
    }
  }
  
  return url;
}