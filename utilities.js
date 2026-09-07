// ===========================================
// UTILITIES.GS - Helper Functions
// ===========================================

/**
 * Hash password using SHA-256
 */
function hashPassword(password) {
  if (!password) return '';
  
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password
  );
  
  const hash = digest.map(function(byte) {
    const v = (byte < 0) ? 256 + byte : byte;
    return ("0" + v.toString(16)).slice(-2);
  }).join('');
  
  return hash;
}

/**
 * Verify password against hash
 */
function verifyPassword(password, hash) {
  if (!password || !hash) return false;
  return hashPassword(password) === hash;
}

/**
 * Generate random token
 */
function generateRandomToken(length = 32) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let token = "";
  
  for (let i = 0; i < length; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  
  return token;
}

/**
 * Generate UUID
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Format date to readable string
 */
function formatDate(dateString) {
  if (!dateString) return '';
  
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy');
  } catch (error) {
    return dateString;
  }
}

/**
 * Format date-time to readable string
 */
function formatDateTime(dateString) {
  if (!dateString) return '';
  
  try {
    const date = new Date(dateString);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
  } catch (error) {
    return dateString;
  }
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email);
}

/**
 * Validate password strength
 */
function isValidPassword(password) {
  if (!password || password.length < 8) return false;
  
  // Check for at least one uppercase letter
  if (!/[A-Z]/.test(password)) return false;
  
  // Check for at least one lowercase letter
  if (!/[a-z]/.test(password)) return false;
  
  // Check for at least one number
  if (!/[0-9]/.test(password)) return false;
  
  return true;
}

/**
 * Sanitize string for HTML output
 */
function sanitizeHtml(str) {
  if (!str) return '';
  
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Handle API errors
 */
function handleApiError(error) {
  console.error('API Error:', error);
  
  return {
    success: false,
    message: 'An unexpected error occurred',
    error: error.toString()
  };
}

/**
 * Get script URL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Generate ID for any table
 */
function generateID(tableName, idColumn, prefix = '') {
  const db = new Database();
  const records = db.getAll(tableName);
  
  let maxId = 0;
  records.forEach(record => {
    const id = record[idColumn];
    if (id && id.startsWith(prefix)) {
      const num = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(num) && num > maxId) {
        maxId = num;
      }
    }
  });
  
  return prefix + (maxId + 1).toString().padStart(4, '0');
}

// Add to utilities.gs
function handleProfileImage(imageData) {
  if (!imageData) return '';
  
  // If it's a Drive file ID
  if (imageData.length > 20 && imageData.length < 50 && !imageData.includes('/')) {
    return `https://drive.google.com/thumbnail?id=${imageData}&sz=s200`;
  }
  
  // If it's a full Drive URL
  if (imageData.includes('drive.google.com')) {
    const fileId = extractFileIdFromUrl(imageData);
    if (fileId) {
      return `https://drive.google.com/thumbnail?id=${fileId}&sz=s200`;
    }
  }
  
  // If it's base64
  if (imageData.startsWith('data:image')) {
    return imageData;
  }
  
  return imageData; // Return as-is
}

function extractFileIdFromUrl(url) {
  if (!url) return '';
  
  // If it's already a file ID (not a URL)
  if (!url.includes('http') && url.length > 20) {
    return url;
  }
  
  // Extract from various Google Drive URL formats
  const patterns = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /\/folders\/([a-zA-Z0-9_-]+)/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return url; // Return original if no match
}