// ===========================================
// PAYMENTS PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check payments permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanEdit')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkPaymentsPermission(token, action) {
    try {
        // Validate session
        const session = Authentication.validateSession(token);
        if (!session) {
            throw new Error("Invalid session. Please log in again.");
        }
        
        // Check if user has payments permissions
        const paymentsPerms = session.permissions && session.permissions["payments"];
        if (!paymentsPerms) {
            throw new Error("You don't have permission to access payment records.");
        }
        
        // Check specific action permission
        const hasPermission = paymentsPerms[action];
        
        if (!hasPermission) {
            const actionText = action.replace('Can', '').toLowerCase();
            throw new Error(`You don't have permission to ${actionText} payment records.`);
        }
        
        return session;
    } catch (error) {
        console.error('Payments permission check error:', error);
        throw error;
    }
}

/**
 * Helper: read a named range as objects. Formats "Date" as MM/dd/yyyy.
 */
function _ptGetRangeData(name) {
    const ss = SpreadsheetApp.getActive();
    const rg = ss.getRangeByName(name);
    if (!rg) throw new Error(`Named range "${name}" not found`);
    const vals = rg.getValues();
    if (vals.length < 2) return [];
    const hdr = vals[0];

    return vals.slice(1)
        .filter(r => r.some(c => c !== '' && c != null))
        .map(r => {
            const obj = {};
            hdr.forEach((h, i) => {
                let v = r[i];
                // If this header contains "Date", format it
                if (h.toLowerCase().includes('date') && v instanceof Date) {
                    v = Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
                }
                obj[h] = v;
            });
            return obj;
        });
}

/**
 * Get all payments data for the dashboard (requires CanView permission)
 */
function ptGetAllPayments(token) {
    try {
        // Check view permission
        checkPaymentsPermission(token, "CanView");
        
        return _ptGetRangeData('RANGEPAYMENTS');
    } catch (error) {
        console.error('Error in ptGetAllPayments:', error);
        throw error;
    }
}

/**
 * Get payment modes from dimensions (requires CanView permission)
 */
function ptGetDimensions(token) {
    try {
        // Check view permission
        checkPaymentsPermission(token, "CanView");
        
        const ss = SpreadsheetApp.getActive();
        const sheet = ss.getSheetByName("Dimensions");
        const range = ss.getRangeByName("RANGEDIMENSIONS");
        
        if (!range) {
            console.error("RANGEDIMENSIONS not found");
            return [];
        }
        
        const data = range.getValues();
        if (data.length < 2) {
            console.error("No data in RANGEDIMENSIONS");
            return [];
        }
        
        const headers = data[0];
        const result = [];
        
        // Log headers for debugging
        console.log("Dimensions headers:", headers);
        
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (row.every(cell => cell === '' || cell == null)) continue;
            
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            result.push(obj);
        }
        
        console.log(`Retrieved ${result.length} dimensions records`);
        return result;
        
    } catch (error) {
        console.error('Error in ptGetDimensions:', error);
        throw error;
    }
}

/**
 * Update payment record (requires CanEdit permission)
 */
function ptUpdatePayment(token, updatedPayment) {
    try {
        // Check edit permission
        checkPaymentsPermission(token, "CanEdit");
        
        const ss = SpreadsheetApp.getActive();
        const rg = ss.getRangeByName('RANGEPAYMENTS');
        const sh = rg.getSheet();
        const vs = rg.getValues();
        const hdr = vs[0];
        const data = vs.slice(1);
        
        // Find the column indexes
        const colPaymentID = hdr.indexOf('Payment ID');
        const colDate = hdr.indexOf('Date');
        const colTransactionType = hdr.indexOf('Transaction Type');
        const colReferenceID = hdr.indexOf('Reference ID');
        const colPartyID = hdr.indexOf('Party ID');
        const colPaymentMode = hdr.indexOf('Payment Mode');
        const colAmount = hdr.indexOf('Amount');
        const colStatus = hdr.indexOf('Status');
        
        // Find the row to update
        const rowIndex = data.findIndex(row => row[colPaymentID] === updatedPayment['Payment ID']);
        
        if (rowIndex === -1) {
            throw new Error('Payment record not found');
        }
        
        // Update the row (add 2 because header is row 1 and data starts at row 2)
        const rowNum = rg.getRow() + rowIndex + 1;
        
        // Convert date string to Date object
        const dateValue = new Date(updatedPayment['Date']);
        
        // Update the cells
        sh.getRange(rowNum, rg.getColumn() + colDate).setValue(dateValue);
        sh.getRange(rowNum, rg.getColumn() + colTransactionType).setValue(updatedPayment['Transaction Type']);
        sh.getRange(rowNum, rg.getColumn() + colReferenceID).setValue(updatedPayment['Reference ID']);
        sh.getRange(rowNum, rg.getColumn() + colPartyID).setValue(updatedPayment['Party ID']);
        sh.getRange(rowNum, rg.getColumn() + colPaymentMode).setValue(updatedPayment['Payment Mode']);
        sh.getRange(rowNum, rg.getColumn() + colAmount).setValue(updatedPayment['Amount']);
        sh.getRange(rowNum, rg.getColumn() + colStatus).setValue(updatedPayment['Status']);
        
        // Recalculate any dependent values if needed
        updatePaymentStatus();
        
        return 'success';
    } catch (error) {
        console.error('Error in ptUpdatePayment:', error);
        throw error;
    }
}

/**
 * Update payment status in related sheets if needed
 */
function updatePaymentStatus() {
    // This function can update related sheets like Purchase Orders or Sales Orders
    // based on the payment updates if needed
    // For now, we'll leave it as a placeholder
    Logger.log('Payment status updated - recalculating dependencies if any');
}

/**
 * Backward compatibility function for existing calls without token
 */
function ptGetAllPaymentsNoToken() {
    console.warn("DEPRECATED: ptGetAllPayments called without token. Using session from context.");
    
    try {
        // Try to get token from URL parameters or context
        // This is a fallback for existing code
        return ptGetAllPayments("deprecated_token_fallback");
    } catch (error) {
        console.error("Error in backward compatibility call:", error);
        return [];
    }
}

/**
 * Backward compatibility function for existing calls without token
 */
function ptGetDimensionsNoToken() {
    console.warn("DEPRECATED: ptGetDimensions called without token. Using session from context.");
    
    try {
        // Try to get token from URL parameters or context
        // This is a fallback for existing code
        return ptGetDimensions("deprecated_token_fallback");
    } catch (error) {
        console.error("Error in backward compatibility call:", error);
        return [];
    }
}

/**
 * Debug function to check permissions (remove in production)
 */
function debugPaymentsPermissions(token) {
    try {
        const session = Authentication.validateSession(token);
        if (!session) {
            return { error: "Invalid session" };
        }
        
        return {
            userId: session.userId,
            userRole: session.userRole,
            permissions: session.permissions ? session.permissions["payments"] : null,
            allPermissions: session.permissions
        };
    } catch (error) {
        return { error: error.toString() };
    }
}