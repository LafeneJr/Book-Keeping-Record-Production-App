// ===========================================
// SUPPLIERS FUNCTIONS WITH PERMISSION CHECKS
// ===========================================

// Helper function to check suppliers permissions
function checkSuppliersPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has suppliers permissions
    const suppliersPerms = session.permissions && session.permissions["suppliers"];
    if (!suppliersPerms) {
      throw new Error("You don't have permission to access suppliers.");
    }
    
    // Check specific action permission
    // action: 'CanView', 'CanAdd', 'CanEdit', or 'CanDelete'
    const hasPermission = suppliersPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} suppliers.`);
    }
    
    return session;
  } catch (error) {
    console.error('Permission check error:', error);
    throw error;
  }
}

// Get all suppliers (requires CanView permission)
function supGetSuppliers(token) {
  try {
    // Check view permission
    checkSuppliersPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const suppliers = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '') continue; // Skip empty rows
      
      suppliers.push({
        id: row[headers.indexOf("Supplier ID")],
        name: row[headers.indexOf("Supplier Name")],
        contact: row[headers.indexOf("Supplier Contact")],
        email: row[headers.indexOf("Supplier Email")],
        state: row[headers.indexOf("State")],
        city: row[headers.indexOf("City")],
        address: row[headers.indexOf("Supplier Address")],
        purchases: row[headers.indexOf("Total Purchases")],
        payments: row[headers.indexOf("Total Payments")],
        balance: row[headers.indexOf("Balance Payable")]
      });
    }
    
    return suppliers;
  } catch (error) {
    console.error('Error in supGetSuppliers:', error);
    throw error;
  }
}

// Get states from Dimensions (requires CanView permission)
function supGetStates(token) {
  try {
    // Check view permission
    checkSuppliersPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const stateIndex = headers.indexOf("State");
    
    if (stateIndex === -1) {
      return [];
    }
    
    const states = [];
    const stateSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const state = data[i][stateIndex];
      if (state && state.trim() !== '' && !stateSet.has(state)) {
        states.push(state);
        stateSet.add(state);
      }
    }
    
    return states;
  } catch (error) {
    console.error('Error in supGetStates:', error);
    throw error;
  }
}

// Get cities from Dimensions (requires CanView permission)
function supGetCities(token) {
  try {
    // Check view permission
    checkSuppliersPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const cityIndex = headers.indexOf("City");
    
    if (cityIndex === -1) {
      return [];
    }
    
    const cities = [];
    const citySet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const city = data[i][cityIndex];
      if (city && city.trim() !== '' && !citySet.has(city)) {
        cities.push(city);
        citySet.add(city);
      }
    }
    
    return cities;
  } catch (error) {
    console.error('Error in supGetCities:', error);
    throw error;
  }
}

// Add new state to Dimensions (requires CanAdd permission)
function supAddNewState(token, stateName) {
  try {
    // Check add permission
    checkSuppliersPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return;
    }
    
    const data = range.getValues();
    const headers = data[0];
    const stateIndex = headers.indexOf("State");
    
    if (stateIndex === -1) {
      return;
    }
    
    // Find first empty row
    let lastRow = sheet.getLastRow() + 1;
    
    // Add to State column
    sheet.getRange(lastRow, stateIndex + 1).setValue(stateName);
  } catch (error) {
    console.error('Error in supAddNewState:', error);
    throw error;
  }
}

// Add new city to Dimensions (requires CanAdd permission)
function supAddNewCity(token, cityName) {
  try {
    // Check add permission
    checkSuppliersPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return;
    }
    
    const data = range.getValues();
    const headers = data[0];
    const cityIndex = headers.indexOf("City");
    
    if (cityIndex === -1) {
      return;
    }
    
    // Find first empty row
    let lastRow = sheet.getLastRow() + 1;
    
    // Add to City column
    sheet.getRange(lastRow, cityIndex + 1).setValue(cityName);
  } catch (error) {
    console.error('Error in supAddNewCity:', error);
    throw error;
  }
}

// Generate unique supplier ID (requires CanAdd permission)
function supGenerateSupplierId(token) {
  try {
    // Check add permission
    checkSuppliersPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");
    
    if (!range) {
      return "P" + Math.floor(10000 + Math.random() * 90000);
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Supplier ID");
    const existingIds = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIndex];
      if (id && id !== '') {
        existingIds.add(id);
      }
    }
    
    let newId;
    do {
      newId = "P" + Math.floor(10000 + Math.random() * 90000);
    } while (existingIds.has(newId));
    
    return newId;
  } catch (error) {
    console.error('Error in supGenerateSupplierId:', error);
    throw error;
  }
}

// Add new supplier (requires CanAdd permission)
function supAddNewSupplier(token, supplier) {
  try {
    // Check add permission
    checkSuppliersPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");
    
    if (!range) {
      return;
    }
    
    const data = range.getValues();
    const headers = data[0];
    
    const newRow = [];
    headers.forEach(header => {
      switch(header) {
        case "Supplier ID": newRow.push(supplier.id); break;
        case "Supplier Name": newRow.push(supplier.name); break;
        case "Supplier Contact": newRow.push(supplier.contact); break;
        case "Supplier Email": newRow.push(supplier.email); break;
        case "State": newRow.push(supplier.state); break;
        case "City": newRow.push(supplier.city); break;
        case "Supplier Address": newRow.push(supplier.address); break;
        case "Total Purchases": newRow.push(0); break;
        case "Total Payments": newRow.push(0); break;
        case "Balance Payable": newRow.push(0); break;
        default: newRow.push(''); break;
      }
    });
    
    // Append to sheet
    sheet.appendRow(newRow);
  } catch (error) {
    console.error('Error in supAddNewSupplier:', error);
    throw error;
  }
}

// Update supplier (requires CanEdit permission)
function supUpdateSupplier(token, supplier) {
  try {
    // Check edit permission
    checkSuppliersPermission(token, "CanEdit");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");
    
    if (!range) {
      return;
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Supplier ID");
    
    // Find row index
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === supplier.id) {
        const rowNum = range.getRow() + i;
        
        // Update editable fields
        sheet.getRange(rowNum, headers.indexOf("Supplier Name") + 1).setValue(supplier.name);
        sheet.getRange(rowNum, headers.indexOf("Supplier Contact") + 1).setValue(supplier.contact);
        sheet.getRange(rowNum, headers.indexOf("Supplier Email") + 1).setValue(supplier.email);
        sheet.getRange(rowNum, headers.indexOf("State") + 1).setValue(supplier.state);
        sheet.getRange(rowNum, headers.indexOf("City") + 1).setValue(supplier.city);
        sheet.getRange(rowNum, headers.indexOf("Supplier Address") + 1).setValue(supplier.address);
        
        break;
      }
    }
  } catch (error) {
    console.error('Error in supUpdateSupplier:', error);
    throw error;
  }
}

// Delete supplier (requires CanDelete permission)
function supDeleteSupplier(token, supplierId) {
  try {
    // Check delete permission
    checkSuppliersPermission(token, "CanDelete");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");
    
    if (!range) {
      return "error";
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Supplier ID");
    const balanceIndex = headers.indexOf("Balance Payable");
    
    // Find row index
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === supplierId) {
        const balance = data[i][balanceIndex] || 0;
        
        // Check balance
        if (balance > 0) {
          return "balance_error";
        }
        
        // Delete row
        const rowNum = range.getRow() + i;
        sheet.deleteRow(rowNum);
        return "success";
      }
    }
    
    return "not_found";
  } catch (error) {
    console.error('Error in supDeleteSupplier:', error);
    throw error;
  }
}

// ===========================================
// BACKWARD COMPATIBILITY FUNCTIONS
// ===========================================

// For backward compatibility with existing calls
function supGetSuppliersNoToken() {
  console.warn("DEPRECATED: supGetSuppliers called without token.");
  throw new Error("This function requires authentication. Please use the token-based version.");
}

// ===========================================
// PERMISSION DEBUG FUNCTION
// ===========================================

function debugSuppliersPermissions(token) {
  try {
    const session = Authentication.validateSession(token);
    if (!session) {
      return { error: "Invalid session" };
    }
    
    return {
      userId: session.userId,
      userRole: session.userRole,
      permissions: session.permissions ? session.permissions["suppliers"] : null,
      allPermissions: session.permissions
    };
  } catch (error) {
    return { error: error.toString() };
  }
}