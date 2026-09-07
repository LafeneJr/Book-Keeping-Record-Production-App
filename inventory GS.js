// ===========================================
// INVENTORY FUNCTIONS WITH PERMISSION CHECKS
// ===========================================

// Helper function to check inventory permissions
function checkInventoryPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has inventory permissions
    const inventoryPerms = session.permissions && session.permissions["inventory"];
    if (!inventoryPerms) {
      throw new Error("You don't have permission to access inventory.");
    }
    
    // Check specific action permission
    // action: 'CanView', 'CanAdd', 'CanEdit', or 'CanDelete'
    const hasPermission = inventoryPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} inventory items.`);
    }
    
    return session;
  } catch (error) {
    console.error('Permission check error:', error);
    throw error;
  }
}

// Get all inventory items (requires CanView permission)
function itemGetInventoryItems(token) {
  try {
    // Check view permission
    checkInventoryPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const items = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '') continue; // Skip empty rows
      
      items.push({
        id: row[headers.indexOf("Item ID")],
        type: row[headers.indexOf("Item Type")],
        category: row[headers.indexOf("Item Category")],
        name: row[headers.indexOf("Item Name")],
        purchasedQty: row[headers.indexOf("Quantity Purchased")] || 0,
        usedQty: row[headers.indexOf("Quantity Used")] || 0,
        productQty: row[headers.indexOf("Product Quantity")] || 0,
        soldQty: row[headers.indexOf("Quantity Sold")] || 0,
        remainingQty: row[headers.indexOf("Remaining Quantity")] || 0,
        reorderLevel: row[headers.indexOf("Reorder Level")] || 0,
        reorderRequired: row[headers.indexOf("Reorder Required")] === "Yes"
      });
    }
    
    return items;
  } catch (error) {
    console.error('Error in itemGetInventoryItems:', error);
    throw error;
  }
}

// Get item types from Dimensions (requires CanView permission)
function itemGetTypes(token) {
  try {
    // Check view permission
    checkInventoryPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const typeIndex = headers.indexOf("Item Type");
    
    if (typeIndex === -1) {
      return [];
    }
    
    const types = [];
    const typeSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const type = data[i][typeIndex];
      if (type && type.trim() !== '' && !typeSet.has(type)) {
        types.push(type);
        typeSet.add(type);
      }
    }
    
    return types;
  } catch (error) {
    console.error('Error in itemGetTypes:', error);
    throw error;
  }
}

// Get item categories from Dimensions (requires CanView permission)
function itemGetCategories(token) {
  try {
    // Check view permission
    checkInventoryPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const categoryIndex = headers.indexOf("Item Category");
    
    if (categoryIndex === -1) {
      return [];
    }
    
    const categories = [];
    const categorySet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const category = data[i][categoryIndex];
      if (category && category.trim() !== '' && !categorySet.has(category)) {
        categories.push(category);
        categorySet.add(category);
      }
    }
    
    return categories;
  } catch (error) {
    console.error('Error in itemGetCategories:', error);
    throw error;
  }
}

// Get item names from Dimensions (requires CanView permission)
function itemGetNames(token) {
  try {
    // Check view permission
    checkInventoryPermission(token, "CanView");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const nameIndex = headers.indexOf("Item Name");
    
    if (nameIndex === -1) {
      return [];
    }
    
    const names = [];
    const nameSet = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const name = data[i][nameIndex];
      if (name && name.trim() !== '' && !nameSet.has(name)) {
        names.push(name);
        nameSet.add(name);
      }
    }
    
    return names;
  } catch (error) {
    console.error('Error in itemGetNames:', error);
    throw error;
  }
}

// Add new item type to Dimensions (requires CanAdd permission)
function itemAddNewType(token, typeName) {
  try {
    // Check add permission
    checkInventoryPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return { success: false, message: "Dimensions range not found" };
    }
    
    const data = range.getValues();
    const headers = data[0];
    const typeIndex = headers.indexOf("Item Type");
    
    if (typeIndex === -1) {
      return { success: false, message: "Item Type column not found" };
    }
    
    // Find first empty row
    let lastRow = sheet.getLastRow() + 1;
    
    // Add to Item Type column
    sheet.getRange(lastRow, typeIndex + 1).setValue(typeName);
    
    return { success: true, message: "Item type added successfully" };
  } catch (error) {
    console.error('Error in itemAddNewType:', error);
    throw error;
  }
}

// Add new item category to Dimensions (requires CanAdd permission)
function itemAddNewCategory(token, categoryName) {
  try {
    // Check add permission
    checkInventoryPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return { success: false, message: "Dimensions range not found" };
    }
    
    const data = range.getValues();
    const headers = data[0];
    const categoryIndex = headers.indexOf("Item Category");
    
    if (categoryIndex === -1) {
      return { success: false, message: "Item Category column not found" };
    }
    
    // Find first empty row
    let lastRow = sheet.getLastRow() + 1;
    
    // Add to Item Category column
    sheet.getRange(lastRow, categoryIndex + 1).setValue(categoryName);
    
    return { success: true, message: "Item category added successfully" };
  } catch (error) {
    console.error('Error in itemAddNewCategory:', error);
    throw error;
  }
}

// Generate unique item ID (requires CanAdd permission)
function itemGenerateInventoryId(token) {
  try {
    // Check add permission
    checkInventoryPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return "P" + Math.floor(10000 + Math.random() * 90000);
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Item ID");
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
    console.error('Error in itemGenerateInventoryId:', error);
    throw error;
  }
}

// Add new inventory item (requires CanAdd permission)
function itemAddNewInventoryItem(token, item) {
  try {
    // Check add permission
    checkInventoryPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return { success: false, message: "Inventory range not found" };
    }
    
    const data = range.getValues();
    const headers = data[0];
    
    const newRow = [];
    headers.forEach(header => {
      switch(header) {
        case "Item ID": newRow.push(item.id); break;
        case "Item Type": newRow.push(item.type); break;
        case "Item Category": newRow.push(item.category); break;
        case "Item Name": newRow.push(item.name); break;
        case "Quantity Purchased": newRow.push(0); break;
        case "Quantity Used": newRow.push(0); break;
        case "Product Quantity": newRow.push(0); break;
        case "Quantity Sold": newRow.push(0); break;
        case "Remaining Quantity": newRow.push(0); break;
        case "Reorder Level": newRow.push(item.reorderLevel); break;
        case "Reorder Required": newRow.push("No"); break;
        default: newRow.push(''); break;
      }
    });
    
    // Append to sheet
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Inventory item added successfully' };
  } catch (error) {
    console.error('Error in itemAddNewInventoryItem:', error);
    throw error;
  }
}

// Update inventory item (requires CanEdit permission)
function itemUpdateInventoryItem(token, item) {
  try {
    // Check edit permission
    checkInventoryPermission(token, "CanEdit");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return { success: false, message: "Inventory range not found" };
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Item ID");
    
    // Find row index
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === item.id) {
        const rowNum = range.getRow() + i;
        
        // Update editable fields
        sheet.getRange(rowNum, headers.indexOf("Item Type") + 1).setValue(item.type);
        sheet.getRange(rowNum, headers.indexOf("Item Category") + 1).setValue(item.category);
        sheet.getRange(rowNum, headers.indexOf("Item Name") + 1).setValue(item.name);
        sheet.getRange(rowNum, headers.indexOf("Reorder Level") + 1).setValue(item.reorderLevel);
        
        // Update reorder required based on remaining quantity
        const remainingQty = sheet.getRange(rowNum, headers.indexOf("Remaining Quantity") + 1).getValue();
        const reorderRequired = remainingQty < item.reorderLevel ? "Yes" : "No";
        sheet.getRange(rowNum, headers.indexOf("Reorder Required") + 1).setValue(reorderRequired);
        
        return { success: true, message: 'Inventory item updated successfully' };
      }
    }
    
    return { success: false, message: "Item not found" };
  } catch (error) {
    console.error('Error in itemUpdateInventoryItem:', error);
    throw error;
  }
}

// Delete inventory item (requires CanDelete permission)
function itemDeleteInventoryItem(token, itemId) {
  try {
    // Check delete permission
    checkInventoryPermission(token, "CanDelete");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return "error";
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Item ID");
    const qtyIndex = headers.indexOf("Remaining Quantity");
    
    // Find row index
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === itemId) {
        const qty = data[i][qtyIndex] || 0;
        
        // Check remaining quantity
        if (qty > 0) {
          return "quantity_error";
        }
        
        // Delete row
        const rowNum = range.getRow() + i;
        sheet.deleteRow(rowNum);
        return "success";
      }
    }
    
    return "not_found";
  } catch (error) {
    console.error('Error in itemDeleteInventoryItem:', error);
    throw error;
  }
}

// Add new item name to Dimensions (requires CanAdd permission)
function itemAddNewName(token, itemName) {
  try {
    // Check add permission
    checkInventoryPermission(token, "CanAdd");
    
    // Original code below - NO CHANGES to functionality
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!range) {
      return { success: false, message: "Dimensions range not found" };
    }
    
    const data = range.getValues();
    const headers = data[0];
    const nameIndex = headers.indexOf("Item Name");
    
    if (nameIndex === -1) {
      return { success: false, message: "Item Name column not found" };
    }
    
    // Find first empty row
    let lastRow = sheet.getLastRow() + 1;
    
    // Add to Item Name column
    sheet.getRange(lastRow, nameIndex + 1).setValue(itemName);
    
    return { success: true, message: 'Item name added successfully' };
  } catch (error) {
    console.error('Error in itemAddNewName:', error);
    throw error;
  }
}

// ===========================================
// BACKWARD COMPATIBILITY FUNCTIONS
// ===========================================

// These functions maintain backward compatibility with existing calls
// They wrap the new token-based functions for existing code

function itemGetInventoryItemsNoToken() {
  console.warn("DEPRECATED: itemGetInventoryItems called without token. Using session from context.");
  
  try {
    // Try to get token from URL parameters or context
    // This is a fallback for existing code
    return itemGetInventoryItems("deprecated_token_fallback");
  } catch (error) {
    console.error("Error in backward compatibility call:", error);
    return [];
  }
}

function itemAddNewInventoryItemNoToken(item) {
  console.warn("DEPRECATED: itemAddNewInventoryItem called without token.");
  // This function can't work without permissions check
  throw new Error("This function requires authentication. Please use the token-based version.");
}

// ===========================================
// PERMISSION DEBUG FUNCTION (REMOVE IN PRODUCTION)
// ===========================================

function debugInventoryPermissions(token) {
  try {
    const session = Authentication.validateSession(token);
    if (!session) {
      return { error: "Invalid session" };
    }
    
    return {
      userId: session.userId,
      userRole: session.userRole,
      permissions: session.permissions ? session.permissions["inventory"] : null,
      allPermissions: session.permissions
    };
  } catch (error) {
    return { error: error.toString() };
  }
}