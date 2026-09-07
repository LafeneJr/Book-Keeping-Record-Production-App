// ===========================================
// PRODUCTION PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check production permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanAdd', 'CanEdit', 'CanDelete')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkProductionPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has production permissions
    const productionPerms = session.permissions && session.permissions["productions"];
    if (!productionPerms) {
      throw new Error("You don't have permission to access production records.");
    }
    
    // Check specific action permission
    const hasPermission = productionPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} production records.`);
    }
    
    return session;
  } catch (error) {
    console.error('Production permission check error:', error);
    throw error;
  }
}

// Get all production records
function productionGetRecords(token) {
  try {
    // Check view permission
    checkProductionPermission(token, "CanView");
        
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Production Records");
    const range = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const records = [];
    
    // Read data from bottom to top (newest records are at the bottom of the sheet)
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      if (row[0] === '') continue; // Skip empty rows
      
      let materialUsedStr = row[headers.indexOf("Material Used")] || '';
      let quantityUsedStr = row[headers.indexOf("Quantity Used")] || '';
      
      // Convert to strings to avoid split errors
      materialUsedStr = String(materialUsedStr);
      quantityUsedStr = String(quantityUsedStr);
      
      // Parse materials and quantities from comma-separated strings
      const materialNames = materialUsedStr.split(',').map(m => m.trim()).filter(m => m !== '');
      const quantities = quantityUsedStr.split(',').map(q => {
        const num = parseFloat(q.trim());
        return isNaN(num) ? 0 : num;
      });
      
      // Create materials array
      const materialsUsed = [];
      let totalQuantityUsed = 0;
      
      for (let j = 0; j < materialNames.length; j++) {
        if (materialNames[j]) {
          const quantity = quantities[j] || 0;
          materialsUsed.push({
            name: materialNames[j],
            quantityUsed: quantity
          });
          totalQuantityUsed += quantity;
        }
      }
      
      records.push({
        date: formatDate(row[headers.indexOf("Date")]),
        id: row[headers.indexOf("Production ID")],
        productName: row[headers.indexOf("Product Name")],
        materialsUsed: materialsUsed,
        totalQuantityUsed: totalQuantityUsed,
        quantityProduced: row[headers.indexOf("Quantity Produced")] || 0
      });
    }
    
    return records;
  } catch (error) {
    console.error('Error in productionGetRecords:', error);
    throw error;
  }
}


// Generate unique production ID
function productionGenerateId(token) {
  try {
    // Check add permission
    checkProductionPermission(token, "CanAdd");
        
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Production Records");
    const range = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    
    if (!range) {
      return "PR" + Math.floor(10000 + Math.random() * 90000);
    }
    
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Production ID");
    const existingIds = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIndex];
      if (id && id !== '') {
        existingIds.add(id);
      }
    }
    
    let newId;
    do {
      newId = "PR" + Math.floor(10000 + Math.random() * 90000);
    } while (existingIds.has(newId));
    
    return newId;
  } catch (error) {
    console.error('Error in productionGenerateId:', error);
    throw error;
  }
}

// Add new production record - SINGLE ROW VERSION
function productionAddRecord(token, record) {
  try {
    // Check add permission
    checkProductionPermission(token, "CanAdd");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const productionSheet = ss.getSheetByName("Production Records");
    const inventorySheet = ss.getSheetByName("Inventory Items");
    const dimensionsSheet = ss.getSheetByName("Dimensions");
    
    const productionRange = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
    const dimensionsRange = ss.getRangeByName("RANGEDIMENSIONS");
    
    if (!productionRange) {
      return;
    }
    
    // Get headers from production records
    const productionData = productionRange.getValues();
    const productionHeaders = productionData[0];
    
    // Create ONE production record with all materials in the same row
    const newProductionRow = [];
    productionHeaders.forEach(header => {
      switch(header) {
        case "Date": 
          // Convert MM/DD/YYYY to Date object
          const dateParts = record.date.split('/');
          const dateObj = new Date(dateParts[2], dateParts[0] - 1, dateParts[1]);
          newProductionRow.push(dateObj); 
          break;
        case "Production ID": newProductionRow.push(record.id); break;
        case "Product Name": newProductionRow.push(record.productName); break;
        case "Material Used": 
          // Store all materials as a comma-separated string
          const materialNames = record.materialsUsed.map(m => m.name).join(', ');
          newProductionRow.push(materialNames); 
          break;
        case "Quantity Used": 
          // Store all quantities as a comma-separated string
          const quantities = record.materialsUsed.map(m => m.quantityUsed).join(',');
          newProductionRow.push(quantities); 
          break;
        case "Quantity Produced": newProductionRow.push(record.quantityProduced); break;
        default: newProductionRow.push(''); break;
      }
    });
    
    // Append ONE row to production records
    productionSheet.appendRow(newProductionRow);
    
    // Update Inventory Items - Reduce raw material stock for each material
    if (inventoryRange) {
      const inventoryData = inventoryRange.getValues();
      const inventoryHeaders = inventoryData[0];
      
      const itemNameIndex = inventoryHeaders.indexOf("Item Name");
      const purchasedQtyIndex = inventoryHeaders.indexOf("Quantity Purchased");
      const usedQtyIndex = inventoryHeaders.indexOf("Quantity Used");
      const productQtyIndex = inventoryHeaders.indexOf("Product Quantity");
      const remainingQtyIndex = inventoryHeaders.indexOf("Remaining Quantity");
      const reorderRequiredIndex = inventoryHeaders.indexOf("Reorder Required");
      const reorderLevelIndex = inventoryHeaders.indexOf("Reorder Level");

      const soldQtyIndex = inventoryHeaders.indexOf("Quantity Sold");
      
      // Create a map to track which materials we've updated to handle duplicates
      const updatedMaterials = new Map();
      
      // 1. Update each raw material (reduce stock)
      record.materialsUsed.forEach(material => {
        // Skip if we've already updated this material (duplicate in same record)
        if (updatedMaterials.has(material.name)) {
          Logger.log(`Skipping duplicate material: ${material.name}`);
          return;
        }
        
        let materialUpdated = false;
        
        // Search through inventory to find the exact row for this material
        for (let i = 1; i < inventoryData.length; i++) {
          const row = inventoryData[i];
          if (row[itemNameIndex] === material.name) {
            const rowNum = inventoryRange.getRow() + i;
            
            // Get CURRENT values directly from the sheet to ensure accuracy
            const currentPurchasedQty = inventorySheet.getRange(rowNum, purchasedQtyIndex + 1).getValue() || 0;
            const currentUsedQty = inventorySheet.getRange(rowNum, usedQtyIndex + 1).getValue() || 0;
            const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
            const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
            const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
            
            // FIXED: Calculate new values for raw material using correct formula
            const newUsedQty = currentUsedQty + material.quantityUsed;
            const newRemainingQty = currentPurchasedQty - newUsedQty - currentSoldQty;
            const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
            
            // Update raw material inventory
            inventorySheet.getRange(rowNum, usedQtyIndex + 1).setValue(newUsedQty);
            inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
            inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
            
            materialUpdated = true;
            updatedMaterials.set(material.name, true);
            Logger.log(`Updated material: ${material.name}, Used: ${material.quantityUsed}, New Remaining: ${newRemainingQty} (Formula: ${currentPurchasedQty} - ${newUsedQty} - ${currentSoldQty})`);
            break;
          }
        }
        
        if (!materialUpdated) {
          Logger.log(`Raw material not found in inventory: ${material.name}`);
        }
      });
      
      // 2. Update or create the finished product
      let productFound = false;

      // Get fresh inventory data again after raw material updates
      const updatedInventoryData = inventoryRange.getValues();

      for (let i = 1; i < updatedInventoryData.length; i++) {
        const row = updatedInventoryData[i];
        if (row[itemNameIndex] === record.productName) {
          const rowNum = inventoryRange.getRow() + i;
          
          // Get CURRENT values directly from the sheet
          const currentProductQty = inventorySheet.getRange(rowNum, productQtyIndex + 1).getValue() || 0;
          const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
          const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
          const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
          
          // FIXED: Calculate new values for finished product using correct formula
          const newProductQty = currentProductQty + record.quantityProduced;
          const newRemainingQty = newProductQty - currentSoldQty; // CORRECT FORMULA
          const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
          
          // Update finished product inventory
          inventorySheet.getRange(rowNum, productQtyIndex + 1).setValue(newProductQty);
          inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
          inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
          
          productFound = true;
          Logger.log(`Updated finished product: ${record.productName}, Produced: ${record.quantityProduced}, New Remaining: ${newRemainingQty} (Formula: ${newProductQty} - ${currentSoldQty})`);
          break;
        }
      }
      
      // If finished product not found in inventory, create new inventory item
      if (!productFound) {
        const newInventoryRow = [];
        
        inventoryHeaders.forEach(header => {
          switch(header) {
            case "Item ID": 
              newInventoryRow.push("P" + Math.floor(10000 + Math.random() * 90000)); 
              break;
            case "Item Type": newInventoryRow.push("Finished Goods"); break;
            case "Item Category": newInventoryRow.push("Production"); break;
            case "Item Name": newInventoryRow.push(record.productName); break;
            case "Quantity Purchased": newInventoryRow.push(0); break;
            case "Quantity Used": newInventoryRow.push(0); break;
            case "Product Quantity": newInventoryRow.push(record.quantityProduced); break;
            case "Quantity Sold": newInventoryRow.push(0); break;
            case "Remaining Quantity": newInventoryRow.push(record.quantityProduced); break;
            case "Reorder Level": newInventoryRow.push(0); break;
            case "Reorder Required": newInventoryRow.push("No"); break;
            default: newInventoryRow.push(''); break;
          }
        });
        
        inventorySheet.appendRow(newInventoryRow);
        Logger.log(`Created new finished product: ${record.productName}`);
      }
    }
    
    // Add product name to Dimensions if it doesn't exist
    if (dimensionsRange) {
      const dimensionsData = dimensionsRange.getValues();
      const dimensionsHeaders = dimensionsData[0];
      const nameIndex = dimensionsHeaders.indexOf("Item Name");
      
      let nameExists = false;
      for (let i = 1; i < dimensionsData.length; i++) {
        if (dimensionsData[i][nameIndex] === record.productName) {
          nameExists = true;
          break;
        }
      }
      
      if (!nameExists) {
        const lastRow = dimensionsSheet.getLastRow() + 1;
        dimensionsSheet.getRange(lastRow, nameIndex + 1).setValue(record.productName);
        // Also set Item Type to "Finished Goods" and Category to "Production"
        const typeIndex = dimensionsHeaders.indexOf("Item Type");
        const categoryIndex = dimensionsHeaders.indexOf("Item Category");
        if (typeIndex !== -1) dimensionsSheet.getRange(lastRow, typeIndex + 1).setValue("Finished Goods");
        if (categoryIndex !== -1) dimensionsSheet.getRange(lastRow, categoryIndex + 1).setValue("Production");
      }
    }
    
    Logger.log("Production record saved successfully");
  
  } catch (error) {
    console.error('Error in productionAddRecord:', error);
    throw error;
  }
}

// Update production record - SINGLE ROW VERSION
function productionUpdateRecord(token, record) {
  try {
    // Check edit permission
    checkProductionPermission(token, "CanEdit");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const productionSheet = ss.getSheetByName("Production Records");
    const inventorySheet = ss.getSheetByName("Inventory Items");
    
    const productionRange = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!productionRange) {
      return "error: production range not found";
    }
    
    try {
      // Get the old record first to revert inventory changes
      const oldRecord = productionGetRecordById(token, record.id);
      if (!oldRecord) {
        return "error: old record not found";
      }
      
      // Step 1: Revert old inventory changes
      revertInventoryChanges(oldRecord);
      
      // Step 2: Update the production record in a single row
      const productionData = productionRange.getValues();
      const productionHeaders = productionData[0];
      
      // Find the row to update
      let rowToUpdate = -1;
      for (let i = 1; i < productionData.length; i++) {
        if (productionData[i][productionHeaders.indexOf("Production ID")] === record.id) {
          rowToUpdate = productionRange.getRow() + i;
          break;
        }
      }
      
      if (rowToUpdate === -1) {
        return "error: record not found";
      }
      
      // Prepare materials and quantities as comma-separated strings
      const materialNames = record.materialsUsed.map(m => m.name).join(', ');
      const quantities = record.materialsUsed.map(m => m.quantityUsed).join(',');
      
      // Update the row
      productionSheet.getRange(rowToUpdate, productionHeaders.indexOf("Date") + 1).setValue(new Date(record.date));
      productionSheet.getRange(rowToUpdate, productionHeaders.indexOf("Product Name") + 1).setValue(record.productName);
      productionSheet.getRange(rowToUpdate, productionHeaders.indexOf("Material Used") + 1).setValue(materialNames);
      productionSheet.getRange(rowToUpdate, productionHeaders.indexOf("Quantity Used") + 1).setValue(quantities);
      productionSheet.getRange(rowToUpdate, productionHeaders.indexOf("Quantity Produced") + 1).setValue(record.quantityProduced);
      
      // Step 3: Apply new inventory changes
      if (inventoryRange) {
        const inventoryData = inventoryRange.getValues();
        const inventoryHeaders = inventoryData[0];
        
        const itemNameIndex = inventoryHeaders.indexOf("Item Name");
        const purchasedQtyIndex = inventoryHeaders.indexOf("Quantity Purchased");
        const usedQtyIndex = inventoryHeaders.indexOf("Quantity Used");
        const productQtyIndex = inventoryHeaders.indexOf("Product Quantity");
        const remainingQtyIndex = inventoryHeaders.indexOf("Remaining Quantity");
        const reorderRequiredIndex = inventoryHeaders.indexOf("Reorder Required");
        const reorderLevelIndex = inventoryHeaders.indexOf("Reorder Level");

        const soldQtyIndex = inventoryHeaders.indexOf("Quantity Sold");

        const updatedMaterials = new Map();
        
        // Update each raw material (reduce stock)
        record.materialsUsed.forEach(material => {
        // Skip if we've already updated this material (duplicate in same record)
        if (updatedMaterials.has(material.name)) {
          Logger.log(`Skipping duplicate material: ${material.name}`);
          return;
        }
        
        let materialUpdated = false;      
        
        // Search through inventory to find the exact row for this material
        for (let i = 1; i < inventoryData.length; i++) {
          const row = inventoryData[i];
          if (row[itemNameIndex] === material.name) {
            const rowNum = inventoryRange.getRow() + i;
            
            // Get CURRENT values directly from the sheet to ensure accuracy
            const currentPurchasedQty = inventorySheet.getRange(rowNum, purchasedQtyIndex + 1).getValue() || 0;
            const currentUsedQty = inventorySheet.getRange(rowNum, usedQtyIndex + 1).getValue() || 0;
            const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
            const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
            const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
            
            // FIXED: Calculate new values for raw material using correct formula
            const newUsedQty = currentUsedQty + material.quantityUsed;
            const newRemainingQty = currentPurchasedQty - newUsedQty - currentSoldQty; // CORRECT FORMULA
            const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
            
            // Update raw material inventory
            inventorySheet.getRange(rowNum, usedQtyIndex + 1).setValue(newUsedQty);
            inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
            inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
            
            materialUpdated = true;
            updatedMaterials.set(material.name, true);
            Logger.log(`Updated material: ${material.name}, Used: ${material.quantityUsed}, New Remaining: ${newRemainingQty} (Formula: ${currentPurchasedQty} - ${newUsedQty} - ${currentSoldQty})`);
            break;
          }
        }
        
        if (!materialUpdated) {
          Logger.log(`Raw material not found in inventory: ${material.name}`);
        }
      });
        
        // Update or create the finished product
        let productFound = false;

        // Get fresh inventory data again after raw material updates
        const updatedInventoryData = inventoryRange.getValues();

        for (let i = 1; i < updatedInventoryData.length; i++) {
          const row = updatedInventoryData[i];
          if (row[itemNameIndex] === record.productName) {
            const rowNum = inventoryRange.getRow() + i;
            
            // Get CURRENT values directly from the sheet
            const currentProductQty = inventorySheet.getRange(rowNum, productQtyIndex + 1).getValue() || 0;
            const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
            const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
            const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
            
            // FIXED: Calculate new values for finished product using correct formula
            const newProductQty = currentProductQty + record.quantityProduced;
            const newRemainingQty = newProductQty - currentSoldQty; // CORRECT FORMULA
            const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
            
            // Update finished product inventory
            inventorySheet.getRange(rowNum, productQtyIndex + 1).setValue(newProductQty);
            inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
            inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
            
            productFound = true;
            Logger.log(`Updated finished product: ${record.productName}, Produced: ${record.quantityProduced}, New Remaining: ${newRemainingQty} (Formula: ${newProductQty} - ${currentSoldQty})`);
            break;
          }
        }
        
        // Create new finished product if not found
        if (!productFound) {
          const newInventoryRow = [];
          
          inventoryHeaders.forEach(header => {
            switch(header) {
              case "Item ID": 
                newInventoryRow.push("P" + Math.floor(10000 + Math.random() * 90000)); 
                break;
              case "Item Type": newInventoryRow.push("Finished Goods"); break;
              case "Item Category": newInventoryRow.push("Production"); break;
              case "Item Name": newInventoryRow.push(record.productName); break;
              case "Quantity Purchased": newInventoryRow.push(0); break;
              case "Quantity Used": newInventoryRow.push(0); break;
              case "Product Quantity": newInventoryRow.push(record.quantityProduced); break;
              case "Quantity Sold": newInventoryRow.push(0); break;
              case "Remaining Quantity": newInventoryRow.push(record.quantityProduced); break;
              case "Reorder Level": newInventoryRow.push(0); break;
              case "Reorder Required": newInventoryRow.push("No"); break;
              default: newInventoryRow.push(''); break;
            }
          });
          
          inventorySheet.appendRow(newInventoryRow);
        }
      }
      
      return "success";
    } catch (error) {
      Logger.log("Error updating production record: " + error.toString());
      return "error: " + error.toString();
    }
  
  } catch (error) {
    console.error('Error in productionUpdateRecord:', error);
    throw error;
  }
}

// Delete production record
function productionDeleteRecord(token, recordId) {
  try {
    // Check delete permission
    checkProductionPermission(token, "CanDelete");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Production Records");
    const range = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    
    if (!range) {
      return "error";
    }
    
    try {
      // Get the record first to revert inventory changes
      const record = productionGetRecordById(recordId);
      if (!record) {
        return "not_found";
      }
      
      // Revert inventory changes
      revertInventoryChanges(record);
      
      // Delete the record
      const data = range.getValues();
      const headers = data[0];
      const idIndex = headers.indexOf("Production ID");
      
      // Find row index
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === recordId) {
          const rowNum = range.getRow() + i;
          sheet.deleteRow(rowNum);
          return "success";
        }
      }
      
      return "not_found";
    } catch (error) {
      Logger.log("Error deleting production record: " + error.toString());
      return "error";
    }
  
  } catch (error) {
    console.error('Error in productionDeleteRecord:', error);
    throw error;
  }
}

// Helper function to format dates
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  if (typeof dateValue === 'string') {
    return dateValue;
  }
  
  try {
    const date = new Date(dateValue);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return dateValue;
  }
}

// Get production record by ID (for editing)
function productionGetRecordById(token, recordId) {
  try {
    // Check view permission
    checkProductionPermission(token, "CanView");
        
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Production Records");
    const range = ss.getRangeByName("RANGEPRODUCTIONRECORDS");
    
    if (!range) {
      return null;
    }
    
    const data = range.getValues();
    const headers = data[0];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '') continue; // Skip empty rows
      
      const productionId = row[headers.indexOf("Production ID")];
      
      if (productionId === recordId) {
        let materialUsedStr = row[headers.indexOf("Material Used")] || '';
        let quantityUsedStr = row[headers.indexOf("Quantity Used")] || '';
        
        // Convert to strings to avoid split errors
        materialUsedStr = String(materialUsedStr);
        quantityUsedStr = String(quantityUsedStr);
        
        // Parse materials and quantities from comma-separated strings
        const materialNames = materialUsedStr.split(',').map(m => m.trim()).filter(m => m !== '');
        const quantities = quantityUsedStr.split(',').map(q => {
          const num = parseFloat(q.trim());
          return isNaN(num) ? 0 : num;
        });
        
        // Create materials array
        const materialsUsed = [];
        let totalQuantityUsed = 0;
        
        for (let j = 0; j < materialNames.length; j++) {
          if (materialNames[j]) {
            const quantity = quantities[j] || 0;
            materialsUsed.push({
              name: materialNames[j],
              quantityUsed: quantity
            });
            totalQuantityUsed += quantity;
          }
        }
        
        return {
          date: formatDate(row[headers.indexOf("Date")]),
          id: productionId,
          productName: row[headers.indexOf("Product Name")],
          materialsUsed: materialsUsed,
          totalQuantityUsed: totalQuantityUsed,
          quantityProduced: row[headers.indexOf("Quantity Produced")] || 0
        };
      }
    }
    
    return null;

  } catch (error) {
    console.error('Error in productionGetRecordById:', error);
    throw error;
  }
}

// Get raw materials from Inventory Items
function productionGetRawMaterials(token) {
  try {
    // Check view permission (since we need to view materials for production)
    checkProductionPermission(token, "CanView");
        
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");
    
    if (!range) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const materials = [];
    
    const itemNameIndex = headers.indexOf("Item Name");
    const purchasedQtyIndex = headers.indexOf("Quantity Purchased");
    const productQtyIndex = headers.indexOf("Product Quantity");
    const remainingQtyIndex = headers.indexOf("Remaining Quantity");
    const typeIndex = headers.indexOf("Item Type");
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '') continue; // Skip empty rows
      
      const purchasedQty = row[purchasedQtyIndex] || 0;
      const productQty = row[productQtyIndex] || 0;
      const itemType = row[typeIndex] || '';
      
      // Include items that are raw materials (have purchased quantity OR are explicitly marked as raw materials)
      if ((purchasedQty > 0 && productQty === 0) || itemType === "Raw Material") {
        materials.push({
          name: row[itemNameIndex],
          remainingQty: row[remainingQtyIndex] || 0,
          purchasedQty: purchasedQty,
          type: itemType
        });
      }
    }
    
    return materials;
  } catch (error) {
    console.error('Error in productionGetRawMaterials:', error);
    throw error;
  }
}

// Helper function to revert inventory changes for a record
function revertInventoryChanges(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName("Inventory Items");
  const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
  
  if (!inventoryRange) {
    return;
  }
  
  const inventoryData = inventoryRange.getValues();
  const inventoryHeaders = inventoryData[0];
  
  const itemNameIndex = inventoryHeaders.indexOf("Item Name");
  const purchasedQtyIndex = inventoryHeaders.indexOf("Quantity Purchased");
  const usedQtyIndex = inventoryHeaders.indexOf("Quantity Used");
  const productQtyIndex = inventoryHeaders.indexOf("Product Quantity");
  const soldQtyIndex = inventoryHeaders.indexOf("Quantity Sold");
  const remainingQtyIndex = inventoryHeaders.indexOf("Remaining Quantity");
  const reorderRequiredIndex = inventoryHeaders.indexOf("Reorder Required");
  const reorderLevelIndex = inventoryHeaders.indexOf("Reorder Level");
  
  // Revert raw material changes (add back used quantities) - FIXED CALCULATION
  record.materialsUsed.forEach(material => {
    for (let i = 1; i < inventoryData.length; i++) {
      const row = inventoryData[i];
      if (row[itemNameIndex] === material.name) {
        const rowNum = inventoryRange.getRow() + i;
        
        // Get current values
        const currentPurchasedQty = inventorySheet.getRange(rowNum, purchasedQtyIndex + 1).getValue() || 0;
        const currentUsedQty = inventorySheet.getRange(rowNum, usedQtyIndex + 1).getValue() || 0;
        const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
        const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
        const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
        
        // FIXED: Revert changes using correct formula
        const newUsedQty = Math.max(0, currentUsedQty - material.quantityUsed);
        const newRemainingQty = currentPurchasedQty - newUsedQty - currentSoldQty; // CORRECT FORMULA
        const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
        
        // Update inventory
        inventorySheet.getRange(rowNum, usedQtyIndex + 1).setValue(newUsedQty);
        inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
        inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
        break;
      }
    }
  });
  
  // Revert finished product changes (subtract produced quantity) - FIXED CALCULATION
  for (let i = 1; i < inventoryData.length; i++) {
    const row = inventoryData[i];
    if (row[itemNameIndex] === record.productName) {
      const rowNum = inventoryRange.getRow() + i;
      
      // Get current values
      const currentProductQty = inventorySheet.getRange(rowNum, productQtyIndex + 1).getValue() || 0;
      const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
      const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
      const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
      
      // FIXED: Revert changes using correct formula
      const newProductQty = Math.max(0, currentProductQty - record.quantityProduced);
      const newRemainingQty = newProductQty - currentSoldQty; // CORRECT FORMULA
      const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
      
      // Update inventory
      inventorySheet.getRange(rowNum, productQtyIndex + 1).setValue(newProductQty);
      inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
      inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
      break;
    }
  }
}

/**
 * UNIVERSAL Remaining Quantity Calculator
 * Use this function to ensure consistent calculations across all modules
 */
function calculateRemainingQuantity(itemType, purchasedQty, usedQty, productQty, soldQty) {
  purchasedQty = Number(purchasedQty) || 0;
  usedQty = Number(usedQty) || 0;
  productQty = Number(productQty) || 0;
  soldQty = Number(soldQty) || 0;
  
  let remaining = 0;
  
  // Determine item type and use appropriate formula
  if (itemType === "Raw Material" || (purchasedQty > 0 && productQty === 0)) {
    // Raw Material: Purchased - Used - Sold
    remaining = purchasedQty - usedQty - soldQty;
  } else if (itemType === "Finished Goods" || productQty > 0) {
    // Finished Good: Product Quantity - Sold
    remaining = productQty - soldQty;
  } else {
    // Fallback: Assume raw material logic
    remaining = purchasedQty - usedQty - soldQty;
  }
  
  // Ensure non-negative
  return Math.max(0, remaining);
}
