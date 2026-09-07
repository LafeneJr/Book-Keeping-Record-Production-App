/**
 * Entry point: show sidebar/modal
 */
function soShowSalesUI() {
  const html = HtmlService.createTemplateFromFile('sales')
      .evaluate()
      .setTitle('Sales Orders')
      .setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ===========================================
// SALES PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check sales permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanAdd', 'CanEdit', 'CanDelete')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkSalesPermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has sales permissions
    const salesPerms = session.permissions && session.permissions["sales"];
    if (!salesPerms) {
      throw new Error("You don't have permission to access sales orders.");
    }
    
    // Check specific action permission
    const hasPermission = salesPerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} sales orders.`);
    }
    
    return session;
  } catch (error) {
    console.error('Sales permission check error:', error);
    throw error;
  }
}


/**
 * Helper: get named range data as array of objects
 */
function soGetRangeDataAsObjects(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    throw new Error(`Named range "${rangeName}" not found. Please verify it exists in your sheet.`);
  }
  
  // Grab all values (first row = headers)
  const values = range.getValues();
  if (values.length < 2) {
    // only header row or empty range
    return [];
  }
  
  const headers = values[0];
  const rows    = values.slice(1)
    .filter(r => r.some(cell => cell !== '' && cell !== null));  // drop blank rows

  return rows.map(r => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = r[i]);
    return obj;
  });
}

/**
 * Fetch all SO rows
 */
function soGetAllSO(token) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Sales Orders"); // Make sure this matches your sheet name
    const range = ss.getRangeByName("RANGESALESORDERS");

    if (!range) {
      console.log("soGetAllSO: RANGESALESORDERS named range not found. Returning empty array.");
      return [];
    }

    try {
      const data = range.getValues();
      console.log(`soGetAllSO: Raw data has ${data.length} rows from range.`);

      if (!data || data.length <= 1) {
        console.log("soGetAllSO: No data found (only headers or empty).");
        return [];
      }

      const headers = data[0];
      const soIdIndex = headers.indexOf("Sales Order ID");

      if (soIdIndex === -1) {
        console.error("soGetAllSO: 'Sales Order ID' header not found.");
        return [];
      }

      const salesOrders = [];
      const getVal = (h, row) => {
        const index = headers.indexOf(h);
        return index !== -1 ? row[index] : '';
      };

      // NEW: Read data from bottom to top (newest records are at the bottom of the sheet)
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const soId = row[soIdIndex];
        
        // Skip empty rows
        if (!soId || soId === '' || soId === null || 
            row.every(cell => cell === '' || cell === null)) {
          continue;
        }

        let dateValue = getVal("Sales Order Date", row);
        // Format date as MM/dd/yyyy
        if (dateValue instanceof Date) {
          const month = (dateValue.getMonth() + 1).toString().padStart(2, '0');
          const day = dateValue.getDate().toString().padStart(2, '0');
          const year = dateValue.getFullYear();
          dateValue = `${month}/${day}/${year}`;
        } else if (typeof dateValue === "string") {
          // If it's already a string in a different format, try to convert it
          try {
            const dateObj = new Date(dateValue);
            if (!isNaN(dateObj)) {
              const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
              const day = dateObj.getDate().toString().padStart(2, '0');
              const year = dateObj.getFullYear();
              dateValue = `${month}/${day}/${year}`;
            }
          } catch (e) {
            // Keep original format if conversion fails
            console.log(`Could not convert date: ${dateValue}`);
          }
        }

        salesOrders.push({
          'Sales Order Date': dateValue,
          'Sales Order ID': soId,
          'Customer ID': getVal("Customer ID", row),
          'Customer Name': getVal("Customer Name", row),
          'State': getVal("State", row),
          'City': getVal("City", row),
          'Total Sales Order Amount': getVal("Total Sales Order Amount", row),
          'Total Received': getVal("Total Received", row),
          'Sales Order Balance': getVal("Sales Order Balance", row),
          'Payment Mode': getVal("Payment Mode", row),
          'Payment Status': getVal("Payment Status", row)
        });
      }
      
      console.log(`soGetAllSO: Returning ${salesOrders.length} sales orders (read from bottom to top, newest first).`);
      return salesOrders;
    } catch (e) {
      console.error("Error in soGetAllSO: " + e.message + " Stack: " + e.stack);
      return [];
    }
  
  } catch (error) {
    console.error('Error in soGetAllSO:', error);
    throw error;
  }
}

/**
 * Fetch customers list
 */
function soGetCustomers(token) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView");
    try {
      return soGetRangeDataAsObjects('RANGECUSTOMERS');
    } catch (err) {
      // log and rethrow so you see it in Execution log
      console.error(err);
      throw new Error('soGetCustomers failed: ' + err.message);
    }
  
  } catch (error) {
    console.error('Error in soGetCustomers:', error);
    throw error;
  }
}

/**
 * Fetch inventory items list
 */
function soGetInventoryItems(token) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView"); 
    try {
      const items = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');    
      return items.map(item => {
        // Calculate current stock based on item type
        let currentStock = 0;
        let remainingQty = 0;
        
        if (item['Item Type'] === 'Raw Material') {
          const purchased = parseFloat(item['Quantity Purchased']) || 0;
          const used = parseFloat(item['Quantity Used']) || 0;
          const sold = parseFloat(item['Quantity Sold']) || 0;
          currentStock = purchased - used - sold;
          remainingQty = currentStock;
        } else {
          // Production Item
          const productQty = parseFloat(item['Product Quantity']) || 0;
          const sold = parseFloat(item['Quantity Sold']) || 0;
          currentStock = productQty - sold;
          remainingQty = currentStock;
        }
        
        return {
          'Item ID': item['Item ID'],
          'Item Name': item['Item Name'],
          'Item Type': item['Item Type'],
          'Item Category': item['Item Category'],
          'Current Stock': Math.max(0, currentStock),
          'Remaining Quantity': Math.max(0, remainingQty)
        };
      }).filter(item => item['Current Stock'] > 0);
    } catch (err) {
      console.error(err);
      throw new Error('soGetInventoryItems failed: ' + err.message);
    }
  
  } catch (error) {
    console.error('Error in soGetInventoryItems:', error);
    throw error;
  }
}


/**
 * Fetch details for one SO
 */
function soGetSODetails(token, soID) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView");

    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    const raw = soGetRangeDataAsObjects('RANGESALESDETAILS')
                .filter(r => r['Sales Order ID'] === soID);
    
    // Remove Customer ID from each detail row
    return raw.map(row => {
      if (row['Sales Order Date'] instanceof Date) {
        row['Sales Order Date'] = Utilities.formatDate(row['Sales Order Date'], tz, 'MM/dd/yyyy');
      }
      // Create a new object without Customer ID
      const { 'Customer ID': removed, ...cleanRow } = row;
      return cleanRow;
    });
  
  } catch (error) {
    console.error('Error in soGetSODetails:', error);
    throw error;
  }
}

/**
 * Generate unique SO ID
 */
function soGenerateSOID(token) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    const data = soGetRangeDataAsObjects('RANGESALESDETAILS');
    const existing = data.map(r=>r['Sales Order ID']);
    let id;
    do {
      id = 'SO' + String(Math.floor(10000 + Math.random()*90000));
    } while (existing.indexOf(id) !== -1);
    return id;
  
  } catch (error) {
    console.error('Error in soGenerateSOID:', error);
    throw error;
  }
}

/**
 * Generate unique Detail ID
 */
function soGenerateSalesDetailID(token) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    const data = soGetRangeDataAsObjects('RANGESALESDETAILS');
    const existing = data.map(r => r['Sales Detail ID']);
    let id;
    do {
      id = 'SD' + String(Math.floor(10000 + Math.random() * 90000));
    } while (existing.includes(id));
    return id;
  
  } catch (error) {
    console.error('Error in soGenerateSalesDetailID:', error);
    throw error;
  }
}

/**
 * Save New SO: master + details + recalc all
 */
function soSaveNewSO(token, payload) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    const ss = SpreadsheetApp.getActive();
    const soRange = ss.getRangeByName('RANGESALESORDERS');
    const sdRange = ss.getRangeByName('RANGESALESDETAILS');
    const soSheet = soRange.getSheet();
    const sdSheet = sdRange.getSheet();

    try {
      // Step 1: append master row to Sales Orders
      const master = payload.master;
      
      // Create date in UTC to avoid timezone issues
      const dateParts = master.date.split('-');
      const soDate = new Date(Date.UTC(
        parseInt(dateParts[0]),
        parseInt(dateParts[1]) - 1,
        parseInt(dateParts[2]),
        12, 0, 0
      ));
      
      soSheet.appendRow([
        soDate,
        master.soID, 
        master.custID, 
        master.custNm,
        master.state, 
        master.city, 
        0,  // Total Sales Order Amount - will be recalculated
        master.amountReceived || 0,  // Total Received
        0,  // Sales Order Balance - will be recalculated
        master.paymentMode || 'Pending',  // Payment Mode
        master.paymentStatus || 'Pending'   // Payment Status
      ]);

      // Step 2: append each detail row (without Customer ID)
      payload.details.forEach(d => {
        const detailDateParts = d['Sales Order Date'].split('-');
        const detailDate = new Date(Date.UTC(
          parseInt(detailDateParts[0]),
          parseInt(detailDateParts[1]) - 1,
          parseInt(detailDateParts[2]),
          12, 0, 0
        ));
        
        sdSheet.appendRow([
          detailDate,
          d['Sales Order ID'], 
          d['Sales Detail ID'],
          d['Customer Name'], 
          d['State'], 
          d['City'], 
          d['Item ID'], 
          d['Item Type'], 
          d['Item Category'], 
          d['Item Name'], 
          d['Quantity Sold'], 
          d['Unit Price'],
          d['Shipping Fee'], 
          d['Total Sales Price']
        ]);
      });

      // Step 3: ALWAYS Create payment record (even for $0 payments)
      const paymentID = soGeneratePaymentID();
      
      const paymentData = {
        date: master.date,
        paymentID: paymentID,
        referenceID: master.soID,
        partyID: master.custID,
        paymentMode: master.paymentMode,
        amount: master.amountReceived || 0,
        status: master.paymentStatus
      };
      
      soCreatePaymentRecord(paymentData);

      // Step 4: recalc all metrics
      _soRecalcAll();
      return true;
    } catch (error) {
      console.error('Error saving sales order:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soSaveNewSO:', error);
    throw error;
  }
}


/**
 * Add new payment mode to Dimensions sheet
 */
function soAddNewPaymentMode(token, mode) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dimensionsRange = ss.getRangeByName('RANGEDIMENSIONS');
      
      if (!dimensionsRange) {
        throw new Error('RANGEDIMENSIONS named range not found. Please check your sheet setup.');
      }
      
      const dimensionsSheet = dimensionsRange.getSheet();
      
      // Get the headers to understand column order
      const headers = dimensionsRange.getValues()[0];
      
      // Find the Payment Mode column
      const paymentModeColIndex = headers.indexOf('Payment Mode');
      
      if (paymentModeColIndex === -1) {
        throw new Error('Payment Mode column not found in RANGEDIMENSIONS');
      }
      
      // Find the next empty row in the Payment Mode column
      const paymentModeCol = dimensionsRange.getColumn() + paymentModeColIndex;
      const paymentModeRange = dimensionsSheet.getRange(dimensionsRange.getRow(), paymentModeCol, dimensionsRange.getNumRows(), 1);
      const paymentModeValues = paymentModeRange.getValues();
      
      let emptyRow = -1;
      for (let i = 0; i < paymentModeValues.length; i++) {
        if (!paymentModeValues[i][0] || paymentModeValues[i][0] === '') {
          emptyRow = dimensionsRange.getRow() + i;
          break;
        }
      }
      
      // If no empty cell found in the range, append to the next row
      if (emptyRow === -1) {
        emptyRow = dimensionsRange.getRow() + dimensionsRange.getNumRows();
      }
      
      // Set the new payment mode
      dimensionsSheet.getRange(emptyRow, paymentModeCol).setValue(mode);
      
      return true;
    } catch (error) {
      console.error('Error adding payment mode:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soAddNewPaymentMode:', error);
    throw error;
  }
}


/**
 * Add new payment status to Dimensions sheet
 */
function soAddNewPaymentStatus(token, status) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const dimensionsRange = ss.getRangeByName('RANGEDIMENSIONS');
      
      if (!dimensionsRange) {
        throw new Error('RANGEDIMENSIONS named range not found. Please check your sheet setup.');
      }
      
      const dimensionsSheet = dimensionsRange.getSheet();
      
      // Get the headers to understand column order
      const headers = dimensionsRange.getValues()[0];
      
      // Find the Payment Status column
      const paymentStatusColIndex = headers.indexOf('Payment Status');
      
      if (paymentStatusColIndex === -1) {
        throw new Error('Payment Status column not found in RANGEDIMENSIONS');
      }
      
      // Find the next empty row in the Payment Status column
      const paymentStatusCol = dimensionsRange.getColumn() + paymentStatusColIndex;
      const paymentStatusRange = dimensionsSheet.getRange(dimensionsRange.getRow(), paymentStatusCol, dimensionsRange.getNumRows(), 1);
      const paymentStatusValues = paymentStatusRange.getValues();
      
      let emptyRow = -1;
      for (let i = 0; i < paymentStatusValues.length; i++) {
        if (!paymentStatusValues[i][0] || paymentStatusValues[i][0] === '') {
          emptyRow = dimensionsRange.getRow() + i;
          break;
        }
      }
      
      // If no empty cell found in the range, append to the next row
      if (emptyRow === -1) {
        emptyRow = dimensionsRange.getRow() + dimensionsRange.getNumRows();
      }
      
      // Set the new payment status
      dimensionsSheet.getRange(emptyRow, paymentStatusCol).setValue(status);
      
      return true;
    } catch (error) {
      console.error('Error adding payment status:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soAddNewPaymentStatus:', error);
    throw error;
  }
}


  /**
   * Update existing SalesDetails rows, then recalc all metrics
   */
function soUpdateSODetails(token, updatedDetails, soID, updatedAmountReceived) {
  try {
    // Check edit permission
    checkSalesPermission(token, "CanEdit");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the OLD details first to revert inventory changes
    const oldDetails = soGetRangeDataAsObjects('RANGESALESDETAILS')
      .filter(r => r['Sales Order ID'] === soID);
    
    // Step 1: Revert old inventory changes (add back previously sold quantities)
    revertSalesInventoryChanges(oldDetails);
    
    // Step 2: Update Sales Details with new data
    const sdRange = ss.getRangeByName('RANGESALESDETAILS');
    const sdSheet = sdRange.getSheet();
    const sdStartRow = sdRange.getRow();
    const sdAllValues = sdRange.getValues();
    const sdHeaders = sdAllValues[0];
    const sdData = sdAllValues.slice(1);
    
    const detailIdCol = sdHeaders.indexOf('Sales Detail ID');
    
    // Calculate new total from updated details
    let newTotalAmount = 0;
    updatedDetails.forEach(update => {
      newTotalAmount += (parseFloat(update['Total Sales Price']) || 0);
      
      const detailId = update['Sales Detail ID'];
      const rowIndex = sdData.findIndex(row => row[detailIdCol] === detailId);
      
      if (rowIndex !== -1) {
        const sheetRow = sdStartRow + 1 + rowIndex;
        
        // Update ALL fields including item details
        const fields = [
          'Item ID', 'Item Type', 'Item Category', 'Item Name',
          'Quantity Sold', 'Unit Price', 'Shipping Fee', 'Total Sales Price'
        ];
        
        fields.forEach(field => {
          const colIndex = sdHeaders.indexOf(field);
          if (colIndex !== -1 && update[field] !== undefined) {
            sdSheet.getRange(sheetRow, sdRange.getColumn() + colIndex).setValue(update[field]);
          }
        });
      }
    });
    
    // Step 3: Apply new inventory changes (deduct new sold quantities)
    applySalesInventoryChanges(updatedDetails);
    
    // Step 4: Update Sales Orders
    const soRange = ss.getRangeByName('RANGESALESORDERS');
    const soSheet = soRange.getSheet();
    const soStartRow = soRange.getRow();
    const soAllValues = soRange.getValues();
    const soHeaders = soAllValues[0];
    const soData = soAllValues.slice(1);
    
    const soIDCol = soHeaders.indexOf('Sales Order ID');
    const amountReceivedCol = soHeaders.indexOf('Total Received');
    const totalAmountCol = soHeaders.indexOf('Total Sales Order Amount');
    const paymentStatusCol = soHeaders.indexOf('Payment Status');
    const balanceCol = soHeaders.indexOf('Sales Order Balance');
    
    const soRowIndex = soData.findIndex(row => row[soIDCol] === soID);
    
    let newPaymentStatus = 'Unpaid';
    
    if (soRowIndex !== -1) {
      const soSheetRow = soStartRow + 1 + soRowIndex;
      
      // Update all fields at once
      soSheet.getRange(soSheetRow, soRange.getColumn() + totalAmountCol).setValue(newTotalAmount);
      soSheet.getRange(soSheetRow, soRange.getColumn() + amountReceivedCol).setValue(updatedAmountReceived);
      
      const newBalance = newTotalAmount - updatedAmountReceived;
      soSheet.getRange(soSheetRow, soRange.getColumn() + balanceCol).setValue(newBalance);
      
      // Calculate Payment Status
      if (updatedAmountReceived === 0) {
        newPaymentStatus = 'Unpaid';
      } else if (updatedAmountReceived >= newTotalAmount) {
        newPaymentStatus = 'Paid';
      } else {
        newPaymentStatus = 'Partial';
      }
      
      soSheet.getRange(soSheetRow, soRange.getColumn() + paymentStatusCol).setValue(newPaymentStatus);
      
      console.log('Updated Sales Order:', soID);
      console.log('Total Amount:', newTotalAmount);
      console.log('Amount Received:', updatedAmountReceived);
      console.log('Balance:', newBalance);
      console.log('Payment Status:', newPaymentStatus);
      
      // Update the payment record properly
      updateSalesPaymentRecord(soID, updatedAmountReceived, newPaymentStatus, newTotalAmount);
    }
    
    // Recalculate all other metrics
    _soRecalcAll();
    return true;
  
  } catch (error) {
    console.error('Error in soUpdateSODetails:', error);
    throw error;
  }
}

/**
 * Helper function to revert inventory changes for sales details
 */
function revertSalesInventoryChanges(details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName("Inventory Items");
  const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
  
  if (!inventoryRange) return;
  
  const inventoryData = inventoryRange.getValues();
  const inventoryHeaders = inventoryData[0];
  
  const itemNameIndex = inventoryHeaders.indexOf("Item Name");
  const soldQtyIndex = inventoryHeaders.indexOf("Quantity Sold");
  const remainingQtyIndex = inventoryHeaders.indexOf("Remaining Quantity");
  const reorderRequiredIndex = inventoryHeaders.indexOf("Reorder Required");
  const reorderLevelIndex = inventoryHeaders.indexOf("Reorder Level");
  
  details.forEach(detail => {
    const itemName = detail['Item Name'];
    const oldQtySold = parseInt(detail['Quantity Sold']) || 0;
    
    for (let i = 1; i < inventoryData.length; i++) {
      const row = inventoryData[i];
      if (row[itemNameIndex] === itemName) {
        const rowNum = inventoryRange.getRow() + i;
        
        // Get current values
        const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
        const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
        const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
        
        // Revert: Subtract the old sold quantity
        const newSoldQty = Math.max(0, currentSoldQty - oldQtySold);
        const newRemainingQty = currentRemainingQty + oldQtySold; // Add back the quantity
        const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
        
        // Update inventory
        inventorySheet.getRange(rowNum, soldQtyIndex + 1).setValue(newSoldQty);
        inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
        inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
        
        break;
      }
    }
  });
}

/**
 * Helper function to apply inventory changes for sales details
 */
function applySalesInventoryChanges(details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventorySheet = ss.getSheetByName("Inventory Items");
  const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
  
  if (!inventoryRange) return;
  
  const inventoryData = inventoryRange.getValues();
  const inventoryHeaders = inventoryData[0];
  
  const itemNameIndex = inventoryHeaders.indexOf("Item Name");
  const soldQtyIndex = inventoryHeaders.indexOf("Quantity Sold");
  const remainingQtyIndex = inventoryHeaders.indexOf("Remaining Quantity");
  const reorderRequiredIndex = inventoryHeaders.indexOf("Reorder Required");
  const reorderLevelIndex = inventoryHeaders.indexOf("Reorder Level");
  
  details.forEach(detail => {
    const itemName = detail['Item Name'];
    const newQtySold = parseInt(detail['Quantity Sold']) || 0;
    
    for (let i = 1; i < inventoryData.length; i++) {
      const row = inventoryData[i];
      if (row[itemNameIndex] === itemName) {
        const rowNum = inventoryRange.getRow() + i;
        
        // Get current values
        const currentSoldQty = inventorySheet.getRange(rowNum, soldQtyIndex + 1).getValue() || 0;
        const currentRemainingQty = inventorySheet.getRange(rowNum, remainingQtyIndex + 1).getValue() || 0;
        const reorderLevel = inventorySheet.getRange(rowNum, reorderLevelIndex + 1).getValue() || 0;
        
        // Apply: Add the new sold quantity
        const newSoldQty = currentSoldQty + newQtySold;
        const newRemainingQty = Math.max(0, currentRemainingQty - newQtySold); // Subtract the quantity
        const newReorderRequired = newRemainingQty < reorderLevel ? "Yes" : "No";
        
        // Update inventory
        inventorySheet.getRange(rowNum, soldQtyIndex + 1).setValue(newSoldQty);
        inventorySheet.getRange(rowNum, remainingQtyIndex + 1).setValue(newRemainingQty);
        inventorySheet.getRange(rowNum, reorderRequiredIndex + 1).setValue(newReorderRequired);
        
        break;
      }
    }
  });
}

function updateSalesPaymentRecord(soID, amountReceived, paymentStatus, totalAmount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsRange = ss.getRangeByName('RANGEPAYMENTS');
    
    if (!paymentsRange) {
      console.error('RANGEPAYMENTS named range not found');
      return false;
    }
    
    const paymentsSheet = paymentsRange.getSheet();
    const startRow = paymentsRange.getRow();
    const allValues = paymentsRange.getValues();
    const headers = allValues[0];
    const data = allValues.slice(1);
    
    const referenceIDCol = headers.indexOf('Reference ID');
    const transactionTypeCol = headers.indexOf('Transaction Type');
    const amountCol = headers.indexOf('Amount');
    const statusCol = headers.indexOf('Status');
    
    // Find the payment record for this Sales Order (Sale transactions only)
    const rowIndex = data.findIndex(row => 
      row[referenceIDCol] === soID && 
      row[transactionTypeCol] === 'Sale'
    );
    
    if (rowIndex !== -1) {
      // Update existing record
      const sheetRow = startRow + 1 + rowIndex;
      
      if (amountCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
          .setValue(amountReceived);
      }
      
      if (statusCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
          .setValue(paymentStatus);
      }
      
      console.log('Updated existing payment record for SO:', soID);
      return true;
    } else {
      // Create new payment record if it doesn't exist
      console.log('Creating new payment record for SO:', soID);
      
      // Get customer info from sales order
      const soData = soGetRangeDataAsObjects('RANGESALESORDERS');
      const salesOrder = soData.find(so => so['Sales Order ID'] === soID);
      
      if (salesOrder) {
        const paymentData = {
          date: salesOrder['Sales Order Date'],
          paymentID: soGeneratePaymentID(),
          referenceID: soID,
          partyID: salesOrder['Customer ID'],
          paymentMode: salesOrder['Payment Mode'] || 'Pending',
          amount: amountReceived,
          status: paymentStatus
        };
        
        soCreatePaymentRecord(paymentData);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Error updating sales payment record:', error);
    return false;
  }
}


/**
 * Delete sales detail by ID
 */
function soDeleteSalesDetail(token, detailId, soId) {
  try {
    // Check delete permission
    checkSalesPermission(token, "CanDelete");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sdRange = ss.getRangeByName('RANGESALESDETAILS');
      const soRange = ss.getRangeByName('RANGESALESORDERS');
      const sdSheet = sdRange.getSheet();
      const soSheet = soRange.getSheet();
      
      const startRow = sdRange.getRow();
      const vals = sdRange.getValues();
      const headers = vals[0];
      const data = vals.slice(1);
      
      const detailCol = headers.indexOf('Sales Detail ID');
      
      // Delete the detail row
      const rowIdx = data.findIndex(r => r[detailCol] === detailId);
      if (rowIdx > -1) {
        sdSheet.deleteRow(startRow + 1 + rowIdx);
      } else {
        throw new Error('Sales detail not found: ' + detailId);
      }
      
      // Check if any details remain for this SO
      const remainingDetails = soGetRangeDataAsObjects('RANGESALESDETAILS')
        .filter(r => r['Sales Order ID'] === soId);
      
      // If no details remain, delete the master record too
      if (remainingDetails.length === 0) {
        const soVals = soRange.getValues();
        const soHeaders = soVals[0];
        const soData = soVals.slice(1);
        const soIDCol = soHeaders.indexOf('Sales Order ID');
        
        const soRowIdx = soData.findIndex(r => r[soIDCol] === soId);
        if (soRowIdx > -1) {
          soSheet.deleteRow(soRange.getRow() + 1 + soRowIdx);
          console.log('Deleted master Sales Order: ' + soId);
          
          // SYNC: Delete the payment record
          soDeletePaymentRecord(soId);
        }
      } else {
        // If details remain, recalc and update payment record
        // Trigger a recalculation which will update the payment record
        _soRecalcAll();
        
        // Get updated sales order data to update payment record
        const soData = soGetRangeDataAsObjects('RANGESALESORDERS');
        const updatedSO = soData.find(so => so['Sales Order ID'] === soId);
        
        if (updatedSO) {
          soUpdatePaymentRecord(
            soId, 
            parseFloat(updatedSO['Total Received']) || 0,
            updatedSO['Payment Status'] || 'Unpaid'
          );
        }
      }
      
      // Recalculate all metrics
      _soRecalcAll();
      return true;
      
    } catch (error) {
      console.error('Error deleting sales detail:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soDeleteSalesDetail:', error);
    throw error;
  }
}

/**
 * Delete entire sales order with all associated records
 */
function soDeleteEntireSalesOrder(token, soID) {
  try {
    // Check delete permission
    checkSalesPermission(token, "CanDelete");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      // Delete from Sales Orders (master)
      const soRange = ss.getRangeByName('RANGESALESORDERS');
      const soSheet = soRange.getSheet();
      const soStartRow = soRange.getRow();
      const soVals = soRange.getValues();
      const soHeaders = soVals[0];
      const soData = soVals.slice(1);
      const soIDCol = soHeaders.indexOf('Sales Order ID');
      
      const soRowIdx = soData.findIndex(r => r[soIDCol] === soID);
      if (soRowIdx > -1) {
        soSheet.deleteRow(soStartRow + 1 + soRowIdx);
        console.log('Deleted master Sales Order: ' + soID);
      }

      // Delete from Sales Details (all line items)
      const sdRange = ss.getRangeByName('RANGESALESDETAILS');
      const sdSheet = sdRange.getSheet();
      const sdStartRow = sdRange.getRow();
      const sdVals = sdRange.getValues();
      const sdHeaders = sdVals[0];
      const sdData = sdVals.slice(1);
      const sdSoIDCol = sdHeaders.indexOf('Sales Order ID');
      
      // Delete from bottom to top to avoid index shifting
      const sdRowsToDelete = [];
      sdData.forEach((row, idx) => {
        if (row[sdSoIDCol] === soID) {
          sdRowsToDelete.push(sdStartRow + 1 + idx);
        }
      });
      
      // Sort in descending order and delete
      sdRowsToDelete.sort((a, b) => b - a).forEach(rowNum => {
        sdSheet.deleteRow(rowNum);
      });
      
      console.log(`Deleted ${sdRowsToDelete.length} sales detail records for SO: ${soID}`);

      // Delete payment record
      soDeletePaymentRecord(soID);
      
      // Recalculate ALL metrics to ensure data integrity
      _soRecalcAll();
      
      console.log(`Successfully deleted entire sales order: ${soID}`);
      return true;
      
    } catch (error) {
      console.error('Error deleting entire sales order:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soDeleteEntireSalesOrder:', error);
    throw error;
  }
}

  /**
   * Internal: run all recalculation routines in order
   */
  function _soRecalcAll() {
    soCalcTotalSOAmount();
    soCalcSOBalance();
    soUpdateQtySold();
    soCalcRemainingQty();
    soCalcReorderRequired();
    soCalcTotalSales();
    soCalcTotalReceipts();
    soCalcBalanceReceivable();
        
  }

  /**
   * Sum Total Sales Price per SO and write to SalesOrders[Total SO Amount]
   */
  function soCalcTotalSOAmount() {
    const ss = SpreadsheetApp.getActive();
    const soRange = ss.getRangeByName('RANGESALESORDERS');
    const sheet = soRange.getSheet();
    const startRow = soRange.getRow();
    const vals = soRange.getValues();
    const headers = vals[0];
    const data = vals.slice(1);

    const soIDCol = headers.indexOf('Sales Order ID');
    const totalCol = headers.indexOf('Total Sales Order Amount');
    const sdData = soGetRangeDataAsObjects('RANGESALESDETAILS');

    data.forEach((row, i) => {
      const soID = row[soIDCol];
      
      // Skip rows without valid Sales Order IDs
      if (!soID || soID === '' || soID === null || soID === undefined) {
        return; // Skip this row
      }

      // Calculate total for this specific sales order
      const sum = sdData
        .filter(d => d['Sales Order ID'] === soID)
        .reduce((acc, cur) => acc + (parseFloat(cur['Total Sales Price']) || 0), 0);
      
      // Only write if we have a valid number
      if (!isNaN(sum)) {
        sheet
          .getRange(startRow + 1 + i, soRange.getColumn() + totalCol)
          .setValue(sum);
      } else {
        // Set to 0 if no sales details found
        sheet
          .getRange(startRow + 1 + i, soRange.getColumn() + totalCol)
          .setValue(0);
      }
    });
  }

  /**
   * Calculate SO Balance = Total SO Amount - Total Received
   */
function soCalcSOBalance() {
  const ss = SpreadsheetApp.getActive();
  const soRange = ss.getRangeByName('RANGESALESORDERS');
  const sheet = soRange.getSheet();
  const startRow = soRange.getRow();
  const vals = soRange.getValues();
  const headers = vals[0];
  const data = vals.slice(1);

  const totalCol = headers.indexOf('Total Sales Order Amount');
  const receivedCol = headers.indexOf('Total Received');
  const balCol = headers.indexOf('Sales Order Balance');
  const soIDCol = headers.indexOf('Sales Order ID');

  data.forEach((row, i) => {
    const soID = row[soIDCol];
    
    // Skip rows without valid Sales Order IDs
    if (!soID || soID === '' || soID === null) {
      return;
    }

    const totalAmt = parseFloat(row[totalCol]) || 0;
    const received = parseFloat(row[receivedCol]) || 0;
    const bal = totalAmt - received;
    
    sheet
      .getRange(startRow + 1 + i, soRange.getColumn() + balCol)
      .setValue(bal);
  });
}

  /**
   * Update InventoryItems[QTY Sold] via SUMIF on SalesDetails[QTY Sold]
   */
  function soUpdateQtySold() {
    const ss = SpreadsheetApp.getActive();
    const invRange = ss.getRangeByName('RANGEINVENTORYITEMS');
    const sheet = invRange.getSheet();
    const startRow = invRange.getRow();
    const vals = invRange.getValues();
    const headers = vals[0];
    const data = vals.slice(1);

    const itemIDCol = headers.indexOf('Item ID');
    const soldCol = headers.indexOf('Quantity Sold');
    const sdData = soGetRangeDataAsObjects('RANGESALESDETAILS');

    // Create a fresh map to track total sold per item
    const itemSoldMap = {};
    
    // Reset all sold quantities to 0 first
    data.forEach((row, i) => {
      const itemID = row[itemIDCol];
      if (itemID && itemID !== '') {
        itemSoldMap[itemID] = 0;
      }
    });

    // Calculate total sold for each item from all sales
    sdData.forEach(sale => {
      const itemId = sale['Item ID'];
      const qtySold = parseInt(sale['Quantity Sold']) || 0;
      
      if (itemId && itemId !== '' && itemSoldMap.hasOwnProperty(itemId)) {
        itemSoldMap[itemId] += qtySold;
      }
    });

    // Update each inventory item
    data.forEach((row, i) => {
      const itemID = row[itemIDCol];
      
      // Skip rows without valid Item IDs
      if (!itemID || itemID === '' || itemID === null || itemID === undefined) {
        return;
      }

      const totalSold = itemSoldMap[itemID] || 0;
      
      // Always update with the calculated total
      sheet.getRange(startRow + 1 + i, invRange.getColumn() + soldCol).setValue(totalSold);
    });
  }

  /**
   * Calculate Remaining QTY = QTY Purchased - QTY Sold
   */
  function soCalcRemainingQty() {
    const ss = SpreadsheetApp.getActive();
    const invRange = ss.getRangeByName('RANGEINVENTORYITEMS');
    const sheet = invRange.getSheet();
    const startRow = invRange.getRow();
    const vals = invRange.getValues();
    const headers = vals[0];
    const data = vals.slice(1);

    const purchasedCol = headers.indexOf('Quantity Purchased');
    const soldCol = headers.indexOf('Quantity Sold');
    const usedCol = headers.indexOf('Quantity Used');
    const productQtyCol = headers.indexOf('Product Quantity');
    const remainCol = headers.indexOf('Remaining Quantity');
    const itemTypeCol = headers.indexOf('Item Type');

    let updateCount = 0;
    
    data.forEach((row, i) => {
      const itemID = row[headers.indexOf('Item ID')];
      const itemType = itemTypeCol !== -1 ? row[itemTypeCol] : '';
      
      // Skip rows without Item IDs
      if (!itemID || itemID === '' || itemID === null) {
        return;
      }

      const purchased = parseFloat(row[purchasedCol]) || 0;
      const used = parseFloat(row[usedCol]) || 0;
      const sold = parseFloat(row[soldCol]) || 0;
      const productQty = parseFloat(row[productQtyCol]) || 0;
      
      let rem = 0;
      
      if (itemType === 'Raw Material' || (purchased > 0 && productQty === 0)) {
        // For Raw Materials: Quantity Purchased - Quantity Used - Quantity Sold
        rem = purchased - used - sold;
      } else {
        // For Production Items: Product Quantity - Quantity Sold
        rem = productQty - sold;
      }
      
      // Ensure remaining quantity doesn't go negative
      rem = Math.max(0, rem);
      
      sheet.getRange(startRow + 1 + i, invRange.getColumn() + remainCol).setValue(rem);
      updateCount++;
    });
    
    console.log(`soCalcRemainingQty: Updated ${updateCount} items with consistent formulas`);
  }

  /**
   * Flag Reorder Required if Remaining QTY < Reorder Level
   */
  function soCalcReorderRequired() {
    const ss = SpreadsheetApp.getActive();
    const invRange = ss.getRangeByName('RANGEINVENTORYITEMS');
    const sheet = invRange.getSheet();
    const startRow = invRange.getRow();
    const vals = invRange.getValues();
    const headers = vals[0];
    const data = vals.slice(1);

    const remainCol = headers.indexOf('Remaining Quantity');
    const levelCol = headers.indexOf('Reorder Level');
    const reqCol = headers.indexOf('Reorder Required');

    data.forEach((row, i) => {
      const itemID = row[headers.indexOf('Item ID')];
      
      // Skip rows without Item IDs
      if (!itemID || itemID === '' || itemID === null) {
        return;
      }

      const remaining = parseFloat(row[remainCol]) || 0;
      const reorderLevel = parseFloat(row[levelCol]) || 0;
      const flag = (remaining < reorderLevel) ? 'Yes' : 'No';
      
      sheet
        .getRange(startRow + 1 + i, invRange.getColumn() + reqCol)
        .setValue(flag);
    });
  }

/**
 * Sum Total Sales per Customer from Sales Orders, write to Customers[Total Sales]
 */
function soCalcTotalSales() {
  const ss = SpreadsheetApp.getActive();
  const custRange = ss.getRangeByName('RANGECUSTOMERS');
  const soRange = ss.getRangeByName('RANGESALESORDERS');
  const custSheet = custRange.getSheet();
  const startRow = custRange.getRow();
  const vals = custRange.getValues();
  const headers = vals[0];
  const data = vals.slice(1);

  const custIDCol = headers.indexOf('Customer ID');
  const totalSalesCol = headers.indexOf('Total Sales');
  
  // Get sales orders data instead of sales details
  const soData = soGetRangeDataAsObjects('RANGESALESORDERS');

  data.forEach((row, i) => {
    const custID = row[custIDCol];
    
    // Skip rows without valid Customer IDs
    if (!custID || custID === '' || custID === null || custID === undefined) {
      return; // Skip this row
    }

    // Calculate total sales for this specific customer from Sales Orders
    const sum = soData
      .filter(so => so['Customer ID'] === custID)
      .reduce((acc, cur) => acc + (parseFloat(cur['Total Sales Order Amount']) || 0), 0);
    
    // Only write if we have a valid number
    if (!isNaN(sum)) {
      custSheet
        .getRange(startRow + 1 + i, custRange.getColumn() + totalSalesCol)
        .setValue(sum);
    }
  });
}

/**
 * Sum Total Receipts per Customer from Sales Orders, write to Customers[Total Receipts]
 */
function soCalcTotalReceipts() {
  const ss = SpreadsheetApp.getActive();
  const custRange = ss.getRangeByName('RANGECUSTOMERS');
  const soRange = ss.getRangeByName('RANGESALESORDERS');
  const custSheet = custRange.getSheet();
  const startRow = custRange.getRow();
  const vals = custRange.getValues();
  const headers = vals[0];
  const data = vals.slice(1);

  const custIDCol = headers.indexOf('Customer ID');
  const totalReceiptsCol = headers.indexOf('Total Receipts');
  
  // Get sales orders data
  const soData = soGetRangeDataAsObjects('RANGESALESORDERS');

  data.forEach((row, i) => {
    const custID = row[custIDCol];
    
    // Skip rows without valid Customer IDs
    if (!custID || custID === '' || custID === null || custID === undefined) {
      return; // Skip this row
    }

    // Calculate total receipts for this specific customer from Sales Orders
    const sum = soData
      .filter(so => so['Customer ID'] === custID)
      .reduce((acc, cur) => acc + (parseFloat(cur['Total Received']) || 0), 0);
    
    // Only write if we have a valid number
    if (!isNaN(sum)) {
      custSheet
        .getRange(startRow + 1 + i, custRange.getColumn() + totalReceiptsCol)
        .setValue(sum);
    } else {
      // Set to 0 if no receipts found
      custSheet
        .getRange(startRow + 1 + i, custRange.getColumn() + totalReceiptsCol)
        .setValue(0);
    }
  });
}

  /**
   * Calculate Balance Receivable = Total Sales - Total Receipts
   */
  function soCalcBalanceReceivable() {
    const ss = SpreadsheetApp.getActive();
    const custRange = ss.getRangeByName('RANGECUSTOMERS');
    const sheet = custRange.getSheet();
    const startRow = custRange.getRow();
    const vals = custRange.getValues();
    const headers = vals[0];
    const data = vals.slice(1);

    const totalSalesCol = headers.indexOf('Total Sales');
    const receiptsCol = headers.indexOf('Total Receipts');
    const balanceRecCol = headers.indexOf('Balance Receivable');
    const custIDCol = headers.indexOf('Customer ID'); // Added to check for valid customers

    data.forEach((row, i) => {
      const custID = row[custIDCol];
      
      // Skip rows without valid Customer IDs
      if (!custID || custID === '' || custID === null) {
        return;
      }

      const totalSales = parseFloat(row[totalSalesCol]) || 0;
      const totalReceipts = parseFloat(row[receiptsCol]) || 0;
      const bal = totalSales - totalReceipts;
      
      sheet
        .getRange(startRow + 1 + i, custRange.getColumn() + balanceRecCol)
        .setValue(bal);
    });
  }

  /**
 * Generate unique Customer ID
 */
function soGenerateCustomerID(token) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    const data = soGetRangeDataAsObjects('RANGECUSTOMERS');
    const existing = data.map(r => r['Customer ID']);
    let id;
    do {
      id = 'C' + String(Math.floor(10000 + Math.random() * 90000));
    } while (existing.includes(id));
    return id;
  
  } catch (error) {
    console.error('Error in soGenerateCustomerID:', error);
    throw error;
  }
}

/**
 * Add new customer to Customers sheet with proper column mapping
 */
function soAddNewCustomer(token, customerData) {
  try {
    // Check add permission
    checkSalesPermission(token, "CanAdd");

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const custRange = ss.getRangeByName('RANGECUSTOMERS');
      const sheet = custRange.getSheet();
      
      // Get the headers to understand column order
      const headers = custRange.getValues()[0];
      
      // Create a row array with exact column matching
      const newRow = [];
      
      // Map each column based on header position
      headers.forEach(header => {
        switch(header) {
          case 'Customer ID':
            newRow.push(customerData['Customer ID'] || '');
            break;
          case 'Customer Name':
            newRow.push(customerData['Customer Name'] || '');
            break;
          case 'Customer Contact':
            newRow.push(''); // Leave empty
            break;
          case 'Customer Email':
            newRow.push(''); // Leave empty
            break;
          case 'State':
            newRow.push(customerData['State'] || '');
            break;
          case 'City':
            newRow.push(customerData['City'] || '');
            break;
          case 'Customer Address':
            newRow.push(''); // Leave empty
            break;
          case 'Total Sales':
            newRow.push(0); // Start with 0
            break;
          case 'Total Receipts':
            newRow.push(0); // Start with 0
            break;
          case 'Balance Receivable':
            newRow.push(0); // Start with 0
            break;
          default:
            newRow.push(''); // For any unexpected columns
        }
      });
      
      // Append the new customer row with proper column alignment
      sheet.appendRow(newRow);
      
      return true;
    } catch (error) {
      console.error('Error adding new customer:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in soAddNewCustomer:', error);
    throw error;
  }
}

/**
 * Get payment modes from Dimensions sheet
 */
function soGetPaymentModes(token) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView");

    try {
      const dimensionsData = soGetRangeDataAsObjects('RANGEDIMENSIONS');
      
      if (!dimensionsData || dimensionsData.length === 0) {
        console.log('No data found in RANGEDIMENSIONS');
        
      }
      
      // Extract payment modes, filtering out empty values
      const paymentModes = dimensionsData
        .map(row => row['Payment Mode'])
        .filter(mode => mode && mode !== '' && mode !== null && mode !== undefined);
      
      console.log('Found payment modes:', paymentModes);
      
      // If no payment modes found, return fallback
      
      
      return paymentModes;
    } catch (error) {
      console.error('Error getting payment modes:', error);
      
    }
  
  } catch (error) {
    console.error('Error in soGetPaymentModes:', error);
    throw error;
  }
}

/**
 * Generate unique Payment ID
 */
function soGeneratePaymentID() {
  const data = soGetRangeDataAsObjects('RANGEPAYMENTS');
  const existing = data.map(r => r['Payment ID']);
  let id;
  do {
    id = 'PMT' + String(Math.floor(10000 + Math.random() * 90000));
  } while (existing.includes(id));
  return id;
}

/**
 * Create payment record in Payments sheet when sales order is created
 */
function soCreatePaymentRecord(paymentData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsSheet = ss.getSheetByName("Payments");
    const paymentsRange = ss.getRangeByName("RANGEPAYMENTS");

    if (!paymentsRange) {
      throw new Error('RANGEPAYMENTS named range not found.');
    }

    // Get headers
    const headers = paymentsRange.getValues()[0];
    const timezone = ss.getSpreadsheetTimeZone();
    
    // Convert new payment date
    const dateParts = paymentData.date.split('-');
    const newPaymentDate = new Date(Date.UTC(
      parseInt(dateParts[0]),
      parseInt(dateParts[1]) - 1,
      parseInt(dateParts[2]),
      12, 0, 0
    ));
    
    // Create the new row data
    const newRow = [];
    headers.forEach(header => {
      switch(header) {
        case 'Date':
          const formattedDate = Utilities.formatDate(newPaymentDate, timezone, 'MM/dd/yyyy');
          newRow.push(formattedDate);
          break;
        case 'Payment ID':
          newRow.push(paymentData.paymentID || '');
          break;
        case 'Transaction Type':
          newRow.push('Sale');
          break;
        case 'Reference ID':
          newRow.push(paymentData.referenceID || '');
          break;
        case 'Party ID':
          newRow.push(paymentData.partyID || '');
          break;
        case 'Payment Mode':
          newRow.push(paymentData.paymentMode || '');
          break;
        case 'Amount':
          newRow.push(paymentData.amount || 0);
          break;
        case 'Status':
          newRow.push(paymentData.status || '');
          break;
        default:
          newRow.push('');
      }
    });

    // FIXED: Use simple append like Purchase system - no complex row finding logic
    paymentsSheet.appendRow(newRow);
    
    console.log('soCreatePaymentRecord: Successfully appended payment record for SO:', paymentData.referenceID);
    
  } catch (error) {
    console.error('Error creating payment record:', error);
    throw error;
  }
}


/**
 * Fetch details for one SO with master record
 */
function soGetSODetailsWithMaster(token, soID) {
  try {
    // Check view permission
    checkSalesPermission(token, "CanView");

    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    
    // Get master record from Sales Orders
    const soData = soGetRangeDataAsObjects('RANGESALESORDERS');
    const masterRecord = soData.find(so => so['Sales Order ID'] === soID);
    
    if (!masterRecord) {
      throw new Error('Sales Order not found: ' + soID);
    }
    
    // Format master date if needed
    if (masterRecord['Sales Order Date'] instanceof Date) {
      masterRecord['Sales Order Date'] = Utilities.formatDate(masterRecord['Sales Order Date'], tz, 'MM/dd/yyyy');
    }
    
    // Get details from Sales Details
    const rawDetails = soGetRangeDataAsObjects('RANGESALESDETAILS')
                .filter(r => r['Sales Order ID'] === soID);
    
    const details = rawDetails.map(row => {
      if (row['Sales Order Date'] instanceof Date) {
        row['Sales Order Date'] = Utilities.formatDate(row['Sales Order Date'], tz, 'MM/dd/yyyy');
      }
      // Create a new object without Customer ID
      const { 'Customer ID': removed, ...cleanRow } = row;
      return cleanRow;
    });
    
    return {
      master: masterRecord,
      details: details
    };
  
  } catch (error) {
    console.error('Error in soGetSODetailsWithMaster:', error);
    throw error;
  }
}

/**
 * Recalculate Total Sales Order Amount for a single Sales Order
 */
function soCalcTotalSOAmountForSingleSO(soID) {
  const ss = SpreadsheetApp.getActive();
  const soRange = ss.getRangeByName('RANGESALESORDERS');
  const soSheet = soRange.getSheet();
  const soStartRow = soRange.getRow();
  const soAllValues = soRange.getValues();
  const soHeaders = soAllValues[0];
  const soData = soAllValues.slice(1);

  const soIDCol = soHeaders.indexOf('Sales Order ID');
  const totalCol = soHeaders.indexOf('Total Sales Order Amount');
  
  // Find the row for this specific SO
  const soRowIndex = soData.findIndex(row => row[soIDCol] === soID);
  
  if (soRowIndex === -1) {
    console.log('Sales Order not found:', soID);
    return;
  }

  // Get current Sales Details data
  const sdData = soGetRangeDataAsObjects('RANGESALESDETAILS');

  // Calculate total for this specific sales order
  const sum = sdData
    .filter(d => d['Sales Order ID'] === soID)
    .reduce((acc, cur) => acc + (parseFloat(cur['Total Sales Price']) || 0), 0);
  
  // Update the Total Sales Order Amount
  if (!isNaN(sum)) {
    soSheet.getRange(soStartRow + 1 + soRowIndex, soRange.getColumn() + totalCol).setValue(sum);
    console.log('Updated Total Sales Order Amount for', soID, 'to', sum);
  }
}

/**
 * Update payment record when sales order is modified
 */
function soUpdatePaymentRecord(soID, updatedAmountReceived, updatedPaymentStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsRange = ss.getRangeByName('RANGEPAYMENTS');
    
    if (!paymentsRange) {
      console.error('RANGEPAYMENTS named range not found');
      return false;
    }
    
    const paymentsSheet = paymentsRange.getSheet();
    const startRow = paymentsRange.getRow();
    const allValues = paymentsRange.getValues();
    const headers = allValues[0];
    const data = allValues.slice(1);
    
    const referenceIDCol = headers.indexOf('Reference ID');
    const transactionTypeCol = headers.indexOf('Transaction Type');
    const amountCol = headers.indexOf('Amount');
    const statusCol = headers.indexOf('Status');
    
    // Find the payment record for this Sales Order (Sale transactions only)
    const rowIndex = data.findIndex(row => 
      row[referenceIDCol] === soID && 
      row[transactionTypeCol] === 'Sale'
    );
    
    if (rowIndex !== -1) {
      const sheetRow = startRow + 1 + rowIndex;
      
      // Update Amount and Status only
      if (amountCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
          .setValue(updatedAmountReceived);
      }
      
      if (statusCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
          .setValue(updatedPaymentStatus);
      }
      
      console.log('Updated payment record for SO:', soID, 'Amount:', updatedAmountReceived, 'Status:', updatedPaymentStatus);
      return true;
    } else {
      console.log('No payment record found for SO:', soID);
      return false;
    }
    
  } catch (error) {
    console.error('Error updating payment record:', error);
    return false;
  }
}

/**
 * Delete payment record when sales order is deleted
 */
function soDeletePaymentRecord(soID) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsRange = ss.getRangeByName('RANGEPAYMENTS');
    
    if (!paymentsRange) {
      console.error('RANGEPAYMENTS named range not found');
      return false;
    }
    
    const paymentsSheet = paymentsRange.getSheet();
    const startRow = paymentsRange.getRow();
    const allValues = paymentsRange.getValues();
    const headers = allValues[0];
    const data = allValues.slice(1);
    
    const referenceIDCol = headers.indexOf('Reference ID');
    
    // Find the payment record for this Sales Order
    const rowIndex = data.findIndex(row => row[referenceIDCol] === soID);
    
    if (rowIndex !== -1) {
      // Delete the row (add 2 because data starts at row 2 and array index starts at 0)
      paymentsSheet.deleteRow(startRow + 1 + rowIndex);
      console.log('Deleted payment record for SO:', soID);
      return true;
    } else {
      console.log('No payment record found for SO:', soID);
      return false;
    }
    
  } catch (error) {
    console.error('Error deleting payment record:', error);
    return false;
  }
}

/**
 * Sync all payment records with current sales order data
 */
function soSyncAllPaymentRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsRange = ss.getRangeByName('RANGEPAYMENTS');
    const soData = soGetRangeDataAsObjects('RANGESALESORDERS');
    
    if (!paymentsRange || !soData || soData.length === 0) {
      return;
    }
    
    const paymentsSheet = paymentsRange.getSheet();
    const paymentsData = paymentsRange.getValues();
    const headers = paymentsData[0];
    const existingPaymentData = paymentsData.slice(1);
    
    const referenceIDCol = headers.indexOf('Reference ID');
    const transactionTypeCol = headers.indexOf('Transaction Type');
    const amountCol = headers.indexOf('Amount');
    const statusCol = headers.indexOf('Status');
    
    // Track which sales orders we've processed
    const processedSOs = new Set();
    
    // STEP 1: Update existing sales payment records
    existingPaymentData.forEach((paymentRow, index) => {
      const referenceID = paymentRow[referenceIDCol];
      const transactionType = paymentRow[transactionTypeCol];
      
      // Only process Sale transactions
      if (referenceID && transactionType === 'Sale') {
        const salesOrder = soData.find(so => so['Sales Order ID'] === referenceID);
        
        if (salesOrder) {
          // Sales order exists - update the payment record
          const sheetRow = paymentsRange.getRow() + 1 + index;
          
          if (amountCol !== -1) {
            paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
              .setValue(parseFloat(salesOrder['Total Received']) || 0);
          }
          
          if (statusCol !== -1) {
            paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
              .setValue(salesOrder['Payment Status'] || 'Unpaid');
          }
          
          processedSOs.add(referenceID);
        }
      }
    });
    
    // STEP 2: Create missing payment records for sales orders that don't have them
    soData.forEach(salesOrder => {
      const soID = salesOrder['Sales Order ID'];
      
      if (!processedSOs.has(soID)) {
        // This sales order doesn't have a payment record - create one
        const paymentData = {
          date: salesOrder['Sales Order Date'],
          paymentID: soGeneratePaymentID(),
          referenceID: soID,
          partyID: salesOrder['Customer ID'],
          paymentMode: salesOrder['Payment Mode'] || 'Pending',
          amount: parseFloat(salesOrder['Total Received']) || 0,
          status: salesOrder['Payment Status'] || 'Unpaid'
        };
        
        // FIXED: Use the corrected append-only function
        soCreatePaymentRecord(paymentData);
        console.log('Created missing payment record for SO:', soID);
      }
    });
    
    console.log('Completed payment records sync for sales orders only');
    
  } catch (error) {
    console.error('Error syncing payment records:', error);
  }
}
