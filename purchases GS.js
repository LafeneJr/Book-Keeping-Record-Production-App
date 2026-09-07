// ===========================================
// PURCHASE PERMISSION HELPER FUNCTION
// ===========================================

/**
 * Check purchase permissions for a user
 * @param {string} token - Session token
 * @param {string} action - Action to check ('CanView', 'CanAdd', 'CanEdit', 'CanDelete')
 * @returns {Object} Session data if permission is granted
 * @throws {Error} If permission is denied or session is invalid
 */
function checkPurchasePermission(token, action) {
  try {
    // Validate session
    const session = Authentication.validateSession(token);
    if (!session) {
      throw new Error("Invalid session. Please log in again.");
    }
    
    // Check if user has purchase permissions
    const purchasePerms = session.permissions && session.permissions["purchases"];
    if (!purchasePerms) {
      throw new Error("You don't have permission to access purchase orders.");
    }
    
    // Check specific action permission
    const hasPermission = purchasePerms[action];
    
    if (!hasPermission) {
      const actionText = action.replace('Can', '').toLowerCase();
      throw new Error(`You don't have permission to ${actionText} purchase orders.`);
    }
    
    return session;
  } catch (error) {
    console.error('Purchase permission check error:', error);
    throw error;
  }
}


// Get all suppliers
function poGetSuppliers(token) {
  try {
    // Check view permission
    checkPurchasePermission(token, "CanView");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Suppliers");
    const range = ss.getRangeByName("RANGESUPPLIERS");

    if (!range) {
      console.log("poGetSuppliers: RANGESUPPLIERS named range not found.");
      return [];
    }

    try {
      const data = range.getValues();
      if (!data || data.length === 0) {
        console.log("poGetSuppliers: No data found in RANGESUPPLIERS.");
        return [];
      }
      const headers = data[0];
      const suppliers = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(cell => cell === '' || cell === null || typeof cell === 'undefined')) continue;

        suppliers.push({
          id: row[headers.indexOf("Supplier ID")],
          name: row[headers.indexOf("Supplier Name")],
          state: row[headers.indexOf("State")],
          city: row[headers.indexOf("City")]
        });
      }
      return suppliers;
    } catch (e) {
      console.error("Error in poGetSuppliers: " + e.message + " Stack: " + e.stack);
      return [];
    }

  } catch (error) {
    console.error('Error in poGetSuppliers:', error);
    throw error;
  }
}

// Get all inventory items
function poGetInventoryItems(token) {
  try {
    // Check view permission
    checkPurchasePermission(token, "CanView");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Inventory Items");
    const range = ss.getRangeByName("RANGEINVENTORYITEMS");

    if (!range) {
      console.log("poGetInventoryItems: RANGEINVENTORYITEMS named range not found.");
      return [];
    }

    try {
      const data = range.getValues();
      if (!data || data.length === 0) {
        console.log("poGetInventoryItems: No data found in RANGEINVENTORYITEMS.");
        return [];
      }
      const headers = data[0];
      const items = [];

      // NEW: Get the Product Quantity column index
      const productQtyIndex = headers.indexOf("Product Quantity");

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(cell => cell === '' || cell === null || typeof cell === 'undefined')) continue;

        // NEW: Check if this is a finished good (has Product Quantity)
        const productQty = row[productQtyIndex];
        const hasProductQty = productQty && productQty !== '' && productQty !== 0 && productQty !== '0';
        
        // ONLY include items that are NOT finished goods (no Product Quantity)
        if (!hasProductQty) {
          items.push({
            id: row[headers.indexOf("Item ID")],
            name: row[headers.indexOf("Item Name")],
            type: row[headers.indexOf("Item Type")],
            category: row[headers.indexOf("Item Category")]
          });
        } else {
          console.log(`Filtered out finished good: ${row[headers.indexOf("Item Name")]} (Product Qty: ${productQty})`);
        }
      }
      
      console.log(`poGetInventoryItems: Returning ${items.length} raw materials (filtered out finished goods)`);
      return items;
    } catch (e) {
      console.error("Error in poGetInventoryItems: " + e.message + " Stack: " + e.stack);
      return [];
    }

  } catch (error) {
    console.error('Error in poGetInventoryItems:', error);
    throw error;
  }
}

// Get Payment statuses
function poGetPMTStatuses(token) {
  try {
    // Check view permission
    checkPurchasePermission(token, "CanView");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");

    if (!range) {
      console.log("poGetPMTStatuses: RANGEDIMENSIONS named range not found.");
      return [];
    }

    try {
      const data = range.getValues();
      if (!data || data.length === 0) {
        console.log("poGetPMTStatuses: No data found in RANGEDIMENSIONS.");
        return [];
      }
      const headers = data[0];
      const statusIndex = headers.indexOf("Payment Status");

      if (statusIndex === -1) {
        console.log("poGetPMTStatuses: 'Payment Status' header not found.");
        return [];
      }

      const statuses = [];
      const statusSet = new Set();

      for (let i = 1; i < data.length; i++) {
        const status = data[i][statusIndex];
        if (status && status.trim() !== '' && !statusSet.has(status)) {
          statuses.push(status);
          statusSet.add(status);
        }
      }
      return statuses;
    } catch (e) {
      console.error("Error in poGetPMTStatuses: " + e.message + " Stack: " + e.stack);
      return [];
    }

  } catch (error) {
    console.error('Error in poGetPMTStatuses:', error);
    throw error;
  }
}


// Add new Payment status
function poAddNewPMTStatus(token,status) {
  try {
   // Check add permission
    checkPurchasePermission(token, "CanAdd"); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dimensions");
    const range = ss.getRangeByName("RANGEDIMENSIONS");

    if (!range) {
      console.error("poAddNewPMTStatus: RANGEDIMENSIONS named range not found.");
      throw new Error("Dimensions range not found.");
    }

    try {
      const data = range.getValues();
      const headers = data[0];
      const statusIndex = headers.indexOf("Payment Status");

      if (statusIndex === -1) {
        console.error("poAddNewPMTStatus: 'Payment Status' header not found.");
        throw new Error("'Payment Status' column not found in Dimensions sheet.");
      }

      let lastRow = sheet.getLastRow();
      // Ensure we append to the first truly empty row after headers
      if (lastRow < 1) lastRow = 1; // If sheet is empty, start at row 1
      let targetRow = lastRow + 1;

      sheet.getRange(targetRow, statusIndex + 1).setValue(status);
      console.log(`Added new Payment status: ${status} at row ${targetRow}`);
    } catch (e) {
      console.error("Error in poAddNewPMTStatus: " + e.message + " Stack: " + e.stack);
      throw new Error("Failed to add new Payment status: " + e.message);
    }
  
  } catch (error) {
    console.error('Error in poAddNewPMTStatus:', error);
    throw error;
  }
}


// Generate Purchase Order ID
function poGeneratePOID(token) {
  try {
    // Check add permission
    checkPurchasePermission(token, "CanAdd");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Purchase Details");
    const range = ss.getRangeByName("RANGEPURCHASEDETAILS");

    if (!range) {
      console.log("poGeneratePOID: RANGEPURCHASEDETAILS named range not found. Generating simple ID.");
      return "Purchase Order" + Math.floor(10000 + Math.random() * 90000);
    }

    try {
      const data = range.getValues();
      const headers = data[0];
      const idIndex = headers.indexOf("Purchase Order ID");
      const existingIds = new Set();

      if (idIndex === -1) {
        console.warn("poGeneratePOID: 'Purchase Order ID' header not found in PurchaseDetails. Generating simple ID.");
        return "Purchase Order" + Math.floor(10000 + Math.random() * 90000);
      }

      for (let i = 1; i < data.length; i++) {
        const id = data[i][idIndex];
        if (id && id !== '') {
          existingIds.add(String(id)); // Ensure string for consistency
        }
      }

      let newId;
      do {
        newId = "PO" + Math.floor(10000 + Math.random() * 90000);
      } while (existingIds.has(newId));

      console.log("Generated new Purchase Order ID: " + newId);
      return newId;
    } catch (e) {
      console.error("Error in poGeneratePOID: " + e.message + " Stack: " + e.stack);
      throw new Error("Failed to generate Purchase Order ID: " + e.message);
    }
  
  } catch (error) {
    console.error('Error in poGeneratePOID:', error);
    throw error;
  }
}

// Get all purchase orders
function poGetPOs(token) {
  try {
    // Check view permission
    checkPurchasePermission(token, "CanView");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Purchase Orders");
    const range = ss.getRangeByName("RANGEPURCHASEORDERS");

    if (!range) {
      console.log("poGetPOs: RANGEPURCHASEORDERS named range not found. Returning empty array.");
      return [];
    }

    try {
      const data = range.getValues();
      console.log(`poGetPOs: Raw data has ${data.length} rows from range.`);

      if (!data || data.length <= 1) {
        console.log("poGetPOs: No data found (only headers or empty).");
        return [];
      }

      const headers = data[0];
      const poIdIndex = headers.indexOf("Purchase Order ID");

      if (poIdIndex === -1) {
        console.error("poGetPOs: 'Purchase Order ID' header not found.");
        return [];
      }

      const pos = [];
      const getVal = (h, row) => {
        const index = headers.indexOf(h);
        return index !== -1 ? row[index] : '';
      };

      // NEW: Read data from bottom to top (newest records are at the bottom of the sheet)
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const poId = row[poIdIndex];
        
        // Skip empty rows
        if (!poId || poId === '' || poId === null || 
            row.every(cell => cell === '' || cell === null)) {
          continue;
        }

        let dateValue = getVal("Date", row);
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

        pos.push({
          date: dateValue,
          id: poId,
          supplierId: getVal("Supplier ID", row),
          supplierName: getVal("Supplier Name", row),
          state: getVal("State", row),
          city: getVal("City", row),
          totalAmount: getVal("Total Amount", row),
          totalPaid: getVal("Total Paid", row),
          poBalance: getVal("Purchase Order Balance", row),
          pmtStatus: getVal("Payment Status", row)
        });
      }
      
      console.log(`poGetPOs: Returning ${pos.length} purchase orders (read from bottom to top, newest first).`);
      return pos;
    } catch (e) {
      console.error("Error in poGetPOs: " + e.message + " Stack: " + e.stack);
      return [];
    }
  
  } catch (error) {
    console.error('Error in poGetPOs:', error);
    throw error;
  }
}

// Get Purchase Order details - UPDATED FOR SIMPLIFIED STRUCTURE
function poGetPODetails(poId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Purchase Details");
  const range = ss.getRangeByName("RANGEPURCHASEDETAILS");

  if (!range) {
    console.log("poGetPODetails: RANGEPURCHASEDETAILS named range not found. Returning empty array.");
    return [];
  }

  try {
    const data = range.getValues();
    if (!data || data.length === 0) {
      console.log("poGetPODetails: No data found in RANGEPURCHASEDETAILS.");
      return [];
    }

    const headers = data[0];
    const details = [];

    // Helper to get value by header name
    const getVal = (h, row) => {
      const index = headers.indexOf(h);
      return index !== -1 ? row[index] : null;
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip empty rows or those that don't match the given Purchase Order ID
      if (row.every(cell => cell === '' || cell === null || typeof cell === 'undefined') || 
          getVal("Purchase Order ID", row) !== poId) {
        continue;
      }

      // UPDATED: Format the date as MM/dd/yyyy
      let dateVal = getVal("Date", row);
      if (dateVal instanceof Date) {
        const month = (dateVal.getMonth() + 1).toString().padStart(2, '0');
        const day = dateVal.getDate().toString().padStart(2, '0');
        const year = dateVal.getFullYear();
        dateVal = `${month}/${day}/${year}`;
      } else if (typeof dateVal === "string" || typeof dateVal === "number") {
        const parsedDate = new Date(dateVal);
        if (!isNaN(parsedDate)) {
          const month = (parsedDate.getMonth() + 1).toString().padStart(2, '0');
          const day = parsedDate.getDate().toString().padStart(2, '0');
          const year = parsedDate.getFullYear();
          dateVal = `${month}/${day}/${year}`;
        }
      }

      details.push({
        date: dateVal,
        poId: getVal("Purchase Order ID", row),
        detailId: getVal("Purchase Detail ID", row),
        supplierName: getVal("Supplier Name", row),
        state: getVal("State", row),
        city: getVal("City", row),
        itemId: getVal("Item ID", row),
        itemType: getVal("Item Type", row),
        itemCategory: getVal("Item Category", row),
        itemName: getVal("Item Name", row),
        qtyPurchased: getVal("Quantity Purchased", row),
        unitCost: getVal("Unit Cost", row),
        shippingFees: getVal("Shipping Fee", row),
        totalPrice: getVal("Total Price", row)
      });
    }

    console.log(`poGetPODetails: Retrieved ${details.length} details for Purchase Order ID: ${poId}`);
    return details;
  } catch (e) {
    console.error("Error in poGetPODetails: " + e.message + " Stack: " + e.stack);
    return [];
  }
}


// Save new Purchase Order
function poSaveNewPO(token,items, amountPaid, paymentTransaction, supplierId, amountPaid2) {
  try {
    // Check add permission
    checkPurchasePermission(token, "CanAdd");
    poLogStart('Save New Purchase Order', {
      itemCount: items.length,
      poId: items[0]?.poId,
      supplierId: supplierId
    });

    if (!Array.isArray(items) || items.length === 0) {
      // ADD LOGGING FOR VALIDATION FAILURE:
      poLogError('Save New Purchase Order', new Error("No items provided"), {
        itemCount: items ? items.length : 0
      });
      throw new Error("Cannot save purchase order: No items provided");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName("Purchase Details");
    const ordersSheet = ss.getSheetByName("Purchase Orders");

    // Save to PurchaseDetails
    const detailsRange = ss.getRangeByName("RANGEPURCHASEDETAILS");
    if (!detailsRange) {
      console.error("poSaveNewPO: RANGEPURCHASEDETAILS named range not found.");
      throw new Error("Purchase Details range not found.");
    }

    try {
      const detailsData = detailsRange.getValues();
      const detailsHeaders = detailsData[0];

      // Prepare new rows for details
      const newRows = items.map(item => {
        const newRow = [];
        detailsHeaders.forEach(header => {
          switch(header) {
            case "Date": newRow.push(item.date); break;
            case "Purchase Order ID": newRow.push(item.poId); break;
            case "Purchase Detail ID": newRow.push(item.detailId); break;
            case "Supplier Name": newRow.push(item.supplierName); break;
            case "State": newRow.push(item.state); break;
            case "City": newRow.push(item.city); break;
            case "Item ID": newRow.push(item.itemId); break;
            case "Item Type": newRow.push(item.itemType); break;
            case "Item Category": newRow.push(item.itemCategory); break;
            case "Item Name": newRow.push(item.itemName); break;
            case "Quantity Purchased": newRow.push(item.qtyPurchased); break;          
            case "Unit Cost": newRow.push(item.unitCost); break;
            case "Shipping Fee": newRow.push(item.shippingFees); break;
            case "Total Price": newRow.push(item.totalPrice); break;
            default: newRow.push(''); break;
          }
        });
        return newRow;
      });

      // Append to details sheet
      detailsSheet.getRange(detailsSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      console.log(`poSaveNewPO: Successfully saved ${newRows.length} item details.`);

      // Calculate total amount for payment status
      let totalAmount = 0;
      items.forEach(item => {
        totalAmount += Number(item.totalPrice) || 0;
      });

      // Calculate payment status
      const totalPaid = Number(amountPaid) || 0;
      let paymentStatus = poCalculatePaymentStatus(totalPaid, totalAmount);

      // Create master record in Purchase Orders
      const firstItem = items[0];
      ordersSheet.appendRow([
        firstItem.date,
        firstItem.poId,
        supplierId,
        firstItem.supplierName,
        firstItem.state,
        firstItem.city,
        totalAmount,
        totalPaid,
        totalAmount - totalPaid,
        paymentStatus
      ]);

      // Update related data
      revisetotalinventory();
      poUpdateRemainingQty();
      poUpdateReorderRequired();
      poUpdateTotalPurchases();
      
      // Update supplier payments and balance (only if amount paid > 0)
      if (amountPaid > 0) {
        poUpdateSupplierPayments(supplierId, amountPaid);
      }
      poUpdateBalancePayable();

      // ALWAYS record payment in Payments sheet (even for $0 payments)
      const paymentData = {
        date: items[0].date,
        paymentId: poGeneratePaymentID(),
        transactionType: "Purchase",
        referenceId: items[0].poId,
        partyId: supplierId,
        paymentMode: '',
        amount: amountPaid,
        status: paymentStatus
      };
      poRecordPayment(paymentData);
      
      poLogSuccess('Save New Purchase Order', {
        poId: items[0].poId,
        itemsSaved: items.length,
        totalAmount: totalAmount
      });
    } catch (e) {
      // ADD LOGGING IN CATCH BLOCK:
      poLogError('Save New Purchase Order', e, {
        poId: items[0]?.poId,
        itemCount: items.length
      });
      console.error("Error in poSaveNewPO: " + e.message + " Stack: " + e.stack);
      throw new Error("Failed to save new Purchase Order: " + e.message);
    }
  
  } catch (error) {
    console.error('Error in poSaveNewPO:', error);
    throw error;
  }
}

// Update total Purchase Order
function poUpdateTotalPO(poId, supplierId, amountPaid = 0) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailsSheet = ss.getSheetByName("Purchase Details");
  const ordersSheet  = ss.getSheetByName("Purchase Orders");
  const detailsRange = ss.getRangeByName("RANGEPURCHASEDETAILS");
  const ordersRange  = ss.getRangeByName("RANGEPURCHASEORDERS");

  if (!detailsRange || !ordersRange) {
    console.warn("poUpdateTotalPO: Named ranges RANGEPURCHASEDETAILS or RANGEPURCHASEORDERS not found. Skipping.");
    return;
  }

  // 1) Gather PO Details and compute totalAmount, keep firstRow for meta
  const detailsData    = detailsRange.getValues();
  const detailsHeaders = detailsData[0];
  const poIdIndex      = detailsHeaders.indexOf("Purchase Order ID");
  const totalPriceIdx  = detailsHeaders.indexOf("Total Price");

  let totalAmount = 0;
  let firstRow    = null;

  for (let i = 1; i < detailsData.length; i++) {
    if (detailsData[i][poIdIndex] === poId) {
      totalAmount += Number(detailsData[i][totalPriceIdx]) || 0;
      if (!firstRow) {
        firstRow = detailsData[i];
      }
    }
  }

  if (!firstRow) {
    console.warn(`poUpdateTotalPO: No details found for Purchase Order ID ${poId}. Nothing to update.`);
    return;
  }

  // 2) Calculate payment status based on amountPaid
  const totalPaid = Number(amountPaid) || 0;
  const poBalance = totalAmount - totalPaid;
  
  let paymentStatus = poCalculatePaymentStatus(totalPaid, totalAmount);

  console.log(`Payment Calculation: Total=${totalAmount}, Paid=${totalPaid}, Balance=${poBalance}, Status=${paymentStatus}`);

  // 3) Get orders data and headers
  const ordersData    = ordersRange.getValues();
  const ordersHeaders = ordersData[0];

  // 4) Find existing Purchase Order row in ordersData
  const ordersPoIdCol  = ordersHeaders.indexOf("Purchase Order ID");
  const existing = ordersData.findIndex((r, i) => i > 0 && r[ordersPoIdCol] == poId);

  if (existing > 0) {
    const startRow = ordersRange.getRow();
    const startCol = ordersRange.getColumn();
    const targetRow = startRow + existing;

    const totalAmountIndex = ordersHeaders.indexOf("Total Amount");
    const totalPaidIndex = ordersHeaders.indexOf("Total Paid");
    const poBalanceIndex = ordersHeaders.indexOf("Purchase Order Balance");
    const pmtStatusIndex = ordersHeaders.indexOf("Payment Status");
    
    // Update individual cells with new payment logic
    if (totalAmountIndex !== -1) {
      ordersSheet.getRange(targetRow, startCol + totalAmountIndex).setValue(totalAmount);
    }
    if (totalPaidIndex !== -1) {
      ordersSheet.getRange(targetRow, startCol + totalPaidIndex).setValue(totalPaid);
    }
    if (poBalanceIndex !== -1) {
      ordersSheet.getRange(targetRow, startCol + poBalanceIndex).setValue(poBalance);
    }
    if (pmtStatusIndex !== -1) {
      ordersSheet.getRange(targetRow, startCol + pmtStatusIndex).setValue(paymentStatus);
    }
    console.log(`poUpdateTotalPO: Updated PO ${poId} with payment data.`);
  } else {
    // Create new PO row with payment data
    const newRow = ordersHeaders.map(header => {
      switch (header) {
        case "Date":               return firstRow[ detailsHeaders.indexOf("Date") ];
        case "Purchase Order ID":  return poId;
        case "Supplier ID":        return supplierId;
        case "Supplier Name":      return firstRow[ detailsHeaders.indexOf("Supplier Name") ];
        case "State":              return firstRow[ detailsHeaders.indexOf("State") ];
        case "City":               return firstRow[ detailsHeaders.indexOf("City") ];
        case "Total Amount":       return totalAmount;
        case "Total Paid":         return totalPaid;
        case "Purchase Order Balance": return poBalance;
        case "Payment Status":     return paymentStatus;
        default:                   return "";
      }
    });
    
    ordersSheet.appendRow(newRow);
    console.log(`poUpdateTotalPO: Created new PO ${poId} with payment data.`);
  }
}

// NEW: Update supplier payments when PO payments are made
function poUpdateSupplierPayments(supplierId, amountChange) {
  console.log("=== STARTING poUpdateSupplierPayments ===");
  console.log(`Supplier: ${supplierId}, Amount Change: ${amountChange}`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suppliersSheet = ss.getSheetByName("Suppliers");
  const suppliersRange = ss.getRangeByName("RANGESUPPLIERS");

  if (!suppliersRange) {
    console.warn("poUpdateSupplierPayments: RANGESUPPLIERS named range not found. Skipping update.");
    return;
  }

  try {
    const suppliersData = suppliersRange.getValues();
    const suppliersHeaders = suppliersData[0];

    const supplierIdIndex = suppliersHeaders.indexOf("Supplier ID");
    const totalPaymentsIndex = suppliersHeaders.indexOf("Total Payments");

    console.log(`Indexes - SupplierID: ${supplierIdIndex}, TotalPayments: ${totalPaymentsIndex}`);

    if (supplierIdIndex === -1 || totalPaymentsIndex === -1) {
      console.error("poUpdateSupplierPayments: Required columns not found");
      return;
    }

    const startRow = suppliersRange.getRow();
    const startCol = suppliersRange.getColumn();
    
    let updated = false;
    for (let i = 1; i < suppliersData.length; i++) {
      // Skip empty rows
      if (suppliersData[i].every(cell => cell === '' || cell === null)) continue;
      
      if (suppliersData[i][supplierIdIndex] === supplierId) {
        const currentTotalPayments = Number(suppliersData[i][totalPaymentsIndex]) || 0;
        const newTotalPayments = currentTotalPayments + (Number(amountChange) || 0);
        
        // Calculate exact cell position
        const targetRow = startRow + i;
        const targetCol = startCol + totalPaymentsIndex;
        
        suppliersSheet.getRange(targetRow, targetCol).setValue(newTotalPayments);
        console.log(`Updated Supplier ${supplierId}: Payments from ${currentTotalPayments} to ${newTotalPayments}`);
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      console.warn(`Supplier ${supplierId} not found in Suppliers sheet`);
    }
    
    console.log("poUpdateSupplierPayments: Completed successfully.");
  } catch (e) {
    console.error("Error in poUpdateSupplierPayments: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update supplier payments: " + e.message);
  }
}


// Update remaining quantity
function poUpdateRemainingQty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName("Inventory Items");
  const itemsRange = ss.getRangeByName("RANGEINVENTORYITEMS");

  if (!itemsRange) {
    console.warn("poUpdateRemainingQty: RANGEINVENTORYITEMS named range not found. Skipping update.");
    return;
  }

  try {
    const itemsData = itemsRange.getValues();
    const itemsHeaders = itemsData[0];

    const purchasedIndex = itemsHeaders.indexOf("Quantity Purchased");
    const usedIndex = itemsHeaders.indexOf("Quantity Used");
    const soldIndex = itemsHeaders.indexOf("Quantity Sold");
    const productQtyIndex = itemsHeaders.indexOf("Product Quantity");
    const remainingIndex = itemsHeaders.indexOf("Remaining Quantity");
    const itemTypeIndex = itemsHeaders.indexOf("Item Type");

    // Validate that we found the required columns
    if (purchasedIndex === -1 || usedIndex === -1 || soldIndex === -1 || remainingIndex === -1 || itemTypeIndex === -1) {
      console.error("poUpdateRemainingQty: Required columns not found in Inventory Items");
      return;
    }

    const startRow = itemsRange.getRow();
    const startCol = itemsRange.getColumn();
    
    for (let i = 1; i < itemsData.length; i++) {
      // Skip empty rows
      if (itemsData[i].every(cell => cell === '' || cell === null)) continue;
      
      const itemType = itemsData[i][itemTypeIndex];
      const purchased = Number(itemsData[i][purchasedIndex]) || 0;
      const used = Number(itemsData[i][usedIndex]) || 0;
      const sold = Number(itemsData[i][soldIndex]) || 0;
      const productQty = Number(itemsData[i][productQtyIndex]) || 0;
      
      let remaining = 0;
      
      // FIXED: Use correct formula based on item type
      if (itemType === "Raw Material" || purchased > 0) {
        // Raw Materials: Purchased - Used - Sold
        remaining = purchased - used - sold;
      } else if (itemType === "Finished Goods" || productQty > 0) {
        // Finished Goods: Product Quantity - Sold
        remaining = productQty - sold;
      } else {
        // Default fallback
        remaining = purchased - used - sold;
      }
      
      // Ensure remaining doesn't go negative
      remaining = Math.max(0, remaining);

      // Calculate exact cell position
      const targetRow = startRow + i;
      const targetCol = startCol + remainingIndex;
      
      itemsSheet.getRange(targetRow, targetCol).setValue(remaining);
      console.log(`Updated row ${i}: Type=${itemType}, Purchased=${purchased}, Used=${used}, Sold=${sold}, Remaining=${remaining}`);
    }
    console.log("poUpdateRemainingQty: Remaining quantities updated successfully with FIXED formulas.");
  } catch (e) {
    console.error("Error in poUpdateRemainingQty: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update remaining quantity: " + e.message);
  }
}

// Update reorder required
function poUpdateReorderRequired() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = ss.getSheetByName("Inventory Items");
  const itemsRange = ss.getRangeByName("RANGEINVENTORYITEMS");

  if (!itemsRange) {
    console.warn("poUpdateReorderRequired: RANGEINVENTORYITEMS named range not found. Skipping update.");
    return;
  }

  try {
    const itemsData = itemsRange.getValues();
    const itemsHeaders = itemsData[0];

    const remainingIndex = itemsHeaders.indexOf("Remaining Quantity");
    const reorderLevelIndex = itemsHeaders.indexOf("Reorder Level");
    const reorderRequiredIndex = itemsHeaders.indexOf("Reorder Required");

    // Validate that we found the required columns
    if (remainingIndex === -1 || reorderLevelIndex === -1 || reorderRequiredIndex === -1) {
      console.error("poUpdateReorderRequired: Required columns not found in Inventory Items");
      return;
    }

    const startRow = itemsRange.getRow();
    const startCol = itemsRange.getColumn();
    
    for (let i = 1; i < itemsData.length; i++) {
      // Skip empty rows
      if (itemsData[i].every(cell => cell === '' || cell === null)) continue;
      
      const remaining = Number(itemsData[i][remainingIndex]) || 0;
      const reorderLevel = Number(itemsData[i][reorderLevelIndex]) || 0;
      const reorderRequired = remaining < reorderLevel ? "Yes" : "No";

      // Calculate exact cell position
      const targetRow = startRow + i;
      const targetCol = startCol + reorderRequiredIndex;
      
      itemsSheet.getRange(targetRow, targetCol).setValue(reorderRequired);
    }
    console.log("poUpdateReorderRequired: Reorder required status updated successfully.");
  } catch (e) {
    console.error("Error in poUpdateReorderRequired: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update reorder required status: " + e.message);
  }
}

// Update total purchases for suppliers
function poUpdateTotalPurchases() {
  console.log("=== STARTING poUpdateTotalPurchases ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suppliersSheet = ss.getSheetByName("Suppliers");
  const detailsSheet = ss.getSheetByName("Purchase Details");

  const suppliersRange = ss.getRangeByName("RANGESUPPLIERS");
  const detailsRange = ss.getRangeByName("RANGEPURCHASEDETAILS");

  if (!suppliersRange || !detailsRange) {
    console.warn("poUpdateTotalPurchases: One or more named ranges not found. Skipping update.");
    return;
  }

  try {
    const suppliersData = suppliersRange.getValues();
    const suppliersHeaders = suppliersData[0];
    const detailsData = detailsRange.getValues();
    const detailsHeaders = detailsData[0];

    console.log("Suppliers Headers:", suppliersHeaders);
    console.log("Purchase Details Headers:", detailsHeaders);

    const supplierIdIndex = suppliersHeaders.indexOf("Supplier ID");
    const totalPurchasesIndex = suppliersHeaders.indexOf("Total Purchases");
    const detailsSupplierNameIndex = detailsHeaders.indexOf("Supplier Name");
    const detailsTotalPriceIndex = detailsHeaders.indexOf("Total Price");

    console.log(`Indexes - SupplierID: ${supplierIdIndex}, TotalPurchases: ${totalPurchasesIndex}, SupplierName: ${detailsSupplierNameIndex}, TotalPrice: ${detailsTotalPriceIndex}`);

    // Validate required columns
    if (supplierIdIndex === -1 || totalPurchasesIndex === -1 || 
        detailsSupplierNameIndex === -1 || detailsTotalPriceIndex === -1) {
      console.error("poUpdateTotalPurchases: Required columns not found");
      console.error(`Missing: ${supplierIdIndex === -1 ? 'Supplier ID, ' : ''}${totalPurchasesIndex === -1 ? 'Total Purchases, ' : ''}${detailsSupplierNameIndex === -1 ? 'Supplier Name, ' : ''}${detailsTotalPriceIndex === -1 ? 'Total Price' : ''}`);
      return;
    }

    // Create a map of supplier totals by Supplier Name (since we don't have Supplier ID in Purchase Details)
    const supplierTotals = {};

    for (let i = 1; i < detailsData.length; i++) {
      const supplierName = detailsData[i][detailsSupplierNameIndex];
      const totalPrice = Number(detailsData[i][detailsTotalPriceIndex]) || 0;

      if (supplierName && supplierName !== '') {
        if (!supplierTotals[supplierName]) supplierTotals[supplierName] = 0;
        supplierTotals[supplierName] += totalPrice;
        console.log(`Added ${totalPrice} to supplier: ${supplierName}`);
      }
    }

    console.log("Supplier Totals Map:", supplierTotals);

    // Update suppliers - target only the Total Purchases column
    const startRow = suppliersRange.getRow();
    const startCol = suppliersRange.getColumn();
    
    let updateCount = 0;
    for (let i = 1; i < suppliersData.length; i++) {
      // Skip empty rows
      if (suppliersData[i].every(cell => cell === '' || cell === null)) continue;
      
      const supplierName = suppliersData[i][suppliersHeaders.indexOf("Supplier Name")];
      if (supplierName && supplierTotals[supplierName]) {
        // Calculate exact cell position for Total Purchases
        const targetRow = startRow + i;
        const targetCol = startCol + totalPurchasesIndex;
        
        suppliersSheet.getRange(targetRow, targetCol).setValue(supplierTotals[supplierName]);
        console.log(`Updated ${supplierName}: ${supplierTotals[supplierName]} at row ${targetRow}, col ${targetCol}`);
        updateCount++;
      }
    }
    
    console.log(`poUpdateTotalPurchases: Updated ${updateCount} suppliers successfully.`);
  } catch (e) {
    console.error("Error in poUpdateTotalPurchases: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update total purchases: " + e.message);
  }
}

// Update balance payable for suppliers
// Update balance payable for suppliers - DEBUGGED VERSION
function poUpdateBalancePayable() {
  console.log("=== STARTING poUpdateBalancePayable ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suppliersSheet = ss.getSheetByName("Suppliers");
  const suppliersRange = ss.getRangeByName("RANGESUPPLIERS");

  if (!suppliersRange) {
    console.warn("poUpdateBalancePayable: RANGESUPPLIERS named range not found. Skipping update.");
    return;
  }

  try {
    const suppliersData = suppliersRange.getValues();
    const suppliersHeaders = suppliersData[0];

    console.log("Suppliers Headers:", suppliersHeaders);

    const totalPurchasesIndex = suppliersHeaders.indexOf("Total Purchases");
    const totalPaymentsIndex = suppliersHeaders.indexOf("Total Payments");
    const balancePayableIndex = suppliersHeaders.indexOf("Balance Payable");

    console.log(`Indexes - TotalPurchases: ${totalPurchasesIndex}, TotalPayments: ${totalPaymentsIndex}, BalancePayable: ${balancePayableIndex}`);

    // Validate that we found the required columns
    if (totalPurchasesIndex === -1 || totalPaymentsIndex === -1 || balancePayableIndex === -1) {
      console.error("poUpdateBalancePayable: Required columns not found in Suppliers");
      console.error(`Missing: ${totalPurchasesIndex === -1 ? 'Total Purchases, ' : ''}${totalPaymentsIndex === -1 ? 'Total Payments, ' : ''}${balancePayableIndex === -1 ? 'Balance Payable' : ''}`);
      return;
    }

    const startRow = suppliersRange.getRow();
    const startCol = suppliersRange.getColumn();
    
    let updateCount = 0;
    for (let i = 1; i < suppliersData.length; i++) {
      // Skip empty rows
      if (suppliersData[i].every(cell => cell === '' || cell === null)) continue;
      
      const totalPurchases = Number(suppliersData[i][totalPurchasesIndex]) || 0;
      const totalPayments = Number(suppliersData[i][totalPaymentsIndex]) || 0;
      const balance = totalPurchases - totalPayments;

      // Calculate exact cell position
      const targetRow = startRow + i;
      const targetCol = startCol + balancePayableIndex;
      
      suppliersSheet.getRange(targetRow, targetCol).setValue(balance);
      console.log(`Updated row ${i}: Purchases=${totalPurchases}, Payments=${totalPayments}, Balance=${balance} at row ${targetRow}, col ${targetCol}`);
      updateCount++;
    }
    
    console.log(`poUpdateBalancePayable: Updated ${updateCount} supplier balances successfully.`);
  } catch (e) {
    console.error("Error in poUpdateBalancePayable: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update balance payable: " + e.message);
  }
}

// Update PO balance
function poUpdatePOBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Purchase Orders");
  const ordersRange = ss.getRangeByName("RANGEPURCHASEORDERS");

  if (!ordersRange) {
    console.warn("poUpdatePOBalance: RANGEPURCHASEORDERS named range not found. Skipping update.");
    return;
  }

  try {
    const ordersData = ordersRange.getValues();
    const ordersHeaders = ordersData[0];

    const totalAmountIndex = ordersHeaders.indexOf("Total Amount"); // Your actual column name
    const totalPaidIndex = ordersHeaders.indexOf("Total Paid"); // Your actual column name
    const poBalanceIndex = ordersHeaders.indexOf("Purchase Order Balance"); // Your actual column name

    // Validate required columns
    if (totalAmountIndex === -1 || totalPaidIndex === -1 || poBalanceIndex === -1) {
      console.error("poUpdatePOBalance: Required columns not found in Purchase Orders");
      return;
    }

    const startRow = ordersRange.getRow();
    const startCol = ordersRange.getColumn();
    
    for (let i = 1; i < ordersData.length; i++) {
      // Skip empty rows
      if (ordersData[i].every(cell => cell === '' || cell === null)) continue;
      
      const totalAmount = Number(ordersData[i][totalAmountIndex]) || 0;
      const totalPaid = Number(ordersData[i][totalPaidIndex]) || 0;
      const balance = totalAmount - totalPaid;

      // Calculate exact cell position
      const targetRow = startRow + i;
      const targetCol = startCol + poBalanceIndex;
      
      ordersSheet.getRange(targetRow, targetCol).setValue(balance);
    }
    console.log("poUpdatePOBalance: Purchase Order balances updated successfully.");
  } catch (e) {
    console.error("Error in poUpdatePOBalance: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to update Purchase Order balance: " + e.message);
  }
}

function poDeletePODetail(token, detailId, poId) {
  try {
    // Check delete permission
    checkPurchasePermission(token, "CanDelete");

    poLogStart('Delete PO Detail', {
      detailId: detailId,
      poId: poId
    });

    console.log(`=== STARTING poDeletePODetail ===`);
    console.log(`Received parameters - detailId: "${detailId}", poId: "${poId}"`);
    
    // Validate parameters
    if (!detailId || detailId === 'undefined' || !poId || poId === 'undefined') {
      console.error(`Invalid parameters: detailId="${detailId}", poId="${poId}"`);
      throw new Error(`Invalid parameters received: detailId="${detailId}", poId="${poId}"`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailsSheet = ss.getSheetByName("Purchase Details");
    const ordersSheet = ss.getSheetByName("Purchase Orders");

    // Get ALL data from both sheets
    const detailsData = detailsSheet.getDataRange().getValues();
    const ordersData = ordersSheet.getDataRange().getValues();
    
    const detailsHeaders = detailsData[0];
    const ordersHeaders = ordersData[0];
    
    const detailIdCol = detailsHeaders.indexOf("Purchase Detail ID");
    const poIdCol = detailsHeaders.indexOf("Purchase Order ID");
    const ordersPoIdCol = ordersHeaders.indexOf("Purchase Order ID");
    
    console.log(`Details sheet has ${detailsData.length} rows`);
    console.log(`Orders sheet has ${ordersData.length} rows`);
    
    // Step 1: Get the PO data BEFORE deletion to know the supplier and amount paid
    let supplierId = '';
    let amountPaid = 0;
    let poRowToDelete = -1;
    
    for (let i = 1; i < ordersData.length; i++) {
      if (ordersData[i][ordersPoIdCol] === poId) {
        supplierId = ordersData[i][ordersHeaders.indexOf("Supplier ID")];
        amountPaid = parseFloat(ordersData[i][ordersHeaders.indexOf("Total Paid")]) || 0;
        poRowToDelete = i + 1; // +1 because sheet rows are 1-indexed
        console.log(`Found PO: Supplier=${supplierId}, Amount Paid=${amountPaid}`);
        break;
      }
    }
    
    if (!supplierId) {
      console.error(`PO ${poId} not found in Purchase Orders`);
      throw new Error(`Purchase Order ${poId} not found`);
    }
    
    // Step 2: Find and delete the detail row
    let detailRowToDelete = -1;
    for (let i = 1; i < detailsData.length; i++) {
      if (detailsData[i][detailIdCol] == detailId) {
        detailRowToDelete = i + 1; // +1 because sheet rows are 1-indexed
        console.log(`Found detail to delete at row ${detailRowToDelete}`);
        break;
      }
    }
    
    if (detailRowToDelete === -1) {
      console.error(`Purchase Detail ID ${detailId} not found in Purchase Details`);
      throw new Error(`Purchase Detail ID ${detailId} not found`);
    }
    
    // Delete the detail row
    detailsSheet.deleteRow(detailRowToDelete);
    console.log(`Deleted detail row ${detailRowToDelete}`);
    
    // Step 3: Count remaining details for this PO
    const updatedDetailsData = detailsSheet.getDataRange().getValues();
    let remainingDetailCount = 0;
    
    for (let i = 1; i < updatedDetailsData.length; i++) {
      if (updatedDetailsData[i][poIdCol] === poId) {
        remainingDetailCount++;
      }
    }
    
    console.log(`PO ${poId} has ${remainingDetailCount} remaining details after deletion`);
    
    // Step 4: If no details remain, delete the PO from Purchase Orders
    if (remainingDetailCount === 0) {
      console.log(`No details remain for PO ${poId}, deleting from Purchase Orders`);
      
      if (poRowToDelete !== -1) {
        ordersSheet.deleteRow(poRowToDelete);
        console.log(`Deleted PO row ${poRowToDelete}`);
        
        // SYNC: Delete the payment record
        poDeletePaymentRecord(poId);
        
        // IMPORTANT: Subtract the paid amount from supplier's Total Payments
        if (amountPaid > 0) {
          poUpdateSupplierPayments(supplierId, -amountPaid); // Negative amount to subtract
          console.log(`Subtracted ${amountPaid} from supplier ${supplierId} payments`);
        }
      } else {
        console.warn(`PO ${poId} not found in Purchase Orders table`);
      }
    } else {
      console.log(`PO ${poId} has remaining details, keeping in Purchase Orders`);
      // SYNC: Update the payment record with remaining data
      _poRecalcAll();
    }
    
    // Step 5: Recalculate everything
    console.log(`Starting recalculations...`);
    _poRecalcAll();
  
  } catch (error) {
    console.error('Error in poDeletePODetail:', error);
    throw error;
  }
}


/**
 * Save all edited detail rows, then recalc in order. 
 */
function poSavePODetails(updates) {
  console.log("=== STARTING poSavePODetails ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Purchase Details");
  const range = ss.getRangeByName("RANGEPURCHASEDETAILS");
  
  if (!range) {
    throw new Error("Purchase Details range not found");
  }
  
  const data = range.getValues();
  const hdr = data[0];
  
  // Helper to find column index by name
  const ci = name => hdr.indexOf(name);

  // Get PO ID from first update
  const poId = updates[0].poId;
  
  updates.forEach(u => {
    console.log(`Processing update for detailId: ${u.detailId}`);
    
    // Find the row by Detail ID
    const rowIdx = data.findIndex((r,i) => i>0 && r[ci("Purchase Detail ID")] == u.detailId);
    if (rowIdx < 1) throw new Error("Purchase Detail ID not found: " + u.detailId);
    
    const sheetRow = range.getRow() + rowIdx;

    // Update all editable fields
    if (ci("Item ID") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Item ID")).setValue(u.itemId);
    }
    if (ci("Item Name") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Item Name")).setValue(u.itemName);
    }
    if (ci("Quantity Purchased") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Quantity Purchased")).setValue(u.qtyPurchased);
    }
    if (ci("Unit Cost") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Unit Cost")).setValue(u.unitCost);
    }
    if (ci("Shipping Fee") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Shipping Fee")).setValue(u.shippingFees);
    }
    if (ci("Total Price") !== -1) {
      sheet.getRange(sheetRow, range.getColumn() + ci("Total Price")).setValue(u.totalPrice);
    }
  });

  // Recalculate everything
  revisetotalpo();
  poUpdatePOBalance();
  revisetotalinventory();
  poUpdateRemainingQty();
  poUpdateReorderRequired();
  poUpdateTotalPurchases();
  poUpdateBalancePayable();
  
  // NEW: Update payment status in Payments sheet if this PO has payments
  poUpdatePaymentStatus(poId);
}



/**
 * Loop through every PO in PurchaseOrders and SUMIF from PurchaseDetails.
 */
function revisetotalpo() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const pdRng   = ss.getRangeByName("RANGEPURCHASEDETAILS");
  const poRng   = ss.getRangeByName("RANGEPURCHASEORDERS");
  
  if (!pdRng || !poRng) {
    console.warn("revisetotalpo: One or more named ranges not found. Skipping update.");
    return;
  }

  const pdVals  = pdRng.getValues();
  const poVals  = poRng.getValues();
  const pdHdr   = pdVals[0];
  const poHdr   = poVals[0];
  const pdPoCol = pdHdr.indexOf("Purchase Order ID");
  const pdTotCol= pdHdr.indexOf("Total Price");
  const poPoCol = poHdr.indexOf("Purchase Order ID");
  const poTotCol= poHdr.indexOf("Total Amount");
  
  if (pdPoCol === -1 || pdTotCol === -1 || poPoCol === -1 || poTotCol === -1) {
    console.error("revisetotalpo: Required columns not found.");
    return;
  }

  // build sum map
  const sums = {};
  for(let i=1;i<pdVals.length;i++){
    // Skip empty rows in Purchase Details
    if (pdVals[i].every(cell => cell === '' || cell === null)) continue;
    
    const id = pdVals[i][pdPoCol];
    const val= Number(pdVals[i][pdTotCol])||0;
    if (id && id !== '') {
      sums[id]=(sums[id]||0)+val;
    }
  }

  const sheet = poRng.getSheet();
  const startRow = poRng.getRow();
  const startCol = poRng.getColumn();
  
  // write back per Purchase Order - ONLY for valid rows
  for(let r=1;r<poVals.length;r++){
    // Skip empty rows in Purchase Orders
    if (poVals[r].every(cell => cell === '' || cell === null)) continue;
    
    const id = poVals[r][poPoCol];
    // Skip if no PO ID or if this row doesn't have a matching sum
    if (!id || id === '' || !sums.hasOwnProperty(id)) continue;
    
    const newSum  = sums[id] || 0;
    const targetRow = startRow + r;
    const targetCol = startCol + poTotCol;
    
    sheet.getRange(targetRow, targetCol).setValue(newSum);
  }
  
  console.log("revisetotalpo: Updated Total Amounts successfully.");
}



function revisetotalinventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get named ranges
  const inventoryRange = ss.getRangeByName("RANGEINVENTORYITEMS");
  const purchaseRange = ss.getRangeByName("RANGEPURCHASEDETAILS");

  const inventoryData = inventoryRange.getValues();
  const purchaseData = purchaseRange.getValues();

  // Get headers
  const inventoryHeaders = inventoryData[0];
  const purchaseHeaders = purchaseData[0];

  // Find column indexes
  const invItemIdIndex = inventoryHeaders.indexOf("Item ID");
  const invQtyPurchasedIndex = inventoryHeaders.indexOf("Quantity Purchased");

  const pdItemIdIndex = purchaseHeaders.indexOf("Item ID");
  const pdQtyIndex = purchaseHeaders.indexOf("Quantity Purchased");

  if (invItemIdIndex === -1 || invQtyPurchasedIndex === -1 ||
      pdItemIdIndex === -1 || pdQtyIndex === -1) {
    throw new Error("One or more required columns (Item ID or Quantity Purchased) not found.");
  }

  // Create a map of Item ID to total Quantity Purchased from PurchaseDetails
  const qtyMap = {};
  for (let i = 1; i < purchaseData.length; i++) {
    const itemId = purchaseData[i][pdItemIdIndex];
    const qty = Number(purchaseData[i][pdQtyIndex]) || 0;
    if (itemId) {
      qtyMap[itemId] = (qtyMap[itemId] || 0) + qty;
    }
  }

  // Update only the Quantity Purchased column in the InventoryItems sheet
  const sheet = inventoryRange.getSheet();
  const startRow = inventoryRange.getRow();
  const startCol = inventoryRange.getColumn();
  
  for (let j = 1; j < inventoryData.length; j++) {
    // Skip empty rows
    if (inventoryData[j].every(cell => cell === '' || cell === null)) continue;
    
    const itemId = inventoryData[j][invItemIdIndex];
    const newQty = qtyMap[itemId] || 0;
    
    // Calculate exact cell position for Quantity Purchased
    const targetRow = startRow + j;
    const targetCol = startCol + invQtyPurchasedIndex;
    
    sheet.getRange(targetRow, targetCol).setValue(newQty);
  }
}

// Generate Payment ID
function poGeneratePaymentID() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Payments");
  const range = ss.getRangeByName("RANGEPAYMENTS");

  if (!range) {
    console.log("poGeneratePaymentID: RANGEPAYMENTS named range not found. Generating simple ID.");
    return "PMT" + Math.floor(10000 + Math.random() * 90000);
  }

  try {
    const data = range.getValues();
    const headers = data[0];
    const idIndex = headers.indexOf("Payment ID");
    const existingIds = new Set();

    if (idIndex === -1) {
      console.warn("poGeneratePaymentID: 'Payment ID' header not found in Payments. Generating simple ID.");
      return "PMT" + Math.floor(10000 + Math.random() * 90000);
    }

    for (let i = 1; i < data.length; i++) {
      const id = data[i][idIndex];
      if (id && id !== '') {
        existingIds.add(String(id));
      }
    }

    let newId;
    do {
      newId = "PMT" + Math.floor(10000 + Math.random() * 90000);
    } while (existingIds.has(newId));

    console.log("Generated new Payment ID: " + newId);
    return newId;
  } catch (e) {
    console.error("Error in poGeneratePaymentID: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to generate Payment ID: " + e.message);
  }
}

// Record payment in Payments sheet
function poRecordPayment(paymentData) {
  console.log("=== STARTING poRecordPayment ===");
  console.log("Payment Data:", paymentData);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName("Payments");
  const paymentsRange = ss.getRangeByName("RANGEPAYMENTS");

  if (!paymentsRange) {
    console.error("poRecordPayment: RANGEPAYMENTS named range not found.");
    throw new Error("Payments range not found.");
  }

  try {
    const paymentsData = paymentsRange.getValues();
    const paymentsHeaders = paymentsData[0];

    console.log("Payments Headers:", paymentsHeaders);

    // UPDATED: Format date as MM/dd/yyyy
    let formattedDate = paymentData.date;
    if (paymentData.date) {
      try {
        const dateObj = new Date(paymentData.date);
        if (!isNaN(dateObj)) {
          const month = (dateObj.getMonth() + 1).toString().padStart(2, '0');
          const day = dateObj.getDate().toString().padStart(2, '0');
          const year = dateObj.getFullYear();
          formattedDate = `${month}/${day}/${year}`;
        }
      } catch (e) {
        console.log(`Could not format payment date: ${paymentData.date}`);
      }
    }

    // Prepare new row
    const newRow = paymentsHeaders.map(header => {
      switch(header) {
        case "Date": return formattedDate;
        case "Payment ID": return paymentData.paymentId;
        case "Transaction Type": return paymentData.transactionType;
        case "Reference ID": return paymentData.referenceId;
        case "Party ID": return paymentData.partyId;
        case "Payment Mode": return paymentData.paymentMode || '';
        case "Amount": return paymentData.amount;
        case "Status": return paymentData.status;
        default: return '';
      }
    });

    // Append to sheet
    paymentsSheet.getRange(paymentsSheet.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);
    console.log("poRecordPayment: Successfully recorded payment with ID: " + paymentData.paymentId);
  } catch (e) {
    console.error("Error in poRecordPayment: " + e.message + " Stack: " + e.stack);
    throw new Error("Failed to record payment: " + e.message);
  }
}

// Update payment status in Payments sheet when PO changes
function poUpdatePaymentStatus(poId) {
  console.log(`=== STARTING poUpdatePaymentStatus for PO: ${poId} ===`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paymentsSheet = ss.getSheetByName("Payments");
  const paymentsRange = ss.getRangeByName("RANGEPAYMENTS");
  const ordersSheet = ss.getSheetByName("Purchase Orders");
  const ordersRange = ss.getRangeByName("RANGEPURCHASEORDERS");

  if (!paymentsRange || !ordersRange) {
    console.warn("poUpdatePaymentStatus: Required named ranges not found.");
    return;
  }

  try {
    const paymentsData = paymentsRange.getValues();
    const paymentsHeaders = paymentsData[0];
    const ordersData = ordersRange.getValues();
    const ordersHeaders = ordersData[0];

    // Get current PO status
    const poIdIndex = ordersHeaders.indexOf("Purchase Order ID");
    const statusIndex = ordersHeaders.indexOf("Payment Status");
    
    let currentStatus = "";
    for (let i = 1; i < ordersData.length; i++) {
      if (ordersData[i][poIdIndex] === poId) {
        currentStatus = ordersData[i][statusIndex];
        break;
      }
    }

    if (!currentStatus) {
      console.log(`PO ${poId} not found or no status available`);
      return;
    }

    // Update corresponding payment record
    const refIdIndex = paymentsHeaders.indexOf("Reference ID");
    const statusColIndex = paymentsHeaders.indexOf("Status");
    
    const startRow = paymentsRange.getRow();
    const startCol = paymentsRange.getColumn();
    
    let updated = false;
    for (let i = 1; i < paymentsData.length; i++) {
      if (paymentsData[i][refIdIndex] === poId) {
        const targetRow = startRow + i;
        const targetCol = startCol + statusColIndex;
        
        paymentsSheet.getRange(targetRow, targetCol).setValue(currentStatus);
        console.log(`Updated payment status for PO ${poId} to: ${currentStatus}`);
        updated = true;
        break;
      }
    }

    if (!updated) {
      console.log(`No payment record found for PO ${poId}`);
    }
  } catch (e) {
    console.error("Error in poUpdatePaymentStatus: " + e.message + " Stack: " + e.stack);
  }
}

function poGetPODetailsWithMaster(token, poId) {
  try {
    // Check view permission
    checkPurchasePermission(token, "CanView");
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    
    try {
      // Get master record from Purchase Orders sheet
      const poData = poGetRangeDataAsObjects('RANGEPURCHASEORDERS');
      const masterRecord = poData.find(po => po['Purchase Order ID'] === poId);
      
      if (!masterRecord) {
        throw new Error('Purchase Order not found: ' + poId);
      }
      
      console.log('Master Record:', masterRecord);
      
      // Format master date if needed
      if (masterRecord['Date'] instanceof Date) {
        masterRecord['Date'] = Utilities.formatDate(masterRecord['Date'], tz, 'MM/dd/yyyy');
      }

      // Get details from Purchase Details sheet
      const rawDetails = poGetRangeDataAsObjects('RANGEPURCHASEDETAILS')
                  .filter(r => r['Purchase Order ID'] === poId);
      
      console.log(`Found ${rawDetails.length} details for PO: ${poId}`);
      
      const details = rawDetails.map(row => {
        // Format date if needed
        if (row['Date'] instanceof Date) {
          row['Date'] = Utilities.formatDate(row['Date'], tz, 'MM/dd/yyyy');
        }
        
        // Return ALL the detail fields from Purchase Details sheet
        return {
          date: row['Date'],
          poId: row['Purchase Order ID'],
          detailId: row['Purchase Detail ID'],
          supplierName: row['Supplier Name'],
          state: row['State'],
          city: row['City'],
          itemId: row['Item ID'],
          itemType: row['Item Type'],
          itemCategory: row['Item Category'],
          itemName: row['Item Name'],
          qtyPurchased: row['Quantity Purchased'],
          unitCost: row['Unit Cost'],
          shippingFees: row['Shipping Fee'],
          totalPrice: row['Total Price']
        };
      });

      // Prepare master data for frontend - combine data from both sheets
      const masterData = {
        // From Purchase Orders sheet
        supplierName: masterRecord['Supplier Name'],
        date: masterRecord['Date'],
        totalAmount: masterRecord['Total Amount'],
        paymentStatus: masterRecord['Payment Status'],
        amountPaid: masterRecord['Total Paid'], // This is "Amount Paid" in the frontend
        
        // From Purchase Details sheet (for reference)
        poId: masterRecord['Purchase Order ID'],
        supplierId: masterRecord['Supplier ID'],
        state: masterRecord['State'],
        city: masterRecord['City'],
        poBalance: masterRecord['Purchase Order Balance']
      };

      console.log('Master Data for frontend:', masterData);
      console.log('Details for frontend:', details);

      return {
        master: masterData,
        details: details
      };
    } catch (error) {
      console.error('Error in poGetPODetailsWithMaster:', error);
      throw error;
    }
  
  } catch (error) {
    console.error('Error in poGetPODetailsWithMaster (permission):', error);
    throw error;
  }
}

// Save Purchase Order details with master updates
function poSavePODetailsWithMaster(token, updates, poId, updatedAmountPaid) {
  try {
    // Check edit permission
    checkPurchasePermission(token, "CanEdit");

    poLogStart('Save PO Details With Master', {
      poId: poId,
      updateCount: updates.length,
      updatedAmountPaid: updatedAmountPaid
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // STEP 1: Get OLD data BEFORE making changes
    const poRange = ss.getRangeByName('RANGEPURCHASEORDERS');
    const poAllValues = poRange.getValues();
    const poHeaders = poAllValues[0];
    const poData = poAllValues.slice(1);
    
    const poIDCol = poHeaders.indexOf('Purchase Order ID');
    const supplierIDCol = poHeaders.indexOf('Supplier ID');
    const oldAmountPaidCol = poHeaders.indexOf('Total Paid');
    
    // Find the PO and get old values
    let oldAmountPaid = 0;
    let supplierId = '';
    const poRowIndex = poData.findIndex(row => row[poIDCol] === poId);
    
    if (poRowIndex !== -1) {
      oldAmountPaid = parseFloat(poData[poRowIndex][oldAmountPaidCol]) || 0;
      supplierId = poData[poRowIndex][supplierIDCol];
      console.log(`Found PO ${poId}: Old Amount Paid = ${oldAmountPaid}, Supplier = ${supplierId}`);
    }

    // STEP 2: Update Purchase Details and calculate new total
    const pdRange = ss.getRangeByName('RANGEPURCHASEDETAILS');
    const pdSheet = pdRange.getSheet();
    const pdStartRow = pdRange.getRow();
    const pdAllValues = pdRange.getValues();
    const pdHeaders = pdAllValues[0];
    const pdData = pdAllValues.slice(1);
    
    const detailIdCol = pdHeaders.indexOf('Purchase Detail ID');
    
    // Calculate new total from updated details
    let newTotalAmount = 0;
    console.log("=== CALCULATING NEW TOTAL AMOUNT ===");
    
    updates.forEach(update => {
      const qty = parseFloat(update.qtyPurchased) || 0;
      const unitCost = parseFloat(update.unitCost) || 0;
      const shippingFees = parseFloat(update.shippingFees) || 0;
      
      const itemTotalPrice = (qty * unitCost) + shippingFees;
      newTotalAmount += itemTotalPrice;
      
      const detailId = update.detailId;
      const rowIndex = pdData.findIndex(row => row[detailIdCol] === detailId);
      
      if (rowIndex !== -1) {
        const sheetRow = pdStartRow + 1 + rowIndex;
        
        // Update fields with PROPERLY PARSED values
        const fields = [
          'Item ID', 'Item Name', 'Quantity Purchased', 'Unit Cost', 'Shipping Fee', 'Total Price'
        ];
        
        const fieldMap = {
          'Item ID': 'itemId',
          'Item Name': 'itemName', 
          'Quantity Purchased': 'qtyPurchased',
          'Unit Cost': 'unitCost',
          'Shipping Fee': 'shippingFees',
          'Total Price': 'totalPrice'
        };
        
        fields.forEach(field => {
          const colIndex = pdHeaders.indexOf(field);
          if (colIndex !== -1 && update[fieldMap[field]] !== undefined) {
            let value = update[fieldMap[field]];
            
            // Ensure numeric fields are properly formatted
            if (field === 'Quantity Purchased' || field === 'Unit Cost' || field === 'Shipping Fee' || field === 'Total Price') {
              value = parseFloat(value) || 0;
              
              // For Total Price, use our calculated value to avoid corruption
              if (field === 'Total Price') {
                value = itemTotalPrice;
              }
            }
            
            pdSheet.getRange(sheetRow, pdRange.getColumn() + colIndex).setValue(value);
          }
        });
      }
    });
    
    console.log(`=== FINAL CALCULATED TOTAL: ${newTotalAmount} ===`);

    // STEP 3: Update Purchase Orders master record
    const poSheet = poRange.getSheet();
    const poStartRow = poRange.getRow();
    
    const amountPaidCol = poHeaders.indexOf('Total Paid');
    const totalAmountCol = poHeaders.indexOf('Total Amount');
    const paymentStatusCol = poHeaders.indexOf('Payment Status');
    const balanceCol = poHeaders.indexOf('Purchase Order Balance');
    
    let newPaymentStatus = 'Unpaid';
    
    if (poRowIndex !== -1) {
      const poSheetRow = poStartRow + 1 + poRowIndex;
      
      // Update all fields at once
      poSheet.getRange(poSheetRow, poRange.getColumn() + totalAmountCol).setValue(newTotalAmount);
      poSheet.getRange(poSheetRow, poRange.getColumn() + amountPaidCol).setValue(updatedAmountPaid);
      
      const newBalance = newTotalAmount - updatedAmountPaid;
      poSheet.getRange(poSheetRow, poRange.getColumn() + balanceCol).setValue(newBalance);
      
      // Calculate Payment Status
      newPaymentStatus = poCalculatePaymentStatus(updatedAmountPaid, newTotalAmount);
      
      poSheet.getRange(poSheetRow, poRange.getColumn() + paymentStatusCol).setValue(newPaymentStatus);
      
      //console.log('Updated Purchase Order:', poId);
      //console.log('Total Amount:', newTotalAmount);
      //console.log('Amount Paid:', updatedAmountPaid);
      //console.log('Balance:', newBalance);
      //console.log('Payment Status:', newPaymentStatus);
      
      // STEP 4: Update supplier payments with the DIFFERENCE
      if (supplierId && supplierId !== '') {
        const paymentDifference = updatedAmountPaid - oldAmountPaid;
        console.log(`Supplier Payment Update: ${supplierId}, Difference = ${paymentDifference} (New: ${updatedAmountPaid} - Old: ${oldAmountPaid})`);
        
        if (paymentDifference !== 0) {
          poUpdateSupplierPayments(supplierId, paymentDifference);
        }
      }
      
      // STEP 5: FIXED - Update the payment record using dedicated function
      updatePurchasePaymentRecord(poId, updatedAmountPaid, newPaymentStatus, newTotalAmount);
    }
    
    // STEP 6: Recalculate all other metrics
    _poRecalcAll();
    
    poLogSuccess('Save PO Details With Master', {
      poId: poId,
      updatesApplied: updates.length,
      newTotalAmount: newTotalAmount,
      newPaymentStatus: newPaymentStatus
    });
    
    return true;
  
  } catch (error) {
    console.error('Error in poSavePODetailsWithMaster:', error);
    throw error;
  }
}

function updatePurchasePaymentRecord(poId, amountPaid, paymentStatus, totalAmount) {
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
    
    // Find the payment record for this Purchase Order (Purchase transactions only)
    const rowIndex = data.findIndex(row => 
      row[referenceIDCol] === poId && 
      row[transactionTypeCol] === 'Purchase'
    );
    
    if (rowIndex !== -1) {
      // Update existing record
      const sheetRow = startRow + 1 + rowIndex;
      
      if (amountCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
          .setValue(amountPaid);
      }
      
      if (statusCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
          .setValue(paymentStatus);
      }
      
      console.log('Updated existing payment record for PO:', poId);
      return true;
    } else {
      // Create new payment record if it doesn't exist
      console.log('Creating new payment record for PO:', poId);
      
      // Get supplier info from purchase order
      const poData = poGetRangeDataAsObjects('RANGEPURCHASEORDERS');
      const purchaseOrder = poData.find(po => po['Purchase Order ID'] === poId);
      
      if (purchaseOrder) {
        const paymentData = {
          date: purchaseOrder['Date'],
          paymentId: poGeneratePaymentID(),
          transactionType: "Purchase",
          referenceId: poId,
          partyId: purchaseOrder['Supplier ID'],
          paymentMode: '',
          amount: amountPaid,
          status: paymentStatus
        };
        
        poRecordPayment(paymentData);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Error updating purchase payment record:', error);
    return false;
  }
}

function poUpdatePaymentRecord(poId, updatedAmountPaid, updatedPaymentStatus) {
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
    const amountCol = headers.indexOf('Amount');
    const statusCol = headers.indexOf('Status');
    
    // Find the payment record for this Purchase Order
    const rowIndex = data.findIndex(row => row[referenceIDCol] === poId);
    
    if (rowIndex !== -1) {
      const sheetRow = startRow + 1 + rowIndex;
      
      // Update Amount and Status
      if (amountCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
          .setValue(updatedAmountPaid);
      }
      
      if (statusCol !== -1) {
        paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
          .setValue(updatedPaymentStatus);
      }
      
      console.log('Updated payment record for PO:', poId, 'Amount:', updatedAmountPaid, 'Status:', updatedPaymentStatus);
      return true;
    } else {
      console.log('No payment record found for PO:', poId);
      return false;
    }
    
  } catch (error) {
    console.error('Error updating payment record:', error);
    return false;
  }
}

function poDeletePaymentRecord(poId) {
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
    
    // Find the payment record for this Purchase Order
    const rowIndex = data.findIndex(row => row[referenceIDCol] === poId);
    
    if (rowIndex !== -1) {
      // Delete the row (add 2 because data starts at row 2 and array index starts at 0)
      paymentsSheet.deleteRow(startRow + 1 + rowIndex);
      console.log('Deleted payment record for PO:', poId);
      return true;
    } else {
      console.log('No payment record found for PO:', poId);
      return false;
    }
    
  } catch (error) {
    console.error('Error deleting payment record:', error);
    return false;
  }
}

function poSyncAllPaymentRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const paymentsRange = ss.getRangeByName('RANGEPAYMENTS');
    const poData = poGetRangeDataAsObjects('RANGEPURCHASEORDERS');
    
    if (!paymentsRange || !poData || poData.length === 0) {
      return;
    }
    
    const paymentsSheet = paymentsRange.getSheet();
    const startRow = paymentsRange.getRow();
    const allValues = paymentsRange.getValues();
    const headers = allValues[0];
    const paymentData = allValues.slice(1);
    
    const referenceIDCol = headers.indexOf('Reference ID');
    const amountCol = headers.indexOf('Amount');
    const statusCol = headers.indexOf('Status');
    
    // Update each payment record
    paymentData.forEach((paymentRow, index) => {
      const poID = paymentRow[referenceIDCol];
      
      if (poID) {
        const purchaseOrder = poData.find(po => po['Purchase Order ID'] === poID);
        
        if (purchaseOrder) {
          // Purchase order exists - update the payment record
          const sheetRow = startRow + 1 + index;
          
          if (amountCol !== -1) {
            paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + amountCol)
              .setValue(parseFloat(purchaseOrder['Total Paid']) || 0);
          }
          
          if (statusCol !== -1) {
            paymentsSheet.getRange(sheetRow, paymentsRange.getColumn() + statusCol)
              .setValue(purchaseOrder['Payment Status'] || 'Unpaid');
          }
        } else {
          // Purchase order doesn't exist - delete the payment record
          paymentsSheet.deleteRow(startRow + 1 + index);
          console.log('Deleted orphaned payment record for PO:', poID);
        }
      }
    });
    
    console.log('Completed payment records sync for Purchase');
    
  } catch (error) {
    console.error('Error syncing payment records:', error);
  }
}

// Internal: run all recalculation routines in order - UPDATED
function _poRecalcAll() {  
  
  revisetotalpo();
  poUpdatePOBalance();
  revisetotalinventory();
  poUpdateRemainingQty();
  poUpdateReorderRequired();
  poUpdateTotalPurchases();
  
  // CRITICAL: Comprehensive supplier payment and balance updates
  poRecalculateAllSupplierPayments(); // This ensures data integrity
  poUpdateBalancePayable();    
    
}

// Helper function to get range data as objects (if not already exists)
function poGetRangeDataAsObjects(rangeName) {
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
 * Helper: get named range data as array of objects (like Sales page)
 */
function poGetRangeDataAsObjects(rangeName) {
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

// NEW: Recalculate all supplier payments from Payments sheet (for data integrity)
function poRecalculateAllSupplierPayments() {
  console.log("=== STARTING poRecalculateAllSupplierPayments ===");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suppliersSheet = ss.getSheetByName("Suppliers");
  const paymentsSheet = ss.getSheetByName("Payments");
  
  const suppliersRange = ss.getRangeByName("RANGESUPPLIERS");
  const paymentsRange = ss.getRangeByName("RANGEPAYMENTS");

  if (!suppliersRange || !paymentsRange) {
    console.warn("Required named ranges not found for payment recalculation");
    return;
  }

  try {
    const suppliersData = suppliersRange.getValues();
    const suppliersHeaders = suppliersData[0];
    const paymentsData = paymentsRange.getValues();
    const paymentsHeaders = paymentsData[0];

    const supplierIdIndex = suppliersHeaders.indexOf("Supplier ID");
    const totalPaymentsIndex = suppliersHeaders.indexOf("Total Payments");
    const paymentsPartyIdIndex = paymentsHeaders.indexOf("Party ID");
    const paymentsAmountIndex = paymentsHeaders.indexOf("Amount");

    // Create a map of supplier total payments
    const supplierPayments = {};
    
    // Sum all payments for each supplier
    for (let i = 1; i < paymentsData.length; i++) {
      const supplierId = paymentsData[i][paymentsPartyIdIndex];
      const amount = parseFloat(paymentsData[i][paymentsAmountIndex]) || 0;
      
      if (supplierId && supplierId !== '') {
        if (!supplierPayments[supplierId]) {
          supplierPayments[supplierId] = 0;
        }
        supplierPayments[supplierId] += amount;
      }
    }

    console.log("Supplier Payments Map:", supplierPayments);

    // Update suppliers sheet
    const startRow = suppliersRange.getRow();
    const startCol = suppliersRange.getColumn();
    
    let updateCount = 0;
    for (let i = 1; i < suppliersData.length; i++) {
      if (suppliersData[i].every(cell => cell === '' || cell === null)) continue;
      
      const supplierId = suppliersData[i][supplierIdIndex];
      const totalPayments = supplierPayments[supplierId] || 0;
      
      const targetRow = startRow + i;
      const targetCol = startCol + totalPaymentsIndex;
      
      suppliersSheet.getRange(targetRow, targetCol).setValue(totalPayments);
      console.log(`Updated ${supplierId}: Total Payments = ${totalPayments}`);
      updateCount++;
    }
    
    console.log(`poRecalculateAllSupplierPayments: Updated ${updateCount} suppliers`);
    
    // Now update the balances
    poUpdateBalancePayable();
    
  } catch (e) {
    console.error("Error in poRecalculateAllSupplierPayments: " + e.message);
  }
}

/**
 * SAFE OPTIMIZATION 1: Centralized payment status calculation
 * Replaces duplicate logic across multiple functions
 */
function poCalculatePaymentStatus(totalPaid, totalAmount) {
  if (totalPaid === 0) return "Unpaid";
  if (totalPaid >= totalAmount) return "Paid";
  return "Partial";
}

/**
 * SAFE OPTIMIZATION 2: Consistent logging
 * Adds standardized logging without changing error handling
 */

// Main logging function
function poLogOperation(operationName, success = true, context = {}) {
  const timestamp = new Date().toISOString();
  const icon = success ? '✅' : '❌';
  const logMessage = `${icon} [${timestamp}] ${operationName}`;
  
  if (Object.keys(context).length > 0) {
    console.log(logMessage, context);
  } else {
    console.log(logMessage);
  }
}

// Helper for operation start
function poLogStart(operationName, context = {}) {
  poLogOperation(`START: ${operationName}`, true, context);
}

// Helper for operation success  
function poLogSuccess(operationName, context = {}) {
  poLogOperation(`SUCCESS: ${operationName}`, true, context);
}

// Helper for operation failure
function poLogError(operationName, error, context = {}) {
  const errorContext = {
    ...context,
    error: error.message,
    stack: error.stack
  };
  poLogOperation(`FAILED: ${operationName}`, false, errorContext);
}

/**
 * Delete entire Purchase Order and all its details
 * Handles proper recalculation of all related data
 */
function poDeleteEntirePurchaseOrder(token, poId) {
  try {
     // Check delete permission
    checkPurchasePermission(token, "CanDelete");

    console.log(`=== STARTING poDeleteEntirePurchaseOrder for: ${poId} ===`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    try {
      // STEP 1: Get PO master data BEFORE deletion (for supplier payments)
      const poData = poGetRangeDataAsObjects('RANGEPURCHASEORDERS');
      const masterRecord = poData.find(po => po['Purchase Order ID'] === poId);
      
      if (!masterRecord) {
        throw new Error(`Purchase Order ${poId} not found`);
      }
      
      const supplierId = masterRecord['Supplier ID'];
      const amountPaid = parseFloat(masterRecord['Total Paid']) || 0;
      
      console.log(`PO Details - Supplier: ${supplierId}, Amount Paid: ${amountPaid}`);

      // STEP 2: Delete all Purchase Details for this PO
      const detailsSheet = ss.getSheetByName("Purchase Details");
      const detailsRange = ss.getRangeByName("RANGEPURCHASEDETAILS");
      const detailsData = detailsRange.getValues();
      const detailsHeaders = detailsData[0];
      
      const poIdCol = detailsHeaders.indexOf("Purchase Order ID");
      const rowsToDelete = [];
      
      // Find all rows to delete (from bottom to top to avoid index shifting)
      for (let i = detailsData.length - 1; i >= 1; i--) {
        if (detailsData[i][poIdCol] === poId) {
          rowsToDelete.push(i + detailsRange.getRow()); // +1 for header row offset
        }
      }
      
      // Delete rows from bottom to top
      rowsToDelete.sort((a, b) => b - a); // Sort descending
      rowsToDelete.forEach(row => {
        detailsSheet.deleteRow(row);
      });
      
      console.log(`Deleted ${rowsToDelete.length} detail rows for PO ${poId}`);

      // STEP 3: Delete PO master record
      const ordersSheet = ss.getSheetByName("Purchase Orders");
      const ordersRange = ss.getRangeByName("RANGEPURCHASEORDERS");
      const ordersData = ordersRange.getValues();
      const ordersHeaders = ordersData[0];
      
      const ordersPoIdCol = ordersHeaders.indexOf("Purchase Order ID");
      let poRowToDelete = -1;
      
      for (let i = 1; i < ordersData.length; i++) {
        if (ordersData[i][ordersPoIdCol] === poId) {
          poRowToDelete = i + ordersRange.getRow();
          break;
        }
      }
      
      if (poRowToDelete !== -1) {
        ordersSheet.deleteRow(poRowToDelete);
        console.log(`Deleted PO master record at row ${poRowToDelete}`);
      }

      // STEP 4: Delete payment record
      poDeletePaymentRecord(poId);

      // STEP 5: Update supplier payments (subtract the paid amount)
      if (supplierId && amountPaid > 0) {
        poUpdateSupplierPayments(supplierId, -amountPaid);
        console.log(`Subtracted ${amountPaid} from supplier ${supplierId} payments`);
      }

      // STEP 6: Recalculate EVERYTHING
      console.log("Starting comprehensive recalculations...");
      _poRecalcAll();

      return {
        success: true,
        message: `Purchase Order ${poId} and all its items deleted successfully`,
        detailsDeleted: rowsToDelete.length
      };
      
    } catch (error) {
      console.error("Error in poDeleteEntirePurchaseOrder:", error);
      return {
        success: false,
        message: error.message
      };
    }
  
  } catch (error) {
    console.error('Error in poDeleteEntirePurchaseOrder:', error);
    throw error;
  }
}