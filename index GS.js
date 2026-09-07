/**
 * Reads a named range and returns rows of objects keyed by header.
 */
function dashGetRows(rangeName) {
  try {
    const ss = SpreadsheetApp.getActive();
    const range = ss.getRangeByName(rangeName);
    if (!range) {
      console.warn('Named range ' + rangeName + ' not found');
      return [];
    }
    const values = range.getValues();
    if (values.length < 2) return []; // No data rows
    
    const [headers, ...rows] = values;
    return rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
  } catch (error) {
    console.error('Error reading range ' + rangeName + ':', error);
    return [];
  }
}

/**
 * Main entrypoint: returns all KPI values and chart data.
 */
function dashGetDashboardData() {
  try {
    // Fetch detail rows with error handling using correct named ranges
    const salesDetails = dashGetRows('RANGESALESDETAILS') || [];
    const purchaseDetails = dashGetRows('RANGEPURCHASEDETAILS') || [];
    const customers = dashGetRows('RANGECUSTOMERS') || [];
    const suppliers = dashGetRows('RANGESUPPLIERS') || [];
    const salesOrders = dashGetRows('RANGESALESORDERS') || [];
    const purchaseOrders = dashGetRows('RANGEPURCHASEORDERS') || [];

    console.log('Data loaded - Sales Details:', salesDetails.length, 'Purchase Details:', purchaseDetails.length);

    // KPI 1 & 2: Total Sales and Purchases
    const totalSales = salesDetails.reduce((sum, r) => sum + (Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0), 0);
    const totalPurchases = purchaseDetails.reduce((sum, r) => sum + (Number(r['Total Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0), 0);
    const netProfit = totalSales - totalPurchases;

    // KPI 4 & 5: Receivables and Payables
    const totalReceivable = customers.reduce((sum, r) => sum + (Number(r['Balance Receivable']) || Number(r['Receivable Balance']) || Number(r['Outstanding Balance']) || 0), 0);
    const totalPayable = suppliers.reduce((sum, r) => sum + (Number(r['Balance Payable']) || Number(r['Payable Balance']) || Number(r['Outstanding Balance']) || 0), 0);

    // KPI 6: Top Sales Location
    const salesByCity = {};
    salesDetails.forEach(r => {
      const city = r['City'] || r['Location'] || r['State'] || 'Unknown';
      const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      salesByCity[city] = (salesByCity[city] || 0) + amount;
    });
    const topLocation = Object.entries(salesByCity).sort((a,b)=>b[1]-a[1])[0]?.[0] || 'N/A';

    // KPI 7: Top Selling Item
    const salesByItem = {};
    salesDetails.forEach(r => {
      const item = r['Item Name'] || r['Item Type'] || r['Product Name'] || 'Unknown';
      const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      salesByItem[item] = (salesByItem[item] || 0) + amount;
    });
    const topItem = Object.entries(salesByItem).sort((a,b)=>b[1]-a[1])[0]?.[0] || 'N/A';

    /*** Chart Data ***/

    // Chart 1: Sales Trend (group by month)
    const trendMap = {};
    salesDetails.forEach(r => {
      try {
        const dateStr = r['Sales Order Date'] || r['Sales Date'] || r['Date'] || r['Order Date'];
        if (dateStr) {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-01');
            const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
            trendMap[key] = (trendMap[key] || 0) + amount;
          }
        }
      } catch (e) {
        // Skip invalid dates
      }
    });
    const salesTrendDates = Object.keys(trendMap).sort();
    const salesTrendValues = salesTrendDates.map(d => trendMap[d]);

    // Chart 2: Sales By Location (State)
    const stateMap = {};
    salesDetails.forEach(r => {
      const st = r['State'] || r['Location'] || 'Unknown';
      const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      stateMap[st] = (stateMap[st] || 0) + amount;
    });
    const salesByLocation = { 
      labels: Object.keys(stateMap), 
      values: Object.values(stateMap) 
    };

    // Chart 3: Sales By Category
    const catMap = {};
    salesDetails.forEach(r => {
      const c = r['Item Type'] || r['Category'] || r['Product Category'] || 'Unknown';
      const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      catMap[c] = (catMap[c] || 0) + amount;
    });
    const totalCat = Object.values(catMap).reduce((a,b)=>a+b,0);
    const salesByCategory = {
      labels: Object.keys(catMap),
      values: totalCat > 0 ? Object.values(catMap).map(v => (v/totalCat)*100) : []
    };

    // Chart 4: Top 10 Customers
    const custMap = {};
    salesDetails.forEach(r => {
      const c = r['Customer Name'] || r['Customer'] || 'Unknown';
      const amount = Number(r['Total Sales Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      custMap[c] = (custMap[c] || 0) + amount;
    });
    const topCustArr = Object.entries(custMap).sort((a,b)=>b[1]-a[1]).slice(0,10);
    const topCustomers = { 
      labels: topCustArr.map(a=>a[0]), 
      values: topCustArr.map(a=>a[1]) 
    };

    // Chart 5: Purchase By Location (State)
    const purStateMap = {};
    purchaseDetails.forEach(r => {
      const st = r['State'] || r['Location'] || 'Unknown';
      const amount = Number(r['Total Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
      purStateMap[st] = (purStateMap[st] || 0) + amount;
    });
    const totalPurState = Object.values(purStateMap).reduce((a,b)=>a+b,0);
    const purchaseByLocation = {
      labels: Object.keys(purStateMap),
      values: totalPurState > 0 ? Object.values(purStateMap).map(v=> (v/totalPurState)*100) : []
    };

    // Chart 6: Purchase By Category stacked by year
    const purCatYear = {};
    purchaseDetails.forEach(r => {
      try {
        const dateStr = r['Date'] || r['Purchase Order Date'] || r['Purchase Date'] || r['Order Date'];
        if (dateStr) {
          const d = new Date(dateStr);
          if (!isNaN(d.getTime())) {
            const y = d.getFullYear();
            const c = r['Item Type'] || r['Category'] || r['Product Category'] || 'Unknown';
            const amount = Number(r['Total Purchase Price']) || Number(r['Total Amount']) || Number(r['Amount']) || 0;
            purCatYear[y] = purCatYear[y] || {};
            purCatYear[y][c] = (purCatYear[y][c] || 0) + amount;
          }
        }
      } catch (e) {
        // Skip invalid dates
      }
    });
    const years = Object.keys(purCatYear).sort();
    const items = Array.from(new Set(purchaseDetails.map(r=>r['Item Type'] || r['Category']).filter(Boolean)));
    const series = items.map(item => ({
      name: item,
      data: years.map(y => purCatYear[y]?.[item] || 0)
    }));
    const purchaseByCategory = { years, series };

    // Chart 7: Sales By City (Treemap)
    const treemapData = Object.entries(salesByCity)
      .sort((a,b)=>b[1]-a[1])
      .map(([city, val])=>({ x: city, y: val }));

    return {
      totalSales,
      totalPurchases,
      netProfit,
      totalReceivable,
      totalPayable,
      topLocation,
      topItem,
      salesTrend: { dates: salesTrendDates, values: salesTrendValues },
      salesByLocation,
      salesByCategory,
      topCustomers,
      purchaseByLocation,
      purchaseByCategory,
      salesByCity: treemapData
    };
  } catch (error) {
    console.error('Error generating dashboard data:', error);
    throw new Error('Failed to generate dashboard data: ' + error.message);
  }
}