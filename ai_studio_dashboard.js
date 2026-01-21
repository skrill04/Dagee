function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('DAGEEB BRIGHT - ERP System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- SALES DATA ---
function getDashboardData(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheetName) sheetName = "CONTAINER RECORDS 2025";
    
    // Smart Search for Sheet
    const sheets = ss.getSheets();
    let sheet = sheets.find(s => s.getName().trim().toUpperCase() === sheetName.trim().toUpperCase());
    if (!sheet) {
       sheet = sheets.find(s => s.getName().toUpperCase().startsWith("CONTAINER RECORDS"));
    }

    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Fetch Cols A to AE
    const range = sheet.getRange(2, 1, lastRow - 1, 31); 
    const data = range.getValues();
    let cleanData = [];

    data.forEach(row => {
      let entryDate = row[0];
      let containerNo = row[1];
      
      // Strict Filter
      if (containerNo && String(containerNo).trim().length > 3) {
        const cNo = String(containerNo).toUpperCase();
        const keywordsToSkip = ["AGENT", "PAYMENT", "NOTE", "STATUS", "TOTAL", "CONTAINER", "SUM", "BALANCE"];
        
        if (!keywordsToSkip.some(w => cNo.includes(w))) {
           
           let finalDateStr = "2025-01-01";
           let timestamp = 0;
           try {
             let d = (entryDate instanceof Date) ? entryDate : new Date(entryDate);
             if (!isNaN(d.getTime())) {
               finalDateStr = Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
               timestamp = d.getTime();
             }
           } catch(e) {}

           // --- MONEY MAPPING ---
           // 1. PRODUCT COST (Column N - Index 13)
           let productCost = cleanMoney(row[13]);

           // 2. LOGISTICS COST (Column V - Index 21 - "Total Clearance Cost")
           // This is usually accurate in the sheet. If 0, we sum manual parts as backup.
           let colV_Clearance = cleanMoney(row[21]);
           
           // Manual parts (Only used for popup details or if Col V is empty)
           let duty = cleanMoney(row[14]);        
           let demurrage = cleanMoney(row[15]);   
           let shipLinePay = cleanMoney(row[16]);    
           let gpha = cleanMoney(row[17]);        
           let transport = cleanMoney(row[18]);   
           let fda = cleanMoney(row[19]);         
           let agent = cleanMoney(row[20]);       
           let storage = cleanMoney(row[24]);     
           let discharging = cleanMoney(row[25]); 

           let manualSum = duty + demurrage + shipLinePay + gpha + transport + fda + agent + storage + discharging;
           
           // Priority: Use Column V. If it's missing, use manual sum.
           let finalLogisticsCost = colV_Clearance !== 0 ? colV_Clearance : manualSum;

           cleanData.push({
             date: finalDateStr,
             timestamp: timestamp, 
             containerNo: row[1], 
             billOfLading: row[2] || "-", 
             shippingLine: row[3] || "-", 
             shipper: row[4] || "-",         // BRAND/MAKE
             port: row[5] || "-",
             product: row[6] || "-", 
             
             quantity: cleanMoney(row[8]),
             unitPrice: cleanMoney(row[9]),
             totalSales: cleanMoney(row[10]),
             costUSD: cleanMoney(row[11]),
             rate: cleanMoney(row[12]),
             productCostGHS: productCost, 
             
             // Details for Popup
             duty: duty, demurrage: demurrage, shipLinePay: shipLinePay,
             gpha: gpha, transport: transport, fda: fda,
             agent: agent, storage: storage, discharging: discharging,
             
             // THE CORRECT LOGISTICS COST (Col V)
             totalAmount: finalLogisticsCost
           });
        }
      }
    });
    
    cleanData.sort((a, b) => b.timestamp - a.timestamp);
    return cleanData;
  } catch (err) {
    return []; 
  }
}

// --- INCOMING DATA ---
function getIncomingData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const sheet = sheets.find(s => s.getName().toUpperCase().includes("INCOMING"));
    if (!sheet) return [{ error: "Sheet 'INCOMING' not found." }];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [{ error: "Sheet found but looks empty." }];

    const range = sheet.getRange(2, 1, lastRow - 1, 10); 
    const data = range.getValues();
    let incomingList = [];

    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      if (!row[1] || String(row[1]).trim() === "") continue;
      if (String(row[1]).toUpperCase().includes("CONTAINER")) continue;

      try {
        let etaStr = "TBD";
        let daysLeft = 999;
        
        if (row[2]) {
           let d = (row[2] instanceof Date) ? row[2] : new Date(row[2]);
           if (!isNaN(d.getTime())) {
             etaStr = Utilities.formatDate(d, "GMT", "MMM dd, yyyy");
             let now = new Date();
             let diff = d.getTime() - now.getTime();
             daysLeft = Math.ceil(diff / (1000 * 60 * 60 * 24));
           }
        }
        let arrivedVal = String(row[7]).toUpperCase();
        let clearedVal = String(row[8]).toUpperCase();
        let isArrived = (arrivedVal.includes("YES") || arrivedVal.includes("Y"));
        let isCleared = (clearedVal.includes("YES") || clearedVal.includes("Y"));

        incomingList.push({
          blNo: row[0] || "-", containerNo: row[1], eta: etaStr, daysLeft: daysLeft,
          product: row[3] || "-", brand: row[4] || "-", country: row[5] || "-",
          shippingLine: row[6] || "-", arrived: isArrived, cleared: isCleared
        });
      } catch (e) { }
    }
    incomingList.sort((a, b) => a.daysLeft - b.daysLeft);
    return incomingList;
  } catch (err) { return [{ error: err.toString() }]; }
}

function getAvailableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName()).filter(n => n.startsWith("CONTAINER RECORDS")).sort();
}

function cleanMoney(raw) {
  if (!raw) return 0;
  if (typeof raw === 'number') return raw;
  let str = String(raw).replace(/GHS/g, '').replace(/,/g, '').replace(/\s/g, '').replace(/-/g, '0');
  let val = parseFloat(str);
  return isNaN(val) ? 0 : val;
}