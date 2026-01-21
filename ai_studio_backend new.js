function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('DAGEEB BRIGHT ENTERPRISE')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- 1. SALES DATA ---
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
      let rawDate = row[0];
      let containerNo = row[1];
      
      // Strict Filter: Must have Container No and not be a header/total row
      if (containerNo && String(containerNo).trim().length > 3) {
        const cNo = String(containerNo).toUpperCase();
        const keywordsToSkip = ["AGENT", "PAYMENT", "NOTE", "STATUS", "TOTAL", "CONTAINER", "SUM", "BALANCE"];
        
        if (!keywordsToSkip.some(w => cNo.includes(w))) {
           
           // --- FIX: Handle "14, November 2025" ---
           let finalDateStr = "2025-01-01";
           let timestamp = 0;
           
           // If it's a string with a comma, remove it so Date() can read it
           let dateVal = rawDate;
           if (typeof rawDate === 'string') {
             dateVal = rawDate.replace(/,/g, ''); 
           }

           try {
             let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
             if (!isNaN(d.getTime())) {
               finalDateStr = Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
               timestamp = d.getTime();
             }
           } catch(e) {}

           // --- COST MAPPING ---
           // Column N: Product Cost
           let productCost = cleanMoney(row[13]); 

           // Column V: Total Clearance (Logistics)
           let colV = cleanMoney(row[21]);
           
           // Breakdown for Popup
           let duty = cleanMoney(row[14]);        
           let demurrage = cleanMoney(row[15]);   
           let shipLine = cleanMoney(row[16]);    
           let gpha = cleanMoney(row[17]);        
           let transport = cleanMoney(row[18]);   
           let fda = cleanMoney(row[19]);         
           let agent = cleanMoney(row[20]);       
           let storage = cleanMoney(row[24]);     
           let discharging = cleanMoney(row[25]); 

           let manualSum = duty + demurrage + shipLine + gpha + transport + fda + agent + storage + discharging;
           
           // FIX: Use Col V. If 0, use manual sum. 
           // This ensures we NEVER accidentally add the Product Cost to Logistics.
           let finalLogistics = colV !== 0 ? colV : manualSum;

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
             productCostGHS: productCost,    // SEPARATE from Logistics
             
             // Details
             duty: duty, demurrage: demurrage, shipLinePay: shipLine,
             gpha: gpha, transport: transport, fda: fda,
             agent: agent, storage: storage, discharging: discharging,
             
             // STRICT LOGISTICS COST
             totalAmount: finalLogistics
           });
        }
      }
    });
    
    // Sort Oldest to Newest (So Graph flows Left-to-Right)
    cleanData.sort((a, b) => a.timestamp - b.timestamp);
    return cleanData;
  } catch (err) {
    return []; 
  }
}

// --- 2. INCOMING DATA ---
function getIncomingData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(s => s.getName().toUpperCase().includes("INCOMING"));
    if (!sheet) return [{ error: "Sheet 'INCOMING' not found." }];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [{ error: "Sheet found but looks empty." }];
    const range = sheet.getRange(2, 1, lastRow - 1, 10); 
    const data = range.getValues();
    let incomingList = [];

    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      if (!row[1] || String(row[1]).trim() === "" || String(row[1]).toUpperCase().includes("CONTAINER")) continue;
      try {
        let etaStr = "TBD"; let daysLeft = 999;
        
        let dateVal = row[2];
        if (dateVal) {
             // Fix Comma date issue
             if (typeof dateVal === 'string') dateVal = dateVal.replace(/,/g, '');
             let d = new Date(dateVal);
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

// --- 3. PAYROLL DATA ---
function getPayrollData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets().find(s => s.getName().toUpperCase().includes("SALARY"));
    if (!sheet) return [{ error: "Sheet 'SALARY RECORDS' not found." }];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [{ error: "No salary records found." }];
    const range = sheet.getRange(2, 1, lastRow - 1, 11); 
    const data = range.getValues();
    let payrollList = [];
    
    data.forEach(row => {
      let name = String(row[0]).trim();
      if (name !== "" && name.toUpperCase() !== "NAME") {
         let periodStr = row[4];
         let monthIndex = -1;
         if (row[4] instanceof Date) {
            periodStr = Utilities.formatDate(row[4], "GMT", "MMM yyyy");
            monthIndex = row[4].getMonth();
         } else {
             // Handle text dates if needed
             let d = new Date(row[4]);
             if (!isNaN(d)) monthIndex = d.getMonth();
         }

         payrollList.push({
           name: row[0], role: row[1], dept: row[2], method: row[3], period: periodStr, monthIndex: monthIndex,
           basic: cleanMoney(row[5]), containers: cleanMoney(row[6]), commRate: cleanMoney(row[7]),
           deductions: cleanMoney(row[8]), netPay: cleanMoney(row[9]), status: row[10]
         });
      }
    });
    return payrollList;
  } catch(err) { return [{ error: err.toString() }]; }
}

function getAvailableSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName()).filter(n => n.startsWith("CONTAINER RECORDS")).sort();
}

// --- MASTER FIX FOR MONEY (Negatives & Text) ---
function cleanMoney(raw) {
  if (!raw) return 0;
  if (typeof raw === 'number') return raw;
  
  let str = String(raw);
  
  // 1. Remove Currency and Commas
  str = str.replace(/GHS|₵|\$|,| /g, ''); 
  
  // 2. Check if it's negative (some formats use () for negative or just -)
  if (str.includes('(') && str.includes(')')) {
     str = '-' + str.replace(/\(|\)/g, '');
  }
  
  let val = parseFloat(str);
  return isNaN(val) ? 0 : val;
}