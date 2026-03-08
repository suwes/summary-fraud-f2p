// ===== KONFIGURASI =====
const CONFIG = {
  SHEET_NAMES: {
    REGISTRY: 'FILE_REGISTRY',
    SUMMARY: 'SUMMARY_FRAUD',
    LOGS: 'LOGS_SUMMARY'
  },
  SUMMARY_HEADERS: [
    'Bulan Periode', 'Nama Hub', 'Kota', 'Provinsi', 'Region', 'Lead', 'Asst Lead', 'PIC Area', 'Koordinator Lapangan', 'LatLon Provinsi', 'LatLong Hub', 'is_hub',
    'CF COD', 'NCF COD', 'CF LND', 'NCF LND', 'CF OTHER', 'NCF OTHER',
    'Total Fraud CF', 'Total Fraud NCF',
    'COD', 'LND', 'OTHER', 'Total Fraud',
    'Rider Mitra Fraud', 'Operator Mitra Fraud', 'Rider Dedicated Fraud', 'Operator Dedicated Fraud', 'Total Mitra Fraud', 'Total Dedicated Fraud',
    'Count COD', 'Count LND', 'Count OTHER', 'Count Case Total', 'Count Case HUB', 'Count Case DC',
    'Total Hold Gaji', 'Total Refund Nek', 'Total Refund E-wallet', 'Total Collect PIC',
    'Total Recovery', 'Total Loss', 'Principal Loss', 'Fraud Hold',
    'Total Invoice CF', 'Total Invoice NCF', 'Total Invoice',
    'Total Karyawan CF', 'Total Karyawan NCF', 'Total Karyawan CF Fraud', 'Total Karyawan NCF Fraud',
    'Total Karyawan Aktif', 'Total Employee Duplicate', 'Total Karyawan Fraud',
    'Rider Mitra', 'Operator Mitra', 'Rider Dedicated', 'Operator Dedicated', 'Total Mitra', 'Total Dedicated'
  ]
};

// ===== FUNGSI UTAMA =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Fraud Summary')
    .addItem('🔄 Generate Summary', 'generateSummary')
    .addItem('📋 View Logs', 'showLogs')
    .addSeparator()
    .addItem('⏰ Set Daily Trigger', 'showTriggerDialog')
    .addItem('📅 Set Weekday Triggers (3x)', 'showWeekdayTriggerDialog')
    .addItem('👁️ View Active Triggers', 'viewActiveTriggers')
    .addItem('🗑️ Remove All Triggers', 'removeAllTriggers')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function generateSummary() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('🔄 Memulai generate summary...', 'Processing', -1);
    
    // Show processing indicator
    showProcessingIndicator(ss);
    
    // 1. Baca registry
    const registry = readRegistry(ss);
    SpreadsheetApp.getActiveSpreadsheet().toast(`📁 Ditemukan ${registry.length} file aktif`, 'Processing', 3);
    
    // 2. Load semua data dari file terpisah
    const allData = loadAllDataFromExternalFiles(registry);
    
    // 3. Generate summary
    SpreadsheetApp.getActiveSpreadsheet().toast('⚙️ Mengkalkulasi summary...', 'Processing', -1);
    const summaryData = generateSummaryData(allData);
    
    // 4. Tulis ke sheet
    SpreadsheetApp.getActiveSpreadsheet().toast('💾 Menulis hasil ke sheet...', 'Processing', -1);
    writeSummaryToSheet(ss, summaryData);
    
    // 5. Rename file
    renameSpreadsheet(ss);
    
    // 6. Log hasil
    const duration = (new Date() - startTime) / 1000;
    const totalRows = Object.values(allData).reduce((sum, arr) => sum + arr.length, 0);
    
    // Hitung total untuk log summary
    const logSummary = calculateLogSummary(summaryData);
    
    logExecution(ss, {
      status: 'Success',
      filesProcessed: registry.length,
      totalRecords: summaryData.length,
      rowsProcessed: totalRows,
      duration: duration,
      error: '-',
      summary: logSummary
    });
    
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ Summary berhasil dibuat!', 'Success', 5);
    
  } catch (error) {
    const duration = (new Date() - startTime) / 1000;
    logExecution(ss, {
      status: 'Failed',
      filesProcessed: 0,
      totalRecords: 0,
      rowsProcessed: 0,
      duration: duration,
      error: error.toString(),
      summary: {cod: 0, lnd: 0, other: 0, totalFraud: 0, holdGaji: 0, refundNek: 0, refundEwallet: 0, collectPic: 0, totalRecovery: 0, totalLoss: 0, activeEmployees: 0, fraudDrivers: 0}
    });
    
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.toString());
    Logger.log('Error detail: ' + error.stack);
  }
}

// ===== FUNGSI BACA REGISTRY =====
function readRegistry(ss) {
  const registrySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.REGISTRY);
  if (!registrySheet) throw new Error('Sheet FILE_REGISTRY tidak ditemukan!');
  
  const data = registrySheet.getDataRange().getValues();
  const headers = data[0];
  const registry = [];
  
  // Cari index kolom
  const fileNameIdx = headers.indexOf('File Name');
  const urlIdx = headers.indexOf('Spreadsheet URL');
  const statusIdx = headers.indexOf('Status');
  const sheetNameIdx = headers.indexOf('Sheets Name');
  const dataTypeIdx = headers.indexOf('Data Type');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusIdx];
    
    if (status && status.toString().toLowerCase() === 'active') {
      const url = row[urlIdx] ? row[urlIdx].toString() : '';
      const spreadsheetId = extractSpreadsheetId(url);
      
      if (spreadsheetId) {
        registry.push({
          fileName: row[fileNameIdx],
          url: url,
          spreadsheetId: spreadsheetId,
          status: status,
          sheetName: row[sheetNameIdx],
          dataType: row[dataTypeIdx]
        });
      }
    }
  }
  Logger.log(`Total file aktif: ${registry.length}`);
  return registry;
}

// ===== EXTRACT SPREADSHEET ID DARI URL =====
function extractSpreadsheetId(url) {
  if (!url) return null;
  
  // Pattern: /spreadsheets/d/{SPREADSHEET_ID}/
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// ===== FUNGSI LOAD DATA DARI FILE EKSTERNAL =====
function loadAllDataFromExternalFiles(registry) {
  const data = {
    claims: [],
    refundNek: [],
    refundEwallet: [],
    collectPic: [],
    payrollHold: [],
    hubs: [],
    invoices: []
  };
  
  let processedCount = 0;
  
  registry.forEach((entry, index) => {
    try {
      processedCount++;
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `📥 Loading ${processedCount}/${registry.length}: ${entry.fileName}...`, 
        'Processing', 
        -1
      );
      
      Logger.log(`Loading file ${processedCount}/${registry.length}: ${entry.fileName} (${entry.dataType})`);
      
      // Buka spreadsheet eksternal
      const externalSs = SpreadsheetApp.openById(entry.spreadsheetId);
      const sheet = externalSs.getSheetByName(entry.sheetName);
      
      if (!sheet) {
        Logger.log(`⚠️ Sheet "${entry.sheetName}" tidak ditemukan di ${entry.fileName}`);
        return;
      }
      
      const sheetData = sheet.getDataRange().getValues();
      if (sheetData.length <= 1) {
        Logger.log(`⚠️ Sheet "${entry.sheetName}" kosong atau hanya ada header`);
        return;
      }
      
      const headers = sheetData[0];
      let rowCount = 0;
      
      // Parse data
      for (let i = 1; i < sheetData.length; i++) {
        const row = {};
        let hasData = false;
        
        headers.forEach((header, idx) => {
          const value = sheetData[i][idx];
          if (value !== '' && value !== null && value !== undefined) {
            hasData = true;
          }
          row[header] = value;
        });
        
        if (!hasData) continue; // Skip baris kosong
        
        rowCount++;
        
        // Kategorikan berdasarkan tipe data
        switch(entry.dataType) {
          case 'Claim':
            data.claims.push(row);
            break;
          case 'Refund NEK':
            data.refundNek.push(row);
            break;
          case 'Refund Ewallet':
            data.refundEwallet.push(row);
            break;
          case 'Collection by PIC':
            data.collectPic.push(row);
            break;
          case 'Payroll Hold':
            data.payrollHold.push(row);
            break;
          case 'HUB RAW':
            data.hubs.push(row);
            break;
          case 'Invoice':
            data.invoices.push(row);
            break;
          default:
            Logger.log(`⚠️ Tipe data tidak dikenali: ${entry.dataType}`);
        }
      }
      
      Logger.log(`✓ Loaded ${rowCount} rows from ${entry.fileName}`);
      
    } catch (e) {
      Logger.log(`❌ Error loading ${entry.fileName}: ${e.toString()}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `⚠️ Gagal load ${entry.fileName}: ${e.message}`, 
        'Warning', 
        3
      );
    }
  });
  
  // Log summary
  Logger.log('=== DATA LOADING SUMMARY ===');
  Logger.log(`Claims: ${data.claims.length} rows`);
  Logger.log(`Refund NEK: ${data.refundNek.length} rows`);
  Logger.log(`Refund E-wallet: ${data.refundEwallet.length} rows`);
  Logger.log(`Collect PIC: ${data.collectPic.length} rows`);
  Logger.log(`Payroll Hold: ${data.payrollHold.length} rows`);
  Logger.log(`Hubs: ${data.hubs.length} rows`);
  Logger.log(`Invoices: ${data.invoices.length} rows`);
  
  return data;
}

// ===== FUNGSI GENERATE SUMMARY =====
function generateSummaryData(allData) {
  const summaryMap = new Map();
  
  // 1. Inisialisasi dari HUB data sebagai patokan
  allData.hubs.forEach(hub => {
    const hubName = (hub['HUB Name'] || hub['Hub Name'] || '').toString().toUpperCase().trim();
    const city = (hub['City'] || '').toString().toUpperCase().trim();
    const province = (hub['Provinsi'] || hub['Province'] || '').toString().toUpperCase().trim();
    
    if (!hubName) return;
    
    const key = hubName;
    if (!summaryMap.has(key)) {
      summaryMap.set(key, {
        hub: hubName,
        city: city,
        province: province,
        region: hub['Region'] || '',
        lead: hub['Lead Region'] || '',
        asstLead: hub['Asst Lead'] || '',
        picArea: hub['PIC Area'] || '',
        koordinatorLapangan: hub['Koordinator Lapangan'] || '',
        provinceLatLon: hub['Province LatLon'] || '',
        hubLatLon: hub['HUB LatLon'] || '',
        months: new Map()
      });
    }
  });
  
  Logger.log(`Total hubs dari HUB RAW: ${summaryMap.size}`);
  
  // 2. Process Claims
  let claimProcessed = 0;
  allData.claims.forEach(claim => {
    const location = (claim['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(claim['Periode Month'] || claim['Periode']);
    const caseType = (claim['Case Type'] || '').toString().toUpperCase();
    const func = claim['Function'] || '';
    const amount = parseFloat(claim['COGS (Amount Claim)'] || claim['Amount Claim'] || 0);
    const driverId = claim['OPS / Driver ID'] || claim['OPS/Driver ID'] || claim['Driver ID'];
    const driverName = claim['Name'] || claim['Driver Name'] || '';
    const contract = claim['Contract'] || '';
    
    if (!location || !month) return;
    
    // Cari hub yang exact match atau buat hub baru jika tidak ada
    let hubData = summaryMap.get(location);
    
    if (!hubData) {
      // Hub tidak ada di HUB RAW, buat hub baru
      Logger.log(`Hub baru ditemukan di Claim: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    claimProcessed++;
    
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    
    const monthData = hubData.months.get(month);
    
    // Unique key untuk count case: ID + Nama + Hub + Case Type + Bulan
    const uniqueCaseKey = `${driverId}|${driverName.trim().toUpperCase()}|${location}|${caseType}`;
    
    // Unique key untuk count contract: ID + Nama + Hub + Contract + Bulan
    const contractClean = contract.trim();
    const uniqueContractKey = `${driverId}|${driverName.trim().toUpperCase()}|${location}|${contractClean}`;
    
    // Kategorikan berdasarkan Case Type dan Function
    if (caseType.includes('COD')) {
      if (func === 'CF') monthData.cfCod += amount;
      else monthData.ncfCod += amount;
      monthData.cod += amount;
      monthData.uniqueCaseCod.add(uniqueCaseKey);
    } else if (caseType.includes('LND') || caseType.includes('L&D')) {
      if (func === 'CF') monthData.cfLnd += amount;
      else monthData.ncfLnd += amount;
      monthData.lnd += amount;
      monthData.uniqueCaseLnd.add(uniqueCaseKey);
    } else {
      if (func === 'CF') monthData.cfOther += amount;
      else monthData.ncfOther += amount;
      monthData.other += amount;
      monthData.uniqueCaseOther.add(uniqueCaseKey);
    }
    
    // Count contract type (unique)
    if (contractClean && driverId && driverId.toString().trim() !== '') {
      if (contractClean === 'Rider Mitra') {
        monthData.uniqueRiderMitra.add(uniqueContractKey);
        monthData.riderMitraFraud += amount;
      } else if (contractClean === 'Operator Mitra') {
        monthData.uniqueOperatorMitra.add(uniqueContractKey);
        monthData.operatorMitraFraud += amount;
      } else if (contractClean === 'Rider Dedicated') {
        monthData.uniqueRiderDedicated.add(uniqueContractKey);
        monthData.riderDedicatedFraud += amount;
      } else if (contractClean === 'Operator Dedicated') {
        monthData.uniqueOperatorDedicated.add(uniqueContractKey);
        monthData.operatorDedicatedFraud += amount;
      }
    }
    
    // Track unique drivers - simple count unique ID, skip empty
    if (driverId && driverId.toString().trim() !== '') {
      const driverIdStr = driverId.toString().trim();
      const driverNameClean = driverName ? driverName.toString().trim().toUpperCase() : '';
      
      if (func === 'CF') {
        monthData.cfFraudDrivers.add(driverIdStr);
        if (driverNameClean) monthData.cfFraudDriverNames.add(driverNameClean);
      } else {
        monthData.ncfFraudDrivers.add(driverIdStr);
        if (driverNameClean) monthData.ncfFraudDriverNames.add(driverNameClean);
      }
      monthData.fraudDrivers.add(driverIdStr);
      if (driverNameClean) monthData.fraudDriverNames.add(driverNameClean);
    }
  });
  Logger.log(`Claims processed: ${claimProcessed}`);
  
  // 3. Process Payroll Hold
  let holdProcessed = 0;
  allData.payrollHold.forEach(payroll => {
    const location = (payroll['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(payroll['Periode'] || payroll['Periode Month']);
    const amount = parseFloat(payroll['Final HOLD'] || payroll['Total Deduction'] || 0);
    
    if (!location || !month || amount === 0) return;
    
    let hubData = summaryMap.get(location);
    if (!hubData) {
      Logger.log(`Hub baru ditemukan di Payroll Hold: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    holdProcessed++;
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    hubData.months.get(month).holdGaji += amount;
  });
  Logger.log(`Payroll Hold processed: ${holdProcessed}`);
  
  // 4. Process Refund NEK
  let refundNekProcessed = 0;
  allData.refundNek.forEach(refund => {
    const location = (refund['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(refund['Month'] || refund['Periode Month']);
    const amount = parseFloat(refund['Amount Claim'] || 0);
    
    if (!location || !month || amount === 0) return;
    
    let hubData = summaryMap.get(location);
    if (!hubData) {
      Logger.log(`Hub baru ditemukan di Refund NEK: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    refundNekProcessed++;
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    hubData.months.get(month).refundNek += amount;
  });
  Logger.log(`Refund NEK processed: ${refundNekProcessed}`);
  
  // 5. Process Refund E-wallet
  let ewalletProcessed = 0;
  allData.refundEwallet.forEach(refund => {
    const location = (refund['Hub'] || refund['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(refund['Periode Month']);
    const amount = parseFloat(refund['closing_balance'] || 0);
    
    if (!location || !month || amount === 0) return;
    
    let hubData = summaryMap.get(location);
    if (!hubData) {
      Logger.log(`Hub baru ditemukan di Refund E-wallet: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    ewalletProcessed++;
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    hubData.months.get(month).refundEwallet += amount;
  });
  Logger.log(`E-wallet processed: ${ewalletProcessed}`);
  
  // 6. Process Collect PIC
  let collectProcessed = 0;
  allData.collectPic.forEach(collect => {
    const location = (collect['HUB Location'] || collect['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(collect['Periode Month']);
    const amount = parseFloat(collect['Amount'] || 0);
    
    if (!location || !month || amount === 0) return;
    
    let hubData = summaryMap.get(location);
    if (!hubData) {
      Logger.log(`Hub baru ditemukan di Collect PIC: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    collectProcessed++;
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    hubData.months.get(month).collectPic += amount;
  });
  Logger.log(`Collect PIC processed: ${collectProcessed}`);
  
  // 7. Process Invoice
  let invoiceProcessed = 0;
  allData.invoices.forEach(invoice => {
    const location = (invoice['Location'] || '').toString().toUpperCase().trim();
    const month = parseMonth(invoice['Periode']);
    const amount = parseFloat(invoice['Subtotal I'] || 0);
    const func = invoice['Function'] || '';
    const empId = invoice['ID'];
    const empName = invoice['Name'] || '';
    
    if (!location || !month) return;
    
    let hubData = summaryMap.get(location);
    if (!hubData) {
      Logger.log(`Hub baru ditemukan di Invoice: ${location}`);
      hubData = {
        hub: location,
        city: '',
        province: '',
        region: '',
        lead: '',
        asstLead: '',
        picArea: '',
        koordinatorLapangan: '',
        provinceLatLon: '',
        hubLatLon: '',
        months: new Map()
      };
      summaryMap.set(location, hubData);
    }
    
    invoiceProcessed++;
    
    if (!hubData.months.has(month)) {
      hubData.months.set(month, initMonthData());
    }
    
    const monthData = hubData.months.get(month);
    monthData.invoice += amount;
    
    // Track invoice per function
    if (func === 'CF') {
      monthData.invoiceCf += amount;
    } else if (func === 'NCF') {
      monthData.invoiceNcf += amount;
    }
    
    // Track unique employees - simple count unique ID, skip empty
    if (empId && empId.toString().trim() !== '') {
      const empIdStr = empId.toString().trim();
      const empNameClean = empName ? empName.toString().trim().toUpperCase() : '';
      
      if (func === 'CF') {
        monthData.cfEmployees.add(empIdStr);
        if (empNameClean) monthData.cfEmployeeNames.add(empNameClean);
      } else if (func === 'NCF') {
        monthData.ncfEmployees.add(empIdStr);
        if (empNameClean) monthData.ncfEmployeeNames.add(empNameClean);
      }
      monthData.activeEmployees.add(empIdStr);
      if (empNameClean) monthData.activeEmployeeNames.add(empNameClean);
    }
  });
  Logger.log(`Invoice processed: ${invoiceProcessed}`);
  
  // 8. Flatten data untuk output
  const result = [];
  for (let [hubKey, hubData] of summaryMap) {
    for (let [month, monthData] of hubData.months) {
      const totalFraud = monthData.cod + monthData.lnd + monthData.other;
      const totalFraudCf = monthData.cfCod + monthData.cfLnd + monthData.cfOther;
      const totalFraudNcf = monthData.ncfCod + monthData.ncfLnd + monthData.ncfOther;
      const totalRecovery = monthData.holdGaji + monthData.refundNek + 
                           monthData.refundEwallet + monthData.collectPic;
      
      // Calculate loss (bisa minus), principal loss (tidak bisa minus), dan fraud hold
      const totalLoss = totalFraud - totalRecovery; // Bisa minus
      const principalLoss = totalFraud >= totalRecovery ? totalFraud - totalRecovery : 0; // Tidak bisa minus
      const fraudHold = totalFraud < totalRecovery ? Math.abs(totalRecovery - totalFraud) : 0; // Surplus
      
      // Count duplicate employees - berapa nama yang punya lebih dari 1 ID
      const duplicateCount = calculateDuplicateNames(monthData.activeEmployees, monthData.activeEmployeeNames);
      
      // Check if hub name contains "HUB"
      const isHub = hubData.hub.includes('HUB') ? 1 : 0;
      
      // Count case total
      const countCaseTotal = monthData.uniqueCaseCod.size + monthData.uniqueCaseLnd.size + monthData.uniqueCaseOther.size;
      
      // Count case HUB and DC
      const countCaseHub = isHub === 1 ? countCaseTotal : 0;
      const countCaseDc = isHub === 0 ? countCaseTotal : 0;
      
      // Calculate fraud amount by contract
      const totalMitraFraud = monthData.riderMitraFraud + monthData.operatorMitraFraud;
      const totalDedicatedFraud = monthData.riderDedicatedFraud + monthData.operatorDedicatedFraud;
      
      // Count contract
      const riderMitra = monthData.uniqueRiderMitra.size;
      const operatorMitra = monthData.uniqueOperatorMitra.size;
      const riderDedicated = monthData.uniqueRiderDedicated.size;
      const operatorDedicated = monthData.uniqueOperatorDedicated.size;
      const totalMitra = riderMitra + operatorMitra;
      const totalDedicated = riderDedicated + operatorDedicated;
      
      result.push([
        month,
        hubData.hub,
        hubData.city,
        hubData.province,
        hubData.region,
        hubData.lead,
        hubData.asstLead,
        hubData.picArea,
        hubData.koordinatorLapangan,
        hubData.provinceLatLon,
        hubData.hubLatLon,
        isHub,
        monthData.cfCod,
        monthData.ncfCod,
        monthData.cfLnd,
        monthData.ncfLnd,
        monthData.cfOther,
        monthData.ncfOther,
        totalFraudCf,
        totalFraudNcf,
        monthData.cod,
        monthData.lnd,
        monthData.other,
        totalFraud,
        monthData.riderMitraFraud,
        monthData.operatorMitraFraud,
        monthData.riderDedicatedFraud,
        monthData.operatorDedicatedFraud,
        totalMitraFraud,
        totalDedicatedFraud,
        monthData.uniqueCaseCod.size,
        monthData.uniqueCaseLnd.size,
        monthData.uniqueCaseOther.size,
        countCaseTotal,
        countCaseHub,
        countCaseDc,
        monthData.holdGaji,
        monthData.refundNek,
        monthData.refundEwallet,
        monthData.collectPic,
        totalRecovery,
        totalLoss,
        principalLoss,
        fraudHold,
        monthData.invoiceCf,
        monthData.invoiceNcf,
        monthData.invoice,
        monthData.cfEmployees.size,
        monthData.ncfEmployees.size,
        monthData.cfFraudDrivers.size,
        monthData.ncfFraudDrivers.size,
        monthData.activeEmployees.size,
        duplicateCount,
        monthData.fraudDrivers.size,
        riderMitra,
        operatorMitra,
        riderDedicated,
        operatorDedicated,
        totalMitra,
        totalDedicated
      ]);
    }
  }
  
  // Sort berdasarkan bulan (ascending) lalu hub name
  result.sort((a, b) => {
    const monthA = parseMonthForSort(a[0]);
    const monthB = parseMonthForSort(b[0]);
    
    if (monthA !== monthB) {
      return monthA - monthB;
    }
    
    return a[1].localeCompare(b[1]);
  });
  
  Logger.log(`Total summary rows generated: ${result.length}`);
  Logger.log(`Total hubs (termasuk hub baru): ${summaryMap.size}`);
  return result;
}

// ===== HELPER FUNCTIONS =====
function initMonthData() {
  return {
    cfCod: 0, ncfCod: 0, cfLnd: 0, ncfLnd: 0, cfOther: 0, ncfOther: 0,
    cod: 0, lnd: 0, other: 0,
    riderMitraFraud: 0, operatorMitraFraud: 0, riderDedicatedFraud: 0, operatorDedicatedFraud: 0,
    uniqueCaseCod: new Set(),
    uniqueCaseLnd: new Set(),
    uniqueCaseOther: new Set(),
    uniqueRiderMitra: new Set(),
    uniqueOperatorMitra: new Set(),
    uniqueRiderDedicated: new Set(),
    uniqueOperatorDedicated: new Set(),
    holdGaji: 0, refundNek: 0, refundEwallet: 0, collectPic: 0, 
    invoiceCf: 0, invoiceNcf: 0, invoice: 0,
    cfEmployees: new Set(),
    ncfEmployees: new Set(),
    cfEmployeeNames: new Set(),
    ncfEmployeeNames: new Set(),
    cfFraudDrivers: new Set(),
    ncfFraudDrivers: new Set(),
    cfFraudDriverNames: new Set(),
    ncfFraudDriverNames: new Set(),
    activeEmployees: new Set(),
    activeEmployeeNames: new Set(),
    fraudDrivers: new Set(),
    fraudDriverNames: new Set()
  };
}

function calculateDuplicateNames(idSet, nameSet) {
  // Duplicate = berapa nama yang punya lebih dari 1 ID
  // Jika ada 10 ID unik tapi hanya 8 nama unik, berarti ada 2 nama yang duplicate
  return idSet.size - nameSet.size;
}

function matchLocation(location, hubName) {
  if (!location || !hubName) return false;
  
  // Exact match only
  return location === hubName;
}

function parseMonth(dateStr) {
  if (!dateStr) return null;
  
  const str = dateStr.toString().trim();
  
  // Format: "November 2025" atau "Nov 2025"
  const monthMap = {
    'JANUARY': 'January', 'JAN': 'January',
    'FEBRUARY': 'February', 'FEB': 'February',
    'MARCH': 'March', 'MAR': 'March',
    'APRIL': 'April', 'APR': 'April',
    'MAY': 'May', 'MEI': 'May',
    'JUNE': 'June', 'JUN': 'June',
    'JULY': 'July', 'JUL': 'July',
    'AUGUST': 'August', 'AUG': 'August', 'AGU': 'August',
    'SEPTEMBER': 'September', 'SEP': 'September',
    'OCTOBER': 'October', 'OCT': 'October',
    'NOVEMBER': 'November', 'NOV': 'November',
    'DECEMBER': 'December', 'DEC': 'December', 'DES': 'December'
  };
  
  const upperStr = str.toUpperCase();
  
  for (let [key, monthName] of Object.entries(monthMap)) {
    if (upperStr.includes(key)) {
      const year = str.match(/\d{4}/);
      if (year) {
        return `${monthName} ${year[0]}`;
      }
    }
  }
  
  return null;
}

function parseMonthForSort(monthStr) {
  if (!monthStr) return 999999;
  
  const monthOrder = {
    'January': 1, 'February': 2, 'March': 3, 'April': 4,
    'May': 5, 'June': 6, 'July': 7, 'August': 8,
    'September': 9, 'October': 10, 'November': 11, 'December': 12
  };
  
  const parts = monthStr.split(' ');
  const month = parts[0];
  const year = parseInt(parts[1]) || 0;
  
  const monthNum = monthOrder[month] || 0;
  
  // Return format: YYYYMM (e.g., 202501 for January 2025)
  return year * 100 + monthNum;
}

// ===== TULIS KE SHEET =====
function showProcessingIndicator(ss) {
  let summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(CONFIG.SHEET_NAMES.SUMMARY);
  }
  
  // Hitung kolom terakhir yang ada data
  const lastCol = CONFIG.SUMMARY_HEADERS.length;
  
  // Clear hanya area data (tidak termasuk header dan kolom tambahan)
  if (summarySheet.getLastRow() > 1) {
    const rangeToClear = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, lastCol);
    rangeToClear.clearContent();
  }
  
  // Tulis "Processing..." di cell A2
  summarySheet.getRange(2, 1).setValue('Processing...').setFontStyle('italic').setFontColor('#999999');
  SpreadsheetApp.flush();
}

function writeSummaryToSheet(ss, summaryData) {
  let summarySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.SUMMARY);
  
  if (!summarySheet) {
    summarySheet = ss.insertSheet(CONFIG.SHEET_NAMES.SUMMARY);
  }
  
  // Hitung kolom terakhir yang ada data
  const lastCol = CONFIG.SUMMARY_HEADERS.length;
  
  // Clear area data termasuk "Processing..." (tidak termasuk header dan kolom tambahan)
  if (summarySheet.getLastRow() > 1) {
    const rangeToClear = summarySheet.getRange(2, 1, summarySheet.getLastRow() - 1, lastCol);
    rangeToClear.clearContent();
  }
  
  // Tulis header jika belum ada atau masih kosong
  if (summarySheet.getLastRow() === 0 || summarySheet.getRange(1, 1).getValue() === '') {
    summarySheet.getRange(1, 1, 1, CONFIG.SUMMARY_HEADERS.length)
      .setValues([CONFIG.SUMMARY_HEADERS])
      .setFontWeight('bold')
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  
  // Tulis data
  if (summaryData.length > 0) {
    summarySheet.getRange(2, 1, summaryData.length, summaryData[0].length)
      .setValues(summaryData);
    
    // Format angka (kolom 13-60 adalah angka, skip kolom 12 is_hub)
    summarySheet.getRange(2, 13, summaryData.length, 48)
      .setNumberFormat('#,##0');
  }
  
  // Freeze header row
  summarySheet.setFrozenRows(1);
  
  Logger.log('Data successfully written to sheet');
}

// ===== RENAME SPREADSHEET =====
function renameSpreadsheet(ss) {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const hh = String(now.getHours()).padStart(2, '0');
  const min = String(now.getMinutes()).padStart(2, '0');
  
  const newName = `SUMMARY_FRAUD_${dd}${mm}${yyyy}_${hh}:${min}`;
  ss.rename(newName);
  Logger.log(`Spreadsheet renamed to: ${newName}`);
}

// ===== CALCULATE LOG SUMMARY =====
function calculateLogSummary(summaryData) {
  const summary = {
    cod: 0,
    lnd: 0,
    other: 0,
    totalFraud: 0,
    holdGaji: 0,
    refundNek: 0,
    refundEwallet: 0,
    collectPic: 0,
    totalRecovery: 0,
    totalLoss: 0,
    activeEmployees: 0,
    fraudDrivers: 0
  };
  
  summaryData.forEach(row => {
    summary.cod += row[19] || 0;           // COD (kolom 20)
    summary.lnd += row[20] || 0;           // LND (kolom 21)
    summary.other += row[21] || 0;         // OTHER (kolom 22)
    summary.totalFraud += row[22] || 0;    // Total Fraud (kolom 23)
    summary.holdGaji += row[33] || 0;      // Total Hold Gaji (kolom 34)
    summary.refundNek += row[34] || 0;     // Total Refund Nek (kolom 35)
    summary.refundEwallet += row[35] || 0; // Total Refund E-wallet (kolom 36)
    summary.collectPic += row[36] || 0;    // Total Collect PIC (kolom 37)
    summary.totalRecovery += row[37] || 0; // Total Recovery (kolom 38)
    summary.totalLoss += row[38] || 0;     // Total Loss (kolom 39)
    summary.activeEmployees += row[46] || 0; // Total Karyawan Aktif (kolom 47)
    summary.fraudDrivers += row[48] || 0;    // Total Karyawan Fraud (kolom 49)
  });
  
  return summary;
}

// ===== LOGGING =====
function logExecution(ss, logData) {
  let logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
  
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEET_NAMES.LOGS);
    const headers = [
      'Time Execute / Run', 'Last Updated', 'Files Processed', 
      'Total Records', 'Rows Processed', 'Duration (min)', 'Status', 
      'COD', 'LND', 'OTHER', 'Total Fraud', 'Hold Gaji', 
      'Refund NEK', 'Refund E-Wallet', 'Collect by PIC', 'Total Recovery', 
      'Total Loss', 'Total Employee', 'Employee Fraud', 'Error Detail'
    ];
    logSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    logSheet.setFrozenRows(1);
  }
  
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 
                                   'EEEE, dd MMMM yyyy HH:mm:ss');
  
  const summary = logData.summary || {};
  
  const rowData = [
    now,                              // Time Execute / Run
    now,                              // Last Updated
    logData.filesProcessed,           // Files Processed
    logData.totalRecords,             // Total Records
    logData.rowsProcessed,            // Rows Processed
    (logData.duration / 60).toFixed(2), // Duration (min)
    logData.status,                   // Status
    summary.cod || 0,                 // COD
    summary.lnd || 0,                 // LND
    summary.other || 0,               // OTHER
    summary.totalFraud || 0,          // Total Fraud
    summary.holdGaji || 0,            // Hold Gaji
    summary.refundNek || 0,           // Refund NEK
    summary.refundEwallet || 0,       // Refund E-Wallet
    summary.collectPic || 0,          // Collect by PIC
    summary.totalRecovery || 0,       // Total Recovery
    summary.totalLoss || 0,           // Total Loss
    summary.activeEmployees || 0,     // Total Employee
    summary.fraudDrivers || 0,        // Employee Fraud
    logData.error                     // Error Detail
  ];
  
  logSheet.appendRow(rowData);
  
  // Format status
  const lastRow = logSheet.getLastRow();
  const statusCell = logSheet.getRange(lastRow, 7);
  if (logData.status === 'Success') {
    statusCell.setBackground('#B7E1CD').setFontWeight('bold');
  } else {
    statusCell.setBackground('#F4C7C3').setFontWeight('bold');
  }
  
  // Format angka di kolom summary (kolom 8-19)
  logSheet.getRange(lastRow, 8, 1, 12).setNumberFormat('#,##0');
  
  // Auto resize
  logSheet.autoResizeColumns(1, 20);
}

// ===== UI FUNCTIONS =====
function showLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.LOGS);
  
  if (logSheet) {
    logSheet.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('📋 Menampilkan log eksekusi', 'Logs', 3);
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ Log sheet belum ada.\n\nGenerate summary terlebih dahulu untuk membuat log.');
  }
}

function showAbout() {
  const html = `
    <div style="font-family: Arial; padding: 20px;">
      <h2 style="color: #4285F4;">📊 Fraud Summary Generator</h2>
      <p><strong>Version:</strong> 2.0.0</p>
      <p><strong>Description:</strong> Generate summary fraud data dari berbagai file Google Sheets terpisah</p>
      <hr>
      <h3>✨ Features:</h3>
      <ul style="line-height: 1.8;">
        <li>🔗 Load data dari multiple Google Sheets (external files)</li>
        <li>📊 Auto-generate summary per hub dan periode</li>
        <li>🎯 Tracking fraud CF & NCF</li>
        <li>💰 Calculate recovery dan loss otomatis</li>
        <li>📝 Logging setiap eksekusi dengan detail</li>
        <li>🔄 Dynamic: Cukup update FILE_REGISTRY untuk tambah data baru</li>
      </ul>
      <hr>
      <h3>📖 Cara Penggunaan:</h3>
      <ol style="line-height: 1.8;">
        <li>Pastikan FILE_REGISTRY sudah terisi dengan benar</li>
        <li>Klik menu <strong>📊 Fraud Summary > 🔄 Generate Summary</strong></li>
        <li>Tunggu proses selesai (akan ada notifikasi)</li>
        <li>Check hasil di sheet <strong>SUMMARY_FRAUD</strong></li>
        <li>Lihat log eksekusi di sheet <strong>LOGS_SUMMARY</strong></li>
      </ol>
      <hr>
      <p style="color: #666; font-size: 12px;">
        <strong>Note:</strong> Untuk menambah data source baru, cukup tambahkan entry baru di FILE_REGISTRY 
        dengan status "Active" tanpa perlu ubah script.
      </p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Fraud Summary Generator');
}

// ===== TRIGGER METADATA MANAGEMENT =====
function saveTriggerMetadata(triggerId, hari, jam) {
  const props = PropertiesService.getScriptProperties();
  const metadata = {
    hari: hari,
    jam: jam
  };
  props.setProperty('trigger_' + triggerId, JSON.stringify(metadata));
}

function getTriggerMetadata(triggerId) {
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty('trigger_' + triggerId);
  if (data) {
    return JSON.parse(data);
  }
  return { hari: 'Setiap hari', jam: '-' };
}

function deleteTriggerMetadata(triggerId) {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('trigger_' + triggerId);
}

// ===== VIEW ACTIVE TRIGGERS =====
function viewActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const summaryTriggers = triggers.filter(t => t.getHandlerFunction() === 'generateSummary');
  
  if (summaryTriggers.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Active Triggers', 'Tidak ada trigger aktif untuk generate summary.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let html = '<div style="font-family: Arial; padding: 20px;">';
  html += '<h3 style="color: #4285F4;">📋 Active Triggers</h3>';
  html += '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">';
  html += '<tr style="background: #f0f0f0; font-weight: bold;">';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">No</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Tipe</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Hari</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Jam</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Status</th>';
  html += '<th style="border: 1px solid #ddd; padding: 8px;">Aksi</th>';
  html += '</tr>';
  
  summaryTriggers.forEach((trigger, index) => {
    const triggerId = trigger.getUniqueId();
    const metadata = getTriggerMetadata(triggerId);
    
    html += '<tr>';
    html += `<td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${index + 1}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">Time-based</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px;">${metadata.hari}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px; text-align: center;">${metadata.jam}</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px; color: green; text-align: center;">✅ Aktif</td>`;
    html += `<td style="border: 1px solid #ddd; padding: 8px; text-align: center;">`;
    html += `<button onclick="deleteTrigger('${triggerId}')" style="background: #dc3545; color: white; border: none; padding: 4px 8px; border-radius: 3px; cursor: pointer; font-size: 11px;">Hapus</button>`;
    html += `</td>`;
    html += '</tr>';
  });
  
  html += '</table>';
  html += '<br><p style="color: #666; font-size: 12px;">Total: <strong>' + summaryTriggers.length + '</strong> trigger(s)</p>';
  html += '<div id="result" style="margin-top: 10px; padding: 10px; display: none;"></div>';
  html += '</div>';
  
  html += `
    <script>
      function deleteTrigger(triggerId) {
        const resultDiv = document.getElementById('result');
        resultDiv.style.display = 'block';
        resultDiv.innerHTML = '⏳ Menghapus trigger...';
        resultDiv.style.background = '#fff3cd';
        resultDiv.style.color = '#856404';
        
        google.script.run
          .withSuccessHandler(function(result) {
            resultDiv.innerHTML = '✅ ' + result;
            resultDiv.style.background = '#d4edda';
            resultDiv.style.color = '#155724';
            setTimeout(function() {
              google.script.host.close();
            }, 1500);
          })
          .withFailureHandler(function(error) {
            resultDiv.innerHTML = '❌ Error: ' + error;
            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
          })
          .deleteTriggerById(triggerId);
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(650)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Active Triggers');
}

function getTriggerInfo(trigger) {
  let hari = 'Setiap hari';
  let jam = '-';
  
  try {
    // Get trigger handler as JSON-like string for parsing
    const handlerFunction = trigger.getHandlerFunction();
    
    // Check if it's a weekday trigger by testing each day
    const days = ['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY'];
    const dayNames = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
    
    // Try to access internal properties via workaround
    const triggerUid = trigger.getUniqueId();
    
    // Get all project triggers and find this one
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let t of allTriggers) {
      if (t.getUniqueId() === triggerUid) {
        // Access trigger properties through toString or other methods
        const triggerStr = JSON.stringify(t);
        
        // Parse the trigger to get schedule info
        // This is a workaround since Apps Script doesn't expose these properties directly
        
        break;
      }
    }
    
  } catch (e) {
    Logger.log('Error getting trigger info: ' + e);
  }
  
  return { hari: hari, jam: jam };
}

function deleteTriggerById(triggerId) {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = false;
  
  triggers.forEach(trigger => {
    if (trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
      deleteTriggerMetadata(triggerId); // Hapus metadata juga
      deleted = true;
    }
  });
  
  if (deleted) {
    return 'Trigger berhasil dihapus!';
  } else {
    throw new Error('Trigger tidak ditemukan');
  }
}

// ===== WEEKDAY TRIGGERS =====
function showWeekdayTriggerDialog() {
  const html = `
    <div style="font-family: Arial; padding: 20px;">
      <h3>📅 Set Weekday Triggers (3x Daily)</h3>
      <p>Script akan otomatis dijalankan <strong>3 kali sehari</strong> pada <strong>hari kerja</strong> (Senin-Jumat).</p>
      <hr>
      
      <h4>Pilih Mode:</h4>
      <input type="radio" id="custom" name="mode" value="custom" checked>
      <label for="custom"><strong>Custom Time</strong> - Pilih 3 jam sendiri</label><br><br>
      
      <input type="radio" id="interval" name="mode" value="interval">
      <label for="interval"><strong>Interval</strong> - Otomatis setiap 8 jam (08:00, 16:00, 00:00)</label><br><br>
      
      <hr>
      <div id="customTime">
        <h4>Pilih 3 Jam untuk Trigger:</h4>
        <label for="hour1">Trigger 1:</label>
        <select id="hour1" style="width: 100%; padding: 8px; margin: 5px 0;">
          <option value="8" selected>08:00 (8 AM)</option>
          <option value="0">00:00 (Midnight)</option>
          <option value="1">01:00 (1 AM)</option>
          <option value="2">02:00 (2 AM)</option>
          <option value="3">03:00 (3 AM)</option>
          <option value="4">04:00 (4 AM)</option>
          <option value="5">05:00 (5 AM)</option>
          <option value="6">06:00 (6 AM)</option>
          <option value="7">07:00 (7 AM)</option>
          <option value="9">09:00 (9 AM)</option>
          <option value="10">10:00 (10 AM)</option>
          <option value="11">11:00 (11 AM)</option>
          <option value="12">12:00 (12 PM)</option>
          <option value="13">13:00 (1 PM)</option>
          <option value="14">14:00 (2 PM)</option>
          <option value="15">15:00 (3 PM)</option>
          <option value="16">16:00 (4 PM)</option>
          <option value="17">17:00 (5 PM)</option>
          <option value="18">18:00 (6 PM)</option>
          <option value="19">19:00 (7 PM)</option>
          <option value="20">20:00 (8 PM)</option>
          <option value="21">21:00 (9 PM)</option>
          <option value="22">22:00 (10 PM)</option>
          <option value="23">23:00 (11 PM)</option>
        </select>
        
        <label for="hour2" style="margin-top: 10px; display: block;">Trigger 2:</label>
        <select id="hour2" style="width: 100%; padding: 8px; margin: 5px 0;">
          <option value="14" selected>14:00 (2 PM)</option>
          <option value="0">00:00 (Midnight)</option>
          <option value="1">01:00 (1 AM)</option>
          <option value="2">02:00 (2 AM)</option>
          <option value="3">03:00 (3 AM)</option>
          <option value="4">04:00 (4 AM)</option>
          <option value="5">05:00 (5 AM)</option>
          <option value="6">06:00 (6 AM)</option>
          <option value="7">07:00 (7 AM)</option>
          <option value="8">08:00 (8 AM)</option>
          <option value="9">09:00 (9 AM)</option>
          <option value="10">10:00 (10 AM)</option>
          <option value="11">11:00 (11 AM)</option>
          <option value="12">12:00 (12 PM)</option>
          <option value="13">13:00 (1 PM)</option>
          <option value="15">15:00 (3 PM)</option>
          <option value="16">16:00 (4 PM)</option>
          <option value="17">17:00 (5 PM)</option>
          <option value="18">18:00 (6 PM)</option>
          <option value="19">19:00 (7 PM)</option>
          <option value="20">20:00 (8 PM)</option>
          <option value="21">21:00 (9 PM)</option>
          <option value="22">22:00 (10 PM)</option>
          <option value="23">23:00 (11 PM)</option>
        </select>
        
        <label for="hour3" style="margin-top: 10px; display: block;">Trigger 3:</label>
        <select id="hour3" style="width: 100%; padding: 8px; margin: 5px 0;">
          <option value="20" selected>20:00 (8 PM)</option>
          <option value="0">00:00 (Midnight)</option>
          <option value="1">01:00 (1 AM)</option>
          <option value="2">02:00 (2 AM)</option>
          <option value="3">03:00 (3 AM)</option>
          <option value="4">04:00 (4 AM)</option>
          <option value="5">05:00 (5 AM)</option>
          <option value="6">06:00 (6 AM)</option>
          <option value="7">07:00 (7 AM)</option>
          <option value="8">08:00 (8 AM)</option>
          <option value="9">09:00 (9 AM)</option>
          <option value="10">10:00 (10 AM)</option>
          <option value="11">11:00 (11 AM)</option>
          <option value="12">12:00 (12 PM)</option>
          <option value="13">13:00 (1 PM)</option>
          <option value="14">14:00 (2 PM)</option>
          <option value="15">15:00 (3 PM)</option>
          <option value="16">16:00 (4 PM)</option>
          <option value="17">17:00 (5 PM)</option>
          <option value="18">18:00 (6 PM)</option>
          <option value="19">19:00 (7 PM)</option>
          <option value="21">21:00 (9 PM)</option>
          <option value="22">22:00 (10 PM)</option>
          <option value="23">23:00 (11 PM)</option>
        </select>
      </div>
      
      <br><br>
      <button onclick="setWeekdayTriggers()" style="background: #4285F4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">
        ✅ Set Triggers
      </button>
      <button onclick="google.script.host.close()" style="background: #ccc; color: black; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-left: 10px;">
        Cancel
      </button>
      <div id="result" style="margin-top: 20px; padding: 10px; display: none;"></div>
    </div>
    <script>
      function setWeekdayTriggers() {
        const mode = document.querySelector('input[name="mode"]:checked').value;
        const resultDiv = document.getElementById('result');
        resultDiv.style.display = 'block';
        resultDiv.innerHTML = '⏳ Setting triggers...';
        resultDiv.style.background = '#fff3cd';
        resultDiv.style.color = '#856404';
        
        let hours = [];
        if (mode === 'custom') {
          hours = [
            parseInt(document.getElementById('hour1').value),
            parseInt(document.getElementById('hour2').value),
            parseInt(document.getElementById('hour3').value)
          ];
        } else {
          hours = [8, 16, 0]; // Interval mode: 08:00, 16:00, 00:00
        }
        
        google.script.run
          .withSuccessHandler(function(result) {
            resultDiv.innerHTML = '✅ ' + result;
            resultDiv.style.background = '#d4edda';
            resultDiv.style.color = '#155724';
            setTimeout(function() {
              google.script.host.close();
            }, 2000);
          })
          .withFailureHandler(function(error) {
            resultDiv.innerHTML = '❌ Error: ' + error;
            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
          })
          .createWeekdayTriggers(hours);
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Set Weekday Triggers');
}

function createWeekdayTriggers(hours) {
  // JANGAN hapus trigger existing - hanya tambah trigger baru
  
  const days = [
    { day: ScriptApp.WeekDay.MONDAY, name: 'Senin' },
    { day: ScriptApp.WeekDay.TUESDAY, name: 'Selasa' },
    { day: ScriptApp.WeekDay.WEDNESDAY, name: 'Rabu' },
    { day: ScriptApp.WeekDay.THURSDAY, name: 'Kamis' },
    { day: ScriptApp.WeekDay.FRIDAY, name: 'Jumat' }
  ];
  
  // Buat trigger untuk setiap hari kerja dan jam
  hours.forEach(hour => {
    days.forEach(dayInfo => {
      const trigger = ScriptApp.newTrigger('generateSummary')
        .timeBased()
        .atHour(hour)
        .onWeekDay(dayInfo.day)
        .create();
      
      // Simpan metadata trigger
      saveTriggerMetadata(trigger.getUniqueId(), dayInfo.name, String(hour).padStart(2, '0') + ':00');
    });
  });
  
  const hourStr = hours.map(h => String(h).padStart(2, '0') + ':00').join(', ');
  return `Berhasil! Script akan berjalan 3x sehari pada hari kerja (Senin-Jumat) di jam: ${hourStr}`;
}

// ===== TRIGGER MANAGEMENT =====
function showTriggerDialog() {
  const html = `
    <div style="font-family: Arial; padding: 20px;">
      <h3>⏰ Set Daily Trigger</h3>
      <p>Script akan otomatis dijalankan setiap hari pada jam yang Anda tentukan.</p>
      <hr>
      <label for="hour"><strong>Pilih Jam:</strong></label><br>
      <select id="hour" style="width: 100%; padding: 8px; margin: 10px 0; font-size: 14px;">
        <option value="2">02:00 (2 AM) - Default</option>
        <option value="0">00:00 (Midnight)</option>
        <option value="1">01:00 (1 AM)</option>
        <option value="3">03:00 (3 AM)</option>
        <option value="4">04:00 (4 AM)</option>
        <option value="5">05:00 (5 AM)</option>
        <option value="6">06:00 (6 AM)</option>
        <option value="7">07:00 (7 AM)</option>
        <option value="8">08:00 (8 AM)</option>
        <option value="9">09:00 (9 AM)</option>
        <option value="10">10:00 (10 AM)</option>
        <option value="11">11:00 (11 AM)</option>
        <option value="12">12:00 (12 PM)</option>
        <option value="13">13:00 (1 PM)</option>
        <option value="14">14:00 (2 PM)</option>
        <option value="15">15:00 (3 PM)</option>
        <option value="16">16:00 (4 PM)</option>
        <option value="17">17:00 (5 PM)</option>
        <option value="18">18:00 (6 PM)</option>
        <option value="19">19:00 (7 PM)</option>
        <option value="20">20:00 (8 PM)</option>
        <option value="21">21:00 (9 PM)</option>
        <option value="22">22:00 (10 PM)</option>
        <option value="23">23:00 (11 PM)</option>
      </select>
      <br><br>
      <button onclick="setTrigger()" style="background: #4285F4; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px;">
        ✅ Set Trigger
      </button>
      <button onclick="google.script.host.close()" style="background: #ccc; color: black; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; margin-left: 10px;">
        Cancel
      </button>
      <div id="result" style="margin-top: 20px; padding: 10px; display: none;"></div>
    </div>
    <script>
      function setTrigger() {
        const hour = document.getElementById('hour').value;
        const resultDiv = document.getElementById('result');
        resultDiv.style.display = 'block';
        resultDiv.innerHTML = '⏳ Setting trigger...';
        resultDiv.style.background = '#fff3cd';
        resultDiv.style.color = '#856404';
        
        google.script.run
          .withSuccessHandler(function(result) {
            resultDiv.innerHTML = '✅ ' + result;
            resultDiv.style.background = '#d4edda';
            resultDiv.style.color = '#155724';
            setTimeout(function() {
              google.script.host.close();
            }, 2000);
          })
          .withFailureHandler(function(error) {
            resultDiv.innerHTML = '❌ Error: ' + error;
            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
          })
          .createDailyTrigger(parseInt(hour));
      }
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Set Daily Trigger');
}

function createDailyTrigger(hour) {
  // JANGAN hapus trigger existing - hanya tambah trigger baru
  
  // Buat trigger baru
  const trigger = ScriptApp.newTrigger('generateSummary')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .create();
  
  // Simpan metadata trigger
  saveTriggerMetadata(trigger.getUniqueId(), 'Setiap hari', String(hour).padStart(2, '0') + ':00');
  
  return `Trigger berhasil dibuat! Script akan berjalan setiap hari jam ${String(hour).padStart(2, '0')}:00`;
}

function removeAllTriggers() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Remove All Triggers',
    'Apakah Anda yakin ingin menghapus semua trigger otomatis?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'generateSummary') {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
    });
    
    if (count > 0) {
      ui.alert(`✅ ${count} trigger berhasil dihapus.`);
    } else {
      ui.alert('ℹ️ Tidak ada trigger yang ditemukan.');
    }
  }
}
