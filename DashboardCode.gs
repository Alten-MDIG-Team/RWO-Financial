// ==========================================
// CONFIGURATION DU FICHIER EXTERNE
// ==========================================
const RWO_LOG_URL = "https://docs.google.com/spreadsheets/d/1hvrnM1uHMAdHqb0YONetIOtl0_PWUxWJCmVs20GJO2s/edit"; 
// ==========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('DashboardIndex')
      .setTitle('RWO - KPI Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getFilterOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DashBoard");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let colExec = -1;
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] && headers[c].toString().trim() === "Execution") {
        colExec = c;
        break;
      }
    }
    
    if (colExec === -1) return []; 
    
    let userMap = {};
    for (let i = 1; i < data.length; i++) {
      let exec = data[i][colExec];
      if (exec) {
        let val = exec.toString().trim(); 
        if (val !== '') userMap[val] = true;
      }
    }
    return Object.keys(userMap).sort();
  } catch (e) {
    return [];
  }
}

function getIsoWeek(date) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) {
    target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
  }
  return 1 + Math.ceil((firstThursday - target) / 604800000);
}

function getDashboardMetrics(filterType, filterValues, priorityFilter, foboFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- 0. SMART SCANNER: PARAMETERS (FINANCIAL / WU RATE) ---
  let gammePrices = {};
  const sheetParams = ss.getSheetByName("Parameters");
  if (sheetParams) {
    const pData = sheetParams.getDataRange().getValues();
    if (pData.length > 1) {
      const pHeaders = pData[0];
      let pColType = 0; 
      let pColMoney = 1; 
      
      for (let c = 0; c < pHeaders.length; c++) {
        let h = (pHeaders[c] || '').toString().toLowerCase();
        if (h.includes('livrable') || h.includes('gamme') || h.includes('type')) pColType = c;
        if (h.includes('rate') || h.includes('wu') || h.includes('prix') || h.includes('cost') || h.includes('$') || h.includes('€')) pColMoney = c;
      }
      
      for (let i = 1; i < pData.length; i++) {
        let gName = pData[i][pColType];
        let rawPrice = pData[i][pColMoney];
        let gPrice = 0;
        
        if (typeof rawPrice === 'number') {
          gPrice = rawPrice;
        } else if (typeof rawPrice === 'string') {
          gPrice = parseFloat(rawPrice.replace(/[^0-9.-]+/g, ""));
        }

        if (gName && !isNaN(gPrice)) {
          gammePrices[gName.toString().trim()] = gPrice;
        }
      }
    }
  }

  const sheetDb = ss.getSheetByName("DashBoard");
  if (!sheetDb) throw new Error("DashBoard tab not found.");

  const dbData = sheetDb.getDataRange().getValues();
  const headers = dbData[0];

  function getCol(name) {
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] && headers[c].toString().trim() === name) return c;
    }
    return -1;
  }

  const colStatus = getCol("Status");
  const colDate = getCol("Delivery Date");
  const colOTD = getCol("OTD");
  const colRFT = getCol("RFT");
  const colProdTarget = getCol("Production Time Target (Hours)");
  const colProdActual = getCol("Production Time (Hours)");
  const colCCTarget = getCol("Crosscheck Time Target (Hours)");
  const colCCActual = getCol("Crosscheck Time (Hours)");
  const colPriority = getCol("Priority");
  const colType = getCol("Type of deliverable");
  const colLeadTime = getCol("Lead Time (Hours)");
  const colResp = getCol("Responsible");
  const colRFTi = getCol("RFTi");
  const colRework = getCol("Rework Time (Hours)");
  const colExecution = getCol("Execution");
  const colBackOffice = getCol("Back Office"); 

  let metrics = {
    otd: { trueCount: 0, falseCount: 0 },
    rft: { trueCount: 0, falseCount: 0 },
    efficiency: { targetHrs: 0, actualHrs: 0 },
    project: { p1: 0, p2: 0, p3: 0, leadTimeSum: 0, total: 0, types: {} },
    performance: { bo: 0, fo: 0, rftiTrue: 0, rftiFalse: 0, reworkHrs: 0 },
    individuals: {}
  };

  // Accumulators for TRM (Weekly) and PRM (Monthly) across 2026
  let weeklyMap = {};
  let monthlyMap = {};

  const today = new Date(); 

  // --- 1. KPI DATA LOOP ---
  for (let i = 1; i < dbData.length; i++) {
    const row = dbData[i];
    if (row[colStatus] !== "Delivered") continue;

    const deliveryDate = row[colDate];
    if (!deliveryDate || !(deliveryDate instanceof Date)) continue;
    if (deliveryDate.getFullYear() !== 2026) continue;

    // Apply Priority and FO/BO filters across all views (including TRM & PRM)
    if (priorityFilter && priorityFilter !== 'All') {
      if (row[colPriority] != priorityFilter) continue;
    }

    if (foboFilter && foboFilter !== 'All') {
      const isBO = row[colBackOffice] === true;
      if (foboFilter === 'BO' && !isBO) continue;
      if (foboFilter === 'FO' && isBO) continue;
    }

    const isOTD = row[colOTD] === true;
    const isNotOTD = row[colOTD] === false;
    const isRFT = row[colRFT] === true;
    const isNotRFT = row[colRFT] === false;

    let targetHrs = (parseFloat(row[colProdTarget]) || 0) + (parseFloat(row[colCCTarget]) || 0);
    let actualHrs = (parseFloat(row[colProdActual]) || 0) + (parseFloat(row[colCCActual]) || 0);

    let t = row[colType];
    let typeName = (t && t.toString().trim() !== '') ? t.toString().trim() : "Non défini";

    // --- Populate TRM (Weekly) & PRM (Monthly) timelines for 2026 ---
    let wk = getIsoWeek(deliveryDate);
    if (!weeklyMap[wk]) {
      weeklyMap[wk] = { week: wk, total: 0, types: {}, otdTrue: 0, otdFalse: 0, rftTrue: 0, rftFalse: 0, targetHrs: 0, actualHrs: 0 };
    }
    weeklyMap[wk].total++;
    weeklyMap[wk].types[typeName] = (weeklyMap[wk].types[typeName] || 0) + 1;
    if (isOTD) weeklyMap[wk].otdTrue++;
    else if (isNotOTD) weeklyMap[wk].otdFalse++;
    if (isRFT) weeklyMap[wk].rftTrue++;
    else if (isNotRFT) weeklyMap[wk].rftFalse++;
    weeklyMap[wk].targetHrs += targetHrs;
    weeklyMap[wk].actualHrs += actualHrs;

    let mo = deliveryDate.getMonth() + 1;
    if (!monthlyMap[mo]) {
      monthlyMap[mo] = { month: mo, total: 0, types: {}, otdTrue: 0, otdFalse: 0, rftTrue: 0, rftFalse: 0, targetHrs: 0, actualHrs: 0 };
    }
    monthlyMap[mo].total++;
    monthlyMap[mo].types[typeName] = (monthlyMap[mo].types[typeName] || 0) + 1;
    if (isOTD) monthlyMap[mo].otdTrue++;
    else if (isNotOTD) monthlyMap[mo].otdFalse++;
    if (isRFT) monthlyMap[mo].rftTrue++;
    else if (isNotRFT) monthlyMap[mo].rftFalse++;
    monthlyMap[mo].targetHrs += targetHrs;
    monthlyMap[mo].actualHrs += actualHrs;

    // --- Apply Period Filter for Résumé / Suivi Projet / Suivi Performance ---
    let includeRow = false;
    if (filterType === 'Global') {
      includeRow = true;
    } else if (filterType === 'Mois') {
      if (filterValues === 'D-31') {
        const d31 = new Date(today.getTime() - (31 * 24 * 60 * 60 * 1000));
        includeRow = (deliveryDate >= d31 && deliveryDate <= today);
      } else if (Array.isArray(filterValues) && filterValues.length > 0) {
        includeRow = filterValues.indexOf(mo) !== -1;
      }
    } else if (filterType === 'Semaine') {
      if (filterValues === 'D-7') {
        const d7 = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
        includeRow = (deliveryDate >= d7 && deliveryDate <= today);
      } else if (Array.isArray(filterValues) && filterValues.length > 0) {
        includeRow = filterValues.indexOf(wk) !== -1;
      }
    }
    if (!includeRow) continue;

    if (isOTD) metrics.otd.trueCount++;
    else if (isNotOTD) metrics.otd.falseCount++;

    if (isRFT) metrics.rft.trueCount++;
    else if (isNotRFT) metrics.rft.falseCount++;

    metrics.efficiency.targetHrs += targetHrs;
    metrics.efficiency.actualHrs += actualHrs;

    metrics.project.total++;
    metrics.project.leadTimeSum += (parseFloat(row[colLeadTime]) || 0);
    
    let prio = row[colPriority];
    if (prio == 1) metrics.project.p1++;
    else if (prio == 2) metrics.project.p2++;
    else if (prio == 3) metrics.project.p3++;

    if (!metrics.project.types[typeName]) {
      metrics.project.types[typeName] = { count: 0, target: 0, actual: 0, pricePerUnit: (gammePrices[typeName] || 0) };
    }
    metrics.project.types[typeName].count++;
    metrics.project.types[typeName].target += targetHrs;
    metrics.project.types[typeName].actual += actualHrs;

    let resp = row[colResp];
    if (resp === "BO") metrics.performance.bo++;
    else if (resp === "FO") metrics.performance.fo++;

    if (row[colRFTi] === true) metrics.performance.rftiTrue++;
    else if (row[colRFTi] === false) metrics.performance.rftiFalse++;

    metrics.performance.reworkHrs += (parseFloat(row[colRework]) || 0);

    let execUser = row[colExecution];
    if (execUser && typeof execUser === 'string' && execUser.trim() !== '') {
      let u = execUser.trim();
      if (!metrics.individuals[u]) {
        metrics.individuals[u] = { volume: 0, otdTrue: 0, otdFalse: 0, rftTrue: 0, rftFalse: 0 };
      }
      metrics.individuals[u].volume++;
      if (isOTD) metrics.individuals[u].otdTrue++;
      else if (isNotOTD) metrics.individuals[u].otdFalse++;
      
      if (isRFT) metrics.individuals[u].rftTrue++;
      else if (isNotRFT) metrics.individuals[u].rftFalse++;
    }
  }

  // --- 2. BO SHIFT HOURLY GRAPH EXTRACTION ---
  let shiftData = [];
  let hourlyData = [];
  for (let h = 0; h < 24; h++) { hourlyData.push({ count: 0, names: [] }); }

  let sheetShift = ss.getSheetByName("BO Shift");
  if (!sheetShift && RWO_LOG_URL.trim() !== "") {
    try {
      const extSs = SpreadsheetApp.openByUrl(RWO_LOG_URL.trim());
      sheetShift = extSs.getSheetByName("BO Shift");
    } catch (e) {}
  }

  if (sheetShift) {
    const sData = sheetShift.getDataRange().getValues();
    if (sData.length > 4) {
      let targetRow = 4; 
      const todayNoTime = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
      
      for (let r = 4; r < Math.min(sData.length, 100); r++) {
        let cell = sData[r][1]; 
        if (cell instanceof Date) {
          let hDate = new Date(cell.getFullYear(), cell.getMonth(), cell.getDate()).getTime();
          if (hDate === todayNoTime) { targetRow = r; break; }
        }
      }
      
      function getHourFromCell(cVal) {
        if (cVal instanceof Date) return cVal.getHours();
        if (typeof cVal === 'string') {
          let pts = cVal.split(':');
          if (pts.length > 0 && !isNaN(pts[0])) return parseInt(pts[0], 10);
        }
        if (typeof cVal === 'number') return Math.floor(cVal * 24);
        return "?";
      }

      const headers = sData[0];
      for (let c = 2; c < headers.length; c++) {
        let shiftName = headers[c];
        if (!shiftName || shiftName.toString().trim() === '') continue;

        let startCell = sData[1][c];
        let endCell = sData[2][c];   
        let namesCell = sData[targetRow][c];
        
        if (namesCell && namesCell.toString().trim() !== '') {
          let people = namesCell.toString().split(/[-,\n]+/).map(p => p.trim()).filter(p => p.length > 0);
          
          let startHour = "?", endHour = "?";
          let activeHours = [];
          
          if (startCell !== '' && endCell !== '') {
            startHour = getHourFromCell(startCell);
            endHour = getHourFromCell(endCell);
            if (startHour !== "?" && endHour !== "?") {
              let tempEnd = endHour === 0 ? 24 : endHour;
              if (tempEnd <= startHour) { 
                for (let h = startHour; h < 24; h++) activeHours.push(h);
                for (let h = 0; h < tempEnd; h++) activeHours.push(h);
              } else {
                for (let h = startHour; h < tempEnd; h++) activeHours.push(h);
              }
            }
          }

          people.forEach(p => {
            shiftData.push({ name: p, shift: shiftName.toString().trim(), start: startHour, end: endHour });
            activeHours.forEach(h => {
              hourlyData[h].count++;
              hourlyData[h].names.push(p);
            });
          });
        }
      }
    }
  }

  // --- 3. NRFT DETAILS SMART SCANNER ---
  let nrftList = [];
  const sheetNrft = ss.getSheetByName("NRFT Tracker");
  if (sheetNrft) {
    const nData = sheetNrft.getDataRange().getValues();
    if (nData.length > 1) {
      const nHeaders = nData[0];
      
      const nColDate = nHeaders.findIndex(h => h && h.toString().toLowerCase().includes('date'));
      const nColId = nHeaders.findIndex(h => h && (h.toString().toLowerCase().includes('id') || h.toString().toLowerCase().includes('msn') || h.toString().toLowerCase().includes('ref')));
      const nColCause = nHeaders.findIndex(h => h && (h.toString().toLowerCase().includes('cause') || h.toString().toLowerCase().includes('root') || h.toString().toLowerCase().includes('motif') || h.toString().toLowerCase().includes('category')));
      const nColDesc = nHeaders.findIndex(h => h && (h.toString().toLowerCase().includes('desc') || h.toString().toLowerCase().includes('comment') || h.toString().toLowerCase().includes('remarque') || h.toString().toLowerCase().includes('detail')));
      
      for (let i = 1; i < nData.length; i++) {
        let row = nData[i];
        if (!row.some(String)) continue; 
        
        let rawDate = nColDate > -1 ? row[nColDate] : null;
        let dDate = null;
        if (rawDate instanceof Date) dDate = rawDate;
        else if (typeof rawDate === 'string' || typeof rawDate === 'number') dDate = new Date(rawDate);
        
        let includeRow = false;
        if (filterType === 'Global') {
          includeRow = true;
        } else if (dDate && !isNaN(dDate.getTime())) {
          if (filterType === 'Mois') {
            if (filterValues === 'D-31') {
              const d31 = new Date(today.getTime() - (31 * 24 * 60 * 60 * 1000));
              includeRow = (dDate >= d31 && dDate <= today);
            } else if (Array.isArray(filterValues) && filterValues.length > 0) {
              includeRow = filterValues.indexOf(dDate.getMonth() + 1) !== -1;
            }
          } else if (filterType === 'Semaine') {
            if (filterValues === 'D-7') {
              const d7 = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
              includeRow = (dDate >= d7 && dDate <= today);
            } else if (Array.isArray(filterValues) && filterValues.length > 0) {
              includeRow = filterValues.indexOf(getIsoWeek(dDate)) !== -1;
            }
          }
        }
        if (!includeRow) continue;
        
        let formattedDate = 'N/A';
        if (dDate && !isNaN(dDate.getTime())) {
          formattedDate = ("0" + dDate.getDate()).slice(-2) + "/" + ("0" + (dDate.getMonth() + 1)).slice(-2) + "/" + dDate.getFullYear();
        } else if (rawDate) {
          formattedDate = rawDate.toString();
        }

        nrftList.push({
          date: formattedDate,
          id: nColId > -1 ? (row[nColId] || 'N/A').toString() : 'N/A',
          cause: nColCause > -1 ? (row[nColCause] || 'N/A').toString() : 'N/A',
          desc: nColDesc > -1 ? (row[nColDesc] || 'N/A').toString() : 'N/A'
        });
      }
    }
  }

  // --- 4. RWO-CAPA PROJECTION EXTRACTOR ---
  let capaData = [];
  let capaByWeek = {};
  let capaByMonth = {};
  const sheetCapa = ss.getSheetByName("RWO-CAPA");
  if (sheetCapa) {
    const cData = sheetCapa.getDataRange().getValues();
    if (cData.length >= 11) {
      let months = cData[2];       // Row 3
      let weeks = cData[3];        // Row 4
      let vacation = cData[5];     // Row 6
      let global = cData[8];       // Row 9
      let meeting = cData[9];      // Row 10
      let operational = cData[10]; // Row 11

      for (let c = 1; c < weeks.length; c++) {
        let w = parseInt(weeks[c], 10);
        let m = parseInt(months[c], 10);
        if (!isNaN(w)) {
          let item = {
            week: w,
            month: !isNaN(m) ? m : null,
            vacation: parseFloat(vacation[c]) || 0,
            global: parseFloat(global[c]) || 0,
            meeting: parseFloat(meeting[c]) || 0,
            operational: parseFloat(operational[c]) || 0
          };
          capaData.push(item);
          capaByWeek[w] = item;
          if (!isNaN(m)) {
            if (!capaByMonth[m]) capaByMonth[m] = { operational: 0, global: 0 };
            capaByMonth[m].operational += item.operational;
            capaByMonth[m].global += item.global;
          }
        }
      }
    }
  }

  // Build sorted TRM (Weekly) array
  let trmArray = Object.keys(weeklyMap).map(Number).sort((a, b) => a - b).map(w => {
    let item = weeklyMap[w];
    let otdTot = item.otdTrue + item.otdFalse;
    let rftTot = item.rftTrue + item.rftFalse;
    let cInfo = capaByWeek[w];
    return {
      week: w,
      total: item.total,
      types: item.types,
      otdTrue: item.otdTrue,
      otdFalse: item.otdFalse,
      otdPct: otdTot > 0 ? (item.otdTrue / otdTot) * 100 : 0,
      rftTrue: item.rftTrue,
      rftFalse: item.rftFalse,
      rftPct: rftTot > 0 ? (item.rftTrue / rftTot) * 100 : 0,
      targetHrs: item.targetHrs,
      actualHrs: item.actualHrs,
      capaOp: cInfo ? cInfo.operational : null
    };
  });

  // Build sorted PRM (Monthly) array
  let prmArray = Object.keys(monthlyMap).map(Number).sort((a, b) => a - b).map(m => {
    let item = monthlyMap[m];
    let otdTot = item.otdTrue + item.otdFalse;
    let rftTot = item.rftTrue + item.rftFalse;
    let cInfo = capaByMonth[m];
    return {
      month: m,
      total: item.total,
      types: item.types,
      otdTrue: item.otdTrue,
      otdFalse: item.otdFalse,
      otdPct: otdTot > 0 ? (item.otdTrue / otdTot) * 100 : 0,
      rftTrue: item.rftTrue,
      rftFalse: item.rftFalse,
      rftPct: rftTot > 0 ? (item.rftTrue / rftTot) * 100 : 0,
      targetHrs: item.targetHrs,
      actualHrs: item.actualHrs,
      capaOp: cInfo ? cInfo.operational : null,
      // chiffre d'affaires estimé du mois = livrables × prix unitaire (onglet Parameters)
      value: Object.keys(item.types).reduce((s, t) => s + item.types[t] * (gammePrices[t] || 0), 0)
    };
  });

  const otdTotal = metrics.otd.trueCount + metrics.otd.falseCount;
  const rftTotal = metrics.rft.trueCount + metrics.rft.falseCount;
  let effPercent = metrics.efficiency.actualHrs > 0 ? (metrics.efficiency.targetHrs / metrics.efficiency.actualHrs) * 100 : 0;
  const rftiTotal = metrics.performance.rftiTrue + metrics.performance.rftiFalse;

  let indArray = [];
  for (let key in metrics.individuals) {
    let stat = metrics.individuals[key];
    let userOtdTotal = stat.otdTrue + stat.otdFalse;
    let userRftTotal = stat.rftTrue + stat.rftFalse;
    indArray.push({
      name: key,
      volume: stat.volume,
      otdPct: userOtdTotal > 0 ? (stat.otdTrue / userOtdTotal) * 100 : 0,
      rftPct: userRftTotal > 0 ? (stat.rftTrue / userRftTotal) * 100 : 0
    });
  }
  indArray.sort((a, b) => b.volume - a.volume);

  return {
    otd: { achieved: metrics.otd.trueCount, missed: metrics.otd.falseCount, total: otdTotal, percent: otdTotal > 0 ? (metrics.otd.trueCount / otdTotal * 100) : 0, targetPercent: 97 },
    rft: { achieved: metrics.rft.trueCount, missed: metrics.rft.falseCount, total: rftTotal, percent: rftTotal > 0 ? (metrics.rft.trueCount / rftTotal * 100) : 0, targetPercent: 97 },
    efficiency: { percent: effPercent, targetPercent: 80, actualHrs: metrics.efficiency.actualHrs, targetHrs: metrics.efficiency.targetHrs },
    project: { total: metrics.project.total, avgLeadTime: metrics.project.total > 0 ? (metrics.project.leadTimeSum / metrics.project.total) : 0, p1: metrics.project.p1, p2: metrics.project.p2, p3: metrics.project.p3, types: metrics.project.types },
    performance: { bo: metrics.performance.bo, fo: metrics.performance.fo, rftiAchieved: metrics.performance.rftiTrue, rftiMissed: metrics.performance.rftiFalse, rftiTotal: rftiTotal, rftiPercent: rftiTotal > 0 ? (metrics.performance.rftiTrue / rftiTotal * 100) : 0, reworkHrs: metrics.performance.reworkHrs },
    individuals: indArray,
    shifts: shiftData,
    hourlyData: hourlyData,
    nrftList: nrftList,
    capaData: capaData,
    trm: trmArray,
    prm: prmArray
  };
}
