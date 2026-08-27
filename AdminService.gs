// =============================================================================
// ADMIN SERVICES
// =============================================================================

function getAdminData(f) {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    let list = [];
    
    for (let i = 1; i < data.length; i++) {
      let r = data[i];
      let rowSem = r[10] ? r[10].toString().trim() : "";
      let rowYr  = r[11] ? r[11].toString().trim() : "";
      let rowLev = r[7]  ? r[7].toString().trim()  : "";
      let rowRom = r[8]  ? r[8].toString().trim()  : "";
      
      let matchSem = (f.sem === "all" || rowSem === f.sem);
      let matchYr  = (f.yr  === "all" || rowYr  === f.yr);
      let matchLev = (f.lev === "all" || rowLev === f.lev);
      let matchRom = (f.rom === "all" || rowRom === f.rom);
      
      if (matchSem && matchYr && matchLev && matchRom) {
        let valF1 = r[16];
        let displayF1 = (valF1 !== undefined && valF1 !== null && valF1 !== "") ? valF1.toString() : "-";

        list.push({
          id: r[1],
          name: `${r[2]}${r[3]} ${r[4]}`,
          lr: `${r[7]}/${r[8]}`,
          no: r[9],
          sem: r[10],
          yr: r[11],
          bmi: r[14], bmiL: r[15],
          f1: displayF1, f1L: r[17],
          f2: r[18], f2L: r[19],
          f3: r[20], f3L: r[21],
          f4: r[22], f4L: r[23],
          total: r[29], rank: r[30]
        });
      }
    }
    return list;
  } catch(e) { return []; }
}

function getPendingStudents(f) {
  try {
    const db = getDb();
    const studentSheet = db.getSheetByName("Student_DB");
    const fitnessSheet = db.getSheetByName("Fitness_Logs");
    
    if (!studentSheet) return [];
    
    const studentData = studentSheet.getDataRange().getValues();
    const fitnessData = fitnessSheet ? fitnessSheet.getDataRange().getValues() : [];
    
    let testedMap = {};
    for (let j = 1; j < fitnessData.length; j++) {
      let sId = fitnessData[j][1] ? fitnessData[j][1].toString().trim() : "";
      let sSem = fitnessData[j][10] ? fitnessData[j][10].toString().trim() : "";
      let sYr = fitnessData[j][11] ? fitnessData[j][11].toString().trim() : "";
      let key = `${sId}_${sSem}_${sYr}`;
      testedMap[key] = true;
    }
    
    let pendingList = [];
    for (let i = 1; i < studentData.length; i++) {
      let r = studentData[i];
      let stId = r[0] ? r[0].toString().trim() : "";
      let stLev = r[6] ? r[6].toString().trim() : "";
      let stRom = r[7] ? r[7].toString().trim() : "";
      
      let matchLev = (f.lev === "all" || stLev === f.lev);
      let matchRom = (f.rom === "all" || stRom === f.rom);
      
      if (matchLev && matchRom && stId !== "") {
        let keyToCheck = `${stId}_${f.sem}_${f.yr}`;
        if (!testedMap[keyToCheck]) {
          pendingList.push({
            id: stId,
            name: `${r[1]}${r[2]} ${r[3]}`,
            lr: `${r[6]}/${r[7]}`,
            no: r[8],
            sem: f.sem,
            yr: f.yr,
            status: "ยังไม่ได้ทดสอบ"
          });
        }
      }
    }
    
    return pendingList.sort((a, b) => (parseInt(a.no) || 0) - (parseInt(b.no) || 0));
  } catch(e) { return []; }
}