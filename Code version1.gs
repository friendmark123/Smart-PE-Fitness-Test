function testMyAnnualReport() {
  Logger.log("--- เริ่มรันฟังก์ชันทดลองระบบ ---");
  var result = generateAnnualPDF("ทั้งหมด", "1", "2569");
  Logger.log("ผลลัพธ์ที่ได้จากระบบ: " + JSON.stringify(result));
}
// =============================================================================
// CONFIGURATION & CORE SETUP
// =============================================================================
const SPREADSHEET_ID = "1rAZyKOI7-C1VHgv6EPXQ5KiD039-DLcaM4xsYEXcSaA";
const FOLDER_PDF_ID   = "1oYAw15QVtc7bc80gXU85geK1tqXHZ39n";

function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบสมรรถภาพทางกาย - บดินทรเดชา ๒')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getYearList() {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    if (!sheet) return [new Date().getFullYear() + 543];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [new Date().getFullYear() + 543];
   
    let years = [];
    for (let i = 1; i < data.length; i++) {
      let yr = data[i][11];
      if (yr && !years.includes(yr)) years.push(yr);
    }
    if (years.length === 0) years.push(new Date().getFullYear() + 543);
    return years.sort((a, b) => b - a);
  } catch(e) { return [new Date().getFullYear() + 543]; }
}

function getBmiCriteriaMatrix() {
  try {
    const sheet = getDb().getSheetByName("Criteria_BMI");
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  } catch(e) { return []; }
}

// =============================================================================
// REAL-TIME SYNC & SEARCH
// =============================================================================
function searchStudentData(studentId, semester, academicYear) {
  const db = getDb();
  const studentSheet = db.getSheetByName("Student_DB");
  const fitnessSheet = db.getSheetByName("Fitness_Logs");
  let result = { found: false, hasOldData: false };
 
  if (!studentSheet) return result;
  const studentData = studentSheet.getDataRange().getValues();
  let studentRow = null;
  for (let i = 1; i < studentData.length; i++) {
    if (studentData[i][0].toString().trim() === studentId.toString().trim()) {
      studentRow = studentData[i];
      break;
    }
  }
  if (!studentRow) return result;
 
  result.found = true;
  result.prefix = studentRow[1]; result.fname = studentRow[2]; result.lname = studentRow[3];
  result.gender = studentRow[4]; result.age = studentRow[5]; result.level = studentRow[6];
  result.room = studentRow[7]; result.no = studentRow[8];
 
  if (fitnessSheet) {
    const fitnessData = fitnessSheet.getDataRange().getValues();
    for (let j = 1; j < fitnessData.length; j++) {
      if (fitnessData[j][1].toString().trim() === studentId.toString().trim() &&
          fitnessData[j][10].toString().trim() === semester.toString().trim() &&
          fitnessData[j][11].toString().trim() === academicYear.toString().trim()) {
        result.hasOldData = true;
        result.weight = fitnessData[j][12]; result.height = fitnessData[j][13];
        
        // แก้ไขข้อ 3: ตรวจสอบค่านั่งงอตัว (f1) ถ้าราคาเป็น 0 หรือ "0" ให้ดึงค่าเป็น 0 ไม่ให้เป็นค่าว่างหรือขีด
        let valF1 = fitnessData[j][16];
        result.f1 = (valF1 === 0 || valF1 === "0") ? 0 : valF1;
        
        result.f2 = fitnessData[j][18];
        result.f3 = fitnessData[j][20];
        result.f4 = fitnessData[j][22];
        break;
      }
    }
  }
  return result;
}

function searchFitnessHistory(studentId, semester, academicYear) {
  const sheet = getDb().getSheetByName("Fitness_Logs");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let history = [];
 
  for(let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === studentId.toString().trim() &&
        (semester === "all" || data[i][10].toString().trim() === semester.toString().trim()) &&
        (academicYear === "all" || data[i][11].toString().trim() === academicYear.toString().trim())) {
      
      // แก้ไขข้อ 3: การแสดงผลการนั่งงอตัวประวัติย้อนหลัง ถ้าเป็น 0 ให้โชว์เลข 0 
      let valF1 = data[i][16];
      let displayF1 = (valF1 === 0 || valF1 === "0" || valF1 === 0.0) ? "0" : (valF1 ? data[i][16].toString() : "-");
     
      history.push({
        info: `${data[i][2]}${data[i][3]} ${data[i][4]} ชั้น ${data[i][7]}/${data[i][8]} เลขที่ ${data[i][9]}`,
        semester: data[i][10], academic_year: data[i][11], bmi: data[i][14],
        bmiL: data[i][15], bmiS: data[i][24],
        f1: displayF1, f1L: data[i][17], f1S: data[i][25],
        f2: data[i][18], f2L: data[i][19], f2S: data[i][26],
        f3: data[i][20], f3L: data[i][21], f3S: data[i][27],
        f4: data[i][22], f4L: data[i][23], f4S: data[i][28],
        total: data[i][29], rank: data[i][30]
      });
    }
  }
  return history;
}

// =============================================================================
// CORE CALCULATION ENGINE
// =============================================================================
function matchTestType(sheetValue, userValue) {
  let s = sheetValue.toString().trim();
  let u = userValue.toString().trim();
  return s === u;
}

function getBmiEvaluation(gender, age, bmiValue) {
  try {
    const sheet = getDb().getSheetByName("Criteria_BMI");
    if (!sheet) return { label: "สมส่วน", score: 5 };
    const data = sheet.getDataRange().getValues();
   
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === gender.trim() && parseInt(data[i][2]) === parseInt(age)) {
        let veryThin = parseFloat(data[i][3]);  
        let thin = parseFloat(data[i][4]);      
        let normalMin = parseFloat(data[i][5]);  
        let normalMax = parseFloat(data[i][6]);  
        let overweight = parseFloat(data[i][7]);
        let obese = parseFloat(data[i][8]);      
       
        if (bmiValue >= obese) return { label: "อ้วน", score: 1 };
        if (bmiValue >= overweight) return { label: "ท้วม", score: 3 };
        if (bmiValue >= normalMin && bmiValue <= normalMax) return { label: "สมส่วน", score: 5 };
        if (bmiValue >= thin && bmiValue < normalMin) return { label: "ผอม", score: 3 };
        return { label: "ผอมมาก", score: 1 };
      }
    }
    if (bmiValue < 18.5) return { label: "ผอม", score: 3 };
    if (bmiValue <= 22.9) return { label: "สมส่วน", score: 5 };
    if (bmiValue <= 24.9) return { label: "ท้วม", score: 3 };
    return { label: "อ้วน", score: 1 };
  } catch(e) { return { label: "สมส่วน", score: 5 }; }
}

function getFitnessEvaluation(testType, gender, age, rawValue) {
  try {
    const sheet = getDb().getSheetByName("Settings_Criteria");
    if (!sheet) return { label: "ไม่พบแผ่นงานเกณฑ์ Settings_Criteria", score: 0 };
   
    const data = sheet.getDataRange().getValues();
    let val = parseFloat(rawValue);
    if (isNaN(val)) return { label: "ไม่มีข้อมูลสถิติ", score: 0 };
   
    let searchTest = testType.toString().trim();
    let searchGender = gender.toString().trim();
    let searchAge = parseInt(age);

    for (let i = 1; i < data.length; i++) {
      let sheetTest = data[i][0].toString().trim();
      let sheetGender = data[i][1].toString().trim();
      let sheetAge = parseInt(data[i][2]);

      if (sheetTest === searchTest && sheetGender === searchGender && sheetAge === searchAge) {
        let s1 = isNaN(parseFloat(data[i][3])) ? 0 : parseFloat(data[i][3]);
        let s2 = isNaN(parseFloat(data[i][4])) ? 0 : parseFloat(data[i][4]);
        let s3 = isNaN(parseFloat(data[i][5])) ? 0 : parseFloat(data[i][5]);
        let s4 = isNaN(parseFloat(data[i][6])) ? 0 : parseFloat(data[i][6]);
        let s5 = isNaN(parseFloat(data[i][7])) ? 0 : parseFloat(data[i][7]);

        if (val >= s5) return { label: "ดีมาก", score: 5 };
        if (val >= s4) return { label: "ดี", score: 4 };
        if (val >= s3) return { label: "ปานกลาง", score: 3 };
        if (val >= s2) return { label: "ต่ำ", score: 2 };
        return { label: "ต่ำมาก", score: 1 };
      }
    }
    return { label: "ไม่พบช่วงเกณฑ์ (" + searchTest + " / " + searchGender + " / " + searchAge + " ปี)", score: 0 };
  } catch(e) { return { label: "ข้อผิดพลาดระบบ: " + e.toString(), score: 0 }; }
}

function calculateRank(totalScore) {
  if (totalScore >= 23) return "S"; if (totalScore >= 19) return "A";
  if (totalScore >= 15) return "B"; if (totalScore >= 11) return "C";
  if (totalScore >= 7)  return "D"; return "E";
}

// =============================================================================
// HELPER FUNCTION: GET TEACHER DATA
// =============================================================================
// ฟังก์ชันใหม่ดึงข้อมูลครูประจำระดับชั้นตาม Job Title เพื่อนำไปแทนที่รหัสสัญลักษณ์ {{}} ใน PDF
function getTeacherNameByJobTitle(jobTitleKey) {
  try {
    const sheet = getDb().getSheetByName("Teachers");
    if (!sheet) return "-";
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][3].toString().trim() === jobTitleKey) {
        let prefix = data[i][0] ? data[i][0].toString().trim() : "";
        let fname = data[i][1] ? data[i][1].toString().trim() : "";
        let lname = data[i][2] ? data[i][2].toString().trim() : "";
        return `${prefix}${fname} ${lname}`.trim();
      }
    }
    return "-";
  } catch(e) {
    return "-";
  }
}

// =============================================================================
// SAVE & DATA PERSISTENCE
// =============================================================================
function saveAndCalc(f) {
  try {
    let ageParsed = parseInt(f.age);
    if (!f.age || isNaN(ageParsed) || ageParsed <= 0 || ageParsed > 100) {
      return { status: "error", message: "ไม่สามารถบันทึกได้ เนื่องจาก ใส่ข้อมูลอายุผิด หรือข้อมูลอายุไม่สมบูรณ์" };
    }

    // แก้ไขข้อ 2: คัดกรองและดักความถูกต้องของค่า "ห้อง" ให้ระบุได้แค่ตัวเลข 1 ถึง 16 เท่านั้น
    let roomParsed = parseInt(f.room);
    if (isNaN(roomParsed) || roomParsed < 1 || roomParsed > 16) {
      return { status: "error", message: "ไม่สามารถบันทึกได้ เนื่องจากข้อมูล 'ห้อง' ต้องเป็นตัวเลข 1 ถึง 16 เท่านั้น (เช่น ใส่ 4 แทน 2/4)" };
    }

    if (!f.Weight || f.Weight.toString().trim() === "" ||
        !f.High || f.High.toString().trim() === "" ||
        !f.Sit_and_Reach || f.Sit_and_Reach.toString().trim() === "" ||
        !f.Step_Up_3 || f.Step_Up_3.toString().trim() === "" ||
        !f.Push_Up_30 || f.Push_Up_30.toString().trim() === "" ||
        !f.Sit_Up_60 || f.Sit_Up_60.toString().trim() === "") {
      return { status: "error", message: "ไม่สามารถบันทึกได้ เนื่องจาก ยังไม่ได้บันทึกผลการทดสอบสมรรถภาพทางกาย หรือข้อมูลสัดส่วนร่างกายไม่ครบถ้วน" };
    }

    const db = getDb();
    let sheet = db.getSheetByName("Fitness_Logs");
   
    const headers = [
      "Timestamp", "Student_id", "Prefix", "First_name", "Last_name", "Gender", "Age", "Level", "Room", "No",
      "Semester", "Academic_year", "Weight", "High", "BMI",
      "BMI_Label", "Sit_and_Reach", "Sit_and_Reach_Label", "Step_Up_3", "Step_Up_3_Label",
      "Push_Up_30", "Push_Up_30_Label", "Sit_up_60", "Sit_up_60_Label",
      "S1", "S2", "S3", "S4", "S5", "Total", "Rank"
    ];

    if (!sheet) {
      sheet = db.insertSheet("Fitness_Logs");
      sheet.appendRow(headers);
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    let w = parseFloat(f.Weight); let h = parseFloat(f.High) / 100;
    let bmi = (w / (h * h)).toFixed(2);
   
    let resBmi = getBmiEvaluation(f.gender, f.age, parseFloat(bmi));
    let resS2  = getFitnessEvaluation("Sit_and_Reach", f.gender, f.age, f.Sit_and_Reach);
    let resS3  = getFitnessEvaluation("Step_Up_3", f.gender, f.age, f.Step_Up_3);
    let resS4  = getFitnessEvaluation("Push_Up_30", f.gender, f.age, f.Push_Up_30);
    let resS5  = getFitnessEvaluation("Sit_Up_60", f.gender, f.age, f.Sit_Up_60);
   
    let total = resBmi.score + resS2.score + resS3.score + resS4.score + resS5.score;
    let rank = calculateRank(total);
    let timestamp = new Date();

    const rowData = [
      timestamp, f.student_id, f.prefix, f.first_name, f.last_name, f.gender, ageParsed, f.class_level, roomParsed, parseInt(f.no),
      f.semester, f.academic_year, w, parseFloat(f.High), parseFloat(bmi),
      resBmi.label,
      parseFloat(f.Sit_and_Reach), resS2.label,
      parseFloat(f.Step_Up_3), resS3.label,      
      parseFloat(f.Push_Up_30), resS4.label,    
      parseFloat(f.Sit_Up_60), resS5.label,      
      resBmi.score, resS2.score, resS3.score, resS4.score, resS5.score,
      total, rank
    ];

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === f.student_id.toString().trim() &&
          data[i][10].toString().trim() === f.semester.toString().trim() &&
          data[i][11].toString().trim() === f.academic_year.toString().trim()) {
        targetRow = i + 1;
        break;
      }
    }

    let status = "insert";
    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).clearContent();
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      status = "update";
    } else {
      sheet.appendRow(rowData);
    }

    // 🎯 [ระบบ Sync ข้อมูลนักเรียนใหม่] ทำงานหลังจากจัดการหน้า Fitness_Logs สำเร็จแล้ว
    syncStudentDatabase(f);

    // แก้ไขข้อ 3: การส่งค่าส่วนของ f1 กลับไปหน้าบ้าน ถ้าคำนวณได้เป็น 0 หรือ 0.0 ให้โชว์ 0 ไม่ให้เป็นขีดว่าง
    let clientF1 = (f.Sit_and_Reach === 0 || f.Sit_and_Reach === "0" || parseFloat(f.Sit_and_Reach) === 0) ? "0" : f.Sit_and_Reach;

    return {
      status: "success",
      action: status,
      info: `${f.prefix}${f.first_name} ${f.last_name} (ม.${f.class_level.replace("ม.", "")}/${roomParsed} เลขที่ ${f.no})`,
      bmi: bmi, bmiLabel: resBmi.label, bmiScore: resBmi.score,
      f1: clientF1, f1Label: resS2.label, f1Score: resS2.score,
      f2: f.Step_Up_3, f2Label: resS3.label, f2Score: resS3.score,
      f3: f.Push_Up_30, f3Label: resS4.label, f3Score: resS4.score,
      f4: f.Sit_Up_60, f4Label: resS5.label, f4Score: resS5.score,
      total: total, rank: rank
    };
  } catch(e) {
    return { status: "error", message: "เกิดข้อผิดพลาดภายในระบบ: " + e.toString() };
  }
}

// =============================================================================
// ADMIN PANEL OPERATIONS (VERSION 3: แก้ไข Multi-level Sort และ Initial Load)
// =============================================================================
function getAdminData(f) {
  if (!f) return [];
  if (f.isInitialLoad === true) return [];

  const sheet = getDb().getSheetByName("Fitness_Logs");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let filtered = [];
 
  for(let i = 1; i < data.length; i++) {
    let rowSem = data[i][10] ? data[i][10].toString().trim() : "";
    let rowYr = data[i][11] ? data[i][11].toString().trim() : "";
    let rowLev = data[i][7] ? data[i][7].toString().trim() : "";
    let rowRom = data[i][8] ? data[i][8].toString().trim() : "";

    if ((!f.sem || f.sem === "all" || rowSem === f.sem.toString().trim()) &&
        (!f.yr || f.yr === "all" || rowYr === f.yr.toString().trim()) &&
        (!f.lev || f.lev === "all" || rowLev === f.lev.toString().trim()) &&
        (!f.rom || f.rom === "all" || rowRom === f.rom.toString().trim())) {
     
      let levelStr = rowLev.replace(/\D/g, '');
      let levelNum = levelStr === "" ? 99 : parseInt(levelStr);
      let roomNum = parseInt(rowRom) || 99;
      let noNum = parseInt(data[i][9]) || 999;

      // แก้ไขข้อ 3: การแสดงผลการนั่งงอตัวในหน้าหลักแอดมิน หากเป็น 0 ให้แสดงเลข 0 ชัดเจนไม่ให้เป็นขีดลบ
      let rawF1 = data[i][16];
      let displayF1 = (rawF1 === 0 || rawF1 === "0" || rawF1 === 0.0) ? "0" : (rawF1 ? rawF1.toString() : "-");

      filtered.push({
        id: data[i][1],
        levelNum: levelNum,
        roomNum: roomNum,  
        no: noNum,          
        lr: `${rowLev}/${rowRom}`,
        name: `${data[i][2]}${data[i][3]} ${data[i][4]}`,
        bmi: data[i][14], f1: displayF1, f2: data[i][18], f3: data[i][20], f4: data[i][22],
        sem: rowSem, yr: rowYr, total: data[i][29], rank: data[i][30]
      });
    }
  }

  return filtered.sort((a, b) => {
    if (a.levelNum !== b.levelNum) return a.levelNum - b.levelNum;
    if (a.roomNum !== b.roomNum) return a.roomNum - b.roomNum;  
    return a.no - b.no;                                          
  });
}

function deleteFitnessRecord(id, sem, yr) {
  const sheet = getDb().getSheetByName("Fitness_Logs");
  if (!sheet) return "ไม่พบชีตข้อมูล";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === id.toString().trim() && data[i][10].toString().trim() === sem.toString().trim() && data[i][11].toString().trim() === yr.toString().trim()) {
      sheet.deleteRow(i + 1); return "ลบข้อมูลเรียบร้อยแล้ว";
    }
  }
  return "ไม่พบข้อมูลที่ต้องการลบ";
}

// =========================================================================
// ส่วนระบบสร้างรายงาน PDF (เวอร์ชันปรับปรุงความแม่นยำดัชนีและเกณฑ์ข้อความ)
// =========================================================================
const DOC_TEMPLATE_ID_INDIVIDUAL = "1foHM3r9xpZg9d34X7AFb6tAvvkmlGORWZdpQ5eL2TfE";
const DOC_TEMPLATE_ID_CLASS = "1G6Wi2wynRGpyETcHOjvvN2vR8htABlnKGvDf_KGEFf8";
const DOC_TEMPLATE_ID_ANNUAL = "1lijYX7evzSlnTbXLIGpd3aPhouc_6eiIMCyUiELtL3s";

function generateIndividualPDF(studentId, semester, academicYear) {
  try {
    const db = getDb();
    const sheet = db.getSheetByName("Fitness_Logs");
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน Fitness_Logs" };
   
    const data = sheet.getDataRange().getValues();
    let rowData = null;
   
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === studentId.toString().trim() &&
          data[i][10].toString().trim() === semester.toString().trim() &&
          data[i][11].toString().trim() === academicYear.toString().trim()) {
        rowData = data[i];
        break;
      }
    }
   
    if (!rowData) return { success: false, message: "ไม่พบข้อมูลการทดสอบของนักเรียนคนนี้ในเทอม/ปีการศึกษาที่ระบุ" };
   
    // แก้ไขข้อ 3: การส่งออกรายงานของเด็กเป็นบุคคลถ้านั่งงอตัวเป็น 0 ให้แสดง 0 
    let f1Raw = rowData[16];
    let f1Display = (f1Raw === 0 || f1Raw === "0" || f1Raw === 0.0) ? "0" : (f1Raw ? f1Raw.toString() : "-");

    const student = {
      student_id: rowData[1] ? rowData[1].toString().trim() : "-",
      prefix: rowData[2] ? rowData[2].toString().trim() : "",
      first_name: rowData[3] ? rowData[3].toString().trim() : "",
      last_name: rowData[4] ? rowData[4].toString().trim() : "",
      gender: rowData[5] ? rowData[5].toString().trim() : "-",
      age: rowData[6] ? rowData[6].toString().trim() : "-",
      class_level: rowData[7] ? rowData[7].toString().trim() : "-",
      room: rowData[8] ? rowData[8].toString().trim() : "-",
      no: rowData[9] ? rowData[9].toString().trim() : "-",
      semester: rowData[10] ? rowData[10].toString().trim() : "-",
      academic_year: rowData[11] ? rowData[11].toString().trim() : "-",
      weight: rowData[12] ? rowData[12].toString().trim() : "-",
      height: rowData[13] ? rowData[13].toString().trim() : "-",
      bmi: rowData[14] ? rowData[14].toString().trim() : "-",
      bmi_label: rowData[15] ? rowData[15].toString().trim() : "-",
      f1: f1Display,
      f1_label: rowData[17] ? rowData[17].toString().trim() : "-",
      f2: rowData[18] ? rowData[18].toString().trim() : "-",
      f2_label: rowData[19] ? rowData[19].toString().trim() : "-",
      f3: rowData[20] ? rowData[20].toString().trim() : "-",
      f3_label: rowData[21] ? rowData[21].toString().trim() : "-",
      f4: rowData[22] ? rowData[22].toString().trim() : "-",
      f4_label: rowData[23] ? rowData[23].toString().trim() : "-",
      total: rowData[29] ? rowData[29].toString().trim() : "-",
      rank: rowData[30] ? rowData[30].toString().trim() : "-",
      test_date: rowData[0] ? Utilities.formatDate(new Date(rowData[0]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "-"
    };

    // แก้ไขข้อ 4, 5, 7: ดึงชื่อครูตามระดับชั้นของเด็กคนนี้ (ดึงเลขอายุจาก ม.1-6 มาแมตช์คีย์)
    let cleanLevelNum = student.class_level.replace(/\D/g, ''); // จะได้เลขเช่น "2"
    let teacherAdvisor = getTeacherNameByJobTitle("test_advisor" + cleanLevelNum);
    let teacherVerifier = getTeacherNameByJobTitle("test_verifier" + cleanLevelNum);

    const folder = DriveApp.getFolderById(FOLDER_PDF_ID);
    const tempCopy = DriveApp.getFileById(DOC_TEMPLATE_ID_INDIVIDUAL).makeCopy("รายงานบุคคล_" + student.first_name, folder);
    const doc = DocumentApp.openById(tempCopy.getId());
    const body = doc.getBody();

    body.replaceText('{{semester}}', student.semester);
    body.replaceText('{{academic_year}}', student.academic_year);
    body.replaceText('{{student_id}}', student.student_id);
    body.replaceText('{{prefix}}', student.prefix);
    body.replaceText('{{first_name}}', student.first_name);
    body.replaceText('{{last_name}}', student.last_name);
    body.replaceText('{{class_level}}', student.class_level);
    body.replaceText('{{room}}', student.room);
    body.replaceText('{{no}}', student.no);
    body.replaceText('{{gender}}', student.gender);
    body.replaceText('{{age}}', student.age);
    body.replaceText('{{weight}}', student.weight);
    body.replaceText('{{height}}', student.height);
    body.replaceText('{{test_date}}', student.test_date);
    body.replaceText('{{bmi}}', student.bmi);
    body.replaceText('{{bmi_label}}', student.bmi_label);
    body.replaceText('{{f1}}', student.f1);
    body.replaceText('{{f1_label}}', student.f1_label);
    body.replaceText('{{f2}}', student.f2);
    body.replaceText('{{f2_label}}', student.f2_label);
    body.replaceText('{{f3}}', student.f3);
    body.replaceText('{{f3_label}}', student.f3_label);
    body.replaceText('{{f4}}', student.f4);
    body.replaceText('{{f4_label}}', student.f4_label);
    body.replaceText('{{total}}', student.total);
    body.replaceText('{{rank}}', student.rank);
    
    // แทนที่สัญลักษณ์ชื่อครูท้ายกระดาษสำหรับบุคคล
    body.replaceText('{{test_advisor}}', teacherAdvisor);
    body.replaceText('{{test_verifier}}', teacherVerifier);

    const sarabunStyle = {};
    sarabunStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    body.setAttributes(sarabunStyle);
   
    const header = doc.getHeader();
    if (header) {
      const headerText = header.editAsText();
      if (headerText && headerText.getText().trim() !== "") {
        headerText.setFontFamily('Sarabun');
      }
    }
   
    const footer = doc.getFooter();
    if (footer) {
      const footerText = footer.editAsText();
      if (footerText && footerText.getText().trim() !== "") {
        footerText.setFontFamily('Sarabun');
      }
    }

    doc.saveAndClose();

    const pdfBlob = tempCopy.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob).setName("Report_Individual_" + studentId + "_" + academicYear + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempCopy.setTrashed(true);

    return { success: true, url: pdfFile.getUrl(), downloadUrl: pdfFile.getDownloadUrl() };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดรายงานบุคคล: " + e.toString() };
  }
}

function generateClassPDF(classLevel, room, semester, academicYear) {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    const data = sheet.getDataRange().getValues();
   
    const folder = DriveApp.getFolderById(FOLDER_PDF_ID);
    const tempCopy = DriveApp.getFileById(DOC_TEMPLATE_ID_CLASS).makeCopy("รายงานชั้นเรียน_ม_" + classLevel + "_" + room, folder);
    const doc = DocumentApp.openById(tempCopy.getId());
    const body = doc.getBody();

    body.replaceText('{{class_level}}', classLevel);
    body.replaceText('{{room}}', room);
    body.replaceText('{{semester}}', semester);
    body.replaceText('{{academic_year}}', academicYear);

    // แก้ไขข้อ 4, 5, 7: ดึงชื่อครูสำหรับรายงานรายห้อง
    let cleanLevelNum = classLevel.toString().replace(/\D/g, '');
    let teacherAdvisor = getTeacherNameByJobTitle("test_advisor" + cleanLevelNum);
    let teacherVerifier = getTeacherNameByJobTitle("test_verifier" + cleanLevelNum);
    
    body.replaceText('{{test_advisor}}', teacherAdvisor);
    body.replaceText('{{test_verifier}}', teacherVerifier);

    let classRows = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][7].toString().trim() === classLevel.toString().trim() &&
          data[i][8].toString().trim() === room.toString().trim() &&
          data[i][10].toString().trim() === semester.toString().trim() &&
          data[i][11].toString().trim() === academicYear.toString().trim()) {
        classRows.push(data[i]);
      }
    }
   
    classRows.sort((a, b) => (parseInt(a[9]) || 0) - (parseInt(b[9]) || 0));

    let statBmi = { "ผอมมาก": 0, "ผอม": 0, "สมส่วน": 0, "ท้วม": 0, "อ้วน": 0 };
    let statF1 = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF2 = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF3 = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF4 = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let totalStudentsInClass = classRows.length;

    const tables = body.getTables();
   
    if (tables.length > 0) {
      const table1 = tables[0];

      for (let k = 0; k < classRows.length; k++) {
        let row = classRows[k];
        let newRow = table1.appendTableRow();
       
        let s_no = row[9] ? row[9].toString().trim() : "-";
        let s_prefix = row[2] ? row[2].toString().trim() : ""; 
        let s_firstname = row[3] ? row[3].toString().trim() : "-";
        let s_lastname = row[4] ? row[4].toString().trim() : "-";
        let s_bmi = row[14] ? row[14].toString().trim() : "-";
        
        // แก้ไขข้อ 3: การโชว์ผลนั่งงอตัวในตารางจดบันทึกรายห้อง
        let f1Raw = row[16];
        let s_f1 = (f1Raw === 0 || f1Raw === "0" || f1Raw === 0.0) ? "0" : (f1Raw ? f1Raw.toString() : "-");
        
        let s_f2 = row[18] ? row[18].toString().trim() : "-";
        let s_f3 = row[20] ? row[20].toString().trim() : "-";
        let s_f4 = row[22] ? row[22].toString().trim() : "-";
        let s_total = row[29] ? row[29].toString().trim() : "-";
        let s_rank = row[30] ? row[30].toString().trim() : "-";

        newRow.appendTableCell(s_no);        
        newRow.appendTableCell(s_prefix);    
        newRow.appendTableCell(s_firstname); 
        newRow.appendTableCell(s_lastname);  
        newRow.appendTableCell(s_bmi);       
        newRow.appendTableCell(s_f1);        
        newRow.appendTableCell(s_f2);        
        newRow.appendTableCell(s_f3);        
        newRow.appendTableCell(s_f4);        
        newRow.appendTableCell(s_total);     
        newRow.appendTableCell(s_rank);      

        let lblBmi = row[15] ? row[15].toString().trim() : "";
        let lblF1 = row[17] ? row[17].toString().trim() : "";
        let lblF2 = row[19] ? row[19].toString().trim() : "";
        let lblF3 = row[21] ? row[21].toString().trim() : "";
        let lblF4 = row[23] ? row[23].toString().trim() : "";

        if (statBmi.hasOwnProperty(lblBmi)) statBmi[lblBmi]++;
        if (statF1.hasOwnProperty(lblF1)) statF1[lblF1]++;
        if (statF2.hasOwnProperty(lblF2)) statF2[lblF2]++;
        if (statF3.hasOwnProperty(lblF3)) statF3[lblF3]++;
        if (statF4.hasOwnProperty(lblF4)) statF4[lblF4]++;
      }
    }
   
    body.replaceText('{{bmi_ผอมมาก}}', statBmi["ผอมมาก"].toString());
    body.replaceText('{{bmi_ผอม}}', statBmi["ผอม"].toString());
    body.replaceText('{{bmi_สมส่วน}}', statBmi["สมส่วน"].toString());
    body.replaceText('{{bmi_ท้วม}}', statBmi["ท้วม"].toString());
    body.replaceText('{{bmi_อ้วน}}', statBmi["อ้วน"].toString());
    body.replaceText('{{total_bmi}}', totalStudentsInClass.toString());

    body.replaceText('{{f1_ต่ำมาก}}', statF1["ต่ำมาก"].toString());
    body.replaceText('{{f1_ต่ำ}}', statF1["ต่ำ"].toString());
    body.replaceText('{{f1_ปานกลาง}}', statF1["ปานกลาง"].toString());
    body.replaceText('{{f1_ดี}}', statF1["ดี"].toString());
    body.replaceText('{{f1_ดีมาก}}', statF1["ดีมาก"].toString());
    body.replaceText('{{total_f1}}', totalStudentsInClass.toString());

    body.replaceText('{{f2_ต่ำมาก}}', statF2["ต่ำมาก"].toString());
    body.replaceText('{{f2_ต่ำ}}', statF2["ต่ำ"].toString());
    body.replaceText('{{f2_ปานกลาง}}', statF2["ปานกลาง"].toString());
    body.replaceText('{{f2_ดี}}', statF2["ดี"].toString());
    body.replaceText('{{f2_ดีมาก}}', statF2["ดีมาก"].toString());
    body.replaceText('{{total_f2}}', totalStudentsInClass.toString());

    body.replaceText('{{f3_ต่ำมาก}}', statF3["ต่ำมาก"].toString());
    body.replaceText('{{f3_ต่ำ}}', statF3["ต่ำ"].toString());
    body.replaceText('{{f3_ปานกลาง}}', statF3["ปานกลาง"].toString());
    body.replaceText('{{f3_ดี}}', statF3["ดี"].toString());
    body.replaceText('{{f3_ดีมาก}}', statF3["ดีมาก"].toString());
    body.replaceText('{{total_f3}}', totalStudentsInClass.toString());

    body.replaceText('{{f4_ต่ำมาก}}', statF4["ต่ำมาก"].toString());
    body.replaceText('{{f4_ต่ำ}}', statF4["ต่ำ"].toString());
    body.replaceText('{{f4_ปานกลาง}}', statF4["ปานกลาง"].toString());
    body.replaceText('{{f4_ดี}}', statF4["ดี"].toString());
    body.replaceText('{{f4_ดีมาก}}', statF4["ดีมาก"].toString());
    body.replaceText('{{total_f4}}', totalStudentsInClass.toString());

    const sarabunStyle = {};
    sarabunStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    body.setAttributes(sarabunStyle);
   
    const header = doc.getHeader();
    if (header) {
      const headerText = header.editAsText();
      if (headerText && headerText.getText().trim() !== "") {
        headerText.setFontFamily('Sarabun');
      }
    }
   
    const footer = doc.getFooter();
    if (footer) {
      const footerText = footer.editAsText();
      if (footerText && footerText.getText().trim() !== "") {
        footerText.setFontFamily('Sarabun');
      }
    }

    doc.saveAndClose();
   
    const pdfBlob = tempCopy.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob).setName("Class_Report_ม_" + classLevel + "_" + room + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempCopy.setTrashed(true);

    return { success: true, url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดคลาสรีพอร์ต: " + e.toString() };
  }
}

function generateAnnualPDF(classLevel, semester, academicYear) {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(FOLDER_PDF_ID);
   
    const safeClassLevel = (classLevel && classLevel.toString().trim() !== "") ? classLevel.toString().trim() : "ทั้งหมด";
    const safeSemester = (semester && semester.toString().trim() !== "") ? semester.toString().trim() : "1";
    const safeAcademicYear = (academicYear && academicYear.toString().trim() !== "") ? academicYear.toString().trim() : "2569";

    let filteredRows = [];
    for (let i = 1; i < data.length; i++) {
      let rowLevel = data[i][7] ? data[i][7].toString().trim() : "";      
      let rowSemester = data[i][10] ? data[i][10].toString().trim() : "";  
      let rowYear = data[i][11] ? data[i][11].toString().trim() : "";      
     
      if (rowSemester == safeSemester && rowYear == safeAcademicYear) {
        if (safeClassLevel === "ทั้งหมด" || safeClassLevel === "all") {
          filteredRows.push(data[i]);
        } else {
          let targetLevel = safeClassLevel.startsWith("ม.") ? safeClassLevel : "ม." + safeClassLevel;
          if (rowLevel == targetLevel) {
            filteredRows.push(data[i]);
          }
        }
      }
    }
   
    if (filteredRows.length === 0) {
      let displayLevel = (safeClassLevel === "ทั้งหมด" || safeClassLevel === "all") ? "ภาพรวมทั้งโรงเรียน" : "ชั้น " + safeClassLevel;
      return { success: false, message: "ไม่พบข้อมูลนักเรียนตามเงื่อนไข ภาคเรียน " + safeSemester + " ปีการศึกษา " + safeAcademicYear + " (" + displayLevel + ") ที่ระบุ" };
    }

    // แก้ไขข้อ 4, 5, 7: ดึงชื่อครูสำหรับรายงานประจำปี (ถ้าเลือก "ทั้งหมด" ให้ดึงผู้ดูแล ม.2 เป็นตัวหลัก หรือปรับแต่งเงื่อนไขตามชอบ)
    let finalLevelNum = (safeClassLevel === "ทั้งหมด" || safeClassLevel === "all") ? "2" : safeClassLevel.replace(/\D/g, '');
    let teacherAdvisor = getTeacherNameByJobTitle("test_advisor" + finalLevelNum);
    let teacherVerifier = getTeacherNameByJobTitle("test_verifier" + finalLevelNum);

    let statBmi = { "ผอมมาก": 0, "ผอม": 0, "สมส่วน": 0, "ท้วม": 0, "อ้วน": 0 };
    let statF1  = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF2  = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF3  = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let statF4  = { "ต่ำมาก": 0, "ต่ำ": 0, "ปานกลาง": 0, "ดี": 0, "ดีมาก": 0 };
    let totalCount = filteredRows.length;

    let countBmiPass = 0, countF1Pass = 0, countF2Pass = 0, countF3Pass = 0, countF4Pass = 0;

    for (let i = 0; i < filteredRows.length; i++) {
      let row = filteredRows[i];
      let lblBmi = row[15] ? row[15].toString().trim() : "";
      let lblF1  = row[17] ? row[17].toString().trim() : "";
      let lblF2  = row[19] ? row[19].toString().trim() : "";
      let lblF3  = row[21] ? row[21].toString().trim() : "";
      let lblF4  = row[23] ? row[23].toString().trim() : "";

      if (statBmi.hasOwnProperty(lblBmi)) statBmi[lblBmi]++;
      if (statF1.hasOwnProperty(lblF1)) statF1[lblF1]++;
      if (statF2.hasOwnProperty(lblF2)) statF2[lblF2]++;
      if (statF3.hasOwnProperty(lblF3)) statF3[lblF3]++;
      if (statF4.hasOwnProperty(lblF4)) statF4[lblF4]++;

      if (lblBmi === "สมส่วน") countBmiPass++;
      if (["ปานกลาง", "ดี", "ดีมาก"].includes(lblF1)) countF1Pass++;
      if (["ปานกลาง", "ดี", "ดีมาก"].includes(lblF2)) countF2Pass++;
      if (["ปานกลาง", "ดี", "ดีมาก"].includes(lblF3)) countF3Pass++;
      if (["ปานกลาง", "ดี", "ดีมาก"].includes(lblF4)) countF4Pass++;
    }

    let pctBmi = totalCount > 0 ? ((countBmiPass / totalCount) * 100).toFixed(2) : "0.00";
    let pctF1  = totalCount > 0 ? ((countF1Pass  / totalCount) * 100).toFixed(2) : "0.00";
    let pctF2  = totalCount > 0 ? ((countF2Pass  / totalCount) * 100).toFixed(2) : "0.00";
    let pctF3  = totalCount > 0 ? ((countF3Pass  / totalCount) * 100).toFixed(2) : "0.00";
    let pctF4  = totalCount > 0 ? ((countF4Pass  / totalCount) * 100).toFixed(2) : "0.00";

    let fileTitle = "รายงานสรุปประจำปี_ภาคเรียน_" + safeSemester + "_ปี_" + safeAcademicYear + "_ชั้น_" + ((safeClassLevel === "ทั้งหมด" || safeClassLevel === "all") ? "ทั้งหมด" : safeClassLevel);
    const tempCopy = DriveApp.getFileById(DOC_TEMPLATE_ID_ANNUAL).makeCopy(fileTitle, folder);
    const doc = DocumentApp.openById(tempCopy.getId());
    const body = doc.getBody();

    body.replaceText('{{semester}}', safeSemester);
    body.replaceText('{{academic_year}}', safeAcademicYear);

    if (safeClassLevel === "ทั้งหมด" || safeClassLevel === "all") {
      body.replaceText('{{class_level}}', "1 - 6");
    } else {
      let cleanNumber = safeClassLevel.replace("ม.", "").trim();
      body.replaceText('{{class_level}}', cleanNumber);
    }

    body.replaceText('{{bmi_percent}}', pctBmi + "%");
    body.replaceText('{{f1_pass}}', countF1Pass.toString());
    body.replaceText('{{f1_percent}}', pctF1 + "%");
    body.replaceText('{{f2_pass}}', countF2Pass.toString());
    body.replaceText('{{f2_percent}}', pctF2 + "%");
    body.replaceText('{{f3_pass}}', countF3Pass.toString());
    body.replaceText('{{f3_percent}}', pctF3 + "%");
    body.replaceText('{{f4_pass}}', countF4Pass.toString());
    body.replaceText('{{f4_percent}}', pctF4 + "%");
    body.replaceText('{{total_students}}', totalCount.toString());

    body.replaceText('{{bmi_ผอมมาก}}', statBmi["ผอมมาก"].toString());
    body.replaceText('{{bmi_ผอม}}', statBmi["ผอม"].toString());
    body.replaceText('{{bmi_สมส่วน}}', statBmi["สมส่วน"].toString());
    body.replaceText('{{bmi_ท้วม}}', statBmi["ท้วม"].toString());
    body.replaceText('{{bmi_อ้วน}}', statBmi["อ้วน"].toString());
    body.replaceText('{{total_bmi}}', totalCount.toString());

    body.replaceText('{{f1_ต่ำมาก}}', statF1["ต่ำมาก"].toString());
    body.replaceText('{{f1_ต่ำ}}', statF1["ต่ำ"].toString());
    body.replaceText('{{f1_ปานกลาง}}', statF1["ปานกลาง"].toString());
    body.replaceText('{{f1_ดี}}', statF1["ดี"].toString());
    body.replaceText('{{f1_ดีมาก}}', statF1["ดีมาก"].toString());
    body.replaceText('{{total_f1}}', totalCount.toString());

    body.replaceText('{{f2_ต่ำมาก}}', statF2["ต่ำมาก"].toString());
    body.replaceText('{{f2_ต่ำ}}', statF2["ต่ำ"].toString());
    body.replaceText('{{f2_ปานกลาง}}', statF2["ปานกลาง"].toString());
    body.replaceText('{{f2_ดี}}', statF2["ดี"].toString());
    body.replaceText('{{f2_ดีมาก}}', statF2["ดีมาก"].toString());
    body.replaceText('{{total_f2}}', totalCount.toString());

    body.replaceText('{{f3_ต่ำมาก}}', statF3["ต่ำมาก"].toString());
    body.replaceText('{{f3_ต่ำ}}', statF3["ต่ำ"].toString());
    body.replaceText('{{f3_ปานกลาง}}', statF3["ปานกลาง"].toString());
    body.replaceText('{{f3_ดี}}', statF3["ดี"].toString());
    body.replaceText('{{f3_ดีมาก}}', statF3["ดีมาก"].toString());
    body.replaceText('{{total_f3}}', totalCount.toString());

    body.replaceText('{{f4_ต่ำมาก}}', statF4["ต่ำมาก"].toString());
    body.replaceText('{{f4_ต่ำ}}', statF4["ต่ำ"].toString());
    body.replaceText('{{f4_ปานกลาง}}', statF4["ปานกลาง"].toString());
    body.replaceText('{{f4_ดี}}', statF4["ดี"].toString());
    body.replaceText('{{f4_ดีมาก}}', statF4["ดีมาก"].toString());
    body.replaceText('{{total_f4}}', totalCount.toString());
    
    // แทนที่สัญลักษณ์ชื่อครูท้ายกระดาษสำหรับรายงานสรุปประจำปี
    body.replaceText('{{test_advisor}}', teacherAdvisor);
    body.replaceText('{{test_verifier}}', teacherVerifier);

    const sarabunStyle = {};
    sarabunStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Sarabun';
    body.setAttributes(sarabunStyle);
   
    const header = doc.getHeader();
    if (header) {
      const headerText = header.editAsText();
      if (headerText && headerText.getText().trim() !== "") {
        headerText.setFontFamily('Sarabun');
      }
    }
   
    const footer = doc.getFooter();
    if (footer) {
      const footerText = footer.editAsText();
      if (footerText && footerText.getText().trim() !== "") {
        footerText.setFontFamily('Sarabun');
      }
    }

    doc.saveAndClose();
    const pdfBlob = tempCopy.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob).setName(fileTitle + ".pdf");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    tempCopy.setTrashed(true);

    return { success: true, url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาดในการสร้างรายงานประจำปี: " + e.toString() };
  }
}

// =============================================================================
// 🎯 SUB-ROUTINE: AUTOMATIC STUDENT DATABASE SYNC (NEW FUNCTION)
// =============================================================================
function syncStudentDatabase(f) {
  try {
    const db = getDb();
    let studentSheet = db.getSheetByName("Student_DB");
   
    const studentHeaders = [
      "Student_id", "Prefix", "First_name", "Last_name", "Gender", "Age", "Level", "Room", "No"
    ];

    if (!studentSheet) {
      studentSheet = db.insertSheet("Student_DB");
      studentSheet.appendRow(studentHeaders);
    }

    const studentData = studentSheet.getDataRange().getValues();
    let targetRowIndex = -1;

    for (let i = 1; i < studentData.length; i++) {
      if (studentData[i][0].toString().trim() === f.student_id.toString().trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    // แก้ไขข้อ 2: ตรวจสอบความถูกต้องของห้องก่อนเซฟลงทะเบียนประวัติ
    let cleanRoom = parseInt(f.room);
    if(isNaN(cleanRoom)) cleanRoom = f.room;

    const updatedRowValues = [
      f.student_id.toString().trim(),         
      f.prefix ? f.prefix.toString().trim() : "",               
      f.first_name ? f.first_name.toString().trim() : "",       
      f.last_name ? f.last_name.toString().trim() : "",         
      f.gender ? f.gender.toString().trim() : "",               
      parseInt(f.age),                        
      f.class_level ? f.class_level.toString().trim() : "",     
      cleanRoom.toString().trim(),                   
      parseInt(f.no)                          
    ];

    if (targetRowIndex > 0) {
      // ดึงข้อมูลแถวเดิมในชีตประวัติเพื่อเช็กว่ามีฐานข้อมูลเดิมอยู่แล้วหรือไม่
      // หากต้องการล็อกระดับความปลอดภัยขั้นสูงสุด สามารถระบุไม่แก้ฟิลด์ที่ล็อกได้ที่ส่วนนี้
      studentSheet.getRange(targetRowIndex, 1, 1, updatedRowValues.length).setValues([updatedRowValues]);
      Logger.log("Sync System: อัปเดตข้อมูลนักเรียนเก่าเรียบร้อย รหัส " + f.student_id);
    } else {
      studentSheet.appendRow(updatedRowValues);
      Logger.log("Sync System: เพิ่มนักเรียนใหม่เข้าทะเบียนประวัติเรียบร้อย รหัส " + f.student_id);
    }
  } catch (e) {
    Logger.log("เกิดข้อผิดพลาดในฟังก์ชันย่อยระบบ SyncStudentDatabase: " + e.toString());
  }
}