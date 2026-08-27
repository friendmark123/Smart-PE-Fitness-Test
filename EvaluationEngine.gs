// =============================================================================
// CORE CALCULATION & EVALUATION ENGINE
// =============================================================================

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

function saveAndCalc(f) {
  try {
    let ageParsed = parseInt(f.age);
    if (!f.age || isNaN(ageParsed) || ageParsed <= 0 || ageParsed > 100) {
      return { status: "error", message: "ไม่สามารถบันทึกได้ เนื่องจาก ใส่ข้อมูลอายุผิด หรือข้อมูลอายุไม่สมบูรณ์" };
    }

    let roomParsed = parseInt(f.room);
    if (isNaN(roomParsed) || roomParsed < 1 || roomParsed > 16) {
      return { status: "error", message: "ไม่สามารถบันทึกได้ เนื่องจากข้อมูล 'ห้อง' ต้องเป็นตัวเลข 1 ถึง 16 เท่านั้น" };
    }

    if (!f.Weight || f.Weight.toString().trim() === "" ||
        !f.High || f.High.toString().trim() === "" ||
        f.Sit_and_Reach === undefined || f.Sit_and_Reach === null || f.Sit_and_Reach.toString().trim() === "" ||
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

    syncStudentDatabase(f);

    return {
      status: "success",
      action: status,
      info: `${f.prefix}${f.first_name} ${f.last_name} (ม.${f.class_level.replace("ม.", "")}/${roomParsed} เลขที่ ${f.no})`,
      bmi: bmi, bmiLabel: resBmi.label, bmiScore: resBmi.score,
      f1: f.Sit_and_Reach, f1Label: resS2.label, f1Score: resS2.score,
      f2: f.Step_Up_3, f2Label: resS3.label, f2Score: resS3.score,
      f3: f.Push_Up_30, f3Label: resS4.label, f3Score: resS4.score,
      f4: f.Sit_Up_60, f4Label: resS5.label, f4Score: resS5.score,
      total: total, rank: rank
    };
  } catch(e) {
    return { status: "error", message: "เกิดข้อผิดพลาดภายในระบบ: " + e.toString() };
  }
}