// =============================================================================
// DATABASE SERVICES & DATA ACCESS
// =============================================================================

function getYearList() {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    const currentYear = (new Date().getFullYear() + 543).toString();
    if (!sheet) return [currentYear];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [currentYear];
   
    let years = [];
    for (let i = 1; i < data.length; i++) {
      let yr = data[i][11] ? data[i][11].toString().trim() : "";
      if (yr && !years.includes(yr)) {
        years.push(yr);
      }
    }
    
    if (!years.includes(currentYear)) {
      years.push(currentYear);
    }
    
    return years.sort((a, b) => b - a);
  } catch(e) { 
    return [(new Date().getFullYear() + 543).toString()]; 
  }
}

function getBmiCriteriaMatrix() {
  try {
    const sheet = getDb().getSheetByName("Criteria_BMI");
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  } catch(e) { return []; }
}

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
        
        let valF1 = fitnessData[j][16];
        result.f1 = (valF1 !== undefined && valF1 !== null && valF1 !== "") ? valF1 : "";
        
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
      
      let valF1 = data[i][16];
      let displayF1 = (valF1 !== undefined && valF1 !== null && valF1 !== "") ? valF1.toString() : "-";
     
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
      studentSheet.getRange(targetRowIndex, 1, 1, updatedRowValues.length).setValues([updatedRowValues]);
    } else {
      studentSheet.appendRow(updatedRowValues);
    }
  } catch (e) {
    Logger.log("Error in syncStudentDatabase: " + e.toString());
  }
}

function deleteFitnessRecord(studentId, semester, academicYear) {
  try {
    const sheet = getDb().getSheetByName("Fitness_Logs");
    if (!sheet) return "ไม่พบแผ่นงาน Fitness_Logs";
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === studentId.toString().trim() &&
          data[i][10].toString().trim() === semester.toString().trim() &&
          data[i][11].toString().trim() === academicYear.toString().trim()) {
        sheet.deleteRow(i + 1);
        return "ลบข้อมูลสำเร็จ";
      }
    }
    return "ไม่พบรายการที่ต้องการลบ";
  } catch (e) {
    return "เกิดข้อผิดพลาดในการลบข้อมูล: " + e.toString();
  }
}