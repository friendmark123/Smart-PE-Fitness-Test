// =============================================================================
// PDF GENERATION SERVICES (Individual, Class, Annual Reports)
// =============================================================================

function generateIndividualPDF(studentId, semester, academicYear) {
  try {
    const db = getDb();
    const sheet = db.getSheetByName("Fitness_Logs");
    if (!sheet) return { status: "error", message: "ไม่พบแผ่นงาน Fitness_Logs" };
    
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
    
    if (!rowData) return { status: "error", message: "ไม่พบข้อมูลสมรรถภาพของนักเรียนในเทอม/ปีการศึกษานี้" };

    const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID_INDIVIDUAL);
    const targetFolder = DriveApp.getFolderById(FOLDER_PDF_ID);
    const newFile = templateFile.makeCopy(`รายงานสมรรถภาพ_${rowData[1]}_${rowData[3]}_${semester}_${academicYear}`, targetFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    let valF1 = rowData[16];
    let displayF1 = (valF1 !== undefined && valF1 !== null && valF1 !== "") ? valF1.toString() : "-";

    const headTeacher = getTeacherNameByJobTitle("หัวหน้ากลุ่มสาระการเรียนรู้สุขศึกษาและพลศึกษา");
    const testTeacher = getTeacherNameByJobTitle("ครูผู้สอน/ผู้ทดสอบ");

    body.replaceText("{{STUDENT_ID}}", rowData[1] || "-");
    body.replaceText("{{NAME}}", `${rowData[2]}${rowData[3]} ${rowData[4]}`);
    body.replaceText("{{GENDER}}", rowData[5] || "-");
    body.replaceText("{{AGE}}", rowData[6] || "-");
    body.replaceText("{{LEVEL}}", rowData[7] || "-");
    body.replaceText("{{ROOM}}", rowData[8] || "-");
    body.replaceText("{{NO}}", rowData[9] || "-");
    body.replaceText("{{SEMESTER}}", rowData[10] || "-");
    body.replaceText("{{YEAR}}", rowData[11] || "-");
    body.replaceText("{{WEIGHT}}", rowData[12] || "-");
    body.replaceText("{{HEIGHT}}", rowData[13] || "-");
    body.replaceText("{{BMI}}", rowData[14] || "-");
    body.replaceText("{{BMI_L}}", rowData[15] || "-");
    body.replaceText("{{F1}}", displayF1);
    body.replaceText("{{F1_L}}", rowData[17] || "-");
    body.replaceText("{{F2}}", rowData[18] || "-");
    body.replaceText("{{F2_L}}", rowData[19] || "-");
    body.replaceText("{{F3}}", rowData[20] || "-");
    body.replaceText("{{F3_L}}", rowData[21] || "-");
    body.replaceText("{{F4}}", rowData[22] || "-");
    body.replaceText("{{F4_L}}", rowData[23] || "-");
    body.replaceText("{{TOTAL}}", rowData[29] || "-");
    body.replaceText("{{RANK}}", rowData[30] || "-");
    body.replaceText("{{HEAD_TEACHER}}", headTeacher);
    body.replaceText("{{TEST_TEACHER}}", testTeacher);

    doc.saveAndClose();

    const pdfBlob = newFile.getAs('application/pdf');
    const pdfFile = targetFolder.createFile(pdfBlob).setName(`FitnessReport_${rowData[1]}_S${semester}_Y${academicYear}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    newFile.setTrashed(true);

    return { status: "success", url: pdfFile.getUrl(), downloadUrl: pdfFile.getDownloadUrl() };
  } catch(e) {
    return { status: "error", message: "สร้าง PDF ไม่สำเร็จ: " + e.toString() };
  }
}

function generateClassPDF(classLevel, room, semester, academicYear) {
  try {
    const db = getDb();
    const sheet = db.getSheetByName("Fitness_Logs");
    if (!sheet) return { status: "error", message: "ไม่พบแผ่นงาน Fitness_Logs" };
    
    const data = sheet.getDataRange().getValues();
    let classStudents = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][7].toString().trim() === classLevel.toString().trim() &&
          data[i][8].toString().trim() === room.toString().trim() &&
          data[i][10].toString().trim() === semester.toString().trim() &&
          data[i][11].toString().trim() === academicYear.toString().trim()) {
        classStudents.push(data[i]);
      }
    }

    if (classStudents.length === 0) {
      return { status: "error", message: "ไม่พบข้อมูลนักเรียนในห้องเรียนที่เลือก" };
    }

    classStudents.sort((a, b) => (parseInt(a[9]) || 0) - (parseInt(b[9]) || 0));

    const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID_CLASS);
    const targetFolder = DriveApp.getFolderById(FOLDER_PDF_ID);
    const newFile = templateFile.makeCopy(`รายงานห้อง_ม.${classLevel}_${room}_เทอม${semester}_${academicYear}`, targetFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    const headTeacher = getTeacherNameByJobTitle("หัวหน้ากลุ่มสาระการเรียนรู้สุขศึกษาและพลศึกษา");
    const testTeacher = getTeacherNameByJobTitle("ครูผู้สอน/ผู้ทดสอบ");

    body.replaceText("{{CLASS}}", classLevel);
    body.replaceText("{{ROOM}}", room);
    body.replaceText("{{SEMESTER}}", semester);
    body.replaceText("{{YEAR}}", academicYear);
    body.replaceText("{{HEAD_TEACHER}}", headTeacher);
    body.replaceText("{{TEST_TEACHER}}", testTeacher);

    let tables = body.getTables();
    if (tables.length > 0) {
      let table = tables[0];
      for (let s = 0; s < classStudents.length; s++) {
        let st = classStudents[s];
        let row = table.appendRow();
        row.addCell((s + 1).toString());
        row.addCell(st[1].toString());
        row.addCell(`${st[2]}${st[3]} ${st[4]}`);
        row.addCell(st[14].toString()); // BMI
        row.addCell(st[15].toString()); // BMI Label
        row.addCell(st[16] ? st[16].toString() : "-"); // F1
        row.addCell(st[18].toString()); // F2
        row.addCell(st[20].toString()); // F3
        row.addCell(st[22].toString()); // F4
        row.addCell(st[29].toString()); // Total
        row.addCell(st[30].toString()); // Rank
      }
    }

    doc.saveAndClose();
    const pdfBlob = newFile.getAs('application/pdf');
    const pdfFile = targetFolder.createFile(pdfBlob).setName(`ClassReport_M${classLevel}_${room}_S${semester}_Y${academicYear}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    newFile.setTrashed(true);
    return { status: "success", url: pdfFile.getUrl() };
  } catch(e) {
    return { status: "error", message: "สร้าง PDF ประจำห้องไม่สำเร็จ: " + e.toString() };
  }
}

function generateAnnualPDF(classLevel, semester, academicYear) {
  try {
    const db = getDb();
    const sheet = db.getSheetByName("Fitness_Logs");
    if (!sheet) return { status: "error", message: "ไม่พบแผ่นงาน Fitness_Logs" };
    
    const data = sheet.getDataRange().getValues();
    let annualStudents = [];
    
    for (let i = 1; i < data.length; i++) {
      let matchClass = (classLevel === "all" || data[i][7].toString().trim() === classLevel.toString().trim());
      if (matchClass &&
          data[i][10].toString().trim() === semester.toString().trim() &&
          data[i][11].toString().trim() === academicYear.toString().trim()) {
        annualStudents.push(data[i]);
      }
    }

    if (annualStudents.length === 0) {
      return { status: "error", message: "ไม่พบข้อมูลสำหรับสรุปประจำปี" };
    }

    const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID_ANNUAL);
    const targetFolder = DriveApp.getFolderById(FOLDER_PDF_ID);
    const newFile = templateFile.makeCopy(`รายงานสรุปภาพรวม_ม.${classLevel}_เทอม${semester}_${academicYear}`, targetFolder);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();

    let rankCount = { S: 0, A: 0, B: 0, C: 0, D: 0, E: 0 };
    annualStudents.forEach(st => {
      let r = st[30];
      if (rankCount[r] !== undefined) rankCount[r]++;
    });

    const headTeacher = getTeacherNameByJobTitle("หัวหน้ากลุ่มสาระการเรียนรู้สุขศึกษาและพลศึกษา");

    body.replaceText("{{CLASS}}", classLevel === "all" ? "ทุกระดับชั้น" : classLevel);
    body.replaceText("{{SEMESTER}}", semester);
    body.replaceText("{{YEAR}}", academicYear);
    body.replaceText("{{TOTAL_COUNT}}", annualStudents.length.toString());
    body.replaceText("{{COUNT_S}}", rankCount.S.toString());
    body.replaceText("{{COUNT_A}}", rankCount.A.toString());
    body.replaceText("{{COUNT_B}}", rankCount.B.toString());
    body.replaceText("{{COUNT_C}}", rankCount.C.toString());
    body.replaceText("{{COUNT_D}}", rankCount.D.toString());
    body.replaceText("{{COUNT_E}}", rankCount.E.toString());
    body.replaceText("{{HEAD_TEACHER}}", headTeacher);

    doc.saveAndClose();
    const pdfBlob = newFile.getAs('application/pdf');
    const pdfFile = targetFolder.createFile(pdfBlob).setName(`AnnualReport_M${classLevel}_S${semester}_Y${academicYear}.pdf`);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    newFile.setTrashed(true);
    return { status: "success", url: pdfFile.getUrl() };
  } catch(e) {
    return { status: "error", message: "สร้าง PDF ภาพรวมไม่สำเร็จ: " + e.toString() };
  }
}