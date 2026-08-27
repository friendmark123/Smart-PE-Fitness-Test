// =============================================================================
// MAIN ENTRYPOINT & HTMLSERVICE WRAPPER
// =============================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบสมรรถภาพทางกาย - บดินทรเดชา ๒')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Wrapper Functions สำหรับเรียกจาก Client (google.script.run)
function getYearListWrapper() {
  return getYearList();
}

function searchStudentDataWrapper(studentId, semester, academicYear) {
  return searchStudentData(studentId, semester, academicYear);
}

function searchFitnessHistoryWrapper(studentId, semester, academicYear) {
  return searchFitnessHistory(studentId, semester, academicYear);
}

function saveAndCalcWrapper(f) {
  return saveAndCalc(f);
}

function getAdminDataWrapper(f) {
  return getAdminData(f);
}

function getPendingStudentsWrapper(f) {
  return getPendingStudents(f);
}

function deleteFitnessRecordWrapper(studentId, semester, academicYear) {
  return deleteFitnessRecord(studentId, semester, academicYear);
}

function generateIndividualPDFWrapper(studentId, semester, academicYear) {
  return generateIndividualPDF(studentId, semester, academicYear);
}

function generateClassPDFWrapper(classLevel, room, semester, academicYear) {
  return generateClassPDF(classLevel, room, semester, academicYear);
}

function generateAnnualPDFWrapper(classLevel, semester, academicYear) {
  return generateAnnualPDF(classLevel, semester, academicYear);
}