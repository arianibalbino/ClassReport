function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Class Report Form');
}

// Root folder containing both "Schools Database" and "Reports"
const ROOT_FOLDER_ID = '1no9PS0VpnlqhkZTBVD0HMVfi-L1q6KjL';

// ============================================================
// Debug function to check folder structure
// ============================================================
function debugGetRegions() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  Logger.log("Root folder: " + root.getName());

  const dbFolder = getSubFolderByName(root, 'Schools Database');
  if (!dbFolder) {
    Logger.log("❌ 'Schools Database' not found inside root folder.");
    const all = root.getFolders();
    while (all.hasNext()) Logger.log("Found: " + all.next().getName());
    return;
  }

  const regions = getSubfolderNames(dbFolder);
  Logger.log("✅ Regions found: " + JSON.stringify(regions));
}

// ============================================================
// Load Bindi logo from Drive and return as Base64 data URL
// ============================================================
function getLogoUrl() {
  try {
    const fileId = '1Nc_bM1zbNFfe2MQAAir7kB19NLcPDOJq'; // seu ID da logo
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const contentType = blob.getContentType();
    return `data:${contentType};base64,${base64}`;
  } catch (err) {
    Logger.log('❌ Error loading logo: ' + err);
    return '';
  }
}


// ============================================================
// Get regions, cities, schools, teachers, and students
// ============================================================
function getRegions() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const dbFolder = getSubFolderByName(root, 'Schools Database');
  if (!dbFolder) return [];
  return getSubfolderNames(dbFolder);
}

function getCities(region) {
  const dbFolder = getSchoolsDatabaseFolder();
  const regionFolder = getSubFolderByName(dbFolder, region);
  if (!regionFolder) return [];
  return getSubfolderNames(regionFolder);
}

function getSchools(region, city) {
  const dbFolder = getSchoolsDatabaseFolder();
  const regionFolder = getSubFolderByName(dbFolder, region);
  if (!regionFolder) return [];
  const cityFolder = getSubFolderByName(regionFolder, city);
  if (!cityFolder) return [];
  const files = cityFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const schools = [];
  while (files.hasNext()) schools.push(files.next().getName());
  return schools;
}

function getTeachers(region, city, school) {
  const ss = openSchoolSheet(region, city, school);
  if (!ss) return [];

  const teacherSheet = ss.getSheetByName("Teacher's name");
  if (!teacherSheet) return [];

  const data = teacherSheet.getRange(2, 1, teacherSheet.getLastRow() - 1, 1).getValues();
  return data.flat().filter(n => n && n.toString().trim() !== "");
}

function getGroups(region, city, school) {
  const ss = openSchoolSheet(region, city, school);
  if (!ss) return [];

  const groupSheet = ss.getSheetByName("Group Class");
  if (!groupSheet) return [];

  // Pega os títulos das colunas (linha 1)
  const headers = groupSheet.getRange(1, 1, 1, groupSheet.getLastColumn()).getValues()[0];
  return headers.filter(n => n && n.toString().trim() !== "");
}

function getStudents(region, city, school, group) {
  const ss = openSchoolSheet(region, city, school);
  if (!ss) return [];

  const groupSheet = ss.getSheetByName("Group Class");
  if (!groupSheet) return [];

  const headers = groupSheet.getRange(1, 1, 1, groupSheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(group) + 1;
  if (colIndex < 1) return [];

  const data = groupSheet.getRange(2, colIndex, groupSheet.getLastRow() - 1, 1).getValues();
  return data.flat().filter(n => n && n.toString().trim() !== "");
}

// ============================================================
// Save report (creates or updates the teacher's tab)
// ============================================================
function submitForm(formData, photos, students) {
  try {
    const { region, city, school, teacher, classDate, grade, subject, notes, gradeType, quarter, month } = formData;

    // 📁 Estrutura de pastas
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const reportsFolder = getOrCreateSubFolder(root, 'Reports');
    const regionFolder = getOrCreateSubFolder(reportsFolder, region);
    const cityFolder = getOrCreateSubFolder(regionFolder, city);
    const schoolFolder = getOrCreateSubFolder(cityFolder, school);

    const dateObj = new Date(classDate || new Date());
    const week = getWeekNumber(dateObj);
    const year = dateObj.getFullYear();
    const weekFolder = getOrCreateSubFolder(schoolFolder, `${year}-W${week}`);

    const reportName = `Assessment_Report_${school.replace(/\s+/g, '')}_W${week}`;
    const existing = weekFolder.getFilesByName(reportName);
    const ss = existing.hasNext() ? SpreadsheetApp.open(existing.next()) : SpreadsheetApp.create(reportName);
    DriveApp.getFileById(ss.getId()).moveTo(weekFolder);

    // 📊 Nome da aba = Professor + Matéria
    const sheetName = `${teacher}_${subject}`;
    let sh = ss.getSheetByName(sheetName);

    if (!sh) {
      sh = ss.insertSheet(sheetName);
      const headers = [
        'Date', 'Region', 'City', 'School', 'Class/Grade', 'Subject',
        'Assessment Type', 'Quarter/Month', 'General Notes',
        'Student Name', 'Assessment', 'Photo Links'
      ];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setHorizontalAlignment('center');
    }

    // 📸 Upload das fotos
    const photoLinks = [];
    if (photos && photos.length) {
      const photosFolder = getOrCreateSubFolder(weekFolder, 'Photos');
      const dateFolder = getOrCreateSubFolder(photosFolder, classDate);
      photos.forEach(p => {
        try {
          const bytes = Utilities.base64Decode(p.data.split(',')[1]);
          const file = dateFolder.createFile(Utilities.newBlob(bytes, 'image/jpeg', p.name));
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoLinks.push({ name: p.name, url: file.getUrl() });
        } catch (e) {
          Logger.log("⚠️ Error saving photo " + p.name + ": " + e);
        }
      });
    }

    // 🧮 Determina período
    const period = gradeType === "Quarterly" ? quarter : month;

    // 🧾 Linhas novas (agora incluindo turma/grade)
    const rows = (students && students.length > 0)
      ? students.map(s => [
          classDate, region, city, school, grade, subject,
          gradeType, period, notes, s.name, s.grade, photoLinks.length ? photoLinks[0].url : ''
        ])
      : [[classDate, region, city, school, grade, subject, gradeType, period, notes, '', '', '']];

    Logger.log(`📋 ${rows.length} rows ready to append in ${sheetName}`);

    // 🔄 Adiciona as novas linhas no final (sem sobrescrever)
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);

    // ✨ Formatação automática e alinhamento
    const totalCols = sh.getLastColumn();
    const totalRows = sh.getLastRow();

    // Auto-ajusta a largura de todas as colunas (como “Ajustar ao texto”)
    sh.autoResizeColumns(1, totalCols);

    // Define o alinhamento à esquerda (inclusive cabeçalhos)
    sh.getRange(1, 1, totalRows, totalCols).setHorizontalAlignment('left');

    // Congela a primeira linha (cabeçalho)
    sh.setFrozenRows(1);


    // 🔍 Adiciona filtros automáticos (somente 1x)
    if (!sh.getFilter()) {
      const dataRange = sh.getDataRange();
      sh.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns()).createFilter();
      Logger.log(`✅ Filter added to sheet ${sheetName}`);
    }

    Logger.log(`✅ Data successfully appended to ${sheetName}`);
    return { success: true };

  } catch (e) {
    Logger.log("❌ submitForm error: " + e);
    return { success: false, message: e.message };
  }
}



// ============================================================
// Utility functions
// ============================================================
function getSchoolsDatabaseFolder() {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  return getSubFolderByName(root, 'Schools Database');
}

function openSchoolSheet(region, city, school) {
  const dbFolder = getSchoolsDatabaseFolder();
  const regionFolder = getSubFolderByName(dbFolder, region);
  if (!regionFolder) return null;
  const cityFolder = getSubFolderByName(regionFolder, city);
  if (!cityFolder) return null;
  const files = cityFolder.getFilesByName(school);
  return files.hasNext() ? SpreadsheetApp.open(files.next()) : null;
}

function getSubFolderByName(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

function getOrCreateSubFolder(parent, name) {
  const existing = getSubFolderByName(parent, name);
  return existing || parent.createFolder(name);
}

function getSubfolderNames(parent) {
  const arr = [];
  const sub = parent.getFolders();
  while (sub.hasNext()) arr.push(sub.next().getName());
  return arr;
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

// ============================================================
// ✅ AUTO-GENERATED REPORTS SECTION
// ============================================================

// Create Attendance Report for the week folder
function generateAttendanceReport(schoolFolder) {
  try {
    const reports = [];
    const weekFolders = schoolFolder.getFolders();
    while (weekFolders.hasNext()) {
      const weekFolder = weekFolders.next();
      const files = weekFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        const ss = SpreadsheetApp.open(files.next());
        ss.getSheets().forEach(sh => {
          const data = sh.getDataRange().getValues();
          if (data.length > 1) {
            const headers = data[0];
            const idx = {
              date: headers.indexOf("Date"),
              student: headers.indexOf("Student Name"),
              attend: headers.indexOf("Attendance")
            };
            data.slice(1).forEach(r => {
              reports.push([
                r[idx.date] || "",
                r[idx.student] || "",
                r[idx.attend] || "",
                sh.getName(), // sheet = date_teacher
                ss.getName()  // report file name
              ]);
            });
          }
        });
      }
    }

    if (!reports.length) return;

    const headers = ["Date", "Student Name", "Attendance", "Sheet", "Source File"];
    const today = new Date().toISOString().split("T")[0];
    const tempSS = SpreadsheetApp.create(`Attendance_Report_${today}`);
    const temp = tempSS.getActiveSheet();
    temp.getRange(1, 1, 1, headers.length).setValues([headers]);
    temp.getRange(2, 1, reports.length, headers.length).setValues(reports);
    temp.autoResizeColumns(1, headers.length);

    const blob = DriveApp.getFileById(tempSS.getId())
      .getAs("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const excelFile = schoolFolder.createFile(blob);
    excelFile.setName(`Attendance_Report-${today}.xlsx`);
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    Logger.log("✅ Attendance report generated: " + excelFile.getName());
  } catch (err) {
    Logger.log("❌ generateAttendanceReport error: " + err);
  }
}


// Create Grades Report for the week folder
function generateGradesReport(schoolFolder) {
  try {
    const grades = [];
    const weekFolders = schoolFolder.getFolders();
    while (weekFolders.hasNext()) {
      const weekFolder = weekFolders.next();
      const files = weekFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        const ss = SpreadsheetApp.open(files.next());
        ss.getSheets().forEach(sh => {
          const data = sh.getDataRange().getValues();
          if (data.length > 1) {
            const headers = data[0];
            const idx = {
              date: headers.indexOf("Date"),
              teacher: headers.indexOf("Teacher") !== -1 ? headers.indexOf("Teacher") : headers.indexOf("Teacher Name"),
              subject: headers.indexOf("Subject"),
              student: headers.indexOf("Student Name"),
              note: headers.indexOf("Note/Grade"),
              school: headers.indexOf("School")
            };
            data.slice(1).forEach(r => {
              grades.push([
                r[idx.date] || "",
                r[idx.teacher] || "",
                r[idx.subject] || "",
                r[idx.student] || "",
                r[idx.note] || "",
                r[idx.school] || "",
                sh.getName(),
                ss.getName()
              ]);
            });
          }
        });
      }
    }

    if (!grades.length) return;

    const headers = ["Date", "Teacher Name", "Subject", "Student Name", "Grade", "School", "Sheet", "Source File"];
    const today = new Date().toISOString().split("T")[0];
    const tempSS = SpreadsheetApp.create(`Grades_Report_${today}`);
    const temp = tempSS.getActiveSheet();
    temp.getRange(1, 1, 1, headers.length).setValues([headers]);
    temp.getRange(2, 1, grades.length, headers.length).setValues(grades);
    temp.autoResizeColumns(1, headers.length);

    const blob = DriveApp.getFileById(tempSS.getId())
      .getAs("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const excelFile = schoolFolder.createFile(blob);
    excelFile.setName(`Grades_Report-${today}.xlsx`);
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);

    Logger.log("✅ Grades report generated: " + excelFile.getName());
  } catch (err) {
    Logger.log("❌ generateGradesReport error: " + err);
  }
}

function getSubjects(region, city, school) {
  try {
    const ss = openSchoolSheet(region, city, school);
    if (!ss) return [];

    const subjSheet = ss.getSheetByName("Subjects");
    if (!subjSheet) {
      Logger.log(`❌ No "Subjects" tab found in ${school}`);
      return [];
    }

    const data = subjSheet.getRange(2, 1, subjSheet.getLastRow() - 1, 1).getValues();
    const subjects = data.flat().filter(v => v && v.toString().trim() !== "");
    Logger.log(`✅ Subjects for ${school}: ${subjects}`);
    return subjects;
  } catch (err) {
    Logger.log("❌ getSubjects error: " + err);
    return [];
  }
}



