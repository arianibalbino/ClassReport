function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Daily Class Report');
}

// ID of the root folder where reports will be saved
const ROOT_FOLDER_ID = '1CM15DCyL0W-RRxpW-ld7Pz-cdNYOSKza';

function submitForm(formData) {
  try {
    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const schoolFolder = getOrCreateFolder(formData.school, rootFolder);

    // ISO Week
    const now = new Date();
    const weekNumber = getWeekNumber(now);
    const weekFolderName = `${now.getFullYear()}-W${weekNumber}`;
    const weekFolder = getOrCreateFolder(weekFolderName, schoolFolder);

    const sheetName = `Report_${formData.school}_${weekFolderName}`;
    const ss = getOrCreateSpreadsheet(sheetName, weekFolder);

    const sheetNameDaily = `${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')} - ${formData.teacher}`;
    let sheet = ss.getSheetByName(sheetNameDaily);
    if (!sheet) {
      sheet = ss.insertSheet(sheetNameDaily);
      sheet.appendRow([
        'Timestamp', 'Teacher', 'School', 'Grade', 'Subject', 'Content',
        'Number of Students', 'Notes', 'Students (Name - Observation)', 'Photos (links)'
      ]);
    }

    // Photo upload (if any)
    let photoLinks = [];
    if (formData.photos && formData.photos.length > 0) {
      const photoFolder = getOrCreateFolder('Photos', weekFolder);
      for (let f of formData.photos) {
        const blob = Utilities.base64Decode(f.data);
        const file = photoFolder.createFile(Utilities.newBlob(blob, 'image/jpeg', f.name));
        file.setDescription(`Uploaded by ${formData.teacher} on ${new Date()}`);
        photoLinks.push(file.getUrl());
      }
    }

    const studentData = formData.students
      .map(s => `${s.name} - ${s.observation}`)
      .join('; ');

    sheet.appendRow([
      new Date(),
      formData.teacher,
      formData.school,
      formData.grade,
      formData.subject,
      formData.content,
      formData.numStudents,
      formData.notes,
      studentData,
      photoLinks.join('\n')
    ]);

    return { success: true, message: '✅ Form and photos sent successfully!' };

  } catch (err) {
    Logger.log('Error in submitForm: ' + err);
    return { success: false, message: '❌ Error sending: ' + err };
  }
}


// === Helper functions ===
function getOrCreateFolder(name, parent) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getOrCreateSpreadsheet(name, folder) {
  const files = folder.getFilesByName(name);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  return ss;
}

function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}
