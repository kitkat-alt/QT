/**
 * ระบบ Backend สำหรับจัดการ Google Drive และ Google Sheets
 * โดย กิต
 */

// ฟังก์ชันเช็คและสร้าง Folder/Sheet อัตโนมัติในครั้งแรก
function getOrCreateSetup() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('MAIN_FOLDER_ID');
  let sheetId = props.getProperty('SPREADSHEET_ID');

  if (!folderId || !sheetId) {
    // 1. สร้างโฟลเดอร์หลักใน Drive
    const mainFolder = DriveApp.createFolder('ระบบออกเอกสาร_กิต');
    folderId = mainFolder.getId();
    props.setProperty('MAIN_FOLDER_ID', folderId);

    // 2. สร้าง Google Sheet
    const sheet = SpreadsheetApp.create('ฐานข้อมูล_ระบบเอกสาร');
    sheetId = sheet.getId();
    props.setProperty('SPREADSHEET_ID', sheetId);
    
    // 3. สร้างหัวตาราง
    sheet.getActiveSheet().appendRow(['ID เอกสาร', 'ชื่อลูกค้า', 'ประเภท', 'ยอดรวม', 'วันที่', 'ลิงก์ PDF']);
  }
  
  return { folderId, sheetId };
}

// ดึงประวัติและรันเลขที่เอกสารล่าสุด
function doGet(e) {
  const { sheetId } = getOrCreateSetup();
  const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const history = [];
  let maxIdNum = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = row[0];
    history.push({
      id: id,
      customerName: row[1],
      type: row[2],
      total: row[3],
      date: row[4],
      pdfUrl: row[5]
    });
    
    // หารันเลขล่าสุด (เช่น QT-005 -> 5)
    const numMatch = String(id).match(/\d+/);
    if (numMatch) {
      const num = parseInt(numMatch[0], 10);
      if (num > maxIdNum) maxIdNum = num;
    }
  }

  const nextId = 'QT-' + String(maxIdNum + 1).padStart(3, '0');

  const result = {
    status: 'success',
    data: history.reverse(), // เอาล่าสุดขึ้นก่อน
    nextId: nextId
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// รับข้อมูลจากหน้าเว็บเพื่อบันทึกไฟล์และข้อมูล
function doPost(e) {
  try {
    const { folderId, sheetId } = getOrCreateSetup();
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'uploadPDF') {
      const mainFolder = DriveApp.getFolderById(folderId);
      
      // จัดการโฟลเดอร์แยกตามลูกค้า
      let customerFolder;
      const folderIter = mainFolder.getFoldersByName(data.customerName || 'ลูกค้าทั่วไป');
      if (folderIter.hasNext()) {
        customerFolder = folderIter.next();
      } else {
        customerFolder = mainFolder.createFolder(data.customerName || 'ลูกค้าทั่วไป');
      }

      // บันทึก PDF
      const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), 'application/pdf', data.fileName);
      const file = customerFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      const fileUrl = file.getUrl();

      // อัปเดตหรือเพิ่มบรรทัดใหม่ใน Sheet
      const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
      const sheetData = sheet.getDataRange().getValues();
      let rowIndex = -1;
      
      for (let i = 1; i < sheetData.length; i++) {
        if (sheetData[i][0] === data.id) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex > -1) {
        sheet.getRange(rowIndex, 2, 1, 5).setValues([[data.customerName, data.type, data.total, data.date, fileUrl]]);
      } else {
        sheet.appendRow([data.id, data.customerName, data.type, data.total, data.date, fileUrl]);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: 'success', url: fileUrl }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
