/**
 * สคริปต์ตั้งค่า Sheet Users
 * รันฟังก์ชัน setupUserSheet() ครั้งเดียวเพื่อสร้าง/รีเซ็ตโครงสร้าง
 */
function setupUserSheet() {
  const spreadsheetId = typeof SETUP_SHEET_ID !== "undefined" ? SETUP_SHEET_ID : SHEET_ID;
  if (!spreadsheetId) {
    throw new Error("กรุณากำหนดค่า SETUP_SHEET_ID หรือ SHEET_ID ก่อน");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetName = typeof SHEET_USERS !== "undefined" ? SHEET_USERS : "Users";
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  sheet.clear();

  const headers = [
    "userid",
    "username",
    "password",
    "name",
    "surname",
    "SchoolID",
    "tel",
    "userline_id",
    "level",
    "email",
    "avatarFileId"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(`✅ สร้าง Sheet '${sheetName}' สำเร็จ`);
  try {
    Browser.msgBox(`✅ สร้าง Sheet '${sheetName}' สำเร็จ`, Browser.Buttons.OK);
  } catch (e) {
    // ในบางสภาพแวดล้อม (เช่น trigger) จะเรียก Browser.msgBox ไม่ได้ จึงปล่อยผ่าน
  }
}
