// !! สำคัญ: ให้ใส่ Google Sheet ID ของคุณที่นี่
const SETUP_SHEET_ID = "1Mu3yzfF7hCd-dtGk-RJV8f-zu_xtjoW9AWpVqtmZY2E"; 

// --- SHEET NAMES ---
const SHEET_ACTIVITIES_NAME = "Activities";
const SHEET_TEAMS_NAME = "Teams";
const SHEET_FILES_NAME = "Files";
const SHEET_SCHOOLS_NAME = "Schools";
const SHEET_SCHOOL_CLUSTERS_NAME = "SchoolCluster";

/**
 * ฟังก์ชันหลักสำหรับรันเพื่อตั้งค่า Sheet ทั้งหมด
 * ให้เลือกฟังก์ชันนี้แล้วกด "Run" จากเมนู Apps Script
 */
function setupSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SETUP_SHEET_ID);
    if (!spreadsheet) {
      throw new Error("ไม่พบ Google Sheet. กรุณาตรวจสอบ SETUP_SHEET_ID");
    }
    
    // สร้าง Sheet กิจกรรม พร้อมข้อมูลตัวอย่าง
    createActivitiesSheet(spreadsheet);
    
    // สร้าง Sheet ทีม (เฉพาะ Header)
    createTeamsSheet(spreadsheet);
    
    // สร้าง Sheet ไฟล์ (เฉพาะ Header)
    createFilesSheet(spreadsheet);

    // สร้าง Sheet เครือข่ายโรงเรียน (ตัวอย่าง)
    createSchoolClusterSheet(spreadsheet);

    // สร้าง Sheet รายชื่อโรงเรียน (ตัวอย่าง)
    createSchoolsSheet(spreadsheet);
    
    SpreadsheetApp.flush();
    Logger.log("🎉 ตั้งค่า Sheet ทั้ง 4 สำเร็จ! (Activities, Teams, Files, Schools)");
    Browser.msgBox("🎉 ตั้งค่า Sheet ทั้ง 4 สำเร็จ!", "กรุณาตรวจสอบ Google Sheet ของคุณ", Browser.Buttons.OK);

  } catch (error) {
    Logger.log(error);
    Browser.msgBox("เกิดข้อผิดพลาด", error.message, Browser.Buttons.OK);
  }
}

/**
 * สร้างและตั้งค่า Sheet 'Activities'
 */
function createActivitiesSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_ACTIVITIES_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_ACTIVITIES_NAME);
  }
  
  sheet.clear();
  
  // ตั้งค่า Header
  const headers = [
    "ID",                   // A1
    "Category",             // B1
    "Name",                 // C1
    "Levels",               // D1
    "Mode",                 // E1
    "ReqTeachers",          // F1 
    "ReqStudents",          // G1
    "MaxTeams",             // H1
    "RegistrationDeadline"  // I1
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  
  // เพิ่มข้อมูลตัวอย่าง
  const sampleData = [
    [
      "act001", 
      "วิชาการ", 
      "แข่งขันตอบปัญหาวิทยาศาสตร์", 
      '["ป.4-ป.6", "ม.1-ม.3"]', 
      "Onsite",
      1, // ครู 1 คน
      3, // นักเรียน 3 คน
      8, // รับสูงสุด 8 ทีม
      new Date(2025, 0, 15, 17, 0, 0) // 15 ม.ค. 2025 17:00
    ],
    [
      "act002", 
      "ศิลปะ", 
      "แข่งขันวาดภาพระบายสี", 
      '["ป.1-ป.3"]', 
      "Onsite",
      1, // ครู 1 คน
      1, // นักเรียน 1 คน
      "", // ไม่จำกัดทีม (เว้นว่าง)
      new Date(2024, 10, 30, 12, 0, 0) // 30 พ.ย. 2024 12:00
    ],
    [
      "act003", 
      "เทคโนโลยี", 
      "แข่งขันเขียนโปรแกรม (Scratch)", 
      '["ป.4-ป.6"]', 
      "Online",
      1, // ครู 1 คน
      2, // นักเรียน 2 คน
      12, // รับสูงสุด 12 ทีม
      ""  // ไม่กำหนดวันปิดรับสมัคร
    ]
  ];
  
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`ตั้งค่า Sheet '${SHEET_ACTIVITIES_NAME}' พร้อมข้อมูลตัวอย่าง 3 รายการ`);
}

/**
 * สร้างและตั้งค่า Sheet 'Teams' (อัปเดต Header - เพิ่ม K: LogoUrl)
 */
function createTeamsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_TEAMS_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_TEAMS_NAME);
  }
  
  sheet.clear();
  
  // ตั้งค่า Header
  const headers = [
    "TeamID",           // A1
    "ActivityID",       // B1
    "TeamName",         // C1
    "School",           // D1
    "Level",            // E1
    "Contact",          // F1 (JSON)
    "Members",          // G1 (JSON)
    "RequiredTeachers", // H1
    "RequiredStudents", // I1
    "Status",           // J1
    "LogoUrl",          // K1 (เก็บ File ID โลโก้ทีม)
    "TeamPhotoId",      // L1 (เก็บ File ID รูปทีม)
    "CreatedByUserId",  // M1 (รหัสผู้ใช้ที่สร้างทีม)
    "CreatedByUsername" // N1 (ชื่อผู้ใช้ที่สร้างทีม)
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`ตั้งค่า Sheet '${SHEET_TEAMS_NAME}' (เฉพาะ Header)`);
}

/**
 * สร้างและตั้งค่า Sheet 'Files'
 */
function createFilesSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_FILES_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_FILES_NAME);
  }
  
  sheet.clear();
  
  // ตั้งค่า Header
  const headers = [
    "FileLogID",    // A1
    "TeamID",       // B1
    "FileType",     // C1
    "Status",       // D1
    "FileUrl",      // E1
    "Remarks",      // F1
    "FileDriveId"   // G1
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`ตั้งค่า Sheet '${SHEET_FILES_NAME}' (เฉพาะ Header)`);
}

/**
 * สร้างและตั้งค่า Sheet 'SchoolCluster'
 */
function createSchoolClusterSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_SCHOOL_CLUSTERS_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_SCHOOL_CLUSTERS_NAME);
  }

  sheet.clear();

  const headers = ["SchoolClusterID", "ClusterName"];
  const clusters = [
    ["CL001", "ภาคเหนือ"],
    ["CL002", "ภาคกลาง"],
    ["CL003", "ภาคใต้"],
    ["CL004", "กรุงเทพฯ"],
    ["CL005", "ภาคเหนือตอนล่าง"]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  if (clusters.length) {
    sheet.getRange(2, 1, clusters.length, headers.length).setValues(clusters);
  }
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`ตั้งค่า Sheet '${SHEET_SCHOOL_CLUSTERS_NAME}' พร้อมข้อมูลเครือข่าย ${clusters.length} รายการ`);
}

/**
 * สร้างและตั้งค่า Sheet 'Schools'
 */
function createSchoolsSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_SCHOOLS_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_SCHOOLS_NAME);
  }

  sheet.clear();

  const headers = ["SchoolID", "SchoolName", "SchoolCluster"];
  const schools = [
    ["SCH001", "โรงเรียนสาธิตมหาวิทยาลัยเชียงใหม่", "CL001"],
    ["SCH002", "โรงเรียนเตรียมอุดมศึกษา", "CL002"],
    ["SCH003", "โรงเรียนสวนกุหลาบวิทยาลัย", "CL002"],
    ["SCH004", "โรงเรียนเบญจมราชูทิศ", "CL003"],
    ["SCH005", "โรงเรียนพิษณุโลกพิทยาคม", "CL001"],
    ["SCH006", "โรงเรียนนครสวรรค์", "CL001"],
    ["SCH007", "โรงเรียนหาดใหญ่วิทยาคาร", "CL003"],
    ["SCH008", "โรงเรียนอนุบาลภูเก็ต", "CL003"],
    ["SCH009", "โรงเรียนสารวิทยา", "CL004"],
    ["SCH010", "โรงเรียนกรุงเทพคริสเตียนวิทยาลัย", "CL004"]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.getRange(2, 1, schools.length, headers.length).setValues(schools);
  sheet.autoResizeColumns(1, headers.length);
  Logger.log(`ตั้งค่า Sheet '${SHEET_SCHOOLS_NAME}' พร้อมรายชื่อโรงเรียนตัวอย่าง ${schools.length} รายการ`);
}
