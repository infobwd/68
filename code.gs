// --- CONFIGURATION ---
const DRIVE_FOLDER_ID = "1rqv8_Uh9SqmvLjsY--9CRwYRPYBCyjAD";
const SHEET_ID = "1Mu3yzfF7hCd-dtGk-RJV8f-zu_xtjoW9AWpVqtmZY2E";
// !! ⬇️ สำคัญ: ใส่อีเมลของคุณที่เป็น Admin ⬇️
const ADMIN_EMAIL = "noppharutlubbuangam@gmail.com";

// --- SHEET NAMES ---
const SHEET_ACTIVITIES = "Activities";
const SHEET_TEAMS = "Teams";
const SHEET_FILES = "Files";
const SHEET_SCHOOLS = "Schools";
const SHEET_USERS = "Users";
const SHEET_SCHOOL_CLUSTER = "SchoolCluster";

const LINE_LIFF_ID = "2006490627-84dBRzwJ";

const USER_LEVELS = ["User", "Admin", "Score", "Area", "School_Admin", "Group_Admin"];
const DEFAULT_USER_LEVEL = USER_LEVELS[0];
const USER_SHEET_HEADERS = [
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

/**
 * ให้บริการหน้าเว็บหลักเมื่อมีการเรียก GET
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("ระบบลงทะเบียนแข่งขัน")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * ดึงข้อมูล HTML จากไฟล์อื่นnoppharut
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper: แปลงค่าจำนวนทีมสูงสุดจาก Sheet ให้เป็นตัวเลข หรือ null (ไม่จำกัด)
 */
function parseMaxTeamsValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number") {
    return value > 0 ? value : null;
  }
  const cleaned = String(value).trim();
  if (cleaned === "" || cleaned.toLowerCase() === "ไม่จำกัด") {
    return null;
  }
  const parsed = parseInt(cleaned, 10);
  return isNaN(parsed) || parsed <= 0 ? null : parsed;
}

/**
 * Helper: แปลงค่ากำหนดปิดรับสมัครให้เป็น Date หรือ null
 */
function parseDeadlineValue(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date) {
    return new Date(value.getTime());
  }
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function safeParseJson(value, fallback) {
  if (value === null || value === undefined || value === "") {
    return fallback;
  }
  if (typeof value === "object") {
    return value;
  }
  try {
    return JSON.parse(value);
  } catch (error) {
    return fallback;
  }
}

function escapeCsvCell(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const stringValue = value.toString().replace(/"/g, '""');
  return /[",\n]/.test(stringValue) ? `"${stringValue}"` : stringValue;
}

function escapeHtml(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return value
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function isValidEmail(value) {
  if (!value) {
    return false;
  }
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value).trim());
}

function buildDriveViewUrl(fileId) {
  if (!fileId) {
    return "";
  }
  return "https://drive.google.com/uc?export=view&id=" + encodeURIComponent(fileId);
}

function hashPassword(password) {
  if (!password) {
    return "";
  }
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  return bytesToHex_(digest);
}

function bytesToHex_(bytes) {
  return bytes
    .map(byte => {
      const value = byte < 0 ? byte + 256 : byte;
      return ("0" + value.toString(16)).slice(-2);
    })
    .join("");
}

function generateUserId() {
  const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMddHHmmss");
  const random = Math.floor(Math.random() * 1000)
    .toString()
    .padStart(3, "0");
  return `USR${timestamp}${random}`;
}

function getUsersSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_USERS);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_USERS);
  }
  ensureUsersSheetHeader_(sheet);
  return sheet;
}

function ensureUsersSheetHeader_(sheet) {
  const neededCols = USER_SHEET_HEADERS.length;
  const headerRange = sheet.getRange(1, 1, 1, neededCols);
  const existing = headerRange.getValues()[0];
  const mismatch = existing.some((value, index) => value !== USER_SHEET_HEADERS[index]);
  if (mismatch) {
    headerRange.setValues([USER_SHEET_HEADERS]).setFontWeight("bold");
  }
}

function getAllUserRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, USER_SHEET_HEADERS.length).getValues();
}

function sanitizeUserRow(row) {
  const record = {};
  USER_SHEET_HEADERS.forEach((header, index) => {
    record[header] = row[index] !== undefined ? row[index] : "";
  });
  return record;
}

function publicUserProfile(user) {
  const result = Object.assign({}, user);
  delete result.password;
  result.schoolId = result.SchoolID || "";
  result.lineUserId = result.userline_id || "";
  result.email = result.email || "";
  result.tel = result.tel || "";
  result.avatarFileId = result.avatarFileId || "";
  if (!result.avatarUrl && result.avatarFileId) {
    result.avatarUrl = buildDriveViewUrl(result.avatarFileId);
  }
  return result;
}

function buildRowFromUser(user) {
  return USER_SHEET_HEADERS.map(header => {
    const value = user[header];
    return value === undefined || value === null ? "" : value;
  });
}

function updateUserRow(sheet, rowIndex, user) {
  const rowValues = buildRowFromUser(user);
  sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
}

function ensureValidUserLevel(level) {
  return USER_LEVELS.includes(level) ? level : DEFAULT_USER_LEVEL;
}

function resolveNameParts_(firstName, surname) {
  let resolvedFirst = (firstName || "").toString().trim();
  let resolvedLast = (surname || "").toString().trim();
  if (!resolvedLast && resolvedFirst.includes(" ")) {
    const parts = resolvedFirst.split(/\s+/);
    resolvedFirst = parts.shift();
    resolvedLast = parts.join(" ");
  }
  return {
    firstName: resolvedFirst,
    surname: resolvedLast
  };
}

// function normalizeKey(value) {
//   return (value || "").toString().trim().toLowerCase();
// }

function normalizeKey(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}


/**
 * อ่านข้อมูลเครือข่ายจากชีต SchoolCluster
 * คอลัมน์: A=SchoolClusterID, B=ClusterName
 */
// function buildSchoolClusterMap_(spreadsheet) {
//   var result = {
//     byId: {},
//     byName: {}
//   };

//   try {
//     var book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
//     var sheet = book.getSheetByName(SHEET_SCHOOL_CLUSTER);
//     if (!sheet) {
//       return result;
//     }

//     var lastRow = sheet.getLastRow();
//     if (lastRow < 2) {
//       return result;
//     }

//     var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A=ID, B=Name
//     for (var i = 0; i < values.length; i++) {
//       var id = String(values[i][0] || "").trim();
//       var name = String(values[i][1] || "").trim();
//       if (!id && !name) {
//         continue;
//       }

//       var label = name || id;
//       if (id) {
//         result.byId[normalizeKey(id)] = {
//           id: id,
//           name: label
//         };
//       }
//       if (name) {
//         result.byName[normalizeKey(name)] = {
//           id: id || name,
//           name: name
//         };
//       }
//     }
//   } catch (e) {
//     Logger.log("buildSchoolClusterMap_ error: " + e);
//   }

//   return result;
// }

function buildSchoolClusterMap_(spreadsheet) {
  var result = { byId: {}, byName: {} };
  try {
    var book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
    var sheet = book.getSheetByName(SHEET_SCHOOL_CLUSTER);
    if (!sheet) return result;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return result;

    var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      var id = String(values[i][0] || "").trim();
      var name = String(values[i][1] || "").trim();
      if (!id && !name) continue;
      var label = name || id;
      if (id)   result.byId[normalizeKey(id)] = { id: id, name: label };
      if (name) result.byName[normalizeKey(name)] = { id: id || name, name: name };
    }
  } catch (e) {
    Logger.log("buildSchoolClusterMap_ error: " + e);
  }
  return result;
}


function findUserByUsername(username) {
  const key = normalizeKey(username);
  if (!key) {
    return null;
  }
  const sheet = getUsersSheet();
  const rows = getAllUserRows_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (normalizeKey(row[1]) === key) {
      return {
        rowIndex: i + 2,
        user: sanitizeUserRow(row)
      };
    }
  }
  return null;
}

function findUserByLineId(lineUserId) {
  const key = normalizeKey(lineUserId);
  if (!key) {
    return null;
  }
  const sheet = getUsersSheet();
  const rows = getAllUserRows_(sheet);
  const lineColumnIndex = USER_SHEET_HEADERS.indexOf("userline_id");
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (normalizeKey(row[lineColumnIndex]) === key) {
      return {
        rowIndex: i + 2,
        user: sanitizeUserRow(row)
      };
    }
  }
  return null;
}

function findUserById(userId) {
  const targetId = (userId || "").toString().trim();
  if (!targetId) {
    return null;
  }
  const sheet = getUsersSheet();
  const rows = getAllUserRows_(sheet);
  const idIndex = USER_SHEET_HEADERS.indexOf("userid");
  if (idIndex === -1) {
    return null;
  }
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const currentId = (row[idIndex] || "").toString().trim();
    if (currentId && currentId === targetId) {
      return {
        rowIndex: i + 2,
        user: sanitizeUserRow(row)
      };
    }
  }
  return null;
}

// --- ADMIN & SECURITY FUNCTIONS ---

/**
 * ตรวจสอบว่าเป็น Admin หรือไม่
 */
function isAdmin() {
  try {
    return Session.getActiveUser().getEmail() === ADMIN_EMAIL;
  } catch (e) {
    return false;
  }
}

/**
 * (ใช้สำหรับ Client-side) ดึงอีเมลผู้ใช้
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (e) {
    return null;
  }
}

// --- USER & AUTH FUNCTIONS ---

function getLineConfig() {
  return {
    success: true,
    liffId: LINE_LIFF_ID,
    levels: USER_LEVELS
  };
}

function registerNormalUser(payload) {
  try {
    const data = payload || {};
    const username = (data.username || "").toString().trim();
    const password = (data.password || "").toString();
    const name = (data.name || "").toString().trim();
    const surname = (data.surname || "").toString().trim();
    const schoolId = (data.schoolId || data.SchoolID || "").toString().trim();
    const tel = (data.tel || data.phone || "").toString().trim();
    const email = (data.email || "").toString().trim();

    if (!username) {
      throw new Error("กรุณากรอกชื่อผู้ใช้");
    }
    if (!password) {
      throw new Error("กรุณากรอกรหัสผ่าน");
    }
    if (!email) {
      throw new Error("กรุณากรอกอีเมล");
    }
    if (!isValidEmail(email)) {
      throw new Error("รูปแบบอีเมลไม่ถูกต้อง");
    }
    if (!tel) {
      throw new Error("กรุณากรอกเบอร์ติดต่อ");
    }

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
    } catch (error) {
      throw new Error("ระบบกำลังใช้งานอยู่ กรุณาลองใหม่อีกครั้ง");
    }

    try {
      const existing = findUserByUsername(username);
      if (existing) {
        throw new Error("ชื่อผู้ใช้ถูกใช้แล้ว");
      }
      const sheet = getUsersSheet();
      const userId = generateUserId();
      let avatarFileId = "";
      if (data.avatarFileData && data.avatarFileData.base64Data) {
        avatarFileId = saveAvatarFile_(data.avatarFileData, userId) || "";
      }
      const newUser = {
        userid: userId,
        username: username,
        password: hashPassword(password),
        name: name,
        surname: surname,
        SchoolID: schoolId,
        tel: tel,
        userline_id: "",
        level: DEFAULT_USER_LEVEL,
        email: email,
        avatarFileId: avatarFileId
      };
      sheet.appendRow(buildRowFromUser(newUser));
      return {
        success: true,
        user: publicUserProfile(newUser)
      };
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log("registerNormalUser error: " + error);
    return {
      success: false,
      error: error.message
    };
  }
}

function registerLineUser(payload) {
  try {
    const data = payload || {};
    const lineUserId = (data.lineUserId || data.userline_id || "").toString().trim();
    if (!lineUserId) {
      throw new Error("กรุณาระบุ userline_id");
    }
    const usernameCandidate = (data.username || data.email || lineUserId).toString().trim();
    const usernameCandidateKey = normalizeKey(usernameCandidate);
    const displayName = (data.displayName || data.name || "").toString().trim();
    const surnameInput = (data.surname || "").toString().trim();
    const { firstName, surname } = resolveNameParts_(displayName, surnameInput);
    const schoolId = (data.schoolId || data.SchoolID || "").toString().trim();
    const tel = (data.tel || "").toString().trim();
    const emailInput = (data.email || data.emailAddress || "").toString().trim();
    const sanitizedEmail = isValidEmail(emailInput) ? emailInput : "";

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
    } catch (error) {
      throw new Error("ระบบกำลังใช้งานอยู่ กรุณาลองใหม่อีกครั้ง");
    }

    try {
      const sheet = getUsersSheet();
      const existingLine = findUserByLineId(lineUserId);
      if (existingLine) {
        const userData = existingLine.user;
        if (usernameCandidate) {
          const currentKey = normalizeKey(userData.username);
          if (usernameCandidateKey && usernameCandidateKey !== currentKey) {
            const duplicateUser = findUserByUsername(usernameCandidate);
            if (duplicateUser && duplicateUser.user.userid !== userData.userid) {
              throw new Error("ชื่อผู้ใช้ถูกใช้แล้ว");
            }
            userData.username = usernameCandidate;
          }
        } else if (!userData.username) {
          userData.username = lineUserId;
        }
        const resolvedNames = resolveNameParts_(firstName || userData.name, surname || userData.surname);
        userData.name = resolvedNames.firstName || userData.name;
        userData.surname = resolvedNames.surname || userData.surname;
        userData.SchoolID = schoolId || userData.SchoolID;
        userData.tel = tel || userData.tel;
        if (sanitizedEmail) {
          userData.email = sanitizedEmail;
        }
        if (data.avatarFileData && data.avatarFileData.base64Data) {
          userData.avatarFileId = saveAvatarFile_(data.avatarFileData, userData.userid) || userData.avatarFileId;
        }
        userData.userline_id = lineUserId;
        userData.level = ensureValidUserLevel(userData.level);
        updateUserRow(sheet, existingLine.rowIndex, userData);
        return {
          success: true,
          user: publicUserProfile(userData),
          mode: "updated"
        };
      }

      const usernameToStore = usernameCandidate || lineUserId;
      const duplicate = findUserByUsername(usernameToStore);
      if (duplicate) {
        throw new Error("ชื่อผู้ใช้ถูกใช้แล้ว");
      }

      const newUserId = generateUserId();
      let avatarFileId = "";
      if (data.avatarFileData && data.avatarFileData.base64Data) {
        avatarFileId = saveAvatarFile_(data.avatarFileData, newUserId) || "";
      }

      const newUser = {
        userid: newUserId,
        username: usernameToStore,
        password: "",
        name: firstName,
        surname: surname,
        SchoolID: schoolId,
        tel: tel,
        userline_id: lineUserId,
        level: DEFAULT_USER_LEVEL,
        email: sanitizedEmail,
        avatarFileId: avatarFileId
      };
      sheet.appendRow(buildRowFromUser(newUser));
      return {
        success: true,
        user: publicUserProfile(newUser),
        mode: "created"
      };
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log("registerLineUser error: " + error);
    return {
      success: false,
      error: error.message
    };
  }
}

function loginUser(credentials) {
  try {
    const data = credentials || {};
    const username = (data.username || "").toString().trim();
    const password = (data.password || "").toString();

    if (!username || !password) {
      throw new Error("กรุณากรอกชื่อผู้ใช้และรหัสผ่าน");
    }

    const record = findUserByUsername(username);
    if (!record) {
      throw new Error("ไม่พบผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
    }

    const hashed = hashPassword(password);
    if (hashed !== (record.user.password || "")) {
      throw new Error("ไม่พบผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
    }

    return {
      success: true,
      user: publicUserProfile(record.user)
    };
  } catch (error) {
    Logger.log("loginUser error: " + error);
    return {
      success: false,
      error: error.message
    };
  }
}

function loginWithLine(payload) {
  try {
    const data = payload || {};
    const lineUserId = (data.lineUserId || data.userline_id || "").toString().trim();
    if (!lineUserId) {
      throw new Error("กรุณาระบุ userline_id");
    }
    const record = findUserByLineId(lineUserId);
    if (!record) {
      return {
        success: false,
        error: "ยังไม่มีการสมัครใช้งานด้วย LINE",
        needsRegistration: true,
        liffId: LINE_LIFF_ID
      };
    }
    return {
      success: true,
      user: publicUserProfile(record.user)
    };
  } catch (error) {
    Logger.log("loginWithLine error: " + error);
    return {
      success: false,
      error: error.message
    };
  }
}

function updateUserProfile(payload) {
  try {
    const data = payload || {};
    const userId = (data.userid || data.userId || "").toString().trim();
    if (!userId) {
      throw new Error("ไม่พบรหัสผู้ใช้");
    }
    const name = (data.name || "").toString().trim();
    const surname = (data.surname || "").toString().trim();
    const tel = (data.tel || data.phone || "").toString().trim();
    const email = (data.email || "").toString().trim();
    const newPassword = (data.newPassword || "").toString();
    if (!name) throw new Error("กรุณากรอกชื่อ");
    if (!surname) throw new Error("กรุณากรอกนามสกุล");
    if (!tel) throw new Error("กรุณากรอกเบอร์โทรศัพท์");
    if (!email) throw new Error("กรุณากรอกอีเมล");
    if (!isValidEmail(email)) throw new Error("รูปแบบอีเมลไม่ถูกต้อง");
    if (newPassword && newPassword.length < 6) {
      throw new Error("รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัวอักษร");
    }
    const avatarData =
      data.avatarFileData && data.avatarFileData.base64Data ? data.avatarFileData : null;

    const sheet = getUsersSheet();
    const rows = getAllUserRows_(sheet);
    const idIndex = USER_SHEET_HEADERS.indexOf("userid");
    if (idIndex === -1) {
      throw new Error("ไม่พบคอลัมน์ userid ในตารางผู้ใช้");
    }
    let targetRowIndex = -1;
    let record = null;
    rows.some((row, idx) => {
      if ((row[idIndex] || "").toString() === userId) {
        targetRowIndex = idx + 2;
        record = sanitizeUserRow(row);
        return true;
      }
      return false;
    });
    if (!record || targetRowIndex === -1) {
      throw new Error("ไม่พบบัญชีผู้ใช้");
    }

    const schoolIdInput = (data.schoolId || data.SchoolID || record.SchoolID || "").toString().trim();
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(5000);
      record.name = name;
      record.surname = surname;
      record.tel = tel;
      record.email = email;
      const canEditSchool = ["Admin", "Area"].includes((record.level || "").toString());
      record.SchoolID = canEditSchool ? (schoolIdInput || record.SchoolID || "") : (record.SchoolID || "");
      if (newPassword) {
        record.password = hashPassword(newPassword);
      }
      if (avatarData) {
        const newAvatarId = saveAvatarFile_(avatarData, userId);
        if (newAvatarId) {
          record.avatarFileId = newAvatarId;
        }
      }
      updateUserRow(sheet, targetRowIndex, record);
    } finally {
      lock.releaseLock();
    }

    return {
      success: true,
      user: publicUserProfile(record)
    };
  } catch (error) {
    Logger.log("updateUserProfile error: " + error);
    return {
      success: false,
      error: error.message
    };
  }
}// --- DATA READ FUNCTIONS (ฝั่ง Server) ---

/**
 * ดึงรายการกิจกรรมทั้งหมดจาก Sheet "Activities"
 */
function getActivities() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_ACTIVITIES);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Activities'");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // นับจำนวนทีมต่อ Activity
    const teamsSheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    const teamCounts = {};
    if (teamsSheet && teamsSheet.getLastRow() >= 2) {
      const teamRows = teamsSheet.getRange(2, 2, teamsSheet.getLastRow() - 1, 1).getValues();
      teamRows.forEach(row => {
        const activityId = row[0];
        if (!activityId) return;
        teamCounts[activityId] = (teamCounts[activityId] || 0) + 1;
      });
    }

    // A:I = 9 คอลัมน์
    const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    const now = new Date();

    const activities = data.map(row => {
      const id = row[0];

      // levels (เก็บเป็น JSON หรือ Array)
      let levels = [];
      if (row[3]) {
        try {
          levels = Array.isArray(row[3]) ? row[3] : JSON.parse(row[3]);
        } catch (parseError) {
          levels = [];
        }
      }

      const maxTeams = parseMaxTeamsValue(row[7]);
      const currentTeams = teamCounts[id] || 0;
      const remainingTeams = maxTeams ? Math.max(maxTeams - currentTeams, 0) : null;

      const deadlineDate = parseDeadlineValue(row[8]);
      const deadlineIso = deadlineDate ? deadlineDate.toISOString() : null;
      const deadlineDisplay = deadlineDate
        ? Utilities.formatDate(deadlineDate, "Asia/Bangkok", "dd MMM yyyy HH:mm")
        : null;
      const deadlinePassed = deadlineDate ? now.getTime() > deadlineDate.getTime() : false;
      const isFull = !!maxTeams && remainingTeams === 0;

      return {
        id: id,
        category: row[1],
        name: row[2],
        levels: levels,
        mode: row[4],
        reqTeachers: row[5],
        reqStudents: row[6],
        maxTeams: maxTeams,
        currentTeams: currentTeams,
        remainingTeams: remainingTeams,
        isUnlimited: maxTeams === null,
        registrationDeadline: deadlineIso,
        registrationDeadlineDisplay: deadlineDisplay,
        deadlinePassed: deadlinePassed,
        isFull: isFull,
        isClosed: deadlinePassed || isFull
      };
    });

    return activities;
  } catch (error) {
    Logger.log(error);
    return { error: error.message };
  }
}

/**
 * ดึงรายการทีมที่ลงทะเบียนทั้งหมดจาก Sheet "Teams"
 * รองรับคอลัมน์ LogoUrl และ TeamPhotoId (ถ้ามี)
 */
function getRegisteredTeams() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TEAMS);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Teams'");
    if (sheet.getLastRow() < 2) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // อ่านสูงสุดถึง 12 คอลัมน์ (รองรับ TeamPhotoId) ถ้ามี
    const colCount = Math.min(lastCol, 14);
    const data = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();

    const teams = data.map(row => ({
      teamId: row[0],
      activity: row[1],
      teamName: row[2],
      school: row[3],
      level: row[4],
      contact: row[5], // JSON string
      members: row[6], // JSON string
      requiredTeachers: row[7],
      requiredStudents: row[8],
      status: row[9],
      logoUrl: row.length > 10 ? row[10] : "",           // File ID โลโก้ (ถ้ามี)
      teamPhotoId: row.length > 11 ? row[11] : "",        // File ID รูปทีม (ถ้ามี)
      createdByUserId: row.length > 12 ? row[12] : "",
      createdByUsername: row.length > 13 ? row[13] : ""
    }));

    return teams;
  } catch (error) {
    Logger.log(error);
    return { error: error.message };
  }
}

/**
 * ดึงข้อมูลสรุปสำหรับ Report
 */
function getReportData() {
  const activities = getActivities();
  if (activities.error) return activities;

  const activityMap = new Map(activities.map(a => [a.id, a.name]));
  const teams = getRegisteredTeams();
  if (teams.error) return teams;

  let totalTeams = teams.length;
  let totalTeachers = 0;
  let totalStudents = 0;

  const activitySummary = {};

  teams.forEach(team => {
    try {
      const activityName = activityMap.get(team.activity) || team.activity;
      const members = JSON.parse(team.members || "{}");
      const teachers = Array.isArray(members.teachers) ? members.teachers.length : 0;
      const students = Array.isArray(members.students) ? members.students.length : 0;

      totalTeachers += teachers;
      totalStudents += students;

      if (!activitySummary[activityName]) {
        activitySummary[activityName] = { teams: 0, teachers: 0, students: 0 };
      }
      activitySummary[activityName].teams++;
      activitySummary[activityName].teachers += teachers;
      activitySummary[activityName].students += students;
    } catch (e) {
      Logger.log("Error parsing members for team " + team.teamId + ": " + e.message);
    }
  });

  return {
    totals: {
      teams: totalTeams,
      teachers: totalTeachers,
      students: totalStudents,
      allMembers: totalTeachers + totalStudents
    },
    summary: activitySummary
  };
}

/**
 * ดึงรายการไฟล์ที่รอตรวจสอบ (สำหรับ Admin)
 */
function getFilesForReview() {
  if (!isAdmin()) {
    return [];
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_FILES);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Files'");
    if (sheet.getLastRow() < 2) return [];

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    const pendingFiles = data
      .filter(row => row[3] === "Pending")
      .map(row => ({
        fileId: row[0],
        teamId: row[1],
        fileType: row[2],
        status: row[3],
        fileUrl: row[4],
        remarks: row[5]
      }));
    return pendingFiles;
  } catch (error) {
    Logger.log(error);
    return { error: error.message };
  }
}

// --- DATA WRITE FUNCTIONS (ฝั่ง Server) ---

/**
 * Helper: อัปโหลดไฟล์ Base64 ไปยัง Drive
 */
function uploadFileToDrive(base64Data, mimeType, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  if (!folder) throw new Error("ไม่พบ Google Drive Folder");

  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);

  const newFile = folder.createFile(blob);
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFile;
}
function saveAvatarFile_(avatarData, userId) {
  try {
    if (!avatarData || !avatarData.base64Data) {
      return "";
    }
    const mimeType = avatarData.mimeType || "image/png";
    const fileName = "avatar_" + (userId || "user") + "_" + (avatarData.fileName || "avatar.png");
    const file = uploadFileToDrive(avatarData.base64Data, mimeType, fileName);
    return file.getId();
  } catch (error) {
    Logger.log("saveAvatarFile_ error: " + error);
    return "";
  }
}

function processMemberPhotosForTeam_(membersObj, teamId) {
  try {
    const result = { teachers: [], students: [] };
    const safeTeamId = (teamId || "team").toString();
    const typeMap = [
      { key: "teachers", label: "teacher" },
      { key: "students", label: "student" }
    ];
    typeMap.forEach(({ key, label }) => {
      const list = Array.isArray(membersObj && membersObj[key]) ? membersObj[key] : [];
      list.forEach((member, index) => {
        const mm = Object.assign({}, member);
        if (mm.photoFileData && mm.photoFileData.base64Data) {
          const pf = mm.photoFileData;
          const file = uploadFileToDrive(
            pf.base64Data,
            pf.mimeType || "image/jpeg",
            "member_" +
              safeTeamId +
              "_" +
              label +
              (index + 1) +
              "_" +
              (pf.fileName || label + (index + 1) + ".jpg")
          );
          mm.photoDriveId = file.getId();
        }
        delete mm.photoFileData;
        result[key].push(mm);
      });
    });
    return result;
  } catch (error) {
    Logger.log("processMemberPhotosForTeam_ error: " + error);
    return membersObj || { teachers: [], students: [] };
  }
}

/**
 * ลงทะเบียนทีมใหม่
 * รองรับ:
 * - logoFileData
 * - teamPhotoFileData
 * - members.teachers[].photoFileData
 * - members.students[].photoFileData
 */
function registerTeam(formData) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const teamsSheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    const activitiesSheet = spreadsheet.getSheetByName(SHEET_ACTIVITIES);
    if (!teamsSheet) throw new Error("ไม่พบ Sheet 'Teams'");
    if (!activitiesSheet) throw new Error("ไม่พบ Sheet 'Activities'");

    // ตรวจสอบกิจกรรมจาก ActivityID
    const activityFinder = activitiesSheet.getRange("A:A").createTextFinder(formData.activity).findNext();
    if (!activityFinder) {
      throw new Error("ไม่พบกิจกรรมที่เลือก");
    }
    const activityRowIndex = activityFinder.getRow();
    const activityRow = activitiesSheet.getRange(activityRowIndex, 1, 1, 9).getValues()[0];

    const maxTeams = parseMaxTeamsValue(activityRow[7]);
    const deadlineDate = parseDeadlineValue(activityRow[8]);
    const now = new Date();

    // ตรวจสอบ deadline
    if (deadlineDate && now.getTime() > deadlineDate.getTime()) {
      const deadlineText = Utilities.formatDate(deadlineDate, "Asia/Bangkok", "dd MMM yyyy HH:mm");
      return {
        success: false,
        error: "กิจกรรมนี้ปิดรับสมัครแล้ว (หมดเขต " + deadlineText + ")"
      };
    }

    // ตรวจสอบจำนวนทีมเต็ม
    if (maxTeams) {
      let existingCount = 0;
      const totalRows = teamsSheet.getLastRow();
      if (totalRows >= 2) {
        const activityValues = teamsSheet.getRange(2, 2, totalRows - 1, 1).getValues();
        existingCount = activityValues.reduce((count, row) => {
          return row[0] === formData.activity ? count + 1 : count;
        }, 0);
      }
      if (existingCount >= maxTeams) {
        return {
          success: false,
          error: "กิจกรรมนี้ปิดรับสมัครแล้ว (จำนวนทีมเต็ม)"
        };
      }
    }

    // สร้าง TeamID
    const newTeamId =
      "T" + Utilities.formatDate(now, "GMT+7", "yyMMdd") + (teamsSheet.getLastRow() + 1);

    // --- Upload Logo ---
    let logoFileDriveId = "";
    if (formData.logoFileData && formData.logoFileData.base64Data) {
      const logo = formData.logoFileData;
      const logoFile = uploadFileToDrive(
        logo.base64Data,
        logo.mimeType || "image/png",
        "logo_" + newTeamId + "_" + (logo.fileName || "logo.png")
      );
      logoFileDriveId = logoFile.getId();
    }

    // --- Upload Team Photo ---
    let teamPhotoDriveId = "";
    if (formData.teamPhotoFileData && formData.teamPhotoFileData.base64Data) {
      const tp = formData.teamPhotoFileData;
      const teamPhotoFile = uploadFileToDrive(
        tp.base64Data,
        tp.mimeType || "image/jpeg",
        "teamphoto_" + newTeamId + "_" + (tp.fileName || "team.jpg")
      );
      teamPhotoDriveId = teamPhotoFile.getId();
    }

    const contactJson = JSON.stringify(formData.contact || {});
    const processedMembers = processMemberPhotosForTeam_(formData.members || {}, newTeamId);
    const membersJson = JSON.stringify(processedMembers);

    const requiredTeachers = formData.requiredTeachers || "";
    const requiredStudents = formData.requiredStudents || "";
    const createdByUserId = (formData.createdByUserId || "").toString();
    const createdByUsername = (formData.createdByUsername || "").toString();

    // --- Append Row ---
    // โครงสร้างแนะนำ:
    // A TeamID
    // B ActivityID
    // C TeamName
    // D School
    // E Level
    // F Contact(JSON)
    // G Members(JSON)
    // H RequiredTeachers
    // I RequiredStudents
    // J Status
    // K LogoUrl (File ID)
    // L TeamPhotoId (File ID)
    // M CreatedByUserId
    // N CreatedByUsername
    const newRow = [
      newTeamId,
      formData.activity,
      formData.teamName,
      formData.school,
      formData.level,
      contactJson,
      membersJson,
      requiredTeachers,
      requiredStudents,
      "Pending",
      logoFileDriveId,
      teamPhotoDriveId,
      createdByUserId,
      createdByUsername
    ];

    teamsSheet.appendRow(newRow);

    return {
      success: true,
      teamId: newTeamId,
      message: "ลงทะเบียนทีม " + newTeamId + " สำเร็จ"
    };
  } catch (error) {
    Logger.log(error);
    return { success: false, error: error.message };
  }
}

function sanitizeMemberList_(list, type) {
  if (!Array.isArray(list)) return [];
  return list.map(member => {
    const record = {};
    record.name = (member && member.name ? member.name : "").toString().trim();
    if (member && member.prefix) {
      record.prefix = member.prefix;
    }
    if (type === "teacher") {
      record.phone = (member && (member.phone || member.tel) ? member.phone || member.tel : "").toString().trim();
    }
    if (type === "student") {
      record.class = (member && (member.class || member.room) ? member.class || member.room : "").toString().trim();
    }
    if (member && member.role) {
      record.role = member.role;
    }
    if (member && member.photoDriveId) {
      record.photoDriveId = member.photoDriveId;
    }
    return record;
  });
}


// function buildSchoolsIndex_(spreadsheet) {
//   var index = {
//     byId: {},
//     byName: {},
//     clustersById: {}
//   };

//   try {
//     var book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
//     var schoolSheet = book.getSheetByName(SHEET_SCHOOLS);
//     if (!schoolSheet) {
//       return index;
//     }

//     var clusterMap = buildSchoolClusterMap_(book);

//     var lastRow = schoolSheet.getLastRow();
//     if (lastRow < 2) {
//       return index;
//     }

//     var rows = schoolSheet.getRange(2, 1, lastRow - 1, 3).getValues();
//     for (var i = 0; i < rows.length; i++) {
//       var id = String(rows[i][0] || "").trim();     // SchoolID
//       var name = String(rows[i][1] || "").trim();   // SchoolName
//       var clusterRaw = String(rows[i][2] || "").trim(); // SchoolClusterID (อาจเป็น ID หรือค่าที่กรอกมา)

//       if (!id && !name) {
//         continue;
//       }

//       var clusterId = "";
//       var clusterName = "";

//       if (clusterRaw) {
//         var norm = normalizeKey(clusterRaw);
//         var hit = clusterMap.byId[norm] || clusterMap.byName[norm];
//         if (hit) {
//           clusterId = hit.id;
//           clusterName = hit.name;
//         } else {
//           // ถ้าไม่เจอใน master ให้เก็บเป็นชื่อดิบ ๆ ไว้แสดง
//           clusterName = clusterRaw;
//         }
//       }

//       var record = {
//         id: id,
//         name: name,
//         clusterId: clusterId,
//         clusterName: clusterName,
//         // field เดิมสำหรับโค้ดหน้าเว็บที่ใช้ school.cluster
//         cluster: clusterName
//       };

//       if (id) {
//         index.byId[normalizeKey(id)] = record;
//       }
//       if (name) {
//         index.byName[normalizeKey(name)] = record;
//       }
//       if (clusterId && clusterName && !index.clustersById[clusterId]) {
//         index.clustersById[clusterId] = clusterName;
//       }
//     }
//   } catch (e) {
//     Logger.log("buildSchoolsIndex_ error: " + e);
//   }

//   return index;
// }

function buildSchoolsIndex_(spreadsheet) {
  var index = { byId: {}, byName: {}, clustersById: {} };
  try {
    var book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
    var schoolSheet = book.getSheetByName(SHEET_SCHOOLS);
    if (!schoolSheet) return index;

    var clusterMap = buildSchoolClusterMap_(book);
    var lastRow = schoolSheet.getLastRow();
    if (lastRow < 2) return index;

    var rows = schoolSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < rows.length; i++) {
      var id = String(rows[i][0] || "").trim();
      var name = String(rows[i][1] || "").trim();
      var clusterRaw = String(rows[i][2] || "").trim();
      if (!id && !name) continue;

      var clusterId = "";
      var clusterName = "";
      if (clusterRaw) {
        var norm = normalizeKey(clusterRaw);
        var hit = clusterMap.byId[norm] || clusterMap.byName[norm];
        if (hit) {
          clusterId = hit.id;
          clusterName = hit.name;
        } else {
          clusterName = clusterRaw;
        }
      }

      var rec = {
        id: id,
        name: name,
        clusterId: clusterId,
        clusterName: clusterName,
        cluster: clusterName
      };

      if (id)   index.byId[normalizeKey(id)] = rec;
      if (name) index.byName[normalizeKey(name)] = rec;
      if (clusterId && clusterName && !index.clustersById[clusterId]) {
        index.clustersById[clusterId] = clusterName;
      }
    }
  } catch (e) {
    Logger.log("buildSchoolsIndex_ error: " + e);
  }
  return index;
}


function lookupSchoolById_(index, schoolId) {
  if (!schoolId || !index || !index.byId) {
    return null;
  }
  const key = normalizeKey(schoolId);
  return index.byId[key] || null;
}

function lookupSchoolByName_(index, schoolName) {
  if (!schoolName || !index || !index.byName) {
    return null;
  }
  const key = normalizeKey(schoolName);
  return index.byName[key] || null;
}

function resolveActorContext_(actorInput, schoolIndex) {
  if (!actorInput) {
    return null;
  }
  const actor = {
    userId: (actorInput.userId || actorInput.userid || actorInput.id || "").toString().trim(),
    username: (actorInput.username || "").toString().trim(),
    level: (actorInput.level || actorInput.role || "").toString().trim(),
    schoolId: (actorInput.schoolId || actorInput.SchoolID || "").toString().trim(),
    schoolName: (actorInput.schoolName || "").toString().trim(),
    cluster: (actorInput.cluster || actorInput.schoolCluster || "").toString().trim()
  };
  if (!actor.userId) {
    return null;
  }
  const matchedUser = findUserById(actor.userId);
  if (!matchedUser || !matchedUser.user) {
    return null;
  }
  actor.level = matchedUser.user.level || actor.level || DEFAULT_USER_LEVEL;
  actor.schoolId = matchedUser.user.SchoolID || actor.schoolId;
  actor.username = matchedUser.user.username || actor.username;
  if (!actor.level) {
    actor.level = DEFAULT_USER_LEVEL;
  }
  const index = schoolIndex || buildSchoolsIndex_();
  if (!actor.schoolName && actor.schoolId) {
    const school = lookupSchoolById_(index, actor.schoolId);
    if (school) {
      actor.schoolName = school.name || actor.schoolName;
      actor.cluster = actor.cluster || school.cluster || "";
    }
  }
  if (!actor.cluster && actor.schoolName) {
    const schoolByName = lookupSchoolByName_(index, actor.schoolName);
    if (schoolByName) {
      actor.cluster = schoolByName.cluster || actor.cluster;
    }
  }
  actor.normalizedLevel = normalizeKey(actor.level);
  actor.schoolNameNormalized = normalizeKey(actor.schoolName);
  actor.clusterNormalized = normalizeKey(actor.cluster);
  return actor;
}

function buildTeamRecordFromRow_(rowValues, schoolIndex) {
  const getValue = (index) => {
    if (!rowValues || rowValues[index] === undefined) {
      return "";
    }
    return (rowValues[index] || "").toString().trim();
  };
  const team = {
    teamId: getValue(0),
    activityId: getValue(1),
    teamName: getValue(2),
    school: getValue(3),
    level: getValue(4),
    status: getValue(9),
    createdByUserId: getValue(12)
  };
  team.schoolNormalized = normalizeKey(team.school);
  const schoolInfo = lookupSchoolByName_(schoolIndex, team.school);
  team.schoolCluster = schoolInfo ? (schoolInfo.cluster || "") : "";
  team.schoolClusterNormalized = normalizeKey(team.schoolCluster);
  return team;
}

function evaluateTeamPermission_(actor, teamRecord) {
  if (!actor || !actor.userId) {
    return { canManage: false, canEditStatus: false };
  }
  const level = actor.normalizedLevel;
  const canEditStatus = level === "admin" || level === "area";
  if (level === "admin" || level === "area" || level === "score") {
    return { canManage: true, canEditStatus: canEditStatus };
  }
  if (teamRecord.createdByUserId && teamRecord.createdByUserId === actor.userId) {
    return { canManage: true, canEditStatus: false };
  }
  if (level === "school_admin") {
    const sameSchool =
      actor.schoolNameNormalized && actor.schoolNameNormalized === teamRecord.schoolNormalized;
    return { canManage: !!sameSchool, canEditStatus: false };
  }
  if (level === "group_admin") {
    const sameCluster =
      actor.clusterNormalized && actor.clusterNormalized === teamRecord.schoolClusterNormalized;
    const sameSchool =
      actor.schoolNameNormalized && actor.schoolNameNormalized === teamRecord.schoolNormalized;
    return { canManage: !!(sameCluster || sameSchool), canEditStatus: false };
  }
  return { canManage: false, canEditStatus: false };
}

function isUserWithinActorScope_(actor, userRecord, schoolsIndex) {
  if (!actor || !userRecord) {
    return false;
  }
  const level = actor.normalizedLevel;
  if (level === "admin" || level === "area") {
    return true;
  }
  const userSchoolId = normalizeKey(userRecord.SchoolID || userRecord.schoolId || "");
  const schoolInfo =
    lookupSchoolById_(schoolsIndex, userRecord.SchoolID) ||
    lookupSchoolByName_(schoolsIndex, userRecord.SchoolID || userRecord.schoolName || "");
  const userCluster = normalizeKey((schoolInfo && schoolInfo.cluster) || userRecord.cluster || "");
  if (level === "group_admin") {
    return actor.clusterNormalized && userCluster === actor.clusterNormalized;
  }
  if (level === "school_admin") {
    return actor.schoolId && normalizeKey(actor.schoolId) === userSchoolId;
  }
  return userRecord.userid === actor.userId;
}

function normalizeUserLevelName_(level) {
  const normalized = normalizeKey(level);
  for (let i = 0; i < USER_LEVELS.length; i++) {
    if (normalizeKey(USER_LEVELS[i]) === normalized) {
      return USER_LEVELS[i];
    }
  }
  return USER_LEVELS[0];
}

function updateTeamMembers(payload) {
  try {
    const data = payload || {};
    const teamId = (data.teamId || "").toString().trim();
    if (!teamId) {
      throw new Error("ไม่พบรหัสทีม");
    }
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Teams'");
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) {
      throw new Error("กรุณาเข้าสู่ระบบก่อนแก้ไขทีม");
    }
    const finder = sheet.getRange("A:A").createTextFinder(teamId).findNext();
    if (!finder) throw new Error("ไม่พบทีมนี้");

    const rowIndex = finder.getRow();
    const lastCol = Math.min(sheet.getLastColumn(), 14);
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const teamRecord = buildTeamRecordFromRow_(rowValues, schoolsIndex);
    const permissions = evaluateTeamPermission_(actor, teamRecord);
    if (!permissions.canManage) {
      throw new Error("ไม่มีสิทธิ์แก้ไขทีมนี้");
    }

    const existingContact = safeParseJson(rowValues[5], {});
    const memberData = safeParseJson(rowValues[6], { teachers: [], students: [] });

    const contactPayload = data.contact || {};
    const nextContact = {
      name: (contactPayload.name || existingContact.name || "").toString().trim(),
      phone: (contactPayload.phone || existingContact.phone || "").toString().trim(),
      email: (contactPayload.email || existingContact.email || "").toString().trim()
    };

    const memberPayload = data.members || {};
    const processedMembers = processMemberPhotosForTeam_(memberPayload, teamId);
    const sanitizedMembers = {
      teachers: sanitizeMemberList_(Array.isArray(processedMembers.teachers) ? processedMembers.teachers : (memberData.teachers || []), "teacher"),
      students: sanitizeMemberList_(Array.isArray(processedMembers.students) ? processedMembers.students : (memberData.students || []), "student")
    };

    const updatedRow = rowValues.slice();
    if (data.teamName) {
      updatedRow[2] = data.teamName;
    }
    updatedRow[5] = JSON.stringify(nextContact);
    updatedRow[6] = JSON.stringify(sanitizedMembers);
    const nextStatus = (data.status || "").toString().trim();
    if (nextStatus) {
      if (!permissions.canEditStatus && nextStatus !== updatedRow[9]) {
        throw new Error("สิทธิ์ไม่เพียงพอสำหรับแก้ไขสถานะทีม");
      }
      if (permissions.canEditStatus) {
        updatedRow[9] = nextStatus;
      }
    }
    const canEditImages =
      permissions.canManage &&
      actor &&
      ["admin", "area", "group_admin", "school_admin"].includes(actor.normalizedLevel);
    if (canEditImages && data.logoFileData && data.logoFileData.base64Data) {
      const logoFile = uploadFileToDrive(
        data.logoFileData.base64Data,
        data.logoFileData.mimeType || "image/png",
        "logo_update_" + teamId + "_" + (data.logoFileData.fileName || "logo.png")
      );
      updatedRow[10] = logoFile.getId();
    }
    if (canEditImages && data.teamPhotoFileData && data.teamPhotoFileData.base64Data) {
      const teamPhotoFile = uploadFileToDrive(
        data.teamPhotoFileData.base64Data,
        data.teamPhotoFileData.mimeType || "image/jpeg",
        "teamphoto_update_" + teamId + "_" + (data.teamPhotoFileData.fileName || "team.jpg")
      );
      updatedRow[11] = teamPhotoFile.getId();
    }

    sheet.getRange(rowIndex, 1, 1, lastCol).setValues([updatedRow]);
    return { success: true };
  } catch (error) {
    Logger.log("updateTeamMembers error: " + error);
    return { success: false, error: error.message };
  }
}

function updateActivityDetails(request) {
  try {
    const data = request || {};
    const activityId = (data.id || data.activityId || "").toString().trim();
    if (!activityId) throw new Error("ไม่พบรหัสกิจกรรม");

    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor || !["admin", "area"].includes(actor.normalizedLevel)) {
      throw new Error("คุณไม่มีสิทธิ์แก้ไขกิจกรรม");
    }

    const sheet = spreadsheet.getSheetByName(SHEET_ACTIVITIES);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Activities'");
    const finder = sheet.getRange("A:A").createTextFinder(activityId).findNext();
    if (!finder) throw new Error("ไม่พบกิจกรรมนี้");
    const rowIndex = finder.getRow();

    const name = (data.name || "").toString().trim();
    const category = (data.category || "").toString().trim();
    const mode = (data.mode || "").toString().trim();
    const reqTeachers = data.reqTeachers === "" ? "" : Number(data.reqTeachers || 0);
    const reqStudents = data.reqStudents === "" ? "" : Number(data.reqStudents || 0);
    const maxTeamsInput = data.maxTeams === "" ? "" : Number(data.maxTeams || 0);
    const levelsArray = Array.isArray(data.levels)
      ? data.levels
      : (data.levels || "").toString().split(",").map(s => s.trim()).filter(Boolean);
    const deadlineInput = (data.registrationDeadline || "").toString().trim();
    let deadlineValue = "";
    if (deadlineInput) {
      const parsed = new Date(deadlineInput);
      if (isNaN(parsed.getTime())) {
        throw new Error("รูปแบบวันปิดรับสมัครไม่ถูกต้อง");
      }
      deadlineValue = parsed;
    }

    sheet.getRange(rowIndex, 2).setValue(category);
    sheet.getRange(rowIndex, 3).setValue(name);
    sheet.getRange(rowIndex, 4).setValue(levelsArray.length ? JSON.stringify(levelsArray) : "");
    sheet.getRange(rowIndex, 5).setValue(mode);
    sheet.getRange(rowIndex, 6).setValue(reqTeachers === "" ? "" : reqTeachers);
    sheet.getRange(rowIndex, 7).setValue(reqStudents === "" ? "" : reqStudents);
    sheet.getRange(rowIndex, 8).setValue(maxTeamsInput > 0 ? maxTeamsInput : "");
    if (deadlineValue) {
      sheet.getRange(rowIndex, 9).setValue(deadlineValue);
    } else {
      sheet.getRange(rowIndex, 9).clearContent();
    }

    return {
      success: true,
      activityId: activityId
    };
  } catch (error) {
    Logger.log("updateActivityDetails error: " + error);
    return { success: false, error: error.message };
  }
}

function deleteTeam(request) {
  try {
    const payload = typeof request === "object" && request !== null ? request : { teamId: request };
    const targetId = (payload.teamId || "").toString().trim();
    if (!targetId) throw new Error("ไม่พบรหัสทีม");
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    if (!sheet) throw new Error("ไม่พบ Sheet 'Teams'");
    const finder = sheet.getRange("A:A").createTextFinder(targetId).findNext();
    if (!finder) throw new Error("ไม่พบทีมนี้");
    const rowIndex = finder.getRow();
    const lastCol = Math.min(sheet.getLastColumn(), 14);
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(payload.actor, schoolsIndex);
    if (!actor) {
      throw new Error("กรุณาเข้าสู่ระบบก่อนจัดการทีม");
    }
    const teamRecord = buildTeamRecordFromRow_(rowValues, schoolsIndex);
    const permissions = evaluateTeamPermission_(actor, teamRecord);
    if (!permissions.canManage) {
      throw new Error("ไม่มีสิทธิ์ลบทีมนี้");
    }
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (error) {
    Logger.log("deleteTeam error: " + error);
    return { success: false, error: error.message };
  }
}

/**
 * อัปโหลดไฟล์ (Section 5)
 */
function uploadFile(fileData) {
  try {
    const { base64Data, mimeType, fileName, teamId, fileType } = fileData;

    const newFile = uploadFileToDrive(
      base64Data,
      mimeType,
      "file_" + teamId + "_" + fileName
    );
    const fileUrl = newFile.getUrl();
    const fileDriveId = newFile.getId();

    const filesSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_FILES);
    const logId = "F" + (filesSheet.getLastRow() + 1);

    filesSheet.appendRow([
      logId, // FileLogID
      teamId,
      fileType,
      "Pending",
      fileUrl,
      "", // Remarks
      fileDriveId
    ]);

    return {
      success: true,
      message: "อัปโหลดไฟล์ " + fileName + " สำเร็จ",
      fileUrl: fileUrl
    };
  } catch (error) {
    Logger.log(error);
    return { success: false, error: error.message };
  }
}

/**
 * ส่งออกข้อมูลทีมเป็น CSV หรือ PDF
 */
function exportTeams(format) {
  try {
    const activities = getActivities();
    if (activities.error) throw new Error(activities.error);

    const teams = getRegisteredTeams();
    if (teams.error) throw new Error(teams.error);

    const activityMap = {};
    activities.forEach(activity => {
      if (activity && activity.id) {
        activityMap[activity.id] = activity;
      }
    });

    const rows = teams.map(team => {
      const activity = activityMap[team.activity] || {};
      const contact = safeParseJson(team.contact, {});
      const members = safeParseJson(team.members, { teachers: [], students: [] });
      const teachers = Array.isArray(members.teachers) ? members.teachers.length : 0;
      const students = Array.isArray(members.students) ? members.students.length : 0;

      const parsedRequiredTeachers = parseInt(team.requiredTeachers, 10);
      const parsedRequiredStudents = parseInt(team.requiredStudents, 10);

      return {
        teamId: team.teamId || "",
        activityId: team.activity || "",
        activityName: activity.name || "",
        category: activity.category || "",
        teamName: team.teamName || "",
        school: team.school || "",
        level: team.level || "",
        contactName: contact.name || "",
        contactPhone: contact.phone || "",
        contactEmail: contact.email || "",
        requiredTeachers: isNaN(parsedRequiredTeachers)
          ? activity.reqTeachers || teachers
          : parsedRequiredTeachers,
        actualTeachers: teachers,
        requiredStudents: isNaN(parsedRequiredStudents)
          ? activity.reqStudents || students
          : parsedRequiredStudents,
        actualStudents: students,
        status: team.status || "Pending"
      };
    });

    const timestamp = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "yyyyMMdd_HHmmss"
    );

    if (format === "csv") {
      const headers = [
        "TeamID",
        "ActivityID",
        "ActivityName",
        "Category",
        "TeamName",
        "School",
        "Level",
        "ContactName",
        "ContactPhone",
        "ContactEmail",
        "RequiredTeachers",
        "ActualTeachers",
        "RequiredStudents",
        "ActualStudents",
        "Status"
      ];

      const csvRows = rows.map(row => [
        row.teamId,
        row.activityId,
        row.activityName,
        row.category,
        row.teamName,
        row.school,
        row.level,
        row.contactName,
        row.contactPhone,
        row.contactEmail,
        row.requiredTeachers,
        row.actualTeachers,
        row.requiredStudents,
        row.actualStudents,
        row.status
      ]);

      const csvContent = [headers, ...csvRows]
        .map(row => row.map(escapeCsvCell).join(","))
        .join("\r\n");

      const fileName = "teams_" + timestamp + ".csv";
      const blob = Utilities.newBlob(csvContent, "text/csv", fileName);

      return {
        success: true,
        fileName: fileName,
        mimeType: "text/csv",
        data: Utilities.base64Encode(blob.getBytes())
      };
    }

    if (format === "pdf") {
      const tableRows = rows
        .map(
          (row, index) => `
        <tr>
          <td>${index + 1}</td>
          <td>${escapeHtml(row.teamId)}</td>
          <td>${escapeHtml(row.teamName)}</td>
          <td>${escapeHtml(row.activityName)}</td>
          <td>${escapeHtml(row.school)}</td>
          <td>${escapeHtml(row.level)}</td>
          <td>${escapeHtml(row.contactName)}<br>${escapeHtml(
            row.contactPhone
          )}<br>${escapeHtml(row.contactEmail)}</td>
          <td>${row.actualTeachers}/${row.requiredTeachers}</td>
          <td>${row.actualStudents}/${row.requiredStudents}</td>
          <td>${escapeHtml(row.status)}</td>
        </tr>`
        )
        .join("");

      const html = `
        <html>
          <head>
            <meta charset="UTF-8">
            <style>
              body { font-family: Arial, "TH SarabunPSK", sans-serif; font-size: 11pt; }
              h2 { margin-bottom: 6px; }
              table { width: 100%; border-collapse: collapse; }
              th, td { border: 1px solid #cccccc; padding: 6px; vertical-align: top; }
              th { background-color: #f1f5f9; }
            </style>
          </head>
          <body>
            <h2>รายงานทีมที่ลงทะเบียน</h2>
            <p>อัปเดตล่าสุด: ${Utilities.formatDate(
              new Date(),
              "Asia/Bangkok",
              "dd MMM yyyy HH:mm"
            )}</p>
            <p>จำนวนทีมทั้งหมด: ${rows.length} ทีม</p>
            <table>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Team ID</th>
                  <th>ชื่อทีม</th>
                  <th>กิจกรรม</th>
                  <th>สถานศึกษา</th>
                  <th>ระดับ</th>
                  <th>ผู้ประสานงาน</th>
                  <th>ครู (จริง/ต้องการ)</th>
                  <th>นักเรียน (จริง/ต้องการ)</th>
                  <th>สถานะ</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </body>
        </html>
      `;

      const pdfBlob = Utilities.newBlob(
        html,
        "text/html",
        "teams_export.html"
      ).getAs("application/pdf");
      const pdfFileName = "teams_" + timestamp + ".pdf";
      pdfBlob.setName(pdfFileName);

      return {
        success: true,
        fileName: pdfFileName,
        mimeType: "application/pdf",
        data: Utilities.base64Encode(pdfBlob.getBytes())
      };
    }



    throw new Error("รูปแบบไฟล์ไม่รองรับ");
  } catch (error) {
    Logger.log(error);
    return {
      success: false,
      error: error.message || "ไม่สามารถส่งออกข้อมูลได้"
    };
  }
}

/**
 * ดึงรายชื่อโรงเรียนทั้งหมด (Teams + Schools)
 */

// function getSchoolList() {
//   try {
//     var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
//     var teamsSheet = spreadsheet.getSheetByName(SHEET_TEAMS);
//     var schoolsSheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
//     var clusterMap = buildSchoolClusterMap_(spreadsheet);

//     // ใช้ object แทน Map เพื่อกันซ้ำ
//     var unique = {};

//     // จาก Teams: ดึงชื่อโรงเรียนที่มีทีม
//     if (teamsSheet && teamsSheet.getLastRow() >= 2) {
//       var lastRowTeams = teamsSheet.getLastRow();
//       var teamSchools = teamsSheet.getRange(2, 4, lastRowTeams - 1, 1).getValues(); // สมมติคอลัมน์ D = School
//       for (var i = 0; i < teamSchools.length; i++) {
//         var sName = String(teamSchools[i][0] || "").trim();
//         if (!sName) {
//           continue;
//         }
//         var key = normalizeKey(sName);
//         if (!unique[key]) {
//           unique[key] = {
//             id: "",
//             name: sName,
//             cluster: "",
//             clusterId: ""
//           };
//         }
//       }
//     }

//     // จาก Schools: เติมข้อมูลรหัสและเครือข่ายให้ครบ
//     if (schoolsSheet && schoolsSheet.getLastRow() >= 2) {
//       var lastRowSch = schoolsSheet.getLastRow();
//       var schValues = schoolsSheet.getRange(2, 1, lastRowSch - 1, 3).getValues();
//       for (var j = 0; j < schValues.length; j++) {
//         var id = String(schValues[j][0] || "").trim();
//         var name = String(schValues[j][1] || "").trim();
//         var clusterRaw = String(schValues[j][2] || "").trim();
//         if (!name) {
//           continue;
//         }

//         var clusterId = "";
//         var clusterName = "";
//         if (clusterRaw) {
//           var normRaw = normalizeKey(clusterRaw);
//           var hit = clusterMap.byId[normRaw] || clusterMap.byName[normRaw];
//           if (hit) {
//             clusterId = hit.id;
//             clusterName = hit.name;
//           } else {
//             clusterName = clusterRaw;
//           }
//         }

//         var keyName = normalizeKey(name);
//         if (!unique[keyName]) {
//           unique[keyName] = {
//             id: id,
//             name: name,
//             cluster: clusterName,
//             clusterId: clusterId
//           };
//         } else {
//           // merge ข้อมูลเพิ่มให้ record ที่มีอยู่
//           if (id && !unique[keyName].id) {
//             unique[keyName].id = id;
//           }
//           if (clusterName && !unique[keyName].cluster) {
//             unique[keyName].cluster = clusterName;
//           }
//           if (clusterId && !unique[keyName].clusterId) {
//             unique[keyName].clusterId = clusterId;
//           }
//         }
//       }
//     }

//     // แปลง object → array แล้ว sort ตามชื่อ
//     var list = [];
//     for (var key in unique) {
//       if (unique.hasOwnProperty(key)) {
//         list.push(unique[key]);
//       }
//     }

//     list.sort(function(a, b) {
//       return a.name.localeCompare(b.name, "th");
//     });

//     return list;
//   } catch (e) {
//     Logger.log("getSchoolList error: " + e);
//     return {
//       error: e.message || "ไม่สามารถดึงรายชื่อโรงเรียนได้"
//     };
//   }
// }

function getSchoolList() {
  try {
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var teamsSheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    var schoolsSheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
    var clusterMap = buildSchoolClusterMap_(spreadsheet);
    var unique = {};

    if (teamsSheet && teamsSheet.getLastRow() >= 2) {
      var lastRowTeams = teamsSheet.getLastRow();
      var teamSchools = teamsSheet.getRange(2, 4, lastRowTeams - 1, 1).getValues();
      for (var i = 0; i < teamSchools.length; i++) {
        var sName = String(teamSchools[i][0] || "").trim();
        if (!sName) continue;
        var key = normalizeKey(sName);
        if (!unique[key]) {
          unique[key] = { id: "", name: sName, cluster: "", clusterId: "" };
        }
      }
    }

    if (schoolsSheet && schoolsSheet.getLastRow() >= 2) {
      var lastRowSch = schoolsSheet.getLastRow();
      var schValues = schoolsSheet.getRange(2, 1, lastRowSch - 1, 3).getValues();
      for (var j = 0; j < schValues.length; j++) {
        var id = String(schValues[j][0] || "").trim();
        var name = String(schValues[j][1] || "").trim();
        var clusterRaw = String(schValues[j][2] || "").trim();
        if (!name) continue;

        var clusterId = "";
        var clusterName = "";
        if (clusterRaw) {
          var normRaw = normalizeKey(clusterRaw);
          var hit = clusterMap.byId[normRaw] || clusterMap.byName[normRaw];
          if (hit) {
            clusterId = hit.id;
            clusterName = hit.name;
          } else {
            clusterName = clusterRaw;
          }
        }

        var keyName = normalizeKey(name);
        if (!unique[keyName]) {
          unique[keyName] = {
            id: id,
            name: name,
            cluster: clusterName,
            clusterId: clusterId
          };
        } else {
          if (id && !unique[keyName].id) unique[keyName].id = id;
          if (clusterName && !unique[keyName].cluster) unique[keyName].cluster = clusterName;
          if (clusterId && !unique[keyName].clusterId) unique[keyName].clusterId = clusterId;
        }
      }
    }

    var list = [];
    for (var k in unique) {
      if (unique.hasOwnProperty(k)) list.push(unique[k]);
    }
    list.sort(function(a, b) { return a.name.localeCompare(b.name, "th"); });
    return list;
  } catch (e) {
    Logger.log("getSchoolList error: " + e);
    return { error: e.message || "ไม่สามารถดึงรายชื่อโรงเรียนได้" };
  }
}

function getManagedUsers(request) {
  try {
    const payload = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(payload.actor, schoolsIndex);
    if (!actor) throw new Error("กรุณาเข้าสู่ระบบ");
    const allowed = ["admin", "area", "group_admin", "school_admin"];
    if (!allowed.includes(actor.normalizedLevel)) {
      throw new Error("คุณไม่มีสิทธิ์เข้าถึงข้อมูลนี้");
    }
    const sheet = getUsersSheet();
    const rows = getAllUserRows_(sheet);
    if (!rows.length) {
      return { success: true, users: [] };
    }
    const users = rows.map(row => {
      const record = sanitizeUserRow(row);
      const schoolInfo =
        lookupSchoolById_(schoolsIndex, record.SchoolID) ||
        lookupSchoolByName_(schoolsIndex, record.SchoolID) ||
        lookupSchoolByName_(schoolsIndex, record.schoolName || "");
      return {
        userId: record.userid || "",
        username: record.username || "",
        name: record.name || "",
        surname: record.surname || "",
        level: (record.level || "").toString(),
        schoolId: record.SchoolID || "",
        schoolName: (schoolInfo && schoolInfo.name) || record.SchoolID || "",
        cluster: (schoolInfo && schoolInfo.cluster) || "",
        tel: record.tel || "",
        email: record.email || ""
      };
    });
    const normalize = value => (value || "").toString().trim().toLowerCase();
    const filtered = users.filter(user => {
      if (actor.normalizedLevel === "admin" || actor.normalizedLevel === "area") return true;
      if (actor.normalizedLevel === "group_admin") {
        return actor.clusterNormalized && normalize(user.cluster) === actor.clusterNormalized;
      }
      if (actor.normalizedLevel === "school_admin") {
        if (!actor.schoolId) return false;
        return normalize(user.schoolId) === normalize(actor.schoolId);
      }
      return user.userId === actor.userId;
    });
    return { success: true, users: filtered };
  } catch (error) {
    Logger.log("getManagedUsers error: " + error);
    return { success: false, error: error.message };
  }
}

function getManagedSchools(request) {
  try {
    const payload = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schools = getSchoolList();
    if (schools.error) throw new Error(schools.error);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(payload.actor, schoolsIndex);
    if (!actor) throw new Error("กรุณาเข้าสู่ระบบ");
    const allowed = ["admin", "area", "group_admin", "school_admin"];
    if (!allowed.includes(actor.normalizedLevel)) {
      throw new Error("คุณไม่มีสิทธิ์เข้าถึงข้อมูลนี้");
    }
    const normalize = value => (value || "").toString().trim().toLowerCase();
    let filtered = Array.isArray(schools) ? [...schools] : [];
    if (actor.normalizedLevel === "school_admin") {
      const target = normalize(actor.schoolId);
      filtered = filtered.filter(school => normalize(school.id) === target);
    } else if (actor.normalizedLevel === "group_admin") {
      const targetCluster = actor.clusterNormalized;
      filtered = targetCluster
        ? filtered.filter(school => normalize(school.cluster) === targetCluster)
        : [];
    }
    return { success: true, schools: filtered };
  } catch (error) {
    Logger.log("getManagedSchools error: " + error);
    return { success: false, error: error.message };
  }
}

function saveManagedUser(request) {
  try {
    const data = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) throw new Error("กรุณาเข้าสู่ระบบ");
    const level = actor.normalizedLevel;
    if (!["admin", "area", "group_admin", "school_admin"].includes(level)) {
      throw new Error("คุณไม่มีสิทธิ์จัดการผู้ใช้");
    }
    const input = data.user || {};
    const userId = (input.userId || input.userid || "").toString().trim();
    const username = (input.username || "").toString().trim();
    const name = (input.name || "").toString().trim();
    const surname = (input.surname || "").toString().trim();
    const tel = (input.tel || input.phone || "").toString().trim();
    const email = (input.email || "").toString().trim();
    const passwordInput = (input.password || "").toString();
    const schoolIdInput = (input.schoolId || input.SchoolID || "").toString().trim();
    if (!userId && !username) throw new Error("กรุณากรอกชื่อผู้ใช้");
    if (!name) throw new Error("กรุณากรอกชื่อ");
    if (!surname) throw new Error("กรุณากรอกนามสกุล");
    if (!tel) throw new Error("กรุณากรอกเบอร์ติดต่อ");
    if (!email) throw new Error("กรุณากรอกอีเมล");
    if (!isValidEmail(email)) throw new Error("รูปแบบอีเมลไม่ถูกต้อง");
    if (!schoolIdInput) throw new Error("กรุณาเลือกโรงเรียน");
    const schoolInfo =
      lookupSchoolById_(schoolsIndex, schoolIdInput) ||
      lookupSchoolByName_(schoolsIndex, schoolIdInput);
    if (!schoolInfo || !(schoolInfo.id || schoolInfo.name)) {
      throw new Error("ไม่พบโรงเรียนที่เลือก");
    }
    const normalizedSchoolId = schoolInfo.id || schoolIdInput;
    if (level === "school_admin") {
      if (!actor.schoolId || normalizeKey(actor.schoolId) !== normalizeKey(normalizedSchoolId)) {
        throw new Error("School Admin จัดการได้เฉพาะโรงเรียนของตัวเอง");
      }
    }
    if (level === "group_admin") {
      const schoolCluster = normalizeKey(schoolInfo.cluster);
      if (actor.clusterNormalized && schoolCluster !== actor.clusterNormalized) {
        throw new Error("Group Admin จัดการได้เฉพาะโรงเรียนในเครือข่ายตัวเอง");
      }
    }
    let requestedLevel = normalizeUserLevelName_(input.level || "User");
    if (level === "school_admin") {
      requestedLevel = "User";
    } else if (level === "group_admin" && !["User", "School_Admin"].includes(requestedLevel)) {
      requestedLevel = "User";
    } else if (
      !["admin", "area"].includes(level) &&
      ["Admin", "Area"].includes(requestedLevel)
    ) {
      requestedLevel = "User";
    }

    const sheet = getUsersSheet();
    const rows = getAllUserRows_(sheet);
    const usernameKey = normalizeKey(username);
    let targetRowIndex = -1;
    let record = null;
    if (userId) {
      rows.some((row, idx) => {
        if ((row[0] || "").toString() === userId) {
          targetRowIndex = idx + 2;
          record = sanitizeUserRow(row);
          return true;
        }
        return false;
      });
      if (!record) throw new Error("ไม่พบบัญชีผู้ใช้");
      if (!isUserWithinActorScope_(actor, record, schoolsIndex)) {
        throw new Error("คุณไม่มีสิทธิ์แก้ไขผู้ใช้นี้");
      }
      if (username && normalizeKey(record.username || "") !== usernameKey) {
        const duplicate = rows.some(row => normalizeKey(row[1]) === usernameKey);
        if (duplicate) throw new Error("ชื่อผู้ใช้ถูกใช้งานแล้ว");
      }
    } else {
      if (!username) throw new Error("กรุณากรอกชื่อผู้ใช้");
      const duplicate = rows.some(row => normalizeKey(row[1]) === usernameKey);
      if (duplicate) throw new Error("ชื่อผู้ใช้ถูกใช้งานแล้ว");
      if (!passwordInput) throw new Error("กรุณากรอกรหัสผ่าน");
      record = {
        userid: generateUserId(),
        username: username,
        password: "",
        name: "",
        surname: "",
        SchoolID: "",
        tel: "",
        userline_id: "",
        level: "",
        email: "",
        avatarFileId: ""
      };
    }

    record.username = username || record.username;
    record.name = name;
    record.surname = surname;
    record.tel = tel;
    record.email = email;
    record.SchoolID = normalizedSchoolId;
    record.level = requestedLevel;
    if (passwordInput) {
      record.password = hashPassword(passwordInput);
    } else if (!record.password) {
      record.password = "";
    }

    if (targetRowIndex > -1) {
      updateUserRow(sheet, targetRowIndex, record);
    } else {
      sheet.appendRow(buildRowFromUser(record));
    }

    return {
      success: true,
      user: publicUserProfile(record)
    };
  } catch (error) {
    Logger.log("saveManagedUser error: " + error);
    return { success: false, error: error.message };
  }
}

function deleteManagedUser(request) {
  try {
    const data = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) throw new Error("กรุณาเข้าสู่ระบบ");
    const level = actor.normalizedLevel;
    if (!["admin", "area", "group_admin", "school_admin"].includes(level)) {
      throw new Error("คุณไม่มีสิทธิ์จัดการผู้ใช้");
    }
    const targetId = (data.userId || data.userid || "").toString().trim();
    if (!targetId) throw new Error("ไม่พบรหัสผู้ใช้");
    if (normalizeKey(targetId) === normalizeKey(actor.userId)) {
      throw new Error("ไม่สามารถลบบัญชีของตนเองได้");
    }
    const sheet = getUsersSheet();
    const rows = getAllUserRows_(sheet);
    let targetRowIndex = -1;
    let record = null;
    rows.some((row, idx) => {
      if ((row[0] || "").toString() === targetId) {
        targetRowIndex = idx + 2;
        record = sanitizeUserRow(row);
        return true;
      }
      return false;
    });
    if (!record) throw new Error("ไม่พบบัญชีผู้ใช้");
    if (!isUserWithinActorScope_(actor, record, schoolsIndex)) {
      throw new Error("คุณไม่มีสิทธิ์ลบผู้ใช้นี้");
    }
    sheet.deleteRow(targetRowIndex);
    return { success: true };
  } catch (error) {
    Logger.log("deleteManagedUser error: " + error);
    return { success: false, error: error.message };
  }
}


// function updateSchoolCluster(request) {
//   try {
//     var data = request || {};
//     var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
//     var schoolsIndex = buildSchoolsIndex_(spreadsheet);
//     var actor = resolveActorContext_(data.actor, schoolsIndex);

//     // จำกัดสิทธิ์
//     if (!actor || (actor.normalizedLevel !== "admin" && actor.normalizedLevel !== "area")) {
//       throw new Error("เฉพาะ Admin หรือ Area เท่านั้นที่แก้ไขเครือข่ายได้");
//     }

//     var schoolId = String(data.schoolId || "").trim();
//     if (!schoolId) {
//       throw new Error("ไม่พบรหัสโรงเรียน");
//     }

//     var input = String(data.cluster || "").trim();
//     if (!input) {
//       throw new Error("กรุณาระบุรหัสหรือชื่อเครือข่าย");
//     }

//     var clusterMap = buildSchoolClusterMap_(spreadsheet);
//     var norm = normalizeKey(input);
//     var hit = clusterMap.byId[norm] || clusterMap.byName[norm];
//     if (!hit) {
//       throw new Error("ไม่พบเครือข่ายที่ระบุในชีต SchoolCluster");
//     }

//     var sheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
//     if (!sheet) {
//       throw new Error("ไม่พบชีต Schools");
//     }

//     // หาแถวที่ schoolId ตรงกับคอลัมน์ A
//     var lastRow = sheet.getLastRow();
//     if (lastRow < 2) {
//       throw new Error("ยังไม่มีข้อมูลโรงเรียน");
//     }

//     var idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
//     var rowIndex = -1;
//     for (var i = 0; i < idRange.length; i++) {
//       var currentId = String(idRange[i][0] || "").trim();
//       if (currentId && normalizeKey(currentId) === normalizeKey(schoolId)) {
//         rowIndex = i + 2; // offset header
//         break;
//       }
//     }

//     if (rowIndex === -1) {
//       throw new Error("ไม่พบโรงเรียนในฐานข้อมูล");
//     }

//     // บันทึกเป็น SchoolClusterID (ตาม requirement)
//     sheet.getRange(rowIndex, 3).setValue(hit.id);

//     return {
//       success: true,
//       schoolId: schoolId,
//       clusterId: hit.id,
//       clusterName: hit.name
//     };
//   } catch (e) {
//     Logger.log("updateSchoolCluster error: " + e);
//     return {
//       success: false,
//       error: e.message || e
//     };
//   }
// }

function updateSchoolCluster(request) {
  try {
    var data = request || {};
    var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    var schoolsIndex = buildSchoolsIndex_(spreadsheet);
    var actor = resolveActorContext_(data.actor, schoolsIndex);

    if (!actor || (actor.normalizedLevel !== "admin" && actor.normalizedLevel !== "area")) {
      throw new Error("เฉพาะ Admin หรือ Area เท่านั้นที่แก้ไขเครือข่ายได้");
    }

    var schoolId = String(data.schoolId || "").trim();
    if (!schoolId) throw new Error("ไม่พบรหัสโรงเรียน");

    var input = String(data.cluster || "").trim();
    if (!input) throw new Error("กรุณาระบุรหัสหรือชื่อเครือข่าย");

    var clusterMap = buildSchoolClusterMap_(spreadsheet);
    var norm = normalizeKey(input);
    var hit = clusterMap.byId[norm] || clusterMap.byName[norm];
    if (!hit) throw new Error("ไม่พบเครือข่ายที่ระบุในชีต SchoolCluster");

    var sheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
    if (!sheet) throw new Error("ไม่พบชีต Schools");

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("ยังไม่มีข้อมูลโรงเรียน");

    var idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var rowIndex = -1;
    for (var i = 0; i < idRange.length; i++) {
      var currentId = String(idRange[i][0] || "").trim();
      if (currentId && normalizeKey(currentId) === normalizeKey(schoolId)) {
        rowIndex = i + 2;
        break;
      }
    }
    if (rowIndex === -1) throw new Error("ไม่พบโรงเรียนในฐานข้อมูล");

    sheet.getRange(rowIndex, 3).setValue(hit.id);

    return { success: true, schoolId: schoolId, clusterId: hit.id, clusterName: hit.name };
  } catch (e) {
    Logger.log("updateSchoolCluster error: " + e);
    return { success: false, error: e.message || e };
  }
}

/**
 * อัปเดตสถานะไฟล์ (Admin)
 */
function updateFileStatus(fileLogId, newStatus, remarks) {
  if (!isAdmin()) {
    throw new Error("คุณไม่มีสิทธิ์ดำเนินการ");
  }

  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_FILES);
    const range = sheet.getRange("A:A").createTextFinder(fileLogId).findNext();
    if (!range) throw new Error("ไม่พบรหัสไฟล์นี้");

    const row = range.getRow();
    sheet.getRange(row, 4).setValue(newStatus); // Status
    sheet.getRange(row, 6).setValue(remarks);   // Remarks

    return {
      success: true,
      message: "อัปเดตสถานะ " + fileLogId + " เป็น " + newStatus
    };
  } catch (error) {
    Logger.log(error);
    return { success: false, error: error.message };
  }
}


// ===== getTeamsForExport (NEW) =====
function getTeamsForExport(request) {
  try {
    var data = request || {};
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var schoolsIndex = buildSchoolsIndex_(ss);
    var actor = resolveActorContext_(data.actor, schoolsIndex) || {};
    var level = String(actor.normalizedLevel || actor.level || "").toLowerCase();

    var teamsSheet = ss.getSheetByName(SHEET_TEAMS);
    if (!teamsSheet) return { success: true, teams: [], activities: [] };

    var lastRow = teamsSheet.getLastRow();
    var lastCol = teamsSheet.getLastColumn();
    if (lastRow < 2) return { success: true, teams: [], activities: [] };

    var headers = teamsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var values = teamsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    function findIdx(cands) {
      for (var c = 0; c < cands.length; c++) {
        var target = normalizeKey(cands[c]);
        for (var i = 0; i < headers.length; i++) {
          if (normalizeKey(headers[i]) === target) return i;
        }
      }
      return -1;
    }

    var idxTeamId    = findIdx(["teamid", "team_id", "รหัสทีม"]);
    var idxTeamName  = findIdx(["teamname", "team_name", "ชื่อทีม"]);
    var idxSchool    = findIdx(["school", "โรงเรียน"]);
    var idxActId     = findIdx(["activityid", "activity_id", "กิจกรรมid", "รหัสกิจกรรม"]);
    var idxActName   = findIdx(["activityname", "activity_name", "กิจกรรม", "ชื่อกิจกรรม"]);
    var idxLevel     = findIdx(["level", "ระดับ"]);
    var idxStatus    = findIdx(["status", "สถานะ"]);
    var idxReqT      = findIdx(["requiredteachers", "reqteachers", "ครูที่ต้องการ"]);
    var idxReqS      = findIdx(["requiredstudents", "reqstudents", "นักเรียนที่ต้องการ"]);
    var idxMembers   = findIdx(["members", "membersjson", "รายชื่อ", "memberjson"]);

    var actorClusterId = "";
    if (actor.clusterId) {
      actorClusterId = String(actor.clusterId).trim();
    } else if (actor.schoolId) {
      var sr = schoolsIndex.byId[normalizeKey(actor.schoolId)];
      if (sr && sr.clusterId) actorClusterId = sr.clusterId;
    }

    var teams = [];
    var activitiesMap = {};

    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      var schoolName = idxSchool >= 0 ? String(row[idxSchool] || "").trim() : "";
      var teamClusterId = "";
      if (schoolName) {
        var schRec = schoolsIndex.byName[normalizeKey(schoolName)];
        if (schRec && schRec.clusterId) teamClusterId = schRec.clusterId;
      }

      var include = false;
      if (level === "admin") {
        include = true;
      } else if (level === "group_admin" || level === "group-admin" || level === "groupadmin") {
        if (actorClusterId && teamClusterId &&
            normalizeKey(actorClusterId) === normalizeKey(teamClusterId)) {
          include = true;
        }
      } else {
        if (actor.schoolId && schoolName) {
          var srec = schoolsIndex.byId[normalizeKey(actor.schoolId)];
          if (srec && normalizeKey(srec.name) === normalizeKey(schoolName)) {
            include = true;
          }
        }
      }
      if (!include) continue;

      var actId = idxActId >= 0 ? String(row[idxActId] || "").trim() : "";
      var actName = idxActName >= 0 ? String(row[idxActName] || "").trim() : "";
      if (actId && !activitiesMap[actId]) {
        activitiesMap[actId] = actName || actId;
      }

      teams.push({
        teamId: idxTeamId >= 0 ? String(row[idxTeamId] || "").trim() : "",
        teamName: idxTeamName >= 0 ? String(row[idxTeamName] || "").trim() : "",
        school: schoolName,
        activityId: actId,
        activityName: actName,
        level: idxLevel >= 0 ? String(row[idxLevel] || "").trim() : "",
        status: idxStatus >= 0 ? String(row[idxStatus] || "").trim() : "",
        requiredTeachers: idxReqT >= 0 ? row[idxReqT] : "",
        requiredStudents: idxReqS >= 0 ? row[idxReqS] : "",
        members: idxMembers >= 0 ? String(row[idxMembers] || "").trim() : ""
      });
    }

    var activities = [];
    for (var id in activitiesMap) {
      if (activitiesMap.hasOwnProperty(id)) {
        activities.push({ id: id, name: activitiesMap[id] });
      }
    }

    return {
      success: true,
      teams: teams,
      activities: activities,
      actorLevel: level,
      actorClusterId: actorClusterId || ""
    };
  } catch (e) {
    Logger.log("getTeamsForExport error: " + e);
    return { success: false, error: e.message || e };
  }
}