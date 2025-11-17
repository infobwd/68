// --- CONFIGURATION ---
const DRIVE_FOLDER_ID = "1rqv8_Uh9SqmvLjsY--9CRwYRPYBCyjAD";
const SHEET_ID = "1Mu3yzfF7hCd-dtGk-RJV8f-zu_xtjoW9AWpVqtmZY2E";
// !! ?? สำคัญ: ใส่อีเมลของคุณที่เป็น Admin ??
const ADMIN_EMAIL = "noppharutlubbuangam@gmail.com";

// --- SHEET NAMES ---
const SHEET_ACTIVITIES = "Activities";
const SHEET_TEAMS = "Teams";
const SHEET_FILES = "Files";
const SHEET_SCHOOLS = "Schools";
const SHEET_USERS = "Users";
const SHEET_SCHOOL_CLUSTER = "SchoolCluster";
const SHEET_SCORE_ASSIGNMENTS = "ScoreAssignments";
const SHEET_SETTINGS = "Settings";
const TEAM_STAGE_COLUMN_INDEX = 20;
const TEAM_AREA_NAME_COLUMN_INDEX = 21;
const TEAM_AREA_CONTACT_COLUMN_INDEX = 22;
const TEAM_AREA_MEMBERS_COLUMN_INDEX = 23;
const TEAM_AREA_SCORE_COLUMN_INDEX = 24;
const TEAM_AREA_RANK_COLUMN_INDEX = 25;
const TEAM_SHEET_MAX_COLS = TEAM_AREA_RANK_COLUMN_INDEX;
const COMPETITION_STAGE_PROPERTY = "COMPETITION_STAGE";
const COMPETITION_STAGE_DEFAULT = "cluster";

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
const SCHOOL_SHEET_HEADERS = [
  "SchoolID",
  "SchoolName",
  "SchoolCluster",
  "RegistrationMode",
  "AssignedActivities"
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

function normalizeCompetitionStage_(value) {
  const normalized = (value || "").toString().trim().toLowerCase();
  return normalized === "area" ? "area" : COMPETITION_STAGE_DEFAULT;
}

function getCompetitionStage_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const stored = props.getProperty(COMPETITION_STAGE_PROPERTY);
    if (stored) {
      return normalizeCompetitionStage_(stored);
    }
    const fallback = getCompetitionStageFromSettings_();
    if (fallback) {
      return normalizeCompetitionStage_(fallback);
    }
    return COMPETITION_STAGE_DEFAULT;
  } catch (error) {
    Logger.log("getCompetitionStage_ error: " + error);
    return COMPETITION_STAGE_DEFAULT;
  }
}

function normalizeBooleanFlag_(value) {
  if (typeof value === "boolean") return value;
  const normalized = (value || "").toString().trim().toLowerCase();
  return ["true", "1", "yes", "y", "t"].includes(normalized);
}

function filterTeamsByCompetitionStage_(teams, stage) {
  if (!Array.isArray(teams)) return [];
  const normalizedStage = normalizeCompetitionStage_(stage);
  if (normalizedStage !== "area") {
    return teams;
  }
  return teams.filter(team => {
    const teamStage = normalizeCompetitionStage_(team && team.stage ? team.stage : "");
    if (teamStage === "area") return true;
    return normalizeBooleanFlag_(team && team.representativeOverride);
  });
}

function syncTeamStagesForCompetitionStage_(sheet, stage) {
  if (!sheet) return;
  try {
    ensureTeamSheetColumns_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const stageCol = TEAM_STAGE_COLUMN_INDEX;
    const repCol = TEAM_STAGE_COLUMN_INDEX - 1;
    const stageRange = sheet.getRange(2, stageCol, lastRow - 1, 1);
    const repValues = sheet.getRange(2, repCol, lastRow - 1, 1).getValues();
    const normalizedStage = normalizeCompetitionStage_(stage);
    const payload = repValues.map(row => {
      const isRepresentative = normalizeBooleanFlag_(row[0] || "");
      if (normalizedStage === "area" && isRepresentative) {
        return ["area"];
      }
      return ["cluster"];
    });
    stageRange.setValues(payload);
  } catch (error) {
    Logger.log("syncTeamStagesForCompetitionStage_ error: " + error);
  }
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

function ensureSchoolsSheetStructure_(sheet) {
  if (!sheet) {
    return;
  }
  const expectedLength = SCHOOL_SHEET_HEADERS.length;
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < expectedLength) {
    sheet.insertColumnsAfter(lastColumn, expectedLength - lastColumn);
  }
  const headerRange = sheet.getRange(1, 1, 1, expectedLength);
  const current = headerRange.getValues()[0];
  const next = current.slice();
  let changed = false;
  for (let i = 0; i < expectedLength; i++) {
    if (next[i] !== SCHOOL_SHEET_HEADERS[i]) {
      next[i] = SCHOOL_SHEET_HEADERS[i];
      changed = true;
    }
  }
  if (changed) {
    headerRange.setValues([next]).setFontWeight("bold");
  }
}

function normalizeRegistrationMode_(value) {
  const raw = normalizeKey(value);
  if (raw === "group" || raw === "group_assigned" || raw === "assigned" || raw === "network") {
    return "group_assigned";
  }
  return "self";
}

function formatRegistrationModeForSheet_(mode) {
  return mode === "group_assigned" ? "Group_Assigned" : "Self";
}

function parseAssignedActivitiesCell_(value) {
  if (value === null || value === undefined || value === "") {
    return [];
  }
  if (Array.isArray(value)) {
    return value
      .map(item => String(item || "").trim())
      .filter(Boolean);
  }
  const raw = String(value).trim();
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed
        .map(item => String(item || "").trim())
        .filter(Boolean);
    }
  } catch (error) {
    // ignore JSON parse error
  }
  return raw
    .split(/[,;\n]/)
    .map(item => item.trim())
    .filter(Boolean);
}

function stringifyAssignedActivities_(activities) {
  if (!Array.isArray(activities) || activities.length === 0) {
    return "";
  }
  return JSON.stringify(activities);
}

function parseSchoolIdentifiers_(value) {
  const cleaned = (value || "").toString().trim();
  if (!cleaned) {
    return { id: "", name: "" };
  }
  const match = cleaned.match(/\[([^\]]+)\]/);
  if (match) {
    const id = match[1].trim();
    const name = cleaned.replace(match[0], "").trim();
    return { id: id, name: name || id };
  }
  return { id: "", name: cleaned };
}

function findSchoolRecordByInput_(index, rawValue) {
  if (!index) {
    return null;
  }
  const parsed = parseSchoolIdentifiers_(rawValue);
  if (parsed.id) {
    const byId = lookupSchoolById_(index, parsed.id);
    if (byId) {
      return byId;
    }
  }
  if (parsed.name) {
    const byName = lookupSchoolByName_(index, parsed.name);
    if (byName) {
      return byName;
    }
  }
  return null;
}

function isGroupAssignedMode_(mode) {
  return normalizeRegistrationMode_(mode) === "group_assigned";
}

function isActivityAllowedForSchool_(schoolRecord, activityId) {
  if (!schoolRecord || !activityId) {
    return true;
  }
  if (!isGroupAssignedMode_(schoolRecord.registrationMode)) {
    return true;
  }
  const assigned = Array.isArray(schoolRecord.assignedActivities)
    ? schoolRecord.assignedActivities
    : [];
  if (!assigned.length) {
    return false;
  }
  const target = normalizeKey(activityId);
  return assigned.some(item => normalizeKey(item) === target);
}

function getSchoolClusterKeys_(schoolRecord) {
  if (!schoolRecord) return [];
  const values = [
    schoolRecord.clusterId,
    schoolRecord.cluster,
    schoolRecord.clusterName,
    schoolRecord.clusterLabel
  ];
  const keys = values
    .map(value => normalizeKey(value || ""))
    .filter(Boolean);
  return Array.from(new Set(keys));
}

function doesSchoolMatchCluster_(schoolRecord, clusterKey) {
  if (!clusterKey) return false;
  const keys = getSchoolClusterKeys_(schoolRecord);
  if (!keys.length) return false;
  return keys.some(key => key === clusterKey);
}

function isSchoolWithinActorCluster_(actor, schoolRecord) {
  if (!actor || actor.normalizedLevel !== "group_admin") return true;
  const actorCluster = actor.clusterNormalized;
  if (!actorCluster) return false;
  return doesSchoolMatchCluster_(schoolRecord, actorCluster);
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

function ensureScoreAssignmentSheet_(spreadsheet) {
  const book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
  let sheet = book.getSheetByName(SHEET_SCORE_ASSIGNMENTS);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_SCORE_ASSIGNMENTS);
    sheet.appendRow(["userId", "activityIds", "updatedAt", "updatedBy"]);
  }
  return sheet;
}

function ensureTeamSheetColumns_(sheet) {
  if (!sheet) return;
  try {
    const currentCols = sheet.getLastColumn();
    if (currentCols < TEAM_SHEET_MAX_COLS) {
      sheet.insertColumnsAfter(currentCols, TEAM_SHEET_MAX_COLS - currentCols);
    }
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const stageRange = sheet.getRange(2, TEAM_STAGE_COLUMN_INDEX, lastRow - 1, 1);
      const stageValues = stageRange.getValues();
      let stageDirty = false;
      stageValues.forEach((row, idx) => {
        if (!row[0]) {
          stageValues[idx][0] = "cluster";
          stageDirty = true;
        }
      });
      if (stageDirty) {
        stageRange.setValues(stageValues);
      }
      const areaColumns = [
        TEAM_AREA_NAME_COLUMN_INDEX,
        TEAM_AREA_CONTACT_COLUMN_INDEX,
        TEAM_AREA_MEMBERS_COLUMN_INDEX,
        TEAM_AREA_SCORE_COLUMN_INDEX,
        TEAM_AREA_RANK_COLUMN_INDEX
      ];
      areaColumns.forEach(col => {
        const areaRange = sheet.getRange(2, col, lastRow - 1, 1);
        const values = areaRange.getValues();
        const needsInit = values.some(row => row[0] === undefined || row[0] === null);
        if (needsInit) {
          const blank = Array.from({ length: lastRow - 1 }, () => [""]);
          areaRange.setValues(blank);
        }
      });
    }
  } catch (error) {
    Logger.log("ensureTeamSheetColumns_ error: " + error);
  }
}

function ensureSettingsSheet_(spreadsheet) {
  const book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
  let sheet = book.getSheetByName(SHEET_SETTINGS);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_SETTINGS);
  }
  const headers = ["Key", "Value", "UpdatedAt"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0] || [];
  let needsHeader = false;
  for (let i = 0; i < headers.length; i++) {
    if (existing[i] !== headers[i]) {
      needsHeader = true;
      break;
    }
  }
  if (needsHeader || !existing.length) {
    headerRange.setValues([headers]).setFontWeight("bold");
  }
  return sheet;
}

function persistCompetitionStageToSettings_(stage) {
  try {
    const book = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ensureSettingsSheet_(book);
    const key = "competition_stage";
    const normalizedStage = normalizeCompetitionStage_(stage);
    const lastRow = sheet.getLastRow();
    const now = new Date();
    if (lastRow < 2) {
      sheet.appendRow([key, normalizedStage, now]);
      return;
    }
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let targetRow = null;
    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || "").toString().trim().toLowerCase() === key) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow) {
      sheet.getRange(targetRow, 2).setValue(normalizedStage);
      sheet.getRange(targetRow, 3).setValue(now);
    } else {
      sheet.appendRow([key, normalizedStage, now]);
    }
  } catch (error) {
    Logger.log("persistCompetitionStageToSettings_ error: " + error);
  }
}

function getCompetitionStageFromSettings_() {
  try {
    const book = SpreadsheetApp.openById(SHEET_ID);
    const sheet = book.getSheetByName(SHEET_SETTINGS);
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if ((values[i][0] || "").toString().trim().toLowerCase() === "competition_stage") {
        return values[i][1] ? values[i][1].toString().trim() : null;
      }
    }
  } catch (error) {
    Logger.log("getCompetitionStageFromSettings_ error: " + error);
  }
  return null;
}

function getScoreAssignmentMap_(spreadsheet) {
  const map = new Map();
  const book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
  const sheet = book.getSheetByName(SHEET_SCORE_ASSIGNMENTS);
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  values.forEach(row => {
    const userId = (row[0] || "").toString().trim();
    if (!userId) return;
    const activities = (row[1] || "")
      .toString()
      .split(",")
      .map(v => v.trim())
      .filter(Boolean);
    map.set(userId, Array.from(new Set(activities)));
  });
  return map;
}

function upsertScoreAssignment_(spreadsheet, userId, activityIds, actorName) {
  const sheet = ensureScoreAssignmentSheet_(spreadsheet);
  const normalizedUserId = (userId || "").toString().trim();
  if (!normalizedUserId) return;
  const cleanedActivities = Array.from(
    new Set(
      (activityIds || []).map(id => (id || "").toString().trim()).filter(Boolean)
    )
  );
  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  if (lastRow >= 2) {
    const userColumn = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    userColumn.some((row, idx) => {
      if ((row[0] || "").toString().trim() === normalizedUserId) {
        targetRow = idx + 2;
        return true;
      }
      return false;
    });
  }
  if (!cleanedActivities.length) {
    if (targetRow > -1) {
      sheet.deleteRow(targetRow);
    }
    return;
  }
  const line = [
    normalizedUserId,
    cleanedActivities.join(","),
    new Date(),
    actorName || ""
  ];
  if (targetRow > -1) {
    sheet.getRange(targetRow, 1, 1, line.length).setValues([line]);
  } else {
    sheet.appendRow(line);
  }
}

function buildActivitySummaryForTeams_(teams, activityMap) {
  const summary = {};
  if (!Array.isArray(teams)) return summary;
  teams.forEach(team => {
    try {
      const activityName = activityMap.get(team.activity) || team.activity;
      const members = JSON.parse(team.members || "{}");
      const teachers = Array.isArray(members.teachers) ? members.teachers.length : 0;
      const students = Array.isArray(members.students) ? members.students.length : 0;
      if (!summary[activityName]) {
        summary[activityName] = { teams: 0, teachers: 0, students: 0 };
      }
      summary[activityName].teams++;
      summary[activityName].teachers += teachers;
      summary[activityName].students += students;
    } catch (error) {
      Logger.log("buildActivitySummaryForTeams_ error: " + error);
    }
  });
  return summary;
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

function findUserByEmail(email) {
  const key = normalizeKey(email);
  if (!key) {
    return null;
  }
  const sheet = getUsersSheet();
  const rows = getAllUserRows_(sheet);
  const emailIndex = USER_SHEET_HEADERS.indexOf("email");
  if (emailIndex === -1) {
    return null;
  }
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (normalizeKey(row[emailIndex]) === key) {
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
      const duplicateEmail = findUserByEmail(email);
      if (duplicateEmail) {
        throw new Error("อีเมลถูกใช้แล้ว");
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

function checkUserAvailability(payload) {
  try {
    const data = payload || {};
    const username = (data.username || "").toString().trim();
    const email = (data.email || "").toString().trim();
    if (!username && !email) {
      return {
        success: false,
        error: "กรุณาระบุชื่อผู้ใช้หรืออีเมล"
      };
    }
    const conflicts = [];
    if (username && findUserByUsername(username)) {
      conflicts.push("username");
    }
    if (email && findUserByEmail(email)) {
      conflicts.push("email");
    }
    if (conflicts.length) {
      let message = "";
      if (conflicts.includes("username") && conflicts.includes("email")) {
        message = "ชื่อผู้ใช้และอีเมลถูกใช้งานแล้ว";
      } else if (conflicts[0] === "username") {
        message = "ชื่อผู้ใช้ถูกใช้งานแล้ว";
      } else {
        message = "อีเมลถูกใช้งานแล้ว";
      }
      return {
        success: false,
        conflicts,
        error: message
      };
    }
    return { success: true };
  } catch (error) {
    Logger.log("checkUserAvailability error: " + error);
    return {
      success: false,
      error: error.message || "ไม่สามารถตรวจสอบข้อมูลได้"
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
    if (!sheet) throw new Error("??? Sheet 'Teams'");
    if (sheet.getLastRow() < 2) return [];
    ensureTeamSheetColumns_(sheet);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // อ่านสูงสุดถึง 12 คอลัมน์ (รองรับ TeamPhotoId) ถ้ามี
    const colCount = Math.min(lastCol, TEAM_SHEET_MAX_COLS);
    const data = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
    const teams = data.map(row => {
      const rawScore = row.length > 15 ? Number(row[15]) : null;
      const parsedScore = typeof rawScore === 'number' && !isNaN(rawScore) ? rawScore : null;
      const stageValue = normalizeCompetitionStage_(row.length > 19 ? row[19] : '');
    const areaTeamName = row.length > 20 ? row[20] : '';
    const areaContact = row.length > 21 ? row[21] : '';
    const areaMembers = row.length > 22 ? row[22] : '';
    const areaScore = row.length > 23 ? row[23] : '';
    const areaRank = row.length > 24 ? row[24] : '';
    return {
      teamId: row[0],
      activity: row[1],
      teamName: row[2],
      teamNameCluster: row[2],
      teamNameArea: areaTeamName,
      school: row[3],
      level: row[4],
      contact: row[5],
      members: row[6],
      contactArea: areaContact,
      membersArea: areaMembers,
      areaScore: areaScore,
      areaRank: areaRank,
      requiredTeachers: row[7],
        requiredStudents: row[8],
        status: row[9],
        logoUrl: row.length > 10 ? row[10] : '',
        teamPhotoId: row.length > 11 ? row[11] : '',
        createdByUserId: row.length > 12 ? row[12] : '',
        createdByUsername: row.length > 13 ? row[13] : '',
        statusReason: row.length > 14 ? row[14] : '',
        scoreTotal: parsedScore,
        scoreManualMedal: row.length > 16 ? row[16] : '',
        rankOverride: row.length > 17 ? row[17] : '',
        representativeOverride: row.length > 18 ? row[18] : '',
        stage: stageValue || 'cluster'
      };
    });

    return teams;
  } catch (error) {
    Logger.log(error);
    return { error: error.message };
  }
}

/**
 * ดึงข้อมูลสรุปสำหรับ Report
 */
function getReportData(options) {
  const opts = options || {};
  const actor = opts.actor || null;
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);

  const activities = getActivities();
  if (activities.error) return activities;

  const activityNameMap = new Map(activities.map(a => [a.id, a.name]));
  const activityInfoMap = new Map(activities.map(a => [a.id, a]));
  const schoolsIndex = buildSchoolsIndex_(spreadsheet);
  const teams = getRegisteredTeams();
  if (teams.error) return teams;

  let totalTeams = teams.length;
  let totalTeachers = 0;
  let totalStudents = 0;

  const activitySummary = {};

  teams.forEach(team => {
    try {
      const activityName = activityNameMap.get(team.activity) || team.activity;
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

  const scoreAssignmentMap = getScoreAssignmentMap_(spreadsheet);
  const competitionStage = getCompetitionStage_();
  const stageTeamsGlobal = filterTeamsByCompetitionStage_(teams, competitionStage);
  const competitionResultsGlobal = buildCompetitionResultsPayload_(stageTeamsGlobal, activityInfoMap, schoolsIndex);
  const clusterLeaderboardGlobal = buildClusterLeaderboard_(stageTeamsGlobal, schoolsIndex);
  const clusterActivitySummaryGlobal = buildClusterActivitySummary_(stageTeamsGlobal, activityInfoMap, schoolsIndex);

  let userTotals = null;
  let userSummary = null;
  let actorScoreActivities = [];
  let userCompetitionResults = null;
  let clusterLeaderboardScoped = null;
  let clusterActivitySummaryScoped = [];
  if (actor) {
    const actorLevel = normalizeKey(actor.level || "");
    const requiresClusterLookup = actorLevel === "group_admin" && normalizeKey(actor.clusterId || actor.cluster || "");
    const schoolLookup = requiresClusterLookup ? buildSchoolNameClusterLookup_() : null;
    const scopedTeams = filterTeamsByActor_(teams, actor, schoolLookup);
    userTotals = summarizeTeamList_(scopedTeams);
    userSummary = buildActivitySummaryForTeams_(scopedTeams, activityNameMap);
    actorScoreActivities = scoreAssignmentMap.get((actor.userId || "").toString().trim()) || [];
    const scopedCompetitionTeams = filterTeamsByCompetitionStage_(scopedTeams, competitionStage);
    userCompetitionResults = buildCompetitionResultsPayload_(scopedCompetitionTeams, activityInfoMap, schoolsIndex);
    clusterLeaderboardScoped = buildClusterLeaderboard_(scopedCompetitionTeams, schoolsIndex);
    clusterActivitySummaryScoped = buildClusterActivitySummary_(scopedCompetitionTeams, activityInfoMap, schoolsIndex);
  }

  return {
    totals: {
      teams: totalTeams,
      teachers: totalTeachers,
      students: totalStudents,
      allMembers: totalTeachers + totalStudents
    },
    summary: activitySummary,
    userTotals: userTotals,
    userSummary: userSummary,
    actorScoreActivities: actorScoreActivities,
    competitionResults: {
      global: competitionResultsGlobal,
      scoped: userCompetitionResults
    },
    clusterLeaderboard: {
      global: clusterLeaderboardGlobal,
      scoped: clusterLeaderboardScoped
    },
    clusterActivitySummary: {
      global: clusterActivitySummaryGlobal,
      scoped: clusterActivitySummaryScoped
    },
    competitionStage: competitionStage
  };
}

function getCompetitionSettings() {
  try {
    const stage = getCompetitionStage_();
    return { success: true, stage };
  } catch (error) {
    Logger.log("getCompetitionSettings error: " + error);
    return { success: false, error: error.message };
  }
}

function updateCompetitionStage(request) {
  try {
    const data = request || {};
    const nextStage = normalizeCompetitionStage_(data.stage || data.value || data.competitionStage);
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor || !["admin", "area"].includes(actor.normalizedLevel)) {
      throw new Error("เฉพาะ Admin หรือ Area เท่านั้นที่กำหนดรอบการแข่งขันได้");
    }
    PropertiesService.getScriptProperties().setProperty(COMPETITION_STAGE_PROPERTY, nextStage);
    persistCompetitionStageToSettings_(nextStage);
    const teamsSheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    if (teamsSheet) {
      syncTeamStagesForCompetitionStage_(teamsSheet, nextStage);
    }
    return { success: true, stage: nextStage };
  } catch (error) {
    Logger.log("updateCompetitionStage error: " + error);
    return { success: false, error: error.message };
  }
}

function summarizeTeamList_(teams) {
  const summary = {
    teams: Array.isArray(teams) ? teams.length : 0,
    teachers: 0,
    students: 0,
    allMembers: 0
  };
  if (!Array.isArray(teams) || !teams.length) return summary;
  teams.forEach(team => {
    try {
      const members = JSON.parse(team.members || "{}");
      if (Array.isArray(members.teachers)) summary.teachers += members.teachers.length;
      if (Array.isArray(members.students)) summary.students += members.students.length;
    } catch (error) {
      Logger.log("summarizeTeamList_ error: " + error);
    }
  });
  summary.allMembers = summary.teachers + summary.students;
  return summary;
}

const MEDAL_CONFIG = [
  { key: "gold", label: "เหรียญทอง", min: 80 },
  { key: "silver", label: "เหรียญเงิน", min: 70 },
  { key: "bronze", label: "เหรียญทองแดง", min: 60 },
  { key: "merit", label: "ชมเชย", min: 50 }
];

function parseScoreValue_(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const numeric = Number(value);
  return isNaN(numeric) ? null : numeric;
}

function resolveMedalFromScore_(score) {
  if (score === null || isNaN(score)) {
    return { key: "participant", label: "เข้าร่วม" };
  }
  for (let i = 0; i < MEDAL_CONFIG.length; i++) {
    if (score >= MEDAL_CONFIG[i].min) {
      return { key: MEDAL_CONFIG[i].key, label: MEDAL_CONFIG[i].label };
    }
  }
  return { key: "participant", label: "เข้าร่วม" };
}

function buildClusterActivitySummary_(teams, activityMap, schoolsIndex) {
  if (!Array.isArray(teams) || !teams.length) {
    return [];
  }
  const clusterMap = {};
  teams.forEach(team => {
    const activityId = (team.activity || "").toString().trim();
    if (!activityId) return;
    const activityInfo = activityMap.get(activityId) || {};
    const activityName = activityInfo.name || team.activity || "";
    if (!activityName) return;
    const stageSpecificMembers =
      normalizeCompetitionStage_(team.stage || "") === "area" && team.membersArea
        ? safeParseJson(team.membersArea, { teachers: [], students: [] })
        : safeParseJson(team.members, { teachers: [], students: [] });
    const members = stageSpecificMembers;
    const teacherCount = Array.isArray(members.teachers) ? members.teachers.length : 0;
    const studentCount = Array.isArray(members.students) ? members.students.length : 0;

    const schoolInfo = lookupSchoolByName_(schoolsIndex, team.school || "") || null;
    const clusterIdRaw =
      (schoolInfo && (schoolInfo.clusterId || schoolInfo.cluster)) ||
      team.schoolClusterId ||
      team.schoolCluster ||
      "";
    const clusterLabel =
      (schoolInfo && (schoolInfo.cluster || schoolInfo.clusterName)) ||
      team.schoolClusterName ||
      team.schoolCluster ||
      "ไม่ระบุเครือข่าย";
    const clusterKey = normalizeKey(clusterIdRaw || clusterLabel || "unassigned") || "unassigned";

    if (!clusterMap[clusterKey]) {
      clusterMap[clusterKey] = {
        clusterKey: clusterKey,
        clusterId: clusterIdRaw || "",
        clusterLabel: clusterLabel,
        totals: { teams: 0, teachers: 0, students: 0, allMembers: 0 },
        activities: {}
      };
    }
    const clusterEntry = clusterMap[clusterKey];
    const activityKey = activityName;
    if (!clusterEntry.activities[activityKey]) {
      clusterEntry.activities[activityKey] = {
        activityId: activityId,
        activityName: activityName,
        category: activityInfo.category || "ไม่ระบุหมวดหมู่",
        teams: 0,
        teachers: 0,
        students: 0
      };
    }
    const activityEntry = clusterEntry.activities[activityKey];
    activityEntry.teams += 1;
    activityEntry.teachers += teacherCount;
    activityEntry.students += studentCount;

    clusterEntry.totals.teams += 1;
    clusterEntry.totals.teachers += teacherCount;
    clusterEntry.totals.students += studentCount;
  });

  return Object.values(clusterMap)
    .map(entry => {
      entry.totals.allMembers = entry.totals.teachers + entry.totals.students;
      entry.activities = Object.values(entry.activities).sort((a, b) =>
        (a.activityName || "").localeCompare(b.activityName || "", "th")
      );
      return entry;
    })
    .sort((a, b) => (a.clusterLabel || "").localeCompare(b.clusterLabel || "", "th"));
}
function buildCompetitionResultsPayload_(teams, activityMap, schoolsIndex) {
  if (!Array.isArray(teams) || !teams.length) {
    return { activities: [], summary: null, representatives: [], areaFinals: [] };
  }
  const activitiesById = new Map();
  const medalCounts = {
    gold: 0,
    silver: 0,
    bronze: 0,
    merit: 0,
    participant: 0
  };
  let totalParticipants = 0;

  teams.forEach(team => {
    const score = parseScoreValue_(team.scoreTotal);
    if (score === null) return;
    const activityId = (team.activity || "").toString();
    if (!activityId) return;
    const activityInfo = activityMap.get(activityId) || {};
    if (!activitiesById.has(activityId)) {
      activitiesById.set(activityId, {
        id: activityId,
        name: activityInfo.name || activityId,
        category: activityInfo.category || "?????????????",
        entries: []
      });
    }
    const stageValue = normalizeCompetitionStage_(team.stage || "");
    const baseMembers = safeParseJson(team.members, { teachers: [], students: [] });
    const areaMembers = safeParseJson(team.membersArea, null);
    const membersSource =
      stageValue === "area" && areaMembers && (Array.isArray(areaMembers.teachers) || Array.isArray(areaMembers.students))
        ? {
            teachers: Array.isArray(areaMembers.teachers) ? areaMembers.teachers : [],
            students: Array.isArray(areaMembers.students) ? areaMembers.students : []
          }
        : {
            teachers: Array.isArray(baseMembers.teachers) ? baseMembers.teachers : [],
            students: Array.isArray(baseMembers.students) ? baseMembers.students : []
          };
    const teachers = membersSource.teachers;
    const students = membersSource.students;
    const clampedScore = Math.max(0, Math.min(100, Number(score.toFixed(2))));
    const medal = resolveMedalFromScore_(clampedScore);
    const schoolInfo = schoolsIndex
      ? lookupSchoolByName_(schoolsIndex, team.school || "")
      : null;
    const clusterKeyRaw =
      (schoolInfo && (schoolInfo.clusterId || schoolInfo.cluster)) ||
      team.schoolClusterId ||
      team.schoolCluster ||
      "";
    const clusterKey = normalizeKey(clusterKeyRaw);
    const clusterLabel =
      (schoolInfo && (schoolInfo.cluster || schoolInfo.clusterName)) ||
      team.schoolClusterName ||
      team.schoolCluster ||
      "ไม่ระบุเครือข่าย";
    const displayTeamName =
      stageValue === "area" && team.teamNameArea
        ? team.teamNameArea
        : team.teamName || team.teamNameCluster || "-";
    activitiesById.get(activityId).entries.push({
      teamId: team.teamId,
      teamName: displayTeamName,
      school: team.school || "-",
      level: team.level || "",
      score: clampedScore,
      medalKey: medal.key,
      medalLabel: medal.label,
      teachers: teachers.map(t => t.name || "").filter(Boolean),
      students: students.map(s => s.name || "").filter(Boolean),
      clusterKey,
      clusterLabel
    });
    medalCounts[medal.key] = (medalCounts[medal.key] || 0) + 1;
    totalParticipants++;
  });

  const activities = Array.from(activitiesById.values())
    .map(activity => {
      activity.entries.sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return (a.teamName || "").localeCompare(b.teamName || "", "th");
      });
      const seenClusters = new Set();
      const clusterChampions = [];
      activity.entries.forEach(entry => {
        const clusterKeyNormalized =
          normalizeKey(entry.clusterKey || entry.clusterLabel || "unassigned") || "unassigned";
        if (!seenClusters.has(clusterKeyNormalized)) {
          seenClusters.add(clusterKeyNormalized);
          entry.isRepresentative = true;
          clusterChampions.push({
            clusterKey: entry.clusterKey || clusterKeyNormalized,
            clusterLabel: entry.clusterLabel || "ไม่ระบุเครือข่าย",
            teamId: entry.teamId,
            teamName: entry.teamName,
            school: entry.school,
            score: entry.score,
            medalKey: entry.medalKey,
            medalLabel: entry.medalLabel
          });
        } else {
          entry.isRepresentative = false;
        }
      });
      activity.entries = activity.entries.map((entry, index) => ({
        ...entry,
        rank: index + 1,
        isRepresentative: Boolean(entry.isRepresentative)
      }));
      return activity;
    })
    .sort((a, b) => a.name.localeCompare(b.name, "th"));
  const areaFinalsCollection = buildAreaFinalsFromTeams_(teams, activityMap, schoolsIndex);

  const representatives = [];
  activities.forEach(activity => {
    activity.entries.forEach(entry => {
      if (entry.isRepresentative) {
        representatives.push({
          activityId: activity.id,
          activityName: activity.name,
          category: activity.category,
          teamId: entry.teamId,
          teamName: entry.teamName,
          school: entry.school,
          score: entry.score,
          medalKey: entry.medalKey,
          medalLabel: entry.medalLabel,
          clusterKey: entry.clusterKey,
          clusterLabel: entry.clusterLabel
        });
      }
    });
  });

  return {
    activities,
    summary: {
      totalActivities: activities.length,
      totalParticipants,
      medals: medalCounts,
      representatives: representatives.length,
      clusterRepresentatives: representatives.length,
      areaFinalActivities: areaFinalsCollection.length
    },
    representatives,
    areaFinals: areaFinalsCollection
  };
}

function buildAreaFinalsFromTeams_(teams, activityMap, schoolsIndex) {
  if (!Array.isArray(teams) || !teams.length) return [];
  const areaMap = new Map();
  teams.forEach(team => {
    const activityId = (team.activity || "").toString().trim();
    if (!activityId) return;
    const activityInfo = activityMap.get(activityId) || {};
    const areaScore = parseScoreValue_(team.areaScore);
    const areaRankRaw = parseInt((team.areaRank || "").toString().trim(), 10);
    const areaRank = Number.isFinite(areaRankRaw) ? areaRankRaw : null;
    if (areaScore === null && areaRank === null) return;

    const schoolInfo = lookupSchoolByName_(schoolsIndex, team.school || "") || {};
    const clusterLabel =
      schoolInfo.cluster ||
      schoolInfo.clusterName ||
      team.schoolCluster ||
      team.schoolClusterName ||
      "ไม่ระบุเครือข่าย";
    const stageValue = normalizeCompetitionStage_(team.stage || "");
    const displayTeamName =
      stageValue === "area" && team.teamNameArea
        ? team.teamNameArea
        : team.teamName || team.teamNameCluster || "-";
    const medal = resolveMedalFromScore_(areaScore);

    if (!areaMap.has(activityId)) {
      areaMap.set(activityId, {
        activityId,
        activityName: activityInfo.name || activityId,
        category: activityInfo.category || "ไม่ระบุหมวดหมู่",
        finalists: []
      });
    }

    areaMap.get(activityId).finalists.push({
      teamId: team.teamId,
      teamName: displayTeamName,
      school: team.school || "-",
      clusterLabel,
      score: areaScore,
      medalKey: medal.key,
      medalLabel: medal.label,
      areaRank
    });
  });

  return Array.from(areaMap.values())
    .map(record => {
      const sorted = record.finalists
        .slice()
        .sort((a, b) => {
          if (a.areaRank !== null && b.areaRank !== null) {
            return a.areaRank - b.areaRank;
          }
          if (a.areaRank !== null) return -1;
          if (b.areaRank !== null) return 1;
          if (a.score !== null && b.score !== null) {
            return b.score - a.score;
          }
          if (a.score !== null) return -1;
          if (b.score !== null) return 1;
          return (a.teamName || "").localeCompare(b.teamName || "", "th");
        })
        .map((entry, index) => ({
          ...entry,
          areaRank: entry.areaRank !== null ? entry.areaRank : index + 1
        }));
      return { ...record, finalists: sorted };
    })
    .filter(record => record.finalists.length)
    .sort((a, b) => (a.activityName || "").localeCompare(b.activityName || "", "th"));
}

function buildClusterLeaderboard_(teams, schoolsIndex) {
  const clusterMap = new Map();
  if (!Array.isArray(teams)) return [];
  teams.forEach(team => {
    const score = parseScoreValue_(team.scoreTotal);
    if (score === null) return;
    const medal = resolveMedalFromScore_(score);
    const schoolName = team.school || "ไม่ระบุสถานศึกษา";
    const schoolInfo =
      lookupSchoolByName_(schoolsIndex, schoolName) ||
      null;
    const clusterKey =
      (schoolInfo && (schoolInfo.clusterId || schoolInfo.cluster)) || "unassigned";
    const clusterLabel =
      (schoolInfo && schoolInfo.cluster) || "ไม่ระบุเครือข่าย";
    if (!clusterMap.has(clusterKey)) {
      clusterMap.set(clusterKey, {
        key: clusterKey,
        label: clusterLabel,
        schools: new Map()
      });
    }
    const clusterEntry = clusterMap.get(clusterKey);
    if (!clusterEntry.schools.has(schoolName)) {
      clusterEntry.schools.set(schoolName, {
        name: schoolName,
        medals: { gold: 0, silver: 0, bronze: 0, merit: 0, participant: 0 },
        totalScore: 0,
        entries: 0
      });
    }
    const schoolEntry = clusterEntry.schools.get(schoolName);
    schoolEntry.medals[medal.key] = (schoolEntry.medals[medal.key] || 0) + 1;
    schoolEntry.totalScore += score;
    schoolEntry.entries += 1;
  });
  const sortSchools = (a, b) => {
    const medals = ["gold", "silver", "bronze", "merit"];
    for (let i = 0; i < medals.length; i++) {
      const diff = (b.medals[medals[i]] || 0) - (a.medals[medals[i]] || 0);
      if (diff !== 0) return diff;
    }
    return (b.totalScore || 0) - (a.totalScore || 0);
  };
  return Array.from(clusterMap.values())
    .map(cluster => {
      const schools = Array.from(cluster.schools.values())
        .map(entry => ({
          name: entry.name,
          medals: entry.medals,
          totalScore: Number(entry.totalScore.toFixed(2)),
          entries: entry.entries
        }))
        .sort(sortSchools);
      return {
        key: cluster.key,
        label: cluster.label,
        schools
      };
    })
    .filter(cluster => cluster.schools.length > 0);
}

function saveActivityScores(request) {
  try {
    const data = request || {};
    const activityId = (data.activityId || "").toString().trim();
    if (!activityId) throw new Error("กรุณาระบุกิจกรรม");
    const scorePayloads = Array.isArray(data.scores) ? data.scores : [];
    if (!scorePayloads.length) throw new Error("กรุณากรอกคะแนนอย่างน้อย 1 ทีม");
    const requestedRepresentativeId = (data.representativeTeamId || "").toString().trim();
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) {
      throw new Error("คุณไม่มีสิทธิ์บันทึกคะแนน");
    }
    const level = actor.normalizedLevel;
    const competitionStage = getCompetitionStage_();
    const isScoreUser = level === "score";
    const isGroupAdmin = level === "group_admin";
    const isAdminAreaScorer = ["admin", "area"].includes(level) && competitionStage === "area";
    if (!isScoreUser && !isGroupAdmin && !isAdminAreaScorer) {
      throw new Error("คุณไม่มีสิทธิ์บันทึกคะแนน");
    }
    const scoreAssignments = isScoreUser ? getScoreAssignmentMap_(spreadsheet) : new Map();
    const assignedActivities = isScoreUser ? scoreAssignments.get(actor.userId) || [] : [];
    const normalizedTarget = normalizeKey(activityId);
    let canEditActivity = true;
    if (isScoreUser) {
      canEditActivity = assignedActivities.some(
        assigned => normalizeKey(assigned) === normalizedTarget
      );
    } else if (isGroupAdmin && !actor.clusterNormalized) {
      throw new Error("ไม่พบข้อมูลเครือข่ายของคุณ");
    }
    if (!canEditActivity) {
      throw new Error("กิจกรรมนี้ไม่ได้มอบหมายให้คุณ");
    }
    const sheet = spreadsheet.getSheetByName(SHEET_TEAMS);
    if (!sheet) throw new Error("??? Sheet 'Teams'");
    ensureTeamSheetColumns_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("ยังไม่มีข้อมูลทีม");
    const normalizedScores = [];
    scorePayloads.forEach(item => {
      const teamId = (item.teamId || item.teamID || "").toString().trim();
      const score = parseFloat(item.score);
      if (!teamId) return;
      if (isNaN(score) || score < 0 || score > 100) return;
      normalizedScores.push({
        teamId,
        score: Math.round(score * 100) / 100
      });
    });
    if (!normalizedScores.length) {
      throw new Error("ข้อมูลคะแนนไม่ถูกต้อง");
    }
    const dataRange = sheet.getRange(2, 1, lastRow - 1, TEAM_SHEET_MAX_COLS);
    const values = dataRange.getValues();
    const teamIndexMap = new Map();
    values.forEach((row, idx) => {
      const rowTeamId = (row[0] || "").toString().trim();
      if (rowTeamId) {
        teamIndexMap.set(rowTeamId, idx);
      }
    });
    const disallowedTeams = [];
    const targetStageMode = competitionStage === "area" ? "area" : "cluster";
    const scoreColumnIndex = targetStageMode === "area" ? TEAM_AREA_SCORE_COLUMN_INDEX - 1 : 15;
    normalizedScores.forEach(entry => {
      if (!teamIndexMap.has(entry.teamId)) return;
      const idx = teamIndexMap.get(entry.teamId);
      const row = values[idx];
      if (normalizeKey(row[1]) !== normalizedTarget) return;
      if (isGroupAdmin) {
        const schoolName = (row[3] || "").toString().trim();
        const schoolInfo = lookupSchoolByName_(schoolsIndex, schoolName);
        const rowClusterKey = normalizeKey(schoolInfo?.cluster || "");
        if (!rowClusterKey || rowClusterKey !== actor.clusterNormalized) {
          disallowedTeams.push(entry.teamId);
          return;
        }
      }
      row[scoreColumnIndex] = entry.score;
      if (targetStageMode === "cluster") {
        row[17] = "";
        row[18] = "";
      }
    });
    if (disallowedTeams.length) {
      throw new Error("พบทีมที่อยู่นอกเครือข่ายของคุณ ไม่สามารถบันทึกคะแนนได้");
    }
    recalculateActivityScores_(values, activityId, requestedRepresentativeId, { stage: targetStageMode });
    dataRange.setValues(values);
    return { success: true };
  } catch (error) {
    Logger.log("saveActivityScores error: " + error);
    return { success: false, error: error.message };
  }
}


function recalculateActivityScores_(rows, activityId, manualRepresentativeId, options = {}) {
  if (!Array.isArray(rows) || !activityId) return;
  const normalizedActivity = normalizeKey(activityId);
  const manualRepKey = normalizeKey(manualRepresentativeId || "");
  const stageMode = options.stage === "area" ? "area" : "cluster";
  const scoreColumnIndex = stageMode === "area" ? TEAM_AREA_SCORE_COLUMN_INDEX - 1 : 15;
  const stageColumnIndex = TEAM_STAGE_COLUMN_INDEX - 1;
  const areaScoreIndex = TEAM_AREA_SCORE_COLUMN_INDEX - 1;
  const areaRankIndex = TEAM_AREA_RANK_COLUMN_INDEX - 1;
  const entries = [];
  rows.forEach((row, idx) => {
    if (normalizeKey(row[1]) !== normalizedActivity) {
      return;
    }
    const score = parseScoreValue_(row[scoreColumnIndex]);
    if (score === null) {
      if (stageMode === "cluster") {
        row[17] = "";
        row[18] = "";
      }
      if (stageColumnIndex >= 0) {
        row[stageColumnIndex] = "cluster";
      }
      return;
    }
    const teamId = (row[0] || "").toString().trim();
    entries.push({ idx, score, teamId });
  });
  if (!entries.length) return;
  entries.sort((a, b) => b.score - a.score);
  const topScore = entries[0].score;
  const tiedTopEntries =
    topScore === null || topScore === undefined
      ? []
      : entries.filter(entry => entry.score === topScore);
  let representativeKey = "";
  if (tiedTopEntries.length > 1 && manualRepKey) {
    const manualMatch = tiedTopEntries.find(
      entry => normalizeKey(entry.teamId) === manualRepKey
    );
    if (manualMatch) {
      representativeKey = normalizeKey(manualMatch.teamId);
    }
  }
  entries.forEach((entry, rankIdx) => {
    const rank = rankIdx + 1;
    const isChampion = representativeKey
      ? normalizeKey(entry.teamId) === representativeKey
      : rank === 1;

    if (stageMode === "cluster") {
      rows[entry.idx][17] = rank;
      rows[entry.idx][18] = isChampion ? "TRUE" : "";
      if (stageColumnIndex >= 0) {
        rows[entry.idx][stageColumnIndex] = "cluster";
      }
      rows[entry.idx][areaScoreIndex] = "";
      rows[entry.idx][areaRankIndex] = "";
    } else {
      if (stageColumnIndex >= 0) {
        rows[entry.idx][stageColumnIndex] = "area";
      }
      rows[entry.idx][areaScoreIndex] = entry.score;
      rows[entry.idx][areaRankIndex] = rank;
    }
  });
}

function filterTeamsByActor_(teams, actor, schoolLookup) {
  if (!Array.isArray(teams) || !teams.length || !actor) return [];
  const level = normalizeKey(actor.level || "");
  const schoolKey = normalizeKey(actor.schoolName || "");
  const clusterKey = normalizeKey(actor.clusterId || actor.cluster || "");
  if (["admin", "area"].includes(level)) {
    return teams;
  }
  if (level === "group_admin") {
    if (clusterKey) {
      const lookup = schoolLookup || buildSchoolNameClusterLookup_();
      return teams.filter(team => {
        const teamSchoolKey = normalizeKey(team.school || "");
        const info = lookup[teamSchoolKey];
        if (!info) return false;
        return info.clusterKey && info.clusterKey === clusterKey;
      });
    }
    if (schoolKey) {
      return teams.filter(team => normalizeKey(team.school || "") === schoolKey);
    }
    return [];
  }
  if (level === "school_admin") {
    if (!schoolKey) return [];
    return teams.filter(team => normalizeKey(team.school || "") === schoolKey);
  }
  const actorUserId = (actor.userId || "").toString().trim();
  if (!actorUserId) return [];
  return teams.filter(team => (team.createdByUserId || "").toString().trim() === actorUserId);
}

function buildSchoolNameClusterLookup_(spreadsheet) {
  const lookup = {};
  try {
    const book = spreadsheet || SpreadsheetApp.openById(SHEET_ID);
    const sheet = book.getSheetByName(SHEET_SCHOOLS);
    if (!sheet) return lookup;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return lookup;
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach(row => {
      const nameKey = normalizeKey(row[1]);
      if (!nameKey) return;
      const clusterRaw = (row[2] || "").toString().trim();
      lookup[nameKey] = {
        clusterKey: normalizeKey(clusterRaw),
        clusterLabel: clusterRaw
      };
    });
  } catch (error) {
    Logger.log("buildSchoolNameClusterLookup_ error: " + error);
  }
  return lookup;
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
    if (!teamsSheet) throw new Error("??? Sheet 'Teams'");
    if (!activitiesSheet) throw new Error("??? Sheet 'Activities'");
    ensureTeamSheetColumns_(teamsSheet);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);

    const totalRows = teamsSheet.getLastRow();
    const existingRows = totalRows >= 2
      ? teamsSheet.getRange(2, 1, totalRows - 1, 2).getValues()
      : [];
    const existingTeamIdSet = new Set(
      existingRows
        .map(row => row[0])
        .filter(id => Boolean(id))
    );

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
      if (existingRows.length) {
        existingCount = existingRows.reduce((count, row) => {
          return row[1] === formData.activity ? count + 1 : count;
        }, 0);
      }
      if (existingCount >= maxTeams) {
        return {
          success: false,
          error: "กิจกรรมนี้ปิดรับสมัครแล้ว (จำนวนทีมเต็ม)"
        };
      }
    }

    const actorContextInput = formData.createdByUserId
      ? { userId: formData.createdByUserId }
      : null;
    const actor = actorContextInput ? resolveActorContext_(actorContextInput, schoolsIndex) : null;
    const actorLevel = actor ? actor.normalizedLevel : "";
    const privilegedLevels = ["admin", "area", "group_admin"];
    const isPrivilegedActor = actor && privilegedLevels.includes(actorLevel);
    const schoolInput = (formData.school || "").toString().trim();
    const schoolIdentifiers = parseSchoolIdentifiers_(schoolInput);
    let effectiveSchool = findSchoolRecordByInput_(schoolsIndex, schoolInput);
    if (!effectiveSchool && actor) {
      effectiveSchool =
        lookupSchoolById_(schoolsIndex, actor.schoolId) ||
        lookupSchoolByName_(schoolsIndex, actor.schoolName);
    }
    if (actor && !isPrivilegedActor) {
      const actorSchoolNameNormalized = normalizeKey(
        actor.schoolName || (effectiveSchool && effectiveSchool.name) || ""
      );
      const inputNameNormalized = normalizeKey(schoolIdentifiers.name || schoolInput);
      if (
        actorSchoolNameNormalized &&
        inputNameNormalized &&
        actorSchoolNameNormalized !== inputNameNormalized
      ) {
        throw new Error("คุณสามารถลงทะเบียนทีมให้เฉพาะโรงเรียนของคุณเท่านั้น");
      }
    }
    if (effectiveSchool && !isPrivilegedActor) {
      if (isGroupAssignedMode_(effectiveSchool.registrationMode)) {
        const assignedList = Array.isArray(effectiveSchool.assignedActivities)
          ? effectiveSchool.assignedActivities
          : [];
        if (!assignedList.length) {
          throw new Error("โรงเรียนของคุณอยู่ในโหมดให้เครือข่ายกำหนดกิจกรรม กรุณาติดต่อ Group Admin");
        }
        if (!isActivityAllowedForSchool_(effectiveSchool, formData.activity)) {
          throw new Error("กิจกรรมนี้ไม่ได้รับอนุญาตให้โรงเรียนของคุณ กรุณาเลือกกิจกรรมที่ได้รับมอบหมาย");
        }
      }
    }

    // สร้าง TeamID ป้องกันซ้ำ (timestamp + suffix)
    const buildTeamId = () => {
      const stamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMMddHHmmss");
      const suffix = Utilities.getUuid().replace(/-/g, "").slice(-4).toUpperCase();
      return `T${stamp}${suffix}`;
    };
    let newTeamId = buildTeamId();
    let guard = 0;
    while (existingTeamIdSet.has(newTeamId) && guard < 5) {
      Utilities.sleep(30);
      newTeamId = buildTeamId();
      guard++;
    }
    if (existingTeamIdSet.has(newTeamId)) {
      throw new Error("ไม่สามารถสร้างรหัสทีมที่ไม่ซ้ำได้ โปรดลองอีกครั้ง");
    }

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
    const initialCompetitionStage = "cluster";

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
    // O StatusReason
    // P ScoreTotal
    // Q MedalOverride
    // R RankOverride
    // S RepresentativeOverride
    // T CompetitionStage
    // U AreaTeamName
    // V AreaContact
    // W AreaMembers
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
      createdByUsername,
      "",
      "",
      "",
      "",
      "",
      initialCompetitionStage,
      "",
      "",
      "",
      "",
      ""
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

    ensureSchoolsSheetStructure_(schoolSheet);
    var clusterMap = buildSchoolClusterMap_(book);
    var lastRow = schoolSheet.getLastRow();
    if (lastRow < 2) return index;

    var rows = schoolSheet.getRange(2, 1, lastRow - 1, SCHOOL_SHEET_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      var id = String(rows[i][0] || "").trim();
      var name = String(rows[i][1] || "").trim();
      var clusterRaw = String(rows[i][2] || "").trim();
      var registrationMode = normalizeRegistrationMode_(rows[i][3] || "");
      var assignedActivities = parseAssignedActivitiesCell_(rows[i][4]);
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
        cluster: clusterName,
        registrationMode: registrationMode,
        assignedActivities: assignedActivities
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

function findSchoolRowIndex_(sheet, schoolId) {
  if (!sheet || !schoolId) {
    return -1;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const target = normalizeKey(schoolId);
  for (var i = 0; i < ids.length; i++) {
    const current = normalizeKey(String(ids[i][0] || ""));
    if (current && current === target) {
      return i + 2;
    }
  }
  return -1;
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
    teamNameCluster: getValue(2),
    teamNameArea: getValue(TEAM_AREA_NAME_COLUMN_INDEX - 1),
    school: getValue(3),
    level: getValue(4),
    status: getValue(9),
    createdByUserId: getValue(12),
    stage: normalizeCompetitionStage_(rowValues[TEAM_STAGE_COLUMN_INDEX - 1] || "")
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
  if (level === "admin" || level === "area") {
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
    if (!sheet) throw new Error("??? Sheet 'Teams'");
    ensureTeamSheetColumns_(sheet);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) {
      throw new Error("กรุณาเข้าสู่ระบบก่อนแก้ไขทีม");
    }
    const finder = sheet.getRange("A:A").createTextFinder(teamId).findNext();
    if (!finder) throw new Error("ไม่พบทีมนี้");

    const rowIndex = finder.getRow();
    const lastCol = Math.min(sheet.getLastColumn(), TEAM_SHEET_MAX_COLS);
    const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
    const stageColumnIndex = TEAM_STAGE_COLUMN_INDEX - 1;
    const teamRecord = buildTeamRecordFromRow_(rowValues, schoolsIndex);
    const permissions = evaluateTeamPermission_(actor, teamRecord);
    if (!permissions.canManage) {
      throw new Error("ไม่มีสิทธิ์แก้ไขทีมนี้");
    }

    const teamStageValue = teamRecord.stage || normalizeCompetitionStage_(rowValues[TEAM_STAGE_COLUMN_INDEX - 1] || "");
    const requestedScope = normalizeCompetitionStage_(data.stageScope || "");
    const competitionStage = getCompetitionStage_();
    const allowAreaScope = competitionStage === "area";
    let stageScope = "cluster";
    if (requestedScope === "area" && (allowAreaScope || teamStageValue === "area")) {
      stageScope = "area";
    } else if (teamStageValue === "area") {
      stageScope = "area";
    }
    const contactColumnIndex = stageScope === "area" ? TEAM_AREA_CONTACT_COLUMN_INDEX - 1 : 5;
    const membersColumnIndex = stageScope === "area" ? TEAM_AREA_MEMBERS_COLUMN_INDEX - 1 : 6;
    const nameColumnIndex = stageScope === "area" ? TEAM_AREA_NAME_COLUMN_INDEX - 1 : 2;

    const existingContact = safeParseJson(rowValues[contactColumnIndex], {});
    const memberData = safeParseJson(rowValues[membersColumnIndex], { teachers: [], students: [] });

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
    if (typeof data.teamName === "string" && data.teamName.trim()) {
      if (stageScope === "area") {
        updatedRow[nameColumnIndex] = data.teamName.trim();
      } else {
        updatedRow[2] = data.teamName.trim();
      }
    }
    updatedRow[contactColumnIndex] = JSON.stringify(nextContact);
    updatedRow[membersColumnIndex] = JSON.stringify(sanitizedMembers);
    const nextStatus = (data.status || "").toString().trim();
    const nextStatusReason =
      typeof data.statusReason === "string" ? data.statusReason.trim() : "";
    const currentStatus = (updatedRow[9] || "").toString();
    const normalizedNextStatus = nextStatus.toLowerCase();
    const actorLevel = actor.normalizedLevel || "";
    if (nextStatus) {
      if (permissions.canEditStatus) {
        updatedRow[9] = nextStatus;
      } else if (
        normalizedNextStatus === "rejected" &&
        actorLevel === "group_admin"
      ) {
        if (!nextStatusReason) {
          throw new Error("กรุณาระบุเหตุผลเมื่อปฏิเสธทีม");
        }
        updatedRow[9] = "Rejected";
      } else if (nextStatus !== currentStatus) {
        throw new Error("สิทธิ์ของคุณไม่สามารถเปลี่ยนสถานะนี้ได้");
      }
    }
    const statusIsRejected =
      (updatedRow[9] || "").toString().toLowerCase() === "rejected";
    if (statusIsRejected) {
      if (nextStatusReason) {
        updatedRow[14] = nextStatusReason;
      } else if (typeof updatedRow[14] !== "string") {
        updatedRow[14] = "";
      }
    } else {
      updatedRow[14] = "";
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

    if (stageColumnIndex >= 0) {
      updatedRow[stageColumnIndex] = stageScope;
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
    if (!sheet) throw new Error("??? Sheet 'Teams'");
    ensureTeamSheetColumns_(sheet);
    const finder = sheet.getRange("A:A").createTextFinder(targetId).findNext();
    if (!finder) throw new Error("ไม่พบทีมนี้");
    const rowIndex = finder.getRow();
    const lastCol = Math.min(sheet.getLastColumn(), TEAM_SHEET_MAX_COLS);
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


      const csvContent = "\uFEFF" + [headers, ...csvRows]
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
 * (ฟังก์ชันใหม่)
 * รับข้อมูลทีมที่กรองแล้ว (จาก Client) มาสร้าง CSV
 * พร้อมแก้ปัญหาภาษาไทยใน Excel
 */
function exportFilteredCsv(teams, activities) {
  try {
    if (!Array.isArray(teams)) {
      teams = [];
    }
    if (!Array.isArray(activities)) {
      activities = [];
    }

    const activityMap = {};
    activities.forEach(activity => {
      if (activity && activity.id) {
        activityMap[activity.id] = activity;
      }
    });

    // ดึงข้อมูลจากฟังก์ชัน helper ที่มีอยู่แล้ว [cite: 15]
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
        activityName: activity.name || team.activityName || "", // ใช้อันที่ Client อาจจะส่งมา
        category: activity.category || "",
        teamName: team.teamName || "",
        school: team.school || "",
        level: team.level || "",
        contactName: contact.name || "",
        contactPhone: contact.phone || "",
        contactEmail: contact.email || "",
        requiredTeachers: isNaN(parsedRequiredTeachers)
          ? (activity.reqTeachers || teachers)
          : parsedRequiredTeachers,
        actualTeachers: teachers,
        requiredStudents: isNaN(parsedRequiredStudents)
          ? (activity.reqStudents || students)
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
    
    // --- ส่วนของ CSV ---
    const headers = [
        "TeamID", "ActivityID", "ActivityName", "Category", "TeamName",
        "School", "Level", "ContactName", "ContactPhone", "ContactEmail",
        "RequiredTeachers", "ActualTeachers", "RequiredStudents", "ActualStudents", "Status"
    ];
    
    const csvRows = rows.map(row => [
        row.teamId, row.activityId, row.activityName, row.category, row.teamName,
        row.school, row.level, row.contactName, row.contactPhone, row.contactEmail,
        row.requiredTeachers, row.actualTeachers, row.requiredStudents, row.actualStudents, row.status
    ]);
    
    // *** แก้ปัญหาภาษาไทยใน Excel (เพิ่ม \uFEFF) ***
    const csvContent = "\uFEFF" + [headers, ...csvRows]
        .map(row => row.map(escapeCsvCell).join(",")) // [cite: 17]
        .join("\r\n");
        
    const fileName = "teams_filtered_" + timestamp + ".csv";
    const blob = Utilities.newBlob(csvContent, "text/csv", fileName);
    
    return {
        success: true,
        fileName: fileName,
        mimeType: "text/csv",
        data: Utilities.base64Encode(blob.getBytes())
    };

  } catch (error) {
    Logger.log("exportFilteredCsv error: " + error);
    return {
      success: false,
      error: error.message || "ไม่สามารถส่งออกข้อมูลได้"
    };
  }
}


/**
 * (ฟังก์ชันแก้ไข)
 * สร้าง CSV (แก้ภาษาไทย) หรือ PDF (แนวนอน)
 * *** เพิ่มการจัดเรียงตามหมวดหมู่และกิจกรรม ***
 */
// function exportFilteredData(teams, activities, format) {
//   try {
//     if (!Array.isArray(teams)) {
//       teams = [];
//     }
//     if (!Array.isArray(activities)) {
//       activities = [];
//     }

//     const activityMap = {};
//     activities.forEach(activity => {
//       if (activity && activity.id) {
//         activityMap[activity.id] = activity;
//       }
//     });

//     // 1. เตรียมข้อมูล (เหมือนเดิม)
//     const rows = teams.map(team => {
//       const activity = activityMap[team.activity] || {};
//       const contact = safeParseJson(team.contact, {});
//       const members = safeParseJson(team.members, { teachers: [], students: [] });
//       const teachers = Array.isArray(members.teachers) ? members.teachers.length : 0;
//       const students = Array.isArray(members.students) ? members.students.length : 0;
      
//       const parsedRequiredTeachers = parseInt(team.requiredTeachers, 10);
//       const parsedRequiredStudents = parseInt(team.requiredStudents, 10);

//       return {
//         teamId: team.teamId || "",
//         activityId: team.activity || "",
//         activityName: activity.name || team.activityName || "",
//         category: activity.category || "ไม่ระบุหมวดหมู่", // <-- ใส่ค่า default
//         teamName: team.teamName || "",
//         school: team.school || "",
//         level: team.level || "",
//         contactName: contact.name || "",
//         contactPhone: contact.phone || "",
//         contactEmail: contact.email || "",
//         requiredTeachers: isNaN(parsedRequiredTeachers)
//           ? (activity.reqTeachers || teachers)
//           : parsedRequiredTeachers,
//         actualTeachers: teachers,
//         requiredStudents: isNaN(parsedRequiredStudents)
//           ? (activity.reqStudents || students)
//           : parsedRequiredStudents,
//         actualStudents: students,
//         status: team.status || "Pending"
//       };
//     });

//     // --- ??? นี่คือส่วนที่เพิ่มเข้ามา ??? ---
//     // 2. จัดเรียงข้อมูล
//     rows.sort((a, b) => {
//       // เรียงลำดับที่ 1: ตามหมวดหมู่ (Category)
//       const categoryCompare = a.category.localeCompare(b.category, 'th');
//       if (categoryCompare !== 0) {
//         return categoryCompare;
//       }
      
//       // เรียงลำดับที่ 2: ตามชื่อกิจกรรม (ActivityName)
//       const activityCompare = a.activityName.localeCompare(b.activityName, 'th');
//       if (activityCompare !== 0) {
//         return activityCompare;
//       }
      
//       // เรียงลำดับที่ 3: ตามชื่อทีม (TeamName) (เผื่อกิจกรรมเดียวกันมีหลายทีม)
//       return (a.teamName || a.teamId).localeCompare(b.teamName || b.teamId, 'th');
//     });
//     // --- ??? สิ้นสุดส่วนที่เพิ่ม ??? ---


//     // 3. สร้างไฟล์ (เหมือนเดิม)
//     const timestamp = Utilities.formatDate(
//       new Date(),
//       "Asia/Bangkok",
//       "yyyyMMdd_HHmmss"
//     );
    
//     // --- แยกการทำงาน CSV / PDF ---

//     if (format === "csv") {
//         const headers = [
//             "TeamID", "ActivityID", "ActivityName", "Category", "TeamName",
//             "School", "Level", "ContactName", "ContactPhone", "ContactEmail",
//             "RequiredTeachers", "ActualTeachers", "RequiredStudents", "ActualStudents", "Status"
//         ];
        
//         const csvRows = rows.map(row => [
//             row.teamId, row.activityId, row.activityName, row.category, row.teamName,
//             row.school, row.level, row.contactName, row.contactPhone, row.contactEmail,
//             row.requiredTeachers, row.actualTeachers, row.requiredStudents, row.actualStudents, row.status
//         ]);
        
//         // *** แก้ปัญหาภาษาไทยใน Excel (เพิ่ม \uFEFF) ***
//         const csvContent = "\uFEFF" + [headers, ...csvRows]
//             .map(row => row.map(escapeCsvCell).join(",")) //
//             .join("\r\n");
            
//         const fileName = "teams_filtered_" + timestamp + ".csv";
//         const blob = Utilities.newBlob(csvContent, "text/csv", fileName);
        
//         return {
//             success: true,
//             fileName: fileName,
//             mimeType: "text/csv",
//             data: Utilities.base64Encode(blob.getBytes())
//         };
//     }

//     if (format === "pdf") {
//         // --- ใช้โค้ดสร้าง PDF แนวนอน (เหมือนเดิม) ---
//         const tableRows = rows
//             .map(
//               (row, index) => `
//             <tr>
//               <td>${index + 1}</td>
//               <td>${escapeHtml(row.teamId)}</td>
//               <td>${escapeHtml(row.teamName)}</td>
//               <td>${escapeHtml(row.activityName)}</td>
//               <td>${escapeHtml(row.school)}</td>
//               <td>${escapeHtml(row.level)}</td>
//               <td>${escapeHtml(row.contactName)}<br>${escapeHtml(
//                 row.contactPhone
//               )}<br>${escapeHtml(row.contactEmail)}</td>
//               <td>${row.actualTeachers}/${row.requiredTeachers}</td>
//               <td>${row.actualStudents}/${row.requiredStudents}</td>
//               <td>${escapeHtml(row.status)}</td>
//             </tr>`
//             )
//             .join("");

//         // (โค้ด HTML นี้ดึงมาจากฟังก์ชัน exportTeams เดิมของคุณ)
//         const html = `
//             <html>
//               <head>
//                 <meta charset="UTF-8">
//                 <style>
//                   /* --- ตั้งค่าแนวนอน --- */
//                   @page { 
//                     size: A4 landscape; 
//                   }
                  
//                   body { 
//                     font-family: Arial, "TH SarabunPSK", sans-serif; 
//                     font-size: 10pt; /* <-- ลดขนาดฟอนต์เล็กน้อย */
//                   }
//                   h2 { margin-bottom: 6px; }
//                   table { width: 100%; border-collapse: collapse; }
//                   th, td { 
//                     border: 1px solid #cccccc; 
//                     padding: 4px; /* <-- ลด Padding เล็กน้อย */
//                     vertical-align: top; 
//                   }
//                   th { background-color: #f1f5f9; }
//                 </style>
//               </head>
//               <body>
//                 <h2>รายงานทีมที่ลงทะเบียน (กรองตามสิทธิ์)</h2>
//                 <p>อัปเดตล่าสุด: ${Utilities.formatDate(
//                   new Date(),
//                   "Asia/Bangkok",
//                   "dd MMM yyyy HH:mm"
//                 )}</p>
//                 <p>จำนวนทีมทั้งหมด: ${rows.length} ทีม</p>
//                 <table>
//                   <thead>
//                     <tr>
//                       <th>#</th>
//                       <th>Team ID</th>
//                       <th>ชื่อทีม</th>
//                       <th>กิจกรรม</th>
//                       <th>สถานศึกษา</th>
//                       <th>ระดับ</th>
//                       <th>ผู้ประสานงาน</th>
//                       <th>ครู (จริง/ต้องการ)</th>
//                       <th>นักเรียน (จริง/ต้องการ)</th>
//                       <th>สถานะ</th>
//                     </tr>
//                   </thead>
//                   <tbody>
//                     ${tableRows}
//                   </tbody>
//                 </table>
//               </body>
//             </html>
//           `;
          
//         const pdfBlob = Utilities.newBlob(
//             html,
//             "text/html",
//             "teams_export.html"
//           ).getAs("application/pdf");
          
//         const pdfFileName = "teams_filtered_" + timestamp + ".pdf";
//         pdfBlob.setName(pdfFileName);

//         return {
//             success: true,
//             fileName: pdfFileName,
//             mimeType: "application/pdf",
//             data: Utilities.base64Encode(pdfBlob.getBytes())
//         };
//     }

//     // ถ้า format ไม่ใช่ทั้ง csv หรือ pdf
//     throw new Error("รูปแบบไฟล์ไม่รองรับ");

//   } catch (error) {
//     Logger.log("exportFilteredData error: " + error);
//     return {
//       success: false,
//       error: error.message || "ไม่สามารถส่งออกข้อมูลได้"
//     };
//   }
// }

/**
 * (ฟังก์ชันแก้ไข - แก้ไข Syntax Error "H" ที่ตกค้าง)
 * สร้าง CSV (แก้ภาษาไทย) หรือ PDF (แนวนอน)
 * จัดเรียงตามหมวดหมู่และกิจกรรม
 */
function exportFilteredData(teams, activities, format) {
  try {
    if (!Array.isArray(teams)) {
      teams = [];
    }
    if (!Array.isArray(activities)) {
      activities = [];
    }

    const activityMap = {};
    activities.forEach(activity => {
      if (activity && activity.id) {
        activityMap[activity.id] = activity;
      }
    });

    // 1. เตรียมข้อมูล
    const rows = teams.map(team => {
      const activity = activityMap[team.activity] || {};
      const contact = safeParseJson(team.contact, {});
      const members = safeParseJson(team.members, { teachers: [], students: [] }); // 
      const teachers = Array.isArray(members.teachers) ? members.teachers.length : 0;
      const students = Array.isArray(members.students) ? members.students.length : 0;
      
      const parsedRequiredTeachers = parseInt(team.requiredTeachers, 10);
      const parsedRequiredStudents = parseInt(team.requiredStudents, 10);

      return {
        teamId: team.teamId || "",
        activityId: team.activity || "",
        activityName: activity.name || team.activityName || "",
        category: activity.category || "ไม่ระบุหมวดหมู่", 
        teamName: team.teamName || "",
        school: team.school || "",
        level: team.level || "",
        contactName: contact.name || "",
        contactPhone: contact.phone || "",
        contactEmail: contact.email || "",
        requiredTeachers: isNaN(parsedRequiredTeachers)
          ? (activity.reqTeachers || teachers)
          : parsedRequiredTeachers,
        actualTeachers: teachers,
        requiredStudents: isNaN(parsedRequiredStudents)
          ? (activity.reqStudents || students)
          : parsedRequiredStudents,
        actualStudents: students,
        status: team.status || "Pending"
      };
    });

    // 2. จัดเรียงข้อมูล
    rows.sort((a, b) => {
      const categoryCompare = a.category.localeCompare(b.category, 'th');
      if (categoryCompare !== 0) {
        return categoryCompare;
      }
      const activityCompare = a.activityName.localeCompare(b.activityName, 'th');
      if (activityCompare !== 0) {
        return activityCompare;
      }
      return (a.teamName || a.teamId).localeCompare(b.teamName || b.teamId, 'th');
    });

    // 3. สร้างไฟล์
    const timestamp = Utilities.formatDate(
      new Date(),
      "Asia/Bangkok",
      "yyyyMMdd_HHmmss"
    );

    if (format === "csv") {
        const headers = [
            "TeamID", "ActivityID", "ActivityName", "Category", "TeamName",
            "School", "Level", "ContactName", "ContactPhone", "ContactEmail",
            "RequiredTeachers", "ActualTeachers", "RequiredStudents", "ActualStudents", "Status"
        ];
        
        const csvRows = rows.map(row => [
            row.teamId, row.activityId, row.activityName, row.category, row.teamName,
            row.school, row.level, row.contactName, row.contactPhone, row.contactEmail,
            row.requiredTeachers, row.actualTeachers, row.requiredStudents, row.actualStudents, row.status
        ]);
        
        const csvContent = "\uFEFF" + [headers, ...csvRows]
            .map(row => row.map(escapeCsvCell).join(","))
            .join("\r\n");
            
        const fileName = "teams_filtered_" + timestamp + ".csv";
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
                  @page { 
                    size: A4 landscape; 
                  }
                  body { 
                    font-family: Arial, "TH SarabunPSK", sans-serif; 
                    font-size: 10pt; 
                  }
                  h2 { margin-bottom: 6px; }
                  table { width: 100%; border-collapse: collapse; }
                  th, td { 
                    border: 1px solid #cccccc; 
                    padding: 4px; 
                    vertical-align: top; 
                  }
                  th { background-color: #f1f5f9; }
                </style>
              </head>
              <body>
                <h2>รายงานทีมที่ลงทะเบียน (กรองตามสิทธิ์)</h2>
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
          
        // --- ??? นี่คือบรรทัดที่แก้ไข (เปลี่ยน H เป็น +) ??? ---
        const pdfFileName = "teams_filtered_" + timestamp + ".pdf";
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
    Logger.log("exportFilteredData error: " + error);
    return {
      success: false,
      error: error.message || "ไม่สามารถส่งออกข้อมูลได้"
    };
  }
}

/**
 * แบบลงทะเบียนรายงานตัว (A4 แนวตั้ง)
 * - ฟอนต์ TH Sarabun New / Sarabun
 * - แยกหน้า "ครูผู้ฝึกสอน" และ "นักเรียน" ต่อกิจกรรม
 * - หัวกระดาษจัดกึ่งกลาง สไตล์มินิมอล
 * - ตาราง: ลำดับ / ชื่อ-นามสกุล / สถานศึกษา / ลายมือชื่อ / หมายเหตุ
 */
function exportAttendanceSheet(teams, activities) {
  try {
    // ------------------------ ตรวจสอบข้อมูลเข้า ------------------------
    if (!Array.isArray(teams)) teams = [];
    if (!Array.isArray(activities)) activities = [];

    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet); // เผื่อใช้ต่อ (lookup โรงเรียน)

    // ------------------------ ฟังก์ชัน: วันที่ไทย (พ.ศ.) ------------------------
    function formatThaiDateTime_(date) {
      const d = Utilities.formatDate(date, "Asia/Bangkok", "d");
      const mIndex = parseInt(Utilities.formatDate(date, "Asia/Bangkok", "M"), 10) - 1;
      const year = parseInt(Utilities.formatDate(date, "Asia/Bangkok", "yyyy"), 10) + 543;
      const time = Utilities.formatDate(date, "Asia/Bangkok", "HH:mm");
      const months = [
        "มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
        "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"
      ];
      const monthName = months[mIndex] || "";
      return `${d} ${monthName} ${year} เวลา ${time} น.`;
    }
    const printedAtTH = formatThaiDateTime_(new Date());

    // ------------------------ Map กิจกรรม ------------------------
    const activityMap = new Map();
    activities.forEach(a => {
      if (a && a.id) activityMap.set(a.id, a);
    });

    // ------------------------ จัดกลุ่มทีมตามกิจกรรม ------------------------
    const teamsByActivity = new Map();
    teams.forEach(team => {
      const activityId = team.activity || 'unknown';
      if (!teamsByActivity.has(activityId)) {
        teamsByActivity.set(activityId, []);
      }
      teamsByActivity.get(activityId).push(team);
    });

    // ------------------------ เรียงกิจกรรมตามชื่อ ------------------------
    const sortedActivityIds = Array.from(teamsByActivity.keys()).sort((a, b) => {
      const nameA = (activityMap.get(a) || {}).name || a;
      const nameB = (activityMap.get(b) || {}).name || b;
      return nameA.localeCompare(nameB, 'th');
    });

    // ------------------------ เตรียม rows ครู / นักเรียน ต่อกิจกรรม ------------------------
    function buildRowsForActivity_(activityId) {
      const activity = activityMap.get(activityId) || {};
      const teamsInActivity = teamsByActivity.get(activityId) || [];
      const teacherRows = [];
      const studentRows = [];

      teamsInActivity.forEach(team => {
        const schoolName = team.school || 'ไม่ระบุสถานศึกษา';
        const members = safeParseJson(team.members, { teachers: [], students: [] }) || {};

        // ครูผู้ฝึกสอน
        (members.teachers || []).forEach(t => {
          if (!t) return;
          teacherRows.push({
            name: ((t.prefix || '') + ' ' + (t.name || '')).trim(),
            schoolName
          });
        });

        // นักเรียน
        (members.students || []).forEach(s => {
          if (!s) return;
          studentRows.push({
            name: ((s.prefix || '') + ' ' + (s.name || '')).trim(),
            schoolName
          });
        });
      });

      // เรียง: สถานศึกษา > ชื่อ
      function rowSort(x, y) {
        return (x.schoolName || '').localeCompare(y.schoolName || '', 'th') ||
               (x.name || '').localeCompare(y.name || '', 'th');
      }
      teacherRows.sort(rowSort);
      studentRows.sort(rowSort);

      return { activity, teacherRows, studentRows };
    }

    // ------------------------ สร้าง HTML ตารางรายชื่อ ------------------------
    function buildTableHtml_(rows) {
      return `
        <div class="table-block">
          <table class="member-table">
            <thead>
              <tr>
                <th class="col-index">ลำดับ</th>
                <th class="col-name">ชื่อ-นามสกุล</th>
                <th class="col-school">สถานศึกษา</th>
                <th class="col-sign">ลายมือชื่อ</th>
                <th class="col-note">หมายเหตุ</th>
              </tr>
            </thead>
            <tbody>
              ${
                rows.length
                  ? rows.map((row, i) => `
                      <tr>
                        <td class="center">${i + 1}</td>
                        <td>${escapeHtml(row.name || '')}</td>
                        <td>${escapeHtml(row.schoolName || '')}</td>
                        <td class="sig-cell"></td>
                        <td class="note-cell"></td>
                      </tr>
                    `).join('')
                  : `<tr><td colspan="5" class="empty">- ไม่มีข้อมูล -</td></tr>`
              }
            </tbody>
          </table>
        </div>
      `;
    }

    // ------------------------ สร้างหน้า ครู + นักเรียน ต่อกิจกรรม ------------------------
    let htmlPages = '';

    sortedActivityIds.forEach(activityId => {
      const { activity, teacherRows, studentRows } = buildRowsForActivity_(activityId);
      const activityName = activity.name || activityId;
      const activityCategory = activity.category || '';

      const subLine = [
        activityCategory ? `หมวดหมู่: ${activityCategory}` : '',
        `พิมพ์เมื่อ: ${printedAtTH}`
      ].filter(Boolean).join('  •  ');

      // หน้า ครูผู้ฝึกสอน
      htmlPages += `
        <div class="page">
          <header class="page-header">
            <div class="top-title">แบบลงทะเบียนรายงานตัว</div>
            <div class="label-pill">👤 ครูผู้ฝึกสอน</div>
            <h1>${escapeHtml(activityName)}</h1>
            <div class="sub">${escapeHtml(subLine)}</div>
            <div class="divider"></div>
          </header>
          <main class="page-body">
            ${buildTableHtml_(teacherRows)}
          </main>
        </div>
      `;

      // หน้า นักเรียน
      htmlPages += `
        <div class="page">
          <header class="page-header">
            <div class="top-title">แบบลงทะเบียนรายงานตัว</div>
            <div class="label-pill">👥 นักเรียน</div>
            <h1>${escapeHtml(activityName)}</h1>
            <div class="sub">${escapeHtml(subLine)}</div>
            <div class="divider"></div>
          </header>
          <main class="page-body">
            ${buildTableHtml_(studentRows)}
          </main>
        </div>
      `;
    });

    // ------------------------ HTML + CSS หลัก ------------------------
    const html = `
      <html>
        <head>
          <meta charset="UTF-8">
          <style>
            @page {
              size: A4 portrait;
              margin: 1.8cm 1.6cm;
            }
            body {
              font-family: "TH Sarabun New", "Sarabun",
                           system-ui, -apple-system, BlinkMacSystemFont,
                           "Segoe UI", sans-serif;
              font-size: 11pt;
              color: #111827;
            }
            .page {
              page-break-after: always;
            }

            /* ส่วนหัวของแต่ละหน้า */
            .page-header {
              display: flex;
              flex-direction: column;
              align-items: center;
              text-align: center;
              gap: 3px;
              padding-bottom: 8px;
            }
            .top-title {
              font-size: 13pt;
              font-weight: 600;
              letter-spacing: 0.06em;
              color: #111827;
            }
            .label-pill {
              display: inline-flex;
              align-items: center;
              justify-content: center;
              padding: 2px 14px;
              margin-top: 1px;
              font-size: 9pt;
              border-radius: 999px;
              border: 1px solid #E5E7EB;
              color: #4B5563;
              background-color: #F9FAFB;
              gap: 6px;
            }
            h1 {
              margin: 0;
              margin-top: 2px;
              font-size: 17pt;
              font-weight: 600;
              color: #111827;
            }
            .sub {
              font-size: 9pt;
              color: #6B7280;
            }
            .divider {
              margin-top: 6px;
              width: 68%;
              height: 1px;
              background: #C4C4C4; /* เส้นแบ่ง block เดียว เรียบ ๆ */
            }

            /* เนื้อหาหลักของหน้า */
            .page-body {
              margin-top: 8px;
            }

            /* กล่องของตาราง (ตัดเส้นรอบออก กันเส้นคู่) */
            .table-block {
              padding: 2px 0;
              border: none;
              background: #FFFFFF;
            }

            /* ตารางรายชื่อ */
            .member-table {
              width: 100%;
              border-collapse: collapse; /* รวมเส้นให้เป็นเส้นเดียว */
              table-layout: fixed;
            }
            .member-table th,
            .member-table td {
              border: 1px solid #C4C4C4; /* เส้นตารางเส้นเดียวแบบปกติ */
              padding: 4px 6px;
              vertical-align: middle;
            }
            .member-table thead {
              background: #F3F4F6;
            }
            .member-table th {
              font-size: 11pt;
              font-weight: 600;
              color: #374151;
            }
            .member-table td {
              font-size: 10pt;
              color: #111827;
            }
            .member-table tr:nth-child(even) td {
              background: #FAFAFA;
            }

            /* ปรับความกว้างคอลัมน์: เซ็นชื่อกว้าง, หมายเหตุแคบ */
            .col-index {
              width: 34px;
              text-align: center;
            }
            .col-name {
              width: 30%;
            }
            .col-school {
              width: 30%;
            }
            .col-sign {
              width: 28%; /* เพิ่มพื้นที่สำหรับลายมือชื่อ */
            }
            .col-note {
              width: 6%;  /* แคบลงสำหรับหมายเหตุ */
            }

            .center { text-align: center; }
            .sig-cell {
              height: 24px;
            }
            .note-cell {
              height: 24px;
            }
            .empty {
              text-align: center;
              color: #9CA3AF;
              font-style: italic;
              font-size: 10.5pt;
            }
          </style>
        </head>
        <body>
          ${htmlPages}
        </body>
      </html>
    `;

    // ------------------------ แปลง HTML เป็น PDF ------------------------
    const blob = Utilities.newBlob(html, "text/html", "attendance.html").getAs("application/pdf");
    const fileName = "attendance_sheet_" +
      Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmm") + ".pdf";
    blob.setName(fileName);

    return {
      success: true,
      fileName: fileName,
      mimeType: "application/pdf",
      data: Utilities.base64Encode(blob.getBytes())
    };

  } catch (error) {
    Logger.log("exportAttendanceSheet error: " + error);
    return {
      success: false,
      error: error.message || "ไม่สามารถสร้างใบลงเวลาได้"
    };
  }
}





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
          unique[key] = {
            id: "",
            name: sName,
            cluster: "",
            clusterId: "",
            registrationMode: "self",
            assignedActivities: []
          };
        }
      }
    }

    if (schoolsSheet && schoolsSheet.getLastRow() >= 2) {
      ensureSchoolsSheetStructure_(schoolsSheet);
      var lastRowSch = schoolsSheet.getLastRow();
      var schValues = schoolsSheet.getRange(2, 1, lastRowSch - 1, SCHOOL_SHEET_HEADERS.length).getValues();
      for (var j = 0; j < schValues.length; j++) {
        var id = String(schValues[j][0] || "").trim();
        var name = String(schValues[j][1] || "").trim();
        var clusterRaw = String(schValues[j][2] || "").trim();
        var registrationMode = normalizeRegistrationMode_(schValues[j][3] || "");
        var assignedActivities = parseAssignedActivitiesCell_(schValues[j][4]);
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
            clusterId: clusterId,
            registrationMode: registrationMode,
            assignedActivities: assignedActivities
          };
        } else {
          if (id && !unique[keyName].id) unique[keyName].id = id;
          if (clusterName && !unique[keyName].cluster) unique[keyName].cluster = clusterName;
          if (clusterId && !unique[keyName].clusterId) unique[keyName].clusterId = clusterId;
          unique[keyName].registrationMode = registrationMode || unique[keyName].registrationMode;
          unique[keyName].assignedActivities = assignedActivities;
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
    const scoreAssignments = getScoreAssignmentMap_(spreadsheet);
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
        email: record.email || "",
        scoreActivities: scoreAssignments.get((record.userid || "").toString().trim()) || []
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
        ? filtered.filter(school => doesSchoolMatchCluster_(school, targetCluster))
        : [];
    }
    return { success: true, schools: filtered };
  } catch (error) {
    Logger.log("getManagedSchools error: " + error);
    return { success: false, error: error.message };
  }
}

function updateSchoolRegistrationMode(request) {
  try {
    const data = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
    if (!sheet) throw new Error("ไม่พบชีต Schools");
    ensureSchoolsSheetStructure_(sheet);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor || !["admin", "area", "group_admin"].includes(actor.normalizedLevel)) {
      throw new Error("ไม่มีสิทธิ์แก้ไขโหมดลงทะเบียน");
    }
    const schoolId = String(data.schoolId || "").trim();
    if (!schoolId) throw new Error("กรุณาระบุรหัสโรงเรียน");
    const schoolRecord = lookupSchoolById_(schoolsIndex, schoolId);
    if (!schoolRecord) throw new Error("ไม่พบโรงเรียนในฐานข้อมูล");
    if (actor.normalizedLevel === "group_admin") {
      const competitionStage = getCompetitionStage_();
      if (competitionStage !== "cluster") {
        throw new Error("ไม่สามารถเปลี่ยนโหมดลงทะเบียนในรอบระดับเขต");
      }
      if (!isSchoolWithinActorCluster_(actor, schoolRecord)) {
        throw new Error("โรงเรียนอยู่นอกเครือข่ายของคุณ");
      }
    }
    const mode = normalizeRegistrationMode_(data.mode || data.registrationMode || "");
    const rowIndex = findSchoolRowIndex_(sheet, schoolId);
    if (rowIndex === -1) throw new Error("ไม่พบโรงเรียนในฐานข้อมูล");
    sheet.getRange(rowIndex, 4).setValue(formatRegistrationModeForSheet_(mode));
    if (mode === "self") {
      sheet.getRange(rowIndex, 5).clearContent();
    }
    return {
      success: true,
      schoolId: schoolId,
      registrationMode: mode
    };
  } catch (error) {
    Logger.log("updateSchoolRegistrationMode error: " + error);
    return { success: false, error: error.message };
  }
}

function updateSchoolAssignments(request) {
  try {
    const data = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_SCHOOLS);
    if (!sheet) throw new Error("ไม่พบชีต Schools");
    ensureSchoolsSheetStructure_(sheet);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor || !["admin", "area", "group_admin"].includes(actor.normalizedLevel)) {
      throw new Error("คุณไม่มีสิทธิ์กำหนดกิจกรรมให้โรงเรียน");
    }
    const schoolId = String(data.schoolId || "").trim();
    if (!schoolId) throw new Error("กรุณาระบุรหัสโรงเรียน");
    const schoolRecord = lookupSchoolById_(schoolsIndex, schoolId);
    if (!schoolRecord) throw new Error("ไม่พบโรงเรียนที่เลือก");
    if (actor.normalizedLevel === "group_admin" && !isSchoolWithinActorCluster_(actor, schoolRecord)) {
      throw new Error("คุณสามารถจัดการได้เฉพาะโรงเรียนในเครือข่ายของคุณ");
    }
    const currentMode = normalizeRegistrationMode_(schoolRecord.registrationMode || "");
    if (currentMode !== "group_assigned" && actor.normalizedLevel !== "admin" && actor.normalizedLevel !== "area") {
      throw new Error("โรงเรียนนี้ยังไม่ได้เปิดโหมดให้ Group Admin กำหนดกิจกรรม");
    }
    const activityIds = Array.isArray(data.activityIds) ? data.activityIds : [];
    const cleaned = [];
    const seen = new Set();
    activityIds.forEach(id => {
      const value = (id || "").toString().trim();
      if (!value) return;
      const key = normalizeKey(value);
      if (!key || seen.has(key)) return;
      seen.add(key);
      cleaned.push(value);
    });
    const rowIndex = findSchoolRowIndex_(sheet, schoolId);
    if (rowIndex === -1) throw new Error("ไม่พบโรงเรียนในฐานข้อมูล");
    sheet.getRange(rowIndex, 5).setValue(stringifyAssignedActivities_(cleaned));
    return {
      success: true,
      schoolId: schoolId,
      assignedActivities: cleaned,
      assignedCount: cleaned.length
    };
  } catch (error) {
    Logger.log("updateSchoolAssignments error: " + error);
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

function saveScoreAssignments(request) {
  try {
    const data = request || {};
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const schoolsIndex = buildSchoolsIndex_(spreadsheet);
    const actor = resolveActorContext_(data.actor, schoolsIndex);
    if (!actor) throw new Error("กรุณาเข้าสู่ระบบ");
    const allowed = ["admin", "area", "group_admin"];
    if (!allowed.includes(actor.normalizedLevel)) {
      throw new Error("คุณไม่มีสิทธิ์มอบหมายผู้ใช้คะแนน");
    }
    const targetUserId = (data.userId || data.userid || "").toString().trim();
    if (!targetUserId) throw new Error("ไม่ระบุผู้ใช้เป้าหมาย");
    const activityInput = Array.isArray(data.activityIds) ? data.activityIds : [];
    const usersSheet = getUsersSheet();
    const userRows = getAllUserRows_(usersSheet);
    let record = null;
    userRows.some(row => {
      if ((row[0] || "").toString() === targetUserId) {
        record = sanitizeUserRow(row);
        return true;
      }
      return false;
    });
    if (!record) throw new Error("ไม่พบผู้ใช้");
    if (!isUserWithinActorScope_(actor, record, schoolsIndex)) {
      throw new Error("คุณไม่มีสิทธิ์จัดการผู้ใช้นี้");
    }
    if (normalizeKey(record.level || "") !== "score") {
      throw new Error("รองรับเฉพาะผู้ใช้ระดับ Score เท่านั้น");
    }
    const activities = getActivities();
    if (activities.error) throw new Error(activities.error);
    const activityIndex = new Map();
    activities.forEach(item => {
      const idKey = normalizeKey(item.id);
      const nameKey = normalizeKey(item.name);
      if (idKey) activityIndex.set(idKey, item.id);
      if (nameKey) activityIndex.set(nameKey, item.id);
    });
    const normalizedActivities = Array.from(
      new Set(
        activityInput
          .map(id => normalizeKey(id))
          .filter(Boolean)
          .map(key => activityIndex.get(key))
          .filter(Boolean)
      )
    );
    upsertScoreAssignment_(spreadsheet, targetUserId, normalizedActivities, actor.username || actor.userId || "");
    return {
      success: true,
      activities: normalizedActivities
    };
  } catch (error) {
    Logger.log("saveScoreAssignments error: " + error);
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

    var rowIndex = findSchoolRowIndex_(sheet, schoolId);
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
