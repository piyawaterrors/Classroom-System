// -----------------------------------------------------------
// 1. ROUTING & TEMPLATE ENGINE
// -----------------------------------------------------------
function doGet(e) {
  var tokenFromUrl = e.parameter.token;

  // CASE A: นักเรียน (เหมือนเดิม)
  if (tokenFromUrl) {
    var template = HtmlService.createTemplateFromFile("Student");
    template.token = tokenFromUrl;
    template.groupName = e.parameter.group || "";
    template.week = e.parameter.week || "";
    template.type = e.parameter.type || "";

    return template
      .evaluate()
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setTitle("Student Check-in")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // CASE B: อาจารย์ (แก้ไขตรงนี้)
  var page = e.parameter.page || "checkin";

  var template = HtmlService.createTemplateFromFile("Main");
  template.activePage = page;

  // *** เพิ่มบรรทัดนี้: ส่ง URL เต็มของ Web App ไปให้หน้าบ้าน ***
  template.url = ScriptApp.getService().getUrl();

  return template
    .evaluate()
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setTitle("Classroom Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันสำหรับดึงไฟล์ HTML ย่อย (Partial View)
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    // ถ้าหาไฟล์ไม่เจอ ให้แสดง Error แทนหน้าขาว
    return (
      "<div style='color:red; padding:20px;'>Error: ไม่พบไฟล์ <b>" +
      filename +
      ".html</b> <br>กรุณาสร้างไฟล์นี้ใน Apps Script</div>"
    );
  }
}

// -----------------------------------------------------------
// 2. GROUP MANAGEMENT (ใช้ร่วมกัน)
// -----------------------------------------------------------
function getGroups() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("SAVED_GROUPS");
  return raw ? JSON.parse(raw) : [];
}

function saveGroup(name, id) {
  try {
    SpreadsheetApp.openById(id);
  } catch (e) {
    return { success: false, msg: "Spreadsheet ID ไม่ถูกต้อง" };
  }
  let groups = getGroups();
  groups.push({ name: name, id: id });
  PropertiesService.getScriptProperties().setProperty(
    "SAVED_GROUPS",
    JSON.stringify(groups)
  );
  return { success: true, data: groups };
}

function updateGroup(index, name, id) {
  try {
    SpreadsheetApp.openById(id);
  } catch (e) {
    return { success: false, msg: "Spreadsheet ID ไม่ถูกต้อง" };
  }
  let groups = getGroups();
  groups[index] = { name: name, id: id };
  PropertiesService.getScriptProperties().setProperty(
    "SAVED_GROUPS",
    JSON.stringify(groups)
  );
  return { success: true, data: groups };
}

function deleteGroup(index) {
  let groups = getGroups();
  groups.splice(index, 1);
  PropertiesService.getScriptProperties().setProperty(
    "SAVED_GROUPS",
    JSON.stringify(groups)
  );
  return { success: true, data: groups };
}

// -----------------------------------------------------------
// 3. CHECK-IN SYSTEM (Logic เดิม)
// -----------------------------------------------------------
function getDashboardData(sheetId, week, type) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const attSheet = ss.getSheetByName("Attendance");
    if (!attSheet) return { success: false, msg: "ไม่พบ Tab 'Attendance'" };
    const lastRow = attSheet.getLastRow();
    if (lastRow < 5)
      return {
        success: true,
        total: 0,
        present: 0,
        absent: 0,
        studentList: [],
      };
    let colIndex = 6 + (parseInt(week) - 1) * 2 + (type === "Lab" ? 1 : 0);
    const studentsData = attSheet.getRange(5, 2, lastRow - 4, 2).getValues();
    const statusValues = attSheet
      .getRange(5, colIndex, lastRow - 4, 1)
      .getValues();
    const statusNotes = attSheet
      .getRange(5, colIndex, lastRow - 4, 1)
      .getNotes();
    let presentCount = 0;
    let validStudentCount = 0;
    let studentList = [];
    for (let i = 0; i < studentsData.length; i++) {
      const id = studentsData[i][0];
      const name = studentsData[i][1];
      const checkVal = statusValues[i][0];
      const checkTime = statusNotes[i][0];
      if (id === "" || name === "") continue;
      validStudentCount++;
      let status = "Absent";
      let displayTime = "-";
      if (checkVal == 1 || checkVal === "1") {
        presentCount++;
        status = "Present";
        displayTime = checkTime ? checkTime : "Checked";
      }
      studentList.push({
        id: id,
        name: name,
        status: status,
        time: displayTime,
      });
    }
    return {
      success: true,
      total: validStudentCount,
      present: presentCount,
      absent: validStudentCount - presentCount,
      studentList: studentList,
    };
  } catch (e) {
    return { success: false, msg: e.toString() };
  }
}

function createSession(data) {
  const props = PropertiesService.getScriptProperties();
  const expireTime = new Date().getTime() + data.timeLimit * 60 * 1000;
  const sessionToken =
    Math.random().toString(36).substring(2, 15) +
    Math.random().toString(36).substring(2, 15);

  const sessionData = {
    active: true,
    token: sessionToken,
    lat: data.lat,
    lng: data.lng,
    expireTime: expireTime,
    targetSheetId: data.sheetId,
    groupName: data.groupName,
    week: data.week,
    type: data.type,
    radius: data.radius || 100,
    requireGPS: data.requireGPS,
  };

  props.setProperty("CURRENT_SESSION", JSON.stringify(sessionData));

  const baseUrl = ScriptApp.getService().getUrl();
  // เช็คว่า URL มี ? อยู่แล้วหรือไม่
  const separator = baseUrl.includes("?") ? "&" : "?";

  // *** จุดสำคัญ: ต่อ String พารามิเตอร์เข้าไป ***
  const params = `token=${sessionToken}&group=${encodeURIComponent(
    data.groupName
  )}&week=${data.week}&type=${data.type}`;
  const sessionUrl = baseUrl + separator + params;

  return {
    success: true,
    url: sessionUrl, // ส่ง URL ตัวเต็มที่มีพารามิเตอร์ครบกลับไป
    expireTime: expireTime,
    groupName: data.groupName,
    week: data.week,
    type: data.type,
  };
}

function getSessionStatus() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("CURRENT_SESSION");
  if (!json) return { active: false };

  const session = JSON.parse(json);
  const now = new Date().getTime();

  if (now > session.expireTime) {
    props.deleteProperty("CURRENT_SESSION");
    return { active: false };
  }

  const baseUrl = ScriptApp.getService().getUrl();
  const separator = baseUrl.includes("?") ? "&" : "?";

  // *** จุดสำคัญ: สร้าง URL ให้ครบเหมือนตอน create ***
  const params = `token=${session.token}&group=${encodeURIComponent(
    session.groupName
  )}&week=${session.week}&type=${session.type}`;
  const sessionUrl = baseUrl + separator + params;

  return {
    active: true,
    expireTime: session.expireTime,
    url: sessionUrl, // ส่ง URL ตัวเต็มกลับไป
    groupName: session.groupName,
    week: session.week,
    type: session.type,
    sheetId: session.targetSheetId,
  };
}

function stopCurrentSession() {
  PropertiesService.getScriptProperties().deleteProperty("CURRENT_SESSION");
  return { success: true };
}

function checkInStudent(studentId, userLat, userLng, clientToken) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, msg: "ระบบยุ่งอยู่" };
  try {
    const props = PropertiesService.getScriptProperties();
    const sessionJson = props.getProperty("CURRENT_SESSION");
    if (!sessionJson) return { success: false, msg: "⛔ ไม่มีการเปิดเช็คชื่อ" };
    const session = JSON.parse(sessionJson);
    if (new Date().getTime() > session.expireTime)
      return { success: false, msg: "⌛ หมดเวลา" };
    if (!clientToken || clientToken !== session.token)
      return { success: false, msg: "🚫 ลิงก์ไม่ถูกต้อง" };
    if (session.requireGPS !== false) {
      if (!userLat || !userLng)
        return { success: false, msg: "❌ ไม่พบพิกัด GPS" };
      const dist =
        calculateDistance(session.lat, session.lng, userLat, userLng) * 1000;
      const maxDist = session.radius || 100;
      if (dist > maxDist)
        return {
          success: false,
          msg: `📍 ไกลเกินไป (${dist.toFixed(0)}m / ${maxDist}m)`,
        };
    }
    const ss = SpreadsheetApp.openById(session.targetSheetId);
    const attSheet = ss.getSheetByName("Attendance");
    const lastRow = attSheet.getLastRow();
    if (lastRow < 5) return { success: false, msg: "ไม่พบรายชื่อ" };
    const ids = attSheet
      .getRange(5, 2, lastRow - 4, 1)
      .getValues()
      .flat()
      .map(String);
    const idx = ids.indexOf(String(studentId));
    if (idx === -1) return { success: false, msg: "❌ ไม่พบรหัสนักศึกษา" };
    const targetRow = 5 + idx;
    const targetCol =
      6 + (parseInt(session.week) - 1) * 2 + (session.type === "Lab" ? 1 : 0);
    const cell = attSheet.getRange(targetRow, targetCol);
    if (cell.getValue() == 1)
      return { success: true, msg: "✅ เช็คชื่อแล้ว", already: true };
    const timeString = Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss");
    cell.setValue(1);
    cell.setNote(timeString);
    const studentName = attSheet.getRange(targetRow, 3).getValue();
    return { success: true, msg: "OK", name: studentName };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function adminManualCheckIn(sheetId, week, type, studentId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const attSheet = ss.getSheetByName("Attendance");
    const lastRow = attSheet.getLastRow();
    const ids = attSheet
      .getRange(5, 2, lastRow - 4, 1)
      .getValues()
      .flat()
      .map(String);
    const idx = ids.indexOf(String(studentId));
    if (idx === -1) return { success: false, msg: "ไม่พบรหัสนี้" };
    const targetRow = 5 + idx;
    const colIndex = 6 + (parseInt(week) - 1) * 2 + (type === "Lab" ? 1 : 0);
    const cell = attSheet.getRange(targetRow, colIndex);
    const timeString =
      Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss") + " (Admin)";
    cell.setValue(1);
    cell.setNote(timeString);
    const name = attSheet.getRange(targetRow, 3).getValue();
    return { success: true, name: name };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371;
  const p = Math.PI / 180;
  const a =
    0.5 -
    Math.cos((lat2 - lat1) * p) / 2 +
    (Math.cos(lat1 * p) *
      Math.cos(lat2 * p) *
      (1 - Math.cos((lon2 - lon1) * p))) /
      2;
  return 12742 * Math.asin(Math.sqrt(a));
}

// -----------------------------------------------------------
// 5. LAB GRADING SYSTEM (ปรับปรุงใหม่ตาม Flow)
// -----------------------------------------------------------

// 1. ดึงข้อมูลตั้งต้น (รายชื่อแลบ + รายชื่อนักเรียน)
function getLabInitData(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Lab Assignments");
    if (!sheet) return { success: false, msg: "ไม่พบ Sheet 'Lab Assignments'" };

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    // --- A. ดึงรายชื่อ Lab (เริ่ม F2) ---
    // สมมติว่า 1 แลบ ใช้ 2 ช่อง (F,G / H,I / ...) ตามที่มี Score 1, Score 2
    let labs = [];
    if (lastCol >= 6) {
      // Col 6 = F
      // ดึงหัวข้อแถวที่ 2
      const headers = sheet.getRange(2, 6, 1, lastCol - 5).getValues()[0];

      // วนลูปทีละ 2 คอลัมน์ (เพราะ 1 Lab มี 2 ช่องคะแนน)
      for (let i = 0; i < headers.length; i += 2) {
        let labName = headers[i];
        if (labName && labName !== "") {
          labs.push({
            name: labName,
            colIndex: 6 + i, // คอลัมน์เริ่มต้น (1-based index) ของแลบนั้น
          });
        }
      }
    }

    // --- B. ดึงรายชื่อนักเรียน (เริ่มแถว 4, Col B=ID, Col C=Name) ---
    let students = [];
    if (lastRow >= 4) {
      const data = sheet.getRange(4, 2, lastRow - 3, 2).getValues(); // Col B, C
      students = data
        .filter((r) => r[0] != "" && r[1] != "") // กรองแถวว่าง
        .map((r) => ({ id: String(r[0]), name: r[1] }));
    }

    return { success: true, labs: labs, data: students };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// 2. บันทึกคะแนน (Update Cell โดยตรง)
function saveLabAssignmentScore(
  sheetId,
  studentId,
  startColIndex,
  score1,
  score2
) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Lab Assignments");
    const lastRow = sheet.getLastRow();

    if (lastRow < 4) return { success: false, msg: "ไม่พบข้อมูลนักศึกษา" };

    // 1. หาแถวของนักศึกษา (ค้นหาใน Col B)
    const ids = sheet
      .getRange(4, 2, lastRow - 3, 1)
      .getValues()
      .flat()
      .map(String);
    const studentIndex = ids.indexOf(String(studentId));

    if (studentIndex === -1)
      return { success: false, msg: "ไม่พบรหัสนักศึกษาใน Sheet นี้" };

    const targetRow = 4 + studentIndex; // แถวจริงใน Sheet

    // 2. บันทึกคะแนน (เฉพาะที่มีค่า)
    // score1 ลง colIndex, score2 ลง colIndex + 1

    if (score1 !== null && score1 !== 0) {
      sheet.getRange(targetRow, startColIndex).setValue(score1);
    }

    if (score2 !== null && score2 !== 0) {
      sheet.getRange(targetRow, startColIndex + 1).setValue(score2);
    }

    return { success: true };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

// -----------------------------------------------------------
// 6. LAB DASHBOARD DATA (Updated: Return List)
// -----------------------------------------------------------
function getLabStats(sheetId, colIndex) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Lab Assignments");
    if (!sheet) return { success: false, msg: "ไม่พบ Sheet" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return { success: true, studentList: [] }; // ไม่มีนักเรียน

    // 1. ดึงข้อมูลนักเรียน (ID, Name) เริ่มแถว 4
    const students = sheet.getRange(4, 2, lastRow - 3, 2).getValues(); // Col B, C

    // 2. ดึงคะแนน (Score 1, Score 2) ตามคอลัมน์ที่เลือก
    const scores = sheet
      .getRange(4, parseInt(colIndex), lastRow - 3, 2)
      .getValues();

    let studentList = [];

    for (let i = 0; i < students.length; i++) {
      const id = students[i][0];
      const name = students[i][1];
      const s1 = scores[i][0];
      const s2 = scores[i][1];

      if (id === "" || name === "") continue; // ข้ามแถวว่าง

      let status = "Missing";
      let displayScore = "-";

      // ถ้ามีคะแนนอย่างน้อย 1 ช่อง ถือว่าส่งแล้ว (หรือจะปรับ logic ตามต้องการ)
      // เงื่อนไข: ต้องมีค่าทั้ง 2 ช่องถึงจะสมบูรณ์ หรือแค่ช่องเดียวก็ได้?
      // ในที่นี้เอาแบบ: ถ้ามีค่าสักช่อง ถือว่า Submitted
      if ((s1 !== "" && s1 !== null) || (s2 !== "" && s2 !== null)) {
        status = "Submitted";
        displayScore = `${s1 === "" ? 0 : s1} / ${s2 === "" ? 0 : s2}`;
      }

      studentList.push({
        id: String(id),
        name: name,
        status: status,
        score: displayScore,
      });
    }

    return {
      success: true,
      studentList: studentList,
    };
  } catch (e) {
    return { success: false, msg: e.message };
  }
}

// -----------------------------------------------------------
// 7. SCRUM MANAGEMENT
// -----------------------------------------------------------

/**
 * ดึงรายการกลุ่มนักศึกษา (Team) จาก Sheet "Team"
 */
function getScrumTeams(groupIndex) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const teamSheet = ss.getSheetByName("Team");

    if (!teamSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'Team'");
    }

    const data = teamSheet.getDataRange().getValues();
    const teams = [];
    const teamNumbers = new Set();

    // อ่านข้อมูลจากแถวที่ 2 เป็นต้นไป (แถว 1 เป็น header)
    for (let i = 1; i < data.length; i++) {
      const teamNumber = data[i][0]; // Column A (index 0)
      const advisor = data[i][11]; // Column L (index 11)

      if (teamNumber && !teamNumbers.has(teamNumber)) {
        teamNumbers.add(teamNumber);
        teams.push({
          teamNumber: String(teamNumber),
          advisor: String(advisor || "ไม่ระบุ"),
        });
      }
    }

    return teams;
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลกลุ่ม: " + e.message);
  }
}

/**
 * ดึงรายชื่อนักศึกษาในกลุ่ม (Team) พร้อมคะแนน
 */
function getScrumTeamStudents(groupIndex, teamNumber) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const teamSheet = ss.getSheetByName("Team");
    const scrumSheet = ss.getSheetByName("Scrum");

    if (!teamSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'Team'");
    }

    // ถ้ายังไม่มี Sheet Scrum ให้สร้างใหม่
    if (!scrumSheet) {
      const newSheet = ss.insertSheet("Scrum");
      newSheet
        .getRange(1, 1, 1, 4)
        .setValues([["รหัสนักศึกษา", "ชื่อ-นามสกุล", "คะแนน 1", "คะแนน 2"]]);
      newSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }

    const teamData = teamSheet.getDataRange().getValues();
    const scrumData = scrumSheet ? scrumSheet.getDataRange().getValues() : [];

    const students = [];
    let advisor = "";

    // อ่านข้อมูลนักศึกษาจาก Team Sheet
    for (let i = 1; i < teamData.length; i++) {
      const rowTeamNumber = String(teamData[i][0]);

      if (rowTeamNumber === String(teamNumber)) {
        const studentId = String(teamData[i][4] || ""); // Column E (index 4)
        const firstName = String(teamData[i][5] || ""); // Column F (index 5)
        const lastName = String(teamData[i][6] || ""); // Column G (index 6)
        advisor = String(teamData[i][11] || "ไม่ระบุ"); // Column L (index 11)

        if (studentId) {
          // ค้นหาคะแนนจาก Scrum Sheet
          let score1 = null;
          let score2 = null;

          for (let j = 1; j < scrumData.length; j++) {
            if (String(scrumData[j][0]) === studentId) {
              score1 = scrumData[j][2] !== "" ? Number(scrumData[j][2]) : null;
              score2 = scrumData[j][3] !== "" ? Number(scrumData[j][3]) : null;
              break;
            }
          }

          students.push({
            id: studentId,
            name: `${firstName} ${lastName}`.trim(),
            score1: score1,
            score2: score2,
          });
        }
      }
    }

    return {
      teamNumber: String(teamNumber),
      advisor: advisor,
      students: students,
    };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลนักศึกษา: " + e.message);
  }
}

/**
 * บันทึกคะแนน Scrum
 */
function saveScrumScore(groupIndex, teamNumber, studentId, score1, score2) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    let scrumSheet = ss.getSheetByName("Scrum");

    // ถ้ายังไม่มี Sheet Scrum ให้สร้างใหม่
    if (!scrumSheet) {
      scrumSheet = ss.insertSheet("Scrum");
      scrumSheet
        .getRange(1, 1, 1, 4)
        .setValues([["รหัสนักศึกษา", "ชื่อ-นามสกุล", "คะแนน 1", "คะแนน 2"]]);
      scrumSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    }

    // ดึงข้อมูลนักศึกษาจาก Team Sheet
    const teamSheet = ss.getSheetByName("Team");
    const teamData = teamSheet.getDataRange().getValues();

    let studentName = "";
    for (let i = 1; i < teamData.length; i++) {
      if (String(teamData[i][4]) === String(studentId)) {
        const firstName = String(teamData[i][5] || "");
        const lastName = String(teamData[i][6] || "");
        studentName = `${firstName} ${lastName}`.trim();
        break;
      }
    }

    // ค้นหาแถวที่มีรหัสนักศึกษานี้อยู่แล้ว
    const scrumData = scrumSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < scrumData.length; i++) {
      if (String(scrumData[i][0]) === String(studentId)) {
        rowIndex = i + 1; // +1 เพราะ getRange เริ่มที่ 1
        break;
      }
    }

    // ถ้าเจอแล้ว ให้อัปเดต ถ้าไม่เจอให้เพิ่มแถวใหม่
    if (rowIndex > 0) {
      scrumSheet.getRange(rowIndex, 3).setValue(score1);
      scrumSheet.getRange(rowIndex, 4).setValue(score2);
    } else {
      scrumSheet.appendRow([studentId, studentName, score1, score2]);
    }

    return { success: true };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการบันทึกคะแนน: " + e.message);
  }
}
