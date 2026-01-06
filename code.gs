// -----------------------------------------------------------
// 1. ROUTING & TEMPLATE ENGINE
// -----------------------------------------------------------
function doGet(e) {
  var tokenFromUrl = e.parameter.token;

  // CASE A: นักเรียน (เพิ่ม GPS parameter)
  if (tokenFromUrl) {
    var template = HtmlService.createTemplateFromFile("Student");
    template.token = tokenFromUrl;
    template.groupName = e.parameter.group || "";
    template.week = e.parameter.week || "";
    template.type = e.parameter.type || "";
    template.requireGPS = e.parameter.gps === "1"; // เพิ่ม GPS parameter

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
  groups.push({ name: name, id: id, isActive: true });
  PropertiesService.getScriptProperties().setProperty(
    "SAVED_GROUPS",
    JSON.stringify(groups)
  );
  return { success: true, data: groups };
}

function updateGroup(index, name, id, isActive) {
  try {
    SpreadsheetApp.openById(id);
  } catch (e) {
    return { success: false, msg: "Spreadsheet ID ไม่ถูกต้อง" };
  }
  let groups = getGroups();
  if (groups[index]) {
    groups[index].name = name;
    groups[index].id = id;
    // ป้องกันกรณีส่ง isActive เป็น undefined/null
    if (isActive !== undefined) {
      groups[index].isActive = isActive;
    } else if (groups[index].isActive === undefined) {
      groups[index].isActive = true;
    }
  }
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

      // ตรวจสอบค่าคะแนน
      const scoreValue = parseFloat(checkVal);

      if (scoreValue === 0.5) {
        // มาสาย (0.5 คะแนน) - นับเป็น Present แต่แสดงสถานะเป็น Late
        presentCount++;
        status = "Late";
        displayTime = checkTime ? checkTime : "Checked";
      } else if (scoreValue === 1 || checkVal == 1 || checkVal === "1") {
        // มาเรียน (1 คะแนน)
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

  // *** จุดสำคัญ: ต่อ String พารามิเตอร์เข้าไป (เพิ่ม gps parameter) ***
  const params = `token=${sessionToken}&group=${encodeURIComponent(
    data.groupName
  )}&week=${data.week}&type=${data.type}&gps=${data.requireGPS ? "1" : "0"}`;
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

  // *** จุดสำคัญ: สร้าง URL ให้ครบเหมือนตอน create (เพิ่ม gps parameter) ***
  const params = `token=${session.token}&group=${encodeURIComponent(
    session.groupName
  )}&week=${session.week}&type=${session.type}&gps=${
    session.requireGPS ? "1" : "0"
  }`;
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

function adminManualCheckIn(sheetId, week, type, studentId, status, score) {
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

    // สร้าง timeString พร้อมสถานะ
    const statusLabel = status === "Late" ? "มาสาย" : "มาเรียน";
    const timeString =
      Utilities.formatDate(new Date(), "GMT+7", "HH:mm:ss") + ` (Admin)`;

    // ใช้ค่า score ที่ส่งมา (0.5 สำหรับมาสาย, 1 สำหรับมาเรียน)
    cell.setValue(score);
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
 * ดึงรายการ Scrum Meeting จาก Sheet "Scrum"
 * อ่านจาก Row 2 (F2, H2, J2, ...) ที่เป็น Scrum Meeting #1, #2, #3, ...
 */
function getScrumMeetings(groupIndex) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const scrumSheet = ss.getSheetByName("Scrum");

    if (!scrumSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'Scrum'");
    }

    // อ่านแถวที่ 2 (header ของ Scrum Meetings)
    const headerRow = scrumSheet
      .getRange(2, 1, 1, scrumSheet.getLastColumn())
      .getDisplayValues()[0];
    const meetings = [];

    // เริ่มจาก column F (index 5) และข้ามทีละ 2 columns (F, H, J, L, ...)
    for (let col = 5; col < headerRow.length; col += 2) {
      const meetingName = String(headerRow[col]).trim();

      // ตรวจสอบว่าเป็น Scrum Meeting หรือไม่
      if (meetingName && meetingName.includes("Scrum Meeting")) {
        meetings.push({
          name: meetingName,
          columnIndex: col, // เก็บ index ของ column (0-based)
        });
      }
    }

    return meetings;
  } catch (e) {
    throw new Error(
      "เกิดข้อผิดพลาดในการดึงข้อมูล Scrum Meetings: " + e.message
    );
  }
}

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

    const data = teamSheet.getDataRange().getDisplayValues();
    const teams = [];
    const teamNumbers = new Set();

    // อ่านข้อมูลจากแถวที่ 2 เป็นต้นไป (แถว 1 เป็น header)
    for (let i = 1; i < data.length; i++) {
      const teamNumber = String(data[i][0] || "").trim(); // Column A

      if (teamNumber && !teamNumbers.has(teamNumber)) {
        teamNumbers.add(teamNumber);
        teams.push({
          teamNumber: teamNumber,
        });
      }
    }

    return teams;
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลกลุ่ม: " + e.message);
  }
}

/**
 * ดึงรายชื่อนักศึกษาตามกลุ่มที่เลือก พร้อมคะแนนจาก Scrum Meeting ที่เลือก
 * @param {number} groupIndex - index ของกลุ่มเรียน
 * @param {number} meetingColumnIndex - column index ของ Scrum Meeting (0-based)
 * @param {string} teamNumber - เลขกลุ่มนักศึกษา เช่น "T01"
 */
function getScrumStudents(groupIndex, meetingColumnIndex, teamNumber) {
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

    // อ่านข้อมูลจาก Team Sheet
    const teamData = teamSheet.getDataRange().getDisplayValues();

    // อ่านข้อมูลจาก Scrum Sheet (ถ้ามี)
    const scrumData = scrumSheet
      ? scrumSheet.getDataRange().getDisplayValues()
      : [];

    const students = [];
    let currentTeamNumber = "";
    let advisor = "";

    // อ่านข้อมูลนักศึกษาจาก Team Sheet (เริ่มจากแถวที่ 2, index 1)
    for (let i = 1; i < teamData.length; i++) {
      // จัดการ merged cells - อ่านค่า Team Number
      const rowTeamNumber = String(teamData[i][0] || "").trim(); // Column A
      if (rowTeamNumber !== "") {
        currentTeamNumber = rowTeamNumber;
      }

      // อ่านค่า Advisor
      const advisorValue = String(teamData[i][11] || "").trim(); // Column L
      if (advisorValue !== "") {
        advisor = advisorValue;
      }

      // ถ้ามีการระบุ teamNumber ให้กรอง, ถ้าไม่มี (null/undefined/empty) ให้เอาหมด
      if (!teamNumber || currentTeamNumber === String(teamNumber)) {
        const studentId = String(teamData[i][4] || "").trim(); // Column E
        const firstName = String(teamData[i][5] || "").trim(); // Column F
        const lastName = String(teamData[i][6] || "").trim(); // Column G

        // ถ้ามีรหัสนักศึกษา
        if (studentId !== "") {
          // ค้นหาคะแนนจาก Scrum Sheet
          let score1 = null;
          let score2 = null;

          if (scrumData.length > 0) {
            // ค้นหาแถวของนักศึกษาใน Scrum Sheet
            // เริ่มจากแถวที่ 5 (index 4) เพราะ Row 1-4 เป็น headers
            for (let j = 4; j < scrumData.length; j++) {
              const scrumStudentId = String(scrumData[j][1] || "").trim(); // Column B

              if (scrumStudentId === studentId) {
                // อ่านคะแนนจาก column ที่เลือก
                // meetingColumnIndex คือ column F (index 5) สำหรับ Meeting #1
                const score1Value = String(
                  scrumData[j][meetingColumnIndex] || ""
                ).trim();
                const score2Value = String(
                  scrumData[j][meetingColumnIndex + 1] || ""
                ).trim();

                score1 =
                  score1Value !== "" && !isNaN(score1Value)
                    ? Number(score1Value)
                    : null;
                score2 =
                  score2Value !== "" && !isNaN(score2Value)
                    ? Number(score2Value)
                    : null;
                break;
              }
            }
          }

          students.push({
            teamNumber: currentTeamNumber,
            id: studentId,
            name: `${firstName} ${lastName}`.trim(),
            advisor: advisor,
            score1: score1,
            score2: score2,
          });
        }
      }
    }

    return {
      teamNumber: String(teamNumber),
      advisor: advisor || "ไม่ระบุ",
      students: students,
    };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลนักศึกษา: " + e.message);
  }
}

/**
 * บันทึกคะแนน Scrum
 */
function saveScrumScore(
  groupIndex,
  meetingColumnIndex,
  studentId,
  score1,
  score2
) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const scrumSheet = ss.getSheetByName("Scrum");

    if (!scrumSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'Scrum'");
    }

    // อ่านข้อมูลทั้งหมด
    const data = scrumSheet.getDataRange().getDisplayValues();

    // ค้นหาแถวของนักศึกษา (เริ่มจากแถวที่ 5, index 4)
    let targetRow = -1;
    for (let row = 4; row < data.length; row++) {
      const currentStudentId = String(data[row][1] || "").trim(); // Column B
      if (currentStudentId === String(studentId)) {
        targetRow = row + 1; // +1 เพราะ Sheet เริ่มที่ 1
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error("ไม่พบนักศึกษารหัส: " + studentId);
    }

    // บันทึกคะแนนลง column ที่ถูกต้อง
    // meetingColumnIndex เป็น 0-based, ต้อง +1 สำหรับ getRange
    const scoreCol1 = meetingColumnIndex + 1; // Column F, H, J, ...
    const scoreCol2 = meetingColumnIndex + 2; // Column G, I, K, ...

    scrumSheet.getRange(targetRow, scoreCol1).setValue(score1);
    scrumSheet.getRange(targetRow, scoreCol2).setValue(score2);

    return { success: true };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการบันทึกคะแนน: " + e.message);
  }
}

// -----------------------------------------------------------
// 8. PROJECT REVIEW SYSTEM
// -----------------------------------------------------------

/**
 * ดึงข้อมูลนักศึกษาจาก Sheet "Project Review"
 * อ่านจาก B5 (รหัสนักศึกษา) และ C5 (ชื่อ)
 * พร้อมตรวจสอบสถานะว่าตรวจแล้วหรือยัง (จาก E5-N5)
 */
function getProjectReviewData(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Project Review");

    if (!sheet) {
      return { success: false, msg: "ไม่พบ Sheet 'Project Review'" };
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 5) {
      return { success: true, students: [] };
    }

    // ดึงข้อมูลนักศึกษา (B5:C และ E:N สำหรับ checkbox data)
    const studentData = sheet.getRange(5, 2, lastRow - 4, 2).getValues(); // B5:C (column 2,3)
    const checkboxData = sheet.getRange(5, 5, lastRow - 4, 10).getValues(); // E5:N (10 columns)

    const students = [];

    for (let i = 0; i < studentData.length; i++) {
      const id = String(studentData[i][0]).trim();
      const name = String(studentData[i][1]).trim();

      if (id === "" || name === "") continue;

      // ตรวจสอบว่ามีข้อมูล checkbox หรือไม่ (ถ้ามีอย่างน้อย 1 ช่อง = ตรวจแล้ว)
      const checkboxes = checkboxData[i];
      const hasAnyCheckbox = checkboxes.some(
        (val) => val === true || val === "TRUE" || val === 1
      );

      students.push({
        id: id,
        name: name,
        reviewed: hasAnyCheckbox,
        checkboxes: checkboxes.map(
          (val) => val === true || val === "TRUE" || val === 1
        ),
      });
    }

    return { success: true, students: students };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

/**
 * บันทึกข้อมูล Project Review
 * @param {string} sheetId - ID ของ Spreadsheet
 * @param {string} studentId - รหัสนักศึกษา
 * @param {Array<boolean>} checkboxValues - Array ของค่า checkbox (10 ตัว)
 */
function saveProjectReview(sheetId, studentId, checkboxValues) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Project Review");

    if (!sheet) {
      return { success: false, msg: "ไม่พบ Sheet 'Project Review'" };
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 5) {
      return { success: false, msg: "ไม่พบข้อมูลนักศึกษา" };
    }

    // ค้นหาแถวของนักศึกษา (จากคอลัมน์ B)
    const ids = sheet
      .getRange(5, 2, lastRow - 4, 1)
      .getValues()
      .flat()
      .map(String);

    const idx = ids.indexOf(String(studentId));

    if (idx === -1) {
      return { success: false, msg: "ไม่พบรหัสนักศึกษา: " + studentId };
    }

    const targetRow = 5 + idx;

    // บันทึกค่า checkbox ลง E:N (10 columns)
    // ถ้า checkbox เป็น true ให้ส่ง 1, ถ้าเป็น false ไม่ต้องส่งอะไร (ค่าว่าง)
    for (let i = 0; i < checkboxValues.length && i < 10; i++) {
      const col = 5 + i; // Column E = 5, F = 6, ..., N = 14
      const value = checkboxValues[i] ? 1 : "";
      sheet.getRange(targetRow, col).setValue(value);
    }

    return { success: true };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// -----------------------------------------------------------
// 9. FINAL EXAM SYSTEM
// -----------------------------------------------------------

/**
 * ดึงข้อมูลนักศึกษาจาก Sheet "Final Exam"
 * อ่านจาก B4 (รหัสนักศึกษา) และ C4 (ชื่อ)
 * พร้อมคะแนนจาก F4-P4 (11 คอลัมน์)
 */
function getFinalExamData(sheetId) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Final Exam");

    if (!sheet) {
      return { success: false, msg: "ไม่พบ Sheet 'Final Exam'" };
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 4) {
      return { success: true, students: [] };
    }

    // ดึงข้อมูลนักศึกษา (B4:C และ F:P สำหรับ score data)
    const studentData = sheet.getRange(4, 2, lastRow - 3, 2).getValues(); // B4:C (column 2,3)
    const scoreData = sheet.getRange(4, 6, lastRow - 3, 11).getValues(); // F4:P (11 columns)

    const students = [];

    for (let i = 0; i < studentData.length; i++) {
      const id = String(studentData[i][0]).trim();
      const name = String(studentData[i][1]).trim();

      if (id === "" || name === "") continue;

      // ตรวจสอบว่ามีข้อมูลคะแนนหรือไม่ (ถ้ามีอย่างน้อย 1 ช่อง = ตรวจแล้ว)
      const scores = scoreData[i];
      const hasAnyScore = scores.some(
        (val) => val !== "" && val !== null && !isNaN(parseFloat(val))
      );

      students.push({
        id: id,
        name: name,
        reviewed: hasAnyScore,
        scores: scores.map((val) =>
          val === "" || val === null ? 0 : parseFloat(val)
        ),
      });
    }

    return { success: true, students: students };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

/**
 * บันทึกข้อมูล Final Exam
 * @param {string} sheetId - ID ของ Spreadsheet
 * @param {string} studentId - รหัสนักศึกษา
 * @param {Array<number>} scoreValues - Array ของคะแนน (11 ตัว)
 */
function saveFinalExam(sheetId, studentId, scoreValues) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName("Final Exam");

    if (!sheet) {
      return { success: false, msg: "ไม่พบ Sheet 'Final Exam'" };
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 4) {
      return { success: false, msg: "ไม่พบข้อมูลนักศึกษา" };
    }

    // ค้นหาแถวของนักศึกษา (จากคอลัมน์ B)
    const ids = sheet
      .getRange(4, 2, lastRow - 3, 1)
      .getValues()
      .flat()
      .map(String);

    const idx = ids.indexOf(String(studentId));

    if (idx === -1) {
      return { success: false, msg: "ไม่พบรหัสนักศึกษา: " + studentId };
    }

    const targetRow = 4 + idx;

    // บันทึกคะแนนลง F:P (11 columns)
    // F4-P4 ตามลำดับของนักศึกษา
    for (let i = 0; i < scoreValues.length && i < 11; i++) {
      const col = 6 + i; // Column F = 6, G = 7, ..., P = 16
      const value = scoreValues[i] || 0;
      sheet.getRange(targetRow, col).setValue(value);
    }

    return { success: true };
  } catch (e) {
    return { success: false, msg: "Error: " + e.message };
  }
}
