// -----------------------------------------------------------
// 10. PROJECT CHECK SYSTEM
// -----------------------------------------------------------

/**
 * ดึงรายชื่อนักศึกษาตามกลุ่มที่เลือก พร้อมคะแนนจาก Project Sheet
 * @param {number} groupIndex - index ของกลุ่มเรียน
 * @param {string} reviewer - ชื่อผู้ตรวจ (teacher, joo, poo, khem, wan, boss)
 * @param {string} teamNumber - เลขกลุ่มนักศึกษา เช่น "T01"
 */
function getProjectCheckStudents(groupIndex, reviewer, teamNumber) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const teamSheet = ss.getSheetByName("Team");
    const projectSheet = ss.getSheetByName("Project");

    if (!teamSheet || !projectSheet) {
      throw new Error("ไม่พบ Sheet 'Team' หรือ 'Project'");
    }

    const teamData = teamSheet.getDataRange().getDisplayValues();
    const projectRange = projectSheet.getDataRange();
    const projectData = projectRange.getValues(); // Use getValues for math
    const projectNotes = projectRange.getNotes(); // Fetch cell notes (for Qualitative feedback)

    // Slot Index ของผู้ตรวจสอดคล้องกับ saveProjectCheckScore
    const reviewerSlots = {
      อาจารย์: 0,
      teacher: 0,
      จู: 1,
      joo: 1,
      ภู: 2,
      poo: 2,
      เข้ม: 3,
      khem: 3,
      วัน: 4,
      wan: 4,
      บอส: 5,
      boss: 5,
    };

    const slotIndex = reviewerSlots[reviewer.toLowerCase()];
    if (slotIndex === undefined) {
      throw new Error(`ไม่พบ Slot สำหรับผู้ตรวจ: ${reviewer}`);
    }

    const students = [];
    let currentTeamNumber = "";
    let advisor = "";

    for (let i = 1; i < teamData.length; i++) {
      const rowTeamNumber = String(teamData[i][0] || "").trim();
      if (rowTeamNumber !== "") currentTeamNumber = rowTeamNumber;

      const advisorValue = String(teamData[i][11] || "").trim(); // Column L
      if (advisorValue !== "") advisor = advisorValue;

      if (!teamNumber || currentTeamNumber === String(teamNumber)) {
        const studentId = String(teamData[i][4] || "").trim(); // Column E
        const firstName = String(teamData[i][5] || "").trim(); // Column F
        const lastName = String(teamData[i][6] || "").trim(); // Column G

        if (studentId !== "") {
          let scores = [0, 0, 0, 0, 0, 0];
          let notes = ["", "", "", "", "", ""];
          let average = null;

          // ค้นหานักศึกษาใน Project Sheet
          for (let j = 2; j < projectData.length; j++) {
            if (String(projectData[j][1]).trim() === studentId) {
              // ดึงคะแนน 6 ข้อตาม Slot
              for (let k = 0; k < 6; k++) {
                // Column G เริ่มที่ index 6
                const colIdx = 6 + k * 6 + slotIndex;
                const val = projectData[j][colIdx];
                scores[k] = val !== "" && !isNaN(val) ? Number(val) : 0;
                notes[k] = projectNotes[j][colIdx] || "";
              }

              // Column E (index 4) คือ AVERAGE
              const avgVal = projectData[j][4];
              average = avgVal !== "" && !isNaN(avgVal) ? Number(avgVal) : null;
              break;
            }
          }

          students.push({
            teamNumber: currentTeamNumber,
            id: studentId,
            name: `${firstName} ${lastName}`.trim(),
            advisor: advisor,
            scores: scores,
            notes: notes,
            average: average,
            // score รวมเพื่อใช้แสดงสถานะ tab
            score: scores.reduce((a, b) => a + b, 0) || null,
          });
        }
      }
    }

    return {
      teamNumber: String(teamNumber),
      advisor: advisor || "ไม่ระบุ",
      reviewer: reviewer,
      students: students,
    };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลนักศึกษา: " + e.message);
  }
}

/**
 * บันทึกคะแนน Project Check ทั้งหมด (ทุกเกณฑ์ ทุกคนในกลุ่ม)
 */
function saveProjectCheckScore(groupIndex, reviewer, teamGrades) {
  try {
    const groups = getGroups();
    if (groupIndex < 0 || groupIndex >= groups.length) {
      throw new Error("ไม่พบกลุ่มเรียน");
    }

    const group = groups[groupIndex];
    const ss = SpreadsheetApp.openById(group.id);
    const projectSheet = ss.getSheetByName("Project");

    if (!projectSheet) {
      throw new Error("ไม่พบ Sheet ชื่อ 'Project'");
    }

    // อ่านข้อมูลทั้งหมดเพื่อหาแถวของนักศึกษา
    const data = projectSheet.getDataRange().getDisplayValues();

    // กำหนด Slot Index ของผู้ตรวจ (0-5) ให้ตรงตามลำดับ G, H, I, J, K, L
    const reviewerSlots = {
      อาจารย์: 0,
      teacher: 0,
      จู: 1,
      joo: 1,
      ภู: 2,
      poo: 2,
      เข้ม: 3,
      khem: 3,
      วัน: 4,
      wan: 4,
      บอส: 5,
      boss: 5,
    };

    const slotIndex = reviewerSlots[reviewer];
    if (slotIndex === undefined) {
      throw new Error(`ไม่พบ Slot สำหรับผู้ตรวจ: ${reviewer}`);
    }

    // เริ่มบันทึกทีละคน
    // teamGrades: { studentId: { scores: [6], notes: [6] } }
    let saveCount = 0;

    for (const studentId in teamGrades) {
      const draft = teamGrades[studentId];

      // ค้นหาแถวของนักศึกษา (เริ่มจากแถวที่ 3, index 2)
      let targetRow = -1;
      for (let row = 2; row < data.length; row++) {
        if (String(data[row][1]).trim() === String(studentId)) {
          targetRow = row + 1;
          break;
        }
      }

      if (targetRow === -1) continue;

      // บันทึกคะแนน 6 ข้อ
      for (let i = 0; i < 6; i++) {
        const criterionNum = i + 1;
        const score = draft.scores[i];
        const note = draft.notes[i] ? draft.notes[i].trim() : "";

        // ถ้าไม่มีคะแนน (>0) และไม่มีหมายเหตุ ไม่ต้องบันทึกอะไรเลย
        if (score === 0 && note === "") {
          continue;
        }

        // คำนวณคอลัมน์: G=7, G-L คือข้อ 1
        // สูตร: 7 + (ข้อ-1)*6 + slot
        const col = 7 + (criterionNum - 1) * 6 + slotIndex;

        const range = projectSheet.getRange(targetRow, col);
        range.setValue(score);

        if (note) {
          range.setNote(note);
        } else {
          range.clearNote();
        }
      }
      saveCount++;
    }

    return { success: true, count: saveCount };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการบันทึกคะแนน: " + e.message);
  }
}
