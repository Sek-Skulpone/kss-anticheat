// xlsx-parser.js - Excel Parser for Student Roster

/**
 * Parses the uploaded student Excel file and extracts student list grouped by class/room.
 * @param {File} file - The Excel file uploaded by the user.
 * @returns {Promise<Array>} A promise that resolves to an array of student objects.
 */
export function parseStudentExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert sheet to 2D array of cells (header: 1 keeps empty rows and structures intact)
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        
        const students = processRawExcelRows(rows);
        if (students.length === 0) {
          reject(new Error("ไม่พบข้อมูลนักเรียน หรือรูปแบบตารางไม่ตรงตามที่กำหนด"));
        } else {
          resolve(students);
        }
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Process raw 2D array from Excel and apply grouping logic
 * @param {Array<Array>} rows - Raw sheet cells
 * @returns {Array} List of parsed student objects
 */
function processRawExcelRows(rows) {
  let grade = "ม.1"; // Fallback grade
  
  // Try to find the Grade (ระดับชั้น) in the first 3 rows
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const rowStr = rows[i].join(" ");
    const gradeMatch = rowStr.match(/ชั้น\s*(ม\.\s*\d+)/);
    if (gradeMatch) {
      grade = gradeMatch[1].replace(/\s+/g, '').trim(); // "ม.1"
      break;
    }
  }
  
  let noColIndex = -1;
  let idColIndex = -1;
  let nameColIndex = -1;
  
  const allStudents = [];
  let tempStudents = [];
  
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    
    // Convert all cells in row to string for searching
    const rowCells = row.map(cell => String(cell).trim());
    
    // Look for column headers if not found yet
    if (noColIndex === -1 || idColIndex === -1 || nameColIndex === -1) {
      for (let j = 0; j < rowCells.length; j++) {
        const text = rowCells[j];
        if (text === "เลขที่") noColIndex = j;
        if (text === "เลขประจำตัว" || text === "รหัสประจำตัว") idColIndex = j;
        if (text === "ชื่อ นามสกุล" || text === "ชื่อ-นามสกุล" || text === "ชื่อนามสกุล") nameColIndex = j;
      }
      continue;
    }
    
    // Detect summary row (indicating end of a room block)
    // Example: "ห้องที่  1  รวม  33  คน ( ช. 16, ญ. 17)"
    const rowText = rowCells.join(" ");
    const summaryMatch = rowText.match(/ห้องที่\s*(\d+)\s*รวม/);
    if (summaryMatch) {
      const roomNum = summaryMatch[1].trim();
      
      // Update room and grade for all collected students in the current room block
      tempStudents.forEach(student => {
        student.room = roomNum;
        student.grade = grade;
      });
      
      // Add to master list and clear buffer for next room block
      allStudents.push(...tempStudents);
      tempStudents = [];
      continue;
    }
    
    // If indices are set, try to parse student row
    const idVal = row[idColIndex] !== undefined ? String(row[idColIndex]).trim() : "";
    const nameVal = row[nameColIndex] !== undefined ? String(row[nameColIndex]).trim() : "";
    const noVal = row[noColIndex] !== undefined ? String(row[noColIndex]).trim() : "";
    
    // Verify it is a student row: ID is a digit code (e.g. 5 digits) and Name is not empty
    const idRegex = /^\d{3,10}$/; // Matches 3 to 10 digit student IDs
    if (idRegex.test(idVal) && nameVal.length > 0) {
      // Clean multiple spaces in name (standardize names like "เด็กชายกฤษฎา  วรรทวี" -> "เด็กชายกฤษฎา วรรทวี")
      const cleanedName = nameVal.replace(/\s+/g, ' ').trim();
      const studentNo = parseInt(noVal, 10) || (tempStudents.length + 1);
      
      tempStudents.push({
        studentId: idVal,
        name: cleanedName,
        no: studentNo,
        password: idVal, // ID is the password
        grade: grade, // Temporary placeholder
        room: "" // Temporary placeholder (will be set by summary row)
      });
    }
  }
  
  // Safeguard: If sheet does not have summary row at the end, assign room "1" to remaining
  if (tempStudents.length > 0) {
    tempStudents.forEach(student => {
      student.room = "1";
      student.grade = grade;
    });
    allStudents.push(...tempStudents);
  }
  
  return allStudents;
}
