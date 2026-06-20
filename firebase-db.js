// firebase-db.js - Database Manager using Firebase Firestore v8 Compatibility CDN

let db = null;

// Default Firebase Configuration for KSS School
const defaultFirebaseConfig = {
  apiKey: "AIzaSyATLsG-Zyeme9jqanzDS3eI5c57pgrwb3g",
  authDomain: "kssanticheat.firebaseapp.com",
  projectId: "kssanticheat",
  storageBucket: "kssanticheat.firebasestorage.app",
  messagingSenderId: "184077478712",
  appId: "1:184077478712:web:9208e293d0f6080ac52fcd"
};

export function isFirebaseInitialized() {
  return db !== null;
}

export function initFirebase(config) {
  try {
    if (!firebase.apps.length) {
      firebase.initializeApp(config);
    }
    db = firebase.firestore();
    
    // Enable offline persistence for better resilience
    db.enablePersistence().catch((err) => {
      console.warn("Firestore persistence failed:", err.code);
    });
    
    return true;
  } catch (error) {
    console.error("Firebase Initialization Error:", error);
    return false;
  }
}

// Auto-init if configuration is saved in localStorage, or fallback to school default
const savedConfig = localStorage.getItem('kss_firebase_config');
if (savedConfig) {
  try {
    initFirebase(JSON.parse(savedConfig));
  } catch (e) {
    console.error("Auto-init from localStorage failed:", e);
    initFirebase(defaultFirebaseConfig);
  }
} else {
  initFirebase(defaultFirebaseConfig);
}

// Generic Helper to check DB initialization
function checkDB() {
  if (!db) {
    throw new Error("Firebase has not been initialized yet. Please configure the database connections.");
  }
  return db;
}

/* ==========================================================================
   1. TEACHER AUTHENTICATION & MANAGEMENT
   ========================================================================== */

export async function loginTeacher(username, password) {
  const store = checkDB();
  const docRef = store.collection('teachers').doc(username.trim());
  const doc = await docRef.get();
  
  if (!doc.exists) {
    // If database has 0 teachers, bootstrap a default admin account
    const snapshot = await store.collection('teachers').limit(1).get();
    if (snapshot.empty && username === 'admin' && password === 'kss12345') {
      const defaultTeacher = {
        username: 'admin',
        password: 'kss12345',
        name: 'ผู้ดูแลระบบเริ่มต้น (Default Admin)',
        role: 'admin'
      };
      await store.collection('teachers').doc('admin').set(defaultTeacher);
      return defaultTeacher;
    }
    throw new Error("ไม่พบชื่อผู้ใช้นี้ในระบบ");
  }
  
  const teacher = doc.data();
  if (teacher.password !== password) {
    throw new Error("รหัสผ่านไม่ถูกต้อง");
  }
  
  // Backwards compatibility for roles
  if (!teacher.role) {
    teacher.role = (username.trim() === 'admin') ? 'admin' : 'teacher';
  }
  
  return teacher;
}

export async function getTeachers() {
  const store = checkDB();
  const snapshot = await store.collection('teachers').get();
  const list = [];
  snapshot.forEach(doc => {
    list.push(doc.data());
  });
  return list;
}

export async function saveTeacher(teacher) {
  const store = checkDB();
  if (!teacher.role) {
    teacher.role = 'teacher';
  }
  await store.collection('teachers').doc(teacher.username.trim()).set(teacher);
}

export async function deleteTeacher(username) {
  const store = checkDB();
  if (username === 'admin') {
    throw new Error("ไม่สามารถลบบัญชีผู้ดูแลระบบหลัก (admin) ได้");
  }
  await store.collection('teachers').doc(username).delete();
}

export async function updateTeacherRole(username, role) {
  const store = checkDB();
  if (username === 'admin') {
    throw new Error("ไม่สามารถเปลี่ยนสิทธิ์ของบัญชีผู้ดูแลระบบหลัก (admin) ได้");
  }
  await store.collection('teachers').doc(username).update({
    role: role
  });
}

export async function updateTeacherPassword(username, newPassword) {
  const store = checkDB();
  await store.collection('teachers').doc(username).update({
    password: newPassword
  });
}

export async function changeOwnPassword(username, oldPassword, newPassword) {
  const store = checkDB();
  const docRef = store.collection('teachers').doc(username.trim());
  const doc = await docRef.get();
  
  if (!doc.exists) {
    throw new Error("ไม่พบข้อมูลผู้ใช้ในระบบ");
  }
  
  const teacher = doc.data();
  if (teacher.password !== oldPassword) {
    throw new Error("รหัสผ่านเดิมไม่ถูกต้อง");
  }
  
  await docRef.update({
    password: newPassword
  });
}

/* ==========================================================================
   2. STUDENT DATABASE (EXCEL IMPORT)
   ========================================================================== */

export async function loginStudent(studentId, password) {
  const store = checkDB();
  const docRef = store.collection('students').doc(studentId.trim());
  const doc = await docRef.get();
  
  if (!doc.exists) {
    throw new Error("ไม่พบรหัสประจำตัวนักเรียนนี้ในระบบ");
  }
  
  const student = doc.data();
  if (student.password !== password) {
    throw new Error("รหัสประจำตัวหรือรหัสผ่านไม่ถูกต้อง");
  }
  return student;
}

// Batch upload students (optimizing with Chunking to prevent Firebase limits)
export async function uploadStudentsBatch(studentsList) {
  const store = checkDB();
  
  // We can upload in batches of 400 to avoid Firestore limits (max 500 per batch)
  const batchSize = 400;
  for (let i = 0; i < studentsList.length; i += batchSize) {
    const chunk = studentsList.slice(i, i + batchSize);
    const batch = store.batch();
    
    chunk.forEach(student => {
      const docRef = store.collection('students').doc(student.studentId);
      batch.set(docRef, student);
    });
    
    await batch.commit();
  }
}

export async function getAllStudents() {
  const store = checkDB();
  const snapshot = await store.collection('students').get();
  const list = [];
  snapshot.forEach(doc => {
    list.push(doc.data());
  });
  return list;
}

export async function deleteStudent(studentId) {
  const store = checkDB();
  await store.collection('students').doc(studentId).delete();
}

export async function clearAllStudents() {
  const store = checkDB();
  const snapshot = await store.collection('students').get();
  const batch = store.batch();
  snapshot.forEach(doc => {
    batch.delete(doc.ref);
  });
  await batch.commit();
}

/* ==========================================================================
   3. EXAM TIMETABLE & QUESTION BANK
   ========================================================================== */

export async function saveExam(exam) {
  const store = checkDB();
  if (!exam.examId) {
    exam.examId = exam.date + "_" + exam.subjectCode + "_" + Math.random().toString(36).substring(2, 7);
  }
  if (!exam.linkStatus) {
    exam.linkStatus = 'auto'; // Default link override status
  }
  await store.collection('exams').doc(exam.examId).set(exam);
  return exam.examId;
}

export async function getExams() {
  const store = checkDB();
  const snapshot = await store.collection('exams').get();
  const list = [];
  snapshot.forEach(doc => {
    list.push(doc.data());
  });
  return list;
}

export async function deleteExam(examId) {
  const store = checkDB();
  await store.collection('exams').doc(examId).delete();
  
  // Also clean up any associated exam sessions
  const sessionsSnapshot = await store.collection('exam_sessions').where('examId', '==', examId).get();
  const batch = store.batch();
  sessionsSnapshot.forEach(doc => {
    batch.delete(doc.ref);
  });
  await batch.commit();
}

/* ==========================================================================
   4. EXAM SESSION TRACKING & REAL-TIME CONTROLS
   ========================================================================== */

// Starts or resumes an exam session for a student
export async function startOrCreateExamSession(student, exam) {
  const store = checkDB();
  const docId = `${student.studentId}_${exam.examId}`;
  const docRef = store.collection('exam_sessions').doc(docId);
  const doc = await docRef.get();
  
  if (doc.exists) {
    return doc.data();
  }
  
  const newSession = {
    studentId: student.studentId,
    studentName: student.name,
    examId: exam.examId,
    subjectCode: exam.subjectCode,
    subjectName: exam.subjectName,
    grade: student.grade,
    room: student.room,
    no: student.no,
    status: 'active', // active | locked | submitted
    violationCount: 0,
    violations: [],
    startedAt: new Date().toISOString(),
    submittedAt: null,
    answers: {},
    score: 0,
    maxScore: exam.questions ? exam.questions.length : 0,
    academicYear: exam.academicYear || '',
    term: exam.term || ''
  };
  
  await docRef.set(newSession);
  return newSession;
}

// Listens to a student's own exam session in real-time (detects unlocking instantly!)
export function listenToExamSession(studentId, examId, onUpdate) {
  const store = checkDB();
  const docId = `${studentId}_${examId}`;
  return store.collection('exam_sessions').doc(docId).onSnapshot((doc) => {
    if (doc.exists) {
      onUpdate(doc.data());
    }
  }, (error) => {
    console.error("Error listening to exam session:", error);
  });
}

// Log a violation client-side and push to Firestore
export async function recordViolation(studentId, examId, violationType, detail) {
  const store = checkDB();
  const docId = `${studentId}_${examId}`;
  const docRef = store.collection('exam_sessions').doc(docId);
  
  await store.runTransaction(async (transaction) => {
    const doc = await transaction.get(docRef);
    if (!doc.exists) return;
    
    const data = doc.data();
    if (data.status === 'submitted') return; // Cannot violate after submission
    
    const newViolationCount = data.violationCount + 1;
    const newViolations = [...(data.violations || [])];
    newViolations.push({
      time: new Date().toLocaleTimeString('th-TH'),
      type: violationType,
      detail: detail
    });
    
    let newStatus = data.status;
    if (newViolationCount >= 3) {
      newStatus = 'locked';
    }
    
    transaction.update(docRef, {
      violationCount: newViolationCount,
      violations: newViolations,
      status: newStatus
    });
  });
}

// Update student answers in Firestore
export async function updateAnswers(studentId, examId, answers) {
  const store = checkDB();
  const docId = `${studentId}_${examId}`;
  await store.collection('exam_sessions').doc(docId).update({
    answers: answers
  });
}

// Submit finished exam
export async function submitExam(studentId, examId, answers, score, maxScore) {
  const store = checkDB();
  const docId = `${studentId}_${examId}`;
  await store.collection('exam_sessions').doc(docId).update({
    answers: answers,
    score: score,
    maxScore: maxScore,
    status: 'submitted',
    submittedAt: new Date().toISOString()
  });
}

// Listens to all active exam sessions in real-time (for Teacher Live Monitor)
export function listenToActiveSessions(onUpdate) {
  const store = checkDB();
  return store.collection('exam_sessions')
    .onSnapshot((snapshot) => {
      const list = [];
      snapshot.forEach(doc => {
        list.push(doc.data());
      });
      onUpdate(list);
    }, (error) => {
      console.error("Error listening to active sessions:", error);
    });
}

// Teacher unlocks a locked student exam
export async function unlockStudentSession(studentId, examId, teacherName) {
  const store = checkDB();
  const docId = `${studentId}_${examId}`;
  const docRef = store.collection('exam_sessions').doc(docId);
  
  await store.runTransaction(async (transaction) => {
    const doc = await transaction.get(docRef);
    if (!doc.exists) return;
    
    const data = doc.data();
    const newViolations = [...(data.violations || [])];
    newViolations.push({
      time: new Date().toLocaleTimeString('th-TH'),
      type: 'teacher_unlock',
      detail: `ครู ${teacherName} ปลดล็อกให้สอบต่อ`
    });
    
    transaction.update(docRef, {
      violationCount: 0,
      status: 'active',
      violations: newViolations
    });
  });
}

// Update exam link override status
export async function updateExamLinkStatus(examId, linkStatus) {
  const store = checkDB();
  await store.collection('exams').doc(examId).update({
    linkStatus: linkStatus
  });
}

/* ==========================================================================
   5. DYNAMIC GRADES MANAGEMENT
   ========================================================================== */

export async function getGrades() {
  const store = checkDB();
  const snapshot = await store.collection('grades').get();
  const list = [];
  snapshot.forEach(doc => {
    list.push(doc.data());
  });
  
  // Bootstrap default grades if empty
  if (list.length === 0) {
    const defaults = ["ม.1", "ม.2", "ม.3"];
    const batch = store.batch();
    const now = new Date().toISOString();
    
    defaults.forEach(g => {
      const docRef = store.collection('grades').doc(g);
      const gradeObj = { name: g, createdAt: now };
      batch.set(docRef, gradeObj);
      list.push(gradeObj);
    });
    await batch.commit();
  }
  
  // Sort grades alphabetically/numerically
  list.sort((a, b) => a.name.localeCompare(b.name, 'th'));
  return list;
}

export async function addGrade(gradeName) {
  const store = checkDB();
  const name = gradeName.trim();
  if (!name) throw new Error("ชื่อระดับชั้นต้องไม่ว่างเปล่า");
  
  const docRef = store.collection('grades').doc(name);
  const doc = await docRef.get();
  if (doc.exists) {
    throw new Error("มีระดับชั้นนี้ในระบบอยู่แล้ว");
  }
  
  await docRef.set({
    name: name,
    createdAt: new Date().toISOString()
  });
}

export async function deleteGrade(gradeName) {
  const store = checkDB();
  await store.collection('grades').doc(gradeName).delete();
}

/* ==========================================================================
   6. DYNAMIC ACADEMIC YEARS MANAGEMENT
   ========================================================================== */

export async function getAcademicYears() {
  const store = checkDB();
  const snapshot = await store.collection('academic_years').get();
  const list = [];
  snapshot.forEach(doc => {
    list.push(doc.data());
  });
  
  // Bootstrap default academic years if empty
  if (list.length === 0) {
    const defaults = ["2568"];
    const batch = store.batch();
    const now = new Date().toISOString();
    
    defaults.forEach(y => {
      const docRef = store.collection('academic_years').doc(y);
      const yearObj = { name: y, createdAt: now };
      batch.set(docRef, yearObj);
      list.push(yearObj);
    });
    await batch.commit();
  }
  
  list.sort((a, b) => b.name.localeCompare(a.name)); // Descending order
  return list;
}

export async function addAcademicYear(yearName) {
  const store = checkDB();
  const name = yearName.trim();
  if (!name) throw new Error("ปีการศึกษาต้องไม่ว่างเปล่า");
  
  const docRef = store.collection('academic_years').doc(name);
  const doc = await docRef.get();
  if (doc.exists) {
    throw new Error("มีปีการศึกษานี้ในระบบอยู่แล้ว");
  }
  
  await docRef.set({
    name: name,
    createdAt: new Date().toISOString()
  });
}

export async function deleteAcademicYear(yearName) {
  const store = checkDB();
  await store.collection('academic_years').doc(yearName).delete();
}

/* ==========================================================================
   7. ACTIVE EXAM PERIOD SETTINGS
   ========================================================================== */

export async function getActiveExamPeriod() {
  const store = checkDB();
  const docRef = store.collection('settings').doc('active_period');
  const doc = await docRef.get();
  if (doc.exists) {
    return doc.data();
  }
  
  const defaultPeriod = {
    activeYear: "2568",
    activeTerm: "ปลายภาค"
  };
  await store.collection('settings').doc('active_period').set(defaultPeriod);
  return defaultPeriod;
}

export async function saveActiveExamPeriod(activeYear, activeTerm) {
  const store = checkDB();
  await store.collection('settings').doc('active_period').set({
    activeYear: activeYear,
    activeTerm: activeTerm,
    updatedAt: new Date().toISOString()
  });
}

