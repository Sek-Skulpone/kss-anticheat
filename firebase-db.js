// firebase-db.js - Database Manager using Firebase Firestore v8 Compatibility CDN

let db = null;

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

// Auto-init if configuration is saved in localStorage
const savedConfig = localStorage.getItem('kss_firebase_config');
if (savedConfig) {
  try {
    initFirebase(JSON.parse(savedConfig));
  } catch (e) {
    console.error("Auto-init failed:", e);
  }
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
        name: 'ผู้ดูแลระบบเริ่มต้น (Default Admin)'
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
  await store.collection('teachers').doc(teacher.username.trim()).set(teacher);
}

export async function deleteTeacher(username) {
  const store = checkDB();
  if (username === 'admin') {
    throw new Error("ไม่สามารถลบบัญชีผู้ดูแลระบบหลัก (admin) ได้");
  }
  await store.collection('teachers').doc(username).delete();
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
    maxScore: exam.questions ? exam.questions.length : 0
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
