// app.js - Main Application controller and SPA state machine
import * as dbManager from './firebase-db.js';
import { parseStudentExcel } from './xlsx-parser.js';
import { parseExamDocx } from './docx-parser.js';

// Application State
let currentUser = null; // Stores either student or teacher object
let currentRole = null; // 'student' or 'teacher'
let currentExam = null; // Exam student is currently taking
let activeRealtimeUnsubscribe = null; // Real-time listener for current student exam session
let activeLiveMonitorUnsubscribe = null; // Real-time listener for teacher live monitor
let examTimerInterval = null; // Interval timer for exam countdown
let studentTimerIntervals = {}; // Countdown timers for student dashboard schedule rows
let localViolationCount = 0;
let isViolationCooldown = false; // Throttle to prevent duplicate violation triggers
let currentParsedQuestions = []; // Temporary storage for docx parsed questions before saving

// Navigation / Router
function showScreen(screenId) {
  // Hide all screens
  document.querySelectorAll('.screen').forEach(s => s.classList.add('hidden'));
  // Show target screen
  const target = document.getElementById(screenId);
  if (target) target.classList.remove('hidden');
}

// UI Notification Helper
function showNotification(message, isError = false) {
  const toast = document.getElementById('toast');
  const toastMsg = document.getElementById('toast-message');
  
  toastMsg.textContent = message;
  if (isError) {
    toast.classList.add('error-toast');
    toast.querySelector('i').className = 'fa-solid fa-circle-exclamation';
  } else {
    toast.classList.remove('error-toast');
    toast.querySelector('i').className = 'fa-solid fa-circle-check';
  }
  
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, 4000);
}

/* ==========================================================================
   1. FIREBASE SETUP & INITIALIZATION
   ========================================================================== */

function checkFirebaseConnection() {
  showScreen('screen-login');
}

/* ==========================================================================
   2. AUTHENTICATION (LOGIN / LOGOUT)
   ========================================================================== */

const formStudentLogin = document.getElementById('form-student-login');
const formTeacherLogin = document.getElementById('form-teacher-login');
const toggleStudentLogin = document.getElementById('toggle-student-login');
const toggleTeacherLogin = document.getElementById('toggle-teacher-login');

// Toggle forms
toggleStudentLogin.addEventListener('click', () => {
  toggleStudentLogin.classList.add('active');
  toggleTeacherLogin.classList.remove('active');
  formStudentLogin.classList.remove('hidden');
  formTeacherLogin.classList.add('hidden');
});

toggleTeacherLogin.addEventListener('click', () => {
  toggleTeacherLogin.classList.add('active');
  toggleStudentLogin.classList.remove('active');
  formTeacherLogin.classList.remove('hidden');
  formStudentLogin.classList.add('hidden');
});

// Student Login
formStudentLogin.addEventListener('submit', async (e) => {
  e.preventDefault();
  const id = document.getElementById('student-username').value.trim();
  const pass = document.getElementById('student-password').value.trim();
  
  try {
    const student = await dbManager.loginStudent(id, pass);
    currentUser = student;
    currentRole = 'student';
    sessionStorage.setItem('kss_user', JSON.stringify(student));
    sessionStorage.setItem('kss_role', 'student');
    
    showNotification(`ยินดีต้อนรับ เข้าสู่ระบบสอบ ชั้น ${student.grade}/${student.room}`);
    setupStudentDashboard();
  } catch (err) {
    showNotification(err.message, true);
  }
});

// Teacher Login
formTeacherLogin.addEventListener('submit', async (e) => {
  e.preventDefault();
  const user = document.getElementById('teacher-username').value.trim();
  const pass = document.getElementById('teacher-password').value.trim();
  
  try {
    const teacher = await dbManager.loginTeacher(user, pass);
    currentUser = teacher;
    currentRole = 'teacher';
    sessionStorage.setItem('kss_user', JSON.stringify(teacher));
    sessionStorage.setItem('kss_role', 'teacher');
    
    showNotification(`ยินดีต้อนรับ คุณครู ${teacher.name}`);
    setupTeacherDashboard();
  } catch (err) {
    showNotification(err.message, true);
  }
});

// Logouts
document.getElementById('btn-student-logout').addEventListener('click', () => {
  // Clear countdown intervals
  Object.values(studentTimerIntervals).forEach(clearInterval);
  studentTimerIntervals = {};
  
  currentUser = null;
  currentRole = null;
  sessionStorage.clear();
  showScreen('screen-login');
});

document.getElementById('btn-teacher-logout').addEventListener('click', () => {
  if (activeLiveMonitorUnsubscribe) {
    activeLiveMonitorUnsubscribe();
    activeLiveMonitorUnsubscribe = null;
  }
  currentUser = null;
  currentRole = null;
  sessionStorage.clear();
  showScreen('screen-login');
});

/* ==========================================================================
   3. STUDENT PORTAL (TIMETABLE VIEW & SCHEDULE COUNTDOWNS)
   ========================================================================== */

async function setupStudentDashboard() {
  document.getElementById('student-display-name').textContent = currentUser.name;
  document.getElementById('student-display-meta').textContent = `ชั้น ${currentUser.grade}/${currentUser.room} | เลขที่ ${currentUser.no}`;
  document.getElementById('student-timetable-title-meta').textContent = `ชั้นมัธยมศึกษาปีที่ ${currentUser.grade.replace('ม.', '')} ห้อง ${currentUser.room}`;
  
  showScreen('screen-student-dashboard');
  await loadStudentTimetable();
}

async function loadStudentTimetable() {
  try {
    // Clear existing countdowns
    Object.values(studentTimerIntervals).forEach(clearInterval);
    studentTimerIntervals = {};
    
    const activePeriod = await dbManager.getActiveExamPeriod();
    const titleEl = document.getElementById('student-timetable-title');
    if (titleEl) {
      titleEl.textContent = `ตารางสอบ${activePeriod.activeTerm} ประจำปีการศึกษา ${activePeriod.activeYear}`;
    }
    
    const exams = await dbManager.getExams();
    // Filter exams for student's grade, active year, and active term
    const studentExams = exams.filter(ex => 
      ex.grade === currentUser.grade &&
      ex.academicYear === activePeriod.activeYear &&
      ex.term === activePeriod.activeTerm
    );
    
    // Sort exams by date, then by start time
    studentExams.sort((a, b) => {
      if (a.date !== b.date) return new Date(a.date) - new Date(b.date);
      return a.startTime.localeCompare(b.startTime);
    });
    
    // Get unique dates for tabs
    const uniqueDates = [...new Set(studentExams.map(ex => ex.date))];
    uniqueDates.sort();
    
    renderDateTabs(uniqueDates, studentExams);
    
    // Default show current day if matches, otherwise show first date
    const todayStr = new Date().toISOString().split('T')[0];
    const defaultDate = uniqueDates.includes(todayStr) ? todayStr : uniqueDates[0];
    
    if (defaultDate) {
      renderTimetableForDate(defaultDate, studentExams);
      // Mark active date tab
      const tabBtn = document.querySelector(`.date-tab-btn[data-date="${defaultDate}"]`);
      if (tabBtn) tabBtn.classList.add('active');
    } else {
      document.getElementById('student-timetable-body').innerHTML = `
        <tr>
          <td colspan="7" style="padding: 2rem; color: var(--text-muted);">ไม่มีรายการสอบของระดับชั้น ${currentUser.grade} ในขณะนี้</td>
        </tr>
      `;
    }
  } catch (err) {
    showNotification("ไม่สามารถดึงข้อมูลตารางสอบได้: " + err.message, true);
  }
}

function renderDateTabs(dates, allExams) {
  const tabsContainer = document.getElementById('student-dates-tabs');
  tabsContainer.innerHTML = '';
  
  dates.forEach(dStr => {
    const formattedDate = formatThaiDateShort(dStr);
    const btn = document.createElement('button');
    btn.className = 'tab-link date-tab-btn';
    btn.setAttribute('data-date', dStr);
    btn.innerHTML = `<i class="fa-regular fa-calendar"></i> ${formattedDate}`;
    btn.addEventListener('click', () => {
      document.querySelectorAll('.date-tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      renderTimetableForDate(dStr, allExams);
    });
    tabsContainer.appendChild(btn);
  });
}

function renderTimetableForDate(dateStr, exams) {
  const tbody = document.getElementById('student-timetable-body');
  tbody.innerHTML = '';
  
  // Clear any counting tickers in this view
  Object.values(studentTimerIntervals).forEach(clearInterval);
  studentTimerIntervals = {};
  
  const dailyExams = exams.filter(ex => ex.date === dateStr);
  
  if (dailyExams.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="padding: 2rem; color: var(--text-muted);">ไม่มีการสอบในวันนี้</td>
      </tr>
    `;
    return;
  }
  
  // Helper to construct Date objects
  const getExamTime = (date, timeStr) => new Date(`${date}T${timeStr}:00`);
  
  dailyExams.forEach((exam, index) => {
    // Add lunch break row if after the morning sessions (e.g. around 11:40)
    // In our timetable templates, lunch starts around 11:40 - 13:00
    if (index > 0 && dailyExams[index - 1].endTime <= "11:40" && exam.startTime >= "13:00") {
      const lunchRow = document.createElement('tr');
      lunchRow.className = 'lunch-row';
      lunchRow.innerHTML = `
        <td colspan="2">11.40 น. - 13.00 น.</td>
        <td>80</td>
        <td colspan="4" style="font-weight: bold; letter-spacing: 0.1em;">พักกลางวัน</td>
      `;
      tbody.appendChild(lunchRow);
    }
    
    const row = document.createElement('tr');
    
    // 1. Day Column (Span all exams or just render)
    const formattedDay = formatThaiDay(dateStr);
    
    // 2. Exam Time Range (24h format)
    const timeRange = `${exam.startTime.replace(':', '.')} น. - ${exam.endTime.replace(':', '.')} น.`;
    
    // 3. Action badge based on type: paper or online link (handles teacher overrides)
    let actionBadge = '';
    const uniqueRowId = `exam-action-${exam.examId}`;
    
    if (exam.examType === 'paper') {
      actionBadge = `<span class="badge badge-paper">Paper</span>`;
    } else {
      if (exam.linkStatus === 'released') {
        actionBadge = `<span class="badge badge-link" onclick="enterExamRoomById('${exam.examId}')"><i class="fa-solid fa-play"></i> เข้าสู่ห้องสอบ</span>`;
      } else if (exam.linkStatus === 'hidden') {
        actionBadge = `<span class="badge badge-ended">ปิดรับการเข้าสอบ</span>`;
      } else {
        // Default: Auto (Scheduled countdown)
        actionBadge = `<span class="badge badge-countdown" id="${uniqueRowId}">กำลังตรวจสอบ...</span>`;
        
        const startDateTime = getExamTime(exam.date, exam.startTime);
        const endDateTime = getExamTime(exam.date, exam.endTime);
        
        const updateRowStatus = () => {
          const now = new Date();
          const element = document.getElementById(uniqueRowId);
          if (!element) return;
          
          if (now < startDateTime) {
            // Future exam - show countdown
            const diffMs = startDateTime - now;
            const diffSec = Math.floor(diffMs / 1000);
            const hrs = String(Math.floor(diffSec / 3600)).padStart(2, '0');
            const mins = String(Math.floor((diffSec % 3600) / 60)).padStart(2, '0');
            const secs = String(diffSec % 60).padStart(2, '0');
            
            element.className = 'badge badge-countdown';
            element.innerHTML = `<i class="fa-regular fa-clock"></i> รอสอบ: ${hrs}:${mins}:${secs}`;
            element.onclick = null;
          } else if (now >= startDateTime && now <= endDateTime) {
            // Active exam - show link
            element.className = 'badge badge-link';
            element.innerHTML = `<i class="fa-solid fa-play"></i> เข้าสู่ห้องสอบ`;
            element.style.cursor = 'pointer';
            element.onclick = () => enterExamRoom(exam);
          } else {
            // Passed exam
            element.className = 'badge badge-ended';
            element.innerHTML = `สิ้นสุดการสอบ`;
            element.onclick = null;
          }
        };
        
        updateRowStatus();
        studentTimerIntervals[exam.examId] = setInterval(updateRowStatus, 1000);
      }
    }
    
    row.innerHTML = `
      <td class="day-cell">${formattedDay}</td>
      <td>${timeRange}</td>
      <td>${exam.duration}</td>
      <td style="font-weight: 500; color: var(--accent-color);">${exam.subjectCode}</td>
      <td style="text-align: left; padding-left: 1.5rem;">${exam.subjectName}</td>
      <td>${exam.room || '-'}</td>
      <td>${actionBadge}</td>
    `;
    
    tbody.appendChild(row);
  });
}

/* ==========================================================================
   4. ACTIVE EXAM TEST ROOM & ANTI-CHEAT TRIGGERS
   ========================================================================== */

async function enterExamRoom(exam) {
  currentExam = exam;
  
  if (!exam.questions || exam.questions.length === 0) {
    showNotification("ข้อสอบนี้ยังไม่มีข้อมูลคำถาม กรุณาแจ้งครูผู้คุมสอบ", true);
    return;
  }
  
  try {
    // 1. Initialize or load existing exam session from Firestore
    const session = await dbManager.startOrCreateExamSession(currentUser, exam);
    
    // Check if session is already completed
    if (session.status === 'submitted') {
      showNotification("ท่านได้ส่งข้อสอบนี้ไปเรียบร้อยแล้ว", true);
      return;
    }
    
    // Check if session is locked
    if (session.status === 'locked') {
      enterLockedScreen(session);
      return;
    }
    
    // 2. Prep exam UI & Shuffle questions
    localViolationCount = session.violationCount || 0;
    
    // Setup and render exam
    setupActiveExamWindow(session);
  } catch (err) {
    showNotification("ไม่สามารถเข้าสู่ห้องสอบได้: " + err.message, true);
  }
}

function setupActiveExamWindow(session) {
  // Update UI top bar
  document.getElementById('exam-display-subject').textContent = `${currentExam.subjectCode} ${currentExam.subjectName}`;
  document.getElementById('exam-display-student-info').textContent = `ผู้เข้าสอบ: ${currentUser.name} (ชั้น ม.${currentUser.grade.replace('ม.', '')}/${currentUser.room} เลขที่ ${currentUser.no})`;
  updateViolationsBadge();
  
  // Render questions
  renderExamQuestions(session);
  
  // Start ticking clock countdown
  startExamTimer();
  
  // Go to exam screen
  showScreen('screen-exam-session');
  
  // Enter Fullscreen mode to secure exam
  requestFullscreenMode();
  
  // Add browser security listeners
  enableAntiCheatListeners();
  
  // Listen to remote unlocking (in case they get locked, but this also watches general db changes)
  listenToExamSessionUpdates();
}

function updateViolationsBadge() {
  const badge = document.getElementById('exam-violations-badge');
  badge.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> สลับหน้าจอ: ${localViolationCount}/3 ครั้ง`;
  
  if (localViolationCount === 1) {
    badge.style.color = 'var(--warning-color)';
    badge.style.borderColor = 'var(--warning-color)';
    badge.style.background = 'rgba(244, 162, 97, 0.1)';
  } else if (localViolationCount >= 2) {
    badge.style.color = 'var(--danger-color)';
    badge.style.borderColor = 'var(--danger-color)';
    badge.style.background = 'rgba(230, 57, 70, 0.1)';
  }
}

// Fisher-Yates Shuffle algorithm
function shuffleArray(array) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

function renderExamQuestions(session) {
  const questionsList = document.getElementById('exam-questions-list');
  questionsList.innerHTML = '';
  
  const navGrid = document.getElementById('exam-nav-grid');
  navGrid.innerHTML = '';
  
  // To keep shuffling consistent across tab reloads during the same session,
  // we can use a seed or store the shuffled sequence in sessionStorage if wanted.
  // Here, we'll store the shuffled sequence in sessionStorage for the duration of the exam.
  const sessionKey = `shuffled_questions_${currentUser.studentId}_${currentExam.examId}`;
  let shuffled = [];
  
  const savedShuffled = sessionStorage.getItem(sessionKey);
  if (savedShuffled) {
    shuffled = JSON.parse(savedShuffled);
  } else {
    // Shuffle questions
    const shuffledQs = shuffleArray(currentExam.questions);
    // Shuffle choices for each question
    shuffled = shuffledQs.map((q) => {
      return {
        ...q,
        choices: shuffleArray(q.choices)
      };
    });
    sessionStorage.setItem(sessionKey, JSON.stringify(shuffled));
  }
  
  // Render each question
  shuffled.forEach((q, index) => {
    const qIndex = index + 1;
    const qCard = document.createElement('div');
    qCard.className = 'glass-card question-card';
    qCard.id = `question-block-${q.id}`;
    
    // Highlight if question is answered
    const selectedAnswerId = session.answers ? session.answers[q.id] : null;
    
    // Choices list HTML
    let choicesHtml = '';
    q.choices.forEach((c) => {
      const isSelected = selectedAnswerId === c.id;
      choicesHtml += `
        <div class="choice-option ${isSelected ? 'selected' : ''}" 
             data-question-id="${q.id}" 
             data-choice-id="${c.id}" 
             onclick="selectChoice(this)">
          <input type="radio" name="radio-${q.id}" class="choice-radio" ${isSelected ? 'checked' : ''}>
          <span class="choice-text">${c.text}</span>
        </div>
      `;
    });
    
    qCard.innerHTML = `
      <div class="question-text">${qIndex}. ${q.questionText.replace(/^\d+\.\s*/, '')}</div>
      <div class="choices-grid">${choicesHtml}</div>
    `;
    questionsList.appendChild(qCard);
    
    // Question Navigator button
    const navBtn = document.createElement('button');
    navBtn.className = `nav-btn ${selectedAnswerId ? 'answered' : ''}`;
    navBtn.id = `nav-btn-${q.id}`;
    navBtn.textContent = qIndex;
    navBtn.addEventListener('click', () => {
      document.getElementById(`question-block-${q.id}`).scrollIntoView({ behavior: 'smooth', block: 'center' });
      // Toggle active states in nav
      document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
      navBtn.classList.add('active');
    });
    navGrid.appendChild(navBtn);
  });
}

// Global scope choice selection wrapper for elements
window.selectChoice = function(element) {
  const qId = element.getAttribute('data-question-id');
  const cId = element.getAttribute('data-choice-id');
  
  // Update UI selection classes
  document.querySelectorAll(`.choice-option[data-question-id="${qId}"]`).forEach(opt => {
    opt.classList.remove('selected');
    opt.querySelector('input').checked = false;
  });
  
  element.classList.add('selected');
  element.querySelector('input').checked = true;
  
  // Highlight navigator button
  const navBtn = document.getElementById(`nav-btn-${qId}`);
  if (navBtn) navBtn.classList.add('answered');
  
  // Save answer locally in sessionStorage first
  const sessionKey = `answers_${currentUser.studentId}_${currentExam.examId}`;
  const answers = JSON.parse(sessionStorage.getItem(sessionKey) || '{}');
  answers[qId] = cId;
  sessionStorage.setItem(sessionKey, JSON.stringify(answers));
  
  // Write to Firebase Firestore (Debounced or instant. We write instant for absolute resilience)
  dbManager.updateAnswers(currentUser.studentId, currentExam.examId, answers).catch(err => {
    console.error("Failed to sync answer with Firestore:", err);
  });
};

function startExamTimer() {
  if (examTimerInterval) clearInterval(examTimerInterval);
  
  const timerBox = document.getElementById('exam-timer');
  const examEndTimeStr = `${currentExam.date}T${currentExam.endTime}:00`;
  const endTime = new Date(examEndTimeStr);
  
  const updateTimer = () => {
    const now = new Date();
    const remainingSec = Math.max(0, Math.floor((endTime - now) / 1000));
    
    if (remainingSec <= 0) {
      clearInterval(examTimerInterval);
      timerBox.textContent = "00:00";
      showNotification("หมดเวลาทำข้อสอบแล้ว ระบบกำลังส่งข้อสอบอัตโนมัติ...", true);
      submitStudentExam(true); // Forced submit
      return;
    }
    
    const hrs = Math.floor(remainingSec / 3600);
    const mins = Math.floor((remainingSec % 3600) / 60);
    const secs = remainingSec % 60;
    
    let displayTime = '';
    if (hrs > 0) {
      displayTime = `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
    } else {
      displayTime = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
    }
    
    timerBox.textContent = displayTime;
    
    // Urgent alarm (less than 5 minutes left)
    if (remainingSec < 300) {
      timerBox.classList.add('urgent');
    } else {
      timerBox.classList.remove('urgent');
    }
  };
  
  updateTimer();
  examTimerInterval = setInterval(updateTimer, 1000);
}

function listenToExamSessionUpdates() {
  if (activeRealtimeUnsubscribe) activeRealtimeUnsubscribe();
  
  activeRealtimeUnsubscribe = dbManager.listenToExamSession(currentUser.studentId, currentExam.examId, (session) => {
    // If exam state is updated to locked from database (teacher locked, or third strike sync)
    if (session.status === 'locked') {
      enterLockedScreen(session);
    }
    // If exam was locked, but is now unlocked by teacher!
    if (session.status === 'active' && document.getElementById('screen-locked').classList.contains('hidden') === false) {
      // Re-initialize exam room
      showNotification("คุณครูได้ทำการปลดล็อกระบบให้คุณแล้ว สามารถทำข้อสอบต่อได้");
      setupActiveExamWindow(session);
    }
  });
}

// Student submits exam
async function submitStudentExam(isForced = false) {
  // Clear intervals & unsubscribe
  if (examTimerInterval) clearInterval(examTimerInterval);
  disableAntiCheatListeners();
  if (activeRealtimeUnsubscribe) {
    activeRealtimeUnsubscribe();
    activeRealtimeUnsubscribe = null;
  }
  
  // Calculate score
  const sessionKey = `answers_${currentUser.studentId}_${currentExam.examId}`;
  const answers = JSON.parse(sessionStorage.getItem(sessionKey) || '{}');
  
  let score = 0;
  currentExam.questions.forEach((q) => {
    const studentChoiceId = answers[q.id];
    const correctChoice = q.choices[q.correctChoiceIndex];
    if (studentChoiceId && correctChoice && studentChoiceId === correctChoice.id) {
      score++;
    }
  });
  
  try {
    await dbManager.submitExam(
      currentUser.studentId, 
      currentExam.examId, 
      answers, 
      score, 
      currentExam.questions.length
    );
    
    // Exit fullscreen
    if (document.fullscreenElement) {
      document.exitFullscreen().catch(e => {});
    }
    
    // Clear session states
    sessionStorage.removeItem(sessionKey);
    sessionStorage.removeItem(`shuffled_questions_${currentUser.studentId}_${currentExam.examId}`);
    
    showNotification("ส่งกระดาษคำตอบเรียบร้อยแล้ว!");
    
    // Return to student schedule dashboard
    setupStudentDashboard();
  } catch (err) {
    showNotification("เกิดข้อผิดพลาดในการส่งข้อสอบ: " + err.message, true);
  }
}

// Confirm submit dialog
document.getElementById('btn-submit-exam').addEventListener('click', () => {
  const sessionKey = `answers_${currentUser.studentId}_${currentExam.examId}`;
  const answers = JSON.parse(sessionStorage.getItem(sessionKey) || '{}');
  const answeredCount = Object.keys(answers).length;
  const totalCount = currentExam.questions.length;
  const unansweredCount = totalCount - answeredCount;
  
  const submitConfirmText = document.getElementById('submit-confirm-unanswered');
  if (unansweredCount > 0) {
    submitConfirmText.textContent = `คำเตือน! ท่านยังไม่ได้ตอบคำถามอีก ${unansweredCount} ข้อ`;
    submitConfirmText.style.color = 'var(--danger-color)';
  } else {
    submitConfirmText.textContent = 'ยอดเยี่ยม! ท่านตอบคำถามครบถ้วนทุกข้อแล้ว';
    submitConfirmText.style.color = 'var(--success-color)';
  }
  
  document.getElementById('modal-confirm-submit').classList.remove('hidden');
});

document.getElementById('btn-confirm-submit-ok').addEventListener('click', () => {
  document.getElementById('modal-confirm-submit').classList.add('hidden');
  submitStudentExam(false);
});

document.getElementById('btn-confirm-submit-cancel').addEventListener('click', () => {
  document.getElementById('modal-confirm-submit').classList.add('hidden');
});

/* ==========================================================================
   5. ANTI-CHEAT ENGINE MECHANICS (STRIKE LOCK SYSTEM)
   ========================================================================== */

function requestFullscreenMode() {
  if (!document.fullscreenElement) {
    document.documentElement.requestFullscreen().catch((err) => {
      console.warn("Could not request Fullscreen mode:", err);
    });
  }
}

function handleViolation(type, detail) {
  // Guard checks
  if (localViolationCount >= 3) return; // Already locked
  if (isViolationCooldown) return; // Prevent double alerts
  if (!document.getElementById('screen-exam-session').classList.contains('hidden') === false && 
      document.getElementById('modal-warning').classList.contains('hidden') === false) return; // Warning popup is already visible
      
  isViolationCooldown = true;
  setTimeout(() => { isViolationCooldown = false; }, 3000); // 3 sec throttle cooldown
  
  localViolationCount++;
  
  // Push violation to firestore
  dbManager.recordViolation(currentUser.studentId, currentExam.examId, type, detail).catch(e => {
    console.error("Failed to log violation to Firestore:", e);
  });
  
  updateViolationsBadge();
  
  if (localViolationCount >= 3) {
    // 3 strikes: lock exam immediately
    const sessionDocId = `${currentUser.studentId}_${currentExam.examId}`;
    dbManager.listenToExamSession(currentUser.studentId, currentExam.examId, (session) => {
      enterLockedScreen(session);
    });
  } else {
    // Show intermediate warning modal
    const modal = document.getElementById('modal-warning');
    const warningContent = document.getElementById('warning-modal-content');
    const warningCount = document.getElementById('warning-modal-count');
    
    warningCount.textContent = `การละเมิดกฎครั้งที่: ${localViolationCount} / 3 ครั้ง`;
    
    if (localViolationCount === 2) {
      warningContent.classList.add('danger-level');
      document.getElementById('warning-modal-desc').innerHTML = `
        <b>คำเตือนครั้งสุดท้าย!</b> การละเมิดสลับหน้าจอครั้งต่อไปจะส่งผลให้ระบบข้อสอบของท่านถูกระงับ (Locked) และต้องให้ผู้คุมสอบเดินมาปลดล็อกที่เครื่องของท่านเท่านั้น
      `;
    } else {
      warningContent.classList.remove('danger-level');
      document.getElementById('warning-modal-desc').textContent = `
        ท่านสลับแท็บบราวเซอร์ ออกจากหน้าต่าง หรือกดออกจาก Fullscreen ซึ่งขัดต่อระเบียบการสอบ กรุณาทำข้อสอบภายในกรอบหน้าต่างสอบที่กำหนด
      `;
    }
    
    modal.classList.remove('hidden');
  }
}

// Ack violation warning
document.getElementById('btn-ack-warning').addEventListener('click', () => {
  document.getElementById('modal-warning').classList.add('hidden');
  // Re-secure screen
  requestFullscreenMode();
});

function enterLockedScreen(session) {
  // Unsubscribe exam updates and clear clocks
  disableAntiCheatListeners();
  if (examTimerInterval) clearInterval(examTimerInterval);
  
  // Render violation log in locked screen
  const logsList = document.getElementById('lock-violations-list');
  logsList.innerHTML = '';
  
  if (session.violations && session.violations.length > 0) {
    session.violations.forEach(v => {
      const item = document.createElement('div');
      item.className = 'violation-log-item';
      item.innerHTML = `[${v.time}] <b>${v.type === 'teacher_unlock' ? 'การปลดล็อก' : 'ผิดกฎ'}</b>: <span>${v.detail}</span>`;
      logsList.appendChild(item);
    });
  } else {
    logsList.innerHTML = '<div class="violation-log-item">ไม่มีบันทึกข้อมูลการสลับจอ</div>';
  }
  
  showScreen('screen-locked');
  
  // Also start watching changes on this document to auto-unlock!
  listenToExamSessionUpdates();
}

// Prevent keys & clicks helper functions
const preventEvent = (e) => e.preventDefault();
function preventKeys(e) {
  // Prevent Alt+Tab, Windows, Command keys or browser dev tools (F12, Ctrl+Shift+I, Ctrl+U)
  if (e.key === 'F12' || 
      (e.ctrlKey && e.shiftKey && e.key === 'I') || 
      (e.ctrlKey && e.key === 'u') || 
      (e.ctrlKey && e.key === 'c') || 
      (e.ctrlKey && e.key === 'v') || 
      (e.ctrlKey && e.key === 'x') || 
      (e.ctrlKey && e.key === 'a')) {
    e.preventDefault();
    showNotification("การกดปุ่มทางลัดเหล่านี้ถูกระงับเพื่อป้องกันการคัดลอกคำตอบ", true);
  }
}

function preventFullscreenExit(e) {
  if (currentExam && !document.fullscreenElement && !document.getElementById('screen-exam-session').classList.contains('hidden')) {
    handleViolation('fullscreen_exit', 'กดออกจากหน้าต่างเต็มจอ (Exit Fullscreen)');
  }
}

function handleTabVisibility() {
  if (document.hidden && currentExam) {
    handleViolation('visibility_hidden', 'สลับแท็บ/ย่อเบราว์เซอร์หรือสลับหน้าต่าง');
  }
}

function handleFocusBlur() {
  if (currentExam) {
    handleViolation('focus_lost', 'ออกจากกรอบข้อสอบ/เบราว์เซอร์สูญเสียจุดโฟกัส');
  }
}

function enableAntiCheatListeners() {
  document.addEventListener('contextmenu', preventEvent);
  document.addEventListener('selectstart', preventEvent);
  document.addEventListener('copy', preventEvent);
  document.addEventListener('cut', preventEvent);
  document.addEventListener('paste', preventEvent);
  document.addEventListener('keydown', preventKeys);
  document.addEventListener('fullscreenchange', preventFullscreenExit);
  document.addEventListener('visibilitychange', handleTabVisibility);
  window.addEventListener('blur', handleFocusBlur);
}

function disableAntiCheatListeners() {
  document.removeEventListener('contextmenu', preventEvent);
  document.removeEventListener('selectstart', preventEvent);
  document.removeEventListener('copy', preventEvent);
  document.removeEventListener('cut', preventEvent);
  document.removeEventListener('paste', preventEvent);
  document.removeEventListener('keydown', preventKeys);
  document.removeEventListener('fullscreenchange', preventFullscreenExit);
  document.removeEventListener('visibilitychange', handleTabVisibility);
  window.removeEventListener('blur', handleFocusBlur);
}

/* ==========================================================================
   6. TEACHER PORTAL & LIVE CONTROLS (REAL-TIME MONITOR)
   ========================================================================== */

async function setupTeacherDashboard() {
  document.getElementById('teacher-display-name').textContent = `คุณครู ${currentUser.name}`;
  showScreen('screen-teacher-dashboard');
  
  // Initialize dynamic tab routing
  initializeTeacherTabRouter();
  
  // Start Realtime Live Exam Monitor listener
  startLiveMonitor();
  
  // Load initial data
  await loadGradesData();
  await loadAcademicYearsData();
  await loadActivePeriodSettings();
  await loadAdminStudentsList();
  await loadAdminExamsList();
  await loadAdminTeachersList();
  await updateDBTotalStudentsStats();
  
  // Apply role security
  checkTeacherPermissions();
}

function initializeTeacherTabRouter() {
  document.querySelectorAll('.tab-link').forEach(link => {
    link.addEventListener('click', (e) => {
      // Toggle tab classes
      document.querySelectorAll('.tab-link').forEach(l => l.classList.remove('active'));
      link.classList.add('active');
      
      // Hide all contents
      document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
      
      // Show targeted content
      const targetId = link.getAttribute('data-target');
      document.getElementById(targetId).classList.remove('hidden');
    });
  });
}

function startLiveMonitor() {
  if (activeLiveMonitorUnsubscribe) activeLiveMonitorUnsubscribe();
  
  activeLiveMonitorUnsubscribe = dbManager.listenToActiveSessions((sessions) => {
    const tbody = document.getElementById('live-sessions-body');
    tbody.innerHTML = '';
    
    // Sort sessions by grade, room, no
    sessions.sort((a, b) => {
      if (a.grade !== b.grade) return a.grade.localeCompare(b.grade);
      if (a.room !== b.room) return a.room.localeCompare(b.room);
      return a.no - b.no;
    });
    
    // Filter down statistics counts
    const activeCount = sessions.filter(s => s.status === 'active').length;
    const lockedCount = sessions.filter(s => s.status === 'locked').length;
    const submittedCount = sessions.filter(s => s.status === 'submitted').length;
    
    document.getElementById('stat-active-students').textContent = activeCount;
    document.getElementById('stat-locked-students').textContent = lockedCount;
    document.getElementById('stat-submitted-students').textContent = submittedCount;
    
    if (sessions.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="8" style="padding: 2rem; color: var(--text-muted); text-align: center;">ไม่มีนักเรียนเริ่มทำข้อสอบในขณะนี้</td>
        </tr>
      `;
      return;
    }
    
    sessions.forEach((s) => {
      const tr = document.createElement('tr');
      
      // Status badge style
      let statusHtml = '';
      let actionHtml = '-';
      
      if (s.status === 'active') {
        statusHtml = `<span class="badge" style="background: rgba(42, 157, 143, 0.15); color: var(--success-color); border: 1px solid var(--success-color);"><i class="fa-solid fa-clock-rotate-left fa-spin"></i> กำลังสอบอยู่</span>`;
      } else if (s.status === 'locked') {
        statusHtml = `<span class="badge" style="background: rgba(230, 57, 70, 0.15); color: var(--danger-color); border: 1px solid var(--danger-color); font-weight: 700;"><i class="fa-solid fa-user-lock"></i> ถูกล็อก (Locked)</span>`;
        // Show unlock action button
        actionHtml = `
          <button class="btn btn-primary btn-action" 
                  style="padding: 0.35rem 0.75rem; background: var(--success-color);"
                  onclick="unlockStudentSession('${s.studentId}', '${s.examId}')">
            <i class="fa-solid fa-unlock-keyhole"></i> ปลดล็อก
          </button>
        `;
      } else if (s.status === 'submitted') {
        statusHtml = `<span class="badge" style="background: rgba(47, 79, 79, 0.08); color: var(--text-secondary);"><i class="fa-regular fa-circle-check"></i> ส่งคำตอบแล้ว (${s.score}/${s.maxScore} คะแนน)</span>`;
      }
      
      tr.innerHTML = `
        <td>${s.no}</td>
        <td style="font-family: monospace; font-weight: bold; color: var(--accent-color);">${s.studentId}</td>
        <td style="text-align: left;">${s.studentName}</td>
        <td>${s.grade}/${s.room}</td>
        <td><b>${s.subjectCode}</b> ${s.subjectName}</td>
        <td style="font-weight: 700; color: ${s.violationCount > 0 ? 'var(--danger-color)' : 'var(--text-muted)'};">${s.violationCount}/3</td>
        <td>${statusHtml}</td>
        <td>${actionHtml}</td>
      `;
      
      tbody.appendChild(tr);
    });
  });
}

// Global unlock click receiver
window.unlockStudentSession = function(studentId, examId) {
  dbManager.unlockStudentSession(studentId, examId, currentUser.name)
    .then(() => {
      showNotification(`ปลดล็อกนักเรียนสำเร็จ ให้ทำการเข้าสอบต่อได้`);
    })
    .catch(err => {
      showNotification("การปลดล็อกล้มเหลว: " + err.message, true);
    });
};

/* ==========================================================================
   7. TIMETABLE MANAGEMENT (ADD/DELETE EXAMS & DOCX PARSING)
   ========================================================================== */

const formManageExam = document.getElementById('form-manage-exam');
const btnResetExamForm = document.getElementById('btn-reset-exam-form');
const docxFileInput = document.getElementById('docx-file-input');
const docxUploadStatus = document.getElementById('docx-upload-status');
const parsedQuestionsJson = document.getElementById('parsed-questions-json');
const filterExamGrade = document.getElementById('filter-exam-grade');

let docxTargetExamId = null;

window.triggerDocxUploadForExam = function(examId) {
  docxTargetExamId = examId;
  if (docxFileInput) {
    docxFileInput.value = '';
    docxFileInput.click();
  }
};

window.triggerEditQuestionsForExam = async function(examId) {
  try {
    const exams = await dbManager.getExams();
    const exam = exams.find(e => e.examId === examId);
    if (exam) {
      docxTargetExamId = examId;
      currentParsedQuestions = exam.questions || [];
      openDocxPreviewModal();
    } else {
      showNotification("ไม่พบวิชาดังกล่าวเพื่อแก้ไขข้อสอบ", true);
    }
  } catch (err) {
    showNotification("ไม่สามารถดึงข้อมูลข้อสอบได้: " + err.message, true);
  }
};

filterExamGrade.addEventListener('change', () => loadAdminExamsList());
document.getElementById('filter-exam-year').addEventListener('change', () => loadAdminExamsList());
document.getElementById('filter-exam-term').addEventListener('change', () => loadAdminExamsList());

// Hide DOCX upload panel if Exam Type is Paper
document.getElementById('exam-type').addEventListener('change', (e) => {
  const uploadSec = document.getElementById('online-exam-upload-section');
  if (e.target.value === 'paper') {
    uploadSec.classList.add('hidden');
    // Clear parsed questions values
    parsedQuestionsJson.value = '';
    docxUploadStatus.textContent = 'สอบด้วยกระดาษคำตอบ ไม่ต้องนำเข้าข้อสอบ';
  } else {
    uploadSec.classList.remove('hidden');
    docxUploadStatus.textContent = 'อัปโหลดไฟล์ข้อสอบเวิร์ดเพื่อแปลงระบบสอบ';
  }
});

// Trigger DOCX parsing when file is selected
docxFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  
  docxUploadStatus.textContent = "กำลังวิเคราะห์โครงสร้างข้อสอบ...";
  try {
    const questions = await parseExamDocx(file);
    currentParsedQuestions = questions;
    
    // Open interactive preview modal to confirm
    openDocxPreviewModal();
    docxUploadStatus.textContent = `นำเข้าสำเร็จ: ตรวจพบข้อสอบทั้งหมด ${questions.length} ข้อ`;
  } catch (err) {
    showNotification(err.message, true);
    docxUploadStatus.textContent = "วิเคราะห์ไฟล์ล้มเหลว กรุณาลองใหม่อีกครั้ง";
    docxFileInput.value = '';
  }
});

function openDocxPreviewModal() {
  const container = document.getElementById('docx-parsed-questions-list');
  container.innerHTML = '';
  
  renderParsedQuestionsInModal();
  document.getElementById('modal-docx-preview').classList.remove('hidden');
}

function renderParsedQuestionsInModal() {
  const container = document.getElementById('docx-parsed-questions-list');
  container.innerHTML = '';
  
  if (currentParsedQuestions.length === 0) {
    container.innerHTML = '<div style="color: var(--text-muted); text-align: center; padding: 2rem;">ไม่มีรายการคำถาม กรุณากดปุ่มเพิ่มข้อสอบใหม่</div>';
    return;
  }
  
  currentParsedQuestions.forEach((q, idx) => {
    const div = document.createElement('div');
    div.className = 'preview-question-item';
    div.id = `docx-q-item-${idx}`;
    
    // Build choices options fields
    let choicesInputs = '';
    for (let cIdx = 0; cIdx < 4; cIdx++) {
      const choice = q.choices[cIdx] || { text: '' };
      // Strip choice prefix for editing convenience
      const choiceVal = choice.text.replace(/^[ก-งa-d1-4]\s*[\.\)]\s*/i, '');
      choicesInputs += `
        <div style="margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.5rem;">
          <span style="font-weight: bold; width: 30px; text-align: right;">${['ก.', 'ข.', 'ค.', 'ง.'][cIdx]}</span>
          <input type="text" class="form-control docx-choice-text-${idx}" 
                 data-choice-index="${cIdx}" value="${choiceVal}" placeholder="กรอกข้อความคำตอบที่ ${cIdx+1}" required>
        </div>
      `;
    }
    
    // Check if warning is active for correct choice
    const isUnverified = q.needsVerification;
    
    // Correct answer dropdown select
    let optionsHtml = '';
    for (let cIdx = 0; cIdx < 4; cIdx++) {
      const isSelected = q.correctChoiceIndex === cIdx;
      optionsHtml += `<option value="${cIdx}" ${isSelected ? 'selected' : ''}>ตัวเลือก ${['ก', 'ข', 'ค', 'ง'][cIdx]}</option>`;
    }
    
    const cleanQText = q.questionText.replace(/^\d+\.\s*/, '');
    
    div.innerHTML = `
      <div class="preview-question-header">
        <span class="preview-question-index"><i class="fa-solid fa-circle-question"></i> ข้อที่ ${idx + 1}</span>
        <div class="preview-question-actions">
          <button type="button" class="btn-icon" onclick="deleteParsedQuestion(${idx})" title="ลบคำถามข้อนี้"><i class="fa-regular fa-trash-can"></i></button>
        </div>
      </div>
      
      <div class="form-group">
        <input type="text" class="form-control docx-q-text" id="docx-q-text-${idx}" value="${cleanQText}" placeholder="กรอกคำถามข้อที่ ${idx + 1}" required>
      </div>
      
      ${choicesInputs}
      
      <div style="display: flex; align-items: center; gap: 1rem; margin-top: 1rem; border-top: 1px dashed rgba(255,255,255,0.05); padding-top: 0.75rem;">
        <label class="form-label" style="margin-bottom: 0;">ข้อที่เฉลยถูกต้อง:</label>
        <select class="form-control docx-q-correct" id="docx-q-correct-${idx}" style="width: auto; padding: 0.4rem 1rem;">
          ${optionsHtml}
        </select>
        ${isUnverified ? `<span style="font-size: 0.8rem; color: var(--warning-color);"><i class="fa-solid fa-triangle-exclamation"></i> ไม่พบเฉลยตัวหนา/ขีดเส้นใต้ในไฟล์เวิร์ด โปรดตรวจสอบเฉลย</span>` : ''}
      </div>
    `;
    
    container.appendChild(div);
  });
}

// Global modal delete parsed question hook
window.deleteParsedQuestion = function(index) {
  currentParsedQuestions.splice(index, 1);
  renderParsedQuestionsInModal();
};

// Add empty question manually inside DOCX modal
document.getElementById('btn-add-question-manually').addEventListener('click', () => {
  currentParsedQuestions.push({
    id: "q_manual_" + Math.random().toString(36).substring(2, 6),
    questionText: "กรอกคำถามเพิ่มเติม",
    choices: [
      { id: "c1", text: "ก. คำตอบ 1" },
      { id: "c2", text: "ข. คำตอบ 2" },
      { id: "c3", text: "ค. คำตอบ 3" },
      { id: "c4", text: "ง. คำตอบ 4" }
    ],
    correctChoiceIndex: 0,
    needsVerification: false
  });
  renderParsedQuestionsInModal();
  
  // Scroll to bottom
  const container = document.getElementById('docx-parsed-questions-list');
  setTimeout(() => { container.scrollTop = container.scrollHeight; }, 100);
});

// Close docx review modal
document.getElementById('btn-close-docx-modal').addEventListener('click', () => {
  document.getElementById('modal-docx-preview').classList.add('hidden');
  // Clear file inputs
  docxFileInput.value = '';
  docxUploadStatus.textContent = 'ยกเลิกการนำเข้าข้อสอบ';
});

// Save docx reviewed questions
document.getElementById('btn-save-docx-questions').addEventListener('click', async () => {
  // Sync form inputs back to currentParsedQuestions array
  for (let idx = 0; idx < currentParsedQuestions.length; idx++) {
    const qTextVal = document.getElementById(`docx-q-text-${idx}`).value.trim();
    const correctVal = parseInt(document.getElementById(`docx-q-correct-${idx}`).value, 10);
    
    const choiceInputs = document.querySelectorAll(`.docx-choice-text-${idx}`);
    const choices = [];
    
    choiceInputs.forEach((input) => {
      const cIdx = parseInt(input.getAttribute('data-choice-index'), 10);
      const textVal = input.value.trim();
      const indicator = ['ก', 'ข', 'ค', 'ง'][cIdx];
      choices.push({
        id: `c_${cIdx + 1}_` + Math.random().toString(36).substring(2, 5),
        text: `${indicator}. ${textVal}`
      });
    });
    
    currentParsedQuestions[idx].questionText = `${idx + 1}. ${qTextVal}`;
    currentParsedQuestions[idx].choices = choices;
    currentParsedQuestions[idx].correctChoiceIndex = correctVal;
    delete currentParsedQuestions[idx].needsVerification; // Clear warnings flags
  }
  
  if (docxTargetExamId) {
    try {
      const exams = await dbManager.getExams();
      const exam = exams.find(e => e.examId === docxTargetExamId);
      if (exam) {
        exam.questions = currentParsedQuestions;
        await dbManager.saveExam(exam);
        showNotification(`บันทึกข้อสอบในวิชา ${exam.subjectCode} สำเร็จ (${currentParsedQuestions.length} ข้อ)`);
        loadAdminExamsList();
      } else {
        showNotification("ไม่พบวิชานี้ในระบบเพื่อนำเข้าข้อสอบ", true);
      }
    } catch (err) {
      showNotification("ไม่สามารถบันทึกข้อสอบได้: " + err.message, true);
    }
    docxTargetExamId = null;
  } else {
    // Store finalized json array string in form
    parsedQuestionsJson.value = JSON.stringify(currentParsedQuestions);
    docxUploadStatus.textContent = `พร้อมบันทึก: ตรวจสอบและอนุมัติข้อสอบจำนวน ${currentParsedQuestions.length} ข้อแล้ว`;
    showNotification("ตรวจสอบและบันทึกโครงสร้างข้อสอบเรียบร้อย");
  }
  
  // Close modal
  document.getElementById('modal-docx-preview').classList.add('hidden');
});

// Timetable schedule forms submissions
formManageExam.addEventListener('submit', async (e) => {
  e.preventDefault();
  
  const examId = document.getElementById('manage-exam-id').value;
  const questionsVal = parsedQuestionsJson.value;
  let questions = [];
  
  if (document.getElementById('exam-type').value === 'online') {
    if (!questionsVal) {
      showNotification("โปรดนำเข้าข้อสอบจากไฟล์เวิร์ด .docx เพื่อทำข้อสอบออนไลน์", true);
      return;
    }
    questions = JSON.parse(questionsVal);
  }
  
  const examData = {
    date: document.getElementById('exam-date').value,
    startTime: document.getElementById('exam-start-time').value,
    endTime: document.getElementById('exam-end-time').value,
    duration: parseInt(document.getElementById('exam-duration').value, 10),
    subjectCode: document.getElementById('exam-subject-code').value.trim(),
    subjectName: document.getElementById('exam-subject-name').value.trim(),
    grade: document.getElementById('exam-grade').value,
    room: document.getElementById('exam-room').value.trim(),
    examType: document.getElementById('exam-type').value,
    questions: questions,
    academicYear: document.getElementById('exam-academic-year').value,
    term: document.getElementById('exam-term').value
  };
  
  if (examId) {
    examData.examId = examId;
  }
  
  try {
    await dbManager.saveExam(examData);
    showNotification("บันทึกข้อมูลตารางสอบลงฐานข้อมูลสำเร็จ");
    
    // Reset Form
    resetExamInputForm();
    
    // Reload schedules lists
    loadAdminExamsList();
  } catch (err) {
    showNotification("ไม่สามารถบันทึกตารางสอบได้: " + err.message, true);
  }
});

btnResetExamForm.addEventListener('click', resetExamInputForm);

function resetExamInputForm() {
  document.getElementById('manage-exam-id').value = '';
  document.getElementById('exam-date').value = '';
  document.getElementById('exam-start-time').value = '';
  document.getElementById('exam-end-time').value = '';
  document.getElementById('exam-duration').value = '';
  document.getElementById('exam-subject-code').value = '';
  document.getElementById('exam-subject-name').value = '';
  document.getElementById('exam-room').value = '';
  document.getElementById('exam-type').value = 'online';
  
  const examGradeSelect = document.getElementById('exam-grade');
  if (examGradeSelect && examGradeSelect.options.length > 0) {
    examGradeSelect.selectedIndex = 0;
  }
  
  const examYearSelect = document.getElementById('exam-academic-year');
  if (examYearSelect && examYearSelect.options.length > 0) {
    examYearSelect.selectedIndex = 0;
  }
  const examTermSelect = document.getElementById('exam-term');
  if (examTermSelect) {
    examTermSelect.value = 'ปลายภาค';
  }
  
  document.getElementById('online-exam-upload-section').classList.remove('hidden');
  docxFileInput.value = '';
  docxUploadStatus.textContent = 'อัปโหลดไฟล์ข้อสอบเวิร์ดเพื่อแปลงระบบสอบ';
  parsedQuestionsJson.value = '';
  currentParsedQuestions = [];
  
  document.getElementById('exam-form-title').innerHTML = `<i class="fa-solid fa-plus-circle"></i> เพิ่มตารางสอบ`;
}

async function loadAdminExamsList() {
  try {
    const exams = await dbManager.getExams();
    const filterGrade = filterExamGrade.value;
    const filterYear = document.getElementById('filter-exam-year').value;
    const filterTerm = document.getElementById('filter-exam-term').value;
    
    let filteredExams = exams;
    
    if (filterGrade !== 'all') {
      filteredExams = filteredExams.filter(ex => ex.grade === filterGrade);
    }
    if (filterYear !== 'all') {
      filteredExams = filteredExams.filter(ex => ex.academicYear === filterYear);
    }
    if (filterTerm !== 'all') {
      filteredExams = filteredExams.filter(ex => ex.term === filterTerm);
    }
    
    // Sort exams by date and start times
    filteredExams.sort((a, b) => {
      if (a.date !== b.date) return new Date(a.date) - new Date(b.date);
      return a.startTime.localeCompare(b.startTime);
    });
    
    const tbody = document.getElementById('admin-exams-table-body');
    tbody.innerHTML = '';
    
    if (filteredExams.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="8" style="padding: 2rem; color: var(--text-muted); text-align: center;">ไม่มีรายการตารางสอบ</td>
        </tr>
      `;
      return;
    }
    
    const isAdmin = currentUser && (currentUser.role === 'admin' || currentUser.username === 'admin');
    
    filteredExams.forEach((ex) => {
      const tr = document.createElement('tr');
      const formattedDate = formatThaiDateShort(ex.date);
      const isOnline = ex.examType === 'online';
      
      const formattedStartTime = ex.startTime.replace(':', '.');
      const formattedEndTime = ex.endTime.replace(':', '.');
      
      let linkOverrideSelect = '-';
      if (isOnline) {
        if (isAdmin) {
          linkOverrideSelect = `
            <select class="form-control" onchange="changeAdminExamLinkStatus('${ex.examId}', this.value)" style="width: auto; padding: 0.25rem 0.4rem; font-size: 0.85rem; margin: 0 auto; background: var(--input-bg); border-color: var(--input-border); color: var(--text-primary);">
              <option value="auto" ${ex.linkStatus === 'auto' || !ex.linkStatus ? 'selected' : ''}>Auto (ตามเวลา)</option>
              <option value="released" ${ex.linkStatus === 'released' ? 'selected' : ''}>ปล่อยลิงก์ทันที</option>
              <option value="hidden" ${ex.linkStatus === 'hidden' ? 'selected' : ''}>ซ่อนลิงก์ข้อสอบ</option>
            </select>
          `;
        } else {
          const statusText = ex.linkStatus === 'released' ? 'ปล่อยลิงก์แล้ว' : (ex.linkStatus === 'hidden' ? 'ซ่อนลิงก์ข้อสอบ' : 'Auto (ตามเวลา)');
          linkOverrideSelect = `<span style="font-size: 0.85rem; color: var(--text-secondary); font-weight: 500;">${statusText}</span>`;
        }
      }
      
      let actionButtons = '-';
      if (isAdmin) {
        actionButtons = `
          <div class="action-buttons-cell">
            <button class="btn btn-secondary btn-action" onclick="editAdminExam('${ex.examId}')" title="แก้ไขวิชานี้"><i class="fa-regular fa-edit"></i></button>
            <button class="btn btn-danger btn-action" onclick="deleteAdminExam('${ex.examId}')" title="ลบวิชานี้"><i class="fa-regular fa-trash-can"></i></button>
          </div>
        `;
      }
      
      tr.innerHTML = `
        <td><b>${ex.grade}</b></td>
        <td>${formattedDate}</td>
        <td>${formattedStartTime} น. - ${formattedEndTime} น.</td>
        <td style="text-align: left;"><b style="color: var(--accent-color);">${ex.subjectCode}</b> ${ex.subjectName} <span style="font-size: 0.75rem; color: var(--text-muted); display: block;">ปีการศึกษา ${ex.academicYear || '2568'} (${ex.term || 'ปลายภาค'})</span></td>
        <td>${ex.room || '-'}</td>
        <td>
          <div style="display: flex; flex-direction: column; align-items: center; gap: 0.35rem;">
            <span class="badge ${isOnline ? 'badge-link' : 'badge-paper'}" style="cursor: default; pointer-events: none; margin-bottom: 0.15rem;">
              ${isOnline ? 'Online' : 'Paper'}
            </span>
            ${isOnline ? `
              <button class="btn btn-secondary btn-action" onclick="${ex.questions && ex.questions.length > 0 ? `triggerEditQuestionsForExam('${ex.examId}')` : `triggerDocxUploadForExam('${ex.examId}')`}" 
                      style="padding: 0.25rem 0.5rem; font-size: 0.75rem; width: auto; margin: 0 auto; display: flex; align-items: center; gap: 0.25rem; font-weight: 500;" 
                      title="${ex.questions && ex.questions.length > 0 ? 'แก้ไขคำถามข้อสอบ' : 'อัปโหลดไฟล์ข้อสอบเวิร์ด'}">
                <i class="fa-solid ${ex.questions && ex.questions.length > 0 ? 'fa-file-signature' : 'fa-cloud-arrow-up'}"></i>
                ${ex.questions && ex.questions.length > 0 ? `แก้ไข (${ex.questions.length} ข้อ)` : 'อัปโหลดข้อสอบ'}
              </button>
            ` : ''}
          </div>
        </td>
        <td>${linkOverrideSelect}</td>
        <td>${actionButtons}</td>
      `;
      tbody.appendChild(tr);
    });
  } catch (err) {
    showNotification("โหลดรายการตารางสอบล้มเหลว", true);
  }
}

// Global action handles for exams list
window.editAdminExam = async function(examId) {
  try {
    const exams = await dbManager.getExams();
    const exam = exams.find(e => e.examId === examId);
    if (!exam) return;
    
    document.getElementById('manage-exam-id').value = exam.examId;
    document.getElementById('exam-date').value = exam.date;
    document.getElementById('exam-start-time').value = exam.startTime;
    document.getElementById('exam-end-time').value = exam.endTime;
    document.getElementById('exam-duration').value = exam.duration;
    document.getElementById('exam-subject-code').value = exam.subjectCode;
    document.getElementById('exam-grade').value = exam.grade;
    document.getElementById('exam-subject-name').value = exam.subjectName;
    document.getElementById('exam-room').value = exam.room || '';
    document.getElementById('exam-type').value = exam.examType;
    document.getElementById('exam-academic-year').value = exam.academicYear || '2568';
    document.getElementById('exam-term').value = exam.term || 'ปลายภาค';
    
    const uploadSec = document.getElementById('online-exam-upload-section');
    if (exam.examType === 'paper') {
      uploadSec.classList.add('hidden');
      parsedQuestionsJson.value = '';
    } else {
      uploadSec.classList.remove('hidden');
      if (exam.questions && exam.questions.length > 0) {
        parsedQuestionsJson.value = JSON.stringify(exam.questions);
        currentParsedQuestions = exam.questions;
        docxUploadStatus.textContent = `มีข้อสอบนำเข้าอยู่ในวิชาแล้ว ${exam.questions.length} ข้อ`;
      } else {
        parsedQuestionsJson.value = '';
        currentParsedQuestions = [];
        docxUploadStatus.textContent = 'ข้อสอบออนไลน์ แต่ยังไม่พบคลังคำถาม';
      }
    }
    
    document.getElementById('exam-form-title').innerHTML = `<i class="fa-solid fa-edit"></i> แก้ไขตารางสอบ`;
    showNotification("ดึงข้อมูลมาแก้ไขเรียบร้อย");
  } catch (err) {
    showNotification("ไม่สามารถดึงข้อมูลมาแก้ไขได้", true);
  }
};

window.deleteAdminExam = function(examId) {
  if (confirm("คุณแน่ใจว่าต้องการลบวิชานี้ออกจากตารางสอบใช่หรือไม่? ข้อมูลคำถามและประวัติการสอบของนักเรียนทั้งหมดในวิชานี้จะถูกลบอย่างถาวร!")) {
    dbManager.deleteExam(examId)
      .then(() => {
        showNotification("ลบรายการสอบสำเร็จ");
        loadAdminExamsList();
      })
      .catch(err => {
        showNotification("ไม่สามารถลบรายการได้: " + err.message, true);
      });
  }
};

/* ==========================================================================
   8. STUDENT DATABASE MANAGEMENT (EXCEL IMPORTER)
   ========================================================================== */

const excelFileInput = document.getElementById('excel-file-input');
const excelPreviewSection = document.getElementById('excel-preview-section');
const excelImportStats = document.getElementById('excel-import-stats');
const excelRoomsBadges = document.getElementById('excel-rooms-badges');

let currentParsedStudents = [];

excelFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  
  try {
    const students = await parseStudentExcel(file);
    currentParsedStudents = students;
    
    // Group parsed students by room to build previews stats
    const stats = {};
    students.forEach((s) => {
      const roomKey = `${s.grade}/${s.room}`;
      stats[roomKey] = (stats[roomKey] || 0) + 1;
    });
    
    // Update badge render counts
    excelRoomsBadges.innerHTML = '';
    Object.entries(stats).forEach(([room, count]) => {
      const badge = document.createElement('span');
      badge.className = 'badge badge-link';
      badge.style.cursor = 'default';
      badge.style.pointerEvents = 'none';
      badge.textContent = `ชั้น ${room} (${count} คน)`;
      excelRoomsBadges.appendChild(badge);
    });
    
    excelImportStats.textContent = `วิเคราะห์สำเร็จ: ทั้งหมด ${students.length} คน`;
    excelPreviewSection.classList.remove('hidden');
  } catch (err) {
    showNotification("วิเคราะห์ไฟล์ Excel ผิดพลาด: " + err.message, true);
    excelFileInput.value = '';
    excelPreviewSection.classList.add('hidden');
    currentParsedStudents = [];
  }
});

// Cancel Import
document.getElementById('btn-cancel-imported-students').addEventListener('click', () => {
  excelFileInput.value = '';
  excelPreviewSection.classList.add('hidden');
  currentParsedStudents = [];
});

// Commit Excel Import to Firestore
document.getElementById('btn-save-imported-students').addEventListener('click', async () => {
  if (currentParsedStudents.length === 0) return;
  
  try {
    showNotification("กำลังอัปโหลดรายชื่อนักเรียนลงฐานข้อมูล...");
    await dbManager.uploadStudentsBatch(currentParsedStudents);
    
    showNotification(`นำเข้าข้อมูลนักเรียน ${currentParsedStudents.length} คน เรียบร้อยแล้ว`);
    
    // Reset file uploads
    excelFileInput.value = '';
    excelPreviewSection.classList.add('hidden');
    currentParsedStudents = [];
    
    // Update counters
    updateDBTotalStudentsStats();
  } catch (err) {
    showNotification("เกิดข้อผิดพลาดในการอัปโหลด: " + err.message, true);
  }
});

// Clear all student database contents
document.getElementById('btn-clear-student-db').addEventListener('click', () => {
  if (confirm("คำเตือนขั้นเด็ดขาด! คุณแน่ใจว่าต้องการล้างฐานข้อมูลนักเรียนทั้งหมดในระบบใช่หรือไม่? นักเรียนทุกคนจะไม่สามารถเข้าสู่ระบบเพื่อสอบได้อีก!")) {
    dbManager.clearAllStudents()
      .then(() => {
        showNotification("ล้างข้อมูลนักเรียนเรียบร้อย");
        updateDBTotalStudentsStats();
      })
      .catch(err => {
        showNotification("ทำรายการไม่สำเร็จ: " + err.message, true);
      });
  }
});

async function updateDBTotalStudentsStats() {
  try {
    const list = await dbManager.getAllStudents();
    document.getElementById('db-total-students-count').textContent = `${list.length} คน`;
  } catch (e) {
    console.error("Failed to load total student stats:", e);
  }
}

/* ==========================================================================
   9. TEACHERS ACCOUNTS MANAGEMENT
   ========================================================================== */

const formManageTeacher = document.getElementById('form-manage-teacher');
formManageTeacher.addEventListener('submit', async (e) => {
  e.preventDefault();
  
  const user = document.getElementById('new-teacher-username').value.trim();
  const pass = document.getElementById('new-teacher-password').value.trim();
  const name = document.getElementById('new-teacher-name').value.trim();
  
  const teacherData = { username: user, password: pass, name: name };
  
  try {
    await dbManager.saveTeacher(teacherData);
    showNotification(`บันทึกข้อมูลครู ${name} สำเร็จ`);
    
    // Reset Form
    document.getElementById('new-teacher-username').value = '';
    document.getElementById('new-teacher-password').value = '';
    document.getElementById('new-teacher-name').value = '';
    
    // Reload Table
    loadAdminTeachersList();
  } catch (err) {
    showNotification("ไม่สามารถบันทึกข้อมูลครูได้: " + err.message, true);
  }
});

async function loadAdminTeachersList() {
  try {
    const list = await dbManager.getTeachers();
    const tbody = document.getElementById('teachers-table-body');
    tbody.innerHTML = '';
    
    list.forEach((t) => {
      const tr = document.createElement('tr');
      const isMainAdmin = t.username === 'admin';
      const role = t.role || (isMainAdmin ? 'admin' : 'teacher');
      
      let roleSelect = '';
      if (isMainAdmin) {
        roleSelect = `<span class="badge badge-paper" style="background: rgba(218, 165, 32, 0.15); color: var(--accent-color); border: 1px solid var(--accent-color);">ผู้ดูแลระบบหลัก (Admin)</span>`;
      } else {
        roleSelect = `
          <select class="form-control" onchange="changeTeacherAdminRole('${t.username}', this.value)" style="width: auto; padding: 0.25rem 0.4rem; font-size: 0.85rem; background: var(--input-bg); border-color: var(--input-border); color: var(--text-primary); cursor: pointer;">
            <option value="teacher" ${role === 'teacher' ? 'selected' : ''}>ครูคุมสอบ (Teacher)</option>
            <option value="admin" ${role === 'admin' ? 'selected' : ''}>ผู้ดูแลระบบ (Admin)</option>
          </select>
        `;
      }
      
      let actions = '';
      if (isMainAdmin) {
        actions = '-';
      } else {
        actions = `
          <div class="action-buttons-cell">
            <button class="btn btn-secondary btn-action" onclick="openAdminEditPasswordModal('${t.username}')" title="แก้ไขรหัสผ่าน" style="background: var(--accent-glow); border-color: var(--accent-color); color: var(--text-primary);"><i class="fa-solid fa-key"></i> รหัสผ่าน</button>
            <button class="btn btn-danger btn-action" onclick="deleteAdminTeacher('${t.username}')" title="ลบครูคนนี้"><i class="fa-regular fa-trash-can"></i> ลบ</button>
          </div>
        `;
      }
      
      tr.innerHTML = `
        <td><b>${t.username}</b></td>
        <td style="text-align: left;">${t.name}</td>
        <td>${roleSelect}</td>
        <td>${actions}</td>
      `;
      tbody.appendChild(tr);
    });
  } catch (err) {
    console.error("Failed to load teachers list:", err);
  }
}

// Global action delete teacher receiver
window.deleteAdminTeacher = function(username) {
  if (confirm(`คุณต้องการลบบัญชีคุณครู "${username}" ออกจากระบบใช่หรือไม่?`)) {
    dbManager.deleteTeacher(username)
      .then(() => {
        showNotification("ลบบัญชีครูเรียบร้อย");
        loadAdminTeachersList();
      })
      .catch(err => {
        showNotification("ไม่สามารถลบได้: " + err.message, true);
      });
  }
};

/* ==========================================================================
   10. FORMATTERS & THAI CALENDAR HELPERS
   ========================================================================== */

function formatThaiDateShort(dateStr) {
  const date = new Date(dateStr);
  const thaiMonthsShort = [
    "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
  ];
  const day = date.getDate();
  const month = thaiMonthsShort[date.getMonth()];
  const year = date.getFullYear() + 543; // BE Year
  
  return `${day} ${month} ${year}`;
}

function formatThaiDay(dateStr) {
  const date = new Date(dateStr);
  const days = [
    "อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"
  ];
  const dayName = days[date.getDay()];
  const formattedDate = formatThaiDateShort(dateStr);
  
  // Custom prefix short days for visual look in timetable
  let shortPrefix = dayName;
  if (dayName === "พฤหัสบดี") shortPrefix = "พฤหัสฯ";
  
  return `${shortPrefix}\n${date.getDate()} มี.ค. ${date.getFullYear() + 543}`;
}

// Global Exam Room trigger by ID (bypassing schedule details)
window.enterExamRoomById = async function(examId) {
  try {
    const exams = await dbManager.getExams();
    const exam = exams.find(e => e.examId === examId);
    if (exam) {
      enterExamRoom(exam);
    }
  } catch (e) {
    showNotification("ดึงข้อมูลข้อสอบล้มเหลว", true);
  }
};

// Global Link Status Change override hook
window.changeAdminExamLinkStatus = function(examId, status) {
  dbManager.updateExamLinkStatus(examId, status)
    .then(() => {
      showNotification("อัปเดตสถานะการปล่อยข้อสอบเรียบร้อย");
      loadAdminExamsList();
    })
    .catch(err => {
      showNotification("การอัปเดตล้มเหลว: " + err.message, true);
    });
};

/* ==========================================================================
   11. DYNAMIC GRADES INTERFACES
   ========================================================================== */

let currentGradesList = [];

async function loadGradesData() {
  try {
    const grades = await dbManager.getGrades();
    currentGradesList = grades;
    renderAdminGradesList(grades);
    populateGradeDropdowns(grades);
  } catch (err) {
    console.error("Failed to load grades:", err);
  }
}

function renderAdminGradesList(grades) {
  const tbody = document.getElementById('db-grades-table-body');
  tbody.innerHTML = '';
  
  grades.forEach((g) => {
    const tr = document.createElement('tr');
    const formattedDate = g.createdAt ? formatThaiDateShort(g.createdAt) : '-';
    tr.innerHTML = `
      <td><b>${g.name}</b></td>
      <td>${formattedDate}</td>
      <td>
        <button class="btn btn-danger btn-action" onclick="deleteAdminGrade('${g.name}')" title="ลบระดับชั้นนี้"><i class="fa-regular fa-trash-can"></i> ลบ</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

window.deleteAdminGrade = function(gradeName) {
  if (confirm(`คุณแน่ใจว่าต้องการลบระดับชั้น "${gradeName}" หรือไม่? ข้อมูลระดับชั้นนี้จะหายไปจากตัวเลือกทั้งหมด`)) {
    dbManager.deleteGrade(gradeName)
      .then(() => {
        showNotification("ลบระดับชั้นสำเร็จ");
        loadGradesData();
      })
      .catch(err => {
        showNotification("ลบล้มเหลว: " + err.message, true);
      });
  }
};

document.getElementById('form-manage-grade').addEventListener('submit', async (e) => {
  e.preventDefault();
  const nameInput = document.getElementById('new-grade-name');
  const name = nameInput.value.trim();
  if (!name) return;
  
  try {
    await dbManager.addGrade(name);
    showNotification(`เพิ่มระดับชั้น ${name} สำเร็จ`);
    nameInput.value = '';
    await loadGradesData();
  } catch (err) {
    showNotification(err.message, true);
  }
});

function populateGradeDropdowns(grades) {
  const examGradeSelect = document.getElementById('exam-grade');
  examGradeSelect.innerHTML = '';
  
  const filterExamGradeSelect = document.getElementById('filter-exam-grade');
  const prevFilterVal = filterExamGradeSelect.value || 'all';
  filterExamGradeSelect.innerHTML = '<option value="all">ทุกระดับชั้น</option>';
  
  const filterStudentGradeSelect = document.getElementById('filter-student-grade');
  const prevFilterStudentVal = filterStudentGradeSelect.value || 'all';
  filterStudentGradeSelect.innerHTML = '<option value="all">ทุกระดับชั้น</option>';
  
  grades.forEach((g) => {
    const opt1 = document.createElement('option');
    opt1.value = g.name;
    opt1.textContent = g.name;
    examGradeSelect.appendChild(opt1);
    
    const opt2 = document.createElement('option');
    opt2.value = g.name;
    opt2.textContent = g.name;
    filterExamGradeSelect.appendChild(opt2);
    
    const opt3 = document.createElement('option');
    opt3.value = g.name;
    opt3.textContent = g.name;
    filterStudentGradeSelect.appendChild(opt3);
  });
  
  filterExamGradeSelect.value = prevFilterVal;
  filterStudentGradeSelect.value = prevFilterStudentVal;
}

/* ==========================================================================
   12. STUDENT DATABASE VIEWER & LIST BROWSING
   ========================================================================== */

let currentStudentsList = [];

async function loadAdminStudentsList() {
  try {
    const list = await dbManager.getAllStudents();
    currentStudentsList = list;
    renderFilteredStudents();
  } catch (err) {
    console.error("Failed to load students:", err);
  }
}

function renderFilteredStudents() {
  const searchQuery = document.getElementById('search-student-query').value.toLowerCase().trim();
  const filterGrade = document.getElementById('filter-student-grade').value;
  const filterRoom = document.getElementById('filter-student-room').value.trim();
  
  let filtered = currentStudentsList;
  
  if (searchQuery) {
    filtered = filtered.filter(s => 
      s.studentId.includes(searchQuery) || 
      s.name.toLowerCase().includes(searchQuery)
    );
  }
  
  if (filterGrade !== 'all') {
    filtered = filtered.filter(s => s.grade === filterGrade);
  }
  
  if (filterRoom) {
    filtered = filtered.filter(s => s.room === filterRoom);
  }
  
  filtered.sort((a, b) => {
    if (a.grade !== b.grade) return a.grade.localeCompare(b.grade);
    if (a.room !== b.room) return a.room.localeCompare(b.room);
    return a.no - b.no;
  });
  
  const tbody = document.getElementById('db-students-table-body');
  tbody.innerHTML = '';
  
  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" style="padding: 2rem; color: var(--text-muted); text-align: center;">ไม่พบข้อมูลนักเรียนที่ค้นหา</td>
      </tr>
    `;
    return;
  }
  
  const isAdmin = currentUser && currentUser.username === 'admin';
  filtered.forEach((s) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${s.no}</td>
      <td style="font-family: monospace; font-weight: 700; color: var(--accent-color);">${s.studentId}</td>
      <td style="text-align: left;">${s.name}</td>
      <td>${s.grade}/${s.room}</td>
      <td style="font-family: monospace; font-size: 0.9rem; color: var(--text-muted);">${s.password}</td>
      <td>
        ${isAdmin ? `<button class="btn btn-danger btn-action" onclick="deleteAdminStudent('${s.studentId}')" title="ลบรายชื่อนี้"><i class="fa-regular fa-trash-can"></i> ลบ</button>` : '-'}
      </td>
    `;
    tbody.appendChild(tr);
  });
}

window.deleteAdminStudent = function(studentId) {
  if (confirm(`คุณแน่ใจว่าต้องการลบนักเรียนรหัส "${studentId}" ออกจากฐานข้อมูลหรือไม่?`)) {
    dbManager.deleteStudent(studentId)
      .then(() => {
        showNotification("ลบรายชื่อนักเรียนสำเร็จ");
        updateDBTotalStudentsStats();
        loadAdminStudentsList();
      })
      .catch(err => {
        showNotification("ลบล้มเหลว: " + err.message, true);
      });
  }
};

document.getElementById('search-student-query').addEventListener('input', renderFilteredStudents);
document.getElementById('filter-student-grade').addEventListener('change', renderFilteredStudents);
document.getElementById('filter-student-room').addEventListener('input', renderFilteredStudents);

/* ==========================================================================
   13. DYNAMIC ACADEMIC YEARS INTERFACES
   ========================================================================== */

let currentYearsList = [];
let activePeriodConfig = null;

async function loadAcademicYearsData() {
  try {
    const years = await dbManager.getAcademicYears();
    currentYearsList = years;
    renderAdminYearsList(years);
    populateYearDropdowns(years);
  } catch (err) {
    console.error("Failed to load academic years:", err);
  }
}

function renderAdminYearsList(years) {
  const tbody = document.getElementById('db-years-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';
  
  years.forEach((y) => {
    const tr = document.createElement('tr');
    const formattedDate = y.createdAt ? formatThaiDateShort(y.createdAt) : '-';
    tr.innerHTML = `
      <td><b>${y.name}</b></td>
      <td>${formattedDate}</td>
      <td>
        <button class="btn btn-danger btn-action" onclick="deleteAdminYear('${y.name}')" title="ลบปีการศึกษานี้"><i class="fa-regular fa-trash-can"></i> ลบ</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

window.deleteAdminYear = function(yearName) {
  if (confirm(`คุณแน่ใจว่าต้องการลบปีการศึกษา "${yearName}" หรือไม่? ข้อมูลตารางสอบที่เกี่ยวข้องจะยังคงอยู่ในระบบแต่ปีนี้จะหายไปจากตัวเลือก`)) {
    dbManager.deleteAcademicYear(yearName)
      .then(() => {
        showNotification("ลบปีการศึกษาสำเร็จ");
        loadAcademicYearsData();
      })
      .catch(err => {
        showNotification("ลบล้มเหลว: " + err.message, true);
      });
  }
};

const formManageYear = document.getElementById('form-manage-year');
if (formManageYear) {
  formManageYear.addEventListener('submit', async (e) => {
    e.preventDefault();
    const nameInput = document.getElementById('new-year-name');
    const name = nameInput.value.trim();
    if (!name) return;
    
    try {
      await dbManager.addAcademicYear(name);
      showNotification(`เพิ่มปีการศึกษา ${name} สำเร็จ`);
      nameInput.value = '';
      await loadAcademicYearsData();
    } catch (err) {
      showNotification(err.message, true);
    }
  });
}

function populateYearDropdowns(years) {
  const examYearSelect = document.getElementById('exam-academic-year');
  if (examYearSelect) {
    examYearSelect.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y.name;
      opt.textContent = y.name;
      examYearSelect.appendChild(opt);
    });
  }
  
  const filterYearSelect = document.getElementById('filter-exam-year');
  if (filterYearSelect) {
    const prevVal = filterYearSelect.value || 'all';
    filterYearSelect.innerHTML = '<option value="all">ทุกปีการศึกษา</option>';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y.name;
      opt.textContent = y.name;
      filterYearSelect.appendChild(opt);
    });
    filterYearSelect.value = prevVal;
  }
  
  const activeYearSelect = document.getElementById('active-setting-year');
  if (activeYearSelect) {
    activeYearSelect.innerHTML = '';
    years.forEach(y => {
      const opt = document.createElement('option');
      opt.value = y.name;
      opt.textContent = y.name;
      activeYearSelect.appendChild(opt);
    });
    if (activePeriodConfig) {
      activeYearSelect.value = activePeriodConfig.activeYear;
    }
  }
}

/* ==========================================================================
   14. ACTIVE EXAM PERIOD INTERFACES & PERMISSIONS
   ========================================================================== */

async function loadActivePeriodSettings() {
  try {
    const config = await dbManager.getActiveExamPeriod();
    activePeriodConfig = config;
    
    const activeYearSelect = document.getElementById('active-setting-year');
    const activeTermSelect = document.getElementById('active-setting-term');
    
    if (activeYearSelect) activeYearSelect.value = config.activeYear;
    if (activeTermSelect) activeTermSelect.value = config.activeTerm;
  } catch (err) {
    console.error("Failed to load active period settings:", err);
  }
}

const formActivePeriod = document.getElementById('form-active-period-settings');
if (formActivePeriod) {
  formActivePeriod.addEventListener('submit', async (e) => {
    e.preventDefault();
    const activeYear = document.getElementById('active-setting-year').value;
    const activeTerm = document.getElementById('active-setting-term').value;
    
    try {
      await dbManager.saveActiveExamPeriod(activeYear, activeTerm);
      showNotification("บันทึกการตั้งค่าช่วงเวลาสอบปัจจุบันสำเร็จ");
      await loadActivePeriodSettings();
    } catch (err) {
      showNotification("บันทึกล้มเหลว: " + err.message, true);
    }
  });
}

function checkTeacherPermissions() {
  const isAdmin = currentUser && (currentUser.role === 'admin' || currentUser.username === 'admin');
  
  // Tab elements
  const tabImport = document.getElementById('tab-btn-students-import');
  const tabGradesYears = document.getElementById('tab-btn-grades-years');
  const tabTeachers = document.getElementById('tab-btn-teachers');
  const examFormContainer = document.getElementById('admin-exam-form-container');
  const timetableGrid = document.getElementById('timetable-tab-grid');
  const btnClearDb = document.getElementById('btn-clear-student-db');
  
  if (isAdmin) {
    if (tabImport) tabImport.classList.remove('hidden');
    if (tabGradesYears) tabGradesYears.classList.remove('hidden');
    if (tabTeachers) tabTeachers.classList.remove('hidden');
    if (examFormContainer) examFormContainer.classList.remove('hidden');
    if (timetableGrid) timetableGrid.style.gridTemplateColumns = '350px 1fr';
    if (btnClearDb) btnClearDb.classList.remove('hidden');
  } else {
    if (tabImport) tabImport.classList.add('hidden');
    if (tabGradesYears) tabGradesYears.classList.add('hidden');
    if (tabTeachers) tabTeachers.classList.add('hidden');
    if (examFormContainer) examFormContainer.classList.add('hidden');
    if (timetableGrid) timetableGrid.style.gridTemplateColumns = '1fr';
    if (btnClearDb) btnClearDb.classList.add('hidden');
  }
}

/* ==========================================================================
   15. TEACHER ROLE & PASSWORD EDIT Event Handlers
   ========================================================================== */

window.changeTeacherAdminRole = function(username, role) {
  dbManager.updateTeacherRole(username, role)
    .then(() => {
      showNotification(`อัปเดตบทบาทของคุณครู "${username}" เรียบร้อยแล้ว`);
      
      // If updating oneself, reload permissions
      if (currentUser && currentUser.username === username) {
        currentUser.role = role;
        sessionStorage.setItem('kss_user', JSON.stringify(currentUser));
        checkTeacherPermissions();
      }
      loadAdminTeachersList();
    })
    .catch(err => {
      showNotification("อัปเดตบทบาทล้มเหลว: " + err.message, true);
    });
};

// Change own password modal triggers
const modalOwnPass = document.getElementById('modal-change-password-own');
const formOwnPass = document.getElementById('form-change-password-own');

document.getElementById('btn-teacher-change-password-own').addEventListener('click', () => {
  if (modalOwnPass) modalOwnPass.classList.remove('hidden');
});

document.getElementById('btn-close-change-password-own').addEventListener('click', () => {
  if (modalOwnPass) {
    modalOwnPass.classList.add('hidden');
    formOwnPass.reset();
  }
});

if (formOwnPass) {
  formOwnPass.addEventListener('submit', async (e) => {
    e.preventDefault();
    const oldPass = document.getElementById('change-pass-old').value;
    const newPass = document.getElementById('change-pass-new').value;
    const confirmPass = document.getElementById('change-pass-confirm').value;
    
    if (newPass !== confirmPass) {
      showNotification("รหัสผ่านใหม่ไม่ตรงกับการยืนยัน", true);
      return;
    }
    
    if (newPass.length < 6) {
      showNotification("รหัสผ่านใหม่ต้องมีความยาวอย่างน้อย 6 ตัวอักษร", true);
      return;
    }
    
    try {
      await dbManager.changeOwnPassword(currentUser.username, oldPass, newPass);
      showNotification("เปลี่ยนรหัสผ่านส่วนตัวสำเร็จเรียบร้อยแล้ว");
      
      // Update session values
      currentUser.password = newPass;
      sessionStorage.setItem('kss_user', JSON.stringify(currentUser));
      
      // Hide modal and reset form
      modalOwnPass.classList.add('hidden');
      formOwnPass.reset();
    } catch (err) {
      showNotification(err.message, true);
    }
  });
}

// Admin change other teacher password modal triggers
const modalAdminEditPass = document.getElementById('modal-admin-edit-password');
const formAdminEditPass = document.getElementById('form-admin-edit-password');

window.openAdminEditPasswordModal = function(username) {
  document.getElementById('admin-edit-pass-username').textContent = username;
  document.getElementById('admin-edit-pass-target-user').value = username;
  if (modalAdminEditPass) modalAdminEditPass.classList.remove('hidden');
};

document.getElementById('btn-close-admin-edit-password').addEventListener('click', () => {
  if (modalAdminEditPass) {
    modalAdminEditPass.classList.add('hidden');
    formAdminEditPass.reset();
  }
});

if (formAdminEditPass) {
  formAdminEditPass.addEventListener('submit', async (e) => {
    e.preventDefault();
    const targetUser = document.getElementById('admin-edit-pass-target-user').value;
    const newPass = document.getElementById('admin-edit-pass-new').value;
    
    if (newPass.length < 6) {
      showNotification("รหัสผ่านต้องมีความยาวอย่างน้อย 6 ตัวอักษร", true);
      return;
    }
    
    try {
      await dbManager.updateTeacherPassword(targetUser, newPass);
      showNotification(`แก้ไขรหัสผ่านของคุณครู "${targetUser}" สำเร็จแล้ว`);
      if (modalAdminEditPass) modalAdminEditPass.classList.add('hidden');
      formAdminEditPass.reset();
      loadAdminTeachersList();
    } catch (err) {
      showNotification("การแก้ไขล้มเหลว: " + err.message, true);
    }
  });
}

// Init bootstrap app load
window.addEventListener('DOMContentLoaded', () => {
  const savedUser = sessionStorage.getItem('kss_user');
  const savedRole = sessionStorage.getItem('kss_role');
  
  if (savedUser && savedRole) {
    currentUser = JSON.parse(savedUser);
    currentRole = savedRole;
    
    if (savedRole === 'student') {
      setupStudentDashboard();
    } else if (savedRole === 'teacher') {
      setupTeacherDashboard();
    }
  } else {
    checkFirebaseConnection();
  }
});
