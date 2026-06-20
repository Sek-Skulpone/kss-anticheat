// docx-parser.js - Exam Docx Parser using Mammoth.js

/**
 * Parses an exam DOCX file and extracts questions and choices.
 * @param {File} file - The uploaded DOCX file.
 * @returns {Promise<Array>} List of parsed question objects.
 */
export function parseExamDocx(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const arrayBuffer = e.target.result;
      
      // Convert docx to HTML using Mammoth to preserve strong/bold and underline tags
      mammoth.convertToHtml({ arrayBuffer: arrayBuffer })
        .then(function(result) {
          const html = result.value;
          const questions = extractQuestionsFromHtml(html);
          if (questions.length === 0) {
            reject(new Error("ไม่สามารถแยกข้อสอบได้ กรุณาตรวจสอบรูปแบบไฟล์ข้อสอบ (เช่น ข้อ 1. และตัวเลือก ก. ข. ค. ง.)"));
          } else {
            resolve(questions);
          }
        })
        .catch(function(err) {
          reject(err);
        });
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Parses Mammoth's output HTML and builds structured questions and choices
 * @param {string} html - HTML output from Mammoth
 * @returns {Array} List of questions
 */
function extractQuestionsFromHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  // Match both normal paragraph tags and list item tags
  const paragraphs = Array.from(doc.querySelectorAll('p, li'));
  
  const questions = [];
  let currentQuestion = null;
  
  // Patterns
  // Questions start with a number like: 1., 1), ข้อ 1., 10.
  const questionRegex = /^(?:ข้อ\s*)?(\d+)\s*[\.\)]\s*(.*)$/i;
  // Choices start with Thai letters (ก, ข, ค, ง), English (a, b, c, d / A, B, C, D), or numbers 1-4
  const choiceRegex = /^\s*([ก-งa-dขคA-D1-4])\s*[\.\)\-]\s*(.*)$/;
  
  paragraphs.forEach((p) => {
    const text = p.textContent.trim();
    if (!text) return;
    
    // Check if it matches question pattern
    const qMatch = text.match(questionRegex);
    const cMatch = text.match(choiceRegex);
    
    // We determine if qMatch is indeed a question. 
    // Sometimes choices are numeric (e.g. 1. choice), which matches qMatch too.
    // If it's a choices list, the numbers are usually 1, 2, 3, 4. 
    // If we already have a question and the matched number is between 1 and 4, and it is sequential, we treat it as a choice.
    let isChoice = false;
    if (cMatch) {
      isChoice = true;
      // Exception: If choice is numeric (1-4) but we don't have an active question, treat it as a question
      if (/^\d+$/.test(cMatch[1]) && !currentQuestion) {
        isChoice = false;
      }
    }
    
    if (qMatch && !isChoice) {
      // It's a new question! Save current active question
      if (currentQuestion && currentQuestion.choices.length > 0) {
        questions.push(currentQuestion);
      }
      
      const qText = qMatch[2] ? qMatch[2].trim() : text;
      const qNum = qMatch[1];
      
      currentQuestion = {
        id: "q_" + (questions.length + 1) + "_" + Math.random().toString(36).substring(2, 5),
        questionText: `${qNum}. ${qText}`,
        choices: [],
        correctChoiceIndex: -1
      };
    } else if (cMatch && currentQuestion) {
      // It's a choice for the current question
      const indicator = cMatch[1].trim();
      const rawChoiceText = cMatch[2] ? cMatch[2].trim() : "";
      
      // Determine if this choice is the correct answer
      // Check 1: Strong/Bold tags anywhere in the paragraph
      const isBold = p.querySelector('strong, b') !== null;
      // Check 2: Underline/Insert tags
      const isUnderline = p.querySelector('u, ins') !== null;
      // Check 3: Check if the text ends with an asterisk '*'
      const endsWithAsterisk = rawChoiceText.endsWith('*');
      
      let cleanChoiceText = rawChoiceText;
      if (endsWithAsterisk) {
        cleanChoiceText = rawChoiceText.slice(0, -1).trim();
      }
      
      const choiceId = `c_${currentQuestion.choices.length + 1}_` + Math.random().toString(36).substring(2, 5);
      
      currentQuestion.choices.push({
        id: choiceId,
        text: `${indicator}. ${cleanChoiceText}`
      });
      
      if (isBold || isUnderline || endsWithAsterisk) {
        currentQuestion.correctChoiceIndex = currentQuestion.choices.length - 1;
      }
    } else if (currentQuestion) {
      // It's continuation text (e.g. image placeholder description, long question text, or multi-line choice)
      if (currentQuestion.choices.length === 0) {
        // Append to question text
        currentQuestion.questionText += "\n" + text;
      } else {
        // Append to the last choice text
        const lastChoice = currentQuestion.choices[currentQuestion.choices.length - 1];
        lastChoice.text += "\n" + text;
      }
    }
  });
  
  // Push the final question
  if (currentQuestion && currentQuestion.choices.length > 0) {
    questions.push(currentQuestion);
  }
  
  // Post-process: ensure all questions have a default correct index (0) if none was detected
  questions.forEach(q => {
    if (q.correctChoiceIndex === -1) {
      q.correctChoiceIndex = 0; // Default to first option
      q.needsVerification = true; // Mark to warn the teacher to check
    }
  });
  
  return questions;
}
