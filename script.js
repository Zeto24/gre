// script.js

// --- DATA SIMULASI EXCEL ---
// ... (tidak ada perubahan di bagian ini) ...
const excelData = [
  // A        B           C           D           E           F           G
  [
    "No.",
    "Nama Produk",
    "Kategori",
    "Harga Satuan",
    "Jumlah Terjual",
    "Diskon (%)",
    "Tanggal Penjualan",
  ], // Baris 1 (Header)
  [1, "Laptop XYZ", "Elektronik", 12500000, 3, 0.05, "2025-01-15"], // Baris 2
  [2, "Keyboard Mech", "Aksesoris", 850000, 5, 0, "2025-01-18"], // Baris 3
  [3, "Mouse Gaming", "Aksesoris", 400000, 8, 0.1, "2025-01-18"], // Baris 4
  [4, "Monitor LED", "Elektronik", 3200000, 2, 0, "2025-01-20"], // Baris 5
  [5, "Headset Nirkabel", "Aksesoris", 1500000, 4, 0.05, "2025-01-22"], // Baris 6
  [6, "Printer Laser", "Elektronik", 4800000, 1, 0, "2025-01-25"], // Baris 7
  [7, "SSD 1TB", "Komponen", 1200000, 6, 0.15, "2025-01-28"], // Baris 8
  [8, "RAM 16GB", "Komponen", 900000, 7, 0.1, "2025-01-28"], // Baris 9
];

const sheet2Data = [
  ["Kode Kategori", "Deskripsi Kategori"],
  ["EL", "Elektronik"],
  ["AK", "Aksesoris"],
  ["KO", "Komponen"],
];

// --- FUNGSI SIMULASI EXCEL UNTUK MENGHITUNG EXPECTED RESULT ---
// ... (tidak ada perubahan di bagian ini) ...
function getVal(rowIdx, colIdx) {
  if (excelData[rowIdx] && excelData[rowIdx][colIdx] !== undefined) {
    return excelData[rowIdx][colIdx];
  }
  return undefined;
}

function getRangeVals(startRef, endRef, dataSrc = excelData) {
  const startCol = startRef.charCodeAt(0) - "A".charCodeAt(0);
  const startRow = parseInt(startRef.substring(1), 10) - 1;
  const endCol = endRef.charCodeAt(0) - "A".charCodeAt(0);
  const endRow = parseInt(endRef.substring(1), 10) - 1;

  const values = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      if (
        dataSrc[r] &&
        dataSrc[r][c] !== undefined &&
        (typeof dataSrc[r][c] === "number" || typeof dataSrc[r][c] === "string")
      ) {
        values.push(dataSrc[r][c]);
      }
    }
  }
  return values;
}

const calculateExcelFunction = {
  SUM: (range) => {
    const numbers = getRangeVals(
      range.split(":")[0],
      range.split(":")[1]
    ).filter((val) => typeof val === "number");
    return numbers.reduce((acc, val) => acc + val, 0);
  },
  AVERAGE: (range) => {
    const numbers = getRangeVals(
      range.split(":")[0],
      range.split(":")[1]
    ).filter((val) => typeof val === "number");
    if (numbers.length === 0) return 0;
    return calculateExcelFunction.SUM(range) / numbers.length;
  },
  COUNT: (range) => {
    return getRangeVals(range.split(":")[0], range.split(":")[1]).length;
  },
  MAX: (range) => {
    const numbers = getRangeVals(
      range.split(":")[0],
      range.split(":")[1]
    ).filter((val) => typeof val === "number");
    if (numbers.length === 0) return 0;
    return Math.max(...numbers);
  },
  MIN: (range) => {
    const numbers = getRangeVals(
      range.split(":")[0],
      range.split(":")[1]
    ).filter((val) => typeof val === "number");
    if (numbers.length === 0) return 0;
    return Math.min(...numbers);
  },
  IF: (condition, valueIfTrue, valueIfFalse) => {
    return condition ? valueIfTrue : valueIfFalse;
  },
  VLOOKUP: (
    lookup_value,
    table_array_ref,
    col_index_num,
    range_lookup = false
  ) => {
    for (let i = 1; i < sheet2Data.length; i++) {
      if (sheet2Data[i][1] === lookup_value) {
        return sheet2Data[i][0];
      }
    }
    return "#N/A";
  },
  CONCAT: (val1, val2, val3) => {
    return `${val1}${val2}${val3}`;
  },
  MONTH: (dateString) => {
    const date = new Date(dateString);
    return date.getMonth() + 1;
  },
  TODAY: () => {
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth(), today.getDate());
  },
  DATEDIF: (startDateObj, endDateObj, unit) => {
    const diffTime = Math.abs(endDateObj.getTime() - startDateObj.getTime());
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    if (unit.toUpperCase() === "D") return diffDays;
    return diffDays;
  },
  DATE_PARSE: (dateString) => {
    const [year, month, day] = dateString.split("-").map(Number);
    return new Date(year, month - 1, day);
  },
};

// --- DATA SOAL PER NOMOR ---
// ... (tidak ada perubahan di bagian ini, sudah benar) ...
const questions = [
  {
    question: "Berapa **Total Harga (Sebelum Diskon)** untuk 'Laptop XYZ'?",
    formulaExample: "H2: =D2*E2",
    resultCell: "H2",
    expectedResult: () => getVal(1, 3) * getVal(1, 4),
    type: "number",
  },
  {
    question: "Berapa **Total Harga (Setelah Diskon)** untuk 'Laptop XYZ'?",
    formulaExample: "I2: =H2-(H2*F2)",
    resultCell: "I2",
    expectedResult: () =>
      getVal(1, 3) * getVal(1, 4) - getVal(1, 3) * getVal(1, 4) * getVal(1, 5),
    type: "number",
  },
  {
    question: "Berapa **total penjualan semua produk sebelum diskon**?",
    formulaExample: "H10: =SUM(H2:H9)",
    resultCell: "H10",
    expectedResult: () =>
      getVal(1, 3) * getVal(1, 4) +
      getVal(2, 3) * getVal(2, 4) +
      getVal(3, 3) * getVal(3, 4) +
      getVal(4, 3) * getVal(4, 4) +
      getVal(5, 3) * getVal(5, 4) +
      getVal(6, 3) * getVal(6, 4) +
      getVal(7, 3) * getVal(7, 4) +
      getVal(8, 3) * getVal(8, 4),
    type: "number",
  },
  {
    question: "Berapa **rata-rata 'Harga Satuan'** dari semua produk?",
    formulaExample: "D10: =AVERAGE(D2:D9)",
    resultCell: "D10",
    expectedResult: () => calculateExcelFunction.AVERAGE("D2:D9"),
    type: "number",
  },
  {
    question: "Apa **'Status Diskon'** untuk 'Laptop XYZ'?",
    formulaExample: 'J2: =JIKA(F2>0;"Ada Diskon";"Tidak Ada Diskon")',
    resultCell: "J2",
    expectedResult: () =>
      calculateExcelFunction.IF(
        getVal(1, 5) > 0,
        "Ada Diskon",
        "Tidak Ada Diskon"
      ),
    type: "string",
  },
  {
    question: "Apa **'ID Produk'** untuk 'Laptop XYZ'?",
    formulaExample: 'M2: =C2&"-"&A2',
    resultCell: "M2",
    expectedResult: () =>
      calculateExcelFunction.CONCAT(getVal(1, 2), "-", getVal(1, 0)),
    type: "string",
  },
  {
    question:
      "Apa **'Kode Kategori'** untuk produk 'Laptop XYZ' jika dicari dari Sheet2?",
    formulaExample: "L2: =INDEX(Sheet2!$A:$A;MATCH(C2;Sheet2!$B:$B;0))",
    resultCell: "L2",
    expectedResult: () =>
      calculateExcelFunction.VLOOKUP(
        getVal(1, 2),
        "Sheet2!$A$1:$B$4",
        1,
        false
      ),
    type: "string",
  },
  {
    question: "Berapa **bulan penjualan** untuk 'Laptop XYZ'?",
    formulaExample: "N2: =MONTH(G2)",
    resultCell: "N2",
    expectedResult: () => calculateExcelFunction.MONTH(getVal(1, 6)),
    type: "number",
  },
  {
    question:
      "Berapa **'Usia Data (Hari)'** untuk 'Laptop XYZ' hingga hari ini?",
    formulaExample: "O2: =TODAY()-G2",
    resultCell: "O2",
    expectedResult: () => {
      const startDate = calculateExcelFunction.DATE_PARSE(getVal(1, 6));
      const today = calculateExcelFunction.TODAY();
      return calculateExcelFunction.DATEDIF(startDate, today, "D");
    },
    type: "number",
  },
];

let currentQuestionIndex = 0;
const feedbackElement = document.getElementById("feedback");
const questionTextElement = document.getElementById("question-text");
const userResultInput = document.getElementById("user-result-input");
const checkBtn = document.getElementById("check-answer-btn");
const nextBtn = document.getElementById("next-question-btn");
const currentLevelSpan = document.getElementById("current-level");

// --- FUNGSI UTAMA GAME ---

function loadQuestion() {
  if (currentQuestionIndex < questions.length) {
    const q = questions[currentQuestionIndex];
    // Pastikan formulaExample ada sebelum mencoba menampilkannya
    let formulaHtml = "";
    if (q.formulaExample) {
      formulaHtml = `<br><small><em>(Rumus di sel ${q.formulaExample})</em></small>`;
    }

    questionTextElement.innerHTML = `
          **Soal ${currentQuestionIndex + 1}:** ${q.question} ${formulaHtml}
      `;
    userResultInput.value = "";
    feedbackElement.textContent = "";
    feedbackElement.className = "feedback";
    nextBtn.style.display = "none";
    checkBtn.style.display = "inline-block";
    currentLevelSpan.textContent = currentQuestionIndex + 1;
  } else {
    questionTextElement.textContent =
      "Selamat! Anda telah menyelesaikan semua soal!";
    userResultInput.style.display = "none";
    checkBtn.style.display = "none";
    nextBtn.style.display = "none";
    feedbackElement.className = "feedback correct";
    feedbackElement.textContent =
      "Anda sangat mahir dalam rumus Excel! Anda bisa mencoba kembali atau melanjutkan belajar Excel di aplikasi aslinya.";
  }
}

function checkAnswer() {
  const userInputResult = userResultInput.value.trim();
  const currentQuestion = questions[currentQuestionIndex];

  if (!userInputResult) {
    feedbackElement.className = "feedback incorrect";
    feedbackElement.textContent = "Silakan masukkan hasil rumus Anda.";
    return;
  }

  let actualExpectedResult = currentQuestion.expectedResult();

  let isCorrect = false;
  if (currentQuestion.type === "number") {
    const parsedUserInput = parseFloat(
      userInputResult.replace(/\./g, "").replace(/,/g, ".")
    );
    if (!isNaN(parsedUserInput)) {
      isCorrect = Math.abs(parsedUserInput - actualExpectedResult) < 0.01;
    }
  } else if (currentQuestion.type === "string") {
    isCorrect =
      userInputResult.toLowerCase() === actualExpectedResult.toLowerCase();
  } else {
    isCorrect = userInputResult == actualExpectedResult;
  }

  if (isCorrect) {
    feedbackElement.className = "feedback correct";
    // Pastikan formulaExample ada sebelum mencoba mengambil bagiannya
    let displayedFormula = currentQuestion.formulaExample
      ? currentQuestion.formulaExample.split(": ")[1]
      : "Tidak ada contoh rumus";
    feedbackElement.innerHTML = `Benar! Hasil Anda tepat.`;
    nextBtn.style.display = "inline-block";
    checkBtn.style.display = "none";
  } else {
    feedbackElement.className = "feedback incorrect";
    let displayedFormula = currentQuestion.formulaExample
      ? currentQuestion.formulaExample.split(": ")[1]
      : "Tidak ada contoh rumus";
    feedbackElement.innerHTML = `
          Salah. Hasil Anda <code>${userInputResult}</code> tidak sesuai harapan. <br>
          Hasil yang benar adalah <code>${
            typeof actualExpectedResult === "number"
              ? actualExpectedResult.toLocaleString("id-ID")
              : actualExpectedResult
          }</code>.Coba lagi!
      `;
  }
}

function nextQuestion() {
  currentQuestionIndex++;
  loadQuestion();
}

// Event Listeners
checkBtn.addEventListener("click", checkAnswer);
nextBtn.addEventListener("click", nextQuestion);

// Inisialisasi game
loadQuestion();
